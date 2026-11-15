# -*- coding: utf-8 -*-
"""
Core exporter for wechat-account-exporter.

This module reads **plaintext SQLite databases that the user has already
decrypted by themselves** and exports the official-account (公众号) list to
Word / Excel / CSV and an AES-256-GCM encrypted HTML page.

It contains NO code that decrypts WeChat databases, reads process memory,
attaches to the WeChat client, or calls any WeChat API. It never uploads data.
The only optional network access is downloading avatar images, which is off by
default and restricted to an explicit host allow-list.

Word/Excel files are written as OOXML directly with the standard library, so
the only third-party dependency is pycryptodome (for the encrypted web page).
"""
from __future__ import annotations

import base64
import csv
import datetime
import hashlib
import json
import os
import re
import sqlite3
import urllib.request
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from xml.sax.saxutils import escape

try:
    from Crypto.Cipher import AES
except ImportError:  # pycryptodome is optional unless --web is used
    AES = None

FONT = 'Microsoft YaHei'
EMU = 360000  # 1.0 cm
COLS = [500, 850, 1750, 1650, 1350, 3100, 1250, 4550]
W = sum(COLS)
DEFAULT_HOSTS = ('mmbiz.qpic.cn',)


class ExportError(Exception):
    """A user-facing, actionable error."""


@dataclass
class ExportOptions:
    db_dir: str                      # directory containing the plaintext DBs
    out_dir: str                     # directory for generated files
    password: str = ''               # required when make_web is True
    make_docx: bool = True
    make_xlsx: bool = True
    make_csv: bool = True
    make_web: bool = True
    fetch_avatars: bool = False      # opt-in; only DEFAULT_HOSTS are contacted
    allowed_hosts: tuple = DEFAULT_HOSTS


def clean(v):
    if v is None:
        return ''
    if isinstance(v, bytes):
        v = v.decode('utf-8', errors='replace')
    return ''.join(ch for ch in str(v) if ch >= ' ' or ch in '\t\n\r')


def ro(path):
    con = sqlite3.connect('file:' + path.replace('\\', '/') + '?mode=ro', uri=True)
    con.text_factory = lambda b: b.decode('utf-8', errors='replace')
    return con


def _find(db_dir, *parts):
    """Accept either <db_dir>/<parts> or <db_dir>/<last part>."""
    p = os.path.join(db_dir, *parts)
    if os.path.exists(p):
        return p
    p2 = os.path.join(db_dir, parts[-1])
    return p2 if os.path.exists(p2) else None


def _is_plain_sqlite(path):
    try:
        with open(path, 'rb') as f:
            return f.read(16).startswith(b'SQLite format 3')
    except OSError:
        return False


def _is_image(d):
    return (d[:8] == b'\x89PNG\r\n\x1a\n' or d[:3] == b'\xff\xd8\xff'
            or d[:6] in (b'GIF87a', b'GIF89a'))


def build_desc(raw):
    """Re-state verifiable local fields; there is no official 简介 field."""
    try:
        d = json.loads(raw)
    except Exception:
        return ''
    parts = []
    pv = d.get('PersonVerifyInfo')
    if isinstance(pv, dict):
        s = pv.get('VerifyDescribe') or pv.get('VerifyDesc')
        if s:
            parts.append('认证说明：' + clean(s).strip())
    vs = d.get('VerifySource')
    if isinstance(vs, dict) and vs.get('Description'):
        parts.append('认证主体：' + clean(vs['Description']).strip())
    rs = d.get('RegisterSource')
    if isinstance(rs, dict) and rs.get('RegisterBody'):
        parts.append('注册主体：' + clean(rs['RegisterBody']).strip())
    if d.get('TrademarkName'):
        parts.append('商标：' + clean(d['TrademarkName']).strip())
    return '；'.join(parts)[:220]


def _collect(opts, log, prog):
    contact_db = _find(opts.db_dir, 'contact', 'contact.db')
    if not contact_db:
        raise ExportError('未找到明文 contact.db。请选择包含 contact\\contact.db 的解密输出目录。')
    if not _is_plain_sqlite(contact_db):
        raise ExportError('contact.db 不是明文 SQLite（文件头不是 SQLite format 3）。本工具不解密数据库，请先自行准备好明文副本。')
    head_db = _find(opts.db_dir, 'head_image', 'head_image.db')
    bizmsg_db = _find(opts.db_dir, 'message', 'biz_message_0.db')
    log('明文库: ' + contact_db)
    log('头像库: ' + (head_db if head_db else '未找到，跳过本地头像'))
    log('公众号消息库: ' + (bizmsg_db if bizmsg_db else '未找到，跳过最早本地消息日期'))

    prog('读取本地头像缓存', 0, 1)
    avatars = {}
    if head_db and _is_plain_sqlite(head_db):
        try:
            hc = ro(head_db)
            for u, buf in hc.execute('SELECT username, image_buffer FROM head_image'):
                if isinstance(buf, bytes) and len(buf) > 64:
                    avatars[clean(u)] = buf
            hc.close()
        except Exception as e:
            log('头像库读取失败（忽略）: ' + str(e)[:100])
    log('本地头像缓存: %d 条' % len(avatars))
    local_gh = set(avatars)

    prog('读取公众号记录', 0, 1)
    con = ro(contact_db)
    con.row_factory = sqlite3.Row
    try:
        rows = con.execute('''
            SELECT b.username AS gh, b.type AS btype, b.brand_icon_url, b.brand_info,
                   b.external_info, c.nick_name, c.alias, c.big_head_url, c.small_head_url,
                   c.quan_pin, c.pin_yin_initial
            FROM biz_info b LEFT JOIN contact c ON b.username = c.username
            ORDER BY b.id
        ''').fetchall()
    except sqlite3.Error as e:
        raise ExportError('读取 biz_info / contact 表失败: %s' % e)
    finally:
        con.close()
    log('公众号记录: %d 条' % len(rows))

    earliest = {}
    if bizmsg_db and _is_plain_sqlite(bizmsg_db):
        try:
            bm = ro(bizmsg_db)
            names = [clean(r[0]) for r in bm.execute(
                "SELECT user_name FROM Name2Id WHERE user_name LIKE 'gh_%'")]
            for u in names:
                t = 'Msg_' + hashlib.md5(u.encode()).hexdigest()
                try:
                    v = bm.execute('SELECT MIN(create_time) FROM "%s"' % t).fetchone()[0]
                    if v and 1000000000 < int(v) < 2000000000:
                        earliest[u] = datetime.datetime.fromtimestamp(int(v)).strftime('%Y-%m-%d')
                except Exception:
                    pass
            bm.close()
        except Exception as e:
            log('最早消息时间读取失败（忽略）: ' + str(e)[:100])
    log('有最早本地消息日期: %d 条' % len(earliest))

    records = []
    for r in rows:
        gh = clean(r['gh'])
        ext = clean(r['external_info'])
        rs = {}
        try:
            rs = json.loads(ext).get('RegisterSource') or {}
        except Exception:
            pass
        about = clean(rs.get('AboutBizUrl')) if isinstance(rs, dict) else ''
        biz = ''
        if r['brand_info']:
            m = re.search(r'__biz=([A-Za-z0-9=_-]+)', clean(r['brand_info']))
            if m:
                biz = m.group(1)
        profile = ('https://mp.weixin.qq.com/mp/profile_ext?action=home&__biz=' + biz) if biz else ''
        avatar_url = (clean(r['brand_icon_url']) or clean(r['big_head_url'])
                      or clean(r['small_head_url']))
        records.append({
            'name': clean(r['nick_name']),
            'gh': gh,
            'alias': clean(r['alias']),
            'pinyin': (clean(r['quan_pin']) or clean(r['pin_yin_initial'])
                       or clean(r['nick_name'])).lower(),
            'desc': build_desc(ext),
            'date': earliest.get(gh, ''),
            'addr': about or profile,
            'profile': profile,
            'avatar_url': avatar_url,
            'btype': r['btype'],
        })
    return records, avatars, local_gh, earliest


def _fill_avatars(opts, records, avatars, log, prog):
    from urllib.parse import urlparse
    cache = os.path.join(opts.out_dir, 'avatar_cache')
    os.makedirs(cache, exist_ok=True)
    hosts = {h.lower() for h in opts.allowed_hosts}

    reused = 0
    for x in records:
        if x['gh'] in avatars:
            continue
        for ext in ('png', 'jpg'):
            p = os.path.join(cache, x['gh'] + '.' + ext)
            if os.path.exists(p):
                avatars[x['gh']] = open(p, 'rb').read()
                reused += 1
                break
    if reused:
        log('复用本地头像缓存: %d 张' % reused)

    missing, skipped = [], 0
    for i, x in enumerate(records):
        if x['gh'] in avatars or not x['avatar_url'].startswith('http'):
            continue
        host = (urlparse(x['avatar_url']).hostname or '').lower()
        if host not in hosts:
            skipped += 1
            continue
        missing.append((i, x['avatar_url']))
    if not opts.fetch_avatars:
        log('联网补齐未开启；可补齐 %d 张（仅 %s），另有 %d 条 URL 不在允许域名内' %
            (len(missing), ', '.join(sorted(hosts)), skipped))
        return skipped

    def fetch(item):
        i, url = item
        urls = ['https://' + url[7:]] if url.startswith('http://') else []
        urls.append(url)
        for u in urls:
            try:
                req = urllib.request.Request(u, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=12) as resp:
                    d = resp.read()
                if 64 < len(d) < 300000 and _is_image(d):
                    return i, d
            except Exception:
                continue
        return i, None

    good = 0
    total = len(missing)
    log('开始联网补齐头像: %d 张' % total)
    with ThreadPoolExecutor(max_workers=10) as pool:
        for done, (i, data) in enumerate(pool.map(fetch, missing), 1):
            if data:
                gh = records[i]['gh']
                avatars[gh] = data
                ext = 'png' if data[:8] == b'\x89PNG\r\n\x1a\n' else 'jpg'
                with open(os.path.join(cache, gh + '.' + ext), 'wb') as f:
                    f.write(data)
                good += 1
            if done % 20 == 0 or done == total:
                prog('下载头像', done, total)
    log('联网补齐成功 %d 张，失败 %d 张' % (good, total - good))
    return skipped


# ------------------------------------------------------------------ OOXML helpers
def rpr(size=13, bold=False, color=None, underline=False):
    out = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
           % (FONT, FONT, FONT, FONT))
    if bold:
        out += '<w:b/>'
    if color:
        out += '<w:color w:val="%s"/>' % color
    if underline:
        out += '<w:u w:val="single"/>'
    out += '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (size, size)
    return out


def para(text, size=18, bold=False, color=None, after=80, before=0):
    ppr = '<w:pPr><w:spacing w:before="%d" w:after="%d" w:line="260" w:lineRule="auto"/></w:pPr>' % (before, after)
    return '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (
        ppr, rpr(size, bold, color), escape(text))


def tcpr(width, shade=None):
    out = '<w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:vAlign w:val="center"/>' % width
    if shade:
        out += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade
    out += ('<w:tcMar><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/>'
            '<w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tcMar></w:tcPr>')
    return out


def text_cell(text, width, size=13, bold=False, align=None, shade=None, color=None):
    ppr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="220" w:lineRule="auto"/>'
    if align:
        ppr += '<w:jc w:val="%s"/>' % align
    ppr += '</w:pPr>'
    body = '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (
        ppr, rpr(size, bold, color), escape(text))
    return '<w:tc>%s%s</w:tc>' % (tcpr(width, shade), body)


def write_docx(path, records, avatars, n_local, skipped, hosts):
    rels, media, docpr_id = [], [], [1]

    def add_rel(url):
        rid = 'rIdL%d' % (len(rels) + 1)
        rels.append((rid, url))
        return rid

    def add_image(data):
        ext = 'png' if data[:8] == b'\x89PNG\r\n\x1a\n' else 'jpg'
        idx = len(media) + 1
        name = 'av%04d.%s' % (idx, ext)
        media.append((name, data))
        return 'rIdP%d' % idx

    def image_cell(width, data):
        if not data:
            return text_cell('—', width, size=12, align='center', color='808080')
        rid = add_image(data)
        did = docpr_id[0]
        docpr_id[0] += 1
        drawing = (
            '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
            '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="avatar%d"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="a%d"/>'
            '<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
            % (EMU, EMU, did, did, did, rid, EMU, EMU))
        return '<w:tc>%s%s</w:tc>' % (tcpr(width), drawing)

    def url_cell(width, url):
        if not url:
            return text_cell('—', width, size=12, align='center', color='808080')
        rid = add_rel(url)
        ppr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="auto"/></w:pPr>'
        body = ('<w:p>%s<w:hyperlink r:id="%s" w:history="1"><w:r>%s<w:t xml:space="preserve">%s</w:t>'
                '</w:r></w:hyperlink></w:p>' % (ppr, rid, rpr(11, color='1F5C99', underline=True), escape(url)))
        return '<w:tc>%s%s</w:tc>' % (tcpr(width), body)

    n = len(records)
    n_pic = sum(1 for x in records if x['gh'] in avatars)
    n_desc = sum(1 for x in records if x['desc'])
    n_date = sum(1 for x in records if x['date'])
    n_addr = sum(1 for x in records if x['addr'])
    n_net = n_pic - n_local

    parts = [para('微信公众号关注清单（横版 · 含头像）', size=30, bold=True,
                  color='1F3864', after=120)]
    parts.append(para(
        '本表由用户自行解密得到的微信本地明文数据库导出，共 %d 条公众号记录。'
        '其中 %d 条嵌入了头像图片（%d 张取自本机 head_image.db 缓存，%d 张按头像 URL 补齐）、'
        '%d 条有描述、%d 条有最早本地消息日期、%d 条有地址链接。'
        '页面为 A4 横向，顺序按名称拼音排序。'
        % (n, n_pic, n_local, n_net, n_desc, n_date, n_addr),
        size=17, after=90))
    parts.append(para(
        '字段说明：「描述」列不是公众号官方简介，本地库没有简介字段；它由 external_info 中的认证说明、'
        '认证主体、注册主体、商标等可核实字段重述而成，没有内容的留空。'
        '「最早本地消息」是公众号消息库中该号最早一条本地留存消息的日期，'
        '它不等于关注时间，也不是公众号创办时间；本地库没有这两个字段，未编造这两列。'
        '「地址」列显示完整 URL 明文并可点击，来源是「关于本公众号」页或 mp/profile_ext 主页；'
        '微信服务市场注明 profile_ext 形式仅支持在微信内或小程序 web-view 打开，普通浏览器不能直接访问。'
        '另有 %d 条因本地无头像 URL、下载失败或域名不在允许范围（%s），头像留「—」，'
        '完整 URL 仍保留在 CSV / Excel。完整字段另见随附 CSV。' % (n - n_pic, ', '.join(hosts)),
        size=17, after=110))

    borders = ('<w:tblBorders>' + ''.join(
        '<w:%s w:val="single" w:sz="4" w:color="C9CDD4"/>' % e
        for e in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) + '</w:tblBorders>')
    tbl = ['<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/>%s</w:tblPr>' % (W, borders)]
    tbl.append('<w:tblGrid>' + ''.join('<w:gridCol w:w="%d"/>' % c for c in COLS) + '</w:tblGrid>')
    headers = ['序号', '头像', '公众号名', '原始 ID', '微信号', '描述', '最早本地消息', '地址（完整 URL）']
    hdr = ''.join(text_cell(h, COLS[i], size=14, bold=True, align='center', shade='EEF2F7')
                  for i, h in enumerate(headers))
    tbl.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>%s</w:tr>' % hdr)
    for i, x in enumerate(records, 1):
        cells = [
            text_cell(str(i), COLS[0], size=12, align='center'),
            image_cell(COLS[1], avatars.get(x['gh'])),
            text_cell(x['name'] or '（本地库无名称）', COLS[2], size=14, bold=True),
            text_cell(x['gh'], COLS[3], size=12),
            text_cell(x['alias'] or '—', COLS[4], size=12, color=None if x['alias'] else '808080'),
            text_cell(x['desc'] or '—', COLS[5], size=12, color=None if x['desc'] else '808080'),
            text_cell(x['date'] or '—', COLS[6], size=12, align='center', color=None if x['date'] else '808080'),
            url_cell(COLS[7], x['addr']),
        ]
        tbl.append('<w:tr>%s</w:tr>' % ''.join(cells))
    tbl.append('</w:tbl>')
    parts.append(''.join(tbl))
    parts.append(para('所有字段均来自本机明文库；空值写「—」，未补写任何未查证内容。', size=16, after=0, before=100))

    sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
            '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
            'w:header="480" w:footer="480" w:gutter="0"/></w:sectPr>')
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<w:body>' + ''.join(parts) + sect + '</w:body></w:document>')
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="jpg" ContentType="image/jpeg"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>')
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="word/styles.xml"/>'
        '</Relationships>')
    rel_items = [
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>'
        % (rid, escape(url)) for rid, url in rels]
    for name, _ in media:
        rid = 'rIdP%d' % int(re.search(r'av(\d+)', name).group(1))
        rel_items.append(
            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>'
            % (rid, name))
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + ''.join(rel_items) + '</Relationships>')
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
        '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="100"/></w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>'
        '<w:rPr><w:color w:val="1F5C99"/><w:u w:val="single"/></w:rPr></w:style>'
        '</w:styles>' % (FONT, FONT, FONT, FONT))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types.encode('utf-8'))
        z.writestr('_rels/.rels', root_rels.encode('utf-8'))
        z.writestr('word/document.xml', document.encode('utf-8'))
        z.writestr('word/_rels/document.xml.rels', doc_rels.encode('utf-8'))
        z.writestr('word/styles.xml', styles.encode('utf-8'))
        for name, data in media:
            z.writestr('word/media/' + name, data)
    return len(media), len(rels)


def write_csv(path, records):
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['序号', '公众号名', '原始ID(gh_)', '自定义微信号',
                    '描述(本地库认证/注册信息)', '最早本地消息日期', '地址(关于本号或主页)',
                    '头像URL', '微信内主页链接', 'biz_info_type'])
        for i, x in enumerate(records, 1):
            w.writerow([i, x['name'], x['gh'], x['alias'], x['desc'], x['date'],
                        x['addr'], x['avatar_url'], x['profile'], x['btype']])


def write_xlsx(path, records):
    def col_letter(i):
        s = ''
        i += 1
        while i:
            i, r = divmod(i - 1, 26)
            s = chr(65 + r) + s
        return s

    heads = ['序号', '公众号名', '原始ID(gh_)', '自定义微信号',
             '描述(本地库认证/注册信息)', '最早本地消息日期',
             '地址(关于本号或主页)', '头像URL', '微信内主页链接']
    widths = [6, 24, 22, 18, 52, 17, 62, 62, 62]
    rows = ['<row r="1">' + ''.join(
        '<c r="%s1" s="1" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
        % (col_letter(j), escape(h)) for j, h in enumerate(heads)) + '</row>']
    for i, x in enumerate(records, start=2):
        vals = [str(i - 1), x['name'], x['gh'], x['alias'], x['desc'], x['date'],
                x['addr'], x['avatar_url'], x['profile']]
        rows.append('<row r="%d">%s</row>' % (
            i, ''.join('<c r="%s%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                       % (col_letter(j), i, escape(v)) for j, v in enumerate(vals))))
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<cols>' + ''.join('<col min="%d" max="%d" width="%d" customWidth="1"/>'
                           % (j + 1, j + 1, w) for j, w in enumerate(widths)) + '</cols>'
        '<sheetData>' + ''.join(rows) + '</sheetData>'
        '<autoFilter ref="A1:I%d"/></worksheet>' % (len(records) + 1))
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Microsoft YaHei"/></font>'
        '<font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Microsoft YaHei"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFEEF2F7"/>'
        '<bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
        '</cellStyles></styleSheet>')
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>')
    root = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')
    wb = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="公众号清单" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wbr = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct.encode('utf-8'))
        z.writestr('_rels/.rels', root.encode('utf-8'))
        z.writestr('xl/workbook.xml', wb.encode('utf-8'))
        z.writestr('xl/_rels/workbook.xml.rels', wbr.encode('utf-8'))
        z.writestr('xl/styles.xml', styles.encode('utf-8'))
        z.writestr('xl/worksheets/sheet1.xml', sheet.encode('utf-8'))


HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="referrer" content="no-referrer">
<title>微信公众号清单</title>
<style>
*{box-sizing:border-box}body{margin:0;font-family:"Microsoft YaHei",system-ui,sans-serif;background:#f5f7fa;color:#1f2937}
.wrap{max-width:1500px;margin:auto;padding:24px}
.card{background:#fff;border:1px solid #e3e8ef;border-radius:12px;padding:28px;box-shadow:0 2px 12px rgba(16,24,40,.05)}
h1{font-size:22px;margin:0 0 8px;color:#1f3864}p.tip{color:#64748b;font-size:13px;line-height:1.7;margin:0 0 18px}
input,button{font:inherit;border-radius:8px;border:1px solid #cbd5e1;padding:10px 14px}
button{background:#1f3864;color:#fff;border-color:#1f3864;cursor:pointer}
button:hover{background:#2b4c86}.err{color:#b91c1c;font-size:13px;margin-top:10px;min-height:18px}
#gate{max-width:420px;margin:12vh auto;text-align:center}
#main{display:none}.bar{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:14px}
#q{flex:1;min-width:240px}.count{color:#64748b;font-size:13px}
table{width:100%;border-collapse:collapse;background:#fff;font-size:13px}
th,td{border:1px solid #e3e8ef;padding:7px 9px;text-align:left;vertical-align:middle;word-break:break-all}
th{background:#eef2f7;color:#1f3864;position:sticky;top:0;z-index:2;font-weight:600}
tr:nth-child(even) td{background:#fbfcfe}
img{width:38px;height:38px;border-radius:6px;object-fit:cover;background:#eef2f7}
a{color:#1f5c99;text-decoration:none}a:hover{text-decoration:underline}
.ph{display:inline-block;width:38px;height:38px;border-radius:6px;background:#eef2f7;text-align:center;line-height:38px;color:#94a3b8;font-size:11px}
</style></head><body>
<div id="gate" class="card"><h1>微信公众号清单</h1>
<p class="tip">本页内容已加密，请输入访问口令。</p>
<input id="pw" type="password" placeholder="访问口令" onkeydown="if(event.key==='Enter')unlock()">
<button onclick="unlock()">解锁</button><div class="err" id="err"></div></div>
<div id="main" class="wrap"><div class="card">
<h1>微信公众号关注清单</h1>
<p class="tip">按名称拼音排序 · 共 <b id="total">0</b> 条 · 数据来自本机明文库。<br>
「描述」为本地库认证/注册信息重述，非官方简介；「最早本地消息」不等于关注时间；地址链接多为微信内页面，普通浏览器可能无法打开。</p>
<div class="bar"><input id="q" placeholder="搜索名称 / 原始ID / 微信号" oninput="render()">
<span class="count" id="shown"></span></div>
<div style="overflow:auto;max-height:78vh"><table><thead><tr>
<th>序号</th><th>头像</th><th>公众号名</th><th>原始ID</th><th>微信号</th><th>描述</th><th>最早本地消息</th><th>地址</th>
</tr></thead><tbody id="tb"></tbody></table></div></div></div>
<script>
const BLOB="__BLOB__"; let DATA=[];
function b64(u){return Uint8Array.from(atob(u),c=>c.charCodeAt(0))}
async function decrypt(pw){
 const raw=b64(BLOB), salt=raw.slice(0,16), nonce=raw.slice(16,28);
 const km=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']);
 const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt:salt,iterations:100000,hash:'SHA-256'},km,{name:'AES-GCM',length:256},false,['decrypt']);
 const tag=raw.slice(28,44), ct=raw.slice(44), buf=new Uint8Array(ct.length+16);
 buf.set(ct);buf.set(tag,ct.length);
 return await crypto.subtle.decrypt({name:'AES-GCM',iv:nonce,tagLength:128},key,buf);
}
async function unlock(){
 const pw=document.getElementById('pw').value, err=document.getElementById('err');err.textContent='';
 if(!pw){err.textContent='请输入口令';return}
 try{
   const plain=await decrypt(pw); DATA=JSON.parse(new TextDecoder().decode(plain));
   document.getElementById('gate').style.display='none';
   document.getElementById('main').style.display='block';
   document.getElementById('total').textContent=DATA.length; render();
 }catch(e){err.textContent='口令不正确，或当前环境不支持 Web Crypto（请用 https 链接打开）'}
}
function esc(s){return String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function render(){
 const q=(document.getElementById('q').value||'').toLowerCase();
 const list=DATA.filter(x=>!q||(x.n+' '+x.g+' '+x.a).toLowerCase().includes(q));
 document.getElementById('shown').textContent='显示 '+list.length+' 条';
 document.getElementById('tb').innerHTML=list.map(x=>'<tr>'+
  '<td>'+x.i+'</td>'+
  '<td>'+(x.p?'<img loading="lazy" src="'+esc(x.p)+'" onerror="this.outerHTML=\'<span class=ph>无图</span>\'">':'<span class="ph">无图</span>')+'</td>'+
  '<td><b>'+esc(x.n||'（本地库无名称）')+'</b></td><td>'+esc(x.g)+'</td><td>'+esc(x.a||'—')+'</td>'+
  '<td>'+esc(x.d||'—')+'</td><td>'+esc(x.t||'—')+'</td>'+
  '<td>'+(x.u?'<a href="'+esc(x.u)+'" target="_blank" rel="noopener">打开链接</a>':'—')+'</td></tr>').join('');
}
</script></body></html>'''


def write_web(path, records, password):
    if AES is None:
        raise ExportError('生成加密网页需要 pycryptodome，请先执行 pip install -r requirements.txt')

    def web_img(u):
        # published pages are usually https; http images would be blocked as mixed content
        return 'https://' + u[7:] if u.startswith('http://') else u

    payload = [{'i': i, 'n': x['name'], 'g': x['gh'], 'a': x['alias'],
                'd': x['desc'], 't': x['date'], 'u': x['addr'],
                'p': web_img(x['avatar_url'])}
               for i, x in enumerate(records, 1)]
    plain = json.dumps(payload, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    salt, nonce = os.urandom(16), os.urandom(12)
    key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 100000, 32)
    cip = AES.new(key, AES.MODE_GCM, nonce=nonce)
    ct, tag = cip.encrypt_and_digest(plain)
    blob = base64.b64encode(salt + nonce + tag + ct).decode('ascii')
    html = HTML_TEMPLATE.replace('__BLOB__', blob)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)
    return len(html)


def export_all(opts: ExportOptions, log=print, progress=None) -> dict:
    """Run the whole export. Returns a summary dict. Raises ExportError."""
    if not opts.db_dir or not os.path.isdir(opts.db_dir):
        raise ExportError('明文库目录不存在：%s' % opts.db_dir)
    if opts.make_web and not opts.password:
        raise ExportError('生成加密网页需要设置访问口令。')
    if opts.make_web and AES is None:
        raise ExportError('未安装 pycryptodome，无法生成加密网页。请执行 pip install -r requirements.txt，或取消勾选「生成加密网页」。')
    os.makedirs(opts.out_dir, exist_ok=True)
    if not (opts.make_docx or opts.make_xlsx or opts.make_csv or opts.make_web):
        raise ExportError('请至少选择一种输出格式。')

    def prog(stage, done=0, total=0):
        if progress:
            progress(stage, done, total)

    log('开始导出')
    records, avatars, local_gh, _ = _collect(opts, log, prog)
    skipped = _fill_avatars(opts, records, avatars, log, prog)

    prog('排序与整理', 0, 1)
    records.sort(key=lambda x: (x['pinyin'] == '', x['pinyin'], x['name']))
    n = len(records)
    result = {'records': n, 'avatars': sum(1 for x in records if x['gh'] in avatars),
              'desc': sum(1 for x in records if x['desc']),
              'date': sum(1 for x in records if x['date']),
              'addr': sum(1 for x in records if x['addr']), 'files': {}}

    if opts.make_csv:
        prog('生成 CSV', 0, 1)
        p = os.path.join(opts.out_dir, '微信公众号清单.csv')
        write_csv(p, records)
        result['files']['csv'] = p
        log('CSV 已生成: %s' % p)
    if opts.make_docx:
        prog('生成 Word（嵌入头像）', 0, 1)
        p = os.path.join(opts.out_dir, '微信公众号清单.docx')
        media, links = write_docx(p, records, avatars,
                                  sum(1 for x in records if x['gh'] in local_gh),
                                  skipped, opts.allowed_hosts)
        result['files']['docx'] = p
        result['embedded_images'] = media
        result['hyperlinks'] = links
        log('Word 已生成: %s（%d 张图片，%d 个链接）' % (p, media, links))
    if opts.make_xlsx:
        prog('生成 Excel', 0, 1)
        p = os.path.join(opts.out_dir, '微信公众号清单.xlsx')
        write_xlsx(p, records)
        result['files']['xlsx'] = p
        log('Excel 已生成: %s' % p)
    if opts.make_web:
        prog('生成加密网页', 0, 1)
        web_dir = os.path.join(opts.out_dir, 'web')
        os.makedirs(web_dir, exist_ok=True)
        p = os.path.join(web_dir, 'index.html')
        size = write_web(p, records, opts.password)
        result['files']['web'] = p
        log('加密网页已生成: %s（AES-256-GCM，%d 字符）' % (p, size))
    prog('完成', 1, 1)
    log('导出完成，共 %d 条记录。' % n)
    return result


def main():
    import argparse
    import sys
    ap = argparse.ArgumentParser(description='从已解密的微信本地明文库导出公众号清单（不解密、不联网上传）')
    ap.add_argument('--db-dir', required=True, help='包含 contact/contact.db 的明文数据库目录')
    ap.add_argument('--out-dir', required=True, help='输出目录')
    ap.add_argument('--password', default='', help='加密网页的访问口令（启用 --web 时必填）')
    ap.add_argument('--fetch-avatars', action='store_true', help='联网补齐缺失头像，仅访问 mmbiz.qpic.cn')
    ap.add_argument('--no-docx', action='store_true')
    ap.add_argument('--no-xlsx', action='store_true')
    ap.add_argument('--no-csv', action='store_true')
    ap.add_argument('--no-web', action='store_true')
    a = ap.parse_args()
    opts = ExportOptions(
        db_dir=a.db_dir, out_dir=a.out_dir, password=a.password,
        make_docx=not a.no_docx, make_xlsx=not a.no_xlsx,
        make_csv=not a.no_csv, make_web=not a.no_web,
        fetch_avatars=a.fetch_avatars)
    try:
        r = export_all(opts)
        print(json.dumps(r, ensure_ascii=False, indent=2))
    except ExportError as e:
        print('错误: %s' % e, file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
