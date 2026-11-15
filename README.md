# wechat-account-exporter

一个带图形界面的小工具：把**你自己已经解密好的**微信本地明文库，整理成一份带头像的公众号清单，一次导出 Word / Excel / CSV 和一个需要口令才能查看的网页。

![界面](https://img.shields.io/badge/UI-tkinter-blue) ![依赖](https://img.shields.io/badge/deps-pycryptodome-green) ![许可](https://img.shields.io/badge/license-MIT-lightgrey)

## 功能

- 从 `contact.db` 的 `biz_info` / `contact` 表导出公众号**名称、原始 ID（`gh_`）、自定义微信号**
- 按名称全拼排序（`quan_pin` → 拼音首字母 → 名称兜底）
- **内嵌真实头像**：优先使用本机 `head_image.db` 缓存；可选联网补齐缺失项
- 导出四种格式：
  - `微信公众号清单.docx` — A4 横向、8 列、内嵌头像缩略图、地址列可点击
  - `微信公众号清单.xlsx` — 冻结首行 + 自动筛选，含完整 URL
  - `微信公众号清单.csv` — UTF-8 BOM，保留全部原始字段
  - `web/index.html` — AES-256-GCM 加密网页，支持搜索
- 附「描述」列（本地认证 / 注册信息重述）与「最早本地消息日期」列
- Word / Excel 由标准库直接写 OOXML，**不依赖 python-docx / openpyxl**；唯一第三方依赖是 `pycryptodome`

## 边界（请先读这一段）

| 这个项目 | 不做 |
|---|---|
| 解密微信数据库 | ✗ |
| 读取进程内存 / 附加到微信客户端 | ✗ |
| 调用微信官方或内部 API | ✗ |
| 模拟登录、抓取、上传数据 | ✗ |
| 内置任何账号路径、口令、公众号数据 | ✗ |

它只读取**用户自行准备并已解密成明文 SQLite** 的本地文件，只读打开，不做任何写回。明文库如何取得由使用者自行决定，本项目不提供、不推荐、也不捆绑任何解密工具。

唯一可选的联网动作是「补齐头像」，**默认关闭**，且只访问白名单域名 `mmbiz.qpic.cn`；下载结果缓存在输出目录，不会上传。

使用者应自行确认对相关数据的处理权限，并遵守所在平台的使用规范与适用法律。

## 快速开始

### A. 可执行文件（Windows）

拿到 `WeChatAccountExporter.exe` 后双击即可，无需安装 Python。自己构建：

```bat
build_exe.bat
```

产物在 `dist\WeChatAccountExporter.exe`，脚本会同时复制一份到项目根目录。

### B. 源码运行

需要 Python 3.10+（含 tkinter）。

```bat
python -m pip install -r requirements.txt
python app.py
```

或双击 `run_gui.bat`。

### C. 命令行（无界面批处理）

```bat
python exporter.py --db-dir "D:\somewhere\decrypted" --out-dir "D:\somewhere\output" --password "your-passphrase"
```

可选参数：`--fetch-avatars`（联网补头像）、`--no-docx` / `--no-xlsx` / `--no-csv` / `--no-web`。

## 输入

```
<明文库目录>/
├── contact/contact.db           必需，明文 SQLite（文件头为 SQLite format 3）
├── head_image/head_image.db     可选，提供本地头像缓存
└── message/biz_message_0.db     可选，提供「最早本地消息日期」
```

工具也接受把这些 `.db` 直接放在根目录。缺少可选库时会跳过对应字段，并在日志里说明。

## 输出字段与已知限制

- **描述**不是公众号官方简介 —— 本地库没有简介字段。它由 `external_info` 里的认证说明、认证主体、注册主体、商标等可核实字段重述，没有内容就留空。
- **最早本地消息**不等于关注时间，也不是公众号创办时间。本地库里没有这两个字段，工具不编造这两列。
- **地址**是「关于本公众号」页或 `mp/profile_ext` 主页；微信官方注明 `profile_ext` 仅支持在微信内或小程序 web-view 打开，普通浏览器通常无法直接访问。工具照实给出完整 URL，不冒充「浏览器可访问主页」。
- **头像**优先本地缓存；未开启联网补齐、无 URL 或域名不在白名单时显示「—」，完整 URL 仍保留在 CSV / Excel。
- **条数口径**：本地 `biz_info` 表可能包含已取关缓存、系统号或企业号，与客户端显示的关注数未必一致。工具如实导出全部记录，不做筛选，CSV 保留 `biz_info_type` 供核对。
- **PDF**：不生成。用 Word / WPS 打开横版 `.docx` 后「另存为 PDF」即可。

## 加密网页是怎么做的

`PBKDF2-HMAC-SHA256`（100000 轮）从口令派生 256 位密钥，`AES-256-GCM` 加密 JSON 数据；浏览器用 Web Crypto 在本地解密。

- 口令**不在**网页源码里，源码里也没有明文清单
- 解密需要安全上下文，请通过 **https** 链接访问；`file://` 下部分浏览器会禁用 Web Crypto
- 忘记口令无法找回，重新运行一次生成新页面即可
- 头像 URL 会升级为 https 并设置 `no-referrer`，避免 HTTPS 页面的混合内容拦截

## 项目结构

```
app.py              GUI 入口（tkinter）
exporter.py         核心导出逻辑、OOXML 生成、加密网页；也可独立作 CLI 运行
help_content.py      GUI 内置帮助文本
使用说明.md          离线帮助文件
build_exe.bat        PyInstaller 一键构建
run_gui.bat          源码方式启动
requirements.txt    依赖
```

## 常见问题

**「不是明文 SQLite」** — 选中的是微信原始加密库。本工具不解密，请先自行准备明文副本。

**头像少了** — 勾选「联网补齐缺失头像」，或检查该号是否本来就没有 URL。

**网页提示口令不正确** — 确认用 https 打开，且输入的是本次生成时设置的口令。

**Excel 打开报「文件已损坏」** — 请确认未被其他程序占用；工具会全部重写该文件。

## 许可

MIT。详见 [LICENSE](LICENSE)。
