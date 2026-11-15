# -*- coding: utf-8 -*-
"""GUI entry point for wechat-account-exporter (tkinter, standard library only)."""
from __future__ import annotations

import os
import queue
import secrets
import string
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from exporter import ExportError, ExportOptions, export_all
from help_content import HELP_TEXT

APP_TITLE = '微信公众号清单导出工具  v1.0'


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.q = queue.Queue()
        self.running = False
        root.title(APP_TITLE)
        root.geometry('920x660')
        root.minsize(800, 580)

        self.db_dir = tk.StringVar()
        self.out_dir = tk.StringVar(value=os.path.join(
            os.path.dirname(os.path.abspath(sys.argv[0])), 'output'))
        self.password = tk.StringVar()
        self.v_docx = tk.BooleanVar(value=True)
        self.v_xlsx = tk.BooleanVar(value=True)
        self.v_csv = tk.BooleanVar(value=True)
        self.v_web = tk.BooleanVar(value=True)
        self.v_fetch = tk.BooleanVar(value=False)

        self._build()
        root.protocol('WM_DELETE_WINDOW', self._on_close)
        root.after(80, self._drain)

    # ---------------------------------------------------------------- UI
    def _build(self):
        f = ttk.Frame(self.root, padding=14)
        f.pack(fill='both', expand=True)
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text='① 输入与输出　（明文库目录 = 已解密的 contact\\contact.db 所在目录）',
                  font=('Microsoft YaHei', 11, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky='w', pady=(0, 8))

        self._dir_row(f, 1, '明文库目录', self.db_dir)
        self._dir_row(f, 2, '输出目录', self.out_dir)

        ttk.Label(f, text='网页访问口令').grid(row=3, column=0, sticky='w', pady=4)
        prow = ttk.Frame(f)
        prow.grid(row=3, column=1, columnspan=2, sticky='ew', pady=4)
        prow.columnconfigure(0, weight=1)
        ttk.Entry(prow, textvariable=self.password, show='').grid(row=0, column=0, sticky='ew')
        ttk.Button(prow, text='随机生成', width=10, command=self._random_pw).grid(
            row=0, column=1, padx=(8, 0))
        ttk.Label(f, text='仅「加密网页」需要；口令不会写进网页源码，忘记只能重新生成。',
                  foreground='#64748b').grid(row=4, column=1, columnspan=2, sticky='w')

        ttk.Separator(f, orient='horizontal').grid(row=5, column=0, columnspan=3,
                                                   sticky='ew', pady=12)
        ttk.Label(f, text='② 导出选项', font=('Microsoft YaHei', 11, 'bold')).grid(
            row=6, column=0, columnspan=3, sticky='w', pady=(0, 8))

        opts = ttk.Frame(f)
        opts.grid(row=7, column=0, columnspan=3, sticky='w')
        for i, (text, var) in enumerate([
                ('Word（横版 · 内嵌头像）', self.v_docx),
                ('Excel', self.v_xlsx),
                ('CSV', self.v_csv),
                ('加密网页（需口令）', self.v_web)]):
            ttk.Checkbutton(opts, text=text, variable=var).grid(row=0, column=i, padx=(0, 18))

        ttk.Checkbutton(f, text='联网补齐缺失头像（默认关闭；仅访问 mmbiz.qpic.cn）',
                        variable=self.v_fetch).grid(row=8, column=0, columnspan=3,
                                                   sticky='w', pady=(10, 0))
        ttk.Label(f, text='本工具不解密微信数据库、不读取微信进程、不上传任何数据。',
                  foreground='#64748b').grid(row=9, column=0, columnspan=3, sticky='w', pady=(4, 0))

        ttk.Separator(f, orient='horizontal').grid(row=10, column=0, columnspan=3,
                                                   sticky='ew', pady=12)
        btns = ttk.Frame(f)
        btns.grid(row=11, column=0, columnspan=3, sticky='w')
        self.btn_run = ttk.Button(btns, text='开始导出', command=self._start)
        self.btn_run.grid(row=0, column=0, ipadx=14, ipady=4)
        ttk.Button(btns, text='打开输出目录', command=self._open_out).grid(
            row=0, column=1, padx=8, ipadx=8, ipady=4)
        ttk.Button(btns, text='使用帮助', command=self._help).grid(
            row=0, column=2, ipadx=8, ipady=4)

        self.status = tk.StringVar(value='就绪')
        ttk.Label(f, textvariable=self.status).grid(row=12, column=0, columnspan=3,
                                                   sticky='w', pady=(12, 4))
        self.pb = ttk.Progressbar(f, mode='determinate', maximum=100, value=0)
        self.pb.grid(row=13, column=0, columnspan=3, sticky='ew')

        ttk.Label(f, text='运行日志').grid(row=14, column=0, columnspan=3,
                                         sticky='w', pady=(12, 4))
        self.txt = scrolledtext.ScrolledText(f, height=12, wrap='word',
                                             font=('Consolas', 9))
        self.txt.grid(row=15, column=0, columnspan=3, sticky='nsew')
        f.rowconfigure(15, weight=1)
        self._log('欢迎使用。请先选择「明文库目录」和「输出目录」，然后点击「开始导出」。')
        self._log('不清楚明文库目录是什么？点右上角「使用帮助」。')

    def _dir_row(self, parent, row, label, var):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', pady=4)
        box = ttk.Frame(parent)
        box.grid(row=row, column=1, columnspan=2, sticky='ew', pady=4)
        box.columnconfigure(0, weight=1)
        ttk.Entry(box, textvariable=var).grid(row=0, column=0, sticky='ew')
        ttk.Button(box, text='浏览…', width=9,
                   command=lambda: self._pick(var)).grid(row=0, column=1, padx=(8, 0))

    def _pick(self, var):
        d = filedialog.askdirectory(title='选择目录', initialdir=var.get() or os.getcwd())
        if d:
            var.set(os.path.normpath(d))

    def _random_pw(self):
        alphabet = string.ascii_letters + string.digits
        self.password.set(''.join(secrets.choice(alphabet) for _ in range(10)))

    def _open_out(self):
        d = self.out_dir.get().strip()
        if d and os.path.isdir(d):
            try:
                os.startfile(d)  # noqa: Windows only; guarded in main()
            except AttributeError:
                messagebox.showinfo('输出目录', d)
        else:
            messagebox.showinfo('输出目录', '目录还不存在，导出成功后会自动创建。\n' + (d or '(未填写)'))

    def _help(self):
        w = tk.Toplevel(self.root)
        w.title('使用帮助')
        w.geometry('780x600')
        w.transient(self.root)
        t = scrolledtext.ScrolledText(w, wrap='word', font=('Microsoft YaHei', 10))
        t.pack(fill='both', expand=True, padx=12, pady=12)
        t.insert('1.0', HELP_TEXT)
        t.configure(state='disabled')
        ttk.Button(w, text='关闭', command=w.destroy).pack(pady=(0, 12))

    def _on_close(self):
        if self.running and not messagebox.askyesno('确认', '导出正在进行，确定要退出吗？'):
            return
        self.root.destroy()

    # ---------------------------------------------------------------- worker
    def _log(self, msg):
        self.txt.insert('end', msg + '\n')
        self.txt.see('end')

    def _start(self):
        if self.running:
            return
        db = self.db_dir.get().strip()
        out = self.out_dir.get().strip()
        pw = self.password.get().strip()
        if not db or not os.path.isdir(db):
            messagebox.showerror('无法开始', '请先选择一个存在的「明文库目录」。')
            return
        if not out:
            messagebox.showerror('无法开始', '请填写「输出目录」。')
            return
        if self.v_web.get() and not pw:
            messagebox.showerror('无法开始', '勾选了「加密网页」，请填写访问口令，或点「随机生成」。')
            return
        if not any([self.v_docx.get(), self.v_xlsx.get(), self.v_csv.get(), self.v_web.get()]):
            messagebox.showerror('无法开始', '请至少选择一种输出格式。')
            return

        opts = ExportOptions(db_dir=db, out_dir=out, password=pw,
                             make_docx=self.v_docx.get(), make_xlsx=self.v_xlsx.get(),
                             make_csv=self.v_csv.get(), make_web=self.v_web.get(),
                             fetch_avatars=self.v_fetch.get())
        self.running = True
        self.btn_run.configure(state='disabled', text='导出中…')
        self.txt.delete('1.0', 'end')
        self.pb['value'] = 0
        self.status.set('正在启动…')

        def work():
            try:
                r = export_all(
                    opts,
                    log=lambda m: self.q.put(('log', m)),
                    progress=lambda s, d, t: self.q.put(('prog', s, d, t)))
                self.q.put(('done', r))
            except ExportError as e:
                self.q.put(('error', str(e)))
            except Exception as e:
                self.q.put(('error', '%s: %s' % (type(e).__name__, e)))

        threading.Thread(target=work, daemon=True).start()

    def _drain(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == 'log':
                    self._log(item[1])
                elif kind == 'prog':
                    self._progress(*item[1:])
                elif kind == 'done':
                    self._done(item[1])
                elif kind == 'error':
                    self._failed(item[1])
        except queue.Empty:
            pass
        self.root.after(80, self._drain)

    def _progress(self, stage, done, total):
        self.status.set(stage + ('（%d/%d）' % (done, total) if total else ''))
        self.pb.configure(maximum=max(total, 1), value=done)

    def _done(self, r):
        self.running = False
        self.btn_run.configure(state='normal', text='开始导出')
        self.pb.configure(maximum=1, value=1)
        self.status.set('完成')
        files = '\n'.join('  %s  %s' % (k.upper(), v) for k, v in r['files'].items())
        messagebox.showinfo(
            '导出完成',
            '共 %d 条公众号记录\n内嵌头像 %d 张\n有描述 %d 条，有地址 %d 条\n\n输出文件：\n%s'
            % (r['records'], r['avatars'], r['desc'], r['addr'], files))

    def _failed(self, msg):
        self.running = False
        self.btn_run.configure(state='normal', text='开始导出')
        self.status.set('失败')
        self._log('错误：' + msg)
        messagebox.showerror('导出失败', msg)


def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
