from pathlib import Path
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from converters import convert


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('轻转 · 本地文件转换')
        self.geometry('900x620')
        self.minsize(760, 520)
        self.events = queue.Queue()
        self.busy = False
        self.paths = {}
        self.mode = tk.StringVar(value='图片互转')
        self.fmt = tk.StringVar(value='PNG')
        self.quality = tk.IntVar(value=90)
        self.output = tk.StringVar(value=str(Path(__file__).parent / '转换结果'))
        self.status = tk.StringVar(value='请先添加文件，再点击右下角“开始转换”。')
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('.', font=('Microsoft YaHei UI', 10))
        style.configure('Title.TLabel', font=('Microsoft YaHei UI', 23, 'bold'))
        style.configure('TButton', padding=(14, 8))
        style.configure('Start.TButton', font=('Microsoft YaHei UI', 12, 'bold'),
                        foreground='white', background='#2563eb', padding=(22, 12))
        style.map('Start.TButton', background=[('disabled', '#94a3b8'), ('active', '#1d4ed8')],
                  foreground=[('disabled', '#f1f5f9')])
        body = ttk.Frame(self, padding=24)
        body.pack(fill='both', expand=True)
        # Reserve the controls first so the file list cannot push them off-screen.
        footer = ttk.Frame(body)
        footer.pack(side='bottom', fill='x')
        ttk.Label(body, text='轻转', style='Title.TLabel').pack(anchor='w')
        ttk.Label(body, text='① 选择转换类型    →    ② 添加文件    →    ③ 点击右下角“开始转换”', foreground='#334155').pack(anchor='w', pady=(4, 20))
        options = ttk.Frame(body)
        options.pack(fill='x')
        self.modebox = ttk.Combobox(options, textvariable=self.mode, values=['图片互转', 'PDF → Word', 'Word → PDF'], state='readonly', width=18)
        self.modebox.pack(side='left')
        self.modebox.bind('<<ComboboxSelected>>', self.change_mode)
        ttk.Label(options, text='    图片格式 ').pack(side='left')
        self.formatbox = ttk.Combobox(options, textvariable=self.fmt, values=['PNG', 'JPEG', 'WEBP', 'BMP', 'TIFF'], state='readonly', width=8)
        self.formatbox.pack(side='left')
        ttk.Label(options, text='    质量 ').pack(side='left')
        self.qualitybox = ttk.Spinbox(options, from_=1, to=100, textvariable=self.quality, width=5)
        self.qualitybox.pack(side='left')
        self.hint = ttk.Label(body, text='', foreground='#617080', wraplength=800)
        self.hint.pack(anchor='w', pady=14)
        toolbar = ttk.Frame(body)
        toolbar.pack(fill='x', pady=(0, 8))
        self.add = ttk.Button(toolbar, text='＋ 添加文件', command=self.add_files)
        self.add.pack(side='left')
        self.clear = ttk.Button(toolbar, text='清空列表', command=self.clear_files)
        self.clear.pack(side='left', padx=8)
        table = ttk.Frame(body)
        table.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(table, columns=('file', 'state'), show='headings', selectmode='extended')
        self.tree.heading('file', text='文件')
        self.tree.heading('state', text='状态 / 结果（双击查看）')
        self.tree.column('file', width=300)
        self.tree.column('state', width=460)
        self.tree.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        scroll.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind('<Double-1>', self.details)
        dest = ttk.Frame(footer)
        dest.pack(fill='x', pady=16)
        ttk.Label(dest, text='保存到 ').pack(side='left')
        self.entry = ttk.Entry(dest, textvariable=self.output)
        self.entry.pack(side='left', fill='x', expand=True, padx=8)
        self.browse = ttk.Button(dest, text='选择目录', command=self.choose_dir)
        self.browse.pack(side='left')
        self.progress = ttk.Progressbar(footer)
        self.progress.pack(fill='x')
        bottom = ttk.Frame(footer)
        bottom.pack(fill='x', pady=(12, 0))
        self.start = ttk.Button(bottom, text='开始转换', style='Start.TButton', command=self.run)
        self.start.pack(side='right')
        ttk.Button(bottom, text='打开结果文件夹', command=self.open_output).pack(side='right', padx=8)
        ttk.Label(footer, textvariable=self.status, wraplength=680).pack(anchor='w', pady=(8, 0))
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.change_mode()
        self.after(100, self.poll)

    def change_mode(self, event=None):
        self.clear_files()
        image = self.mode.get() == '图片互转'
        self.formatbox.configure(state='readonly' if image else 'disabled')
        self.qualitybox.configure(state='normal' if image else 'disabled')
        self.hint.configure(text={'图片互转': '支持 JPG / PNG / WebP / BMP / TIFF；转 JPG、BMP 时透明背景填白。质量仅影响 JPG、WebP，暂不支持动画或多页图片。', 'PDF → Word': '输出可编辑 DOCX，适用于有文字层的 PDF。复杂表格、公式和多栏排版需检查；扫描页不会自动识别。', 'Word → PDF': '通过本机 WPS 导出 PDF，支持 DOC / DOCX。转换期间请勿操作正在转换的文档。'}[self.mode.get()])

    def add_files(self):
        patterns = {'图片互转': '*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff', 'PDF → Word': '*.pdf', 'Word → PDF': '*.doc *.docx'}
        for name in filedialog.askopenfilenames(filetypes=[('支持的文件', patterns[self.mode.get()])]):
            if name not in self.paths.values():
                row = self.tree.insert('', 'end', values=(Path(name).name, '等待转换'))
                self.paths[row] = name
        if self.paths:
            self.status.set(f'已添加 {len(self.paths)} 个文件，请点击右下角蓝色“开始转换”按钮。')

    def clear_files(self):
        self.tree.delete(*self.tree.get_children())
        self.paths.clear()
        self.status.set('请先添加文件，再点击右下角“开始转换”。')

    def choose_dir(self):
        value = filedialog.askdirectory()
        if value:
            self.output.set(value)

    def open_output(self):
        path = Path(self.output.get())
        if path.is_dir():
            os.startfile(path)
        else:
            messagebox.showinfo('结果目录', '目录尚未创建，请先完成一次转换。')

    def details(self, event=None):
        selected = self.tree.selection()
        if selected:
            messagebox.showinfo('转换详情', self.tree.item(selected[0], 'values')[1])

    def set_busy(self, value):
        self.busy = value
        self.start.configure(text='正在转换…' if value else '开始转换')
        for widget in [self.add, self.clear, self.entry, self.browse, self.start]:
            widget.configure(state='disabled' if value else 'normal')
        self.modebox.configure(state='disabled' if value else 'readonly')
        self.formatbox.configure(state='readonly' if not value and self.mode.get() == '图片互转' else 'disabled')
        self.qualitybox.configure(state='normal' if not value and self.mode.get() == '图片互转' else 'disabled')

    def run(self):
        if not self.paths:
            messagebox.showinfo('添加文件', '请先添加需要转换的文件。')
            return
        try:
            quality = self.quality.get()
            if not 1 <= quality <= 100:
                raise ValueError()
            if not self.output.get().strip():
                raise ValueError()
        except (ValueError, tk.TclError):
            messagebox.showerror('参数错误', '请填写输出目录和 1 到 100 的图片质量。')
            return
        args = (list(self.paths.items()), self.output.get(), self.mode.get(), self.fmt.get(), quality)
        self.set_busy(True)
        self.progress.configure(value=0, maximum=len(self.paths))
        self.status.set('正在转换…')
        threading.Thread(target=self.worker, args=args, daemon=True).start()

    def worker(self, items, directory, mode, fmt, quality):
        successes = 0
        for row, source in items:
            self.events.put(('row', row, '正在转换…'))
            try:
                result = convert(source, directory, mode, fmt, quality)
                msg = f'完成：{result}'
                successes += 1
            except Exception as exc:
                msg = f'失败：{exc}'
            self.events.put(('result', row, msg))
        self.events.put(('done', successes, len(items)))

    def poll(self):
        while not self.events.empty():
            kind, a, b = self.events.get()
            if kind == 'done':
                self.set_busy(False)
                self.status.set(f'完成 {a} 个，失败 {b-a} 个')
            else:
                self.tree.set(a, 'state', b)
                if kind == 'result':
                    self.progress['value'] += 1
        self.after(100, self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo('正在转换', '请等待当前批次完成后再关闭。')
        else:
            self.destroy()


if __name__ == '__main__':
    App().mainloop()
