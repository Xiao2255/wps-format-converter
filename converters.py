from pathlib import Path
import sys
import tempfile
import os

sys.path.insert(0, str(Path(__file__).parent / '.vendor'))


def image_convert(source, target, fmt, quality=90):
    from PIL import Image, ImageOps
    with Image.open(source) as original:
        if getattr(original, 'n_frames', 1) > 1:
            raise ValueError('暂不支持动画或多页图片，请先拆分为单张图片。')
        img = ImageOps.exif_transpose(original)
        if fmt in ('JPEG', 'BMP'):
            rgba = img.convert('RGBA')
            img = Image.new('RGB', rgba.size, 'white')
            img.paste(rgba, mask=rgba.getchannel('A'))
        elif img.mode not in ('RGB', 'RGBA', 'L', 'LA'):
            img = img.convert('RGBA')
        img.save(target, format=fmt, quality=quality)


def pdf_to_word(source, target):
    import pymupdf
    from pdf2docx import Converter
    # WPS PDFs may report unusually tall CJK font boxes; use glyph heights
    # so adjacent text lines are not discarded as overlapping duplicates.
    pymupdf.TOOLS.set_small_glyph_heights(True)
    with pymupdf.open(source) as pdf:
        if pdf.needs_pass:
            raise ValueError('PDF 有密码保护，请先解除密码后重试。')
        if not any(page.get_text().strip() for page in pdf):
            raise ValueError('未检测到可提取文字，此版本暂不支持扫描件 OCR。')
    cv = Converter(str(source))
    try:
        cv.convert(str(target), multi_processing=False)
    finally:
        cv.close()


def word_to_pdf(source, target):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = doc = None
    previous_security = None
    try:
        try:
            app = win32com.client.DispatchEx('KWPS.Application')
        except Exception as exc:
            raise RuntimeError('无法连接 WPS，请确认已安装且可以正常打开文字文档。') from exc
        # Do not quit WPS: some editions share the user's existing application.
        try:
            previous_security = app.AutomationSecurity
            app.AutomationSecurity = 3
        except Exception as exc:
            raise RuntimeError('此 WPS 版本不支持禁用文档宏，无法安全自动转换。') from exc
        doc = app.Documents.Open(str(Path(source).resolve()), False, True, False)
        doc.ExportAsFixedFormat(str(Path(target).resolve()), 17)
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        finally:
            try:
                if app is not None and previous_security is not None:
                    app.AutomationSecurity = previous_security
            finally:
                doc = app = None
                pythoncom.CoUninitialize()


def convert(source, directory, mode, fmt='PNG', quality=90):
    source, directory = Path(source), Path(directory)
    if not source.is_file():
        raise ValueError('源文件不存在。')
    suffix = {'JPEG': '.jpg', 'PNG': '.png', 'WEBP': '.webp', 'BMP': '.bmp', 'TIFF': '.tiff'}[fmt] if mode == '图片互转' else ('.docx' if mode == 'PDF → Word' else '.pdf')
    directory.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.convert-', dir=directory) as tmp:
        staged = Path(tmp) / ('result' + suffix)
        if mode == '图片互转':
            image_convert(source, staged, fmt, quality)
        elif mode == 'PDF → Word':
            pdf_to_word(source, staged)
        elif mode == 'Word → PDF':
            word_to_pdf(source, staged)
        else:
            raise ValueError('未知转换模式')
        if not staged.exists() or not staged.stat().st_size:
            raise RuntimeError('转换引擎没有生成有效文件。')
        # Windows rename fails if destination exists: never overwrite even in a race.
        for n in range(10000):
            dest = directory / (source.stem + (f' ({n})' if n else '') + suffix)
            try:
                if dest.exists():
                    continue
                os.rename(staged, dest)
                return dest
            except FileExistsError:
                continue
        raise RuntimeError('同名文件过多，请选择其他目录。')
