# WPS格式转换工具 · 轻转

**WPS格式转换、Word 转 PDF、PDF 转 Word、图片格式批量转换，一个本地桌面工具完成。**

轻转（Qingzhuan）是一款 Windows 本地文件转换工具：利用已安装的 WPS 将 Word 导出为 PDF，将文字型 PDF 转为可在 WPS 中编辑的 Word 文档，并支持 JPG、PNG、WebP、BMP、TIFF 图片互转。

当前为 v0.1.0 桌面初版，使用 Python + Tkinter 开发，尚未提供独立 EXE。使用前需安装 Python 和依赖；只有 Word 转 PDF 需要本机 WPS。

本项目为独立开发工具，与 WPS 官方无隶属关系。

## 功能

- 图片：JPG、PNG、WebP、BMP、TIFF 互转，支持批量和质量设置。透明图片转 JPG/BMP 填白；暂不支持动画及多页图片。
- PDF 转 Word：生成可编辑的 DOCX，可用 WPS 打开；尝试保留布局、图片和表格。复杂布局和公式需要人工检查；不提供 OCR，混合扫描页不会识别为文字。
- Word 转 PDF：调用本机 WPS，支持 DOC/DOCX，需要 WPS 自动化接口可用。

文件在本机处理，无上传功能。保留原文件，同名输出自动增加序号。失败详情双击列表查看。默认保存到本目录的“转换结果”。转换中的 WPS 文档请勿手动操作。

## 安装与运行

已在 Windows、Python 3.13 和带自动化接口的 WPS 环境验证。Word 转 PDF 依赖本机 WPS；其他转换无需启动 WPS。

1. 安装 Python 3.13，并将 Python 加入 PATH。
2. 下载本仓库并解压，在项目目录打开终端。
3. 建议创建虚拟环境并安装依赖：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

4. 双击 `启动工具.bat`，或运行：

```powershell
.\.venv\Scripts\python.exe app.py
```

启动脚本优先使用项目的 `.venv`，否则使用 PATH 中的 Python。开发电脑的 `.vendor` 为本地依赖目录，不随源码上传。

## 使用步骤

1. 选择图片互转、PDF → Word 或 Word → PDF。
2. 点击“添加文件”，可多选。切换转换类型会清空列表。
3. 图片转换可选目标格式和质量；确认保存目录。
4. 点击右下角蓝色“开始转换”按钮。
5. 完成后点击“打开结果文件夹”；双击列表中的结果可查看完整路径或失败原因。

默认保存到项目目录的 `转换结果` 文件夹。同名文件自动编号，不覆盖原文件。转换期间请等待批次完成再关闭程序。

## 已知限制

- PDF 转 Word 尝试还原布局，不保证与原稿完全一致；复杂表格、公式、多栏和字体需人工复核。
- 不支持扫描件 OCR；混合 PDF 中的扫描页不会识别为可编辑文字。
- 密码保护 PDF 需先解除密码。
- 动画与多页图片会提示不支持，避免只转换首帧造成遗漏。
- 图片质量设置仅影响 JPEG / WebP；JPEG / BMP 的透明背景填充为白色。
- 不同 WPS 版本的自动化接口可能不同；连接失败时会显示错误信息。

## 验证

```powershell
.\.venv\Scripts\python.exe smoke_test.py
```

集成检查仅使用自行生成的样例，输出位于 `tmp/smoke`。覆盖五种图片格式、透明背景、同名保护、WPS 导出、中文与表格的 PDF / DOCX 转换及无文字 PDF 拒绝。此检查需要本机 WPS，可能启动 WPS 窗口；样例通过不代表所有文档都能完整还原。

## 依赖

使用 Pillow、pdf2docx、PyMuPDF、python-docx 和 pywin32 等组件。第三方依赖保留各自的许可证，未捆绑 WPS 或依赖二进制文件。

PDF 转换引擎文档：https://pdf2docx.readthedocs.io/en/stable/
