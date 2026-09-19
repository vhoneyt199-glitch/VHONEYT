# PDF 文献中文批量重命名工具

本仓库提供 `scripts/rename_pdfs_chinese.py`，用于把 Windows 桌面文件夹 `鼠标与键盘` 中的 PDF 文献按中文题名批量重命名。

## 功能

1. 读取指定文件夹内所有 PDF 文件。
2. 尽量从 PDF 元数据和首页文本提取论文标题。
3. 可选使用 OpenAI 将英文标题翻译成规范中文标题；也支持在 Excel 中人工填写/修改中文题名。
4. 按 `序号-中文题名.pdf` 格式生成目标文件名。
5. 重命名前先生成 `rename_preview.xlsx` 预览表。
6. 只有在预览表中把 `confirm` 列改为 `Y` 的行才会被重命名。
7. 遇到同名目标文件时自动追加 `（2）`、`（3）` 等后缀，避免覆盖。

## 安装依赖

```powershell
pip install -r requirements.txt
```

## 推荐使用流程

在 Windows PowerShell 中进入本仓库目录后运行：

```powershell
python scripts\rename_pdfs_chinese.py preview --folder "C:\Users\你的用户名\Desktop\鼠标与键盘"
```

这会在 PDF 文件夹内生成：

```text
rename_preview.xlsx
```

打开该表格，检查并修改：

- `extracted_title`：脚本从 PDF 中提取到的题名。
- `chinese_title`：中文题名；如果为空，请人工填写。
- `new_filename`：最终文件名，可人工调整。
- `confirm`：确认无误后，把需要重命名的行改为 `Y`。

确认后先试运行：

```powershell
python scripts\rename_pdfs_chinese.py apply --folder "C:\Users\你的用户名\Desktop\鼠标与键盘" --dry-run
```

如果输出无误，再正式重命名：

```powershell
python scripts\rename_pdfs_chinese.py apply --folder "C:\Users\你的用户名\Desktop\鼠标与键盘"
```

## 可选：自动翻译英文题名

如果你有 OpenAI API Key，可以先设置环境变量：

```powershell
setx OPENAI_API_KEY "你的 API Key"
```

重新打开 PowerShell 后运行：

```powershell
python scripts\rename_pdfs_chinese.py preview --folder "C:\Users\你的用户名\Desktop\鼠标与键盘" --translator openai
```

脚本会尽量把英文标题自动翻译到 `chinese_title` 列，但仍建议打开 `rename_preview.xlsx` 人工检查后再执行重命名。
