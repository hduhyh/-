# CSV 列名提取合并工具

这是一个 Windows GUI 工具：根据模板中的列名，从输入文件夹里的 CSV/XLSX 文件提取对应列，并合并输出为一个 CSV 文件。

## 功能

- 选择输入文件夹、输出文件夹和列名模板。
- 自动记住上一次选择的路径、文件名和选项。
- 支持 CSV、XLSX，CSV 会尝试常见中文编码并自动识别常见分隔符。
- 可选递归扫描子文件夹、读取 XLSX 的全部工作表。
- 缺失列自动补空；输出增加 `Source_File` 和 `Source_Sheet` 来源列。
- 后台处理、进度显示、失败文件日志和安全的临时文件写入。

## 模板格式

模板支持 `.xlsx` 或 `.csv`。把需要提取的列名逐行放在模板的第一个非空列中；可以有“列名”“字段名”等说明性表头。程序会按模板顺序输出列。

## 本地运行

```bash
python -m pip install -r requirements.txt
python csv_column_merger_gui.py
```

## 在 GitHub 打包 Windows EXE

1. 把本项目文件上传到 GitHub 仓库的 `main` 或 `master` 分支。
2. 推送代码后，GitHub Actions 会自动运行 `Build Windows application`；也可以在仓库的 **Actions** 页面手动运行。
3. 构建完成后，在该次运行页面的 **Artifacts** 区域下载 `CSV列名提取合并-Windows`。
4. 解压后运行 `CSV列名提取合并.exe`，无需另外安装 Python。

程序未进行商业代码签名，因此 Windows 首次运行时可能显示 SmartScreen 提示。配置保存在 `%APPDATA%\CSVColumnMerger\settings.json`。

