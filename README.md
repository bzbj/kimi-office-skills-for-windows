# Kimi Office Skills for Windows

三个适用于 Codex 的本地 Office 技能，使用 Python 处理 Word、PDF 和 Excel。

这是个人维护的 Windows 适配版本，保留 Kimi 技能名称用于识别；不是 Kimi 官方发布，也不需要 Kimi API 或账号。不包含原版 Linux 二进制工具。

| 技能 | 功能 | Python 工具 |
| --- | --- | --- |
| `kimi-docx` | Word 报告、段落、表格、中英混排 | python-docx、lxml |
| `kimi-pdf` | 文本生成 PDF、合并、拆分、提取文本 | ReportLab、pypdf、pdfplumber |
| `kimi-xlsx` | CSV 转 Excel、表头样式、列宽、数据分析 | pandas、openpyxl |

## 安装

需要 Python 3.10 或更高版本，以及可运行 Python 的 Codex 环境。以下命令在 PowerShell 中执行。

```powershell
git clone https://github.com/bzbj/kimi-office-skills.git
cd kimi-office-skills
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

将所需技能复制到 Codex 技能目录。默认目录是用户主目录下的 `.codex/skills`；设置了 `CODEX_HOME` 时使用该目录。下面的命令遇到同名技能会停止，不会覆盖：

```powershell
$skillHome = if ($env:CODEX_HOME) { Join-Path $env:CODEX_HOME 'skills' } else { Join-Path $env:USERPROFILE '.codex/skills' }
[System.IO.Directory]::CreateDirectory($skillHome) | Out-Null
$names = @('kimi-docx', 'kimi-pdf', 'kimi-xlsx')
foreach ($name in $names) {
    $destination = Join-Path $skillHome $name
    if (Test-Path -LiteralPath $destination) { throw "Skill already exists: $name" }
}
foreach ($name in $names) {
    Copy-Item -LiteralPath (Join-Path '.\skills' $name) -Destination $skillHome -Recurse
}
```

启动新的 Codex 会话，以 `$kimi-docx`、`$kimi-pdf` 或 `$kimi-xlsx` 调用。依赖安装在仓库的虚拟环境中，运行脚本时应使用该环境的 Python；也可在自己的 Python 环境中安装同一份依赖。告诉 Codex 所选 Python 解释器的位置即可。

## 命令示例

以下命令从仓库根目录执行。输入文件需由你准备，输出默认放在被 Git 忽略的 `output` 目录。

```powershell
# Word 输入是 UTF-8 JSON
.\.venv\Scripts\python.exe skills/kimi-docx/scripts/docx_windows.py create-report report.json output/report.docx
.\.venv\Scripts\python.exe skills/kimi-docx/scripts/docx_windows.py validate output/report.docx

# PDF 输入是 UTF-8 文本
.\.venv\Scripts\python.exe skills/kimi-pdf/scripts/pdf_windows.py create-from-text input.txt output/report.pdf
.\.venv\Scripts\python.exe skills/kimi-pdf/scripts/pdf_windows.py merge output/merged.pdf first.pdf second.pdf
.\.venv\Scripts\python.exe skills/kimi-pdf/scripts/pdf_windows.py split output/merged.pdf output/pages
.\.venv\Scripts\python.exe skills/kimi-pdf/scripts/pdf_windows.py extract-text output/merged.pdf

# Excel 输入是 UTF-8 CSV
.\.venv\Scripts\python.exe skills/kimi-xlsx/scripts/xlsx_windows.py create-from-csv input.csv output/report.xlsx
.\.venv\Scripts\python.exe skills/kimi-xlsx/scripts/xlsx_windows.py validate output/report.xlsx
```

Word 的 `report.json` 示例：

```json
{
  "title": "示例报告",
  "paragraphs": ["第一段内容。", "中文与 English 混排。"],
  "tables": [{"headers": ["项目", "数量"], "rows": [["示例", 3]]}]
}
```

## 限制与验证

- Word 默认使用 Microsoft YaHei 中文字体；其他系统应改用已安装的中文字体。
- PDF 文本助手使用 ReportLab 的 STSong-Light CID 字体，中文显示效果仍需在目标阅读器验证；不承诺完整嵌入中文字体。
- PDF 文本助手将每行交给 ReportLab Paragraph，输入使用其标记语法；包含字面的 `<`、`>`、`&` 时需转义。
- Excel 可以写入公式，但 openpyxl 不计算公式结果；必要时使用 Excel 或其他经验证的计算引擎重新计算。
- 验证命令只检查文件可读取和基本结构，不替代逐页视觉检查、公式核验或 Office 兼容性测试。
- 不包含旧版 KimiXlsx 的 Linux/OpenXML 二进制文件或原生透视表创建命令。

## 隐私

发布内容只包含技能说明、通用脚本和依赖清单，不包含个人资料、凭据、实际文档或本机绝对路径。
脚本的 `env` 命令会在运行时显示当前 Python 路径，其他命令也可能显示输入内容或输出路径；公开分享日志前请自行脱敏。

## 验证方式

```powershell
.\.venv\Scripts\python.exe tools/smoke_test.py
```

此检查会在系统临时目录中创建合成数据，验证 Word、PDF 和 Excel 的基本读写及 PDF 合并、拆分、提取流程，结束后自动清理。
