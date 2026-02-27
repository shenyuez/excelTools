# ExcelTools — Excel 出生年月格式统一工具

将 Excel 中各种格式的出生年月批量转换为统一格式（默认 `YYYY.MM`）。

## 功能

- 支持多种输入格式自动识别：`YYYY.MM`、`YYYY-MM-DD`、`YYYYMMDD`、`YYYYMM` 等
- 输出格式可选：`YYYY.MM` / `YYYY-MM` / `YYYY/MM` / `YYYYMM` / 自定义模板
- 输入规则可在界面中增删改、启用/禁用，无需改代码
- 支持指定表头所在行号（应对表头不在第一行的情况）
- 规则编辑器内置实时正则测试
- 多 Sheet 自动遍历处理

## 使用方式

### 直接运行（需要 Python 环境）

```bash
pip install openpyxl
python format_birthday_gui.py
```

### 打包为 exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "ExcelBirthdayConverter" format_birthday_gui.py
# 输出在 dist/ExcelBirthdayConverter.exe
```

或者直接双击项目根目录的 `build.bat`，效果相同。

## 界面说明

| 区域 | 说明 |
|------|------|
| 输入文件 | 选择要处理的 Excel 文件 |
| 输出文件 | 留空则自动生成 `原文件名_fixed.xlsx` |
| 列名 / 表头行号 | 填写目标列的表头文字，以及表头所在行（默认第 1 行）|
| 输出格式 | 选择预设格式或填写自定义模板，如 `{year}/{month}` |
| 输入格式规则 | 双击行切换启用/禁用；可添加自定义正则规则 |

## 自定义规则说明

每条规则包含：

- **名称**：便于识别的描述
- **正则模式**：用于匹配输入值的 Python 正则表达式
- **年份组号 / 月份组号**：正则捕获组的编号（从 1 开始）
- **全匹配**：勾选则使用 `re.fullmatch`，否则使用 `re.search`

规则按顺序匹配，第一个命中的规则生效。

## 环境要求

- Python 3.10+
- openpyxl
