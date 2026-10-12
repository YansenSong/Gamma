# Gamma

**基于 PDF、Excel 和 Word 业务材料，生成离线 Windows 评分应用的 Codex Skill。**

Gamma 并非现成的评分软件，而是一套应用构建工具：以 Excel 中的评分单元格确定计算范围，从 PDF 评分规程核对规则，并将结果映射到 Word 检验报告模板。

## 核心功能

- **规则与公式**：按 Excel 现有评分项定位 PDF 条款，编写评分细则和 Excel 公式。
- **原生算分**：由 Microsoft Excel 重算并提供分数，Python 不另写一套评分算法。
- **报告生成**：只填充 DOCX 模板中的 `{{ }}` 位置，保留其他文字与格式。
- **应用交付**：构建可离线运行的 Windows 桌面应用，附带验证与审计记录；原始材料保持不变。

## 快速开始

**环境要求：** Windows、Microsoft Excel、Python 3.13 和 [uv](https://docs.astral.sh/uv/)。

在仓库根目录创建环境并执行预检：

```powershell
uv venv .venv --python 3.13
uv pip install --python .venv/Scripts/python.exe -r autotesting-app-builder/requirements.txt
cd autotesting-app-builder
../.venv/Scripts/python.exe -B scripts/check_environment.py --require-excel
```

准备同一项目的 **评分规程 PDF、评分表 XLSX、带 `{{ }}` 填充标记的空白报告 DOCX**，然后在 Codex 中按 [SKILL.md](autotesting-app-builder/SKILL.md) 执行构建流程，提供这三份文件、项目名称和输出位置。遇到材料无法明确的计分规则时，需要人工确认。

仓库内的 [test-materials/](test-materials/) 是智能行车示例输入，仅用于验证流程。详细操作与约束见 [技能使用说明](autotesting-app-builder/README.md) 和 [环境说明](autotesting-app-builder/references/environment.md)。

## 仓库结构

| 路径 | 用途 |
| --- | --- |
| [`autotesting-app-builder/`](autotesting-app-builder/) | Codex Skill、分析与验证脚本、Windows 应用模板及测试 |
| [`test-materials/`](test-materials/) | PDF、Excel、Word 三份示例材料 |
| [`docs/`](docs/技能工作流.md) | 完整中文工作流文档 |

## 构建产物

生成项目默认位于本地 `generated_apps/<项目名>/`，主要关注：

- **Windows 应用**：`dist/<应用名>/<应用名>.exe`（交付时需保留整个同名目录）。
- **公式评分表**：`templates/*.xlsx`，以 `config/project.json` 中的 `excel_output.template` 为准。

`generated_apps/` 和 `.venv/` 不纳入版本控制。**生成后的应用无需 AI 服务或网络，但计算评分仍依赖 Windows 上安装的 Microsoft Excel。**
