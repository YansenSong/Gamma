# Gamma 智能行车评分应用构建项目

本仓库提供一个用于生成离线 Windows 评分应用的 Codex 技能，以及一组智能行车业务材料。技能以 Excel 评分表中需要计算的单元格为范围，从 PDF 查找对应规则，为项目副本编写公式，并将计算结果填入 Word 检验报告模板的可填区域。

## 仓库内容

| 目录 | 内容 |
| --- | --- |
| [`autotesting-app-builder/`](autotesting-app-builder/README.md) | 技能说明、执行规范、脚本、应用模板和测试。 |
| [`docs/`](docs/) | 本项目的评分章程 PDF、评分 Excel 和检验报告 Word 模板。 |

三份业务材料分别是：[评分章程](docs/智能行车评分章程.pdf)、[评分 Excel](docs/智能行车评分Excel.xlsx)和[检验报告模板](docs/智能行车检验报告模板.docx)。其中 Word 模板的 `{{ }}` 位置用于自动填充；报告其他内容和格式保持原样。

## 使用入口

从[技能 README](autotesting-app-builder/README.md)了解使用方法与整体流程。实际执行规则见[技能说明](autotesting-app-builder/SKILL.md)，运行环境见[环境说明](autotesting-app-builder/references/environment.md)。生成的应用需要在安装 Microsoft Excel 的 Windows 电脑上完成原生重算。

生成项目保存在本地 `generated_apps/`，Python 环境保存在 `.venv/`；这两个目录不属于仓库提交内容。生成项目的程序和已填公式评分表位置见该项目自己的 `README.md`。
