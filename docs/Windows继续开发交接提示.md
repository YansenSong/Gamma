# Windows Codex 继续开发交接提示（2026-10-10）

将**整个 Gamma 项目目录**复制到 Windows 后，在该项目中打开 Codex，并把下面的「给 Codex 的提示」作为首条消息。请保留未提交文件；本轮实现尚未提交。`projects/智能行车-审核工作区/` 也必须一同复制，其 `sources/` 内已有三份测试文档的副本。Windows 上不必沿用 Linux 的 `/home/young/桌面/...` 路径。

## 给 Codex 的提示

> 请继续开发当前 Gamma 项目。先读 `docs/交接文档.md`、本文件、`autotesting-app-builder/IMPLEMENTATION_REPORT.md`、`autotesting-app-builder/SKILL.md`、`projects/智能行车-审核工作区/PROJECT_STAGE.md`，再检查 `git status`、实际源码和测试。保留所有未提交修改及样本产物，不要 reset、clean 或覆盖原始文档。
>
> 核心约束：Gamma 是适用于**任意合规评分 PDF、Excel、Word 三文档**的通用 Skill。`projects/智能行车-审核工作区/` 仅是用户提供的真实测试项目；它的业务条款、单元格地址和报告字段不能写死进通用脚本。AutoTesting 仅作只读参考。Excel 原生公式是评分权威，日常程序必须读取用户本次上传的完整已填工作簿，在安全副本上用 Microsoft Excel 重算，再按确认的映射生成 Word。
>
> 目前已完成：上传整份 Excel 的运行链与防篡改/原件保护；通用三文档项目初始化、规则发布、公式准备、Word 位置上下文导出和项目冻结工具；83 个 Linux 单元测试通过。样本 PDF 69 页、原 Excel 3 个工作表且无公式、Word 320 个匿名 `{{}}`。仅 PDF 第 21–28 页有 10 条待审规则候选；六个罚分格 N16:N21 在 `analysis/score-staged.xlsx` 中有待审公式。`templates/score.xlsx` 仍是原始无公式副本，不能当作正式模板。320 个 Word 标记已有位置和上下文，尚未确认字段来源。样本状态是 `review_required`，没有正式业务报告或 EXE。
>
> 请先在 Windows + Microsoft Excel 环境复核依赖和测试，再从待审规则、Excel 映射、公式及 Word 映射继续推进。不要把 mock 测试当原生 Excel 证据，不要把 `pending` 改成 `confirmed` 来绕过门禁；有多种合理解释的条款要记录问题并等待真实业务确认。完成可独立实现的通用工程部分，持续记录实际测试、失败和阻断。最终以交接文档的 G3–G7 门禁为准，只有真实 Windows Excel 重算、完整项目配置、PyInstaller onedir 构建与隔离运行通过后才能声称 EXE 验收。

## Windows 首轮操作

在 PowerShell 中进入复制后的 Gamma 根目录，以实际路径为准：

```powershell
git status --short
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r .\autotesting-app-builder\requirements.txt
cd .\autotesting-app-builder
..\.venv\Scripts\python.exe -B scripts\check_environment.py --require-excel
..\.venv\Scripts\python.exe -B -m unittest discover -s tests -t tests
```

若 `py` 不可用，改用本机 Python 启动器。安装依赖后先核对 Microsoft Excel COM 确实可调用。测试数可能随后续开发变化；记录实际结果即可。

样本工作区的阶段清单是 `projects/智能行车-审核工作区/analysis/formula-stage.json`，其中记录了 **Linux 绝对路径**；Windows 上应从 `review/formula-plan.json` **重新执行 `prepare`**，生成 Windows 本地阶段清单，然后运行 `scripts/validate_formula_stage.ps1`。不要直接拿 Linux 阶段清单作为 Windows 验证证据。候选公式仍待业务审核；原生边界测试通过也不自动等于业务批准。

相关文件：

- `docs/交接文档.md`：完整目标与 G0–G7 门禁。
- `autotesting-app-builder/IMPLEMENTATION_REPORT.md`：改动、测试、当前缺口。
- `projects/智能行车-审核工作区/analysis/评分细则.md`、`open_questions.json`：有来源的待审规则与问题。
- `projects/智能行车-审核工作区/review/`：规则、公式、Excel、Word 和项目配置计划。
- `projects/智能行车-审核工作区/analysis/report-positions.json`、`review/report-positions.csv`：320 个 Word 标记的上下文。
- `projects/智能行车-审核工作区/analysis/formula-diff.json`：六个候选公式写入阶段文件后的结构差异。

## 当前必须保留的事实

1. Linux 只完成工程回归：83 项单元测试通过，`git diff --check` 通过；没有原生 Excel COM、Windows EXE 或隔离目录验收。
2. 样本 `analysis/score-staged.xlsx` 是**待审阶段产物**，原始 `sources/` 与 `templates/score.xlsx` 未被覆盖。
3. 320 个 Word 标记仍是 `pending`；上下文帮助人工判断，不能单凭表格位置推断业务来源。
4. Windows 目标机需安装 Microsoft Excel；PyInstaller `onedir` 必须分发整个目录。
5. 若只复制 Git 已跟踪文件，当前未提交的新脚本、测试和样本工作区会丢失。请复制整个 Gamma 目录，或先用可信方式完整保存工作区，再迁移。
