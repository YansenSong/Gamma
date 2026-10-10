# Gamma 实施记录（2026-10-10，持续更新）

## 基线与差异

- Gamma 基线：`c3c48945ace6b3bb11887d796bc5d43f45d5a8f3`，分支 `main`，开始时工作区干净。
- AutoTesting 目录作为只读参考；该本地目录没有 `.git` 元数据，未作修改。
- 需求：上传人工填写的整份 Excel；旧实现：从上传件提取事实，再复制内置空白模板；缺口：报告可能不来自本次上传的工作簿。
- 需求：允许事实格改变但校验关键公式；旧实现：只有内置模板 SHA 校验；本次增加上传件逐格基线检查、输出公式绑定核验和重算后复核。
- 需求：无需 JSON 或表单；旧实现：`scoring.inputs` 必须非空；本次允许空输入列表，界面直接接受 `.xlsx`。
- 需求：原件保护及审计；本次在安全副本重算，记录上传件 SHA 和输出件 SHA，检查原件运行前后哈希。

## 本次实现

- `src/app_service.py`：增加 `run_job_from_workbook()`，校验项目和上传工作簿、复制本次上传件、调用 Microsoft Excel 原生重算、读取该副本得分单元格、生成报告与审计包。失败时清理暂存目录。旧 `run_job_from_facts()` 保留兼容。
- `src/main.py`：批量/拖入 `.xlsx` 走新入口，CLI 增加 `--run-workbook`；界面提示需 Microsoft Excel。
- 运行时与构建时两个项目验证器：允许只读上传工作簿项目的 `scoring.inputs=[]`。
- Excel PowerShell 脚本：显式禁止更新外部链接。
- 测试：`tests/test_uploaded_workbook.py` 使用清楚标记的重算替身，验证数据流、两份工作簿的不同得分、原件哈希、防篡改和失败阻断。生成项目测试骨架增加原生 Excel 上传门禁。
- `scripts/publish_scoring_rules.py`：从已解释、带 PDF 引用的规则计划发布 `评分细则.md` 和 `scoring_rules.json`，核对原文位置与审核状态；该工具不自行解释 PDF 业务含义。
- 通用构建入口 `scripts/init_project.py`：接受任意合规 PDF/XLSX/DOCX，复制原件并生成盘点和待审计划。`scripts/export_report_positions.py` 导出 Word 标记的表格行列上下文。`scripts/freeze_project.py` 在映射与公式均已确认后生成工程项目配置；没有任何 AutoTesting 业务常量。
- 公式写入器已支持未格式化的空白目标格及原 XML 中不存在的行；相对路径公式计划以计划文件目录为基准，仍保持源工作簿原件与非目标内容校验。

## 实测

- 环境：Linux，Python 3.10.12；通过 `UV_CACHE_DIR=/tmp/gamma-uv-cache uv pip install --python .venv/bin/python -r autotesting-app-builder/requirements.txt` 安装依赖。
- `../.venv/bin/python -B scripts/check_environment.py`：`status=ok`，vendored 模块与依赖齐全；Excel COM 未注册（Linux 平台）。
- 命令：从 `autotesting-app-builder` 执行 `../.venv/bin/python -B -m unittest discover -s tests -t tests`：83 tests，OK。
- `git diff --check`：通过。
- 只读真实泊车样本盘点：规程 PDF 53 页；Excel 2 个工作表、3169 个有内容或样式的单元格，其中 206 个公式候选；报告 DOCX 无任何 `{{}}`，检查器按“非空白模板”阻断。该样本不能直接作为本次三文档构建输入，需在副本中明确标出待自动填写的位置并审核其来源。
- 用户提供的 `/home/young/桌面/docs` 三份新材料用于下一阶段测试：PDF 69 页；Excel 3 个工作表、1078 个有内容或样式的单元格，原表无公式；Word 有 320 个匿名 `{{}}`。已生成 `Gamma/projects/智能行车-审核工作区/`，包含三原件副本、盘点、待审计划和 320 行上下文清单。该目录仍是 `review_required`，没有擅自填写 320 项来源或编造评分公式。
- 已对用户样本 PDF 第 21–28 页编写 10 条带位置引用的评分规则候选，发布 `评分细则.md` 和 `open_questions.json`。全部保持 `pending`；其余页仍需覆盖审查，文中不授予业务批准。
- 对 PDF 第 21、22、23、25、26、28 页作原图核对，修正了文本提取失真的表 16 智能拟人化阈值与舒适性得分公式。项目工作区的公式计划暂覆盖 N16:N21 六个封闭场地罚分格；`prepare` 生成阶段 Excel，`diff_workbooks` 显示仅这六格变化、结构变化 0。原始 Excel 与项目评分模板 hash 均保持 `d2e364a73174429d34d927c68f8d964aa62c6ccc637aab3f2f10bb7170c6bcad`。阶段 Excel 未经 Excel COM 验证，不视为交付模板。
- 通用项目初始化现在将 Word 标记的表格行列、整行文字、前一行文字和邻近文本一并写入待审映射计划。已将用户样本的 320 个位置补齐这些上下文；字段名、来源和审核状态仍待确认。
- 再次核对用户三份原件与工作区 `sources/` 副本的 SHA-256：三个角色均与初始化记录一致，原件未被改写。
- 原生 Microsoft Excel COM：本机无法执行，未验收。上述新测试的 `TEST_DOUBLE_ONLY` 是工程回归测试，不构成原生计算证据。
- Windows PyInstaller、隔离目录 EXE：未构建、未验收，暂无可交付 EXE 路径。

## 未完成门禁

- 三份真实业务文档的规则解释与审核签署；目前仅有可追溯规则发布通道，未对真实规程生成已审 `评分细则.md`。
- 真实 Excel 公式写入、Microsoft Excel 边界测试、全部 Word `{{}}` 的业务映射，以及一套完整业务项目产物。
- Windows Excel COM 端到端和 PyInstaller onedir 隔离运行；目标机器仍须安装 Microsoft Excel。

当前状态为 `engineering_sample / release_blocked`。不能据此生成或宣称正式业务报告。下一步在 Windows + Excel 环境用确认过的项目配置和 `examples/filled_score.xlsx` 运行生成项目测试，再打包并做隔离目录验收。
