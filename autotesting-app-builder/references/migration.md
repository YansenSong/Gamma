# 复用与迁移关系

## 直接复用的模块

- `ivista_engine.charter.pdf_extractor.PdfPlumberExtractor`：供 PDF 检查脚本调用，保留页码定位信息和来源哈希。
- `ivista_engine.authoring.excel_analyzer`：供工作簿检查脚本分析工作表结构、公式、数据验证、保护、合并单元格及不支持的功能；`validate_workbook.py` 还使用其 `_archive_features` 辅助函数。
- `ivista_reporting.template_inspector`：现有 DOCX 模板检查器通过包装脚本对外提供，报告文档结构和占位符候选位置；`scripts/report_markers.py` 另外识别冻结版检查器无法发现的无名 `{{ }}` 位置。
- Windows 构建流程：通过每个项目独立的构建脚本，复用原仓库的 PyInstaller 方案和 Qt 依赖系列。

上述模块均通过 `<skill-root>/vendor/` 中的冻结副本使用，由 `scripts/_bootstrap.py` 定位，不依赖实时源码检出。技能不再要求位于 AutoTesting 仓库内部；参见 [environment.md](environment.md)。

## 生成项目中的通用能力

- 将 IVISTA 前端的配色、卡片布局和拖放交互抽取为独立界面设计系统。名称、标志、分类和动态字段由具体项目提供，不复用固定的评分或报告业务代码。
- **评分只依赖 Excel 原生公式。** 输入校验、Excel 自动化和报告生成均独立提供，不复制 Python 评分器。
- Excel 输出基于用户评审后的模板副本，只将明确映射的数据写入独立输出文件。带公式的工作簿必须经过 Microsoft Excel 原生重算，才能信任其缓存结果。
- 若输入工作簿缺少分项评分公式，技能根据章程推导公式，在分阶段验证、Excel 原生边界检查和独立备份之后，仅写入 `init_project.py` 工作区副本 `templates/score.xlsx`。原始工作簿不变；公式仍由 Excel 而非 Python 执行。参见 [formula-authoring.md](formula-authoring.md)。
- DOCX 输出使用显式来源映射和 Jinja 占位符。输入必须是带 `{{ }}` 标记的空白模板，不接受已完成的报告；每个标记在填充前都需要映射并评审。
- 项目验证通过规范化的 Excel 公式绑定哈希，将每项已批准的业务选择与具体公式关联。选择、公式或映射发生变化时，原批准失效。

## 仅供旧版兼容的行为

`ivista_desktop`、固定车辆分类、IVISTA 专用报告版式、现有 RulePack 和当前 EXE 入口仅作为兼容工作流保留。生成应用不会复制这些内容，不应修改其固定业务行为来实现所谓通用化。

## 迁移步骤

1. 将旧版 IVISTA 代码及产物作为历史兼容参考保留，不得在技能或生成应用中沿用其 Python 评分路径。
2. 每组新的章程、工作簿和报告模板，都建立独立的 `generated_apps/<project-id>/`。
3. 确认决定前，核对规则解释、Excel 单元格映射、报告字段、计算权威和来源哈希。
4. 运行生成项目的测试并检查输出。只有真实证据和明确的评审符合 [acceptance.md](acceptance.md)，才能将项目提升为 `business_validated`。
