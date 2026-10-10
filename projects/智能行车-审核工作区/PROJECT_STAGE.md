# 项目构建状态

当前为 `review_required`，不是可运行或已批准的评分应用。

1. 核对 `analysis/` 中的三文件盘点及原始文件哈希。
2. 已形成第 21–28 页评分条款的 10 条待审候选，见 `analysis/评分细则.md` 与 `analysis/open_questions.json`。已查看第 21、22、23、25、26、28 页原图，核实表 16 阈值及第 6.3.8 节公式。其余 PDF 页尚未完成条款覆盖审查；所有规则仍为 `pending`。
3. `review/excel-map-plan.json` 已列出 6 个封闭场地罚分输入/得分格候选；`review/formula-plan.json` 已为 N16:N21 制定待审公式。`analysis/score-staged.xlsx` 确实含 6 个 Excel 公式，`analysis/formula-diff.json` 记录仅这 6 格变化。尚未运行 Microsoft Excel 原生边界验证，`templates/score.xlsx` 仍是无公式原样副本，不能用于正式评分。
4. 参照 `review/report-positions.csv` 的表格行列上下文，逐个解释 320 个 Word 标记并填写 `review/report-map-plan.json`；生成具名模板。
5. 冻结项目配置、运行生成项目测试，再执行 Windows onedir 构建与隔离验收。
