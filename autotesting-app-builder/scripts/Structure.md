# scripts 文件结构

| 文件或目录 | 作用 |
| --- | --- |
| `_bootstrap.py` | 从技能目录加载随附模块。 |
| `check_environment.py` | 检查依赖和 Excel 计算环境。 |
| `init_project.py` | 从三份材料初始化生成项目。 |
| `inspect_pdf.py` | 盘点 PDF 页面与内容。 |
| `inspect_excel.py` | 盘点工作簿结构和评分单元格。 |
| `inspect_formula_empty.py` | 检查公式输出的空值状态。 |
| `inspect_report.py` | 检查 DOCX 模板标记。 |
| `build_review_inventory.py` | 汇总需要评审的字段与来源。 |
| `export_field_review.py` | 导出字段评审清单。 |
| `export_report_positions.py` | 导出报告可填位置。 |
| `report_markers.py` | 识别与定位 Word 模板标记。 |
| `map_report_template.py` | 建立报告位置与来源映射。 |
| `name_report_template.py` | 给匿名模板标记命名。 |
| `publish_scoring_rules.py` | 发布经核对的结构化评分规则。 |
| `validate_scoring_spec.py` | 验证评分细则来源与覆盖。 |
| `write_scoring_formulas.py` | 分阶段准备、验证和写回 Excel 公式。 |
| `validate_formula_stage.ps1` | 通过 Microsoft Excel 验证公式边界案例。 |
| `diff_workbooks.py` | 比较公式写回前后的工作簿差异。 |
| `validate_workbook.py` | 验证工作簿及重算结果。 |
| `validate_report.py` | 验证报告填充值与标记。 |
| `validate_project_spec.py` | 验证生成项目配置。 |
| `freeze_project.py` | 冻结来源哈希与项目配置。 |
| `review_workflow.py` | 评审包导出及提交记录核验。 |
| `write_manifest.py` | 写入测试、构建和已知限制的交付清单。 |
| `verify_release.py` | 核查交付状态和项目绑定。 |

`README.md` 说明本目录用途；`Structure.md` 为本文件。
