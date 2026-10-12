# Excel 评分规范

每个生成项目使用 `scoring.mode = excel_native`。不允许使用 Python 评分器、表达式树、参考分数或备用评分权威。

提供的工作簿可能根本没有分项评分公式：得分单元格为空，仅汇总 `SUM` 行可能含有公式。此时必须先编写公式，才能建立绑定。首先盘点全部工作簿评分目标，再定位对应章程规则并推导公式。经过分阶段验证、Microsoft Excel 原生边界检查及唯一备份后，仅向 `init_project.py` 创建的工作区副本 `templates/score.xlsx` 写入公式，随后按常规方式绑定；参见 [formula-authoring.md](formula-authoring.md)。用户的源工作簿必须保持不变。编写公式不会改变评分权威：所有分数都由 Excel 计算，技能编写的公式与用户提供的公式同样需要评审、来源引用和批准。

声明用于数据校验的输入，并为每个 `excel_output.output_cells` 条目提供 `scoring.formula_bindings`。每条绑定记录输出 ID、工作表、单元格、精确公式和章程/来源引用。验证器将公式与哈希绑定的工作簿进行比对。只有尚无输出映射、且处于阻塞评审状态的项目才允许空绑定；此类项目不能执行评分。

```json
{
  "mode": "excel_native",
  "authority": {"type": "Microsoft Excel", "require_full_recalculation": true},
  "inputs": [{"id": "latency_ms", "type": "number", "minimum": 0}],
  "formula_bindings": [{
    "output": "score", "sheet": "Score", "cell": "B6",
    "formula": "=IF(B4<=200,100,IF(B4<=300,80,60))",
    "source_refs": ["synthetic charter, page 1"]
  }]
}
```

上例仅为合成示意，不是经过批准的业务规则。需要用户决定的事项，通过 `decision_bindings` 绑定工作表/单元格/公式的规范化 SHA-256。来源哈希以及四项限定范围的批准内容哈希，必须覆盖当前材料。评审者标记为 AI 还是人工，都不会自行修改公式。当业务决定需要改变公式时，只能修改明确获授权的工作簿副本，并重新取得批准。

必须在 Microsoft Excel 完整原生重算、保存并重新打开工作簿后，才能读取评分结果。记录 Excel 版本及工作簿和报告的实际哈希。缺失、非有限值、错误值或类型不符的映射结果都应拒绝。Excel 计算失败时，绝不能以 Python 计算值或示例值替代。

章程评审可能发现阈值、单位、上限、例外或舍入冲突。评审应记录 AI 的解释证据及必要的用户决定，但不得实现第二套评分算法。
