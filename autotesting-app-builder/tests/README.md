# 技能回归测试

这里测试技能脚本、随技能附带的检查模块和生成流程。测试只证明工程流程行为，不替代具体项目的业务批准或 Microsoft Excel 原生验证。

在技能根目录运行：

```powershell
../.venv/Scripts/python.exe -B -m unittest discover -s tests -t tests
```

`-B` 用于避免在技能目录生成 `__pycache__`。涉及 Excel 公式的纯单元测试可使用测试替身；真实项目仍须在 Windows 上用 Microsoft Excel 重算。
