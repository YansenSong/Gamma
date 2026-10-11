# Windows 评分应用模板

此目录是生成应用的骨架，包含项目配置、源代码、构建脚本和测试。初始化后的项目会有自己的 `config/project.json`、Excel 与 DOCX 模板及业务测试；`build.ps1` 用于打包 Windows 应用。

应用读取操作员填写的工作簿，在安全副本上校验输入并调用 Microsoft Excel 原生重算。报告只填写已映射的 Word 模板位置，不改其他内容。运行机器需要安装 Microsoft Excel；程序运行还依赖 `dist/<应用名>/` 目录中的配套文件。

打包后的 EXE 通常位于 `dist/<应用名>/<应用名>.exe`；公式评分表位于 `config/project.json` 的 `excel_output.template` 指定位置，通常是 `templates/*.xlsx`。
