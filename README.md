# Gamma

**基于 PDF、Excel 和 Word 业务材料，构建离线 Windows 评分应用的通用 Skill。**

Gamma 并非现成的评分软件，而是一套可复用的应用构建技能与工具：以 Excel 中的评分单元格确定计算范围，从 PDF 评分规程核对规则，并将结果映射到 Word 检验报告模板。

## 在 AI 智能体中使用（Agent Skill）

**技能本体是 [`autotesting-app-builder/`](autotesting-app-builder/)，入口为其中的 [`SKILL.md`](autotesting-app-builder/SKILL.md)**；仓库根目录的 README 只是项目介绍。请将**整个技能目录**安装到支持 Agent Skills 的智能体中，不要只复制 `SKILL.md`：运行时还需要同目录下的 `scripts/`、`references/`、`assets/`、`vendor/` 和 `requirements.txt`。

### 1. 安装技能

以下以 **Windows PowerShell + Codex（用户级安装）**为例；首次执行时需安装 Git：

```powershell
git clone https://github.com/YansenSong/Gamma.git
cd Gamma
$skillsHome = Join-Path $HOME ".agents/skills"
New-Item -ItemType Directory -Force -Path $skillsHome | Out-Null
Copy-Item -Path .\autotesting-app-builder -Destination $skillsHome -Recurse
```

安装后应存在 `~/.agents/skills/autotesting-app-builder/SKILL.md`。上述复制步骤适用于**首次安装**；如果目标目录已存在，请先备份或自行更新该目录，避免嵌套复制。不同智能体可将同一技能目录复制到对应位置：

| 智能体 | 用户级技能目录 | 项目级技能目录 |
| --- | --- | --- |
| Codex | `~/.agents/skills/` | `<项目根目录>/.agents/skills/` |
| Cursor | `~/.cursor/skills/` | `<项目根目录>/.cursor/skills/` |
| Claude Code | `~/.claude/skills/` | `<项目根目录>/.claude/skills/` |

每个目录下的最终结构都应是 `<技能目录>/autotesting-app-builder/SKILL.md`。安装后打开或重新启动智能体会话；如果没有发现该技能，检查 `SKILL.md` 路径及智能体是否支持本地 Skills。其他支持 `SKILL.md` 的智能体可按各自文档配置技能搜索路径。

### 2. 准备工作环境

使用此技能生成和验证最终程序，需要**Windows、Microsoft Excel、Python 3.13、[uv](https://docs.astral.sh/uv/)**，以及能够读取本地文件、执行命令的智能体。**在处理业务文件之前**，先在已安装技能目录的**父目录**创建虚拟环境；不要把 `.venv/` 放进技能目录。以下命令承接上面的 Codex 安装示例：

```powershell
cd $skillsHome
uv venv .venv --python 3.13
uv pip install --python .venv/Scripts/python.exe -r autotesting-app-builder/requirements.txt
cd autotesting-app-builder
../.venv/Scripts/python.exe -B scripts/check_environment.py --require-excel
```

初次安装 Python 依赖需要联网；预检必须通过后才能继续。若选择直接在克隆的仓库中使用、不复制技能，则改在**仓库根目录**创建 `.venv/`，并在 `autotesting-app-builder/` 内执行预检。详细要求见 [环境说明](autotesting-app-builder/references/environment.md)。

### 3. 在对话中调用技能

在 Codex 中可以通过 `/skills` 查看已发现的技能，并在提示词里使用 `$autotesting-app-builder` 显式调用；在其他支持 Skills 的智能体中，也可直接描述任务并要求按该技能的 `SKILL.md` 执行。例如：

> 请使用 `autotesting-app-builder` 技能，严格遵循其 `SKILL.md`，根据以下三份同一项目的业务材料生成离线 Windows 评分应用：
>
> - 评分规程：`D:\项目资料\评分规程.pdf`
> - 评分表：`D:\项目资料\评分表.xlsx`
> - 空白报告模板：`D:\项目资料\检验报告模板.docx`（包含 `{{ }}` 填充标记）
>
> 项目名称：`示例项目`；输出目录：`D:\项目输出\示例项目`。
> 请先检查环境，再按 Excel 得分单元格核对 PDF 条款，编写并验证原生 Excel 公式、映射报告字段、生成及测试 Windows 应用。原始三份文件保持不变；业务上无法从材料确定、且会影响得分或报告含义的问题，请列出依据并向我确认。最后告知实际交付路径和验证状态。

将示例路径替换为自己的**真实本地路径**，并确保智能体可访问三个文件。它会按 `SKILL.md` 指导完成分析、公式验证、应用构建和验收，而不是仅根据一句提示直接“生成 EXE”。如果当前环境无法使用 Windows 原生 Excel，智能体可以记录分析或待审结果，**但不能声称评分验证或正式交付已完成**。

更多细节见 [技能使用说明](autotesting-app-builder/README.md) 与 [完整工作流](docs/技能工作流.md)；智能体的技能加载规则以其官方文档为准（[Codex](https://developers.openai.com/codex/skills/)、[Cursor](https://cursor.com/docs/skills)、[Claude Code](https://code.claude.com/docs/en/skills)）。

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

准备同一项目的 **评分规程 PDF、评分表 XLSX、带 `{{ }}` 填充标记的空白报告 DOCX**，使用本技能时，按 [SKILL.md](autotesting-app-builder/SKILL.md) 中的流程执行构建，提供这三份文件、项目名称和输出位置。遇到材料无法明确的计分规则时，需要人工确认。

仓库内的 [test-materials/](test-materials/) 是智能行车示例输入，仅用于验证流程。详细操作与约束见 [技能使用说明](autotesting-app-builder/README.md) 和 [环境说明](autotesting-app-builder/references/environment.md)。

## 仓库结构

| 路径 | 用途 |
| --- | --- |
| [`autotesting-app-builder/`](autotesting-app-builder/) | 通用构建技能、分析与验证脚本、Windows 应用模板及测试 |
| [`test-materials/`](test-materials/) | PDF、Excel、Word 三份示例材料 |
| [`docs/`](docs/技能工作流.md) | 完整中文工作流文档 |

## 构建产物

生成项目默认位于本地 `generated_apps/<项目名>/`，主要关注：

- **Windows 应用**：`dist/<应用名>/<应用名>.exe`（交付时需保留整个同名目录）。
- **公式评分表**：`templates/*.xlsx`，以 `config/project.json` 中的 `excel_output.template` 为准。

`generated_apps/` 和 `.venv/` 不纳入版本控制。**生成后的应用无需 AI 服务或网络，但计算评分仍依赖 Windows 上安装的 Microsoft Excel。**
