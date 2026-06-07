# AGENTS.md

> 请用中文与用户交互。代码符号、路径、命令、API 名称保持原文。

本文件是仓库级 Agent 导航图，帮助 AI Agent 快速判断“要改某类功能应该看哪里”。工程级硬规则、构建验证、VB.NET/VSTO 避坑见根目录 `CLAUDE.md`。

## Purpose

- `CLAUDE.md`：项目级工作协议，规定必须遵守的规则、避坑和验证清单。
- `AGENTS.md`：仓库级导航，说明目录职责、任务入口和局部规则位置。
- 子目录 `AGENTS.md`：进入某个模块后的局部补充规则。距离目标文件最近的 `AGENTS.md` 优先作为局部上下文，但不得覆盖根 `CLAUDE.md` 的工程级安全/质量规则。

## Read Order for Agents

处理任务前按顺序读取：

1. 根目录 `CLAUDE.md`
2. 根目录 `AGENTS.md`
3. 目标目录最近的 `AGENTS.md`
4. 相关 `.vbproj` / `packages.config`
5. 目标功能源码

如果任务跨多个项目，例如同时修改 `ShareRibbon/` 和 `WordAi/`，需要同时阅读这些目录下的局部 `AGENTS.md`。

## Top-level Structure

```text
AiHelper/
├── AiHelper.sln                         # Visual Studio 解决方案
├── CLAUDE.md                            # AI Agent 项目级规则
├── AGENTS.md                            # 仓库级导航
├── ExcelAi/                             # Excel VSTO 插件
├── WordAi/                              # Word VSTO 插件
├── PowerPointAi/                        # PowerPoint VSTO 插件
├── ShareRibbon/                         # 共享组件库
├── OfficeAgent/                         # 安装包项目 (.vdproj)
├── OfficeAgentSetupCustomActions/        # 安装自定义动作 VB.NET 项目
├── docs/                                # 调试、迁移、说明文档
├── openspec/                            # 需求/变更规格与归档
├── packages/                            # NuGet packages 目录
└── .codegraph/                          # CodeGraph 语义索引（如存在）
```

## Where to Look

| 任务 | 位置 | 说明 |
|---|---|---|
| 共享 Chat UI / WebView2 / 服务 | `ShareRibbon/Controls/` | `BaseChatControl`、HTTP/streaming 服务、共享 UI 组件 |
| 前端 HTML/JS/CSS 资源 | `ShareRibbon/Resources/` 及相关项目资源项 | Office Virtual Server 加载入口，改动后检查 `.vbproj` 注册和 HTML 引用 |
| 配置、API Key、Prompt | `ShareRibbon/Config/` | 配置管理和持久化逻辑 |
| MCP 协议集成 | `ShareRibbon/Mcp/` | StreamJsonRpc / MCP client 相关实现 |
| Excel 功能 | `ExcelAi/` | 单元格、Range、Sheet、图表、ExcelDna、批量数据生成 |
| Word 功能 | `WordAi/` | 文档、段落、续写、翻译、OpenXml 处理 |
| PowerPoint 功能 | `PowerPointAi/` | 演示文稿、幻灯片、形状、续写、翻译 |
| 安装包定义 | `OfficeAgent/` | `.vdproj` 安装项目，慎改，保持最小 diff |
| 安装自定义动作 | `OfficeAgentSetupCustomActions/` | 安装流程辅助逻辑，优先在这里承载可代码化安装行为 |
| 调试/迁移文档 | `docs/` | Visual Studio/VSTO 调试、迁移指南等 |
| 需求/变更规格 | `openspec/` | 变更说明、规格草案、归档记录 |
| NuGet 依赖 | 各项目 `packages.config` 与根 `packages/` | 判断技术栈、依赖版本、目标框架 |

## Local AGENTS.md Files

当前仓库存在多层局部 Agent 文档：

- `ExcelAi/AGENTS.md`
- `WordAi/AGENTS.md`
- `PowerPointAi/AGENTS.md`
- `ShareRibbon/AGENTS.md`
- `ShareRibbon/Controls/AGENTS.md`
- `ShareRibbon/Mcp/AGENTS.md`
- `ShareRibbon/Config/AGENTS.md`

进入对应目录前必须阅读局部文件。局部文件用于说明该模块的入口、约束和反模式；根目录 `CLAUDE.md` 仍是通用质量规则来源。

## Project Ownership Boundaries

### ShareRibbon

`ShareRibbon/` 是跨 Office 插件共享的核心库。以下能力优先放这里：

- 共享 Chat 控件与 WebView2 UI 基础设施
- AI 请求、streaming、HTTP 通信
- MCP client 与工具调用协议
- 配置、Prompt、API Key 管理
- 可被 Excel/Word/PowerPoint 共同复用的服务、模型、工具类

不要把共享服务复制到单个 Office 插件项目中。

### ExcelAi

`ExcelAi/` 只处理 Excel 相关能力：

- Workbook / Worksheet / Range / Cell 操作
- 数据分析、公式辅助、图表生成
- ExcelDna 自定义函数
- Excel 专属 Ribbon、任务窗格、命令 schema

不要访问 Word 或 PowerPoint 对象模型。

### WordAi

`WordAi/` 只处理 Word 相关能力：

- Document / Paragraph / Range / Selection 操作
- 文档生成、续写、翻译
- OpenXml 文档处理
- Word 专属 Ribbon、任务窗格、命令 schema

不要访问 Excel 或 PowerPoint 对象模型。

### PowerPointAi

`PowerPointAi/` 只处理 PowerPoint 相关能力：

- Presentation / Slide / Shape 操作
- 幻灯片生成、排版、续写、翻译
- PowerPoint 专属 Ribbon、任务窗格、命令 schema

不要访问 Excel 或 Word 对象模型。

### OfficeAgent 与 OfficeAgentSetupCustomActions

- `OfficeAgent/` 的 `.vdproj` 是安装包定义，自动化修改风险高，除非必要不要大范围编辑。
- `OfficeAgentSetupCustomActions/` 是 VB.NET 自定义动作项目，适合承载安装过程中需要代码实现的逻辑。
- 如果安装逻辑可以通过自定义动作表达，优先改 `OfficeAgentSetupCustomActions/`，避免直接重写 `.vdproj`。

## Common Agent Mistakes

修改本仓库时重点避免：

1. 把 C# 语法写进 VB.NET，例如 `var`、`List<T>`、`??`、`=>`、`new()`。
2. 新增 `.vb` 文件后忘记加入 `.vbproj` 的 `<Compile Include="...">`。
3. 新增 `.js` / `.css` / `.html` 后忘记加入 `.vbproj`，或没有被 HTML 引用。
4. 修改前端资源但没有考虑 Office Virtual Server 访问路径，导致 WebView2 控制台 `ERR_FILE_NOT_FOUND`。
5. 在 Excel/Word/PowerPoint 项目里复制本应位于 `ShareRibbon/` 的共享逻辑。
6. 跨 Office 应用直接访问对象模型，例如在 `WordAi/` 中访问 Excel 对象。
7. 新增 SQLite 字段只改建表逻辑，没有写升级迁移。
8. 大范围格式化 `.vdproj`、`.vbproj`、Designer 文件或生成文件。
9. 多 Agent 并行创建类型时，没有先统一共享模型/接口，导致 API 不一致。

## Commands

常用命令索引如下；更完整规则见 `CLAUDE.md`。

```bash
# 还原 NuGet 包
msbuild AiHelper.sln -t:Restore

# 构建整个解决方案
msbuild AiHelper.sln

# 构建共享库
msbuild ShareRibbon/ShareRibbon.vbproj

# 构建单个 Office 插件
msbuild ExcelAi/ExcelAi.vbproj
msbuild WordAi/WordAi.vbproj
msbuild PowerPointAi/PowerPointAi.vbproj

# 构建安装自定义动作项目
msbuild OfficeAgentSetupCustomActions/OfficeAgentSetupCustomActions.vbproj
```

## Documentation Maintenance Rules

- 不要在本文件写容易过时的生成日期、当前分支、临时 commit 信息。
- 新增顶层目录、重要子系统或局部 `AGENTS.md` 后，应更新本文件的导航表。
- 技术栈、构建命令、验证清单的权威位置是 `CLAUDE.md`；本文件只保留导航级摘要。
- 如果局部 `AGENTS.md` 与实际代码明显不一致，优先相信代码和项目文件，并修正文档。
