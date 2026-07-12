# 构建与安装包链路（code vs installer）

> P0-5：把「代码项目构建」和「安装包构建」明确分开，避免把 `.vdproj` 失败误判成代码编译失败。

## 1. 两条链路

| 链路 | 目的 | 入口 | 产物 |
|---|---|---|---|
| **Code build** | 日常开发 / 验证 VB.NET + VSTO 插件 | `build-code.bat` 或 `scripts/build-code-projects.ps1` | `*/bin/Debug` 或 `*/bin/Release` 中的 DLL |
| **Installer prep** | 打 MSI 前准备与预检 | `build-installer-prep.bat` 或 `scripts/build-installer-prep.ps1` | Release 代码产物 + SourcePath 审计通过 |
| **Installer MSI** | 生成安装包 | Visual Studio 中打开 `OfficeAgent/OfficeAgent.vdproj` 并 Build | MSI（依赖 Installer Projects 扩展） |

**不要**用「整解决方案 Debug Rebuild」作为代码是否可编的判断：`OfficeAgent.vdproj` 的 `SourcePath` 大量指向 `bin\Release\...`，Debug 整编时安装项目预验证常会失败，**这与 ShareRibbon/Word/Excel/PPT 代码能否编译无关**。

## 2. 日常开发（只编代码）

```powershell
# 推荐：四代码项目 Debug
.\build-code.bat
# 或
powershell -File .\scripts\build-code-projects.ps1 -Configuration Debug

# 仅共享库
.\build-shareribbon.bat

# Excel + PowerPoint
.\build-excel-ppt.bat
```

兼容别名：`build-all.bat` 等同于 Debug 代码四项目构建（**不含** MSI）。

成功标准：四个 `.vbproj` 全部 PASS。  
**不要**要求 `OfficeAgent.vdproj` 在此步骤成功。

## 3. 发布 / 打安装包（强制顺序）

```text
1) build-code Release
2) AuditInstallerInputs（SourcePath 文件是否存在）
3) 在 VS 中 Build OfficeAgent.vdproj（需 Installer Projects 扩展）
4) 可选：签名、RunReleaseChecks
```

一键准备（步骤 1–2）：

```powershell
.\build-installer-prep.bat
# 或
powershell -File .\scripts\build-installer-prep.ps1
```

若 Release 产物已存在、只想重跑审计：

```powershell
powershell -File .\scripts\build-installer-prep.ps1 -SkipBuild
```

完整发布门禁（版本审计 + Release 构建 + 安装输入审计 + smoke 等）：

```powershell
powershell -File .\build\RunReleaseChecks.ps1
# 已有产物时可：
powershell -File .\build\RunReleaseChecks.ps1 -SkipBuild
```

然后：

1. 安装并启用 [Microsoft Visual Studio Installer Projects](https://marketplace.visualstudio.com/)（VS 2022/2026 实例需匹配）。  
2. 打开 `OfficeAgent/OfficeAgent.vdproj`（或从 `AiHelper.sln` 加载）。  
3. 配置选 **Release**，Build 安装项目生成 MSI。  

VS 打不开 `.vdproj` 时见 [VS2026-InstallerProjects.md](./VS2026-InstallerProjects.md)。

## 4. 为什么 vdproj 依赖 Release

`OfficeAgent/OfficeAgent.vdproj` 中大量：

```text
"SourcePath" = "8:..\\WordAi\\bin\\Release\\...."
"SourcePath" = "8:..\\ExcelAi\\bin\\Release\\...."
"SourcePath" = "8:..\\PowerPointAi\\bin\\Release\\...."
```

因此：

- Debug 日常开发：**只跑 code build**。  
- 打包装箱：**必须先 Release code build**，再审计，再编 MSI。

审计脚本：`build/AuditInstallerInputs.ps1`  
（默认只检查带目录的 `SourcePath`，避免把 GAC/框架 DLL 误判为缺失；严格模式加 `-IncludePlainFiles`。）

## 5. 自定义安装动作

- 仓库中 **当前不存在** `OfficeAgentSetupCustomActions/` 项目。  
- 文档中「如存在」表示可选扩展位，不是现成目录。  
- 若需要安装过程代码逻辑，应**新建**该 VB.NET 自定义动作项目，而不是大改 `.vdproj`。  
- `.vdproj` 保持最小 diff，禁止批量格式化。

## 6. 中长期方向（不在本轮实现）

- 评估 WiX / MSIX / 可脚本化 bootstrapper，降低对 VS Installer Projects 扩展的依赖。  
- 减少 MSI 内重复 DLL，统一共享依赖拷贝策略。  
- CI：PR 只跑 code build + smoke；Release 流水线再跑 installer prep（MSI 仍可能需 Windows + VS 扩展）。

## 7. 命令速查

| 场景 | 命令 |
|---|---|
| 开发编译 | `.\build-code.bat` |
| 开发编译 Release | `.\build-code.bat Release` |
| 打 MSI 前准备 | `.\build-installer-prep.bat` |
| 完整 Release 门禁 | `powershell -File .\build\RunReleaseChecks.ps1` |
| 仅审计安装输入 | `powershell -File .\build\AuditInstallerInputs.ps1` |
| 仅代码四项目 | `powershell -File .\scripts\build-code-projects.ps1` |
| 生成本地 VSTO TemporaryKey | `powershell -File .\scripts\ensure-vsto-temp-keys.ps1` |
| Release 产物签名 | `powershell -File .\build\SignArtifacts.ps1`（需 `OFFICE_AI_SIGN_*`） |

签名与证书策略详见 [signing-and-certificates.md](./signing-and-certificates.md)。

## 8. 常见误判

| 现象 | 正确理解 |
|---|---|
| `devenv AiHelper.sln /Rebuild Debug` 在 OfficeAgent 失败 | 多半是 vdproj 找 Release 路径，不是 Word/Excel 代码坏了 |
| 只有 Debug 输出却去编 MSI | SourcePath 指向 Release，必然缺文件 |
| 未装 Installer Projects 扩展 | 代码项目仍可编；只是打不开/编不了 `.vdproj` |
| 大改 `.vdproj` 格式 | 高风险，易导致 VS 加载失败；保持最小 diff |
