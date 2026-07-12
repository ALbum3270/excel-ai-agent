# 签名与证书策略（Debug vs Release）

> P0-6：仓库内曾跟踪 VSTO 临时 `.pfx` 与未使用的 `.snk`。正式发布**不得**依赖这些 TemporaryKey。

## 1. 核实结论（2026-07-11）

| 文件 | Git 跟踪 | 用途 | 说明 |
|---|---|---|---|
| `ExcelAi/ExcelAi_TemporaryKey.pfx` | 是（应取消跟踪） | VSTO 清单签名（开发） | 自签名，Subject 为本机用户，含私钥 |
| `WordAi/WordAi_TemporaryKey.pfx` | 是（应取消跟踪） | 同上 | 同上 |
| `PowerPointAi/PowerPointAi_TemporaryKey.pfx` | 是（应取消跟踪） | 同上 | 同上 |
| `OfficeAiAgent.snk` | 是（应取消跟踪） | **无任何 `.vbproj` 引用** | 可安全从版本库移除 |

复核命令：

```powershell
git ls-files "*.pfx" "*.snk"
```

若仍有输出，说明密钥仍在版本库历史或当前索引中，需按第 4 节处理。

**定级**：真实仓库已跟踪含私钥的 TemporaryKey → 按 **P0** 治理。这些是开发机 ClickOnce/VSTO 临时证书，不是商业代码签名证书，但仍不应进入公开仓库。

## 2. 两类签名（不要混用）

| 场景 | 签什么 | 用什么 | 谁信任 |
|---|---|---|---|
| **Debug / 本机开发** | VSTO 清单（`.vsto` / `.dll.manifest`） | 各项目 `*_TemporaryKey.pfx`（本地生成） | 开发者本机「信任一次」或受信任发布者 |
| **Release / 对外分发** | MSI、DLL、可选 XLL | `build/SignArtifacts.ps1` + **受控证书**（环境变量） | 企业/公开 CA 或内网受信任根 |

Release **禁止**使用 `*_TemporaryKey.pfx` 作为对外分发签名。

## 3. Debug：VSTO TemporaryKey

### 3.1 行为

- 各宿主 `.vbproj` 中：`SignManifests=true`，`ManifestKeyFile=*_TemporaryKey.pfx`。
- 证书过期或不存在时，清单签名/调试加载会失败或反复提示。

### 3.2 生成本地密钥（推荐）

```powershell
# 若缺少或已过期，为三端生成自签名 pfx 并写回 ManifestCertificateThumbprint
powershell -File .\scripts\ensure-vsto-temp-keys.ps1

# 仅检查，不写入
powershell -File .\scripts\ensure-vsto-temp-keys.ps1 -WhatIf
```

生成结果仅保留在本机，已被 `.gitignore` 忽略（`*.pfx`）。

### 3.3 本机信任

1. 用 Visual Studio 启动对应 Office 宿主调试插件。  
2. 首次可能提示「无法验证发布者」——开发机可选择信任该发布者，或把证书装到「受信任的根/受信任的发布者」（仅限本机自签名）。  
3. 不要把本机 TemporaryKey 提交回 Git。

## 4. 从版本库移除密钥（维护者操作）

本地文件保留，仅取消 Git 跟踪：

```powershell
git rm --cached -- ExcelAi/ExcelAi_TemporaryKey.pfx `
  WordAi/WordAi_TemporaryKey.pfx `
  PowerPointAi/PowerPointAi_TemporaryKey.pfx `
  OfficeAiAgent.snk

# 确认 .gitignore 含 *.pfx 与 snk 规则后
git status
# 审查 diff 后由维护者 commit（本流程不自动 push）
```

**历史清理**（可选、破坏性）：若仓库曾公开，仅 `rm --cached` 不能从历史删除私钥。需要 `git filter-repo` / BFG 并轮换任何曾用于正式签名的证书。TemporaryKey 为自签名开发证时，通常重新生成本地证即可。

## 5. Release：正式签名

### 5.1 环境变量（任选其一）

| 变量 | 含义 |
|---|---|
| `OFFICE_AI_SIGN_CERT_THUMBPRINT` | 证书存储中的 SHA1 指纹（推荐企业机） |
| `OFFICE_AI_SIGN_PFX` | PFX 路径（**勿放进仓库**） |
| `OFFICE_AI_SIGN_PFX_PASSWORD` | PFX 密码（CI secret） |

### 5.2 命令

```powershell
# 先准备 Release 产物
.\build-installer-prep.bat

# 签名（需 Windows SDK signtool）
powershell -File .\build\SignArtifacts.ps1

# 仅校验
powershell -File .\build\VerifySignatures.ps1

# 完整门禁（可选开签名校验）
powershell -File .\build\RunReleaseChecks.ps1 -VerifySignatures
```

`SignArtifacts.ps1` 默认目标包括 MSI 与四项目 Release DLL。证书必须来自环境变量/安全存储，**不要**指向仓库内 TemporaryKey。

### 5.3 Excel XLL

Debug 下 `ExcelAi-AddIn64.xll` 无签名属预期。公开分发时再对 XLL/MSI 使用同一受控证书策略。

## 6. 禁止事项

1. 不要把正式代码签名 `.pfx` / 密码 / snk 提交到 Git。  
2. 不要在文档或脚本里硬编码密码。  
3. 不要用 TemporaryKey 签对外 MSI。  
4. 不要为大范围「消 diff」而改写 `.vdproj` 里的签名相关二进制资源（若有）。

## 7. 与构建链路的关系

- 日常：`.\build-code.bat`（清单仍用本地 TemporaryKey）。  
- 发版：`.\build-installer-prep.bat` → VS 编 MSI → `SignArtifacts.ps1`。  
- 总览：`docs/build-and-installer.md`。

## 8. 检查清单

- [ ] `git ls-files "*.pfx" "*.snk"` 为空（或仅文档说明的例外）  
- [ ] 本机可 `ensure-vsto-temp-keys.ps1` 后 Debug 启动 Word/Excel/PPT  
- [ ] Release 签名仅通过环境变量证书  
- [ ] CI secrets 不落盘到工作区明文文件  
