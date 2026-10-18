# 安装 / 启用 / 停用 / 卸载步骤（任务 6.3）

本文件是**正式安装路径**的步骤与回退说明。集中验收仍使用 `acceptance/preview-session.md` 的 `Register-Preview.ps1`；二者注册同样的四个键，区别只在于 `Register-Preview.ps1` 额外承担验收模式标记。

**当前状态：未执行正式安装。** 脚本与回退隔离性已离线验证通过（`tools/Test-InstallRollback.ps1`，20 项断言），实际安装需在关卡开放并获得授权后执行。

## 0. 前提

- 已完成离线构建：`.\build.ps1 -Test`（生成 `artifacts/preview/SHA256SUMS.txt`）。
- **OneNote 已正常退出**：脚本检测到 `ONENOTE.EXE` 时会拒绝安装，**不会结束任何进程**。
- 仅支持**桌面版** OneNote；商店版（UWP）不加载本插件。
- 使用 **64 位 Windows PowerShell**。

## 1. 安装（仅本插件，按用户注册）

```powershell
.\tools\Install-Addin.ps1 -DryRun    # 预演：只打印将要创建的键，不落盘
.\tools\Install-Addin.ps1            # 实际安装
.\tools\Install-Addin.ps1 -Verify    # 复核：STATE=REGISTERED / LoadBehavior=0
```

安装只创建以下四个 HKCU 键（不写 HKLM、不改系统级设置）：

| 键 | 作用 |
|---|---|
| `HKCU\Software\Classes\CLSID\{E036DE73-418B-45F7-973F-A4B09D58E515}\InprocServer32` | mscoree + 托管类 + CodeBase |
| `HKCU\Software\Classes\AppID\{E036DE73-...}` | `DllSurrogate`（代理承载） |
| `HKCU\Software\Classes\OneNoteMarkdown.Preview` | ProgID 别名 |
| `HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteMarkdown.Preview` | 加载项项 + `LoadBehavior=0` |

**不会设置自动加载**：脚本刻意不提供 `LoadBehavior=3`，启用/停用由 OneNote 自身管理。请勿手工改 `LoadBehavior`。

先校验构建完整性（`SHA256SUMS.txt` 全部匹配）、拒绝已存在的注册、安装后回读校验四个键与 `LoadBehavior=0`，并比对同级的**其他加载项键**是否被改动。

## 2. 启用 / 停用

1. 正常启动 OneNote → 文件 → 选项 → 加载项 → 管理“COM 加载项” → 转到。
2. 勾选 **OneNote Markdown Preview (manual load)** → 确定（此时只是加载，不是自动启动）。
3. 功能区出现 **Markdown 预览** 选项卡；点击“启用”后快捷键 `Ctrl+Shift+D` 才生效。
4. 停用：点击“停用”，或取消勾选该加载项；停用会释放键盘钩子，快捷键不再拦截。
5. 若出现“上次操作未完成”，转换会被停用，需手动重新启用（见 `usage.md` 第 7 节）。

## 3. 卸载

```powershell
.\tools\Install-Addin.ps1 -Uninstall
.\tools\Install-Addin.ps1 -Verify      # 期望 STATE=NOT_REGISTERED
```

- 只删除上表四个键，**不删除笔记、不删除其他加载项、不改其他注册表项**。
- 卸载后需**正常断开或重启 OneNote** 才会卸载已加载的 DLL；脚本不删除正在加载的文件、不结束进程。
- 本地诊断/偏好目录 `%LOCALAPPDATA%\OneNoteMarkdownPreview` 不由脚本删除；需要时手动清理。

## 4. 失败安装与回退（只影响本插件）

安装是**事务式**的：

1. 安装前对四个键做导出备份（若已存在）。
2. 任一阶段失败（注册表写入、回读校验、或其他加载项被意外改动）→ **自动回退**：删除本次创建的四个键，并从备份导入此前存在的键。
3. 失败时不会留下半套注册，也不会影响其他加载项。

隔离性验证：

```powershell
.\tools\Test-InstallRollback.ps1
```

在 `HKCU\Software\OneNoteMarkdownInstallTest` 沙箱中预置两个“其他加载项”键和一个“其他 CLSID”键，覆盖：预演不改动、安装、`-Verify`、卸载、中途失败回退、以及“已存在键在失败后被原样恢复”。断言全部通过且沙箱最终删除；**不触碰真实注册**。

## 5. 已知限制

- 未签名、HKCU 单用户；没有 MSI/安装包。
- 不提供自动加载（关卡未授权）。
- 商店版 OneNote 永远不会加载本插件。
- 本步骤未覆盖“宿主更新后注册项失效”等场景。
