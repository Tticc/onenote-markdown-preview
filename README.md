# OneNote 局部 Markdown 格式化预览

> 预览状态（2026-10-18）：宿主退出残留问题已修复并多次复验通过；真实宿主功能验收（转换、一次撤销、剪贴板来源、跳过提示）已通过。仍为**手动加载预览**（LoadBehavior=0，每次启动需在 OneNote 的 COM 加载项对话框勾选）。**已于 2026-10-18 在本人机器执行正式安装**（HKCU 单用户）并验证 `STATE=REGISTERED`；`LoadBehavior=3`（自动加载）仍未授权，也不要使用旧探针脚本绕过集中验收。

Windows x64 / OneNote 桌面 COM Add-in。当前是**仅离线构建的预览代码**，尚未通过完整交付验收，不应安装到日常使用环境。构建不会注册插件、启动 OneNote 或操作剪贴板。

用户使用说明见 [`usage.md`](usage.md)。

## 离线构建与自检

在项目根目录运行 64 位 Windows PowerShell：

```powershell
.\build.ps1 -Test
```

使用本机 .NET Framework 4.x 编译器、Office 15 PIA（当前 Office 16 安装提供），目标 x64、.NET Framework 4.8 环境。不要求 .NET SDK。首次构建下载固定 NuGet 版本到 `.packages/`，结果输出到 `artifacts/preview/`。运行阶段无网络访问。Markdig 使用 BSD-2-Clause，System.* 依赖各自许可证随原 NuGet 包保留在缓存；正式分发前需整理许可证文件。

当前自检含真实解析器/状态测试和官方 COM/Ribbon 契约审计，不触发 OnConnection，不安装或加载宿主。旧验证工具入口为 `validation/OneNoteProbe/Run-AutomatedValidation.ps1`；不要将旧固定载荷探针误认为正式 Markdown 转换器。

## 预览支持边界

- `#`～`###` ATX 标题；`**粗体**`、`*斜体*`、`~~删除线~~`；行内代码等宽加 #eeeeee 底色。
- `- ` 基础列表、从 `1.` 开始的基础有序列表；任务列表显示 ☐/☑；基础引用。
- 不支持围栏/缩进代码块、链接、图片、表格、原始 HTML、嵌套列表、扩展语法；遇到未支持项整段跳过。
- 不支持空选区、标题、搜索框、表格、混合对象、跨容器、复杂原生样式、快速笔记/停靠视图、多窗口。正文内部默认 span 包装可解析，不能把未验证复杂富文本当普通文字。
- 输入上限 32768 UTF-16 字符、节点 4096、渲染深度 16、输出 256 KiB；异步解析 2 秒等待预算。解析线程超时不强杀，未结束前拒绝重新开始。

## 设计行为

插件由 COM surrogate 承载，转换合并在 dllhost 中，无额外转换 EXE。入口和剪贴板工作运行专用 STA 消息循环，解析运行后台任务。默认禁用；启停偏好和不含正文的未完成操作标记保存在 `%LOCALAPPDATA%\OneNoteMarkdownPreview`。

快捷键 Ctrl+Shift+D；按钮触发不强行恢复正文焦点，若焦点已移走会提示改用快捷键。功能区反馈不弹窗。禁用释放键盘钩子；不自动撤销/重试，不强行结束 OneNote/dllhost。

默认等待 200 ms 恢复剪贴板，可在上述配置目录 `restore-delay-ms.txt` 写入 `500`。其他值使用 200。经验等待不保证宿主已经完成粘贴。恢复前检查所有者和序列；新内容不覆盖。快照覆盖常用文本/HTML/RTF/PNG/DIB，条件省略 OneNote 专属格式和带位图后备的图形表示；未知格式、EnterpriseDataProtectionId 等未支持标记仍跳过。

恢复失败时保留内存快照，停用转换，用户可点击“重试恢复剪贴板”。代理崩溃/退出会使内存快照丢失，不承诺崩溃持久恢复。不将正文/快照写磁盘。

## 安装与发布状态

**已执行正式安装（2026-10-18）。** 正式安装路径为 `tools/Install-Addin.ps1`（HKCU 单用户，`LoadBehavior=0`，宿主必须已退出），当前 `-Verify` 为 `STATE=REGISTERED`；卸载用同脚本 `-Uninstall`。受控手动加载脚本 `tools/Register-Preview.ps1` 仅用于集中验收，回退为同脚本 `-Remove`，预检为 `-CheckOnly`。步骤见 `acceptance/preview-session.md`。早期探针曾引发退出异常，已卸载并恢复；不得使用旧 Register-Probe.ps1/手动修改 LoadBehavior 绕过集中验收。

正式安装路径（任务 6.3）：`tools/Install-Addin.ps1`（`-DryRun` / `-Verify` / `-Uninstall`，事务式回退），步骤见 `acceptance/install.md`；回退隔离性由 `tools/Test-InstallRollback.ps1` 在 HKCU 沙箱中验证。

已完成：宿主退出残留修复（多次复验通过）、真实宿主转换与一次撤销、选区严格解析修复、剪贴板来源分类与友好提示、非激活跳过提示、启动开销采样。

仍未闭环：安装/卸载/回退脚本（6.3）、故障注入与恢复失败持久化（1.6/5.5）、`worker-exited` 幂等修复的真机复验、外部延迟渲染来源（`gate-review.md:20`）、空闲 CPU/内存与转换耗时。通过离线测试不等于可安全发布。

一次性验收计划：`openspec/changes/onenote-safe-markdown-addin/manual-acceptance.md`。

## 开源发布、打包与签名

- 许可证 **MIT**（`LICENSE`）；第三方组件（Markdig BSD-2-Clause、System.* MIT）见 `THIRD-PARTY-NOTICES.md`，其许可证文件随 `bin/` 分发。
- 版本号来自根目录 `VERSION`（当前 `0.1.0`），`build.ps1` 通过生成的 `Version.g.cs` 写入 `AssemblyVersion`/`AssemblyFileVersion`。
- 打包：

```powershell
.\tools\Package-Release.ps1                 # 构建 + 校验完整性 + 打包
.\tools\Package-Release.ps1 -SkipBuild      # 复用已有构建
.\tools\Package-Release.ps1 -Thumbprint <cert>   # 可选 Authenticode 签名
```

产物在 `artifacts/release/OneNoteMarkdown-<version>-x64.zip`，并生成 `.zip.sha256`。归档含 `bin/`（DLL 与许可证）、`tools/`（安装/预览脚本）、`docs/`、`usage.md`、`LICENSE`、`THIRD-PARTY-NOTICES.md`、`VERSION-INFO.txt` 与 `SHA256SUMS.txt`。

- **签名**为可选：提供 `-Thumbprint` 时用 Windows SDK 的 `signtool` 对每个 DLL 做 Authenticode 签名并校验；没有证书时发布包**不签名**（开源项目常见做法），以 SHA-256 校验和作为完整性凭据。
- 仍需注意：Microsoft Office PIA（`Extensibility.dll`/`office.dll`/`Microsoft.Office.Interop.OneNote.dll`）**不随包分发**，构建时从本机 GAC 引用。
