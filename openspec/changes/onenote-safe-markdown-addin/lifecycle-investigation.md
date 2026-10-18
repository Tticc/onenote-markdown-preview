# 第二次宿主启动故障：离线调查结论

## 已证实

- 实际预览断开后，预览 dllhost 已退出，但原 OneNote 进程无窗口仍存在；随后重新打开出现清理/启动失败。
- 删除插件注册并在用户授权后结束残留 OneNote，用户确认可打开、页面完整。
- 原代码持有 IRibbonUI，专用 STA 会同步调用 InvalidateControl；关闭信号仅队列投递，启动竞态可能丢失；关闭后迟到任务可能无法结束 Busy。
- 上述代码隐患已离线修正：不保留 Ribbon 对象、不跨线程回调 Office 刷新、立即取消、STA 本地计时退出及迟到任务完成清理。原生钩子只在自己线程卸载，不杀代理或宿主。
- 最新 142 项真实 Addin 本地生命周期检查通过；真实 Addin 跨进程官方接口连接、功能区 XML、断开、工作 STA 退出与客户端引用释放通过。
- 跨进程测试运行于受控可执行文件，不是 OneNote，不使用真实 Office 应用对象，不注册 COM、不安装钩子、不操作剪贴板。已有日志中没有找到能确认根因的相关 Application 错误事件。

## 尚不能证明

- 不能断言 IRibbonUI 是宿主残留/启动错误的唯一根因。没有事发时 wait-chain 或进程转储。
- 离线“代理可退出”不能证明真实 OneNote 会退出：用户配置、其他插件、同步、Office COM 服务行为都可能影响。
- 不能仅因后台进程 Responding=True 或预览 dllhost 消失就允许继续打开第二次宿主实例。
- 不能因为用户看到窗口关闭，就断言完整进程退出。必须检查无 ONENOTE 进程。

## 当前防线

- 注册脚本保持封锁；-Remove 与 -CheckOnly 可用。
- 不再在运行中的 Office 上覆写 DLL 或自动试验；不改 OneNote 缓存/用户配置来掩盖生命周期故障。
- 在没有插件注册时建立一次正常退出基线。仅当用户需要正常关闭时检查，不自动强制关闭、不调用 COM Quit（OneNote 也无通用此类退出接口）。
- 若无插件时依然留下进程，先调查 Office/COM 基线，不再通过重复安装来猜。
- 若基线干净，下一次加载也必须明确审核：只做生命周期，无正文转换/快捷键启用，和基线对照；失败即回退。

## UI 能力变化待审核

临时移除功能区实时状态刷新，结果写本地 `%LOCALAPPDATA%\OneNoteMarkdownPreview\status.txt`。这不是最终完成的反馈交互，也不是已获批准的永久需求变更。未来可用插件自身非模态窗口显示结果，避免持有宿主 Ribbon，但需单独实现/验证。不能为了恢复标签立即加回危险跨线程 COM 引用。

## 所需下一次用户操作（一次）

请先完成正常笔记工作及同步，然后在无预览插件注册状态下正常退出 OneNote。运行只读 tools/Get-RecoveryStatus.ps1 检查完全退出。基线操作不需要加载插件或重复格式化样例。

### 基线结果

用户正常关闭后，立即与间隔 10 秒两次检查均为：ONENOTE 0、预览代理 0、注册不存在。无插件完整退出基线通过。它支持把下一次试验限定为插件生命周期对照，但不能独立证明原事故根因已修复。

### 下一受控试验审核点

仅在用户明确同意后解锁一次生命周期试验注册：默认强制功能禁用（不使用此前 enabled 文件）、LoadBehavior=0，手动加载后不启用转换/键盘钩子、不访问笔记或剪贴板；手动断开并关闭，检查两个进程均退出。任何残留立即回退，不启动第二个 OneNote、不进入格式化验收。该试验应使用可识别的生命周期诊断状态，不复用能自动启用的旧预览注册安排。

### 受控试验结果：未通过，已回退

用户批准后实际试验：插件 STA 与 dllhost 都退出，但 OneNote PID 9760 无窗口仍留存，十秒复查不变。没有启用解析/快捷键/剪贴板功能，因此不能再把故障归因于这些工作。注册已全部移除，生命周期专用注册也重新封锁。

只读 WCT 观察 48 个线程未报告等待环；没有 COM 专用等待回调或堆栈，不能据此断言无卡住、无引用泄漏。没有强制结束宿主。不得再次让用户带着残留进程重新启动，也不重复同样加载试验。

### 授权恢复及原生引用实验

用户随后授权结束 PID 9760，执行后确认宿主/代理均为 0、注册不存在。未自行重启。

新增 NativeReferenceTests：手工创建独立的原生 IDispatch vtable 与 AddRef/Release 计数器，经 CLR 包装为真实 RCW，而不是用托管对象假装宿主。真实 Addin OnConnection 接收两个参数且不保存；返回时两个原生对象引用计数均为 2（fixture 自身基线为 1）。工作线程退出后，不应假定参数对应 RCW 已即时释放。显式测试 GC 后引用回到 1，弱引用失活；这证明“不存字段”并不等于确定性释放。

**不能因此认定已找到真实故障根因。** 已观察到 dllhost 整体退出，而跨进程代理可能已正常 rundown；该实验尚未覆盖实际 Office 对象/回调栈/RPC 服务端。不能粗暴对传入共享 RCW 循环 ReleaseComObject/FinalReleaseComObject，微软文档警告这样可能破坏其他用户或造成访问异常。也不在宿主回调中 GC.WaitForPendingFinalizers，以免制造新等待。

下一步需要原生堆栈或受控最小集成对照，而不是继续修复无证据的引用猜测。注册保持封锁。若要捕获转储，需先说明转储可能包含笔记内容并获得单独决策；本次未生成任何转储或读取进程内存。

### 真实退出堆栈证据已取得

用户保留现场后，自动观察器对 PID 14500 两次采集本地非侵入式线程堆栈，间隔约 4 秒，最后自动清理注册。未生成转储，未结束宿主。

两次主线程地址一致，停在 Windows 消息/多对象等待，往下是 Office 与 OneNote 本机帧；另有 OneNote ExeServer 附近工作线程等待。主线程没有直接显示插件/CLR 或可确认的跨进程 RPC 调用栈。无匹配 PDB 时导出函数名加大偏移只能当邻近符号，不能据此断言正在同步、注销 COM 或等待某个插件。

结论仅是“代理退出后宿主本机等待仍持续”。故障仍未修复，不能用结束宿主/共享 RCW 强制清零作为产品补丁。重复同一生命周期试验无益，所有注册重新封锁。后续须先恢复残留宿主，再决定更深的原生诊断或架构替代；不得自动扩大到内存转储。

参考：https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject?view=netframework-4.8.1

## OneMore 同架构对照证据（离线收集）

- OneMore 的注册样例（`register-onemore.reg`）证明它是与本项目完全相同的架构：`AppID {88AB88AB-CDFB-4C68-9C3A-F10B75A5BC61}` 含空值 `DllSurrogate`，`InprocServer32=mscoree.dll`（ThreadingModel=Both）、`LoadBehavior=3`。不是 in-proc 插件。
- OneMore 官方技术笔记《TechNote - COM Surrogate》：代理按定义运行在 MTA；直接在 MTA 线程使用 STA-only 组件（打开文件对话框、剪贴板等）会"同时挂起代理和 OneNote"；OneMore 用自建 STA 线程（SingleThreaded helper）+ 全异步回避。与本项目的专用 STA + 异步设计一致。
- OneMore `AddIn.cs` 的 `OnDisconnection` 结尾**无条件执行**：`GC.Collect();` 然后（注释原文）"this is a hack, modeless dialogs seem to keep OneNote open :-(" 后 `Environment.Exit(0);`——即断开时立即终结承载插件的 dllhost 进程。该实现存在于 2016 年至今的持续维护生产代码中。
- OneNoteDev/VanillaAddIn issue #2（微软官方示例插件）：点击插件按钮后关闭 OneNote，ONENOTE.EXE 无窗口残留，下次启动报 "OneNote is cleaning up" 并失败；杀残留进程或重启后恢复；至今未修复。说明该症状类并非本项目特有。
- 结论：宿主残留疑为 OneNote 平台级生命周期缺陷；同架构成熟插件用"断开时自毁代理进程"实际规避了它。

## 证据缺口与代码修改（离线完成）

- 缺口修正：此前试验记录"代理已退出但宿主仍等待"未在挂起时点核实过代理存在性——watcher 原本只查 ONENOTE，`Get-RecoveryStatus` 只在事后方执行。`Watch-ExitStacks.ps1` 现在在窗口关闭、PASS、留存三个时点记录预览 dllhost 数量。
- `src/Addin/Addin.cs`：`OnDisconnection` 现在在 `Shutdown()` 后调用 `ForceExitSurrogate(mode)`：仅真实宿主模式下，若操作进行中等候最多 5 秒（仅此期间会在途剪贴板恢复得以完成），写 `surrogate-exit mode=... busy=...` 生命周期日志，`GC.Collect()` 后 `Environment.Exit(0)` 终结本代理进程；绝不触碰宿主进程。离线构造（生命周期测试/探针）直接返回，`OnDisconnection` 的 500ms 回调约束保持。
- `tools/Register-Preview.ps1`：加载路径改为 `-FixTrial` 闸门（要求引用本计划）；`-Remove`/`-CheckOnly` 不变；试验仍写 `lifecycle-only` 标记（转换/快捷键锁定）。
- 上方"所有注册重新封锁"的论断由本计划取代；重复无计划的生命周期试验仍不被允许。

## 修复验证试验（已审核，待用户一次会话执行）

前提：`build.ps1 -Test` 全绿（离线）。

轮次 1（宿主关闭路径）：
1. 启动 `tools/Start-ExitStackWatch.ps1`（或等价观察器）。
2. `Register-Preview.ps1 -FixTrial`。
3. 用户打开 OneNote，手动加载预览插件（不启用转换），直接关闭 OneNote（不手动断开）。
4. 预期：ONENOTE 与预览 dllhost 均在数秒内完全退出；watcher 记 PASS 并自动移除注册；lifecycle.log 出现 `surrogate-exit mode=ext_dm_HostShutdown`。

轮次 2（用户断开路径）：重新 `Register-Preview.ps1 -FixTrial` → 打开 OneNote → 加载插件 → 手动取消勾选（断开）→ 关闭 OneNote → 预期同上；断开瞬间应出现 `surrogate-exit mode=ext_dm_UserClosed`。

判定规则：
- 双轮 PASS → 修复有效，记录证据；安装/发布仍按原门禁另行审核。
- 任一残留 → watcher 留存（含 dllhost 计数）；若日志已有 `surrogate-exit` 且留存时代理不存在，则证实"宿主自身等待、与代理无关"，须返回架构/规避决策；若代理仍存活且日志无 `surrogate-exit`，则修复未生效，回退排查。

约束：转换锁定、LoadBehavior=0、不访问笔记或剪贴板、不杀进程、失败不启动第二个实例、失败即回退。

### 修复验证试验结果：两轮 × 两会话全部通过

- 会话 A（原环境）：轮次 1 `07:26:10 pid=26000 surrogate-exit mode=ext_dm_HostShutdown idle`，watcher `PASS host process exited naturally; preview dllhost=0`；轮次 2 `07:30:09 pid=10968 surrogate-exit mode=ext_dm_UserClosed idle`，`07:30:23 PASS ... dllhost=0`。
- 会话 B（用户重启后冷启动重做）：轮次 1 `10:23:24 pid=8816 surrogate-exit mode=ext_dm_HostShutdown idle`，`10:23:27 PASS ... dllhost=0`；轮次 2 `10:24:41 pid=19856 surrogate-exit mode=ext_dm_UserClosed idle`，`10:24:46 PASS ... dllhost=0`。
- 四次均：预览 dllhost=0、注册由 watcher finally 自动移除、事后 `Get-RecoveryStatus` 全零（ONENOTE/代理/注册）。宿主关闭后约 1–3 秒内完全退出；多数轮次连"窗口关闭但进程仍在"的中间态都未进入。
- 对照：修复前同一动作残留 ≥10 秒并被捕获到宿主原生等待栈；修复后不再复现。
- 结论修正：原记录"代理已退出但宿主仍等待"在当时未核实代理存在性（watcher 只查 ONENOTE）。修复后代理在宿主关闭前退出、宿主随即退出，支持"存活代理阻塞宿主关闭"的工作解释；根因仍是 OneNote 内部生命周期，未经私有符号证实。此为可复现故障已消除的证据，不等于完整根因证明。
- 遗留：本轮仅覆盖加载/断开/关闭生命周期；转换、快捷键、剪贴板在 lifecycle-only 下未启用。剪贴板安全由有界等待（仅操作进行中 ≤5s）保障，未在真实挂起中演练；升级现有"严禁 Environment.Exit"规格见下节。
