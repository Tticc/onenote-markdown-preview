# 故障注入清单（任务 5.5 / 1.6）

对照 `src/Addin/Addin.cs` 的 `Format()` 流程与 `OperationState`，逐项验证失败路径。**核心规则**：提交前失败 → 取消（不写正文）；提交后不确定 → 明确标注、不自动撤销/重试；任何失败都保留原剪贴板或明确提示需手动恢复。

观察通道：右下角非激活弹窗 + `%LOCALAPPDATA%\OneNoteMarkdownPreview\status.txt`（保留原始原因码）。全程**只用一个宿主实例**（桌面版路径启动），不要强杀进程（第 6 项除外，且需单独授权）。

## 前置

```powershell
.\tools\Register-Preview.ps1 -Acceptance      # 宿主已退出；LoadBehavior=0
# 用桌面版路径启动：C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE
.\tools\Start-ExitStackWatch.ps1              # 记录退出/残留
```

## 已经离线证明（宿主复测可选）

| 故障 | 机制 | 证据 |
|---|---|---|
| 重复请求 | `Begin()` 单飞，第二次返回 0 | CoreTests「duplicate request rejected」 |
| 停用/故障后的迟到结果 | `Disable()/Fault()` 使 generation 失效，`Valid(ticket)=false` | CoreTests「late response after disable rejected」 |
| 故障后需手动恢复 | `Fault()` 后 `Enable()` 才可再次 `Begin()` | CoreTests「manual recovery after fault」 |
| 恢复延迟只接受 200/500 | `OperationState.Delay` 其余值→200 | CoreTests「restore delay policy」 |
| 第三方复制竞争 | 锁内核验 owner+序列，检测到变化即不覆盖 | 剪贴板协调离线测试 |
| 恢复失败 | 保留内存快照，停用，提示手动恢复 | LifecycleTests「pending marker」+ `RestoreRecovery` 路径 |

## 需要真实宿主（逐项记录弹窗/status.txt 内容）

| # | 故障注入 | 操作 | 期望 |
|---|---|---|---|
| 1 | **切页** | 长选区触发后，在粘贴前立刻切到别的页 | 提交前取消：`选区变化，已取消`（或 `操作已取消`）；正文不变；剪贴板恢复 |
| 2 | **改选区** | 触发后立刻改动选区 | 同上，`选区变化，已取消` |
| 3 | **输入** | 触发后立刻在页面输入字符 | 同上；不覆盖新内容 |
| 4 | **第二窗口** | 开第二个 OneNote 窗口后触发 | 拒绝多窗口选区；不写正文；有中文提示 |
| 5 | **转换异常/卡死** | 选超过 32768 字的正文 | 立即拒绝（`选中内容过长`），不卡死；非超时停用 |
| 6 | **代理意外退出** | 运行中强制结束该加载项的 `dllhost`（**需单独授权**） | 操作中止；无部分写入承诺；内存快照丢失→提示需重新复制；重新加载插件后可用 |
| 7 | **剪贴板竞争** | 恢复窗口（200ms）内于记事本复制新内容 | 新内容保留，不被覆盖（宿主复测） |
| 8 | **停用（进行中）** | 转换进行中点击“停用” | 取消；不写正文；钩子释放 |
| 9 | **粘贴结果不明** | 观察 `sent==4` 与否则的消息前缀 | `已提交；` vs `提交结果不确定；` 与恢复结果拼接 |

## 判定

- 第 1–5、7–9 项可在一个会话内做完（人工操作 + 自动 watcher）。
- 第 6 项（结束代理进程）**需要明确授权**，且仅结束该加载项的代理进程，不结束 OneNote 主进程。
- 完成后把每项的 `status.txt` 片段与结果记入 `validation.md`；未通过项不得标记支持。
