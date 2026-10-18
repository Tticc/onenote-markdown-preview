## ADDED Requirements

### Requirement: Integrated explicit formatting entry
系统 SHALL 提供 OneNote 功能区格式化与启用/禁用入口，并支持 Ctrl+Shift+D；最近结果 SHALL 以不含正文的本地诊断记录（`status.txt`）保存，跳过或失败时以不抢焦点的轻量提示呈现原因；只有已启用且符合条件的 OneNote 编辑环境接管格式化快捷键。系统 MUST NOT 自动监听页面并格式化。

#### Scenario: Disabled shortcut
- **WHEN** 插件已禁用且用户按 Ctrl+Shift+D
- **THEN** 插件不拦截该组合键、不执行转换、不修改笔记

#### Scenario: Other application active
- **WHEN** 其他应用处于前台
- **THEN** 插件不占用该应用的 Ctrl+Shift+D 输入

### Requirement: Persistent user-controlled merged surrogate lifecycle
系统 SHALL 通过已验证的 DllSurrogate 注册将插件入口与转换模块合并于同一 COM 代理进程，不新增专用转换进程。转换 SHALL 按用户触发异步执行，不在加载或同步功能区回调中等待解析。启用状态与可检测故障停用状态 SHALL 持久化，恢复 MUST 由用户明确启用。禁用 SHALL 停止转换并释放快捷键，且 MUST NOT 结束代理进程。OnDisconnection（宿主关闭或用户卸载）时，插件 SHALL 在有限优雅等待（仅操作进行中，最多 5 秒，以保全在途剪贴板恢复）后结束自身代理进程，以解除存活代理对宿主关闭的阻塞（依据：本项目四次验证与 OneMore 同架构生产实践）。系统 MUST NOT 自动重启代理，MUST NOT 强制结束 OneNote 或其他 dllhost，MUST NOT 按进程名结束进程。

#### Scenario: Disabled across restart
- **WHEN** 用户停用后重新启动 OneNote
- **THEN** 转换保持禁用，直到用户明确启用；不以代理进程是否仍存在判断功能状态

#### Scenario: Detectable conversion fault
- **WHEN** 转换发生可捕获故障或超时且代理仍能响应
- **THEN** 未提交的结果作废，保存故障停用状态并显示原因；未结束的任务不得因用户重复启用而被叠加

#### Scenario: Abrupt surrogate failure
- **WHEN** 代理崩溃或失去响应
- **THEN** 系统不承诺即时反馈或清理成功，不自动重试文档写入；用户手动重新连接或重启后，初始化发现持久化未完成操作标记时保持故障停用

#### Scenario: Normal enabled restart
- **WHEN** 启用中的 OneNote 正常退出后再次启动
- **THEN** 启用偏好保留，通过已验证 COM 生命周期恢复入口；加载时不解析笔记或启动专用转换进程，正常宿主退出应释放本插件的 COM 引用

#### Scenario: Disconnect or host close
- **WHEN** 用户卸载插件或宿主开始关闭（OnDisconnection）
- **THEN** 插件仅在操作进行中等待至多 5 秒以完成在途剪贴板恢复，随后结束自身代理进程，不触碰宿主进程；宿主随后在数秒内完全退出

### Requirement: Serialized operations and explicit stopping
系统 SHALL 同时只处理一次转换，不排队重复请求。用户停用 SHALL 取消尚未提交的操作并释放快捷键；提交后的停用 MUST NOT 引发自动重试或自动撤销。

#### Scenario: Repeated shortcut
- **WHEN** 转换尚未结束时用户再次触发格式化
- **THEN** 不新增或排队操作

#### Scenario: Disable during conversion
- **WHEN** 用户在粘贴提交前停用
- **THEN** 本次结果不再提交到 OneNote，后续转换等待用户手动启用且之前任务已经结束；禁用本身不结束代理进程（仅断开/宿主关闭时按生命周期要求经有限优雅等待后结束自身代理）

### Requirement: Nonintrusive status and content-free diagnostics
系统 SHALL 提供不抢焦点的结果或跳过原因，不自动弹出打断式对话框。诊断日志 MUST NOT 包含笔记正文、生成 HTML 或剪贴板内容。

#### Scenario: Unsupported syntax feedback
- **WHEN** 转换因不支持的语法被拒绝
- **THEN** 结果原因以不含正文的本地记录与（跳过/失败时）不抢焦点的轻量提示呈现，编辑焦点不被提示抢走，日志不包含原文

#### Scenario: Skip or failure feedback
- **WHEN** 一次操作因语法、选区、剪贴板或错误被跳过或未能完成
- **THEN** 通过不抢焦点的轻量提示（非激活、自动消失的弹窗或系统通知）说明原因，MUST NOT 使用模态对话框，且 MUST NOT 抢走编辑焦点；提示内容 MUST NOT 包含笔记正文、生成 HTML 或剪贴板字节，只允许原因与剪贴板格式名等元数据；MUST NOT 依赖可被系统通知设置抑制的机制作为唯一通道；断开/禁用后不得遗留提示窗口
