## ADDED Requirements

### Requirement: Restricted and identifiable selection
系统 SHALL 仅处理单个 OneNote 主窗口中、单一正文文本容器内连续的纯文本选区。块级转换 MUST 确认完整段落边界。无法确认范围、混合对象、跨容器、跨单元格及非正文输入 SHALL 拒绝处理。

#### Scenario: Multiple windows
- **WHEN** 触发或提交前检查发现多个 OneNote 主窗口
- **THEN** 拒绝或取消本次操作，不发送粘贴

#### Scenario: Mixed object selection
- **WHEN** 选区包含图片、附件、墨迹或不能可靠识别的对象
- **THEN** 即使复制结果包含纯文本，也不执行替换

#### Scenario: Partial paragraph with block syntax
- **WHEN** 检测到块级语法但不能确认完整段落边界
- **THEN** 整次操作跳过并显示范围不支持

#### Scenario: Formatted content surrounding the selection
- **WHEN** 选区自身为纯文本，但同一段落选区之外存在加粗、斜体或删除线等行内格式
- **THEN** 仅要求选区文本为纯文本；选区外的格式 MUST NOT 作为拒绝理由，且 MUST 保持不变

### Requirement: Revalidate edit context before commit
系统 SHALL 在提交前复核窗口、页面、正文目标与选区位置和内容证据。证据不足或检测到变化 SHALL 取消；MUST NOT 通过强行导航、恢复旧选区或抢焦点继续操作。

#### Scenario: Same text at a different position
- **WHEN** 用户将选区移到同页另一处相同文字
- **THEN** 系统不能仅凭文字相同提交，位置改变或无法证明一致时取消

#### Scenario: User edits during conversion
- **WHEN** 转换期间发生可检测的相关编辑环境变化
- **THEN** 不粘贴转换结果并提示重新触发

### Requirement: Preserve clipboard without overwriting newer data
系统 SHALL 在首次改写剪贴板前保存受支持格式的独立快照；无法可靠保存全部现有需保留格式时 SHALL 跳过。系统 SHALL 在单次粘贴提交后默认异步等待 200 毫秒，可配置为 500 毫秒，再尝试恢复。恢复 MUST 在锁内确认所有者与序列仍匹配本次载荷，MUST NOT 覆盖后来写入的内容。系统 MUST NOT 将等待结束或按键提交作为已确认粘贴成功的证据；用户接受经验延迟的提前恢复残余风险。

保留范围 SHALL 以常用用户内容为目标，包括经验证的文本、HTML/RTF 和图片表示，不保证所有应用专属粘贴能力。系统/OLE 辅助元数据 SHALL 通过有依据的格式策略逐类处理，MUST NOT 将其所有权句柄或进程私有字节直接重放。未知格式、文件/虚拟文件及尚未验证的保护标记 SHALL 拒绝，不得笼统忽略私有格式。只有建立分类依据并验证不会丢失承诺内容或改变内容保护含义的辅助格式才可省略/重建。

#### Scenario: Unsupported original clipboard
- **WHEN** 原剪贴板包含未知/未批准格式，或承诺保留的内容保存超时/超限/失败
- **THEN** 保留原剪贴板并停止，不修改笔记

#### Scenario: Classified auxiliary metadata
- **WHEN** 原剪贴板同时包含支持的用户内容与已分类为可省略的应用内元数据格式（首版：`Chromium internal source RFH token`、`Chromium internal source URL`）
- **THEN** 用户内容照常快照并恢复；该类元数据不保存也不恢复，不因此拒绝操作；未分类的未知格式仍按上述规则拒绝

#### Scenario: Unverified protection metadata
- **WHEN** 剪贴板含有未验证保护语义的标记，例如尚未分类处理的 EnterpriseDataProtectionId
- **THEN** 不删除该标记后继续复制，保持原剪贴板并跳过

#### Scenario: Third-party copy during operation
- **WHEN** 用户或其他程序在操作期间写入新的剪贴板内容
- **THEN** 系统不以旧快照覆盖新内容；尚未提交时取消本次粘贴

#### Scenario: Normal restoration
- **WHEN** 提交粘贴后的配置等待时间已到，且恢复锁内所有者与序列仍匹配
- **THEN** 尝试恢复原快照并报告恢复结果，不报告已验证粘贴成功

#### Scenario: Slower compatibility setting
- **WHEN** 用户将恢复等待配置为 500 毫秒
- **THEN** 按该等待时间安排恢复，等待期间不持有剪贴板锁

#### Scenario: Invalid delay configuration
- **WHEN** 恢复等待配置不是 200 或 500
- **THEN** 使用默认 200 毫秒并提供不含正文的配置诊断

### Requirement: Single native replacement and undo
系统 SHALL 以一次原生粘贴作为唯一正文修改动作，不预先删除或剪切。受支持场景 MUST 在实际插件路径验证一次 Ctrl+Z 恢复原文、原格式和段落结构，且选区外内容不变。

#### Scenario: Successful replacement and undo
- **WHEN** 支持的选区通过所有检查并完成格式化后，用户立即按一次 Ctrl+Z
- **THEN** 原始 Markdown、格式和段落结构恢复，选区外内容保持不变

#### Scenario: Failure before submission
- **WHEN** 解析、校验、环境复核或其他步骤在粘贴提交前失败
- **THEN** 插件不发送任何修改正文的命令

#### Scenario: Uncertain result after submission
- **WHEN** 粘贴已提交但结果无法确认
- **THEN** 停止追加编辑命令，报告结果不确定，不自动重试或撤销，也不宣称原文未变

### Requirement: Evidence-gated implementation
系统安装和发布 SHALL 以精确选区识别、实际插件粘贴撤销、经用户接受的 200/500 毫秒剪贴板恢复及所有权保护、快捷键作用范围验证通过为前提。用户已批准在关卡全部通过前开展离线实现/构建/自动测试；MUST NOT 将此批准视为自动注册、加载或修改真实笔记的授权。未通过的集成关卡 MUST 保持阻塞发布，不静默放宽要求。

#### Scenario: Ownership protection fails validation
- **WHEN** 延迟恢复测试发现覆盖第三方新复制内容或未保存支持范围内的原格式
- **THEN** 暂停放行并修复，不以用户接受经验延迟为理由放宽快照或所有权保护
