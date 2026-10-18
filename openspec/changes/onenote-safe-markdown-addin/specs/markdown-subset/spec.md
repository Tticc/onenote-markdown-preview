## ADDED Requirements

### Requirement: Minimal supported Markdown rendering
系统 SHALL 支持 # 至 ### 标题、**粗体**、*斜体*、~~删除线~~、行内代码、基础无序/有序列表、任务列表和引用，以及其间的普通文字。围栏/缩进代码块**不在本版本范围内（已知限制，计划后续版本支持）**，因其严格空白保真尚未验证。标题 SHALL 呈现视觉层级，任务项 SHALL 呈现 ☐/☑，行内代码 SHALL 使用等宽显示；不要求原生标题样式或可交互待办标签。

行内代码 SHALL 采用等宽字体和浅灰底色，底色限于代码范围；MUST NOT 为实现该样式转换成图片或改变同段其他文字的底色。

#### Scenario: Inline code visual distinction
- **WHEN** 选区中含有受支持的行内代码
- **THEN** 移除定界反引号，代码文字呈现等宽字体与浅灰底色，同段非代码文字保持原有显示

#### Scenario: Supported mixed content
- **WHEN** 完整段落选区只包含已定义子集内的标题、列表和行内样式
- **THEN** 返回保留正文与顺序的富文本结果，不要求 OneNote 原生标签语义

#### Scenario: Inline code content preservation
- **WHEN** 行内代码包含中文、emoji 和看似 Markdown 的符号
- **THEN** 行内代码正文保持，内部符号不作为外层 Markdown 转换

#### Scenario: Code block is out of scope (known limitation)
- **WHEN** 选区含围栏或缩进代码块
- **THEN** 整次转换跳过并显示代码块不在本版本范围；不用 NBSP、全角空格或图片替代原始字符；与其它语法混合时整段跳过，不部分转换

### Requirement: Reject unsupported syntax for the whole selection
系统 SHALL 使用明确的子集语法与上下文识别未支持结构、确定不完整的结构及资源超限；发现任一情况 SHALL 整段拒绝并给出类别，不返回可提交的部分结果。普通标点、合法转义和代码内容 MUST NOT 仅因含有 Markdown 符号而被拒绝。

#### Scenario: Supported and unsupported syntax mixed
- **WHEN** 选区含粗体以及未支持的表格或图片语法
- **THEN** 整段不转换，返回未支持语法类别，不只转换粗体部分

#### Scenario: No formatting needed
- **WHEN** 选区仅含普通文本且没有可转换节点
- **THEN** 返回无需转换状态，不执行粘贴

### Requirement: Validated bounded rich-text output
系统 SHALL 校验 AST、标签与属性白名单、内容保真、结构及资源限额；原始 HTML MUST NOT 直接透传，外部资源 MUST NOT 加载。CF_HTML SHALL 使用 UTF-8 字节偏移和合法片段边界；代理内操作协调模块 SHALL 拒绝异常、超限或操作标识不匹配的转换结果。

#### Scenario: Unicode clipboard payload
- **WHEN** 富文本包含中文或多字节 Unicode 字符
- **THEN** CF_HTML 偏移按 UTF-8 字节计算，片段可完整解码且文本不截断

#### Scenario: Invalid conversion result
- **WHEN** 合并转换模块返回空缺、畸形、超限或操作标识不匹配的结果
- **THEN** 操作协调模块不发布或粘贴该结果，报告失败且不记录正文
