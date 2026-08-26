# BLCU-PPT-Template

一个 agent skill：用户投入素材（文档 + 图片）与提示词，产出使用 BLCU 汇报模板的可演示成果——PPTX 文件、web ppt 程序、独立演讲稿。服务计算机研究生组会汇报场景。

## Language

### 素材与产物

**素材（Material）**:
用户首次调用时提供的输入：文档（md/txt/docx）与图片。图片放在 `material/images/`，文件命名本身说明图片内容。
_Avoid_: 资源、源文件、输入文件

**模板（Template）**:
一个 `.pptx` 原件及其衍生物（manifest、提取的品牌资产）的集合，住在 `templates/<id>/`。当前唯一模板为 `blcu-report`。完整 pptx 原件永不入库（gitignored），衍生物入库。
_Avoid_: 主题（theme）、皮肤

**框架（Frame）**:
模板中固定照搬的部分：母版品牌层（绿条、logo、页码、封面版式）及绿条内标题样式。样式决策只到"保留"，不做再设计。
_Avoid_: 母版（那是 OOXML 部件名）、模板整体

**体裁（Genre）**:
从模板学到的"汇报长什么样"：页型构成（方法靠公式、实验靠图）、密度水平、结构惯例。只学内容与形式，不学具体样式与坐标。
_Avoid_: 风格、格式

**内容形式（Content form）**:
框架之外内容区的全部样式决策域：字体字号字重、栅格间距、颜色、图表规范、图文排布。归本项目自行设计，永不从模板作者的习惯继承。
_Avoid_: 排版（口语泛称）、布局（仅指几何摆放）

**页型（Archetype）**:
一类可克隆页面形态的布局规范（cover / agenda / text-formula / text-image / chart-focus / closing），是布局知识的单一真相源。deck.json 只引用页型，不携带坐标。
_Avoid_: 版式（layout，指 OOXML 的 slideLayout）、页面类型

**deck.json**:
演示内容的中间表示与单一真相源：页面序列 + 每页的语义块（标题、文本、公式、图片引用）。两个渲染器都只消费它。
_Avoid_: 大纲文件（那是 outline.md）、中间格式

**大纲（outline.md）**:
CP1 之前产出的规划文档：页面序列、每页页型与要点、容量预算分配、配图分配。经用户确认后固化为 deck.json。
_Avoid_: 计划书、目录

**演讲稿（演讲稿.md）**:
独立交付的口头讲稿文档，与 PPTX/Web 产物并列。永不写入 pptx 的演讲者备注。
_Avoid_: 备注、notes、旁白

### 流程与质量

**检查点（Checkpoint / CP1-CP3）**:
流程的硬停点，用户在此逐项决策（CP1 大纲与输出模式、CP2 样张验收、CP3 交付）。以 brief 式决策摘要呈现，决策项独立成题。
_Avoid_: 确认点、审批

**门（Gate / O1、G1-G3）**:
一道以零缺陷为通过标准的完成判据（机器检查、内联自查或 fresh subagent 审查）。不过门不推进；fail loop 要求修复全部缺陷后整门重跑，禁止抽查修补项。门不是建议。
_Avoid_: 检查步骤、review

**AI 味门（AI-flavor gate）**:
对文本内容的 AI 腔审查：优先调用已安装的 no-ai-slop skill，缺失或无法调用时降级到本仓库 `references/ai-slop-patterns.md` 内联清单。发现以具名 pattern + 原文引证 + 一句修法呈现。
_Avoid_: 文风检查

**容量预算（Budget）**:
每页型对文本量的可数上限（要点条数、每条字数），大纲与 deck.json 阶段都要合规，是防溢出的主要手段。
_Avoid_: 字数限制（不完整——还包括条数与图数）

**渲染器（Renderer）**:
把 deck.json 落成产物的确定性程序：renderer-pptx（python-pptx，以模板原件为基底的 Clone & Fill，公式走 LaTeX→OMML 原生注入）与 renderer-web（Vite + React，公式走 KaTeX）。
_Avoid_: 导出器、生成器

**样张（Sample pages）**:
CP1 之后、全稿之前渲染的 2-3 页关键页真模板产物，是最便宜的纠偏锚点。
_Avoid_: demo、预览
