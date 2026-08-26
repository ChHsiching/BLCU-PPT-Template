---
name: blcu-ppt
description: Generate BLCU-template presentations from user material — use when the user supplies documents/images and asks for a presentation (组会汇报) on the BLCU template, or asks to revise or extend deliverables this skill produced (outline.md / deck.json / 演讲稿.md / out/ 产物).
---

# blcu-ppt

素材 + 一句提示词 → BLCU 模板保真的组会汇报产物：PPTX、（可选）web 放映工程、独立演讲稿。流程是编号步骤，每步以门/检查点收口。

术语（素材/页型/deck.json/大纲/演讲稿/检查点/门/AI 味门/容量预算/渲染器/样张）的含义以 `CONTEXT.md` 词汇表为准。本仓库的 Python 脚本在仓库根目录执行，演示项目路径作为参数传入；npm 命令在对应 web 工程目录内执行。

## 契约

**输入（素材，首次调用一次性提供）**
- 文档：md / txt 直读；docx 先转 markdown（convert-documents-to-markdown skill）。
- 图片：放演示项目 `material/images/`，文件命名本身说明图的内容。
- 提示词：用途（默认组会汇报）+ 期望规模或时长（可缺省）。

**输出（全部写入演示项目目录）**
- `素材要点.md`（步骤 1 的工作记录：要点与数字出处，供 O1 溯源与 G3 审查）
- `outline.md`（大纲）→ `deck.json`（单一真相源）
- `out/<名>.pptx`；输出模式含 web 时加 `out/web/`（`npm run dev` 放映；导出 pptx 复用同一条 render_pptx CLI）
- `演讲稿.md`（独立交付物）

**依赖与降级**（Phase 0 逐项探测）
- `templates/blcu-report/blcu-report.pptx`（本地手动放置，gitignored）：缺失 = 硬停并给放置指引，用替代品糊弄属于事故。
- Python：python-pptx、latex2mathml、matplotlib；web 模式的 QA 门另需 playwright + Chromium。
- Node + npm：仅 web 模式需要。
- no-ai-slop skill：可选；AI 味门的选路与降级见 `references/ai-slop-patterns.md`。

**不做什么**
- TTS/口播音频、视频渲染、PDF 导出；代码/伪代码页型；PPT 内动画。
- 演讲内容写入 pptx 备注——演讲稿永远是独立文件。
- AI 生成素材配图；示意图仅限自绘且学术严谨优先（画不准就不画）。
- 向 git 提交任何 pptx（生成物留在 `out/`，gitignore 兜底）。

## 执行纪律

1. **门是完成判据**：O1 / G1 / G2 / G3 未达零缺陷，停在当前步骤。
2. **fail loop 整门重跑**：修完全部 findings 后重跑整道门；只查修补项不算过门。
3. **审查者是 fresh subagent**：生产者不审自己的产出；派遣提示词以 `references/reviewer-contract.md` 为准（派遣能力不可用时按其降级条款执行）。
4. **按证据判成败**：门的结论是命令输出或审查报告原文，退出码 0 / No findings 才算过。
5. **按规模伸缩**：≤10 页的 deck 每道门单遍判定（机器跑一遍、审查派发一次），clean 即放行——豁免的是循环，不是 fresh；更大 deck 的样张覆盖规则见步骤 4。
6. **样张先行**：全稿渲染前必过样张锚点（步骤 4），在 2-3 页上便宜地纠偏。

## 流程

每步以**完成判据**收口。渲染产物（pptx / web 工程）统一落演示项目的 `out/`，首次用 render_pptx 前创建该目录（render_pptx 不代建）。已有产物的修改请求按影响面进入对应步骤：改文案/换图 → 从步骤 6 的 deck.json 起步重过 G2+G3；改结构/页数 → 从步骤 3 起步。

### 0. Phase 0 · 依赖探测

逐项探测契约头的依赖与降级清单；探测结论带进 CP1 呈报。

完成判据：每项依赖有「可用」或「已定降级路径」结论；模板原件缺失即停，输出放置指引后结束本次调用。

### 1. ingest 素材

文档读入（docx 先转）；图片清点，图片语义 = 文件名。素材没有但页面需要的示意图，标记「可自绘」（学术严谨优先）。

完成判据：`素材要点.md` 成形——（a）素材要点清单，含全部数字与其在素材中的出处；（b）图片清单，每图一句语义；素材缺口（缺什么、自绘与否）有明确记录并带进大纲。

### 2. 大纲 + O1 + AI 味门

按 `references/outline-format.md` 写 `outline.md`（其头部列出页型语义与预算真相源的必读指针）。

- **O1 内联自查**：outline-format.md 的自查清单逐项过。
- **AI 味门**：选路与规程见 `references/ai-slop-patterns.md`。

完成判据：outline.md 按 outline-format 完整成形；O1 清单全绿；AI 味门零 finding。

### 3. CP1 · 用户确认

brief 式逐项决策，决策项清单与改动处理规则见 `references/outline-format.md`。

完成判据：每个决策项有用户明确答复，回写 outline.md 头部固化。

### 4. 样张 + G1

选 2-3 关键页（必含 cover 与最满的内容页；有公式必含公式页；>10 页尽量覆盖每种内容页型）：写样张 deck.json，放项目根与全稿同目录（图片相对路径同规则解析；格式见 `references/deck-json-schema.md`）→ `validate_deck` 零 finding → `render_pptx` 出样张（输出模式含 web 时 `scaffold_web` 出 web 样张到 `out/web-sample/`）。

**G1** = 机器门（`qa_check_pptx <pptx> --deck <deck.json> --com-screenshots <目录>`，COM 缺失自动降级为 note；web 样张先 `npm install`，再 `qa_check_web <web目录> --screenshots <目录>`）+ fresh subagent 审查（reviewer-contract.md）。fail loop 按纪律 2。

完成判据：机器门退出码 0；审查者 No findings；样张路径就绪供 CP2。

### 5. CP2 · 样张验收

用户逐页验收样张的视觉气质、容量、节奏。

完成判据：用户对每张样张明确认可；改动意见已落地并回走——改大纲回步骤 2（重跑 O1+AI 味门），改样张重跑 G1。

### 6. 全稿 + G2 + G3

1. 全稿 `deck.json`（deck-json-schema.md）→ `validate_deck` 零 finding。
2. 渲染所选产物；**G2 机器门**：`qa_check_pptx <pptx> --deck <deck.json> --com-screenshots <目录>`；web 为 `npm install` 后 `qa_check_web <web目录> --screenshots <目录> --export-pptx <out>`（全页截图，顺带验证导出链）。web 工程持有 deck.json 副本：deck.json 有变（含 fail loop 修复）即先 `scaffold_web --force` 重建并重新 `npm install`，再过 web 门。fail loop。
3. 按下方演讲稿规范写 `演讲稿.md`。
4. **G3 终审**：fresh subagent 审查全稿（派遣提示词与输入范围以 reviewer-contract.md 为准）。fail loop。

完成判据：G2 所选产物对应的 QA 脚本退出码 0；G3 No findings；按输出模式全部产物就位。

### 7. CP3 · 交付

交付清单，每项给出路径或可直接执行的命令：
- pptx：`out/<名>.pptx`
- web（如有）：放映 `cd out/web && npm run dev`；导出 `python scripts/render_pptx.py <项目>/deck.json -o <out.pptx>`
- 演讲稿：`演讲稿.md`

完成判据：用户收到完整清单；每项产物存在、命令原样可执行。

## 演讲稿规范

- 独立文件 `演讲稿.md`，与 PPTX / web 产物并列交付。
- 结构：头部一行元数据（汇报人 · 日期 · 预估时长；汇报人/日期取 CP1 答复与封面一致，时长取提示词的期望时长、未给则按页数与内容密度自估）；以页为节、按放映顺序组织；末尾可附「备问」小节（预期提问与答法，不进正稿）。
- 内容从大纲与素材生成：页面留结论，讲稿补细节与数字的来龙去脉；每个数字可溯源到素材，报数与页面一致。
- 质量随 G3 把关：AI 味引证与事实核对同样适用于讲稿。
