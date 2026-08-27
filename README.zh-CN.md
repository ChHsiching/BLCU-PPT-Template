# BLCU-PPT-Template

一个 [agent skill](SKILL.md)：用户投入素材（文档 + 图片）与一句提示词，产出忠实继承 BLCU 汇报模板的可演示成果——PPTX 文件、（可选）web 放映工程、独立演讲稿。服务计算机研究生组会汇报场景；为将来接入多模板设计。

[English (README.md)](README.md)

## 产出什么

| 输入 | 输出 |
|---|---|
| `material/` — 一份 markdown/docx 文档 + 图片 | `outline.md` → `deck.json`（单一真相源） |
| 一句提示词：「用它做一份组会汇报」 | `out/<名>.pptx` —— 原生 OMML 公式、内嵌字体 |
| 可选：web 输出模式 | `out/web/` —— Vite+React 放映工程，`npm run dev` 即放映 |
| 总是产出 | `演讲稿.md` —— 独立的口头讲稿文档 |

## 工作方式

skill 驱动 agent 走编号步骤，每步以**门**收口（零缺陷完成判据）：素材 ingest → 大纲（含 AI 味门）→ CP1 用户确认 → 样张 → CP2 验收 → 全稿 + 机器 QA + 终审 → 交付。两个渲染器消费同一份 `deck.json`：**renderer-pptx** 在模板原件上 Clone & Fill，**renderer-web** 是 Vite + React 工程。机器门复核一切交付物：容量预算、占位残留、标题、几何、样式 token、品牌层，web 侧还有绿条像素级抽查。

## 仓库结构

```
SKILL.md                  skill 本体：契约、执行纪律、8 步流程
CONTEXT.md                领域词汇表（素材/页型/门/容量预算 …）
references/               大纲格式 · deck.json schema · 页型语义 ·
                          审查者契约 · AI 味 pattern 清单
scripts/                  render_pptx · scaffold_web · validate_deck ·
                          qa_check_pptx · qa_check_web · embed_fonts …
templates/blcu-report/    模板衍生物：manifest.json（几何/预算/样式 token
                          的唯一真相源）、提取的品牌媒体
assets/web-template/      预置 web 放映脚手架，scaffold 进演示项目
fonts/                    Noto Sans SC 400/700 TTF + OFL 许可证（子集内嵌进每份 pptx）
examples/qat-lsq-repro/   端到端示例：素材 → 大纲 → deck.json → pptx + web + 演讲稿
tests/                    scripts 与门的 pytest 测试套件
```

## 环境要求

- Python 3.11+：`python-pptx`、`latex2mathml`、`matplotlib`；web QA 门另需 `playwright`（+ Chromium），可选 PIL。
- Node 24 + npm：仅 web 输出模式需要。
- PowerPoint（Windows + COM）：仅 COM 截图抽查用；缺失时门自动降级。
- `templates/blcu-report/blcu-report.pptx` **不在仓库里**（有意 gitignored）：请手动放置模板原件的副本。其余全部随仓库分发。

## 快速上手

```bash
# 跑测试套件
python -m pytest tests/

# 再生示例的三产物（需先放置模板原件）
python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx
python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force

# 之后即可拿自己的素材调用 skill 本体
```

使用 skill = 把素材和提示词交给 agent——流程、门、检查点见 [SKILL.md](SKILL.md)。

## 设计教义

- **模板 = 母版**：模板母版不可动；母版之外的一切（字体字重、层级、强调色、间距节奏）均自行设计，由 `manifest.json` 的 token 单一真相源驱动。
- 层级靠字重 + 颜色表达（唯一强调绿），不新增字号差；正文保持纯黑以扛投影。
- 公式保持原生 OMML 可编辑（LaTeX → MathML → OMML）；字体按实际用字子集化并内嵌（<1 MB），任何机器打开都是设计时观感。
- 门是证据不是感觉：机器检查 + fresh subagent 审查循环至零缺陷。

## 状态

v1 完成：skill 流水线（T1–T7）与视觉样式轮（S1–S6）均已落地，并在随仓示例上端到端验证。历史见 [Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues)。
