<h1 align="center">BLCU-PPT-Template</h1>

<p align="center">
  <strong>素材 + 一句提示词，生成继承 BLCU 汇报模板的组会演示：PPTX、Web 放映、演讲稿</strong>
</p>

<p align="center"><a href="README.md">English</a> · <b>简体中文</b></p>

---

一个面向研究生组会汇报的 [agent skill](SKILL.md)。把文档、图片和一句提示词交给 agent，skill 会依次执行大纲规划、样张确认、带门渲染，产出基于 BLCU 模板原件的演示成品。

流程为线性步骤：素材整理、含容量预算的大纲、大纲确认（CP1）、机器检查与审查通过的样张（CP2）、全稿渲染。两个渲染器消费同一份中间文件：renderer-pptx 在 XML 层克隆模板 pptx 的页面，renderer-web 从同一份 deck.json 生成 Vite + React 放映工程。每份交付物都要通过脚本化质量门（文字预算、占位残留、几何、run 级字排、品牌层；web 侧另含绿条像素抽查）；审查门由 fresh subagent 执行，循环至零缺陷。

## 成品为何与模板一致

- 页面从模板原件克隆而来，顶部绿条、logo、页码就是模板母版自带的形状，不是重新绘制的近似物。
- 母版之外的全部样式收敛在一张 token 表里
  （[`templates/blcu-report/manifest.json`](templates/blcu-report/manifest.json)）：字族、字重、字号、颜色、间距。两个渲染器与两个 QA 脚本读取同一份数据。
- 公式从 LaTeX 转为原生 OMML，在 PowerPoint 中保持可编辑。Noto Sans SC 按实际用字子集化并嵌入文件，未安装该字体的机器也能正确显示。
- QA 脚本是工具而非说明：`validate_deck.py`、`qa_check_pptx.py`、`qa_check_web.py` 对交付物逐项复查，发现缺陷即以非零码退出。

## 安装

```bash
# skills CLI，支持 Claude Code、Cursor、Codex 等多数 agent
npx skills add ChHsiching/BLCU-PPT-Template

# 或注册为 Claude Code 的本地插件目录
git clone https://github.com/ChHsiching/BLCU-PPT-Template.git
claude --plugin-dir BLCU-PPT-Template
```

之后提供素材与提示词即可；agent 的下一步行为由 [SKILL.md](SKILL.md) 定义。

### 环境要求

| 要求 | 用途 | 缺失时 |
|---|---|---|
| Python 3.11+ 与 python-pptx、latex2mathml、matplotlib | 全部渲染 | 硬停 |
| 模板原件置于 `templates/blcu-report/blcu-report.pptx` | 渲染 | 硬停（手动放置；有意不入库） |
| playwright + Chromium，可选 PIL | web QA 门 | 该门降级为 note |
| Node 24 + npm | 仅 web 输出 | 仅 pptx 模式不受影响 |
| Windows PowerPoint COM | 截图抽查 | 降级为 note |

## 产出物

| 文件 | 内容 |
|---|---|
| `outline.md` | 页面规划与预算核算，CP1 确认后固化 |
| `deck.json` | 两个渲染器共同消费的单一真相源 |
| `out/<名>.pptx` | 演示文稿：OMML 可编辑公式，内嵌字体 |
| `out/web/` | Vite + React 步进放映工程（`npm run dev`），可导回 pptx |
| `演讲稿.md` | 独立的口头讲稿，不写入 pptx 备注 |

仓库附带完整实例 [`examples/qat-lsq-repro/`](examples/qat-lsq-repro/)：量化感知训练复现汇报，11 页，5 个原生公式，4 张实验图，实验数字全部可溯源至 `material/make_figures.py`。

## 仓库结构

```
SKILL.md                  契约、执行纪律、8 步流程（skill 本体）
CONTEXT.md                领域词汇表：素材、门、容量预算、页型等
references/               大纲格式、deck schema、审查者契约、AI 味 pattern
scripts/                  render_pptx, scaffold_web, validate_deck,
                          qa_check_pptx, qa_check_web, embed_fonts 等
templates/blcu-report/    manifest.json（几何/预算/token 真相源）、提取的
                          品牌媒体 [pptx 原件仅存本地]
assets/web-template/      复制进演示项目的 web 放映脚手架
fonts/                    Noto Sans SC 400/700 TTF 及 OFL 许可证
examples/qat-lsq-repro/   端到端示例
tests/                    覆盖 scripts 与门的 pytest 测试套件
```

## 试一试

```bash
python -m pytest tests/

python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx

python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force
```

## 状态

v1 已完成：流水线（T1–T7）与视觉样式轮（S1–S6），并已在随仓示例上端到端验证。历史与进行中的工作见 [Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues)。
