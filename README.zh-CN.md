<h1 align="center">BLCU-PPT-Template</h1>

<p align="center">
  <strong>素材 + 一句提示词 → BLCU 模板保真的组会汇报：PPTX · Web 放映 · 演讲稿</strong>
</p>

<p align="center"><a href="README.md">English</a> · <b>简体中文</b></p>

---

一个面向研究生组会汇报的 [agent skill](SKILL.md)。把你的论文笔记和一句提示词交给 agent，它按流程走完大纲 → 门 → 成品：成品克隆自真实的 BLCU 模板——原生 OMML 公式、内嵌思源黑体、母版原封不动。

```
 素材 ingest      大纲 + AI 味门       CP1 你确认        样张先行
┌──────────┐   ┌───────────┐   ┌────────────┐   ┌────────────┐
│ doc / md │──▶│ outline.md│──▶│  决策回写    │──▶│ cover+最满页 │
│ images/  │   │ 预算核对   │   │            │   │ G1 双门     │
└──────────┘   └───────────┘   └────────────┘   └────────────┘
                                                      │ CP2 你验收
      三产物 ◀── 渲染 ◀── deck.json（单一真相源）◀─────┘
   ┌─────────┐  ┌─────────┐  ┌──────────┐
   │  PPTX   │  │ web 放映 │  │ 演讲稿.md │
   │ 公式可编辑│  │ npm dev │  │ 口头讲稿  │
   └─────────┘  └─────────┘  └──────────┘
        ▲ 每步过门：机器检查 + fresh-subagent 审查，循环至零缺陷
```

## 和「拿模板手搓」的区别

- **模板是神圣的。** 每页从原件 pptx 克隆而来——绿条、logo、页码就是模板自己的母版形状，不是照着画的替代品。
- **内容形式自成体系，token 驱动。** 母版之外的一切（字重层级、强调绿 `#548235`、间距节奏）收敛在 [`templates/blcu-report/manifest.json`](templates/blcu-report/manifest.json) 一张机器可读 token 表里；两个渲染器、两道 QA 门读的是同一份。
- **门是证据不是感觉。** 所有交付物重过脚本化检查（预算/残留/几何/run 级字排/品牌层/绿条像素抽查），外加 fresh-subagent 审查循环至零缺陷。
- **公式保持可编辑。** LaTeX → MathML → OMML，拒绝截图。字体按用字子集化内嵌（<1 MB），任何机器打开都是设计时观感。

## 安装

### 接入你的编码 agent

```bash
# skills CLI 支持的任意 agent（Claude Code / Cursor / Codex …）
npx skills add ChHsiching/BLCU-PPT-Template

# 或 Claude Code 本地插件目录方式
git clone https://github.com/ChHsiching/BLCU-PPT-Template.git
claude --plugin-dir BLCU-PPT-Template
```

然后把素材丢给 agent，说一句「用这份素材做一份组会汇报」——剩下的由 [SKILL.md](SKILL.md) 接管。

### 环境要求

| 要求 | 用在哪 | 缺失时 |
|---|---|---|
| Python 3.11+ 与 python-pptx、latex2mathml、matplotlib | 全流程 | 硬停 |
| 模板原件放 `templates/blcu-report/blcu-report.pptx` | 渲染 | 硬停（手动放置；有意不入库） |
| playwright + Chromium，可选 PIL | web QA 门 | 该门降级为 note |
| Node 24 + npm | 仅 web 输出 | 仅 pptx 模式照常 |
| PowerPoint COM（Windows） | 截图抽查 | 降级为 note |

## 产出什么

| 文件 | 是什么 |
|---|---|
| `outline.md` | 页面规划 + 预算核算，CP1 你确认后固化 |
| `deck.json` | 两个渲染器共同消费的单一真相源 |
| `out/<名>.pptx` | 正式交付物：公式可编辑、字体已内嵌 |
| `out/web/` | Vite + React 步进放映（`npm run dev`），可导回 pptx |
| `演讲稿.md` | 独立口头讲稿——绝不塞进 pptx 备注 |

完整实例随仓库附带：[`examples/qat-lsq-repro/`](examples/qat-lsq-repro/) ——量化感知训练复现汇报，11 页、5 个原生公式、4 张实验图，所有数字从 `make_figures.py` 端到端可溯源。

## 仓库结构

```
SKILL.md                  契约 + 执行纪律 + 8 步流程（skill 本体）
CONTEXT.md                领域词汇表：素材/门/容量预算/页型 …
references/               大纲格式 · deck schema · 审查者契约
scripts/                  render_pptx · scaffold_web · validate_deck ·
                          qa_check_pptx · qa_check_web · embed_fonts …
templates/blcu-report/    manifest.json（几何/预算/token 真相源）、
                          提取的品牌媒体         [pptx 原件仅本地]
assets/web-template/      预置 web 放映脚手架
fonts/                    Noto Sans SC 400/700 + OFL 许可证
examples/qat-lsq-repro/   端到端示例
tests/                    覆盖 scripts 与门的 130 项测试
```

## 试一试

```bash
python -m pytest tests/                                  # 130 项测试

python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx        # 重渲示例 pptx

python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force             # 重建 web 工程
```

## 状态

v1 已完成：流水线（T1–T7）与视觉样式轮（S1–S6），并已在随仓示例上端到端验证。历史与路线图见 [Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues)。
