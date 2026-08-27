# QAT-LSQ 复现：量化感知训练组会汇报（端到端示例）

端到端示例（接替已移除的 tracer-bullet）：一份真实小素材（组会风格：算法 +
公式 + 实验图）走通素材 → 大纲 → deck.json → 三产物（pptx / web / 演讲稿）
的完整流程，并接受两道 QA 门（qa_check_pptx / qa_check_web）。主题与旧
AdamW 示例不同领域：量化感知训练（QAT）复现，W4A4 下 PTQ / QAT-STE /
QAT-LSQ 三条路径对照，3 随机种子均值 ± std 的实验体裁。

## 目录

- `material/报告素材.md` — 素材文档（汇报的原始输入）
- `material/make_figures.py` — 实验图生成脚本；数字的单一真相源
- `material/images/` — 素材图片（qat-pipeline 示意 / val-curves 曲线 / bitwidth-bars 柱状 / stepscale-ablation 消融）
- `outline.md` — 大纲（CP1 决策记录 + 页面规划 + 预算核对 + 配图分配）
- `deck.json` — 中间表示（validate_deck 全绿的单一真相源）
- `演讲稿.md` — 独立讲稿交付物
- `out/` — 生成产物（gitignored：pptx 与 web 工程不入库）

## 再生三产物

前置：模板原件在 `templates/blcu-report/blcu-report.pptx`（本地放置），
Node 24 + npm，Python 依赖见根 AGENTS.md。

```bash
cd <repo>
# 实验图（改实验数字从这一步开始，再同步其余三处）
python examples/qat-lsq-repro/material/make_figures.py

# pptx（+ PowerPoint COM 截图，COM 可用即出图）
mkdir -p examples/qat-lsq-repro/out
python scripts/render_pptx.py examples/qat-lsq-repro/deck.json -o examples/qat-lsq-repro/out/qat-report.pptx
python scripts/qa_check_pptx.py examples/qat-lsq-repro/out/qat-report.pptx \
    --deck examples/qat-lsq-repro/deck.json --com-screenshots examples/qat-lsq-repro/out/com-shots

# web（scaffold 后安装并过 Playwright 门；--export-pptx 同时验证导出链）
python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json -o examples/qat-lsq-repro/out/web
cd examples/qat-lsq-repro/out/web && npm install && cd -
python scripts/qa_check_web.py examples/qat-lsq-repro/out/web \
    --screenshots examples/qat-lsq-repro/out/web-shots \
    --export-pptx examples/qat-lsq-repro/out/qat-report-from-web.pptx

# 放映：cd examples/qat-lsq-repro/out/web && npm run dev
```

演讲稿是手工交付物（`演讲稿.md`），不从 deck.json 生成。

## 数字一致性约束

素材表格、make_figures.py 的数据数组、deck 文字、演讲稿报数
（FP16 基线 78.4 ± 0.2；W4A4：PTQ 61.3 ± 0.4 / STE 73.1 ± 0.3 /
LSQ 76.2 ± 0.3，差值 14.9 / 3.1 / 2.2；W6A6：76.5 / 76.9 / 77.9；
W8A8：78.1 / 78.2 / 78.3；步长消融 r 0.5/1.0/2.0/4.0 → 73.6/75.4/76.2/74.8，
r=2.0 锚定主结果）四处必须一致；改实验数字时从 make_figures.py 开始改，
再同步其余三处。
