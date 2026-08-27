# Tracer bullet：AdamW 复现组会汇报

端到端穿膛弹（T6）：一份真实小素材（组会风格：算法 + 公式 + 实验图）走通
素材 → 大纲 → deck.json → 三产物（pptx / web / 演讲稿）的完整流程，
并接受两道 QA 门（qa_check_pptx / qa_check_web）。它同时是本仓库
skill 流程的首个可执行先例：后续生成的演示向它看齐。

## 目录

- `material/报告素材.md` — 素材文档（汇报的原始输入）
- `material/make_figures.py` — 实验图生成脚本；数字的单一真相源
- `material/images/` — 素材图片（pipeline 示意 / 损失曲线 / 精度柱状图 / λ 消融柱状图）
- `outline.md` — 大纲（CP1 固化：页面规划 + 预算核对 + 配图分配）
- `deck.json` — 中间表示（validate_deck 全绿的单一真相源）
- `演讲稿.md` — 独立讲稿交付物
- `out/` — 生成产物（gitignored：pptx 与 web 工程不入库）

## 再生三产物

前置：模板原件在 `templates/blcu-report/blcu-report.pptx`（本地放置），
Node 24 + npm，Python 依赖见根 AGENTS.md。

```bash
cd <repo>
# pptx（+ PowerPoint COM 截图抽查，COM 可用即出图）
mkdir -p examples/tracer-bullet/out
python scripts/render_pptx.py examples/tracer-bullet/deck.json -o examples/tracer-bullet/out/adamw-report.pptx
python scripts/qa_check_pptx.py examples/tracer-bullet/out/adamw-report.pptx \
    --deck examples/tracer-bullet/deck.json --com-screenshots examples/tracer-bullet/out/com-shots

# web（scaffold 后安装并过 Playwright 门；--export-pptx 同时验证导出链）
python scripts/scaffold_web.py examples/tracer-bullet/deck.json -o examples/tracer-bullet/out/web
cd examples/tracer-bullet/out/web && npm install && cd -
python scripts/qa_check_web.py examples/tracer-bullet/out/web \
    --screenshots examples/tracer-bullet/out/web-shots \
    --export-pptx examples/tracer-bullet/out/adamw-report-from-web.pptx

# 放映：cd examples/tracer-bullet/out/web && npm run dev
```

演讲稿是手工交付物（`演讲稿.md`），不从 deck.json 生成。

## 数字一致性约束

素材表格、make_figures.py 的数据数组、deck 文字、演讲稿报数
（95.2 / 93.9 / 94.6、+1.3 个百分点、std 0.1 vs 0.2；λ 消融
94.7 / 94.9 / 95.2 / 94.4）四处必须一致；
改实验数字时从 make_figures.py 开始改，再同步其余三处。
