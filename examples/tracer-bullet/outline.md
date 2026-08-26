# 大纲：AdamW 优化器复现与实验分析

汇报人：陈晨 · 日期：2026-08-26 · 输出模式：both（pptx + web + 演讲稿）
素材：`material/报告素材.md` + `material/images/`（pipeline / loss-curves / acc-bars）
页型语义与容量上限：`templates/blcu-report/manifest.json`（字数按 CJK 等宽计）

## 页面规划

| # | 页型 | 标题 | 内容要点 | 配图 |
|---|------|------|----------|------|
| 1 | cover | AdamW 优化器复现与实验分析 | 汇报人行：汇报人:陈晨 2026-08-26 | — |
| 2 | agenda | 汇报大纲 | 五节：背景 / 方法 / 复现设置 / 实验 / 结论 | — |
| 3 | text-formula | 背景：L2 正则与自适应学习率的耦合 | 1 公式（耦合梯度）+ 1 段说明 | — |
| 4 | text-formula | 方法：Adam 更新与 AdamW 解耦 | 2 公式（矩估计 + AdamW 更新）+ 1 段说明 | — |
| 5 | text-image | 复现设置：数据、网络与超参 | 小标题 + 1 段文字 + 4 条要点 | pipeline.png（主位，带图注） |
| 6 | chart-focus | 实验：训练损失曲线 | 1 句看图说话 | loss-curves.png |
| 7 | chart-focus | 实验：最终测试精度 | 1 句看图说话 | acc-bars.png |
| 8 | closing | 谢谢聆听 | 汇报人行 | — |

## 容量预算核对（text_total = 小标题 + 文本块 + 列表条目；数值为 validate_deck.text_width 实算）

| # | 页型 | text_total | 上限 | 备注 |
|---|------|-----------|------|------|
| 1 | cover | 11 | 17 | 汇报人行（标题 13 ≤ 16） |
| 2 | agenda | 73.5 | 180 | 5 条 ×≤17；标签「汇报大纲」= 4 ≤ 4 |
| 3 | text-formula | 61.5 | 150 | 公式不计字数 |
| 4 | text-formula | 56 | 150 | 同上 |
| 5 | text-image | 99.5 | 120 | 含小标题 9；图注 17 ≤ 30 单列 |
| 6 | chart-focus | 24 | 60 | 看图说话结论句 |
| 7 | chart-focus | 31 | 60 | 同上 |
| 8 | closing | 5.5 | 16 | 标题 4 ≤ 8 |

全部在预算内；公式 3 个（页 3 一个、页 4 两个，formulas_max 6）；图片 3 张各归其位。

## 演讲稿要点

- 页 3-4 用讲稿补足公式细节（m̂/v̂ 偏差修正的直觉），页面只留结论性文字。
- 页 6-7 报数：95.2% vs 93.9%（+1.3 个百分点），种子间 std 0.1 vs 0.2。
- 页 8 后进入问答；补充消融（λ 扫描）留在素材里备问。
