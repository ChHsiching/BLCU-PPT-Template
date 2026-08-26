# 审查者契约（G1 / G3 派遣）

G1 与 G3 的审查者必须是 fresh subagent（执行纪律：生产者不审自己的产出）。派遣时把下方提示词模板填入当次输入清单整体发出；审查者只读、defect-first、按五类具名缺陷报告，No findings 是合法结论。

## 派遣提示词模板

```text
你是独立审查者，只读审查：不修改任何文件，不跑渲染器（cat/grep 等只读命令可用）。
逐页通读全部输入：截图逐张看、deck 逐块读、讲稿逐节读，再对照素材。

输入：
- <outline.md 路径>、<deck.json 路径>、<素材目录 material/>（G3 另附 <素材要点.md 路径>——数字与出处台账，及 <演讲稿.md 路径>）
- <产物截图：COM 截图目录（pptx 产物）/ web 截图目录，按当次范围>
- 真相源：templates/blcu-report/manifest.json（预算）、references/archetypes.md（页型语义）
- AI 味 pattern：`references/ai-slop-patterns.md`

只报告以下五类具名缺陷：
1. 占位残留 —— xxx / TODO / lorem / [insert / 待补充 等未定内容。
2. 超预算 —— 超出 manifest budget 的页或块，引出双方数值。
3. AI 味引证 —— 命中 pattern 清单的原文，引 pattern 名 + 原文行。
4. 与素材事实不符 —— 页面或讲稿的数字/结论在素材中无出处或相抵触，引两侧原文。
5. 页型错配 —— 内容形态与页型语义冲突（如 chart-focus 放两张图、text-formula 塞代码）。

输出：findings 列表，每条一行：
  [P0-P3] 类别 | 位置（页/块/行）| 证据原文 | 一句修法
无缺陷时输出「No findings.」——这是合法结论，绝不为交差发明缺陷。
```

## 严重度（P0–P3）

- **P0** 产物不可用：文件打不开、页数与大纲不符、占位残留。
- **P1** 内容失真：与素材事实不符、超容量预算。
- **P2** 形态与文风：页型错配、AI 味引证。
- **P3** 瑕疵：不改不损害事实与结构（如措辞可再紧）。

分级只排修复顺序；零缺陷标准不分级，P3 也要修。

## 各门输入范围

- **G1（样张）**：样张 deck.json + 样张产物截图 + outline.md + 素材。
- **G3（全稿）**：全稿 deck.json + outline.md + 素材 + 素材要点.md + 全部产物截图（COM / web）+ 演讲稿。

## fail loop

按 SKILL.md 执行纪律 2、3：修完全部 findings，由**新的** fresh subagent 重审整道门，循环至 No findings。

## 环境降级

运行环境无 subagent 派遣能力时，由生产者按本契约全文自查替代（同样的五类缺陷、全量输入、No findings 门槛），并在门证据里记录「自查替代」；机器门结论不受影响。
