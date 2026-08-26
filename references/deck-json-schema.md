# deck.json 格式规范

deck.json 是演示内容的中间表示与单一真相源（词汇表见 `CONTEXT.md`）：页面序列 + 每页语义块。本仓库的两个消费者是 `scripts/validate_deck.py`（写完必跑）与两个渲染器；渲染器只认本格式。**写 deck.json 之前**先读 `references/archetypes.md` 的页型语义；**所有容量数值**（字数、条数、图数上限）以 `templates/<id>/manifest.json` 的 `budget` 为真相源，本文不重复任何数字。`tests/fixtures/deck.json` 是覆盖全部 6 页型的权威样例。

## 顶层结构

```json
{
  "template": "blcu-report",
  "meta": { "presenter": "张三", "date": "2026-08-25" },
  "pages": [ { "archetype": "cover", "blocks": [ ... ] } ]
}
```

- `template`：模板 id，对应 `templates/<id>/`；校验器据此定位 manifest。
- `meta`：可整个省略；只接受 `presenter` / `date`（非空单行字符串）。供演讲稿与 web 头部使用，不占页面预算。
- `pages`：非空数组，每页 `{ archetype, blocks }`。页序即放映序。

未知字段一律拒绝（防止 `tex`/`texts` 这类拼写错误静默丢内容）；字段内容不接受换行/制表符（`latex` 除外，源码可换行）——多段内容拆成多个块或列表项。

## 语义块

| 块类型 | 字段 | 语义 |
|---|---|---|
| `title` | `text` | 页标题。每页**恰好一个**。 |
| `subhead` | `text` | 小标题。每页至多一个；仅允许 `subhead_max_chars > 0` 的页型。 |
| `text` | `text` | 段落文本。落位由页型决定（见下表）。 |
| `list` | `items[]` | 要点列表。每页至多一个块；条目单行。 |
| `formula` | `latex` | 显示公式，LaTeX 源码，pptx 走原生 OMML 注入、web 走 KaTeX。 |
| `image` | `path`，可选 `caption` | 图片引用 + 可选图注（图注渲染在所属图位内部底边）。 |

`text` 块在不同页型的落位（渲染器按此映射，写内容时按此理解字数花在哪）：

| 页型 | `text` 块落位 |
|---|---|
| cover / closing | 汇报人行（`汇报人:姓名 日期` 单行；closing 可省略） |
| agenda | 不允许（只有列表） |
| text-formula | 底部文字区（楷体段落）；公式为 0 时用全高文字区 |
| text-image | 图旁文字块 |
| chart-focus | 右侧评注（看图说话的结论句） |

## 字数计算

字数按 CJK 等宽计：East Asian W/F 字符计 1，其余（含半角标点、空格、数字）计 0.5。绑定关系：`title`/`subhead`/`text` 块与 `list` 条目、`caption` 分别对 `*_max_chars` 上限；**`text_total`** = 同页 subhead + 全部 text 块 + 全部 list 条目之和，是物理容量的主约束。预算 = 各文字 region 扣内边距后按声明字号的物理排字容量 × 0.85（安全系数：PowerPoint 实际换行早于理想排字、中英混排宽度有波动）。部分页型（如 text-formula）声明了 `text_total_max_chars_full`：**页面没有公式块时**（文字落入更高的全高区）块与总量上限放宽到该值。`latex` 不计字数。个数约束：text 块数 / 列表条数 / 公式数 / 图片数对 `*_max`、`images_min`。

## 校验器

```
python scripts/validate_deck.py <deck.json> [--format text|json] [--manifest PATH]
```

stdin 传 `-`。退出码：`0` 合法，`1` 有 findings，`2` 用法/IO/JSON 解析错误。报告每行 `<json路径>: [<code>] <message>`，可直接定位到字段；`--format json` 输出 `{"valid","page_count","findings"}` 供门与脚本消费。门与渲染器复用：进程外用 CLI + 退出码，进程内 `from validate_deck import validate_deck`（纯函数：`(deck, manifest, image_root) -> list[Finding]`，无副作用，图片存在性除外）。

finding code 前缀：`schema.*` 结构（缺字段/类型/未知字段或页型/块类型/多重性）、`budget.*` 容量（哪个预算键超了写在 message 里）、`image.path_missing`、`manifest.missing_key`（manifest 本身缺预算键）、`template.unknown|missing`（模板定位失败，仅 CLI 层）。

## 图片路径

相对 `path` 相对 **deck.json 所在目录**解析（stdin 时相对当前目录）；绝对路径亦可。图片本体按素材契约放 `material/images/`，文件命名即语义。校验只查存在性，不查尺寸/格式。
