# web-template（renderer-web 预置脚手架）

BLCU 演示的 Vite + React 步进式放映脚手架。本目录是**预置模板**，不是直接交付的工程：`scripts/scaffold_web.py <deck.json> -o <out-dir>` 把它整体拷贝到输出目录并注入 `src/deck.json`、`src/manifest.json`（deck 所用模板的 manifest）与 `public/material/images/`（deck 引用的素材图片，按 basename 平铺）。本仓库拒绝交互式向导——脚手架只是被拷贝。

单跑本模板（自带样例 deck，即 T2 全页型夹具）：

```bash
cd assets/web-template
npm install
npm run dev
```

正式产物永远从 deck.json 生成：

```bash
python scripts/scaffold_web.py path/to/deck.json -o path/to/web-deck
cd path/to/web-deck && npm install && npm run dev
```

## 约定

- **单一真相源**：布局与字体全部读 `src/manifest.json`（regions/typography 的 inch/pt 由 `src/lib/layout.js` 换算为 1280×720 舞台像素），脚手架不含自己的几何数值。renderer-pptx 与本渲染器消费同一 manifest。
- **步进模型**：每页 1 + 列表条数 个 step；方向键 / 空格 / 点击前进，逐步揭示列表项（agenda 与 text-image 的列表）。`#页码` hash 支持直达与刷新定位。
- **公式**：KaTeX display 模式（对应 pptx 线的 LaTeX→OMML 链）。
- **字体**：manifest `typography.web_fallbacks` 的本地字体栈（SimHei/STZhongsong/KaiTi）优先，bundled Noto Sans/Serif SC 兜底，无字体缺失方块。
- **品牌层**（#17）：母版品牌元素按 `manifest.brand_layer` 实测几何复刻——内容页顶条 + 左下 logo + 右下页码；封面/结束页三 logo + 中部大带 + 右下 logo 条。条色取 `typography.tokens.colors.band`，logo 资产在 `public/brand/`（由 `scripts/export_brand_assets.py` 从 `templates/<id>/extracted/media/` 按 manifest media 字段复制）。`src/components/BrandLayer.jsx` 不含任何自创几何。
- **图片**：deck 图片路径在 scaffold 时改写为 `material/images/<basename>`（Vite public 目录，`assetUrl()` 兼容任意 base）。

`src/deck.json`、`src/manifest.json`、`public/material/` 在 scaffold 时会被整体覆盖——样例只服务本模板单跑。
