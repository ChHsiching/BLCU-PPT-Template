<h1 align="center">BLCU-PPT-Template</h1>

<p align="center">
  <strong>一个 agent skill：把素材和一句提示词变成 BLCU 模板的组会演示——
  PPTX、web 放映、演讲稿。</strong>
</p>

<p align="center"><a href="README.md">English</a> · <b>简体中文</b></p>

---

把文档、图片和一句「用这份素材做一份组会汇报」交给 agent。skill 先规划大纲并经你确认，再渲染样张供你验收，最后产出完整成品。页面直接克隆自模板 pptx 原件，品牌层原样保留；公式转为 PowerPoint 可编辑数学对象；字体按用字子集化嵌入，文件在任何机器上都能正确显示。

## 安装

```bash
# skills CLI —— 支持 Claude Code、Cursor、Codex 等多数 agent
npx skills add ChHsiching/BLCU-PPT-Template

# 或注册为 Claude Code 的本地插件目录
git clone https://github.com/ChHsiching/BLCU-PPT-Template.git
claude --plugin-dir BLCU-PPT-Template
```

渲染前请把模板原件放到 `templates/blcu-report/blcu-report.pptx`（有意不入库），其余内容随仓库分发。

### 环境要求

| 要求 | 用途 | 缺失时 |
|---|---|---|
| Python 3.11+ 与 python-pptx、latex2mathml、matplotlib | 全部渲染 | 硬停 |
| playwright + Chromium（可选 PIL） | web QA 门 | 门记录跳过并继续 |
| Node 24 + npm | 仅 web 输出 | 仅 pptx 模式不受影响 |
| Windows PowerPoint COM | 截图抽查 | 降级为 note |

## 使用

提供素材与提示词，其余由 [SKILL.md](SKILL.md) 驱动：

```text
素材：文档 + 图片（图片文件名描述其内容）
提示词：用这份素材做一份组会汇报，pptx 和 web 都要
```

流程中你要参与两步：确认大纲（CP1）、验收样张（CP2）；全稿经脚本化质量门与 fresh subagent 审查，零缺陷后交付（CP3）。

## 产出物

| 文件 | 内容 |
|---|---|
| `outline.md` | 页面规划与预算，CP1 后固化 |
| `deck.json` | 两个渲染器共同消费的单一真相源 |
| `out/<名>.pptx` | 演示文稿：原生 OMML 公式，内嵌字体 |
| `out/web/` | Vite + React 步进放映工程（`npm run dev`） |
| `演讲稿.md` | 独立的口头讲稿 |

仓库附带完整实例 [`examples/qat-lsq-repro/`](examples/qat-lsq-repro/)，可一键再生：

```bash
python scripts/render_pptx.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/qat-report.pptx
python scripts/scaffold_web.py examples/qat-lsq-repro/deck.json \
    -o examples/qat-lsq-repro/out/web --force
```

## 开发

```bash
python -m pytest tests/
```

架构与环境说明见 [AGENTS.md](AGENTS.md)，领域词汇表见 [CONTEXT.md](CONTEXT.md)。历史与进行中的工作见 [Issues](https://github.com/ChHsiching/BLCU-PPT-Template/issues)。
