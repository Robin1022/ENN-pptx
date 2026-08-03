# ENN-SKILL

给 Claude Code 用的公司 PPTX 生成 skill —— 通用 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 的公司专属分支。底层机制（pptxgenjs 踩坑清单、XML 直接编辑、校验器）继承自官方版本，新增的是**默认套用公司模板**，以及一套"先规划叙事、再自适应布局"的生成流程。

生成时不逐页套模块，而是先对用户素材做整体叙事规划，显式判断是否需要外部知识检索，再按每页的内容量和信息层级自适应选择版式。

## 目录

| 文件 | 作用 |
|---|---|
| [`SKILL.md`](SKILL.md) | 入口。任务分流、脚本清单、生成流程、QA 关卡 |
| [`PLANNING.md`](PLANNING.md) | 整份 PPT 的页数、叙事、逐页内容规划 + 生成前的确认关卡 |
| [`RESEARCH.md`](RESEARCH.md) | 外部检索决策、来源验证、事实标记与引用要求 |
| [`LAYOUT.md`](LAYOUT.md) | 按内容密度分配区域、选择网格、改造模块、检测无效留白 |
| [`COMPANY_STYLE.md`](COMPANY_STYLE.md) | 视觉规范（画布/配色/字体/Logo/版式一览） |
| [`MODULES.md`](MODULES.md) | 三种可选的结构化正文原型及其复用边界 |

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/thumbnail.py` | 生成带标签的缩略图网格，用来挑版式 |
| `scripts/add_slide.py` | 复制幻灯片或从版式新建，自动处理全部包级注册 |
| `scripts/clean.py` | 清理不再被引用的幻灯片、媒体和关系条目 |
| `scripts/office/validate.py` | Schema、关系、内容类型、图表和幻灯片检查 |
| `scripts/office/soffice.py` | LibreOffice 封装（沙箱环境下直接裸调 `soffice` 会卡住） |

## 安装

克隆到 Claude Code 的 skills 目录：

```bash
git clone https://github.com/Robin1022/ENN-SKILL.git ~/.claude/skills/enn-pptx
```

然后按 [`templates/README.md`](templates/README.md) 放置公司模板 —— **模板未随仓库分发**，需要自行提供，否则所有"从模板实例化"的流程无法执行。

## 依赖

`pptxgenjs`（npm）· `markitdown[pptx]`、`Pillow`、`defusedxml`、`lxml`（pip）· LibreOffice（`soffice`）· `pdftoppm`（Poppler）
