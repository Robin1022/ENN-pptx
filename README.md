# ENN-pptx

给 Claude Code 用的**公司模板 PPTX 生成 skill** —— 通用 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 的定制分支。底层机制（pptxgenjs 踩坑清单、XML 直接编辑、校验器）继承自官方版本，新增的是**默认套用你自己的公司模板**，以及一套"先规划叙事、再自适应布局"的生成流程。

生成时不逐页套模块，而是先对用户素材做整体叙事规划，显式判断是否需要外部知识检索，再按每页的内容量和信息层级自适应选择版式。核心主张是**内容决定结构，模板决定风格**：栏数、条目数、图分几层跟着实际内容走，配色/字体/Logo/图形语言严格跟模板一致。

---

## ⚠️ 使用前必读：这个仓库是脱敏版

仓库里**不含模板文件**，而且**所有文档中的数值都是从作者自己的公司模板实测得出的**。直接克隆下来是跑不通的，也不该照搬。

具体来说：

- `templates/company-template.pptx` **未随仓库分发**（被 `.gitignore` 排除，因为模板内含真实内部示例文案）
- `COMPANY_STYLE.md` 里的画布尺寸、Logo 坐标、配色表、字体、9 个版式一览 —— **全部是作者模板的实测值，对你的模板一律无效**
- `MODULES.md` 定义的三个内容模块 —— 绑定在作者模板的 `slide1/2/3.xml` 上，**你的模板没有这三页**

换句话说：**你拿到的是一套方法论 + 一套工具脚本，不是一个开箱即用的成品。** 视觉规范部分需要你按自己的模板重做一遍。

下面的[首次适配](#首次适配)一节给出了完整步骤。

---

## 快速开始

### 1. 安装 skill

克隆到 Claude Code 的 skills 目录：

```bash
git clone https://github.com/Robin1022/ENN-pptx.git ~/.claude/skills/enn-pptx
```

目录名可以自己改，Claude Code 识别的是 `SKILL.md` 里 frontmatter 的 `name` 字段。

### 2. 安装依赖

```bash
# Python 依赖：文本提取 / 缩略图 / 清理 / 校验
pip install "markitdown[pptx]" Pillow defusedxml lxml

# Node 依赖：pptxgenjs 用于从零绘制内容区
npm install -g pptxgenjs

# 系统依赖：LibreOffice 用于转 PDF，Poppler 用于转图片做视觉 QA
brew install --cask libreoffice   # macOS
brew install poppler
```

Linux 下把最后两行换成对应的包管理器命令（`libreoffice`、`poppler-utils`）。

### 3. 放置你的模板

```bash
cp /path/to/你的公司模板.pptx ~/.claude/skills/enn-pptx/templates/company-template.pptx
```

文件名必须是 `company-template.pptx`，否则 `SKILL.md` 里所有"从模板实例化"的流程都找不到它。

### 4. 完成首次适配

见下一节。**跳过这一步会生成跟你公司品牌完全不符的 PPT。**

---

## 首次适配

只需做一次。目的是把文档里作者模板的数值，换成你自己模板的数值。

### 第 1 步：盘点你的模板结构

```bash
cd ~/.claude/skills/enn-pptx
python scripts/thumbnail.py templates/company-template.pptx template-thumbs
```

会生成 `template-thumbs.jpg`（超过 12 页拆成多张），是带编号标签的缩略图网格。对着它确认：

- 模板有几个版式（`slideLayoutN.xml`），各是什么用途
- Logo 在哪一层——母版（`slideMaster`）还是每页单独放
- 有没有可以当作"内容原型"复用的示例正文页

同时解压模板看 XML 结构：

```bash
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('/tmp/tpl')" templates/company-template.pptx
ls /tmp/tpl/ppt/slideLayouts/ /tmp/tpl/ppt/slides/
```

### 第 2 步：重写 `COMPANY_STYLE.md`

这是适配工作量最大的一步。需要**逐项重新实测**并替换，下面每一项在现有文档里都是作者模板的值：

| 项目 | 怎么测 | 现有文档里的值（需替换） |
|---|---|---|
| **画布尺寸** | `/tmp/tpl/ppt/presentation.xml` 的 `<p:sldSz>` | `12192000 x 6858000` EMU（13.33in × 7.5in） |
| **Logo 位置与尺寸** | 在 `slideMaster1.xml` 里找 `<p:pic>` 的 `<a:off>` / `<a:ext>` | 位置 `(10.56in, 0.21in)`，尺寸 `2.41in × 0.85in` |
| **Logo 继承方式** | 确认 Logo 在母版还是版式上 | 母版承载，9 个版式自动继承 |
| **配色表** | grep 各页 `srgbClr`，统计高频色值 | 蓝 `2B5BD7`、青绿 `2E9482`、警示红 `FF0000` 等 |
| **强调色选用规则** | 观察不同主题页面用了什么色 | "组织/认知类用蓝、技术/架构类用青绿" |
| **字体** | grep `latin typeface` / `ea typeface` | 微软雅黑 / 等线 / Arial |
| **版式一览表** | 逐个看 `slideLayoutN.xml` 的名称和占位符 | 9 个版式的名称、占位符、适用场景 |
| **常见图形语言** | 从示例页归纳视觉习惯 | 圆角卡片、虚线边框、编号徽标、分层色块 |

> **建议**：把文件顶部的来源说明也一并改掉，写清楚你的模板来源和实测日期，方便日后换模板时知道该重测什么。

### 第 3 步：重定义或删除 `MODULES.md`

现有的三个模块（反思模块 / 技术图谱模块 / 技术方案详情模块）是**作者公司 PPT 的常用结构**，绑定在作者模板的三个示例正文页上。

你有两个选择：

- **有自己的原型页** → 按同样的格式重写：每个模块写清楚主标题性质、副标题性质、正文结构、强调色、以及"结构不匹配时怎么改造"的指引
- **没有原型页**（模板只有空白版式）→ 直接把 `MODULES.md` 内容清空或删除，并去掉 `SKILL.md`、`PLANNING.md`、`LAYOUT.md` 里对它的引用。skill 依然可用，只是每页都走"从版式实例化 + 按图形语言搭建"的路径

**不要保留原有的三个模块定义不改**——它们指向的 `slide1/2/3.xml` 在你的模板里不存在或内容完全不同，会让模型按错误的假设去复制页面。

### 第 4 步：改 `SKILL.md` 的触发词

`SKILL.md` 顶部 frontmatter 的 `description` 决定了 Claude Code 什么时候激活这个 skill。现有描述里写死了作者公司的名称，需要换成你自己的：

```yaml
---
name: enn-pptx          # 可改成你自己的 skill 名
description: "……触发场景：用户要求「用公司模板」「公司PPT」「内部汇报」「<你的公司名/品牌名>」相关的演示文稿制作……"
---
```

不改的话，用你公司名触发时不会命中这个 skill。

### 适配完成检查

- [ ] `templates/company-template.pptx` 已放置
- [ ] 依赖已装齐（`python scripts/office/validate.py --help` 能跑通）
- [ ] `COMPANY_STYLE.md` 八项数值全部替换为实测值
- [ ] `MODULES.md` 已重定义或已删除（并清理了引用）
- [ ] `SKILL.md` 的 `description` 触发词已换成自己公司名
- [ ] 跑一遍完整流程生成一份测试 deck，视觉检查 Logo 位置、配色、字体是否正确

---

## 目录

| 文件 | 作用 | 是否需要适配 |
|---|---|---|
| [`SKILL.md`](SKILL.md) | 入口。任务分流、脚本清单、生成流程、QA 关卡 | 只改 frontmatter 触发词 |
| [`PLANNING.md`](PLANNING.md) | 整份 PPT 的页数、叙事、逐页内容规划 + 生成前的确认关卡 | 通用，无需改 |
| [`RESEARCH.md`](RESEARCH.md) | 外部检索决策、来源验证、事实标记与引用要求 | 通用，无需改 |
| [`LAYOUT.md`](LAYOUT.md) | 按内容密度分配区域、选择网格、改造模块、检测无效留白 | 通用，无需改 |
| [`COMPANY_STYLE.md`](COMPANY_STYLE.md) | 视觉规范（画布/配色/字体/Logo/版式一览） | **必须整篇重测** |
| [`MODULES.md`](MODULES.md) | 内容原型模块及其复用边界 | **必须重定义或删除** |

`PLANNING.md` / `RESEARCH.md` / `LAYOUT.md` 这三份是与品牌无关的方法论，可以原样复用——它们也是这个仓库相对官方 pptx skill 的主要增量。

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/thumbnail.py` | 生成带标签的缩略图网格，用来挑版式（第二个参数一定要传，否则会覆盖同目录的旧缩略图） |
| `scripts/add_slide.py` | 复制幻灯片或从版式新建，自动处理全部包级注册 |
| `scripts/clean.py` | 清理不再被引用的幻灯片、媒体和关系条目 |
| `scripts/office/validate.py` | Schema、关系、内容类型、图表和幻灯片检查（从模板派生的 deck 记得传 `--original`） |
| `scripts/office/soffice.py` | LibreOffice 封装（沙箱环境下直接裸调 `soffice` 会卡住） |

## 依赖

`pptxgenjs`（npm）· `markitdown[pptx]`、`Pillow`、`defusedxml`、`lxml`（pip）· LibreOffice（`soffice`）· `pdftoppm`（Poppler）

## 来源

底层机制与 `scripts/` 下的工具脚本来自 Anthropic 官方 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx)。该 skill 为 source-available，其许可条款见上游仓库的 `skills/pptx/LICENSE.txt`，使用前请自行确认。

本仓库新增的部分是 `PLANNING.md` / `RESEARCH.md` / `LAYOUT.md` / `MODULES.md` 所描述的规划与布局方法论，以及围绕公司模板的适配流程。
