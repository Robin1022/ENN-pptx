# ENN-pptx

给 Claude Code 用的**新奥集团 PPT 模板生成 skill** —— 通用 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 的定制分支。底层机制（pptxgenjs 踩坑清单、XML 直接编辑、校验器）继承自官方版本，新增的是**默认套用新奥的 PPT 模板**，以及一套"先规划叙事、再自适应布局"的生成流程。

生成时不逐页套模块，而是先对用户素材做整体叙事规划，显式判断是否需要外部知识检索，再按每页的内容量和信息层级自适应选择版式。核心主张是**内容决定结构，模板决定风格**：栏数、条目数、图分几层跟着实际内容走，配色/字体/Logo/图形语言严格跟模板一致。

面向新奥集团各板块使用。当前视觉规范基线取自**新奥新智（ENNEW）**的模板；其他板块如果用的是另一套模板，需要按下面的[模板适配](#模板适配)重测一遍。

---

## ⚠️ 使用前必读

**`company-template.pptx` 不在仓库里。** 它被 `.gitignore` 排除——内含真实内部示例文案，不适合放在公开仓库中。需要你自行放置，见 [`templates/README.md`](templates/README.md)。

`templates/` 下的**场景模板**不受此限制，可以直接取用，见下面的[场景模板](#场景模板)。

**文档里的视觉数值有板块归属。** `COMPANY_STYLE.md` 的画布尺寸、Logo 坐标、配色表、字体、9 个版式一览，以及 `MODULES.md` 的三个内容模块，全部是从**新奥新智模板**实测得出的。如果你所在板块用的是同一套模板，可以直接用；如果不是，这些数值需要重测，否则生成的 PPT 在 Logo 位置、配色和版式编号上都会对不上。

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

### 3. 放置模板

```bash
cp /path/to/模板.pptx ~/.claude/skills/enn-pptx/templates/company-template.pptx
```

文件名必须是 `company-template.pptx`，否则 `SKILL.md` 里所有"从模板实例化"的流程都找不到它。

### 4. 确认是否需要适配

- **用的是新奥新智那套模板** → 到这里就可以开始用了
- **用的是其他板块的模板** → 继续下一节

---

## 模板适配

只在换用不同模板时需要做，做一次即可。目的是把 `COMPANY_STYLE.md` 和 `MODULES.md` 里新智模板的数值，换成你这套模板的数值。

### 第 1 步：盘点模板结构

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

### 第 2 步：重测 `COMPANY_STYLE.md` 的八项

这是适配工作量最大的一步。下表右列是新智模板的实测值，需要逐项换成你这套模板的：

| 项目 | 怎么测 | 新智模板的基线值 |
|---|---|---|
| **画布尺寸** | `/tmp/tpl/ppt/presentation.xml` 的 `<p:sldSz>` | `12192000 x 6858000` EMU（13.33in × 7.5in） |
| **Logo 位置与尺寸** | 在 `slideMaster1.xml` 里找 `<p:pic>` 的 `<a:off>` / `<a:ext>` | 位置 `(10.56in, 0.21in)`，尺寸 `2.41in × 0.85in` |
| **Logo 继承方式** | 确认 Logo 在母版还是版式上 | 母版承载，9 个版式自动继承 |
| **配色表** | grep 各页 `srgbClr`，统计高频色值 | 蓝 `2B5BD7`、青绿 `2E9482`、警示红 `FF0000` 等 |
| **强调色选用规则** | 观察不同主题页面用了什么色 | "组织/认知类用蓝、技术/架构类用青绿" |
| **字体** | grep `latin typeface` / `ea typeface` | 微软雅黑 / 等线 / Arial |
| **版式一览表** | 逐个看 `slideLayoutN.xml` 的名称和占位符 | 9 个版式的名称、占位符、适用场景 |
| **常见图形语言** | 从示例页归纳视觉习惯 | 圆角卡片、虚线边框、编号徽标、分层色块 |

> 改完记得把文件顶部的来源说明也更新，写清楚是哪个板块的模板、实测日期是什么，方便日后换模板时知道该重测什么。

### 第 3 步：重定义或删除 `MODULES.md`

现有的三个模块（反思模块 / 技术图谱模块 / 技术方案详情模块）绑定在新智模板的三个示例正文页（`slide1/2/3.xml`）上。

- **你的模板也有可复用的原型页** → 按同样的格式重写：每个模块写清楚主标题性质、副标题性质、正文结构、强调色，以及"结构不匹配时怎么改造"的指引
- **模板只有空白版式，没有原型页** → 把 `MODULES.md` 内容清空或删除，并去掉 `SKILL.md`、`PLANNING.md`、`LAYOUT.md` 里对它的引用。skill 依然可用，只是每页都走"从版式实例化 + 按图形语言搭建"的路径

**不要保留原有的三个模块定义不改**——它们指向的 `slide1/2/3.xml` 在你的模板里不存在或内容完全不同，会让模型按错误的假设去复制页面。

### 适配完成检查

- [ ] `templates/company-template.pptx` 已放置
- [ ] 依赖已装齐（`python scripts/office/validate.py --help` 能跑通）
- [ ] `COMPANY_STYLE.md` 八项数值已换成本板块模板的实测值
- [ ] `MODULES.md` 已重定义或已删除（并清理了引用）
- [ ] 跑一遍完整流程生成一份测试 deck，视觉检查 Logo 位置、配色、字体是否正确

---

## 场景模板

`templates/` 下除了需要自行放置的 `company-template.pptx`，还收了开箱即用的场景模板。它们已清空全部个人信息，复制一份改文字即可，**不需要**经过"从模板实例化"的生成流程。

| 模板 | 场景 | 页数 |
|---|---|---|
| [`templates/campus-hire-defence.pptx`](templates/campus-hire-defence.pptx) | 校招生转正答辩 | 14 |

**校招生转正答辩**：从 HR 的「新进人员任前答辩」表单改造而来，针对校招生场景调整了措辞与结构——`过往业绩亮点及资源情况` 改为 `项目与科研经历`、`目标承诺` 改为 `转正后的工作计划`、`年度重点工作落实思路`（Q1–Q4）改为 `能力提升路径`（首月 / 3 / 6 / 12 个月），并新增 `2.3 从校园到职场的转变与自我复盘`。逐页结构和改造理由见 [`templates/README.md`](templates/README.md)。

```bash
cp templates/campus-hire-defence.pptx ~/Desktop/我的转正答辩.pptx
```

---

## 目录

| 文件 | 作用 | 换模板时是否要改 |
|---|---|---|
| [`SKILL.md`](SKILL.md) | 入口。任务分流、脚本清单、生成流程、QA 关卡 | 否 |
| [`PLANNING.md`](PLANNING.md) | 整份 PPT 的页数、叙事、逐页内容规划 + 生成前的确认关卡 | 否 |
| [`RESEARCH.md`](RESEARCH.md) | 外部检索决策、来源验证、事实标记与引用要求 | 否 |
| [`LAYOUT.md`](LAYOUT.md) | 按内容密度分配区域、选择网格、改造模块、检测无效留白 | 否 |
| [`COMPANY_STYLE.md`](COMPANY_STYLE.md) | 视觉规范（画布/配色/字体/Logo/版式一览） | **是，整篇重测** |
| [`MODULES.md`](MODULES.md) | 内容原型模块及其复用边界 | **是，重定义或删除** |

`PLANNING.md` / `RESEARCH.md` / `LAYOUT.md` 这三份是与具体模板无关的方法论，任何板块都可以原样用——它们也是这个仓库相对官方 pptx skill 的主要增量。

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
