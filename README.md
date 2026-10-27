# ENN-pptx

给 Claude Code 用的**新奥集团 PPT 模板生成 skill**，是通用 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 的定制分支。底层机制（pptxgenjs 踩坑清单、XML 直接编辑、校验器）继承自官方版本，新增的是**默认套用新奥的 PPT 模板**，以及一套"先规划叙事、再自适应布局"的生成流程。

生成时不逐页套模块，而是先对用户素材做整体叙事规划，判断是否需要外部检索，再按每页的内容量和信息层级选择版式。原则是**内容决定结构，模板决定风格**：栏数、条目数、图分几层跟着实际内容走，配色、字体、Logo、图形语言与模板保持一致。

当前视觉规范的基线取自**新奥新智（ENNEW）**的模板；其他板块用的是另一套模板时，按下面的[模板适配](#模板适配)重测一遍。

---

## 版本

仓库有三个分支，对应三个 skill，可以同时安装、按名字调用：

| 分支 | skill 名 | 说明 |
|---|---|---|
| `main`（本分支） | `enn-pptx-origin` | 长期使用的稳定版，留作备份和对照 |
| [`ENNEW`](https://github.com/Robin1022/ENN-pptx/tree/ENNEW) | `enn-pptx-new` | 在稳定版基础上，按近期真实使用会话优化：修复校验漏报、新增版面几何检查、改稿流程默认原地写回、页面不再出现推测标记 |
| [`enn-pptx-2.0`](https://github.com/Robin1022/ENN-pptx/tree/enn-pptx-2.0) | `enn-pptx-new-2` | 按 skill 写作规则重构后，再做同一批优化；多了交付检查脚本、模板写保护 hook、测试任务和场景模板 |

三个 skill 的触发描述相同，使用时用 `/enn-pptx-origin`、`/enn-pptx-new`、`/enn-pptx-new-2` 指定。两个优化分支各自的改动依据见分支里的 `EXPERIENCE.md`。

---

## 快速开始

### 1. 安装 skill

```bash
git clone -b main https://github.com/Robin1022/ENN-pptx.git ~/.claude/skills/enn-pptx-origin
```

目录名可以自己改，Claude Code 识别的是 `SKILL.md` 里 frontmatter 的 `name` 字段。

### 2. 安装依赖

```bash
# Python 依赖：文本提取 / 缩略图 / 清理 / 校验
pip install "markitdown[pptx]" Pillow defusedxml lxml

# Node 依赖：pptxgenjs 用于给正文区补画内容
npm install -g pptxgenjs

# 系统依赖：LibreOffice 转 PDF，Poppler 转图片做视觉检查
brew install --cask libreoffice   # macOS
brew install poppler
```

Linux 下把最后两行换成对应的包管理器命令（`libreoffice`、`poppler-utils`）。

### 3. 确认模板

仓库已收录新奥新智模板 `templates/company-template.pptx`，用同一套模板的板块装完即可使用。

用其他板块的模板时，用同名文件覆盖，再做下一节的模板适配：

```bash
cp /path/to/模板.pptx ~/.claude/skills/enn-pptx-origin/templates/company-template.pptx
```

文件名必须是 `company-template.pptx`，`SKILL.md` 里所有"从模板实例化"的流程都按这个名字找它。

---

## 模板适配

只在换用不同模板时需要做，做一次即可：把 `COMPANY_STYLE.md` 和 `MODULES.md` 里新奥新智模板的数值，换成你这套模板的数值。不重测直接用，生成的 PPT 在 Logo 位置、配色和版式编号上都会对不上。

### 第 1 步：盘点模板结构

```bash
cd ~/.claude/skills/enn-pptx-origin
python3 scripts/thumbnail.py templates/company-template.pptx template-thumbs
```

生成带编号标签的缩略图网格 `template-thumbs.jpg`（超过 12 页拆成多张）。对着它确认：

- 模板有几个版式（`slideLayoutN.xml`），各是什么用途
- Logo 在母版（`slideMaster`）上还是每页单独放
- 有没有可以当作内容原型复用的示例正文页

同时解压模板看 XML 结构：

```bash
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('/tmp/tpl')" templates/company-template.pptx
ls /tmp/tpl/ppt/slideLayouts/ /tmp/tpl/ppt/slides/
```

### 第 2 步：重测 `COMPANY_STYLE.md` 的八项

右列是新奥新智模板的实测值，逐项换成你这套模板的：

| 项目 | 怎么测 | 新奥新智模板的基线值 |
|---|---|---|
| 画布尺寸 | `/tmp/tpl/ppt/presentation.xml` 的 `<p:sldSz>` | `12192000 x 6858000` EMU（13.33in × 7.5in） |
| Logo 位置与尺寸 | `slideMaster1.xml` 里 `<p:pic>` 的 `<a:off>` / `<a:ext>` | 位置 `(10.56in, 0.21in)`，尺寸 `2.41in × 0.85in` |
| Logo 继承方式 | Logo 在母版还是版式上 | 母版承载，9 个版式自动继承 |
| 配色表 | 统计各页 `srgbClr` 的高频色值 | 蓝 `2B5BD7`、青绿 `2E9482`、警示红 `FF0000` 等 |
| 强调色选用规则 | 观察不同主题的页面用什么色 | 组织/认知类用蓝，技术/架构类用青绿 |
| 字体 | 查 `latin typeface` / `ea typeface` | 微软雅黑 / 等线 / Arial |
| 版式一览表 | 逐个看 `slideLayoutN.xml` 的名称和占位符 | 9 个版式的名称、占位符、适用场景 |
| 常见图形语言 | 从示例页归纳视觉习惯 | 圆角卡片、虚线边框、编号徽标、分层色块 |

改完把 `COMPANY_STYLE.md` 顶部的来源说明一并更新，写清楚是哪个板块的模板、实测日期。

### 第 3 步：重定义或删除 `MODULES.md`

现有三个模块（反思 / 技术图谱 / 技术方案详情）绑定在新奥新智模板的三个示例正文页（`slide1/2/3.xml`）上。

- **你的模板也有可复用的原型页**：按同样的格式重写，每个模块写清主标题、副标题、正文结构、强调色，以及结构对不上时怎么改造
- **模板只有空白版式**：清空或删除 `MODULES.md`，并去掉 `SKILL.md`、`PLANNING.md`、`LAYOUT.md` 里对它的引用。每页改走"从版式实例化 + 按图形语言搭建"

不要保留原有的三个模块定义不改：它们指向的 `slide1/2/3.xml` 在你的模板里不存在或内容不同，会让模型按错误的假设复制页面。

### 适配完成检查

- [ ] `templates/company-template.pptx` 已换成本板块的模板
- [ ] 依赖已装齐（`python3 scripts/office/validate.py --help` 能跑通）
- [ ] `COMPANY_STYLE.md` 八项数值已换成本板块模板的实测值
- [ ] `MODULES.md` 已重定义或删除，并清理了引用
- [ ] 跑一遍完整流程生成测试 deck，检查 Logo 位置、配色、字体

---

## 目录

| 文件 | 作用 | 换模板时是否要改 |
|---|---|---|
| [`SKILL.md`](SKILL.md) | 入口：任务分流、脚本清单、生成流程、质量检查 | 否 |
| [`PLANNING.md`](PLANNING.md) | 整份 PPT 的页数、叙事、逐页内容规划，以及生成前的确认关卡 | 否 |
| [`RESEARCH.md`](RESEARCH.md) | 外部检索决策、来源验证、事实标记与引用要求 | 否 |
| [`LAYOUT.md`](LAYOUT.md) | 按内容密度分配区域、选择网格、改造模块、检测无效留白 | 否 |
| [`COMPANY_STYLE.md`](COMPANY_STYLE.md) | 视觉规范：画布、配色、字体、Logo、版式一览 | 是，整篇重测 |
| [`MODULES.md`](MODULES.md) | 内容原型模块及其复用边界 | 是，重定义或删除 |
| [`templates/company-template.pptx`](templates/company-template.pptx) | 新奥新智公司模板 | 是，换成本板块模板 |

`PLANNING.md` / `RESEARCH.md` / `LAYOUT.md` 是与具体模板无关的方法论，任何板块都可以原样使用。

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/thumbnail.py` | 生成带标签的缩略图网格，用来挑版式（第二个参数一定要传，否则会覆盖同目录的旧缩略图） |
| `scripts/add_slide.py` | 复制幻灯片或从版式新建，自动处理全部包级注册 |
| `scripts/clean.py` | 清理不再被引用的幻灯片、媒体和关系条目 |
| `scripts/office/validate.py` | Schema、关系、内容类型、图表和幻灯片检查（从模板派生的 deck 记得传 `--original`） |
| `scripts/office/soffice.py` | LibreOffice 封装（直接调用 `soffice` 可能卡住） |

## 来源

底层机制与 `scripts/` 下的工具脚本来自 Anthropic 官方 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx)。该 skill 为 source-available，许可条款见上游仓库的 `skills/pptx/LICENSE.txt`，使用前请自行确认。

本仓库新增的是 `PLANNING.md` / `RESEARCH.md` / `LAYOUT.md` / `MODULES.md` 描述的规划与布局方法论，以及围绕公司模板的适配流程。
