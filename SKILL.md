---
name: enn-pptx-2-0
description: "用公司自有 PPT 模板生成/编辑内部汇报、方案材料等 .pptx 文件。触发场景：用户要求「用公司模板」「公司PPT」「内部汇报」「新奥」「新奥新智/ENNEW」相关的演示文稿制作,或明确引用 templates/company-template.pptx。与通用 pptx skill 的区别:这个 skill 默认套用公司品牌(配色/字体/Logo/版式),不做与品牌无关的通用/个人 PPT——那种场景交给通用 pptx skill。"
hooks:
  PreToolUse:
    - matcher: "Bash|Edit|Write|MultiEdit"
      hooks:
        - type: command
          command: "python3 \"${CLAUDE_SKILL_DIR:-$HOME/.claude/skills/enn-pptx-2-0}/scripts/guard_templates.py\""
---

# 公司 PPTX 生成、编辑与分析

## 不可省的规则

1. **新建 deck 先交整体规划,用户确认后才动手**(`PLANNING.md`)。素材缺口自主推理补全,但推测内容必须标 `[推理补充]` / `[待确认假设]` / `[示意数据]`,不得写成用户提供的事实。
2. **模板原文件只读。** 一律 `cp` 出工作副本再操作;写入 `templates/` 下模板的命令会被 hook 拦截。
3. **成稿不得带模板示例文案。** 所有 `add_slide.py` 复制完成后,从 `<p:sldIdLst>` 移除示例页,再跑 `clean.py`。
4. **交付前三道检查全过才算完成**:`validate.py --original`、`check_delivery.py`、渲染后的视觉 QA。任何一道失败,修完从失败的那一步重跑。

## 配套文档

动手前按任务读对应文档:

- [`PLANNING.md`](PLANNING.md) —— 页数、叙事、逐页内容与布局规划 + 确认关卡
- [`RESEARCH.md`](RESEARCH.md) —— 是否外部检索、来源验证、事实标记与引用
- [`LAYOUT.md`](LAYOUT.md) —— 按内容密度分配区域、选网格、改造模块、检测无效留白
- [`COMPANY_STYLE.md`](COMPANY_STYLE.md) —— 画布/配色/字体/Logo/版式一览
- [`MODULES.md`](MODULES.md) —— 三种可选的结构化正文原型
- [`templates/README.md`](templates/README.md) —— 场景模板逐页结构
- [`README.md` 的「模板适配」](README.md#模板适配) —— 换用其他板块模板时的重测步骤(`COMPANY_STYLE.md` 的数值只对新奥新智模板成立)

本 skill 是通用 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 的公司分支,脚本和踩坑清单继承自官方版本。下文路径都相对本 skill 目录。

## 任务分流

| 任务 | 方法 |
|---|---|
| **场景已有现成模板**(见下表) | 复制 `templates/` 下的场景模板改文字,**不走** `company-template.pptx` 实例化流程——那会把做好的版面重做一遍 |
| **新建公司风格演示文稿** | 走下面的默认工作流 |
| **编辑现有的公司 PPT** | 见「编辑现有演示文稿和模板」,同样遵循公司规范 |
| **读取内容** | `markitdown deck.pptx`(每页一个 `<!-- Slide number: N -->` 区块);缩略图:`python3 scripts/thumbnail.py deck.pptx <prefix>` |

用户没有要求公司品牌/模板、且是跟公司汇报无关的私人演示文稿时,改用通用 `pptx` skill。

### 现成的场景模板

| 模板 | 命中场景 | 页数 |
|---|---|---|
| `templates/campus-hire-defence.pptx` | 校招生转正答辩 / 新进人员任前答辩 / 试用期转正述职 | 14 |

`cp templates/campus-hire-defence.pptx <目标路径>` 后逐页替换文字;改动时同样适用「填充模板时」的全部注意事项,尤其是**填充元素只能替换不能新增**(`<p:spPr>` 的填充是互斥选择组,插出两个并列的 `<a:solidFill>` 会让 PowerPoint 报损坏,而 XSD、LibreOffice、python-pptx 都不报错)。交付前检查用 `--scaffold` 指向这份场景模板。

## 默认工作流:用公司模板新建

把这份清单复制到回复里,逐项打勾:

```text
- [ ] 1. 规划:按 PLANNING.md 出整体规划(含 RESEARCH.md 的检索决策、LAYOUT.md 的逐页密度判断),用户确认
- [ ] 2. 建页:cp 模板 → 解压 → add_slide.py 建出全部目标页
- [ ] 3. 定稿结构:调整 <p:sldIdLst> 顺序,移除示例页引用,跑 clean.py
- [ ] 4. 填内容:改 ppt/slides/slideN.xml,或用 pptxgenjs 重搭不匹配的区域
- [ ] 5. 打包 + validate.py --original          失败 → 回到第 4 步修
- [ ] 6. check_delivery.py --template           失败 → docProps 类加 --fix;文案残留回到第 4 步
- [ ] 7. 渲染成图片做视觉 QA                   有缺陷 → 回到第 4 步,修完从第 5 步重跑
```

第 1 步的规划要求、标记规则和输出格式全在 `PLANNING.md`。三个示例正文页是可选原型:每页由模型按内容语义判断是否复制对应源页,不匹配就用 `COMPANY_STYLE.md` 的其他版式,不让用户逐页挑模块。

```bash
# 2. 解压工作副本
cp templates/company-template.pptx /tmp/deck.pptx
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('/tmp/unpacked')" /tmp/deck.pptx
#    匹配三种模块的页 = 复制 slide1/2/3.xml(见 MODULES.md);封面/目录/其他 = 从 slideLayoutN.xml 实例化。
#    Content_Types / rels / <p:sldIdLst> 自动处理,Logo 随版式/母版自动带出
python3 scripts/add_slide.py /tmp/unpacked/ slide1.xml --after slide3.xml        # 例:反思模块
# python3 scripts/add_slide.py /tmp/unpacked/ slideLayout3.xml --after slide3.xml # 例:从"正文页"版式新开

# 3. 全部页面建好后:改 ppt/presentation.xml 的 <p:sldIdLst> 定顺序,删掉 slide1/2/3 示例页引用
#    (用户明确要保留的除外)。不要直接删 slide 文件,交给 clean.py
python3 scripts/clean.py /tmp/unpacked/

# 4. 编辑最终保留页:结构跟源页接近 = 安全文字替换(见「填充模板时」);
#    栏数/条目数/层级对不上 = 保留 COMPANY_STYLE.md 的图形语言,增删形状子组或用 pptxgenjs 重搭该区域

# 5-6. 打包与检查
(cd /tmp/unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)
python3 scripts/office/validate.py /tmp/out.pptx --original templates/company-template.pptx
python3 scripts/check_delivery.py /tmp/out.pptx --template templates/company-template.pptx
```

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/thumbnail.py deck.pptx [prefix]` | 每页带标签的缩略图网格,用来挑版式。只认 `.pptx`。**一定传 `prefix`**——默认 `thumbnails`,同目录处理另一份 deck 会覆盖 |
| `scripts/add_slide.py unpacked/ slide2.xml [--after slideN.xml]` | 复制一页(或从 `slideLayoutN.xml` 新建),自动处理包级注册,并摘掉 `notesSlide`/`tags` 引用。两页形状指向同一个 `<p:tags>` 部件时,真实 PowerPoint 会静默清空整页,而 XSD/python-pptx/LibreOffice 都看不出来——自己手写复制逻辑时要防这个坑。直接处理 `.pptx` 时**必须传 `-o out.pptx`**,否则原地改写输入 |
| `scripts/clean.py unpacked/` | 删除不再被引用的幻灯片、媒体和关系。`<p:sldIdLst>` 定稿**之后**运行 |
| `scripts/office/validate.py deck.pptx [--original src.pptx]` | Schema、关系、内容类型、图表和幻灯片检查,每条失败附修复方法。模板派生的 deck 必须传 `--original` |
| `scripts/check_delivery.py deck.pptx [--template ...] [--scaffold ...] [--fix]` | 交付前检查:XML 声明 `standalone="yes"`、docProps 作者与旧页标题、模板示例文案残留、场景模板占位、通用占位符与规划标签。`--fix` 自动修 docProps |
| `scripts/office/soffice.py --headless --convert-to pdf deck.pptx` | LibreOffice 封装,自动配置运行环境,避免裸调 `soffice` 卡住 |
| `scripts/guard_templates.py` | hook 脚本,由 frontmatter 自动挂载,拦截写入原始模板的操作,不需要手动调用 |

## 用 pptxgenjs 补充内容——踩坑清单

只在**版式实例化之后、给正文区补内容**时用 pptxgenjs,不要拿它从零画整份公司 PPT。颜色/字号一律取 `COMPANY_STYLE.md`。先试 `require('pptxgenjs')`,失败再 `npm install pptxgenjs`。

- **加幻灯片之前先设 `pres.layout = "LAYOUT_WIDE"`**(13.33" × 7.5",公司模板尺寸)。默认 `LAYOUT_16x9` 只有 10" × 5.625",超界坐标会被写进去但不显示。
- **十六进制颜色不带 `#`、不用 8 位。** `"#FF0000"` 和 `"00000020"` 都会导致文件损坏。半透明:填充和图片用 `transparency: 0-100`,阴影用 `opacity: 0.0-1.0`,用错会被静默忽略。
- **pptxgenjs 会原地修改传入的 options 对象**(转成 EMU)。不要在两次 `add*` 之间共享 `shadow`/options 对象。
- **阴影 `offset` 必须 ≥ 0**,负值损坏文件;向上投影用 `angle: 270` 配正 offset。
- **`letterSpacing` 被静默忽略**,用 `charSpacing`。
- **列表:** 每项 `bullet: true`,不手打 `•`;除最后一项外都设 `breakLine: true`;段间距用 `paraSpaceAfter`,不用 `lineSpacing`。
- **每个输出文件一个 `new pptxgen()`**,不复用实例。
- **`rectRadius` 只对 `ROUNDED_RECTANGLE` 生效。**
- **不支持渐变填充**,改用渐变图片当背景。
- **文本框自带内部留白**,需要跟形状/线条对齐时设 `margin: 0`。
- **讲者备注写进 `slide.addNotes("...")`**,不写进页面文本框。
- **图表尽量原生 `addChart()`**(组合图传 `{type, data, options}` 数组);趋势线、误差线自己算额外系列或后处理 OOXML;只有桑基图、网络图、和弦图这类 PowerPoint 没有原生形式的才用图片。
- **默认图表很朴素**:设 `showTitle` + `title`、`showValue: true` + `dataLabelPosition`、`chartColors` 取 `COMPANY_STYLE.md` 配色,弱化坐标轴和网格线(`catAxisLabelColor`/`valAxisLabelColor`、`valGridLine: { color, size }`、`catGridLine: { style: "none" }`),单系列 `showLegend: false`。
- **堆叠柱状/条形图的 `dataLabelPosition` 只能是 `ctr`、`inEnd`、`inBase`**,`outEnd` 损坏文件。
- **用 `secondaryValAxis`/`secondaryCatAxis` 时,`valAxes` 和 `catAxes` 都要各给两条**,否则 PowerPoint 丢弃整个图表并报损坏。
- **`writeFile()` 之后跑 `validate.py`**,它能报出上面两类图表问题;去生成脚本里改,不手改打包好的 XML。
- **不要挪动 `<p:presentation>` 子元素顺序。** pptxgenjs 把 `<p:notesMasterIdLst>` 紧跟在 `<p:sldIdLst>` 后,PowerPoint 能读,挪动后反而打不开。
- **图标:** `react-icons` 渲染 SVG(`ReactDOMServer.renderToStaticMarkup`)→ `sharp` 栅格化到 ≥256px → `addImage({ data: "image/png;base64," + buf.toString("base64") })`,前缀 `image/png;base64,` 必须有。require 失败时 `npm install react-icons react react-dom sharp`。

## 编辑现有演示文稿和模板

先挑版式:`python3 scripts/thumbnail.py template.pptx template-thumbs` 生成缩略图网格(超过 12 页拆成 `template-thumbs-N.jpg`)。这只用于挑版式,视觉检查要用「转换成图片」的全分辨率渲染。`.potx` 先复制一份改成 `.pptx` 扩展名再跑缩略图。搭配 `markitdown` 把内容板块对应到模板页,换着用不同版式。

```bash
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" deck.pptx
python3 scripts/add_slide.py unpacked/ slide2.xml --after slide2.xml   # 会打印新页面路径
# 调整顺序/删除页面 = 改 ppt/presentation.xml 的 <p:sldIdLst>
python3 scripts/clean.py unpacked/
# 编辑 ppt/slides/slideN.xml
(cd unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)           # 先 rm 再打包,否则删掉的部件还留在包里
python3 scripts/office/validate.py out.pptx --original deck.pptx
python3 scripts/check_delivery.py out.pptx --template deck.pptx       # deck.pptx 本身是用户文件时省略 --template
```

- **增、删、调顺序全部做完,再编辑内容。** `add_slide.py` 原样复制文件,先改后复制会把改动一起带走;`clean.py` 会删掉不在 `<p:sldIdLst>` 里的页,包括刚写好的。
- **不要手动复制幻灯片文件**,用 `add_slide.py`。复制出的页仍*引用*源页的图表/SmartArt/嵌入对象部件,改一页会连带另一页。
- **`python-pptx` 的三个限制**:不能复制幻灯片(只有 `add_slide(layout)`);`text_frame.text = "..."` 会把段落坍缩成无样式 run(改赋值 `run.text`);读不了模板常用的 SVG/EMF(`add_picture` 抛 `UnidentifiedImageError`)。
- 老版 `.ppt` 先转:`python3 scripts/office/soffice.py --headless --convert-to pptx file.ppt`。`.potx` 解压打包方式相同,输出保留 `.potx` 扩展名。
- 复用模板里的图标或图片,复制一份已经包含它的幻灯片或版式。

### 填充模板时

- 写脚本做 XML 变换用 `defusedxml.minidom`;`xml.etree.ElementTree` 会重写命名空间前缀导致损坏。
- **`minidom` 的 `Document.toxml(encoding=...)` 会丢掉 `standalone="yes"`。** 缺了它,XSD、python-pptx、LibreOffice 都正常,但真实 PowerPoint 会静默"修复"并清空该页。改用:
  ```python
  body = dom.documentElement.toxml()
  with open(path, 'wb') as f:
      f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
      f.write(body.encode('utf-8'))
  ```
  `check_delivery.py` 会检查全部部件的声明。
- **模板槽位数 ≠ 内容条数。** 模板 4 个成员、你只有 3 个,删掉第 4 个的整个形状组(图片+文本框),不只删文字;QA 时检查孤立的视觉元素。
- 每条列表项一个 `<a:p>`,不把多条拼进一段。复制相邻 `<a:pPr>` 保留间距;标题、板块标题、行内标签的 `<a:rPr>` 加 `b="1"`。
- **安全的纯文字替换**:把新字符串整段放进该段落*第一个* `<a:r>` 的 `<a:t>`,同段其他 run 的文字清空,保留 run/段落数量和格式。注意**字数对齐不等于换行一致**:中英混排里的英文短语(如 `Data For AI`)不可拆分,卡在行尾时整体换行,替换文字更短也可能多出一行,压到标题下分隔线等固定装饰上,视觉上表现为"压线"或"内容缺失"。换行变了就明显缩短一截,重新渲染确认行数。
- 项目符号从版式继承,只在需要覆盖时加 `<a:buChar>`/`<a:buAutoNum>`/`<a:buNone>`,文字里不写 `•`。
- 文字前后有空格时,`<a:t>` 加 `xml:space="preserve"`。

## 设计约束

配色、字体、Logo 以 `COMPANY_STYLE.md` 为准。以下约束只管**新画的内容**;母版/版式自带的元素(如 slide1-3 标题下那条细分隔线,来自 `slideLayout3` 所属母版)是公司既定设计,不要改或删。

**版式**:每页要有视觉元素(图表、图标、形状、架构图),不做纯"标题+项目符号"页;相邻页换着用分栏、卡片、信息突出块、时间线、大号数字(60-72pt 配小标签)。

**字号**

| 元素 | 字号 |
|---|---|
| 幻灯片标题 | 36-44pt 加粗 |
| 板块标题 | 20-24pt 加粗 |
| 正文 | 14-16pt |
| 图注 | 10-12pt 弱化色 |

14-16pt 是信息量较高页面的正文基准。`LAYOUT.md` 判定为低密度页时,卡片标题 24-28pt、主要正文 18-22pt,并同步收缩卡片高度。

**间距**:页边距 ≥ 0.5";内容块之间固定用 0.3" 或 0.5",全篇一致。

**避免**

- 正文居中(段落和列表左对齐,只有标题居中)
- 低对比度元素(浅底浅字、深底深字、深色图标配深色背景)
- 在新画的标题下加装饰线,或加装饰性色条、侧边条、卡片单边强调条——想突出卡片用浅色底、投影或图标
- 米色/米黄色背景,没指定时用 `FFFFFF`
- 文字溢出形状:缩字号、拆页或放大容器,不能截断
- 只精心做一页、其余潦草

## 质量检查

第一次渲染通常有真实问题。按「检查 → 修复 → 重新检查」循环,直到三道检查全部通过才交付。

### 文件检查

```bash
python3 scripts/office/validate.py output.pptx --original src.pptx   # 从模板生成;从零生成时去掉 --original
python3 scripts/check_delivery.py output.pptx --template src.pptx    # 场景模板改用 --scaffold
```

**来自模板的 deck 一定传 `--original`。** 模板自带的 XSD 问题会被压制,不会淹没真正的问题;结构性检查(关系、内容类型、图表)不受 `--original` 影响,模板自带问题要按是非判断。

`check_delivery.py` 查的是 validate、markitdown、LibreOffice 都看不到的问题:`docProps/core.xml` 的作者带着内部邮箱或模板原作者、`docProps/app.xml` 留着旧页标题和旧页数、示例文案没删干净、XML 声明缺 `standalone`。docProps 类问题加 `--fix` 自动修;标为"需确认"的条目逐条判断是有意保留的标题还是没改的提示语。

### 内容检查

`markitdown output.pptx`,检查内容缺失、错别字、顺序错乱。

### 视觉检查

按「转换成图片」渲染后逐页看。盯着生成代码久了容易看到预期而非实际,所以用新鲜的眼光看图(有子 agent 时交给它效果更好)。要找的缺陷:

- **文字溢出或被容器/页面边界截断——先查这个**
- 元素重叠(文字穿过形状、线条穿过文字)、与 Logo 冲突
- 引用来源或页脚跟上方内容碰撞
- 元素间距 < 0.3",或距页边 < 0.5"
- 间距不均(一处空很大、一处很挤)、分栏未对齐
- 命中 `LAYOUT.md`「无效留白检测」任一阈值
- 替换文字后模板装饰错位(如标题从一行变两行,分隔线仍按单行定位)
- 文本框太窄导致过度换行
- 残留的占位符内容

发现问题就修,重新渲染改动过的页,再看一遍;直到这一轮没有新问题为止。

## 转换成图片

```bash
python3 scripts/office/soffice.py --headless --convert-to pdf output.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 output.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

把打印出的绝对路径直接交给看图工具。`rm` 清掉上一轮的旧图;`pdftoppm` 按总页数补零(10 页内 `slide-1.jpg`,10-99 页 `slide-01.jpg`)。每次修改后四条命令全部重跑,PDF 必须从改过的 `.pptx` 重新生成。

## 依赖

```bash
pip install "markitdown[pptx]" Pillow defusedxml lxml   # 文本提取、缩略图、清理、校验
npm install pptxgenjs                                   # 仅在 require 失败时
brew install --cask libreoffice                         # soffice,转 PDF(Linux: libreoffice)
brew install poppler                                    # pdftoppm,转图片(Linux: poppler-utils)
```

`check_delivery.py` 和 `guard_templates.py` 只用 Python 标准库(3.9+)。

## 测试

`evals/evals.json` 有 3 个测试任务。改动本 skill 后,在新会话里分别开启和关闭本 skill 各跑一遍,对比结果,不变差才保留改动。
