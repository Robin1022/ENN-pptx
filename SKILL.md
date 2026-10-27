---
name: enn-pptx-new-2
description: "用公司自有 PPT 模板生成/编辑内部汇报、方案材料等 .pptx 文件。触发场景：用户要求「用公司模板」「公司PPT」「内部汇报」「新奥」「新奥新智/ENNEW」相关的演示文稿制作,或明确引用 templates/company-template.pptx。与通用 pptx skill 的区别:这个 skill 默认套用公司品牌(配色/字体/Logo/版式),不做与品牌无关的通用/个人 PPT——那种场景交给通用 pptx skill。"
hooks:
  PreToolUse:
    - matcher: "Bash|Edit|Write|MultiEdit"
      hooks:
        - type: command
          command: "f=\"${CLAUDE_SKILL_DIR:-$HOME/.claude/skills/enn-pptx-new-2}/scripts/guard_templates.py\"; [ ! -f \"$f\" ] || python3 \"$f\""
---

# 公司 PPTX 生成、编辑与分析

## 不可省的规则

1. **新建 deck 先交整体规划,用户确认后才动手**(`PLANNING.md`)。素材缺口自主推理补全,但推测内容必须标 `[推理补充]` / `[待确认假设]` / `[示意数据]`,不得写成用户提供的事实。这些标记只写在规划和回复里,不写到幻灯片上(见「内容来源」)。
2. **模板原文件只读。** 一律 `cp` 出工作副本再操作;写入 `templates/` 下模板的命令会被 hook 拦截。
3. **成稿不得带模板示例文案。** 所有 `add_slide.py` 复制完成后,从 `<p:sldIdLst>` 移除示例页,再跑 `clean.py`。
4. **交付前四道检查全过才算完成**:`validate.py --original`、`check_delivery.py`、`check_layout.py`、渲染后的视觉 QA。任何一道失败,修完从失败的那一步重跑。

## 配套文档

动手前按任务读对应文档:

- [`PLANNING.md`](PLANNING.md) —— 页数、叙事、逐页内容与布局规划 + 确认关卡
- [`RESEARCH.md`](RESEARCH.md) —— 是否外部检索、来源验证、事实标记与引用
- [`LAYOUT.md`](LAYOUT.md) —— 按内容密度分配区域、选网格、改造模块、检测无效留白
- [`COMPANY_STYLE.md`](COMPANY_STYLE.md) —— 画布/配色/字体/Logo/版式一览
- [`MODULES.md`](MODULES.md) —— 三种可选的结构化正文原型
- [`templates/README.md`](templates/README.md) —— 场景模板逐页结构
- [`PPTXGENJS.md`](PPTXGENJS.md) —— 用 pptxgenjs 给正文区补画内容时的踩坑清单
- [`EXPERIENCE.md`](EXPERIENCE.md) —— 历次优化的依据和结论,改这个 skill 之前先读
- [`README.md` 的「模板适配」](README.md#模板适配) —— 换用其他板块模板时的重测步骤(`COMPANY_STYLE.md` 的数值只对新奥新智模板成立)

本 skill 是通用 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 的公司分支,脚本和踩坑清单继承自官方版本。下文路径都相对本 skill 目录。

## 任务分流

| 任务 | 方法 |
|---|---|
| **场景已有现成模板**(见下表) | 复制 `templates/` 下的场景模板改文字,**不走** `company-template.pptx` 实例化流程——那会把做好的版面重做一遍 |
| **新建公司风格演示文稿** | 走下面的默认工作流 |
| **修改现有 PPT 的某几页(最常见)** | 见「修改现有页面」:改前备份,改完校验和几何检查,写回原文件 |
| **读取内容** | `markitdown deck.pptx`(每页一个 `<!-- Slide number: N -->` 区块);缩略图:`python3 scripts/thumbnail.py deck.pptx <prefix>` |

用户没有要求公司品牌/模板、且是跟公司汇报无关的私人演示文稿时,改用通用 `pptx` skill。

### 现成的场景模板

| 模板 | 命中场景 | 页数 |
|---|---|---|
| `templates/campus-hire-defence.pptx` | 校招生转正答辩 / 新进人员任前答辩 / 试用期转正述职 | 14 |

`cp templates/campus-hire-defence.pptx <目标路径>` 后逐页替换文字;改动时同样适用「填充模板时」的全部注意事项,尤其是**填充元素只能替换不能新增**(`<p:spPr>` 的填充是互斥选择组,插出两个并列的 `<a:solidFill>` 会让 PowerPoint 报损坏,而 XSD、LibreOffice、python-pptx 都不报错)。交付前检查用 `--scaffold` 指向这份场景模板。

## 内容来源

用户给了材料(文档、已有页面、口述要点)时,页面上的痛点、技术点、方法和组件名都以材料为准。

- 材料之外补充的内容(推理补全、外部检索、示意数据、你认为该有的技术点)可以写,但每次改完都在回复里逐条列出:补了什么、写在哪、依据是什么。用户问"材料里有没有"时,引用材料原文整句回答;没有就直说是补充的。
- 推理、假设、示意、来源这些标记只写在规划和回复里。幻灯片上不出现标记、来源脚注、注释和状态括注,用户要求标注时才加。

## 文案

这类 PPT 用于技术汇报和评审,页面文字按正式的书面技术语言写:

- 直接陈述做法、机制和结论。不用设问式标签、装饰性引号、口号式短语,也不用口语和网络化表达。
- 每句话落到具体的对象和动作上,读者不靠讲解也能看懂。概括性的说法后面要跟上它具体指什么。
- 术语沿用材料和已有页面的叫法。新引入的术语要么当页解释清楚,要么换成材料里已有的说法;同一个概念全页只用一种叫法。

## 修改现有页面

用户最常见的用法是拿一份已有的 PPT,围绕其中一两页反复修改。每一轮:

1. **确定写到哪个文件。** 默认原地写回用户给的那份文件,不另存新版本。用户说"新的 ppt""单起一个""新加一页"时,才另存文件或用 `add_slide.py` 加页;新加的页放在被参照页的后面,原页保持不动。
2. **动手前备份。** 第一次修改前把原文件复制到工作目录,作为 `--original` 的基线和出错时的回退版本。
3. **改内容。** 解压 → 改 `ppt/slides/slideN.xml` → 打包,写法见「填充模板时」。用生成脚本重画整块区域时,传给 XML 的每个属性值都要是 schema 允许的取值(对齐、换行、锚点这类枚举不能传数字或 `None`)。
4. **检查后再写回。**
   ```bash
   python3 scripts/office/validate.py out.pptx --original backup.pptx
   python3 scripts/check_delivery.py out.pptx
   python3 scripts/check_layout.py out.pptx --original backup.pptx --slides N
   ```
   然后只渲染改动的页(「转换成图片」),对照几何检查的结果逐条看。都处理完再覆盖原文件。
5. **回复里说清楚**改了哪一页的哪些地方,以及按「内容来源」列出这轮补充的内容。

## 默认工作流:用公司模板新建

把这份清单复制到回复里,逐项打勾:

```text
- [ ] 1. 规划:按 PLANNING.md 出整体规划(含 RESEARCH.md 的检索决策、LAYOUT.md 的逐页密度判断),用户确认
- [ ] 2. 建页:cp 模板 → 解压 → add_slide.py 建出全部目标页
- [ ] 3. 定稿结构:调整 <p:sldIdLst> 顺序,移除示例页引用,跑 clean.py
- [ ] 4. 填内容:改 ppt/slides/slideN.xml,或用 pptxgenjs 重搭不匹配的区域(见 PPTXGENJS.md)
- [ ] 5. 打包 + validate.py --original          失败 → 回到第 4 步修
- [ ] 6. check_delivery.py --template           失败 → docProps 类加 --fix;文案残留回到第 4 步
- [ ] 7. check_layout.py --original             对照渲染图确认是可见问题 → 回到第 4 步
- [ ] 8. 渲染成图片做视觉 QA                   有缺陷 → 回到第 4 步,修完从第 5 步重跑
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

# 5-7. 打包与检查
(cd /tmp/unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)
python3 scripts/office/validate.py /tmp/out.pptx --original templates/company-template.pptx
python3 scripts/check_delivery.py /tmp/out.pptx --template templates/company-template.pptx
python3 scripts/check_layout.py /tmp/out.pptx --original templates/company-template.pptx
```

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/thumbnail.py deck.pptx [prefix]` | 每页带标签的缩略图网格,用来挑版式。只认 `.pptx`。**一定传 `prefix`**——默认 `thumbnails`,同目录处理另一份 deck 会覆盖 |
| `scripts/add_slide.py unpacked/ slide2.xml [--after slideN.xml]` | 复制一页(或从 `slideLayoutN.xml` 新建),自动处理包级注册,并摘掉 `notesSlide`/`tags` 引用。两页形状指向同一个 `<p:tags>` 部件时,真实 PowerPoint 会静默清空整页,而 XSD/python-pptx/LibreOffice 都看不出来——自己手写复制逻辑时要防这个坑。直接处理 `.pptx` 时**必须传 `-o out.pptx`**,否则原地改写输入 |
| `scripts/clean.py unpacked/` | 删除不再被引用的幻灯片、媒体和关系。`<p:sldIdLst>` 定稿**之后**运行 |
| `scripts/office/validate.py deck.pptx [--original src.pptx]` | Schema、关系、内容类型、图表和幻灯片检查,每条失败附修复方法。模板派生的 deck 必须传 `--original` |
| `scripts/check_delivery.py deck.pptx [--template ...] [--scaffold ...] [--fix]` | 交付前检查:XML 声明 `standalone="yes"`、docProps 作者与旧页标题、模板示例文案残留、场景模板占位、通用占位符与规划标签。`--fix` 自动修 docProps |
| `scripts/check_layout.py deck.pptx [--original src.pptx] [--slides 5,6]` | 版面几何检查:形状超出画布、图片或方块压到边框、箭头短到只剩箭头头部,以及按字宽估算的文字溢出、文字压边框、连线穿过文字。结果要对照渲染图判断,见「版面几何检查」 |
| `scripts/office/soffice.py --headless --convert-to pdf deck.pptx` | LibreOffice 封装,自动配置运行环境,避免裸调 `soffice` 卡住 |
| `scripts/guard_templates.py` | hook 脚本,由 frontmatter 自动挂载,拦截写入原始模板的操作,不需要手动调用 |

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

第一次渲染通常有真实问题。按「检查 → 修复 → 重新检查」循环,直到四道检查全部通过才交付。

### 文件检查

```bash
python3 scripts/office/validate.py output.pptx --original src.pptx   # 从模板生成;从零生成时去掉 --original
python3 scripts/check_delivery.py output.pptx --template src.pptx    # 场景模板改用 --scaffold
```

**来自模板的 deck 一定传 `--original`。** 模板自带的 XSD 问题会被压制,不会淹没真正的问题;结构性检查(关系、内容类型、图表)不受 `--original` 影响,模板自带问题要按是非判断。

`check_delivery.py` 查的是 validate、markitdown、LibreOffice 都看不到的问题:`docProps/core.xml` 的作者带着内部邮箱或模板原作者、`docProps/app.xml` 留着旧页标题和旧页数、示例文案没删干净、XML 声明缺 `standalone`。docProps 类问题加 `--fix` 自动修;标为"需确认"的条目逐条判断是有意保留的标题还是没改的提示语。

### 内容检查

`markitdown output.pptx`,检查内容缺失、错别字、顺序错乱。

### 版面几何检查

```bash
python3 scripts/check_layout.py output.pptx --original src.pptx --slides 5   # 只查改过的页
```

从坐标算出形状超出画布、图片或方块压到边框、箭头短到只剩箭头头部,再按字宽估算文字溢出、文字压边框、连线穿过文字。渲染图上这类问题小而分散,用户反复指出过,所以每次改完先跑这一步。`--original` 传改动前的文件或模板,原来就有的不再列出。

列出来的是几何事实,不一定是缺陷:有意骑在边框上的页签和编号、刻意画成三角的小箭头也会出现。每一条都到渲染图里核对,是用户能看到的问题才改。

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
