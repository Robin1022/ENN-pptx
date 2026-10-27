---
name: enn-pptx-new
description: "用公司自有 PPT 模板生成/编辑内部汇报、方案材料等 .pptx 文件。触发场景：用户要求「用公司模板」「公司PPT」「内部汇报」「新奥新智/ENNEW」相关的演示文稿制作,或明确引用 templates/company-template.pptx。与通用 pptx skill 的区别:这个 skill 默认套用公司品牌(配色/字体/Logo/版式),不做与品牌无关的通用/个人 PPT——那种场景交给通用 pptx skill。"
---

# 公司 PPTX 生成、编辑与分析

这是通用 [pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) 的公司专属分支,底层机制(pptxgenjs 踩坑清单、XML 直接编辑、校验器)完全继承自官方版本,新增的是**默认套用公司模板** `templates/company-template.pptx`。生成时不是逐页套模块,而是先对用户素材进行整体叙事规划,自主判断是否需要外部知识检索,再根据每页的内容量和信息层级自适应选择版式或模块。配套文档按任务读:修改现有页面一般只需要 `COMPANY_STYLE.md`,新建整份 deck 时前五份都要读:

- [`PLANNING.md`](PLANNING.md) —— 整份 PPT 的页数、叙事、逐页内容和布局规划 + 生成前的整体确认关卡
- [`RESEARCH.md`](RESEARCH.md) —— 外部知识检索决策、来源验证、事实标记和引用要求
- [`LAYOUT.md`](LAYOUT.md) —— 根据内容密度分配区域、选择网格、改造模块和检测无效留白
- [`COMPANY_STYLE.md`](COMPANY_STYLE.md) —— 视觉规范(画布/配色/字体/Logo/版式一览)
- [`MODULES.md`](MODULES.md) —— 三种可选的结构化正文原型;由模型根据内容和主题自主判断是否使用
- [`GENERIC.md`](GENERIC.md) —— 不套公司模板时的通用配色/字体规范,以及 pptxgenjs 踩坑清单

`.pptx` 本质是 ZIP 打包的 XML 文件集合。按任务选方法:

| 任务 | 方法 |
|---|---|
| **新建公司风格演示文稿** | 先按 `PLANNING.md` 分析用户全部内容,规划总页数、叙事顺序、每页核心结论、内容和布局,并让用户一次性确认整体方案;然后由模型为每页自主判断是否使用 `MODULES.md` 的三种模块,或选用其他公司版式 |
| **修改现有 PPT 的某几页(最常见)** | 按[修改现有页面](#修改现有页面):改前备份,改完校验和几何检查,写回原文件 |
| **读取内容** | `markitdown deck.pptx`(每页一个 `<!-- Slide number: N -->` 区块);可视化缩略图:`python scripts/thumbnail.py deck.pptx` |

**不确定是不是该用这个 skill?** 只要用户没有明确要求公司品牌/模板,且这是个跟公司汇报无关的临时/私人演示文稿,改用通用 `pptx` skill——那边的配色/字体规范是通用设计建议,不会被这里的公司规范覆盖。

## 内容来源

用户给了材料(文档、已有页面、口述要点)时,页面上的痛点、技术点、方法和组件名都以材料为准。

- 材料之外补充的内容(推理补全、外部检索、示意数据、你认为该有的技术点)可以写,但每次改完都在回复里逐条列出:补了什么、写在哪、依据是什么。用户问"材料里有没有"时,引用材料原文整句回答;没有就直说是补充的。
- 推理、假设、示意、来源这些标记只写在规划和回复里。幻灯片上不出现标记、来源脚注、注释和状态括注,用户要求标注时才加。

## 文案

这类 PPT 用于技术汇报和评审,页面文字按正式的书面技术语言写:

- 直接陈述做法、机制和结论。不用设问式标签、装饰性引号、口号式短语,也不用口语和网络化表达。
- 每句话落到具体的对象和动作上,读者不靠讲解也能看懂。概括性的说法后面要跟上它具体指什么。
- 术语沿用材料和已有页面的叫法。新引入的术语要么当页解释清楚,要么换成材料里已有的说法;同一个概念全页只用一种叫法。

## 脚本

路径都是相对这个 skill 目录的。其他一律是普通的 Python、`node` 或 shell。

| 脚本 | 作用 |
|---|---|
| `scripts/thumbnail.py deck.pptx [prefix]` | 生成每页幻灯片的带标签缩略图网格,用来挑模板版式。只认 `.pptx`。记得传 `prefix` 参数——默认是 `thumbnails`,同一目录下处理另一份 deck 会把之前的缩略图覆盖掉 |
| `scripts/add_slide.py unpacked/ slide2.xml [--after slideN.xml]` | 复制一页幻灯片(或从 `slideLayoutN.xml` 新建),自动处理全部包级注册工作——包括摘掉 `notesSlide`/`tags` 引用,让复制出来的页面不会跟源页共享讲者备注或形状级自定义数据标签。两个不同页面的形状指向同一个 `<p:tags r:id="...">` 部件,对 XSD/python-pptx/LibreOffice 完全不可见,但真实 PowerPoint 的严格解析器会拒绝它,静默清空*整页内容*——如果你自己手写 XML 变换来复制形状/关系而不是用这个脚本,要留意同一类坑。也可以直接对 `.pptx` 操作,加 `-o out.pptx` |
| `scripts/clean.py unpacked/` | 删除不再被引用的幻灯片、媒体文件和关系条目。要在 `<p:sldIdLst>` 定稿**之后**运行 |
| `scripts/office/validate.py deck.pptx [--original src.pptx]` | Schema、关系、内容类型、图表和幻灯片检查;每条失败都会给出具体修复方法。任何从模板派生的 deck 都要传 `--original`——它会把 schema 检查基线定在模板本身,这样模板自带的 XSD 错误不会被算到你头上 |
| `scripts/check_layout.py deck.pptx [--original src.pptx] [--slides 5,6]` | 版面几何检查:形状超出画布、图片或方块压到边框、箭头短到只剩箭头头部,以及按字宽估算的文字溢出、文字压边框、连线穿过文字。结果要对照渲染图判断,见[版面几何检查](#版面几何检查) |
| `scripts/office/soffice.py --headless --convert-to pdf deck.pptx` | LibreOffice 的封装——在这个沙箱环境里直接裸调 `soffice` 会卡住 |

## 修改现有页面

用户最常见的用法是拿一份已有的 PPT,围绕其中一两页反复修改。每一轮:

1. **确定写到哪个文件。** 默认原地写回用户给的那份文件,不另存新版本。用户说"新的 ppt""单起一个""新加一页"时,才另存文件或用 `add_slide.py` 加页;新加的页放在被参照页的后面,原页保持不动。
2. **动手前备份。** 第一次修改前把原文件复制到工作目录,作为 `--original` 的基线和出错时的回退版本。
3. **改内容。** 解压 → 改 `ppt/slides/slideN.xml` → 打包,写法见[填充模板时](#编辑现有演示文稿和模板)。用生成脚本重画整块区域时,传给 XML 的每个属性值都要是 schema 允许的取值(对齐、换行、锚点这类枚举不能传数字或 `None`)。
4. **检查后再写回。**
   ```bash
   python3 scripts/office/validate.py out.pptx --original backup.pptx
   python3 scripts/check_layout.py out.pptx --original backup.pptx --slides N
   ```
   然后只渲染改动的页([转换成图片](#转换成图片)),对照几何检查的结果逐条看。都处理完再覆盖原文件。
5. **回复里说清楚**改了哪一页的哪些地方,以及按[内容来源](#内容来源)列出这轮补充的内容。

## 用公司模板新建演示文稿

先读 [`COMPANY_STYLE.md`](COMPANY_STYLE.md) 确认版式怎么选、配色怎么跟内容类型走。

**Step 0 —— 整份 PPT 规划与确认(新建 deck 必做):** 先完整阅读用户提供的素材,按 [`PLANNING.md`](PLANNING.md) 确定演示目标、受众、核心结论、总页数和叙事顺序。按 [`RESEARCH.md`](RESEARCH.md) 显式判断是否需要外部检索:需要时先完成检索、验证和来源记录,再定稿规划;不需要时说明原因后继续。素材缺失时要基于上下文主动推理补全,在规划里用 `[推理补充]`、`[待确认假设]` 或 `[示意数据]` 明确标记;外部检索获得的内容用 `[外部资料]` 标记并记录来源。不得把推测写成用户已提供的真实事实。这些标记只出现在规划和回复里,不写到幻灯片上(见[内容来源](#内容来源))。对每页再按 [`LAYOUT.md`](LAYOUT.md) 评估信息任务、条目数、文字量、主次和预计占用高度,得出区域比例、网格和模块改造方案。然后输出逐页规划:页码、页标题、该页要证明的一句话、主要内容、布局/可视化方式、区域比例/网格、版式或模块选择、推理补充/假设/外部来源。用户确认整体规划后才开始制作。

三个示例正文页只是可选原型,不是用户必须选择的三选一。模型要根据每页内容的语义和主题自主判断:匹配反思/技术图谱/技术方案详情原型时,复制对应源页并按内容调整;不匹配时,选用 `COMPANY_STYLE.md` 的其他版式并使用公司图形语言搭建。不要为了复用模板而把内容硬塞进三个模块,也不要要求用户逐页挑选模块。

基本流程(Step 0 确认完之后):

```bash
# 1. 解压公司模板(不要直接改 templates/ 下的原始文件,拷出来一份工作副本)
cp templates/company-template.pptx /tmp/deck.pptx
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('/tmp/unpacked')" /tmp/deck.pptx

# 2. 按已确认的整体规划创建页面。匹配三种可选模块的页 = 复制对应 slide1/2/3.xml(见 MODULES.md);
#    封面/目录/其他版式 = 从对应的 slideLayoutN.xml 实例化。这一步自动处理
#    Content_Types / presentation.xml.rels / <p:sldIdLst>,Logo 也随版式/母版自动带出来
python scripts/add_slide.py /tmp/unpacked/ slide1.xml --after slide3.xml   # 例:反思模块,复制 slide1.xml
# python scripts/add_slide.py /tmp/unpacked/ slideLayout3.xml --after slide3.xml   # 例:非三模块页面,从"正文页"版式新开

# 3. 把所有目标页都创建好后,先定稿 <p:sldIdLst>:
#    - 调整最终页面顺序;
#    - 从 <p:sldIdLst> 移除模板自带的 slide1/2/3 示例页引用,除非用户明确要保留某一页示例;
#    - 只保留最终成稿需要的页。不要直接删除 ppt/slides/slide1.xml 等文件,由 clean.py 统一清理。
python scripts/clean.py /tmp/unpacked/   # sldIdLst 定稿后,清理已无引用的示例页、媒体和关系

# 4. 再编辑最终保留页的内容——
#    内容结构跟源页面接近 = 直接改 ppt/slides/slideN.xml 做文字替换(见"填充模板时"一节的安全替换技巧);
#    内容的栏数/条目数/层级跟源页面对不上 = 在保留强调色和图形语言(COMPANY_STYLE.md"常见图形语言")的前提下,
#    增删对应的形状子组,或用 pptxgenjs 重新搭建那部分区域——不要为了省事把内容硬塞进不匹配的结构里

# 5. 打包 + 校验(--original 指向公司模板,过滤掉模板自身的 XSD 已知问题)
(cd /tmp/unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)
python scripts/office/validate.py /tmp/out.pptx --original templates/company-template.pptx
python3 scripts/check_layout.py /tmp/out.pptx --original templates/company-template.pptx
```

原始的 3 个示例正文页(slide1-3)只是**工作阶段的复制源和视觉参照**, 不是最终成稿内容。必须先完成所有 `add_slide.py` 复制操作,再将未被用户明确要求保留的示例页从 `<p:sldIdLst>` 移除,然后运行 `clean.py`。最终交付的 deck 不得混入模板示例文案。三个模块各自的槽位定义、复用边界见 [`MODULES.md`](MODULES.md)。

## 编辑现有演示文稿和模板

先挑版式:`python scripts/thumbnail.py template.pptx template-thumbs` 会生成每页幻灯片的带标签缩略图网格,并打印生成了哪些文件——`template-thumbs.jpg`,超过 12 页会拆成 `template-thumbs-N.jpg`。**第二个参数一定要传,按这份 deck 命名。** 它默认是 `thumbnails`,同一目录下处理两份 deck 会互相覆盖对方的缩略图——第一份的直接就没了(这里只是做模板分析——真正的视觉检查需要[转换成图片](#转换成图片)里的全分辨率渲染图;缩略图脚本只认 `.pptx`,所以要先把 `.potx` 复制一份改成 `.pptx` 扩展名)。搭配 `markitdown` 使用,把每个内容板块对应到模板里的某一页,并且要换着用不同版式——不要把所有内容都塞进同一种"标题+项目符号"版式里。

```bash
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" deck.pptx
python scripts/add_slide.py unpacked/ slide2.xml --after slide2.xml   # 复制一页幻灯片(或 slideLayoutN.xml);会打印新页面的路径
# 调整顺序/删除页面 = 改 ppt/presentation.xml 里的 <p:sldIdLst>
python scripts/clean.py unpacked/                                     # 删除页面之后:清理孤立的 slide/media/rels
# 在 ppt/slides/slideN.xml 里编辑幻灯片内容
(cd unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)           # 从目录内部打包;先 rm 再打包,不然删掉的部件还会留在压缩包里
python scripts/office/validate.py out.pptx --original deck.pptx
```

- **所有结构性操作——增、删、调顺序——都要在编辑任何一页的内容之前做完。** `add_slide.py` 是原样复制幻灯片文件,如果先编辑了内容再复制,会把编辑过的内容也一起复制走;而 `clean.py` 会删掉任何不在 `<p:sldIdLst>` 里的幻灯片,包括你刚写好的那一页。
- **永远不要手动复制幻灯片文件**——`add_slide.py` 会做新页面需要的每一项注册工作,并且报告它做了什么(`Created ppt/slides/slide17.xml from slide2.xml`)。它也能直接对文件操作:`add_slide.py deck.pptx slide2.xml -o out.pptx`——**记得传 `-o`,不然会原地重写输入的 deck。** 复制出来的页面依然*引用*源页面的图表/SmartArt/嵌入对象部件,而不是各自克隆一份,所以改一页的图表会连带改到另一页。
- **如果用 `python-pptx`**,有三件事它做不到:复制幻灯片(它唯一的入口是 `add_slide(layout)`)、通过 `text_frame.text = "..."` 保留格式(这会把整个段落坍缩成一个无样式的 run——应该改赋值 `run.text`)、读取大多数模板美术素材用的 SVG/EMF 格式(`add_picture` 会抛出 `UnidentifiedImageError`)。
- 老版本的 `.ppt` 要先转换:`python scripts/office/soffice.py --headless --convert-to pptx file.ppt`。`.potx` 模板的解压打包方式完全一样——输出文件保留 `.potx` 扩展名。
- 要复用模板里的图标或图片,复制一份已经包含它的幻灯片或版式。

填充模板时:

- 如果要写脚本做 XML 变换,用 `defusedxml.minidom` 解析——用 `xml.etree.ElementTree` 来回解析/序列化 OOXML 会重写命名空间前缀,导致文件损坏。
- **`minidom` 自己的 `Document.toxml(encoding=...)` 会静默丢掉 XML 声明里的 `standalone="yes"`。** 每一个由 PowerPoint/WPS 写出的 OOXML 部件都带 `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;缺了 `standalone="yes"` 的部件依然能通过 XSD 校验、在 `python-pptx` 里正常打开、在 LibreOffice 里正常渲染——但真实 PowerPoint 的严格解析器会把它当成损坏文件,静默"修复"并清空那一页的内容(表现为空白页,除非你注意到修复对话框,否则完全看不出报错)。这个 skill 自带的所有 QA 步骤(`validate.py`、`markitdown`、经 LibreOffice 的视觉检查)都抓不到这个问题——只有在真实 PowerPoint 里才会暴露。修复方法:不要调用 `Document.toxml(encoding=...)`;改用 `dom.documentElement.toxml()` 拿到不带声明的正文,自己在前面拼上正确的声明:
  ```python
  body = dom.documentElement.toxml()
  with open(path, 'wb') as f:
      f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
      f.write(body.encode('utf-8'))
  ```
  任何基于 minidom 的编辑之后,先做字节层面的健全性检查,确认文件头跟没被动过的同级部件一致(`head -c 60 file.xml | xxd`),再相信这份文件在真实 PowerPoint 里能正常打开。
- **模板槽位数 ≠ 你的内容条数。** 如果模板展示 4 个团队成员,而你只有 3 个,要删掉第 4 个成员的整个形状组(图片+文本框),不能只删文字——然后在 QA 阶段检查有没有留下孤立的视觉元素。
- 每条列表项对应一个 `<a:p>`——不要把多条内容拼进一个段落。复制相邻的 `<a:pPr>` 以保留间距,标题、板块标题、行内标签(`Status:`、`Owner:`)的 `<a:rPr>` 上要加 `b="1"`。
- **安全的纯文字替换方式**:要在不动形状/位置的前提下替换一个段落的内容,把完整的新字符串放进该段落*第一个* `<a:r>` 的 `<a:t>` 里,同一段落里其他所有 run 的文字清空——这样能保留 run/段落的数量(以及 schema 要求的其他兄弟节点),不用重建 run 级别的格式。但**替换文字的字数跟原文对齐,不代表换行结果也会一致**——一段中英文混排的文字如果夹了一个英文/拉丁短语(比如 `Data For AI`、产品名等),这个短语会被当成一个不可拆分的整体来换行;如果它恰好卡在行尾附近,整个短语会被推到下一行,让上一行留白没填满,所以哪怕替换文字比原文明显*更短*,依然可能比原文多换出一行。这类溢出会跟固定位置的装饰元素(标题下的分隔线、旁边的形状)重叠——这些装饰不会随着文本框变高而跟着挪动,所以在视觉检查里表现为"文字被压线"或"内容缺失",而不是一眼看出来的溢出。如果一段中英混排的文字改完之后换行结果变了,不要死抠跟原文字数对齐——直接明显缩短一截(不是一两个字的程度),然后重新渲染确认行数真的对上了,因为字数对齐在这里并不能可靠地预测换行是否一致。
- 项目符号让它从版式继承;只有需要覆盖默认行为时才加 `<a:buChar>`、`<a:buAutoNum>`(编号列表)或 `<a:buNone>`——文字里永远不要写字面的 `•`。
- 文字前后有空格时,对应的 `<a:t>` 需要加 `xml:space="preserve"`。

## 设计思路

**配色/字体/Logo 以 `COMPANY_STYLE.md` 为准。** 不套公司模板时的通用配色表和字体安全列表在 [`GENERIC.md`](GENERIC.md)。下面几节不分场景,公司模板页面也要遵守——**但仅约束你新画的内容**。模板母版/版式自带、原样继承下来的设计元素(比如标题下方那条细分隔线,`slide1.xml`/`slide2.xml`/`slide3.xml` 三页都有,来自 `slideLayout3` 所属母版的既有设计)不受这份清单约束,不要因为清单里写了"标题下永远不要用装饰线"就去改或删这类模板自带元素——那是公司已经定下来的设计,不是 AI 生成时加的装饰。

**不要做无聊的幻灯片。** 白底配纯文字项目符号打动不了任何人。每一页都可以从下面这份清单里找灵感。

### 每一页的设计

**每一页都需要一个视觉元素**——图片、图表、图标或形状。纯文字的页面很容易被遗忘。

**版式选项:**
- 双栏(左侧文字,右侧插图)
- 图标+文字行(彩色圆圈里的图标,加粗标题,下面配说明)
- 2x2 或 2x3 网格(一侧放图片,另一侧是内容块网格)
- 半通栏图片(左侧或右侧铺满)配文字叠加

**数据展示:**
- 大号统计数字(60-72pt 的大数字,下面配小号标签)
- 对比列(前后对比、优缺点、并列选项)
- 时间线或流程图(编号步骤、箭头)

**视觉细节:**
- 板块标题旁边配小号彩色圆圈图标
- 关键数据或标语用斜体强调

### 字号

| 元素 | 字号 |
|---------|------|
| 幻灯片标题 | 36-44pt 加粗 |
| 板块标题 | 20-24pt 加粗 |
| 正文 | 14-16pt |
| 图注 | 10-12pt 弱化色 |

`14-16pt` 是信息量较高页面的正文基准,不是不分场景的固定字号。按 `LAYOUT.md` 判定为低密度页时,卡片标题使用 24-28pt,主要正文使用 18-22pt,并同步收缩卡片高度;不得用小字号配合超高空容器。

### 间距

- 最小边距 0.5"
- 内容块之间 0.3-0.5"
- 留足呼吸空间——不要填满每一寸

### 避免踩的坑(常见错误)

- **不要重复用同一种版式**——不同页之间要换着用分栏、卡片和信息突出块
- **不要让正文居中**——段落和列表要左对齐;只有标题才居中
- **不要吝啬字号对比**——标题要 36pt 以上才能跟 14-16pt 的正文形成反差
- **不要随意混用间距**——固定选 0.3" 或 0.5" 的间隔并保持一致
- **不要只精心设计一页、其余页潦草了事**——要么全篇都下功夫,要么全篇都保持简洁
- **不要做纯文字页**——加图片、图标、图表或其他视觉元素;避免"标题+项目符号"的单调组合
- **不要忘记文本框的内部留白**——需要文字边缘跟形状/线条对齐时,给文本框设 `margin: 0`,或者相应地偏移形状位置
- **不要用低对比度元素**——图标和文字都要跟背景形成强对比;避免浅色背景配浅色文字、深色背景配深色文字
- **永远不要在标题下加装饰线**——这是 AI 生成幻灯片的标志性特征;用留白或背景色来做区隔
- **永远不要加装饰性色条或强调条**——包括:横跨页面宽度的页眉/页脚色条、沿页面一侧的竖向侧边条、卡片或内容块一侧的细长强调条、矩形的"单边框"。这些都会让人一眼看出是 AI 生成的填充物。想让某张卡片突出,用浅色背景色调、投影或图标——不要用边缘色条。
- **不要默认用米色/米黄色背景**——没有指定背景时用白色(`FFFFFF`)或用户的品牌配色;避免暖色调的默认色,比如 `F5F5DC`、`FAF0E6`、`FAEBD7`、`FFF8E1`
- **不要让文字溢出形状**——如果文字放不下,缩小字号、拆到多页,或者放大容器;绝不能让内容被截断或溢出边界

## 质量检查(必须做)

第一次渲染出来通常会有一些真实问题——重叠、溢出、错位。找出来并修好,只重新渲染改动过的那几页,然后停下。

### 内容检查

```bash
markitdown output.pptx
```

检查有没有内容缺失、错别字、顺序错乱。

**用模板的场景下,检查有没有残留的占位符文字:**

```bash
markitdown output.pptx | grep -iE "\bx{3,}\b|lorem|ipsum|\bTODO|\[insert|this.*(page|slide).*layout"
```

如果 grep 有结果,先修好再宣布完成。

### 文件检查(必须)

```bash
python scripts/office/validate.py output.pptx                      # 从零生成的场景
python scripts/office/validate.py output.pptx --original src.pptx  # 从模板生成的场景
```

**如果 deck 来自模板,一定要传 `--original`。** 模板本身可能就含有 XSD 会拒绝的部分,裸跑一遍可能会报出不是你造成的失败——真正的问题反而可能藏在这堆误报里。`--original` 会把 schema 和幻灯片检查的基线定在模板本身,压制掉模板自带的错误。结构性检查——关系、内容类型、图表——不受 `--original` 影响,不管有没有传都会报出模板自带的问题,这类问题要按它本身的是非来看。

pptxgenjs 会生成 PowerPoint 拒绝打开、但其他所有工具都能接受的图表 XML:python-pptx 能打开这些 deck,LibreOffice 能渲染,XSD 也能通过。每条失败都会给出具体修复方法。去生成脚本里修,然后重新生成。

### 版面几何检查

```bash
python3 scripts/check_layout.py output.pptx --original src.pptx --slides 5   # 只查改过的页
```

从坐标算出形状超出画布、图片或方块压到边框、箭头短到只剩箭头头部,再按字宽估算文字溢出、文字压边框、连线穿过文字。渲染图上这类问题小而分散,用户反复指出过,所以每次改完先跑这一步。`--original` 传改动前的文件或模板,原来就有的不再列出。

列出来的是几何事实,不一定是缺陷:有意骑在边框上的页签和编号、刻意画成三角的小箭头也会出现。每一条都到下一步的渲染图里核对,是用户能看到的问题才改。

### 视觉检查

把幻灯片转换成图片(见[转换成图片](#转换成图片))并逐页检查。盯着生成代码看久了,容易看到自己预期的东西而不是实际渲染出来的东西,所以要用新鲜的眼光看这些图片(如果有子 agent 可用,交给它做效果不错)。要找的用户可见缺陷:

- **文字溢出或者在容器/幻灯片边界被截断——先查这个。** 这是最常见、也永远是用户可见的缺陷。(对于 `GENERIC.md` 字体一节标注为渲染不可靠的字体,预览只是近似值:相信你留出的约 10% 余量,不要相信预览看起来是否放得下。)
- 元素重叠(文字穿过形状、线条穿过文字、元素堆叠)
- 引用来源或页脚跟上方内容碰撞
- 元素靠得太近(间距 < 0.3")或卡片/板块几乎贴在一起
- 间距不均匀(一处留白很大,另一处很拥挤)
- 主要容器的实际内容占用高度低于 45%,并在容器下方留下超过 35% 的连续空白——这通常是固定容器未随内容重排,必须调整栏宽、改为上下结构或更换网格
- 全页只有 3 个或更少的内容组,每组不超过 4 条短句,却使用整高卡片和 16pt 或更小的正文——必须改为紧凑版式、放大字号、补充有价值信息或合并页面
- 文字卡片高度超过实际内容需求的 1.8 倍,或文字/有意义的视觉元素占用高度低于 55%
- 最密集区域与最稀疏区域的内容占用高度比超过 1.6:1,导致一侧拥挤、一侧空洞
- 页面存在超过画布 20% 的连续空白区,但它没有用于突出核心结论、图表或视觉焦点
- 距离幻灯片边缘的边距不够(< 0.5")
- 分栏或类似元素没有对齐
- 低对比度文字(比如浅灰文字配米色背景)
- 替换文字后模板装饰错位——比如标题下划线是按单行文字定位的,但替换后的标题换成了两行
- 低对比度图标(比如深色图标配深色背景,又没有对比色圆圈衬底)
- 文本框太窄导致过度换行
- 残留的占位符内容

## 转换成图片

把演示文稿转换成单页图片,方便视觉检查:

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 output.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

**把上面打印出来的绝对路径直接传给查看图片的工具。** 那条 `rm` 命令是清掉上一轮遗留的旧图片。`pdftoppm` 会按总页数决定补零位数:10 页以内是 `slide-1.jpg`,10-99 页是 `slide-01.jpg`,100 页以上是 `slide-001.jpg`。

**每次修完之后,把上面四条命令重新跑一遍**——PDF 必须从改过的 `.pptx` 重新生成,`pdftoppm` 才能反映出你的改动。

## 依赖

`pptxgenjs`(npm,已预装——只有 `require('pptxgenjs')` 失败时才需要装)· `markitdown[pptx]`、`Pillow`、`defusedxml`、`lxml`(pip——分别用于文本提取、缩略图、清理、校验)· LibreOffice(`soffice`,通过 `scripts/office/soffice.py` 针对沙箱环境自动配置)· `pdftoppm`(Poppler)
