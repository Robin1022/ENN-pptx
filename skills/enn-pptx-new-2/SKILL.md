---
name: enn-pptx-new-2
description: "使用新奥／ENNEW 公司模板制作 PPT，或沿用现有公司演示文稿的样式修改页面、合并选页、将截图重画为可编辑幻灯片。适用于公司汇报、技术方案、转正答辩；不用于无公司品牌要求的私人演示文稿。"
---

# 公司 PPTX 生成、编辑与分析

## 不可省的规则

1. **新建 deck 先规划再生成**(`PLANNING.md`)。用户要求先审规划时等待确认；用户已明确要求直接制作、已确认规划或授权自主决定时，说明合理假设后继续，不重复索要确认。素材缺口自主推理补全,但推测内容必须标 `[推理补充]` / `[待确认假设]` / `[示意数据]`,不得写成用户提供的事实。推理和假设标记默认只写在规划和回复里；用户明确要求的模拟／示例数据必须在页面可见位置标明为示意，详见「内容来源」。
2. **模板原文件只读。** 一律 `cp` 出工作副本再操作;不要向 `templates/` 写入输出；工作开始时记录模板 SHA-256，交付前复核。完整性检查只能发现改动，不能阻止写入，不能声称存在自动 hook 保护。
3. **成稿不得带模板示例文案。** 所有 `add_slide.py` 复制完成后,从 `<p:sldIdLst>` 移除示例页,再跑 `clean.py`。
4. **交付前四道检查全过才算完成**:`validate.py --original`、`check_delivery.py`、`check_layout.py`、渲染后的视觉 QA。任何一道失败,修完从失败的那一步重跑。

## Codex 运行方式

- 本技能自包含，不依赖已删除的通用 `pptx` 技能，也不依赖 Claude Code、`~/.claude` 或 Claude hook。技能名称和安装目录统一为 `enn-pptx-new-2`。
- 使用本次实际读取的 `SKILL.md` 所在目录定位资源，不从项目当前目录猜测路径。脚本用绝对路径调用，或通过执行工具的工作目录参数指定技能目录。
- 每个任务在项目可写目录或系统临时目录新建唯一工作目录；所有解包、备份、PDF 和预览图均写在其中。不要复用固定的 `/tmp/deck.pptx`、`/tmp/unpacked`，也不要在技能目录安装依赖或写输出。
- 只读检查和已授权的生成、修改可继续执行；遵守实际工具的沙箱权限，权限拒绝时按工具要求申请授权，不通过换工具规避。除用户明确要求外，不安装软件、不改全局配置、不启动 PowerPoint。
- 先按 [CODEX.md](CODEX.md) 检查任务所需依赖、创建工作目录和模板指纹；缺少渲染工具时可继续完成不依赖它的工作，但必须报告未完成的视觉检查，不能宣布通过全部交付检查。
- 预览本地渲染图片使用当前环境提供的图片查看工具；最终回复提供输出文件的绝对路径链接，简述改动、验证结果及未完成项。不要为了检查而自动启动子代理；仅在当前会话明确允许时委派。

## 配套文档

动手前按任务读对应文档:

- [`PLANNING.md`](PLANNING.md) —— 页数、叙事、逐页内容与布局规划及授权处理
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
| **从几份 PPT 里挑页合成新文件** | `python3 scripts/merge_slides.py -o 新文件.pptx a.pptx:5 b.pptx:2,8`,页码按放映顺序,输出按参数顺序排列。版式、母版、图片自动带上,不手工拷贝部件。之后照常跑四道检查,docProps 用 `check_delivery.py --fix` 修 |
| **把图片(截图、草图、别处的页面)转成一页 PPT** | 在公司模板「正文页」版式(`slideLayout3`)上,用可编辑的形状和文字按图重画,不整张贴图;文字以图上为准。走「默认工作流」第 2-8 步,规划只在回复里简要列出页面结构,不用等确认 |
| **读取内容** | `markitdown deck.pptx`(每页一个 `<!-- Slide number: N -->` 区块);缩略图:`python3 scripts/thumbnail.py deck.pptx <prefix>` |

用户没有要求公司品牌/模板、且是跟公司汇报无关的私人演示文稿时,不强制套用公司品牌；使用当前可用的通用演示文稿工具，缺少时明确说明，不调用不存在的 `pptx` 技能。

### 现成的场景模板

| 模板 | 命中场景 | 页数 |
|---|---|---|
| `templates/campus-hire-defence.pptx` | 校招生转正答辩 / 新进人员任前答辩 / 试用期转正述职 | 14 |

`cp templates/campus-hire-defence.pptx <目标路径>` 后逐页替换文字;改动时同样适用「填充模板时」的全部注意事项,尤其是**填充元素只能替换不能新增**(`<p:spPr>` 的填充是互斥选择组,插出两个并列的 `<a:solidFill>` 会让 PowerPoint 报损坏,而 XSD、LibreOffice、python-pptx 都不报错)。交付前检查用 `--scaffold` 指向这份场景模板。

## 内容来源

用户给了材料(文档、已有页面、口述要点)时,页面上的痛点、技术点、方法和组件名都以材料为准。

- 材料之外补充的定性推理、已验证的外部资料和建议，每次改完都在回复里说明补了什么、写在哪、依据是什么。用户问“材料里有没有”时，依据材料原文回答；没有就直说是补充的。
- 真实业绩、金额、比例、客户和经历等事实缺失时，不主动填示意值。优先使用有依据的定性表述、删除无依据栏目；关键数据不可缺少时向用户询问。外部基准不能当作本项目实绩或承诺。
- 只有用户明确要求演示、模拟或示例数据时才生成示意值。必须在数据所在页的可见位置注明“示意数据”或“模拟数据，非实际成果”，不得只在回复、讲者备注或参考页说明。即使用户要求只在回复说明，也须保留页内标识，或改用不含模拟数字的表达。
- 推理、假设和来源标记默认只写在规划和回复里；用户要求来源脚注或注释时再加。前述模拟／示例数据的页内标识不受这一默认规则限制。

## 文案

这类 PPT 用于技术汇报和评审,页面文字按正式的书面技术语言写:

- 直接陈述做法、机制和结论。不用设问式标签、装饰性引号、口号式短语,也不用口语和网络化表达。
- 每句话落到具体的对象和动作上,读者不靠讲解也能看懂。概括性的说法后面要跟上它具体指什么。
- 术语沿用材料和已有页面的叫法。新引入的术语要么当页解释清楚,要么换成材料里已有的说法;同一个概念全页只用一种叫法。
- 版面放不下时,缩短句子、删掉次要信息或调整版面,不要把一句话压缩成自造的缩略词或四字短语。页面上每个短语,不熟悉这个项目的人也要能直接看懂。
- 分析内容时用到的框架、维度名和内部字段名只用来帮助思考,不写到页面上,除非材料本身就是这么写的。
- 问题和难点写出原因:什么因素导致了什么困难。不要只罗列表面现象。
- 用户在本次对话里否掉的措辞和说法,之后每次改文字前都要先对照一遍。改其他地方时也不能再用回去。

## 页内逻辑

技术汇报页常有一条论证链,比如"问题 → 目标 → 难点 → 方案",或者"环节 → 难点 → 解决方案"。链上每一环都要和前后两环接得上:

- **一一对应**:每个难点对应一个问题或环节,每个方案正好解决它对应的那个难点。对不上时,先改难点或方案的表述,不要硬连。
- **互不重叠**:同一页的几个难点、几个例子各讲一件事,也不重复相邻页已经讲过的内容。两条讲的是同一件事时,合并或者换掉一条。
- **立场一致**:难点写成"要做到什么、需要克服什么",不写成否定本方方案的话。比如说本方生成的结果"易出错",就和后面讲方案的页面冲突了。
- **目标写成状态**:目标写成要达到的状态或能力,不写成要做的动作,而且要回应上面列出的问题。

改了链上任何一环,都要把整条链从头读一遍再交付,并在回复里用一两句话说明改完后的对应关系。

## 修改现有页面

用户最常见的用法是拿一份已有的 PPT,围绕其中一两页反复修改。每一轮:

1. **确定写到哪个文件。** 默认原地写回用户给的那份文件,不另存新版本。用户说"新的 ppt""单起一个""新加一页"时,才另存文件或用 `add_slide.py` 加页;新加的页放在被参照页的后面,原页保持不动。
2. **动手前备份。** 第一次修改前把原文件复制到工作目录,作为 `--original` 的基线和出错时的回退版本。
3. **改内容。** 解压 → 改 `ppt/slides/slideN.xml` → 打包,写法见「填充模板时」。用生成脚本重画整块区域时,传给 XML 的每个属性值都要是 schema 允许的取值(对齐、换行、锚点这类枚举不能传数字或 `None`)。
   新画或重画的内容,字号、配色、字体、间距先从这份文件里取:同类元素(页标题、板块标题、卡片标题、正文、图注)沿用其他页已有的取值,强调色沿用该页或相邻页已经在用的颜色。「设计约束」里的数值只在文件里找不到可参照的同类元素时使用。
4. **检查后再写回。**
   ```bash
   python3 scripts/office/validate.py out.pptx --original backup.pptx
   python3 scripts/check_delivery.py out.pptx
   python3 scripts/check_layout.py out.pptx --original backup.pptx --slides N
   ```
   然后只渲染改动的页(「转换成图片」),对照几何检查的结果逐条看。都处理完再覆盖原文件。
5. **回复里说清楚**改了哪一页的哪些地方,以及按「内容来源」列出这轮补充的内容。

## 默认工作流:用公司模板新建

按任务执行以下清单；进度回复只说明关键发现和下一步，不必每次粘贴整份清单:

```text
- [ ] 1. 规划:按 PLANNING.md 出整体规划(含 RESEARCH.md 的检索决策、LAYOUT.md 的逐页密度判断),按已有授权执行
- [ ] 2. 建页:cp 模板 → 解压 → add_slide.py 建出全部目标页
- [ ] 3. 定稿结构:调整 <p:sldIdLst> 顺序,移除示例页引用,跑 clean.py
- [ ] 4. 填内容:改 ppt/slides/slideN.xml,或用 pptxgenjs 重搭不匹配的区域(见 PPTXGENJS.md)
- [ ] 5. 打包 + validate.py --original          失败 → 回到第 4 步修
- [ ] 6. check_delivery.py --template           失败 → docProps 类加 --fix;文案残留回到第 4 步
- [ ] 7. check_layout.py --original             对照渲染图确认是可见问题 → 回到第 4 步
- [ ] 8. 渲染成图片做视觉 QA                   有缺陷 → 回到第 4 步,修完从第 5 步重跑
```

第 1 步的规划要求、标记规则和输出格式全在 `PLANNING.md`。下面代码使用 `CODEX.md` 已建立的 `skill_dir` 与 `work_dir`，每个任务使用不同工作目录。三个示例正文页是可选原型:每页由模型按内容语义判断是否复制对应源页,不匹配就用 `COMPANY_STYLE.md` 的其他版式,不让用户逐页挑模块。

```bash
# 2. 解压工作副本
cp "$skill_dir/templates/company-template.pptx" "$work_dir/deck.pptx"
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" "$work_dir/deck.pptx" "$work_dir/unpacked"
#    匹配三种模块的页 = 复制 slide1/2/3.xml(见 MODULES.md);封面/目录/其他 = 从 slideLayoutN.xml 实例化。
#    Content_Types / rels / <p:sldIdLst> 自动处理,Logo 随版式/母版自动带出
python3 "$skill_dir/scripts/add_slide.py" "$work_dir/unpacked/" slide1.xml --after slide3.xml        # 例:反思模块
# python3 "$skill_dir/scripts/add_slide.py" "$work_dir/unpacked/" slideLayout3.xml --after slide3.xml # 例:从"正文页"版式新开

# 3. 全部页面建好后:改 ppt/presentation.xml 的 <p:sldIdLst> 定顺序,删掉 slide1/2/3 示例页引用
#    (用户明确要保留的除外)。不要直接删 slide 文件,交给 clean.py
python3 "$skill_dir/scripts/clean.py" "$work_dir/unpacked/"

# 4. 编辑最终保留页:结构跟源页接近 = 安全文字替换(见「填充模板时」);
#    栏数/条目数/层级对不上 = 保留 COMPANY_STYLE.md 的图形语言,增删形状子组或用 pptxgenjs 重搭该区域

# 5-7. 打包与检查
(cd "$work_dir/unpacked" && zip -Xr "$work_dir/out.pptx" .) # 首次打包；重打包前只删除本任务旧 out.pptx
python3 "$skill_dir/scripts/office/validate.py" "$work_dir/out.pptx" --original "$skill_dir/templates/company-template.pptx"
python3 "$skill_dir/scripts/check_delivery.py" "$work_dir/out.pptx" --template "$skill_dir/templates/company-template.pptx"
python3 "$skill_dir/scripts/check_layout.py" "$work_dir/out.pptx" --original "$skill_dir/templates/company-template.pptx"
```

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/thumbnail.py deck.pptx [prefix]` | 每页带标签的缩略图网格,用来挑版式。只认 `.pptx`。**一定传 `prefix`**——默认 `thumbnails`,同目录处理另一份 deck 会覆盖 |
| `scripts/add_slide.py unpacked/ slide2.xml [--after slideN.xml]` | 复制一页(或从 `slideLayoutN.xml` 新建),自动处理包级注册,并摘掉 `notesSlide`/`tags` 引用。两页形状指向同一个 `<p:tags>` 部件时,真实 PowerPoint 会静默清空整页,而 XSD/python-pptx/LibreOffice 都看不出来——自己手写复制逻辑时要防这个坑。直接处理 `.pptx` 时**必须传 `-o out.pptx`**,否则原地改写输入 |
| `scripts/merge_slides.py -o out.pptx a.pptx:5 b.pptx:2,8` | 从一份或多份 deck 挑页合成新文件。第一份是底稿(母版、主题、docProps 沿用它);其他 deck 的页面带上所需的版式和母版,同一版式已存在时直接复用,不重复导入;不带讲者备注和 `<p:tags>`。输出必须是新文件 |
| `scripts/clean.py unpacked/` | 删除不再被引用的幻灯片、媒体和关系。`<p:sldIdLst>` 定稿**之后**运行 |
| `scripts/office/validate.py deck.pptx [--original src.pptx]` | Schema、关系、内容类型、图表和幻灯片检查,每条失败附修复方法。模板派生的 deck 必须传 `--original` |
| `scripts/check_delivery.py deck.pptx [--template ...] [--scaffold ...] [--fix]` | 交付前检查:XML 声明 `standalone="yes"`、docProps 作者与旧页标题、模板示例文案残留、场景模板占位、通用占位符与规划标签。`--fix` 自动修 docProps |
| `scripts/check_layout.py deck.pptx [--original src.pptx] [--slides 5,6]` | 版面几何检查:形状超出画布、图片或方块压到边框、箭头短到只剩箭头头部,以及按字宽估算的文字溢出、文字贴着或跨过边框、连线穿过文字。结果要对照渲染图判断,见「版面几何检查」 |
| `scripts/office/soffice.py --headless --convert-to pdf deck.pptx` | LibreOffice 封装,自动配置运行环境,避免裸调 `soffice` 卡住 |
| `scripts/guard_templates.py --snapshot <manifest.json>` / `--verify <manifest.json>` | 显式记录／复核模板 SHA-256，不依赖自动 hook；详见 CODEX.md |

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

配色、字体、Logo 以 `COMPANY_STYLE.md` 为准。以下约束只管**新画的内容**;修改用户已有的 PPT 时,字号、配色和间距先沿用该文件已有页面的取值(见「修改现有页面」第 3 步),下面的数值只作缺省值。母版/版式自带的元素(如 slide1-3 标题下那条细分隔线,来自 `slideLayout3` 所属母版)是公司既定设计,不要改或删。

**版式**：根据表达需要选择视觉形式。关系、流程、架构和数据优先可视化；结论、决策事项、简短说明等允许使用排版清晰的纯文字页，可用 `slideLayout7`。不为满足形式要求硬加图标、卡片或装饰；避免大段文字堆积，并按 `LAYOUT.md` 检查内容密度和留白。相邻页的版式变化应服务于内容，不强制每页换一种布局。

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

按「转换成图片」渲染后逐页看。盯着生成代码久了容易看到预期而非实际,所以重新对照实际图像检查，不依赖对生成代码的预期。要找的缺陷:

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

先确认本机装了这份 deck 用到的中文字体。LibreOffice 找不到字体时会换用其他字体,字宽和换行都和 PowerPoint 不一样,渲染图上的压线、溢出就不能作为判断依据:

```bash
unzip -p output.pptx 'ppt/slides/*.xml' | grep -o 'typeface="[^"+][^"]*"' | sort | uniq -c   # deck 用到的字体
fc-list :lang=zh family | grep -iE "Microsoft YaHei|DengXian"                              # 本机有没有
```

缺字体时,LibreOffice 每次渲染挑的替代字体都可能不一样,有时是手写体,同一份文件前后两次渲染的字宽也不同,所以改前改后的对比图不能拿来比换行。在回复里说明渲染图的换行可能和 PowerPoint 不一致,并告诉用户怎么安装,不要自行安装。装了 Office for Mac 的,字体在 `/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/`,把 `msyh*.ttc`、`Deng*.ttf` 复制到 `~/Library/Fonts/` 就行。

```bash
python3 "$skill_dir/scripts/office/soffice.py" --headless --convert-to pdf --outdir "$work_dir" "$work_dir/output.pptx"
# 只在本任务目录清理旧预览；不要在项目或用户目录运行通配删除
pdftoppm -jpeg -r 150 "$work_dir/output.pdf" "$work_dir/slide"
```

把打印出的绝对路径直接交给看图工具。重渲染前仅清理本任务目录的旧图;`pdftoppm` 按总页数补零(10 页内 `slide-1.jpg`,10-99 页 `slide-01.jpg`)。每次修改后重新导出 PDF 并渲染,PDF 必须从改过的 `.pptx` 重新生成。

## 依赖

```bash
pip install "markitdown[pptx]" Pillow defusedxml lxml   # 文本提取、缩略图、清理、校验
npm install pptxgenjs                                   # 仅在 require 失败时
brew install --cask libreoffice                         # soffice,转 PDF(Linux: libreoffice)
brew install poppler                                    # pdftoppm,转图片(Linux: poppler-utils)
```

`check_delivery.py` 和 `guard_templates.py` 只用 Python 标准库(3.9+)。

## 测试

`evals/evals.json` 提供行为回归场景。修改脚本时运行相关功能验证；完整行为评估按任务范围在隔离工作目录执行，报告实际覆盖范围，不把静态校验或脚本冒烟测试称为整份 PPT 的视觉验收。
