# 通用设计规范与 pptxgenjs 踩坑清单

这里是从 SKILL.md 挪出来的通用内容，公司模板场景一般用不到：

- 用户明确要求不套公司模板、自由设计时，参考“配色与风格”和“字体”两节；
- 只有在版式实例化之后、需要用 pptxgenjs 给正文区补画内容时，才看“pptxgenjs 踩坑清单”。

公司模板场景下，配色、字体、Logo 以 `COMPANY_STYLE.md` 为准。

## pptxgenjs 踩坑清单

只在**版式实例化之后、给正文区补充内容**时才用 pptxgenjs(比如往 `slideLayout7` 的通用内容占位符里画东西),而不是拿它从零画整份公司 PPT——封面/Logo/版式定位这些交给上面的模板实例化流程。用的时候颜色/字号一律取 `COMPANY_STYLE.md`,不要用下面"设计思路"里的通用配色表(那是没有公司模板时的兜底方案)。

`pptxgenjs` 已经预装好了——不要先跑 `npm install`;直接写脚本、`require('pptxgenjs')` 就行。只有 require 失败时才需要 `npm install pptxgenjs`。模型本身懂这个 API,下面这些是容易踩的坑:

- **加幻灯片之前先设置 `pres.layout`。** 默认画布是 `LAYOUT_16x9` = **10" × 5.625"**,不是 13.3" 宽。超出边界的坐标会被写进去,不会被裁剪——形状只是不显示在幻灯片上而已。(`LAYOUT_WIDE` 是 13.3" × 7.5"。)
- **十六进制颜色:不要带 `#`,不要用 8 位。** 写 `color: "FF0000"`。`"#FF0000"` 和把透明度编进十六进制(`"00000020"`)**都会导致文件损坏**。要做半透明效果:填充和图片用 `transparency: 0-100`,阴影用 `opacity: 0.0-1.0`——两个参数互不生效,用错了会被静默忽略。
- **pptxgenjs 会原地修改传入的 options 对象**(第一次用到时就把数值转成 EMU)。不要在两次 `add*` 调用之间共享同一个 `shadow`/options 对象——每次都新建一个。
- **阴影的 `offset` 必须 ≥ 0**——负值会导致文件损坏。要让阴影往上投,用 `angle: 270` 配合正数 offset。
- **`letterSpacing` 会被静默忽略**——真正生效的参数是 `charSpacing`。
- **列表:** 每一项都要设 `bullet: true`,不要手打字面的 `•`(会渲染成两个项目符号)。除最后一项外,数组里每项都设 `breakLine: true`。项目符号段落之间的间距用 `paraSpaceAfter`,不要用 `lineSpacing`(间距会大得离谱)。
- **每个输出文件对应一个 `new pptxgen()`**——不要复用同一个实例。
- **`rectRadius` 只对 `ROUNDED_RECTANGLE` 生效**,对 `RECTANGLE` 无效。
- **不支持渐变填充**——改用渐变图片当背景。
- **文本框自带内部留白**——需要文字跟形状/线条/图标在同一个 x 坐标对齐时,设 `margin: 0`。
- **讲者备注要写进 `slide.addNotes("...")`**(纯文本,每页一次),不要写进幻灯片上的文本框。
- **图表尽量用原生实现。** PowerPoint 能画的图表类型都用 `addChart()`(组合图传一个 `{type, data, options}` 数组)。库没暴露的 PowerPoint 原生功能(趋势线、误差线)要自己算出额外的数据系列或者后处理生成的 OOXML——不要退而求其次用渲染好的图片。只有 PowerPoint 本身就没有原生形式的图表类型(桑基图、网络图、和弦图)才用图片。
- **默认生成的图表很朴素**——没有标题、没有数据标签、配色也过时。要设 `showTitle` + `title`、`showValue: true` + `dataLabelPosition`、从你的配色方案取 `chartColors: [...]`,并且把边框元素调低调调(`catAxisLabelColor`/`valAxisLabelColor`、`valGridLine: { color, size }`、`catGridLine: { style: "none" }`、单数据系列时 `showLegend: false`)。
- **堆叠柱状图或条形图上,`dataLabelPosition` 只能是 `ctr`、`inEnd` 或 `inBase`。** `outEnd` **会导致文件损坏**。
- **组合系列用 `secondaryValAxis`/`secondaryCatAxis` 时,图表选项里 `valAxes` 和 `catAxes` 都要各给两条。** 不给的话 pptxgenjs 会写出它自己都没声明过的坐标轴 *id*,PowerPoint 会**直接丢弃那个图表**并报告文件损坏。只给 `valAxes` 是不够的。
- **`writeFile()` 之后,跑一遍 `python scripts/office/validate.py deck.pptx`。** 它会报出上面这两类图表问题,以及 PowerPoint 拒绝打开的幻灯片 XML 缺陷,并且每条都给出修复方法。去生成脚本里改,不要手改打包好的 XML。
- **永远不要打乱 `<p:presentation>` 子元素的顺序。** pptxgenjs 会把 `<p:notesMasterIdLst>` 紧跟着写在 `<p:sldIdLst>` 后面,两个母版都指向同一个 theme 部件。PowerPoint 能正常读这种顺序——一旦挪动这个元素,同一份 deck 就打不开了。
- **图标:** 用 `react-icons` 渲染成 SVG(`ReactDOMServer.renderToStaticMarkup`),用 `sharp` 栅格化到 ≥256px,再通过 `addImage({ data: "image/png;base64," + buf.toString("base64") })` 插入——`image/png;base64,` 这个前缀是必须的(`react-icons`、`react`、`react-dom`、`sharp` 都已预装——只有 require 失败时才需要 `npm install react-icons react react-dom sharp`)。

## 配色与风格

### 动手之前

- **选一套大胆的、贴合内容的配色方案**:配色应该让人感觉是专门为*这个*主题设计的。如果把你的配色换到一份完全不同的演示文稿里依然"说得过去",说明你的选择还不够具体。
- **主次分明,而不是平均用力**:一种颜色应该占主导地位(60-70% 的视觉权重),搭配 1-2 个辅助色调和一个鲜明的强调色。不要让所有颜色权重相同。
- **深色/浅色对比**:标题页和总结页用深色背景,内容页用浅色("三明治"结构)。或者整篇都坚持深色,营造高级感。
- **确定一个视觉母题并坚持下去**:选一个独特的元素反复使用——圆角图片框、彩色圆圈里的图标。贯穿每一页。**不要用色条或装饰条当作母题**(见"避免踩的坑"一节)。

### 配色方案

选跟主题贴合的颜色——不要默认用泛泛的蓝色。可以用下面这些配色方案找灵感:

| 主题 | 主色 | 辅助色 | 强调色 |
|-------|---------|-----------|--------|
| **午夜行政风** | `1E2761`(藏青) | `CADCFC`(冰蓝) | `FFFFFF`(白) |
| **森林苔藓** | `2C5F2D`(森林绿) | `97BC62`(苔藓绿) | `F5F5F5`(米白) |
| **珊瑚活力** | `F96167`(珊瑚色) | `F9E795`(金色) | `2F3C7E`(藏青) |
| **暖调赤陶** | `B85042`(赤陶色) | `E7E8D1`(沙色) | `A7BEAE`(鼠尾草绿) |
| **海洋渐变** | `065A82`(深蓝) | `1C7293`(青色) | `21295C`(午夜蓝) |
| **炭灰极简** | `36454F`(炭灰) | `F2F2F2`(米白) | `212121`(黑) |
| **青色信赖** | `028090`(青色) | `00A896`(海泡绿) | `02C39A`(薄荷绿) |
| **莓果奶油** | `6D2E46`(莓果色) | `A26769`(灰玫瑰) | `ECE2D0`(奶油色) |
| **鼠尾草沉静** | `84B59F`(鼠尾草绿) | `69A297`(桉树绿) | `50808E`(石板灰) |
| **樱桃大胆** | `990011`(樱桃红) | `FCF6F5`(米白) | `2F3C7E`(藏青) |

## 字体

**你写进 `.pptx` 里的字体名称,是由用户的 PowerPoint 渲染的,不是这个环境本身渲染的。** 视觉检查是通过 LibreOffice 渲染的,它会用替代字体顶替本地没有的字体——某些字体的替代版本宽度不一样,所以你的 QA 预览可能会显示出真实 deck 里不存在的文字溢出(或者反过来,显示"刚好放得下"但实际会溢出)。要让 QA 结果可信:

- **安全字体**(在 QA 里渲染宽度准确、*并且* Office 自带):**Arial、Calibri、Cambria、Times New Roman、Courier New、Bookman Old Style、Century Schoolbook**。正文和任何对是否放得下敏感的地方都用这些。
- **零 QA 风险又有个性的标题搭配**:安全列表里的衬线字体做标题(Cambria、Bookman Old Style、Century Schoolbook),配安全列表里的无衬线字体做正文(Calibri 或 Arial)。这样能拿到视觉对比,又不放弃可靠的溢出检查。
- **如果用户要求安全列表之外的字体**(比如 Georgia 或 Trebuchet MS):在用户指定的地方就用它,但给这些容器的尺寸多留一些余量(约 10%),不要相信 QA 对这些元素的"是否放得下"判断——这类字体的预览只是近似值。如果用户没指定,正文优先用安全列表里的字体。
- **QA 不可靠的字体**(替代字体宽度不一样——溢出检查可能是错的):Georgia、Trebuchet MS、Impact、Arial Black、Garamond、Consolas、Palatino Linotype。Calibri Light 的替代效果因环境而异,也按不可靠处理。用在标题/强调文字上、并留足余量没问题;但不要相信 QA 对这些字体的"是否放得下"判断。
- **永远不要默认用 Aptos**——Office 2023 之后的这个默认字体在这里没有宽度兼容的替代字体,而且在较老的 Office 安装里根本没有,两头都不可靠。
