# pptxgenjs 踩坑清单

从 SKILL.md 挪出来:实际使用中很少用 pptxgenjs,只在需要它给正文区补画内容时读这里。

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
