# templates/

这个目录本应放 `company-template.pptx` —— 整个 skill 的视觉约束来源（画布尺寸、配色、字体、Logo、9 个可实例化版式，以及 3 个可复用的内容原型页）。

**它没有入库。** 模板文件里带着真实的内部示例文案，不适合放在公开仓库中，因此被 `.gitignore` 排除。

## 使用前需要自行放置

把公司模板放到这里并命名为 `company-template.pptx`：

```
templates/company-template.pptx
```

没有这个文件时，`SKILL.md` 里所有"从模板实例化"的流程都无法执行。

## 模板需要满足的前提

`COMPANY_STYLE.md` 里的数值是从原始模板实测得出的，换用其他模板时需要同步核对：

- 画布 `12192000 x 6858000` EMU（13.33in x 7.5in，16:9）
- `slideMaster1` 承载 Logo，9 个版式 `slideLayout1.xml` ~ `slideLayout9.xml` 会自动继承
- `slide1.xml` / `slide2.xml` / `slide3.xml` 为三个内容原型页，是 `MODULES.md` 的复制源

模板结构不同时，先运行 `python scripts/thumbnail.py <模板>.pptx template-thumbs` 重新做一次版式盘点，再据此更新 `COMPANY_STYLE.md` 的版式一览表和配色表。
