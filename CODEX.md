# Codex 执行说明

## 资源与工作目录

从本次读取的 `SKILL.md` 定位技能根目录，不依赖项目目录或 Claude 环境变量。示例的 `skill_dir` 替换为实际绝对路径；自定义安装位置时同样适用。

```bash
skill_dir="${CODEX_HOME:-$HOME/.codex}/skills/enn-pptx-new-2"
work_dir=$(mktemp -d "${TMPDIR:-/tmp}/enn-pptx-new-2.XXXXXX")
python3 "$skill_dir/scripts/guard_templates.py" --snapshot "$work_dir/templates-before.json"
cp "$skill_dir/templates/company-template.pptx" "$work_dir/source.pptx"
```

记录并复用同一个工作目录；不同 shell 调用不一定保留变量，后续使用实际绝对路径。

- 解压工作副本到工作目录内，再运行建页、清理及原有编辑流程。
- 改用户现有文件时先备份到工作目录，作为 `--original`；检查通过后才写回用户指定文件，不能覆盖技能模板。
- PDF、预览、缩略图全部输出到任务目录。删除旧预览只限本任务目录，不在用户目录执行宽泛通配删除。
- 完成四道交付检查后复核同一份指纹清单。校验仅检测原始模板修改、丢失或新增，不拦截工具，也不自动恢复文件。失败时停止覆盖成稿并排查，不能重建清单掩盖失败。

```bash
python3 "$skill_dir/scripts/guard_templates.py" --verify "$work_dir/templates-before.json"
```

## 按任务探测依赖

只检查当前任务用到的组件。不要为只读检查安装软件；需要安装时用隔离环境或用户已有环境，不在技能目录或系统 Python 写依赖。

```bash
python3 -c "import defusedxml, lxml; from PIL import Image"
command -v markitdown
command -v soffice
command -v pdftoppm
```

- 文字提取需要 markitdown；XML 修改／校验需要 defusedxml、lxml；缩略图需要 Pillow。
- 仅用 PptxGenJS 新画内容时，才从生成脚本目录检查 node 与 `require.resolve('pptxgenjs')`。Node 模块按脚本目录解析，不因切换 shell 目录就一定可用。
- 渲染需要 LibreOffice 和 Poppler。macOS 先检查 `/Applications/LibreOffice.app/Contents/MacOS/soffice`；已安装但不在 PATH 时，仅将其目录加入当前进程 PATH，不修改用户 shell 配置。
- 缺依赖时先检查已有环境。不能渲染时说明视觉检查未完成，不能把结构检查当视觉检查。
- LibreOffice 因实际沙箱限制失败时，按当前执行工具的授权机制处理；不要依赖 Linux 的 LD_PRELOAD 绕过 macOS 限制。

## 渲染与交付

用 `scripts/office/soffice.py` 导出 PDF，显式传 `--outdir` 为本任务目录，再用 pdftoppm 输出带唯一前缀的逐页图片。通过当前本地图片查看工具检查实际渲染图。

沿用 SKILL.md 的四道交付检查及字体检查，记录实际结果。最终回复提供 PPTX 的绝对路径 Markdown 链接，说明改动页、重要补充及未完成项。用户请求时再附 PDF 或预览，不把中间文件当最终交付。
