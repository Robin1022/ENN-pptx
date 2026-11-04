# enn-pptx-new-2 · Codex 插件

使用新奥／ENNEW 公司模板制作和修改可编辑 PPT，包含两份模板、页面编辑与合并脚本，以及交付质量检查。

本分支 `enn-pptx-new-2-codex` 提供 Codex 版。原 Claude Code 版本保留在 `enn-pptx-new-2` 分支。插件和独立 Skill 共用 `skills/enn-pptx-new-2/`，只维护一份模板、脚本和制作规则。

## 安装插件（推荐给同事）

需要支持 `codex plugin` 的 Codex CLI，以及 Git。添加此仓库的 Codex 分支为插件市场，再安装：

```bash
codex plugin marketplace add Robin1022/ENN-pptx --ref enn-pptx-new-2-codex
codex plugin add enn-pptx-new-2@enn-pptx
```

安装后新开会话，使用 `$enn-pptx-new-2`，例如：

> 使用 $enn-pptx-new-2，帮我修改这份公司汇报的第 3 页，沿用原文件的样式。

在支持插件的 Codex 桌面入口，也可打开插件浏览器检查安装结果。如果看不到新插件，重启客户端。上述命令添加的是本项目的 Git 插件市场，不是向公共插件目录发布。

本插件只提供本地 Skill、模板和脚本，不包含 MCP 服务，不要求连接账号，不会在安装时自动安装软件。制作和渲染所需依赖见下文。

## 更新和卸载

GitHub 上推送更新后，用户需要刷新 Git 市场；不要将推送等同于本机立即生效：

```bash
codex plugin marketplace upgrade enn-pptx
codex plugin list --json
```

刷新后新开会话，检查插件版本；如旧版本仍被使用，可通过插件浏览器重新安装，或执行：

```bash
codex plugin remove enn-pptx-new-2@enn-pptx
codex plugin add enn-pptx-new-2@enn-pptx
```

卸载插件只需上面的 `remove` 命令，不会删除用户已经生成的 PPT。

## 仅安装 Skill

不使用插件时，可克隆仓库并把其中的 Skill 链接到本地技能目录。以下为 macOS／Linux 示例：

```bash
git clone -b enn-pptx-new-2-codex https://github.com/Robin1022/ENN-pptx.git "$HOME/ENN-pptx-codex"
mkdir -p "${CODEX_HOME:-$HOME/.codex}/skills"
ln -s "$HOME/ENN-pptx-codex/skills/enn-pptx-new-2" "${CODEX_HOME:-$HOME/.codex}/skills/enn-pptx-new-2"
```

若目标目录已有旧版，先备份并移走旧目录，不要覆盖或嵌套复制。Windows 可复制该 Skill 文件夹到本地技能目录，之后更新时重新复制。

通过符号链接安装后，在仓库目录执行 `git pull --ff-only` 即可更新文件。独立 Skill 和插件选择一种安装即可；同时安装可能出现同名技能重复。

## 本机依赖

插件安装完成不代表生成与渲染依赖已经安装。按实际任务准备：

| 任务 | 依赖 |
|---|---|
| XML 编辑、结构检查 | Python 3.10+、defusedxml、lxml |
| 文本提取与缩略图 | markitdown[pptx]、Pillow |
| PDF 与图片渲染 | LibreOffice、Poppler（pdftoppm）及 PPT 所用字体 |
| 用 PptxGenJS 新画内容（可选） | Node.js、pptxgenjs |

建议 Python 依赖安装在已有环境或独立虚拟环境中。缺少依赖时不会自动安装；缺少字体时渲染换行可能与 PowerPoint 不同。具体安装命令和字体说明见 [Skill 使用说明](skills/enn-pptx-new-2/README.md)及 [Codex 执行说明](skills/enn-pptx-new-2/CODEX.md)。

## 版本维护与打包

1. 在 `skills/enn-pptx-new-2/` 修改唯一的 Skill 内容，不编辑用户插件缓存。
2. 更新根目录 `plugin.json` 的版本。打包脚本自动同步 Codex 兼容清单，避免维护两套版本号。
3. 运行下面命令，检查清单并生成 ZIP：

```bash
python3 tools/build_plugin.py
```

输出为 `dist/enn-pptx-new-2-<版本>.zip`，可分享给需要离线拷贝的同事。解压后可从解压目录添加本地插件市场，再安装；ZIP 中包含市场清单。Git 安装优先使用上面的分支命令。

```bash
codex plugin marketplace add /absolute/path/to/unpacked-plugin
codex plugin add enn-pptx-new-2@enn-pptx
```

维护时同步提交生成的 `.codex-plugin/plugin.json`。发布插件包不会自动安装 LibreOffice、字体或其他依赖。

## 内容与来源

- [Skill 入口与制作规范](skills/enn-pptx-new-2/SKILL.md)
- [公司视觉规范](skills/enn-pptx-new-2/COMPANY_STYLE.md)
- [模板说明](skills/enn-pptx-new-2/templates/README.md)
- [历史优化记录](skills/enn-pptx-new-2/EXPERIENCE.md)

底层工具来自 Anthropic 的 PPTX Skill，沿用原有来源说明；本项目没有为上游代码另行声明许可。公司模板与相关素材的使用范围由其权利人决定。

插件格式和分发方式参考 [OpenAI 官方插件打包文档](https://developers.openai.com/plugins/build/plugins)。
