#!/usr/bin/env python3
"""PreToolUse hook:阻止任何写入 templates/ 下原始模板文件的操作。

模板是所有成稿的视觉来源,被改一次后面每份 deck 都会带上错误。读取(cp 出去、解压、
markitdown、validate.py --original)不受影响;只拦截以模板为写入目标的操作。

从 stdin 读 hook JSON。exit 2 = 拦截,stderr 会作为反馈交给 Claude;exit 0 = 放行。
"""
import json
import os
import re
import shlex
import sys

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates")
TEMPLATES = {f for f in os.listdir(TEMPLATE_DIR) if f.endswith((".pptx", ".potx"))} if os.path.isdir(TEMPLATE_DIR) else set()


def is_template(path):
    path = path.strip("'\"")
    name = os.path.basename(path)
    if name not in TEMPLATES:
        return False
    parent = os.path.basename(os.path.dirname(path))
    return parent == "templates" or os.path.abspath(os.path.expanduser(path)) == os.path.join(TEMPLATE_DIR, name)


def written_by_bash(command):
    """返回命令里被当作写入目标的模板路径。按 ; && || | 拆成子命令逐个判断。"""
    hits = []
    for segment in re.split(r"&&|\|\||;|\|", command):
        for target in re.findall(r">>?\s*(\S+)", segment):
            if is_template(target):
                hits.append(target)
        try:
            argv = shlex.split(segment)
        except ValueError:
            argv = segment.split()
        if not argv:
            continue
        prog = os.path.basename(argv[0])
        args = [a for a in argv[1:] if not a.startswith("-")]
        if prog in ("cp", "mv", "install", "rsync") and args and is_template(args[-1]):
            hits.append(args[-1])
        elif prog in ("rm", "truncate", "touch") and any(is_template(a) for a in args):
            hits += [a for a in args if is_template(a)]
        elif prog == "zip" and args and is_template(args[0]):
            hits.append(args[0])
        elif prog == "sed" and "-i" in " ".join(argv) and any(is_template(a) for a in args):
            hits += [a for a in args if is_template(a)]
        elif any(a.endswith("add_slide.py") for a in argv) and "-o" not in argv:
            # add_slide.py deck.pptx ... 不传 -o 会原地改写输入的 deck
            hits += [a for a in args if is_template(a)]
    return hits


def main():
    try:
        data = json.load(sys.stdin)
    except json.JSONDecodeError:
        sys.exit(0)
    tool = data.get("tool_name", "")
    inp = data.get("tool_input", {}) or {}
    if tool in ("Edit", "Write", "MultiEdit", "NotebookEdit"):
        hits = [inp.get("file_path", "")] if is_template(inp.get("file_path", "")) else []
    elif tool == "Bash":
        hits = written_by_bash(inp.get("command", ""))
    else:
        hits = []
    if hits:
        print(f"已拦截:这一步会改写原始模板 {hits[0]}。先 cp 一份工作副本到别处(例如 /tmp/deck.pptx),"
              f"再对副本操作;add_slide.py 直接处理 .pptx 时要传 -o 指定输出文件。", file=sys.stderr)
        sys.exit(2)
    sys.exit(0)


if __name__ == "__main__":
    main()
