"""Build a new deck from slides picked out of one or more decks.

    python3 scripts/merge_slides.py -o out.pptx a.pptx:5 b.pptx:2,8 a.pptx:7

Slides are numbered in show order (as markitdown and the rendered PDF number them),
and land in the output in the order given. The first deck is the base: its masters,
theme, size and docProps carry over, and every slide of it that is not picked is
dropped. A slide from another deck brings along what it references:

  - its layout. When the base already has the same layout under the same master
    (same XML, ignoring relationship ids), the slide is pointed at the base's copy;
    otherwise the layout comes over with its master, the master's other layouts and
    its theme, registered as an extra master
  - images, charts (with their embedded workbooks), diagrams and embedded objects,
    renamed so they do not collide with the base's parts
  - not its speaker notes or <p:tags> custom data, for the same reasons add_slide.py
    strips them

Every XML part is copied as bytes, so declarations and line endings stay as Office
wrote them. Run validate.py / check_delivery.py on the output as usual; docProps still
describe the base deck, so check_delivery.py --fix is normally needed.
"""

import argparse
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from add_slide import STRIP_ON_DUPLICATE_TYPE_RE, _purge_dangling_tag_refs
from clean import clean_unused_files
from office.helpers import opc_target, rezip, safe_extract

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_SLIDE = R_NS + "/slide"
REL_LAYOUT = R_NS + "/slideLayout"
REL_MASTER = R_NS + "/slideMaster"
CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
LAYOUT_ID_MIN = 2147483648


def die(msg):
    print(f"Error: {msg}", file=sys.stderr)
    sys.exit(1)


def read(p: Path) -> str:
    return p.read_bytes().decode("utf-8")


def write(p: Path, s: str):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(s.encode("utf-8"))


def rels_path(root: Path, part: str) -> Path:
    d, n = posixpath.split(part)
    return root / d / "_rels" / (n + ".rels")


def parse_rels(root: Path, part: str):
    """[(element text, Id, Type, resolved part or None)] for a part's relationships."""
    p = rels_path(root, part)
    if not p.exists():
        return []
    out = []
    for el in re.findall(r"<Relationship\b[^>]*/?>", read(p)):
        attr = dict(re.findall(r'(\w+)="([^"]*)"', el))
        tgt = opc_target(attr.get("Target", ""), part, attr.get("TargetMode", ""))
        out.append((el, attr.get("Id"), attr.get("Type", ""), tgt))
    return out


def show_order(root: Path):
    """Slide parts in <p:sldIdLst> order."""
    pres = read(root / "ppt/presentation.xml")
    by_rid = {rid: tgt for _, rid, typ, tgt in parse_rels(root, "ppt/presentation.xml") if typ == REL_SLIDE}
    lst = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", pres, re.S)
    rids = re.findall(r'r:id="([^"]+)"', lst.group(1)) if lst else []
    return [by_rid[r] for r in rids if r in by_rid]


def slide_size(root: Path):
    m = re.search(r"<p:sldSz\b[^>]*>", read(root / "ppt/presentation.xml"))
    return dict(re.findall(r'(cx|cy)="(\d+)"', m.group(0))) if m else {}


def normalized(root: Path, part: str) -> str:
    """Part XML minus relationship ids and layout-list ids, which differ between copies."""
    xml = re.sub(r'\sr:(id|embed|link|pict)="[^"]*"', "", read(root / part))
    return re.sub(r'(<p:sldLayoutId\b[^>]*?)\sid="\d+"', r"\1", xml)


class ContentTypes:
    def __init__(self, root: Path):
        self.path = root / "[Content_Types].xml"
        self.xml = read(self.path)

    def lookup(self, part: str):
        m = re.search(rf'<Override\b[^>]*PartName="/{re.escape(part)}"[^>]*ContentType="([^"]+)"', self.xml)
        if m:
            return m.group(1)
        ext = part.rsplit(".", 1)[-1].lower()
        m = re.search(rf'<Default\b[^>]*Extension="{re.escape(ext)}"[^>]*ContentType="([^"]+)"', self.xml, re.I)
        return m.group(1) if m else None

    def register(self, part: str, ctype: str | None):
        if not ctype or self.lookup(part) == ctype:
            return
        ext = part.rsplit(".", 1)[-1].lower()
        has_default = re.search(rf'<Default\b[^>]*Extension="{re.escape(ext)}"', self.xml, re.I)
        if not has_default and not ext == "xml":
            entry = f'<Default Extension="{ext}" ContentType="{ctype}"/>'
        else:
            entry = f'<Override PartName="/{part}" ContentType="{ctype}"/>'
        self.xml = self.xml.replace("</Types>", entry + "</Types>")

    def save(self):
        write(self.path, self.xml)


class Merger:
    def __init__(self, base_dir: Path):
        self.out = base_dir
        self.ct = ContentTypes(base_dir)
        self.taken = {p.relative_to(base_dir).as_posix() for p in base_dir.rglob("*") if p.is_file()}
        self.order = []  # output slide parts, in show order
        self.base_layouts = self._layout_index(base_dir)

    def _layout_index(self, root: Path):
        """normalized (layout xml, master xml) -> layout part, for matching."""
        idx = {}
        for lay in sorted(root.glob("ppt/slideLayouts/slideLayout*.xml")):
            part = lay.relative_to(root).as_posix()
            master = next((t for _, _, typ, t in parse_rels(root, part) if typ == REL_MASTER), None)
            if master:
                idx.setdefault((normalized(root, part), normalized(root, master)), part)
        return idx

    def fresh_name(self, part: str) -> str:
        d, n = posixpath.split(part)
        stem, ext = posixpath.splitext(n)
        stem = re.sub(r"\d+$", "", stem)
        i = 1
        while True:
            cand = posixpath.join(d, f"{stem}{i}{ext}")
            if cand not in self.taken:
                self.taken.add(cand)
                return cand
            i += 1

    def import_slide(self, src: Path, slide: str) -> str:
        """Copy slide (and its dependency closure) from unpacked src into the output."""
        mapping = {}  # src part -> output part
        new_masters = []
        queue = [slide]
        mapping[slide] = self.fresh_name(f"ppt/slides/{posixpath.basename(slide)}")
        seen = set()
        while queue:
            part = queue.pop()
            if part in seen:
                continue
            seen.add(part)
            for el, rid, typ, tgt in parse_rels(src, part):
                if tgt is None or tgt in mapping:
                    continue
                if part == slide and STRIP_ON_DUPLICATE_TYPE_RE.search(el):
                    continue
                if typ == REL_LAYOUT:
                    master = next((t for _, _, ty, t in parse_rels(src, tgt) if ty == REL_MASTER), None)
                    key = (normalized(src, tgt), normalized(src, master)) if master else None
                    if key in self.base_layouts:
                        mapping[tgt] = self.base_layouts[key]  # reuse, don't copy
                        continue
                if typ == REL_MASTER and tgt not in mapping:
                    new_masters.append(tgt)
                if not (src / tgt).exists():
                    continue
                mapping[tgt] = self.fresh_name(tgt)
                queue.append(tgt)

        stripped = []
        for part in seen:
            dest = mapping[part]
            (self.out / dest).parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(src / part, self.out / dest)
            self.ct.register(dest, ContentTypes(src).lookup(part) if part != slide else CT_SLIDE)
            rels = rels_path(src, part)
            if not rels.exists():
                continue
            xml = read(rels)
            for el, rid, typ, tgt in parse_rels(src, part):
                if part == slide and STRIP_ON_DUPLICATE_TYPE_RE.search(el):
                    xml = xml.replace(el, "")
                    stripped.append(rid)
                    continue
                if tgt is None or tgt not in mapping:
                    continue
                new_t = posixpath.relpath(mapping[tgt], posixpath.dirname(dest))
                xml = xml.replace(el, re.sub(r'Target="[^"]*"', f'Target="{new_t}"', el))
            write(rels_path(self.out, dest), xml)
        if stripped:
            _purge_dangling_tag_refs(self.out / mapping[slide], stripped)
        for m in new_masters:
            self._register_master(mapping[m])
        if new_masters:
            # later slides from the same deck reuse what this one brought over
            self.base_layouts = self._layout_index(self.out)
        return mapping[slide]

    def _register_master(self, master: str):
        pres_p = self.out / "ppt/presentation.xml"
        pres = read(pres_p)
        ids = [int(x) for x in re.findall(r'<p:sld(?:Master|Layout)Id\b[^>]*\bid="(\d+)"', pres)]
        for p in self.out.glob("ppt/slideMasters/*.xml"):
            ids += [int(x) for x in re.findall(r'<p:sldLayoutId\b[^>]*\bid="(\d+)"', read(p))]
        nxt = max(ids + [LAYOUT_ID_MIN - 1]) + 1

        # the imported master's layout ids must not collide with any other master's
        mp = self.out / master
        mxml = read(mp)

        def renum(m):
            nonlocal nxt
            nxt += 1
            return re.sub(r'\bid="\d+"', f'id="{nxt}"', m.group(0))

        mxml = re.sub(r"<p:sldLayoutId\b[^>]*>", renum, mxml)
        write(mp, mxml)
        nxt += 1

        rels_p = self.out / "ppt/_rels/presentation.xml.rels"
        rels = read(rels_p)
        used = {int(n) for n in re.findall(r'\bId="rId(\d+)"', rels)}
        rid = f"rId{max(used | {0}) + 1}"
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="{rid}" Type="{REL_MASTER}" '
                            f'Target="{posixpath.relpath(master, "ppt")}"/></Relationships>')
        write(rels_p, rels)
        pres = pres.replace("</p:sldMasterIdLst>", f'<p:sldMasterId id="{nxt}" r:id="{rid}"/></p:sldMasterIdLst>', 1)
        write(pres_p, pres)

    def finish(self):
        """Point <p:sldIdLst> at self.order (registering imported slides), then clean."""
        self.ct.save()
        rels_p = self.out / "ppt/_rels/presentation.xml.rels"
        rels = read(rels_p)
        by_target = {t: rid for _, rid, typ, t in parse_rels(self.out, "ppt/presentation.xml") if typ == REL_SLIDE}
        used = {int(n) for n in re.findall(r'\bId="rId(\d+)"', rels)}
        entries, next_id = [], 256
        for part in self.order:
            rid = by_target.get(part)
            if rid is None:
                n = max(used | {0}) + 1
                used.add(n)
                rid = f"rId{n}"
                rels = rels.replace("</Relationships>",
                                    f'<Relationship Id="{rid}" Type="{REL_SLIDE}" '
                                    f'Target="{posixpath.relpath(part, "ppt")}"/></Relationships>')
                by_target[part] = rid
            entries.append(f'<p:sldId id="{next_id}" r:id="{rid}"/>')
            next_id += 1
        write(rels_p, rels)
        pres_p = self.out / "ppt/presentation.xml"
        pres = read(pres_p)
        lst = "<p:sldIdLst>" + "".join(entries) + "</p:sldIdLst>"
        if re.search(r"<p:sldIdLst>.*?</p:sldIdLst>", pres, re.S):
            pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", lambda _: lst, pres, count=1, flags=re.S)
        elif re.search(r"<p:sldIdLst\s*/>", pres):
            pres = re.sub(r"<p:sldIdLst\s*/>", lambda _: lst, pres, count=1)
        else:
            pres = pres.replace("</p:sldMasterIdLst>", "</p:sldMasterIdLst>" + lst, 1)
        write(pres_p, pres)
        clean_unused_files(self.out)


def parse_spec(spec: str):
    m = re.match(r"^(.*):([\d,\s-]+)$", spec)
    if not m:
        die(f"{spec!r}: expected deck.pptx:N or deck.pptx:N,M")
    nums = []
    for chunk in m.group(2).split(","):
        chunk = chunk.strip()
        if "-" in chunk:
            a, b = (int(x) for x in chunk.split("-"))
            nums += list(range(a, b + 1))
        elif chunk:
            nums.append(int(chunk))
    return Path(m.group(1)).expanduser(), nums


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("specs", nargs="+", help="deck.pptx:N[,M,...] (show-order slide numbers, ranges like 3-5 allowed)")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()
    specs = [parse_spec(s) for s in args.specs]
    out = Path(args.output).expanduser()
    for deck, _ in specs:
        if not deck.exists():
            die(f"{deck} not found")
        if deck.resolve() == out.resolve():
            die("the output must be a new file, not one of the inputs")

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        unpacked = {}
        for deck, _ in specs:
            key = deck.resolve()
            if key not in unpacked:
                d = tmp / f"src{len(unpacked)}"
                with zipfile.ZipFile(deck) as z:
                    safe_extract(z, d)
                unpacked[key] = d
        base_key = specs[0][0].resolve()
        base_src = unpacked[base_key]
        work = tmp / "out"
        shutil.copytree(base_src, work)
        merger = Merger(work)
        base_size = slide_size(base_src)
        used_base = set()

        for deck, nums in specs:
            src = unpacked[deck.resolve()]
            order = show_order(src)
            if slide_size(src) != base_size:
                print(f"Warning: {deck.name} has a different slide size from {specs[0][0].name}; "
                      "its slides are copied unscaled", file=sys.stderr)
            for n in nums:
                if not 1 <= n <= len(order):
                    die(f"{deck.name} has {len(order)} slides; there is no slide {n}")
                part = order[n - 1]
                if deck.resolve() == base_key and part not in used_base:
                    used_base.add(part)
                    merger.order.append(part)
                    dest = part
                else:
                    dest = merger.import_slide(src, part)
                    merger.order.append(dest)
                print(f"{deck.name} slide {n} -> {dest}")

        merger.finish()
        rezip(work, out)
    print(f"Wrote {out} ({len(merger.order)} slides)")


if __name__ == "__main__":
    main()
