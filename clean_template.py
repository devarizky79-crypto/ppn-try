"""Clean and simplify the crude-availability Excel template.

Works directly on the .xlsx package (zip + XML) instead of re-saving it
through openpyxl, so tables, threaded comments, data-validation extensions
and other Excel features in the kept sheets survive unchanged.

What it does:
  1. Removes hidden sheets that no kept sheet references (old monthly copies).
  2. Removes external workbook links that no kept formula or name uses.
  3. Removes defined names that are broken (#REF!), point to external files,
     belong to removed sheets, or are not used anywhere.
  4. Deletes calcChain.xml (Excel rebuilds it on open) and any package parts
     left orphaned by the steps above.

The original file is never modified; a new file is written.

Usage:
    python clean_template.py INPUT.xlsx [OUTPUT.xlsx] [--keep-hidden "Sheet A,Sheet B"]
"""

from __future__ import annotations

import argparse
import html
import posixpath
import re
import sys
import zipfile
from pathlib import Path

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------------------------------------------------------- helpers
def rels_path(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def parse_rels(xml: str) -> list[dict]:
    out = []
    for m in re.finditer(r"<Relationship\b[^>]*/?>", xml):
        tag = m.group(0)
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', tag))
        attrs["_tag"] = tag
        out.append(attrs)
    return out


def resolve(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def sheet_ref_present(text: str, name: str) -> bool:
    quoted = "'" + name.replace("'", "''") + "'!"
    if quoted in text:
        return True
    if re.fullmatch(r"[A-Za-z_][\w.]*", name):
        return re.search(r"(?<![\w'.\]])" + re.escape(name) + r"!", text) is not None
    return False


# ---------------------------------------------------------------- main
def clean(src: Path, dst: Path, keep_hidden: set[str]) -> dict:
    zin = zipfile.ZipFile(src)
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    infos = {i.filename: i for i in zin.infolist()}
    txt = lambda p: files[p].decode("utf-8")

    wb_part = "xl/workbook.xml"
    wb = txt(wb_part)
    wb_rels_p = rels_path(wb_part)
    wb_rels = parse_rels(txt(wb_rels_p))
    rid_to_target = {r["Id"]: resolve(wb_part, r["Target"]) for r in wb_rels if r.get("TargetMode") != "External"}

    # ---- sheets in workbook order
    sheets = []
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        tag = m.group(0)
        name = html.unescape(re.search(r'\bname="([^"]*)"', tag).group(1))
        rid = re.search(r'\br:id="([^"]*)"', tag).group(1)
        st = re.search(r'\bstate="([^"]*)"', tag)
        sheets.append({"tag": tag, "name": name, "rid": rid, "part": rid_to_target[rid],
                       "state": st.group(1) if st else "visible"})
    sheet_text = {s["name"]: html.unescape(txt(s["part"])) for s in sheets}

    # ---- which sheets to keep: visible + anything they (transitively) reference
    keep = {s["name"] for s in sheets if s["state"] == "visible"} | keep_hidden
    changed = True
    while changed:
        changed = False
        for s in sheets:
            if s["name"] in keep:
                continue
            if any(sheet_ref_present(sheet_text[k], s["name"]) for k in keep):
                keep.add(s["name"])
                changed = True
    removed_sheets = [s["name"] for s in sheets if s["name"] not in keep]
    kept = [s for s in sheets if s["name"] in keep]
    old_index = {s["name"]: i for i, s in enumerate(sheets)}
    new_index = {s["name"]: i for i, s in enumerate(kept)}

    for s in sheets:
        if s["name"] not in keep:
            wb = wb.replace(s["tag"], "")

    # ---- defined names
    kept_text = "\n".join(sheet_text[s["name"]] for s in kept)
    dn_block = re.search(r"<definedNames>.*?</definedNames>", wb, re.S)
    names = []
    if dn_block:
        for m in re.finditer(r"<definedName\b([^>]*)>(.*?)</definedName>", dn_block.group(0), re.S):
            attrs, val = m.group(1), html.unescape(m.group(2))
            nm = re.search(r'\bname="([^"]*)"', attrs).group(1)
            ls = re.search(r'\blocalSheetId="(\d+)"', attrs)
            names.append({"full": m.group(0), "attrs": attrs, "val": val, "name": nm,
                          "local": int(ls.group(1)) if ls else None})

    def broken(n):
        v = n["val"]
        if "#REF!" in v or v.strip() == "#N/A":
            return True
        if re.search(r"\[\d+\]", v) and not re.search(r"\[#|\[\[", v):  # external workbook ref
            return True
        if any(sheet_ref_present(v, r) for r in removed_sheets):
            return True
        if n["local"] is not None and sheets[n["local"]]["name"] not in keep:
            return True
        return False

    candidates = [n for n in names if not broken(n)]
    # keep built-in names (_xlnm.*) that survive, plus user names that are actually used
    used = set()
    pool = kept_text
    changed = True
    while changed:
        changed = False
        for n in candidates:
            if n["name"] in used:
                continue
            if n["name"].startswith("_xlnm.") or re.search(r"(?<![\w.])" + re.escape(n["name"]) + r"(?![\w(])", pool):
                used.add(n["name"])
                pool += "\n" + n["val"]
                changed = True
    final_names = [n for n in candidates if n["name"] in used]

    def renumber(n):
        if n["local"] is None:
            return n["full"]
        new = new_index[sheets[n["local"]]["name"]]
        return n["full"].replace(f'localSheetId="{n["local"]}"', f'localSheetId="{new}"', 1)

    if dn_block:
        new_block = ("<definedNames>" + "".join(renumber(n) for n in final_names) + "</definedNames>") if final_names else ""
        wb = wb.replace(dn_block.group(0), new_block)

    # ---- external links: keep only those still referenced as [n] outside structured refs
    ext_block = re.search(r"<externalReferences>.*?</externalReferences>", wb, re.S)
    removed_links = 0
    if ext_block:
        ext_rids = re.findall(r'r:id="([^"]+)"', ext_block.group(0))
        all_text = kept_text + "\n" + "\n".join(n["val"] for n in final_names)
        used_idx = {int(i) for i in re.findall(r"(?<![\[\w])\[(\d+)\](?=[^\]\[]*!)", all_text)}
        if used_idx:
            # renumbering external indices inside formulas is risky; keep them all
            print("NOTE: kept formulas still use external links", sorted(used_idx), "- leaving links in place")
        else:
            wb = wb.replace(ext_block.group(0), "")
            removed_links = len(ext_rids)
            for rid in ext_rids:
                wb_rels = [r for r in wb_rels if r["Id"] != rid]

    # ---- drop removed-sheet + calcChain relationships from the workbook
    removed_rids = {s["rid"] for s in sheets if s["name"] not in keep}
    wb_rels = [r for r in wb_rels if r["Id"] not in removed_rids and not r["Type"].endswith("/calcChain")]

    # ---- book view: first visible tab, keep the active sheet if it survived
    def fix_view(m):
        tag = m.group(0)
        act = re.search(r'activeTab="(\d+)"', tag)
        act_name = sheets[int(act.group(1))]["name"] if act else kept[0]["name"]
        new_act = new_index.get(act_name, 0)
        tag = re.sub(r'\sfirstSheet="\d+"', "", tag)
        tag = re.sub(r'\sactiveTab="\d+"', "", tag)
        return tag.replace("<workbookView", f'<workbookView activeTab="{new_act}"', 1)
    wb = re.sub(r"<workbookView\b[^>]*/?>", fix_view, wb, count=1)

    files[wb_part] = wb.encode("utf-8")
    rel_xml = txt(wb_rels_p)
    head = rel_xml[: rel_xml.index("<Relationship ")] if "<Relationship " in rel_xml else rel_xml.replace("</Relationships>", "")
    files[wb_rels_p] = (head + "".join(r["_tag"] for r in wb_rels) + "</Relationships>").encode("utf-8")

    # make sure exactly one kept sheet is selected (deleted sheet may have been the selected one)
    act_i = int(re.search(r'activeTab="(\d+)"', wb).group(1))
    for i, s in enumerate(kept):
        x = txt(s["part"])
        x2 = re.sub(r'\stabSelected="1"', "", x, count=1)
        if i == act_i:
            x2 = x2.replace("<sheetView ", '<sheetView tabSelected="1" ', 1)
        if x2 != x:
            files[s["part"]] = x2.encode("utf-8")

    # ---- garbage-collect parts no longer reachable from the package root
    reachable, stack = set(), ["_rels/.rels"]
    reachable.add("_rels/.rels")
    while stack:
        rp = stack.pop()
        base = rp.replace("_rels/", "").removesuffix(".rels")
        for r in parse_rels(files[rp].decode("utf-8")):
            if r.get("TargetMode") == "External":
                continue
            tgt = resolve(base, r["Target"])
            if tgt in files and tgt not in reachable:
                reachable.add(tgt)
                rr = rels_path(tgt)
                if rr in files:
                    reachable.add(rr)
                    stack.append(rr)
    keep_always = {"[Content_Types].xml"}
    dropped = [p for p in files if p not in reachable and p not in keep_always and not p.endswith("/")]
    for p in dropped:
        del files[p]

    # ---- content types: drop overrides for deleted parts
    ct = files["[Content_Types].xml"].decode("utf-8")
    def keep_override(m):
        part = re.search(r'PartName="/([^"]+)"', m.group(0)).group(1)
        return m.group(0) if part in files else ""
    files["[Content_Types].xml"] = re.sub(r"<Override\b[^>]*/>", keep_override, ct).encode("utf-8")

    # ---- docProps/app.xml sheet list
    if "docProps/app.xml" in files:
        app = files["docProps/app.xml"].decode("utf-8")
        app = re.sub(r"<HeadingPairs>.*?</HeadingPairs>", "", app, flags=re.S)
        app = re.sub(r"<TitlesOfParts>.*?</TitlesOfParts>", "", app, flags=re.S)
        files["docProps/app.xml"] = app.encode("utf-8")

    # ---- write
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        order = ["[Content_Types].xml"] + [p for p in infos if p in files and p != "[Content_Types].xml"]
        for p in order:
            zout.writestr(p, files[p])

    return {
        "sheets_before": len(sheets), "sheets_after": len(kept), "removed_sheets": removed_sheets,
        "names_before": len(names), "names_after": len(final_names),
        "external_links_removed": removed_links, "parts_dropped": len(dropped),
        "size_before": src.stat().st_size, "size_after": dst.stat().st_size,
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path, nargs="?")
    ap.add_argument("--keep-hidden", default="", help="comma-separated hidden sheet names to keep anyway")
    a = ap.parse_args(argv)
    out = a.output or a.input.with_name(a.input.stem + "_CLEAN.xlsx")
    if out.resolve() == a.input.resolve():
        sys.exit("Refusing to overwrite the input file.")
    keep_hidden = {s.strip() for s in a.keep_hidden.split(",") if s.strip()}
    r = clean(a.input, out, keep_hidden)
    print(f"Wrote {out}")
    print(f"  Sheets:          {r['sheets_before']} -> {r['sheets_after']}  (removed {len(r['removed_sheets'])} hidden)")
    print(f"  Defined names:   {r['names_before']} -> {r['names_after']}")
    print(f"  External links removed: {r['external_links_removed']}")
    print(f"  Orphan parts dropped:   {r['parts_dropped']}")
    print(f"  File size: {r['size_before']/1e6:.2f} MB -> {r['size_after']/1e6:.2f} MB")
    print("  Removed sheets:")
    for s in r["removed_sheets"]:
        print("   -", s)


if __name__ == "__main__":
    main()
