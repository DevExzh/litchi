#!/usr/bin/env python3
"""Decompose an OOXML open's relationship-part reads into the two mechanisms.

`PackageReader::source_catalog` reads a `*/_rels/*.rels` member through exactly
two mechanisms, and they differ in what the open uses the result for:

  BFS      `walk_relationship_graph` -- runs BEFORE `classify_part_members`.
           Its key set (the reachable internal target names) is what decides
           whether an untyped ZIP member is tolerated junk or a hard
           `ContentTypeNotFound`. Reading these is load-bearing for the open's
           own verdict.

  ORPHAN   the post-classification fallback loop -- runs AFTER
           `classify_part_members`, for each admitted typed part the BFS never
           visited. Its result cannot influence admission, which has already
           happened. It is retained for later query and charged to the
           relationship ledger.

This script reproduces both mechanisms from the package bytes alone. It never
asks the library anything.
"""
import sys, os, zipfile, json, posixpath
import xml.etree.ElementTree as ET

NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
NS_CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"


def is_rels_member(name):
    if not name.lower().endswith(".rels"):
        return False
    d = posixpath.dirname(name)
    return d == "_rels" or d.endswith("/_rels")


def rels_member_for(partname):
    """`/word/document.xml` -> `word/_rels/document.xml.rels`; `/` -> `_rels/.rels`."""
    p = partname.lstrip("/")
    d, b = posixpath.dirname(p), posixpath.basename(p)
    return posixpath.join(d, "_rels", b + ".rels") if d else \
        (f"_rels/{b}.rels" if b else "_rels/.rels")


def parse_rels(data, base_uri):
    """Return the list of internal target part names, resolved against base_uri."""
    out = []
    root = ET.fromstring(data)
    for r in root:
        if r.tag != NS_REL + "Relationship":
            continue
        if r.get("TargetMode") == "External":
            continue
        t = r.get("Target", "")
        if t.startswith("/"):
            out.append(posixpath.normpath(t))
        else:
            out.append(posixpath.normpath(posixpath.join(base_uri, t)))
    return out


def base_uri_of(partname):
    d = posixpath.dirname(partname)
    return d if d else "/"


def analyse(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    nameset = set(names)
    rels_members = [n for n in names if is_rels_member(n)]

    # --- content types ------------------------------------------------------
    ct_member = next(n for n in names if n.lower() == "[content_types].xml")
    ct_root = ET.fromstring(z.read(ct_member))
    defaults, overrides = {}, {}
    for e in ct_root:
        if e.tag == NS_CT + "Default":
            defaults[e.get("Extension", "").lower()] = e.get("ContentType", "")
        elif e.tag == NS_CT + "Override":
            overrides[e.get("PartName", "")] = e.get("ContentType", "")

    def content_type_of(partname):
        if partname in overrides:
            return overrides[partname]
        ext = posixpath.splitext(partname)[1].lstrip(".").lower()
        return defaults.get(ext)

    # --- mechanism 1: the BFS ----------------------------------------------
    bfs_read = []          # rels members the BFS actually reads
    visited = set()        # reachable internal target part names (the key set)
    root_rels = "_rels/.rels"
    queue = []
    if root_rels in nameset:
        for t in parse_rels(z.read(root_rels), "/"):
            if t not in visited:
                visited.add(t); queue.append(t)
    while queue:
        part = queue.pop()
        rm = rels_member_for(part)
        if rm in nameset:
            bfs_read.append(rm)
            try:
                targets = parse_rels(z.read(rm), base_uri_of(part))
            except ET.ParseError:
                targets = []
            for t in targets:
                if t not in visited:
                    visited.add(t); queue.append(t)

    # --- mechanism 2: classification, then the orphan fallback --------------
    typed_parts = []
    for n in names:
        if not n or n.endswith("/") or n == ct_member or is_rels_member(n):
            continue
        pn = "/" + n
        if content_type_of(pn) is None and pn not in visited:
            continue        # UntypedAndUnreferenced: tolerated archive junk
        typed_parts.append(pn)

    orphan_read = []
    for pn in typed_parts:
        if pn in visited:
            continue        # already carried by the BFS
        rm = rels_member_for(pn)
        if rm in nameset:
            orphan_read.append(rm)

    read = set(bfs_read) | set(orphan_read)
    main = {"xl/workbook.xml", "ppt/presentation.xml"} & nameset
    # `_rels/.rels` is read by its own `load_rels_lazy(package_uri)` call
    # before the walk, so it is a structural member but not a walk read.
    package_rels = 1 if root_rels in nameset else 0
    return {
        "fixture": os.path.basename(path),
        "members": len(names),
        "rels_members_present": len(rels_members),
        "rels_read_total": len(read),
        "rels_read_by_bfs": len(bfs_read),
        "rels_read_by_orphan_fallback": len(orphan_read),
        "rels_never_read": sorted(set(rels_members) - read - {root_rels}),
        "structural_members_modelled":
            1 + package_rels + len(read) + len(main),
    }


if __name__ == "__main__":
    rows = [analyse(p) for p in sys.argv[1:]]
    print(json.dumps(rows, indent=2))
    print()
    hdr = ("fixture", "mem", "rels", "read", "bfs", "orphan", "struct")
    print("%-34s %5s %5s %5s %5s %7s %7s" % hdr)
    for r in rows:
        print("%-34s %5d %5d %5d %5d %7d %7d" % (
            r["fixture"], r["members"], r["rels_members_present"],
            r["rels_read_total"], r["rels_read_by_bfs"],
            r["rels_read_by_orphan_fallback"], r["structural_members_modelled"]))
