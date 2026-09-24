#!/usr/bin/env python3
"""Census of per-element attribute and namespace-declaration counts in OOXML packages.

Scans every ZIP package (detected by content: local-file magic plus a
[Content_Types].xml member) under the given roots, parses every member whose
name ends in .xml, .rels or .vml with expat (namespace processing off, so
xmlns declarations are ordinary attributes, as quick-xml sees them), and
records per element: attributes, namespace declarations, the declarations in
scope across the open-element stack (shadowed re-declarations included) and
depth. Writes a JSON summary.
"""
import hashlib
import json
import os
import sys
import zipfile
import xml.parsers.expat

XML_SUFFIXES = (".xml", ".rels", ".vml")


def is_package(path):
    try:
        with open(path, "rb") as handle:
            if handle.read(4) != b"PK\x03\x04":
                return False
        with zipfile.ZipFile(path) as archive:
            names = {info.filename for info in archive.infolist()}
            return "[Content_Types].xml" in names or "META-INF/manifest.xml" in names
    except Exception:
        return False


def census_member(data):
    stats = {"elements": 0, "max_attributes": 0, "max_declarations": 0,
             "max_in_scope_declarations": 0, "max_in_scope_distinct": 0, "max_depth": 0,
             "top_element": None, "top_declarations_element": None}
    stack = []  # per open element: list of declared prefixes
    in_scope = [0]
    parser = xml.parsers.expat.ParserCreate()
    parser.ordered_attributes = True

    def start(name, attrs):
        count = len(attrs) // 2
        declared = [attrs[i] for i in range(0, len(attrs), 2)
                    if attrs[i] == "xmlns" or attrs[i].startswith("xmlns:")]
        stats["elements"] += 1
        if count > stats["max_attributes"]:
            stats["max_attributes"] = count
            stats["top_element"] = name
        if len(declared) > stats["max_declarations"]:
            stats["max_declarations"] = len(declared)
            stats["top_declarations_element"] = name
        stack.append(declared)
        in_scope[0] += len(declared)
        stats["max_in_scope_declarations"] = max(stats["max_in_scope_declarations"], in_scope[0])
        distinct = set()
        for frame in stack:
            distinct.update(frame)
        stats["max_in_scope_distinct"] = max(stats["max_in_scope_distinct"], len(distinct))
        stats["max_depth"] = max(stats["max_depth"], len(stack))

    def end(_name):
        declared = stack.pop()
        in_scope[0] -= len(declared)

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(data, True)
    return stats


def main():
    roots = sys.argv[2:]
    output = sys.argv[1]
    packages = []
    for root in roots:
        for directory, _dirs, files in os.walk(root):
            for name in files:
                path = os.path.join(directory, name)
                if is_package(path):
                    packages.append(path)
    packages.sort()
    summary = {"packages": 0, "members": 0, "elements": 0, "parse_failures": [],
               "max_attributes": 0, "max_declarations": 0, "max_in_scope_declarations": 0,
               "max_in_scope_distinct": 0, "max_depth": 0,
               "top_attributes": [], "top_declarations": [], "top_in_scope": [],
               "attribute_histogram": {}}
    rows = []
    seen_digests = set()
    for path in packages:
        try:
            archive = zipfile.ZipFile(path)
        except Exception as error:  # noqa: BLE001
            summary["parse_failures"].append({"package": path, "error": str(error)})
            continue
        with archive:
            summary["packages"] += 1
            for info in archive.infolist():
                if not info.filename.lower().endswith(XML_SUFFIXES):
                    continue
                try:
                    data = archive.read(info)
                except Exception as error:  # noqa: BLE001
                    summary["parse_failures"].append({"package": path, "member": info.filename, "error": str(error)})
                    continue
                digest = hashlib.sha256(data).hexdigest()
                duplicate_member = digest in seen_digests
                seen_digests.add(digest)
                try:
                    stats = census_member(data)
                except Exception as error:  # noqa: BLE001
                    summary["parse_failures"].append({"package": path, "member": info.filename, "error": str(error)})
                    continue
                summary["members"] += 1
                summary["elements"] += stats["elements"]
                for key in ("max_attributes", "max_declarations", "max_in_scope_declarations",
                            "max_in_scope_distinct", "max_depth"):
                    summary[key] = max(summary[key], stats[key])
                rows.append((path, info.filename, stats, duplicate_member))
    def top(key, element_key):
        ranked = sorted(rows, key=lambda row: row[2][key], reverse=True)[:12]
        return [{"package": row[0], "member": row[1], key: row[2][key],
                 "element": row[2].get(element_key) if element_key else None} for row in ranked]
    summary["top_attributes"] = top("max_attributes", "top_element")
    summary["top_declarations"] = top("max_declarations", "top_declarations_element")
    summary["top_in_scope"] = top("max_in_scope_declarations", None)
    histogram = {}
    for _path, _member, stats, _dup in rows:
        bucket = stats["max_attributes"]
        bucket = 1 << (bucket.bit_length()) if bucket else 0
        histogram[bucket] = histogram.get(bucket, 0) + 1
    summary["attribute_histogram"] = {str(k): v for k, v in sorted(histogram.items())}
    summary["unique_member_digests"] = len(seen_digests)
    with open(output, "w", encoding="utf-8") as handle:
        json.dump(summary, handle, indent=1)
    print(json.dumps({k: summary[k] for k in ("packages", "members", "unique_member_digests", "elements",
                                                "max_attributes", "max_declarations",
                                                "max_in_scope_declarations", "max_in_scope_distinct",
                                                "max_depth")}, indent=1))
    print("parse failures:", len(summary["parse_failures"]))
    for item in summary["top_attributes"][:6]:
        print("  top attrs", item)
    for item in summary["top_declarations"][:4]:
        print("  top decls", item)
    for item in summary["top_in_scope"][:4]:
        print("  top in-scope", item)


if __name__ == "__main__":
    main()
