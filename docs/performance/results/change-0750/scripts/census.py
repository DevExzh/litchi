"""Change 0750: audit every OOXML XML member in the repository's fixtures with
the before and after `xml-minifier`, and join the verdicts.

Usage: census.py <repo-root> <before-probe> <after-probe> <out-dir>

The corpus is every file under `test-data/`, `docs/performance/` and
`crates/*/tests/` that is a ZIP archive holding a `[Content_Types].xml`
member (an OPC package, whatever its extension), plus every loose `.xml` or
`.rels` file under `test-data/ooxml/`. A member is audited when publication
would audit it: `xml_minifier::audit::package::is_xml_part(name, media_type)`,
with the media type taken from the package's own content-type map.
"""

import collections
import json
import os
import struct
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ElementTree
import xml.parsers.expat

ROOTS = ["test-data", "docs/performance", "crates"]
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"


def is_xml_name(name):
    leaf = name.rsplit("/", 1)[-1]
    if leaf.lower() == "[content_types].xml":
        return True
    if "." not in leaf:
        return False
    return leaf.rsplit(".", 1)[1].lower() in ("xml", "rels", "rdf")


def is_xml_media_type(media_type):
    essence = media_type.split(";", 1)[0].strip().lower()
    return essence in ("application/xml", "text/xml") or essence.endswith("+xml")


def content_types(archive):
    names = {name.lower(): name for name in archive.namelist()}
    member = names.get("[content_types].xml")
    defaults, overrides = {}, {}
    if member is None:
        return defaults, overrides, False
    try:
        root = ElementTree.fromstring(archive.read(member))
    except Exception:
        return defaults, overrides, False
    for element in root:
        tag = element.tag.replace(CT_NS, "")
        if tag == "Default":
            defaults[element.get("Extension", "").lower()] = element.get("ContentType", "")
        elif tag == "Override":
            overrides[element.get("PartName", "").lower()] = element.get("ContentType", "")
    return defaults, overrides, True


def media_type(name, defaults, overrides):
    part = "/" + name.lstrip("/")
    if part.lower() in overrides:
        return overrides[part.lower()]
    leaf = name.rsplit("/", 1)[-1]
    if "." in leaf:
        return defaults.get(leaf.rsplit(".", 1)[1].lower(), "")
    return ""


def is_package(path):
    try:
        with open(path, "rb") as handle:
            if handle.read(4) != b"PK\x03\x04":
                return False
        with zipfile.ZipFile(path) as archive:
            return any(name.lower() == "[content_types].xml" for name in archive.namelist())
    except Exception:
        return False


def corpus(repo):
    packages, loose = [], []
    for root in ROOTS:
        base = os.path.join(repo, root)
        for directory, subdirectories, files in os.walk(base):
            subdirectories.sort()
            if root == "crates" and "/tests" not in directory.replace(base, ""):
                continue
            for leaf in sorted(files):
                path = os.path.join(directory, leaf)
                relative = os.path.relpath(path, repo)
                if is_package(path):
                    packages.append(relative)
                elif relative.startswith("test-data/ooxml/") and leaf.lower().endswith((".xml", ".rels")):
                    loose.append(relative)
    return packages, loose


def members(repo, packages, loose):
    rows, payload = [], bytearray()
    unreadable = []
    for relative in packages:
        path = os.path.join(repo, relative)
        try:
            archive = zipfile.ZipFile(path)
            defaults, overrides, typed = content_types(archive)
            for info in archive.infolist():
                name = info.filename
                if name.endswith("/"):
                    continue
                kind = media_type(name, defaults, overrides)
                if not (is_xml_name(name) or is_xml_media_type(kind)):
                    continue
                try:
                    data = archive.read(name)
                except Exception as error:
                    unreadable.append((relative, name, str(error)))
                    continue
                rows.append({"file": relative, "member": name, "media_type": kind,
                             "typed": typed, "bytes": len(data)})
                payload += struct.pack("<I", len(data)) + data
        except Exception as error:
            unreadable.append((relative, "", str(error)))
    for relative in loose:
        with open(os.path.join(repo, relative), "rb") as handle:
            data = handle.read()
        rows.append({"file": relative, "member": "", "media_type": "", "typed": False,
                     "bytes": len(data)})
        payload += struct.pack("<I", len(data)) + data
    return rows, bytes(payload), unreadable


def run(binary, payload):
    output = subprocess.run([binary], input=payload, stdout=subprocess.PIPE, check=True)
    return output.stdout.decode().splitlines()


def namespace_depth(data):
    """The most `xmlns:prefix` declarations in scope at once, by expat."""
    parser = xml.parsers.expat.ParserCreate()
    stack, state = [], {"now": 0, "most": 0}

    def start(_name, attributes):
        declared = sum(1 for key in attributes if key.startswith("xmlns:") and key != "xmlns:xml")
        stack.append(declared)
        state["now"] += declared
        state["most"] = max(state["most"], state["now"])

    def end(_name):
        state["now"] -= stack.pop()

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    try:
        parser.Parse(data, True)
    except Exception:
        return None
    return state["most"]


def main(repo, before, after, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    packages, loose = corpus(repo)
    rows, payload, unreadable = members(repo, packages, loose)
    before_lines = run(before, payload)
    after_lines = run(after, payload)
    assert len(before_lines) == len(after_lines) == len(rows)
    policies = ["source", "authored", "compact", "reader"]
    offset = 0
    for row, left, right in zip(rows, before_lines, after_lines):
        length = row["bytes"]
        data = payload[offset + 4: offset + 4 + length]
        offset += 4 + length
        for policy, value in zip(policies, left.split("\t")):
            row["before_" + policy] = value
        for policy, value in zip(policies, right.split("\t")):
            row["after_" + policy] = value
        if row["after_source"] == "OK":
            row["namespaces_in_scope"] = namespace_depth(data)
    with open(os.path.join(out_dir, "census.jsonl"), "w") as handle:
        for row in rows:
            handle.write(json.dumps(row, sort_keys=True) + "\n")
    with open(os.path.join(out_dir, "corpus.txt"), "w") as handle:
        handle.write("# packages\n")
        handle.writelines(path + "\n" for path in packages)
        handle.write("# loose XML parts\n")
        handle.writelines(path + "\n" for path in loose)
        handle.write("# unreadable\n")
        handle.writelines("\t".join(item) + "\n" for item in unreadable)
    print(f"packages {len(packages)}, loose parts {len(loose)}, XML members {len(rows)}, "
          f"unreadable {len(unreadable)}")


if __name__ == "__main__":
    main(*sys.argv[1:5])
