#!/usr/bin/env python3
"""Census of OOXML package members that begin with a UTF-8 byte-order mark.

Walks the given roots for .docx/.docm/.dotx/.dotm/.xlsx/.xlsm/.xltx/.xltm/
.pptx/.pptm/.potx/.ppsx/.xlsb files, opens each as a ZIP, and lists every
member whose first three decompressed bytes are EF BB BF. Unreadable
archives are counted, not listed.
"""
import os
import sys
import zipfile

SUFFIXES = (".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm", ".xltx", ".xltm",
            ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".xlsb")
packages = 0
unreadable = 0
marked = []
for root in sys.argv[1:]:
    for directory, _, files in os.walk(root):
        for name in files:
            if not name.lower().endswith(SUFFIXES):
                continue
            path = os.path.join(directory, name)
            packages += 1
            try:
                with zipfile.ZipFile(path) as archive:
                    for info in archive.infolist():
                        if info.is_dir():
                            continue
                        try:
                            with archive.open(info) as member:
                                head = member.read(3)
                        except Exception:
                            continue
                        if head == b"\xef\xbb\xbf":
                            marked.append((path, info.filename))
            except Exception:
                unreadable += 1
for path, member in sorted(marked):
    print(f"{path}\t{member}")
print(f"# packages={packages} unreadable={unreadable} marked_members={len(marked)} "
      f"marked_packages={len({p for p, _ in marked})}", file=sys.stderr)
