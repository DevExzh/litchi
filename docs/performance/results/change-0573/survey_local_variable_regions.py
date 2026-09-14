#!/usr/bin/env python3
"""Size the strict local-header window from the repository fixture corpus.

Usage: survey_local_variable_regions.py <root> [<root> ...]

For every ZIP container under the given roots, walks the central directory,
reads each member's LOCAL header, and reports the distribution of
`file_name_len + extra_field_len`. That value plus the 30-byte fixed header is
the smallest window that proves a member's layout in one positional read.
"""

import collections
import os
import struct
import sys


def members(path):
    with open(path, "rb") as handle:
        data = handle.read()
    eocd = data.rfind(b"PK\x05\x06")
    if eocd < 0:
        return None
    total = struct.unpack_from("<H", data, eocd + 10)[0]
    directory = struct.unpack_from("<I", data, eocd + 16)[0]
    if directory == 0xFFFFFFFF:
        return None  # ZIP64 locator; none in this corpus
    out, cursor = [], directory
    for _ in range(total):
        if data[cursor : cursor + 4] != b"PK\x01\x02":
            return None
        name, extra, comment = struct.unpack_from("<HHH", data, cursor + 28)
        local = struct.unpack_from("<I", data, cursor + 42)[0]
        cursor += 46 + name + extra + comment
        if data[local : local + 4] != b"PK\x03\x04":
            return None
        local_name, local_extra = struct.unpack_from("<HH", data, local + 26)
        out.append((local_name, local_extra))
    return out


def main(roots):
    sizes = collections.Counter()
    extras = collections.Counter()
    archives = entries = 0
    per_file_max = {}
    for root in roots:
        for directory, _, names in os.walk(root):
            for name in sorted(names):
                path = os.path.join(directory, name)
                try:
                    found = members(path)
                except Exception:
                    found = None
                if not found:
                    continue
                archives += 1
                entries += len(found)
                per_file_max[path] = max(n + e for n, e in found)
                for local_name, local_extra in found:
                    sizes[local_name + local_extra] += 1
                    extras[local_extra] += 1
    print(f"roots={roots} archives={archives} members={entries}")
    print(f"max local variable region = {max(sizes)} bytes "
          f"(one-read window {30 + max(sizes)})")
    print()
    print("local variable-region size -> member count")
    for size in sorted(sizes):
        print(f"  var={size:5d}  window={30 + size:5d}  members={sizes[size]}")
    print()
    print("local extra_field_len -> member count")
    for size in sorted(extras):
        print(f"  extra={size:5d}  members={extras[size]}")
    print()
    print("candidate window -> members proved in one read")
    for window in (128, 256, 320, 512, 576, 640, 1024, 2048):
        fits = sum(c for size, c in sizes.items() if 30 + size <= window)
        print(f"  window={window:5d}: {fits}/{entries} fit, {entries - fits} fall back")
    print()
    print("archives whose largest member exceeds a 640-byte window")
    for path, largest in sorted(per_file_max.items()):
        if 30 + largest > 640:
            print(f"  {path} max_var={largest} window_needed={30 + largest}")


if __name__ == "__main__":
    main(sys.argv[1:] or ["test-data"])
