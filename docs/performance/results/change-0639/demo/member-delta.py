#!/usr/bin/env python3
"""Census two OPC packages member by member.

Change 0639 uses this to show what the `office_crud_demo` PPTX update step
rewrites once it goes through `opened_presentation_transaction` instead of the
refused `presentation_mut`: every member the transaction did not name keeps its
exact source bytes, and the archive's member order is unchanged.

Usage: member-delta.py BEFORE.pptx AFTER.pptx
"""

from __future__ import annotations

import hashlib
import sys
import zipfile

# The two strings the demo's update writes, so the census also states that the
# edit landed in the one member that changed.
EXPECTED_STRINGS = (
    "Financial Performance — Q4 Final",
    "Thank you for your attention! Questions?",
)


def digest(archive: zipfile.ZipFile, name: str) -> str:
    return hashlib.sha256(archive.read(name)).hexdigest()


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print(__doc__.strip(), file=sys.stderr)
        return 2
    with zipfile.ZipFile(argv[1]) as before, zipfile.ZipFile(argv[2]) as after:
        names_before = [item.filename for item in before.infolist()]
        names_after = [item.filename for item in after.infolist()]
        print(
            f"members before={len(names_before)} after={len(names_after)} "
            f"order_identical={names_before == names_after}"
        )
        shared = [name for name in names_before if name in names_after]
        changed = [name for name in shared if digest(before, name) != digest(after, name)]
        identical = [name for name in shared if name not in changed]
        print(f"identical members={len(identical)} changed members={len(changed)}")
        for name in changed:
            print(
                f"  changed: {name} {len(before.read(name))} -> "
                f"{len(after.read(name))} bytes"
            )
        print("added:", [name for name in names_after if name not in names_before])
        print("removed:", [name for name in names_before if name not in names_after])
        for name in changed:
            text = after.read(name).decode("utf-8", "replace")
            for expected in EXPECTED_STRINGS:
                if expected in text:
                    print(f"  {name} carries: {expected!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
