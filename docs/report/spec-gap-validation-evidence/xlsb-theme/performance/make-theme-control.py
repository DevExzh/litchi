#!/usr/bin/env python3
"""Create a deterministic valid large-theme control from a native XLSB fixture.

The control adds an XML comment outside the modeled theme elements.  The
comment exercises package/source-byte retention while leaving the typed theme
model at the same bounded size.  The resulting workbook is validated by the
existing vendored Transitional DrawingML schema checker before profiling.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import zipfile


THEME_SUFFIX = "/theme/theme1.xml"
ROOT_END = b"</a:theme>"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--comment-bytes", type=int, default=1 << 20)
    args = parser.parse_args()
    if args.comment_bytes < 32:
        raise SystemExit("--comment-bytes must be at least 32")
    args.output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(args.base, "r") as source:
        theme_names = [
            info.filename
            for info in source.infolist()
            if info.filename.endswith(THEME_SUFFIX)
        ]
        if len(theme_names) != 1:
            raise SystemExit(f"expected exactly one theme1.xml, found {theme_names}")
        theme_name = theme_names[0]
        comment_prefix = b"<!--litchi-theme-profile:"
        comment_suffix = b"-->"
        payload_len = args.comment_bytes - len(comment_prefix) - len(comment_suffix)
        comment = comment_prefix + (b"x" * payload_len) + comment_suffix
        original = source.read(theme_name)
        if original.count(ROOT_END) != 1:
            raise SystemExit("theme XML did not contain one a:theme closing tag")
        replacement = original.replace(ROOT_END, comment + ROOT_END, 1)
        with zipfile.ZipFile(args.output, "w") as target:
            for info in source.infolist():
                data = replacement if info.filename == theme_name else source.read(info)
                target.writestr(info, data)

    print(args.output)
    print(f"theme_name={theme_name}")
    print(f"comment_bytes={len(comment)}")
    print(f"theme_bytes={len(replacement)}")


if __name__ == "__main__":
    main()
