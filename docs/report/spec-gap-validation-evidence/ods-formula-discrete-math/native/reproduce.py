#!/usr/bin/env python3
"""Reproduce the discrete native receipt from pinned raw GitHub fixtures.

The downloaded source tree and generated output are temporary and removed
automatically when this process exits.  The repository's LibreOffice fixture
checkout is never modified.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
import tempfile
import urllib.request

sys.dont_write_bytecode = True

from extract import RAW_URL, SOURCES, UPSTREAM_COMMIT, extract


ROOT = Path(__file__).resolve().parent
MAX_FILE_BYTES = 2_000_000


def main() -> None:
    expected = json.loads((ROOT / "provenance.json").read_text())
    expected_inputs = expected["inputs"]
    with tempfile.TemporaryDirectory(prefix="litchi-ods-discrete-pinned-") as temporary:
        source = Path(temporary) / "libreoffice-core"
        output = Path(temporary) / "evidence"
        for relative in SOURCES.values():
            url = f"{RAW_URL}/{relative}"
            with urllib.request.urlopen(url, timeout=60) as response:
                data = response.read(MAX_FILE_BYTES + 1)
            if len(data) > MAX_FILE_BYTES:
                raise RuntimeError(f"pinned fixture exceeds bounded download size: {relative}")
            digest = hashlib.sha256(data).hexdigest()
            if digest != expected_inputs[relative]:
                raise RuntimeError(
                    f"raw SHA-256 mismatch for {relative}: {digest} != {expected_inputs[relative]}"
                )
            path = source / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(data)

        output.mkdir(parents=True, exist_ok=True)
        extract(source, output)
        for name in ("cached-results.json", "provenance.json"):
            generated = (output / name).read_bytes()
            retained = (ROOT / name).read_bytes()
            if generated != retained:
                raise RuntimeError(f"reproduced {name} differs from retained evidence")

    print(
        json.dumps(
            {
                "commit": UPSTREAM_COMMIT,
                "files": len(SOURCES),
                "selected_observations": expected["selected_observations"],
                "temporary_tree_cleaned": True,
            },
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
