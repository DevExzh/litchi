#!/usr/bin/env python3
"""Reproduce the pinned paired-statistics native receipt.

Raw FODS inputs are downloaded into a bounded temporary tree, checked against
the retained SHA-256 manifest, extracted, and compared byte-for-byte with the
committed evidence.  The helper never modifies a source checkout and removes
the temporary tree on exit.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
import tempfile
import urllib.request

sys.dont_write_bytecode = True

from extract import RAW_URL, UPSTREAM_COMMIT, extract


ROOT = Path(__file__).resolve().parent
MAX_FILE_BYTES = 2_000_000


def main() -> None:
    expected = json.loads((ROOT / "provenance.json").read_text())
    expected_inputs = expected["inputs"]
    with tempfile.TemporaryDirectory(prefix="litchi-ods-paired-statistics-pinned-") as temporary_name:
        temporary = Path(temporary_name)
        source = temporary / "libreoffice-core"
        output = temporary / "evidence"
        for relative, expected_digest in expected_inputs.items():
            url = f"{RAW_URL}/{relative}"
            with urllib.request.urlopen(url, timeout=60) as response:
                data = response.read(MAX_FILE_BYTES + 1)
            if len(data) > MAX_FILE_BYTES:
                raise RuntimeError(f"pinned fixture exceeds bounded download size: {relative}")
            digest = hashlib.sha256(data).hexdigest()
            if digest != expected_digest:
                raise RuntimeError(
                    f"raw SHA-256 mismatch for {relative}: {digest} != {expected_digest}"
                )
            path = source / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(data)

        output.mkdir(parents=True, exist_ok=True)
        extract(source, output)
        for name in (
            "cached-results.json",
            "provenance.json",
            "steyx-deviation-proof.json",
        ):
            generated = (output / name).read_bytes()
            retained = (ROOT / name).read_bytes()
            if generated != retained:
                raise RuntimeError(f"reproduced {name} differs from retained evidence")

    print(
        json.dumps(
            {
                "commit": UPSTREAM_COMMIT,
                "files": len(expected_inputs),
                "selected_observations": expected["selected_observations"],
                "temporary_tree_cleaned": True,
                "converter": False,
            },
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
