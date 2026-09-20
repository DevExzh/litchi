#!/usr/bin/env python3
"""Record the frozen 0704 driver/probe tree and its source binding."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


paths = sorted(
    [
        *P.glob("*.py"),
        *(P / "probe").rglob("*"),
        *(P / "refusal-probe").rglob("*"),
        *((P / "retention-probe").rglob("*") if (P / "retention-probe").exists() else []),
        *((P / "mechanism").rglob("*") if (P / "mechanism").exists() else []),
        P / "README.md",
        P / "cases.json",
        P / "control-manifest.json",
        P / "functional-test-inventory.json",
        *[P / name for name in ("retention-observer-README.md",)
          if (P / name).exists()],
    ]
)
paths = [
    path for path in paths
    if path.is_file()
    and path.name != "scripts-manifest.py"
    # Observer stdout/stderr are raw evidence, not frozen driver inputs.  If
    # they were included here, rerunning this manifest after the observer
    # would make the manifest self-invalidating.
    and not any(
        str(path.relative_to(P)).startswith(prefix)
        for prefix in (
            "retention-probe/results/",
            "mechanism/results/",
            "mechanism/trace-runs/",
            "mechanism/raw/",
        )
    )
]
baseline = json.loads((P / "baseline.json").read_text())
receipt = {
    "baseline_head": baseline["baseline_head"],
    "baseline_source_file_count": len(baseline["source_sha256"]),
    "packet_files": {
        str(path.relative_to(P)): {"bytes": path.stat().st_size, "sha256": sha(path)}
        for path in paths
    },
    "scope": "0704 scripts, standalone probes, controls, and frozen manifests; raw measurements are excluded",
}
(P / "scripts-manifest.json").write_text(json.dumps(receipt, indent=2) + "\n")
print("recorded", len(paths), "frozen packet files")
