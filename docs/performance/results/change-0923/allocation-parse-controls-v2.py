#!/usr/bin/env python3
"""Exercise fail-closed Heaptrack parser controls without running a benchmark."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import subprocess
import sys
import os
from pathlib import Path


HERE = Path(__file__).resolve().parent
PARSER_PATH = HERE / "allocation-parse-v2.py"
ZSTD = Path("/usr/bin/zstd")
OUTPUT = HERE / "allocation-parser-controls-v2.json"
CONTROL_ROOT = Path(os.environ.get(
    "LITCHI_PERF_0923_PARSER_CONTROL_ROOT",
    "/tmp/litchi-performance-0923/parser-controls-v2",
)).resolve()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, object]:
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def parser_module():
    spec = importlib.util.spec_from_file_location("litchi_0923_allocation_parser", PARSER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import {PARSER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


VALID = """v 10500 3
X parser-control-fixture
s 1 x
i 0 1 0
t 0 0
a a 0
+ 0
- 0
R 0
# strings: 1
# ips: 1
"""


def expect_rejection(name: str, text: str, temporary: Path, parser) -> dict[str, object]:
    source = temporary / f"{name}.interpreted"
    profile = temporary / f"{name}.zst"
    source.write_text(text, encoding="utf-8")
    result = subprocess.run(
        [str(ZSTD), "-q", "-f", "-o", str(profile), str(source)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"zstd fixture compression failed for {name}: {result.stderr}")
    parse_result = subprocess.run(
        [sys.executable, "-B", str(PARSER_PATH), str(profile)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )
    if parse_result.returncode == 0:
        raise RuntimeError(f"parser accepted negative control {name}")
    # Parse once in this process as well.  This exercises the generator's
    # finally path directly, where an early ParseError must reap zstdcat.
    try:
        parser.parse(profile)
    except parser.ParseError:
        pass
    else:
        raise RuntimeError(f"in-process parser accepted negative control {name}")
    return {
        "name": name,
        "expected": "reject",
        "exit_code": parse_result.returncode,
        "stderr": parse_result.stderr[-300:],
        "profile": descriptor(profile),
    }


def main() -> int:
    parser = parser_module()
    cases = {
        "malformed_version": VALID.replace("v 10500 3", "v 99999 3"),
        "negative_string_reference": VALID.replace("i 0 1 0", "i 0 1 -1"),
        "negative_trace_reference": VALID.replace("t 0 0", "t -1 0"),
        "negative_allocation_reference": VALID.replace("+ 0", "+ -1"),
        "out_of_range_trace_reference": VALID.replace("t 0 0", "t 2 0"),
        "out_of_range_allocation_trace": VALID.replace("a a 0", "a a 2"),
        "out_of_range_allocation_reference": VALID.replace("+ 0", "+ 1"),
        "truncated_profile": VALID.replace("# ips: 1\n", ""),
        "deallocation_underflow": VALID.replace("- 0\nR 0", "- 0\n- 0\nR 0"),
        "double_free_with_nonnegative_total": VALID.replace(
            "a a 0\n+ 0\n- 0", "a 0 0\n+ 0\n- 0\n- 0"
        ),
    }
    if CONTROL_ROOT.exists():
        raise RuntimeError(f"refusing to replace retained control profiles: {CONTROL_ROOT}")
    CONTROL_ROOT.mkdir(parents=True)
    controls = [expect_rejection(case, text, CONTROL_ROOT, parser)
                for case, text in cases.items()]
    valid_source = CONTROL_ROOT / "valid.interpreted"
    valid_profile = CONTROL_ROOT / "valid.zst"
    valid_source.write_text(VALID, encoding="utf-8")
    result = subprocess.run([str(ZSTD), "-q", "-f", "-o", str(valid_profile), str(valid_source)],
                            check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"zstd valid fixture compression failed: {result.stderr}")
    parsed = parser.parse(valid_profile)
    if parsed["allocation_event_count"] != 1 or parsed["deallocation_event_count"] != 1:
        raise RuntimeError("valid fixture totals changed")
    value = {
        "schema": "litchi.performance.0923.heaptrack-parser-controls.v2",
        "status": "pass",
        "fixture_root": str(CONTROL_ROOT),
        "parser": {"path": str(PARSER_PATH), "bytes": PARSER_PATH.stat().st_size,
                    "sha256": sha256(PARSER_PATH)},
        "compression_tool": descriptor(ZSTD),
        "valid_fixture": {
            "profile": descriptor(valid_profile),
            "format": parsed["format"],
            "allocation_event_count": parsed["allocation_event_count"],
            "deallocation_event_count": parsed["deallocation_event_count"],
            "requested_bytes": parsed["requested_bytes"],
            "peak_live_heap_bytes": parsed["peak_live_heap_bytes"],
            "live_bytes_at_end": parsed["live_bytes_at_end"],
        },
        "controls": controls,
        "scope": "parser_integrity_only; no benchmark or operation claim",
    }
    if OUTPUT.exists():
        raise RuntimeError(f"refusing to replace retained control receipt: {OUTPUT}")
    OUTPUT.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps(value, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
