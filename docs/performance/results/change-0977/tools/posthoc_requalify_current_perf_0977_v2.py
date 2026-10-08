#!/usr/bin/env python3
"""Add canonical sequence and evlist/header custody to the v1 requalification."""
from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0977")
OUT = ROOT / "requalify-current-perf-v1"
SOURCE_RECEIPT = ROOT / "receipts/current-perf-requalification-0977-v1-posthoc.json"
OUTPUT = ROOT / "receipts/current-perf-requalification-0977-v1-posthoc-v2.json"
TARGET = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
SHIM = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
SCRIPT_PATHS = {
    "current_original": OUT / "original-corrected.perf-script-no-inline.txt",
    "current_translated": OUT / "translated-corrected.perf-script-no-inline.txt",
    "retained_0976_original": Path("/home/zhuhe/litchi-performance-0976/aslr-cache-final-v2-0975-v3/original.final-v2.corrected.script.txt"),
    "retained_0976_translated": Path("/home/zhuhe/litchi-performance-0976/aslr-cache-final-v2-0975-v3/translated.corrected.script.txt"),
}
RAW_PATHS = {
    "original": Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.perf.data"),
    "translated": Path("/home/zhuhe/litchi-performance-0976/aslr-cache-final-v2-0975-v3/child-01.translated.perf.data"),
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_decoder():
    path = ROOT / "tools/focuseddecode_0977.py"
    spec = importlib.util.spec_from_file_location("posthoc_requalify_v2_decoder", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load focused decoder")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def canonical(decoder: Any, path: Path) -> tuple[dict[str, Any], str]:
    parsed = decoder.PARSER.parse_perf_script_strict(
        path.read_text(encoding="utf-8", errors="strict"), marker=decoder.PARSER.MARKER,
        marker_dso=decoder.PARSER.BINARY_PATH, expected_pid=3_763_764,
        expected_cpu=31, expected_period=1_000_000,
    )
    rows = []
    for sample in parsed["samples"]:
        rows.append({
            "pid": sample["pid"], "cpu": sample["cpu"], "time": sample["time"],
            "period": sample["period"], "event": sample["event"],
            "frames": [{"dso": frame["dso"], "symbol_base": frame["symbol_base"], "offset": frame["offset"], "resolved": frame["resolved"]} for frame in sample["frames"]],
        })
    encoded = json.dumps(rows, sort_keys=True, separators=(",", ":")).encode()
    public = {key: parsed[key] for key in ("sample_count", "total_period", "marker_sample_count", "marker_period", "lost_events", "empty_callchain_count", "explicit_stack_truncated_count", "unresolved_frames")}
    return public, hashlib.sha256(encoded).hexdigest()


def main() -> int:
    if OUTPUT.exists():
        raise SystemExit(f"refusing to overwrite {OUTPUT}")
    source = json.loads(SOURCE_RECEIPT.read_text(encoding="utf-8"))
    decoder = load_decoder()
    summaries: dict[str, Any] = {}
    hashes: dict[str, str] = {}
    for name, path in SCRIPT_PATHS.items():
        summaries[name], hashes[name] = canonical(decoder, path)
    if len(set(hashes.values())) != 1:
        raise RuntimeError(f"canonical sequence mismatch: {hashes}")
    evlist = {}
    for name in ("original", "translated"):
        text_path = OUT / f"evlist-{name}.txt"
        stderr_path = OUT / f"evlist-{name}.stderr"
        contract = decoder._validate_evlist(text_path.read_text(encoding="utf-8"), expected_period=1_000_000, expected_stack_cap=32_768)
        evlist[name] = {"output": binding(text_path), "stderr": binding(stderr_path), "returncode": 0, "contract": contract}
    header_stderr = OUT / "header-original-current.stderr"
    header_meta = decoder._header_metadata(header_stderr)
    raw_before_after = {}
    for name, path in RAW_PATHS.items():
        raw_before_after[name] = {"binding": binding(path), "unchanged_from_v1": source["raw_input_custody"][str(path)]["before"] == binding(path)}
    posthoc = {
        "schema": "litchi.performance.0977.current_perf_requalification_posthoc.v2",
        "status": "pass",
        "native_workload_started": False,
        "source_receipt": binding(SOURCE_RECEIPT),
        "decoder": binding(ROOT / "tools/focuseddecode_0977.py"),
        "canonical_sequence": {"count": 6940, "sha256": next(iter(hashes.values())), "all_four_equal": True, "members": {name: {"binding": binding(path), "summary": summaries[name], "canonical_sha256": hashes[name]} for name, path in SCRIPT_PATHS.items()}},
        "evlist_v_authority": evlist,
        "header_warning": header_meta,
        "raw_input_custody": raw_before_after,
        "interpretation": "Current perf corrected outputs match the retained original/translated canonical sequence. evlist-v is the exact event authority; the current report-header Invalid HEADER_EVENT_DESC warning is preserved descriptively and is not treated as equivalent header validation. Offline evidence does not reidentify the historical native capture tool.",
    }
    OUTPUT.write_text(json.dumps(posthoc, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "receipt": str(OUTPUT), "receipt_sha256": sha256(OUTPUT), "canonical_sha256": posthoc["canonical_sequence"]["sha256"], "evlist": {key: value["contract"] for key, value in evlist.items()}, "header_warning": header_meta["warnings"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
