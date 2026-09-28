"""Pure offline reader preflight for the 0829 phase-partition contract.

The preflight does not read live captures, execute a workload, or create
derived packet outputs. It exercises the partition helpers with synthetic
stacks and checks that the retained 0822 audit fixture can still be inspected
without being treated as a new timing result.
"""

from __future__ import annotations

import json
import importlib.util
from pathlib import Path
import sys

import analysis
import frame_audit
import custody


PACKET = Path(__file__).resolve().parent
BINARY = {"path": "/tmp/litchi-perf-0829-fp"}
OWNER = analysis.OWNER
DSO = BINARY["path"]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def custody_preflight() -> None:
    """Replay static custody before any build or capture is authorized."""
    cv = analysis.load_custody()
    frozen = custody.freeze()
    require(frozen.get("schema") == "litchi.performance.0829.frozen-inputs.v1",
            "static freeze witness schema changed")
    require(frozen.get("source", {}).get("revision") == custody.BASE_COMMIT,
            "static freeze source revision changed")
    for key in ("source", "tool", "probe", "root_inputs", "locks", "architecture",
                "corpus", "provenance", "host", "unrelated", "previous_seal",
                "plan", "origin", "drivers", "static_packet"):
        require(key in frozen, f"static freeze witness omits {key}")

    # quality.py is deliberately imported only for its read-only
    # production_reuse() proof. Its top-level guard refuses an existing
    # quality.json, so skip this import when the fresh quality gate has already
    # run and let analysis.load_quality validate the retained result later.
    if (PACKET / "quality.json").exists():
        quality = analysis.load_quality(cv)
        if (PACKET / "build.json").exists():
            build = analysis.load_build(cv, quality)
            custody.stable(build["frozen"])
            if (PACKET / "symbols/symbol.json").exists():
                analysis.symbol_evidence(PACKET / "perf", build, cv["plan"])
        return
    path = PACKET / "quality.py"
    spec = importlib.util.spec_from_file_location("litchi_perf_0829_quality_preflight", path)
    require(spec is not None and spec.loader is not None, "quality module cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    reuse = module.production_reuse()
    require(isinstance(reuse, dict)
            and reuse.get("schema") == "litchi.performance.0829.quality-reuse.v1"
            and reuse.get("cargo_commands_executed") is False,
            "quality production reuse preflight changed")


def frame(symbol: str, dso: str = DSO) -> dict[str, str]:
    return {"symbol": symbol, "dso": dso}


def synthetic_partition_cases() -> None:
    capture = analysis.PHASE_WRAPPERS["capture"]
    set_text = analysis.PHASE_WRAPPERS["set_text"]
    publish = analysis.PHASE_WRAPPERS["publish"]

    result = analysis.classify_phase_frames([frame("leaf"), frame(capture)], BINARY)
    require((result["status"], result["phase"]) == ("phase", "capture"),
            "exact capture phase was not classified")

    result = analysis.classify_phase_frames(
        [frame("leaf"), frame(capture), frame(set_text)], BINARY)
    require(result["status"] == "ambiguous" and result["phase"] is None
            and len(result["all_hits"]) == len(result["exact_hits"]) == 2
            and not result["other_dso_hits"],
            "multiple exact phase wrappers were not ambiguous")

    result = analysis.classify_phase_frames([frame("leaf"), frame("[unknown]")], BINARY)
    require(result["status"] == "unclassified" and result["phase"] is None
            and not result["all_hits"] and not result["exact_hits"]
            and not result["other_dso_hits"],
            "missing phase marker was not unclassified")

    result = analysis.classify_phase_frames(
        [frame("leaf"), frame(publish, "/tmp/another-dso")], BINARY)
    require(result["status"] == "unclassified" and result["phase"] is None
            and len(result["all_hits"]) == len(result["other_dso_hits"]) == 1
            and not result["exact_hits"],
            "wrong-DSO phase marker was assigned")

    # The independent implementation must make the same classification.
    for frames in (
        [frame("leaf"), frame(capture)],
        [frame("leaf"), frame(capture), frame(set_text)],
        [frame("leaf")],
        [frame("leaf"), frame(publish, "/tmp/another-dso")],
    ):
        left = analysis.classify_phase_frames(frames, BINARY)
        right = frame_audit.phase_partition(frames, DSO)
        require((left["status"], left["phase"], left["all_hits"],
                 left["exact_hits"], left["other_dso_hits"])
                == (right[0], right[1], right[2], right[3], right[4]),
                "primary and independent phase partition disagree")


def owner_partition_case() -> None:
    parsed = {"samples": [
        {"index": 0, "period": 10,
         "timestamp": "0", "frames": [frame("leaf"),
                                           frame(analysis.PHASE_WRAPPERS["capture"]),
                                           frame(OWNER)]},
        {"index": 1, "period": 20,
         "timestamp": "1", "frames": [frame("[unknown]"), frame(OWNER)]},
        {"index": 2, "period": 30,
         "timestamp": "2", "frames": [frame("leaf"),
                                           frame(analysis.PHASE_WRAPPERS["capture"]),
                                           frame(analysis.PHASE_WRAPPERS["set_text"]),
                                           frame(OWNER)]},
    ], "whole_process_samples": 3, "whole_process_period": 60}
    summary = analysis.exact_owner_summary(parsed, BINARY, "synthetic")
    partition = summary["phase_partition"]
    require(partition["sample_counts"] == {
        "capture": 1, "set_text": 0, "publish": 0,
        "unclassified": 1, "ambiguous": 1,
    }, "synthetic owner partition changed")
    require(sum(partition["sample_counts"].values()) == partition["owner_samples"] == 3,
            "synthetic owner partition is not exhaustive")
    require(sum(partition["periods"].values()) == summary["owner_qualified_period"] == 60,
            "synthetic owner period partition is not exhaustive")
    require(partition["additive_timing_claim"] is False
            and partition["causal_fraction_claim"] is False,
            "synthetic phase partition makes an additive or causal claim")


def empty_stack_repair_case() -> None:
    data = (
        b"probe 1 1.0: 10 cycles:u:\n\n"
        b"probe 1 2.0: 20 cycles:u:\n"
        b"  1 leaf (/tmp/litchi-perf-0829-fp)\n\n"
    )
    parsed = frame_audit.parse_frames(data)
    require(parsed["whole_samples"] == 2
            and len(parsed["samples"][0]["frames"]) == 0
            and len(parsed["samples"][1]["frames"]) == 1,
            "header-only empty stack was lost")


def historical_fixture_case() -> None:
    # This is a structural compatibility check only. The old packet remains
    # a prior experiment and contributes no 0829 timing values or claims.
    old = PACKET.parent / "change-0822" / "frame-audit.json"
    if not old.is_file():
        return
    value = json.loads(old.read_text(encoding="utf-8"))
    require(value.get("schema") == "litchi.performance.0822.frame-audit.v1",
            "old frame fixture schema changed")
    require(isinstance(value.get("repeats"), list),
            "old frame fixture repeats are malformed")


def symbol_fixture_case() -> None:
    old = PACKET.parent / "change-0828"
    value = custody.read(old / "symbol-diagnostic/result.json")
    for proof in value["proofs"].values():
        body = custody.verify_descriptor(proof["assembly"]).read_text()
        row = proof["demangled"]
        analysis.verify_owner_assembly(body, proof["owner"], row["address"], row["size"])
        for invalid in ("", (old / "symbols/edit_region-assembly.txt").read_text(),
                        body.replace(proof["owner"], "wrong_owner")):
            try:
                analysis.verify_owner_assembly(invalid, proof["owner"], row["address"], row["size"])
            except analysis.ReplayError:
                pass
            else:
                raise RuntimeError("invalid disassembly fixture admitted")


def main() -> int:
    if sys.argv[1:] not in (["--write"], ["--check"]):
        print("usage: reader_preflight.py --write|--check", file=sys.stderr)
        return 2
    try:
        custody_preflight()
        symbol_fixture_case()
        synthetic_partition_cases()
        owner_partition_case()
        empty_stack_repair_case()
        historical_fixture_case()
    except (AssertionError, KeyError, TypeError, ValueError, OSError, RuntimeError) as error:
        print(f"0829 reader preflight failed: {error}", file=sys.stderr)
        return 1
    print("0829 reader preflight PASS: phase partition, DSO/owner, empty-stack, and historical fixture checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
