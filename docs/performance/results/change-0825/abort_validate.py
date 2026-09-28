"""Offline validation of the failed 0825 qualification, without timing analysis."""
import hashlib
import json
from pathlib import Path
import re
import sys

import artifact_audit
import custody
import preservation

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def read(path):
    return json.loads(path.read_text())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify(value):
    path = Path(value["path"])
    assert path.is_absolute() and path.is_file() and not path.is_symlink(), path
    assert value == {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}, path
    return path


def check(final=False):
    plan = read(P / "plan.json")
    frozen = read(P / "freeze.json")
    assert plan["base"] == frozen["base"] == custody.BASE
    assert frozen["source"] == frozen["after_source"]
    for key in ("plan", "quality", "host", "toolchain", "root_input_manifest",
                "architecture_manifest", "corpus_manifest", "lock_parity", "origin", "provenance"):
        verify(frozen[key])
    for group in ("drivers", "support"):
        for name, digest in frozen[group].items():
            assert sha(P / name) == digest, name
    for group in ("tool", "root_inputs", "architecture", "unrelated"):
        for name, digest in frozen[group].items():
            assert sha(ROOT / name) == digest, name
    assert len(frozen["architecture"]) == 35
    for name, value in frozen["corpus"].items():
        assert value == {"bytes": (ROOT / name).stat().st_size, "sha256": sha(ROOT / name)}
    for value in frozen["candidate_archives"].values():
        verify(value)
    after = frozen["source"]
    before = {"revision": plan["base"], "files": dict(after["files"])}
    for name, digest in after["files"].items():
        assert sha(ROOT / name) == digest, name
    for name in plan["source_allowlist"]:
        before["files"][name] = sha(P / "candidate/before" / Path(name).name)
        assert sha(ROOT / name) == sha(P / "candidate/after" / Path(name).name)
    for leg, expected, opposite in (("before", before, after), ("after", after, before)):
        transition = read(P / f"source-transition-{leg}.json")
        assert transition["leg"] == leg and transition["head"] == plan["base"]
        assert transition["only_allowlisted_files_changed"] is True
        assert transition["after"] == {n: expected["files"][n] for n in plan["source_allowlist"]}
        assert transition["before"] == {n: opposite["files"][n] for n in plan["source_allowlist"]}
    quality = read(P / "quality.json")
    assert quality["status"] == "pass" and quality["gate_count"] == 6 and len(quality["rows"]) == 6
    quality_source = read(verify(quality["source"]))
    assert all(quality_source[n] == d for n, d in after["files"].items())
    counts = [0, 0, 0]
    spans = []
    for row in quality["rows"]:
        assert row["exit_code"] == 0
        log = verify(row["log"]).read_text()
        for match in re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", log):
            counts = [a + int(b) for a, b in zip(counts, match)]
        spans.append((row["started"], row["ended"]))
    assert counts == [641, 0, 1], counts
    for leg in ("before", "after"):
        reused = quality["pptx_reuse"][leg]
        receipt = read(verify(reused["receipt"]))
        assert receipt["status"] == "pass"
        verify(reused["source"])
        for name, digest in reused["archives"].items():
            assert sha(P / "candidate" / leg / Path(name).name) == digest
        for row in receipt["rows"]:
            assert row["exit_code"] == 0
            verify({"path": str(ROOT / row["log"]), "bytes": row["log_bytes"],
                    "sha256": row["log_sha256"]})
    build = read(P / "build-before/build.json")
    assert build["schema"] == "litchi.performance.0825.build-before.v1"
    assert read(verify(build["source"])) == before
    assert read(verify(build["frozen_inputs"])) == frozen
    assert build["quality"] == frozen["quality"]
    assert len(build["rows"]) == 3
    for logical, row in zip(("native", "artifacts", "observer"), build["rows"]):
        spec = plan["binaries"][logical]
        command = ["cargo", "build", "--offline", "--locked", "--release", "--manifest-path",
                   str(ROOT / "tools/perf-baseline/Cargo.toml"), "--bin", spec["cargo_bin"]]
        if spec["features"]:
            command += ["--features", ",".join(spec["features"])]
        assert row["command"] == command and row["exit_code"] == 0
        verify(row["log"])
        binary = build["binaries"][logical]["artifact"]
        if Path(binary["path"]).exists():
            verify(binary)
        else:
            assert binary in read(P / "cleanup.json")["binaries"]
        spans.append((row["started"], row["ended"]))
    exported = read(P / "artifacts-before-receipt.json")
    assert exported["exit_code"] == 0
    assert exported["binary"] == build["binaries"]["artifacts"]
    assert read(verify(exported["source"])) == before
    verify(exported["log"])
    verify(exported["manifest"])
    spans.append((exported["started"], exported["ended"]))
    admission = read(P / "artifact-admission-before.json")
    assert admission["accepted"] is True and admission["leg"] == "before"
    assert admission["plan_sha256"] == sha(P / "plan.json")
    for key in ("artifact_complete", "manifest", "audit", "auditor", "zip_preservation"):
        verify(admission[key])
    for value in admission["fresh_bound"]["attempt_files"].values():
        verify(value)
    replay = artifact_audit.Audit(ROOT).run(P / "artifacts-before")
    assert replay == read(verify(admission["audit"])) and replay["ok"] is True
    preserved = preservation.analyze("before")
    assert json.dumps(preserved, indent=2, sort_keys=True) + "\n" == verify(admission["zip_preservation"]).read_text()
    assert len(preserved["all_files"]) == 37 and len(preserved["cases"]) == 6
    rows = read(P / "qualification-before/receipts.json")
    assert len(rows) == 1
    row = rows[0]
    assert row["exit_code"] == 0 and row["case"] == "docx_real_file_ordinary_save_lifecycle"
    assert row["samples"] == 1 and row["warmup"] == 0 and row["leg"] == "before"
    assert row["binary"] == build["binaries"]["observer"]["artifact"]
    assert read(verify(row["source"])) == before
    assert read(verify(row["frozen_inputs"])) == frozen
    assert read(verify(row["artifact_admission"])) == admission
    verify(row["log"])
    assert int(verify(row["rss"]).read_text()) > 0
    report = read(verify(row["report"]))
    result = report["results"][0]
    assert len(report["results"]) == len(result["elapsed_ns"]["samples"]) == 1
    allocation = result["operation_metrics"]["allocation"]
    assert allocation["status"] == "measured"
    for field in custody.RAW_ALLOC:
        metric = allocation[field]
        assert metric["status"] == "measured" and metric["scope"] == "operation_global_system_allocator"
        assert len(metric["values"]) == 1 and type(metric["values"][0]) is int and metric["values"][0] >= 0
    try:
        custody.check_alloc(allocation, row["case"])
    except AssertionError as error:
        assert str(error) == row["case"] + ".allocation_calls"
    else:
        raise AssertionError("frozen checker no longer reproduces its documented defect")
    scalar = {name: allocation[name]["values"][0] for name in custody.RAW_ALLOC}
    assert scalar["failed_allocation_calls"] == 0
    assert scalar["live_bytes_after"] - scalar["live_bytes_before"] == scalar["allocated_bytes"] - scalar["deallocated_bytes"]
    assert scalar["region_peak_live_bytes"] >= max(scalar["live_bytes_before"], scalar["live_bytes_after"])
    spans.append((row["started"], row["ended"]))
    spans.sort()
    assert all(start <= end for start, end in spans)
    assert all(left[1] <= right[0] for left, right in zip(spans, spans[1:]))
    assert admission["ended"] <= row["started"]
    assert row["ended"] <= read(P / "source-transition-after.json")["started"]
    for name in ("qualification-before/complete.json", "qualification-admission-before.json",
                 "build-after", "artifacts-after", "qualification-after", "native", "observer",
                 "analysis.json", "analysis.md", "native.csv", "raw-audit.json"):
        assert not (P / name).exists(), f"unexpected comparative output: {name}"
    failures = read(P / "reader-failures.json")["failures"]
    assert len(failures) == 4
    for failure in failures:
        assert failure["exit_code"] == 1
        verify(failure["log"])
        for source in failure["sources"]:
            verify(source)
    assert "not JSON serializable" in (P / "admission-before/preservation-write.log").read_text()
    assert "allocation_calls" in (P / "qualification-before-run.log").read_text()
    if final:
        cleanup = read(P / "cleanup.json")
        assert cleanup["schema"] == "litchi.performance.0825.aborted-cleanup.v1"
        assert cleanup["verified"] is True
        assert cleanup["target_absent_after_removal"] is True and cleanup["scratch_absent_after_removal"] is True
        assert not Path(plan["target"]).exists() and not Path(plan["scratch"]).exists()
        assert cleanup["binaries"] == [build["binaries"][n]["artifact"] for n in ("native", "artifacts", "observer")]
        assert {r["path"] for r in cleanup["roots"]} == {plan["target"], plan["scratch"]}
        assert cleanup["scratch_marker_sha256"] == hashlib.sha256(plan["scratch_marker"].encode()).hexdigest()
    print("0825 aborted-attempt validation PASS: six quality gates; admitted baseline artifacts; one failed qualification; zero comparative captures; restored production" + ("; owned roots removed" if final else ""))


if __name__ == "__main__":
    assert sys.argv[1:] in ([], ["--final"])
    check(final=sys.argv[1:] == ["--final"])
