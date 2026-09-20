#!/usr/bin/env python3
"""Check source restoration, evidence-gate receipts and packet integrity."""
import hashlib
import json
import runpy
import subprocess
import tempfile
from pathlib import Path
P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    baseline = json.loads((P / "baseline.json").read_text())
    for group in ("source_sha256", "constraints_sha256", "build_inputs_sha256"):
        for name, digest in baseline[group].items():
            assert sha(ROOT / name) == digest, (group, name)
    checks = json.loads((P / "validation.json").read_text())
    assert {row["name"] for row in checks} == {
        "crate-boundaries", "claims", "claims-structural", "report", "coverage", "non-iwork"}
    for row in checks:
        assert row["exit_code"] == 0, row
        assert sha(P / "validation" / (row["name"] + ".log")) == row["log_sha256"]
    build = json.loads((P / "build.json").read_text())
    assert build["status"] == "completed"
    assert build["original_codec_sha256"] == build["restored_codec_sha256"]
    assert sha(P / "0703-mce-trace.patch") == build["patch_sha256"]
    for name, digest in build["probe_sha256"].items():
        assert sha(P / name) == digest, name
    for step in build["steps"]:
        assert step["exit_code"] == 0
        assert sha(P / (step["name"] + ".log")) == step["log_sha256"]
    with tempfile.TemporaryDirectory(prefix="litchi-0703-audit-") as directory:
        base = Path(directory)
        relative = Path("crates/litchi-ooxml-common/src/mce/codec.rs")
        target = base / relative
        target.parent.mkdir(parents=True)
        target.write_bytes((ROOT / relative).read_bytes())
        subprocess.run(["git", "apply", str(P / "0703-mce-trace.patch")],cwd=base,check=True)
        assert sha(target) == build["instrumented_codec_sha256"]
    recomputed = runpy.run_path(str(P / "focused-summary.py"))["summarize"]()
    assert recomputed == json.loads((P / "focused-summary.json").read_text())
    corpus = runpy.run_path(str(P / "prepare-corpus.py"))["prepare"]()
    assert corpus == json.loads((P / "corpus.json").read_text())
    quality = json.loads((P / "probe-quality.json").read_text())
    assert quality["exit_code"] == 0
    assert sha(P / "probe-clippy.log") == quality["log_sha256"]
    runs = json.loads((P / "runs.json").read_text())
    for row in runs:
        assert row["binary_sha256"] == build["binary_sha256"]
        if row["case"] == "real":
            assert row["source_archive_sha256"] == corpus["real"]["archive_sha256"]
    cleanup_path = P / "cleanup.json"
    if cleanup_path.exists():
        cleanup = json.loads(cleanup_path.read_text())
        assert cleanup["workspace_lock_preserved"]
        assert set(cleanup["removed_paths"]) == {str(ROOT.parent / "litchi-target-0703"), str(ROOT.parent / "litchi-0703-bin")}
        for path in cleanup["removed_paths"]:
            assert not Path(path).exists()
    assert not list(P.rglob("*.phase"))
    scripts = sorted(P.rglob("*.py"))
    for script in scripts:
        compile(script.read_bytes(), str(script), "exec")
    manifest = json.loads((P / "artifact-hashes.json").read_text())
    census = runpy.run_path(str(P / "seal.py"))["files"]()
    expected = {str(path.relative_to(P)) for path in census}
    expected.add("../../0703-pptx-capture-projection-reuse-diagnostic.md")
    assert set(manifest) == expected
    for name, digest in manifest.items():
        assert sha(P / name) == digest, name
    assert not list(P.rglob("__pycache__"))
    print(f"PASS: restored source, {len(checks)} evidence gates, {len(scripts)} scripts, {len(manifest)} artifacts")

if __name__ == "__main__":
    main()
