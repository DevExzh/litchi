"""Replay the retained failed candidate quality gate and bounded test repair."""
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P / "pre-repair-0"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path):
    return json.loads(path.read_text())


def check():
    manifest = read(OLD / "manifest.json")
    for name, digest in manifest["files"].items():
        assert sha(OLD / name) == digest, name
    quality_manifest = read(OLD / "quality-manifest.json")
    assert sha(OLD / "manifest.json") == quality_manifest["original_manifest_sha256"]
    for name, digest in quality_manifest["files"].items():
        assert sha(OLD / name) == digest == sha(P / name), name
    original = read(OLD / "freeze.json")
    current = read(P / "freeze.json")
    assert original.keys() == current.keys()
    for name in original:
        if name != "candidate":
            assert original[name] == current[name], name
    changed = {name for name in original["candidate"]
               if original["candidate"][name] != current["candidate"][name]}
    assert changed == {"candidate/after/transaction.rs", "candidate/combinedcandidate.patch"}
    for name, digest in original["candidate"].items():
        assert sha(OLD / name) == digest, name
    for name, digest in current["candidate"].items():
        assert sha(P / name) == digest, name
    old_text = (OLD / "candidate/after/transaction.rs").read_text()
    old_assertion = '            0,\n            "byte-identical compaction needs no scene read"'
    new_assertion = '            1,\n            "cold byte-identical compaction keeps its initial validation read"'
    assert old_text.count(old_assertion) == 1
    assert old_text.replace(old_assertion, new_assertion) == (P / "candidate/after/transaction.rs").read_text()
    receipt = read(P / "quality-after-0/receipt.json")
    assert receipt["schema"] == "litchi.performance.0824.quality.v1"
    assert receipt["leg"] == "after"
    assert receipt["environment"] == {
        "CARGO_TARGET_DIR": str(ROOT.parent / "litchi-target-0824/quality"),
        "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_DEV_DEBUG": "0", "RUSTDOCFLAGS": "-D warnings",
        "PYTHONDONTWRITEBYTECODE": "1",
    }
    assert [row["command"] for row in receipt["rows"]] == [
        ["cargo", "fmt", "-p", "litchi-pptx", "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "-p", "litchi-pptx", "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", "-p", "litchi-pptx", "--all-features", "--", "--test-threads=2"],
    ]
    assert receipt["status"] == "failed" and len(receipt["rows"]) == 3
    assert [row["exit_code"] for row in receipt["rows"]] == [0, 0, 101]
    assert receipt["driver_sha256"] == sha(OLD / "quality.py") == sha(P / "quality.py")
    source = read(ROOT / receipt["source"])
    for name, digest in original["source"]["files"].items():
        expected = original["candidate"].get("candidate/after/" + Path(name).name, digest) if name in (
            "crates/litchi-pptx/src/opened/transaction.rs",
            "crates/litchi-pptx/src/opened/xml.rs",
        ) else digest
        assert source[name] == expected, name
    for row in receipt["rows"]:
        log = ROOT / row["log"]
        assert sha(log) == row["log_sha256"] and log.stat().st_size == row["log_bytes"]
    log = (ROOT / receipt["rows"][-1]["log"]).read_text()
    assert 'left: 1' in log and 'right: 0' in log
    assert 'test result: FAILED. 837 passed; 1 failed; 1 ignored' in log
    cleanup = read(OLD / "cleanup.json")
    build = read(OLD / "build-before/build.json")
    assert cleanup["binaries"] == [x["artifact"] for x in build["binaries"].values()]
    assert cleanup["quality_target_retained"] is True
    assert cleanup["ended"] <= read(P / "build-before/commands.json")[0]["started"]
    assert read(OLD / "qualification-before/complete.json")["reports"] == 19
    assert receipt["rows"][-1]["ended"] < cleanup["started"]
    print("0824 failure audit PASS: original failure retained; only cold-path test assertion repaired; fresh baseline restarted")


if __name__ == "__main__":
    check()
