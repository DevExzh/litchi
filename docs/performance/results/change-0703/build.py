#!/usr/bin/env python3
"""Build temporary default-MCE instrumentation and restore exact production bytes."""
import hashlib
import json
import os
import shutil
import subprocess
import time
from pathlib import Path
P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / "litchi-target-0703"
BIN = ROOT.parent / "litchi-0703-bin"
CODEC = ROOT / "crates/litchi-ooxml-common/src/mce/codec.rs"

def sha(data):
    return hashlib.sha256(data).hexdigest()

def main():
    original = CODEC.read_bytes()
    baseline = json.loads((P / "baseline.json").read_text())
    assert sha(original) == baseline["source_sha256"][str(CODEC.relative_to(ROOT))]
    workspace_lock = (ROOT / "Cargo.lock").read_bytes()
    patch = P / "0703-mce-trace.patch"
    manifest = P / "probe/Cargo.toml"
    env = os.environ | {"RUSTFLAGS": "-D warnings"}
    steps = []
    def run(name, command):
        log = P / (name + ".log")
        start = time.monotonic()
        with log.open("w") as out:
            result = subprocess.run(command, cwd=ROOT, env=env, stdout=out, stderr=subprocess.STDOUT)
        steps.append(dict(name=name, command=command, exit_code=result.returncode,
                          seconds=time.monotonic()-start, log_sha256=sha(log.read_bytes())))
        if result.returncode:
            raise RuntimeError(log.read_text())
    if not (P / "probe/Cargo.lock").exists():
        run("probe-lock", ["cargo", "generate-lockfile", "--offline", "--manifest-path", str(manifest)])
    assert (ROOT / "Cargo.lock").read_bytes() == workspace_lock
    run("probe-fmt", ["cargo", "fmt", "--manifest-path", str(manifest), "--", "--check"])
    run("patch-check", ["git", "apply", "--check", str(patch)])
    receipt = dict(diagnostic_only=True, performance_claim="none", baseline_head=baseline["baseline_head"],
                   original_codec_sha256=sha(original), patch_sha256=sha(patch.read_bytes()),
                   environment={"RUSTFLAGS": env["RUSTFLAGS"]}, steps=steps)
    try:
        run("patch-apply", ["git", "apply", str(patch)])
        receipt["instrumented_codec_sha256"] = sha(CODEC.read_bytes())
        receipt["probe_sha256"] = {str(path.relative_to(P)): sha(path.read_bytes())
                                  for path in sorted((P / "probe").rglob("*")) if path.is_file()}
        run("build-trace", ["cargo", "build", "--release", "--locked", "--offline",
                            "--manifest-path", str(manifest), "--target-dir", str(TARGET), "-j", "2"])
        BIN.mkdir(exist_ok=True)
        binary = BIN / "probe0703"
        shutil.copy2(TARGET / "release/probe0703", binary)
        receipt["binary"] = str(binary)
        receipt["binary_sha256"] = sha(binary.read_bytes())
        receipt["status"] = "completed"
    finally:
        CODEC.write_bytes(original)
        receipt["restored_codec_sha256"] = sha(CODEC.read_bytes())
        receipt["workspace_lock_sha256"] = sha((ROOT / "Cargo.lock").read_bytes())
        assert (ROOT / "Cargo.lock").read_bytes() == workspace_lock
        (P / "build.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print("PASS: trace binary built; exact production codec and workspace lock restored", flush=True)

if __name__ == "__main__":
    main()
