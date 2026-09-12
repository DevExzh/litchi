"""Run the remaining quality matrix, reusing exact-source preflight checks."""
import json
from checks import COMMANDS
from run import HERE, TARGET, run, sha

candidate = HERE / "candidate" / "source-manifest.json"
preflight = HERE / "preflight-3" / "source-manifest.json"
assert candidate.read_bytes() == preflight.read_bytes()
for name, command in COMMANDS:
    if name in ("xlsx-tests", "clippy"):
        receipt = json.loads((preflight.parent / ("check-" + name + ".receipt.json")).read_text())
        assert receipt["exit_code"] == 0
        assert receipt["source_manifest_sha256"] == sha(candidate)
        continue
    run("candidate", "check-" + name, [
        "env", "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings",
    ] + command)
