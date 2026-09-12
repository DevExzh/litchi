"""Freeze and test one source-reviewed candidate attempt without replacing evidence."""
import argparse
import subprocess
from run import HERE, REPO, TARGET, manifest_paths, run, sha, write

parser = argparse.ArgumentParser()
parser.add_argument("stage")
args = parser.parse_args()
assert args.stage.startswith("preflight-") and args.stage[10:].isdigit()
folder = HERE / args.stage
folder.mkdir(exist_ok=False)
write(folder / "source-manifest.json", {
    name: sha(REPO / name) for name in sorted(manifest_paths())
})
(folder / "source.patch").write_bytes(subprocess.check_output(
    ["git", "diff", "--", "crates", "tools/perf-baseline"], cwd=REPO))
run(args.stage, "check-xlsx-tests", [
    "env", "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
    "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings",
    "cargo", "test", "--locked", "-p", "litchi-xlsx", "--all-features",
    "--", "--test-threads=2",
])
