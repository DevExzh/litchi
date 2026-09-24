from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import time

root, perf, results, target = (Path(value).resolve() for value in sys.argv[1:])
results.mkdir(parents=True, exist_ok=True)
head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=root, text=True).strip()
expected_head = "3e9ef2d014a2f70bbbca18b824aa50a1567cdcc7"
semantic_owner = "cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd"
production_baseline = "2a2ffa1cae4e6b7070082768ce84483e5d411dc8"
design_path = "docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md"
design_sha = "30b78cca84c4ca24ae44f3d3694c3097f54b5e5a1f2004af9ce5007bcaf4173d"
design_blob = "597400950b1027c47cd6e4cbbedd23915bc0980e"
helper_path = "crates/litchi-pptx/tests/pptx_ink_actions.rs"
helper_sha = "bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e"
helper_blob = "ad7e43e8c2c362c1b9e1938806f59b9fb3ab1dea"
if head != expected_head:
    raise SystemExit(f"unexpected capture HEAD: {head}")

for variable in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP", "RUSTDOCFLAGS"):
    if os.environ.get(variable):
        raise SystemExit(f"refusing inherited {variable}")
env = os.environ.copy()
env.update({
    "CARGO_TARGET_DIR": str(target),
    "CARGO_INCREMENTAL": "0",
    "LC_ALL": "C",
    "PPTX_INK_ACTIONS_CAPTURE_HEAD": head,
})
for variable in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP", "RUSTDOCFLAGS"):
    env.pop(variable, None)

commands_path = results / "commands.jsonl"
failures: list[str] = []

def run(label: str, argv: list[str], stdout_name: str | None = None, stderr_name: str | None = None) -> int:
    stdout = results / (stdout_name or f"{label}.stdout")
    stderr = results / (stderr_name or f"{label}.stderr")
    start = time.time_ns()
    with stdout.open("wb") as out, stderr.open("wb") as err:
        completed = subprocess.run(argv, cwd=root, env=env, stdout=out, stderr=err)
    finish = time.time_ns()
    record = {
        "event": "command",
        "label": label,
        "argv": argv,
        "cwd": str(root),
        "exit_code": completed.returncode,
        "stdout": str(stdout),
        "stderr": str(stderr),
        "started_unix_ns": start,
        "finished_unix_ns": finish,
        "environment": {
            key: env.get(key)
            for key in (
                "CARGO_TARGET_DIR",
                "CARGO_INCREMENTAL",
                "LC_ALL",
                "PPTX_INK_ACTIONS_CAPTURE_HEAD",
                "RUSTFLAGS",
                "CARGO_ENCODED_RUSTFLAGS",
                "RUSTC_BOOTSTRAP",
                "RUSTDOCFLAGS",
            )
        },
    }
    with commands_path.open("a") as stream:
        stream.write(json.dumps(record, sort_keys=True) + "\n")
    if completed.returncode:
        failures.append(f"{label}: exit {completed.returncode}")
    return completed.returncode

def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

def git_status(name: str) -> Path:
    path = results / name
    path.write_text(subprocess.check_output(
        ["git", "status", "--porcelain", "--untracked-files=all"], cwd=root, text=True
    ))
    return path

def host_receipt(name: str) -> None:
    path = results / name
    model = ""
    try:
        for line in Path("/proc/cpuinfo").read_text().splitlines():
            if line.startswith("model name") or line.startswith("Hardware"):
                model = line.split(":", 1)[1].strip()
                break
    except OSError:
        pass
    memory = ""
    try:
        for line in Path("/proc/meminfo").read_text().splitlines():
            if line.startswith("MemTotal:"):
                memory = " ".join(line.split()[1:3])
                break
    except OSError:
        pass
    path.write_text("\n".join([
        f"capture_head={head}",
        f"semantic_owner_commit={semantic_owner}",
        f"production_source_baseline_commit={production_baseline}",
        f"uname={platform.uname()}",
        f"hostname={platform.node()}",
        f"loadavg={' '.join(str(value) for value in os.getloadavg())}",
        f"cpu_model={model}",
        f"memory={memory}",
        f"cwd={root}",
        f"cargo_target_dir={target}",
        f"RUSTFLAGS={env.get('RUSTFLAGS')}",
        f"CARGO_ENCODED_RUSTFLAGS={env.get('CARGO_ENCODED_RUSTFLAGS')}",
        f"RUSTC_BOOTSTRAP={env.get('RUSTC_BOOTSTRAP')}",
        f"RUSTDOCFLAGS={env.get('RUSTDOCFLAGS')}",
        "",
    ]))

def manifest_args(metadata: Path, output: Path) -> list[str]:
    context_extras = [
        root / "docs/adr/0001-priorities-and-api-layers.md",
        root / "docs/adr/0003-snapshots-edits-and-patches.md",
        root / "docs/adr/0005-io-memory-and-performance.md",
        root / "docs/adr/0006-validation-security-and-compatibility.md",
    ]
    extras = [
        root / "Cargo.toml", root / "rust-toolchain.toml", root / ".cargo/config.toml",
        root / "rustfmt.toml", root / "clippy.toml", root / "deny.toml",
        perf / "harness/Cargo.toml", perf / "harness/Cargo.lock", perf / "harness/main.rs",
        perf / "harness/adapter.rs", perf / "harness/support.rs", perf / "source_manifest.py",
        perf / "committed_inputs.py", perf / "verify.py", perf / "cleanup_target.sh",
        perf / "run_profile.sh", perf / "README.md", perf / "PLAN.md", perf / "report.md",
        perf / "requirements.md", perf / "corpus-manifest.json", perf / "source-contract.json",
        perf / "scaffold_tests.py", perf / "root-review.md",
    ]
    args = [
        "python3", str(perf / "source_manifest.py"), "--metadata", str(metadata),
        "--root", str(root), "--output", str(output), "--git-commit", head,
        "--semantic-owner-commit", semantic_owner,
        "--production-source-baseline-commit", production_baseline,
        "--production-exclude-package", "litchi-pptx-ink-actions-performance",
        "--semantic-owner-extra", str(root / design_path),
        "--semantic-owner-extra-sha256", design_sha,
        "--semantic-owner-extra-blob", design_blob,
    ]
    for extra in extras:
        args.extend(("--extra", str(extra)))
    for extra in context_extras:
        args.extend(("--context-extra", str(extra)))
    for extra in (root / "Cargo.toml", root / "rust-toolchain.toml", root / ".cargo/config.toml",
                  root / "rustfmt.toml", root / "clippy.toml", root / "deny.toml"):
        args.extend(("--production-extra", str(extra)))
    return args

def guard_args() -> list[str]:
    paths = [
        "crates/litchi-pptx/src/lib.rs",
        "crates/litchi-pptx/src/package/codec.rs",
        "crates/litchi-pptx/src/package/model.rs",
        "crates/litchi-pptx/src/presentation/model.rs",
        "crates/litchi-pptx/src/presentation/package.rs",
        "crates/litchi-pptx/src/presentation/embedded/ink_actions/model.rs",
        "crates/litchi-pptx/src/presentation/embedded/ink_actions/codec.rs",
        "crates/litchi-pptx/src/presentation/embedded/ink_actions/package.rs",
        "crates/litchi-pptx/src/presentation/embedded/ink_actions/transaction.rs",
        helper_path,
    ]
    args = [
        "python3", str(perf / "committed_inputs.py"), "--root", str(root),
        "--semantic-owner-commit", semantic_owner,
        "--production-source-baseline-commit", production_baseline,
        "--helper", helper_path, "--helper-sha256", helper_sha,
        "--manifest-path", str(perf / "harness/Cargo.toml"),
        "--exclude-package", "litchi-pptx-ink-actions-performance",
        "--semantic-owner-extra", str(root / design_path),
        "--semantic-owner-extra-sha256", design_sha,
        "--semantic-owner-extra-blob", design_blob,
    ]
    for path in paths:
        args.extend(("--path", path))
    for extra in (root / "Cargo.toml", root / "rust-toolchain.toml", root / ".cargo/config.toml",
                  root / "rustfmt.toml", root / "clippy.toml", root / "deny.toml"):
        args.extend(("--production-extra", str(extra)))
    return args

git_status("git-status-before.txt")
if (results / "git-status-before.txt").read_text():
    failures.append("checkout dirty before capture")
if not (perf / "harness/Cargo.lock").is_file():
    failures.append("isolated harness lock missing")
(results / "cargo-lock.sha256").write_text(f"{sha256(perf / 'harness/Cargo.lock')}  {perf / 'harness/Cargo.lock'}\n")
(results / "corpus-manifest.sha256").write_text(f"{sha256(perf / 'corpus-manifest.json')}  {perf / 'corpus-manifest.json'}\n")
(results / "helper.sha256").write_text(f"{sha256(root / helper_path)}  {root / helper_path}\n")
(results / "design.sha256").write_text(f"{sha256(root / design_path)}  {root / design_path}\n")
host_receipt("host-before.txt")
run("cargo-metadata-before", ["cargo", "metadata", "--format-version=1", "--locked", "--offline", "--manifest-path", str(perf / "harness/Cargo.toml")], "cargo-metadata-before.json", "cargo-metadata-before.stderr")
run("rustc-version-before", ["rustc", "-vV"], "rustc-vV-before.txt", "rustc-vV-before.stderr")
run("cargo-version-before", ["cargo", "-V"], "cargo-version-before.txt", "cargo-version-before.stderr")
run("rustfmt-version-before", ["rustfmt", "--version"], "rustfmt-version-before.txt", "rustfmt-version-before.stderr")
run("committed-inputs-guard", guard_args(), "committed-inputs-guard.stdout", "committed-inputs-guard.stderr")
run("source-manifest-before", manifest_args(results / "cargo-metadata-before.json", results / "source-manifest-before.txt"), "source-manifest-before.stdout", "source-manifest-before.stderr")

binary = target / "release/pptx-ink-actions-performance"
build_code = run("cargo-build-release", ["cargo", "build", "--release", "--locked", "--offline", "--manifest-path", str(perf / "harness/Cargo.toml")], "cargo-build-release.stdout", "cargo-build-release.stderr")
if build_code == 0 and binary.is_file():
    binary_hash = sha256(binary)
    (results / "binary.sha256").write_text(f"{binary_hash}  {binary}\n")
    (results / "binary-stat.txt").write_text(f"binary_size_bytes={binary.stat().st_size}\nbinary_mode={oct(binary.stat().st_mode & 0o777)[2:]}\n")
    file_info = subprocess.run(["file", "-b", str(binary)], cwd=root, env=env, capture_output=True, text=True)
    (results / "binary-file.txt").write_text(file_info.stdout)
else:
    failures.append("release binary missing after successful build")

if binary.is_file():
    run("host-probe", [str(binary), "--host-probe"], "host-probe.stdout", "host-probe.stderr")
    run("matrix-correctness", [str(binary), "--matrix-correctness"], "matrix-correctness.json", "matrix-correctness.stderr")
    lanes = [
        "package_read_tiny_shared", "presentation_read_small_shared", "package_read_medium_shared",
        "package_read_small_distinct", "package_read_large_shared", "package_read_large_distinct",
        "package_read_near_shared", "package_read_near_distinct", "package_read_multislide_shared",
        "package_read_multislide_distinct", "package_scalar_edit_small_shared",
        "presentation_scalar_edit_small_shared", "package_noop_small_shared", "presentation_noop_small_shared",
        "package_apply_medium_shared", "package_apply_medium_distinct", "package_inverse_small_shared",
        "package_inverse_case_equivalent", "package_save_medium_shared", "package_save_reopen_medium_shared",
        "stale_owner", "stale_owner_rels", "stale_target", "stale_content_type", "signed_noop",
        "signed_changed_refusal", "opaque_mce_scalar_edit", "opaque_mce_save_reopen",
        "unknown_outbound_read_edit", "strict_shared_edit", "limit_anchor_one_under", "limit_anchor_exact",
        "limit_anchor_one_over", "limit_target_one_under", "limit_target_exact", "limit_target_one_over",
        "limit_aggregate_one_under", "limit_aggregate_exact", "limit_aggregate_one_over", "limit_graph_one_under",
        "limit_graph_exact", "limit_graph_one_over",
    ]
    for lane in lanes:
        run(f"lane-{lane}", [str(binary), "--lane", lane, "--warmup", "1", "--samples", "1"], f"lane-{lane}.json", f"lane-{lane}.stderr")
    (results / "lane-list.txt").write_text("\n".join(lanes) + "\n")
else:
    failures.append("skipped host/matrix/lanes because binary is absent")

run("cargo-metadata-after", ["cargo", "metadata", "--format-version=1", "--locked", "--offline", "--manifest-path", str(perf / "harness/Cargo.toml")], "cargo-metadata-after.json", "cargo-metadata-after.stderr")
run("source-manifest-after", manifest_args(results / "cargo-metadata-after.json", results / "source-manifest-after.txt"), "source-manifest-after.stdout", "source-manifest-after.stderr")
host_receipt("host-after.txt")
git_status("git-status-after.txt")

if binary.is_file():
    (results / "binary-after.sha256").write_text(f"{sha256(binary)}  {binary}\n")
    if sha256(binary) != sha256(Path(results / "binary.sha256")):
        failures.append("binary hash changed")
if (results / "source-manifest-before.txt").is_file() and (results / "source-manifest-after.txt").is_file():
    if (results / "source-manifest-before.txt").read_bytes() != (results / "source-manifest-after.txt").read_bytes():
        failures.append("source manifest changed")
if (results / "cargo-metadata-before.json").is_file() and (results / "cargo-metadata-after.json").is_file():
    if (results / "cargo-metadata-before.json").read_bytes() != (results / "cargo-metadata-after.json").read_bytes():
        failures.append("cargo metadata changed")
if (results / "git-status-after.txt").read_text():
    failures.append("checkout dirty after capture")

provenance = {
    "schema": "pptx-ink-actions-correctness-capture-provenance-v1",
    "capture_head": head,
    "semantic_owner_commit": semantic_owner,
    "production_source_baseline_commit": production_baseline,
    "semantic_owner_design": {"path": design_path, "sha256": design_sha, "git_blob": design_blob},
    "helper": {"path": helper_path, "sha256": helper_sha, "git_blob": helper_blob},
    "target": str(target),
    "binary": str(binary),
    "binary_sha256": sha256(binary) if binary.is_file() else None,
    "cargo_lock_sha256": sha256(perf / "harness/Cargo.lock"),
    "metadata_before_sha256": sha256(results / "cargo-metadata-before.json") if (results / "cargo-metadata-before.json").is_file() else None,
    "metadata_after_sha256": sha256(results / "cargo-metadata-after.json") if (results / "cargo-metadata-after.json").is_file() else None,
    "source_manifest_before_sha256": sha256(results / "source-manifest-before.txt") if (results / "source-manifest-before.txt").is_file() else None,
    "source_manifest_after_sha256": sha256(results / "source-manifest-after.txt") if (results / "source-manifest-after.txt").is_file() else None,
    "warmup": 1,
    "samples": 1,
    "lanes_requested": 42,
    "usr_bin_time": False,
    "timed_performance_claim": False,
    "failures": failures,
}
(results / "capture-provenance.json").write_text(json.dumps(provenance, indent=2, sort_keys=True) + "\n")
print(json.dumps({"results": str(results), "target": str(target), "failures": failures}, sort_keys=True))
if failures:
    raise SystemExit(1)
