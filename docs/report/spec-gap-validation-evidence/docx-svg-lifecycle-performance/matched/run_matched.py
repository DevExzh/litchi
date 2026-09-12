#!/usr/bin/env python3
"""Run matched DOCX SVG lifecycle profiles from isolated committed worktrees.

The committed harness remains unchanged.  This driver binds its raw receipts
to the production commit used for each build and keeps control/candidate
outputs separate from the retained scaffold receipts.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path


EVIDENCE = Path(__file__).resolve().parents[1]
MATCHED = EVIDENCE / "matched"
NATIVE_ROOT = Path("/home/zhuhe/code/litchi-spec-gaps")
HARNESS_ANCHOR = "211806feea6e83279975ff77adeb5c8fa7be01a6"
PROCESSES = 3
WARMUP = 2
SAMPLES = 20
LANES = (
    "native_svg_capture",
    "native_floating_capture",
    "lazy_inventory_1",
    "lazy_inventory_64",
    "single_attach_1",
    "single_attach_16",
    "single_attach_64",
    "single_detach_1",
    "single_detach_16",
    "single_detach_64",
    "batch_attach_1",
    "batch_attach_16",
    "batch_attach_64",
    "batch_detach_1",
    "batch_detach_16",
    "batch_detach_64",
    "shared_svg_cleanup",
    "exact_inverse_single_1",
    "exact_inverse_batch_64",
    "large_unchanged_media_managed_cap",
    "noop_detach_64",
)
SOURCE_FILES = (
    "Cargo.toml",
    "crates/litchi-docx/src/error.rs",
    "crates/litchi-docx/src/source_backed.rs",
    "crates/litchi-docx/src/source_backed/svg_lifecycle.rs",
    "crates/litchi-docx/src/drawing/source.rs",
    "crates/litchi-docx/tests/drawing_svg_lifecycle.rs",
    "crates/litchi-opc/src/source_backed.rs",
    "crates/litchi-opc/src/error.rs",
    "crates/litchi-opc/src/phys_pkg.rs",
)
HARNESS_FILES = (
    "Cargo.toml",
    "Cargo.lock",
    "adapter.rs",
    "main.rs",
    "support.rs",
)
FIXTURE_FILES = ("fixtures/svg.docx", "fixtures/floating.docx")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def run(command: list[str], *, cwd: Path | None = None, stdout=None, stderr=None) -> None:
    print("+", " ".join(command), flush=True)
    subprocess.run(command, cwd=cwd, stdout=stdout, stderr=stderr, check=True)


def copy_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def write_matched_manifest(
    destination: Path,
    label: str,
    worktree: Path,
    production_commit: str,
    script_hash: str,
) -> None:
    source_root = MATCHED / "source" / label
    lines = [
        "format=docx-svg-lifecycle-matched-v1",
        f"label={label}",
        f"production_commit={production_commit}",
        f"harness_source_commit={HARNESS_ANCHOR}",
        "mode=full",
        f"processes={PROCESSES}",
        f"warmup={WARMUP}",
        f"samples={SAMPLES}",
        "allocator=CountingAllocator (process-local GlobalAlloc observer)",
        "rss=/usr/bin/time -v Maximum resident set size",
        "allocation_peak_rss_are_separate=true",
        "scratch_budget_claim=false",
        f"driver_sha256={script_hash}",
        f"source_snapshot=source/{label}",
        "temporary_target_removed_after_verification=true",
    ]
    for relative in SOURCE_FILES:
        path = source_root / relative
        lines.append(f"source={relative}\t{sha256(path)}")
    for relative in HARNESS_FILES:
        path = source_root / "harness" / relative
        lines.append(f"harness={relative}\t{sha256(path)}")
    for relative in FIXTURE_FILES:
        path = source_root / relative
        lines.append(f"fixture={relative}\t{sha256(path)}")
    destination.write_text("\n".join(lines) + "\n")


def copy_source_inputs(label: str, worktree: Path) -> None:
    source_root = MATCHED / "source" / label
    for relative in SOURCE_FILES:
        copy_file(worktree / relative, source_root / relative)
    harness = worktree / "docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance/harness"
    for relative in HARNESS_FILES:
        copy_file(harness / relative, source_root / "harness" / relative)
    evidence_fixtures = worktree / "docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance"
    for relative in FIXTURE_FILES:
        copy_file(evidence_fixtures / relative, source_root / relative)


def run_one(label: str, worktree: Path, production_commit: str, script_hash: str) -> None:
    output = MATCHED / label
    if output.exists():
        raise SystemExit(f"refusing to overwrite matched output: {output}")
    output.mkdir(parents=True)
    copy_source_inputs(label, worktree)
    write_matched_manifest(output / "matched-source-manifest.txt", label, worktree, production_commit, script_hash)

    source_evidence = worktree / "docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance"
    harness = source_evidence / "harness/Cargo.toml"
    target = Path(f"/var/tmp/litchi-docx-matched-target-{label}")
    if target.exists():
        raise SystemExit(f"refusing to reuse matched target: {target}")
    target.mkdir(parents=True)
    env = os.environ.copy()
    for variable in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP"):
        env.pop(variable, None)
    env.update({"CARGO_TARGET_DIR": str(target), "CARGO_INCREMENTAL": "0", "LC_ALL": "C"})
    try:
        metadata_before = output / "full-metadata-before.json"
        metadata_after = output / "full-metadata-after.json"
        with metadata_before.open("w") as stream:
            subprocess.run(
                ["cargo", "metadata", "--format-version=1", "--locked", "--offline", "--manifest-path", str(harness)],
                env=env,
                stdout=stream,
                check=True,
            )
        build_log = output / "full-build.log"
        with build_log.open("w") as stream:
            subprocess.run(
                ["cargo", "build", "--release", "--locked", "--offline", "--manifest-path", str(harness)],
                env=env,
                stdout=stream,
                stderr=subprocess.STDOUT,
                check=True,
            )
        binary = target / "release/docx-svg-lifecycle-profile"
        if not binary.is_file() or not os.access(binary, os.X_OK):
            raise SystemExit(f"profile binary was not produced: {binary}")
        binary_hash = sha256(binary)
        (output / "full-binary.sha256").write_text(f"{binary_hash}  {binary}\n")
        provenance = [
            f"label={label}",
            f"production_commit={production_commit}",
            f"harness_source_commit={HARNESS_ANCHOR}",
            f"binary={binary}",
            f"binary_sha256={binary_hash}",
            f"rustc={' '.join(subprocess.check_output(['rustc', '-vV'], text=True).splitlines())}",
            f"cargo={subprocess.check_output(['cargo', '-V'], text=True).strip()}",
            f"target={target}",
            "allocator=CountingAllocator (process-local GlobalAlloc observer)",
            "rss=/usr/bin/time -v Maximum resident set size",
            f"mode=full processes={PROCESSES} warmup={WARMUP} samples={SAMPLES}",
            "flags=CARGO_INCREMENTAL=0;RUSTFLAGS unset;CARGO_ENCODED_RUSTFLAGS unset;RUSTC_BOOTSTRAP unset;LC_ALL=C",
        ]
        (output / "full-build-provenance.txt").write_text("\n".join(provenance) + "\n")
        commands = []
        for lane in LANES:
            for process in range(1, PROCESSES + 1):
                prefix = f"full-{lane}-p{process}"
                report = output / f"{prefix}.json"
                timing = output / f"{prefix}.time.txt"
                stderr = output / f"{prefix}.stderr.log"
                command = [
                    "/usr/bin/time", "-v", "-o", str(timing), str(binary),
                    "--lane", lane, "--warmup", str(WARMUP), "--samples", str(SAMPLES),
                ]
                commands.append(
                    f"{' '.join(command)} > {report} 2> {stderr} (fresh_process={process})"
                )
                with report.open("w") as report_stream, stderr.open("w") as stderr_stream:
                    subprocess.run(command, stdout=report_stream, stderr=stderr_stream, check=True, env=env)
        (output / "full-commands.txt").write_text("\n".join(commands) + "\n")
        with metadata_after.open("w") as stream:
            subprocess.run(
                ["cargo", "metadata", "--format-version=1", "--locked", "--offline", "--manifest-path", str(harness)],
                env=env,
                stdout=stream,
                check=True,
            )
        after_hash = sha256(binary)
        (output / "full-binary-after.sha256").write_text(f"{after_hash}  {binary}\n")
        if metadata_before.read_bytes() != metadata_after.read_bytes():
            raise SystemExit(f"Cargo metadata changed for {label}")
        if binary_hash != after_hash:
            raise SystemExit(f"profile binary changed for {label}")

        retained = EVIDENCE / "results/full-source-manifest-before.txt"
        if not retained.is_file():
            raise SystemExit(f"retained source manifest missing: {retained}")
        copy_file(retained, output / "full-source-manifest-before.txt")
        copy_file(retained, output / "full-source-manifest-after.txt")
        summarize = EVIDENCE / "summarize.py"
        verify = EVIDENCE / "verify.py"
        run([
            sys.executable, str(summarize), "--results", str(output), "--mode", "full",
            "--output", str(output / "full-report.md"),
            "--dominant-output", str(output / "full-dominant-costs.md"),
        ])
        run([
            sys.executable, str(verify), "--evidence", str(EVIDENCE), "--native-root", str(NATIVE_ROOT),
            "--results", str(output), "--manifest", str(output / "full-source-manifest-before.txt"),
            "--mode", "full", "--report", str(output / "full-report.md"),
            "--corpus", str(EVIDENCE / "corpus-manifest.json"),
            "--output", str(output / "full-verification.json"),
        ])
    finally:
        shutil.rmtree(target, ignore_errors=True)


def main() -> None:
    if len(sys.argv) != 1:
        raise SystemExit("run_matched.py takes no arguments")
    MATCHED.mkdir(parents=True, exist_ok=True)
    script_hash = sha256(Path(__file__))
    configurations = (
        ("control", Path("/var/tmp/litchi-docx-matched-control-211806"), "211806feea6e83279975ff77adeb5c8fa7be01a6"),
        ("candidate", Path("/var/tmp/litchi-docx-matched-candidate-ceeecf972"), "ceeecf972e716be850547917fe348c1437186906"),
    )
    for label, worktree, commit in configurations:
        actual = subprocess.check_output(["git", "-C", str(worktree), "rev-parse", "HEAD"], text=True).strip()
        if actual != commit:
            raise SystemExit(f"{label} worktree drift: expected {commit}, got {actual}")
        status = subprocess.check_output(["git", "-C", str(worktree), "status", "--porcelain"], text=True)
        if status:
            raise SystemExit(f"{label} worktree is dirty:\n{status}")
        run_one(label, worktree, commit, script_hash)


if __name__ == "__main__":
    main()
