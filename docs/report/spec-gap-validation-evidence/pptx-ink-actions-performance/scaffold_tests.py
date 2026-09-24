#!/usr/bin/env python3
"""Fast, timing-free executable checks for the bounded profile scaffold."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


HERE = Path(__file__).resolve().parent
ROOT = (HERE / "../../../..").resolve()
RUNNER = HERE / "run_profile.sh"
CLEANUP = HERE / "cleanup_target.sh"
VERIFY = HERE / "verify.py"
COMMITTED_INPUTS = HERE / "committed_inputs.py"
SOURCE_MANIFEST = HERE / "source_manifest.py"
SEMANTIC_OWNER_DESIGN_PATH = "docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md"
SEMANTIC_OWNER_DESIGN_SHA256 = "30b78cca84c4ca24ae44f3d3694c3097f54b5e5a1f2004af9ce5007bcaf4173d"
SEMANTIC_OWNER_DESIGN_GIT_BLOB = "597400950b1027c47cd6e4cbbedd23915bc0980e"

CHECKS = 0


def check(condition: bool, message: str) -> None:
    global CHECKS
    CHECKS += 1
    if not condition:
        raise SystemExit(f"scaffold test failed: {message}")


def verify_module():
    spec = importlib.util.spec_from_file_location("pptx_ink_actions_verify", VERIFY)
    check(spec is not None and spec.loader is not None, "load verifier module")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def sentinel(root: Path, target: Path) -> str:
    return f"pptx-ink-actions-profile-target-v1\nroot={root}\ntarget={target}\n"


def success_sentinel(root: Path, target: Path) -> str:
    return f"pptx-ink-actions-profile-success-v1\nroot={root}\ntarget={target}\n"


def run_cleanup(status: int, root: Path, target: Path, sentinel_path: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [str(CLEANUP), str(status), str(root), str(target), str(sentinel_path)],
        check=False,
        capture_output=True,
        text=True,
    )


def cleanup_control_tests() -> None:
    """Exercise the same cleanup executable used by run_profile.sh."""
    with tempfile.TemporaryDirectory(prefix="pptx-ink-actions-cleanup-") as raw:
        temporary = Path(raw)
        root = temporary / "root"
        targets = temporary / "targets"
        root.mkdir()
        targets.mkdir()
        sibling = targets / "sibling"
        sibling.mkdir()
        (sibling / "keep").write_text("outside target")

        owned = targets / "owned"
        owned.mkdir()
        (owned / "payload").write_text("owned target")
        owned_sentinel = owned / ".pptx-ink-actions-profile-target.sentinel"
        owned_sentinel.write_text(sentinel(root, owned))
        (owned / ".pptx-ink-actions-profile-success.sentinel").write_text(success_sentinel(root, owned))
        result = run_cleanup(0, root, owned, owned_sentinel)
        check(result.returncode == 0, "verified sentinel cleanup status")
        check(not owned.exists(), "verified sentinel deletes owned target")
        check(not owned_sentinel.exists(), "verified cleanup removes owned sentinel with target")
        check(sibling.exists() and (sibling / "keep").is_file(), "cleanup stayed inside target")
        check(root.is_dir() and targets.is_dir(), "verified cleanup keeps containing directories")

        cases = {
            "missing sentinel": None,
            "malformed sentinel": "not-a-profile-sentinel\n",
            "mismatching root": sentinel(temporary / "other-root", sibling),
            "mismatching target": sentinel(root, sibling),
        }
        for label, content in cases.items():
            target = targets / label.replace(" ", "-")
            target.mkdir()
            marker = target / "payload"
            marker.write_text(label)
            sentinel_path = target / ".pptx-ink-actions-profile-target.sentinel"
            if content is not None:
                sentinel_path.write_text(content)
            (target / ".pptx-ink-actions-profile-success.sentinel").write_text(success_sentinel(root, target))
            result = run_cleanup(0, root, target, sentinel_path)
            check(result.returncode != 0, f"{label} is rejected")
            check(target.is_dir() and marker.is_file(), f"{label} retains target")

        success_cases = {
            "missing success sentinel": None,
            "malformed success sentinel": "not-a-profile-success-sentinel\n",
            "mismatching success root": success_sentinel(temporary / "other-root", sibling),
            "mismatching success target": success_sentinel(root, sibling),
        }
        for label, content in success_cases.items():
            target = targets / label.replace(" ", "-")
            target.mkdir()
            marker = target / "payload"
            marker.write_text(label)
            target_sentinel = target / ".pptx-ink-actions-profile-target.sentinel"
            target_sentinel.write_text(sentinel(root, target))
            success_path = target / ".pptx-ink-actions-profile-success.sentinel"
            if content is not None:
                success_path.write_text(content)
            result = run_cleanup(0, root, target, target_sentinel)
            check(result.returncode != 0, f"{label} is rejected")
            check(target.is_dir() and marker.is_file(), f"{label} retains target")

        failed = targets / "failed-profile"
        failed.mkdir()
        (failed / "payload").write_text("diagnostic")
        failed_sentinel = failed / ".pptx-ink-actions-profile-target.sentinel"
        failed_sentinel.write_text(sentinel(root, failed))
        (failed / ".pptx-ink-actions-profile-success.sentinel").write_text(success_sentinel(root, failed))
        result = run_cleanup(7, root, failed, failed_sentinel)
        check(result.returncode == 7, "failed profile status is preserved")
        check(failed.is_dir(), "failed profile retains valid target")

        symlink_target = targets / "symlink-target"
        symlink_payload = targets / "symlink-payload"
        symlink_payload.mkdir()
        (symlink_payload / "keep").write_text("symlink target")
        symlink_target.symlink_to(symlink_payload, target_is_directory=True)
        symlink_sentinel = symlink_target / ".pptx-ink-actions-profile-target.sentinel"
        symlink_sentinel.write_text(sentinel(root, symlink_target))
        (symlink_target / ".pptx-ink-actions-profile-success.sentinel").write_text(success_sentinel(root, symlink_target))
        result = run_cleanup(0, root, symlink_target, symlink_sentinel)
        check(result.returncode != 0, "symlink target is rejected")
        check(symlink_target.is_symlink() and (symlink_payload / "keep").is_file(), "symlink target retained")

        symlink_sentinel_target = targets / "symlink-sentinel-target"
        symlink_sentinel_target.mkdir()
        (symlink_sentinel_target / "payload").write_text("sentinel link target")
        real_sentinel = targets / "real-sentinel"
        real_sentinel.write_text(sentinel(root, symlink_sentinel_target))
        linked_sentinel = symlink_sentinel_target / ".pptx-ink-actions-profile-target.sentinel"
        linked_sentinel.symlink_to(real_sentinel)
        (symlink_sentinel_target / ".pptx-ink-actions-profile-success.sentinel").write_text(
            success_sentinel(root, symlink_sentinel_target)
        )
        result = run_cleanup(0, root, symlink_sentinel_target, linked_sentinel)
        check(result.returncode != 0, "symlink sentinel is rejected")
        check(symlink_sentinel_target.is_dir() and (symlink_sentinel_target / "payload").is_file(), "symlink sentinel retains target")
        check(real_sentinel.is_file(), "symlink sentinel target file is retained")

        success_link_target = targets / "success-symlink-target"
        success_link_target.mkdir()
        (success_link_target / "payload").write_text("success sentinel link")
        success_target_sentinel = success_link_target / ".pptx-ink-actions-profile-target.sentinel"
        success_target_sentinel.write_text(sentinel(root, success_link_target))
        real_success_sentinel = targets / "real-success-sentinel"
        real_success_sentinel.write_text(success_sentinel(root, success_link_target))
        linked_success = success_link_target / ".pptx-ink-actions-profile-success.sentinel"
        linked_success.symlink_to(real_success_sentinel)
        result = run_cleanup(0, root, success_link_target, success_target_sentinel)
        check(result.returncode != 0, "symlink success sentinel is rejected")
        check(success_link_target.is_dir() and (success_link_target / "payload").is_file(), "symlink success sentinel retains target")
        check(real_success_sentinel.is_file(), "symlink success sentinel target file is retained")

        file_target = targets / "file-target"
        file_target.write_text("target is not a directory")
        file_sentinel = targets / "file-target.sentinel"
        file_sentinel.write_text(sentinel(root, file_target))
        result = run_cleanup(0, root, file_target, file_sentinel)
        check(result.returncode != 0, "non-directory target is rejected")
        check(file_target.is_file(), "non-directory target retained")


def capture_gate_test() -> None:
    """The frozen gate must refuse before creating results or target paths."""
    with tempfile.TemporaryDirectory(prefix="pptx-ink-actions-gate-") as raw:
        temporary = Path(raw)
        results = temporary / "results"
        target = temporary / "target"
        environment = {
            key: value
            for key, value in os.environ.items()
            if not key.startswith("PPTX_INK_ACTIONS_")
        }
        environment.update(
            {
                "PPTX_INK_ACTIONS_PROFILE_RESULTS": str(results),
                "PPTX_INK_ACTIONS_PROFILE_TARGET": str(target),
            }
        )
        result = subprocess.run(
            [str(RUNNER)],
            cwd=ROOT,
            env=environment,
            check=False,
            capture_output=True,
            text=True,
        )
        check(result.returncode == 2, "capture gate refuses without authorization")
        check(not results.exists() and not target.exists(), "capture gate created no paths")


def git_fixture() -> tuple[Path, str, Path, str]:
    """Create a tiny committed source closure for guard probes.

    The real guard and source-manifest commands are run against this isolated
    repository.  This lets the tests mutate local Cargo sources and root build
    inputs without dirtying the production checkout.
    """
    temporary = Path(tempfile.mkdtemp(prefix="pptx-ink-actions-source-"))
    root = temporary / "repo"
    (root / ".cargo").mkdir(parents=True)
    (root / "crates/local/src").mkdir(parents=True)
    files = {
        "Cargo.toml": "[workspace]\nmembers = [\"crates/local\"]\n",
        "rust-toolchain.toml": "[toolchain]\nchannel = \"stable\"\n",
        ".cargo/config.toml": "[build]\ntarget-dir = \"target\"\n",
        "crates/local/Cargo.toml": (
            "[package]\nname = \"fixture-local\"\nversion = \"0.1.0\"\n"
            "edition = \"2021\"\n\n[lib]\npath = \"src/lib.rs\"\n"
        ),
        "crates/local/src/lib.rs": "pub mod extra;\npub fn value() -> u32 { 1 }\n",
        "crates/local/src/extra.rs": "pub const EXTRA: u32 = 2;\n",
        "fixture_helper.rs": "// committed fixture helper\n",
        SEMANTIC_OWNER_DESIGN_PATH: "# committed semantic-owner design fixture\n",
    }
    for relative, content in files.items():
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content)

    metadata_path = root / "metadata.json"
    metadata_path.write_text(
        json.dumps(
            {
                "packages": [
                    {
                        "name": "fixture-local",
                        "version": "0.1.0",
                        "source": None,
                        "manifest_path": str(root / "crates/local/Cargo.toml"),
                        "targets": [
                            {"src_path": str(root / "crates/local/src/lib.rs")},
                        ],
                    }
                ]
            },
            sort_keys=True,
        )
        + "\n"
    )

    def git(*args: str) -> None:
        subprocess.run(["git", *args], cwd=root, check=True, capture_output=True, text=True)

    git("init", "--quiet", "--initial-branch=main")
    git("config", "user.email", "scaffold@example.invalid")
    git("config", "user.name", "PPTX scaffold test")
    git("add", "--all")
    git("commit", "--quiet", "-m", "fixture source")
    commit = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=root, text=True).strip()
    helper_sha256 = hashlib.sha256((root / "fixture_helper.rs").read_bytes()).hexdigest()
    return temporary, commit, metadata_path, helper_sha256


def semantic_owner_design_args(root: Path, semantic_owner_commit: str) -> list[str]:
    """Return the real semantic-owner receipt for the classified design path."""
    blob = subprocess.check_output(
        ["git", "rev-parse", f"{semantic_owner_commit}:{SEMANTIC_OWNER_DESIGN_PATH}"],
        cwd=root,
        text=True,
    ).strip()
    content = subprocess.check_output(
        ["git", "cat-file", "blob", f"{semantic_owner_commit}:{SEMANTIC_OWNER_DESIGN_PATH}"],
        cwd=root,
    )
    digest = hashlib.sha256(content).hexdigest()
    return [
        "--semantic-owner-extra",
        SEMANTIC_OWNER_DESIGN_PATH,
        "--semantic-owner-extra-sha256",
        digest,
        "--semantic-owner-extra-blob",
        blob,
    ]


def source_guard_args(
    root: Path,
    semantic_owner_commit: str,
    production_source_baseline_commit: str,
    metadata: Path,
    helper_sha256: str,
    *,
    include_root_paths: bool = True,
    include_semantic_owner_design: bool = False,
    production_extras: tuple[str, ...] = (),
) -> list[str]:
    args = [
        sys.executable,
        "-B",
        str(COMMITTED_INPUTS),
        "--root",
        str(root),
        "--semantic-owner-commit",
        semantic_owner_commit,
        "--production-source-baseline-commit",
        production_source_baseline_commit,
        "--helper",
        "fixture_helper.rs",
        "--helper-sha256",
        helper_sha256,
        "--metadata",
        str(metadata),
    ]
    if include_root_paths:
        for path in ("Cargo.toml", "rust-toolchain.toml", ".cargo/config.toml"):
            args.extend(("--path", path))
    else:
        # argparse requires at least one explicit path; keep this probe's
        # ordinary source input stable so owner-extra is the failing seam.
        args.extend(("--path", "crates/local/src/lib.rs"))
    if include_semantic_owner_design:
        args.extend(semantic_owner_design_args(root, semantic_owner_commit))
    for path in production_extras:
        args.extend(("--production-extra", path))
    return args


def source_manifest_args(
    root: Path,
    capture_head: str,
    metadata: Path,
    output: Path,
    *,
    semantic_owner_commit: str,
    production_source_baseline_commit: str,
    include_semantic_owner_design: bool = False,
    production_extras: tuple[str, ...] = (),
) -> list[str]:
    args = [
        sys.executable,
        "-B",
        str(SOURCE_MANIFEST),
        "--metadata",
        str(metadata),
        "--root",
        str(root),
        "--output",
        str(output),
        "--git-commit",
        capture_head,
        "--semantic-owner-commit",
        semantic_owner_commit,
        "--production-source-baseline-commit",
        production_source_baseline_commit,
        "--extra",
        "Cargo.toml",
        "--extra",
        "rust-toolchain.toml",
        "--extra",
        ".cargo/config.toml",
        "--extra",
        "fixture_helper.rs",
    ]
    if include_semantic_owner_design:
        args.extend(semantic_owner_design_args(root, semantic_owner_commit))
    for path in production_extras:
        args.extend(("--production-extra", path))
    return args


def source_pin_tamper_tests() -> None:
    """Exercise separate semantic-owner, production-baseline, and capture pins."""
    temporary, semantic_owner_commit, metadata, helper_sha256 = git_fixture()
    try:
        root = temporary / "repo"
        manifest_output = temporary / "source-manifest.txt"
        production_extras = ("Cargo.toml", "rust-toolchain.toml", ".cargo/config.toml")
        baseline_guard = subprocess.run(
            source_guard_args(
                root,
                semantic_owner_commit,
                semantic_owner_commit,
                metadata,
                helper_sha256,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(baseline_guard.returncode == 0, "source guard accepts committed fixture")
        baseline_manifest = subprocess.run(
            source_manifest_args(
                root,
                semantic_owner_commit,
                metadata,
                manifest_output,
                semantic_owner_commit=semantic_owner_commit,
                production_source_baseline_commit=semantic_owner_commit,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(baseline_manifest.returncode == 0, "source manifest accepts committed fixture")
        check(manifest_output.is_file(), "source manifest was materialized for fixture")

        mutations = (
            ("crates/local/src/extra.rs", "pub const EXTRA: u32 = 99;\n"),
            ("Cargo.toml", "[workspace]\nmembers = [\"crates/local\"]\n# tampered\n"),
            (
                "rust-toolchain.toml",
                "[toolchain]\nchannel = \"nightly\"\n",
            ),
            (
                ".cargo/config.toml",
                "[build]\ntarget-dir = \"tampered-target\"\n",
            ),
        )
        for relative, content in mutations:
            path = root / relative
            original = path.read_bytes()
            path.write_text(content)
            try:
                guard = subprocess.run(
                    source_guard_args(
                        root,
                        semantic_owner_commit,
                        semantic_owner_commit,
                        metadata,
                        helper_sha256,
                        production_extras=production_extras,
                    ),
                    cwd=root,
                    check=False,
                    capture_output=True,
                    text=True,
                )
                check(guard.returncode != 0, f"committed-input guard rejects {relative} tamper")
                check(relative in guard.stderr, f"guard identifies {relative} tamper")

                manifest = subprocess.run(
                    source_manifest_args(
                        root,
                        semantic_owner_commit,
                        metadata,
                        manifest_output,
                        semantic_owner_commit=semantic_owner_commit,
                        production_source_baseline_commit=semantic_owner_commit,
                        production_extras=production_extras,
                    ),
                    cwd=root,
                    check=False,
                    capture_output=True,
                    text=True,
                )
                check(manifest.returncode != 0, f"source manifest rejects {relative} tamper")
                check(relative in manifest.stderr, f"source manifest identifies {relative} tamper")
            finally:
                path.write_bytes(original)

        untracked = root / "crates/local/src/untracked.rs"
        untracked.write_text("pub const UNTRACKED: u32 = 3;\n")
        try:
            guard = subprocess.run(
                source_guard_args(
                    root,
                    semantic_owner_commit,
                    semantic_owner_commit,
                    metadata,
                    helper_sha256,
                    production_extras=production_extras,
                ),
                cwd=root,
                check=False,
                capture_output=True,
                text=True,
            )
            check(guard.returncode != 0, "guard rejects untracked transitive source")
            check("not tracked" in guard.stderr, "guard identifies untracked transitive source")
        finally:
            untracked.unlink(missing_ok=True)

        production_changes = {
            "Cargo.toml": "[workspace]\nmembers = [\"crates/local\"]\n# production baseline\n",
            "rust-toolchain.toml": "[toolchain]\nchannel = \"stable\"\n# production baseline\n",
            ".cargo/config.toml": "[build]\ntarget-dir = \"production-target\"\n",
        }
        for relative, content in production_changes.items():
            (root / relative).write_text(content)
        subprocess.run(
            ["git", "add", *production_changes],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        subprocess.run(
            ["git", "commit", "--quiet", "-m", "capture descendant workspace inputs"],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        production_baseline_commit = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=root, text=True
        ).strip()

        production_mutation = root / "crates/local/src/extra.rs"
        production_mutation.write_text("pub const EXTRA: u32 = 99;\n")
        subprocess.run(
            ["git", "add", "crates/local/src/extra.rs"],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        subprocess.run(
            ["git", "commit", "--quiet", "-m", "capture production mutation"],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        production_mutation_capture_head = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=root, text=True
        ).strip()
        check(
            production_mutation_capture_head != production_baseline_commit,
            "capture HEAD is distinct from production baseline",
        )

        production_guard = subprocess.run(
            source_guard_args(
                root,
                semantic_owner_commit,
                production_baseline_commit,
                metadata,
                helper_sha256,
                include_root_paths=False,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(
            production_guard.returncode != 0,
            "committed production mutation is rejected against baseline",
        )
        check("production baseline" in production_guard.stderr, "guard identifies production baseline drift")
        tracked_production_mutation = subprocess.run(
            ["git", "ls-files", "--error-unmatch", "--", "crates/local/src/extra.rs"],
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(
            tracked_production_mutation.returncode == 0,
            "production mutation is tracked before baseline rejection",
        )

        production_manifest = subprocess.run(
            source_manifest_args(
                root,
                production_mutation_capture_head,
                metadata,
                manifest_output,
                semantic_owner_commit=semantic_owner_commit,
                production_source_baseline_commit=production_baseline_commit,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(
            production_manifest.returncode != 0,
            "source manifest rejects committed production mutation",
        )
        check(
            "production baseline" in production_manifest.stderr,
            "source manifest identifies production baseline drift",
        )

        wrong_baseline_manifest = subprocess.run(
            source_manifest_args(
                root,
                production_mutation_capture_head,
                metadata,
                manifest_output,
                semantic_owner_commit=semantic_owner_commit,
                production_source_baseline_commit=semantic_owner_commit,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(wrong_baseline_manifest.returncode != 0, "wrong production baseline receipt is rejected")
        check(
            "production baseline" in wrong_baseline_manifest.stderr,
            "wrong production baseline is identified",
        )

        wrong_capture_manifest = subprocess.run(
            source_manifest_args(
                root,
                production_baseline_commit,
                metadata,
                manifest_output,
                semantic_owner_commit=semantic_owner_commit,
                production_source_baseline_commit=production_baseline_commit,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(wrong_capture_manifest.returncode != 0, "wrong capture HEAD receipt is rejected")
        check("source HEAD changed" in wrong_capture_manifest.stderr, "wrong capture HEAD is identified")

        owner_mutation = root / "fixture_helper.rs"
        owner_mutation.write_text("// committed semantic owner mutation\n")
        subprocess.run(
            ["git", "add", "fixture_helper.rs"],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        subprocess.run(
            ["git", "commit", "--quiet", "-m", "capture semantic owner mutation"],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        owner_mutation_capture_head = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=root, text=True
        ).strip()
        owner_guard = subprocess.run(
            source_guard_args(
                root,
                semantic_owner_commit,
                production_baseline_commit,
                metadata,
                helper_sha256,
                include_root_paths=False,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(owner_guard.returncode != 0, "semantic owner guard rejects committed owner mutation")
        check(
            "semantic owner" in owner_guard.stderr,
            "semantic owner guard identifies owner drift",
        )
        tracked_owner_mutation = subprocess.run(
            ["git", "ls-files", "--error-unmatch", "--", "fixture_helper.rs"],
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(
            tracked_owner_mutation.returncode == 0,
            "semantic owner mutation is tracked before owner rejection",
        )
    finally:
        # The fixture is the only temporary tree owned by this test.
        shutil.rmtree(temporary)


def semantic_owner_design_tamper_test() -> None:
    """Reject a committed mutation of the classified semantic design input."""
    temporary, semantic_owner_commit, metadata, helper_sha256 = git_fixture()
    try:
        root = temporary / "repo"
        manifest_output = temporary / "source-manifest.txt"
        production_extras = ("Cargo.toml", "rust-toolchain.toml", ".cargo/config.toml")
        baseline_guard = subprocess.run(
            source_guard_args(
                root,
                semantic_owner_commit,
                semantic_owner_commit,
                metadata,
                helper_sha256,
                include_semantic_owner_design=True,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(baseline_guard.returncode == 0, "guard accepts pinned semantic design input")

        baseline_manifest = subprocess.run(
            source_manifest_args(
                root,
                semantic_owner_commit,
                metadata,
                manifest_output,
                semantic_owner_commit=semantic_owner_commit,
                production_source_baseline_commit=semantic_owner_commit,
                include_semantic_owner_design=True,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(
            baseline_manifest.returncode == 0,
            "source manifest accepts pinned semantic design input",
        )

        design = root / SEMANTIC_OWNER_DESIGN_PATH
        design.write_text(design.read_text() + "\n# committed semantic-owner tamper\n")
        subprocess.run(
            ["git", "add", SEMANTIC_OWNER_DESIGN_PATH],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        subprocess.run(
            ["git", "commit", "--quiet", "-m", "capture semantic design mutation"],
            cwd=root,
            check=True,
            capture_output=True,
            text=True,
        )
        capture_head = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=root, text=True
        ).strip()
        check(capture_head != semantic_owner_commit, "design tamper advances capture HEAD")
        tracked = subprocess.run(
            ["git", "ls-files", "--error-unmatch", "--", SEMANTIC_OWNER_DESIGN_PATH],
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(tracked.returncode == 0, "design tamper is committed and tracked")

        guard = subprocess.run(
            source_guard_args(
                root,
                semantic_owner_commit,
                semantic_owner_commit,
                metadata,
                helper_sha256,
                include_semantic_owner_design=True,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(guard.returncode != 0, "guard rejects committed semantic design tamper")
        check(
            SEMANTIC_OWNER_DESIGN_PATH in guard.stderr,
            "guard identifies committed semantic design tamper",
        )

        manifest = subprocess.run(
            source_manifest_args(
                root,
                capture_head,
                metadata,
                manifest_output,
                semantic_owner_commit=semantic_owner_commit,
                production_source_baseline_commit=semantic_owner_commit,
                include_semantic_owner_design=True,
                production_extras=production_extras,
            ),
            cwd=root,
            check=False,
            capture_output=True,
            text=True,
        )
        check(manifest.returncode != 0, "source manifest rejects committed semantic design tamper")
        check(
            SEMANTIC_OWNER_DESIGN_PATH in manifest.stderr,
            "source manifest identifies committed semantic design tamper",
        )
    finally:
        shutil.rmtree(temporary)


def run_matrix_acceptance_guard() -> None:
    """Evaluate the live Rust matrix predicate without building or timing it."""
    adapter = (HERE / "harness/adapter.rs").read_text()
    matrix_start = adapter.find("pub fn run_matrix()")
    check(matrix_start >= 0, "matrix helper is present")
    match = re.search(
        r"\bif\s+(receipt\.[a-z_]+(?:\s*&&\s*receipt\.[a-z_]+)*)\s*\{",
        adapter[matrix_start:],
    )
    check(match is not None, "matrix helper exposes an acceptance predicate")
    expression = " ".join(match.group(1).split())
    fields = tuple(re.findall(r"receipt\.([a-z_]+)", expression))
    expected_fields = (
        "semantic_ok",
        "preservation_ok",
        "inverse_ok",
        "baseline_reopenable",
        "retained_baseline_balance_ok",
    )
    check(fields == expected_fields, "matrix predicate covers semantic and preservation invariants")

    translated = re.sub(
        r"receipt\.([a-z_]+)",
        lambda item: f"values[{item.group(1)!r}]",
        expression,
    ).replace("&&", " and ")

    def accepts(values: dict[str, bool]) -> bool:
        return bool(eval(translated, {"__builtins__": {}}, {"values": values}))

    all_true = {field: True for field in expected_fields}
    check(accepts(all_true), "matrix predicate accepts a fully valid receipt")
    for field in ("baseline_reopenable", "retained_baseline_balance_ok"):
        rejected = dict(all_true, **{field: False})
        check(not accepts(rejected), f"matrix predicate rejects false {field}")


def generator_receipt_preflight_test(corpus: dict[str, object]) -> None:
    """Check the retained adapter bytes and receipt before any build step."""
    verifier = verify_module()
    generator = corpus["retained_opc_generator"]
    adapter = HERE / str(generator["path"])
    actual_sha256 = hashlib.sha256(adapter.read_bytes()).hexdigest()
    actual_git_blob = subprocess.check_output(
        ["git", "hash-object", str(adapter)],
        cwd=ROOT,
        text=True,
    ).strip()
    check(actual_sha256 == generator["sha256"], "current adapter SHA-256 matches generator receipt")
    check(actual_git_blob == generator["git_blob"], "current adapter Git blob matches generator receipt")
    verifier.verify_corpus(HERE / "corpus-manifest.json")
    check(True, "generator receipt preflight accepts current adapter")

    def rejected_receipt(payload: dict[str, object], label: str, message: str) -> None:
        receipt = HERE / f".scaffold-test-generator-{label}.json"
        receipt.write_text(json.dumps(payload))
        try:
            probe = subprocess.run(
                [
                    sys.executable,
                    "-c",
                    "import importlib.util, pathlib, sys; p=pathlib.Path(sys.argv[1]); s=importlib.util.spec_from_file_location('v', pathlib.Path(sys.argv[2])); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); m.verify_corpus(p)",
                    str(receipt),
                    str(VERIFY),
                ],
                cwd=ROOT,
                check=False,
                capture_output=True,
                text=True,
            )
            check(probe.returncode != 0, f"{label} generator receipt tamper is rejected before build")
            check(message in probe.stderr, f"{label} generator receipt failure is identified")
        finally:
            receipt.unlink(missing_ok=True)

    for field, message in (
        ("sha256", "retained generator hash changed"),
        ("git_blob", "retained generator Git blob changed"),
    ):
        payload = json.loads(json.dumps(corpus))
        payload["retained_opc_generator"][field] = "0" * 64
        rejected_receipt(payload, f"receipt-{field}", message)

    with tempfile.TemporaryDirectory(prefix="pptx-ink-actions-generator-") as raw:
        isolated_root = Path(raw)
        (isolated_root / ".git").symlink_to(ROOT / ".git", target_is_directory=True)
        (isolated_root / "crates").symlink_to(ROOT / "crates", target_is_directory=True)
        design = isolated_root / SEMANTIC_OWNER_DESIGN_PATH
        design.parent.mkdir(parents=True)
        design.symlink_to(ROOT / SEMANTIC_OWNER_DESIGN_PATH)
        performance = isolated_root / "docs/report/spec-gap-validation-evidence/pptx-ink-actions-performance"
        harness = performance / "harness"
        harness.mkdir(parents=True)
        tampered_adapter = harness / "adapter.rs"
        tampered_adapter.write_bytes(adapter.read_bytes() + b"\n// preflight adapter tamper\n")
        (harness / "Cargo.lock").symlink_to(HERE / "harness/Cargo.lock")
        tampered_corpus = performance / "corpus-manifest.json"
        tampered_corpus.write_text(json.dumps(corpus))
        probe = subprocess.run(
            [
                sys.executable,
                "-c",
                "import importlib.util, pathlib, sys; p=pathlib.Path(sys.argv[1]); s=importlib.util.spec_from_file_location('v', pathlib.Path(sys.argv[2])); module=importlib.util.module_from_spec(s); s.loader.exec_module(module); module.verify_corpus(p)",
                str(tampered_corpus),
                str(VERIFY),
            ],
            cwd=ROOT,
            check=False,
            capture_output=True,
            text=True,
        )
        check(probe.returncode != 0, "tampered adapter is rejected before build")
        check("retained generator hash changed" in probe.stderr, "tampered adapter hash failure is identified")


def owner_extra_runner_wiring_test() -> None:
    """Keep the runner's dual source pins coupled to the real guards."""
    runner = RUNNER.read_text()
    check("SEMANTIC_OWNER_COMMIT=" in runner, "runner declares semantic owner pin")
    check(
        "PRODUCTION_SOURCE_BASELINE_COMMIT=" in runner,
        "runner declares production baseline pin",
    )
    check(
        "PRODUCTION_WORKSPACE_EXTRAS=(" in runner,
        "runner declares production workspace extras",
    )
    for path in ("$ROOT/Cargo.toml", "$ROOT/rust-toolchain.toml", "$ROOT/.cargo/config.toml"):
        check(path in runner, f"runner owns {path} as a production extra")
    check(
        'GUARD_ARGS+=(--production-extra "$path")' in runner,
        "runner passes production extras to committed-input guard",
    )
    check(
        'MANIFEST_ARGS+=(--production-extra "$extra")' in runner
        and 'MANIFEST_ARGS_AFTER+=(--production-extra "$extra")' in runner,
        "runner passes production extras to both source manifests",
    )
    check(
        '--semantic-owner-commit "$SEMANTIC_OWNER_COMMIT"' in runner
        and '--production-source-baseline-commit "$PRODUCTION_SOURCE_BASELINE_COMMIT"' in runner,
        "runner passes both source pins to committed-input and manifest commands",
    )
    check(
        "SEMANTIC_OWNER_DESIGN_PATH=docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md"
        in runner,
        "runner declares the classified semantic design path",
    )
    check(
        "SEMANTIC_OWNER_DESIGN_SHA256=" in runner
        and "SEMANTIC_OWNER_DESIGN_GIT_BLOB=" in runner,
        "runner declares semantic design hash and Git blob receipts",
    )
    check(
        'GUARD_ARGS+=(--semantic-owner-extra "$ROOT/$SEMANTIC_OWNER_DESIGN_PATH"' in runner,
        "runner passes semantic design receipt to committed-input guard",
    )
    check(
        'MANIFEST_ARGS+=(--semantic-owner-extra "$ROOT/$SEMANTIC_OWNER_DESIGN_PATH"' in runner
        and 'MANIFEST_ARGS_AFTER+=(--semantic-owner-extra "$ROOT/$SEMANTIC_OWNER_DESIGN_PATH"'
        in runner,
        "runner passes semantic design receipt to both source manifests",
    )


def verifier_tamper_tests(corpus: dict[str, object]) -> None:
    verifier = verify_module()
    expected = "Error::Limit { resource: ink-action anchor count, limit: 7 }"
    typed = {
        "actual_error_type": "Error::Limit",
        "actual_error_resource": "ink-action anchor count",
        "actual_error_limit": 7,
        "actual_error_debug": 'Limit { resource: "ink-action anchor count", limit: 7 }',
    }
    check(verifier.error_matches(expected, typed), "typed error receipt is accepted")
    wrong_type = dict(typed, actual_error_type="Error::StaleSource")
    check(not verifier.error_matches(expected, wrong_type), "wrong typed error is rejected")
    wrong_resource = dict(typed, actual_error_resource="ink-action target bytes")
    check(not verifier.error_matches(expected, wrong_resource), "wrong typed error resource is rejected")
    wrong_limit = dict(typed, actual_error_limit=8)
    check(not verifier.error_matches(expected, wrong_limit), "wrong limit boundary is rejected")
    wrong_limit_type = dict(typed, actual_error_limit="7")
    check(not verifier.error_matches(expected, wrong_limit_type), "non-numeric typed limit is rejected")
    debug_only = dict(typed, actual_error_debug='Limit { resource: "ink-action anchor count", limit: 8 }')
    check(verifier.error_matches(expected, debug_only), "debug-only typed receipt change is ignored")
    check(verifier.error_matches(expected, dict(typed, actual_error_debug="")), "empty debug text is ignored")

    def expect_bad_corpus(mutator, label: str) -> None:
        payload = json.loads(json.dumps(corpus))
        mutator(payload)
        temporary = HERE / f".scaffold-test-{label}.json"
        temporary.write_text(json.dumps(payload))
        try:
            probe = subprocess.run(
                [
                    sys.executable,
                    "-c",
                    "import importlib.util, pathlib, sys; p=pathlib.Path(sys.argv[1]); s=importlib.util.spec_from_file_location('v', pathlib.Path(sys.argv[2])); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); m.verify_corpus(p)",
                    str(temporary),
                    str(VERIFY),
                ],
                cwd=ROOT,
                check=False,
                capture_output=True,
                text=True,
            )
            check(probe.returncode != 0, f"{label} verifier failure is nonzero")
            if label.startswith("matrix"):
                with tempfile.TemporaryDirectory(prefix="pptx-ink-actions-verifier-") as raw:
                    target = Path(raw) / "target"
                    target.mkdir()
                    (target / "diagnostic").write_text("retain after verifier rejection")
                    sentinel_path = target / ".pptx-ink-actions-profile-target.sentinel"
                    sentinel_path.write_text(sentinel(ROOT, target))
                    (target / ".pptx-ink-actions-profile-success.sentinel").write_text(
                        success_sentinel(ROOT, target)
                    )
                    result = run_cleanup(1, ROOT, target, sentinel_path)
                    check(result.returncode == 1, "verifier rejection preserves failure status")
                    check(target.is_dir(), "verifier rejection retains target for diagnosis")
        finally:
            temporary.unlink(missing_ok=True)

    expect_bad_corpus(
        lambda payload: payload["lanes"][0].update(recipe="missing_recipe"),
        "matrix-unresolved-recipe",
    )
    expect_bad_corpus(
        lambda payload: payload["lanes"][1].update(id=payload["lanes"][0]["id"]),
        "matrix-duplicate-lane",
    )
    expect_bad_corpus(
        lambda payload: payload["recipes"].append(dict(payload["recipes"][0], id="orphan_recipe")),
        "matrix-orphan-recipe",
    )
    expect_bad_corpus(
        lambda payload: payload["retained_opc_generator"].update(sha256="0" * 64),
        "generator",
    )

    opaque_probe = {
        "schema": "pptx-ink-actions-host-probe-v1",
        "source_commit": verifier.SEMANTIC_OWNER_COMMIT,
        "semantic_owner_commit": verifier.SEMANTIC_OWNER_COMMIT,
        "production_source_baseline_commit": verifier.PRODUCTION_SOURCE_BASELINE_COMMIT,
        "capture_head": "capture-head-test",
        "helper_sha256": verifier.HELPER_SHA256,
        "native_powerpoint_claim": False,
        "synthetic_complete_opc": True,
        "package_anchors": 1,
        "presentation_anchors": 1,
        "opaque_choice_preserved": True,
        "opaque_fallback_preserved": True,
        "opaque_payload_preserved": True,
        "opaque_default_namespace_preserved": True,
        "opaque_prefix_preserved": True,
        "opaque_unknown_requires_preserved": True,
        "opaque_internal_outbound_preserved": True,
        "opaque_external_outbound_preserved": True,
        "opaque_manifest_preserved": True,
        "opaque_patch_publication_preserved": True,
    }
    with tempfile.TemporaryDirectory(prefix="pptx-ink-actions-receipt-") as raw:
        receipt = Path(raw) / "host-probe.json"
        receipt.write_text(json.dumps(opaque_probe, sort_keys=True))
        verifier.verify_host_probe(receipt, "capture-head-test")
        check(True, "valid opaque/source host receipt is accepted")
        receipt_mutations = (
            ("opaque_choice_preserved", False),
            ("opaque_fallback_preserved", False),
            ("opaque_payload_preserved", False),
            ("opaque_default_namespace_preserved", False),
            ("opaque_prefix_preserved", False),
            ("opaque_unknown_requires_preserved", False),
            ("opaque_internal_outbound_preserved", False),
            ("opaque_external_outbound_preserved", False),
            ("opaque_manifest_preserved", False),
            ("opaque_patch_publication_preserved", False),
            ("source_commit", "0" * 40),
            ("semantic_owner_commit", "0" * 40),
            ("production_source_baseline_commit", "0" * 40),
            ("capture_head", "wrong-capture-head"),
            ("helper_sha256", "0" * 64),
        )
        for field, value in receipt_mutations:
            mutated = dict(opaque_probe, **{field: value})
            receipt.write_text(json.dumps(mutated, sort_keys=True))
            try:
                verifier.verify_host_probe(receipt, "capture-head-test")
            except AssertionError:
                check(True, f"modified {field} receipt is rejected")
            else:
                check(False, f"modified {field} receipt is rejected")


def main() -> None:
    corpus = json.loads((HERE / "corpus-manifest.json").read_text())
    recipes = {item["id"] for item in corpus["recipes"]}
    lanes = corpus["lanes"]
    check(len(recipes) == 23, "recipe count")
    check(len(lanes) == 42, "lane count")
    check({item["recipe"] for item in lanes} == recipes, "every recipe is used")
    check(len({item["id"] for item in lanes}) == 42, "lane IDs are unique")
    check(corpus["native_powerpoint_claim"] is False, "native host claim")
    check(
        corpus["semantic_owner_commit"] == "cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd",
        "semantic owner pin",
    )
    check(
        corpus["production_source_baseline_commit"] == "2a2ffa1cae4e6b7070082768ce84483e5d411dc8",
        "production source baseline pin",
    )
    check(corpus["owner_commit"] == corpus["semantic_owner_commit"], "owner alias pin")

    contract = json.loads((HERE / "source-contract.json").read_text())
    check(contract["owner_commit"] == corpus["owner_commit"], "contract owner pin")
    check(
        contract["semantic_owner_commit"] == corpus["semantic_owner_commit"],
        "contract semantic owner pin",
    )
    check(
        contract["production_source_baseline_commit"] == corpus["production_source_baseline_commit"],
        "contract production source baseline pin",
    )
    design_corpus = corpus["semantic_owner_design"]
    check(
        design_corpus["commit"] == corpus["semantic_owner_commit"],
        "corpus semantic design owner pin",
    )
    check(design_corpus["path"] == SEMANTIC_OWNER_DESIGN_PATH, "corpus semantic design path")
    check(
        design_corpus["sha256"] == SEMANTIC_OWNER_DESIGN_SHA256,
        "corpus semantic design SHA-256",
    )
    check(
        design_corpus["git_blob"] == SEMANTIC_OWNER_DESIGN_GIT_BLOB,
        "corpus semantic design Git blob",
    )
    design_contract = contract["semantic_owner_design"]
    check(
        design_contract["commit"] == corpus["semantic_owner_commit"],
        "contract semantic design owner pin",
    )
    check(
        design_contract["path"] == SEMANTIC_OWNER_DESIGN_PATH,
        "contract classified semantic design path",
    )
    check(
        design_contract["sha256"] == SEMANTIC_OWNER_DESIGN_SHA256,
        "contract semantic design SHA-256",
    )
    check(
        design_contract["git_blob"] == SEMANTIC_OWNER_DESIGN_GIT_BLOB,
        "contract semantic design Git blob",
    )
    check(contract["goal_reference"]["committed_at_owner_pin"] is False, "untracked GOAL admission")
    check(contract["goal_reference"]["profile_input"] is False, "untracked GOAL source input")

    check(RUNNER.is_file() and os.access(RUNNER, os.X_OK), "runner executable")
    check(CLEANUP.is_file() and os.access(CLEANUP, os.X_OK), "cleanup helper executable")

    help_result = subprocess.run(
        [str(COMMITTED_INPUTS), "--help"],
        check=False,
        capture_output=True,
        text=True,
    )
    check(help_result.returncode == 0, "committed-input guard help")
    check("--manifest-path" in help_result.stdout, "transitive guard option")
    cleanup_control_tests()
    capture_gate_test()
    owner_extra_runner_wiring_test()
    source_pin_tamper_tests()
    semantic_owner_design_tamper_test()
    run_matrix_acceptance_guard()
    generator_receipt_preflight_test(corpus)
    verifier_tamper_tests(corpus)
    print(
        "PPTX InkAction scaffold tests passed (timing-free); "
        f"behavioral checks={CHECKS}; cleanup/capture/source-pin/matrix/typed-receipt seams covered"
    )


if __name__ == "__main__":
    main()
