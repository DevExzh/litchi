"""Adversarial tests for the stylesWithEffects smoke verifier.

These tests use small temporary receipts and disposable Git checkouts.  They
exercise the verifier's evidence binding without running the Cargo harness.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import shutil
import subprocess
from copy import deepcopy
from contextlib import contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest


VERIFY_PATH = Path(__file__).with_name("verify.py")
VERIFY_SPEC = importlib.util.spec_from_file_location("styles_effects_verify", VERIFY_PATH)
assert VERIFY_SPEC is not None and VERIFY_SPEC.loader is not None
VERIFY = importlib.util.module_from_spec(VERIFY_SPEC)
VERIFY_SPEC.loader.exec_module(VERIFY)


def digest_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


@contextmanager
def source_commit(value: str):
    original = VERIFY.SOURCE_COMMIT
    VERIFY.SOURCE_COMMIT = value
    try:
        yield
    finally:
        VERIFY.SOURCE_COMMIT = original


class ManifestCheckout:
    """A tiny committed source/evidence tree matching verify_manifest's pin."""

    EVIDENCE_FILES = (
        "docs/evidence/harness/Cargo.toml",
        "docs/evidence/harness/Cargo.lock",
        "docs/evidence/harness/adapter.rs",
        "docs/evidence/harness/main.rs",
        "docs/evidence/harness/support.rs",
        "docs/evidence/source_manifest.py",
        "docs/evidence/test_source_snapshot.py",
        "docs/evidence/test_verify.py",
        "docs/evidence/verify.py",
        "docs/evidence/requirements.md",
        "docs/evidence/corpus-manifest.json",
    )
    PACKAGE_FILES = (
        "crates/litchi-docx/Cargo.toml",
        "crates/litchi-docx/src/styles/effects.rs",
        "crates/litchi-docx/src/package/package/styles_with_effects.rs",
    )
    EXTRA_FILES = (
        "crates/litchi-docx/tests/styles_with_effects.rs",
        "crates/litchi-opc/src/phys_pkg.rs",
        "crates/litchi-opc/src/limits.rs",
    )

    def __init__(self, directory: Path):
        self.directory = directory
        self.root = directory / "checkout"
        self.root.mkdir()
        self.evidence = self.root / "docs/evidence"
        self.metadata = directory / "metadata.json"
        self.manifest = directory / "source-manifest.txt"
        self._write_inputs()
        self._git("add", ".")
        self._git(
            "-c",
            "user.name=Verifier Test",
            "-c",
            "user.email=verifier@example.invalid",
            "commit",
            "-qm",
            "initial source pin",
        )
        self.commit = self._git("rev-parse", "HEAD").stdout.strip()
        self.metadata.write_text("{}\n")
        self.write_manifest()

    def _git(self, *arguments: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            ["git", "-C", str(self.root), *arguments],
            check=True,
            capture_output=True,
            text=True,
        )

    def _write_inputs(self) -> None:
        subprocess.run(["git", "-C", str(self.root), "init", "-q"], check=True)
        (self.root / "crates/litchi-docx/src/styles").mkdir(parents=True)
        (self.root / "crates/litchi-docx/src/package/package").mkdir(parents=True)
        (self.root / "crates/litchi-docx/tests").mkdir(parents=True)
        (self.root / "crates/litchi-opc/src").mkdir(parents=True)
        for shown in self.PACKAGE_FILES:
            path = self.root / shown
            path.parent.mkdir(parents=True, exist_ok=True)
            if shown.endswith("Cargo.toml"):
                path.write_text('[package]\nname = "fixture-docx"\nversion = "0.1.0"\n')
            else:
                path.write_text(f"// committed fixture: {shown}\n")
        for shown in (*self.EVIDENCE_FILES, *self.EXTRA_FILES):
            path = self.root / shown
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(f"committed evidence input: {shown}\n")
        self.metadata.write_text("{}\n")

    def write_manifest(self, evidence_root: str = "docs/evidence") -> None:
        package_shown = sorted(self.PACKAGE_FILES)
        package_file_lines = [
            f"{shown}\t{VERIFY.sha256(self.root / shown)}" for shown in package_shown
        ]
        tree_hash = VERIFY.sha256_bytes("\n".join(package_file_lines).encode())
        lines = [
            f"format={VERIFY.MANIFEST_FORMAT}",
            f"source_commit={self.commit}",
            f"git_head={self.commit}",
            f"metadata_sha256={VERIFY.sha256(self.metadata)}",
            f"evidence_root={evidence_root}",
            "package="
            + "\t".join(
                (
                    "fixture-docx",
                    "0.1.0",
                    "path",
                    "crates/litchi-docx/Cargo.toml",
                    VERIFY.sha256(self.root / "crates/litchi-docx/Cargo.toml"),
                    str(len(package_shown)),
                    tree_hash,
                )
            ),
        ]
        lines.extend(
            "file="
            + "\t".join(
                (
                    "fixture-docx",
                    "0.1.0",
                    "crates/litchi-docx/Cargo.toml",
                    shown,
                    digest,
                )
            )
            for shown, digest in (line.split("\t", 1) for line in package_file_lines)
        )
        lines.extend(
            f"extra=\t{shown}\t{VERIFY.sha256(self.root / shown)}"
            for shown in (*self.EVIDENCE_FILES, *self.EXTRA_FILES)
        )
        self.manifest.write_text("\n".join(lines) + "\n")

    def recompute_extra_digest(self, shown: str) -> None:
        digest = VERIFY.sha256(self.root / shown)
        lines = self.manifest.read_text().splitlines()
        prefix = f"extra=\t{shown}\t"
        replaced = False
        for index, line in enumerate(lines):
            if line.startswith(prefix):
                lines[index] = prefix + digest
                replaced = True
                break
        if not replaced:
            raise AssertionError(f"test fixture has no extra record for {shown}")
        self.manifest.write_text("\n".join(lines) + "\n")


class ManifestBindingTests(unittest.TestCase):
    def checkout(self) -> ManifestCheckout:
        temporary = TemporaryDirectory(prefix="styles-effects-verify-git-")
        self.addCleanup(temporary.cleanup)
        return ManifestCheckout(Path(temporary.name))

    def verify_manifest(self, fixture: ManifestCheckout) -> None:
        with source_commit(fixture.commit):
            VERIFY.verify_manifest(
                fixture.manifest,
                fixture.root,
                fixture.metadata,
                fixture.evidence,
            )

    def test_recomputed_changed_evidence_blob_is_rejected_by_bound_git_blob(self):
        fixture = self.checkout()
        self.verify_manifest(fixture)

        changed = fixture.root / "docs/evidence/harness/main.rs"
        changed.write_text("changed after the evidence commit\n")
        fixture.recompute_extra_digest("docs/evidence/harness/main.rs")

        with self.assertRaisesRegex(AssertionError, "Git blob changed"):
            self.verify_manifest(fixture)

    def test_manifest_with_wrong_retained_evidence_root_is_rejected(self):
        fixture = self.checkout()
        lines = fixture.manifest.read_text().replace(
            "evidence_root=docs/evidence\n", "evidence_root=docs/other\n"
        )
        fixture.manifest.write_text(lines)

        with self.assertRaisesRegex(AssertionError, "staged evidence root changed"):
            self.verify_manifest(fixture)


def make_build_receipts(directory: Path, lanes: list[str]) -> dict[str, Path | str]:
    results = directory / "results"
    target = directory / "target"
    results.mkdir()
    binary = target / "debug/docx-styles-effects-smoke"
    binary.parent.mkdir(parents=True)
    binary.write_bytes(b"small disposable smoke executable\n")
    digest = VERIFY.sha256(binary)
    manifest_before = directory / "manifest-before.txt"
    manifest_after = directory / "manifest-after.txt"
    metadata_before = directory / "metadata-before.json"
    metadata_after = directory / "metadata-after.json"
    manifest_before.write_text("source manifest\n")
    manifest_after.write_text(manifest_before.read_text())
    metadata_before.write_text('{"packages":[]}\n')
    metadata_after.write_text(metadata_before.read_text())
    current_head = "a" * 40

    (results / "smoke-build.log").write_text(
        "Finished `dev` profile [unoptimized + debuginfo] target(s)\n"
    )
    for name in ("smoke-binary.sha256", "smoke-binary-after.sha256"):
        (results / name).write_text(f"{digest} {binary}\n")
    (results / "smoke-build-provenance.txt").write_text(
        "\n".join(
            (
                f"{digest} {binary}",
                f"binary={binary}",
                f"source_commit={VERIFY.SOURCE_COMMIT}",
                f"git_head={current_head}",
                "git_status=clean",
                "rustc=rustc 1.80.0",
                "cargo=cargo 1.80.0",
                f"target={target}",
                "allocator=CountingAllocator (process-local GlobalAlloc observer)",
                "rss=/usr/bin/time -v Maximum resident set size",
                "mode=smoke processes=1 warmup=0 samples=1",
            )
        )
        + "\n"
    )
    command_lines = [
        "command=cargo metadata --format-version=1 --locked --offline",
        "command=python3 source_manifest.py --metadata metadata.json",
        "command=cargo build --locked --offline",
        "environment="
        "CARGO_TARGET_DIR=/tmp/styles-effects-target CARGO_INCREMENTAL=0 "
        "LC_ALL=C RUSTFLAGS=unset CARGO_ENCODED_RUSTFLAGS=unset "
        "RUSTC_BOOTSTRAP=unset RUSTDOCFLAGS=unset",
    ]
    command_lines.extend(
        f"run=/usr/bin/time -v smoke --lane {lane} --warmup 0 --samples 1 "
        "(fresh_process=1)"
        for lane in lanes
    )
    (results / "smoke-commands.txt").write_text("\n".join(command_lines) + "\n")
    (results / "smoke-source-provenance.txt").write_text(
        "\n".join(
            (
                f"source_manifest_sha256={VERIFY.sha256(manifest_before)}",
                f"metadata_before_sha256={VERIFY.sha256(metadata_before)}",
                f"metadata_after_sha256={VERIFY.sha256(metadata_after)}",
                f"git_head={current_head}",
            )
        )
        + "\n"
    )
    return {
        "results": results,
        "target": target,
        "binary": binary,
        "manifest_before": manifest_before,
        "manifest_after": manifest_after,
        "metadata_before": metadata_before,
        "metadata_after": metadata_after,
        "current_head": current_head,
        "lanes": lanes,
    }


def verify_build_state(state: dict[str, Path | str]) -> None:
    VERIFY.verify_build_receipts(
        state["results"],
        state["current_head"],
        state["manifest_before"],
        state["manifest_after"],
        state["metadata_before"],
        state["metadata_after"],
        state["lanes"],
    )


class BuildReceiptTests(unittest.TestCase):
    def state(self) -> tuple[TemporaryDirectory[str], dict[str, Path | str]]:
        temporary = TemporaryDirectory(prefix="styles-effects-build-receipts-")
        return temporary, make_build_receipts(Path(temporary.name), ["lane_one"])

    def test_missing_build_log_is_rejected(self):
        temporary, state = self.state()
        self.addCleanup(temporary.cleanup)
        state["results"].joinpath("smoke-build.log").unlink()

        with self.assertRaisesRegex(AssertionError, "build log"):
            verify_build_state(state)

    def test_tampered_binary_digest_receipt_is_rejected(self):
        temporary, state = self.state()
        self.addCleanup(temporary.cleanup)
        binary = state["binary"]
        state["results"].joinpath("smoke-binary-after.sha256").write_text(
            f"{'0' * 64} {binary}\n"
        )

        with self.assertRaisesRegex(AssertionError, "binary changed"):
            verify_build_state(state)

    def test_tampered_build_provenance_is_rejected(self):
        temporary, state = self.state()
        self.addCleanup(temporary.cleanup)
        provenance = state["results"].joinpath("smoke-build-provenance.txt")
        provenance.write_text(provenance.read_text().replace("git_status=clean", "git_status=dirty"))

        with self.assertRaisesRegex(AssertionError, "clean"):
            verify_build_state(state)

    def test_missing_or_tampered_command_receipt_is_rejected(self):
        for tamper in ("missing", "tampered"):
            with self.subTest(tamper=tamper):
                temporary, state = self.state()
                self.addCleanup(temporary.cleanup)
                commands = state["results"].joinpath("smoke-commands.txt")
                if tamper == "missing":
                    commands.unlink()
                else:
                    commands.write_text(commands.read_text().replace("--warmup 0", "--warmup 1"))

                with self.assertRaisesRegex(AssertionError, "command"):
                    verify_build_state(state)

    def test_replay_is_valid_after_disposable_executable_cleanup(self):
        temporary, state = self.state()
        self.addCleanup(temporary.cleanup)
        shutil.rmtree(state["target"])

        # The retained hash/provenance receipts still identify the executable;
        # replay is valid when its separately disposable target is also gone.
        verify_build_state(state)


def metrics(value: int = 1) -> dict[str, int]:
    return {field: value for field in VERIFY.METRIC_FIELDS}


def native_success_sample(input_sha256: str) -> dict[str, object]:
    sample: dict[str, object] = {
        field: 0
        for field in VERIFY.COUNTER_FIELDS
    }
    sample.update(
        {
            "elapsed_ns": 1,
            "output_bytes": 1,
            "actual_success": True,
            "ingress_refusal": False,
            "no_output_ok": True,
            "semantic_ok": True,
            "opaque_ok": True,
            "exact_inverse_ok": True,
            "alloc_balance_ok": True,
            "alloc_invalid": False,
            "phases": {field: 0 for field in VERIFY.PHASE_FIELDS},
            "input_metrics": metrics(),
            "output_metrics": metrics(),
            "input_sha256": input_sha256,
            "output_sha256": "a" * 64,
            "input_member_digest": "b" * 64,
            "output_member_digest": "b" * 64,
            "error": None,
            "source_readback_physical_ok": None,
            "source_readback_metadata_ok": None,
            "cap_refusal": None,
            "cap_commit_refusal": None,
            "cap_source_metrics": None,
            "cap_projected_metrics": None,
            "cap_existing": None,
            "cap_exact_fit_ok": None,
            "cap_under_refused_ok": None,
        }
    )
    return sample


def read_limit_error(resource: str) -> dict[str, object]:
    return {
        "class": "DocxError::Opc::ReadLimit",
        "variant": "DocxError::Opc::ReadLimit",
        "message": f"read limit exceeded for {resource}",
        "typed_match": True,
        "resource": resource,
        "actual": 2,
        "maximum": 1,
    }


def cap_success_sample() -> dict[str, object]:
    resource = VERIFY.CAP_RESOURCES["cap_parts"]
    refusal = read_limit_error(resource)
    sample = native_success_sample("a" * 64)
    sample.update(
        {
            "output_sha256": "b" * 64,
            "output_metrics": metrics(2),
            "source_readback_physical_ok": True,
            "source_readback_metadata_ok": True,
            "cap_exact_fit_ok": True,
            "cap_under_refused_ok": True,
            "cap_refusal": refusal,
            "cap_commit_refusal": deepcopy(refusal),
            "cap_source_metrics": metrics(),
            "cap_projected_metrics": metrics(2),
            "cap_existing": {
                "applicable": True,
                "source_metrics": metrics(),
                "projected_metrics": metrics(2),
                "exact_fit_ok": True,
                "exact_opaque_ok": True,
                "source_unrelated_member_digest": "c" * 64,
                "exact_unrelated_member_digest": "c" * 64,
                "under_refused_ok": True,
                "commit_stage_checked": True,
                "refusal": deepcopy(refusal),
                "commit_refusal": deepcopy(refusal),
            },
        }
    )
    return sample


def write_cap_receipt(results: Path, sample: dict[str, object]) -> None:
    lane = "cap_parts"
    receipt = {
        "schema": VERIFY.RECEIPT_SCHEMA,
        "source_commit": VERIFY.SOURCE_COMMIT,
        "opc_source_label": "styles-effects-source-committed",
        "lane": lane,
        "expected_success": True,
        "warmup": 0,
        "sample_count": 1,
        "source_backed_api": True,
        "fixture": "ComplexNumberedLists.docx:main-effects-absent",
        "fixture_expected_package_sha256": None,
        "fixture_native": False,
        "fixture_signed": False,
        "fixture_main_present": False,
        "fixture_glossary_present": False,
        "samples": [sample],
    }
    (results / f"smoke-{lane}-p1.json").write_text(json.dumps(receipt) + "\n")
    (results / f"smoke-{lane}-p1.time.txt").write_text(
        "Exit status: 0\nMaximum resident set size (kbytes): 1\n"
    )
    (results / f"smoke-{lane}-p1.stderr.log").write_bytes(b"")


def write_native_receipt(results: Path, input_sha256: str) -> None:
    lane = "native_capture_bug_main"
    expected = VERIFY.NATIVE_FIXTURES["Bug54849.docx"]["sha256"]
    receipt = {
        "schema": VERIFY.RECEIPT_SCHEMA,
        "source_commit": VERIFY.SOURCE_COMMIT,
        "opc_source_label": "styles-effects-source-committed",
        "lane": lane,
        "expected_success": True,
        "warmup": 0,
        "sample_count": 1,
        "source_backed_api": True,
        "fixture": "Bug54849.docx",
        "fixture_expected_package_sha256": expected,
        "fixture_native": True,
        "fixture_signed": False,
        "fixture_main_present": True,
        "fixture_glossary_present": True,
        "samples": [native_success_sample(input_sha256)],
    }
    (results / f"smoke-{lane}-p1.json").write_text(json.dumps(receipt) + "\n")
    (results / f"smoke-{lane}-p1.time.txt").write_text(
        "Exit status: 0\nMaximum resident set size (kbytes): 1\n"
    )
    (results / f"smoke-{lane}-p1.stderr.log").write_bytes(b"")


class NativeReceiptTests(unittest.TestCase):
    def test_native_input_digest_mismatch_is_rejected(self):
        temporary = TemporaryDirectory(prefix="styles-effects-native-receipt-")
        self.addCleanup(temporary.cleanup)
        results = Path(temporary.name)
        expected = VERIFY.NATIVE_FIXTURES["Bug54849.docx"]["sha256"]
        write_native_receipt(results, expected)
        self.assertEqual(VERIFY.verify_receipts(results, ["native_capture_bug_main"]), {"lanes": 1, "samples": 1})

        receipt_path = results / "smoke-native_capture_bug_main-p1.json"
        receipt = json.loads(receipt_path.read_text())
        receipt["samples"][0]["input_sha256"] = "0" * 64
        receipt_path.write_text(json.dumps(receipt) + "\n")

        with self.assertRaisesRegex(AssertionError, "native input package hash changed"):
            VERIFY.verify_receipts(results, ["native_capture_bug_main"])

    def test_native_fixture_copy_digest_mismatch_is_rejected(self):
        temporary = TemporaryDirectory(prefix="styles-effects-native-copy-")
        self.addCleanup(temporary.cleanup)
        evidence = Path(temporary.name) / "evidence"
        copy = evidence / "fixtures/Bug54849.docx"
        copy.parent.mkdir(parents=True)
        source = Path(__file__).resolve().parents[4] / "test-data/poi/test-data/document/Bug54849.docx"
        shutil.copyfile(source, copy)
        corpus = {
            "schema": "docx-styles-effects-corpus-v1",
            "source_commit": VERIFY.SOURCE_COMMIT,
            "fixtures": [
                {
                    "name": "Bug54849.docx",
                    "source_path": "test-data/poi/test-data/document/Bug54849.docx",
                    "copy_path": "fixtures/Bug54849.docx",
                    "bytes": source.stat().st_size,
                    "sha256": VERIFY.sha256(source),
                    "members": [
                        {
                            "name": "word/stylesWithEffects.xml",
                            "bytes": 19883,
                            "sha256": "799de1f7a4ce43f0ca101dc750e8a8bd6e75bcb721f7744d4787dad576cda3b1",
                        }
                    ],
                }
            ],
            "lanes": ["native_capture_bug_main"],
            "expected_refusals": sorted(VERIFY.EXPECTED_REFUSALS),
        }
        corpus_path = Path(temporary.name) / "corpus.json"
        corpus_path.write_text(json.dumps(corpus))
        VERIFY.verify_corpus(corpus_path, Path(__file__).resolve().parents[4], evidence)

        changed = bytearray(copy.read_bytes())
        changed[-1] ^= 1
        copy.write_bytes(changed)
        with self.assertRaisesRegex(AssertionError, "fixture hash changed"):
            VERIFY.verify_corpus(corpus_path, Path(__file__).resolve().parents[4], evidence)


class CapOpaqueReceiptTests(unittest.TestCase):
    def test_exact_cap_opaque_evidence_is_required_for_both_cap_paths(self):
        temporary = TemporaryDirectory(prefix="styles-effects-cap-opaque-")
        self.addCleanup(temporary.cleanup)
        results = Path(temporary.name)
        baseline = cap_success_sample()
        write_cap_receipt(results, baseline)
        self.assertEqual(VERIFY.verify_receipts(results, ["cap_parts"]), {"lanes": 1, "samples": 1})

        cases = (
            (
                "source-less exact path false",
                lambda sample: sample.__setitem__("opaque_ok", False),
                "opaque gate failed",
            ),
            (
                "source-less exact path omitted",
                lambda sample: sample.pop("opaque_ok"),
                "missing opaque_ok",
            ),
            (
                "existing-owner exact path false",
                lambda sample: sample["cap_existing"].__setitem__("exact_opaque_ok", False),
                "existing-owner exact-fit opaque members changed",
            ),
            (
                "existing-owner exact path omitted",
                lambda sample: sample["cap_existing"].pop("exact_opaque_ok"),
                "cap schema changed",
            ),
        )
        for label, mutate, message in cases:
            with self.subTest(case=label):
                receipt_path = results / "smoke-cap_parts-p1.json"
                receipt = json.loads(receipt_path.read_text())
                sample = receipt["samples"][0]
                mutate(sample)
                receipt_path.write_text(json.dumps(receipt) + "\n")
                with self.assertRaisesRegex(AssertionError, message):
                    VERIFY.verify_receipts(results, ["cap_parts"])
                write_cap_receipt(results, deepcopy(baseline))


if __name__ == "__main__":
    unittest.main()
