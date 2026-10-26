#!/usr/bin/env python3
"""Source-only tests for integrated-final-C ODF raw-codec admission."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import subprocess
import tempfile
import unittest

from tools.native_qualification import odf_raw_codec_paired_runner as runner
from tools.native_qualification import prepare_odf_raw_codec_inputs as prep

SOURCE_ROOT = Path(__file__).resolve().parents[1]


def git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.stdout.strip()


def repository(
    root: Path,
    *,
    consumer: bool,
    protocol_root: Path | None = None,
    consumer_manifest: bytes | None = None,
    probe: bytes | None = None,
    lock: bytes | None = None,
    production_marker: str = "candidate",
    baseline: bool = False,
) -> str:
    if protocol_root is None:
        root.mkdir(parents=True)
        git(root, "init", "-q", "--initial-branch=main")
    else:
        subprocess.run(["git", "clone", "-q", str(protocol_root), str(root)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    git(root, "config", "user.email", "odf-raw-prep@example.invalid")
    git(root, "config", "user.name", "ODF raw preparation test")
    if protocol_root is None:
        (root / "Cargo.toml").write_text("[workspace]\nmembers=[]\n", encoding="utf-8")
        production = root / "crates/litchi-ops/Cargo.toml"
        production.parent.mkdir(parents=True, exist_ok=True)
        production.write_text(
            "[package]\nname = \"litchi-ops\"\nversion = \"0.0.0\"\n"
            f"description = \"{production_marker}\"\n",
            encoding="utf-8",
        )
    elif production_marker != "protocol-p":
        (root / "crates/litchi-ops/Cargo.toml").write_text(
            "[package]\nname = \"litchi-ops\"\nversion = \"0.0.0\"\n"
            f"description = \"{production_marker}\"\n",
            encoding="utf-8",
        )
    perf = root / prep.CONSUMER_MANIFEST
    perf.parent.mkdir(parents=True, exist_ok=True)
    examples = root / prep.PROBE_PATH.parent
    examples.mkdir(parents=True, exist_ok=True)
    if baseline:
        # A tracked pre-existing consumer surface makes the overlay status
        # precise (manifest/lock edits plus the new probe file).
        perf.write_text(
            '[package]\nname = "litchi-perf-baseline"\nversion = "0.0.1"\n',
            encoding="utf-8",
        )
        (root / prep.LOCK_PATH).write_bytes(b"baseline old lock\n")
        (examples / ".keep").write_text("\n", encoding="utf-8")
    elif consumer:
        if consumer_manifest is None:
            consumer_manifest = (SOURCE_ROOT / prep.CONSUMER_MANIFEST).read_bytes()
        if probe is None:
            probe = (SOURCE_ROOT / prep.PROBE_PATH).read_bytes()
        if lock is None:
            lock = (SOURCE_ROOT / prep.LOCK_PATH).read_bytes()
        perf.write_bytes(consumer_manifest)
        (root / prep.LOCK_PATH).write_bytes(lock)
        (root / prep.PROBE_PATH).write_bytes(probe)
    git(root, "add", ".")
    git(root, "commit", "-qm", "fixture")
    return git(root, "rev-parse", "HEAD")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def candidate_identity(candidate: Path, protocol_root: Path) -> dict:
    p_commit = git(protocol_root, "rev-parse", "HEAD")
    p_tree = git(protocol_root, "rev-parse", "HEAD^{tree}")
    candidate_object = {
        "mode": prep.CANDIDATE_MODE,
        "commit": git(candidate, "rev-parse", "HEAD"),
        "tree": git(candidate, "rev-parse", "HEAD^{tree}"),
        "production_source_tree_sha256": prep._production_source_tree(candidate),
        "consumer_manifest": {
            "path": str(prep.CONSUMER_MANIFEST),
            "sha256": sha(candidate / prep.CONSUMER_MANIFEST),
            "bytes": (candidate / prep.CONSUMER_MANIFEST).stat().st_size,
        },
        "probe": {
            "path": str(prep.PROBE_PATH),
            "sha256": sha(candidate / prep.PROBE_PATH),
            "bytes": (candidate / prep.PROBE_PATH).stat().st_size,
        },
        "perf_lock": {
            "path": str(prep.LOCK_PATH),
            "sha256": sha(candidate / prep.LOCK_PATH),
            "bytes": (candidate / prep.LOCK_PATH).stat().st_size,
        },
    }
    return {
        "schema": prep.IDENTITY_SCHEMA,
        "sealed_by": {
            "role": "protocol-P",
            "commit": p_commit,
            "tree": p_tree,
            "candidate_sha256": prep._candidate_identity_digest(candidate_object),
        },
        "candidate": candidate_object,
    }


class OdfRawPreparationTests(unittest.TestCase):
    def setUp(self) -> None:
        self._old_baseline = prep.BASELINE_COMMIT
        self._old_prep_p_commit = prep.PROTOCOL_P_COMMIT
        self._old_prep_p_tree = prep.PROTOCOL_P_TREE
        self._old_runner_p_commit = runner.PROTOCOL_P_COMMIT
        self._old_runner_p_tree = runner.PROTOCOL_P_TREE

    def tearDown(self) -> None:
        prep.BASELINE_COMMIT = self._old_baseline
        prep.PROTOCOL_P_COMMIT = self._old_prep_p_commit
        prep.PROTOCOL_P_TREE = self._old_prep_p_tree
        runner.PROTOCOL_P_COMMIT = self._old_runner_p_commit
        runner.PROTOCOL_P_TREE = self._old_runner_p_tree

    def fixture(self, root: Path) -> tuple[Path, Path, Path, Path, Path, Path, dict]:
        baseline = root / "baseline"
        protocol_p = root / "protocol-p"
        candidate = root / "candidate"
        adapter = root / "adapter"
        p_ref = repository(protocol_p, consumer=False, production_marker="protocol-p")
        baseline_ref = repository(baseline, consumer=False, baseline=True, production_marker="baseline")
        repository(candidate, consumer=True, protocol_root=protocol_p, production_marker="integrated-final-c")
        repository(adapter, consumer=True, production_marker="adapter-3401")
        external_lock = root / "baseline.lock"
        external_lock.write_bytes(b"baseline lock\n")
        p_tree = git(protocol_p, "rev-parse", "HEAD^{tree}")
        # The synthetic fixture uses a real Git P commit/tree and shares the
        # candidate commit object with the P repository for ancestry checks.
        git(protocol_p, "fetch", "-q", str(candidate), git(candidate, "rev-parse", "HEAD"))
        prep.PROTOCOL_P_COMMIT = p_ref
        prep.PROTOCOL_P_TREE = p_tree
        runner.PROTOCOL_P_COMMIT = p_ref
        runner.PROTOCOL_P_TREE = p_tree
        identity = candidate_identity(candidate, protocol_p)
        identity_path = root / "candidate-identity.json"
        identity_path.write_text(json.dumps(identity, indent=2) + "\n", encoding="utf-8")
        prep.BASELINE_COMMIT = baseline_ref
        return baseline, candidate, adapter, protocol_p, external_lock, identity_path, identity

    def run_prepare(self, root: Path) -> tuple[dict, Path, Path, Path, Path, Path, Path]:
        baseline, candidate, adapter, protocol_p, external_lock, identity_path, identity = self.fixture(root)
        output = root / "preparation.json"
        value = prep.prepare(
            baseline_root=baseline,
            candidate_root=candidate,
            adapter_root=adapter,
            protocol_p_root=protocol_p,
            candidate_identity_path=identity_path,
            baseline_lock=external_lock,
            output=output,
        )
        return value, baseline, candidate, adapter, protocol_p, identity_path, output

    def test_integrated_candidate_is_clean_and_candidate_is_never_written(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-integrated-") as directory:
            root = Path(directory)
            value, baseline, candidate, adapter, protocol_p, identity_path, output = self.run_prepare(root)
            self.assertEqual(value["candidate"]["mode"], prep.CANDIDATE_MODE)
            self.assertEqual(value["candidate"]["overlay_paths"], [])
            self.assertEqual(value["candidate"]["before"], value["candidate"]["after"])
            self.assertEqual(value["candidate_identity"]["path"], str(identity_path.resolve()))
            self.assertEqual(value["adapter"]["consumer_provenance_scope"], "three consumer files only; production closure is not equated to final C")
            self.assertNotEqual(
                prep._production_source_tree(candidate),
                prep._production_source_tree(adapter),
                "adapter provenance must not require the whole production closure to match final C",
            )
            self.assertEqual((candidate / prep.PROBE_PATH).read_bytes(), (adapter / prep.PROBE_PATH).read_bytes())
            self.assertEqual(
                sorted(prep._status_paths(value["baseline"]["after_overlay"]["status"])),
                sorted([str(prep.CONSUMER_MANIFEST), str(prep.LOCK_PATH), str(prep.PROBE_PATH)]),
            )
            self.assertEqual(value["candidate"]["consumer_manifest"]["source_path"], str(candidate / prep.CONSUMER_MANIFEST))
            self.assertTrue(output.is_file())

    def test_six_admission_refusals(self) -> None:
        # old lock, missing probe, stale C, unexpected changes, symlink, and
        # proof drift are deliberately separate fail-closed admissions.
        cases = ("old-lock", "missing-probe", "stale-c", "unexpected", "symlink", "proof-drift")
        for case in cases:
            with self.subTest(case=case), tempfile.TemporaryDirectory(prefix=f"odf-raw-{case}-") as directory:
                root = Path(directory)
                baseline, candidate, adapter, protocol_p, external_lock, identity_path, identity = self.fixture(root)
                mutated = copy.deepcopy(identity)
                if case == "old-lock":
                    old = b"old pre-final-C lock\n"
                    (adapter / prep.LOCK_PATH).write_bytes(old)
                    git(adapter, "add", str(prep.LOCK_PATH))
                    git(adapter, "commit", "-qm", "old lock")
                    mutated["candidate"]["perf_lock"] = {
                        "path": str(prep.LOCK_PATH), "sha256": sha(adapter / prep.LOCK_PATH), "bytes": len(old)
                    }
                elif case == "missing-probe":
                    (candidate / prep.PROBE_PATH).unlink()
                elif case == "stale-c":
                    mutated["candidate"]["commit"] = "c" * 40
                elif case == "unexpected":
                    (candidate / "unexpected.txt").write_text("dirty\n", encoding="utf-8")
                elif case == "symlink":
                    probe = adapter / prep.PROBE_PATH
                    probe.unlink()
                    probe.symlink_to(adapter / prep.LOCK_PATH)
                elif case == "proof-drift":
                    mutated["sealed_by"]["candidate_sha256"] = "0" * 64
                if case in {"old-lock", "stale-c"}:
                    mutated["sealed_by"]["candidate_sha256"] = prep._candidate_identity_digest(mutated["candidate"])
                identity_path.write_text(json.dumps(mutated, indent=2) + "\n", encoding="utf-8")
                with self.assertRaises(prep.PreparationError):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        adapter_root=adapter,
                        protocol_p_root=protocol_p,
                        candidate_identity_path=identity_path,
                        baseline_lock=external_lock,
                        output=root / f"{case}.json",
                    )
                self.assertFalse((baseline / prep.PROBE_PATH).exists(), "refusal must precede baseline writes")

    def test_rejects_fake_protocol_p_object_and_ignored_adapter_inputs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-trust-guards-") as directory:
            root = Path(directory)
            baseline, candidate, adapter, protocol_p, external_lock, identity_path, identity = self.fixture(root)
            fake = copy.deepcopy(identity)
            fake["sealed_by"]["commit"] = "a" * 40
            fake["sealed_by"]["tree"] = "b" * 40
            identity_path.write_text(json.dumps(fake, indent=2) + "\n", encoding="utf-8")
            with self.assertRaises(prep.PreparationError):
                prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    adapter_root=adapter,
                    protocol_p_root=protocol_p,
                    candidate_identity_path=identity_path,
                    baseline_lock=external_lock,
                    output=root / "fake-p.json",
                )

            # Leave the identity valid, then remove the adapter files from its
            # index while retaining identical ignored bytes on disk.
            identity_path.write_text(json.dumps(identity, indent=2) + "\n", encoding="utf-8")
            for relative in (prep.CONSUMER_MANIFEST, prep.LOCK_PATH, prep.PROBE_PATH):
                git(adapter, "rm", "-q", "--cached", str(relative))
            (adapter / ".gitignore").write_text(
                "tools/perf-baseline/Cargo.toml\n"
                "tools/perf-baseline/Cargo.lock\n"
                "tools/perf-baseline/examples/odf_raw_codec_qualification.rs\n",
                encoding="utf-8",
            )
            git(adapter, "add", ".gitignore")
            git(adapter, "commit", "-qm", "hide adapter files")
            with self.assertRaises(prep.PreparationError):
                prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    adapter_root=adapter,
                    protocol_p_root=protocol_p,
                    candidate_identity_path=identity_path,
                    baseline_lock=external_lock,
                    output=root / "ignored-adapter.json",
                )

    def test_rejects_assume_unchanged_production_edit_at_prepare(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-assume-unchanged-") as directory:
            root = Path(directory)
            baseline, candidate, adapter, protocol_p, external_lock, identity_path, _ = self.fixture(root)
            production = candidate / "crates/litchi-ops/Cargo.toml"
            production.write_text(production.read_text(encoding="utf-8") + "# hidden edit\n", encoding="utf-8")
            git(candidate, "update-index", "--assume-unchanged", str(production.relative_to(candidate)))
            with self.assertRaises(prep.PreparationError):
                prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    adapter_root=adapter,
                    protocol_p_root=protocol_p,
                    candidate_identity_path=identity_path,
                    baseline_lock=external_lock,
                    output=root / "assume-unchanged.json",
                )

    def test_public_prepare_to_runner_positive_and_identity_rejection(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-cli-composition-") as directory:
            root = Path(directory)
            baseline, candidate, adapter, protocol_p, external_lock, identity_path, identity = self.fixture(root)
            preparation_path = root / "preparation.json"
            self.assertEqual(
                prep.main(
                    [
                        "--baseline-root", str(baseline),
                        "--candidate-root", str(candidate),
                        "--adapter-root", str(adapter),
                        "--protocol-p-root", str(protocol_p),
                        "--candidate-identity", str(identity_path),
                        "--baseline-lock", str(external_lock),
                        "--output", str(preparation_path),
                    ]
                ),
                0,
            )
            value = json.loads(preparation_path.read_text(encoding="utf-8"))
            manifest = json.loads((Path(__file__).resolve().parents[1] / "tools/native_qualification/manifests/odf-raw-codec.json").read_text(encoding="utf-8"))
            manifest["baseline_commit"] = git(baseline, "rev-parse", "HEAD")
            manifest["source"]["candidate_consumer_manifest_sha256"] = value["candidate"]["consumer_manifest"]["sha256"]
            manifest["source"]["candidate_consumer_manifest_bytes"] = value["candidate"]["consumer_manifest"]["bytes"]
            manifest["source"]["candidate_lock_sha256"] = value["candidate"]["lock"]["sha256"]
            manifest["source"]["candidate_lock_bytes"] = value["candidate"]["lock"]["bytes"]
            manifest["source"]["protocol_p_commit"] = identity["sealed_by"]["commit"]
            manifest["source"]["protocol_p_tree"] = identity["sealed_by"]["tree"]
            manifest["probe"]["sha256"] = value["candidate"]["probe"]["sha256"]
            manifest["probe"]["bytes"] = value["candidate"]["probe"]["bytes"]
            manifest_path = root / "manifest.json"
            manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")

            old_build = runner.PairedRunner.build
            instances: list[runner.PairedRunner] = []

            def admit_endpoints(instance: runner.PairedRunner) -> None:
                instances.append(instance)
                instance._endpoint("baseline", full=True)
                instance._endpoint("candidate", full=True)

            runner.PairedRunner.build = admit_endpoints  # type: ignore[method-assign]
            try:
                args = [
                    "--manifest", str(manifest_path),
                    "--preparation", str(preparation_path),
                    "--baseline-root", str(baseline),
                    "--candidate-root", str(candidate),
                    "--adapter-root", str(adapter),
                    "--protocol-p-root", str(protocol_p),
                    "--candidate-identity", str(identity_path),
                    "--output", str(root / "runner-output"),
                    "--target-root", str(root / "target"),
                    "--tmp-root", str(root / "tmp"),
                    "--cpu", "0",
                    "--phase", "build",
                ]
                self.assertEqual(runner.main(args), 0)
                self.assertEqual(len(instances), 1)
                drifted = json.loads(identity_path.read_text(encoding="utf-8"))
                identity_path.write_text(json.dumps(drifted, separators=(",", ":")) + "\n", encoding="utf-8")
                with self.assertRaises(runner.RunnerError):
                    instances[0]._endpoint("candidate", full=True)
                broken = json.loads(identity_path.read_text(encoding="utf-8"))
                broken["candidate"]["tree"] = "d" * 40
                broken["sealed_by"]["candidate_sha256"] = prep._candidate_identity_digest(broken["candidate"])
                broken_path = root / "broken-identity.json"
                broken_path.write_text(json.dumps(broken, indent=2) + "\n", encoding="utf-8")
                bad_args = list(args)
                bad_args[bad_args.index(str(identity_path))] = str(broken_path)
                bad_args[bad_args.index(str(root / "runner-output"))] = str(root / "bad-output")
                bad_args[bad_args.index(str(root / "target"))] = str(root / "bad-target")
                bad_args[bad_args.index(str(root / "tmp"))] = str(root / "bad-tmp")
                self.assertEqual(runner.main(bad_args), 2)
            finally:
                runner.PairedRunner.build = old_build


if __name__ == "__main__":
    unittest.main()
