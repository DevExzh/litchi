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
    consumer_manifest: bytes | None = None,
    probe: bytes | None = None,
    lock: bytes | None = None,
    production_marker: str = "candidate",
    baseline: bool = False,
) -> str:
    root.mkdir(parents=True)
    git(root, "init", "-q", "--initial-branch=main")
    git(root, "config", "user.email", "odf-raw-prep@example.invalid")
    git(root, "config", "user.name", "ODF raw preparation test")
    (root / "Cargo.toml").write_text("[workspace]\nmembers=[]\n", encoding="utf-8")
    production = root / "crates/litchi-ops/Cargo.toml"
    production.parent.mkdir(parents=True, exist_ok=True)
    production.write_text(
        "[package]\nname = \"litchi-ops\"\nversion = \"0.0.0\"\n"
        f"description = \"{production_marker}\"\n",
        encoding="utf-8",
    )
    perf = root / prep.CONSUMER_MANIFEST
    perf.parent.mkdir(parents=True, exist_ok=True)
    examples = root / prep.PROBE_PATH.parent
    examples.mkdir(parents=True, exist_ok=True)
    if baseline:
        perf.write_text('[package]\nname = "litchi-perf-baseline"\nversion = "0.0.1"\n', encoding="utf-8")
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


def candidate_identity(candidate: Path, *, p1_commit: str, p1_tree: str, pin_sha256: str) -> dict:
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
            "role": "protocol-P1",
            "commit": p1_commit,
            "tree": p1_tree,
            "p1_pin_sha256": pin_sha256,
            "candidate_sha256": prep._candidate_identity_digest(candidate_object),
        },
        "candidate": candidate_object,
    }


class OdfRawPreparationTests(unittest.TestCase):
    def setUp(self) -> None:
        self._old_baseline = prep.BASELINE_COMMIT

    def tearDown(self) -> None:
        prep.BASELINE_COMMIT = self._old_baseline

    def fixture(self, root: Path) -> tuple[Path, Path, Path, Path, Path, Path, Path, dict]:
        baseline = root / "baseline"
        candidate = root / "candidate"
        protocol_p = root / "protocol-p"
        adapter = root / "adapter"
        candidate_ref = repository(candidate, consumer=True, production_marker="integrated-final-c")
        # P1 is derived from frozen C; P2 is the later binding endpoint.
        subprocess.run(["git", "clone", "-q", str(candidate), str(protocol_p)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        git(protocol_p, "config", "user.email", "odf-raw-prep@example.invalid")
        git(protocol_p, "config", "user.name", "ODF raw preparation test")
        (protocol_p / "p1-source-seal.txt").write_text("P1 source seal\n", encoding="utf-8")
        git(protocol_p, "add", "p1-source-seal.txt")
        git(protocol_p, "commit", "-qm", "seal P1 from frozen C")
        p1_ref = git(protocol_p, "rev-parse", "HEAD")
        p1_tree = git(protocol_p, "rev-parse", "HEAD^{tree}")
        (protocol_p / "p2-binding-install.txt").write_text("P2 binding\n", encoding="utf-8")
        git(protocol_p, "add", "p2-binding-install.txt")
        git(protocol_p, "commit", "-qm", "install P2 binding")
        self.assertNotEqual(p1_ref, git(protocol_p, "rev-parse", "HEAD"))
        baseline_ref = repository(baseline, consumer=False, baseline=True, production_marker="baseline")
        repository(adapter, consumer=True, production_marker="adapter-3401")
        external_lock = root / "baseline.lock"
        external_lock.write_bytes(b"baseline lock\n")
        pin = {
            "schema": prep.PROTOCOL_P_PIN_SCHEMA,
            "root": str(protocol_p.resolve()),
            "source_candidate": {"commit": candidate_ref, "tree": git(candidate, "rev-parse", "HEAD^{tree}")},
            "p1": {"commit": p1_ref, "tree": p1_tree},
        }
        pin_path = root / "protocol-p1-pin.json"
        pin_path.write_text(json.dumps(pin, indent=2) + "\n", encoding="utf-8")
        identity = candidate_identity(candidate, p1_commit=p1_ref, p1_tree=p1_tree, pin_sha256=sha(pin_path))
        identity_path = root / "candidate-identity.json"
        identity_path.write_text(json.dumps(identity, indent=2) + "\n", encoding="utf-8")
        prep.BASELINE_COMMIT = baseline_ref
        return baseline, candidate, adapter, protocol_p, pin_path, external_lock, identity_path, identity

    def run_prepare(self, root: Path) -> tuple[dict, Path, Path, Path, Path, Path, Path, Path]:
        baseline, candidate, adapter, protocol_p, pin_path, external_lock, identity_path, _ = self.fixture(root)
        output = root / "preparation.json"
        value = prep.prepare(
            baseline_root=baseline,
            candidate_root=candidate,
            adapter_root=adapter,
            protocol_p_root=protocol_p,
            protocol_p_pin_path=pin_path,
            candidate_identity_path=identity_path,
            baseline_lock=external_lock,
            output=output,
        )
        return value, baseline, candidate, adapter, protocol_p, pin_path, identity_path, output

    def test_integrated_candidate_is_clean_and_candidate_is_never_written(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-integrated-") as directory:
            root = Path(directory)
            value, baseline, candidate, adapter, protocol_p, pin_path, identity_path, output = self.run_prepare(root)
            self.assertEqual(value["candidate"]["mode"], prep.CANDIDATE_MODE)
            self.assertEqual(value["candidate"]["overlay_paths"], [])
            self.assertEqual(value["candidate"]["before"], value["candidate"]["after"])
            self.assertEqual(value["candidate_identity"]["path"], str(identity_path.resolve()))
            self.assertEqual(value["protocol_p_pin"]["path"], str(pin_path.resolve()))
            self.assertEqual(value["protocol_p_anchor"]["source_candidate_commit"], git(candidate, "rev-parse", "HEAD"))
            self.assertEqual(value["protocol_p_anchor"]["current_commit"], git(protocol_p, "rev-parse", "HEAD"))
            self.assertNotEqual(value["protocol_p_anchor"]["p1_commit"], value["protocol_p_anchor"]["current_commit"])
            self.assertEqual(value["adapter"]["consumer_provenance_scope"], "three consumer files only; production closure is not equated to final C")
            self.assertNotEqual(prep._production_source_tree(candidate), prep._production_source_tree(adapter))
            self.assertEqual((candidate / prep.PROBE_PATH).read_bytes(), (adapter / prep.PROBE_PATH).read_bytes())
            self.assertEqual(sorted(prep._status_paths(value["baseline"]["after_overlay"]["status"])), sorted([str(prep.CONSUMER_MANIFEST), str(prep.LOCK_PATH), str(prep.PROBE_PATH)]))
            self.assertTrue(output.is_file())

    def test_six_admission_refusals(self) -> None:
        cases = ("old-lock", "missing-probe", "stale-c", "unexpected", "symlink", "proof-drift")
        for case in cases:
            with self.subTest(case=case), tempfile.TemporaryDirectory(prefix=f"odf-raw-{case}-") as directory:
                root = Path(directory)
                baseline, candidate, adapter, protocol_p, pin_path, external_lock, identity_path, identity = self.fixture(root)
                mutated = copy.deepcopy(identity)
                if case == "old-lock":
                    old = b"old pre-final-C lock\n"
                    (adapter / prep.LOCK_PATH).write_bytes(old)
                    git(adapter, "add", str(prep.LOCK_PATH))
                    git(adapter, "commit", "-qm", "old lock")
                    mutated["candidate"]["perf_lock"] = {"path": str(prep.LOCK_PATH), "sha256": sha(adapter / prep.LOCK_PATH), "bytes": len(old)}
                    mutated["sealed_by"]["candidate_sha256"] = prep._candidate_identity_digest(mutated["candidate"])
                elif case == "missing-probe":
                    (candidate / prep.PROBE_PATH).unlink()
                elif case == "stale-c":
                    mutated["candidate"]["commit"] = "c" * 40
                    mutated["sealed_by"]["candidate_sha256"] = prep._candidate_identity_digest(mutated["candidate"])
                elif case == "unexpected":
                    (candidate / "unexpected.txt").write_text("dirty\n", encoding="utf-8")
                elif case == "symlink":
                    probe = adapter / prep.PROBE_PATH
                    probe.unlink()
                    probe.symlink_to(adapter / prep.LOCK_PATH)
                elif case == "proof-drift":
                    mutated["sealed_by"]["candidate_sha256"] = "0" * 64
                identity_path.write_text(json.dumps(mutated, indent=2) + "\n", encoding="utf-8")
                with self.assertRaises(prep.PreparationError):
                    prep.prepare(baseline_root=baseline, candidate_root=candidate, adapter_root=adapter, protocol_p_root=protocol_p, protocol_p_pin_path=pin_path, candidate_identity_path=identity_path, baseline_lock=external_lock, output=root / f"{case}.json")
                self.assertFalse((baseline / prep.PROBE_PATH).exists(), "refusal must precede baseline writes")

    def test_rejects_mismatched_p1_envelope_root_tree_and_ancestry(self) -> None:
        for field in ("root", "p1", "source_candidate", "ancestry"):
            with self.subTest(field=field), tempfile.TemporaryDirectory(prefix=f"odf-raw-p1-{field}-") as directory:
                root = Path(directory)
                baseline, candidate, adapter, protocol_p, pin_path, external_lock, identity_path, identity = self.fixture(root)
                pin = json.loads(pin_path.read_text(encoding="utf-8"))
                if field == "root":
                    pin["root"] = str(root / "other-p")
                elif field == "p1":
                    pin["p1"]["tree"] = "b" * 40
                elif field == "source_candidate":
                    pin["source_candidate"]["commit"] = "a" * 40
                else:
                    unrelated = git(baseline, "rev-parse", "HEAD")
                    git(protocol_p, "fetch", "-q", str(baseline), "main")
                    pin["p1"] = {"commit": unrelated, "tree": git(baseline, "rev-parse", "HEAD^{tree}")}
                    identity["sealed_by"]["commit"] = unrelated
                    identity["sealed_by"]["tree"] = pin["p1"]["tree"]
                pin_path.write_text(json.dumps(pin, indent=2) + "\n", encoding="utf-8")
                if field == "ancestry":
                    identity["sealed_by"]["p1_pin_sha256"] = sha(pin_path)
                    identity_path.write_text(json.dumps(identity, indent=2) + "\n", encoding="utf-8")
                with self.assertRaises(prep.PreparationError):
                    prep.prepare(baseline_root=baseline, candidate_root=candidate, adapter_root=adapter, protocol_p_root=protocol_p, protocol_p_pin_path=pin_path, candidate_identity_path=identity_path, baseline_lock=external_lock, output=root / f"{field}.json")

    def test_rejects_assume_unchanged_production_edit_at_prepare(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-assume-unchanged-") as directory:
            root = Path(directory)
            baseline, candidate, adapter, protocol_p, pin_path, external_lock, identity_path, _ = self.fixture(root)
            production = candidate / "crates/litchi-ops/Cargo.toml"
            production.write_text(production.read_text(encoding="utf-8") + "# hidden edit\n", encoding="utf-8")
            git(candidate, "update-index", "--assume-unchanged", str(production.relative_to(candidate)))
            with self.assertRaises(prep.PreparationError):
                prep.prepare(baseline_root=baseline, candidate_root=candidate, adapter_root=adapter, protocol_p_root=protocol_p, protocol_p_pin_path=pin_path, candidate_identity_path=identity_path, baseline_lock=external_lock, output=root / "assume-unchanged.json")

    def test_rejects_p1_seal_fields_that_do_not_equal_the_pin(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-p1-seal-") as directory:
            root = Path(directory)
            baseline, candidate, adapter, protocol_p, pin_path, external_lock, identity_path, identity = self.fixture(root)
            identity["sealed_by"]["commit"] = "f" * 40
            identity["sealed_by"]["tree"] = "e" * 40
            identity_path.write_text(json.dumps(identity, indent=2) + "\n", encoding="utf-8")
            with self.assertRaises(prep.PreparationError):
                prep.prepare(baseline_root=baseline, candidate_root=candidate, adapter_root=adapter, protocol_p_root=protocol_p, protocol_p_pin_path=pin_path, candidate_identity_path=identity_path, baseline_lock=external_lock, output=root / "bad-seal.json")

    def test_public_prepare_to_runner_positive_and_identity_rejection(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-cli-composition-") as directory:
            root = Path(directory)
            baseline, candidate, adapter, protocol_p, pin_path, external_lock, identity_path, identity = self.fixture(root)
            preparation_path = root / "preparation.json"
            self.assertEqual(prep.main(["--baseline-root", str(baseline), "--candidate-root", str(candidate), "--adapter-root", str(adapter), "--protocol-p-root", str(protocol_p), "--protocol-p-pin", str(pin_path), "--candidate-identity", str(identity_path), "--baseline-lock", str(external_lock), "--output", str(preparation_path)]), 0)
            value = json.loads(preparation_path.read_text(encoding="utf-8"))
            manifest = json.loads((SOURCE_ROOT / "tools/native_qualification/manifests/odf-raw-codec.json").read_text(encoding="utf-8"))
            manifest["baseline_commit"] = git(baseline, "rev-parse", "HEAD")
            manifest["source"]["candidate_consumer_manifest_sha256"] = value["candidate"]["consumer_manifest"]["sha256"]
            manifest["source"]["candidate_consumer_manifest_bytes"] = value["candidate"]["consumer_manifest"]["bytes"]
            manifest["source"]["candidate_lock_sha256"] = value["candidate"]["lock"]["sha256"]
            manifest["source"]["candidate_lock_bytes"] = value["candidate"]["lock"]["bytes"]
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
                args = ["--manifest", str(manifest_path), "--preparation", str(preparation_path), "--baseline-root", str(baseline), "--candidate-root", str(candidate), "--adapter-root", str(adapter), "--protocol-p-root", str(protocol_p), "--protocol-p-pin", str(pin_path), "--candidate-identity", str(identity_path), "--output", str(root / "runner-output"), "--target-root", str(root / "target"), "--tmp-root", str(root / "tmp"), "--cpu", "0", "--phase", "build"]
                self.assertEqual(runner.main(args), 0)
                self.assertEqual(len(instances), 1)
                bad_seal = copy.deepcopy(identity)
                bad_seal["sealed_by"]["commit"] = "f" * 40
                bad_seal["sealed_by"]["tree"] = "e" * 40
                bad_seal_path = root / "bad-seal-identity.json"
                bad_seal_path.write_text(json.dumps(bad_seal, indent=2) + "\n", encoding="utf-8")
                bad_seal_args = list(args)
                bad_seal_args[bad_seal_args.index(str(identity_path))] = str(bad_seal_path)
                bad_seal_args[bad_seal_args.index(str(root / "runner-output"))] = str(root / "bad-seal-output")
                bad_seal_args[bad_seal_args.index(str(root / "target"))] = str(root / "bad-seal-target")
                bad_seal_args[bad_seal_args.index(str(root / "tmp"))] = str(root / "bad-seal-tmp")
                self.assertEqual(runner.main(bad_seal_args), 2)
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
