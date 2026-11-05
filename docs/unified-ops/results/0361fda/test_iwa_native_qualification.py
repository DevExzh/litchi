"""Unit tests for the fail-closed IWA native qualification runner.

These tests use a disposable executable that emits probe-shaped JSON.  They
do not build a Rust crate, run a benchmark, or create evidence under the
checkout.
"""

from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path
import stat
import shutil
import subprocess
import tempfile
import unittest
from unittest import mock


HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("iwa_native_qualification", HERE / "iwa_native_qualification.py")
assert SPEC and SPEC.loader
RUNNER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(RUNNER)


def _fake_probe(
    path: Path, *, owner: bool, valid: bool = True, allocation: bool = False
) -> None:
    source = "a" * 64
    output = "b" * 64
    if owner:
        correctness = {
            "status": "ok",
            "replay_matches_changed": valid,
            "inverse_matches_source": valid,
            "noop_is_noop": valid,
            "noop_matches_expected": valid,
            "semantic_ok": valid,
            "changed_sha256": output,
            "replay_sha256": output,
            "inverse_sha256": source,
            "noop_sha256": source,
            "source_bytes": 12,
            "changed_bytes": 12,
            "noop_bytes": 12,
        }
        body = {
            "schema": "litchi.iwa.owner.probe.v2",
            "route": "numbers_headers",
            "iterations": 1,
            "ns": 500_000_000,
            "checksum": 11,
            "output_bytes": 12,
            "source_sha256": source,
            "output_sha256": output,
            "output_differs": True,
            "correctness": correctness,
        }
        if allocation:
            body.update(
                {
                    "peak": 8,
                    "retained": 4,
                    "allocated": 16,
                    "allocations": 2,
                }
            )
    else:
        body = {
            "schema": "litchi.iwa.constructor.probe.v3",
            "route": "field_path",
            "iterations": 1,
            "ns": 500_000_000,
            "checksum": 11,
            "output_bytes": 12,
            "source_sha256": source,
            "output_sha256": source,
            "output_matches_source": True,
            "correctness": {
                "status": "ok",
                "roundtrip": True,
                "replay": "not-applicable",
                "noop": "not-applicable",
            },
        }
    path.write_text(
        "#!/usr/bin/env python3\n"
        "import json, sys\n"
        f"print({json.dumps(json.dumps(body))})\n",
        encoding="utf-8",
    )
    path.chmod(path.stat().st_mode | stat.S_IXUSR)


class NativeQualificationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)

    def test_artifact_identity_is_content_bound(self) -> None:
        path = self.directory / "probe"
        path.write_bytes(b"one")
        identity = RUNNER.artifact(path)
        RUNNER.verify_artifact(identity)
        path.write_bytes(b"two")
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.verify_artifact(identity)

    def test_output_inside_checkout_is_rejected(self) -> None:
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.outside_checkout(RUNNER.ROOT / "docs/unified-ops/results/0361fda")

    def test_owner_fixture_manifest_is_content_bound(self) -> None:
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.fixture_identity(self.directory)
        fixture_dir = self.directory / "fixtures"
        fixture_dir.mkdir()
        provenance = RUNNER._fixture_provenance()
        sources = provenance["files"]
        for name in ("basic.numbers", "media-replacement-native.key", "tone.wav", "basic.pages"):
            source = RUNNER.ROOT / sources[name]["path"]
            shutil.copyfile(source, fixture_dir / name)
        shutil.copyfile(fixture_dir / "basic.pages", fixture_dir / "footnote-fixture.pages")
        files = {}
        for name in RUNNER.OWNER_FIXTURES:
            row = RUNNER.artifact(fixture_dir / name)
            if name == "footnote-fixture.pages":
                row["generated"] = {
                    "generator": RUNNER.artifact(RUNNER.FIXTURE_GENERATOR),
                    "sources": [
                        RUNNER._tracked_identity(source, name)
                        for source in sources[name]["sources"]
                    ],
                    "command": sources[name]["command"],
                    "method": sources[name]["method"],
                }
            else:
                row["source"] = RUNNER.artifact(RUNNER.ROOT / sources[name]["path"])
            files[name] = row
        RUNNER.save_json(
            fixture_dir / "manifest.json",
            {
                "schema": "litchi.iwa.owner-fixtures.v1",
                "provenance": RUNNER.artifact(RUNNER.FIXTURE_PROVENANCE),
                "files": files,
            },
        )
        identity = RUNNER.fixture_identity(fixture_dir)
        self.assertEqual(set(identity["files"]), set(RUNNER.OWNER_FIXTURES))
        regenerated_paths: list[Path] = []

        def replay_generator(output: Path) -> None:
            regenerated_paths.append(output)
            shutil.copyfile(fixture_dir / "footnote-fixture.pages", output)

        # The replay output is intentionally at a different temporary path;
        # fixture validation must compare content identity while retaining the
        # manifest's original path and generated provenance.
        with mock.patch.object(RUNNER, "_run_fixture_generator", side_effect=replay_generator):
            replayed = RUNNER.fixture_identity(fixture_dir, regenerate=True)
        self.assertEqual(replayed, identity)
        self.assertEqual(len(regenerated_paths), 1)
        self.assertNotEqual(regenerated_paths[0].resolve(), (fixture_dir / "footnote-fixture.pages").resolve())
        (fixture_dir / "tone.wav").write_bytes(b"changed")
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.fixture_identity(fixture_dir)

    def test_fixed_protocol_rejects_weakening_without_diagnostic(self) -> None:
        args = type(
            "Args",
            (),
            {
                "start_count": RUNNER.START_COUNTS["constructor"],
                "blocks": RUNNER.FIXED_BLOCKS - 1,
                "warmups": RUNNER.FIXED_WARMUPS,
                "target_ns": RUNNER.TARGET_NS,
                "minimum_ns": RUNNER.MINIMUM_NS,
                "growth": RUNNER.FIXED_GROWTH,
                "max_iterations": RUNNER.MAX_ITERATIONS,
                "timeout_s": RUNNER.FIXED_TIMEOUT_S,
            },
        )()
        protocol = RUNNER._protocol(args, "constructor", RUNNER.ROUTES["constructor"])
        self.assertFalse(RUNNER._fixed_protocol(protocol, "constructor"))
        args.blocks = RUNNER.FIXED_BLOCKS
        self.assertTrue(RUNNER._fixed_protocol(protocol | {"blocks": RUNNER.FIXED_BLOCKS}, "constructor"))

    def test_integrated_source_delta_records_full_nonempty_change_set(self) -> None:
        base = {
            "crates/litchi-iwa-core/same.rs": {"bytes": 5, "sha256": "a" * 64},
            "crates/litchi-opc/changed.rs": {"bytes": 7, "sha256": "b" * 64},
            "crates/litchi-pages/removed.rs": {"bytes": 8, "sha256": "c" * 64},
        }
        candidate = {
            "crates/litchi-iwa-core/same.rs": {"bytes": 5, "sha256": "a" * 64},
            "crates/litchi-opc/changed.rs": {"bytes": 9, "sha256": "d" * 64},
            "crates/litchi-pages/added.rs": {"bytes": 6, "sha256": "e" * 64},
        }
        base_commit = "1" * 40
        candidate_commit = "2" * 40
        delta = RUNNER.source_delta(base, candidate, base_commit, candidate_commit)
        self.assertEqual(delta["comparison"], "integrated-route")
        self.assertEqual(delta["base_commit"], base_commit)
        self.assertEqual(delta["candidate_commit"], candidate_commit)
        self.assertEqual(len(delta["changed"]), 1)
        self.assertEqual(len(delta["base_only"]), 1)
        self.assertEqual(len(delta["candidate_only"]), 1)
        RUNNER.verify_source_delta(
            delta, {"base": base, "candidate": candidate}, base_commit, candidate_commit
        )

    def test_integrated_source_delta_rejects_ref_and_artifact_tampering(self) -> None:
        base = {"crates/shared.rs": {"bytes": 5, "sha256": "a" * 64}}
        candidate = {"crates/shared.rs": {"bytes": 6, "sha256": "b" * 64}}
        delta = RUNNER.source_delta(base, candidate, "1" * 40, "2" * 40)
        tampered_deltas = []
        tampered = copy.deepcopy(delta)
        tampered["base_commit"] = "3" * 40
        tampered_deltas.append(tampered)
        tampered = copy.deepcopy(delta)
        tampered["changed"][0]["base"]["sha256"] = "0" * 64
        tampered_deltas.append(tampered)
        tampered = copy.deepcopy(delta)
        tampered["changed"][0]["candidate"]["bytes"] = 99
        tampered_deltas.append(tampered)
        tampered = copy.deepcopy(delta)
        tampered["changed"].append(
            {
                "path": "extra.rs",
                "base": base["crates/shared.rs"],
                "candidate": candidate["crates/shared.rs"],
            }
        )
        tampered_deltas.append(tampered)
        for tampered in tampered_deltas:
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.verify_source_delta(
                    tampered, {"base": base, "candidate": candidate}, "1" * 40, "2" * 40
                )
        tampered_maps = copy.deepcopy({"base": base, "candidate": candidate})
        tampered_maps["candidate"]["crates/shared.rs"]["sha256"] = "0" * 64
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.verify_source_delta(
                delta, tampered_maps, "1" * 40, "2" * 40
            )

    def _write_dependency_source(self, root: Path, *, candidate: bool, owner: bool) -> None:
        packages = {
            "litchi-iwa-common": "[package]\nname = \"litchi-iwa-common\"\nversion = \"0.1.0\"\n",
            "litchi-iwa-core": (
                "[package]\nname = \"litchi-iwa-core\"\nversion = \"0.1.0\"\n\n"
                "[dependencies]\nlitchi-iwa-common = { path = \"../litchi-iwa-common\" }\n"
                "snap = \"1\"\n"
                + ('smallvec = { version = "1", features = ["union"] }\n' if candidate else "")
            ),
        }
        if owner:
            for name in ("litchi-numbers", "litchi-keynote", "litchi-pages"):
                packages[name] = (
                    f"[package]\nname = \"{name}\"\nversion = \"0.1.0\"\n\n"
                    "[dependencies]\nlitchi-iwa-core = { path = \"../litchi-iwa-core\" }\n\n"
                    "[target.'cfg(windows)'.dependencies]\n"
                    "litchi-iwa-common = { path = \"../litchi-iwa-common\" }\n"
                )
            if candidate:
                packages["litchi-ops"] = (
                    "[package]\nname = \"litchi-ops\"\nversion = \"0.1.0\"\n\n"
                    "[dependencies]\nsha2 = \"0.11\"\n"
                )
        for name, contents in packages.items():
            path = root / "crates" / name / "Cargo.toml"
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(contents, encoding="utf-8")

    def _write_dependency_lock(
        self,
        path: Path,
        *,
        candidate: bool,
        owner: bool = False,
        rogue: bool = False,
        extra_registry_dependency: str | None = None,
        extra_registry_version: str = "9.9.9",
        extra_registry_checksum: str = "c" * 64,
        omit_snap_edge: bool = False,
        omit_snap_registry: bool = False,
        duplicate_smallvec: bool = False,
        snap_version: str = "1.1.2",
        snap_source: str = RUNNER.REGISTRY_SOURCE,
        snap_checksum: str = "a" * 64,
        extra_local_edge: bool = False,
    ) -> None:
        dependencies = ["litchi-iwa-common"]
        if not omit_snap_edge:
            dependencies.append("snap")
        if candidate:
            dependencies.append("smallvec")
        extra_registry_dependency = extra_registry_dependency or ("rogue" if rogue else None)
        if extra_registry_dependency:
            dependencies.append(extra_registry_dependency)
        rows = [
            "version = 4",
            "",
            "[[package]]",
            'name = "litchi-iwa-common"',
            'version = "0.1.0"',
            "",
            "[[package]]",
            'name = "litchi-iwa-core"',
            'version = "0.1.0"',
            "dependencies = [",
            *[f' "{dependency}",' for dependency in dependencies],
            "]",
            "",
        ]
        if not omit_snap_registry:
            rows.extend(
                [
                    "[[package]]",
                    'name = "snap"',
                    f'version = "{snap_version}"',
                    f'source = "{snap_source}"',
                    f'checksum = "{snap_checksum}"',
                ]
            )
        if owner:
            owner_rows = []
            for name in ("litchi-numbers", "litchi-keynote", "litchi-pages"):
                owner_rows.extend(
                    [
                        "",
                        "[[package]]",
                        f'name = "{name}"',
                        'version = "0.1.0"',
                        'dependencies = [',
                        ' "litchi-iwa-core",',
                        *([' "litchi-iwa-common",'] if extra_local_edge else []),
                        "]",
                    ]
                )
            rows.extend(owner_rows)
        if candidate or owner:
            rows.extend(
                [
                    "",
                    "[[package]]",
                    'name = "smallvec"',
                    'version = "1.16.3"',
                    'source = "registry+https://github.com/rust-lang/crates.io-index"',
                    'checksum = "'
                    + RUNNER.REVIEWED_REGISTRY_ADDITIONS["constructor"][0]["checksum"]
                    + '"',
                ]
            )
            if duplicate_smallvec:
                rows.extend(
                    [
                        "",
                        "[[package]]",
                        'name = "smallvec"',
                        'version = "1.16.3"',
                        'source = "registry+https://github.com/rust-lang/crates.io-index"',
                        'checksum = "' + "c" * 64 + '"',
                    ]
                )
        if extra_registry_dependency:
            rows.extend(
                [
                    "",
                    "[[package]]",
                    f'name = "{extra_registry_dependency}"',
                    f'version = "{extra_registry_version}"',
                    'source = "registry+https://github.com/rust-lang/crates.io-index"',
                    f'checksum = "{extra_registry_checksum}"',
                ]
            )
        path.write_text("\n".join(rows) + "\n", encoding="utf-8")

    def _write_seed_lock(
        self,
        path: Path,
        *,
        snap_source: str = RUNNER.REGISTRY_SOURCE,
        rogue: bool = False,
        extra_registry_dependency: str | None = None,
        extra_registry_version: str = "9.9.9",
        extra_registry_checksum: str = "c" * 64,
    ) -> None:
        checksum = RUNNER.REVIEWED_REGISTRY_ADDITIONS["constructor"][0]["checksum"]
        contents = (
            "version = 3\n\n"
            "[[package]]\n"
            'name = "snap"\n'
            'version = "1.1.2"\n'
            f'source = "{snap_source}"\n'
            'checksum = "' + "a" * 64 + '"\n\n'
            "[[package]]\n"
            'name = "smallvec"\n'
            'version = "1.16.3"\n'
            'source = "registry+https://github.com/rust-lang/crates.io-index"\n'
            f'checksum = "{checksum}"\n'
        )
        extra_registry_dependency = extra_registry_dependency or ("rogue" if rogue else None)
        if extra_registry_dependency:
            contents += (
                "\n[[package]]\n"
                f'name = "{extra_registry_dependency}"\n'
                f'version = "{extra_registry_version}"\n'
                'source = "registry+https://github.com/rust-lang/crates.io-index"\n'
                f'checksum = "{extra_registry_checksum}"\n'
            )
        path.write_text(contents, encoding="utf-8")

    def test_endpoint_lock_delta_binds_source_edges_registry_and_features(self) -> None:
        base_source = self.directory / "constructor-base"
        candidate_source = self.directory / "constructor-candidate"
        self._write_dependency_source(base_source, candidate=False, owner=False)
        self._write_dependency_source(candidate_source, candidate=True, owner=False)
        base_lock = self.directory / "constructor-base.lock"
        candidate_lock = self.directory / "constructor-candidate.lock"
        seed_lock = self.directory / "seed.lock"
        self._write_dependency_lock(base_lock, candidate=False)
        self._write_dependency_lock(candidate_lock, candidate=True)
        self._write_seed_lock(seed_lock)
        delta = RUNNER.lock_delta_manifest(
            "constructor", base_source, candidate_source, base_lock, candidate_lock, seed_lock
        )
        self.assertEqual(delta["registry"]["added"][0]["name"], "smallvec")
        self.assertEqual(delta["source_dependencies"]["registry_added"], ["smallvec"])
        self.assertEqual(
            delta["source_dependencies"]["added"][0]["features"], ["union"]
        )
        RUNNER.verify_lock_delta(
            delta,
            "constructor",
            base_source,
            candidate_source,
            base_lock,
            candidate_lock,
            seed_lock,
        )

        # Even when an undeclared registry edge is present in both endpoint
        # locks and the seed, it is outside the selected source closure.
        extra_base = self.directory / "constructor-extra-base.lock"
        extra_candidate = self.directory / "constructor-extra-candidate.lock"
        extra_seed = self.directory / "constructor-extra-seed.lock"
        self._write_dependency_lock(
            extra_base,
            candidate=False,
            extra_registry_dependency="chrono",
            extra_registry_version="1.0.0",
        )
        self._write_dependency_lock(
            extra_candidate,
            candidate=True,
            extra_registry_dependency="chrono",
            extra_registry_version="1.0.0",
        )
        self._write_seed_lock(
            extra_seed,
            extra_registry_dependency="chrono",
            extra_registry_version="1.0.0",
        )
        with self.assertRaisesRegex(RUNNER.QualificationError, "undeclared"):
            RUNNER.lock_delta_manifest(
                "constructor",
                base_source,
                candidate_source,
                extra_base,
                extra_candidate,
                extra_seed,
            )

        # A row shared by both endpoints is still bound to the pinned seed;
        # comparing endpoints alone would accept this changed version.
        shared_bad_base = self.directory / "constructor-shared-bad-base.lock"
        shared_bad_candidate = self.directory / "constructor-shared-bad-candidate.lock"
        self._write_dependency_lock(
            shared_bad_base,
            candidate=False,
            snap_version="9.9.9",
            snap_checksum="b" * 64,
        )
        self._write_dependency_lock(
            shared_bad_candidate,
            candidate=True,
            snap_version="9.9.9",
            snap_checksum="b" * 64,
        )
        with self.assertRaisesRegex(RUNNER.QualificationError, "common"):
            RUNNER.lock_delta_manifest(
                "constructor",
                base_source,
                candidate_source,
                shared_bad_base,
                shared_bad_candidate,
                seed_lock,
            )

        # A seed with the right name/version but a different registry source
        # is not an equivalent pinned identity.
        evil_seed = self.directory / "evil-seed.lock"
        self._write_seed_lock(evil_seed, snap_source="registry+https://evil.invalid/index")
        with self.assertRaisesRegex(RUNNER.QualificationError, "common"):
            RUNNER.lock_delta_manifest(
                "constructor",
                base_source,
                candidate_source,
                base_lock,
                candidate_lock,
                evil_seed,
            )

        tampered = copy.deepcopy(delta)
        tampered["registry"]["common"][0]["checksum"] = "0" * 64
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.verify_lock_delta(
                tampered,
                "constructor",
                base_source,
                candidate_source,
                base_lock,
                candidate_lock,
                seed_lock,
            )
        tampered = copy.deepcopy(delta)
        tampered["source_dependencies"]["added"][0]["features"] = []
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.verify_lock_delta(
                tampered,
                "constructor",
                base_source,
                candidate_source,
                base_lock,
                candidate_lock,
                seed_lock,
            )

        rogue_lock = self.directory / "constructor-rogue.lock"
        self._write_dependency_lock(rogue_lock, candidate=True, rogue=True)
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.lock_delta_manifest(
                "constructor", base_source, candidate_source, base_lock, rogue_lock, seed_lock
            )
        for suffix, old, new in (
            ("wrong-version", 'version = "1.16.3"', 'version = "9.9.9"'),
            (
                "wrong-checksum",
                "5b3dc8af474f516a851ff4bd12db780f948b9250ad37211e4eec0bccea54e01b",
                "d" * 64,
            ),
        ):
            bad_identity_lock = self.directory / f"constructor-{suffix}.lock"
            contents = candidate_lock.read_text(encoding="utf-8").replace(old, new, 1)
            bad_identity_lock.write_text(contents, encoding="utf-8")
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.lock_delta_manifest(
                    "constructor",
                    base_source,
                    candidate_source,
                    base_lock,
                    bad_identity_lock,
                    seed_lock,
                )
        bad_seed = self.directory / "bad-seed.lock"
        bad_seed.write_text(
            seed_lock.read_text(encoding="utf-8").replace(
                RUNNER.REVIEWED_REGISTRY_ADDITIONS["constructor"][0]["checksum"],
                "e" * 64,
                1,
            ),
            encoding="utf-8",
        )
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.lock_delta_manifest(
                "constructor",
                base_source,
                candidate_source,
                base_lock,
                candidate_lock,
                bad_seed,
            )

        missing_edge_lock = self.directory / "constructor-missing-edge.lock"
        self._write_dependency_lock(missing_edge_lock, candidate=True, omit_snap_edge=True)
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.lock_delta_manifest(
                "constructor",
                base_source,
                candidate_source,
                base_lock,
                missing_edge_lock,
                seed_lock,
            )
        missing_package_lock = self.directory / "constructor-missing-package.lock"
        self._write_dependency_lock(missing_package_lock, candidate=True, omit_snap_registry=True)
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.lock_delta_manifest(
                "constructor",
                base_source,
                candidate_source,
                base_lock,
                missing_package_lock,
                seed_lock,
            )
        duplicate_lock = self.directory / "constructor-duplicate.lock"
        self._write_dependency_lock(duplicate_lock, candidate=True, duplicate_smallvec=True)
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.lock_delta_manifest(
                "constructor",
                base_source,
                candidate_source,
                base_lock,
                duplicate_lock,
                seed_lock,
            )

    def test_owner_lock_delta_uses_owner_closure_not_unrelated_workspace_edges(self) -> None:
        base_source = self.directory / "owner-base"
        candidate_source = self.directory / "owner-candidate"
        self._write_dependency_source(base_source, candidate=False, owner=True)
        self._write_dependency_source(candidate_source, candidate=True, owner=True)
        base_lock = self.directory / "owner-base.lock"
        candidate_lock = self.directory / "owner-candidate.lock"
        seed_lock = self.directory / "owner-seed.lock"
        self._write_dependency_lock(
            base_lock, candidate=False, owner=True, extra_local_edge=True
        )
        self._write_dependency_lock(
            candidate_lock, candidate=True, owner=True, extra_local_edge=True
        )
        self._write_seed_lock(seed_lock)
        delta = RUNNER.lock_delta_manifest(
            "owner", base_source, candidate_source, base_lock, candidate_lock, seed_lock
        )
        self.assertEqual(delta["source_dependencies"]["registry_added"], ["smallvec"])
        self.assertEqual(delta["registry"]["added"], [])
        self.assertEqual(delta["reviewed_registry_additions"], [])
        self.assertNotIn("litchi-ops", delta["source_dependencies"]["packages"])
        RUNNER.verify_lock_delta(
            delta, "owner", base_source, candidate_source, base_lock, candidate_lock, seed_lock
        )

    def test_validation_report_requires_all_routes_and_typed_boundaries(self) -> None:
        value = {
            "schema": "litchi.iwa.constructor.validation.v1",
            "status": "ok",
            "routes": list(RUNNER.ROUTES["constructor"]),
            "route_checks": [
                {"route": route, "status": "ok", "output_bytes": 12}
                for route in RUNNER.ROUTES["constructor"]
            ],
            "source_sha256": "a" * 64,
            "encoded_sha256": "a" * 64,
            "source_bytes": 12,
            "encoded_bytes": 12,
            "roundtrip_equal": True,
            "retained_source_sha256": "b" * 64,
            "retained_encoded_sha256": "b" * 64,
            "retained_roundtrip_equal": True,
            "header_boundary": {
                "kind": "HeaderBytes",
                "accepted": 3,
                "below": {"observed": 4, "maximum": 2},
            },
            "retained_header_boundary": {
                "kind": "HeaderFields",
                "accepted": 3,
                "below": {"observed": 4, "maximum": 2},
            },
        }
        RUNNER._validate_report_value(value, "constructor")
        value["routes"] = value["routes"][:-1]
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._validate_report_value(value, "constructor")

    def test_duplicate_json_keys_are_rejected(self) -> None:
        path = self.directory / "duplicate.json"
        path.write_text('{"a": 1, "a": 2}', encoding="utf-8")
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.load_json(path)

    def test_constructor_probe_requires_exact_roundtrip(self) -> None:
        value = {
            "schema": "litchi.iwa.constructor.probe.v3",
            "route": "field_path",
            "iterations": 3,
            "ns": 10,
            "checksum": 1,
            "output_bytes": 12,
            "source_sha256": "a" * 64,
            "output_sha256": "a" * 64,
            "output_matches_source": True,
            "correctness": {"status": "ok", "roundtrip": True},
        }
        self.assertIs(RUNNER._strict_probe_value(value, "constructor", "field_path", 3), value)
        value["correctness"]["roundtrip"] = False
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._strict_probe_value(value, "constructor", "field_path", 3)

    def test_owner_probe_requires_replay_inverse_and_noop(self) -> None:
        value = {
            "schema": "litchi.iwa.owner.probe.v2",
            "route": "numbers_headers",
            "iterations": 3,
            "ns": 10,
            "checksum": 1,
            "output_bytes": 12,
            "source_sha256": "a" * 64,
            "output_sha256": "b" * 64,
            "output_differs": True,
            "correctness": {
                "status": "ok",
                "replay_matches_changed": True,
                "inverse_matches_source": True,
                "noop_is_noop": True,
                "noop_matches_expected": True,
                "semantic_ok": True,
                "changed_sha256": "b" * 64,
                "replay_sha256": "b" * 64,
                "inverse_sha256": "a" * 64,
                "noop_sha256": "a" * 64,
                "source_bytes": 12,
                "changed_bytes": 12,
                "noop_bytes": 12,
            },
        }
        self.assertIs(RUNNER._strict_probe_value(value, "owner", "numbers_headers", 3), value)
        value["correctness"]["inverse_matches_source"] = False
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._strict_probe_value(value, "owner", "numbers_headers", 3)

    def test_allocation_probe_requires_typed_metrics(self) -> None:
        value = {
            "schema": "litchi.iwa.owner.probe.v2",
            "route": "numbers_headers",
            "iterations": 3,
            "ns": 10,
            "checksum": 1,
            "output_bytes": 12,
            "source_sha256": "a" * 64,
            "output_sha256": "b" * 64,
            "output_differs": True,
            "correctness": {
                "status": "ok",
                "replay_matches_changed": True,
                "inverse_matches_source": True,
                "noop_is_noop": True,
                "noop_matches_expected": True,
                "semantic_ok": True,
                "changed_sha256": "b" * 64,
                "replay_sha256": "b" * 64,
                "inverse_sha256": "a" * 64,
                "noop_sha256": "a" * 64,
                "source_bytes": 12,
                "changed_bytes": 12,
                "noop_bytes": 12,
            },
        }
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._strict_probe_value(
                value, "owner", "numbers_headers", 3, require_allocations=True
            )
        value.update({"peak": 8, "retained": 4, "allocated": 16, "allocations": 2})
        self.assertIs(
            RUNNER._strict_probe_value(
                value, "owner", "numbers_headers", 3, require_allocations=True
            ),
            value,
        )

    def test_allocation_abba_requires_pairwise_metric_identity(self) -> None:
        values = [
            {"checksum": 1, "output_sha256": "b", "source_sha256": "a", **dict.fromkeys(RUNNER.ALLOCATION_METRICS, 4)},
            {"checksum": 1, "output_sha256": "b", "source_sha256": "a", **dict.fromkeys(RUNNER.ALLOCATION_METRICS, 5)},
            {"checksum": 1, "output_sha256": "b", "source_sha256": "a", **dict.fromkeys(RUNNER.ALLOCATION_METRICS, 5)},
            {"checksum": 1, "output_sha256": "b", "source_sha256": "a", **dict.fromkeys(RUNNER.ALLOCATION_METRICS, 4)},
        ]
        RUNNER._allocation_abba_equal(values)
        values[3]["peak"] = 9
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._allocation_abba_equal(values)

    def test_invoke_records_receipt_and_rejects_wrong_identity(self) -> None:
        binary = self.directory / "probe.py"
        _fake_probe(binary, owner=False)
        expected = RUNNER.artifact(binary)
        with mock.patch.object(RUNNER, "host_snapshot"):
            value = RUNNER.invoke(
                self.directory,
                "constructor",
                "field_path",
                binary,
                "field_path",
                1,
                min(RUNNER.os.sched_getaffinity(0)),
                None,
                expected,
            )
        self.assertEqual(value["route"], "field_path")
        receipt = RUNNER.load_json(self.directory / "field_path.receipt.json")
        self.assertEqual(receipt["exit_code"], 0)
        self.assertEqual(receipt["binary"], expected)

    def test_invoke_allocation_requires_metrics(self) -> None:
        binary = self.directory / "probe.py"
        _fake_probe(binary, owner=True, allocation=True)
        expected = RUNNER.artifact(binary)
        with mock.patch.object(RUNNER, "host_snapshot"):
            value = RUNNER.invoke(
                self.directory,
                "owner",
                "numbers_headers-allocation",
                binary,
                "numbers_headers",
                1,
                min(RUNNER.os.sched_getaffinity(0)),
                self.directory,
                expected,
                allocation=True,
            )
        self.assertEqual(value["allocations"], 2)

    def test_invoke_fails_closed_on_owner_correctness(self) -> None:
        binary = self.directory / "probe.py"
        _fake_probe(binary, owner=True, valid=False)
        expected = RUNNER.artifact(binary)
        with mock.patch.object(RUNNER, "host_snapshot"):
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.invoke(
                    self.directory,
                    "owner",
                    "numbers_headers",
                    binary,
                    "numbers_headers",
                    1,
                    min(RUNNER.os.sched_getaffinity(0)),
                    self.directory,
                    expected,
                )

    def test_invoke_enforces_timed_child_bound(self) -> None:
        binary = self.directory / "slow_probe.py"
        binary.write_text(
            "#!/usr/bin/env python3\nimport time\ntime.sleep(1)\n",
            encoding="utf-8",
        )
        binary.chmod(binary.stat().st_mode | stat.S_IXUSR)
        expected = RUNNER.artifact(binary)
        with mock.patch.object(RUNNER, "host_snapshot"):
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.invoke(
                    self.directory,
                    "constructor",
                    "timeout",
                    binary,
                    "field_path",
                    1,
                    min(RUNNER.os.sched_getaffinity(0)),
                    None,
                    expected,
                    timeout_s=0.05,
                )
        receipt = RUNNER.load_json(self.directory / "timeout.receipt.json")
        self.assertTrue(receipt["timed_out"])
        self.assertEqual(receipt["timeout_s"], 0.05)

    def test_command_runner_enforces_build_bound(self) -> None:
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.run_command(
                self.directory,
                "slow-command",
                ["/usr/bin/python3", "-c", "import time; time.sleep(1)"],
                self.directory,
                timeout_s=0.05,
            )
        receipt = RUNNER.load_json(self.directory / "slow-command.receipt.json")
        self.assertTrue(receipt["timed_out"])
        self.assertEqual(receipt["timeout_s"], 0.05)

    def test_endpoint_lock_preparation_uses_metadata_not_refresh(self) -> None:
        self.assertEqual(
            RUNNER.LOCK_PREPARATION_COMMAND,
            ("cargo", "metadata", "--offline", "--format-version=1"),
        )
        self.assertNotIn("--locked", RUNNER.LOCK_PREPARATION_COMMAND)
        with mock.patch.object(RUNNER, "run_command", return_value={}) as run:
            result = RUNNER.prepare_endpoint_lock(self.directory, self.directory, {})
        self.assertEqual(result, {})
        run.assert_called_once_with(
            self.directory,
            "metadata",
            list(RUNNER.LOCK_PREPARATION_COMMAND),
            self.directory,
            {},
        )

        RUNNER.verify_lock_preparation_command(
            [{"command": list(RUNNER.LOCK_PREPARATION_COMMAND)}]
        )
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.verify_lock_preparation_command(
                [{"command": ["cargo", "generate-lockfile", "--offline"]}]
            )
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.verify_lock_preparation_command(
                [{"command": [*RUNNER.LOCK_PREPARATION_COMMAND, "--locked"]}]
            )

    def test_generator_timeout_uses_coordinated_process_cleanup(self) -> None:
        child = mock.Mock()
        child.pid = 12345
        child.wait.side_effect = subprocess.TimeoutExpired("generator", 1)
        with mock.patch.object(RUNNER.subprocess, "Popen", return_value=child):
            with mock.patch.object(RUNNER, "_terminate_generator") as terminate:
                with self.assertRaises(RUNNER.QualificationError):
                    RUNNER._run_fixture_generator(self.directory / "footnote-fixture.pages")
        terminate.assert_called_once_with(child)

    def test_abba_identity_check_rejects_mismatched_outputs(self) -> None:
        values = [
            {"checksum": 1, "output_sha256": "a", "source_sha256": "s"},
            {"checksum": 1, "output_sha256": "b", "source_sha256": "s"},
            {"checksum": 1, "output_sha256": "b", "source_sha256": "s"},
            {"checksum": 1, "output_sha256": "a", "source_sha256": "s"},
        ]
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._ratios_equal(values)

    def test_host_guard_fails_closed_on_foreign_build(self) -> None:
        foreign = {"pid": 99999, "parent": 1, "exe": "/usr/bin/cargo", "args": "cargo build"}
        with mock.patch.object(RUNNER, "process_rows", return_value=[foreign]):
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.host_snapshot(self.directory, "blocked")
        snapshot = RUNNER.load_json(self.directory / "blocked.host.json")
        self.assertEqual(snapshot["blockers"], [foreign])


if __name__ == "__main__":
    unittest.main()
