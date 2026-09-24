"""Git-blob and provenance regressions for the XLSB smoke scaffold."""

from __future__ import annotations

import subprocess
import tempfile
import unittest
import copy
import hashlib
import json
from pathlib import Path
from unittest.mock import patch

import source_manifest
import verify


class GitBlobSnapshotTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory(prefix="litchi-xlsb-profile-git-")
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        subprocess.run(["git", "-C", str(self.root), "init", "-q"], check=True)
        subprocess.run(
            ["git", "-C", str(self.root), "config", "user.name", "Profile Test"],
            check=True,
        )
        subprocess.run(
            [
                "git",
                "-C",
                str(self.root),
                "config",
                "user.email",
                "profile@example.invalid",
            ],
            check=True,
        )
        self.source = self.root / "input.rs"
        self.source.write_text("fn source() {}\n")
        subprocess.run(["git", "-C", str(self.root), "add", "input.rs"], check=True)
        subprocess.run(
            ["git", "-C", str(self.root), "commit", "-qm", "source"],
            check=True,
        )
        self.commit = subprocess.check_output(
            ["git", "-C", str(self.root), "rev-parse", "HEAD"],
            text=True,
        ).strip()

    def assert_rejected(self) -> None:
        with self.assertRaisesRegex(SystemExit, "committed Git blobs"):
            source_manifest.verify_git_inputs([self.source], self.root, self.commit)
        with self.assertRaisesRegex(AssertionError, "committed snapshot"):
            verify.verify_git_snapshot([self.source], self.root, self.commit)

    def test_clean_input_is_accepted(self) -> None:
        source_manifest.verify_git_inputs([self.source], self.root, self.commit)
        verify.verify_git_snapshot([self.source], self.root, self.commit)

    def test_unstaged_dirty_input_is_rejected(self) -> None:
        self.source.write_text("fn changed() {}\n")
        self.assert_rejected()

    def test_staged_dirty_input_is_rejected(self) -> None:
        self.source.write_text("fn staged() {}\n")
        subprocess.run(["git", "-C", str(self.root), "add", "input.rs"], check=True)
        self.assert_rejected()

    def test_assume_unchanged_does_not_hide_dirty_input(self) -> None:
        subprocess.run(
            ["git", "-C", str(self.root), "update-index", "--assume-unchanged", "input.rs"],
            check=True,
        )
        self.source.write_text("fn hidden_change() {}\n")
        self.assert_rejected()

    def test_untracked_input_is_rejected(self) -> None:
        untracked = self.root / "untracked.rs"
        untracked.write_text("fn untracked() {}\n")
        with self.assertRaisesRegex(SystemExit, "not tracked"):
            source_manifest.verify_git_inputs([untracked], self.root, self.commit)
        with self.assertRaisesRegex(AssertionError, "not tracked"):
            verify.verify_git_snapshot([untracked], self.root, self.commit)

    def test_manifest_commit_mismatch_is_rejected(self) -> None:
        wrong_commit = "0" * 40
        with self.assertRaisesRegex(SystemExit, "Git head changed"):
            source_manifest.verify_git_inputs([self.source], self.root, wrong_commit)
        with self.assertRaisesRegex(AssertionError, "Git snapshot commit changed"):
            verify.verify_git_snapshot([self.source], self.root, wrong_commit)

    def test_metadata_provenance_mismatch_is_rejected(self) -> None:
        results = self.root / "results"
        results.mkdir()
        metadata = results / "metadata-before.json"
        metadata.write_text("{}\n")
        (results / "metadata-after.json").write_text("{}\n")
        source_manifest_hash = "a" * 64
        binary = results / "xlsb-model-identity-profile"
        binary.write_bytes(b"binary")
        binary_digest = hashlib.sha256(binary.read_bytes()).hexdigest()
        (results / "provenance.txt").write_text(
            "\n".join(
                (
                    f"git_head={self.commit}",
                    "git_baseline=host",
                    "neutral_baseline=neutral",
                    "git_status_relevant=",
                    "metadata_before_sha256=" + "b" * 64,
                    "metadata_after_sha256=" + "b" * 64,
                    f"source_manifest_before_sha256={source_manifest_hash}",
                    f"source_manifest_after_sha256={source_manifest_hash}",
                    f"binary_sha256={binary_digest}",
                    f"binary_after_sha256={binary_digest}",
                )
            )
            + "\n"
        )
        with self.assertRaisesRegex(SystemExit, "metadata hashes"):
            verify.verify_provenance(
                results,
                self.root,
                {"host_feature_commit": "host", "neutral_baseline_commit": "neutral"},
                self.commit,
                hashlib.sha256(metadata.read_bytes()).hexdigest(),
                source_manifest_hash,
                binary_digest,
            )

    def test_binary_receipt_mismatch_is_rejected(self) -> None:
        results = self.root / "binary-results"
        results.mkdir()
        binary = results / "xlsb-model-identity-profile"
        binary.write_bytes(b"binary")
        digest = hashlib.sha256(binary.read_bytes()).hexdigest()
        receipt = f"{digest}  {binary}\n"
        (results / "binary.sha256").write_text(receipt)
        (results / "binary-after.sha256").write_text("0" * 64 + f"  {binary}\n")
        (results / "build-provenance.txt").write_text(f"binary={binary}\n{receipt}")
        with self.assertRaisesRegex(SystemExit, "changed during smoke"):
            verify.verify_binary_receipts(results)

    def test_matrix_capture_rejects_tampered_argv_and_receipt(self) -> None:
        hashed_binary = "/tmp/xlsb-profile-binary"
        digest = "a" * 64

        def write_capture(directory: Path, argv: list[str], stdout: str = "{}", stderr: str = "") -> None:
            directory.mkdir()
            (directory / "binary.sha256").write_text(f"{digest}  {hashed_binary}\n")
            (directory / "matrix-correctness.argv.json").write_text(
                json.dumps({"argv": argv}) + "\n"
            )
            (directory / "matrix-correctness.stdout.json").write_text(stdout)
            (directory / "matrix-correctness.stderr.log").write_text(stderr)
            (directory / "matrix-correctness.exit.txt").write_text("0\n")

        valid = self.root / "matrix-valid"
        write_capture(valid, [hashed_binary, "--matrix-correctness"])
        with patch.object(verify, "verify_correctness_receipt") as verifier:
            verify.verify_matrix_runner_capture(valid, self.root / "manifest.json")
            verifier.assert_called_once_with(
                valid / "matrix-correctness.stdout.json", self.root / "manifest.json"
            )

        wrong_binary = self.root / "matrix-wrong-binary"
        write_capture(wrong_binary, ["/tmp/other-binary", "--matrix-correctness"])
        with self.assertRaisesRegex(SystemExit, "does not match the hashed binary"):
            verify.verify_matrix_runner_capture(
                wrong_binary, self.root / "manifest.json"
            )

        extra_argument = self.root / "matrix-extra-argument"
        write_capture(
            extra_argument,
            [hashed_binary, "--matrix-correctness", "--case", "tiny:1:0"],
        )
        with self.assertRaisesRegex(SystemExit, "does not match the hashed binary"):
            verify.verify_matrix_runner_capture(
                extra_argument, self.root / "manifest.json"
            )

        malformed_receipt = self.root / "matrix-malformed-receipt"
        write_capture(malformed_receipt, [hashed_binary, "--matrix-correctness"], "[]")
        with self.assertRaisesRegex(SystemExit, "receipt is not an object"):
            verify.verify_matrix_runner_capture(
                malformed_receipt, self.root / "manifest.json"
            )

    def test_correctness_semantic_check_requires_every_gate(self) -> None:
        with self.assertRaisesRegex(SystemExit, "semantic equality check fields"):
            verify.verify_matrix_check({"all_equal": True}, "matrix")

    def test_correctness_semantic_check_rejects_false_relationship_endpoint(self) -> None:
        check = {
            field: True for field in verify.MATRIX_CHECK_FIELDS
        }
        check["relationship_endpoints_equal"] = False
        with self.assertRaisesRegex(SystemExit, "semantic equality check failed"):
            verify.verify_matrix_check(check, "matrix")

    def test_correctness_semantic_vector_recomputes_changed_table_name(self) -> None:
        source = {
            "tables": [
                {
                    "table_id": "T1",
                    "xml_name": "Table1",
                    "metadata_path": "Model.1.db/T1.0.dim/T1.1.tbl.xml",
                    "dimension_object_id": "11111111-2222-3333-4444-555555555501",
                }
            ],
            "relationships": [],
            "time_groupings": [
                {"table_name": "Table1", "column_id": "Key", "column_ids": ["Year"]}
            ],
        }
        expected = copy.deepcopy(source)
        expected["tables"][0]["xml_name"] = "TableX"
        expected["time_groupings"][0]["table_name"] = "TableX"
        result = {
            "name_profile": "same_length_ascii",
            "endpoint_layout": "selected_table",
            "source_semantic": source,
            "semantic_observed": expected,
            "reopened_semantic_observed": expected,
            "semantic": verify.recompute_matrix_check(expected, expected),
            "reopened_semantic": verify.recompute_matrix_check(expected, expected),
        }
        verify.verify_matrix_semantic_transition(result, 1, 0)
        mutated = copy.deepcopy(result)
        mutated["semantic_observed"]["tables"][0]["xml_name"] = "unverified"
        with self.assertRaisesRegex(SystemExit, "independently recomputed identity"):
            verify.verify_matrix_semantic_transition(mutated, 1, 0)
        source_mutated = copy.deepcopy(result)
        source_mutated["source_semantic"]["tables"][0][
            "dimension_object_id"
        ] = "11111111-2222-3333-4444-555555555599"
        source_mutated["semantic_observed"]["tables"][0][
            "dimension_object_id"
        ] = source_mutated["source_semantic"]["tables"][0]["dimension_object_id"]
        source_mutated["reopened_semantic_observed"]["tables"][0][
            "dimension_object_id"
        ] = source_mutated["source_semantic"]["tables"][0]["dimension_object_id"]
        with self.assertRaisesRegex(SystemExit, "dimension object ID"):
            verify.verify_matrix_semantic_transition(source_mutated, 1, 0)

    def test_correctness_source_recipe_rejects_coherent_identity_mutations(self) -> None:
        source = {
            "tables": [
                {
                    "table_id": f"T{index}",
                    "xml_name": f"Table{index}",
                    "metadata_path": f"Model.1.db/T{index}.0.dim/T{index}.1.tbl.xml",
                    "dimension_object_id": (
                        "11111111-2222-3333-4444-"
                        f"{0x5555_5555_5500 + index:012X}"
                    ),
                }
                for index in range(1, 5)
            ],
            "relationships": verify.expected_matrix_relationships(
                4, 3, "selected_table"
            ),
            "time_groupings": [
                {
                    "table_name": f"Table{index}",
                    "column_id": "Key",
                    "column_ids": ["Year"],
                }
                for index in range(1, 5)
            ],
        }
        expected = verify.expected_matrix_semantic(source, "same_length_ascii")
        result = {
            "name_profile": "same_length_ascii",
            "endpoint_layout": "selected_table",
            "source_semantic": source,
            "semantic_observed": expected,
            "reopened_semantic_observed": copy.deepcopy(expected),
            "semantic": verify.recompute_matrix_check(expected, expected),
            "reopened_semantic": verify.recompute_matrix_check(expected, expected),
        }
        verify.verify_matrix_semantic_transition(result, 4, 3)

        def assert_coherent_mutation(mutator, message: str) -> None:
            mutated = copy.deepcopy(result)
            for field in (
                "source_semantic",
                "semantic_observed",
                "reopened_semantic_observed",
            ):
                mutator(mutated[field])
            with self.assertRaisesRegex(SystemExit, message):
                verify.verify_matrix_semantic_transition(mutated, 4, 3)

        assert_coherent_mutation(
            lambda vector: vector["tables"][0].update(
                metadata_path="Model.1.db/T1.0.dim/T99.1.tbl.xml"
            ),
            "table metadata path",
        )
        assert_coherent_mutation(
            lambda vector: vector["relationships"][0].update(
                relationship_id="Rel99",
                metadata_path="Model.1.db/T1.0.dim/R$T1$Rel99.1.tbl.xml",
                expected_index_key="R$T1$Rel99",
            ),
            "relationship topology",
        )
        assert_coherent_mutation(
            lambda vector: vector["time_groupings"][0].update(table_name="Table99"),
            "time-grouping identity",
        )

    def test_correctness_member_preservation_rejects_native_change(self) -> None:
        digest = "a" * 64
        changed = "b" * 64
        preservation = {
            "source": {
                "parts": {"/xl/model/item.data": digest, "/xl/opaque.bin": digest},
                "relationships": {"/xl/model/item.data": digest},
                "content_types": digest,
                "inner_members": {
                    "native.idf": digest,
                    "BackupLog": digest,
                    "Model.1.db.xml": digest,
                    "T1.1.tbl.xml": digest,
                    "T2.1.tbl.xml": digest,
                },
            },
            "candidate": {
                "parts": {"/xl/model/item.data": changed, "/xl/opaque.bin": digest},
                "relationships": {"/xl/model/item.data": digest},
                "content_types": digest,
                "inner_members": {
                    "native.idf": changed,
                    "BackupLog": digest,
                    "Model.1.db.xml": digest,
                    "T1.1.tbl.xml": digest,
                    "T2.1.tbl.xml": digest,
                },
            },
            "mutable_inner_paths": ["BackupLog", "T1.1.tbl.xml"],
        }
        source = {
            "tables": [
                {
                    "table_id": "T1",
                    "xml_name": "Table1",
                    "metadata_path": "T1.1.tbl.xml",
                    "dimension_object_id": "dimension-1",
                },
                {
                    "table_id": "T2",
                    "xml_name": "Table2",
                    "metadata_path": "T2.1.tbl.xml",
                    "dimension_object_id": "dimension-2",
                },
            ],
            "relationships": [],
            "time_groupings": [],
        }
        with self.assertRaisesRegex(SystemExit, "unrelated XLDM member changed"):
            verify.verify_matrix_preservation(preservation, source)

        t2_changed = copy.deepcopy(preservation)
        t2_changed["candidate"]["inner_members"]["T2.1.tbl.xml"] = changed
        with self.assertRaisesRegex(SystemExit, "unrelated XLDM member changed"):
            verify.verify_matrix_preservation(t2_changed, source)

        broad = copy.deepcopy(preservation)
        broad["candidate"]["inner_members"]["native.idf"] = digest
        broad["mutable_inner_paths"] = [
            "BackupLog",
            "Model.1.db.xml",
            "T1.1.tbl.xml",
        ]
        with self.assertRaisesRegex(SystemExit, "mutable inner path scope"):
            verify.verify_matrix_preservation(broad, source)

    def test_olap_proof_defaults_bind_work_limit(self) -> None:
        limits = copy.deepcopy(verify.EXPECTED_OLAP_PROOF_LIMITS)
        receipt = {"olap_proof_limits": limits}
        recipe = {"olap_proof_limits": limits}
        manifest = {"olap_proof_limits": limits, "recipe": {"olap_proof_limits": limits}}
        verify.verify_olap_proof_limits(receipt, recipe, manifest)
        limits["max_work"] = 4_000_001
        with self.assertRaisesRegex(SystemExit, "reviewed defaults"):
            verify.verify_olap_proof_limits(receipt, recipe, manifest)

    def test_full_manifest_parser_binds_local_cargo_files_to_git_blobs(self) -> None:
        package_root = self.root / "fixture"
        source_dir = package_root / "src"
        source_dir.mkdir(parents=True)
        manifest = package_root / "Cargo.toml"
        source = source_dir / "lib.rs"
        manifest.write_text(
            "[package]\nname = \"fixture\"\nversion = \"0.1.0\"\nedition = \"2021\"\n"
        )
        source.write_text("pub fn identity() -> &'static str { \"source\" }\n")
        subprocess.run(
            ["git", "-C", str(self.root), "add", "fixture"], check=True
        )
        subprocess.run(
            ["git", "-C", str(self.root), "commit", "-qm", "fixture"], check=True
        )
        commit = subprocess.check_output(
            ["git", "-C", str(self.root), "rev-parse", "HEAD"],
            text=True,
        ).strip()
        metadata = self.root / "metadata.json"
        metadata.write_text(
            json.dumps(
                {
                    "packages": [
                        {
                            "name": "fixture",
                            "version": "0.1.0",
                            "source": None,
                            "manifest_path": str(manifest),
                            "targets": [{"src_path": str(source)}],
                        }
                    ]
                }
            )
        )
        output = self.root / "source-manifest.txt"
        with patch(
            "sys.argv",
            [
                "source_manifest.py",
                "--metadata",
                str(metadata),
                "--root",
                str(self.root),
                "--output",
                str(output),
                "--git-commit",
                commit,
            ],
        ):
            source_manifest.main()

        source.write_text("pub fn identity() -> &'static str { \"dirty\" }\n")
        lines = output.read_text().splitlines()
        entries = []
        rewritten = []
        for line in lines:
            parts = line.split("\t")
            if line.startswith("file=") and parts[3] == "fixture/src/lib.rs":
                parts[4] = source_manifest.sha256(source)
                line = "\t".join(parts)
            if line.startswith("file="):
                entries.append((parts[3], parts[4]))
            rewritten.append(line)
        tree_hash = hashlib.sha256(
            "\n".join(f"{shown}\t{digest}" for shown, digest in entries).encode()
        ).hexdigest()
        for index, line in enumerate(rewritten):
            if line.startswith("package="):
                parts = line.split("\t")
                parts[6] = tree_hash
                rewritten[index] = "\t".join(parts)
                break
        output.write_text("\n".join(rewritten) + "\n")

        with self.assertRaisesRegex(SystemExit, "committed snapshot"):
            verify.verify_source_manifest(output, self.root, require_transitive=False)


if __name__ == "__main__":
    unittest.main()
