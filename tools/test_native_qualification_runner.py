#!/usr/bin/env python3
"""Static/unit checks for the native qualification launcher.

These tests exercise manifest and guard logic only.  They do not invoke Cargo,
Rust, a benchmark binary, or a timing workload.
"""

from __future__ import annotations

import json
import hashlib
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import native_qualification_runner as runner  # noqa: E402
from native_qualification import prepare_w4b_w5_inputs as source_prep  # noqa: E402
from native_qualification import w5_analyzer_adapter as w5_adapter  # noqa: E402


class NativeQualificationRunnerTests(unittest.TestCase):
    def test_cargo_flags_are_added_once_before_rustc_separator(self) -> None:
        argv = runner.normalize_cargo_argv(
            ["cargo", "build", "--locked", "--", "-C", "link-arg=--offline"]
        )
        self.assertEqual(argv, ["cargo", "build", "--locked", "--offline", "--", "-C", "link-arg=--offline"])

    def test_python_literal_does_not_trigger_heavy_process_detection(self) -> None:
        python = {
            "comm": "python3",
            "exe": "/usr/bin/python3",
            "cwd": str(ROOT),
            "cmdline": ["python3", "-c", "print('cargo')"],
        }
        cargo = {
            "comm": "cargo",
            "exe": "/home/toolchain/bin/cargo",
            "cwd": "/home/zhuhe/code/litchi-worktrees/other",
            "cmdline": ["cargo", "build"],
        }
        unrelated = {
            "comm": "cargo",
            "exe": "/home/toolchain/bin/cargo",
            "cwd": "/tmp/unrelated-repository",
            "cmdline": ["cargo", "build"],
        }
        self.assertFalse(runner._is_heavy(python))
        self.assertTrue(runner._is_heavy(cargo))
        self.assertTrue(runner._in_litchi_scope(cargo))
        self.assertFalse(runner._in_litchi_scope(unrelated))

    def test_manifests_are_parameterized_and_keep_workload_order(self) -> None:
        paths = [
            ROOT / "tools/native_qualification/manifests/w4b-xlsb.json",
            ROOT / "tools/native_qualification/manifests/legacy-w5.json",
        ]
        for path in paths:
            text = path.read_text(encoding="utf-8")
            self.assertNotRegex(text, r"(?<![A-Za-z0-9])[0-9a-f]{40}(?![A-Za-z0-9])")
            data = json.loads(text)
            self.assertEqual(data["abba_order"], list(runner.ABBA_ORDER))
            self.assertEqual(
                [item["id"] for item in data["legs"]["before"]["programs"]],
                [item["id"] for item in data["legs"]["after"]["programs"]],
            )
            self.assertEqual(data["rounds"], 31)

    def test_manifest_shape_loads_without_source_ids(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-manifest-") as directory:
            target_root = Path(directory) / "targets"
            output = Path(directory) / "output"
            environment = {
                "NATIVE_QUAL_OUTPUT": str(output),
                "W4B_BEFORE_SOURCE": str(ROOT),
                "W4B_AFTER_SOURCE": str(Path(directory) / "w4b-after-source"),
                "W4B_TARGET_ROOT": str(target_root),
                "W4B_XLSB_FIXTURE": str(ROOT / "Cargo.lock"),
                "W5_BEFORE_SOURCE": str(ROOT),
                "W5_AFTER_SOURCE": str(Path(directory) / "w5-after-source"),
                "W5_TARGET_ROOT": str(target_root),
                "LITCHI_COORD_BIN": directory,
                "LITCHI_NATIVE_QUAL_ROOT": str(ROOT),
                "W4B_ANALYZER": str(ROOT / "tools/native_qualification_runner.py"),
                "W5_ANALYZER": str(ROOT / "tools/native_qualification_runner.py"),
            }
            old = {key: os.environ.get(key) for key in environment}
            try:
                os.environ.update(environment)
                for name in ("w4b-xlsb.json", "legacy-w5.json"):
                    manifest_path = ROOT / "tools/native_qualification/manifests" / name
                    data = runner.load_manifest(manifest_path)
                    self.assertEqual(data["schema"], runner.SCHEMA)
            finally:
                for key, value in old.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value

    def test_xlsx_raw_manifest_preserves_original_and_candidate_routes(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xlsx-raw-manifest-") as directory:
            root = Path(directory)
            environment = {
                "NATIVE_QUAL_OUTPUT": str(root / "output"),
                "LITCHI_COORD_BIN": str(root),
                "LITCHI_NATIVE_QUAL_ROOT": str(ROOT),
                "XLSX_BEFORE_SOURCE": str(ROOT),
                "XLSX_AFTER_SOURCE": str(root / "after-source"),
                "XLSX_TARGET_ROOT": str(root / "targets"),
                "XLSX_PREPARATION_RECEIPT": str(root / "preparation.json"),
                "XLSX_BEFORE_REVISION": "0befb112d96d2409e6f40c9abfeb4813665905ad",
                "XLSX_AFTER_REVISION": "1" * 40,
            }
            old = {key: os.environ.get(key) for key in environment}
            try:
                os.environ.update(environment)
                manifest = runner.load_manifest(
                    ROOT / "tools/native_qualification/manifests/w4-xlsx-raw.json"
                )
                self.assertEqual(manifest["lane"], "w4-xlsx-raw")
                self.assertEqual(
                    [program["id"] for program in manifest["legs"]["before"]["programs"]],
                    ["raw-alloc", "raw-native"],
                )
                self.assertEqual(
                    manifest["legs"]["before"]["build"][0]["argv"][8],
                    "allocator-metrics,legacy-xlsx-raw-opc",
                )
                self.assertEqual(
                    manifest["legs"]["after"]["build"][0]["argv"][7],
                    "allocator-metrics,current-xlsx-raw-opc",
                )
                self.assertEqual(
                    manifest["legs"]["before"]["programs"][0]["argv"][4],
                    environment["XLSX_BEFORE_REVISION"],
                )
                self.assertEqual(
                    manifest["legs"]["after"]["programs"][0]["argv"][4],
                    environment["XLSX_AFTER_REVISION"],
                )
                self.assertIn(
                    "tools/perf-baseline/src/bin/support/counting_allocator.rs",
                    manifest["shared"]["probe_paths"],
                )
                runner._validate_xlsx_raw_workload(manifest)
                manifest["validators"][0]["argv"][-1] = str(root / "other-receipt.json")
                with self.assertRaisesRegex(runner.RunnerError, "bind its reports"):
                    runner._validate_xlsx_raw_workload(manifest)
                os.environ["XLSX_PREPARATION_RECEIPT"] = str(root / "output" / "preparation.json")
                with self.assertRaisesRegex(runner.RunnerError, "outside the fresh output"):
                    runner.load_manifest(ROOT / "tools/native_qualification/manifests/w4-xlsx-raw.json")
                for leg in ("before", "after"):
                    for field in ("source_root", "target_dir"):
                        receipt = Path(manifest["legs"][leg][field]) / "preparation.json"
                        with self.subTest(leg=leg, field=field):
                            os.environ["XLSX_PREPARATION_RECEIPT"] = str(receipt)
                            reason = "outside source worktrees" if field == "source_root" else "pairwise"
                            with self.assertRaisesRegex(runner.RunnerError, reason):
                                runner.load_manifest(
                                    ROOT / "tools/native_qualification/manifests/w4-xlsx-raw.json"
                                )
            finally:
                for key, value in old.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value

    def test_manifest_cannot_weaken_fixed_limits(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-limits-") as directory:
            data = json.loads(
                (ROOT / "tools/native_qualification/manifests/legacy-w5.json").read_text(encoding="utf-8")
            )
            data["target_limit_bytes"] = 1
            path = Path(directory) / "manifest.json"
            path.write_text(json.dumps(data), encoding="utf-8")
            old = {key: os.environ.get(key) for key in ("NATIVE_QUAL_OUTPUT", "W5_BEFORE_SOURCE", "W5_AFTER_SOURCE", "W5_TARGET_ROOT", "LITCHI_COORD_BIN", "LITCHI_NATIVE_QUAL_ROOT", "W5_ANALYZER")}
            try:
                os.environ.update(
                    {
                        "NATIVE_QUAL_OUTPUT": str(Path(directory) / "output"),
                        "W5_BEFORE_SOURCE": str(ROOT),
                        "W5_AFTER_SOURCE": str(Path(directory) / "after"),
                        "W5_TARGET_ROOT": str(Path(directory) / "targets"),
                        "LITCHI_COORD_BIN": directory,
                        "LITCHI_NATIVE_QUAL_ROOT": str(ROOT),
                        "W5_ANALYZER": str(ROOT / "tools/native_qualification_runner.py"),
                    }
                )
                with self.assertRaisesRegex(runner.RunnerError, "fixed"):
                    runner.load_manifest(path)
            finally:
                for key, value in old.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value

    def test_pairwise_source_target_output_overlap_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-overlap-") as directory:
            data = json.loads(
                (ROOT / "tools/native_qualification/manifests/legacy-w5.json").read_text(encoding="utf-8")
            )
            data["output_dir"] = str(Path(directory) / "out")
            data["legs"]["before"]["source_root"] = str(Path(directory) / "source")
            data["legs"]["after"]["source_root"] = str(Path(directory) / "source" / "nested")
            data["legs"]["before"]["target_dir"] = str(Path(directory) / "target-before")
            data["legs"]["after"]["target_dir"] = str(Path(directory) / "target-after")
            data["shared"]["probe_paths"] = []
            data["path_prepend"] = [directory]
            data["legs"]["before"]["lock_paths"] = ["Cargo.lock"]
            data["legs"]["after"]["lock_paths"] = ["Cargo.lock"]
            data["validators"][0]["script"] = str(ROOT / "tools/native_qualification_runner.py")
            data["validators"][0]["argv"] = ["python3", str(ROOT / "tools/native_qualification_runner.py"), "{output_dir}"]
            path = Path(directory) / "manifest.json"
            path.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaisesRegex(runner.RunnerError, "pairwise"):
                runner.load_manifest(path)

    def test_cargo_output_override_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-cargo-") as directory:
            target = Path(directory) / "target"
            with self.assertRaisesRegex(runner.RunnerError, "output-directory"):
                runner._validate_cargo_paths(
                    ["cargo", "build", "--target-dir", str(Path(directory) / "other")],
                    cwd=ROOT,
                    source=ROOT,
                    target=target,
                )
            config_dir = Path(directory) / ".cargo"
            config_dir.mkdir()
            (config_dir / "config.toml").write_text(
                '[build]\ntarget-dir = "/tmp/escape-target"\n', encoding="utf-8"
            )
            with self.assertRaisesRegex(runner.RunnerError, "redirects target"):
                runner._validate_cargo_paths(["cargo", "build"], cwd=Path(directory), source=ROOT, target=target)

    def test_build_artifact_must_stay_in_its_leg_target(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-artifact-") as directory:
            root = Path(directory)
            target = root / "target"
            safe = runner.SafeRunner(
                {
                    "lane": "w5-legacy",
                    "output_dir": str(root / "output"),
                    "legs": {"before": {"target_dir": str(target)}},
                },
                ROOT / "tools/native_qualification_runner.py",
                "build",
                None,
            )
            with self.assertRaisesRegex(runner.RunnerError, "escapes its leg target"):
                safe._build_artifact_paths(
                    "before",
                    {"artifacts": [str(root / "outside")], "id": "escape"},
                    {},
                )

    def test_shared_runner_process_observer_is_inert_by_default(self) -> None:
        """Existing lanes do not opt into the DrawingML child hook."""

        with tempfile.TemporaryDirectory(prefix="native-qual-observer-") as directory:
            root = Path(directory)
            safe = runner.SafeRunner(
                {
                    "lane": "w5-legacy",
                    "output_dir": str(root / "output"),
                    "legs": {"before": {"target_dir": str(root / "target")}},
                },
                ROOT / "tools/native_qualification_runner.py",
                "build",
                None,
            )
            self.assertIsNone(safe.process_observer)

    def test_environment_uses_fixed_decimal_memory_cap(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-env-") as directory:
            root = Path(directory)
            safe = runner.SafeRunner(
                {
                    "lane": "w5-legacy",
                    "output_dir": str(root / "output"),
                    "legs": {"before": {"target_dir": str(root / "target")}},
                },
                ROOT / "tools/native_qualification_runner.py",
                "build",
                None,
            )
            env = safe._environment({}, {"target_dir": root / "target"})
            self.assertEqual(env["LITCHI_MEM_MAX"], "20G")
            self.assertNotIn("CARGO_HOME", env)
            self.assertNotIn("CARGO_CONFIG", env)

    def test_environment_replaces_inherited_target_dir(self) -> None:
        # Session env.sh files and the coordinator gate export
        # CARGO_TARGET_DIR; that is not a manifest override (record 0365).
        with tempfile.TemporaryDirectory(prefix="native-qual-env-") as directory:
            root = Path(directory)
            safe = runner.SafeRunner(
                {
                    "lane": "w5-legacy",
                    "output_dir": str(root / "output"),
                    "legs": {"before": {"target_dir": str(root / "target")}},
                },
                ROOT / "tools/native_qualification_runner.py",
                "build",
                None,
            )
            with mock.patch.dict(os.environ, {"CARGO_TARGET_DIR": str(root / "elsewhere")}):
                env = safe._environment({}, {"target_dir": root / "target"})
            self.assertEqual(env["CARGO_TARGET_DIR"], str(root / "target"))
            with self.assertRaisesRegex(runner.RunnerError, "cannot override"):
                safe._environment(
                    {"environment": {"CARGO_TARGET_DIR": str(root / "elsewhere")}},
                    {"target_dir": root / "target"},
                )

    def test_build_output_must_be_fresh(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-stale-") as directory:
            root = Path(directory)
            output = root / "output"
            output.mkdir()
            (output / "source-before-1.json").write_text("stale", encoding="utf-8")
            with self.assertRaisesRegex(runner.RunnerError, "fresh empty"):
                runner.SafeRunner(
                    {
                        "lane": "w5-legacy",
                        "output_dir": str(output),
                        "legs": {},
                    },
                    ROOT / "tools/native_qualification_runner.py",
                    "build",
                    None,
                )

    def test_build_handoff_rejects_stale_analyzer_inputs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-handoff-") as directory:
            root = Path(directory)
            output = root / "output"
            output.mkdir()
            (output / "source-before-1.json").write_text("stale", encoding="utf-8")
            safe = runner.SafeRunner(
                {
                    "lane": "w5-legacy",
                    "output_dir": str(output),
                    "legs": {},
                },
                ROOT / "tools/native_qualification_runner.py",
                "abba",
                None,
            )
            with self.assertRaisesRegex(runner.RunnerError, "stale entries"):
                safe._validate_build_handoff(set())

    def test_w5_adapter_preserves_valid_analysis_in_qualified_summary(self) -> None:
        analysis = {
            "source": [{"scenario": "doc/source_set", "before_us": 1.0, "after_us": 2.0, "ratio": 2.0}],
            "ops": [{"scenario": "xls/get", "before_us": 3.0, "after_us": 4.0, "ratio": 1.333}],
            "doc_batch1k_unified_over_leaf": {},
            "max_rss_kb": {},
        }
        self.assertEqual(w5_adapter.validate_analysis(analysis), analysis)
        with self.assertRaisesRegex(w5_adapter.AdapterError, "missing required"):
            w5_adapter.validate_analysis({"source": []})
        with tempfile.TemporaryDirectory(prefix="native-qual-adapter-") as directory:
            root = Path(directory)
            analyzer = root / "analyze.py"
            output = root / "output"
            output.mkdir()
            analyzer.write_text(
                "import json\n"
                f"print(json.dumps({analysis!r}))\n",
                encoding="utf-8",
            )
            self.assertEqual(w5_adapter.main([str(analyzer), str(output)]), 0)
            qualified = json.loads((output / "summary.json").read_text(encoding="utf-8"))
            self.assertEqual(qualified["schema"], w5_adapter.SUMMARY_SCHEMA)
            self.assertEqual(qualified["lane"], "w5-legacy")
            self.assertEqual(qualified["analysis"], analysis)

    def test_w5_source_overlay_translates_common_api_and_retains_doc_controls(self) -> None:
        candidate = (ROOT / "crates/litchi-ops/examples/legacy_w5_source_probe.rs").read_text(encoding="utf-8")
        routes = (ROOT / "docs/unified-ops/results/0359a/routes_before.rs").read_text(encoding="utf-8")
        adapted = source_prep._adapt_w5_probe(candidate, routes)
        self.assertIn("SourceSnapshot::from_bytes", adapted)
        self.assertIn("source_backed::Package::from_bytes", candidate)
        self.assertEqual(adapted.count("mod routes {"), 1)
        self.assertIn("PLAN_ONLY_ROUTE: bool = true", adapted)
        self.assertIn(".edit_with_limits(&limits,", adapted)
        self.assertIn("black_box(commit.file);", adapted)
        self.assertNotIn("litchi_ops::EditOptions", adapted)
        self.assertNotIn("commit.file()", adapted)
        self.assertEqual(adapted.count("pub(super) fn doc_batch_leaf_write"), 1)
        self.assertEqual(adapted.count("pub(super) fn doc_leaf_write"), 1)
        with self.assertRaisesRegex(source_prep.PreparationError, "doc_leaf_write"):
            source_prep._adapt_w5_probe(
                candidate.replace(
                    "pub(super) fn doc_leaf_write",
                    "pub(super) fn removed_doc_leaf_write",
                    1,
                ),
                routes,
            )

    def test_w4b_crud_overlay_is_only_the_recorded_trait_import(self) -> None:
        baseline = "use litchi::sheet::{WorkbookTrait, Worksheet};\nfn main() {}\n"
        adapted = source_prep._adapt_w4b_crud(baseline)
        self.assertEqual(
            adapted,
            "use litchi::sheet::{WorkbookTrait, Worksheet};\n"
            "use litchi_core::sheet::Cell as _;\nfn main() {}\n",
        )
        with self.assertRaisesRegex(source_prep.PreparationError, "already contains"):
            source_prep._adapt_w4b_crud(
                baseline.replace(
                    "fn main", "use litchi_core::sheet::Cell as _;\nfn main"
                )
            )

    def test_w4b_probe_overlay_translates_exactly_three_lazy_resolutions(self) -> None:
        candidate = "".join(
            [
                "fn first() {\n",
                "    let edit = package\n",
                "        .edit_cell_values(0_usize)\n",
                "        .expect(\"cell edit\")\n",
                "        .resolve()\n",
                "        .expect(\"worksheet resolution\")\n",
                "        .expect(\"worksheet\");\n",
                "}\n",
                "fn second() {\n",
                "        let edit = package\n",
                "            .edit_cell_values(0_usize)\n",
                "            .expect(\"cell edit\")\n",
                "            .resolve()\n",
                "            .expect(\"worksheet resolution\")\n",
                "            .expect(\"worksheet\");\n",
                "}\n",
                "fn third() {\n",
                "        let edit = package\n",
                "            .edit_cell_values(0_usize)\n",
                "            .expect(\"cell edit\")\n",
                "            .resolve()\n",
                "            .expect(\"worksheet resolution\")\n",
                "            .expect(\"worksheet\");\n",
                "}\n",
                "fn commit_result(commit: litchi_ops::transaction::Commit) {\n",
                "    black_box(commit.into_file());\n",
                "}\n",
            ]
        )
        adapted = source_prep._adapt_w4b_probe(candidate)
        self.assertEqual(adapted.count(".resolve()"), 0)
        self.assertEqual(adapted.count('.expect("worksheet");'), 3)
        self.assertNotIn("commit.into_file()", adapted)
        self.assertIn("commit.file", adapted)
        with self.assertRaisesRegex(source_prep.PreparationError, "exactly three"):
            source_prep._adapt_w4b_probe(candidate.replace(".resolve()", "", 1))

    def test_source_preparation_requires_external_disjoint_provenance(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-prep-paths-") as directory:
            root = Path(directory)
            before = root / "before"
            after = root / "after"
            before.mkdir()
            after.mkdir()
            with self.assertRaisesRegex(source_prep.PreparationError, "pairwise disjoint"):
                source_prep._require_disjoint_sources(before, after, before / "provenance.json")

    def test_w4b_preparation_transforms_existing_baseline_crud(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-prep-crud-") as directory:
            root = Path(directory)
            before = root / "before"
            after = root / "after"
            baseline_crud = "use litchi::sheet::{WorkbookTrait, Worksheet};\nfn main() {}\n"
            candidate_crud = baseline_crud + "// candidate API body\n"
            candidate_probe = "".join(
                [
                    "fn first() {\n",
                    "    let edit = package\n",
                    "        .edit_cell_values(0_usize)\n",
                    "        .expect(\"cell edit\")\n",
                    "        .resolve()\n",
                    "        .expect(\"worksheet resolution\")\n",
                    "        .expect(\"worksheet\");\n",
                    "}\n",
                    "fn second() {\n",
                    "    let edit = package\n",
                    "        .edit_cell_values(0_usize)\n",
                    "        .expect(\"cell edit\")\n",
                    "        .resolve()\n",
                    "        .expect(\"worksheet resolution\")\n",
                    "        .expect(\"worksheet\");\n",
                    "}\n",
                    "fn third() {\n",
                    "    let edit = package\n",
                    "        .edit_cell_values(0_usize)\n",
                    "        .expect(\"cell edit\")\n",
                    "        .resolve()\n",
                    "        .expect(\"worksheet resolution\")\n",
                    "        .expect(\"worksheet\");\n",
                    "}\n",
                    "fn commit_result(commit: litchi_ops::transaction::Commit) {\n",
                    "    black_box(commit.into_file());\n",
                    "}\n",
                ]
            )

            def git(repo: Path, *args: str) -> str:
                result = subprocess.run(
                    ["git", "-C", str(repo), *args],
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                )
                return result.stdout.strip()

            for repo, crud in ((before, baseline_crud), (after, candidate_crud)):
                repo.mkdir()
                git(repo, "init", "--quiet")
                git(repo, "config", "user.email", "native-qualification@test.invalid")
                git(repo, "config", "user.name", "native qualification test")
                crud_path = repo / source_prep.W4B_CRUD
                crud_path.parent.mkdir(parents=True)
                crud_path.write_text(crud, encoding="utf-8")
                probe_path = repo / source_prep.W4B_PROBE
                probe_path.parent.mkdir(parents=True, exist_ok=True)
                if repo == after:
                    probe_path.write_text(candidate_probe, encoding="utf-8")
                git(repo, "add", ".")
                git(repo, "commit", "--quiet", "-m", "source pair")

            baseline_tree = git(before, "rev-parse", "HEAD^{tree}")
            with mock.patch.object(source_prep, "W4B_BASE_TREE", baseline_tree):
                prepared = source_prep._prepare_w4b(before, after)

            adapted_probe = (before / source_prep.W4B_PROBE).read_text(encoding="utf-8")
            self.assertEqual(adapted_probe.count(".resolve()"), 0)
            self.assertEqual(adapted_probe.count('.expect("worksheet");'), 3)
            self.assertEqual(
                prepared["overlays"][0]["adaptation"],
                "recorded three-call historical selector-resolution and commit-field translations",
            )
            adapted = (before / source_prep.W4B_CRUD).read_text(encoding="utf-8")
            self.assertEqual(adapted, source_prep._adapt_w4b_crud(baseline_crud))
            crud_overlay = prepared["overlays"][1]
            self.assertEqual(
                crud_overlay["before_sha256"],
                hashlib.sha256(baseline_crud.encode()).hexdigest(),
            )
            self.assertEqual(
                crud_overlay["after_sha256"],
                source_prep._sha256(before / source_prep.W4B_CRUD)[0],
            )
            self.assertEqual(
                crud_overlay["adaptation"],
                "recorded one-line litchi_core::sheet::Cell import",
            )

            unrelated = root / "unrelated.rs"
            unrelated.write_text("changed\n", encoding="utf-8")
            with self.assertRaisesRegex(source_prep.PreparationError, "changed unexpectedly"):
                source_prep._transform_text_exact(
                    unrelated,
                    baseline_crud,
                    source_prep._adapt_w4b_crud(baseline_crud),
                    after / source_prep.W4B_CRUD,
                )

    def test_symlink_artifact_is_never_hashed(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-link-") as directory:
            root = Path(directory)
            real = root / "real"
            real.write_text("artifact", encoding="utf-8")
            link = root / "link"
            link.symlink_to(real)
            with self.assertRaisesRegex(runner.RunnerError, "symlinked"):
                runner._sha256_file(link)

    def test_dangling_log_link_is_not_recorded_as_empty(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-dangling-") as directory:
            root = Path(directory)
            link = root / "log"
            link.symlink_to(root / "missing")
            with self.assertRaisesRegex(runner.RunnerError, "symlinked log"):
                runner._json_hash(link)

    def test_receipt_phase_is_checked_before_any_timing(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-receipt-") as directory:
            root = Path(directory)
            manifest = {
                "schema": runner.SCHEMA,
                "lane": "w5-legacy",
                "output_dir": str(root / "output"),
                "abba_order": list(runner.ABBA_ORDER),
                "legs": {
                    "before": {"source_root": str(ROOT), "target_dir": str(root / "before")},
                    "after": {"source_root": str(ROOT), "target_dir": str(root / "after")},
                },
            }
            safe = runner.SafeRunner(manifest, ROOT / "tools/native_qualification_runner.py", "abba", None)
            with self.assertRaisesRegex(runner.RunnerError, "build-phase"):
                safe._validate_build_receipt({"schema": runner.SCHEMA, "status": "passed", "phase": "abba", "lane": "w5-legacy"})

    def test_failed_child_is_journaled_before_error(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-child-") as directory:
            root = Path(directory)
            helper = root / "fail.py"
            helper.write_text("raise SystemExit(7)\n", encoding="utf-8")
            output = root / "output"
            target = root / "target"
            class Observer:
                def __init__(self) -> None:
                    self.started: list[tuple[str, int]] = []
                    self.finished: list[tuple[str, int | None]] = []

                def process_started(self, **kwargs):
                    event_id = kwargs["event_id"]
                    process = kwargs["process"]
                    self.started.append((event_id, process.pid))
                    return {"event_id": event_id, "pid": process.pid}

                def process_finished(self, token, *, returncode):
                    self.finished.append((token["event_id"], returncode))
                    return {
                        "schema": "test.process-observer.v1",
                        "event_id": token["event_id"],
                    }

            observer = Observer()
            manifest = {
                "schema": runner.SCHEMA,
                "output_dir": str(output),
                "abba_order": list(runner.ABBA_ORDER),
                "shared": {"probe_paths": [str(ROOT / "tools/native_qualification_runner.py")]},
                "legs": {
                    "before": {
                        "source_root": str(ROOT),
                        "target_dir": str(target),
                        "lock_paths": ["tools/native_qualification_runner.py"],
                        "allow_dirty": True,
                    },
                    "after": {
                        "source_root": str(ROOT),
                        "target_dir": str(root / "after-target"),
                        "lock_paths": ["tools/native_qualification_runner.py"],
                        "allow_dirty": True,
                    },
                },
            }
            safe = runner.SafeRunner(
                manifest,
                ROOT / "tools/native_qualification_runner.py",
                "build",
                None,
                process_observer=observer,
            )
            # The real host currently has unrelated wave builders; this unit
            # test isolates child-journaling behavior from the host guard.
            safe.guard.require_quiet = lambda: None
            safe.guard.offenders = lambda _allowed_root=None: []
            identity = safe._identity("before")
            with self.assertRaises(runner.RunnerError):
                safe._run_command(
                    event_id="failed-child",
                    leg="before",
                    argv=[sys.executable, str(helper)],
                    cwd=ROOT,
                    env=os.environ.copy(),
                    identity=identity,
                )
            journal = output / "events-build.jsonl"
            self.assertTrue(journal.is_file())
            event = json.loads(journal.read_text(encoding="utf-8").splitlines()[0])
            self.assertEqual(event["returncode"], 7)
            self.assertEqual(event["event_id"], "failed-child")
            self.assertEqual(observer.started[0][0], "failed-child")
            self.assertEqual(observer.finished, [("failed-child", 7)])
            self.assertEqual(event["process_observation"]["event_id"], "failed-child")

    def test_control_handoff_is_compared_before_measurement_and_not_measured(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-control-") as directory:
            root = Path(directory)
            helper = root / "control.py"
            helper.write_text(
                "import json, sys\n"
                "output = sys.argv[sys.argv.index('--json') + 1]\n"
                "with open(output, 'w', encoding='utf-8') as stream:\n"
                "    json.dump({'measured': False, 'identity': 'same'}, stream)\n",
                encoding="utf-8",
            )
            output = root / "output"
            manifest = {
                "schema": runner.SCHEMA,
                "lane": "w5-legacy",
                "output_dir": str(output),
                "abba_order": list(runner.ABBA_ORDER),
                "shared": {},
                "legs": {
                    "before": {
                        "source_root": str(ROOT),
                        "target_dir": str(root / "before"),
                        "lock_paths": ["tools/perf-baseline/Cargo.lock"],
                        "allow_dirty": True,
                    },
                    "after": {
                        "source_root": str(ROOT),
                        "target_dir": str(root / "after"),
                        "lock_paths": ["tools/perf-baseline/Cargo.lock"],
                        "allow_dirty": True,
                    },
                },
                "controls": [{
                    "id": "test-control",
                    "compare": ["identity"],
                    "programs": {
                        leg: {
                            "binary": str(Path(sys.executable).resolve()),
                            "argv": [str(helper), "--json", "{output}"],
                            "outputs": [f"{{output_dir}}/controls/{leg}.json"],
                            "stdout": f"{{output_dir}}/controls/{leg}.stdout",
                            "stderr": f"{{output_dir}}/controls/{leg}.stderr",
                            "require_stdout": False,
                        }
                        for leg in ("before", "after")
                    },
                }],
            }
            from native_qualification.drawingml_raw_paired_runner import ExternalResourceObserver

            observer = ExternalResourceObserver(output)
            safe = runner.SafeRunner(
                manifest,
                ROOT / "tools/native_qualification_runner.py",
                "abba",
                None,
                process_observer=observer,
            )
            safe.guard.require_quiet = lambda: None
            safe.guard.offenders = lambda _allowed_root=None: []
            identities = {leg: safe._identity(leg) for leg in ("before", "after")}
            safe._run_controls(identities)
            events = [json.loads(line) for line in (output / "events-abba.jsonl").read_text().splitlines()]
            self.assertEqual([event["event_kind"] for event in events], ["control", "control"])
            self.assertEqual([event["measured"] for event in events], [False, False])
            self.assertTrue(all("process_observation" not in event for event in events))
            self.assertFalse(observer.path.exists())

    def _overlap_fixture(self, root: Path, *, max_jobs: int = 4, expires: str | None = None):
        return self._overlap_fixture_with_sources(
            root,
            max_jobs=max_jobs,
            expires=expires,
            source_roots=(ROOT, ROOT),
        )

    def _overlap_fixture_with_sources(
        self,
        root: Path,
        *,
        max_jobs: int = 4,
        expires: str | None = None,
        source_roots: tuple[Path, Path],
    ):
        permit = root / "permit.json"
        manifest_path = root / "manifest.json"
        lease_dir = root / "leases"
        manifest = {
            "schema": runner.SCHEMA,
            "lane": "w5-legacy",
            "output_dir": str(root / "output"),
            "build_execution": {"class": "compile_only", "permit_path": str(permit)},
            "legs": {
                "before": {
                    "source_root": str(source_roots[0]),
                    "target_dir": str(root / "target-before"),
                    "build": [{"id": "build", "argv": ["cargo", "build"], "artifacts": ["{target_dir}/probe"]}],
                },
                "after": {
                    "source_root": str(source_roots[1]),
                    "target_dir": str(root / "target-after"),
                    "build": [{"id": "build", "argv": ["cargo", "build"], "artifacts": ["{target_dir}/probe"]}],
                },
            },
        }
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        manifest_hash = hashlib.sha256(manifest_path.read_bytes()).hexdigest()
        expires = expires or (
            runner._dt.datetime.now(runner._dt.timezone.utc) + runner._dt.timedelta(minutes=5)
        ).isoformat()
        permit.write_text(
            json.dumps(
                {
                    "schema": runner.BUILD_PERMIT_SCHEMA,
                    "grant_id": "test-grant",
                    "job_group": "test-group",
                    "expires_utc": expires,
                    "max_jobs": max_jobs,
                    "allowed_lanes": ["w5-legacy"],
                    "allowed_manifest_sha256": [manifest_hash],
                    "allowed_command_classes": ["cargo-build", "cargo-check", "cargo-clippy", "cargo-test"],
                    "manifest_bindings": [{
                        "path": str(manifest_path),
                        "sha256": manifest_hash,
                        "lane": "w5-legacy",
                        "source_roots": [str(source_roots[0].resolve()), str(source_roots[1].resolve())],
                        "target_dirs": [str(root / "target-before"), str(root / "target-after")],
                    }],
                    "lease_dir": str(lease_dir),
                }
            ),
            encoding="utf-8",
        )
        return manifest, manifest_path, permit, lease_dir

    @staticmethod
    def _make_fake_cargo_source(
        root: Path, *, sleep_seconds: float | None = None, hold: bool = False
    ) -> Path:
        """Create a tiny committed source tree whose cargo symlink is bash.

        Invoking ``cargo build`` with cwd at this tree makes bash execute the
        tracked ``build`` script while /proc still exposes the exact Cargo
        argv that the overlap protocol classifies.  This exercises the real
        SafeRunner subprocess/lease path without invoking Cargo or native
        code.
        """

        if (sleep_seconds is None) == (not hold):
            raise ValueError("choose exactly one fake-cargo wait mode")
        wait = "sleep infinity" if hold else f"sleep {sleep_seconds}"

        root.mkdir(parents=True, exist_ok=False)
        (root / "Cargo.lock").write_text("# fake lock for runner tests\n", encoding="utf-8")
        build = root / "build"
        build.write_text(
            "#!/bin/bash\n"
            "set -eu\n"
            "trap 'exit 143' TERM\n"
            "mkdir -p \"$CARGO_TARGET_DIR\"\n"
            "printf '#!/bin/sh\\nexit 0\\n' > \"$CARGO_TARGET_DIR/probe\"\n"
            "chmod +x \"$CARGO_TARGET_DIR/probe\"\n"
            f"{wait}\n",
            encoding="utf-8",
        )
        build.chmod(0o755)
        (root / "cargo").symlink_to("/bin/bash")
        subprocess.run(["git", "init", "--quiet", str(root)], check=True, stdout=subprocess.DEVNULL)
        subprocess.run(
            ["git", "-C", str(root), "add", "Cargo.lock", "build", "cargo"],
            check=True,
            stdout=subprocess.DEVNULL,
        )
        subprocess.run(
            [
                "git",
                "-C",
                str(root),
                "-c",
                "user.name=Native runner test",
                "-c",
                "user.email=native-runner-test@example.invalid",
                "commit",
                "--quiet",
                "-m",
                "fake cargo source",
            ],
            check=True,
            stdout=subprocess.DEVNULL,
        )
        return root / "cargo"

    @staticmethod
    def _peer_entry(pid: int, starttime: str, argv: list[str]) -> dict[str, object]:
        argv = runner.normalize_cargo_argv(argv)
        return {
            "pid": pid,
            "ppid": pid,
            "pgrp": pid,
            "starttime": starttime,
            "comm": "cargo",
            "exe": "/usr/bin/cargo",
            "cwd": str(ROOT),
            "cmdline": argv,
        }

    def _write_peer_lease(
        self,
        lease: runner.BuildOverlapLease,
        path: Path,
        *,
        pid: int = 999991,
        starttime: str = "peer-start",
        argv: list[str] | None = None,
    ) -> None:
        value = {
            "schema": runner.BUILD_PERMIT_SCHEMA,
            "status": "active",
            "lease_id": "peer-lease",
            "grant_id": lease.grant_id,
            "job_group": lease.job_group,
            "permit_sha256": lease.permit_hash,
            "root_pid": pid,
            "root_starttime": starttime,
            "controller_pid": pid,
            "controller_starttime": starttime,
            "lane": "w5-legacy",
            "manifest_path": str(lease.manifest_path),
            "manifest_sha256": hashlib.sha256(lease.manifest_path.read_bytes()).hexdigest(),
            "command_class": "cargo-build",
            "allowed_command_classes": ["cargo-build"],
            "phase": "build",
            "phase_id": "build-before-build",
            "phase_argv": ["cargo", "build", "--offline", "--locked"],
            "phase_cwd": str(ROOT),
            "phase_target_dir": str(Path(lease.manifest["legs"]["before"]["target_dir"]).resolve()),
            "source_roots": [str(ROOT), str(lease.manifest["legs"]["after"]["source_root"])],
            "target_dirs": [
                str(Path(lease.manifest["legs"]["before"]["target_dir"]).resolve()),
                str(Path(lease.manifest["legs"]["after"]["target_dir"]).resolve()),
            ],
            "created_utc": runner._now(),
            "expires_utc": lease.permit["expires_utc"],
        }
        path.write_text(json.dumps(value), encoding="utf-8")

    def test_compile_overlap_default_guard_rejects_unregistered_builder(self) -> None:
        entry = self._peer_entry(999991, "peer-start", ["cargo", "build"])
        with mock.patch.object(runner, "_proc_snapshot", return_value={999991: entry}):
            with self.assertRaisesRegex(runner.RunnerError, "unrelated Litchi"):
                runner.ProcessGuard().require_quiet()

    def test_proc_monitor_distinguishes_gone_process_from_scan_failure(self) -> None:
        with mock.patch.object(
            Path,
            "read_text",
            side_effect=FileNotFoundError(runner.errno.ENOENT, "gone"),
        ):
            self.assertIsNone(runner._read_proc(991981))
        with mock.patch.object(
            Path,
            "read_text",
            side_effect=PermissionError(runner.errno.EACCES, "denied"),
        ):
            with self.assertRaisesRegex(runner.ProcMonitorError, "cannot inspect /proc/991982"):
                runner._read_proc(991982)
        with mock.patch.object(
            Path,
            "iterdir",
            side_effect=PermissionError(runner.errno.EACCES, "denied"),
        ):
            with self.assertRaisesRegex(runner.ProcMonitorError, "cannot enumerate /proc"):
                runner._proc_snapshot()

    def test_owned_record_rejects_unproved_leader_subtree(self) -> None:
        def entry(pid: int, ppid: int, starttime: str) -> dict[str, object]:
            return {
                "pid": pid,
                "ppid": ppid,
                "pgrp": pid,
                "starttime": starttime,
                "comm": "cargo",
                "exe": "/usr/bin/cargo",
                "cwd": str(ROOT),
                "cmdline": ["cargo", "build"],
            }

        snapshot = {
            1: entry(1, 0, "runner"),
            10: entry(10, 1, "replacement-leader"),
            20: entry(20, 10, "unproved-child"),
        }
        owned: dict[int, dict[str, object]] = {}
        with self.assertRaisesRegex(runner.RunnerError, "cannot establish owned leader identity"):
            runner.SafeRunner._record_owned_processes(
                snapshot,
                process_pid=10,
                runner_pid=1,
                preexisting={},
                owned=owned,
                pid_reuse=[],
            )
        self.assertNotIn(20, owned)

    def test_compile_only_manifest_rejects_non_build_command(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-command-") as directory:
            root = Path(directory)
            manifest, _manifest_path, _permit, _lease_dir = self._overlap_fixture(root)
            manifest["legs"]["after"]["build"][0]["argv"] = ["cargo", "check"]
            with self.assertRaisesRegex(runner.RunnerError, "only cargo build"):
                runner._validate_build_execution(
                    manifest,
                    output=root / "output",
                    source_roots=[ROOT, ROOT],
                )

    def test_compile_overlap_lease_rejects_stale_hash_pid_reuse_expiry_and_excess(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-negative-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root, max_jobs=1)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            peer_path = lease_dir / "peer.json"
            self._write_peer_lease(lease, peer_path)
            peer = self._peer_entry(999991, "peer-start", ["cargo", "build"])
            with mock.patch.object(runner, "_proc_snapshot", return_value={os.getpid(): own, 999991: peer}):
                with self.assertRaisesRegex(runner.RunnerError, "exceeds max_jobs"):
                    lease.ensure_valid()
            peer_path.unlink()
            self._write_peer_lease(lease, peer_path)
            with mock.patch.object(runner, "_proc_snapshot", return_value={os.getpid(): own}):
                with self.assertRaisesRegex(runner.RunnerError, "process is gone"):
                    lease.peer_snapshots({os.getpid(): own})
            peer_path.unlink()
            with mock.patch.object(runner, "_proc_snapshot", return_value={os.getpid(): own}):
                permit_path = Path(lease.permit_path)
                permit_path.write_text(permit_path.read_text(encoding="utf-8") + "\n", encoding="utf-8")
                with self.assertRaisesRegex(runner.RunnerError, "permit changed"):
                    lease.ensure_valid()
            lease.release()

    def test_compile_overlap_lease_allows_only_declared_live_peer(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-peer-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            peer_path = lease_dir / "peer.json"
            self._write_peer_lease(lease, peer_path)
            peer = self._peer_entry(999991, "peer-start", ["cargo", "build"])
            snapshot = {os.getpid(): own, 999991: peer}
            with mock.patch.object(runner, "_proc_snapshot", return_value=snapshot):
                self.assertEqual(lease.peer_snapshots(snapshot)[0]["root_pid"], 999991)
                self.assertEqual(lease.allowed_peer_processes(snapshot), {999991})
                self.assertEqual(runner.ProcessGuard().offenders(overlap=lease), [])
            lease.release()

        with tempfile.TemporaryDirectory(prefix="native-build-overlap-expired-") as directory:
            root = Path(directory)
            expired = (runner._dt.datetime.now(runner._dt.timezone.utc) - runner._dt.timedelta(seconds=1)).isoformat()
            manifest, manifest_path, _permit, _lease_dir = self._overlap_fixture(root, expires=expired)
            with self.assertRaisesRegex(runner.RunnerError, "expired"):
                runner.BuildOverlapLease(manifest, manifest_path)

    def test_compile_overlap_lease_rejects_bound_peer_reparented_to_init(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-parent-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            peer_path = lease_dir / "peer.json"
            self._write_peer_lease(lease, peer_path)
            peer = self._peer_entry(999991, "peer-start", ["cargo", "build"])
            peer["ppid"] = 1
            snapshot = {os.getpid(): own, 999991: peer}
            with mock.patch.object(runner, "_proc_snapshot", return_value=snapshot):
                with self.assertRaisesRegex(runner.RunnerError, "parent changed"):
                    lease.peer_snapshots(snapshot)
            lease.release()

    def test_compile_overlap_lease_rejects_tampered_peer_manifest_bindings(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-binding-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            peer_path = lease_dir / "peer.json"
            self._write_peer_lease(lease, peer_path)
            peer = self._peer_entry(999991, "peer-start", ["cargo", "build"])
            original = json.loads(peer_path.read_text(encoding="utf-8"))
            mutations = (
                ("manifest_path", str(root / "unadmitted-manifest.json"), "manifest binding is not admitted"),
                ("source_roots", [str(ROOT), str(root / "wrong-source")], "source_roots differ"),
                ("target_dirs", [str(root / "wrong-before"), str(root / "target-after")], "target_dirs differ"),
            )
            for key, replacement, reason in mutations:
                value = dict(original)
                value[key] = replacement
                peer_path.write_text(json.dumps(value), encoding="utf-8")
                snapshot = {os.getpid(): own, 999991: peer}
                with mock.patch.object(runner, "_proc_snapshot", return_value=snapshot):
                    with self.assertRaisesRegex(runner.RunnerError, reason):
                        lease.peer_snapshots(snapshot)
            lease.release()

    def test_compile_overlap_lease_rejects_phase_argv_redirection(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-phase-path-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            peer_path = lease_dir / "peer.json"
            original = json.loads(
                json.dumps(
                    {
                        "manifest_path": str(lease.manifest_path),
                        "manifest_sha256": lease.manifest_hash,
                    }
                )
            )
            self._write_peer_lease(lease, peer_path)
            base = json.loads(peer_path.read_text(encoding="utf-8"))
            for extra in (
                ["--manifest-path", str(root / "outside" / "Cargo.toml")],
                ["--target-dir", str(root / "outside-target")],
            ):
                value = dict(base)
                phase_argv = runner.normalize_cargo_argv(["cargo", "build", *extra])
                value["phase_argv"] = phase_argv
                peer_path.write_text(json.dumps(value), encoding="utf-8")
                peer = self._peer_entry(999991, "peer-start", phase_argv)
                snapshot = {os.getpid(): own, 999991: peer}
                with mock.patch.object(runner, "_proc_snapshot", return_value=snapshot):
                    with self.assertRaisesRegex(runner.RunnerError, "phase differs"):
                        lease.peer_snapshots(snapshot)
            self.assertEqual(original["manifest_path"], str(lease.manifest_path))
            lease.release()

    def test_compile_overlap_lease_rejects_peer_expiry_after_permit(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-expiry-binding-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            peer_path = lease_dir / "peer.json"
            self._write_peer_lease(lease, peer_path)
            value = json.loads(peer_path.read_text(encoding="utf-8"))
            value["expires_utc"] = (
                lease.expires + runner._dt.timedelta(seconds=1)
            ).isoformat()
            peer_path.write_text(json.dumps(value), encoding="utf-8")
            peer = self._peer_entry(999991, "peer-start", ["cargo", "build"])
            snapshot = {os.getpid(): own, 999991: peer}
            with mock.patch.object(runner, "_proc_snapshot", return_value=snapshot):
                with self.assertRaisesRegex(runner.RunnerError, "expiry differs"):
                    lease.peer_snapshots(snapshot)
            lease.release()

    def test_compile_overlap_lease_rejects_expired_peer_record(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-expired-record-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            expired = (runner._dt.datetime.now(runner._dt.timezone.utc) - runner._dt.timedelta(seconds=1)).isoformat()
            lease.permit["expires_utc"] = expired
            peer_path = lease_dir / "peer.json"
            self._write_peer_lease(lease, peer_path)
            peer = self._peer_entry(999991, "peer-start", ["cargo", "build"])
            snapshot = {os.getpid(): own, 999991: peer}
            with mock.patch.object(runner, "_proc_snapshot", return_value=snapshot):
                with self.assertRaisesRegex(runner.RunnerError, "active build lease is expired"):
                    lease.peer_snapshots(snapshot)
            lease.release()

    def test_compile_overlap_lease_rejects_pid_reuse_and_disallowed_phase(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-phase-") as directory:
            root = Path(directory)
            manifest, manifest_path, _permit, lease_dir = self._overlap_fixture(root)
            lease = runner.BuildOverlapLease(manifest, manifest_path)
            lease.acquire()
            own = runner._read_proc(os.getpid())
            assert own is not None
            peer_path = lease_dir / "peer.json"
            self._write_peer_lease(lease, peer_path)
            for starttime, argv, reason in (
                ("other-start", ["cargo", "build"], "PID was reused"),
                ("peer-start", ["cargo", "doc"], "phase changed"),
                ("peer-start", ["cargo", "check"], "phase changed"),
            ):
                peer = self._peer_entry(999991, starttime, argv)
                with mock.patch.object(runner, "_proc_snapshot", return_value={os.getpid(): own, 999991: peer}):
                    with self.assertRaisesRegex(runner.RunnerError, reason):
                        lease.peer_snapshots({os.getpid(): own, 999991: peer})
            lease.release()

    def test_control_false_does_not_select_compile_overlap_guard(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-control-") as directory:
            root = Path(directory)
            safe = runner.SafeRunner(
                {
                    "lane": "w5-legacy",
                    "output_dir": str(root / "output"),
                    "build_execution": {"class": "compile_only", "permit_path": str(root / "permit.json")},
                    "legs": {
                        "before": {
                            "source_root": str(ROOT),
                            "target_dir": str(root / "target"),
                            "build": [{"id": "build", "argv": ["cargo", "build"], "artifacts": ["{target_dir}/probe"]}],
                        }
                    },
                },
                ROOT / "tools/native_qualification_runner.py",
                "abba",
                None,
            )
            sentinel = object()
            safe.overlap_lease = sentinel  # type: ignore[assignment]
            self.assertEqual(safe._event_execution_class("control"), "exclusive")
            self.assertEqual(safe._event_execution_class("measurement"), "exclusive")
            self.assertEqual(safe._event_execution_class("build"), "compile_only")

    def test_expired_compile_lease_reaps_child_and_writes_failed_receipt(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-expiry-child-") as directory:
            root = Path(directory)
            source = self._make_fake_cargo_source(root / "source", hold=True)
            permit_expiry = (
                runner._dt.datetime.now(runner._dt.timezone.utc)
                + runner._dt.timedelta(minutes=5)
            )
            manifest, manifest_path, permit, _lease_dir = self._overlap_fixture_with_sources(
                root,
                expires=permit_expiry.isoformat(),
                source_roots=(source.parent, source.parent),
            )
            for leg in manifest["legs"].values():
                leg["build"][0]["argv"] = [str(source), "build"]
                leg["allow_dirty"] = True
            manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            permit_value = json.loads(permit.read_text(encoding="utf-8"))
            manifest_hash = hashlib.sha256(manifest_path.read_bytes()).hexdigest()
            permit_value["allowed_manifest_sha256"] = [manifest_hash]
            permit_value["manifest_bindings"][0]["sha256"] = manifest_hash
            permit.write_text(json.dumps(permit_value), encoding="utf-8")
            # The permit expiry is comfortably beyond setup.  The controlled
            # clock below advances only after the child is bound, so this
            # test never relies on startup latency or arbitrary sleep time.
            permit_value["expires_utc"] = permit_expiry.isoformat()
            permit.write_text(json.dumps(permit_value), encoding="utf-8")
            permit_hash = hashlib.sha256(permit.read_bytes()).hexdigest()

            # Advance the lease clock only after the real child has been
            # bound.  This uses the runner's process-start hook as a
            # deterministic hand-off, so setup latency cannot consume the
            # permit before the event journal exists.  The permit file and
            # its hash remain unchanged; the monitor still exercises the
            # actual expiry/reap path while the fake Cargo child is live.
            state: dict[str, bool] = {
                "started": False,
                "permit_was_future": False,
                "clock_advanced": False,
            }
            safe_ref: dict[str, runner.SafeRunner] = {}
            clock_value = {"now": runner._dt.datetime.now(runner._dt.timezone.utc)}

            def clock() -> runner._dt.datetime:
                return clock_value["now"]

            class ExpireAfterStart:
                def process_started(self, **_kwargs: object) -> dict[str, bool]:
                    safe = safe_ref["safe"]
                    lease = safe.overlap_lease
                    assert lease is not None
                    state["started"] = True
                    state["permit_was_future"] = lease.expires > clock()
                    clock_value["now"] = permit_expiry + runner._dt.timedelta(seconds=1)
                    state["clock_advanced"] = True
                    return {"expired_after_start": True}

                def process_finished(
                    self, token: dict[str, bool], *, returncode: int | None
                ) -> dict[str, object]:
                    return {**token, "returncode": returncode}

            safe = runner.SafeRunner(
                manifest,
                manifest_path,
                "build",
                None,
                process_observer=ExpireAfterStart(),
                clock=clock,
            )
            safe_ref["safe"] = safe

            def isolated_offenders(*args: object, **kwargs: object) -> list[dict[str, object]]:
                # Keep the host-builder isolation from the historical test,
                # but retain the lease validation that the monitor performs
                # through the overlap argument.
                overlap = kwargs.get("overlap")
                if overlap is not None:
                    overlap.ensure_valid()
                return []

            safe.guard.offenders = isolated_offenders
            with self.assertRaisesRegex(runner.RunnerError, "expired"):
                safe.build()
            self.assertEqual(state, {
                "started": True,
                "permit_was_future": True,
                "clock_advanced": True,
            })
            self.assertEqual(hashlib.sha256(permit.read_bytes()).hexdigest(), permit_hash)
            receipt = json.loads((root / "output" / "build-receipt.json").read_text(encoding="utf-8"))
            self.assertEqual(receipt["status"], "failed")
            self.assertEqual(receipt["build_execution"]["class"], "compile_only")
            self.assertEqual(receipt["build_execution"]["lease_status"], "released")
            events = (root / "output" / "events-build.jsonl").read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(events), 1)
            event = json.loads(events[0])
            self.assertIn("expired", event["runner_error"])
            self.assertEqual(event["post_exit_descendants"], [])
            self.assertEqual(event["process_observation"]["expired_after_start"], True)

    def test_runner_reaps_setsid_descendant_after_rapid_leader_exit(self) -> None:
        """A detached helper cannot escape the generic runner's ownership."""

        with tempfile.TemporaryDirectory(prefix="native-build-overlap-setsid-") as directory:
            root = Path(directory)
            source = self._make_escaping_fake_cargo_source(root / "source")
            manifest, manifest_path, _permit, _lease_dir = self._overlap_fixture_with_sources(
                root,
                source_roots=(source.parent, source.parent),
            )
            manifest["build_execution"] = {"class": "exclusive"}
            manifest["environment"] = {"ESCAPED_PID_FILE": str(root / "escaped.pid")}
            for leg in ("before", "after"):
                manifest["legs"][leg]["build"][0]["argv"] = [str(source), "build"]
                manifest["legs"][leg]["allow_dirty"] = True
            manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            safe = runner.SafeRunner(manifest, manifest_path, "build", None)
            safe.guard.require_quiet = lambda: None
            safe.guard.offenders = lambda *_args, **_kwargs: []
            with self.assertRaisesRegex(runner.RunnerError, "escaped its owned process tree"):
                safe.build()

            receipt = json.loads((root / "output" / "build-receipt.json").read_text(encoding="utf-8"))
            self.assertEqual(receipt["status"], "failed")
            events = [
                json.loads(line)
                for line in (root / "output" / "events-build.jsonl").read_text(encoding="utf-8").splitlines()
            ]
            self.assertEqual(len(events), 1)
            event = events[0]
            self.assertTrue(event["escaped_descendants"])
            self.assertEqual(event["post_exit_descendants"], [])
            self.assertEqual(event["owned_pid_reuse"], [])
            escaped_pid = int((root / "escaped.pid").read_text(encoding="utf-8"))
            self.assertIsNone(runner._read_proc(escaped_pid))

    def test_owned_cleanup_does_not_signal_a_reused_pid(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-pid-reuse-") as directory:
            root = Path(directory)
            safe = runner.SafeRunner(
                {
                    "schema": runner.SCHEMA,
                    "lane": "w5-legacy",
                    "output_dir": str(root / "output"),
                    "legs": {"before": {"target_dir": str(root / "target")}},
                },
                ROOT / "tools/native_qualification_runner.py",
                "build",
                None,
            )
            owned = {
                991991: {
                    "pid": 991991,
                    "starttime": "old-start",
                    "ppid": 1,
                    "pgrp": 991991,
                    "comm": "cargo",
                    "exe": "/usr/bin/cargo",
                    "cwd": str(ROOT),
                    "cmdline": ["cargo", "build"],
                }
            }
            snapshot = {
                991991: {
                    **owned[991991],
                    "starttime": "new-start",
                }
            }
            with mock.patch.object(runner, "_proc_snapshot", return_value=snapshot), mock.patch.object(
                runner.os, "kill"
            ) as kill:
                safe._signal_owned_processes(owned, leader_pid=123, sig=runner.signal.SIGTERM)
            kill.assert_not_called()
            with mock.patch.object(runner.os, "waitpid") as waitpid:
                safe._reap_owned_processes(owned, leader_pid=123, snapshot=snapshot)
            waitpid.assert_not_called()

    def test_owned_record_rejects_reused_leader_tree(self) -> None:
        def entry(pid: int, ppid: int, starttime: str) -> dict[str, object]:
            return {
                "pid": pid,
                "ppid": ppid,
                "pgrp": pid,
                "starttime": starttime,
                "comm": "cargo",
                "exe": "/usr/bin/cargo",
                "cwd": str(ROOT),
                "cmdline": ["cargo", "build"],
            }

        leader = entry(991992, 1, "leader-old")
        snapshot = {
            1: entry(1, 0, "runner"),
            991992: entry(991992, 1, "leader-new"),
            991993: entry(991993, 991992, "foreign-child"),
        }
        owned = {991992: runner._process_record(leader)}
        pid_reuse: list[dict[str, object]] = []
        runner.SafeRunner._record_owned_processes(
            snapshot,
            process_pid=991992,
            runner_pid=1,
            preexisting={},
            owned=owned,
            pid_reuse=pid_reuse,
        )
        self.assertNotIn(991993, owned)
        self.assertTrue(any(item["pid"] == 991992 for item in pid_reuse))

    def test_owned_record_discards_reused_preexisting_subtree(self) -> None:
        def entry(pid: int, ppid: int, starttime: str) -> dict[str, object]:
            return {
                "pid": pid,
                "ppid": ppid,
                "pgrp": pid,
                "starttime": starttime,
                "comm": "cargo",
                "exe": "/usr/bin/cargo",
                "cwd": str(ROOT),
                "cmdline": ["cargo", "build"],
            }

        leader = entry(991994, 1, "leader")
        snapshot = {
            1: entry(1, 0, "runner"),
            991994: leader,
            991995: entry(991995, 991994, "preexisting-new"),
            991996: entry(991996, 991995, "preexisting-grandchild"),
        }
        owned = {991994: runner._process_record(leader)}
        pid_reuse: list[dict[str, object]] = []
        runner.SafeRunner._record_owned_processes(
            snapshot,
            process_pid=991994,
            runner_pid=1,
            preexisting={991995: "preexisting-old"},
            owned=owned,
            pid_reuse=pid_reuse,
        )
        self.assertNotIn(991995, owned)
        self.assertNotIn(991996, owned)
        self.assertTrue(any(item["pid"] == 991995 for item in pid_reuse))

    def test_final_kill_recaptures_a_child_forked_after_first_snapshot(self) -> None:
        def entry(pid: int, ppid: int, starttime: str) -> dict[str, object]:
            return {
                "pid": pid,
                "ppid": ppid,
                "pgrp": pid,
                "starttime": starttime,
                "comm": "cargo",
                "exe": "/usr/bin/cargo",
                "cwd": str(ROOT),
                "cmdline": ["cargo", "build"],
            }

        with tempfile.TemporaryDirectory(prefix="native-build-overlap-final-kill-") as directory:
            root = Path(directory)
            safe = runner.SafeRunner(
                {
                    "schema": runner.SCHEMA,
                    "lane": "w5-legacy",
                    "output_dir": str(root / "output"),
                    "legs": {"before": {"target_dir": str(root / "target")}},
                },
                ROOT / "tools/native_qualification_runner.py",
                "build",
                None,
            )
            leader = entry(991997, 1, "leader")
            child = entry(991998, 1, "child")
            forked = entry(991999, 991998, "forked-after-kill-snapshot")
            snapshots = [
                {1: entry(1, 0, "runner"), 991998: child},
                {1: entry(1, 0, "runner"), 991998: child},
                {1: entry(1, 0, "runner"), 991998: child, 991999: forked},
                {1: entry(1, 0, "runner")},
            ]
            signal_calls: list[set[int]] = []
            monotonic_values = iter((0.0, 0.0, 6.0, 6.0, 6.1, 6.2))

            def record_signal(owned, *, leader_pid, sig):
                signal_calls.append(set(owned))

            with mock.patch.object(runner, "_proc_snapshot", side_effect=snapshots), mock.patch.object(
                runner.time, "monotonic", side_effect=monotonic_values
            ), mock.patch.object(runner.time, "sleep"), mock.patch.object(
                safe, "_signal_owned_processes", side_effect=record_signal
            ), mock.patch.object(safe, "_reap_owned_processes"):
                lifecycle = safe._cleanup_owned_descendants(
                    leader_pid=991997,
                    runner_pid=1,
                    preexisting={},
                    owned={991997: runner._process_record(leader)},
                    pid_reuse=[],
                )

            self.assertEqual(lifecycle["survivors"], [])
            self.assertEqual(
                {record["pid"] for record in lifecycle["escaped"]},
                {991998, 991999},
            )
            self.assertIn(991999, signal_calls[-2])

    @staticmethod
    def _make_escaping_fake_cargo_source(root: Path) -> Path:
        """Create a fake Cargo that leaves a real setsid child behind."""

        root.mkdir(parents=True, exist_ok=False)
        (root / "Cargo.lock").write_text("# fake lock for runner tests\n", encoding="utf-8")
        build = root / "build"
        build.write_text(
            "#!/bin/bash\n"
            "set -eu\n"
            "mkdir -p \"$CARGO_TARGET_DIR\"\n"
            "printf '#!/bin/sh\\nexit 0\\n' > \"$CARGO_TARGET_DIR/probe\"\n"
            "chmod +x \"$CARGO_TARGET_DIR/probe\"\n"
            "setsid /bin/sh -c 'sleep 30' >/dev/null 2>&1 &\n"
            "printf '%s\\n' \"$!\" > \"$ESCAPED_PID_FILE\"\n"
            "exit 0\n",
            encoding="utf-8",
        )
        build.chmod(0o755)
        (root / "cargo").symlink_to("/bin/bash")
        subprocess.run(["git", "init", "--quiet", str(root)], check=True, stdout=subprocess.DEVNULL)
        subprocess.run(
            ["git", "-C", str(root), "add", "Cargo.lock", "build", "cargo"],
            check=True,
            stdout=subprocess.DEVNULL,
        )
        subprocess.run(
            [
                "git",
                "-C",
                str(root),
                "-c",
                "user.name=Native runner test",
                "-c",
                "user.email=native-runner-test@example.invalid",
                "commit",
                "--quiet",
                "-m",
                "fake cargo setsid source",
            ],
            check=True,
            stdout=subprocess.DEVNULL,
        )
        return root / "cargo"

    def test_compile_overlap_two_safe_runners_bind_real_cargo_children(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-build-overlap-integration-") as directory:
            root = Path(directory)
            source_parent = root / "sources"
            source_parent.mkdir()
            sources = [
                self._make_fake_cargo_source(source_parent / f"source-{index}", sleep_seconds=2.0)
                for index in range(4)
            ]
            permit = root / "permit.json"
            lease_dir = root / "leases"
            manifests: list[Path] = []
            manifest_values: list[dict[str, object]] = []
            for index in range(2):
                manifest_path = root / f"manifest-{index}.json"
                manifest = {
                    "schema": runner.SCHEMA,
                    "lane": "w5-legacy",
                    "output_dir": str(root / f"output-{index}"),
                    "build_execution": {"class": "compile_only", "permit_path": str(permit)},
                    "legs": {
                        "before": {
                            "source_root": str(sources[index * 2].parent),
                            "target_dir": str(root / f"target-{index}-before"),
                            "build": [{
                                "id": "build",
                                "argv": [str(sources[index * 2]), "build"],
                                "artifacts": ["{target_dir}/probe"],
                            }],
                        },
                        "after": {
                            "source_root": str(sources[index * 2 + 1].parent),
                            "target_dir": str(root / f"target-{index}-after"),
                            "build": [{
                                "id": "build",
                                "argv": [str(sources[index * 2 + 1]), "build"],
                                "artifacts": ["{target_dir}/probe"],
                            }],
                        },
                    },
                }
                manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
                manifests.append(manifest_path)
                manifest_values.append(manifest)
            expires = (runner._dt.datetime.now(runner._dt.timezone.utc) + runner._dt.timedelta(seconds=30)).isoformat()
            bindings = []
            hashes = []
            for manifest_path, manifest in zip(manifests, manifest_values):
                manifest_hash = hashlib.sha256(manifest_path.read_bytes()).hexdigest()
                hashes.append(manifest_hash)
                bindings.append({
                    "path": str(manifest_path.resolve()),
                    "sha256": manifest_hash,
                    "lane": "w5-legacy",
                    "source_roots": [
                        str(Path(manifest["legs"]["before"]["source_root"]).resolve()),
                        str(Path(manifest["legs"]["after"]["source_root"]).resolve()),
                    ],
                    "target_dirs": [
                        str(Path(manifest["legs"]["before"]["target_dir"]).resolve()),
                        str(Path(manifest["legs"]["after"]["target_dir"]).resolve()),
                    ],
                })
            permit.write_text(
                json.dumps({
                    "schema": runner.BUILD_PERMIT_SCHEMA,
                    "grant_id": "integration-grant",
                    "job_group": "integration-group",
                    "expires_utc": expires,
                    "max_jobs": 2,
                    "allowed_lanes": ["w5-legacy"],
                    "allowed_manifest_sha256": hashes,
                    "allowed_command_classes": ["cargo-build"],
                    "manifest_bindings": bindings,
                    "lease_dir": str(lease_dir),
                }),
                encoding="utf-8",
            )
            helper = root / "run-safe-build.py"
            helper.write_text(
                "import json, pathlib, sys\n"
                "import time\n"
                "tools = pathlib.Path(sys.argv[1]).resolve()\n"
                "sys.path.insert(0, str(tools))\n"
                "import native_qualification_runner as runner\n"
                "manifest_path = pathlib.Path(sys.argv[2]).resolve()\n"
                "runner.LITCHI_ROOTS = (pathlib.Path(sys.argv[3]).resolve(),)\n"
                "time.sleep(float(sys.argv[4]))\n"
                "manifest = json.loads(manifest_path.read_text(encoding='utf-8'))\n"
                "runner.SafeRunner(manifest, manifest_path, 'build', None).build()\n",
                encoding="utf-8",
            )
            processes = [
                subprocess.Popen(
                    [
                        sys.executable,
                        str(helper),
                        str(ROOT / "tools"),
                        str(manifest_path),
                        str(source_parent),
                        "0.0" if index == 0 else "0.35",
                    ],
                    cwd=str(ROOT),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                )
                for manifest_path in manifests
            ]
            results = []
            for process in processes:
                stdout, stderr = process.communicate(timeout=30)
                results.append((process.returncode, stdout, stderr))
            for returncode, stdout, stderr in results:
                self.assertEqual(returncode, 0, msg=f"helper failed: stdout={stdout}\nstderr={stderr}")
            event_ranges: list[tuple[runner._dt.datetime, runner._dt.datetime]] = []
            for index in range(2):
                receipt = json.loads((root / f"output-{index}" / "build-receipt.json").read_text(encoding="utf-8"))
                self.assertEqual(receipt["status"], "passed")
                self.assertEqual(receipt["build_execution"]["class"], "compile_only")
                owned_jobs = receipt["build_execution"]["owned_jobs"]
                self.assertEqual(len(owned_jobs), 2)
                self.assertTrue(all(job["command_class"] == "cargo-build" for job in owned_jobs))
                self.assertTrue(all(job["phase"] == "build" for job in owned_jobs))
                events = [
                    json.loads(line)
                    for line in (root / f"output-{index}" / "events-build.jsonl").read_text(encoding="utf-8").splitlines()
                ]
                self.assertEqual(len(events), 2)
                for event in events:
                    owned = event["overlap"]["owned_job"]
                    self.assertEqual(owned["command_class"], "cargo-build")
                    self.assertEqual(owned["phase_id"], event["event_id"])
                    event_ranges.append(
                        (
                            runner._parse_utc(event["start_utc"], "event start"),
                            runner._parse_utc(event["end_utc"], "event end"),
                        )
                    )
            self.assertTrue(
                any(
                    left_start < right_end and right_start < left_end
                    for index, (left_start, left_end) in enumerate(event_ranges)
                    for right_start, right_end in event_ranges[index + 1 :]
                ),
                "the two SafeRunner Cargo child event intervals never overlapped",
            )


if __name__ == "__main__":
    unittest.main()
