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

    def test_legacy_external_manifest_is_rendered_and_bound_to_validator(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-external-manifest-") as directory:
            root = Path(directory)
            manifest_path = root / "legacy-rtf-final.json"
            source_manifest = ROOT / "tools/native_qualification/manifests/legacy-rtf.json"
            manifest_path.write_text(source_manifest.read_text(encoding="utf-8"), encoding="utf-8")
            environment = {
                "LITCHI_COORD_BIN": str(root / "coord-bin"),
                "LITCHI_NATIVE_QUAL_ROOT": str(ROOT),
                "RTF_NATIVE_QUAL_OUTPUT": str(root / "output"),
                "RTF_BEFORE_SOURCE": str(root / "before"),
                "RTF_AFTER_SOURCE": str(root / "after"),
                "RTF_TARGET_ROOT": str(root / "targets"),
                "RTF_NATIVE_TMP": str(root / "tmp"),
                "RTF_NATIVE_PREPARATION_RECEIPT": str(root / "preparation.json"),
                "RTF_BASELINE_REVISION": "0" * 40,
                "RTF_CANDIDATE_REVISION": "1" * 40,
            }
            old = {key: os.environ.get(key) for key in environment}
            try:
                os.environ.update(environment)
                loaded = runner.load_manifest(manifest_path)
                validator_argv = loaded["validators"][0]["argv"]
                self.assertEqual(validator_argv[-1], "{manifest_path}")
                context = {
                    "manifest_path": manifest_path.resolve(),
                    "output_dir": root / "output",
                }
                rendered = runner._render_argv(validator_argv, context)
                self.assertEqual(rendered[-1], str(manifest_path.resolve()))

                invalid = json.loads(manifest_path.read_text(encoding="utf-8"))
                invalid["validators"][0]["argv"][-1] = str(root / "wrong.json")
                manifest_path.write_text(json.dumps(invalid), encoding="utf-8")
                with self.assertRaisesRegex(runner.RunnerError, "current manifest path"):
                    runner.load_manifest(manifest_path)
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


if __name__ == "__main__":
    unittest.main()
