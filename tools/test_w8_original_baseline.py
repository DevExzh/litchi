"""Focused checks for the pristine-f2 W8 consumer adapter."""

from __future__ import annotations

import collections
from pathlib import Path
import tarfile
import tempfile
import unittest

from tools import w8_original_baseline as overlay


ROOT = Path(__file__).resolve().parents[1]
PROBE = ROOT / overlay.PROBE_RELATIVE


class AdapterSource(unittest.TestCase):
    def test_only_four_old_accessor_calls_are_translated(self):
        source = PROBE.read_text(encoding="utf-8")
        adapted, records = overlay.consumer_adapter(source)
        self.assertEqual(len(records), 4)
        self.assertEqual(source.count(".into_file()"), 4)
        self.assertNotIn(".into_file()", adapted.decode("utf-8"))
        self.assertEqual(
            [record["function"] for record in records],
            ["unified_batch", "unified_import", "unified_merge", "unified_refresh"],
        )
        self.assertEqual(
            [record["new_accessor"] for record in records],
            [".file"] * 4,
        )
        # The adapter is a consumer example, never a production source patch;
        # its only accepted API spelling is the old public result field.
        self.assertEqual(adapted.decode("utf-8").count(".file"), source.count(".file") + 4)

    def test_history_is_pristine_f2_and_probe_is_independently_pinned(self):
        tree = overlay.git_tree(ROOT)
        history = overlay.verify_history(ROOT, tree)
        self.assertEqual(history["baseline_commit"], overlay.BASELINE_COMMIT)
        self.assertEqual(history["probe_source_commit"], overlay.PROBE_SOURCE_COMMIT)
        self.assertNotIn(overlay.PROBE_RELATIVE.as_posix(), tree)
        self.assertNotIn(overlay.ADAPTER_RELATIVE.as_posix(), tree)
        self.assertEqual(history["corrected_probe"]["sha256"], overlay.REVIEWED_PROBE_SHA256)


class RouteProof(unittest.TestCase):
    def test_current_probe_proves_all_native_facade_routes(self):
        value = overlay.route_map(PROBE)
        self.assertEqual(value["route_count"], 109)
        self.assertEqual(
            collections.Counter(row["family"] for row in value["routes"]),
            collections.Counter({
                "docx": 11, "odt": 9, "doc": 9, "rtf": 9,
                "xlsx": 11, "xlsb": 11, "xls": 11, "ods": 11,
                "pptx": 9, "ppt": 9, "odp": 9,
            }),
        )
        self.assertEqual(value["routes"][0]["route"], "docx/open_get")
        self.assertEqual(value["routes"][-1]["route"], "odp/merge_publication")
        self.assertTrue(value["provenance"]["production_source_unchanged"])
        self.assertTrue(value["provenance"]["no_workload_substitution"])
        self.assertTrue(all(row["source_assertions"] for row in value["routes"]))
        self.assertTrue(all(row["native_timing"] and row["facade_timing"] for row in value["routes"]))

    def test_bindings_are_route_specific_and_preserve_pending_status(self):
        value = overlay.route_map(PROBE)
        set_commit = next(row for row in value["routes"] if row["route"] == "docx/set_commit")
        import_commit = next(row for row in value["routes"] if row["route"] == "xlsx/import_commit")
        refresh = next(row for row in value["routes"] if row["route"] == "docx/refresh_commit")
        self.assertIn("source.native_set", set_commit["native_dispatch"])
        self.assertIn("source.native_import", import_commit["native_dispatch"])
        self.assertIn("source.native_refresh", refresh["native_dispatch"])
        self.assertEqual(set_commit["adapter_api"]["old_accessor"], ".file")
        self.assertEqual(import_commit["adapter_api"]["consumer_transform"], ".into_file() -> .file")
        self.assertFalse(refresh["changed"])
        self.assertEqual(value["provenance"]["native_validation"], "pending")
        self.assertEqual(value["provenance"]["performance_qualification"], "pending")


class SourceClosure(unittest.TestCase):
    def test_archive_rejects_an_extra_production_file(self):
        files = {"Cargo.toml": b"[workspace]\n", "crates/litchi-ops/Cargo.toml": b"[package]\n"}
        tree = {name: ("100644", overlay.git_blob_sha1(payload)) for name, payload in files.items()}
        with tempfile.TemporaryDirectory() as temporary:
            archive = Path(temporary) / "source.tar"
            with tarfile.open(archive, "w") as stream:
                for name, payload in files.items():
                    path = Path(temporary) / name
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(payload)
                    stream.add(path, arcname=name)
                extra = Path(temporary) / "unexpected.rs"
                extra.write_bytes(b"not in f2")
                stream.add(extra, arcname=extra.name)
            with self.assertRaises(overlay.OverlayError):
                overlay.verify_source_archive(archive, tree)


if __name__ == "__main__":
    unittest.main()
