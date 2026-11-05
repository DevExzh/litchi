"""Focused checks for the pristine-f2 W8 consumer adapter."""

from __future__ import annotations

import collections
import hashlib
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
        accessor_records = [record for record in records if record["old_accessor"] == ".into_file()"]
        resolution_records = [record for record in records if record["old_accessor"] == overlay.XLSB_RESOLUTION_LABEL]
        self.assertEqual(len(records), 7)
        self.assertEqual(len(accessor_records), 4)
        self.assertEqual(len(resolution_records), 3)
        self.assertEqual(source.count(".into_file()"), 4)
        self.assertEqual(source.count(".resolve()"), 3)
        self.assertEqual(source.count(overlay.XLSB_RESOLUTION_SEQUENCE), 3)
        self.assertNotIn(".into_file()", adapted.decode("utf-8"))
        self.assertNotIn(".resolve()", adapted.decode("utf-8"))
        self.assertEqual(
            [record["function"] for record in accessor_records],
            ["unified_batch", "unified_import", "unified_merge", "unified_refresh"],
        )
        self.assertEqual(
            [record["new_accessor"] for record in accessor_records],
            [".file"] * 4,
        )
        self.assertEqual(
            [record["function"] for record in resolution_records],
            ["native_set", "native_import", "native_merge"],
        )
        # The adapter is a consumer example, never a production source patch;
        # its only accepted API spelling is the old public result field.
        self.assertEqual(adapted.decode("utf-8").count(".file"), source.count(".file") + 4)

    def test_workspace_rejects_a_fake_short_historical_lock(self):
        source = PROBE.read_text(encoding="utf-8")
        adapted, _records = overlay.consumer_adapter(source)
        with tempfile.TemporaryDirectory() as tmp:
            workspace = Path(tmp)
            probe = workspace / overlay.PROBE_RELATIVE
            adapter = workspace / overlay.ADAPTER_RELATIVE
            lock = workspace / overlay.LOCK_RELATIVE
            probe.parent.mkdir(parents=True, exist_ok=True)
            probe.write_bytes(source.encode("utf-8"))
            adapter.write_bytes(adapted)
            lock.write_bytes(b"fake lock accepted by the old verifier\n")
            with self.assertRaises(overlay.OverlayError):
                overlay.verify_workspace(workspace, {}, probe, adapter, lock, adapted)

    def test_baseline_adapter_matches_pinned_d39_adapter_identity(self):
        source = PROBE.read_text(encoding="utf-8")
        adapted, _records = overlay.consumer_adapter(source)
        self.assertEqual(len(adapted), 74_719)
        self.assertEqual(
            hashlib.sha256(adapted).hexdigest(),
            "3e14376c366fa911025212d7e7081ab5c3b5aaac804c1e01a1415fd5cd826ca0",
        )

    def test_accidental_extra_xlsb_resolution_is_rejected(self):
        source = PROBE.read_text(encoding="utf-8")
        altered = source.replace(
            "let mut edit = package\n                        .edit_cell_values(0_usize)",
            "let mut edit = package\n                        .resolve()\n                        .edit_cell_values(0_usize)",
            1,
        )
        self.assertEqual(altered.count(".resolve()"), 4)
        with self.assertRaises(overlay.OverlayError):
            overlay.consumer_adapter(altered)

    def test_compatibility_patch_contains_only_probe_and_adapter_additions(self):
        source = PROBE.read_bytes()
        adapted, _records = overlay.consumer_adapter(source.decode("utf-8"))
        patch = overlay.adapter_patch(source, adapted).decode("utf-8")
        self.assertEqual(patch.count("--- /dev/null\n"), 2)
        self.assertIn(f"+++ b/{overlay.PROBE_RELATIVE.as_posix()}\n", patch)
        self.assertIn(f"+++ b/{overlay.ADAPTER_RELATIVE.as_posix()}\n", patch)
        self.assertNotIn("crates/litchi-ops/src/", patch)

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
    @staticmethod
    def _archive(path: Path, files: dict[str, bytes], extra: list[tarfile.TarInfo] | None = None) -> None:
        with tarfile.open(path, "w") as stream:
            for name, payload in files.items():
                file_path = path.parent / name
                file_path.parent.mkdir(parents=True, exist_ok=True)
                file_path.write_bytes(payload)
                stream.add(file_path, arcname=name)
            for member in extra or []:
                stream.addfile(member)

    def test_archive_rejects_an_extra_production_file(self):
        files = {"Cargo.toml": b"[workspace]\n", "crates/litchi-ops/Cargo.toml": b"[package]\n"}
        tree = {name: ("100644", overlay.git_blob_sha1(payload)) for name, payload in files.items()}
        with tempfile.TemporaryDirectory() as temporary:
            archive = Path(temporary) / "source.tar"
            self._archive(archive, files)
            with tarfile.open(archive, "a") as stream:
                extra = Path(temporary) / "unexpected.rs"
                extra.write_bytes(b"not in f2")
                stream.add(extra, arcname=extra.name)
            with self.assertRaises(overlay.OverlayError):
                overlay.verify_source_archive(archive, tree)

    def test_archive_rejects_missing_and_mutated_production_files(self):
        files = {"Cargo.toml": b"[workspace]\n", "crates/litchi-ops/Cargo.toml": b"[package]\n"}
        tree = {name: ("100644", overlay.git_blob_sha1(payload)) for name, payload in files.items()}
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            missing = root / "missing.tar"
            self._archive(missing, {"Cargo.toml": files["Cargo.toml"]})
            with self.assertRaises(overlay.OverlayError):
                overlay.verify_source_archive(missing, tree)
            mutated = root / "mutated.tar"
            self._archive(mutated, {**files, "Cargo.toml": b"changed\n"})
            with self.assertRaises(overlay.OverlayError):
                overlay.verify_source_archive(mutated, tree)

    def test_archive_rejects_links_and_devices(self):
        files = {"Cargo.toml": b"[workspace]\n", "crates/litchi-ops/Cargo.toml": b"[package]\n"}
        tree = {name: ("100644", overlay.git_blob_sha1(payload)) for name, payload in files.items()}
        members = []
        for member_type, name in ((tarfile.SYMTYPE, "symlink"), (tarfile.LNKTYPE, "hardlink"), (tarfile.CHRTYPE, "device")):
            member = tarfile.TarInfo(name)
            member.type = member_type
            member.linkname = "Cargo.toml"
            member.devmajor = 1
            member.devminor = 3
            members.append(member)
        with tempfile.TemporaryDirectory() as temporary:
            for member in members:
                archive = Path(temporary) / f"{member.name}.tar"
                self._archive(archive, files, [member])
                with self.subTest(member=member.name), self.assertRaises(overlay.OverlayError):
                    overlay.verify_source_archive(archive, tree)


if __name__ == "__main__":
    unittest.main()
