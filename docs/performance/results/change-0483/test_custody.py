#!/usr/bin/env python3
"""Tamper tests for the 0483 retained fuzz and freeze custody checks."""

from __future__ import annotations

from contextlib import contextmanager
import copy
import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import unittest

import verify


BUNDLE = Path(__file__).resolve().parent
SOURCE_MANIFEST = "30d3ad60b0ae90911992b9855c5328913de2741f27090ba6c31b760c2ebcaeaf.json"


def metadata(path: Path) -> dict[str, int | str]:
    data = path.read_bytes()
    return {"bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def copy_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


@contextmanager
def copied_fuzz_bundle():
    """Provide only the accepted fuzz inputs needed by verify.check_fuzz_custody."""

    with tempfile.TemporaryDirectory(prefix="change-0483-custody-") as directory:
        root = Path(directory)
        for relative in (
            "fuzz.py",
            "fuzz-seeds.py",
            "fuzz/seed-manifest-v2.json",
            "fuzz/generator-v2.json",
            f"validation-sources/{SOURCE_MANIFEST}",
        ):
            copy_file(BUNDLE / relative, root / relative)
        shutil.copytree(BUNDLE / "fuzz" / "seeds", root / "fuzz" / "seeds")
        shutil.copytree(BUNDLE / "fuzz" / "accepted" / "build-inputs", root / "fuzz" / "accepted" / "build-inputs")
        shutil.copytree(BUNDLE / "fuzz" / "accepted" / "post-run", root / "fuzz" / "accepted" / "post-run")
        for kind in ("prepared", "build", "smoke"):
            copy_file(
                BUNDLE / "fuzz" / "accepted" / f"{kind}.json",
                root / "fuzz" / "accepted" / f"{kind}.json",
            )
        original_root = verify.ROOT
        verify.ROOT = root
        try:
            yield root
        finally:
            verify.ROOT = original_root


def fuzz_plan(root: Path) -> dict[str, object]:
    def reference(relative: str) -> dict[str, int | str]:
        return {"path": relative, **metadata(root / relative)}

    receipts = [
        {
            "label": f"custody-{kind}-accepted",
            "kind": kind,
            **reference(f"fuzz/accepted/{kind}.json"),
        }
        for kind in ("prepared", "build", "smoke")
    ]
    return {
        "attempt": "accepted",
        "required_labels": [item["label"] for item in receipts],
        "seed_root": "fuzz/seeds",
        "seed_manifest": reference("fuzz/seed-manifest-v2.json"),
        "generator": reference("fuzz/generator-v2.json"),
        "receipts": receipts,
    }


class CustodyTests(unittest.TestCase):
    def test_changed_build_input_cannot_rebind_prepared_record(self) -> None:
        with copied_fuzz_bundle() as root:
            build_path = root / "fuzz/accepted/build.json"
            build = json.loads(build_path.read_text(encoding="utf-8"))
            cargo_toml = root / "fuzz/accepted/build-inputs/Cargo.toml"
            cargo_toml.write_bytes(cargo_toml.read_bytes() + b"# tampered\n")
            build["inputs"] = copy.deepcopy(build["inputs"])
            build["inputs"]["manifest"] = metadata(cargo_toml)
            write_json(build_path, build)
            with self.assertRaisesRegex(verify.VerificationError, "retained prepared record"):
                verify.check_fuzz_custody({"fuzz": fuzz_plan(root)})

    def test_smoke_options_and_corpus_are_bound_to_prepared_run(self) -> None:
        mutations = (
            ("options", lambda smoke: smoke["argv"].__setitem__(3, "-seed=484"), "smoke argv differs from fuzz.py"),
            (
                "corpus",
                lambda smoke: smoke["corpus_before"].__setitem__("plain-stored.docx", {"bytes": 1, "sha256": "0" * 64}),
                "smoke corpus-before differs from prepared corpus",
            ),
        )
        for name, mutate, expected in mutations:
            with self.subTest(name=name), copied_fuzz_bundle() as root:
                smoke_path = root / "fuzz/accepted/smoke.json"
                smoke = json.loads(smoke_path.read_text(encoding="utf-8"))
                mutate(smoke)
                write_json(smoke_path, smoke)
                with self.assertRaisesRegex(verify.VerificationError, expected):
                    verify.check_fuzz_custody({"fuzz": fuzz_plan(root)})

    def test_pre_freeze_rejects_late_retained_custody(self) -> None:
        with copied_fuzz_bundle() as root:
            (root / "builds").mkdir()
            write_json(root / "builds/normal.json", {"copied_utc": "2026-09-09T03:00:02+00:00"})
            write_json(root / "builds/allocator.json", {"copied_utc": "2026-09-09T03:00:02+00:00"})
            protocol = {"frozen_utc": "2026-09-09T03:00:05+00:00"}
            binaries = {
                "binaries": {
                    "normal": {"build_path": "builds/normal.json"},
                    "allocator": {"build_path": "builds/allocator.json"},
                }
            }
            validation = {
                "times": {
                    "pilot": (
                        verify.timestamp("2026-09-09T03:00:03+00:00", "pilot.started"),
                        verify.timestamp("2026-09-09T03:00:04+00:00", "pilot.finished"),
                    )
                }
            }
            fuzz = {
                "times": {
                    "smoke": (
                        verify.timestamp("2026-09-09T03:00:04+00:00", "smoke.started"),
                        verify.timestamp("2026-09-09T03:00:06+00:00", "smoke.finished"),
                    )
                }
            }
            with self.assertRaisesRegex(verify.VerificationError, "finished after protocol freeze"):
                verify.check_pre_freeze_chronology(protocol, binaries, validation, fuzz)


if __name__ == "__main__":
    unittest.main()
