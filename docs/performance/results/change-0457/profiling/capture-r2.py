#!/usr/bin/env python3
"""R2 source-custody and sample-count amendment for the bounded 0457 profile runner.

The original capture.py and protocol.json remain immutable.  This wrapper
reuses the original recorder/postprocessor, changing only the custody rule:
candidate production sources must equal the bound build subset after excluding
the four newly archived fuzz inputs; the old control binary is authenticated
by its retained build/binary bindings while the current ambient source epoch
is recorded separately.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any


SCRIPT = Path(__file__).resolve()
PROFILE_ROOT = SCRIPT.parent
BUNDLE = PROFILE_ROOT.parent
ORIGINAL_PATH = PROFILE_ROOT / "capture.py"
PROTOCOL_PATH = PROFILE_ROOT / "protocol-r2.json"

spec = importlib.util.spec_from_file_location("change0457_profile_original", ORIGINAL_PATH)
if spec is None or spec.loader is None:
    raise SystemExit(f"cannot import original profile runner: {ORIGINAL_PATH}")
original = importlib.util.module_from_spec(spec)
spec.loader.exec_module(original)

# Make the original runner identify this amendment and keep its evidence in a
# distinct directory.  Its workload/perf argv and postprocessing remain the
# frozen implementation from capture.py.
original.__file__ = str(SCRIPT)
original.SCRIPT = SCRIPT
original.PROFILE_ROOT = PROFILE_ROOT / "r2"
original.BUNDLE = BUNDLE
original.REPO = BUNDLE.parents[3]
original.PROTOCOL_PATH = PROTOCOL_PATH
ORIGINAL_ROLE_BINDINGS = original.role_bindings
ORIGINAL_WRITE_JSON = original.write_json

PROVENANCE: dict[str, Any] = {}


def _load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _source_map(record: dict[str, Any]) -> dict[str, str]:
    path = BUNDLE / record["path"]
    return _load(path)


def _bound_record(role: dict[str, Any]) -> dict[str, Any]:
    item = role["source_manifest"]
    manifest_path = BUNDLE / item["path"]
    source_map = _load(manifest_path)
    return {
        "path": f"sources/{item['sha256']}.json",
        "sha256": item["sha256"],
        "files": len(source_map),
    }


def _r1_role_bindings(protocol: dict[str, Any], role_name: str):
    # The original binder still authenticates schema, binary, build receipt,
    # oracle, and bound source-manifest identity.  It returns its old bound
    # source record; R1 substitutes the current ambient record only after the
    # candidate subset check below.
    role, bindings, binary, bound_source = ORIGINAL_ROLE_BINDINGS(protocol, role_name)
    custody = original.import_custody()
    ambient = custody.sources()
    ambient_map = _source_map(ambient)
    bound = _bound_record(role)
    bound_map = _source_map(bound)

    if role_name == "candidate":
        extras = sorted(set(ambient_map) - set(bound_map))
        missing = sorted(set(bound_map) - set(ambient_map))
        changed = sorted(
            path for path in set(bound_map) & set(ambient_map)
            if bound_map[path] != ambient_map[path]
        )
        expected_extras = protocol["amendment"]["candidate_ambient_exclusions"]
        if extras != [item["path"] for item in expected_extras] or missing or changed:
            raise original.ProfileError(
                "candidate ambient production sources do not equal bound subset; "
                f"extras={extras!r} missing={missing!r} changed={changed!r}"
            )
        exclusions = [
            {"path": path, "sha256": ambient_map[path]}
            for path in extras
        ]
    else:
        exclusions = []

    PROVENANCE.clear()
    PROVENANCE.update({
        "amendment": "r2",
        "ambient_source_epoch": ambient,
        "bound_source": bound,
        "source_policy": (
            "candidate bound production subset equals current ambient source after the explicitly archived fuzz inputs are excluded"
            if role_name == "candidate" else
            "control old binary is authenticated by retained build and binary bindings; current ambient source is recorded separately"
        ),
        "candidate_ambient_exclusions": exclusions,
        "control_ambient_comparison": "withheld-by-design" if role_name == "control" else None,
    })
    # The original main compares source_before with this returned record.  It
    # must see the current ambient epoch, while PROVENANCE retains the old
    # build-bound subset for review.
    return role, bindings, binary, ambient


def _write_json(path: Path, value: Any) -> None:
    if isinstance(value, dict) and value.get("change") == 457:
        value["r1_provenance"] = PROVENANCE.copy()
        if value.get("role") in ("candidate", "control"):
            value["check_tag"] = f"profile-{value['role']}-large-r2"
    ORIGINAL_WRITE_JSON(path, value)


ORIGINAL_RUN_ORACLE = original.run_oracle


def _run_oracle(argv, cwd, env, output):
    # Mutate the recorded command itself so the receipt records the exact call.
    argv.extend(["--samples", "100", "--warmups", "3"])
    return ORIGINAL_RUN_ORACLE(argv, cwd, env, output)


original.run_oracle = _run_oracle
original.role_bindings = _r1_role_bindings
original.write_json = _write_json


if __name__ == "__main__":
    raise SystemExit(original.main())
