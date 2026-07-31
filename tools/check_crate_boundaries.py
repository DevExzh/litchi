#!/usr/bin/env python3
"""Reject workspace dependency edges that violate the accepted crate topology."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

# These crates have exact internal dependency ceilings. External dependencies
# are governed separately so the rules stay readable as the workspace grows.
INTERNAL_ALLOWLIST: dict[str, set[str]] = {
    "litchi-core": set(),
    "litchi-word": {"litchi-core"},
    "litchi-slide": {"litchi-core"},
    "litchi-sheet": {"litchi-core"},
    "litchi-ooxml-common": {"litchi-core", "litchi-opc"},
    "litchi-drawingml": {
        "litchi-core",
        "litchi-opc",
        "litchi-ooxml-common",
        "litchi-sheet",
    },
    "litchi-xlsx": {
        "litchi-core",
        "litchi-ooxml-common",
        "litchi-opc",
        "litchi-sheet",
    },
    "litchi-ole-common": {"litchi-cfb", "litchi-core"},
    "litchi-odraw": {"litchi-cfb", "litchi-core", "litchi-ole-common"},
}

OOXML_FORMATS = {"litchi-docx", "litchi-pptx", "litchi-xlsb", "litchi-xlsx"}
OLE_FORMATS = {"litchi-doc", "litchi-ppt", "litchi-xls"}

RUNTIME_PACKAGES = {"rayon", "reqwest", "tokio"}
RUNTIME_NEUTRAL_CRATES = {
    "litchi-core",
    "litchi-word",
    "litchi-slide",
    "litchi-sheet",
    "litchi-ooxml-common",
    "litchi-drawingml",
    "litchi-ole-common",
    "litchi-odraw",
}

# Phase-one code moved into litchi-core before this ADR existed. Keep this debt
# explicit so CI rejects additions while later extraction can remove entries.
CORE_FORBIDDEN = {
    "encoding_rs",
    "litchi-cfb",
    "litchi-opc",
    "quick-xml",
    "rayon",
    "reqwest",
    "soapberry-zip",
    "tokio",
}
CORE_DEPENDENCY_DEBT = {"encoding_rs", "quick-xml", "soapberry-zip"}
CORE_FEATURE_DEBT = {"odf", "ole", "rtf"}
CORE_FORMAT_FEATURES = {
    "doc",
    "docx",
    "odf",
    "ole",
    "ppt",
    "pptx",
    "rtf",
    "xls",
    "xlsb",
    "xlsx",
}


def metadata() -> dict[str, object]:
    result = subprocess.run(
        ["cargo", "metadata", "--format-version", "1", "--no-deps"],
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.stderr.write(result.stderr)
        raise SystemExit(result.returncode)
    return json.loads(result.stdout)


def main() -> int:
    data = metadata()
    workspace_ids = set(data["workspace_members"])
    packages = {
        package["name"]: package
        for package in data["packages"]
        if package["id"] in workspace_ids
    }
    workspace_names = set(packages)
    errors: list[str] = []

    direct: dict[str, set[str]] = {}
    for name, package in packages.items():
        dependencies = {dependency["name"] for dependency in package["dependencies"]}
        direct[name] = dependencies

        internal = dependencies & workspace_names
        if name in INTERNAL_ALLOWLIST:
            unexpected = internal - INTERNAL_ALLOWLIST[name]
            if name == "litchi-core":
                unexpected -= CORE_DEPENDENCY_DEBT
            if unexpected:
                errors.append(
                    f"{name} has forbidden internal dependencies: "
                    f"{', '.join(sorted(unexpected))}"
                )

        if name in RUNTIME_NEUTRAL_CRATES:
            runtimes = dependencies & RUNTIME_PACKAGES
            if runtimes:
                errors.append(
                    f"{name} is runtime-neutral but depends on: "
                    f"{', '.join(sorted(runtimes))}"
                )

    for family in (OOXML_FORMATS, OLE_FORMATS):
        for name in family & workspace_names:
            peers = direct[name] & (family - {name})
            if peers:
                errors.append(
                    f"{name} depends on concrete peer formats: "
                    f"{', '.join(sorted(peers))}"
                )

    for common in ("litchi-ooxml-common", "litchi-drawingml"):
        if common in direct:
            concrete = direct[common] & OOXML_FORMATS
            if concrete:
                errors.append(
                    f"{common} depends upward on concrete formats: "
                    f"{', '.join(sorted(concrete))}"
                )
    for common in ("litchi-ole-common", "litchi-odraw"):
        if common in direct:
            concrete = direct[common] & OLE_FORMATS
            if concrete:
                errors.append(
                    f"{common} depends upward on concrete formats: "
                    f"{', '.join(sorted(concrete))}"
                )

    core_dependencies = direct.get("litchi-core", set())
    current_dependency_debt = core_dependencies & CORE_FORBIDDEN
    added_debt = current_dependency_debt - CORE_DEPENDENCY_DEBT
    stale_dependency_debt = CORE_DEPENDENCY_DEBT - current_dependency_debt
    if added_debt:
        errors.append(
            "litchi-core added forbidden format/container dependencies: "
            + ", ".join(sorted(added_debt))
        )
    if stale_dependency_debt:
        errors.append(
            "remove resolved litchi-core dependency debt from the boundary checker: "
            + ", ".join(sorted(stale_dependency_debt))
        )

    core = packages.get("litchi-core")
    if core is not None:
        current_feature_debt = set(core["features"]) & CORE_FORMAT_FEATURES
        added_feature_debt = current_feature_debt - CORE_FEATURE_DEBT
        stale_feature_debt = CORE_FEATURE_DEBT - current_feature_debt
        if added_feature_debt:
            errors.append(
                "litchi-core added forbidden format features: "
                + ", ".join(sorted(added_feature_debt))
            )
        if stale_feature_debt:
            errors.append(
                "remove resolved litchi-core feature debt from the boundary checker: "
                + ", ".join(sorted(stale_feature_debt))
            )

    if errors:
        for error in errors:
            print(f"crate-boundary error: {error}", file=sys.stderr)
        return 1

    print(
        f"crate boundaries valid for {len(workspace_names)} workspace packages "
        f"({len(CORE_DEPENDENCY_DEBT) + len(CORE_FEATURE_DEBT)} explicit debt items)"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
