#!/usr/bin/env python3
"""Reject workspace dependency edges that violate the accepted crate topology."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_POLICY = Path(__file__).with_name("crate_boundaries.json")

OOXML_FORMATS = frozenset(
    {"litchi-docx", "litchi-pptx", "litchi-xlsb", "litchi-xlsx"}
)
OLE_FORMATS = frozenset({"litchi-doc", "litchi-ppt", "litchi-xls"})
ODF_FORMATS = frozenset(
    {
        "litchi-odb",
        "litchi-odc",
        "litchi-odf",
        "litchi-odf-formula",
        "litchi-odg",
        "litchi-odi",
        "litchi-odm",
        "litchi-odp",
        "litchi-ods",
        "litchi-odt",
        "litchi-oth",
    }
)
COMMON_FAMILY_GUARDS = {
    "litchi-cfb": OLE_FORMATS,
    "litchi-drawingml": OOXML_FORMATS,
    "litchi-ooxml-common": OOXML_FORMATS,
    "litchi-odf-common": ODF_FORMATS,
    "litchi-odraw": OLE_FORMATS,
    "litchi-ole-common": OLE_FORMATS,
    "litchi-opc": OOXML_FORMATS,
}
RETIRED_MONOLITHS = frozenset({"litchi-ooxml", "litchi-ole"})
RETIRED_FACADE_FEATURES = frozenset({"ole"})
XLSB_SOURCE_ROOT = Path("crates/litchi-xlsb/src")
XLSX_SOURCE_ROOT = Path("crates/litchi-xlsx/src")
PUBLIC_XLSX_MODULE = re.compile(r"^\s*pub(?:\([^)]*\))?\s+mod\s+xlsx\s*;")
PACKAGE_XLSX_PATH = re.compile(r"(?<![A-Za-z0-9_])package::xlsx\b")
RETIRED_XLSX_CHART_FILES = (
    Path("chart/anchor.rs"),
    Path("chart/codec.rs"),
    Path("chart/model.rs"),
    Path("chart/relationship.rs"),
)
SPREADSHEET_CHART_FACADES = {
    "litchi-xlsb": (XLSB_SOURCE_ROOT / "chart.rs", XLSB_SOURCE_ROOT / "chart/mod.rs"),
    "litchi-xlsx": (XLSX_SOURCE_ROOT / "chart.rs", XLSX_SOURCE_ROOT / "chart/mod.rs"),
}
SHARED_SPREADSHEET_CHART_TYPES = (
    "Anchor",
    "Chart",
    "ExternalDataPart",
    "ExternalDataTarget",
    "Relationship",
    "RelationshipTarget",
    "Target",
    "UserShapesPart",
)
LOCAL_SHARED_CHART_TYPE = re.compile(
    r"^\s*(?:pub(?:\([^)]*\))?\s+)?(?:struct|enum|type)\s+(?:"
    + "|".join(SHARED_SPREADSHEET_CHART_TYPES)
    + r")\b"
)
DRAWINGML_CHART_CODEC = re.compile(
    r"\blitchi_drawingml::chart::(?:"
    r"reader|writer|read_chart|write_chart|ChartReader|ChartWriter)\b"
)
MAX_SPREADSHEET_CHART_FACADE_LINES = 200
RETIRED_XLSX_SHAPE_FILES = (
    Path("shapes/codec.rs"),
    Path("shapes/model.rs"),
    Path("shapes/tests.rs"),
)
SPREADSHEET_SHAPE_FACADES = {
    "litchi-xlsb": (
        XLSB_SOURCE_ROOT / "shapes.rs",
        XLSB_SOURCE_ROOT / "writer/shape.rs",
    ),
    "litchi-xlsx": (
        XLSX_SOURCE_ROOT / "shapes/mod.rs",
        XLSX_SOURCE_ROOT / "writer/shape.rs",
    ),
}
MAX_SPREADSHEET_SHAPE_FACADE_LINES = 200
LEGACY_HOST_SHAPE_NAMES = (
    "DrawingObject",
    "DrawingObjectSpec",
    "DrawingOleObject",
    "OleObjectAspect",
    "ShapeAnchor",
    "ShapeEmitter",
)
LOCAL_HOST_SHAPE_TYPE = re.compile(
    r"^\s*(?:pub(?:\([^)]*\))?\s+)?(?:struct|enum|type|union)\s+\w+\b"
)
LEGACY_HOST_SHAPE_NAME = re.compile(
    r"\b(?:" + "|".join(LEGACY_HOST_SHAPE_NAMES) + r")\b"
)
QUICK_XML_USE = re.compile(r"\bquick_xml\b")
XDR_XML_EMISSION = re.compile(r"(?<![A-Za-z0-9_])xdr:")


@dataclass(frozen=True, order=True)
class Edge:
    """A dependency-direction edge, from dependent to dependency."""

    dependent: str
    dependency: str

    def display(self) -> str:
        return f"{self.dependent} -> {self.dependency}"


@dataclass(frozen=True)
class Debt:
    order: int
    edge: Edge
    reason: str
    exit: str


@dataclass(frozen=True)
class NamedDebt:
    order: int
    name: str
    reason: str
    exit: str


@dataclass(frozen=True)
class Policy:
    packages: frozenset[str]
    canonical_edges: frozenset[Edge]
    migration_hosts: frozenset[str]
    migration_debt: tuple[Debt, ...]
    runtime_neutral: frozenset[str]
    runtime_packages: frozenset[str]
    core_forbidden_dependencies: frozenset[str]
    core_dependency_debt: tuple[NamedDebt, ...]
    core_format_features: frozenset[str]
    core_feature_debt: tuple[NamedDebt, ...]

    @property
    def migration_edges(self) -> frozenset[Edge]:
        return frozenset(item.edge for item in self.migration_debt)


@dataclass(frozen=True)
class Snapshot:
    packages: frozenset[str]
    manifests: frozenset[Path]
    edges: dict[Edge, tuple[str, ...]]
    dependencies: dict[str, frozenset[str]]
    normal_dependencies: dict[str, frozenset[str]]
    features: dict[str, frozenset[str]]


class PolicyError(ValueError):
    pass


def _require_string(record: dict[str, Any], key: str, context: str) -> str:
    value = record.get(key)
    if not isinstance(value, str) or not value:
        raise PolicyError(f"{context}.{key} must be a non-empty string")
    return value


def _require_order(record: dict[str, Any], context: str) -> int:
    value = record.get("order")
    if not isinstance(value, int) or isinstance(value, bool) or value < 0:
        raise PolicyError(f"{context}.order must be a non-negative integer")
    return value


def _require_sorted_strings(value: Any, context: str) -> tuple[str, ...]:
    invalid_item = isinstance(value, list) and any(
        not isinstance(item, str) or not item for item in value
    )
    if not isinstance(value, list) or invalid_item:
        raise PolicyError(f"{context} must be a list of non-empty strings")
    if len(value) != len(set(value)):
        raise PolicyError(f"{context} contains duplicates")
    if value != sorted(value):
        raise PolicyError(f"{context} must be sorted")
    return tuple(value)


def _parse_named_debt(value: Any, context: str) -> tuple[NamedDebt, ...]:
    if not isinstance(value, list):
        raise PolicyError(f"{context} must be a list")
    result: list[NamedDebt] = []
    for index, item in enumerate(value):
        item_context = f"{context}[{index}]"
        if not isinstance(item, dict):
            raise PolicyError(f"{item_context} must be an object")
        result.append(
            NamedDebt(
                order=_require_order(item, item_context),
                name=_require_string(item, "name", item_context),
                reason=_require_string(item, "reason", item_context),
                exit=_require_string(item, "exit", item_context),
            )
        )
    keys = [(item.order, item.name) for item in result]
    if keys != sorted(keys):
        raise PolicyError(f"{context} must be ordered by order, then name")
    if len({item.name for item in result}) != len(result):
        raise PolicyError(f"{context} contains duplicate names")
    return tuple(result)


def parse_policy(raw: Any) -> Policy:
    """Parse and self-check the checked-in topology policy."""

    if not isinstance(raw, dict):
        raise PolicyError("policy root must be an object")
    if raw.get("schema") != 1:
        raise PolicyError("policy schema must be 1")

    package_map = raw.get("packages")
    if not isinstance(package_map, dict):
        raise PolicyError("packages must be an object")
    retired = RETIRED_MONOLITHS & package_map.keys()
    if retired:
        raise PolicyError(
            "retired monoliths cannot return as workspace packages: "
            + ", ".join(sorted(retired))
        )
    package_names = list(package_map)
    if package_names != sorted(package_names):
        raise PolicyError("packages must be sorted by package name")

    canonical: set[Edge] = set()
    for dependent, dependencies in package_map.items():
        if not isinstance(dependent, str) or not dependent:
            raise PolicyError("package names must be non-empty strings")
        for dependency in _require_sorted_strings(dependencies, f"packages.{dependent}"):
            canonical.add(Edge(dependent, dependency))

    migration_hosts = frozenset(
        _require_sorted_strings(raw.get("migration_hosts"), "migration_hosts")
    )
    migration_raw = raw.get("migration_debt")
    if not isinstance(migration_raw, list):
        raise PolicyError("migration_debt must be a list")
    migration: list[Debt] = []
    for index, item in enumerate(migration_raw):
        context = f"migration_debt[{index}]"
        if not isinstance(item, dict):
            raise PolicyError(f"{context} must be an object")
        migration.append(
            Debt(
                order=_require_order(item, context),
                edge=Edge(
                    _require_string(item, "dependent", context),
                    _require_string(item, "dependency", context),
                ),
                reason=_require_string(item, "reason", context),
                exit=_require_string(item, "exit", context),
            )
        )
    migration_keys = [(item.order, item.edge) for item in migration]
    if migration_keys != sorted(migration_keys):
        raise PolicyError("migration_debt must be ordered by order, then edge")
    if len({item.order for item in migration}) != len(migration):
        raise PolicyError("migration_debt orders must be unique")
    migration_edges = {item.edge for item in migration}
    if len(migration_edges) != len(migration):
        raise PolicyError("migration_debt contains duplicate edges")
    overlap = canonical & migration_edges
    if overlap:
        joined = ", ".join(edge.display() for edge in sorted(overlap))
        raise PolicyError(f"canonical and migration edges overlap: {joined}")
    self_edges = sorted(
        edge for edge in canonical | migration_edges if edge.dependent == edge.dependency
    )
    if self_edges:
        raise PolicyError(
            "policy contains self dependencies: "
            + ", ".join(edge.display() for edge in self_edges)
        )

    packages = frozenset(package_names)
    referenced = {
        name
        for edge in canonical | migration_edges
        for name in (edge.dependent, edge.dependency)
    }
    unknown = referenced - packages
    if unknown:
        raise PolicyError("edges reference unknown packages: " + ", ".join(sorted(unknown)))
    if not migration_hosts <= packages:
        raise PolicyError(
            "migration_hosts references unknown packages: "
            + ", ".join(sorted(migration_hosts - packages))
        )
    host_canonical = sorted(edge for edge in canonical if edge.dependent in migration_hosts)
    if host_canonical:
        raise PolicyError(
            "migration-host edges must be debt, not canonical: "
            + ", ".join(edge.display() for edge in host_canonical)
        )

    runtime_neutral = frozenset(
        _require_sorted_strings(raw.get("runtime_neutral"), "runtime_neutral")
    )
    if not runtime_neutral <= packages:
        raise PolicyError(
            "runtime_neutral references unknown packages: "
            + ", ".join(sorted(runtime_neutral - packages))
        )
    runtime_packages = frozenset(
        _require_sorted_strings(raw.get("runtime_packages"), "runtime_packages")
    )

    core = raw.get("core")
    if not isinstance(core, dict):
        raise PolicyError("core must be an object")
    core_forbidden = frozenset(
        _require_sorted_strings(
            core.get("forbidden_dependencies"), "core.forbidden_dependencies"
        )
    )
    core_dependency_debt = _parse_named_debt(
        core.get("dependency_debt"), "core.dependency_debt"
    )
    core_features = frozenset(
        _require_sorted_strings(core.get("format_features"), "core.format_features")
    )
    core_feature_debt = _parse_named_debt(core.get("feature_debt"), "core.feature_debt")

    debt_orders = (
        [item.order for item in migration]
        + [item.order for item in core_dependency_debt]
        + [item.order for item in core_feature_debt]
    )
    if len(debt_orders) != len(set(debt_orders)):
        raise PolicyError("debt orders must be unique across the complete policy")

    dependency_debt_names = {item.name for item in core_dependency_debt}
    if not dependency_debt_names <= core_forbidden:
        raise PolicyError("core dependency debt must also be forbidden")
    internal_named_debt = dependency_debt_names & packages
    if internal_named_debt:
        raise PolicyError(
            "internal core debt must use migration_debt edges: "
            + ", ".join(sorted(internal_named_debt))
        )
    canonical_core_forbidden = sorted(
        edge
        for edge in canonical
        if edge.dependent == "litchi-core" and edge.dependency in core_forbidden
    )
    if canonical_core_forbidden:
        raise PolicyError(
            "forbidden core edges must be migration debt, not canonical: "
            + ", ".join(edge.display() for edge in canonical_core_forbidden)
        )
    feature_debt_names = {item.name for item in core_feature_debt}
    if not feature_debt_names <= core_features:
        raise PolicyError("core feature debt must also be a format feature")

    return Policy(
        packages=packages,
        canonical_edges=frozenset(canonical),
        migration_hosts=migration_hosts,
        migration_debt=tuple(migration),
        runtime_neutral=runtime_neutral,
        runtime_packages=runtime_packages,
        core_forbidden_dependencies=core_forbidden,
        core_dependency_debt=core_dependency_debt,
        core_format_features=core_features,
        core_feature_debt=core_feature_debt,
    )


def load_policy(path: Path) -> Policy:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise PolicyError(f"cannot read {path}: {error}") from error
    return parse_policy(raw)


def cargo_metadata() -> dict[str, Any]:
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


def snapshot_from_metadata(data: dict[str, Any]) -> Snapshot:
    workspace_ids = set(data["workspace_members"])
    packages = [package for package in data["packages"] if package["id"] in workspace_ids]
    names = frozenset(package["name"] for package in packages)

    evidence: dict[Edge, set[str]] = {}
    dependencies: dict[str, frozenset[str]] = {}
    normal_dependencies: dict[str, frozenset[str]] = {}
    features: dict[str, frozenset[str]] = {}
    manifests: set[Path] = set()
    for package in packages:
        name = package["name"]
        manifests.add(Path(package["manifest_path"]).resolve())
        package_dependencies = frozenset(item["name"] for item in package["dependencies"])
        dependencies[name] = package_dependencies
        normal_dependencies[name] = frozenset(
            item["name"]
            for item in package["dependencies"]
            if (item.get("kind") or "normal") == "normal"
        )
        features[name] = frozenset(package["features"])
        for dependency in package["dependencies"]:
            if dependency["name"] not in names:
                continue
            edge = Edge(name, dependency["name"])
            kind = dependency.get("kind") or "normal"
            target = dependency.get("target") or "*"
            optional = str(bool(dependency.get("optional"))).lower()
            rename = dependency.get("rename") or "-"
            evidence.setdefault(edge, set()).add(
                f"kind={kind}, optional={optional}, target={target}, rename={rename}"
            )

    return Snapshot(
        packages=names,
        manifests=frozenset(manifests),
        edges={edge: tuple(sorted(items)) for edge, items in evidence.items()},
        dependencies=dependencies,
        normal_dependencies=normal_dependencies,
        features=features,
    )


def _first_cycle(packages: Iterable[str], edges: Iterable[Edge]) -> tuple[str, ...] | None:
    graph = {name: [] for name in packages}
    for edge in edges:
        graph[edge.dependent].append(edge.dependency)
    for dependencies in graph.values():
        dependencies.sort()

    state: dict[str, int] = {name: 0 for name in graph}
    stack: list[str] = []
    stack_positions: dict[str, int] = {}

    def visit(name: str) -> tuple[str, ...] | None:
        state[name] = 1
        stack_positions[name] = len(stack)
        stack.append(name)
        for dependency in graph[name]:
            if state[dependency] == 0:
                cycle = visit(dependency)
                if cycle is not None:
                    return cycle
            elif state[dependency] == 1:
                start = stack_positions[dependency]
                return tuple(stack[start:] + [dependency])
        stack.pop()
        stack_positions.pop(name)
        state[name] = 2
        return None

    for name in sorted(graph):
        if state[name] == 0:
            cycle = visit(name)
            if cycle is not None:
                return cycle
    return None


def audit_snapshot(snapshot: Snapshot, policy: Policy) -> list[str]:
    """Return deterministic violations for one resolved workspace snapshot."""

    violations: list[str] = []
    missing_policy = snapshot.packages - policy.packages
    stale_policy = policy.packages - snapshot.packages
    if missing_policy:
        violations.append(
            "workspace packages lack topology policy: " + ", ".join(sorted(missing_policy))
        )
    if stale_policy:
        violations.append(
            "topology policy names absent workspace packages: "
            + ", ".join(sorted(stale_policy))
        )

    retired_facade_features = (
        snapshot.features.get("litchi", frozenset()) & RETIRED_FACADE_FEATURES
    )
    if retired_facade_features:
        violations.append(
            "retired litchi facade features returned: "
            + ", ".join(sorted(retired_facade_features))
        )

    actual_edges = frozenset(snapshot.edges)
    known_edges = policy.canonical_edges | policy.migration_edges
    for edge in sorted(actual_edges - known_edges):
        evidence = "; ".join(snapshot.edges[edge])
        violations.append(f"unclassified internal edge {edge.display()} ({evidence})")
    for edge in sorted(policy.canonical_edges - actual_edges):
        violations.append(
            f"resolved canonical edge still listed: {edge.display()}; remove its policy entry"
        )
    for edge in sorted(policy.migration_edges - actual_edges):
        violations.append(
            f"resolved migration debt still listed: {edge.display()}; remove its policy entry"
        )

    cycle = _first_cycle(snapshot.packages, actual_edges)
    if cycle is not None:
        violations.append("workspace dependency cycle: " + " -> ".join(cycle))

    for family_name, family in (("OOXML", OOXML_FORMATS), ("OLE", OLE_FORMATS)):
        for name in sorted(family & snapshot.packages):
            peers = snapshot.dependencies.get(name, frozenset()) & (family - {name})
            if peers:
                violations.append(
                    f"{family_name} concrete peer edge from {name}: " + ", ".join(sorted(peers))
                )

    for common, family in sorted(COMMON_FAMILY_GUARDS.items()):
        if common not in snapshot.packages:
            continue
        concrete = snapshot.dependencies.get(common, frozenset()) & family
        if concrete:
            violations.append(
                f"foundation crate {common} depends upward on: "
                + ", ".join(sorted(concrete))
            )

    for name in sorted(policy.runtime_neutral & snapshot.packages):
        runtimes = (
            snapshot.normal_dependencies.get(name, frozenset())
            & policy.runtime_packages
        )
        if runtimes:
            violations.append(
                f"runtime-neutral crate {name} depends on: " + ", ".join(sorted(runtimes))
            )

    core_dependencies = snapshot.dependencies.get("litchi-core", frozenset())
    active_forbidden = core_dependencies & policy.core_forbidden_dependencies
    internal_core_debt = {
        debt.edge.dependency
        for debt in policy.migration_debt
        if debt.edge.dependent == "litchi-core"
    }
    named_core_debt = {item.name for item in policy.core_dependency_debt}
    approved_core_debt = internal_core_debt | named_core_debt
    added_core_debt = active_forbidden - approved_core_debt
    if added_core_debt:
        violations.append(
            "litchi-core added forbidden dependencies: " + ", ".join(sorted(added_core_debt))
        )
    stale_core_debt = named_core_debt - active_forbidden
    if stale_core_debt:
        violations.append(
            "resolved litchi-core dependency debt still listed: "
            + ", ".join(sorted(stale_core_debt))
        )

    core_features = snapshot.features.get("litchi-core", frozenset())
    active_format_features = core_features & policy.core_format_features
    feature_debt = {item.name for item in policy.core_feature_debt}
    added_feature_debt = active_format_features - feature_debt
    if added_feature_debt:
        violations.append(
            "litchi-core added forbidden format features: "
            + ", ".join(sorted(added_feature_debt))
        )
    stale_feature_debt = feature_debt - active_format_features
    if stale_feature_debt:
        violations.append(
            "resolved litchi-core feature debt still listed: "
            + ", ".join(sorted(stale_feature_debt))
        )

    return sorted(set(violations))


def audit_manifest_inventory(snapshot: Snapshot) -> list[str]:
    manifests = frozenset(path.resolve() for path in (ROOT / "crates").glob("*/Cargo.toml"))
    missing = manifests - snapshot.manifests
    outside = snapshot.manifests - manifests
    violations: list[str] = []
    if missing:
        violations.append(
            "crate manifests are not audited workspace packages: "
            + ", ".join(str(path.relative_to(ROOT)) for path in sorted(missing))
        )
    if outside:
        violations.append(
            "workspace package manifests fall outside crates/*/Cargo.toml: "
            + ", ".join(str(path) for path in sorted(outside))
        )
    return violations


def audit_xlsb_source_topology(root: Path = ROOT) -> list[str]:
    """Reject retired XLSX implementation paths from the XLSB crate."""

    source_root = root / XLSB_SOURCE_ROOT
    violations: list[str] = []
    host_xlsx = source_root / "host" / "xlsx"
    if host_xlsx.exists():
        violations.append(
            "retired XLSB host XLSX source returned: "
            + str(host_xlsx.relative_to(root))
        )

    package_root = source_root / "package"
    if package_root.is_dir():
        for path in sorted(package_root.rglob("*.rs")):
            for line_number, line in enumerate(
                path.read_text(encoding="utf-8").splitlines(), start=1
            ):
                if PUBLIC_XLSX_MODULE.match(line):
                    violations.append(
                        "retired XLSB package XLSX module: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            for line_number, line in enumerate(
                path.read_text(encoding="utf-8").splitlines(), start=1
            ):
                if PACKAGE_XLSX_PATH.search(line):
                    violations.append(
                        "retired XLSB package::xlsx path: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    return sorted(set(violations))


def audit_spreadsheet_chart_source_topology(root: Path = ROOT) -> list[str]:
    """Keep shared spreadsheet-chart ownership out of OOXML format hosts."""

    violations: list[str] = []
    for retired in RETIRED_XLSX_CHART_FILES:
        path = root / XLSX_SOURCE_ROOT / retired
        if path.exists():
            violations.append(
                "retired XLSX chart owner source returned: "
                + str(path.relative_to(root))
            )

    for host, facades in SPREADSHEET_CHART_FACADES.items():
        for path in facades:
            absolute_path = root / path
            if not absolute_path.is_file():
                continue
            lines = absolute_path.read_text(encoding="utf-8").splitlines()
            if len(lines) > MAX_SPREADSHEET_CHART_FACADE_LINES:
                violations.append(
                    f"{host} chart facade exceeds "
                    f"{MAX_SPREADSHEET_CHART_FACADE_LINES} lines: "
                    f"{path}"
                )
            for line_number, line in enumerate(lines, start=1):
                if LOCAL_SHARED_CHART_TYPE.match(line):
                    violations.append(
                        f"{host} chart facade defines shared chart type: "
                        f"{path}:{line_number}"
                    )
                if DRAWINGML_CHART_CODEC.search(line):
                    violations.append(
                        f"{host} chart facade directly uses litchi_drawingml chart codec: "
                        f"{path}:{line_number}"
                    )

    return sorted(set(violations))


def audit_spreadsheet_shape_source_topology(root: Path = ROOT) -> list[str]:
    """Keep canonical spreadsheet-shape ownership out of OOXML format hosts."""

    violations: list[str] = []
    for retired in RETIRED_XLSX_SHAPE_FILES:
        path = root / XLSX_SOURCE_ROOT / retired
        if path.exists():
            violations.append(
                "retired XLSX shape owner source returned: "
                + str(path.relative_to(root))
            )

    for host, facades in SPREADSHEET_SHAPE_FACADES.items():
        for path in facades:
            absolute_path = root / path
            if not absolute_path.is_file():
                continue
            lines = absolute_path.read_text(encoding="utf-8").splitlines()
            if len(lines) > MAX_SPREADSHEET_SHAPE_FACADE_LINES:
                violations.append(
                    f"{host} shape facade exceeds "
                    f"{MAX_SPREADSHEET_SHAPE_FACADE_LINES} lines: {path}"
                )
            for line_number, line in enumerate(lines, start=1):
                if LOCAL_HOST_SHAPE_TYPE.match(line):
                    violations.append(
                        f"{host} shape facade defines local shape type: "
                        f"{path}:{line_number}"
                    )
                if QUICK_XML_USE.search(line):
                    violations.append(
                        f"{host} shape facade directly uses quick_xml: "
                        f"{path}:{line_number}"
                    )
                if XDR_XML_EMISSION.search(line):
                    violations.append(
                        f"{host} shape facade directly emits xdr XML: "
                        f"{path}:{line_number}"
                    )

    for host, source_root in (
        ("litchi-xlsb", XLSB_SOURCE_ROOT),
        ("litchi-xlsx", XLSX_SOURCE_ROOT),
    ):
        absolute_source_root = root / source_root
        if not absolute_source_root.is_dir():
            continue
        for path in sorted(absolute_source_root.rglob("*.rs")):
            for line_number, line in enumerate(
                path.read_text(encoding="utf-8").splitlines(), start=1
            ):
                for match in LEGACY_HOST_SHAPE_NAME.finditer(line):
                    violations.append(
                        f"{host} legacy shape host name {match.group(0)}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    return sorted(set(violations))


def debt_report(policy: Policy, explain: bool) -> list[str]:
    lines = ["ordered migration debt:"]
    items: list[tuple[int, str, str, str]] = []
    for item in policy.core_dependency_debt:
        items.append(
            (item.order, f"litchi-core dependency {item.name}", item.reason, item.exit)
        )
    for item in policy.core_feature_debt:
        items.append(
            (item.order, f"litchi-core feature {item.name}", item.reason, item.exit)
        )
    for item in policy.migration_debt:
        items.append((item.order, item.edge.display(), item.reason, item.exit))
    for order, label, reason, exit_condition in sorted(items):
        lines.append(f"  [{order:03}] {label}")
        if explain:
            lines.extend((f"        reason: {reason}", f"        exit: {exit_condition}"))
    return lines


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--policy",
        type=Path,
        default=DEFAULT_POLICY,
        help="checked-in JSON topology policy",
    )
    parser.add_argument(
        "--explain",
        action="store_true",
        help="print reasons and exit conditions for every migration debt item",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        policy = load_policy(args.policy)
    except PolicyError as error:
        print(f"crate-boundary policy error: {error}", file=sys.stderr)
        return 2

    snapshot = snapshot_from_metadata(cargo_metadata())
    violations = (
        audit_manifest_inventory(snapshot)
        + audit_snapshot(snapshot, policy)
        + audit_xlsb_source_topology()
        + audit_spreadsheet_chart_source_topology()
        + audit_spreadsheet_shape_source_topology()
    )
    if violations:
        for violation in sorted(set(violations)):
            print(f"crate-boundary error: {violation}", file=sys.stderr)
        return 1

    declaration_count = sum(len(items) for items in snapshot.edges.values())
    debt_count = (
        len(policy.migration_debt)
        + len(policy.core_dependency_debt)
        + len(policy.core_feature_debt)
    )
    print(
        f"crate boundaries valid for {len(snapshot.packages)} workspace packages and "
        f"{declaration_count} internal dependency declarations ({debt_count} explicit debt items)"
    )
    for line in debt_report(policy, args.explain):
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
