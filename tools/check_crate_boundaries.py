#!/usr/bin/env python3
"""Reject workspace dependency edges that violate the accepted crate topology."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass, field
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
RETIRED_FACADE_FEATURES = frozenset(
    {
        "eval_engine",
        "eval_engine_web_functions",
        "full",
        "imgconv",
        "iwa",
        "ooxml_encryption",
    }
)
XLSB_SOURCE_ROOT = Path("crates/litchi-xlsb/src")
XLSX_SOURCE_ROOT = Path("crates/litchi-xlsx/src")
RETIRED_SHEET_VIEW_OWNER_SOURCES = (
    ("litchi-xlsb", XLSB_SOURCE_ROOT / "views.rs"),
    ("litchi-xlsx", XLSX_SOURCE_ROOT / "views.rs"),
)
RETIRED_XLSX_SHEET_VIEW_OWNER_TREE = XLSX_SOURCE_ROOT / "views"
XLSB_SHEET_VIEW_ADAPTER = XLSB_SOURCE_ROOT / "host/sheet_view.rs"
XLSX_SHEET_VIEW_MODEL = XLSX_SOURCE_ROOT / "sheet_view/model.rs"
LEGACY_XLSB_SHEET_VIEW_NAMES = (
    "SheetPane",
    "SheetPanePosition",
    "SheetPaneState",
    "SheetSelection",
    "SheetView",
    "SheetViewType",
)
LEGACY_XLSB_SHEET_VIEW_NAME = re.compile(
    r"\b(?:" + "|".join(LEGACY_XLSB_SHEET_VIEW_NAMES) + r")\b"
)
LEGACY_XLSB_SHEET_VIEW_METHOD = re.compile(
    r"^\s*pub(?:\([^)]*\))?\s+fn\s+(?:set_sheet_view|sheet_views)\b"
)
CANONICAL_SHEET_VIEW_TYPES = (
    "Display",
    "Mode",
    "Pane",
    "Position",
    "Selection",
    "State",
    "View",
    "Zoom",
)
LOCAL_CANONICAL_SHEET_VIEW_TYPE = re.compile(
    r"^\s*(?:pub(?:\([^)]*\))?\s+)?(?:struct|enum|type|union)\s+(?:"
    + "|".join(CANONICAL_SHEET_VIEW_TYPES)
    + r")\b"
)
FACADE_PACKAGE = "litchi"
FACADE_REQUIRED_NORMAL_DEPENDENCIES = frozenset({"litchi-core"})
RETIRED_FACADE_DEPENDENCIES = frozenset({"litchi-iwa"})
IWA_KEYNOTE_SOURCE_ROOT = Path("crates/litchi-iwa/src/keynote")
RETIRED_IWA_KEYNOTE_METHODS = (
    "set_slide_title",
    "replace_slide_title",
    "clear_slide_title",
    "set_slide_body",
    "replace_slide_body",
    "clear_slide_body",
    "set_slide_notes",
    "replace_slide_notes",
    "clear_slide_notes",
    "slide_storage",
    "slide_notes_storage",
)
RETIRED_IWA_KEYNOTE_METHOD_SET = frozenset(RETIRED_IWA_KEYNOTE_METHODS)
IWA_NUMBERS_SOURCE_ROOT = Path("crates/litchi-iwa/src/numbers")
IWA_TABLE_LOCK_SOURCE = Path("crates/litchi-iwa/src/table_lock.rs")
IWA_NUMBERS_TABLE_INFO_SOURCE = (
    IWA_NUMBERS_SOURCE_ROOT / "editor" / "semantic" / "model.rs"
)
NUMBERS_SOURCE_ROOT = Path("crates/litchi-numbers/src")
NUMBERS_TABLE_LOCK_IMPLEMENTATION_SOURCES = (
    NUMBERS_SOURCE_ROOT / "package" / "table_lock.rs",
    NUMBERS_SOURCE_ROOT / "table" / "lock.rs",
)
NUMBERS_TABLE_LOCK_EXPORT_SOURCES = (
    NUMBERS_SOURCE_ROOT / "lib.rs",
    NUMBERS_SOURCE_ROOT / "package.rs",
    NUMBERS_SOURCE_ROOT / "table.rs",
)
RETIRED_IWA_NUMBERS_TABLE_LOCK_METHODS = (
    "table_lock_state",
    "set_table_lock_state",
    "table_lock_context",
    "set_table_lock_state_for_model",
    "table_lock_state_for_model",
)
RETIRED_IWA_NUMBERS_HOST_TABLE_LOCK_METHODS = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_LOCK_METHODS[:3]
)
RETIRED_IWA_NUMBERS_SHARED_TABLE_LOCK_METHODS = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_LOCK_METHODS[3:]
)
RETIRED_IWA_NUMBERS_TABLE_INFO_FIELDS = frozenset({"lock_state"})
RUST_FUNCTION_DECLARATION = re.compile(
    r"(?<![A-Za-z0-9_])fn\s+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)\b"
)
RUST_PUBLIC_DECLARATION = re.compile(
    r"(?<![A-Za-z0-9_#])pub(?![ \t\r\n]*\()[ \t\r\n]+"
    r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
)
RUST_IMPL_DECLARATION = re.compile(r"^[ \t]*impl\b", re.MULTILINE)
RUST_IDENTIFIER = re.compile(r"(?:r#)?([A-Za-z_][A-Za-z0-9_]*)")
RUST_ITEM_KEYWORDS = frozenset(
    {"const", "enum", "fn", "mod", "static", "struct", "trait", "type", "union", "use"}
)
RUST_FUNCTION_QUALIFIERS = frozenset({"async", "const", "extern", "unsafe"})
RUST_BRACED_ITEM_KEYWORDS = frozenset({"enum", "trait"})
RUST_SEMICOLON_ITEM_KEYWORDS = frozenset({"const", "mod", "static", "type", "use"})
NUMBERS_TABLE_LOCK_PUBLIC_MARKERS = frozenset(
    {
        "LockState",
        "State",
        "TableLockCommit",
        "TableLockDiagnostics",
        "TableLockEdit",
        "TableLockError",
        "TableLockLimitKind",
        "TableLockPatch",
    }
)
NUMBERS_TABLE_LOCK_ALLOWED_COMMON_REEXPORT = (
    "pub",
    "use",
    "litchi_iwa_common",
    "table",
    "lock",
    "State",
)
IWA_PAGES_SOURCE_ROOT = Path("crates/litchi-iwa/src/pages")
IWA_PAGES_EDITOR_SOURCE = IWA_PAGES_SOURCE_ROOT / "editor.rs"
RETIRED_IWA_PAGES_PAGE_LAYOUT_SOURCE = IWA_PAGES_SOURCE_ROOT / "editor" / "page_layout.rs"
RETIRED_IWA_PAGES_PAGE_LAYOUT_METHODS = ("page_layout", "set_page_layout")
RETIRED_IWA_PAGES_PAGE_LAYOUT_METHOD_SET = frozenset(
    RETIRED_IWA_PAGES_PAGE_LAYOUT_METHODS
)
PAGES_SOURCE_ROOT = Path("crates/litchi-pages/src")
PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "page_layout.rs"
)
PAGES_PAGE_LAYOUT_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
)
PAGES_PAGE_LAYOUT_PUBLIC_MARKERS = frozenset(
    {
        "PageLayoutCommit",
        "PageLayoutDiagnostics",
        "PageLayoutEdit",
        "PageLayoutError",
        "PageLayoutLimitKind",
        "PageLayoutPatch",
    }
)
IWA_PAGES_PAGE_LAYOUT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?page_layout\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
CAMEL_CASE_WORD = re.compile(r"[A-Z]+(?=[A-Z][a-z]|$)|[A-Z]?[a-z]+|[0-9]+")
FACADE_DEFAULT_FEATURE = "default"
FACADE_ALL_FEATURE = "all"
FACADE_SOURCE_ROOT = Path("crates/litchi/src")
PUBLIC_FACADE_IWA_MODULE = re.compile(
    r"^[ \t]*pub(?:\([^()]*\))?[ \t\r\n]+mod[ \t\r\n]+iwa\b",
    re.MULTILINE,
)
PUBLIC_FACADE_IWA_REEXPORT = re.compile(
    r"^[ \t]*pub(?:\([^()]*\))?[ \t\r\n]+use[ \t\r\n]+"
    r"(?:(?:crate[ \t\r\n]*::[ \t\r\n]*)?iwa|litchi_iwa)\b",
    re.MULTILINE,
)
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
    dev_only_edges: frozenset[Edge]
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
    dependency_kinds: dict[Edge, frozenset[str]]
    dependencies: dict[str, frozenset[str]]
    normal_dependencies: dict[str, frozenset[str]]
    features: dict[str, frozenset[str]]
    feature_definitions: dict[str, dict[str, frozenset[str]]] = field(
        default_factory=dict
    )
    normal_optional_dependencies: dict[str, frozenset[str]] = field(
        default_factory=dict
    )


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

    dev_only_raw = raw.get("dev_only_edges")
    if not isinstance(dev_only_raw, list):
        raise PolicyError("dev_only_edges must be a list")
    dev_only: list[Edge] = []
    for index, item in enumerate(dev_only_raw):
        context = f"dev_only_edges[{index}]"
        if not isinstance(item, dict):
            raise PolicyError(f"{context} must be an object")
        dev_only.append(
            Edge(
                _require_string(item, "dependent", context),
                _require_string(item, "dependency", context),
            )
        )
    if dev_only != sorted(dev_only):
        raise PolicyError("dev_only_edges must be ordered by edge")
    dev_only_edges = frozenset(dev_only)
    if len(dev_only_edges) != len(dev_only):
        raise PolicyError("dev_only_edges contains duplicate edges")
    noncanonical_dev_only = sorted(dev_only_edges - canonical)
    if noncanonical_dev_only:
        raise PolicyError(
            "dev-only annotations must reference canonical edges: "
            + ", ".join(edge.display() for edge in noncanonical_dev_only)
        )

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
    incoming_host_edges = sorted(
        edge
        for edge in canonical | migration_edges
        if edge.dependency in migration_hosts
    )
    if incoming_host_edges:
        raise PolicyError(
            "migration hosts cannot be workspace dependencies: "
            + ", ".join(edge.display() for edge in incoming_host_edges)
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
        dev_only_edges=dev_only_edges,
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
    dependency_kinds: dict[Edge, set[str]] = {}
    dependencies: dict[str, frozenset[str]] = {}
    normal_dependencies: dict[str, frozenset[str]] = {}
    features: dict[str, frozenset[str]] = {}
    feature_definitions: dict[str, dict[str, frozenset[str]]] = {}
    normal_optional_dependencies: dict[str, frozenset[str]] = {}
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
        feature_definitions[name] = {
            feature: frozenset(references)
            for feature, references in package["features"].items()
        }
        normal_optional_dependencies[name] = frozenset(
            item["name"]
            for item in package["dependencies"]
            if (item.get("kind") or "normal") == "normal" and item.get("optional")
        )
        for dependency in package["dependencies"]:
            if dependency["name"] not in names:
                continue
            edge = Edge(name, dependency["name"])
            kind = dependency.get("kind") or "normal"
            dependency_kinds.setdefault(edge, set()).add(kind)
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
        dependency_kinds={
            edge: frozenset(kinds) for edge, kinds in dependency_kinds.items()
        },
        dependencies=dependencies,
        normal_dependencies=normal_dependencies,
        features=features,
        feature_definitions=feature_definitions,
        normal_optional_dependencies=normal_optional_dependencies,
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

    actual_dev_only = frozenset(
        edge
        for edge, kinds in snapshot.dependency_kinds.items()
        if kinds == frozenset({"dev"})
    )
    for edge in sorted(actual_dev_only - policy.dev_only_edges):
        violations.append(
            f"dev-only internal edge lacks policy annotation: {edge.display()}"
        )
    for edge in sorted(policy.dev_only_edges - actual_edges):
        violations.append(
            f"resolved dev-only edge still annotated: {edge.display()}; "
            "remove its policy annotation"
        )
    for edge in sorted(policy.dev_only_edges & actual_edges):
        kinds = snapshot.dependency_kinds.get(edge, frozenset())
        if kinds != frozenset({"dev"}):
            evidence = "; ".join(snapshot.edges[edge])
            violations.append(
                f"dev-only internal edge used outside dev: {edge.display()} "
                f"({evidence})"
            )

    for edge in sorted(
        edge for edge in actual_edges if edge.dependency in policy.migration_hosts
    ):
        violations.append(f"workspace dependency targets migration host: {edge.display()}")

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

    violations.extend(audit_litchi_facade(snapshot))

    return sorted(set(violations))


def _facade_all_dependencies(
    feature_definitions: dict[str, frozenset[str]],
) -> frozenset[str]:
    """Resolve direct and aggregate feature ownership from litchi's `all` feature."""

    dependencies: set[str] = set()
    pending = [FACADE_ALL_FEATURE]
    visited: set[str] = set()
    while pending:
        feature = pending.pop()
        if feature in visited:
            continue
        visited.add(feature)
        for reference in feature_definitions.get(feature, frozenset()):
            if reference.startswith("dep:"):
                dependencies.add(reference.removeprefix("dep:"))
            elif "/" not in reference:
                pending.append(reference)
    return frozenset(dependencies)


def audit_litchi_facade(snapshot: Snapshot) -> list[str]:
    """Enforce litchi's all-optional, feature-owned facade contract."""

    if FACADE_PACKAGE not in snapshot.packages:
        return []

    normal_dependencies = snapshot.normal_dependencies.get(FACADE_PACKAGE, frozenset())
    optional_dependencies = snapshot.normal_optional_dependencies.get(
        FACADE_PACKAGE, frozenset()
    )
    feature_definitions = snapshot.feature_definitions.get(FACADE_PACKAGE, {})
    violations: list[str] = []

    required = normal_dependencies - optional_dependencies
    unexpected_required = required - FACADE_REQUIRED_NORMAL_DEPENDENCIES
    if unexpected_required:
        violations.append(
            "litchi facade has non-optional normal dependencies: "
            + ", ".join(sorted(unexpected_required))
        )
    missing_required = FACADE_REQUIRED_NORMAL_DEPENDENCIES - required
    if missing_required:
        violations.append(
            "litchi facade is missing required normal dependencies: "
            + ", ".join(sorted(missing_required))
        )

    retired_dependencies = (
        snapshot.dependencies.get(FACADE_PACKAGE, frozenset())
        & RETIRED_FACADE_DEPENDENCIES
    )
    if retired_dependencies:
        violations.append(
            "litchi facade depends on retired packages: "
            + ", ".join(sorted(retired_dependencies))
        )

    if feature_definitions.get(FACADE_DEFAULT_FEATURE) != frozenset():
        violations.append("litchi default feature must be exactly empty")
    if FACADE_ALL_FEATURE not in feature_definitions:
        violations.append("litchi facade is missing the all feature")

    owned_dependencies: set[str] = set()
    stale_dependencies: set[str] = set()
    unknown_feature_references: list[str] = []
    unknown_dependency_references: list[str] = []
    for feature, references in sorted(feature_definitions.items()):
        for reference in sorted(references):
            if reference.startswith("dep:"):
                dependency = reference.removeprefix("dep:")
                if dependency in optional_dependencies:
                    owned_dependencies.add(dependency)
                else:
                    stale_dependencies.add(dependency)
                continue
            if "/" in reference:
                dependency = reference.split("/", maxsplit=1)[0].removesuffix("?")
                if dependency not in normal_dependencies:
                    unknown_dependency_references.append(
                        f"{feature} -> {reference}"
                    )
                continue
            if reference not in feature_definitions:
                unknown_feature_references.append(f"{feature} -> {reference}")

    missing_ownership = optional_dependencies - owned_dependencies
    if missing_ownership:
        violations.append(
            "litchi facade optional dependencies lack dep: feature ownership: "
            + ", ".join(sorted(missing_ownership))
        )
    if stale_dependencies:
        violations.append(
            "litchi facade has stale dep: feature references: "
            + ", ".join(sorted(stale_dependencies))
        )
    if unknown_feature_references:
        violations.append(
            "litchi facade feature references unknown features: "
            + ", ".join(unknown_feature_references)
        )
    if unknown_dependency_references:
        violations.append(
            "litchi facade feature references unknown dependencies: "
            + ", ".join(unknown_dependency_references)
        )

    omitted_from_all = optional_dependencies - _facade_all_dependencies(feature_definitions)
    if omitted_from_all:
        violations.append(
            "litchi all feature omits optional dependencies: "
            + ", ".join(sorted(omitted_from_all))
        )

    return sorted(set(violations))


def audit_litchi_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Reject public re-exports of the retired monolithic iWork facade."""

    source_root = root / FACADE_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_non_code(path.read_text(encoding="utf-8"))
        for pattern, label in (
            (PUBLIC_FACADE_IWA_MODULE, "module"),
            (PUBLIC_FACADE_IWA_REEXPORT, "re-export"),
        ):
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start()) + 1
                violations.append(
                    f"retired litchi facade public iwa {label}: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _mask_rust_non_code(source: str) -> str:
    """Mask Rust comments and literals while preserving offsets and newlines."""

    masked = list(source)

    def mask(start: int, end: int) -> None:
        for offset in range(start, end):
            if masked[offset] != "\n":
                masked[offset] = " "

    def raw_string_end(start: int) -> int | None:
        cursor = start
        if source.startswith(("br", "cr"), cursor):
            cursor += 2
        elif source.startswith("r", cursor):
            cursor += 1
        else:
            return None
        hashes_start = cursor
        while cursor < len(source) and source[cursor] == "#":
            cursor += 1
        if cursor >= len(source) or source[cursor] != '"':
            return None
        terminator = '"' + source[hashes_start:cursor]
        closing = source.find(terminator, cursor + 1)
        return len(source) if closing < 0 else closing + len(terminator)

    def quoted_string_end(start: int) -> int:
        cursor = start + 1
        while cursor < len(source):
            if source[cursor] == "\\":
                cursor += 2
            elif source[cursor] == '"':
                return cursor + 1
            else:
                cursor += 1
        return len(source)

    def character_literal_end(start: int) -> int | None:
        cursor = start + 1
        if cursor >= len(source) or source[cursor] in ("\n", "\r", "'"):
            return None
        if source[cursor] == "\\":
            cursor += 1
            if cursor >= len(source):
                return len(source)
            if source[cursor] == "u" and cursor + 1 < len(source) and source[cursor + 1] == "{":
                closing = source.find("}", cursor + 2)
                if closing < 0:
                    return len(source)
                cursor = closing + 1
            else:
                cursor += 1
        else:
            cursor += 1
        return cursor + 1 if cursor < len(source) and source[cursor] == "'" else None

    cursor = 0
    while cursor < len(source):
        if source.startswith("//", cursor):
            end = source.find("\n", cursor + 2)
            end = len(source) if end < 0 else end
            mask(cursor, end)
            cursor = end
            continue
        if source.startswith("/*", cursor):
            depth = 1
            end = cursor + 2
            while end < len(source) and depth:
                if source.startswith("/*", end):
                    depth += 1
                    end += 2
                elif source.startswith("*/", end):
                    depth -= 1
                    end += 2
                else:
                    end += 1
            mask(cursor, end)
            cursor = end
            continue
        raw_end = raw_string_end(cursor)
        if raw_end is not None:
            mask(cursor, raw_end)
            cursor = raw_end
            continue
        if source[cursor] == '"':
            end = quoted_string_end(cursor)
            mask(cursor, end)
            cursor = end
            continue
        if source[cursor] == "'":
            end = character_literal_end(cursor)
            if end is not None:
                mask(cursor, end)
                cursor = end
                continue
        cursor += 1

    return "".join(masked)


def _rust_function_declarations(source: str) -> list[tuple[str, int]]:
    """Return exact Rust function declaration names and source line numbers."""

    code = _mask_rust_non_code(source)
    declarations: list[tuple[str, int]] = []
    line_number = 1
    previous_offset = 0
    for match in RUST_FUNCTION_DECLARATION.finditer(code):
        name_offset = match.start(1)
        line_number += code.count("\n", previous_offset, name_offset)
        declarations.append((match.group(1), line_number))
        previous_offset = name_offset
    return declarations


def _rust_public_declarations(source: str) -> list[tuple[str, int]]:
    """Return public Rust declaration text without descending into function bodies."""

    code = _mask_rust_non_code(source)
    declarations: list[tuple[str, int]] = []
    for match in RUST_PUBLIC_DECLARATION.finditer(code):
        leading_identifier = list(RUST_IDENTIFIER.finditer(match.group(0)))[-1].group(1)
        item_keyword = (
            leading_identifier if leading_identifier in RUST_ITEM_KEYWORDS else None
        )
        if leading_identifier in RUST_FUNCTION_QUALIFIERS:
            for identifier_match in RUST_IDENTIFIER.finditer(code, match.end()):
                identifier = identifier_match.group(1)
                if identifier == "fn":
                    item_keyword = identifier
                elif leading_identifier == "const" and identifier not in RUST_FUNCTION_QUALIFIERS:
                    item_keyword = "const"
                elif identifier not in RUST_FUNCTION_QUALIFIERS:
                    item_keyword = None
                if identifier not in RUST_FUNCTION_QUALIFIERS:
                    break
        parentheses = 0
        brackets = 0
        cursor = match.end()
        end = len(code)
        while cursor < len(code):
            character = code[cursor]
            if character == "(":
                parentheses += 1
            elif character == ")" and parentheses:
                parentheses -= 1
            elif character == "[":
                brackets += 1
            elif character == "]" and brackets:
                brackets -= 1
            elif not parentheses and not brackets:
                if character == ";":
                    end = cursor + 1
                    break
                if character == "," and item_keyword is None:
                    end = cursor + 1
                    break
                if character == "{":
                    if item_keyword in RUST_SEMICOLON_ITEM_KEYWORDS:
                        cursor += 1
                        continue
                    if item_keyword in RUST_BRACED_ITEM_KEYWORDS:
                        depth = 1
                        cursor += 1
                        while cursor < len(code) and depth:
                            if code[cursor] == "{":
                                depth += 1
                            elif code[cursor] == "}":
                                depth -= 1
                            cursor += 1
                        end = cursor
                    else:
                        end = cursor
                    break
            cursor += 1
        declarations.append(
            (code[match.start() : end], code.count("\n", 0, match.start()) + 1)
        )
    return declarations


def _rust_impl_headers(source: str) -> list[tuple[str, int]]:
    """Return Rust impl headers, whose public trait relationships lack `pub`."""

    code = _mask_rust_non_code(source)
    headers: list[tuple[str, int]] = []
    for match in RUST_IMPL_DECLARATION.finditer(code):
        cursor = match.end()
        parentheses = 0
        brackets = 0
        while cursor < len(code):
            character = code[cursor]
            if character == "(":
                parentheses += 1
            elif character == ")" and parentheses:
                parentheses -= 1
            elif character == "[":
                brackets += 1
            elif character == "]" and brackets:
                brackets -= 1
            elif character == "{" and not parentheses and not brackets:
                break
            cursor += 1
        headers.append(
            (
                code[match.start() : cursor],
                code.count("\n", 0, match.start()) + 1,
            )
        )
    return headers


def _rust_named_struct_body(source: str, name: str) -> tuple[str, int] | None:
    """Return one exact struct body and its source offset, ignoring Rust trivia."""

    code = _mask_rust_non_code(source)
    declaration = re.search(
        rf"(?<![A-Za-z0-9_])struct[ \t\r\n]+{re.escape(name)}\b", code
    )
    if declaration is None:
        return None
    opening = code.find("{", declaration.end())
    semicolon = code.find(";", declaration.end())
    if opening < 0 or (semicolon >= 0 and semicolon < opening):
        return None
    depth = 1
    cursor = opening + 1
    while cursor < len(code) and depth:
        if code[cursor] == "{":
            depth += 1
        elif code[cursor] == "}":
            depth -= 1
        cursor += 1
    if depth:
        return None
    return code[opening + 1 : cursor - 1], opening + 1


def _iwork_public_leak(identifier: str) -> str | None:
    """Classify physical vocabulary forbidden in focused iWork facades."""

    if identifier.startswith("litchi_iwa"):
        return "archive/IWA type"
    if identifier in {"buffa", "prost", "prost_types"}:
        return "protobuf type"
    if identifier == "IWorkPackage":
        return "archive/IWA type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(
        word in {"guid", "guids", "id", "ids", "identifier", "identifiers", "uuid", "uuids"}
        for word in words
    ):
        return "raw identifier"
    if identifier[:1].islower() and any(
        word in {"object", "objects"} for word in words
    ):
        return "native object"
    if identifier[:1].isupper() and "object" in words:
        return "native object"
    if any(word in {"proto", "protobuf"} for word in words) or identifier.endswith(
        "Message"
    ):
        return "protobuf type"
    if identifier.endswith(("Archive", "ArchiveView", "MessageInfo")) or (
        "raw" in words and identifier[:1].isupper()
    ):
        return "archive/IWA type"
    return None


def _is_numbers_table_lock_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & NUMBERS_TABLE_LOCK_PUBLIC_MARKERS) or any(
        "table_lock" in identifier.lower() for identifier in identifiers
    )


def _is_pages_page_layout_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & PAGES_PAGE_LAYOUT_PUBLIC_MARKERS) or any(
        "page_layout" in identifier.lower() for identifier in identifiers
    )


def audit_iwa_keynote_source_topology(root: Path = ROOT) -> list[str]:
    """Prevent retired Keynote function declarations from returning to the host."""

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        source = path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_KEYNOTE_METHOD_SET:
                continue
            violations.append(
                "retired litchi-iwa Keynote method "
                f"{name}: {path.relative_to(root)}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_numbers_table_lock_source_topology(root: Path = ROOT) -> list[str]:
    """Keep retired Numbers table-lock APIs out of their former host scopes."""

    scoped_sources = (
        (
            root / IWA_NUMBERS_SOURCE_ROOT,
            RETIRED_IWA_NUMBERS_HOST_TABLE_LOCK_METHODS,
        ),
        (
            root / IWA_TABLE_LOCK_SOURCE,
            RETIRED_IWA_NUMBERS_SHARED_TABLE_LOCK_METHODS,
        ),
    )
    violations: list[str] = []
    for source_path, retired_names in scoped_sources:
        paths = (
            sorted(source_path.rglob("*.rs"))
            if source_path.is_dir()
            else [source_path]
            if source_path.is_file()
            else []
        )
        for path in paths:
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in retired_names:
                    continue
                violations.append(
                    "retired litchi-iwa Numbers table-lock method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    table_info_path = root / IWA_NUMBERS_TABLE_INFO_SOURCE
    if table_info_path.is_file():
        source = table_info_path.read_text(encoding="utf-8")
        body = _rust_named_struct_body(source, "NumbersTableInfo")
        if body is not None:
            body_source, body_offset = body
            for name in sorted(RETIRED_IWA_NUMBERS_TABLE_INFO_FIELDS):
                field = re.search(
                    rf"(?<![A-Za-z0-9_#])pub(?![ \t\r\n]*\()[ \t\r\n]+"
                    rf"(?:r#)?{re.escape(name)}[ \t\r\n]*:",
                    body_source,
                )
                if field is None:
                    continue
                line_number = source.count("\n", 0, body_offset + field.start()) + 1
                violations.append(
                    "retired litchi-iwa Numbers table-info field "
                    f"{name}: {table_info_path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_numbers_table_lock_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Reject physical identifiers and implementation types from the lock facade."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_LOCK_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path for path in NUMBERS_TABLE_LOCK_EXPORT_SOURCES if (root / path).is_file()
    }
    violations: list[str] = []
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for declaration, line_number, complete_source_scope in declarations:
            if not _is_numbers_table_lock_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use = declaration_identifiers[:2] == ["pub", "use"]
            if (
                path == root / NUMBERS_SOURCE_ROOT / "table" / "lock.rs"
                and tuple(declaration_identifiers)
                == NUMBERS_TABLE_LOCK_ALLOWED_COMMON_REEXPORT
            ):
                continue
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                if public_use and not (
                    identifier in NUMBERS_TABLE_LOCK_PUBLIC_MARKERS
                    or "lock" in identifier.lower()
                    or identifier.startswith("litchi_iwa")
                    or identifier in {"buffa", "prost", "prost_types"}
                ):
                    continue
                reason = _iwork_public_leak(identifier)
                if reason is None:
                    continue
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                violations.append(
                    "focused litchi-numbers table-lock public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )

    return sorted(set(violations))


def audit_iwa_pages_page_layout_source_topology(root: Path = ROOT) -> list[str]:
    """Keep the retired Pages page-layout API and module out of the host."""

    violations: list[str] = []
    retired_source = root / RETIRED_IWA_PAGES_PAGE_LAYOUT_SOURCE
    if retired_source.exists():
        violations.append(
            "retired litchi-iwa Pages page-layout source returned: "
            + str(RETIRED_IWA_PAGES_PAGE_LAYOUT_SOURCE)
        )

    source_root = root / IWA_PAGES_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_PAGES_PAGE_LAYOUT_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Pages page-layout method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    editor_path = root / IWA_PAGES_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_PAGES_PAGE_LAYOUT_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Pages page-layout module declaration: "
                f"{IWA_PAGES_EDITOR_SOURCE}:{line_number}"
            )

    return sorted(set(violations))


def audit_pages_page_layout_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Reject physical identifiers and implementation types from the layout facade."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    implementation_source = root / PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCE
    dedicated_sources = {implementation_source} if implementation_source.is_file() else set()
    export_sources = {
        root / path for path in PAGES_PAGE_LAYOUT_EXPORT_SOURCES if (root / path).is_file()
    }
    violations: list[str] = []
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for declaration, line_number, complete_source_scope in declarations:
            if not _is_pages_page_layout_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                reason = _iwork_public_leak(identifier)
                if reason is None:
                    continue
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                violations.append(
                    "focused litchi-pages page-layout public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
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


def audit_spreadsheet_sheet_view_source_topology(root: Path = ROOT) -> list[str]:
    """Keep canonical worksheet-view ownership out of OOXML format hosts."""

    violations: list[str] = []
    for host, path in RETIRED_SHEET_VIEW_OWNER_SOURCES:
        absolute_path = root / path
        if absolute_path.exists():
            violations.append(
                f"retired {host} sheet-view owner source returned: "
                f"{path}"
            )

    retired_tree = root / RETIRED_XLSX_SHEET_VIEW_OWNER_TREE
    if retired_tree.exists():
        violations.append(
            "retired litchi-xlsx sheet-view owner tree returned: "
            + str(RETIRED_XLSX_SHEET_VIEW_OWNER_TREE)
        )

    xlsb_source_root = root / XLSB_SOURCE_ROOT
    if xlsb_source_root.is_dir():
        for path in sorted(xlsb_source_root.rglob("*.rs")):
            for line_number, line in enumerate(
                path.read_text(encoding="utf-8").splitlines(), start=1
            ):
                for match in LEGACY_XLSB_SHEET_VIEW_NAME.finditer(line):
                    violations.append(
                        f"litchi-xlsb legacy sheet-view name {match.group(0)}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
                method = LEGACY_XLSB_SHEET_VIEW_METHOD.search(line)
                if method:
                    violations.append(
                        "litchi-xlsb legacy sheet-view public method "
                        f"{method.group(0).rsplit(None, 1)[-1]}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    for host, path, role in (
        ("litchi-xlsb", XLSB_SHEET_VIEW_ADAPTER, "adapter"),
        ("litchi-xlsx", XLSX_SHEET_VIEW_MODEL, "model"),
    ):
        absolute_path = root / path
        if not absolute_path.is_file():
            continue
        for line_number, line in enumerate(
            absolute_path.read_text(encoding="utf-8").splitlines(), start=1
        ):
            match = LOCAL_CANONICAL_SHEET_VIEW_TYPE.match(line)
            if match:
                type_name = match.group(0).rsplit(None, 1)[-1]
                violations.append(
                    f"{host} sheet-view {role} defines canonical view type {type_name}: "
                    f"{path}:{line_number}"
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
        + audit_litchi_facade_source_topology()
        + audit_iwa_keynote_source_topology()
        + audit_iwa_numbers_table_lock_source_topology()
        + audit_numbers_table_lock_facade_source_topology()
        + audit_iwa_pages_page_layout_source_topology()
        + audit_pages_page_layout_facade_source_topology()
        + audit_xlsb_source_topology()
        + audit_spreadsheet_sheet_view_source_topology()
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
