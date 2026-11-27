from __future__ import annotations

import copy
import re
from pathlib import Path
from tools import check_crate_boundaries as boundaries


KEYNOTE_CHART_LEGEND_RETAINED_STYLE_METHODS = frozenset(
    {
        "slide_chart_legend_fill",
        "set_slide_chart_legend_fill",
        "slide_chart_legend_frame",
        "set_slide_chart_legend_frame",
        "slide_chart_legend_font",
        "set_slide_chart_legend_font",
        "slide_chart_legend_font_size",
        "set_slide_chart_legend_font_size",
        "slide_chart_legend_stroke",
        "set_slide_chart_legend_stroke",
        "slide_chart_legend_shadow",
        "set_slide_chart_legend_shadow",
    }
)


SYNTHETIC_HOST = "litchi-legacy-host"


def synthetic_host_raw_policy(raw_policy: dict) -> dict:
    """The checked-in policy plus one synthetic migration host and debt edge.

    The workspace no longer has a migration host, but the
    policy keeps the mechanism; these tests exercise it on a synthetic host.
    """

    raw = copy.deepcopy(raw_policy)
    raw["packages"][SYNTHETIC_HOST] = []
    raw["packages"] = dict(sorted(raw["packages"].items()))
    raw["migration_hosts"] = [SYNTHETIC_HOST]
    raw["migration_debt"] = [
        {
            "order": 1,
            "dependent": SYNTHETIC_HOST,
            "dependency": "litchi-core",
            "reason": "Synthetic migration host used by the policy tests.",
            "exit": "Delete the synthetic host.",
        }
    ]
    return raw


def valid_snapshot(policy: boundaries.Policy) -> boundaries.Snapshot:
    edges = policy.canonical_edges | policy.migration_edges
    dependencies = {name: set() for name in policy.packages}
    normal_dependencies = {name: set() for name in policy.packages}
    for edge in edges:
        dependencies[edge.dependent].add(edge.dependency)
        if edge not in policy.dev_only_edges:
            normal_dependencies[edge.dependent].add(edge.dependency)
    dependencies["litchi-core"].update(item.name for item in policy.core_dependency_debt)
    normal_dependencies["litchi-core"].update(
        item.name for item in policy.core_dependency_debt
    )
    frozen_dependencies = {
        name: frozenset(items) for name, items in dependencies.items()
    }
    frozen_normal_dependencies = {
        name: frozenset(items) for name, items in normal_dependencies.items()
    }
    features = {name: frozenset() for name in policy.packages}
    features["litchi-core"] = frozenset(item.name for item in policy.core_feature_debt)
    facade_optional_dependencies = frozenset(
        edge.dependency
        for edge in edges
        if edge.dependent == "litchi" and edge.dependency != "litchi-core"
    )
    return boundaries.Snapshot(
        packages=policy.packages,
        manifests=frozenset(),
        edges={
            edge: (
                f"kind={'dev' if edge in policy.dev_only_edges else 'normal'}, "
                "optional=false, target=*, rename=-",
            )
            for edge in edges
        },
        dependency_kinds={
            edge: frozenset({"dev" if edge in policy.dev_only_edges else "normal"})
            for edge in edges
        },
        dependencies=frozen_dependencies,
        normal_dependencies=frozen_normal_dependencies,
        features=features,
        feature_definitions={
            "litchi": {
                "default": frozenset(),
                "all": frozenset(
                    f"dep:{dependency}" for dependency in facade_optional_dependencies
                ),
            }
        },
        normal_optional_dependencies={"litchi": facade_optional_dependencies},
    )


def write_cargo_manifest_fixture(root: Path, relative: str, source: str) -> Path:
    path = root / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(source, encoding="utf-8")
    return path


def copy_numbers_table_merge_reader_fixture(root: Path) -> None:
    """Copy the Numbers merge owner and its optional private reader child."""

    for relative in (
        boundaries.NUMBERS_TABLE_MERGE_SOURCE,
        boundaries.NUMBERS_TABLE_MERGE_READER_SOURCE,
    ):
        destination = root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(
            (boundaries.ROOT / relative).read_text(encoding="utf-8"),
            encoding="utf-8",
        )


def copy_focused_table_merge_reader_fixture(root: Path, format_name: str) -> None:
    """Copy one format's merge owner and whichever reader layout it uses."""

    if format_name == "Pages":
        owner = boundaries.PAGES_TABLE_MERGE_SOURCE
        reader = boundaries.PAGES_TABLE_MERGE_READER_SOURCE
    elif format_name == "Keynote":
        owner = boundaries.KEYNOTE_TABLE_MERGE_SOURCE
        reader = boundaries.KEYNOTE_TABLE_MERGE_READER_SOURCE
    else:
        raise AssertionError(f"unknown focused merge-reader format: {format_name}")

    paths = [owner]
    if (boundaries.ROOT / reader).is_file():
        paths.append(reader)
    for relative in paths:
        destination = root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(
            (boundaries.ROOT / relative).read_text(encoding="utf-8"),
            encoding="utf-8",
        )


def add_iwa_common_formula_render_ownership_scaffold(
    root: Path,
    *,
    common_source: str | None = None,
    reader_suffix: str = "",
) -> None:
    """Install the shared formula arena and both migrated reader imports."""

    common = root / boundaries.IWA_COMMON_FORMULA_RENDER_SOURCE
    common.parent.mkdir(parents=True, exist_ok=True)
    common.write_text(
        common_source
        if common_source is not None
        else (
            "use std::ops::Range;\n"
            "pub trait FormulaRenderBudget {}\n"
            "pub struct FormulaExpr;\n"
            "pub struct FormulaRenderer;\n"
            "impl FormulaRenderer { pub fn render() {} }\n"
        ),
        encoding="utf-8",
    )
    reader = (
        "use litchi_iwa_common::formula::render::{"
        "FormulaExpr, FormulaRenderBudget, FormulaRenderer};\n"
        "fn render_route() {}\n"
        + reader_suffix
    )
    for relative in boundaries.IWA_FORMULA_RENDER_READER_SOURCES:
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(reader, encoding="utf-8")


def add_iwa_numbers_wire_formula_render_scaffold(
    root: Path,
    *,
    wire_source: str | None = None,
    reader_suffix: str = "",
) -> None:
    """Install the low-level formula owner and migrated reader adapters."""

    wire = root / boundaries.IWA_NUMBERS_WIRE_FORMULA_RENDER_SOURCE
    wire.parent.mkdir(parents=True, exist_ok=True)
    wire.write_text(
        wire_source
        if wire_source is not None
        else (
            "use litchi_iwa_common::formula::render::{FormulaExpr, "
            "FormulaRenderBudget, FormulaRenderer};\n"
            "use litchi_iwa_protos::numbers_formula_codec;\n"
            "pub trait FormulaEventRenderBudget: FormulaRenderBudget {}\n"
            "pub trait ReferenceResolver {}\n"
            "pub fn render_scalar_formula_nodes() {}\n"
            "pub struct CompatibilityFormulaVisitor;\n"
            "pub struct FormulaRenderCodecVisitor;\n"
            "impl numbers_formula_codec::FormulaRenderVisitor "
            "for FormulaRenderCodecVisitor {}\n"
            "fn render_arena(_: FormulaExpr, _: FormulaRenderer) {}\n"
        ),
        encoding="utf-8",
    )
    reader = (
        "use litchi_numbers_wire::formula_render::{"
        "CompatibilityFormulaVisitor, FormulaRenderCodecVisitor, "
        "FormulaEventRenderBudget, ReferenceResolver, "
        "render_scalar_formula_nodes};\n"
        "fn preflight_formula_archive_envelope() {}\n"
        "fn decode_formula_archive_with_visitor() {}\n"
        "fn decode_formula_archive_for_render() {}\n"
        "fn charge_formula_wire() {}\n"
        "fn charge_formula_render_work() {}\n"
        + reader_suffix
    )
    for relative in boundaries.IWA_FORMULA_RENDER_READER_SOURCES:
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(reader, encoding="utf-8")


def add_iwa_numbers_wire_formula_envelope_scaffold(
    root: Path,
    *,
    envelope_source: str | None = None,
    reader_suffix: str = "",
) -> None:
    """Install the shared FormulaArchive envelope owner and reader imports."""

    envelope = root / boundaries.IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE
    envelope.parent.mkdir(parents=True, exist_ok=True)
    envelope.write_text(
        envelope_source
        if envelope_source is not None
        else (
            "use litchi_iwa_common::wire::{WireDescent, WireLimits};\n"
            "use litchi_iwa_protos::numbers_formula_codec;\n"
            "pub struct FormulaEnvelopeLimits;\n"
            "pub struct AttemptedFormulaEnvelopeCost;\n"
            "pub struct FormulaEnvelopeReport;\n"
            "pub struct FormulaEnvelopeFailure;\n"
            "pub struct FormulaEnvelopeCost;\n"
            "struct FormulaEnvelopeField;\n"
            "enum FormulaScalar {}\n"
            "pub fn preflight_formula_envelope() {}\n"
            "fn required_formula_ast_node_type() {}\n"
            "fn require_formula_fields() {}\n"
            "fn formula_field_is_repeated() {}\n"
            "fn validate_formula_scalar() {}\n"
            "fn formula_envelope_field() {}\n"
        ),
        encoding="utf-8",
    )
    reader = (
        "use litchi_numbers_wire::formula_envelope::"
        "preflight_formula_envelope;\n"
        "struct FormulaArchiveBytes;\n"
        "fn decode_formula_archive_with_visitor() {}\n"
        "fn decode_formula_archive_for_render() {}\n"
        + reader_suffix
    )
    for relative in boundaries.IWA_FORMULA_RENDER_READER_SOURCES:
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(reader, encoding="utf-8")


def add_iwa_table_cell_read_scaffold(
    root: Path,
    *,
    common_source: str | None = None,
    common_module_source: str | None = None,
    cells_source: str | None = None,
    sidecars_source: str | None = None,
    wire_lib_source: str | None = None,
    pages_source: str | None = None,
    pages_package_source: str | None = None,
    keynote_source: str | None = None,
    keynote_package_source: str | None = None,
    include_pages: bool = True,
    include_keynote: bool = True,
) -> None:
    """Install minimal common, wire, and focused table-cell owners."""

    common = root / boundaries.IWA_COMMON_TABLE_READ_SOURCE
    common.parent.mkdir(parents=True, exist_ok=True)
    common.write_text(
        common_source
        if common_source is not None
        else (
            "pub type Result<T> = std::result::Result<T, Error>;\n"
            "pub struct CommentTimestamp(u64);\n"
            "pub struct CommentAuthor;\n"
            "pub struct CommentReply;\n"
            "pub struct Comment;\n"
            "pub struct CellComment;\n"
            "pub struct TableRead;\n"
            "pub struct Builder;\n"
        ),
        encoding="utf-8",
    )
    module = root / boundaries.NEUTRAL_TABLE_MODULE_SOURCE
    module.parent.mkdir(parents=True, exist_ok=True)
    module.write_text(
        common_module_source if common_module_source is not None else "pub mod read;\n",
        encoding="utf-8",
    )

    cells = root / boundaries.IWA_NUMBERS_WIRE_TABLE_CELLS_SOURCE
    cells.parent.mkdir(parents=True, exist_ok=True)
    cells.write_text(
        cells_source
        if cells_source is not None
        else (
            "use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;\n"
            "use crate::cell_value::decode_cell_value;\n"
            "pub struct TableDimensions;\n"
            "pub struct TileReference;\n"
            "pub struct CellSource<'source> { bytes: &'source [u8] }\n"
            "pub enum TableCellIssue { CellStorageOutOfBounds, CellValueDecode }\n"
            "pub struct TableCellReadReport;\n"
            "pub trait TableCellReadBudget { type Error; }\n"
            "pub trait CellValueSink<Budget: TableCellReadBudget> {}\n"
            "pub fn read_table_cells() {\n"
            "    let _ = decode_cell_value;\n"
            "    let _ = storage::decode_tile_with_visitor;\n"
            "}\n"
        ),
        encoding="utf-8",
    )

    sidecars = root / boundaries.IWA_NUMBERS_WIRE_TABLE_SIDECARS_SOURCE
    sidecars.parent.mkdir(parents=True, exist_ok=True)
    sidecars.write_text(
        sidecars_source
        if sidecars_source is not None
        else (
            "use litchi_iwa_common::table::read::TableRead;\n"
            "use litchi_iwa_protos::comment_storage_codec;\n"
            "use crate::table_data_list::read_list_with_decoder;\n"
            "pub enum SidecarKind { Strings, Comments }\n"
            "pub struct SidecarReference;\n"
            "pub enum SidecarValue { Text }\n"
            "pub struct SidecarList;\n"
            "pub struct SidecarTables;\n"
            "pub trait SidecarReadBudget { type Error; }\n"
            "pub trait CellSidecarResolver { type Error; }\n"
            "pub fn read_sidecar_list() { let _ = read_list_with_decoder; }\n"
            "pub fn read_comment_storage() { let _ = comment_storage_codec; }\n"
            "impl SidecarTables { pub fn materialize_cell() { let _ = TableRead::builder; } }\n"
        ),
        encoding="utf-8",
    )
    wire_lib = root / boundaries.IWA_NUMBERS_WIRE_TABLE_CELLS_LIB_SOURCE
    wire_lib.parent.mkdir(parents=True, exist_ok=True)
    wire_lib.write_text(
        wire_lib_source
        if wire_lib_source is not None
        else "pub mod table_cells;\npub mod table_sidecars;\n",
        encoding="utf-8",
    )

    def focused_source(method: str, selector: str) -> str:
        return (
            "use litchi_iwa_common::table::read::TableRead;\n"
            "use litchi_numbers_wire::table_cells::{read_table_cells, TableCellReadBudget};\n"
            "use litchi_numbers_wire::table_sidecars::{read_sidecar_list, SidecarReadBudget};\n"
            "struct Budget;\n"
            "impl TableCellReadBudget for Budget { type Error = (); }\n"
            "impl SidecarReadBudget for Budget { type Error = (); }\n"
            "struct Package;\n"
            f"impl Package {{\n    pub fn {method}<S>(&self, {selector}: S) -> Result<TableRead, ()> {{\n"
            "        let mut budget = Budget;\n"
            "        read_sidecar_list();\n"
            "        read_table_cells();\n"
            "        let _ = (&mut budget, " + selector + ");\n"
            "        Ok(TableRead::builder())\n"
            "    }\n}\n"
        )

    if include_pages:
        pages = root / boundaries.PAGES_TABLE_CELLS_SOURCE
        pages.parent.mkdir(parents=True, exist_ok=True)
        pages.write_text(
            pages_source
            if pages_source is not None
            else focused_source("body_table_cells", "selector"),
            encoding="utf-8",
        )
        package = root / boundaries.PAGES_TABLE_CELLS_PACKAGE_SOURCE
        package.parent.mkdir(parents=True, exist_ok=True)
        package.write_text(
            pages_package_source
            if pages_package_source is not None
            else "pub(crate) mod table_cells;\n",
            encoding="utf-8",
        )

    if include_keynote:
        keynote = root / boundaries.KEYNOTE_TABLE_CELLS_SOURCE
        keynote.parent.mkdir(parents=True, exist_ok=True)
        keynote.write_text(
            keynote_source
            if keynote_source is not None
            else focused_source("slide_table_cells", "slide"),
            encoding="utf-8",
        )
        package = root / boundaries.KEYNOTE_TABLE_CELLS_PACKAGE_SOURCE
        package.parent.mkdir(parents=True, exist_ok=True)
        package.write_text(
            keynote_package_source
            if keynote_package_source is not None
            else "pub(crate) mod slide_table_cells;\n",
            encoding="utf-8",
        )


def add_iwa_numbers_wire_formula_names_scaffold(
    root: Path,
    *,
    source: str | None = None,
    lib_source: str | None = None,
) -> None:
    """Install a minimal bounded formula-name wire owner."""

    owner = root / boundaries.IWA_NUMBERS_WIRE_FORMULA_NAMES_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        source
        if source is not None
        else (
            "use litchi_iwa_common::wire::{WireDescent, preflight_wire_tree};\n"
            "use litchi_iwa_protos::group_node_category_codec;\n"
            "pub struct ReadLimits;\n"
            "pub struct ReadReport;\n"
            "pub struct FormulaOwnerDependencies;\n"
            "pub struct FormulaOwnerReadError;\n"
            "pub struct CategoryReadLimits;\n"
            "pub enum CategoryLabel<'source> { String(&'source str) }\n"
            "pub struct CategoryName<'source> { label: CategoryLabel<'source> }\n"
            "pub struct CategoryRead<'source> { entries: Vec<CategoryName<'source>> }\n"
            "pub struct CategoryReadError;\n"
            "pub fn read_formula_owner_dependencies(source: &[u8], limits: ReadLimits)\n"
            "    -> Result<FormulaOwnerDependencies, FormulaOwnerReadError> {\n"
            "    let _ = (source, limits, preflight_wire_tree, WireDescent);\n"
            "    todo!()\n"
            "}\n"
            "pub fn read_formula_category_names<'source>(source: &'source [u8], limits: CategoryReadLimits)\n"
            "    -> Result<CategoryRead<'source>, CategoryReadError> {\n"
            "    let _ = (source, limits, group_node_category_codec);\n"
            "    todo!()\n"
            "}\n"
        ),
        encoding="utf-8",
    )
    lib = root / boundaries.IWA_NUMBERS_WIRE_FORMULA_NAMES_LIB_SOURCE
    lib.parent.mkdir(parents=True, exist_ok=True)
    lib.write_text(
        lib_source if lib_source is not None else "pub mod formula_names;\n",
        encoding="utf-8",
    )


def add_keynote_family_module(
    root: Path, family: str, extra: tuple[str, ...] = ()
) -> None:
    """Publish one scaffolded Keynote family from its focused module."""

    path = root / getattr(boundaries, f"KEYNOTE_{family}_FAMILY_SOURCE")
    module = getattr(boundaries, f"KEYNOTE_{family}_FAMILY_MODULE")
    names = ", ".join(
        sorted(set(getattr(boundaries, f"KEYNOTE_{family}_CANONICAL_TYPES")) | set(extra))
    )
    body = f"pub use crate::package::family::{{{names}}};\n"
    path.parent.mkdir(parents=True, exist_ok=True)
    existing = path.read_text(encoding="utf-8") if path.is_file() else ""
    published = f"pub mod {module} {{\n    {body}}}\n" if module else body
    path.write_text(existing + published, encoding="utf-8")


def add_keynote_slide_transition_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES[0]
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES[0]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod transition;\n", encoding="utf-8")


def add_keynote_slide_delete_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_DELETE_CANONICAL_TYPES
        )
        + "impl Edit { pub fn remove_slide(&mut self, selector: SlideSelector) "
        "-> Result<&mut Self, Error> { todo!() } }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_DELETE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        "pub fn edit_slide_deletion(&self) -> slide::delete::Edit<'_> { todo!() }\n"
        "pub fn apply_slide_deletion(&self, patch: &slide::delete::Patch) "
        "-> Result<slide::delete::Commit, slide::delete::Error> { todo!() }\n"
        "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, slide_export = (
        root / path for path in boundaries.KEYNOTE_SLIDE_DELETE_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod slide;\n", encoding="utf-8")
    package_export.write_text("mod slide_delete;\n", encoding="utf-8")
    slide_export.write_text("pub mod delete;\n", encoding="utf-8")


def add_numbers_table_cell_comment_reply_scaffold(
    root: Path,
    *,
    owner_source: str | None = None,
    lib_source: str | None = None,
    codec_source: str | None = None,
) -> None:
    owner = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        owner_source
        or (
            "pub struct CommentReply { text: String }\n"
            "impl fmt::Debug for CommentReply {\n"
            "    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {\n"
            "        f.debug_struct(\"CommentReply\")\n"
            "            .field(\"text\", &\"<redacted>\")\n"
            "            .finish()\n"
            "    }\n"
            "}\n"
            "struct CommentOwnershipCensus {\n"
            "    comment_list_ids: Vec<u64>, list_entries: Vec<u64>,\n"
            "    table_references: Vec<u64>, cell_comment_keys: Vec<u64>,\n"
            "    storages: Vec<u64>, reply_ids: Vec<u64>,\n"
            "    entry_storage_ids: Vec<u64>, author_ids: Vec<u64>,\n"
            "    storage_uuid: Vec<u64>, uuids: Vec<u64>,\n"
            "}\n"
            "fn census_comment_ownership() { census_models_and_cells(); census_alias_checks(); }\n"
            "fn census_models_and_cells() {}\n"
            "fn census_alias_checks() {}\n"
            "fn strict_comment_reply_codec() {\n"
            "    comment_storage_codec::decode_comment_storage_archive_with_visitor();\n"
            "    visit_reply(); reply_references();\n"
            "    let _ = DecodeOptions; let _ = DecodeLimit;\n"
            "    let _ = CommentStorageVisitor; let _ = CommentStorageSnapshot;\n"
            "}\n"
            "impl Package {\n"
            "    pub fn table_cell_comment_replies<'sheet, 'table>(\n"
            "        &self, sheet: impl Into<SheetSelector<'sheet>>,\n"
            "        table: impl Into<TableSelector<'table>>, position: At,\n"
            "    ) -> Result<Box<[CommentReply]>, Error> { todo!() }\n"
            "    pub fn table_cell_comment_replies_a1<'sheet, 'table>(\n"
            "        &self, sheet: impl Into<SheetSelector<'sheet>>,\n"
            "        table: impl Into<TableSelector<'table>>, address: &str,\n"
            "    ) -> Result<Box<[CommentReply]>, Error> {\n"
            "        self.table_cell_comment_replies(sheet, table, At::from_a1(address))\n"
            "    }\n"
            "}\n"
        ),
        encoding="utf-8",
    )

    codec = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        codec_source
        or (
            "pub struct DecodeOptions;\n"
            "pub enum DecodeLimit { Bytes }\n"
            "pub trait CommentStorageVisitor {}\n"
            "pub struct CommentStorageSnapshot<'source> {\n"
            "    source_order: &'source str,\n"
            "}\n"
            "pub fn decode_comment_storage_archive_with_visitor() {}\n"
            "fn visit_reply() {}\n"
            "fn reply_references() {}\n"
        ),
        encoding="utf-8",
    )

    lib = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]
    lib.parent.mkdir(parents=True, exist_ok=True)
    lib.write_text(
        lib_source
        or "pub mod cell;\n",
        encoding="utf-8",
    )
    # ADR 0034: the value is published by `cell::comment`.
    comment = root / boundaries.NUMBERS_SOURCE_ROOT / "cell" / "comment.rs"
    comment.parent.mkdir(parents=True, exist_ok=True)
    if not comment.exists():
        comment.write_text(
            "pub use crate::package::comments::CommentReply;\n",
            encoding="utf-8",
        )


def add_numbers_table_cell_comment_reply_metadata_scaffold(root: Path) -> None:
    """Extend the read facade scaffold with its typed metadata contract."""

    add_numbers_table_cell_comment_reply_scaffold(root)
    owner = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE
    source = owner.read_text(encoding="utf-8")
    source = source.replace(
        "pub struct CommentReply { text: String }",
        "pub struct CommentReply {\n"
        "    text: String,\n"
        "    timestamp: Option<CommentTimestamp>,\n"
        "    author: Option<CommentAuthor>,\n"
        "}",
        1,
    )
    source += (
        "pub struct Comment {\n"
        "    text: String,\n"
        "    timestamp: Option<CommentTimestamp>,\n"
        "    author: Option<CommentAuthor>,\n"
        "}\n"
        "pub struct CommentTimestamp(f64);\n"
        "impl CommentTimestamp {\n"
        "    pub const fn new(seconds: f64) -> Option<Self> {\n"
        "        seconds.is_finite().then_some(Self(seconds))\n"
        "    }\n"
        "    pub const fn as_f64(self) -> f64 { self.0 }\n"
        "}\n"
        "pub struct CommentAuthor {\n"
        "    display_name: Option<Box<str>>,\n"
        "    public_id: Option<Box<str>>,\n"
        "}\n"
        "impl CommentAuthor {\n"
        "    pub fn display_name(&self) -> Option<&str> { todo!() }\n"
        "    pub fn public_id(&self) -> Option<&str> { todo!() }\n"
        "}\n"
        "impl fmt::Debug for CommentAuthor {\n"
        "    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {\n"
        "        f.debug_struct(\"CommentAuthor\").finish_non_exhaustive()\n"
        "    }\n"
        "}\n"
        "impl Comment {\n"
        "    pub fn timestamp(&self) -> Option<CommentTimestamp> { todo!() }\n"
        "    pub fn author(&self) -> Option<&CommentAuthor> { todo!() }\n"
        "}\n"
        "impl CommentReply {\n"
        "    pub fn timestamp(&self) -> Option<CommentTimestamp> { todo!() }\n"
        "    pub fn author(&self) -> Option<&CommentAuthor> { todo!() }\n"
        "}\n"
    )
    owner.write_text(source, encoding="utf-8")

    lib = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]
    lib.write_text(
        lib.read_text(encoding="utf-8")
        + "pub use package::{ comments::{"
        "CommentAuthor as TableCellCommentAuthor, "
        "CommentTimestamp as TableCellCommentTimestamp} };\n",
        encoding="utf-8",
    )


def add_numbers_table_cell_comment_reply_codec_rewrite_scaffold(root: Path) -> None:
    codec = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        """/// Strict canonical known-field/reference validation rejects duplicate,
/// wrong-wire, and noncanonical values.
/// Unknown/raw fields, source_order, raw framing, and balanced unknown groups
/// remain preserved.
/// Resource/preflight accounting covers max_input, max_output, max_fields,
/// max_work, max_depth, max_references, max_text, try_reserve, scratch,
/// retained bytes, and allocations.
pub struct CommentStorageReplyRewrite;
pub struct PreparedCommentStorageReplyRewrite;
pub struct RewriteExecutionRequirements;
pub struct RewriteExecutionLimits;
pub fn prepare_comment_storage_reply_rewrite() -> PreparedCommentStorageReplyRewrite {
    todo!()
}
impl PreparedCommentStorageReplyRewrite {
    pub fn prepare_report(&self) {}
    pub fn execution_requirements(&self) -> RewriteExecutionRequirements { todo!() }
    pub fn execute(&self, _limits: RewriteExecutionLimits) {}
}
#[cfg(test)]
mod tests {
    #[test]
    fn comment_storage_reply_rewrite_round_trip_preserves_unknown_raw_and_limits() {
        let prepared = prepare_comment_storage_reply_rewrite();
        prepared.prepare_report();
        prepared.execution_requirements();
        prepared.execute(RewriteExecutionLimits);
    }
}
""",
        encoding="utf-8",
    )


def add_numbers_table_cell_comment_reply_mutation_scaffold(root: Path) -> None:
    """Create the frozen reply mutation contract in a temp tree."""

    semantic = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction {\n"
        + "\n".join(
            f"pub struct {name};"
            for name in boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TYPES[1:]
        )
        + "\n}\n"
        "pub struct CommentReply { pub text: String }\n"
        "impl fmt::Debug for CommentReply {\n"
        "    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {\n"
        "        f.debug_struct(\"CommentReply\").field(\"text\", &\"<redacted>\").finish()\n"
        "    }\n"
        "}\n"
        "pub struct CommentReplyIndex(u32);\n"
        "impl CommentReplyIndex {\n"
        "    pub fn new(value: u32) -> Self { Self(value) }\n"
        "    pub fn index(self) -> u32 { self.0 }\n"
        "    pub fn get(self) -> u32 { self.0 }\n"
        "    pub fn try_from_usize(value: usize) -> Result<Self, CommentReplyError> { Ok(Self(value as u32)) }\n"
        "}\n",
        encoding="utf-8",
    )
    cell_mod = root / boundaries.NUMBERS_SOURCE_ROOT / "cell" / "mod.rs"
    cell_mod.parent.mkdir(parents=True, exist_ok=True)
    cell_mod.write_text("pub mod comment;\n", encoding="utf-8")

    owner = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "use crate::cell::comment::{CommentReply, CommentReplyIndex};\n"
        "use crate::cell::comment::transaction::{CommentReplyCommit, CommentReplyEdit, CommentReplyPatch};\n"
        "struct CommentOwnershipCensus;\n"
        "fn resolve_comment_reply_rooted(_: SheetSelector, _: TableSelector, _: At, _: CommentReplyIndex) {}\n"
        "fn comment_graph() { CommentList; comment_list_ids; cell_comment_keys; list_entries; entry_storage_ids; table_references; comment_storage; }\n"
        "fn census_comment_reply_ownership() { CommentOwnershipCensus; global; all_components; ArchiveReference; storages; reply_ids; storage_uuid; author_ids; storage_occurrences; refcount; reference_count; }\n"
        "fn copy_on_write_shared_clone_new_storage_cull() { copy_on_write; shared; clone; new_storage; cull; }\n"
        "fn reject_duplicate_alias_cycle_depth_dependency() { duplicate; alias; cycle; depth; invalid; reject; dependency; }\n"
        "fn archive_info_authority() { ArchiveInfo; MessageInfo; FieldInfo; object_references; aggregate; field_infos; preserve; }\n"
        "fn metadata_ownership() { metadata; uuid; object_uuid; save_token; watermark; versioned; ambiguous; data_owner; root_map; external; }\n"
        "fn atomic_source_bytes_inverse_patch_candidate_reopen_locality_same_content_previews() { atomic; source_bytes; inverse; Patch; candidate; reopen; locality; same_content; previews; publication; }\n"
        "fn transaction_budget_preflight_charge() { TransactionBudget; WireBudget; budget; residual; preflight; charge; max_input; max_output; fields; work; depth; references; allocations; scratch; retained; reassembly; Snappy; ZIP; try_reserve; }\n"
        "fn strict_comment_storage_route() { CommentStorageReplyRewrite; prepare_comment_storage_reply_rewrite(); CommentStorage; visitor; execution_requirements; execute; }\n"
        "impl Package {\n"
        "pub fn table_cell_comment_reply(&self, sheet: SheetSelector, table: TableSelector, position: At, index: CommentReplyIndex) -> Result<CommentReply, CommentReplyError> { todo!() }\n"
        "pub fn table_cell_comment_reply_a1(&self, sheet: SheetSelector, table: TableSelector, address: &str, index: CommentReplyIndex) -> Result<CommentReply, CommentReplyError> { self.table_cell_comment_reply(sheet, table, At::from_a1(address), index) }\n"
        "pub fn edit_table_cell_comment_replies(&self, sheet: SheetSelector, table: TableSelector, position: At) -> Result<CommentReplyEdit, CommentReplyError> { todo!() }\n"
        "pub fn edit_table_cell_comment_replies_a1(&self, sheet: SheetSelector, table: TableSelector, address: &str) -> Result<CommentReplyEdit, CommentReplyError> { self.edit_table_cell_comment_replies(sheet, table, At::from_a1(address)) }\n"
        "pub fn set_table_cell_comment_reply(&self, sheet: SheetSelector, table: TableSelector, position: At, index: CommentReplyIndex, value: CommentReply) -> Result<CommentReplyCommit, CommentReplyError> { todo!() }\n"
        "pub fn set_table_cell_comment_reply_a1(&self, sheet: SheetSelector, table: TableSelector, address: &str, index: CommentReplyIndex, value: CommentReply) -> Result<CommentReplyCommit, CommentReplyError> { self.set_table_cell_comment_reply(sheet, table, At::from_a1(address), index, value) }\n"
        "pub fn remove_table_cell_comment_reply(&self, sheet: SheetSelector, table: TableSelector, position: At, index: CommentReplyIndex) -> Result<CommentReplyCommit, CommentReplyError> { todo!() }\n"
        "pub fn remove_table_cell_comment_reply_a1(&self, sheet: SheetSelector, table: TableSelector, address: &str, index: CommentReplyIndex) -> Result<CommentReplyCommit, CommentReplyError> { self.remove_table_cell_comment_reply(sheet, table, At::from_a1(address), index) }\n"
        "pub fn add_table_cell_comment_reply(&self, sheet: SheetSelector, table: TableSelector, position: At, value: CommentReply) -> Result<CommentReplyCommit, CommentReplyError> { todo!() }\n"
        "pub fn add_table_cell_comment_reply_a1(&self, sheet: SheetSelector, table: TableSelector, address: &str, value: CommentReply) -> Result<CommentReplyCommit, CommentReplyError> { self.add_table_cell_comment_reply(sheet, table, At::from_a1(address), value) }\n"
        "pub fn apply_table_cell_comment_reply(&self, patch: &CommentReplyPatch) -> Result<CommentReplyCommit, CommentReplyError> { todo!() }\n"
        "}\n"
        "impl CommentReplyEdit {\n"
        "pub fn set(self, value: CommentReply) -> Result<Self, CommentReplyError> { let _ = value; Ok(self) }\n"
        "pub fn remove(self, index: CommentReplyIndex) -> Result<Self, CommentReplyError> { let _ = index; Ok(self) }\n"
        "pub fn commit(self) -> Result<CommentReplyCommit, CommentReplyError> { todo!() }\n"
        "}\n",
        encoding="utf-8",
    )
    package = root / boundaries.NUMBERS_SOURCE_ROOT / "package.rs"
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text("pub(crate) mod comments;\n", encoding="utf-8")

    lib = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]
    lib.parent.mkdir(parents=True, exist_ok=True)
    lib.write_text(
        "pub mod cell;\n"
        "pub use cell::comment::{CommentReply as TableCellCommentReply, CommentReplyIndex as TableCellCommentReplyIndex};\n"
        "pub use cell::comment::transaction::{CommentReplyEdit as TableCellCommentReplyEdit, CommentReplyPatch as TableCellCommentReplyPatch, CommentReplyCommit as TableCellCommentReplyCommit, CommentReplyDiagnostics as TableCellCommentReplyDiagnostics, CommentReplyError as TableCellCommentReplyError, CommentReplyLimitKind as TableCellCommentReplyLimitKind, CommentReplyPath as TableCellCommentReplyPath};\n",
        encoding="utf-8",
    )
    codec = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "pub struct CommentStorageReplyRewrite;\n"
        "pub struct PreparedCommentStorageReplyRewrite;\n"
        "pub struct RewriteExecutionRequirements;\n"
        "pub struct RewriteExecutionLimits;\n"
        "pub fn prepare_comment_storage_reply_rewrite() {}\n"
        "impl PreparedCommentStorageReplyRewrite { pub fn prepare_report(&self) {} pub fn execution_requirements(&self) -> RewriteExecutionRequirements { todo!() } pub fn execute(&self, _: RewriteExecutionLimits) {} }\n"
        "// strict canonical duplicate wrong-wire validation; unknown raw balanced groups preserve source order; max_input max_output max_fields max_work max_depth max_references try_reserve scratch retained allocations\n",
        encoding="utf-8",
    )
    codec_fuzz = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_FUZZ_SOURCE
    codec_fuzz.parent.mkdir(parents=True, exist_ok=True)
    codec_fuzz.write_text(
        "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
        "fuzz_target!(|data: &[u8]| { let _ = data; prepare_comment_storage_reply_rewrite(); execute; unknown; inverse; });\n",
        encoding="utf-8",
    )
    package_fuzz = root / boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_FUZZ_SOURCE
    package_fuzz.parent.mkdir(parents=True, exist_ok=True)
    package_fuzz.write_text(
        "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
        "fuzz_target!(|data: &[u8]| { let _ = data; add_table_cell_comment_reply; set_table_cell_comment_reply; remove_table_cell_comment_reply; CommentReplyIndex; inverse; conflict; atomic; limit; });\n",
        encoding="utf-8",
    )
    for corpus in (
        boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_FUZZ_CORPUS,
        boundaries.NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_FUZZ_CORPUS,
    ):
        directory = root / corpus
        directory.mkdir(parents=True, exist_ok=True)
        (directory / "seed.hex").write_text("00\n", encoding="utf-8")


def add_numbers_table_header_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source + "pub mod transaction;\n",
        encoding="utf-8",
    )
    transaction = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE
    transaction.parent.mkdir(parents=True, exist_ok=True)
    transaction_source = (
        transaction.read_text(encoding="utf-8") if transaction.is_file() else ""
    )
    transaction.write_text(
        transaction_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    table_export = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[2]
    table_export.parent.mkdir(parents=True, exist_ok=True)
    table_source = (
        table_export.read_text(encoding="utf-8") if table_export.is_file() else ""
    )
    table_export.write_text(table_source + "pub mod headers;\n", encoding="utf-8")
    lib_export = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[0]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod table;\n", encoding="utf-8")


def add_keynote_placeholder_visibility_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "pub enum Kind { Title, Body, SlideNumber }\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES
            if name != "Kind"
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    if not owner.is_file():
        owner.write_text("// Private package owner.\n", encoding="utf-8")
    lib_export = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[0]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod slide;\n", encoding="utf-8")
    package_export = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[1]
    package_source = (
        package_export.read_text(encoding="utf-8")
        if package_export.is_file()
        else ""
    )
    package_export.write_text(
        package_source + "mod slide_placeholder_visibility;\n", encoding="utf-8"
    )
    slide_export = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[2]
    slide_source = (
        slide_export.read_text(encoding="utf-8") if slide_export.is_file() else ""
    )
    slide_export.write_text(
        slide_source + "pub mod placeholder;\n", encoding="utf-8"
    )


def add_keynote_soundtrack_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    if not owner.is_file():
        owner.write_text("// Private package owner.\n", encoding="utf-8")
    lib_export = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES[0]
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod soundtrack;\n", encoding="utf-8")
    package_export = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES[1]
    package_source = (
        package_export.read_text(encoding="utf-8")
        if package_export.is_file()
        else ""
    )
    package_export.write_text(
        package_source + "mod soundtrack_settings;\n", encoding="utf-8"
    )


def add_keynote_soundtrack_items_canonical_scaffold(root: Path) -> None:
    """Install a minimal focused soundtrack-item owner for checker tests."""

    semantic = root / boundaries.KEYNOTE_SOUNDTRACK_ITEMS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SOUNDTRACK_ITEMS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )

    owner = root / boundaries.KEYNOTE_SOUNDTRACK_ITEMS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        "    pub fn soundtrack_items(&self) {}\n"
        "    pub fn edit_soundtrack_items(&self) {}\n"
        "    pub fn apply_soundtrack_items(&self) {}\n"
        "}\n",
        encoding="utf-8",
    )

    lib = root / boundaries.KEYNOTE_SOUNDTRACK_ITEMS_EXPORT_SOURCES[0]
    lib.parent.mkdir(parents=True, exist_ok=True)
    lib.write_text("pub mod soundtrack;\n", encoding="utf-8")
    package = root / boundaries.KEYNOTE_SOUNDTRACK_ITEMS_EXPORT_SOURCES[1]
    package.write_text("mod soundtrack_items;\n", encoding="utf-8")
    soundtrack = root / boundaries.KEYNOTE_SOUNDTRACK_ITEMS_EXPORT_SOURCES[2]
    soundtrack.write_text("pub mod items;\n", encoding="utf-8")


def add_numbers_sheet_order_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_SHEET_ORDER_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_SHEET_ORDER_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    if not owner.is_file():
        owner.write_text("// Private package owner.\n", encoding="utf-8")
    lib_export = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[0]
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod sheet;\n", encoding="utf-8")
    package_export = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[1]
    package_source = (
        package_export.read_text(encoding="utf-8")
        if package_export.is_file()
        else ""
    )
    package_export.write_text(
        package_source + "mod sheet_order;\n", encoding="utf-8"
    )
    sheet_export = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[2]
    sheet_source = (
        sheet_export.read_text(encoding="utf-8") if sheet_export.is_file() else ""
    )
    sheet_export.write_text(sheet_source + "pub mod order;\n", encoding="utf-8")


def add_numbers_table_sort_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_SORT_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub use litchi_iwa_common::table::sort::{"
        + ", ".join(boundaries.NUMBERS_TABLE_SORT_SEMANTIC_TYPES)
        + "};\n"
        "pub mod transaction;\n",
        encoding="utf-8",
    )
    transaction = root / boundaries.NUMBERS_TABLE_SORT_TRANSACTION_SOURCE
    transaction.parent.mkdir(parents=True, exist_ok=True)
    transaction.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_SORT_TRANSACTION_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_SORT_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_SORT_TRANSACTION_TYPES
        )
        + "fn table_model_field44_field45() {}\n"
        + "fn transaction_budget_preflight() {}\n"
        + "fn verify_locality_candidate_reopen_atomic_inverse() {}\n"
        + "impl Package {\n"
        + "pub fn table_sort_order(&self, sheet: SheetSelector, table: TableSelector) "
        "-> Result<Order, Error> {}\n"
        + "pub fn edit_table_sort_order(&self, sheet: SheetSelector, table: TableSelector) "
        "-> Result<Edit, Error> {}\n"
        + "pub fn apply_table_sort_order(&self, patch: &Patch) "
        "-> Result<Commit, Error> {}\n"
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, table_export = (
        root / path for path in boundaries.NUMBERS_TABLE_SORT_EXPORT_SOURCES
    )
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text("pub mod table;\n", encoding="utf-8")
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "pub(crate) mod table_sort;\n", encoding="utf-8"
    )
    table_export.parent.mkdir(parents=True, exist_ok=True)
    table_export.write_text("pub mod sort;\n", encoding="utf-8")

    codec = root / boundaries.NUMBERS_TABLE_SORT_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "pub struct SortOrderSnapshot;\n"
        "pub struct PreparedTableSortOrderRewrite;\n"
        "pub struct RewriteExecutionRequirements;\n"
        "pub struct RewriteExecutionLimits;\n"
        "pub fn decode_table_sort_order() {}\n"
        "pub fn decode_table_sort_order_with_report() {}\n"
        "pub fn prepare_table_sort_order_rewrite() {}\n"
        "pub fn canonical_table_sort_order() {}\n"
        "pub fn rewrite_table_sort_order() {}\n"
        "const SORT_ORDER_FIELD: u32 = 44;\n"
        "const SORT_RULES_FIELD: u32 = 2;\n"
        "// unknown overlong fields are retained\n"
        "fn report_field44_field45() {}\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.NUMBERS_TABLE_SORT_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.NUMBERS_TABLE_SORT_CODEC_MODULE};\n"
        "\n#[doc(hidden)]\n"
        f"pub mod {boundaries.NUMBERS_TABLE_SORT_CODEC_NEUTRAL_MODULE} {{\n"
        f"    pub use super::{boundaries.NUMBERS_TABLE_SORT_CODEC_MODULE}::*;\n"
        "}\n",
        encoding="utf-8",
    )
    for fuzz_target in (
        boundaries.NUMBERS_TABLE_SORT_CODEC_FUZZ_SOURCE,
        boundaries.NUMBERS_TABLE_SORT_FUZZ_SOURCE,
    ):
        absolute = root / fuzz_target
        absolute.parent.mkdir(parents=True, exist_ok=True)
        absolute.write_text(
            "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
            "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
            encoding="utf-8",
        )
    for corpus in (
        boundaries.NUMBERS_TABLE_SORT_CODEC_FUZZ_CORPUS,
        boundaries.NUMBERS_TABLE_SORT_FUZZ_CORPUS,
    ):
        (root / corpus).mkdir(parents=True, exist_ok=True)


def add_numbers_table_title_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text("impl Package {}\n", encoding="utf-8")
    lib_export, package_export, table_export = (
        root / path for path in boundaries.NUMBERS_TABLE_TITLE_SETTINGS_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod table;\n", encoding="utf-8")
    package_export.write_text("pub(crate) mod table_title;\n", encoding="utf-8")
    table_export.write_text("pub mod title;\n", encoding="utf-8")


def add_numbers_table_appearance_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_APPEARANCE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_APPEARANCE_SEMANTIC_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_APPEARANCE_TRANSACTION_TYPES
        )
        + "fn copy_on_write_style_variation() {}\n"
        + "fn metadata_save_token_watermark() {}\n"
        + "fn transaction_budget_preflight() {}\n"
        + "fn verify_locality_candidate_reopen() {}\n"
        + "impl Package {\n"
        + "pub fn table_appearance<'sheet, 'table>(&self, "
        "sheet: impl Into<SheetSelector<'sheet>>, "
        "table: impl Into<TableSelector<'table>>) -> Result<Appearance, Error> {}\n"
        + "pub fn edit_table_appearance<'sheet, 'table>(&self, "
        "sheet: impl Into<SheetSelector<'sheet>>, "
        "table: impl Into<TableSelector<'table>>, value: Appearance) -> Result<Edit, Error> {}\n"
        + "pub fn apply_table_appearance(&self, patch: &Patch) -> Result<Commit, Error> {}\n"
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, table_export = (
        root / path for path in boundaries.NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES
    )
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text("pub mod table;\n", encoding="utf-8")
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "pub(crate) mod table_appearance;\n", encoding="utf-8"
    )
    table_export.parent.mkdir(parents=True, exist_ok=True)
    table_export.write_text("pub mod appearance;\n", encoding="utf-8")

    codec = root / boundaries.NUMBERS_TABLE_APPEARANCE_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "pub fn decode_table_model_with_report() {}\n"
        "pub fn prepare_table_model_style_rewrite() {}\n"
        "pub fn canonical_table_style_variation() {}\n"
        "pub fn append_stylesheet_style() {}\n"
        "fn canonical_unknown_report() {}\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.NUMBERS_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.NUMBERS_TABLE_APPEARANCE_CODEC_MODULE};\n",
        encoding="utf-8",
    )


def add_numbers_table_cell_pop_up_menu_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction;\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_TYPES
        )
        + "pub enum InitialSelection { FirstItem, Blank }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_TYPES
        )
        + "fn control_cell_data_list_refcount() {}\n"
        + "fn copy_on_write_shared_model_cull() {}\n"
        + "fn metadata_save_token_object_uuid_external_reference_watermark() {}\n"
        + "fn transaction_budget_preflight_charge() {}\n"
        + "fn verify_locality_candidate_reopen_string_table() {}\n"
        + "impl Package {\n"
        + "pub fn table_cell_pop_up_menu_format<'sheet, 'table, 'cell>(&self, "
        "sheet: SheetSelector<'sheet>, table: TableSelector<'table>, "
        "position: At<'cell>) -> Result<PopUpMenu, Error> {}\n"
        + "pub fn edit_table_cell_pop_up_menu_format<'sheet, 'table, 'cell>(&self, "
        "sheet: SheetSelector<'sheet>, table: TableSelector<'table>, "
        "position: At<'cell>, value: PopUpMenu) -> Result<Edit, Error> {}\n"
        + "pub fn apply_table_cell_pop_up_menu_format(&self, patch: &Patch) "
        "-> Result<Commit, Error> {}\n"
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, data_format_export = (
        root / path
        for path in boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[:3]
    )
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text("pub mod cell;\n", encoding="utf-8")
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "pub(crate) mod table_cell_pop_up_menu;\n", encoding="utf-8"
    )
    data_format_export.parent.mkdir(parents=True, exist_ok=True)
    data_format_export.write_text("pub mod pop_up_menu;\n", encoding="utf-8")

    codec = root / boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "pub struct PreparedPopUpMenuModelWrite;\n"
        "pub struct PreparedCellSpecWrite;\n"
        "pub struct RewriteExecutionRequirements;\n"
        "pub struct RewriteExecutionLimits;\n"
        "pub fn decode_popup_menu_model_with_report() {}\n"
        "pub fn decode_cell_spec_with_report() {}\n"
        "pub fn prepare_popup_menu_model_write() {}\n"
        "pub fn prepare_cell_spec_write() {}\n"
        "pub fn canonical_popup_menu_model() {}\n"
        "pub fn canonical_cell_spec() {}\n"
        "fn canonical_unknown_report_nil_deprecated_refcount() {}\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_MODULE};\n",
        encoding="utf-8",
    )
    for fuzz_target in (
        boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_FUZZ_SOURCE,
        boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_FUZZ_SOURCE,
    ):
        absolute = root / fuzz_target
        absolute.parent.mkdir(parents=True, exist_ok=True)
        absolute.write_text(
            "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
            "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
            encoding="utf-8",
        )
    for corpus in (
        boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_FUZZ_CORPUS,
        boundaries.NUMBERS_TABLE_CELL_POP_UP_MENU_FUZZ_CORPUS,
    ):
        (root / corpus).mkdir(parents=True, exist_ok=True)


def add_numbers_table_cell_control_canonical_scaffold(root: Path) -> None:
    """Create the smallest complete cell-control owner for boundary unit tests."""

    semantic = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction;\n"
        + "\n".join(
            f"pub struct {name};"
            for name in boundaries.NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_TYPES
        )
        + "\npub enum CellControl { Checkbox(Checkbox), StarRating(StarRating), "
        "Slider(Slider), Stepper(Stepper), PopUpMenu(PopUpMenu) }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "\n".join(
            f"pub struct {name};"
            for name in boundaries.NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_TYPES
        )
        + "\n"
        "fn control_cell_graph() {}\n"
        "fn table_data_list_graph() {}\n"
        "fn refcount_graph() {}\n"
        "fn copy_on_write_shared_cull() {}\n"
        "fn metadata_save_token_object_uuid_external_reference_watermark() {}\n"
        "fn transaction_budget_preflight_residual() {}\n"
        "fn resource_allocation_retained_max_minus_one() {}\n"
        "fn prepared_execution_requirements() {}\n"
        "fn verify_locality_candidate_reopen_atomic_inverse() {}\n"
        "fn prove_cross_component_reference() {}\n"
        "fn reject_cross_component_write() {}\n"
        "fn changed_route() { reject_cross_component_write(); let mut budget = (); }\n"
        "fn effective_locator() {}\n"
        "fn require_external_edge() {}\n"
        "fn object_identifier() {}\n"
        "fn has_physical_alias() {}\n"
        "impl Package {\n"
        "pub fn table_cell_control_format<'sheet, 'table, 'cell>(&self, "
        "sheet: SheetSelector<'sheet>, table: TableSelector<'table>, "
        "position: At<'cell>) -> Result<Option<CellControl>, Error> {}\n"
        "pub fn edit_table_cell_control_format<'sheet, 'table, 'cell>(&self, "
        "sheet: SheetSelector<'sheet>, table: TableSelector<'table>, "
        "position: At<'cell>) -> Result<Edit, Error> {}\n"
        "pub fn apply_table_cell_control_format(&self, patch: &Patch) "
        "-> Result<Commit, Error> {}\n"
        "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, data_format_export = (
        root / path
        for path in boundaries.NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[:3]
    )
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text("pub mod cell;\n", encoding="utf-8")
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "pub(crate) mod table_cell_control;\n", encoding="utf-8"
    )
    data_format_export.parent.mkdir(parents=True, exist_ok=True)
    data_format_export.write_text("pub mod control;\n", encoding="utf-8")

    codec = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "pub struct PreparedCellSpecWrite;\n"
        "pub struct RewriteExecutionRequirements;\n"
        "pub struct RewriteExecutionLimits;\n"
        "pub fn decode_cell_spec_with_report() {}\n"
        "pub fn prepare_cell_spec_write() {}\n"
        "pub fn canonical_cell_spec() {}\n"
        "pub fn rewrite_cell_spec() {}\n"
        "fn canonical_unknown_report_prepared_fixed64_interaction_refcount() {}\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.NUMBERS_TABLE_CELL_CONTROL_CODEC_MODULE};\n",
        encoding="utf-8",
    )
    for fuzz_target in (
        boundaries.NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_SOURCE,
        boundaries.NUMBERS_TABLE_CELL_CONTROL_FUZZ_SOURCE,
    ):
        absolute = root / fuzz_target
        absolute.parent.mkdir(parents=True, exist_ok=True)
        absolute.write_text(
            "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
            "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
            encoding="utf-8",
        )
    for corpus in (
        boundaries.NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_CORPUS,
        boundaries.NUMBERS_TABLE_CELL_CONTROL_FUZZ_CORPUS,
    ):
        (root / corpus).mkdir(parents=True, exist_ok=True)
    integration = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE
    integration.parent.mkdir(parents=True, exist_ok=True)
    integration.write_text(
        "#[test]\n"
        "fn split_components_read_noop_and_all_control_transitions_are_atomic() {}\n"
        "fn assert_split_read_noop_and_write_reject() { no_op; is_noop; exact_bytes; write_reject; changed_edit_rejects; }\n"
        "fn split_components_reject_bad_edges_aliases_and_opaque_inbound_refs_atomically() { alias; external; opaque; atomic; }\n",
        encoding="utf-8",
    )


def add_numbers_table_cell_custom_format_canonical_scaffold(root: Path) -> None:
    """Create the complete focused Custom owner and private codec seam."""

    semantic = root / boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction {\n"
        "    pub use crate::package::table_cell_custom_format::{"
        "Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path};\n"
        "}\n"
        "pub enum Custom { Number(Number), Text(Text), DateTime(DateTime) }\n"
        "pub struct Number;\n"
        "pub struct Text;\n"
        "pub struct DateTime;\n",
        encoding="utf-8",
    )

    owner = root / boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "pub enum Path { Package, Cell { sheet: usize, table: usize, position: At } }\n"
        "pub enum LimitKind { PayloadBytes }\n"
        "pub struct Edit;\n"
        "pub struct Patch;\n"
        "pub struct Commit;\n"
        "pub struct Diagnostics;\n"
        "pub struct Error;\n"
        "fn custom_owner_policy() {\n"
        "    document_scoped_registry; rooted; registry; field_9; message_type;\n"
        "    270; 271; 272; Number; Text; Date_Time;\n"
        "    format_list; refcount; reuse; cull; copy_on_write; reference;\n"
        "    exact_source; source_bound; PatchConflict; inverse; no_op;\n"
        "    candidate; reopen; readback; locality; same_content;\n"
        "    budget; preflight; bounded; allocation; unsupported; fail_closed;\n"
        "}\n"
        "impl Package {\n"
        "pub fn table_cell_custom_format<'sheet, 'table>(&self, "
        "sheet: impl Into<SheetSelector<'sheet>>, "
        "table: impl Into<TableSelector<'table>>, position: At) "
        "-> Result<Option<Custom>, Error> {}\n"
        "pub fn edit_table_cell_custom_format<'sheet, 'table>(&self, "
        "sheet: impl Into<SheetSelector<'sheet>>, "
        "table: impl Into<TableSelector<'table>>, position: At) "
        "-> Result<Edit, Error> {}\n"
        "pub fn apply_table_cell_custom_format(&self, patch: &Patch) "
        "-> Result<Commit, Error> {}\n"
        "}\n",
        encoding="utf-8",
    )

    lib = root / (boundaries.NUMBERS_SOURCE_ROOT / "lib.rs")
    lib.parent.mkdir(parents=True, exist_ok=True)
    lib.write_text("pub mod cell;\n", encoding="utf-8")
    cell = root / (boundaries.NUMBERS_SOURCE_ROOT / "cell" / "mod.rs")
    cell.parent.mkdir(parents=True, exist_ok=True)
    cell.write_text("pub mod data_format;\n", encoding="utf-8")
    data_format = root / (boundaries.NUMBERS_SOURCE_ROOT / "cell" / "data_format.rs")
    data_format.parent.mkdir(parents=True, exist_ok=True)
    data_format.write_text(
        "pub mod custom;\npub use custom::Custom;\n", encoding="utf-8"
    )
    package = root / boundaries.NUMBERS_PACKAGE_SOURCE
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text(
        "pub(crate) mod table_cell_custom_format;\n", encoding="utf-8"
    )

    codec = root / boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "use buffa::DecodeOptions;\n"
        "pub const CUSTOM_FORMAT_REGISTRY_MESSAGE_TYPE: u32 = 222;\n"
        "pub const CUSTOM_FORMAT_REGISTRY_REFERENCE_FIELD: u32 = 9;\n"
        "pub const CUSTOM_FORMAT_LIST_KIND: u32 = 6;\n"
        "pub const NATIVE_CUSTOM_NUMBER_FORMAT_TYPE: u32 = 270;\n"
        "pub const NATIVE_CUSTOM_TEXT_FORMAT_TYPE: u32 = 271;\n"
        "pub const NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE: u32 = 272;\n"
        "pub struct DecodeOptions;\n"
        "pub struct DecodeReport;\n"
        "pub struct DecodeError;\n"
        "pub struct PreparedCustomFormatRewrite;\n"
        "pub struct PreparedCustomFormatListRewrite;\n"
        "pub struct RewriteExecutionRequirements;\n"
        "pub struct RewriteExecutionLimits;\n"
        "pub struct RewriteOutput;\n"
        "pub fn decode_custom_format_with_report() {}\n"
        "pub fn decode_custom_format_list_with_report() {}\n"
        "pub fn prepare_custom_format_rewrite() {}\n"
        "pub fn prepare_custom_format_list_rewrite() {}\n"
        "pub fn canonical_custom_format() {}\n"
        "pub fn canonical_custom_format_list() {}\n"
        "pub fn rewrite_custom_format() {}\n"
        "pub fn rewrite_custom_format_list() {}\n"
        "fn scanner() {\n"
        "    Budget::new; preflight; scan_custom_format; scan_custom_format_list;\n"
        "    decode_lazy_view; LazyView; buffa; projection; generated;\n"
        "    duplicate; unique; noncanonical; canonical; wire; unknown; raw;\n"
        "    uuid; UUID; count; validate_unique_uuids; zero;\n"
        "    extend_from_slice; execution_requirements; Prepared; execute;\n"
        "    retained; scratch; allocation;\n"
        "}\n"
        "#[cfg(test)]\n"
        "mod tests { #[test] fn hostile_wire_and_registry_rejection() {} }\n",
        encoding="utf-8",
    )
    public = root / boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE
    public.parent.mkdir(parents=True, exist_ok=True)
    public.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_MODULE};\n"
        "#[doc(hidden)]\n"
        f"mod {boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_GENERATED_MODULE} {{\n"
        "    include!(concat!(env!(\"OUT_DIR\"), "
        "\"/buffa-numbers-table-cell-custom-format/generated.rs\"));\n"
        "}\n",
        encoding="utf-8",
    )

    integration = root / boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_SOURCE
    integration.parent.mkdir(parents=True, exist_ok=True)
    integration.write_text(
        "#[test] fn custom_number_text_date_time_set_clear_reset_noop_inverse() {\n"
        "    Number; Text; Date_Time; Custom; set; clear; reset; automatic;\n"
        "    no_op; inverse; byte_for_byte; uuid; registry; refcount; reference;\n"
        "    reuse; cull; unknown; member; nested; source_span; candidate; reopen;\n"
        "    readback; locality; stale; foreign; wrong_family; malformed; unsupported;\n"
        "}\n",
        encoding="utf-8",
    )
    example = root / boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_EXAMPLE
    example.parent.mkdir(parents=True, exist_ok=True)
    example.write_text(
        "fn main() { package.table_cell_custom_format(); "
        "package.edit_table_cell_custom_format(); "
        "package.apply_table_cell_custom_format(); }\n",
        encoding="utf-8",
    )

    for fuzz_path in (
        boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_FUZZ_SOURCE,
        boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_FUZZ_SOURCE,
    ):
        target = root / fuzz_path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("fuzz_target!(|data: &[u8]| { let _ = data; });\n", encoding="utf-8")
    for corpus in (
        boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_FUZZ_CORPUS,
        boundaries.NUMBERS_TABLE_CELL_CUSTOM_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        absolute.mkdir(parents=True, exist_ok=True)
        (absolute / "hostile.seed").write_bytes(b"custom")


def complete_numbers_table_cell_control_popup_split_scaffold(root: Path) -> None:
    """Upgrade the cell-control fixture to the positive split popup contract."""

    add_numbers_table_cell_control_canonical_scaffold(root)
    owner = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE
    owner_source = owner.read_text(encoding="utf-8")
    owner_source = owner_source.replace(
        "fn reject_cross_component_write() {}\n"
        "fn changed_route() { reject_cross_component_write(); let mut budget = (); }\n",
        "",
    )
    owner_source += (
        "struct NativePopUpGraphInput; struct NativePopUpMember; "
        "struct NativePopUpObjectRoute; struct NativeMemberEdit;\n"
        "fn rewrite_native_popup_menu_multi() { NativePopUpGraphInput; member_edits; }\n"
        "fn member_edit_deduplication() { NativeControlOutput; from_edits; deduplicated; }\n"
        "fn component_uuid_ownership() { require_current_uuid; uuid_addition; uuid_removal; IdentifierCensus; reserved_identifier; }\n"
        "fn external_edge_and_save_token_batch() { require_external_edge; prove_cross_component_reference; external_reference; selectors_for_components; SaveTokenBatch; save_tokens; touched_components; }\n"
        "fn cow_reuse_cull_refcount_census() { copy_on_write; prepare_desired_state; existing_popup_model_identifiers; archive_has_popup_reference; census_bnc_references; validate_bnc_refcounts; refcount; cull; }\n"
        "fn aggregate_field_info_transition() { ArchiveInfo; FieldInfo; object_references; replace_control_message_with_transition; aggregate; }\n"
        "fn aggregate_transaction_budget() { TransactionBudget; charge_wire_fields; residual_storage_options; preflight; }\n"
        "fn prepared_archive_zip_reassembly() { prepare_reassembly; execution_requirements; Prepared; reassembly; }\n"
        "fn candidate_reopen_object_locality() { candidate; reopen; verify_package_locality; same_content; }\n"
    )
    owner.write_text(owner_source, encoding="utf-8")

    integration = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE
    integration.write_text(
        "#[test]\n"
        "fn split_popup_menu_lifecycle() {\n"
        "    split cross_component PopUpMenu read no_op replacement copy_on_write create reuse cull refcount inverse apply conflict member locality exact_bytes metadata uuid save_token external alias opaque versioned atomic budget limit reassembly allocation;\n"
        "}\n",
        encoding="utf-8",
    )

    package_fuzz = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_FUZZ_SOURCE
    package_fuzz.write_text(
        package_fuzz.read_text(encoding="utf-8")
        + "fn split_popup_fuzz_lifecycle() { split member PopUpMenu create clear reset inverse conflict atomic; }\n",
        encoding="utf-8",
    )
    codec_fuzz = root / boundaries.NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_SOURCE
    codec_fuzz.write_text(
        codec_fuzz.read_text(encoding="utf-8")
        + "fn prepared_popup_codec_fuzz_replay() { prepare replay unknown limit depth allocation scratch; }\n",
        encoding="utf-8",
    )
    (root / boundaries.NUMBERS_TABLE_CELL_CONTROL_FUZZ_CORPUS / "split_popup.seed").write_bytes(
        b"split-popup"
    )
    (root / boundaries.NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_CORPUS / "prepared_popup.hex").write_bytes(
        b"00"
    )


def add_keynote_movie_playback_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod playback {\n"
        "pub mod transaction;\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_TYPES
        )
        + "}\n"
        "pub use playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume, TimeField};\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_MOVIE_PLAYBACK_CANONICAL_TYPES
        )
        + "struct TransactionBudget;\n"
        + "fn transaction_budget_preflight() { budget; }\n"
        + "fn preflight_candidate_reopen_locality() { ExactArtifacts; inverse; }\n"
        + "fn charge_codec_report_fields_work() { codec_report; wire; fields; work; }\n"
        + "fn charge_archive_snappy_zip_output() { archive; Snappy; ZIP; output; }\n"
        + "fn movie_playback_codec_route() { movie_playback_codec; }\n"
        + "impl Package {\n"
        + "pub fn slide_movie_playback_settings<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector) -> Result<MediaPlaybackSettings, Error> { let _ = (slide, movie); todo!() }\n"
        + "pub fn edit_slide_movie_playback_settings<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector, value: MediaPlaybackSettings) -> Result<Edit, Error> { let _ = (slide, movie, value); todo!() }\n"
        + "pub fn apply_slide_movie_playback_settings(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit { pub fn set(self, value: MediaPlaybackSettings) -> Self { let _ = value; self } pub fn commit(self) -> Result<Commit, Error> { todo!() } }\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_movie_playback;\n"
        "pub use slide_movie_playback::{"
        + ", ".join(sorted(boundaries.KEYNOTE_MOVIE_PLAYBACK_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod slide;\n"
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_MOVIE_PLAYBACK_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::movie::MovieSelector;\n"
        "pub use slide::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume, TimeField};\n",
        encoding="utf-8",
    )
    codec = root / boundaries.KEYNOTE_MOVIE_PLAYBACK_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_MOVIE_PLAYBACK_CODEC_TYPES
        )
        + "pub fn decode_movie_playback_with_report() {}\n"
        + "pub fn prepare_movie_playback_rewrite() {}\n"
        + "pub fn rewrite_movie_playback() {}\n"
        + "// private Buffa lazy sidecar ingress; unknown balanced raw fields; max_input max_output max_fields max_work max_depth max_references try_reserve\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.KEYNOTE_MOVIE_PLAYBACK_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.KEYNOTE_MOVIE_PLAYBACK_CODEC_MODULE};\n",
        encoding="utf-8",
    )
    for fuzz_target in boundaries.KEYNOTE_MOVIE_PLAYBACK_FUZZ_SOURCES:
        absolute = root / fuzz_target
        absolute.parent.mkdir(parents=True, exist_ok=True)
        absolute.write_text(
            "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
            "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
            encoding="utf-8",
        )
    for corpus in boundaries.KEYNOTE_MOVIE_PLAYBACK_FUZZ_CORPORA:
        (root / corpus).mkdir(parents=True, exist_ok=True)
    add_keynote_family_module(root, "MOVIE_PLAYBACK")


def add_keynote_movie_geometry_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod geometry {\n"
        "pub mod transaction;\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_TYPES
        )
        + "}\n"
        "pub use geometry::MovieGeometry;\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_MOVIE_GEOMETRY_CANONICAL_TYPES
        )
        + "struct TransactionBudget;\n"
        + "fn transaction_budget_preflight() { budget; }\n"
        + "fn preflight_candidate_reopen_locality() { ExactArtifacts; inverse; locality; }\n"
        + "fn charge_codec_report_fields_work() { codec_report; wire; fields; work; }\n"
        + "fn charge_archive_snappy_zip_output() { archive; Snappy; ZIP; output; }\n"
        + "fn invalidate_previews_after_geometry_change() { preview; invalidated_previews; }\n"
        + "fn keynote_movie_geometry_codec_route() { keynote_movie_geometry_codec; }\n"
        + "impl Package {\n"
        + "pub fn slide_movie_geometry<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector) -> Result<MovieGeometry, Error> { let _ = (slide, movie); todo!() }\n"
        + "pub fn edit_slide_movie_geometry<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector, value: MovieGeometry) -> Result<Edit, Error> { let _ = (slide, movie, value); todo!() }\n"
        + "pub fn apply_slide_movie_geometry(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit { pub fn set(self, value: MovieGeometry) -> Self { let _ = value; self } pub fn restore_original_size(self) -> Result<Self, Error> { Ok(self) } pub fn set_transform(self, value: MovieTransform) -> Result<Self, Error> { let _ = value; Ok(self) } pub fn flip(self, axis: MovieFlipAxis) -> Result<Self, Error> { let _ = axis; Ok(self) } pub fn commit(self) -> Result<Commit, Error> { todo!() } }\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_movie_geometry;\n"
        "pub use slide_movie_geometry::{"
        + ", ".join(sorted(boundaries.KEYNOTE_MOVIE_GEOMETRY_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod slide;\n"
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_MOVIE_GEOMETRY_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::movie::MovieSelector;\n"
        "pub use slide::media::geometry::MovieGeometry;\n",
        encoding="utf-8",
    )
    codec = root / boundaries.KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_MOVIE_GEOMETRY_CODEC_TYPES
        )
        + "pub fn decode_movie_geometry_with_report() {}\n"
        + "pub fn prepare_movie_geometry_rewrite() {}\n"
        + "pub fn rewrite_movie_geometry() {}\n"
        + "// private Buffa lazy sidecar ingress; unknown balanced raw fields; max_input max_output max_fields max_work max_depth max_references try_reserve\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.KEYNOTE_MOVIE_GEOMETRY_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.KEYNOTE_MOVIE_GEOMETRY_CODEC_MODULE};\n",
        encoding="utf-8",
    )
    for fuzz_target in boundaries.KEYNOTE_MOVIE_GEOMETRY_FUZZ_SOURCES:
        absolute = root / fuzz_target
        absolute.parent.mkdir(parents=True, exist_ok=True)
        absolute.write_text(
            "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
            "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
            encoding="utf-8",
        )
    for corpus in boundaries.KEYNOTE_MOVIE_GEOMETRY_FUZZ_CORPORA:
        (root / corpus).mkdir(parents=True, exist_ok=True)
    add_keynote_family_module(root, "MOVIE_GEOMETRY")

def add_keynote_audio_position_canonical_scaffold(root: Path) -> None:
    """Create the focused selector-first audio-position admission fixture."""

    semantic = root / boundaries.KEYNOTE_AUDIO_POSITION_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub struct Point { pub x: f32, pub y: f32 }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_AUDIO_POSITION_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_AUDIO_POSITION_CANONICAL_TYPES
        )
        + "struct AudioPositionBudget;\n"
        + "fn position_budget_for_package() { let _budget = AudioPositionBudget::for_package(); }\n"
        + "fn select_position(slide: SlideSelector, movie: MovieSelector) { let _ = (slide, movie); }\n"
        + "fn candidate_reopen_readback_verify() { candidate; reopen; readback; verify; }\n"
        + "fn exact_source_inverse() { ExactArtifacts; source_fingerprint; inverse; PatchConflict; }\n"
        + "fn invalidate_previews_after_position_change() { preview; invalidate_preview; deleted_previews; }\n"
        + "fn codec_position_route() { decode_movie_position_with_report; prepare_movie_position_rewrite; }\n"
        + "fn execute_prepared() { let prepared = prepare_movie_position_rewrite(); let output = prepared.execute(); let _ = output; }\n"
        + "impl Package {\n"
        + "    pub fn slide_audio_position<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector) -> Result<Point, Error> { let _ = (slide, movie); todo!() }\n"
        + "    pub fn edit_slide_audio_position<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector, position: Point) -> Result<Edit, Error> { let _ = (slide, movie, position); todo!() }\n"
        + "    pub fn apply_slide_audio_position(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit { pub fn set(self, value: Point) -> Self { let _ = value; self } pub fn commit(self) -> Result<Commit, Error> { todo!() } }\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_AUDIO_POSITION_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_audio_position;\n"
        "pub use slide_audio_position::{"
        + ", ".join(sorted(boundaries.KEYNOTE_AUDIO_POSITION_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_AUDIO_POSITION_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_AUDIO_POSITION_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::movie::MovieSelector;\n"
        "pub use slide::media::Point;\n",
        encoding="utf-8",
    )
    codec = root / boundaries.KEYNOTE_AUDIO_POSITION_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "use buffa::DecodeOptions;\n"
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_AUDIO_POSITION_CODEC_TYPES
        )
        + "pub fn decode_movie_position_with_report() {}\n"
        + "pub fn prepare_movie_position_rewrite() {}\n"
        + "pub fn decode_lazy_view() {}\n"
        + "fn execute() { output_bytes; source; unknown; raw; max_message_bytes; max_fields; max_work_bytes; try_reserve; }\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.KEYNOTE_AUDIO_POSITION_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.KEYNOTE_AUDIO_POSITION_CODEC_MODULE};\n",
        encoding="utf-8",
    )
    integration = root / boundaries.KEYNOTE_AUDIO_POSITION_TEST_SOURCE
    integration.parent.mkdir(parents=True, exist_ok=True)
    integration.write_text(
        "#[test]\n"
        "fn audio_position_round_trip_and_inverse_preview_atomic_limit() {\n"
        "    slide_audio_position; edit_slide_audio_position; apply_slide_audio_position;\n"
        "    inverse; preview; atomic; limit;\n"
        "}\n",
        encoding="utf-8",
    )
    fuzz = root / boundaries.KEYNOTE_AUDIO_POSITION_FUZZ_SOURCE
    fuzz.parent.mkdir(parents=True, exist_ok=True)
    fuzz.write_text(
        "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
        "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
        encoding="utf-8",
    )
    (root / boundaries.KEYNOTE_AUDIO_POSITION_FUZZ_CORPUS).mkdir(parents=True, exist_ok=True)
    add_keynote_family_module(root, "AUDIO_POSITION")


def add_keynote_slide_media_data_canonical_scaffold(root: Path) -> None:
    """Create the focused selector-first slide-media replacement fixture."""

    core = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CORE_SOURCE_ROOT / "archive.rs"
    core.parent.mkdir(parents=True, exist_ok=True)
    core.write_text(
        "impl ArchiveObject {\n"
        "    pub fn clone_with_identity_remap(&self, limits: Limits) {\n"
        "        let _ = (archive_info, messages, original_header, original_canonical_header, remap);\n"
        "    }\n"
        "}\n",
        encoding="utf-8",
    )

    owner = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "mod budget;\n"
        "mod closure;\n"
        "use budget::MediaBudget;\n"
        "pub enum MediaPart { Content, Poster }\n"
        "pub type SlideMediaData<'a> = &'a [u8];\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CANONICAL_TYPES
        )
        + "fn select_media() {}\n"
        + "fn strict_media_metadata_closure() { keynote_media_codec; package_metadata_media_codec; DataInfo; digest; materialized_length; owner; }\n"
        + "fn bounded_candidate_publication() { WireLimits; preflight; candidate; reopen; Verification; locality; }\n"
        + "fn exact_reversible_artifacts() { ExactArtifacts; PatchConflict; inverse; source_fingerprint; }\n"
        + "fn invalidate_previews() { preview; deleted_previews; root_previews_absent; }\n"
        + "fn select_media(budget: &mut MediaBudget) { let _ = budget; }\n"
        + "fn read_selected_media(budget: &mut MediaBudget) { let _ = budget; }\n"
        + "fn validate_media_closure(budget: &mut MediaBudget) { validate_selected_media_closure(budget); }\n"
        + "fn rewrite_metadata_entry(budget: &mut MediaBudget) { let prepared = prepare_package_metadata_media_rewrite(); let requirements = prepared.execution_requirements(); budget.metadata_requirements(requirements); let output = prepared.execute(requirements); let _ = output; }\n"
        + "fn rewrite_media(budget: &mut MediaBudget) { rewrite_metadata_entry(budget); }\n"
        + "fn commit_edit(budget: &mut MediaBudget) { rewrite_media(budget); }\n"
        + "impl Package {\n"
        + "    pub fn slide_media_data<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector, part: MediaPart) -> Result<SlideMediaData<'_>, Error> { let mut budget = MediaBudget::for_package(self)?; select_media(&mut budget); read_selected_media(&mut budget); let _ = (slide, movie, part); todo!() }\n"
        + "    pub fn edit_slide_media_data<'slide, 'movie>(&self, slide: SlideSelector<'slide>, movie: MovieSelector, part: MediaPart) -> Result<Edit<'_>, Error> { let mut budget = MediaBudget::for_package(self)?; Edit::new(&mut budget); let _ = (slide, movie, part); todo!() }\n"
        + "    pub fn apply_slide_media_data(&self, patch: &Patch) -> Result<Commit, Error> { let mut budget = MediaBudget::for_package(self)?; select_media(&mut budget); read_selected_media(&mut budget); let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "    pub const fn before(&self) -> &[u8] { &[] }\n"
        + "    pub const fn after(&self) -> Option<&[u8]> { None }\n"
        + "    pub const fn part(&self) -> MediaPart { MediaPart::Content }\n"
        + "    pub fn set(self, replacement: &[u8]) -> Result<Self, Error> { let mut budget = MediaBudget::for_package(self.source)?; budget.entry_bytes(replacement.len()); let _ = replacement; Ok(self) }\n"
        + "    pub fn commit(self) -> Result<Commit, Error> { let mut budget = MediaBudget::for_package(self.source)?; commit_edit(&mut budget); todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )

    budget = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE
    budget.parent.mkdir(parents=True, exist_ok=True)
    budget.write_text(
        "pub(super) struct MediaBudget;\n"
        "impl MediaBudget {\n"
        "    pub(super) fn for_package(package: &Package) -> Result<Self, Error> { let _ = package; todo!() }\n"
        "    fn charges() {\n"
        + "        "
        + "; ".join(
            f"LimitKind::{variant}"
            for variant in sorted(boundaries.KEYNOTE_SLIDE_MEDIA_DATA_LIMIT_VARIANTS)
        )
        + ";\n"
        "}\n}\n",
        encoding="utf-8",
    )
    closure = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE
    closure.parent.mkdir(parents=True, exist_ok=True)
    closure.write_text(
        "pub(super) fn validate_selected_media_closure(\n"
        "    budget: &mut MediaBudget,\n"
        ") -> Result<(), Error> { let _ = budget; Ok(()) }\n",
        encoding="utf-8",
    )

    package = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[0]
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text(
        "mod slide_media_replacement;\n"
        "pub use slide_media_replacement::{"
        + ", ".join(
            sorted(
                boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CANONICAL_TYPES
                | boundaries.KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_TYPES
            )
        )
        + "};\n",
        encoding="utf-8",
    )
    library = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[1]
    library.parent.mkdir(parents=True, exist_ok=True)
    library.write_text(
        "pub use package::{"
        + ", ".join(
            sorted(
                boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CANONICAL_TYPES
                | boundaries.KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_TYPES
            )
        )
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::movie::MovieSelector;\n",
        encoding="utf-8",
    )
    semantic = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub enum MediaPart { Content, Poster }\n"
        "pub type SlideMediaData<'a> = &'a [u8];\n",
        encoding="utf-8",
    )

    keynote_codec = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES[0]
    keynote_codec.parent.mkdir(parents=True, exist_ok=True)
    keynote_codec.write_text(
        "pub struct DecodeOptions;\n"
        "pub struct DecodeReport;\n"
        "pub struct DataReferenceSnapshot;\n"
        "pub fn decode_data_reference() {}\n"
        "pub fn decode_data_reference_with_report() {}\n"
        "struct DataReferenceLazyView;\n"
        "fn decode_lazy_view() {}\n"
        "// private Buffa lazy ingress; unknown fields; max_message_bytes; max_fields; max_work_bytes; try_reserve\n",
        encoding="utf-8",
    )
    metadata_codec = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES[1]
    metadata_codec.parent.mkdir(parents=True, exist_ok=True)
    metadata_codec.write_text(
        "pub struct DataMetadataMapSource<'source> { payload: &'source [u8], fields: usize, work_bytes: usize, entries: usize, scratch_bytes: usize }\n"
        "pub struct DataMetadataMapEntry;\n"
        "pub trait DataMetadataMapVisitor {}\n"
        "pub struct MediaRewriteBatch { data_removals: &[DataInfoRemoval], owner_removals: &[DataReferenceOwnerRemoval], data_metadata_map_source: Option<DataMetadataMapSource<'source>> }\n"
        "pub struct DataInfoContentReplacement;\n"
        "pub struct DataInfoRemoval;\n"
        "pub struct DataReferenceOwnerRemoval;\n"
        "pub struct RewriteExecutionRequirements;\n"
        "fn from_source() { scan_data_metadata_map(payload, options); }\n"
        "fn scan_data_metadata_map(payload: &[u8], options: DecodeOptions, state: &mut State) { reserve_identity_slot; finish_u64_identities; options.max_fields; let _ = (payload, state); }\n"
        "fn validate_batch(source: &[u8], batch: MediaRewriteBatch, report: DecodeReport) {\n"
        "    if report.data_metadata_map_present { DataInfoMetadataMapDependency; }\n"
        "    for (index, removal) in batch.data_removals.iter().copied().enumerate() {\n"
        "        let map_source = batch.data_metadata_map_source; count_data_owners(source, removal.identifier); owner_removals;\n"
        "        let _ = (index, map_source);\n"
        "    }\n"
        "    for (index, replacement) in batch.data_replacements.iter().copied().enumerate() { let _ = (index, replacement); }\n"
        "}\n"
        "fn verify_postconditions() { data_removals; owner_removals; }\n"
        "pub fn visit_package_metadata_media() {}\n"
        "pub fn prepare_package_metadata_media_rewrite() {}\n"
        "pub fn rewrite_package_metadata_media() {}\n"
        "fn decode_lazy_view() {}\n"
        "struct PackageMetadataMedia;\n"
        "fn prepared_rewrite() { unknown; raw; extend_from_slice; }\n"
        "fn bounded_resources() { max_message_bytes; max_fields; max_work_bytes; try_reserve; }\n",
        encoding="utf-8",
    )
    proto_lib = root / boundaries.KEYNOTE_SLIDE_MEDIA_DATA_CODEC_PUBLIC_SOURCE
    proto_lib.parent.mkdir(parents=True, exist_ok=True)
    proto_lib.write_text(
        "pub mod keynote_media_codec;\n"
        "pub mod package_metadata_media_codec;\n",
        encoding="utf-8",
    )
    add_keynote_family_module(root, "SLIDE_MEDIA_DATA", ("Data", "MediaPart"))

def add_keynote_slide_media_lifecycle_canonical_scaffold(root: Path) -> None:
    """Create the focused Keynote duplicate/remove boundary fixture."""

    owner = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "mod budget;\n"
        "mod clone_payload;\n"
        "mod comment_graph;\n"
        "mod comment_removal;\n"
        "mod graph;\n"
        "mod identifier_watermark;\n"
        "mod metadata;\n"
        "mod node_cache;\n"
        "use budget::MediaLifecycleBudget;\n"
        "pub struct LifecycleCommit;\n"
        "pub struct LifecyclePatch;\n"
        "pub enum LifecycleError { InvalidSource, UnsupportedSource }\n"
        "fn typed_kind(expected_kind: MovieKind, actual: MovieKind) { let _ = (KindMismatch, expected_kind, MovieKind, expected, actual); }\n"
        "fn count_slide_media() { let mut count = 0; let _ = MOVIE_MESSAGE_TYPE; let Some(message) = find() else { continue }; let _ = message; count = count.checked_add(1); let _ = (source_media_count, target_media_count, actual_count); }\n"
        "fn unique_payload() {}\n"
        "fn rewrite_node_cache(budget: &mut MediaLifecycleBudget) { node_cache::prepare_slide_node_build_cache(source, selection, archive, budget); node_cache::validate_candidate_package_node_cache(candidate, selection, budget); }\n"
        "fn select_media() {}\n"
        "fn clone_object() {}\n"
        "fn metadata_transition() {}\n"
        "fn candidate_reopen_verify() {}\n"
        "fn allocate_clone_identities(package: &Package, snapshot: &metadata::MetadataSnapshot<'_>, budget: &mut MediaLifecycleBudget) {\n"
        "    let mut global_max = snapshot.last_identifier().max(snapshot.metadata_identifier_maximum());\n"
        "    let mut object_count = 0usize;\n"
        "    for component in package.state.source.components().iter() { for object in &component.archive().objects { object_count = object_count.checked_add(1); if let Some(identifier) = object.archive_info.identifier { global_max = global_max.max(identifier); } } }\n"
        "    budget.charge_entries(object_count);\n"
        "    let mut remap = Vec::new(); let _ = (&mut remap, global_max);\n"
        "}\n"
        "fn remove_comment_graph(action: LifecycleAction, budget: &mut MediaLifecycleBudget) {\n"
        "    if action == LifecycleAction::Remove {\n"
        "        let plan = comment_removal::plan_comment_removal(package, component_name, selected_object_ids, comment_graph, archive_limits, budget);\n"
        "        let source_ids = plan.removed_object_ids;\n"
        "        let external_authors = plan.unused_external_author_ids;\n"
        "        metadata::IdentityBatch::removals(last_identifier, &source_ids).with_external_reference_removals(&external_authors);\n"
        "        let released_watermark = if action == LifecycleAction::Remove { identifier_watermark::plan_release(source, &source_ids, snapshot.last_identifier(), budget) } else { None };\n"
        "        let identity_batch = match action { LifecycleAction::Remove => { let batch = metadata::IdentityBatch::removals(last_identifier, &source_ids).with_external_reference_removals(&external_authors); match released_watermark { Some(new_last_identifier) => batch.with_new_last_identifier(new_last_identifier), None => batch } }, LifecycleAction::Duplicate => metadata::IdentityBatch::additions(last_identifier, last_identifier + 1, &[]) };\n"
        "        if let Some(new_last_identifier) = released_watermark { identifier_watermark::verify_release(&candidate, new_last_identifier, budget); }\n"
        "    }\n"
        "}\n"
        "fn exact_source_fingerprint_inverse() { ExactArtifacts; source_fingerprint; inverse; PatchConflict; }\n"
        "pub struct Edit<'a> { source: &'a Package, staged: Option<LifecycleAction> }\n"
        "impl<'a> Edit<'a> {\n"
        "    pub fn duplicate(self, slide: SlideSelector, movie: MovieSelector) -> Result<Self, LifecycleError> { let _ = (slide, movie); Ok(self) }\n"
        "    pub fn remove(self, slide: SlideSelector, movie: MovieSelector) -> Result<Self, LifecycleError> { let _ = (slide, movie); Ok(self) }\n"
        "    pub fn duplicate_movie(self, slide: SlideSelector, movie: MovieSelector) -> Result<Self, LifecycleError> { self.duplicate(slide, movie) }\n"
        "    pub fn duplicate_audio(self, slide: SlideSelector, movie: MovieSelector) -> Result<Self, LifecycleError> { self.duplicate(slide, movie) }\n"
        "    pub fn remove_movie(self, slide: SlideSelector, movie: MovieSelector) -> Result<Self, LifecycleError> { self.remove(slide, movie) }\n"
        "    pub fn remove_audio(self, slide: SlideSelector, movie: MovieSelector) -> Result<Self, LifecycleError> { self.remove(slide, movie) }\n"
        "    pub fn commit(self) -> Result<LifecycleCommit, LifecycleError> { let mut budget = MediaLifecycleBudget::for_package(self.source); select_media(&mut budget); clone_object(&mut budget); replace_slide_message_with_lifecycle_refs(&mut budget); metadata_transition(); candidate_reopen_verify(); let _ = budget; todo!() }\n"
        "}\n"
        "impl Package {\n"
        "    pub fn edit_slide_media_lifecycle(&self) -> Edit<'_> { Edit { source: self, staged: None } }\n"
        "    pub fn apply_slide_media_lifecycle(&self, patch: &LifecyclePatch) -> Result<LifecycleCommit, LifecycleError> { let mut budget = MediaLifecycleBudget::for_package(self); replace_slide_message_with_lifecycle_refs(&mut budget); metadata_transition(); candidate_reopen_verify(); let _ = (patch, budget); todo!() }\n"
        "}\n"
        "impl LifecyclePatch { pub fn inverse(&self) -> Self { Self } }\n"
        "fn bounded_allocations() { try_reserve(); checked_add(); checked_mul(); allocation(); }\n",
        encoding="utf-8",
    )
    for child, source in {
        "budget.rs": (
            "struct MediaLifecycleBudget { allocation_bytes: u64, allocations: usize }\n"
            "impl MediaLifecycleBudget {\n"
            "    fn for_package(package: &Package) -> Self { let _ = package; Self { allocation_bytes: 0, allocations: 0 } }\n"
            "    fn charge_allocations(&mut self, amount: usize) { self.charge_allocation_plan(amount, 1); }\n"
            "    fn charge_allocation_plan(&mut self, amount: usize, events: usize) {\n"
            "        let observed_bytes = self.allocation_bytes + amount as u64;\n"
            "        let observed_events = self.allocations + events;\n"
            "        if observed_bytes > max_output || observed_events > max_allocations { return; }\n"
            "        self.allocation_bytes = observed_bytes;\n"
            "        self.allocations = observed_events;\n"
            "    }\n"
            "}\n"
            "fn bounded_allocations() { try_reserve(); checked_add(); checked_mul(); allocation(); }\n"
        ),
        "clone_payload.rs": (
            "fn clone_payload() { clone_with_identity_remap(); candidate; reopen; verify; }\n"
        ),
        "graph.rs": (
            "fn graph() { SlideSelector; MovieSelector; select_media(); "
            "let mut charge = |amount| budget.charge_allocations(amount); "
            "remap_clone_payload_with_budget(source, &mut charge); "
            "IdentifierRewrite; UuidRewrite; replace_slide_message_with_lifecycle_refs(); "
            "rewrite_slide_lifecycle(); rewrite_build(); rewrite_build_chunk(); }\n"
            "fn comment_graph_route() { match direct_drawable_comment() { "
            "Some(root) => plan_comment_graph(package, component_name, root, limits, budget), "
            "None => None }; let _ = (storage_ids, author_ids); continue; "
            "super::comment_clone::rewrite_comment_payload(message, object_remap, budget); }\n"
            "fn private_graph(comment_graph: Option<&CommentGraphPlan>) { if comment_graph.is_none() { return Err(UnsupportedComment); } }\n"
            "fn clone_object() { let limits = reserve_core_header_work(source, object_remap.len(), limits, budget)?; source.clone_with_identity_remap(new_identifier, object_remap, replacements, limits); }\n"
            "fn replace_slide_message_with_lifecycle_refs() { let limits = reserve_core_header_work(object, removed.len(), limits, budget)?; object.replace_message_pruning_object_references_preserving_header(limits); let limits = reserve_core_header_work(object, growth, limits, budget)?; object.replace_message_transitioning_object_references_preserving_header(limits); }\n"
            "fn reserve_core_header_inspection_work(source: &ArchiveObject, limits: ArchiveObjectLimits, budget: &mut MediaLifecycleBudget) { reserve_core_header_work_inner(source, 0, limits, budget, false); }\n"
            "fn reserve_core_header_work(source: &ArchiveObject, additional_references: usize, limits: ArchiveObjectLimits, budget: &mut MediaLifecycleBudget) { reserve_core_header_work_inner(source, additional_references, limits, budget, true); }\n"
            "fn reserve_core_header_work_inner(source: &ArchiveObject, additional_references: usize, limits: ArchiveObjectLimits, budget: &mut MediaLifecycleBudget, rewrite: bool) { let bytes = source.header_length; let events = additional_references; budget.charge_allocation_plan(bytes, events); budget.charge_wire_work(bytes); let _ = (limits, rewrite); }\n"
        ),
        "comment_graph.rs": (
            "enum CommentGraphPayloadPolicy { Strict, PreserveRootExtensions }\n"
            "impl CommentGraphPayloadPolicy { fn reject_unknown_root_fields(self) -> bool { Strict; PreserveRootExtensions; true } }\n"
            "enum CommentGraphComponentPolicy { SameComponent, Package }\n"
            "struct CommentStorageIdentity { identifier: u64, uuid: Option<SuperUuid> }\n"
            "struct CommentStorageLocation { identifier: u64, component_name: Box<str> }\n"
            "struct CommentAuthorDependency;\n"
            "struct CommentGraphPlan { storage_ids: Vec<u64>, storage_locations: Vec<CommentStorageLocation>, storage_identities: Vec<CommentStorageIdentity>, author_ids: Vec<u64>, author_dependencies: Vec<CommentAuthorDependency> }\n"
            "impl CommentGraphPlan { fn root_storage_uuid() { storage_uuid; root_storage_uuid; uuid; } fn storage_component(&self, identifier: u64) -> Option<&str> { storage_locations; binary_search; identifier; } }\n"
            "fn plan_comment_graph(package: &Package, component_name: &str, root: u64, limits: WireLimits, budget: &mut LifecycleBudget) {\n"
            "    comment_storage_codec; decode_comment_storage_archive_with_visitor; CommentStorageVisitor; ReplyCollector;\n"
            "    storage_ids; storage_locations; storage_identities; author_ids; author_dependencies; binary_search; InvalidSource;\n"
            "    LifecycleBudget; charge_entries; charge_references; charge_allocations; try_reserve; CommentGraphPayloadPolicy::Strict; CommentGraphComponentPolicy::SameComponent; plan_comment_graph_with_policy;\n"
            "    let _ = (package, component_name, root, limits, budget);\n"
            "}\n"
            "fn plan_comment_graph_cross_component(package: &Package, component_name: &str, root: u64, limits: WireLimits, budget: &mut LifecycleBudget) { plan_comment_graph_with_policy(package, component_name, root, limits, budget, CommentGraphPayloadPolicy::PreserveRootExtensions, CommentGraphComponentPolicy::Package); }\n"
            "fn plan_comment_graph_with_policy(package: &Package, component_name: &str, root: u64, limits: WireLimits, budget: &mut LifecycleBudget, policy: CommentGraphPayloadPolicy, component_policy: CommentGraphComponentPolicy) { if matches(component_policy, CommentGraphComponentPolicy::SameComponent) && storage_component != component_name { return Err(InvalidSource); } let _ = (package, component_name, root, limits, budget, policy, component_policy, reject_unknown_root_fields); }\n"
        ),
            "comment_removal.rs": (
            "struct CommentRemovalPlan { removed_object_ids: Vec<u64>, retained_comment_storage_ids: Vec<u64>, unused_external_author_ids: Vec<u64> }\n"
            "fn plan_comment_removal(package: &Package, component_name: &str, selected_object_ids: &[u64], comment_graph: &CommentGraphPlan, archive_limits: ArchiveLimits, budget: &mut LifecycleBudget) {\n"
            "    component_name; selected_object_ids; comment_graph; archive_limits; InvalidSource;\n"
            "    LifecycleBudget; budget.charge_entries(1); budget.charge_wire_fields(1); budget.charge_wire_work(1); budget.charge_allocations(1);\n"
            "    validate_known_payload_relationships(package, component.name(), object, comment_graph, limits, budget); reserve_core_header_inspection_work(object, archive_limits, budget); strict_reference_census(object, archive_limits, budget); source_component;\n"
            "    retained_comment_storage_ids; comment_edges; selected_comment_author_edges; pending; processed;\n"
            "    lower_bound_source; insert_sorted_unique; author_dependencies; external_author_ids;\n"
            "    unused_external_author_ids; removed_object_ids; let _ = package;\n"
            "    MOVIE_MESSAGE_TYPE; direct_drawable_comment; object_references; count; comment_storage_ids;\n"
            "    validate_movie_payload_relationship; validate_comment_storage_payload_relationship; COMMENT_STORAGE_MESSAGE_TYPE; MessageInfo; message.data; comment_storage_codec; author; replies;\n"
            "}\n"
            "fn validate_known_payload_relationships(package: &Package, component: &str, object: &ArchiveObject, comment_graph: &CommentGraphPlan, limits: WireLimits, budget: &mut LifecycleBudget) {\n"
            "    validate_movie_payload_relationship(package, component, payload, info, comment_graph.storage_ids, limits, budget);\n"
            "    validate_comment_storage_relationship(package, component, storage_identifier, info, payload, limits, budget);\n"
            "    let _ = (package, component, object, comment_graph, limits, budget);\n"
            "}\n"
            "fn validate_movie_payload_relationship(package: &Package, component_name: &str, payload: &[u8], info: &MessageInfo, comment_storage_ids: &[u64], limits: WireLimits, budget: &mut LifecycleBudget) {\n"
            "    validate_comment_storage_target(package, component_name, identifier);\n"
            "    let _ = (payload, info, comment_storage_ids, limits, budget, MOVIE_MESSAGE_TYPE, direct_drawable_comment, object_references, count);\n"
            "}\n"
            "fn validate_comment_storage_relationship(package: &Package, component_name: &str, storage_identifier: u64, info: &MessageInfo, payload: &[u8], limits: WireLimits, budget: &mut LifecycleBudget) {\n"
            "    let facts = super::comment_graph::validate_comment_storage_payload_relationship(storage_identifier, info, payload, limits, budget);\n"
            "    if let Some(author_identifier) = facts.author_identifier { validate_annotation_author_target(package, author_identifier); }\n"
            "    for reply_identifier in facts.reply_identifiers { validate_comment_storage_target(package, component_name, reply_identifier); }\n"
            "}\n"
            "fn validate_comment_storage_target(package: &Package, component_name: &str, identifier: u64) {\n"
            "    let Some((actual_component, object)) = package.object_with_component(identifier) else { return };\n"
            "    if actual_component != component_name { return; }\n"
            "    validate_single_message_type(object, identifier, COMMENT_STORAGE_MESSAGE_TYPE);\n"
            "}\n"
            "fn validate_annotation_author_target(package: &Package, identifier: u64) {\n"
            "    let Some((_, object)) = package.object_with_component(identifier) else { return };\n"
            "    validate_single_message_type(object, identifier, 212);\n"
            "}\n"
            "fn validate_single_message_type(object: &ArchiveObject, identifier: u64, expected_type: u32) {\n"
            "    if object.archive_info.identifier != Some(identifier) || object.messages.len() != 1 || object.archive_info.message_infos.len() != 1 { return; }\n"
            "    let message = object.messages.first().unwrap();\n"
            "    let info = object.archive_info.message_infos.first().unwrap();\n"
            "    if message.type_ != expected_type || info.type_ != expected_type || usize::try_from(info.length).ok() != Some(message.data.len()) { return; }\n"
            "}\n"
        ),
        "metadata.rs": (
            "fn metadata() { MetadataSnapshot; IdentityBatch; MediaBatch; "
            "package_metadata; candidate; reopen; readback; verify; }\n"
            "fn optional_watermark() { if batch.is_removal_transition() { if let Some(new_last) = batch.new_last_identifier() { removals = removals.with_new_last_object_identifier(new_last); } } }\n"
            "struct IdentityCollector { metadata_identifier_maximum: u64 }\n"
            "impl IdentityCollector<'_> { fn observe_identifier(&mut self, identifier: u64) { self.metadata_identifier_maximum = self.metadata_identifier_maximum.max(identifier); } }\n"
            "impl identity_codec::PackageMetadataVisitor for IdentityCollector<'_> {\n"
            "    fn visit_component(&mut self, component: identity_codec::ComponentDescriptor<'_>) { self.observe_identifier(component.identifier()); }\n"
            "    fn visit_object_uuid(&mut self, binding: identity_codec::ObjectUuidDescriptor<'_>) { self.observe_identifier(binding.object_identifier()); }\n"
            "    fn visit_external_reference(&mut self, reference: identity_codec::ExternalReferenceDescriptor<'_>) { if let Some(identifier) = reference.object_identifier() { self.observe_identifier(identifier); } }\n"
            "    fn visit_data_reference_owner(&mut self, owner: identity_codec::DataReferenceOwnerDescriptor<'_>) { self.observe_identifier(owner.object_identifier()); }\n"
            "    fn visit_ambiguous_object_identifier(&mut self, _component: identity_codec::ComponentDescriptor<'_>, identifier: u64) { self.observe_identifier(identifier); }\n"
            "    fn visit_data_metadata_map(&mut self, object_identifier: u64, _has_unknown_fields: bool) { self.observe_identifier(object_identifier); }\n"
            "}\n"
            "struct NoopIdentityVisitor;\n"
            "impl MetadataSnapshot { fn metadata_identifier_maximum(&self) -> u64 { self.metadata_identifier_maximum } }\n"
        ),
        "identifier_watermark.rs": (
            "fn plan_release(source: &Package, removed_ids: &[u64], expected_last_identifier: u64, budget: &mut LifecycleBudget) -> Result<Option<u64>, Error> {\n"
            "    budget.charge_entries(1); budget.charge_wire_work(1); if removed_ids.binary_search(&expected_last_identifier).is_err() { return Ok(None); }\n"
            "    for component in source.state.source.components().iter() { let objects = &component.archive().objects; for object in objects { let identifier = object.archive_info.identifier; if removed_ids.binary_search(&identifier.unwrap_or(0)).is_ok() { continue; } } }\n"
            "    let maximum_remaining = 0; Ok(Some(maximum_remaining))\n"
            "}\n"
            "fn verify_release(candidate: &Package, new_last_identifier: u64, budget: &mut LifecycleBudget) -> Result<(), Error> {\n"
            "    let mut maximum_remaining = 0; for component in candidate.state.source.components().iter() { let objects = &component.archive().objects; budget.charge_entries(objects.len()); budget.charge_wire_work(objects.len().max(1)); for object in objects { let identifier = object.archive_info.identifier.unwrap_or(0); maximum_remaining = maximum_remaining.max(identifier); } }\n"
            "    if maximum_remaining != new_last_identifier { return Err(Error::Verification); } Ok(())\n"
            "}\n"
        ),
        "node_cache.rs": (
            "struct NodeCacheComponentEdit;\n"
            "fn prepare_slide_node_build_cache() { LifecycleBudget; reserve_core_header_work; "
            "lifecycle_codec::node_cache::decode_slide_node_build_cache_with_report; "
            "SlideNodeBuildCacheEdit; prepare_slide_node_build_cache_rewrite; }\n"
            "fn validate_candidate_node_cache() { unique_slide_node_message; InvalidSource; Verification; }\n"
            "fn validate_candidate_package_node_cache() { unique_slide_node_message; }\n"
            "fn unique_slide_node_message(object: ArchiveObject) {\n"
            "    if object.archive_info.message_infos.len() != object.messages.len() { return InvalidSource; }\n"
            "    let mut index = None; for (position, message) in object.messages.iter().enumerate() {\n"
            "        if message.type_ != SLIDE_NODE_MESSAGE_TYPE { continue; }\n"
            "        if index.replace(position).is_some() { return InvalidSource; }\n"
            "    }\n"
            "}\n"
            "fn charge_node_cache_decode_report() {} fn charge_node_cache_rewrite_report() {}\n"
        ),
    }.items():
        path = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / child
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(source, encoding="utf-8")

    package = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_EXPORT_SOURCES[0]
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text(
        "mod slide_media_lifecycle;\n"
        "pub use slide_media_lifecycle::{LifecycleCommit, LifecycleError, LifecyclePatch};\n",
        encoding="utf-8",
    )
    library = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_EXPORT_SOURCES[1]
    library.parent.mkdir(parents=True, exist_ok=True)
    library.write_text(
        "pub use package::{LifecycleCommit, LifecycleError, LifecyclePatch};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::movie::MovieSelector;\n",
        encoding="utf-8",
    )

    codec = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "use buffa::DecodeOptions as BuffaDecodeOptions;\n"
        "use crate::buffa_keynote_media_lifecycle_generated::LitchiIwaProjection;\n"
        "pub struct DecodeError; pub enum DecodeLimit { Work } pub struct DecodeOptions;\n"
        "pub struct Uuid; pub struct Reference<'a> { raw: &'a [u8] } pub struct UuidSnapshot<'a> { raw: &'a [u8] }\n"
        "pub struct DecodeReport; pub struct SlideLifecycleSnapshot<'a> { source: &'a [u8] }\n"
        "pub struct BuildLifecycleSnapshot<'a> { source: &'a [u8] }\n"
        "pub struct BuildChunkLifecycleSnapshot<'a> { source: &'a [u8] }\n"
        "pub struct IdentifierRewrite; pub struct UuidRewrite; pub struct SlideLifecycleEdit<'a> { source: &'a [u8] }\n"
        "pub struct BuildLifecycleEdit; pub struct BuildChunkLifecycleEdit; pub struct RewriteReport;\n"
        "pub struct PreparedSlideLifecycleRewrite<'a> { source: &'a [u8] }\n"
        "pub fn decode_slide_lifecycle() {} pub fn decode_slide_lifecycle_with_report() {}\n"
        "pub fn decode_build() {} pub fn decode_build_with_report() {}\n"
        "pub fn decode_build_chunk() {} pub fn decode_build_chunk_with_report() {}\n"
        "pub fn prepare_slide_lifecycle_rewrite() {} pub fn rewrite_slide_lifecycle() {}\n"
        "pub fn rewrite_slide_lifecycle_with_report() {}\n"
        "pub fn rewrite_build() {} pub fn rewrite_build_with_report() {}\n"
        "pub fn rewrite_build_chunk() {} pub fn rewrite_build_chunk_with_report() {}\n"
        "fn bounded() { DecodeOptions; max_message_bytes; max_output_bytes; max_fields; max_work_bytes; try_reserve(); }\n"
        "fn strict() { validate_input(); canonical; unknown; raw; decode_lazy_view(); source; output_bytes; readback; }\n"
        "#[doc(hidden)]\npub mod node_cache;\n",
        encoding="utf-8",
    )
    codec_child = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE
    codec_child.parent.mkdir(parents=True, exist_ok=True)
    codec_child.write_text(
        "pub struct SlideNodeBuildCacheSnapshot<'a> { source: &'a [u8] }\n"
        "pub struct SlideNodeBuildCacheEdit;\n"
        "pub struct PreparedSlideNodeBuildCacheRewrite<'a> { source: &'a [u8], output_bytes: usize }\n"
        "pub fn decode_slide_node_build_cache() {}\n"
        "pub fn decode_slide_node_build_cache_with_report() {}\n"
        "pub fn prepare_slide_node_build_cache_rewrite() {}\n"
        "pub fn rewrite_slide_node_build_cache() {}\n"
        "pub fn rewrite_slide_node_build_cache_with_report() {}\n"
        "fn cross_check_projection() { projection::SlideNodeBuildCacheArchiveLazyView; buffa(); decode_lazy_view(); }\n"
        "fn strict() { validate_input(); canonical_bool(); expected_values(); Presence; unknown; raw; source; readback; try_reserve(); output_bytes; noncanonical; }\n",
        encoding="utf-8",
    )
    projection = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE
    projection.parent.mkdir(parents=True, exist_ok=True)
    projection.write_text(
        'syntax = "proto2";\npackage LitchiIwaProjection;\n'
        "message Reference { required uint64 identifier = 1; optional int32 deprecated_type = 2; optional bool deprecated_is_external = 3; }\n"
        "message UUID { required uint64 lower = 1; required uint64 upper = 2; }\n"
        "message BuildChunkIdentifierArchive { optional UUID build_id = 1; optional int32 build_chunk_id = 2; }\n"
        "message BuildArchive { optional Reference drawable = 1; }\n"
        "message BuildChunkArchive { optional Reference build = 1; optional BuildChunkIdentifierArchive build_chunk_identifier = 7; optional UUID build_id = 8; }\n"
        "message SlideNodeBuildCacheArchive { optional uint32 build_event_count = 15; optional uint32 build_event_count_cache_version = 26; optional bool build_event_count_is_up_to_date = 22; optional bool has_explicit_builds = 20; optional uint32 has_explicit_builds_cache_version = 27; optional bool has_explicit_builds_is_up_to_date = 23; }\n",
        encoding="utf-8",
    )
    proto_lib = root / boundaries.KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_PUBLIC_SOURCE
    proto_lib.parent.mkdir(parents=True, exist_ok=True)
    proto_lib.write_text(
        "#[doc(hidden)]\nmod buffa_keynote_media_lifecycle_generated {\n"
        'include!(concat!(env!("OUT_DIR"), "/buffa-keynote-media-lifecycle/iwa_keynote_media_lifecycle_buffa_protos.rs"));\n}\n'
        "#[doc(hidden)]\npub mod keynote_media_lifecycle_codec;\n",
        encoding="utf-8",
    )
    build = root / "crates/litchi-iwa-protos/build.rs"
    build.parent.mkdir(parents=True, exist_ok=True)
    build.write_text(
        "let buffa_keynote_media_lifecycle_out_directory = out_dir.join("
        '"buffa-keynote-media-lifecycle");\n'
        "buffa_build::Config::new().files(&["
        '"KNMediaLifecycleArchive.proto"]).out_dir(&buffa_keynote_media_lifecycle_out_directory) '
        '.include_file("iwa_keynote_media_lifecycle_buffa_protos.rs") '
        ".generate_views(true).lazy_views(true).preserve_unknown_fields(false).compile()?;\n"
        "enforce_keynote_media_lifecycle_projection_budget("
        "&buffa_keynote_media_lifecycle_out_directory,);\n"
        "fn enforce_keynote_media_lifecycle_projection_budget() {\n"
        "const EXPECTED_FILES: usize = 5; const MAX_GENERATED_BYTES: u64 = 192 * 1024;\n"
        "let repeated_views = 0; let lazy_repeated_views = 0;\n"
        "if repeated_views != 0 || lazy_repeated_views != 0 { RepeatedView LazyRepeatedView }\n"
        "}\n",
        encoding="utf-8",
    )


def add_keynote_chart_axis_title_canonical_scaffold(root: Path) -> None:
    """Install a complete axis-title boundary fixture."""

    semantic = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub enum Axis { Category, Value }\n"
        "pub struct ChartSelector;\n",
        encoding="utf-8",
    )
    selector = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text("pub enum SlideSelector { Index }\n", encoding="utf-8")

    owner = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_CHART_AXIS_TITLE_CANONICAL_TYPES
        )
        + "struct ChartAxisTitlePackageBudget;\n"
        + "fn account() { let mut budget = ChartAxisTitlePackageBudget; budget.charge(); budget.account(); }\n"
        + "fn strict_prepared() { keynote_chart_axis_title_codec; decode_axis_titles_with_report(); prepare_axis_title_rewrite(); let prepared = prepared; let requirements = prepared.execution_requirements(); let output = prepared.execute(requirements); let _ = output; }\n"
        + "fn exact_source_authority() { ExactArtifacts; source_fingerprint; inverse; PatchConflict; is_noop; candidate; reopen; verify_locality; same_content; authority; root_preview_deletions; deleted_previews; }\n"
        + "impl Package {\n"
        + "    pub fn slide_chart_axis_title<'slide, 'chart>(&self, slide: SlideSelector<'slide>, chart: ChartSelector<'chart>, axis: Axis) -> Result<Option<String>, Error> { let _ = (slide, chart, axis); todo!() }\n"
        + "    pub fn edit_slide_chart_axis_title<'slide, 'chart>(&self, slide: SlideSelector<'slide>, chart: ChartSelector<'chart>, axis: Axis) -> Result<Edit, Error> { let _ = (slide, chart, axis); todo!() }\n"
        + "    pub fn apply_slide_chart_axis_title(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "    pub fn axis(self) -> Axis { Axis::Category }\n"
        + "    pub fn before(&self) -> Option<&str> { None }\n"
        + "    pub fn after(&self) -> Option<&str> { None }\n"
        + "    pub fn set(self, title: impl AsRef<str>) -> Result<Self, Error> { let _ = title; Ok(self) }\n"
        + "    pub fn clear(self) -> Result<Self, Error> { Ok(self) }\n"
        + "    pub fn commit(self) -> Result<Commit, Error> { todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )
    package = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[0]
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text(
        "mod slide_chart_axis_title;\n"
        "pub use slide_chart_axis_title::{"
        + ", ".join(sorted(boundaries.KEYNOTE_CHART_AXIS_TITLE_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    library = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[1]
    library.write_text(
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_CHART_AXIS_TITLE_CANONICAL_TYPES))
        + "};\n"
        "pub use chart::{Axis, ChartSelector};\n"
        "pub use selector::SlideSelector;\n",
        encoding="utf-8",
    )
    codec = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "pub struct DecodeError;\n"
        "pub enum DecodeLimit { Bytes }\n"
        "pub struct DecodeOptions;\n"
        "pub struct AxisTitleSnapshot;\n"
        "pub struct AxisTitleWrite;\n"
        "pub struct DecodeReport;\n"
        "pub struct RewriteReport;\n"
        "pub enum AxisTitleKind { Category, Value }\n"
        "pub struct PreparedAxisTitleRewrite;\n"
        "pub struct RewriteExecutionRequirements;\n"
        "pub struct RewriteExecutionLimits;\n"
        "const CHART_AXIS_CATEGORY_VISIBLE_FIELD: u32 = 13;\n"
        "const CHART_AXIS_VALUE_VISIBLE_FIELD: u32 = 14;\n"
        "const CHART_AXIS_CATEGORY_TEXT_FIELD: u32 = 15;\n"
        "const CHART_AXIS_VALUE_TEXT_FIELD: u32 = 16;\n"
        "const MAX_RECURSION_LIMIT: u32 = 64;\n"
        "pub fn decode_axis_titles() {}\n"
        "pub fn decode_axis_titles_with_report() {}\n"
        "pub fn prepare_axis_title_rewrite() {}\n"
        "fn preflight_axis_titles() {}\n"
        "fn next_strict_field() {}\n"
        "fn require_canonical_bool() {}\n"
        "pub fn rewrite_axis_title() {}\n"
        "impl PreparedAxisTitleRewrite { fn execution_requirements(self) -> RewriteExecutionRequirements { RewriteExecutionRequirements } fn execute(self) {} }\n"
        "impl AxisTitleKind { pub fn visible_title(self, kind: AxisTitleKind) {} }\n"
        "fn strict() { duplicate; noncanonical; unknown; raw; extend_from_slice; execution_requirements; execute; }\n"
        "#[cfg(test)] mod tests { #[test] fn codec_round_trip() {} }\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.KEYNOTE_CHART_AXIS_TITLE_CODEC_MODULE};\n",
        encoding="utf-8",
    )
    test_source = root / boundaries.KEYNOTE_CHART_AXIS_TITLE_TEST_SOURCES[0]
    test_source.parent.mkdir(parents=True, exist_ok=True)
    test_source.write_text(
        "#[test]\n"
        "fn chart_axis_title_lifecycle() { package.slide_chart_axis_title(slide, chart, axis); package.edit_slide_chart_axis_title(slide, chart, axis); package.apply_slide_chart_axis_title(&patch); deleted_previews = 1; }\n",
        encoding="utf-8",
    )
    for fuzz_target in boundaries.KEYNOTE_CHART_AXIS_TITLE_FUZZ_SOURCES:
        absolute = root / fuzz_target
        absolute.parent.mkdir(parents=True, exist_ok=True)
        absolute.write_text(
            "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
            "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
            encoding="utf-8",
        )
    for corpus in boundaries.KEYNOTE_CHART_AXIS_TITLE_FUZZ_CORPORA:
        (root / corpus).mkdir(parents=True, exist_ok=True)
    add_keynote_family_module(root, "CHART_AXIS_TITLE")


def add_focused_chart_arrangement_owner_scaffolds(root: Path) -> None:
    """Install minimal Numbers/Pages focused Arrange owner fixtures."""

    numbers_owner = root / boundaries.NUMBERS_CHART_ARRANGEMENT_OWNER_SOURCE
    numbers_owner.parent.mkdir(parents=True, exist_ok=True)
    numbers_owner.write_text(
        "use litchi_iwa_protos::chart_arrangement_codec as codec;\n"
        "impl Package {\n"
        "    pub fn sheet_chart_arrangement<'sheet>(&self, sheet_selector: impl Into<SheetSelector<'sheet>>, chart_selector: impl Into<ChartSelector>) -> Result<ChartArrangement, Error> {\n"
        "        select_chart_with_budget(self, sheet_selector.into(), chart_selector.into())\n"
        "    }\n"
        "    pub fn sheet_chart_arrangements<'sheet>(&self, sheet_selector: impl Into<SheetSelector<'sheet>>) -> Result<Box<[ChartArrangement]>, Error> {\n"
        "        discover_sheet_charts(self, resolve_sheet_position(self, sheet_selector.into()))\n"
        "    }\n"
        "    pub fn edit_sheet_chart_arrangement<'sheet>(&self, sheet_selector: impl Into<SheetSelector<'sheet>>, chart_selector: impl Into<ChartSelector>) -> Result<Edit, Error> {\n"
        "        Edit::new(self, sheet_selector.into(), chart_selector.into())\n"
        "    }\n"
        "    pub fn apply_sheet_chart_arrangement(&self, patch: &Patch) -> Result<Commit, Error> {\n"
        "        let selected = select_chart_with_budget(self, SheetSelector::position(patch.position()), ChartSelector::position(patch.position()));\n"
        "        verify_locality(self, selected)\n"
        "    }\n"
        "}\n"
        "fn select_chart_with_budget(_: &Package, _: SheetSelector<'_>, _: ChartSelector) -> Result<ChartArrangement, Error> { todo!() }\n"
        "fn resolve_sheet_position(_: &Package, _: SheetSelector<'_>) -> usize { 0 }\n"
        "fn discover_sheet_charts(_: &Package, _: usize) -> Result<Box<[ChartArrangement]>, Error> { todo!() }\n"
        "fn verify_locality(_: &Package, _: Result<ChartArrangement, Error>) -> Result<Commit, Error> { todo!() }\n"
        "fn codec_path() {\n"
        "    codec::decode_chart_arrangement_with_report();\n"
        "    codec::prepare_chart_arrangement_rewrite();\n"
        "    execution_requirements; execute;\n"
        "}\n",
        encoding="utf-8",
    )
    numbers_package = root / boundaries.NUMBERS_PACKAGE_SOURCE
    numbers_package.parent.mkdir(parents=True, exist_ok=True)
    numbers_package.write_text(
        "mod chart_arrangement;\n",
        encoding="utf-8",
    )
    numbers_public = root / (boundaries.NUMBERS_SOURCE_ROOT / "sheet" / "chart.rs")
    numbers_public.parent.mkdir(parents=True, exist_ok=True)
    numbers_public.write_text(
        "pub mod arrangement {\n"
        "    pub use crate::ChartArrangement;\n"
        "    pub use crate::package::families::chart_arrangement::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};\n"
        "}\n",
        encoding="utf-8",
    )

    pages_owner = root / boundaries.PAGES_BODY_CHART_ARRANGEMENT_OWNER_SOURCE
    pages_owner.parent.mkdir(parents=True, exist_ok=True)
    pages_owner.write_text(
        "use litchi_iwa_protos::chart_arrangement_codec as codec;\n"
        "impl Package {\n"
        "    pub fn body_chart_arrangement(&self, selector: impl Into<BodyChartSelector>) -> Result<ChartArrangement, Error> {\n"
        "        resolve_target(self, selector.into())\n"
        "    }\n"
        "    pub fn body_chart_arrangements(&self) -> Result<Box<[ChartArrangement]>, Error> {\n"
        "        resolve_targets(self)\n"
        "    }\n"
        "    pub fn edit_body_chart_arrangement(&self, selector: impl Into<BodyChartSelector>) -> Result<Edit, Error> {\n"
        "        resolve_target(self, selector.into())\n"
        "    }\n"
        "    pub fn apply_body_chart_arrangement(&self, patch: &Patch) -> Result<Commit, Error> {\n"
        "        let selected = resolve_target(self, BodyChartSelector::position(patch.position()));\n"
        "        verify_locality(self, selected)\n"
        "    }\n"
        "}\n"
        "fn resolve_target(_: &Package, _: BodyChartSelector) -> Result<ChartArrangement, Error> { todo!() }\n"
        "fn resolve_targets(_: &Package) -> Result<Box<[ChartArrangement]>, Error> { todo!() }\n"
        "fn verify_locality(_: &Package, _: Result<ChartArrangement, Error>) -> Result<Commit, Error> { todo!() }\n"
        "fn codec_path() {\n"
        "    codec::decode_chart_arrangement_with_report();\n"
        "    codec::prepare_chart_arrangement_rewrite();\n"
        "    execution_requirements; execute;\n"
        "}\n",
        encoding="utf-8",
    )
    pages_package = root / boundaries.PAGES_PACKAGE_SOURCE
    pages_package.parent.mkdir(parents=True, exist_ok=True)
    pages_package.write_text(
        "mod body_chart_arrangement;\n"
        "pub use body_chart_arrangement::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};\n",
        encoding="utf-8",
    )
    pages_public = root / (boundaries.PAGES_SOURCE_ROOT / "lib.rs")
    pages_public.parent.mkdir(parents=True, exist_ok=True)
    pages_public.write_text(
        "pub use litchi_iwa_common::chart::arrangement::ChartArrangement;\n"
        "pub use package::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};\n",
        encoding="utf-8",
    )
    add_pages_transaction_family_scaffold(root, "body_chart_arrangement")


def add_keynote_slide_table_title_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text("pub struct Settings;\n", encoding="utf-8")
    selector = root / boundaries.KEYNOTE_SLIDE_TABLE_TITLE_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub mod title;\n"
        "pub struct TableSelector;\n"
        "impl TableSelector { pub fn index(index: usize) -> Self { let _ = index; Self } }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TABLE_TITLE_CANONICAL_TYPES
        )
        + "struct TitleBudget;\n"
        + "fn strict_projections() { numbers_table_title_codec; table_info_codec; }\n"
        + "fn raw_preserving() { patch_nested_fields_batched; NestedFieldReplacement; }\n"
        + "fn exact() { ExactArtifacts; inverse; }\n"
        + "fn publish() { prepare_reassembly_with_deletions; root_preview_deletions; }\n"
        + "fn verify() { candidate.validate; verify_locality; }\n"
        + "impl Package {\n"
        + "pub fn slide_table_title_settings<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Settings, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn edit_slide_table_title<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Edit, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn apply_slide_table_title(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit { pub fn set(self, settings: Settings) -> Self { let _ = settings; self } pub fn commit(self) -> Result<Commit, Error> { todo!() } }\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_table_title;\n"
        "pub use slide_table_title::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_TITLE_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_TITLE_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::table::TableSelector;\n"
        "pub use slide::table::title::Settings;\n",
        encoding="utf-8",
    )
    add_keynote_family_module(root, "SLIDE_TABLE_TITLE")

def add_keynote_slide_table_sort_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction {\n"
        "pub struct Edit; pub struct Patch; pub struct Commit; pub struct Diagnostics;\n"
        "pub struct Error; pub struct LimitKind; pub struct Path;\n"
        "}\n"
        "pub use litchi_iwa_common::table::sort::{ColumnIndex, Direction, Order, RowRange, Rule, Scope};\n",
        encoding="utf-8",
    )
    selector = root / boundaries.KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub struct SlideSelector;\n"
        "pub struct TableSelector;\n"
        "impl TableSelector { pub fn index(index: usize) -> Self { let _ = index; Self } }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TABLE_SORT_CANONICAL_TYPES
        )
        + "use super :: slide_table_core :: { Budget,\n"
        + "    select_table, same_target, verify_locality,\n"
        + "};\n"
        + "fn strict_codec() { numbers_table_sort_order_codec; decode_table_model_sort_order_with_report; prepare_table_model_sort_order_rewrite; }\n"
        + "fn prepared() { PreparedTableSortOrderRewrite; execution_requirements; report(); execute; }\n"
        + "fn fields() { TABLE_MODEL_MESSAGE_TYPE; sort_order; }\n"
        + "fn rewrite() { replace_message_preserving_header; }\n"
        + "fn shared_graph() { select_table(); same_target(); verify_locality(); Budget; }\n"
        + "fn transaction() { residual; budget; ExactArtifacts; inverse; is_noop; prepare_reassembly; candidate; reopen; validate; same_content; let deleted_previews: usize = 0; }\n"
        + "impl Package {\n"
        + "pub fn slide_table_sort_order<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Option<Order>, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn edit_slide_table_sort_order<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Edit, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn apply_slide_table_sort_order(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit { pub fn order(&self) -> Option<&Order> { todo!() } pub fn set(self, order: Order) -> Self { let _ = order; self } pub fn clear(self) -> Self { self } pub fn reset(self) -> Self { self } pub fn commit(self) -> Result<Commit, Error> { todo!() } }\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_table_sort_order;\n"
        "pub use slide_table_sort_order::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_SORT_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_SORT_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::table::TableSelector;\n"
        "pub use slide::table::sort::{ColumnIndex, Direction, Order, RowRange, Rule, Scope};\n",
        encoding="utf-8",
    )
    codec = root / boundaries.KEYNOTE_SLIDE_TABLE_SORT_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in (
                "SortOrderSnapshot",
                "PreparedTableSortOrderRewrite",
                "RewriteExecutionRequirements",
                "RewriteExecutionLimits",
            )
        )
        + "pub type PreparedTableModelSortOrderRewrite<'a> = PreparedTableSortOrderRewrite;\n"
        + "pub fn decode_table_model_sort_order_with_report() {}\n"
        + "pub fn prepare_table_model_sort_order_rewrite() {}\n"
        + "pub fn canonical_table_sort_order() {}\n"
        + "pub fn rewrite_table_model_sort_order() {}\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.KEYNOTE_SLIDE_TABLE_SORT_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        "pub mod numbers_table_sort_order_codec;\n"
        "#[doc(hidden)]\n"
        "pub mod table_sort_order_codec;\n",
        encoding="utf-8",
    )


def add_keynote_slide_table_lock_state_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub use litchi_iwa_common::table::lock::State;\n"
        "pub mod transaction {\n"
        "    pub use crate::package::slide_table_lock_state::{\n"
        "        Commit as Commit,\n"
        "        Diagnostics as Diagnostics,\n"
        "        Edit as Edit,\n"
        "        Error as Error,\n"
        "        LimitKind as LimitKind,\n"
        "        Patch as Patch,\n"
        "        Path as Path,\n"
        "    };\n"
        "}\n",
        encoding="utf-8",
    )
    selector = root / boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub struct SlideSelector;\n"
        "pub struct TableSelector;\n"
        "impl TableSelector {\n"
        "    pub const fn index(index: usize) -> Self { let _ = index; Self }\n"
        "    pub const fn position(position: Position) -> Self { let _ = position; Self }\n"
        "}\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_CANONICAL_TYPES
        )
        + "struct LockBudget;\n"
        + "fn strict() { table_info_codec; WireLimits; ExactArtifacts; }\n"
        + "fn transaction() { budget; residual; candidate; reopen; verify_locality; inverse; let deleted_previews: usize = 0; }\n"
        + "impl Package {\n"
        + "pub fn slide_table_lock_state<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<State, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn edit_slide_table_lock_state<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Edit, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn apply_slide_table_lock_state(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "pub fn before(&self) -> State { todo!() } pub fn after(&self) -> State { todo!() } pub fn state(&self) -> State { todo!() }\n"
        + "pub fn set(self, state: State) -> Self { let _ = state; self } pub fn set_state(&mut self, state: State) -> &mut Self { let _ = state; self }\n"
        + "pub fn lock(&mut self) -> &mut Self { self } pub fn unlock(&mut self) -> &mut Self { self } pub fn commit(self) -> Result<Commit, Error> { todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_table_lock_state;\n"
        "pub use slide_table_lock_state::{"
        + ", ".join(
            sorted(boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_CANONICAL_TYPES)
        )
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub use package::{"
        + ", ".join(
            sorted(boundaries.KEYNOTE_SLIDE_TABLE_LOCK_STATE_CANONICAL_TYPES)
        )
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::table::TableSelector;\n"
        "pub use slide::table::lock::State;\n",
        encoding="utf-8",
    )


def add_keynote_slide_table_dimension_canonical_scaffold(root: Path) -> None:
    """Install the smallest complete Keynote dimension facade."""

    semantic = root / boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub use litchi_iwa_common::table::dimension::{Dimension, Error, Points, Size};\n"
        "pub mod transaction {\n"
        "    pub use crate::package::slide_table_dimension::{\n"
        "        Commit as Commit,\n"
        "        Diagnostics as Diagnostics,\n"
        "        Edit as Edit,\n"
        "        Error as Error,\n"
        "        LimitKind as LimitKind,\n"
        "        Patch as Patch,\n"
        "        Path as Path,\n"
        "    };\n"
        "}\n",
        encoding="utf-8",
    )
    selector = root / boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub struct SlideSelector;\n"
        "pub struct TableSelector;\n"
        "impl TableSelector {\n"
        "    pub const fn index(index: usize) -> Self { let _ = index; Self }\n"
        "    pub const fn position(position: Position) -> Self { let _ = position; Self }\n"
        "}\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_CANONICAL_TYPES
        )
        + "struct DimensionBudget;\n"
        + "fn strict() { table_dimension_codec; ExactArtifacts; }\n"
        + "fn transaction() { budget; residual; prepare; execute; candidate; reopen; verify_locality; inverse; let deleted_previews: usize = 0; }\n"
        + "impl Package {\n"
        + "pub fn slide_table_dimension_size<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>, dimension: Dimension) -> Result<Size, Error> { let _ = (slide, table, dimension); todo!() }\n"
        + "pub fn edit_slide_table_dimension_size<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>, dimension: Dimension) -> Result<Edit, Error> { let _ = (slide, table, dimension); todo!() }\n"
        + "pub fn apply_slide_table_dimension_size(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "pub fn size(&self) -> Size { todo!() }\n"
        + "pub fn set(self, size: Size) -> Self { let _ = size; self }\n"
        + "pub fn set_points(self, points: Points) -> Self { let _ = points; self }\n"
        + "pub fn reset(self) -> Self { self }\n"
        + "pub fn commit(self) -> Result<Commit, Error> { todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_table_dimension;\n"
        "pub use slide_table_dimension::{"
        + ", ".join(
            sorted(boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_CANONICAL_TYPES)
        )
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub use package::{"
        + ", ".join(
            sorted(boundaries.KEYNOTE_SLIDE_TABLE_DIMENSION_CANONICAL_TYPES)
        )
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::table::TableSelector;\n"
        "pub use slide::table::dimension::{Dimension, Error, Points, Size};\n",
        encoding="utf-8",
    )


def add_keynote_slide_table_name_canonical_scaffold(root: Path) -> None:
    """Install the smallest complete Keynote name-owner fixture."""

    semantic = root / boundaries.KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub enum Error { Empty, ContainsNul }\n"
        "pub struct Name;\n"
        "pub mod transaction {\n"
        "    pub use crate::package::slide_table_name::{\n"
        "        Commit as Commit,\n"
        "        Diagnostics as Diagnostics,\n"
        "        Edit as Edit,\n"
        "        Error as Error,\n"
        "        LimitKind as LimitKind,\n"
        "        Patch as Patch,\n"
        "        Path as Path,\n"
        "    };\n"
        "}\n",
        encoding="utf-8",
    )
    selector = root / boundaries.KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub struct SlideSelector;\n"
        "pub struct TableSelector;\n"
        "impl TableSelector {\n"
        "    pub const fn index(index: usize) -> Self { let _ = index; Self }\n"
        "    pub const fn position(position: Position) -> Self { let _ = position; Self }\n"
        "}\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TABLE_NAME_CANONICAL_TYPES
        )
        + "struct NameBudget;\n"
        + "fn budget() { residual; budget; }\n"
        + "fn codec() { table_model_discovery_codec; prepare_table_model_name_rewrite; execution_requirements; execute; TABLE_MODEL_MESSAGE_TYPE; table_name; codec_report; requirements; exact_limits; }\n"
        + "fn authority() { validate_package_metadata; validate_slide_metadata; validate_table_info_metadata; validate_global_inbound_references; authority; }\n"
        + "fn publish() { ExactArtifacts; source_fingerprint; inverse; is_noop; PatchConflict; prepare_reassembly; EntryEdit; candidate; reopen; validate; verify_locality; same_content_ignoring_offsets; let deleted_previews = root_preview_deletions(); let diagnostics = Diagnostics { deleted_previews }; }\n"
        + "fn no_op() { if patch.is_noop() { let diagnostics = Diagnostics { deleted_previews: 0 }; } }\n"
        + "impl Package {\n"
        + "pub fn slide_table_name<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Name, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn edit_slide_table_name<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Edit, Error> { let _ = (slide, table); todo!() }\n"
        + "pub fn apply_slide_table_name(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "pub fn path(&self) -> Path { todo!() }\n"
        + "pub fn before(&self) -> &Name { todo!() }\n"
        + "pub fn name(&self) -> &Name { todo!() }\n"
        + "pub fn set(self, name: Name) -> Self { let _ = name; self }\n"
        + "pub fn set_name(self, name: &str) -> Result<Self, Error> { let _ = name; todo!() }\n"
        + "pub fn commit(self) -> Result<Commit, Error> { todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_table_name;\n"
        "pub use slide_table_name::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_NAME_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_NAME_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::table::TableSelector;\n"
        "pub use slide::table::name::{Error, Name};\n",
        encoding="utf-8",
    )
    codec = root / boundaries.KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "const TABLE_NAME_FIELD: u32 = 8;\n"
        + "pub struct DecodeError;\n"
        + "pub struct DecodeOptions;\n"
        + "pub struct TableModelSnapshot;\n"
        + "pub struct TableModelNameWrite;\n"
        + "pub struct TableModelNameRewriteReport;\n"
        + "pub struct TableModelNameRewriteRequirements;\n"
        + "pub struct TableModelNameRewriteLimits;\n"
        + "pub struct TableModelNameRewriteOutput;\n"
        + "pub struct PreparedTableModelNameRewrite;\n"
        + "pub fn decode_table_model_with_report() {}\n"
        + "pub fn prepare_table_model_name_rewrite() {}\n"
        + "pub fn table_model_source_fingerprint() {}\n"
        + "fn known_field() {} fn duplicate() {} fn noncanonical() {}\n"
        + "fn skip_unknown() {} fn nested_wire() {} fn extend_from_slice() {} fn require_wire() {} fn wrong_wire_type() {}\n"
        + "fn readback() {} fn source_fingerprint() {} fn table_name() {}\n"
        + "fn execution_requirements() {} fn execute() {}\n"
        + "#[cfg(test)] mod tests { #[test] fn canonical() {} }\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.KEYNOTE_SLIDE_TABLE_NAME_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        f"pub mod {boundaries.KEYNOTE_SLIDE_TABLE_NAME_CODEC_MODULE};\n",
        encoding="utf-8",
    )
    for test_path in boundaries.KEYNOTE_SLIDE_TABLE_NAME_TEST_SOURCES:
        test = root / test_path
        test.parent.mkdir(parents=True, exist_ok=True)
        test.write_text(
            "#[test]\n"
            "fn owner_name_lifecycle() { package.slide_table_name(slide, table); }\n"
            "#[test]\n"
            "fn exact_noop_reports_zero_previews() {\n"
            "    let noop = package.edit_slide_table_name(slide, table).commit();\n"
            "    assert!(noop.patch().is_noop());\n"
            "    assert_eq!(noop.diagnostics().deleted_previews(), 0);\n"
            "}\n"
            "#[test]\n"
            "fn changed_rename_reports_root_preview_deletions() {\n"
            "    let changed = package.edit_slide_table_name(slide, table).set_name(name).commit();\n"
            "    assert!(changed.diagnostics().changed());\n"
            "    assert_eq!(changed.diagnostics().deleted_previews(), PREVIEWS.len());\n"
            "}\n",
            encoding="utf-8",
        )
    for fuzz_path in boundaries.KEYNOTE_SLIDE_TABLE_NAME_FUZZ_SOURCES:
        fuzz = root / fuzz_path
        fuzz.parent.mkdir(parents=True, exist_ok=True)
        fuzz.write_text(
            "#![no_main]\n"
            "use libfuzzer_sys::fuzz_target;\n"
            "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
            encoding="utf-8",
        )
    for corpus in boundaries.KEYNOTE_SLIDE_TABLE_NAME_FUZZ_CORPORA:
        (root / corpus).mkdir(parents=True, exist_ok=True)


def add_keynote_slide_table_headers_canonical_scaffold(
    root: Path,
    *,
    include_codec: bool = True,
    include_fuzz: bool = True,
    include_tests: bool = True,
) -> None:
    """Install the smallest complete header-owner fixture.

    This fixture deliberately models the source topology rather than the
    native implementation.  It is used to prove that the checker has a
    dormant gate, a private package owner, selector-first methods, a strict
    prepared codec seam, and test/fuzz inventory checks without copying any
    production Rust into the policy suite.
    """

    semantic = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction {\n"
        "    pub use crate::package::slide_table_headers::{\n"
        "        Commit as Commit,\n"
        "        Diagnostics as Diagnostics,\n"
        "        Edit as Edit,\n"
        "        Error as Error,\n"
        "        InvalidReason as InvalidReason,\n"
        "        LimitKind as LimitKind,\n"
        "        Patch as Patch,\n"
        "        Path as Path,\n"
        "    };\n"
        "}\n"
        "pub use litchi_iwa_common::table::headers::{Count, Error, Settings};\n",
        encoding="utf-8",
    )

    selector = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub struct SlideSelector;\n"
        "pub struct TableSelector;\n"
        "impl TableSelector {\n"
        "    pub fn index(index: usize) -> Self { let _ = index; Self }\n"
        "    pub fn position(index: usize) -> Self { let _ = index; Self }\n"
        "}\n",
        encoding="utf-8",
    )

    owner = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_CANONICAL_TYPES
        )
        + "struct HeaderBudget;\n"
        + "enum TransactionWork { One }\n"
        + "fn budget_residual() { let budget = HeaderBudget; let residual = budget; }\n"
        + "fn strict_codec() { header_codec; decode_header; prepare_table_header_settings_rewrite; RewriteExecutionRequirements; }\n"
        + "fn model_rewrite() { TABLE_MODEL_MESSAGE_TYPE; 6_001; rewrite_headers; }\n"
        + "fn prepared_execution() { execution_requirements; exact_limits; execute; codec_report; }\n"
        + "fn exact_transaction() { ExactArtifacts; inverse; is_noop; PatchConflict; }\n"
        + "fn candidate_checks() { candidate.validate; reopen; verify_locality; same_selection; }\n"
        + "fn preserve_previews() { let diagnostics = Diagnostics { deleted_previews: 0 }; }\n"
        + "impl Package {\n"
        + "    pub fn slide_table_header_settings<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Settings, Error> { let _ = (slide, table); todo!() }\n"
        + "    pub fn edit_slide_table_headers<'slide>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector>) -> Result<Edit, Error> { let _ = (slide, table); todo!() }\n"
        + "    pub fn apply_slide_table_headers(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "    pub fn slide_position(&self) -> Position { todo!() }\n"
        + "    pub fn table_position(&self) -> Position { todo!() }\n"
        + "    pub fn path(&self) -> Path { todo!() }\n"
        + "    pub fn before(&self) -> Settings { todo!() }\n"
        + "    pub fn after(&self) -> Settings { todo!() }\n"
        + "    pub fn settings(&self) -> Settings { todo!() }\n"
        + "    pub fn set(self, settings: Settings) -> Self { let _ = settings; self }\n"
        + "    pub fn commit(self) -> Result<Commit, Error> { todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )

    package_export = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[0]
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod slide_table_headers;\n"
        "pub use slide_table_headers::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[1]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::table::TableSelector;\n",
        encoding="utf-8",
    )

    if include_codec:
        codec = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE
        codec.parent.mkdir(parents=True, exist_ok=True)
        codec.write_text(
            "pub const HEADER_ROWS_FIELD: u32 = 9;\n"
            "pub const HEADER_COLUMNS_FIELD: u32 = 10;\n"
            "pub const FOOTER_ROWS_FIELD: u32 = 11;\n"
            "pub const HEADER_ROWS_FROZEN_FIELD: u32 = 12;\n"
            "pub const HEADER_COLUMNS_FROZEN_FIELD: u32 = 13;\n"
            "pub const REPEATING_HEADER_ROWS_FIELD: u32 = 29;\n"
            "pub const REPEATING_HEADER_COLUMNS_FIELD: u32 = 32;\n"
            "pub struct DecodeOptions;\n"
            "pub enum DecodeError { WireTypeMismatch }\n"
            "pub struct TableHeaderSettingsSnapshot;\n"
            "pub struct TableHeaderSettingsWrite;\n"
            "pub struct RewriteReport;\n"
            "pub struct RewriteOutput;\n"
            "pub struct PreparedTableHeaderSettingsRewrite<'a>(&'a [u8]);\n"
            "pub struct RewriteExecutionRequirements;\n"
            "pub struct RewriteExecutionLimits;\n"
            "pub fn decode_table_header_settings() {}\n"
            "pub fn prepare_table_header_settings_rewrite() {}\n"
            "pub fn rewrite_table_header_settings() {}\n"
            "pub fn rewrite_table_header_settings_with_report() {}\n"
            "fn known_field() {}\n"
            "fn duplicate() {}\n"
            "fn noncanonical() {}\n"
            "fn strict_unknown() {}\n"
            "fn unknown() {}\n"
            "fn group() {}\n"
            "fn preserv() {}\n"
            "fn check_execution_limits() {}\n"
            "fn execution_requirements() {}\n"
            "fn execute() {}\n"
            "fn with_max_output_bytes() {}\n"
            "fn with_max_allocations() {}\n"
            "fn with_max_retained_bytes() {}\n"
            "fn with_max_scratch_bytes() {}\n"
            "fn readback() {}\n"
            "fn optional_write() {}\n"
            "fn source() {}\n"
            "// noncanonical duplicate(wire) group policy preserves unknown spans\n"
            "#[cfg(test)]\n"
            "mod tests { #[test] fn strict_codec_roundtrip() {} }\n",
            encoding="utf-8",
        )
        codec_lib = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_PUBLIC_SOURCE
        codec_lib.parent.mkdir(parents=True, exist_ok=True)
        codec_lib.write_text(
            "#[doc(hidden)]\n"
            "pub mod numbers_table_header_settings_codec;\n"
            "#[doc(hidden)]\n"
            "pub mod table_header_settings_codec;\n",
            encoding="utf-8",
        )

    if include_tests:
        integration = root / boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_TEST_SOURCES[0]
        integration.parent.mkdir(parents=True, exist_ok=True)
        integration.write_text(
            "#[test]\n"
            "fn exercises_owner() { slide_table_header_settings(); }\n",
            encoding="utf-8",
        )

    if include_fuzz:
        for fuzz_path in boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_FUZZ_SOURCES:
            absolute = root / fuzz_path
            absolute.parent.mkdir(parents=True, exist_ok=True)
            absolute.write_text(
                "#![no_main]\n"
                "use libfuzzer_sys::fuzz_target;\n"
                "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
                encoding="utf-8",
            )
        for corpus in boundaries.KEYNOTE_SLIDE_TABLE_HEADERS_FUZZ_CORPORA:
            (root / corpus).mkdir(parents=True, exist_ok=True)


def add_keynote_slide_table_appearance_canonical_scaffold(
    root: Path,
    *,
    include_codec: bool = True,
    include_fuzz: bool = True,
    include_tests: bool = True,
) -> None:
    """Install the smallest complete appearance-owner fixture."""

    semantic = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction {\n"
        "    pub use crate::package::slide_table_appearance::{\n"
        "        Commit as Commit,\n"
        "        Diagnostics as Diagnostics,\n"
        "        Edit as Edit,\n"
        "        Error as Error,\n"
        "        LimitKind as LimitKind,\n"
        "        Patch as Patch,\n"
        "        Path as Path,\n"
        "    };\n"
        "}\n"
        "pub use litchi_iwa_common::table::appearance::{Appearance, Banding, "
        "GridlineVisibility, Gridlines, RowSizing};\n",
        encoding="utf-8",
    )
    selector = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub struct SlideSelector;\n"
        "pub struct TableSelector;\n"
        "impl TableSelector { pub fn index(index: usize) -> Self { let _ = index; Self } }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES
        )
        + ""
        + "".join(
            f"fn marker_{index}() {{ {' '.join(markers)} }}\n"
            for index, markers in enumerate(
                boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_MARKER_GROUPS.values()
            )
        )
        + "fn preserve_previews() { let deleted_previews = 0; }\n"
        + "impl Package {\n"
        + "    pub fn slide_table_appearance<'slide, 'table>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector<'table>>) -> Result<Appearance, Error> { let _ = (slide, table); todo!() }\n"
        + "    pub fn edit_slide_table_appearance<'slide, 'table>(&self, slide: impl Into<SlideSelector<'slide>>, table: impl Into<TableSelector<'table>>) -> Result<Edit, Error> { let _ = (slide, table); todo!() }\n"
        + "    pub fn apply_slide_table_appearance(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "    pub fn path(&self) -> Path { todo!() }\n"
        + "    pub fn appearance(&self) -> Appearance { todo!() }\n"
        + "    pub fn set(self, appearance: Appearance) -> Self { let _ = appearance; self }\n"
        + "    pub fn commit(self) -> Result<Commit, Error> { todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )
    package = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[0]
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text(
        "mod slide_table_appearance;\n"
        "pub use slide_table_appearance::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES))
        + "};\n",
        encoding="utf-8",
    )
    lib = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[1]
    lib.write_text(
        "pub use package::{"
        + ", ".join(sorted(boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES))
        + "};\n"
        "pub use selector::SlideSelector;\n"
        "pub use slide::table::TableSelector;\n",
        encoding="utf-8",
    )
    if include_codec:
        codec = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_SOURCE
        codec.parent.mkdir(parents=True, exist_ok=True)
        codec.write_text(
            "pub struct DecodeOptions;\n"
            "pub struct DecodeError;\n"
            "pub struct AppearanceSnapshot;\n"
            "pub struct TableModelSnapshot;\n"
            "pub struct PreparedRewrite;\n"
            "pub type PreparedTableModelStyleRewrite<'a> = PreparedRewrite;\n"
            "pub struct RewriteExecutionRequirements;\n"
            "pub struct RewriteExecutionLimits;\n"
            "pub fn decode_table_model_with_report() {}\n"
            "pub fn prepare_table_model_style_rewrite() {}\n"
            "pub fn canonical_table_style_variation() {}\n"
            "pub fn append_stylesheet_style() {}\n"
            "pub fn resolve_table_style_appearance() {}\n"
            "fn scan_group_fields() {}\n"
            "fn find_group_end() {}\n"
            "fn copy_from_slice() {}\n"
            "fn prepare_execution_requirements_execute() {}\n"
            "fn max_input_bytes() {} fn max_output_bytes() {} fn max_fields() {}\n"
            "fn max_work_bytes() {} fn max_allocations() {}\n"
            "#[cfg(test)]\n"
            "mod tests { #[test] fn strict_appearance_roundtrip() {} }\n",
            encoding="utf-8",
        )
        codec_lib = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE
        codec_lib.parent.mkdir(parents=True, exist_ok=True)
        codec_lib.write_text(
            "#[doc(hidden)]\n"
            "pub mod table_appearance_codec;\n",
            encoding="utf-8",
        )
    if include_tests:
        integration = root / boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_TEST_SOURCES[0]
        integration.parent.mkdir(parents=True, exist_ok=True)
        integration.write_text(
            "#[test]\n"
            "fn exercises_owner() { slide_table_appearance(); edit_slide_table_appearance(); apply_slide_table_appearance(); }\n",
            encoding="utf-8",
        )
    if include_fuzz:
        for fuzz_path in boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_FUZZ_SOURCES:
            absolute = root / fuzz_path
            absolute.parent.mkdir(parents=True, exist_ok=True)
            absolute.write_text(
                "#![no_main]\n"
                "use libfuzzer_sys::fuzz_target;\n"
                "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
                encoding="utf-8",
            )
        for corpus in boundaries.KEYNOTE_SLIDE_TABLE_APPEARANCE_FUZZ_CORPORA:
            (root / corpus).mkdir(parents=True, exist_ok=True)


def add_numbers_table_dimension_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction;\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES
        ),
        encoding="utf-8",
    )
    transaction = root / boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE
    transaction.parent.mkdir(parents=True, exist_ok=True)
    transaction.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_DIMENSION_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {name}(&self, selector: TableSelector, dimension: Dimension) {{}}\n"
            for name in boundaries.NUMBERS_TABLE_DIMENSION_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, table_export, _semantic_export = (
        root / path for path in boundaries.NUMBERS_TABLE_DIMENSION_EXPORT_SOURCES
    )
    lib_export.write_text(
        "pub mod table;\n"
        "pub use table::dimension::{Dimension, Points, Size};\n",
        encoding="utf-8",
    )
    package_export.write_text("pub(crate) mod table_dimension;\n", encoding="utf-8")
    table_export.write_text("pub mod dimension;\n", encoding="utf-8")


def add_numbers_table_cells_read_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub struct State;\n"
        "pub enum Storage { Empty, Stored }\n"
        "pub struct Error;\n"
        "pub struct LimitKind;\n"
        "pub struct Path;\n",
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_CELLS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        "pub fn table_cell() {}\n"
        "pub fn table_cells() {}\n"
        "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, table_export = (
        root / path for path in boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod table;\n", encoding="utf-8")
    package_export.write_text("pub(crate) mod table_cells;\n", encoding="utf-8")
    table_export.write_text("pub mod cells;\n", encoding="utf-8")


def add_numbers_table_cells_mutation_scaffold(root: Path) -> None:
    add_numbers_table_cells_read_scaffold(root)
    semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {name}() {{}}\n"
            for name in boundaries.NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES[1]
    package_export.write_text(
        package_export.read_text(encoding="utf-8")
        + "pub(crate) mod table_cell_edit;\n",
        encoding="utf-8",
    )


def add_numbers_formula_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_FORMULA_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_FORMULA_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    transaction_limit = root / boundaries.NUMBERS_FORMULA_TRANSACTION_LIMIT_SOURCE
    transaction_limit.parent.mkdir(parents=True, exist_ok=True)
    transaction_limit.write_text(
        "pub enum LimitKind { "
        + ", ".join(boundaries.NUMBERS_FORMULA_TRANSACTION_LIMIT_VARIANTS)
        + " }\n",
        encoding="utf-8",
    )
    lib = root / boundaries.NUMBERS_SOURCE_ROOT / "lib.rs"
    lib.parent.mkdir(parents=True, exist_ok=True)
    lib.write_text("pub mod formula;\n", encoding="utf-8")


# ADR 0034 Pages transaction families: modules whose canonical re-export is
# the whole file rather than an inline ``pub mod`` (see the Rust sources).
PAGES_TRANSACTION_FILE_SCOPED_OWNERS = frozenset({"table_lock", "body_table_title"})


def _remove_rust_statements(source: str, pattern: str) -> str:
    return re.sub(pattern, "", source, flags=re.DOTALL)


def _remove_inline_rust_module(source: str, name: str) -> str:
    match = re.search(rf"(?m)^[ \t]*pub[ \t]+mod[ \t]+{re.escape(name)}[ \t]*\{{", source)
    if match is None:
        return source
    depth = 0
    for index in range(match.end() - 1, len(source)):
        if source[index] == "{":
            depth += 1
        elif source[index] == "}":
            depth -= 1
            if depth == 0:
                end = index + 1
                if end < len(source) and source[end] == "\n":
                    end += 1
                return source[: match.start()] + source[end:]
    return source


def add_pages_transaction_family_scaffold(root: Path, owner: str) -> None:
    """Install the ADR 0034 layout of one Pages transaction family.

    ``package.rs`` declares ``pub(crate) mod <owner>;``, the owner defines the
    canonical short names, and only the family's public module re-exports them
    with ``pub use crate::package::<owner>::{...};``.  Prefixed root/package
    re-exports written by older scaffolds are removed.
    """

    _, _, segments, shorts, _ = boundaries.PAGES_TRANSACTION_FAMILY_BY_OWNER[owner]
    legacy = boundaries._pages_transaction_legacy_names(owner)
    source_root = root / boundaries.PAGES_SOURCE_ROOT
    source_root.mkdir(parents=True, exist_ok=True)

    package = root / boundaries.PAGES_PACKAGE_SOURCE
    text = package.read_text(encoding="utf-8") if package.is_file() else ""
    text = _remove_rust_statements(
        text, rf"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+{owner}[ \t]*;\n?"
    )
    for statement in re.findall(
        rf"pub[ \t]+use[ \t]+(?:crate::package::)?{owner}::[^;]*;\n?", text
    ):
        names = set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", statement)) - {
            "pub",
            "use",
            "crate",
            "package",
            "as",
            owner,
        }
        if names and names <= set(shorts) | legacy:
            text = text.replace(statement, "", 1)
    package.write_text(f"pub(crate) mod {owner};\n" + text, encoding="utf-8")

    lib = source_root / "lib.rs"
    text = lib.read_text(encoding="utf-8") if lib.is_file() else ""
    for statement in re.findall(r"pub[ \t]+use[ \t]+package::[^;]*;\n?", text):
        names = set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", statement)) - {
            "pub",
            "use",
            "package",
            "as",
        }
        if names & legacy or (names and names <= set(shorts)):
            text = text.replace(statement, "", 1)
    lib.write_text(text, encoding="utf-8")

    owner_path = root / boundaries.PAGES_SOURCE_ROOT / "package" / f"{owner}.rs"
    owner_path.parent.mkdir(parents=True, exist_ok=True)
    text = owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
    shared_types = boundaries.PAGES_TRANSACTION_SHARED_TYPES.get(owner, frozenset())
    missing = [
        name
        for name in shorts
        if name not in shared_types
        and re.search(rf"\b(?:struct|enum|type|union|trait)[ \t]+{name}\b", text) is None
    ]
    if "Diagnostics" in shared_types and "litchi_core::edit::Diagnostics" not in text:
        text = "pub use litchi_core::edit::Diagnostics;\n" + text
    owner_path.write_text(
        "".join(f"pub struct {name};\n" for name in missing) + text, encoding="utf-8"
    )

    if shared_types:
        core_edit = root / boundaries.PAGES_TRANSACTION_SHARED_TYPE_SOURCES["Diagnostics"]
        core_edit.parent.mkdir(parents=True, exist_ok=True)
        core_edit.write_text(
            "pub struct Diagnostics { changed: bool }\n"
            "impl Diagnostics { pub const fn changed(&self) -> bool { self.changed } }\n",
            encoding="utf-8",
        )

    file_segments = (
        segments if owner in PAGES_TRANSACTION_FILE_SCOPED_OWNERS else segments[:-1]
    )
    current = lib
    directory = source_root
    for segment in file_segments:
        text = current.read_text(encoding="utf-8") if current.is_file() else ""
        if re.search(rf"(?m)^[ \t]*pub[ \t]+mod[ \t]+{segment}[ \t]*[;{{]", text) is None:
            current.write_text(text + f"pub mod {segment};\n", encoding="utf-8")
        nested = directory / segment / "mod.rs"
        current = nested if nested.is_file() or segment == "table" else directory / f"{segment}.rs"
        directory = directory / segment
        current.parent.mkdir(parents=True, exist_ok=True)
    text = current.read_text(encoding="utf-8") if current.is_file() else ""
    text = _remove_rust_statements(
        text, rf"pub[ \t]+use[ \t]+crate::package::{owner}::[^;]*;\n?"
    )
    statement = f"pub use crate::package::{owner}::{{{', '.join(shorts)}}};\n"
    if owner in PAGES_TRANSACTION_FILE_SCOPED_OWNERS:
        text += statement
    else:
        text = _remove_inline_rust_module(text, segments[-1])
        text += f"pub mod {segments[-1]} {{\n    {statement}}}\n"
    current.write_text(text, encoding="utf-8")


def add_pages_section_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_SECTION_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_SECTION_SETTINGS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_SECTION_SETTINGS_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, section_export = (
        root / path for path in boundaries.PAGES_SECTION_SETTINGS_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod section;\n", encoding="utf-8")
    package_export.write_text("pub(crate) mod section_settings;\n", encoding="utf-8")
    section_export.write_text(
        "pub mod settings;\npub struct Settings;\n", encoding="utf-8"
    )


def add_pages_section_background_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_SECTION_BACKGROUND_CANONICAL_TYPES
        )
        + "impl Edit {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_SECTION_BACKGROUND_EDIT_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_SECTION_BACKGROUND_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_SECTION_BACKGROUND_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, section_export = (
        root / path for path in boundaries.PAGES_SECTION_BACKGROUND_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod section;\n", encoding="utf-8")
    package_export.write_text(
        "pub(crate) mod section_background;\n", encoding="utf-8"
    )
    section_export.write_text(
        "pub mod background;\npub struct Background;\n", encoding="utf-8"
    )


def add_pages_output_canonical_scaffold(
    root: Path,
    *,
    write_signature: str = (
        "pub fn write_to<W: Write + ?Sized>(self, writer: &mut W) "
        "-> Result<(), WriteError>"
    ),
    error_declaration: str = "pub struct WriteError;",
    root_export: str = "pub use package::WriteError;",
) -> None:
    package = root / boundaries.PAGES_PACKAGE_SOURCE
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text(
        f"{error_declaration}\n"
        "impl Package {\n"
        f"    {write_signature} {{ todo!() }}\n"
        "    fn source_bytes(&self) -> &[u8] { &[] }\n"
        "}\n"
        "#[cfg(test)]\n"
        "mod tests {\n"
        "    pub fn source_bytes() {}\n"
        "    pub fn from_archive_bytes() {}\n"
        "}\n",
        encoding="utf-8",
    )
    export = root / boundaries.PAGES_PACKAGE_OUTPUT_EXPORT_SOURCE
    export.parent.mkdir(parents=True, exist_ok=True)
    export.write_text(root_export + "\n", encoding="utf-8")


def add_pages_table_lock_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_TABLE_LOCK_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub use litchi_iwa_common::table::lock::State;\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_TABLE_LOCK_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_TABLE_LOCK_CANONICAL_TYPES
        )
        + "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_TABLE_LOCK_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_TABLE_LOCK_EXPORT_SOURCES[0]
    table_export = root / boundaries.PAGES_TABLE_LOCK_EXPORT_SOURCES[2]
    package_export = root / boundaries.PAGES_TABLE_LOCK_PACKAGE_EXPORT_SOURCE
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod table;\n"
        "pub use package::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod table_lock;\n"
        "pub use table_lock::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    table_export.write_text(
        "pub mod lock;\n", encoding="utf-8"
    )
    add_pages_transaction_family_scaffold(root, "table_lock")


def add_pages_table_title_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_TABLE_TITLE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub use crate::package::body_table_title::{"
        "Commit as Commit, Diagnostics as Diagnostics, "
        "Edit as Edit, Error as Error, "
        "LimitKind as LimitKind, Patch as Patch};\n"
        "pub use litchi_iwa_common::table::title::Settings;\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_TABLE_TITLE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_TABLE_TITLE_CANONICAL_TYPES
        )
        + "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_TABLE_TITLE_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_TABLE_TITLE_EXPORT_SOURCES[0]
    package_export = root / boundaries.PAGES_TABLE_TITLE_EXPORT_SOURCES[1]
    table_export = root / boundaries.PAGES_TABLE_TITLE_EXPORT_SOURCES[2]
    selector = root / boundaries.PAGES_TABLE_TITLE_SELECTOR_SOURCE
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod table;\n"
        "pub use selector::BodyTableSelector;\n"
        "pub use package::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod body_table_title;\n"
        "pub use body_table_title::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    table_export.write_text("pub mod title;\n", encoding="utf-8")
    selector.write_text("pub enum BodyTableSelector<'a> { Name(&'a str) }\n", encoding="utf-8")
    add_pages_transaction_family_scaffold(root, "body_table_title")


def add_pages_table_name_owner_scaffold(root: Path) -> None:
    """Install the minimum future owner needed to activate name retirement."""

    owner = root / boundaries.PAGES_TABLE_NAME_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        "    pub fn body_table_name(selector: BodyTableSelector) {}\n"
        "    pub fn edit_body_table_name(selector: BodyTableSelector) {}\n"
        "    pub fn apply_body_table_name(patch: BodyTableNamePatch) {}\n"
        "}\n",
        encoding="utf-8",
    )
    package = root / boundaries.PAGES_TABLE_NAME_EXPORT_SOURCES[1]
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text("mod body_table_name;\n", encoding="utf-8")
    lib = root / boundaries.PAGES_TABLE_NAME_EXPORT_SOURCES[0]
    lib.write_text(
        "pub use selector::BodyTableSelector;\n",
        encoding="utf-8",
    )
    selector = root / boundaries.PAGES_TABLE_NAME_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text("pub struct BodyTableSelector;\n", encoding="utf-8")


def add_pages_table_headers_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_TABLE_HEADERS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub use crate::package::body_table_headers::{"
        "Commit as Commit, Diagnostics as Diagnostics, "
        "Edit as Edit, Error as Error, "
        "LimitKind as LimitKind, Patch as Patch};\n"
        "pub use litchi_iwa_common::table::headers::{Count, Error, Settings};\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_TABLE_HEADERS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_TABLE_HEADERS_CANONICAL_TYPES
        )
        + "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_TABLE_HEADERS_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_TABLE_HEADERS_EXPORT_SOURCES[0]
    package_export = root / boundaries.PAGES_TABLE_HEADERS_EXPORT_SOURCES[1]
    table_export = root / boundaries.PAGES_TABLE_HEADERS_EXPORT_SOURCES[2]
    selector = root / boundaries.PAGES_TABLE_HEADERS_SELECTOR_SOURCE
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod table;\n"
        "pub use selector::BodyTableSelector;\n"
        "pub use package::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod body_table_headers;\n"
        "pub use body_table_headers::{Commit, "
        "Diagnostics, Edit, Error, "
        "LimitKind, Patch};\n",
        encoding="utf-8",
    )
    table_export.write_text("pub mod headers;\n", encoding="utf-8")
    selector.write_text("pub struct BodyTableSelector;\n", encoding="utf-8")
    add_pages_transaction_family_scaffold(root, "body_table_headers")


def add_pages_table_dimension_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_TABLE_DIMENSION_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub use crate::package::body_table_dimension::{"
        "Commit as Commit, Diagnostics as Diagnostics, "
        "Edit as Edit, Error as Error, "
        "LimitKind as LimitKind, Patch as Patch};\n"
        "pub use litchi_iwa_common::table::dimension::{Dimension, Points, Size};\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_TABLE_DIMENSION_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_TABLE_DIMENSION_CANONICAL_TYPES
        )
        + "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_TABLE_DIMENSION_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_TABLE_DIMENSION_EXPORT_SOURCES[0]
    package_export = root / boundaries.PAGES_TABLE_DIMENSION_EXPORT_SOURCES[1]
    table_export = root / boundaries.PAGES_TABLE_DIMENSION_EXPORT_SOURCES[2]
    selector = root / boundaries.PAGES_TABLE_DIMENSION_SELECTOR_SOURCE
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod table;\n"
        "pub use selector::BodyTableSelector;\n"
        "pub use package::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod body_table_dimension;\n"
        "pub use body_table_dimension::{Commit, "
        "Diagnostics, Edit, Error, "
        "LimitKind, Patch};\n",
        encoding="utf-8",
    )
    table_export.write_text("pub mod dimension;\n", encoding="utf-8")
    selector.write_text("pub struct BodyTableSelector;\n", encoding="utf-8")
    add_pages_transaction_family_scaffold(root, "body_table_dimension")


def add_pages_table_sort_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_TABLE_SORT_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction {\n"
        "pub use crate::package::body_table_sort::{"
        "Commit as Commit, Diagnostics as Diagnostics, "
        "Edit as Edit, Error as Error, "
        "LimitKind as LimitKind, Patch as Patch};\n}\n"
        "pub use litchi_iwa_common::table::sort::{"
        + ", ".join(boundaries.PAGES_TABLE_SORT_SEMANTIC_TYPES)
        + "};\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_TABLE_SORT_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_TABLE_SORT_CANONICAL_TYPES
        )
        + "impl Package {\n"
        "pub fn body_table_sort_order(selector: BodyTableSelector) {}\n"
        "pub fn edit_body_table_sort_order(selector: BodyTableSelector) {}\n"
        "pub fn apply_body_table_sort_order(patch: &Patch) {}\n"
        "}\n"
        "fn table_model_sort_order_budget_candidate_reopen_locality_reassembly_atomic_inverse_source_bytes_unchanged() {}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_TABLE_SORT_EXPORT_SOURCES[0]
    package_export = root / boundaries.PAGES_TABLE_SORT_EXPORT_SOURCES[1]
    table_export = root / boundaries.PAGES_TABLE_SORT_EXPORT_SOURCES[2]
    selector = root / boundaries.PAGES_TABLE_SORT_SELECTOR_SOURCE
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod table;\n"
        "pub use selector::BodyTableSelector;\n"
        "pub use package::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod body_table_sort;\n"
        "pub use body_table_sort::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    table_export.write_text("pub mod sort;\n", encoding="utf-8")
    selector.write_text("pub struct BodyTableSelector;\n", encoding="utf-8")
    codec = root / boundaries.PAGES_TABLE_SORT_CODEC_SOURCE
    codec.parent.mkdir(parents=True, exist_ok=True)
    codec.write_text(
        "".join(
            f"pub struct {name};\n"
            if name[0].isupper() and not name.startswith("decode") and not name.startswith("prepare") and not name.startswith("canonical") and not name.startswith("rewrite")
            else f"pub fn {name}() {{}}\n"
            for name in boundaries.PAGES_TABLE_SORT_CODEC_REQUIRED_APIS
        )
        + "// unknown unknown overlong report SORT_ORDER_FIELD SORT_RULES_FIELD\n",
        encoding="utf-8",
    )
    codec_lib = root / boundaries.PAGES_TABLE_SORT_CODEC_PUBLIC_SOURCE
    codec_lib.parent.mkdir(parents=True, exist_ok=True)
    codec_lib.write_text(
        "#[doc(hidden)]\n"
        "pub mod table_sort_order_codec { pub use super::numbers_table_sort_order_codec::*; }\n",
        encoding="utf-8",
    )
    for target in (
        boundaries.PAGES_TABLE_SORT_CODEC_FUZZ_SOURCE,
        boundaries.PAGES_TABLE_SORT_FUZZ_SOURCE,
    ):
        path = root / target
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("fuzz_target!(|data: &[u8]| { let _ = data; });\n", encoding="utf-8")
    for corpus in (
        boundaries.PAGES_TABLE_SORT_CODEC_FUZZ_CORPUS,
        boundaries.PAGES_TABLE_SORT_FUZZ_CORPUS,
    ):
        (root / corpus).mkdir(parents=True, exist_ok=True)
    add_pages_transaction_family_scaffold(root, "body_table_sort")


def add_pages_table_appearance_canonical_scaffold(
    root: Path,
    *,
    include_codec: bool = True,
    include_fuzz: bool = True,
    include_tests: bool = True,
) -> None:
    """Install a complete synthetic body-table appearance owner."""

    semantic = root / boundaries.PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction {\n"
        "    pub use crate::package::body_table_appearance::{\n"
        "        Commit as Commit,\n"
        "        Diagnostics as Diagnostics,\n"
        "        Edit as Edit,\n"
        "        Error as Error,\n"
        "        LimitKind as LimitKind,\n"
        "        Patch as Patch,\n"
        "        Path as Path,\n"
        "    };\n"
        "}\n"
        "pub use litchi_iwa_common::table::appearance::{Appearance, Banding, "
        "GridlineVisibility, Gridlines, RowSizing};\n",
        encoding="utf-8",
    )
    selector = root / boundaries.PAGES_TABLE_APPEARANCE_SELECTOR_SOURCE
    selector.parent.mkdir(parents=True, exist_ok=True)
    selector.write_text(
        "pub struct BodyTableSelector<'a> { name: &'a str }\n"
        "impl<'a> BodyTableSelector<'a> { pub fn name(name: &'a str) -> Self { Self { name } } }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_TABLE_APPEARANCE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    marker_functions = "".join(
        f"fn marker_{index}() {{ {' '.join(markers)} }}\n"
        for index, markers in enumerate(
            boundaries.PAGES_TABLE_APPEARANCE_OWNER_MARKER_GROUPS.values()
        )
    )
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_TABLE_APPEARANCE_CANONICAL_TYPES
        )
        + marker_functions
        + "fn preserve_previews() { let deleted_previews = 0; let root_preview_deletions = (); }\n"
        + "impl Package {\n"
        + "    pub fn body_table_appearance<'table>(&self, table: impl Into<BodyTableSelector<'table>>) -> Result<Appearance, Error> { let _ = table; todo!() }\n"
        + "    pub fn edit_body_table_appearance<'table>(&self, table: impl Into<BodyTableSelector<'table>>) -> Result<Edit, Error> { let _ = table; todo!() }\n"
        + "    pub fn apply_body_table_appearance(&self, patch: &Patch) -> Result<Commit, Error> { let _ = patch; todo!() }\n"
        + "}\n"
        + "impl Edit {\n"
        + "    pub fn set(self, appearance: Appearance) -> Self { let _ = appearance; self }\n"
        + "    pub fn commit(self) -> Result<Commit, Error> { todo!() }\n"
        + "}\n",
        encoding="utf-8",
    )
    lib = root / boundaries.PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[0]
    package = root / boundaries.PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[1]
    table = root / boundaries.PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[2]
    lib.parent.mkdir(parents=True, exist_ok=True)
    lib.write_text(
        "pub mod table;\n"
        "pub use selector::BodyTableSelector;\n"
        "pub use package::{"
        + ", ".join(boundaries.PAGES_TABLE_APPEARANCE_CANONICAL_TYPES)
        + "};\n",
        encoding="utf-8",
    )
    package.parent.mkdir(parents=True, exist_ok=True)
    package.write_text(
        "mod body_table_appearance;\n"
        "pub use body_table_appearance::{"
        + ", ".join(boundaries.PAGES_TABLE_APPEARANCE_CANONICAL_TYPES)
        + "};\n",
        encoding="utf-8",
    )
    table.write_text("pub mod appearance;\n", encoding="utf-8")
    if include_codec:
        codec = root / boundaries.PAGES_TABLE_APPEARANCE_CODEC_SOURCE
        codec.parent.mkdir(parents=True, exist_ok=True)
        api_lines = []
        for api in boundaries.PAGES_TABLE_APPEARANCE_CODEC_REQUIRED_APIS:
            if api == "PreparedTableModelStyleRewrite":
                api_lines.append("pub type PreparedTableModelStyleRewrite<'a> = PreparedRewrite<'a>;\n")
            elif api.startswith("decode_") or api.startswith("prepare_") or api.startswith("canonical_") or api.startswith("append_") or api.startswith("resolve_"):
                api_lines.append(f"pub fn {api}() {{}}\n")
            else:
                api_lines.append(f"pub struct {api};\n")
        codec.write_text(
            "".join(api_lines)
            + "fn scan_group_fields() {} fn find_group_end() {} fn copy_from_slice() {}\n"
            + "fn max_input_bytes() {} fn max_output_bytes() {} fn max_fields() {}\n"
            + "fn max_work_bytes() {} fn max_allocations() {}\n"
            + "fn execution_requirements() {} fn execute() {}\n"
            + "#[cfg(test)] mod tests { #[test] fn strict_appearance_roundtrip() {} }\n",
            encoding="utf-8",
        )
        codec_lib = root / boundaries.PAGES_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE
        codec_lib.parent.mkdir(parents=True, exist_ok=True)
        codec_lib.write_text(
            "#[doc(hidden)]\n"
            "pub mod table_appearance_codec;\n",
            encoding="utf-8",
        )
    if include_tests:
        integration = root / boundaries.PAGES_TABLE_APPEARANCE_TEST_SOURCES[0]
        integration.parent.mkdir(parents=True, exist_ok=True)
        integration.write_text(
            "#[test]\n"
            "fn exercises_owner() { body_table_appearance(); edit_body_table_appearance(); apply_body_table_appearance(); }\n",
            encoding="utf-8",
        )
    if include_fuzz:
        for fuzz_path in boundaries.PAGES_TABLE_APPEARANCE_FUZZ_SOURCES:
            path = root / fuzz_path
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(
                "#![no_main]\nuse libfuzzer_sys::fuzz_target;\n"
                "fuzz_target!(|data: &[u8]| { let _ = data; });\n",
                encoding="utf-8",
            )
        for corpus in boundaries.PAGES_TABLE_APPEARANCE_FUZZ_CORPORA:
            (root / corpus).mkdir(parents=True, exist_ok=True)
    add_pages_transaction_family_scaffold(root, "body_table_appearance")


def add_pages_header_footer_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_HEADER_FOOTER_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub struct HeaderFooter;\n"
        "pub enum Template { First, Even, Odd }\n"
        "pub enum Kind { Header, Footer }\n"
        "pub struct HeaderFooterSelector<'a> { section: &'a str }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_HEADER_FOOTER_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_HEADER_FOOTER_CANONICAL_TYPES
        )
        + "impl Package {\n"
        + "pub fn header_footers() {}\n"
        + "pub fn edit_header_footer_text() {}\n"
        + "pub fn apply_header_footer_text() {}\n"
        + "}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_HEADER_FOOTER_EXPORT_SOURCES[0]
    package_export = root / boundaries.PAGES_HEADER_FOOTER_EXPORT_SOURCES[1]
    semantic_export = root / boundaries.PAGES_HEADER_FOOTER_EXPORT_SOURCES[2]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod header_footer;\n"
        "pub use header_footer::HeaderFooterSelector;\n"
        "pub use package::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod header_footer_text;\n"
        "pub use header_footer_text::{Commit, "
        "Diagnostics, Edit, Error, "
        "LimitKind, Patch};\n",
        encoding="utf-8",
    )
    semantic_export.write_text(
        "pub struct HeaderFooter;\n"
        "pub enum Template { First, Even, Odd }\n"
        "pub enum Kind { Header, Footer }\n"
        "pub struct HeaderFooterSelector<'a> { section: &'a str }\n",
        encoding="utf-8",
    )
    add_pages_transaction_family_scaffold(root, "header_footer_text")


def add_pages_section_text_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_SECTION_TEXT_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub enum SectionSelector { Index(usize), Name(String) }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_SECTION_TEXT_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_SECTION_TEXT_CANONICAL_TYPES
        )
        + "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_SECTION_TEXT_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_SECTION_TEXT_EXPORT_SOURCES[0]
    package_export = root / boundaries.PAGES_SECTION_TEXT_EXPORT_SOURCES[1]
    selector_export = root / boundaries.PAGES_SECTION_TEXT_EXPORT_SOURCES[2]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod selector;\n"
        "pub use selector::SectionSelector;\n"
        "pub use package::{Commit, Diagnostics, Edit, "
        "Error, LimitKind, Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod section_text;\n"
        "pub use section_text::{Commit, Diagnostics, Edit, "
        "Error, LimitKind, Patch};\n",
        encoding="utf-8",
    )
    selector_export.write_text(
        "pub enum SectionSelector { Index(usize), Name(String) }\n",
        encoding="utf-8",
    )
    add_pages_transaction_family_scaffold(root, "section_text")


def add_pages_footnote_lifecycle_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub struct Position;\n"
        "pub enum Selector { At(Position), Index(usize) }\n"
        "pub struct Footnote { position: Position }\n"
        "pub use crate::package::body_footnote::{"
        "Commit as Commit, Diagnostics as Diagnostics, "
        "Edit as Edit, Error as Error, "
        "LimitKind as LimitKind, Patch as Patch};\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_FOOTNOTE_LIFECYCLE_CANONICAL_TYPES
        )
        + "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_FOOTNOTE_LIFECYCLE_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export = root / boundaries.PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[0]
    package_export = root / boundaries.PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[1]
    footnote_export = root / boundaries.PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[2]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_export.write_text(
        "pub mod footnote;\n"
        "pub use package::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    package_export.parent.mkdir(parents=True, exist_ok=True)
    package_export.write_text(
        "mod body_footnote;\n"
        "pub use body_footnote::{Commit, Diagnostics, "
        "Edit, Error, LimitKind, "
        "Patch};\n",
        encoding="utf-8",
    )
    footnote_export.write_text("pub mod body;\n", encoding="utf-8")
    add_pages_transaction_family_scaffold(root, "body_footnote")
