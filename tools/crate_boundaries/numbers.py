"""Dependency and public API checks for Numbers."""

from __future__ import annotations

from .common import (
    CAMEL_CASE_WORD,
    CARGO_SECTION_HEADER,
    IWA_CODEC_BUILD_SOURCE,
    IWA_CODEC_PRODUCTION_GUARD_SOURCE,
    IWA_COMMON_TABLE_READ_IMPORT_STATEMENT,
    IWA_DEPENDENCY_CODEC_PUBLIC_PHYSICAL,
    IWA_FOCUSED_TABLE_MERGES_RAW_ID_PARAMETER,
    IWA_FORMULA_ENVELOPE_LOCAL_SCHEMA_CONSTANTS,
    IWA_FORMULA_ENVELOPE_LOCAL_SCHEMA_DECLARATIONS,
    IWA_FORMULA_ENVELOPE_WIRE_IMPORT,
    IWA_FORMULA_RENDER_LOCAL_VISITOR_DECLARATIONS,
    IWA_FORMULA_RENDER_READER_SOURCES,
    IWA_FORMULA_RENDER_WIRE_IMPORT,
    IWA_NUMBERS_WIRE_TABLE_CELLS_PEER_IMPORT,
    IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_SOURCE,
    IWA_TABLE_CELL_FORMAT_ENGINE_SOURCES,
    IWA_TABLE_MERGE_FORMULA_CODEC_MODULE,
    IWA_TABLE_MERGE_TRANSACTION_SPECS,
    Iterable,
    NUMBERS_CHART_ARRANGEMENT_OWNER_SOURCE,
    NUMBERS_CHART_ARRANGEMENT_PACKAGE_MODULE,
    NUMBERS_EXTRACTOR_SOURCE,
    NUMBERS_PACKAGE_SOURCE,
    NUMBERS_SOURCE_ROOT,
    NUMBERS_TABLE_MERGE_PACKAGE_SOURCE,
    NUMBERS_TABLE_MERGE_READER_BYTE_RETURN,
    NUMBERS_TABLE_MERGE_READER_FUNCTION_NAME,
    NUMBERS_TABLE_MERGE_READER_PUBLIC_FIELD,
    NUMBERS_TABLE_MERGE_SOURCE,
    PACKAGE_METADATA_BARE_CALL,
    PACKAGE_METADATA_CODEC_NO_EAGER_PATTERNS,
    PACKAGE_METADATA_CODEC_SOURCE,
    PACKAGE_METADATA_COMBINED_BATCH,
    PACKAGE_METADATA_COMBINED_FUNCTION,
    PACKAGE_METADATA_NO_EAGER_DECODE_PATTERNS,
    PACKAGE_METADATA_REMOVE_BATCH,
    PACKAGE_METADATA_REMOVE_FUNCTION,
    PACKAGE_METADATA_SAVE_TOKEN_BATCH,
    PACKAGE_METADATA_SAVE_TOKEN_FUNCTION,
    PACKAGE_METADATA_SELECTOR,
    PAGES_TABLE_MERGE_COMMON_TABLE_MODULE,
    Path,
    ROOT,
    RUST_BYTE_SLICE,
    RUST_FUNCTION_DECLARATION,
    RUST_IDENTIFIER,
    _audit_focused_chart_arrangement_owner_source_topology,
    _audit_focused_table_merge_source_topology,
    _audit_iwa_component_removal_codec_source_topology,
    _audit_table_merge_transaction_source_topology,
    _boundary_marker_present,
    _iwa_table_cell_format_engine_code,
    _iwork_public_leak,
    _mask_rust_cfg_test_items,
    _mask_rust_comments,
    _mask_rust_non_code,
    _rust_any_function_body,
    _rust_balanced_delimited_end,
    _rust_canonical_exports,
    _rust_generated_module_include_is_valid,
    _rust_impl_headers,
    _rust_named_struct_body,
    _rust_public_declarations,
    _rust_public_enum_variants,
    _rust_public_methods_in_impl,
    _rust_public_module_body,
    _rust_root_canonical_exports,
    _rust_root_generated_public_leaks,
    _rust_root_level_matches,
    _rust_root_level_module_declarations,
    _rust_statement_end,
    _rust_top_level_function_bodies,
    _table_merge_transaction_has_marker,
    field,
    re,
)


RETIRED_IWA_NUMBERS_IMAGE_ADJUSTMENTS_METHODS = (
    "sheet_image_adjustments",
    "set_sheet_image_adjustments",
)
RETIRED_IWA_NUMBERS_IMAGE_ADJUSTMENTS_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_IMAGE_ADJUSTMENTS_METHODS
)
NUMBERS_NAMES_IMPLEMENTATION_SOURCES = (
    Path("crates/litchi-numbers/src/names.rs"),
    Path("crates/litchi-numbers/src/package/names.rs"),
)
# Keep the publication-route audit narrow to the names owner and any helper
# files split out beneath it.  In particular, do not walk all of
# ``package/**/*.rs``: unrelated package editors may legitimately have their
# own private helpers and should not become part of the names call graph.
NUMBERS_NAMES_PUBLICATION_HELPER_ROOT = (
    Path("crates/litchi-numbers/src/package/names")
)
NUMBERS_NAMES_PUBLICATION_FUNCTION = "rewrite_names"
NUMBERS_NAMES_SAVE_TOKEN_REWRITE = "rewrite_package_metadata_save_tokens"
NUMBERS_NAMES_SAVE_TOKEN_FORBIDDEN_PUBLIC_TYPES = frozenset(
    {
        "Batch",
        "ComponentInfo",
        "ComponentSelector",
        "ComponentDescriptor",
        "DataReferenceOwnerRemoval",
        "DecodeError",
        "DecodeOptions",
        "ExternalReferenceAddition",
        "ExternalReferenceDescriptor",
        "ExternalReferenceRemoval",
        "ObjectUuidAddition",
        "ObjectUuidDescriptor",
        "ObjectUuidRemoval",
        "PackageMetadata",
        "PackageMetadataInspection",
        "PackageMetadataVisitor",
        "PreparedPackageMetadataRewrite",
        "RemovalBatch",
        "RewriteError",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteLimit",
        "RewriteOptions",
        "RewriteOutput",
        "RewriteReport",
        "SaveTokenBatch",
        "SaveTokenSelector",
        "UuidBits",
    }
)
NUMBERS_NAMES_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
)
NUMBERS_NAMES_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
NUMBERS_NAMES_OPTIONAL_TYPES = ("Path", "InvalidReason")
NUMBERS_NAMES_SHORT_NAMES = frozenset(
    NUMBERS_NAMES_CANONICAL_TYPES + NUMBERS_NAMES_OPTIONAL_TYPES
)
NUMBERS_NAMES_FLAT_ALIAS_PREFIXES = (
    "Name",
    "Names",
    "SheetName",
    "TableName",
)
NUMBERS_NAMES_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_NAMES_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_NAMES_SHORT_NAMES
)
NUMBERS_NAMES_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?names[ \t\r\n]*::"
)
PUBLIC_NUMBERS_PACKAGE_NAMES_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?names\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_NAMES_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ComponentCatalog",
        "EntryEdit",
        "RawMessage",
        "Resolved",
        "SnappyStream",
    }
)
NUMBERS_NAMES_WIRE_TYPES = frozenset(
    {"WireDescent", "WireError", "WireLimits", "WireResourceLimit", "WireView"}
)
NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE = (
    Path("crates/litchi-numbers/src/table/headers.rs")
)
NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/table/headers/transaction.rs"
)
NUMBERS_TABLE_HEADER_SETTINGS_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_headers.rs"
)
NUMBERS_TABLE_HEADER_SETTINGS_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE,
    NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE,
    NUMBERS_TABLE_HEADER_SETTINGS_OWNER_SOURCE,
    Path("crates/litchi-numbers/src/package/table_headers/api.rs"),
    Path("crates/litchi-numbers/src/package/table_headers/dependencies.rs"),
    Path("crates/litchi-numbers/src/package/table_headers/error.rs"),
    Path("crates/litchi-numbers/src/package/table_headers/ownership.rs"),
    Path("crates/litchi-numbers/src/package/table_headers/resolve.rs"),
    Path("crates/litchi-numbers/src/package/table_headers/rewrite.rs"),
)
NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/table.rs"),
)
NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
    "InvalidReason",
)
NUMBERS_TABLE_HEADER_SETTINGS_PACKAGE_METHODS = (
    "table_header_settings",
    "edit_table_headers",
    "apply_table_headers",
)
NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_TYPES = (
    "Count",
    "Error",
    "Settings",
)
NUMBERS_TABLE_HEADER_SETTINGS_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES
)
NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIAS_PREFIXES = (
    "HeaderSettings",
    "TableHeader",
    "TableHeaders",
    "TableHeaderSettings",
)
NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_HEADER_SETTINGS_SHORT_NAMES
)
NUMBERS_TABLE_HEADER_SETTINGS_FLAT_SEMANTIC_ALIASES = frozenset(
    {
        "HeaderCount",
        "HeaderSettings",
        "TableHeaderCount",
        "TableHeaderSettings",
    }
)
NUMBERS_TABLE_HEADER_SETTINGS_ROOT_ALIASES = frozenset(
    NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES
    + NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_TYPES
)
NUMBERS_TABLE_HEADER_SETTINGS_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_headers|headers)"
    r"[ \t\r\n]*::"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_HEADERS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_headers\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_HEADERS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?headers\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_HEADER_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_HEADER_SETTINGS_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableHeaderSettingsSnapshot",
        "TableInfoSnapshot",
    }
)
NUMBERS_TABLE_HEADER_SETTINGS_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_WIRE_HELPERS = (
    "read_table_title_settings_wire",
    "write_table_title_settings_wire",
)
RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_WIRE_HELPER_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_WIRE_HELPERS
)
NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE = (
    Path("crates/litchi-numbers/src/table/title.rs")
)
NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE = (
    Path("crates/litchi-numbers/src/package/table_title.rs")
)
NUMBERS_TABLE_TITLE_SETTINGS_OWNER_HELPER_ROOT = (
    Path("crates/litchi-numbers/src/package/table_title")
)
NUMBERS_TABLE_TITLE_SETTINGS_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE,
    NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE,
)
NUMBERS_TABLE_TITLE_SETTINGS_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/table.rs"),
)
NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES = (
    "Settings",
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_TITLE_SETTINGS_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES
)
NUMBERS_TABLE_TITLE_SETTINGS_PACKAGE_METHODS = (
    "table_title_settings",
    "edit_table_title",
    "apply_table_title",
)
NUMBERS_TABLE_TITLE_SETTINGS_FLAT_ALIAS_PREFIXES = (
    "Title",
    "TitleSettings",
    "TableTitle",
    "TableTitleSettings",
)
NUMBERS_TABLE_TITLE_SETTINGS_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_TITLE_SETTINGS_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_TITLE_SETTINGS_SHORT_NAMES
)
NUMBERS_TABLE_TITLE_SETTINGS_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_title|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?title)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_TITLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_title\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_TITLE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_title\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_TITLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?title\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_TITLE_SETTINGS_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableTitleSettingsArchive",
        "TableTitleSettingsSnapshot",
    }
)
NUMBERS_TABLE_TITLE_SETTINGS_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_TITLE_SETTINGS_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tswp"}
)

NUMBERS_TABLE_APPEARANCE_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/table/appearance.rs"
)
NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_appearance.rs"
)
# Source-built Numbers catalogs still enter through the monolithic host.  The
# only raw-ID/borrowed-payload exception is the explicitly hidden seam below;
# keeping its declaration exact prevents a second public/native appearance API
# from being smuggled into the focused crate.
NUMBERS_TABLE_APPEARANCE_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_appearance"
)
NUMBERS_TABLE_APPEARANCE_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/table_appearance_codec.rs"
)
NUMBERS_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_APPEARANCE_CODEC_MODULE = "table_appearance_codec"
NUMBERS_TABLE_APPEARANCE_CODEC_REQUIRED_APIS = (
    "decode_table_model_with_report",
    "prepare_table_model_style_rewrite",
    "canonical_table_style_variation",
    "append_stylesheet_style",
)
NUMBERS_TABLE_APPEARANCE_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_APPEARANCE_SEMANTIC_SOURCE,
    NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE,
)
NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/table.rs"),
)
NUMBERS_TABLE_APPEARANCE_SEMANTIC_TYPES = (
    "Appearance",
    "Banding",
    "GridlineVisibility",
    "Gridlines",
    "RowSizing",
)
NUMBERS_TABLE_APPEARANCE_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_APPEARANCE_CANONICAL_TYPES = (
    *NUMBERS_TABLE_APPEARANCE_SEMANTIC_TYPES,
    *NUMBERS_TABLE_APPEARANCE_TRANSACTION_TYPES,
)
NUMBERS_TABLE_APPEARANCE_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_APPEARANCE_CANONICAL_TYPES
)
NUMBERS_TABLE_APPEARANCE_PACKAGE_METHODS = (
    "table_appearance",
    "edit_table_appearance",
    "apply_table_appearance",
)
NUMBERS_TABLE_APPEARANCE_FLAT_ALIAS_PREFIXES = (
    "Appearance",
    "TableAppearance",
    "TableStyleAppearance",
)
NUMBERS_TABLE_APPEARANCE_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_APPEARANCE_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_APPEARANCE_TRANSACTION_TYPES
)
NUMBERS_TABLE_APPEARANCE_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_appearance|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?appearance)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_APPEARANCE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_appearance\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_APPEARANCE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_appearance\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_APPEARANCE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?appearance\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_APPEARANCE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableAppearanceArchive",
        "TableAppearanceSnapshot",
        "TableStyleArchive",
        "TableStylePresetArchive",
    }
)
NUMBERS_TABLE_APPEARANCE_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_APPEARANCE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tst", "tswp"}
)
NUMBERS_TABLE_APPEARANCE_CODEC_REQUIRED_MARKERS = (
    "canonical",
    "unknown",
    "report",
)
NUMBERS_TABLE_APPEARANCE_OWNER_REQUIRED_MARKERS = {
    "copy-on-write": re.compile(
        r"(?<![A-Za-z0-9])(?:copy_on_write|copy-on-write|style_variation|"
        r"shared_style|Cow)(?![A-Za-z0-9])"
    ),
    "metadata": re.compile(
        r"(?<![A-Za-z0-9])(?:metadata|save_token|object_uuid|"
        r"external_reference|watermark)(?![A-Za-z0-9])"
    ),
    "budget": re.compile(
        r"(?<![A-Za-z0-9])(?:budget|WireBudget|TransactionBudget|"
        r"charge_|preflight)(?![A-Za-z0-9])"
    ),
    "locality": re.compile(
        r"(?<![A-Za-z0-9])(?:locality|verify_locality|candidate|reopen|"
        r"same_content)(?![A-Za-z0-9])"
    ),
}

NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/pop_up_menu.rs"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_pop_up_menu.rs"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_pop_up_menu"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_MODULE = (
    "numbers_table_cell_pop_up_menu_codec"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_REQUIRED_APIS = (
    "decode_popup_menu_model_with_report",
    "decode_cell_spec_with_report",
    "prepare_popup_menu_model_write",
    "prepare_cell_spec_write",
    "PreparedPopUpMenuModelWrite",
    "PreparedCellSpecWrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "canonical_popup_menu_model",
    "canonical_cell_spec",
)
NUMBERS_TABLE_CELL_POP_UP_MENU_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_SOURCE,
    NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE,
)
NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/cell/data_format.rs"),
    NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_SOURCE,
)
NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_TYPES = (
    "Item",
    "InitialSelection",
    "PopUpMenu",
    "Error",
)
NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_POP_UP_MENU_PACKAGE_METHODS = (
    "table_cell_pop_up_menu_format",
    "edit_table_cell_pop_up_menu_format",
    "apply_table_cell_pop_up_menu_format",
)
NUMBERS_TABLE_CELL_POP_UP_MENU_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_POP_UP_MENU_FLAT_ALIAS_PREFIXES = (
    "PopUpMenu",
    "TableCellPopUpMenu",
    "CellPopUpMenu",
    "PopUpMenuFormat",
)
NUMBERS_TABLE_CELL_POP_UP_MENU_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_CELL_POP_UP_MENU_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_pop_up_menu|cell[ \t\r\n]*::"
    r"[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?pop_up_menu)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_POP_UP_MENU_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_pop_up_menu\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_POP_UP_MENU_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_pop_up_menu\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_CELL_POP_UP_MENU_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?pop_up_menu\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_POP_UP_MENU_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "BncCell",
        "CellValueArchive",
        "ComponentCatalog",
        "ControlCellSpec",
        "ExactArtifacts",
        "FormatStructArchive",
        "HostPackage",
        "PhysicalSource",
        "PopUpMenuModel",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableDataList",
    }
)
NUMBERS_TABLE_CELL_POP_UP_MENU_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_CELL_POP_UP_MENU_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tsce", "tst", "tswp"}
)
NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_REQUIRED_MARKERS = (
    "canonical",
    "unknown",
    "report",
    "nil",
    "deprecated",
    "refcount",
)
NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_REQUIRED_MARKERS = {
    "control-cell": re.compile(
        r"(?<![A-Za-z0-9])(?:control[_-]?cell|ControlCellSpec)(?![A-Za-z0-9])"
    ),
    "table-data-list": re.compile(
        r"(?<![A-Za-z0-9])(?:table[_-]?data[_-]?list|TableDataList|data[_-]?list)(?![A-Za-z0-9])"
    ),
    "refcount": re.compile(
        r"(?<![A-Za-z0-9])(?:refcount|ref_count|reference_count)(?![A-Za-z0-9])"
    ),
    "copy-on-write": re.compile(
        r"(?<![A-Za-z0-9])(?:copy_on_write|copy-on-write|shared|cull|cow)(?![A-Za-z0-9])"
    ),
    "metadata": re.compile(
        r"(?<![A-Za-z0-9])(?:metadata|save_token|object_uuid|external_reference|watermark)(?![A-Za-z0-9])"
    ),
    "budget": re.compile(
        r"(?<![A-Za-z0-9])(?:budget|WireBudget|TransactionBudget|charge_|preflight)(?![A-Za-z0-9])"
    ),
    "locality": re.compile(
        r"(?<![A-Za-z0-9])(?:locality|verify_locality|candidate|reopen|same_content)(?![A-Za-z0-9])"
    ),
    "string-table": re.compile(
        r"(?<![A-Za-z0-9])(?:string[_-]?table|StringTable|InitialSelection)(?![A-Za-z0-9])"
    ),
}
NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_pop_up_menu_codec.rs"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_pop_up_menu_codec"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_pop_up_menu.rs"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_pop_up_menu"
)
NUMBERS_TABLE_CELL_POP_UP_MENU_RAW_PARAMETER = re.compile(
    r"\b(?:table_id|row|column|object_id|native_id|model_id|"
    r"control_spec_id|popup_model_id|source_bytes|bytes)\b[ \t\r\n]*:\s*"
    r"(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])\b"
)


NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/control.rs"
)
NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_control.rs"
)
NUMBERS_TABLE_CELL_CONTROL_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_control"
)
NUMBERS_TABLE_CELL_CONTROL_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_control_codec.rs"
)
NUMBERS_TABLE_CELL_CONTROL_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_CELL_CONTROL_CODEC_MODULE = "numbers_table_cell_control_codec"
NUMBERS_TABLE_CELL_CONTROL_CODEC_REQUIRED_APIS = (
    "decode_cell_spec_with_report",
    "prepare_cell_spec_write",
    "PreparedCellSpecWrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "canonical_cell_spec",
    "rewrite_cell_spec",
)
NUMBERS_TABLE_CELL_CONTROL_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE,
    NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE,
)
NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/cell/data_format.rs"),
    NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE,
)
NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_TYPES = (
    "CellControl",
    "Checkbox",
    "StarRating",
    "Slider",
    "Stepper",
    "PopUpMenu",
    "Range",
    "DisplayFormat",
)
NUMBERS_TABLE_CELL_CONTROL_VARIANTS = (
    "Checkbox",
    "StarRating",
    "Slider",
    "Stepper",
    "PopUpMenu",
)
NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_CONTROL_PACKAGE_METHODS = (
    "table_cell_control_format",
    "edit_table_cell_control_format",
    "apply_table_cell_control_format",
)
NUMBERS_TABLE_CELL_CONTROL_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_CONTROL_FLAT_ALIAS_PREFIXES = (
    "CellControl",
    "TableCellControl",
    "CellControlFormat",
    "TableCellControlFormat",
)
NUMBERS_TABLE_CELL_CONTROL_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_CELL_CONTROL_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_CONTROL_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_control|cell[ \t\r\n]*::"
    r"[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?control)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_CONTROL_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_control\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_CONTROL_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_control\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_CELL_CONTROL_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?control\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_CONTROL_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "BncCell",
        "CellSpecArchive",
        "CellValueArchive",
        "ComponentCatalog",
        "ControlCellSpec",
        "FormatStructArchive",
        "HostPackage",
        "PhysicalSource",
        "PopUpMenuModel",
        "RawMessage",
        "SourceCatalog",
        "StringCellValueArchive",
        "TableDataList",
    }
)
NUMBERS_TABLE_CELL_CONTROL_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_CELL_CONTROL_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "bnc", "tsce", "tst", "tsp"}
)
NUMBERS_TABLE_CELL_CONTROL_CODEC_REQUIRED_MARKERS = (
    "canonical",
    "unknown",
    "report",
    "prepared",
    "fixed64",
    "interaction",
    "refcount",
)
NUMBERS_TABLE_CELL_CONTROL_OWNER_REQUIRED_MARKERS = {
    "control-cell": re.compile(
        r"(?<![A-Za-z0-9])(?:control[_-]?cell|CellControl)(?![A-Za-z0-9])"
    ),
    "table-data-list": re.compile(
        r"(?<![A-Za-z0-9])(?:table[_-]?data[_-]?list|TableDataList|data[_-]?list)(?![A-Za-z0-9])"
    ),
    "refcount": re.compile(
        r"(?<![A-Za-z0-9])(?:refcount|ref_count|reference_count)(?![A-Za-z0-9])"
    ),
    "copy-on-write": re.compile(
        r"(?<![A-Za-z0-9])(?:copy_on_write|copy-on-write|shared|cull|cow)(?![A-Za-z0-9])"
    ),
    "metadata": re.compile(
        r"(?<![A-Za-z0-9])(?:metadata|save_token|object_uuid|external_reference|watermark)(?![A-Za-z0-9])"
    ),
    "budget": re.compile(
        r"(?<![A-Za-z0-9])(?:budget|WireBudget|TransactionBudget|charge_|preflight)(?![A-Za-z0-9])"
    ),
    "resource": re.compile(
        r"(?<![A-Za-z0-9])(?:resource|allocation|retained|max_minus_one)(?![A-Za-z0-9])"
    ),
    "prepared": re.compile(
        r"(?<![A-Za-z0-9])(?:prepared|execution_requirements|scratch)(?![A-Za-z0-9])"
    ),
    "locality": re.compile(
        r"(?<![A-Za-z0-9])(?:locality|verify_locality|candidate|reopen|same_content)(?![A-Za-z0-9])"
    ),
}
# The bounded split-graph contract is a cross-component slice. Reads and exact
# no-ops may traverse a split graph, but every changed operation must refuse
# before candidate/native allocation unless the popup lifecycle contract below
# owns multi-member mutation. Require strict read/edge/locator proof and the
# early refusal guard; do not require successful split writes, token batching,
# or cross-component candidate/locality machinery.
NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_control.rs"
)
NUMBERS_TABLE_CELL_CONTROL_BOUNDED_OWNER_REQUIRED_MARKERS = {
    "strict cross-component read/no-op": re.compile(
        r"(?<![A-Za-z0-9_])(?:prove[_-]?cross[_-]?component[_-]?reference|"
        r"read[_-]?popup[_-]?with[_-]?(?:policy|budget))(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "pre-candidate changed-operation refusal": re.compile(
        r"(?<![A-Za-z0-9_])reject[_-]?cross[_-]?component[_-]?write(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "refusal precedes candidate allocation": re.compile(
        r"(?is)reject[_-]?cross[_-]?component[_-]?write\s*\([^;]{0,500}\)"
        r"\s*(?:\?\s*)?;[\s\S]{0,1800}(?:let\s+mut\s+budget|"
        r"physical[_-]?source|preflight|rewrite[_-]?(?:native|scalar)|"
        r"prepare[_-]?reassembly)"
    ),
    "effective locator coverage": re.compile(
        r"(?<![A-Za-z0-9_])(?:effective[_-]?locator|effective[_-]?location|"
        r"resolved[_-]?location)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "current external-edge coverage": re.compile(
        r"(?<![A-Za-z0-9_])(?:require[_-]?external[_-]?edge|"
        r"prove[_-]?cross[_-]?component[_-]?reference|external[_-]?edge)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "object/component edge scope": re.compile(
        r"(?<![A-Za-z0-9_])(?:object[_-]?identifier|component[_-]?identifier|"
        r"Option\s*<\s*u64\s*>)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "target alias rejection": re.compile(
        r"(?<![A-Za-z0-9_])(?:has[_-]?physical[_-]?alias|physical[_-]?alias)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}
# The popup lifecycle contract completes the split Pop-Up Menu graph under the
# generic CellControl facade.  The bounded split-graph contract deliberately
# required a refusal guard; that guard must disappear once the owner can stage
# every changed member in a single transaction.  Keep this contract separate
# from the scalar-control bridge so Pages/Keynote adapters and the shared
# private cell-data-format implementation remain out of scope.
NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_OWNER_REQUIRED_MARKERS = {
    "per-member graph ownership": re.compile(
        r"(?<![A-Za-z0-9_])(?:NativePopUpGraphInput|NativePopUpMember|"
        r"NativePopUpObjectRoute|NativeMemberEdit|member_edits)(?![A-Za-z0-9_])"
    ),
    "multi-member native transition": re.compile(
        r"(?<![A-Za-z0-9_])rewrite_native_popup_menu_multi(?![A-Za-z0-9_])"
    ),
    "member edit deduplication": re.compile(
        r"(?<![A-Za-z0-9_])(?:NativeControlOutput|from_edits|"
        r"deduplicat(?:e|ed|ion)|conflicting bytes)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "component UUID ownership": re.compile(
        r"(?<![A-Za-z0-9_])(?:require_current_uuid|uuid_(?:addition|removal)|"
        r"IdentifierCensus|reserved[_-]?identifier)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "external edge and save-token batch": re.compile(
        r"(?<![A-Za-z0-9_])(?:require_external_edge|"
        r"prove_cross_component_reference|external[_-]?reference|"
        r"selectors_for_components|SaveTokenBatch|save_tokens|"
        r"touched_components)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "COW reuse cull and refcount census": re.compile(
        r"(?<![A-Za-z0-9_])(?:copy[_-]?on[_-]?write|"
        r"prepare_desired_state|existing_popup_model_identifiers|"
        r"archive_has_popup_reference|census_bnc_references|"
        r"validate_bnc_refcounts|ref[_-]?count|cull)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "aggregate and FieldInfo transition": re.compile(
        r"(?<![A-Za-z0-9_])(?:ArchiveInfo|FieldInfo|object_references|"
        r"replace_(?:control|message).*transition|aggregate)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "aggregate transaction budget": re.compile(
        r"(?<![A-Za-z0-9_])(?:TransactionBudget|WireBudget|charge_|"
        r"residual_|preflight)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "prepared archive and ZIP reassembly": re.compile(
        r"(?<![A-Za-z0-9_])(?:prepare[_-]?reassembly|"
        r"execution[_-]?requirements|Prepared|prepared|reassembly)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "candidate reopen and object locality": re.compile(
        r"(?<![A-Za-z0-9_])(?:candidate|reopen|verify[_-]?package[_-]?locality|"
        r"same[_-]?content)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}
NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_FORBIDDEN_MARKERS = {
    "cross-component write refusal": re.compile(
        r"(?<![A-Za-z0-9_])reject[_-]?cross[_-]?component[_-]?write(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}
NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_INTEGRATION_MARKERS = {
    "split PopUpMenu read/no-op": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,12000}"
        r"pop[_-]?up[\s\S]{0,5000}(?:read|no[_-]?op)"
    ),
    "split PopUpMenu replacement/COW": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,12000}"
        r"pop[_-]?up[\s\S]{0,5000}(?:replace|replacement|copy[_-]?on[_-]?write|cow)"
    ),
    "split PopUpMenu create/reuse/cull/refcount": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,16000}"
        r"pop[_-]?up[\s\S]{0,8000}(?:create|reuse|cull|ref[_-]?count)"
    ),
    "split PopUpMenu inverse/apply/conflict": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,16000}"
        r"pop[_-]?up[\s\S]{0,8000}(?:inverse|apply|conflict)"
    ),
    "split PopUpMenu member locality": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,16000}"
        r"pop[_-]?up[\s\S]{0,8000}(?:member|locality|exact[_-]?bytes)"
    ),
    "split PopUpMenu metadata ownership": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,16000}"
        r"pop[_-]?up[\s\S]{0,8000}(?:metadata|uuid|save[_-]?token|external)"
    ),
    "split PopUpMenu hostile ownership atomicity": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,16000}"
        r"pop[_-]?up[\s\S]{0,8000}(?:alias|opaque|versioned|atomic|reject)"
    ),
    "split PopUpMenu resource/reassembly": re.compile(
        r"(?is)(?:split|cross[_-]?component)[\s\S]{0,16000}"
        r"pop[_-]?up[\s\S]{0,8000}(?:budget|limit|reassembl|allocation|scratch)"
    ),
}
NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_FUZZ_MARKERS = {
    "package fuzz split lifecycle": re.compile(
        r"(?is)(?:split|member)[\s\S]{0,10000}(?:popup|pop[_-]?up)[\s\S]{0,10000}"
        r"(?:create|clear|reset|inverse|conflict|atomic)"
    ),
    "codec fuzz prepared replay": re.compile(
        r"(?is)(?:prepare|execution[_-]?requirements|replay)[\s\S]{0,10000}"
        r"(?:unknown|limit|depth|allocation|scratch)"
    ),
}
NUMBERS_TABLE_CELL_CONTROL_SPLIT_INTEGRATION_MARKERS = {
    "split-component read coverage": re.compile(
        r"\bfn\s+split_components?[_a-z0-9]*read[_a-z0-9]*\s*\(", re.IGNORECASE
    ),
    "split-component no-op exact bytes": re.compile(
        r"(?is)(?:split_components?|assert_split_read_noop_and_write_reject)"
        r"[\s\S]{0,5000}(?:is_noop|no_op)[\s\S]{0,2000}exact_bytes"
    ),
    "split-component changed refusal": re.compile(
        r"(?is)(?:split_components?|assert_split_read_noop_and_write_reject)"
        r"[\s\S]{0,6000}(?:write_reject;|changed_edit_rejects|\.is_err\(\))"
    ),
    "split-component alias/edge atomicity": re.compile(
        r"(?is)split_components?[\s\S]{0,8000}(?:assert_changed_edit_rejects|"
        r"assert_split_owner_rejects|\.is_err\(\)|external;|opaque;|atomic;)"
    ),
}

# The existing-cell decimal Number format is the next focused Numbers
# transaction.  Keep this boundary deliberately narrower than the generic
# ``DataFormat``/interactive-control routes: only ``Option<Number>`` and its
# explicit-to-automatic transition are admitted here.  The package owner
# resolves selectors and owns the exact-source transaction; native BNC cells,
# format-table/refcount records, archive members, generated messages, and
# Buffa views remain below this boundary.
NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/number.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/number/transaction.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_number_format.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_number_format"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_number_format_codec.rs"
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_display_format_codec.rs"
)
# The shared display-format core lives in `litchi-iwa-table`;
# the Numbers facades reach it through the `table_cell_display_format_native`
# alias of `package.rs`.
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_NATIVE_SOURCE = Path(
    "crates/litchi-iwa-table/src/cell_format/display.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_number_format_codec"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_number_format_generated"
)
# Number-format decoding/rewrite intentionally reuses the strict control-cell
# codec's bounded wire walker.  The small public wrapper may re-export the
# focused APIs while the implementation and its Buffa projection live in the
# existing private source.
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE,
    Path("crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"),
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULE,
    "buffa_numbers_table_cell_pop_up_menu_generated",
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE,
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_SOURCE,
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE,
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/cell/data_format.rs"),
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE,
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_TYPES = ("Number",)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_PACKAGE_METHODS = (
    "table_cell_number_format",
    "edit_table_cell_number_format",
    "apply_table_cell_number_format",
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_EDIT_METHODS = (
    "before",
    "after",
    "set",
    "clear",
    "commit",
)
# These aliases were used by earlier raw-ID/data-format drafts.  They must
# not become a second public owner when the canonical nested transaction lands.
NUMBERS_TABLE_CELL_NUMBER_FORMAT_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in (
        "NumberFormat",
        "TableCellNumberFormat",
        "CellNumberFormat",
        "NumbersTableCellNumberFormat",
    )
    for suffix in NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_FLAT_ALIAS_PATTERN = re.compile(
    r"(?:Number|TableCellNumber|CellNumber|NumbersTableCellNumber)"
    r"(?:Format)?(?:Edit|Patch|Commit|Diagnostics|Error|LimitKind|Path)\b"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_number_format|cell[ \t\r\n]*::"
    r"[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*(?:r#)?number)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_CELL_DATA_FORMAT_NUMBER_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?number\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_CELL_NUMBER_FORMAT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_NUMBER_FORMAT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_number_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_NUMBER_FORMAT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_number_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "BncCell",
        "BncCellView",
        "CellValueArchive",
        "ComponentCatalog",
        "ControlCellSpec",
        "EntryEdit",
        "ExactArtifacts",
        "FormatEntry",
        "FormatStructArchive",
        "FormatTableArchive",
        "HostPackage",
        "NumberFormatArchive",
        "NumberFormatSnapshot",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableDataList",
    }
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "bnc", "tn", "tsp", "tst", "tswp"}
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"table_id|format_id|format_key|registry_id|entry_id|component_id|member_id|"
    r"archive_id|message_id|uuid|source_bytes|bytes|row|column)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:source_bytes|raw_bytes|wire_bytes|payload|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:Vec\s*<\s*u8\s*>|Box\s*<\s*\[\s*u8\s*\]\s*>|"
    r"(?:std\s*::\s*borrow\s*::\s*)?Cow\s*<[^>]*\[\s*u8\s*\]|"
    r"(?:std\s*::\s*sync\s*::\s*)?Arc\s*<\s*\[\s*u8\s*\]\s*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE = re.compile(
    r"(?<![A-Za-z0-9_])(?:(?:bytes\s*::\s*)?Bytes(?:Mut)?\b|"
    r"Vec\s*<\s*u8\s*>|Box\s*<\s*\[\s*u8\s*\]\s*>|"
    r"(?:std\s*::\s*borrow\s*::\s*)?Cow\s*<[^>]*\[\s*u8\s*\]|"
    r"(?:std\s*::\s*(?:sync|rc)\s*::\s*)?(?:Arc|Rc)\s*<\s*\[\s*u8\s*\]\s*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|table|model|native|component|archive|message|"
    r"format|registry|entry|resource|uuid)[A-Za-z_]*(?:id|identifier|key))"
    r"[ \t\r\n]*:[ \t\r\n]*u64\b"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|key|"
    r"(?:cell|object|native|model|table|format|registry|entry|component|archive|"
    r"message|storage|resource)_(?:id|identifier|key))"
    r"[ \t\r\n]*:[ \t\r\n]*"
    r"(?:[A-Z][A-Za-z0-9_]*(?:Id|ID|Key|Identifier)|"
    r"(?:u64|u32|usize|i64|i32))\b"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_HELPER = re.compile(
    r"(?<![A-Za-z0-9_])(?:cell_number_format|set_cell_number_format|"
    r"reset_cell_number_format|read_native_number_format|set_native_number_format|"
    r"read_number_format_wire|write_number_format_wire|raw_number_format|"
    r"patch_number_format|parse_wire_fields|patch_varint_field|"
    r"patch_length_delimited_field|FormatStructArchive)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_REQUIRED_APIS = (
    ("decode_number_format_with_report", "decode_cell_number_format_with_report"),
    ("decode_cell_spec_with_report", "decode_control_cell_spec_with_report"),
    ("prepare_number_format_rewrite", "prepare_cell_number_format_rewrite"),
    ("PreparedNumberFormatRewrite", "PreparedCellNumberFormatRewrite"),
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    ("canonical_number_format", "canonical_cell_number_format"),
    ("rewrite_number_format", "rewrite_cell_number_format"),
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_REQUIRED_MARKER_GROUPS = {
    "lazy Buffa projection": (
        NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULES,
        ("decode_lazy_view", "decode_view", "LazyView", "visitor"),
        "buffa",
    ),
    "strict canonical ingress": (
        ("preflight", "scan_number_message", "scan_decimal_message", "scan_display_format"),
        ("duplicate", "is_some", "wrapper_seen"),
        ("noncanonical", "non-canonical", "varint_canonical"),
    ),
    "unknown/raw preservation": ("unknown", "raw", "extend_from_slice"),
    "bounded prepared rewrite": (
        ("MAX_RECURSION_LIMIT", "MAX_RECURSION"),
        "execution_requirements",
        "execute",
    ),
    "number format fields": (
        ("decimal_places", "decimal"),
        ("negative_style", "negative"),
        ("thousands_separator", "thousands"),
    ),
    "registry/refcount closure": (
        ("format_table", "registry", "refcount", "ref_count", "references"),
        ("reference", "entry"),
    ),
}
NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_REQUIRED_MARKERS = {
    "existing-cell resolution": re.compile(
        r"(?<![A-Za-z0-9])(?:existing[_-]?cell|existing|resolve|locator)(?![A-Za-z0-9])",
        re.IGNORECASE,
    ),
    "format-table/refcount closure": re.compile(
        r"(?<![A-Za-z0-9])(?:format[_-]?table|registry|refcount|ref_count|reference|"
        r"WireReferenceBytes|PayloadReferences)(?![A-Za-z0-9])",
        re.IGNORECASE,
    ),
    "exact-source transaction": re.compile(
        r"(?<![A-Za-z0-9])(?:exact[_-]?source|source[_-]?bound|PatchConflict|inverse|no[_-]?op)(?![A-Za-z0-9])",
        re.IGNORECASE,
    ),
    "candidate verification/locality": re.compile(
        r"(?<![A-Za-z0-9])(?:candidate|reopen|readback|locality|same[_-]?content)(?![A-Za-z0-9])",
        re.IGNORECASE,
    ),
    "bounded staging": re.compile(
        r"(?<![A-Za-z0-9])(?:budget|limit|preflight|bounded|allocation)(?![A-Za-z0-9])",
        re.IGNORECASE,
    ),
    "fail-closed unsupported graph": re.compile(
        r"(?<![A-Za-z0-9])(?:unsupported|UnsupportedDependency|UnsupportedSource|"
        r"ambiguous|cross[_-]?component|fail[_-]?closed)(?![A-Za-z0-9])",
        re.IGNORECASE,
    ),
}
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_HIDDEN_MODULE = re.compile(
    rf"(?m)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
    rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_MODULE)}\b"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_GENERATED_MODULE = re.compile(
    rf"(?m)^\s*pub\s+mod\s+"
    rf"{re.escape(NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULE)}\b"
)

RETIRED_IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_HELPERS = (
    "cell_number_format",
    "set_cell_number_format",
    "reset_cell_number_format",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_HELPER_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_HELPERS
)
RETIRED_IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_TESTS = (
    "source_built_table_roundtrips_and_reuses_number_formats",
    "invalid_number_format_coordinate_is_transactional",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_TEST_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_TESTS
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXAMPLE = Path(
    "crates/litchi-numbers/examples/edit_table_cell_number_format.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_INTEGRATION_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_number_format.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_number_format_codec.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_number_format_codec"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_number_format.rs"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_number_format"
)
NUMBERS_TABLE_CELL_NUMBER_FORMAT_INTEGRATION_MARKERS = {
    "explicit and automatic readback": re.compile(
        r"(?:Option|automatic|explicit|read).*Number", re.IGNORECASE
    ),
    "set/reset transaction": re.compile(
        r"(?:edit_table_cell_number_format|set\(|clear\(|reset|automatic)",
        re.IGNORECASE,
    ),
    "exact no-op and inverse": re.compile(
        r"(?:no[_ -]?op|unchanged|inverse|byte[_ -]?for[_ -]?byte)",
        re.IGNORECASE,
    ),
    "conflict and fail-closed refusal": re.compile(
        r"(?:stale|foreign|ambiguous|conflict|unsupported|malformed)",
        re.IGNORECASE,
    ),
    "selector/locality verification": re.compile(
        r"(?:SheetSelector|TableSelector|(?-i:\bAt\b)|locality|reopen|readback)",
        re.IGNORECASE,
    ),
}

# Plain Text is the next focused Numbers display-format owner.  Keep its
# boundary separate from the generic ``DataFormat`` compatibility surface:
# the semantic value is a zero-sized, archive-free marker while the package
# owner is responsible for resolving selectors, preserving the exact native
# format-table graph, and publishing a checked candidate.  The explicit
# fallback is useful while the marker type is moved out of ``data_format.rs``
# into its dedicated leaf.
NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/text.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_FALLBACK_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/text/transaction.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_text_format.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_text_format"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_text_format_codec.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_PUBLIC_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_text_format_codec"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_text_format_generated"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_GENERATED_MODULES = (
    NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_GENERATED_MODULE,
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE,
    Path("crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"),
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_SOURCE,
    NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_SOURCE,
    NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_SOURCE,
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/cell/data_format.rs"),
    NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_SOURCE,
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_TYPES = ("Text",)
NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_PACKAGE_METHODS = (
    "table_cell_text_format",
    "edit_table_cell_text_format",
    "apply_table_cell_text_format",
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_EDIT_METHODS = (
    "before",
    "after",
    "set",
    "clear",
    "reset",
    "commit",
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in (
        "TextFormat",
        "TableCellTextFormat",
        "CellTextFormat",
        "NumbersTableCellTextFormat",
    )
    for suffix in NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_FLAT_ALIAS_PATTERN = re.compile(
    r"(?:Text|TableCellText|CellText|NumbersTableCellText)"
    r"(?:Format)?(?:Edit|Patch|Commit|Diagnostics|Error|LimitKind|Path)\b"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_text_format|"
    r"cell[ \t\r\n]*::[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?text)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_CELL_DATA_FORMAT_TEXT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?text\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_CELL_TEXT_FORMAT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_TEXT_FORMAT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_text_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_TEXT_FORMAT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_text_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_PHYSICAL_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PHYSICAL_TYPES
    | frozenset(
        {
            "TextFormatArchive",
            "TextFormatSnapshot",
            "TextFormatWrite",
            "DisplayFormatArchive",
            "DisplayFormatSnapshot",
            "DisplayFormatWrite",
            "FormatSnapshot",
            "FormatWrite",
        }
    )
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_WIRE_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_WIRE_TYPES
    | frozenset(
        {
            "TextFormatDecodeError",
            "TextFormatDecodeReport",
            "TextFormatWireView",
            "DisplayFormatDecodeError",
            "DisplayFormatDecodeReport",
            "DisplayFormatWireView",
        }
    )
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_PROTO_ORIGINS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PROTO_ORIGINS
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_PUBLIC_RAW_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_PUBLIC_RAW_CONTAINER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_TYPED_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_RAW_HELPER = re.compile(
    r"(?<![A-Za-z0-9_])(?:cell_(?:number|percentage|currency|scientific|"
    r"fraction|text)_format|set_cell_(?:number|percentage|currency|scientific|"
    r"fraction|text)_format|reset_cell_(?:number|percentage|currency|scientific|"
    r"fraction|text)_format|read_native_(?:number|percentage|currency|scientific|"
    r"fraction|text|display)_format|set_native_(?:number|percentage|currency|"
    r"scientific|fraction|text|display)_format|read_(?:number|percentage|currency|"
    r"scientific|fraction|text|display)_format_wire|write_(?:number|percentage|"
    r"currency|scientific|fraction|text|display)_format_wire|raw_(?:number|"
    r"percentage|currency|scientific|fraction|text|display)_format|patch_(?:number|"
    r"percentage|currency|scientific|fraction|text|display)_format|parse_wire_fields|"
    r"patch_varint_field|patch_length_delimited_field|FormatStructArchive)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_REQUIRED_APIS = (
    (
        "decode_text_format_with_report",
        "decode_cell_text_format_with_report",
        "decode_display_format_with_report",
    ),
    (
        "prepare_text_format_rewrite",
        "prepare_cell_text_format_rewrite",
        "prepare_display_format_rewrite",
    ),
    ("TextFormatSnapshot", "DisplayFormatSnapshot"),
    ("TextFormatWrite", "DisplayFormatWrite"),
    ("PreparedTextFormatRewrite", "PreparedDisplayFormatRewrite"),
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    ("canonical_text_format", "canonical_display_format"),
    ("rewrite_text_format", "rewrite_display_format"),
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_REQUIRED_MARKER_GROUPS = {
    "lazy Buffa projection": (
        NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_GENERATED_MODULES,
        ("decode_lazy_view", "decode_view", "LazyView", "visitor"),
        "buffa",
    ),
    "strict canonical ingress": (
        (
            "preflight",
            "scan_text_format",
            "scan_text_message",
            "scan_format_message",
            "scan_display_format",
            "Budget::new",
        ),
        ("duplicate", "is_some", "wrapper_seen"),
        ("noncanonical", "non-canonical", "varint_canonical"),
    ),
    "unknown/raw preservation": ("unknown", "raw", "extend_from_slice"),
    "bounded prepared rewrite": (
        ("MAX_RECURSION_LIMIT", "MAX_RECURSION"),
        "execution_requirements",
        "execute",
    ),
    "text format family": (
        ("NATIVE_TEXT_FORMAT_TYPE", "TEXT_FORMAT_TYPE", "text_format_type"),
        ("format_type", "text_format", "canonical_text"),
        ("validate_text", "plain_text", "Text"),
    ),
    "registry/refcount closure": (
        ("format_table", "registry", "refcount", "ref_count", "references"),
        ("reference", "entry"),
    ),
}
NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_REQUIRED_MARKERS = {
    **NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_REQUIRED_MARKERS,
    "text-format ownership": re.compile(
        r"(?<![A-Za-z0-9_])(?:text[_ -]?format|plain[_ -]?text|"
        r"NATIVE_TEXT_FORMAT_TYPE|Text)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}
RETIRED_IWA_NUMBERS_TABLE_CELL_TEXT_FORMAT_METHODS = (
    "table_cell_text_format",
    "set_table_cell_text_format",
    "reset_table_cell_text_format",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_TEXT_FORMAT_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_TEXT_FORMAT_METHODS
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_EXAMPLE = Path(
    "crates/litchi-numbers/examples/edit_table_cell_text_format.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_INTEGRATION_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_text_format.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_text_format_codec.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_text_format_codec"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_text_format.rs"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_text_format"
)
NUMBERS_TABLE_CELL_TEXT_FORMAT_INTEGRATION_MARKERS = {
    "explicit and automatic readback": re.compile(
        r"(?:Option|automatic|explicit|read).*Text", re.IGNORECASE
    ),
    "set/reset transaction": re.compile(
        r"(?:edit_table_cell_text_format|set\(|clear\(|reset|automatic)",
        re.IGNORECASE,
    ),
    "exact no-op and inverse": re.compile(
        r"(?:no[_ -]?op|unchanged|inverse|byte[_ -]?for[_ -]?byte)",
        re.IGNORECASE,
    ),
    "conflict and fail-closed refusal": re.compile(
        r"(?:stale|foreign|ambiguous|conflict|unsupported|malformed)",
        re.IGNORECASE,
    ),
    "selector/locality verification": re.compile(
        r"(?:SheetSelector|TableSelector|(?-i:\bAt\b)|locality|reopen|readback)",
        re.IGNORECASE,
    ),
}
IWA_NUMBERS_TABLE_CELL_TEXT_FORMAT_FOCUSED_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:litchi_numbers[ \t\r\n]*::[ \t\r\n]*Package|"
    r"FocusedNumbersPackage|focused_text_format_package|"
    r"focused_number_format_package|focused_numbers_package)[^;{}]*\.[ \t\r\n]*"
    r"(?:table_cell_text_format|edit_table_cell_text_format|"
    r"apply_table_cell_text_format)\s*\("
)
IWA_NUMBERS_TABLE_CELL_TEXT_FORMAT_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?P<method>table_cell_text_format|"
    r"set_table_cell_text_format|reset_table_cell_text_format)"
    r"(?![A-Za-z0-9_])[ \t\r\n]*\("
)

# Date & Time is a focused Numbers display-format owner. Keep this ratchet
# deliberately topology-oriented: DateTime's host compatibility methods are
# not retired here, but its archive-free semantic value, private transaction
# owner, strict Buffa/lazy codec seam, and generated-module visibility must not
# regress as the shared display-format core evolves.
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/date_time.rs"
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_date_time_format.rs"
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_date_time_format_codec.rs"
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_date_time_format_codec"
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_date_time_format_generated"
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE,
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_REQUIRED_APIS = (
    "DateTimeFormatSnapshot",
    "DateTimeFormatWrite",
    "PreparedDateTimeFormatRewrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "decode_date_time_format",
    "decode_date_time_format_with_report",
    "prepare_date_time_format_rewrite",
    "prepare_date_time_format_write",
    "canonical_date_time_format",
    "rewrite_date_time_format",
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_OWNER_METHODS = (
    "table_cell_date_time_format",
    "edit_table_cell_date_time_format",
    "apply_table_cell_date_time_format",
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_RAW_MARKERS = frozenset(
    {
        "RawMessage",
        "HostPackage",
        "SourceCatalog",
        "ObjectId",
        "ObjectID",
        "ModelId",
        "ModelID",
        "Identifier",
        "UUID",
        "Uuid",
    }
)

# The focused Numbers package now owns the selector-first Date & Time
# transaction.  Keep the migration-host retirement inventory deliberately
# narrow: the generic ``DataFormat`` helpers (``cell_date_time_format`` and
# friends) remain a compatibility surface for source-built packages and for
# the Pages/Keynote attached-table adapters.  Only these Draft
# table-cell methods are retired once the focused owner is present.
RETIRED_IWA_NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_METHODS = (
    "table_cell_date_time_format",
    "set_table_cell_date_time_format",
    "reset_table_cell_date_time_format",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_METHODS
)
IWA_NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?P<method>table_cell_date_time_format|"
    r"set_table_cell_date_time_format|reset_table_cell_date_time_format)"
    r"(?![A-Za-z0-9_])[ \t\r\n]*\("
)

# Duration is the next focused Numbers display-format owner.  Keep the
# contract separate from Date & Time even though both formats use the shared
# source-preserving display-format core: Duration has a typed style/unit
# value, a native type-268 envelope, and a dedicated migration-host route that
# must disappear once the selector-first owner is active.  The semantic leaf
# and transaction exports remain archive-free; native IDs, raw payloads, and
# generated views stay below the package/codec boundary.
NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/duration.rs"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/duration/transaction.rs"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_duration_format.rs"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_duration_format_codec.rs"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_duration_format_codec"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_duration_format_generated"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE,
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_REQUIRED_APIS = (
    "DurationStyle",
    "DurationUnit",
    "DurationFormatSnapshot",
    "DurationFormatWrite",
    "PreparedDurationFormatRewrite",
    "PreparedDurationFormatWrite",
    "RewriteOutput",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "decode_duration_format",
    "decode_duration_format_with_report",
    "prepare_duration_format_rewrite",
    "prepare_duration_format_write",
    "canonical_duration_format",
    "rewrite_duration_format",
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_OWNER_METHODS = (
    "table_cell_duration_format",
    "edit_table_cell_duration_format",
    "apply_table_cell_duration_format",
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_TYPES = (
    "Duration",
    "Style",
    "Unit",
    "UnitRange",
    "Units",
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_RAW_MARKERS = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "HostPackage",
        "RawMessage",
        "SourceCatalog",
        "ObjectId",
        "ObjectID",
        "ModelId",
        "ModelID",
        "Identifier",
        "UUID",
        "Uuid",
        "DurationFormatArchive",
        "DurationFormatSnapshot",
        "DurationFormatWrite",
        "PreparedDurationFormatRewrite",
    }
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "bnc", "tn", "tsp", "tst", "tswp"}
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "FormatStructArchive",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
    }
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_PUBLIC_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"table_id|format_id|format_key|registry_id|entry_id|component_id|member_id|"
    r"archive_id|message_id|uuid|source_bytes|raw_bytes|wire_bytes|payload|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_PUBLIC_RAW_CONTAINER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:source_bytes|raw_bytes|wire_bytes|payload|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:Vec\s*<\s*u8\s*>|Box\s*<\s*\[\s*u8\s*\]\s*>|"
    r"(?:std\s*::\s*borrow\s*::\s*)?Cow\s*<[^>]*\[\s*u8\s*\]|"
    r"(?:std\s*::\s*sync\s*::\s*)?Arc\s*<\s*\[\s*u8\s*\]\s*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_PUBLIC_SELECTOR_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:sheet|table|row|column|position|cell)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|"
    r"[A-Z][A-Za-z0-9_]*(?:Id|ID|Key|Identifier))\b"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_TYPED_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|key|"
    r"(?:cell|object|native|model|table|format|registry|entry|component|archive|"
    r"message|storage|resource)_(?:id|identifier|key))"
    r"[ \t\r\n]*:[ \t\r\n]*"
    r"(?:[A-Z][A-Za-z0-9_]*(?:Id|ID|Key|Identifier)|"
    r"(?:u64|u32|usize|i64|i32))\b"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_HIDDEN_MODULE = re.compile(
    rf"(?m)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
    rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_MODULE)}\b"
)
NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_GENERATED_MODULE = re.compile(
    rf"(?m)^\s*pub\s+mod\s+"
    rf"{re.escape(NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_GENERATED_MODULE)}\b"
)

# These are deliberately production-code markers rather than prose checks.
# They make the owner gate prove that the selector-first route is actually
# source-bound, bounded, and candidate-verified; comments, literals, and
# test-only decoys are removed before the markers are evaluated.
NUMBERS_TABLE_CELL_DURATION_FORMAT_OWNER_REQUIRED_MARKERS = {
    "exact-source admission": re.compile(
        r"(?<![A-Za-z0-9_])source_is_exact(?![A-Za-z0-9_])"
    ),
    "source-owner identity": re.compile(
        r"(?<![A-Za-z0-9_])source_owner(?![A-Za-z0-9_])"
    ),
    "exact-artifact patch capability": re.compile(
        r"(?<![A-Za-z0-9_])OwnedExactArtifacts(?![A-Za-z0-9_])"
    ),
    "patch-owner authorization": re.compile(
        r"(?<![A-Za-z0-9_])authorizes_owner(?![A-Za-z0-9_])"
    ),
    "selector-first cell resolution": re.compile(
        r"(?<![A-Za-z0-9_])resolve_cell(?![A-Za-z0-9_])"
    ),
    "strict Duration reader route": re.compile(
        r"(?<![A-Za-z0-9_])read_duration_format(?![A-Za-z0-9_])"
    ),
    "strict Duration writer route": re.compile(
        r"(?<![A-Za-z0-9_])rewrite_duration_format(?![A-Za-z0-9_])"
    ),
    "transaction budget": re.compile(
        r"(?<![A-Za-z0-9_])TransactionBudget(?![A-Za-z0-9_])"
    ),
    "bounded reassembly preflight": re.compile(
        r"(?<![A-Za-z0-9_])preflight_reassembly(?![A-Za-z0-9_])"
    ),
    "charged transaction resources": re.compile(
        r"(?<![A-Za-z0-9_])charge_[A-Za-z0-9_]+(?![A-Za-z0-9_])"
    ),
    "candidate reopening": re.compile(
        r"(?<![A-Za-z0-9_])(?:candidate_reopen|charge_candidate_reopen)"
        r"(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "candidate locality verification": re.compile(
        r"(?<![A-Za-z0-9_])verify_duration_package_locality(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}

# Only the dedicated Draft convenience methods are retired.  The
# generic ``DataFormat`` helpers intentionally remain available to source-
# built packages and to the attached Pages/Keynote table adapters.
RETIRED_IWA_NUMBERS_TABLE_CELL_DURATION_FORMAT_METHODS = (
    "table_cell_duration_format",
    "set_table_cell_duration_format",
    "reset_table_cell_duration_format",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_DURATION_FORMAT_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_DURATION_FORMAT_METHODS
)
IWA_NUMBERS_TABLE_CELL_DURATION_FORMAT_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?P<method>table_cell_duration_format|"
    r"set_table_cell_duration_format|reset_table_cell_duration_format)"
    r"(?![A-Za-z0-9_])[ \t\r\n]*\("
)

# Custom is the document-scoped Numbers display-format owner.  Unlike the
# scalar format owners above, its semantic value is an enum backed by a
# private UUID-indexed registry.  Keep the checker explicit about every
# ingress/export seam: a similarly named local helper must not be enough to
# activate the focused owner, and none of the physical registry graph may
# leak through the archive-free facade.  There is intentionally no
# ``RETIRED_IWA_*CUSTOM*`` inventory here: the migration-host Custom route is
# still compatibility-only and remains outside this ratchet.
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/custom.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TRANSACTION_SOURCE = (
    NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_custom_format.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_custom_format"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_NATIVE_SOURCE = Path(
    "crates/litchi-iwa-table/src/cell_format/display.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_custom_format_codec.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_custom_format_codec"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_custom_format_generated"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_REQUIRED_APIS = (
    "decode_custom_format_with_report",
    "decode_custom_format_list_with_report",
    "prepare_custom_format_rewrite",
    "PreparedCustomFormatRewrite",
    "prepare_custom_format_list_rewrite",
    "PreparedCustomFormatListRewrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "canonical_custom_format",
    "canonical_custom_format_list",
    "rewrite_custom_format",
    "rewrite_custom_format_list",
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_METHODS = (
    "table_cell_custom_format",
    "edit_table_cell_custom_format",
    "apply_table_cell_custom_format",
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_TYPES = (
    "Custom",
    "Number",
    "Text",
    "DateTime",
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in (
        "CustomFormat",
        "TableCellCustomFormat",
        "CellCustomFormat",
        "NumbersTableCellCustomFormat",
    )
    for suffix in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_FLAT_ALIAS_PATTERN = re.compile(
    r"(?:Custom|TableCellCustom|CellCustom|NumbersTableCellCustom)"
    r"(?:Format)?(?:Edit|Patch|Commit|Diagnostics|Error|LimitKind|Path)\b"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_custom_format|"
    r"cell[ \t\r\n]*::[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?custom)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "BncCell",
        "ComponentCatalog",
        "CustomFormatArchive",
        "CustomFormatListArchive",
        "CustomFormatListSnapshot",
        "CustomFormatSnapshot",
        "EntryEdit",
        "ExactArtifacts",
        "FormatStructArchive",
        "HostPackage",
        "OwnedExactArtifacts",
        "RawMessage",
        "SourceCatalog",
        "SnappyStream",
        "Uuid",
        "UUID",
    }
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "RewriteOutput",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "bnc", "tn", "tst", "tsp", "tswp"}
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"table_id|format_id|format_key|registry_id|entry_id|component_id|member_id|"
    r"archive_id|message_id|uuid|uuid_pair|member|member_name|payload|"
    r"payload_bytes|archive|archive_bytes|source_bytes|raw_bytes|wire_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\]|"
    r"Vec\s*<\s*u8\s*>|Box\s*<\s*\[\s*u8\s*\]\s*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_RAW_CONTAINER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:source_bytes|raw_bytes|wire_bytes|payload|"
    r"payload_bytes|archive_bytes|bytes)[ \t\r\n]*:[ \t\r\n]*(?:Vec\s*<\s*u8\s*>|"
    r"Box\s*<\s*\[\s*u8\s*\]\s*>|"
    r"(?:std\s*::\s*borrow\s*::\s*)?Cow\s*<[^>]*\[\s*u8\s*\]|"
    r"(?:std\s*::\s*sync\s*::\s*)?Arc\s*<\s*\[\s*u8\s*\]\s*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_UUID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:uuid|uuid_pair|registry_uuid|custom_uuid)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:Uuid|UUID|[A-Z][A-Za-z0-9_]*(?:Uuid|UUID))\b"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TYPED_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|key|registry_id|format_id|"
    r"entry_id|member_id|component_id|archive_id|message_id|uuid)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:[A-Z][A-Za-z0-9_]*(?:Id|ID|Key|Identifier)|"
    r"(?:u64|u32|usize|i64|i32))\b"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_RAW_HELPER = re.compile(
    r"(?<![A-Za-z0-9_])(?:read_custom_format|write_custom_format|"
    r"rewrite_custom_format|read_custom_registry|rewrite_custom_registry|"
    r"locate_custom_registry|registry_payload|parse_custom_registry|"
    r"custom_format_type|custom_uid|format_identifier|object_identifier|"
    r"member_name|payload_bytes|RawMessage|CustomFormatArchive)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_REQUIRED_MARKERS = {
    "document-scoped registry": re.compile(
        r"(?<![A-Za-z0-9_])(?:document[_ -]?scoped|rooted|registry|"
        r"custom[_ -]?format[_ -]?list|field[_ -]?9|message[_ -]?type)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "custom family discriminators": re.compile(
        r"(?<![A-Za-z0-9_])(?:270|271|272|CUSTOM[_ -]?FORMAT[_ -]?TYPE|"
        r"Number|Text|Date[_ -]?Time)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "format-list/refcount closure": re.compile(
        r"(?<![A-Za-z0-9_])(?:format[_ -]?list|registry|refcount|ref_count|"
        r"reference|reuse|cull|copy[_ -]?on[_ -]?write)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "exact-source transaction": re.compile(
        r"(?<![A-Za-z0-9_])(?:exact[_ -]?source|source[_ -]?bound|"
        r"PatchConflict|inverse|no[_ -]?op)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "candidate verification/locality": re.compile(
        r"(?<![A-Za-z0-9_])(?:candidate|reopen|readback|locality|same[_ -]?content)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "bounded staging": re.compile(
        r"(?<![A-Za-z0-9_])(?:budget|limit|preflight|bounded|allocation)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
    "fail-closed unsupported graph": re.compile(
        r"(?<![A-Za-z0-9_])(?:unsupported|UnsupportedDependency|UnsupportedSource|"
        r"ambiguous|cross[_ -]?component|fail[_ -]?closed)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_custom_format.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_EXAMPLE = Path(
    "crates/litchi-numbers/examples/edit_table_cell_custom_format.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_custom_format_codec.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_custom_format_codec"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_custom_format.rs"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_custom_format"
)
NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_MARKERS = {
    "all custom families": re.compile(
        r"(?:Number|Text|Date[_ -]?Time|Custom)", re.IGNORECASE
    ),
    "set/clear/reset lifecycle": re.compile(
        r"(?:set|clear|reset|automatic)", re.IGNORECASE
    ),
    "exact no-op and inverse": re.compile(
        r"(?:no[_ -]?op|unchanged|inverse|byte[_ -]?for[_ -]?byte)",
        re.IGNORECASE,
    ),
    "registry UUID/refcount closure": re.compile(
        r"(?:uuid|registry|refcount|reference|reuse|cull)", re.IGNORECASE
    ),
    "unknown/member preservation": re.compile(
        r"(?:unknown|member|nested|source[_ -]?span)", re.IGNORECASE
    ),
    "candidate/locality verification": re.compile(
        r"(?:candidate|reopen|readback|locality|unchanged)", re.IGNORECASE
    ),
    "typed refusal coverage": re.compile(
        r"(?:stale|foreign|wrong[_ -]?family|malformed|unsupported|locked|refus)",
        re.IGNORECASE,
    ),
}

# Percentage is deliberately audited as a separate focused owner.  It shares
# the archive-free semantic family with Number, but its native format
# discriminator is 258 and its public transaction namespace must not become a
# second set of flat aliases.  The semantic implementation may live in its
# eventual ``percentage.rs`` leaf or, during the split, in ``number.rs`` with
# an explicit public percentage namespace; the focused package owner and
# native graph core have stable, dedicated paths.
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/percentage.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_FALLBACK_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/percentage/transaction.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_percentage_format.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_percentage_format"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_percentage_format_codec.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_PUBLIC_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_percentage_format_codec"
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_display_format_codec"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_percentage_format_generated"
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_display_format_generated"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULES = (
    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULE,
    "buffa_numbers_table_cell_pop_up_menu_generated",
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE,
    Path("crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"),
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_TYPES = ("Percentage",)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PACKAGE_METHODS = (
    "table_cell_percentage_format",
    "edit_table_cell_percentage_format",
    "apply_table_cell_percentage_format",
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EDIT_METHODS = (
    "before",
    "after",
    "set",
    "clear",
    "reset",
    "commit",
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in (
        "PercentageFormat",
        "TableCellPercentageFormat",
        "CellPercentageFormat",
        "NumbersTableCellPercentageFormat",
    )
    for suffix in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIAS_PATTERN = re.compile(
    r"(?:Percentage|TableCellPercentage|CellPercentage|"
    r"NumbersTableCellPercentage)(?:Format)?"
    r"(?:Edit|Patch|Commit|Diagnostics|Error|LimitKind|Path)\b"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_percentage_format|"
    r"cell[ \t\r\n]*::[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?percentage)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_CELL_DATA_FORMAT_PERCENTAGE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?percentage\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_CELL_PERCENTAGE_FORMAT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_percentage_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_percentage_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PHYSICAL_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PHYSICAL_TYPES
    | frozenset(
        {
            "PercentageFormatArchive",
            "PercentageFormatSnapshot",
            "PercentageFormatWrite",
            "DisplayFormatArchive",
            "DisplayFormatSnapshot",
            "DisplayFormatWrite",
            "FormatSnapshot",
            "FormatWrite",
        }
    )
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_WIRE_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_WIRE_TYPES
    | frozenset(
        {
            "DisplayFormatDecodeError",
            "DisplayFormatDecodeReport",
            "DisplayFormatWireView",
            "PercentageFormatDecodeError",
            "PercentageFormatDecodeReport",
            "PercentageFormatWireView",
        }
    )
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PROTO_ORIGINS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PROTO_ORIGINS
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_CONTAINER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TYPED_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_RAW_HELPER = re.compile(
    r"(?<![A-Za-z0-9_])(?:cell_(?:number|percentage)_format|"
    r"set_cell_(?:number|percentage)_format|reset_cell_(?:number|percentage)_format|"
    r"read_native_(?:number|percentage|display)_format|"
    r"set_native_(?:number|percentage|display)_format|"
    r"read_(?:number|percentage|display)_format_wire|"
    r"write_(?:number|percentage|display)_format_wire|raw_(?:number|percentage|display)_format|"
    r"patch_(?:number|percentage|display)_format|parse_wire_fields|patch_varint_field|"
    r"patch_length_delimited_field|FormatStructArchive)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_REQUIRED_APIS = (
    ("decode_percentage_format_with_report", "decode_cell_percentage_format_with_report", "decode_display_format_with_report"),
    ("prepare_percentage_format_rewrite", "prepare_cell_percentage_format_rewrite", "prepare_display_format_rewrite"),
    ("PercentageFormatSnapshot", "DisplayFormatSnapshot"),
    ("PercentageFormatWrite", "DisplayFormatWrite"),
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    ("canonical_percentage_format", "canonical_display_format"),
    ("rewrite_percentage_format", "rewrite_display_format"),
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_REQUIRED_MARKER_GROUPS = {
    "lazy Buffa projection": (
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULES,
        ("decode_lazy_view", "decode_view", "LazyView", "visitor"),
        "buffa",
    ),
    "strict canonical ingress": (
        ("preflight", "scan_percentage_message", "scan_decimal_message", "scan_display_format"),
        ("duplicate", "is_some", "wrapper_seen"),
        ("noncanonical", "non-canonical", "varint_canonical"),
    ),
    "unknown/raw preservation": ("unknown", "raw", "extend_from_slice"),
    "bounded prepared rewrite": (
        ("MAX_RECURSION_LIMIT", "MAX_RECURSION"),
        "execution_requirements",
        "execute",
    ),
    "percentage format fields": (
        ("decimal_places", "decimal"),
        ("negative_style", "negative"),
        ("thousands_separator", "thousands"),
    ),
    "registry/refcount closure": (
        ("format_table", "registry", "refcount", "ref_count", "references"),
        ("reference", "entry"),
    ),
}
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_REQUIRED_MARKERS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_REQUIRED_MARKERS
)
RETIRED_IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_HELPERS = (
    "cell_percentage_format",
    "set_cell_percentage_format",
    "reset_cell_percentage_format",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_HELPER_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_HELPERS
)
RETIRED_IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TESTS = (
    "source_built_table_converts_and_reuses_percentage_formats",
    "source_built_table_roundtrips_percentage_data_format",
    "invalid_percentage_format_coordinate_is_transactional",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TEST_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TESTS
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EXAMPLE = Path(
    "crates/litchi-numbers/examples/edit_table_cell_percentage_format.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_percentage_format.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_percentage_format_codec.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_percentage_format_codec"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_percentage_format.rs"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_percentage_format"
)
NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_MARKERS = {
    "explicit and automatic readback": re.compile(
        r"(?:Option|automatic|explicit|read).*Percentage", re.IGNORECASE
    ),
    "set/reset transaction": re.compile(
        r"(?:edit_table_cell_percentage_format|set\(|clear\(|reset|automatic)",
        re.IGNORECASE,
    ),
    "exact no-op and inverse": re.compile(
        r"(?:no[_ -]?op|unchanged|inverse|byte[_ -]?for[_ -]?byte)",
        re.IGNORECASE,
    ),
    "conflict and fail-closed refusal": re.compile(
        r"(?:stale|foreign|ambiguous|conflict|unsupported|malformed)",
        re.IGNORECASE,
    ),
    "selector/locality verification": re.compile(
        r"(?:SheetSelector|TableSelector|(?-i:\bAt\b)|locality|reopen|readback)",
        re.IGNORECASE,
    ),
}
IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FOCUSED_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:litchi_numbers[ \t\r\n]*::[ \t\r\n]*Package|"
    r"FocusedNumbersPackage|focused_percentage_format_package|"
    r"focused_number_format_package|focused_numbers_package)[^;{}]*\.[ \t\r\n]*"
    r"(?:table_cell_percentage_format|edit_table_cell_percentage_format|"
    r"apply_table_cell_percentage_format)\s*\("
)
IWA_NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_README_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor|editor)"
    r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?(?P<method>"
    r"table_cell_percentage_format|set_table_cell_percentage_format|"
    r"reset_table_cell_percentage_format)\b[ \t\r\n]*\("
)

# Currency is the next focused decimal family after Number and Percentage.
# Keep its public semantic/transaction namespace independent while sharing the
# private native display-format graph. During the split the semantic value may
# still be defined in ``number.rs``; the explicit fallback keeps that
# transitional ownership visible without permitting a second package owner.
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/currency.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_FALLBACK_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/currency/transaction.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_currency_format.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_currency_format"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_currency_format_codec.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_PUBLIC_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_currency_format_codec"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_currency_format_generated"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_GENERATED_MODULES = (
    NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULE,
    "buffa_numbers_table_cell_pop_up_menu_generated",
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE,
    Path("crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"),
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_TYPES = (
    "Currency",
    "CurrencyCode",
    "CurrencyStyle",
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PACKAGE_METHODS = (
    "table_cell_currency_format",
    "edit_table_cell_currency_format",
    "apply_table_cell_currency_format",
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_EDIT_METHODS = (
    "before",
    "after",
    "set",
    "clear",
    "reset",
    "commit",
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in (
        "CurrencyFormat",
        "TableCellCurrencyFormat",
        "CellCurrencyFormat",
        "NumbersTableCellCurrencyFormat",
    )
    for suffix in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FLAT_ALIAS_PATTERN = re.compile(
    r"(?:Currency|TableCellCurrency|CellCurrency|NumbersTableCellCurrency)"
    r"(?:Format)?(?:Edit|Patch|Commit|Diagnostics|Error|LimitKind|Path)\b"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_currency_format|"
    r"cell[ \t\r\n]*::[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?currency)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_CELL_DATA_FORMAT_CURRENCY_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?currency\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_CELL_CURRENCY_FORMAT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_CURRENCY_FORMAT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_currency_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_CURRENCY_FORMAT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_currency_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PHYSICAL_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PHYSICAL_TYPES
    | frozenset(
        {
            "CurrencyFormatArchive",
            "CurrencyFormatSnapshot",
            "CurrencyFormatWrite",
            "DisplayFormatArchive",
            "DisplayFormatSnapshot",
            "DisplayFormatWrite",
            "FormatSnapshot",
            "FormatWrite",
        }
    )
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_WIRE_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_WIRE_TYPES
    | frozenset(
        {
            "CurrencyFormatDecodeError",
            "CurrencyFormatDecodeReport",
            "CurrencyFormatWireView",
            "DisplayFormatDecodeError",
            "DisplayFormatDecodeReport",
            "DisplayFormatWireView",
        }
    )
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PROTO_ORIGINS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PROTO_ORIGINS
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PUBLIC_RAW_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PUBLIC_RAW_CONTAINER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TYPED_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_RAW_HELPER = re.compile(
    r"(?<![A-Za-z0-9_])(?:cell_(?:number|percentage|currency)_format|"
    r"set_cell_(?:number|percentage|currency)_format|"
    r"reset_cell_(?:number|percentage|currency)_format|"
    r"read_native_(?:number|percentage|currency|display)_format|"
    r"set_native_(?:number|percentage|currency|display)_format|"
    r"read_(?:number|percentage|currency|display)_format_wire|"
    r"write_(?:number|percentage|currency|display)_format_wire|"
    r"raw_(?:number|percentage|currency|display)_format|"
    r"patch_(?:number|percentage|currency|display)_format|"
    r"parse_wire_fields|patch_varint_field|patch_length_delimited_field|"
    r"FormatStructArchive)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_REQUIRED_APIS = (
    (
        "decode_currency_format_with_report",
        "decode_cell_currency_format_with_report",
        "decode_display_format_with_report",
    ),
    (
        "prepare_currency_format_rewrite",
        "prepare_cell_currency_format_rewrite",
        "prepare_display_format_rewrite",
    ),
    ("CurrencyFormatSnapshot", "DisplayFormatSnapshot"),
    ("CurrencyFormatWrite", "DisplayFormatWrite"),
    ("PreparedCurrencyFormatRewrite", "PreparedDisplayFormatRewrite"),
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    ("canonical_currency_format", "canonical_display_format"),
    ("rewrite_currency_format", "rewrite_display_format"),
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_REQUIRED_MARKER_GROUPS = {
    "lazy Buffa projection": (
        NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_GENERATED_MODULES,
        ("decode_lazy_view", "decode_view", "LazyView", "visitor"),
        "buffa",
    ),
    "strict canonical ingress": (
        (
            "preflight",
            "scan_currency_message",
            "scan_decimal_message",
            "scan_display_format",
        ),
        ("duplicate", "is_some", "wrapper_seen"),
        ("noncanonical", "non-canonical", "varint_canonical"),
    ),
    "unknown/raw preservation": ("unknown", "raw", "extend_from_slice"),
    "bounded prepared rewrite": (
        ("MAX_RECURSION_LIMIT", "MAX_RECURSION"),
        "execution_requirements",
        "execute",
    ),
    "currency format fields": (
        ("currency_code", "currency", "code"),
        ("style", "accounting", "currency_style"),
        ("decimal_places", "decimal"),
        ("negative_style", "negative"),
        ("thousands_separator", "thousands"),
    ),
    "registry/refcount closure": (
        ("format_table", "registry", "refcount", "ref_count", "references"),
        ("reference", "entry"),
    ),
}
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_REQUIRED_MARKERS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_REQUIRED_MARKERS
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_EXAMPLE = Path(
    "crates/litchi-numbers/examples/edit_table_cell_currency_format.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_INTEGRATION_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_currency_format.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_currency_format_codec.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_currency_format_codec"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_currency_format.rs"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_currency_format"
)
NUMBERS_TABLE_CELL_CURRENCY_FORMAT_INTEGRATION_MARKERS = {
    "explicit and automatic readback": re.compile(
        r"(?:Option|automatic|explicit|read).*Currency", re.IGNORECASE
    ),
    "set/reset transaction": re.compile(
        r"(?:edit_table_cell_currency_format|set\(|clear\(|reset|automatic)",
        re.IGNORECASE,
    ),
    "currency code and style": re.compile(
        r"(?:CurrencyCode|CurrencyStyle|accounting|currency_code|style)",
        re.IGNORECASE,
    ),
    "exact no-op and inverse": re.compile(
        r"(?:no[_ -]?op|unchanged|inverse|byte[_ -]?for[_ -]?byte)",
        re.IGNORECASE,
    ),
    "conflict and fail-closed refusal": re.compile(
        r"(?:stale|foreign|ambiguous|conflict|unsupported|malformed)",
        re.IGNORECASE,
    ),
    "selector/locality verification": re.compile(
        r"(?:SheetSelector|TableSelector|(?-i:\bAt\b)|locality|reopen|readback)",
        re.IGNORECASE,
    ),
}
IWA_NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FOCUSED_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:litchi_numbers[ \t\r\n]*::[ \t\r\n]*Package|"
    r"FocusedNumbersPackage|focused_currency_format_package|"
    r"focused_number_format_package|focused_numbers_package)[^;{}]*\.[ \t\r\n]*"
    r"(?:table_cell_currency_format|edit_table_cell_currency_format|"
    r"apply_table_cell_currency_format)\s*\("
)

# Scientific notation is the fourth focused decimal family.  It has a
# deliberately independent semantic/transaction namespace while reusing the
# private native display-format graph.  During the split the value may still
# be re-exported from ``number.rs``; the explicit fallback keeps that
# transitional ownership visible without permitting a second package owner.
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/scientific.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_FALLBACK_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/scientific/transaction.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_scientific_format.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_scientific_format"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_scientific_format_codec.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_PUBLIC_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_scientific_format_codec"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_scientific_format_generated"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_GENERATED_MODULES = (
    NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULE,
    "buffa_numbers_table_cell_currency_format_generated",
    "buffa_numbers_table_cell_pop_up_menu_generated",
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE,
    Path("crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"),
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_TYPES = ("Scientific",)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PACKAGE_METHODS = (
    "table_cell_scientific_format",
    "edit_table_cell_scientific_format",
    "apply_table_cell_scientific_format",
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_EDIT_METHODS = (
    "before",
    "after",
    "set",
    "clear",
    "reset",
    "commit",
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in (
        "ScientificFormat",
        "TableCellScientificFormat",
        "CellScientificFormat",
        "NumbersTableCellScientificFormat",
    )
    for suffix in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FLAT_ALIAS_PATTERN = re.compile(
    r"(?:Scientific|TableCellScientific|CellScientific|NumbersTableCellScientific)"
    r"(?:Format)?(?:Edit|Patch|Commit|Diagnostics|Error|LimitKind|Path)\b"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_scientific_format|"
    r"cell[ \t\r\n]*::[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?scientific)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_CELL_DATA_FORMAT_SCIENTIFIC_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?scientific\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_CELL_SCIENTIFIC_FORMAT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_SCIENTIFIC_FORMAT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_scientific_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_SCIENTIFIC_FORMAT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_scientific_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PHYSICAL_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PHYSICAL_TYPES
    | frozenset(
        {
            "ScientificFormatArchive",
            "ScientificFormatSnapshot",
            "ScientificFormatWrite",
            "DisplayFormatArchive",
            "DisplayFormatSnapshot",
            "DisplayFormatWrite",
            "FormatSnapshot",
            "FormatWrite",
        }
    )
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_WIRE_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_WIRE_TYPES
    | frozenset(
        {
            "ScientificFormatDecodeError",
            "ScientificFormatDecodeReport",
            "ScientificFormatWireView",
            "DisplayFormatDecodeError",
            "DisplayFormatDecodeReport",
            "DisplayFormatWireView",
        }
    )
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PROTO_ORIGINS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PROTO_ORIGINS
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PUBLIC_RAW_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PUBLIC_RAW_CONTAINER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TYPED_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_RAW_HELPER = re.compile(
    r"(?<![A-Za-z0-9_])(?:cell_(?:number|percentage|currency|scientific)_format|"
    r"set_cell_(?:number|percentage|currency|scientific)_format|"
    r"reset_cell_(?:number|percentage|currency|scientific)_format|"
    r"read_native_(?:number|percentage|currency|scientific|display)_format|"
    r"set_native_(?:number|percentage|currency|scientific|display)_format|"
    r"read_(?:number|percentage|currency|scientific|display)_format_wire|"
    r"write_(?:number|percentage|currency|scientific|display)_format_wire|"
    r"raw_(?:number|percentage|currency|scientific|display)_format|"
    r"patch_(?:number|percentage|currency|scientific|display)_format|"
    r"parse_wire_fields|patch_varint_field|patch_length_delimited_field|"
    r"FormatStructArchive)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_REQUIRED_APIS = (
    (
        "decode_scientific_format_with_report",
        "decode_cell_scientific_format_with_report",
        "decode_display_format_with_report",
    ),
    (
        "prepare_scientific_format_rewrite",
        "prepare_cell_scientific_format_rewrite",
        "prepare_display_format_rewrite",
    ),
    ("ScientificFormatSnapshot", "DisplayFormatSnapshot"),
    ("ScientificFormatWrite", "DisplayFormatWrite"),
    ("PreparedScientificFormatRewrite", "PreparedDisplayFormatRewrite"),
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    ("canonical_scientific_format", "canonical_display_format"),
    ("rewrite_scientific_format", "rewrite_display_format"),
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_REQUIRED_MARKER_GROUPS = {
    "lazy Buffa projection": (
        NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_GENERATED_MODULES,
        ("decode_lazy_view", "decode_view", "LazyView", "visitor"),
        "buffa",
    ),
    "strict canonical ingress": (
        (
            "preflight",
            "scan_scientific_message",
            "scan_decimal_message",
            "scan_display_format",
        ),
        ("duplicate", "is_some", "wrapper_seen"),
        ("noncanonical", "non-canonical", "varint_canonical"),
    ),
    "unknown/raw preservation": ("unknown", "raw", "extend_from_slice"),
    "bounded prepared rewrite": (
        ("MAX_RECURSION_LIMIT", "MAX_RECURSION"),
        "execution_requirements",
        "execute",
    ),
    "scientific format fields": (
        ("format_type", "NATIVE_SCIENTIFIC_FORMAT_TYPE"),
        ("decimal_places", "decimal"),
        ("negative_style", "NATIVE_SCIENTIFIC_NEGATIVE_STYLE"),
        ("show_thousands_separator", "NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR"),
    ),
    "registry/refcount closure": (
        ("format_table", "registry", "refcount", "ref_count", "references"),
        ("reference", "entry"),
    ),
}
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_REQUIRED_MARKERS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_REQUIRED_MARKERS
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_EXAMPLE = Path(
    "crates/litchi-numbers/examples/edit_table_cell_scientific_format.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_INTEGRATION_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_scientific_format.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_scientific_format_codec.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_scientific_format_codec"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_scientific_format.rs"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_scientific_format"
)
NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_INTEGRATION_MARKERS = {
    "explicit and automatic readback": re.compile(
        r"(?:Option|automatic|explicit|read).*Scientific", re.IGNORECASE
    ),
    "set/reset transaction": re.compile(
        r"(?:edit_table_cell_scientific_format|set\(|clear\(|reset|automatic)",
        re.IGNORECASE,
    ),
    "scientific precision and exponent": re.compile(
        r"(?:Scientific|decimal_places|exponent|fixed[_ -]?precision)",
        re.IGNORECASE,
    ),
    "exact no-op and inverse": re.compile(
        r"(?:no[_ -]?op|unchanged|inverse|byte[_ -]?for[_ -]?byte)",
        re.IGNORECASE,
    ),
    "conflict and fail-closed refusal": re.compile(
        r"(?:stale|foreign|ambiguous|conflict|unsupported|malformed)",
        re.IGNORECASE,
    ),
    "selector/locality verification": re.compile(
        r"(?:SheetSelector|TableSelector|(?-i:\bAt\b)|locality|reopen|readback)",
        re.IGNORECASE,
    ),
}
IWA_NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FOCUSED_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:litchi_numbers[ \t\r\n]*::[ \t\r\n]*Package|"
    r"FocusedNumbersPackage|focused_scientific_format_package|"
    r"focused_number_format_package|focused_numbers_package)[^;{}]*\.[ \t\r\n]*"
    r"(?:table_cell_scientific_format|edit_table_cell_scientific_format|"
    r"apply_table_cell_scientific_format)\s*\("
)
IWA_NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_README_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor|editor)"
    r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?(?P<method>"
    r"table_cell_scientific_format|set_table_cell_scientific_format|"
    r"reset_table_cell_scientific_format)\b[ \t\r\n]*\("
)

# Fraction notation is the next focused decimal family.  It has a
# deliberately independent semantic/transaction namespace while reusing the
# private native display-format graph.  During the split the value may still
# be re-exported from ``number.rs``; the explicit fallback keeps that
# transitional ownership visible without permitting a second package owner.
NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/fraction.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_FALLBACK_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/data_format/fraction/transaction.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_fraction_format.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_cell_fraction_format"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_fraction_format_codec.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_PUBLIC_SOURCE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_MODULE = (
    "numbers_table_cell_fraction_format_codec"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_GENERATED_MODULE = (
    "buffa_numbers_table_cell_fraction_format_generated"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_GENERATED_MODULES = (
    NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULE,
    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULE,
    "buffa_numbers_table_cell_currency_format_generated",
    "buffa_numbers_table_cell_pop_up_menu_generated",
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_SOURCE,
    NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE,
    Path("crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"),
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_TYPES = (
    "Fraction",
    "FractionAccuracy",
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_TYPES
    + NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_PACKAGE_METHODS = (
    "table_cell_fraction_format",
    "edit_table_cell_fraction_format",
    "apply_table_cell_fraction_format",
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_EDIT_METHODS = (
    "before",
    "after",
    "set",
    "clear",
    "reset",
    "commit",
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in (
        "FractionFormat",
        "TableCellFractionFormat",
        "CellFractionFormat",
        "NumbersTableCellFractionFormat",
    )
    for suffix in NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_FLAT_ALIAS_PATTERN = re.compile(
    r"(?:Fraction|TableCellFraction|CellFraction|NumbersTableCellFraction)"
    r"(?:Format)?(?:Edit|Patch|Commit|Diagnostics|Error|LimitKind|Path)\b"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_fraction_format|"
    r"cell[ \t\r\n]*::[ \t\r\n]*data_format[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?fraction)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_CELL_DATA_FORMAT_FRACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?fraction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_CELL_FRACTION_FORMAT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_FRACTION_FORMAT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cell_fraction_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_FRACTION_FORMAT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?table_cell_fraction_format\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_PHYSICAL_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PHYSICAL_TYPES
    | frozenset(
        {
            "FractionFormatArchive",
            "FractionFormatSnapshot",
            "FractionFormatWrite",
            "DisplayFormatArchive",
            "DisplayFormatSnapshot",
            "DisplayFormatWrite",
            "FormatSnapshot",
            "FormatWrite",
        }
    )
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_WIRE_TYPES = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_WIRE_TYPES
    | frozenset(
        {
            "FractionFormatDecodeError",
            "FractionFormatDecodeReport",
            "FractionFormatWireView",
            "DisplayFormatDecodeError",
            "DisplayFormatDecodeReport",
            "DisplayFormatWireView",
        }
    )
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_PROTO_ORIGINS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PROTO_ORIGINS
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_PUBLIC_RAW_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_PUBLIC_RAW_CONTAINER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_TYPED_RAW_ID_PARAMETER = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_RAW_HELPER = re.compile(
    r"(?<![A-Za-z0-9_])(?:cell_(?:number|percentage|currency|fraction)_format|"
    r"set_cell_(?:number|percentage|currency|fraction)_format|"
    r"reset_cell_(?:number|percentage|currency|fraction)_format|"
    r"read_native_(?:number|percentage|currency|fraction|display)_format|"
    r"set_native_(?:number|percentage|currency|fraction|display)_format|"
    r"read_(?:number|percentage|currency|fraction|display)_format_wire|"
    r"write_(?:number|percentage|currency|fraction|display)_format_wire|"
    r"raw_(?:number|percentage|currency|fraction|display)_format|"
    r"patch_(?:number|percentage|currency|fraction|display)_format|"
    r"parse_wire_fields|patch_varint_field|patch_length_delimited_field|"
    r"FormatStructArchive)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_REQUIRED_APIS = (
    (
        "decode_fraction_format_with_report",
        "decode_cell_fraction_format_with_report",
        "decode_display_format_with_report",
    ),
    (
        "prepare_fraction_format_rewrite",
        "prepare_cell_fraction_format_rewrite",
        "prepare_display_format_rewrite",
    ),
    ("FractionFormatSnapshot", "DisplayFormatSnapshot"),
    ("FractionFormatWrite", "DisplayFormatWrite"),
    ("PreparedFractionFormatRewrite", "PreparedDisplayFormatRewrite"),
    ("PreparedFractionFormatWrite", "PreparedDisplayFormatWrite"),
    (
        "prepare_fraction_format_write",
        "prepare_fraction_format_append",
        "prepare_display_format_write",
    ),
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    ("canonical_fraction_format", "canonical_display_format"),
    ("rewrite_fraction_format", "rewrite_display_format"),
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_REQUIRED_MARKER_GROUPS = {
    "lazy Buffa projection": (
        NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_GENERATED_MODULES,
        ("decode_lazy_view", "decode_view", "LazyView", "visitor"),
        "buffa",
    ),
    "strict canonical ingress": (
        (
            "preflight",
            "scan_fraction_message",
            "scan_decimal_message",
            "scan_display_format",
        ),
        ("duplicate", "is_some", "wrapper_seen"),
        ("noncanonical", "non-canonical", "varint_canonical"),
    ),
    "unknown/raw preservation": ("unknown", "raw", "extend_from_slice"),
    "bounded prepared rewrite": (
        ("MAX_RECURSION_LIMIT", "MAX_RECURSION"),
        "execution_requirements",
        "execute",
    ),
    "fraction format fields": (
        ("format_type", "NATIVE_FRACTION_FORMAT_TYPE"),
        (
            "fraction_accuracy",
            "FORMAT_FRACTION_ACCURACY_FIELD",
            "FractionAccuracy",
            "accuracy",
        ),
    ),
    "fraction replacement tri-state": (
        ("requires_fraction_replacement", "FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD"),
        ("Some(false)", "Some(value != 0)", "value != 0", "value == 0"),
        ("Some(true)", "== Some(true)", "value != 0", "return Err"),
        (
            "snapshot.requires_fraction_replacement",
            "view.requires_fraction_replacement",
            "source-authoritative",
            "preserved",
        ),
    ),
    "registry/refcount closure": (
        ("format_table", "registry", "refcount", "ref_count", "references"),
        ("reference", "entry"),
    ),
}
NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_REQUIRED_MARKERS = (
    NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_REQUIRED_MARKERS
)
RETIRED_IWA_NUMBERS_TABLE_CELL_FRACTION_FORMAT_HELPERS = (
    "cell_fraction_format",
    "set_cell_fraction_format",
    "reset_cell_fraction_format",
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_EXAMPLE = Path(
    "crates/litchi-numbers/examples/edit_table_cell_fraction_format.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_INTEGRATION_SOURCE = Path(
    "crates/litchi-numbers/tests/table_cell_fraction_format.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_fraction_format_codec.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_fraction_format_codec"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_fraction_format.rs"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_fraction_format"
)
NUMBERS_TABLE_CELL_FRACTION_FORMAT_INTEGRATION_MARKERS = {
    "explicit and automatic readback": re.compile(
        r"(?:Option|automatic|explicit|read).*Fraction", re.IGNORECASE
    ),
    "set/reset transaction": re.compile(
        r"(?:edit_table_cell_fraction_format|set\(|clear\(|reset|automatic)",
        re.IGNORECASE,
    ),
    "fraction accuracy and denominator": re.compile(
        r"(?:Fraction|FractionAccuracy|fraction_accuracy|accuracy|denominator)",
        re.IGNORECASE,
    ),
    "exact no-op and inverse": re.compile(
        r"(?:no[_ -]?op|unchanged|inverse|byte[_ -]?for[_ -]?byte)",
        re.IGNORECASE,
    ),
    "conflict and fail-closed refusal": re.compile(
        r"(?:stale|foreign|ambiguous|conflict|unsupported|malformed)",
        re.IGNORECASE,
    ),
    "selector/locality verification": re.compile(
        r"(?:SheetSelector|TableSelector|(?-i:\bAt\b)|locality|reopen|readback)",
        re.IGNORECASE,
    ),
}
IWA_NUMBERS_TABLE_CELL_FRACTION_FORMAT_FOCUSED_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:litchi_numbers[ \t\r\n]*::[ \t\r\n]*Package|"
    r"FocusedNumbersPackage|focused_fraction_format_package|"
    r"focused_number_format_package|focused_numbers_package)[^;{}]*\.[ \t\r\n]*"
    r"(?:table_cell_fraction_format|edit_table_cell_fraction_format|"
    r"apply_table_cell_fraction_format)\s*\("
)
IWA_NUMBERS_TABLE_CELL_FRACTION_FORMAT_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?P<method>table_cell_fraction_format|"
    r"set_table_cell_fraction_format|reset_table_cell_fraction_format)"
    r"(?![A-Za-z0-9_])[ \t\r\n]*\("
)
IWA_NUMBERS_TABLE_CELL_FRACTION_FORMAT_README_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor|editor)"
    r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?(?P<method>"
    r"table_cell_fraction_format|set_table_cell_fraction_format|"
    r"reset_table_cell_fraction_format)\b[ \t\r\n]*\("
)


# Fraction is the next focused decimal family.
def _numbers_table_cell_fraction_format_owner_present(root: Path) -> bool:
    """Return whether the independent Fraction package owner is active."""

    owner_path = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and NUMBERS_PACKAGE_TABLE_CELL_FRACTION_FORMAT_MODULE.search(package_source)
        is not None
    )


def _numbers_table_cell_fraction_format_semantic_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return the canonical Fraction leaf or its explicit split fallback."""

    canonical = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_SOURCE
    if canonical.is_file():
        return (canonical,)
    fallback = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_FALLBACK_SOURCE
    return (fallback,) if fallback.is_file() else ()


def _numbers_table_cell_fraction_format_transaction_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return sources that can own the Fraction transaction namespace."""

    candidates = (
        NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_SOURCE,
        NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_SOURCE,
    )
    return tuple(root / path for path in candidates if (root / path).is_file())


def _numbers_table_cell_fraction_format_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden by the Fraction facade."""

    if identifier in {
        "Draft",
        "Draft",
        "Draft",
        "LitchiIwaPackage",
    }:
        return "monolithic host type"
    if identifier in NUMBERS_TABLE_CELL_FRACTION_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_FRACTION_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_FRACTION_FORMAT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    if identifier in {"Projection", "LitchiIwaProjection"}:
        return "generated type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_fraction_format_owner_declaration(
    declaration: str,
) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_FRACTION_FORMAT_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_fraction_format_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_CELL_FRACTION_FORMAT_FLAT_ALIASES
            | {"Fraction", "fraction", "fraction_format"}
        )
    ) or any(
        NUMBERS_TABLE_CELL_FRACTION_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(identifier)
        for identifier in identifiers
    ) or _numbers_table_cell_fraction_format_owner_declaration(declaration)


# Both focused owners must use this private native seam.  The alternatives
# accommodate the first generalized implementation while keeping the
# contract about ownership, not a particular spelling.  All listed writer /
# scanner names are implementation names: wrappers may call or re-export
# them, but only the shared core may define them.
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CORE_APIS = {
    "read": (
        "read_display_format",
        "read_display_format_with_budget",
        "read_display_format_with_limits",
    ),
    "rewrite": (
        "rewrite_display_format",
        "rewrite_display_format_with_budget",
    ),
    "preflight": (
        "preflight_display_format",
        "preflight_copy_on_write_display_format",
        "preflight_copy_on_write_scalar_control",
    ),
    "canonical": (
        "canonical_display_format",
        "display_format_write",
    ),
}
# The Number facade still enters the shared display-format core through the
# existing private control-native module.  That module only re-exports the
# typed Number wrappers; it must not become a second implementation owner.
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_SOURCE = (
    Path("crates/litchi-iwa-table/src/cell_format/control.rs")
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_MODULE = (
    "table_cell_control_native"
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_READ_APIS = frozenset(
    {"read_number_format", "read_number_format_with_budget"}
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_REWRITE_APIS = frozenset(
    {"rewrite_number_format"}
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CORE_WRITER_DEFINITIONS = re.compile(
    r"(?m)^[ \t]*(?:(?:pub|pub\([^()]*\)|pub\([^()]*\))[ \t]+)?"
    r"(?:async[ \t]+|const[ \t]+|unsafe[ \t]+)*fn[ \t]+"
    r"(?:[A-Za-z0-9_]+_)?(?:rewrite|canonical|prepare|emit|scan|decode|read|write)_"
    r"(?:display|number|percentage|currency|scientific|fraction)_format(?:_with_budget|_write)?\b|"
    r"(?:[A-Za-z0-9_]+_)?(?:number|percentage|currency|scientific)_from_[a-z_]+\b"
)
NUMBERS_TABLE_CELL_DISPLAY_FORMAT_OWNER_LOCAL_GRAPH_MARKERS = (
    "Archive",
    "RawMessage",
    "BncCell",
    "BncCellView",
    "FormatEntry",
    "format_table",
    "registry",
    "refcount",
    "ref_count",
    "reassembly",
    "SnappyStream",
    "decode_lazy_view",
    "WireView",
)
IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_FOCUSED_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:litchi_numbers[ \t\r\n]*::[ \t\r\n]*Package|"
    r"FocusedNumbersPackage|focused_number_format_package|"
    r"focused_numbers_package)[^;{}]*\.[ \t\r\n]*(?:table_cell_number_format|"
    r"edit_table_cell_number_format|apply_table_cell_number_format)\s*\("
)

# The focused Custom owner now replaces the three Draft convenience
# methods.  Keep this inventory separate from generic ``DataFormat`` helper
# names: source-built Numbers packages and attached Pages/Keynote tables still
# use those compatibility routes.  The shared retirement scanner also masks
# comments, literals, and cfg-only items, while allowing only a lexically bound
# focused ``litchi-numbers::Package`` call with the same method spelling.
RETIRED_IWA_NUMBERS_TABLE_CELL_CUSTOM_FORMAT_METHODS = (
    "table_cell_custom_format",
    "set_table_cell_custom_format",
    "reset_table_cell_custom_format",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_CUSTOM_FORMAT_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_CELL_CUSTOM_FORMAT_METHODS
)
IWA_NUMBERS_TABLE_CELL_NUMBER_FORMAT_README_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor|editor)"
    r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?(?P<method>"
    r"table_cell_number_format|set_table_cell_number_format|"
    r"reset_table_cell_number_format)\b[ \t\r\n]*\("
)
NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_cell_control_codec.rs"
)
NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_cell_control_codec"
)
NUMBERS_TABLE_CELL_CONTROL_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_control.rs"
)
NUMBERS_TABLE_CELL_CONTROL_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_control"
)
NUMBERS_TABLE_CELL_CONTROL_RAW_PARAMETER = re.compile(
    r"\b(?:table_id|row|column|object_id|native_id|model_id|control_spec_id|"
    r"popup_model_id|cell_spec_id|source_bytes|bytes)\b[ \t\r\n]*:\s*"
    r"(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])\b"
)
NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/table/dimension.rs"
)
NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/table/dimension/transaction.rs"
)
NUMBERS_TABLE_DIMENSION_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_dimension.rs"
)
NUMBERS_TABLE_DIMENSION_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package/table_dimension"
)
NUMBERS_TABLE_DIMENSION_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE,
    NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE,
    NUMBERS_TABLE_DIMENSION_OWNER_SOURCE,
)
NUMBERS_TABLE_DIMENSION_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/table.rs"),
    NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE,
)
NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES = (
    "Dimension",
    "Points",
    "Size",
)
NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Path",
    "LimitKind",
    "Error",
)
NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPE_SET = frozenset(
    NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES
)
NUMBERS_TABLE_DIMENSION_PACKAGE_METHODS = (
    "table_dimension_size",
    "edit_table_dimension_size",
    "apply_table_dimension_size",
)
NUMBERS_TABLE_DIMENSION_PACKAGE_METHOD_SET = frozenset(
    NUMBERS_TABLE_DIMENSION_PACKAGE_METHODS
)
NUMBERS_TABLE_DIMENSION_FLAT_ALIAS_PREFIXES = (
    "Dimension",
    "DimensionSize",
    "TableDimension",
    "TableDimensionSize",
)
NUMBERS_TABLE_DIMENSION_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_DIMENSION_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES
)
NUMBERS_TABLE_DIMENSION_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_dimension|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?dimension)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_DIMENSION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_dimension\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_DIMENSION_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_dimension\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_DIMENSION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?dimension\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_DIMENSION_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_DIMENSION_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableDimensionSnapshot",
        "TableInfoSnapshot",
    }
)
NUMBERS_TABLE_DIMENSION_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_DIMENSION_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tswp"}
)
NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/table/cells.rs"
)
NUMBERS_TABLE_CELLS_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cells.rs"
)
NUMBERS_TABLE_CELLS_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE,
    NUMBERS_TABLE_CELLS_OWNER_SOURCE,
)
NUMBERS_TABLE_CELLS_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/table.rs"),
)
NUMBERS_TABLE_CELLS_CANONICAL_TYPES = (
    "State",
    "Storage",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELLS_SHORT_NAMES = frozenset(
    NUMBERS_TABLE_CELLS_CANONICAL_TYPES
)
NUMBERS_TABLE_CELLS_PACKAGE_METHODS = (
    "table_cell",
    "table_cells",
)
NUMBERS_TABLE_CELLS_FLAT_ALIAS_PREFIXES = (
    "Cell",
    "Cells",
    "TableCell",
    "TableCells",
)
NUMBERS_TABLE_CELLS_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_CELLS_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_CELLS_SHORT_NAMES
)
NUMBERS_TABLE_CELLS_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cells|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?cells)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELLS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_cells\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELLS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_cells\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_CELLS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?cells\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELLS_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tst", "tswp"}
)
NUMBERS_TABLE_CELLS_WIRE_TYPES = frozenset(
    {
        "BncCell",
        "BncCellView",
        "DecodeOptions",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_FORMULA_SEMANTIC_SOURCE = Path("crates/litchi-numbers/src/formula.rs")
NUMBERS_FORMULA_PUBLIC_API_SOURCES = (
    NUMBERS_FORMULA_SEMANTIC_SOURCE,
    Path("crates/litchi-numbers/src/lib.rs"),
    Path("crates/litchi-numbers/src/package.rs"),
    Path("crates/litchi-numbers/src/package/table_cells.rs"),
    Path("crates/litchi-numbers/src/table/cells.rs"),
    Path("crates/litchi-numbers/src/package/table_cell_edit.rs"),
)
NUMBERS_FORMULA_CANONICAL_TYPES = (
    "Expression",
    "CachedValue",
    "CellReference",
    "AxisReference",
    "BinaryOperator",
    "Table",
    "Error",
    "LimitKind",
)
NUMBERS_FORMULA_CANONICAL_TYPE_SET = frozenset(NUMBERS_FORMULA_CANONICAL_TYPES)
NUMBERS_FORMULA_TRANSACTION_LIMIT_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cells.rs"
)
NUMBERS_FORMULA_TRANSACTION_LIMIT_VARIANTS = ("WireBytes", "AllocationEvents")
RETIRED_NUMBERS_FORMULA_FACADE_TYPES = frozenset(
    {
        "FormulaExpression",
        "FormulaCachedValue",
        "FormulaCellReference",
        "FormulaAxisReference",
        "FormulaBinaryOperator",
        "FormulaPivotCategoryReference",
        "FormulaUuid",
    }
)
NUMBERS_FORMULA_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tst", "tswp"}
)
NUMBERS_FORMULA_WIRE_TYPES = frozenset(
    {
        "BncCell",
        "BncCellView",
        "DecodeOptions",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_FORMULA_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?formula[ \t\r\n]*::"
)
PUBLIC_NUMBERS_FORMULA_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?formula\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_CELLS_MUTATION_TYPES = (
    "Input",
    "Change",
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "DependencyKind",
)
NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES = (
    "Input",
    "Change",
    "State",
    "Storage",
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
    "DependencyKind",
)
NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS = (
    "edit_table_cells",
    "apply_table_cells",
)
NUMBERS_TABLE_CELLS_FULL_PACKAGE_METHODS = (
    *NUMBERS_TABLE_CELLS_PACKAGE_METHODS,
    *NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS,
)
NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_cell_edit.rs"
)
NUMBERS_TABLE_CELLS_MUTATION_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE,
    NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE,
)
NUMBERS_TABLE_CELLS_FULL_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_CELLS_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES
)
NUMBERS_TABLE_CELLS_MUTATION_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_cell_edit|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?cells)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_EDIT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_cell_edit\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_CELL_EDIT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_cell_edit\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_CHART_ARRANGEMENT_PACKAGE_METHODS = (
    "sheet_chart_arrangement",
    "sheet_chart_arrangements",
    "edit_sheet_chart_arrangement",
    "apply_sheet_chart_arrangement",
)
NUMBERS_CHART_ARRANGEMENT_SELECTOR_TYPES = ("SheetSelector", "ChartSelector")
NUMBERS_CHART_ARRANGEMENT_OWNER_MARKERS = (
    "select_chart_with_budget",
    "resolve_sheet_position",
    "verify_locality",
)
NUMBERS_CHART_ARRANGEMENT_LAZY_CODEC_MARKERS = (
    "chart_arrangement_codec",
    "decode_chart_arrangement_with_report",
    "prepare_chart_arrangement_rewrite",
    "execution_requirements",
    "execute",
)
NUMBERS_INDEX_SOURCE = NUMBERS_SOURCE_ROOT / "package" / "index.rs"
NUMBERS_CORE_ARCHIVE_SOURCE = Path("crates/litchi-iwa-core/src/archive.rs")
NUMBERS_PACKAGE_MANIFEST = Path("crates/litchi-numbers/Cargo.toml")
NUMBERS_NAMES_PACKAGE_SOURCE = NUMBERS_SOURCE_ROOT / "package" / "names.rs"
NUMBERS_INDEX_SEMANTIC_ALIAS_WORD = re.compile(
    r"\b(?:alias\w*|coalesc\w*|dedup\w*|equivalent|unique|canonical)\b",
    re.IGNORECASE,
)
NUMBERS_INDEX_RAW_ID_PUBLIC_METHOD = re.compile(
    r"\bfn[ \t\r\n]+(?:r#)?(?:object_id|object_identifier|native_id|"
    r"native_object(?:_id|_identifier)?|resolve_ref_id|resolve_object(?:_id)?|"
    r"raw_object(?:_id|_identifier)|object_locator)\b"
)
NUMBERS_PACKAGE_NO_EAGER_PROST_SOURCE_PATTERNS = (
    (
        "DocumentArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])DocumentArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
    (
        "SheetArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])SheetArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
    (
        "FormBasedSheetArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])FormBasedSheetArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
    (
        "StorageArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])StorageArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
)
NUMBERS_EXTRACTOR_NO_EAGER_RICH_TEXT_SOURCE_PATTERNS = (
    (
        "RichTextPayloadArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])RichTextPayloadArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
)
NUMBERS_EXTRACTOR_NO_EAGER_TILE_SOURCE_PATTERNS = (
    (
        "Tile::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])Tile[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "TileRowInfo::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])TileRowInfo[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
)
NUMBERS_EXTRACTOR_NO_EAGER_TABLE_DATA_LIST_SOURCE_PATTERNS = (
    (
        "TableDataList::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?:r#)?TableDataList[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
    (
        "TableDataListSegment::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?:r#)?TableDataListSegment[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
)
NUMBERS_COMMENT_SOURCE = Path("crates/litchi-numbers/src/package/comments.rs")
# Direct table-cell comment replies have a read-only, selector-first
# projection.  Keep this ratchet dormant until the canonical semantic
# type is present: the existing root-comment owner and the raw-ID
# Draft reply mutation/identity APIs remain compatibility surfaces.
NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE = NUMBERS_COMMENT_SOURCE
NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/comment_storage_codec.rs"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_ACTIVATION = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?CommentStorageReplyRewrite(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_REQUIRED_APIS = (
    "CommentStorageReplyRewrite",
    "prepare_comment_storage_reply_rewrite",
    "PreparedCommentStorageReplyRewrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "prepare_report",
    "execution_requirements",
    "execute",
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_REQUIRED_MARKERS = {
    "strict known-field/reference validation": re.compile(
        r"(?i)\b(?:canonical|duplicate|wrong[_ -]?wire|noncanonical|strict)\b"
    ),
    "unknown/raw preservation": re.compile(
        r"(?i)\b(?:unknown[_ -]?(?:field|scalar|group)|raw[_ -]?"
        r"(?:field|byte|framing|wire)|source[_ -]?order|preserv(?:e|ation)|"
        r"balanced[_ -]?group)\b"
    ),
    "resource/preflight accounting": re.compile(
        r"(?i)\b(?:max[_ -]?(?:input|output|fields|work|depth|references|text)|"
        r"try[_ -]?reserve|preflight|scratch|retained|alloc(?:ation|ations)?|"
        r"resource)\b"
    ),
}
NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_TEST_MARKERS = {
    "direct cfg(test) coverage": re.compile(r"#[ \t\r\n]*\[[ \t]*test\b"),
    "prepared replay calls": re.compile(
        r"(?=[\s\S]*\bprepare_comment_storage_reply_rewrite\b)"
        r"(?=[\s\S]*\bprepare_report\b)"
        r"(?=[\s\S]*\bexecution_requirements\b)"
        r"(?=[\s\S]*\bexecute\b)"
    ),
}
NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES = (
    Path("crates/litchi-numbers/src/lib.rs"),
    NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE,
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_CANONICAL_TYPE = "CommentReply"
NUMBERS_TABLE_CELL_COMMENT_REPLY_PUBLIC_ALIAS = "TableCellCommentReply"
NUMBERS_TABLE_CELL_COMMENT_REPLY_PACKAGE_METHODS = (
    "table_cell_comment_replies",
    "table_cell_comment_replies_a1",
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_ACTIVATION = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?CommentReply(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REQUIRED_MARKERS = {
    "strict comment-storage codec visitor": re.compile(
        r"(?<![A-Za-z0-9_])(?:comment_storage_codec|"
        r"decode_comment_storage_archive_with_visitor|"
        r"CommentStorageVisitor)(?![A-Za-z0-9_])"
    ),
    "bounded decode options": re.compile(
        r"(?<![A-Za-z0-9_])(?:DecodeOptions|DecodeLimit|"
        r"max_(?:message_)?bytes|max_(?:wire_)?fields|max_references)(?![A-Za-z0-9_])"
    ),
    "streamed reply cardinality": re.compile(
        r"(?<![A-Za-z0-9_])(?:visit_reply|reply_references|"
        r"report[ \t\r\n]*\.[ \t\r\n]*(?:replies|reply_references))(?![A-Za-z0-9_])"
    ),
    "strict borrowed snapshot": re.compile(
        r"(?<![A-Za-z0-9_])(?:CommentStorageSnapshot|"
        r"decode_comment_storage_archive_with_report|borrowed|source_order)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}
NUMBERS_TABLE_CELL_COMMENT_REPLY_CENSUS_REQUIRED_MARKERS = {
    "ownership census type": re.compile(
        r"(?<![A-Za-z0-9_])(?:CommentOwnershipCensus|"
        r"CommentReplyOwnershipCensus)(?![A-Za-z0-9_])"
    ),
    "global ownership census": re.compile(
        r"(?<![A-Za-z0-9_])(?:census_comment_ownership|"
        r"census_comment_reply_ownership|prove_global_comment_ownership)(?![A-Za-z0-9_])"
    ),
    "table/list/cell references": re.compile(
        r"(?<![A-Za-z0-9_])(?:table_references|cell_comment_keys|"
        r"list_entries|comment_list_ids)(?![A-Za-z0-9_])"
    ),
    "storage/reply identities": re.compile(
        r"(?<![A-Za-z0-9_])(?:storages|reply_ids|entry_storage_ids|"
        r"storage_occurrences)(?![A-Za-z0-9_])"
    ),
    "author/UUID identities": re.compile(
        r"(?<![A-Za-z0-9_])(?:author_ids|storage_uuid|uuids|"
        r"object_uuid)(?![A-Za-z0-9_])"
    ),
    "duplicate/alias rejection": re.compile(
        r"(?<![A-Za-z0-9_])(?:census_alias_checks|duplicate|"
        r"alias|unique)(?![A-Za-z0-9_])",
        re.IGNORECASE,
    ),
}
NUMBERS_TABLE_CELL_COMMENT_REPLY_PUBLIC_LEAKS = frozenset(
    {
        "u64",
        "StorageId",
        "AuthorId",
        "Uuid",
        "RawMessage",
        "IWA",
        "Wire",
        "Prost",
        "Buffa",
        "bytes",
        "Archive",
        "ArchiveInfo",
        "ArchiveObject",
        "Component",
        "DataStore",
        "FieldInfo",
        "MessageInfo",
        "ObjectId",
        "ObjectIdentifier",
        "PackageMetadata",
        "RawBytes",
        "RawMessage",
        "StorageObject",
    }
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_PUBLIC_LEAK_PREFIXES = (
    "Wire",
    "Prost",
    "Buffa",
    "Iwa",
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_TYPES = frozenset(
    {"CommentAuthor", "CommentTimestamp"}
)
# ADR 0034: no prefixed root aliases; the values are
# published by `litchi_numbers::cell::comment`.
NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_ROOT_ALIASES: dict[str, str] = {}
NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_ALLOWED_IDENTIFIERS = frozenset(
    {"public_id"}
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_NATIVE_IDENTIFIERS = frozenset(
    {
        "Archive",
        "ArchiveInfo",
        "ArchiveObject",
        "Component",
        "DataStore",
        "FieldInfo",
        "MessageInfo",
        "ObjectId",
        "ObjectIdentifier",
        "PackageMetadata",
        "RawBytes",
        "RawMessage",
        "StorageId",
        "StorageObject",
        "Uuid",
        "UUID",
        "author_id",
        "object_id",
        "object_identifier",
        "raw_bytes",
        "raw_message",
        "storage_id",
        "storage_uuid",
        "wire",
    }
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_DEBUG_IMPL = re.compile(
    r"(?<![A-Za-z0-9_#])impl(?:[ \t]+<[^>{}]*>)?[ \t]+"
    r"(?:std[ \t]*::[ \t]*|core[ \t]*::[ \t]*)?"
    r"(?:fmt[ \t]*::[ \t]*)?Debug[ \t]+for[ \t]+"
    r"(?:r#)?CommentAuthor(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_A1_ADDRESS = re.compile(
    r"(?<![A-Za-z0-9_])(?:address|a1)(?![A-Za-z0-9_])|"
    r"&[ \t\r\n]*str\b"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_DEBUG_IMPL = re.compile(
    r"(?<![A-Za-z0-9_])impl(?:[ \t\r\n]+<[^>{}]*>)?[ \t\r\n]+"
    r"(?:std[ \t\r\n]*::[ \t\r\n]*|core[ \t\r\n]*::[ \t\r\n]*)?"
    r"(?:fmt[ \t\r\n]*::[ \t\r\n]*)?Debug[ \t\r\n]+for[ \t\r\n]+"
    r"(?:r#)?CommentReply(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_DEBUG_REDACTION = re.compile(
    r"(?i)(?:redact(?:ed|ion)?|omitted|hidden|finish_non_exhaustive|"
    r"authored[ \t\r\n_-]*(?:content|text)|content[ \t\r\n_-]*redact)"
)

# The reply mutation owner owns mutation of an existing cell-comment reply
# thread.  The read projection above intentionally remains a separate,
# already-landed contract: seeing ``CommentReply`` alone must not retire the
# compatibility host or activate these stronger checks.  Activation therefore
# requires one of the mutation-only names below in non-test production source.
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE = Path(
    "crates/litchi-numbers/src/cell/comment.rs"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TRANSACTION_SOURCE = Path(
    "crates/litchi-numbers/src/cell/comment/transaction.rs"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE = (
    NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_HELPER_ROOT = Path(
    "crates/litchi-numbers/src/package"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_SHARED_ENGINE_ROOT = Path(
    "crates/litchi-iwa-table/src/comments_native"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_SOURCE = (
    NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/comment_storage_reply_codec.rs"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/comment_storage_reply_codec"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_cell_comment_reply.rs"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_cell_comment_reply"
)
# ADR 0034: the reply family keeps its focused
# `cell::comment::transaction` module and short canonical names; the
# prefixed root aliases are deleted.
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TYPES = (
    "CommentReplyIndex",
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_INDEX_METHODS = (
    "new",
    "index",
    "get",
    "try_from_usize",
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ROOT_ALIASES: tuple[str, ...] = ()
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_PACKAGE_METHODS = (
    "table_cell_comment_reply",
    "table_cell_comment_reply_a1",
    "edit_table_cell_comment_replies",
    "edit_table_cell_comment_replies_a1",
    "set_table_cell_comment_reply",
    "set_table_cell_comment_reply_a1",
    "remove_table_cell_comment_reply",
    "remove_table_cell_comment_reply_a1",
    "add_table_cell_comment_reply",
    "add_table_cell_comment_reply_a1",
    "apply_table_cell_comment_reply",
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ACTIVATION = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:edit_table_cell_comment_replies|"
    r"set_table_cell_comment_reply|"
    r"remove_table_cell_comment_reply|add_table_cell_comment_reply|"
    r"apply_table_cell_comment_reply)(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_REQUIRED_MARKERS = {
    "rooted selector graph": re.compile(
        r"(?i)\b(?:rooted|resolve[_ -]?(?:comment|reply)|SheetSelector|"
        r"TableSelector|(?-i:At)|CommentReplyIndex)\b"
    ),
    "comment list and cell references": re.compile(
        r"(?i)\b(?:CommentList|comment[_ -]?list[_ -]?ids|cell[_ -]?comment[_ -]?keys|"
        r"list[_ -]?entries|entry[_ -]?storage[_ -]?ids|table[_ -]?references|"
        r"comment[_ -]?storage)\b"
    ),
    "global ownership census": re.compile(
        r"(?i)\b(?:CommentOwnershipCensus|CommentReplyOwnershipCensus|"
        r"census[_ -]?comment(?:[_ -]?reply)?[_ -]?ownership|global|"
        r"all[_ -]?components|ArchiveReference)\b"
    ),
    "storage/reply identity census": re.compile(
        r"(?i)\b(?:storages|reply[_ -]?ids|storage[_ -]?uuid|author[_ -]?ids|"
        r"storage[_ -]?occurrences|refcount|reference[_ -]?count)\b"
    ),
    "copy-on-write/shared-thread lifecycle": re.compile(
        r"(?i)\b(?:copy[_ -]?on[_ -]?write|cow|shared|clone|new[_ -]?storage|"
        r"cull)\b"
    ),
    "duplicate/alias/dependency rejection": re.compile(
        r"(?i)\b(?:duplicate|alias|cycle|depth|invalid|reject|dependency)\b"
    ),
    "ArchiveInfo authority": re.compile(
        r"(?i)\b(?:ArchiveInfo|MessageInfo|FieldInfo|object[_ -]?references|"
        r"aggregate|field[_ -]?infos|preserve)\b"
    ),
    "metadata ownership": re.compile(
        r"(?i)\b(?:metadata|uuid|object[_ -]?uuid|save[_ -]?token|watermark|"
        r"versioned|ambiguous|data[_ -]?owner|root[_ -]?map|external)\b"
    ),
    "atomic candidate/inverse locality": re.compile(
        r"(?i)\b(?:atomic|source(?:[_ -]?bytes)?|inverse|Patch|candidate|reopen|"
        r"locality|same[_ -]?content|previews|publication)\b"
    ),
    "aggregate transaction budget": re.compile(
        r"(?i)\b(?:TransactionBudget|WireBudget|budget|residual|preflight|charge|"
        r"max[_ -]?(?:input|output)|fields|work|depth|references|alloc(?:ation|ations)?|"
        r"scratch|retained|reassembly|Snappy|ZIP|try[_ -]?reserve)\b"
    ),
    "prepared strict codec route": re.compile(
        r"(?i)\b(?:CommentStorageReplyRewrite|prepare_comment_storage_reply_rewrite|"
        r"CommentStorage|visitor|execution[_ -]?requirements|execute)\b"
    ),
}
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_PUBLIC_LEAKS = frozenset(
    {
        "u64",
        "usize",
        "StorageId",
        "AuthorId",
        "ReplyStorageId",
        "ObjectId",
        "ObjectIdentifier",
        "Uuid",
        "RawMessage",
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "IWA",
        "Wire",
        "WireView",
        "WireLimits",
        "MessageInfo",
        "FieldInfo",
        "ObjectReference",
        "SourceCatalog",
        "Prost",
        "Buffa",
        "bytes",
    }
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_PUBLIC_LEAK_PREFIXES = (
    "Wire",
    "Prost",
    "Buffa",
    "Archive",
    "SourceCatalog",
    "CommentStorage",
    "ReplyStorage",
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_DEBUG_SENSITIVE = re.compile(
    r"(?i)\b(?:text|content|body|message|author)\b"
)
NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_A1_ADDRESS = re.compile(
    r"(?<![A-Za-z0-9_])(?:address|a1)(?![A-Za-z0-9_])|"
    r"&[ \t\r\n]*str\b"
)
NUMBERS_COMMENT_CLEAR_FUNCTION = "clear_table_cell_comment"
NUMBERS_COMMENT_COMMIT_FUNCTION = "commit_edit"
NUMBERS_COMMENT_CLEAR_UNSUPPORTED_ERROR = "UnsupportedDependency"
NUMBERS_EXTRACTOR_NO_EAGER_COMMENT_STORAGE_SOURCE_PATTERNS = (
    (
        "CommentStorageArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])CommentStorageArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
)
NUMBERS_EXTRACTOR_NO_EAGER_FORMULA_SOURCE_PATTERNS = (
    (
        "FormulaArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])FormulaArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated FormulaArchive type",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?:litchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*)?"
            r"tsce[ \t\r\n]*::[ \t\r\n]*FormulaArchive\b"
            r"(?![ \t\r\n]*::[ \t\r\n]*decode\b)"
        ),
    ),
)
RETIRED_IWA_NUMBERS_TABLE_READER_MODULES = frozenset(
    {"formula_renderer", "table", "table_extractor"}
)
# The generated-free formula event visitor has a low-level home in
# ``litchi-numbers-wire``.  The wire crate is allowed to depend on the
# generated Numbers codec because it is the codec's immediate adapter; format
# readers must consume this semantic owner instead of carrying another event
# visitor or scalar expression walk.  Keep this ratchet independent from the
# legacy host renderer audit above: the host still owns archive admission and
# decoder-report charging while the shared visitor migrates out.
IWA_NUMBERS_WIRE_FORMULA_RENDER_SOURCE = Path(
    "crates/litchi-numbers-wire/src/formula_render.rs"
)
IWA_NUMBERS_WIRE_FORMULA_RENDER_REQUIRED_MARKERS = (
    "FormulaEventRenderBudget",
    "ReferenceResolver",
    "render_scalar_formula_nodes",
    "CompatibilityFormulaVisitor",
    "FormulaRenderCodecVisitor",
    "FormulaRenderVisitor",
)
IWA_NUMBERS_WIRE_FORMULA_RENDER_COMMON_IMPORT = re.compile(
    r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*formula"
    r"[ \t\r\n]*::[ \t\r\n]*render\b"
)
IWA_NUMBERS_WIRE_FORMULA_RENDER_PROTOS_IMPORT = re.compile(
    r"\blitchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*"
    r"numbers_formula_codec\b"
)
# The Numbers cell-storage value envelope has one low-level semantic owner.
# The owner returns only scalar values and unresolved table references; it
# borrows the input bytes and leaves sidecar lookup to each format adapter.
# Keep this separate from the existing BNC/pre-BNC view audits: those views are
# still valid wire-owned building blocks, while this module owns their shared
# version dispatch and value interpretation.
IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE = Path(
    "crates/litchi-numbers-wire/src/cell_value.rs"
)
IWA_NUMBERS_WIRE_CELL_VALUE_LIB_SOURCE = Path(
    "crates/litchi-numbers-wire/src/lib.rs"
)
IWA_NUMBERS_WIRE_CELL_VALUE_MODULE_EXPORT = re.compile(
    r"(?m)^[ \t]*pub[ \t]+mod[ \t]+(?:r#)?cell_value[ \t]*;"
)
IWA_NUMBERS_WIRE_CELL_VALUE_PUBLIC_DECODER = re.compile(
    r"(?ms)^[ \t]*pub[ \t]+fn[ \t]+(?:r#)?decode_cell_value\b"
    r"[ \t\r\n]*\([^)]*&[ \t\r\n]*\[[ \t\r\n]*u8[ \t\r\n]*\][^)]*\)"
    r"[ \t\r\n]*->[^{;]*\bCellValueSource\b"
)
IWA_NUMBERS_WIRE_CELL_VALUE_PUBLIC_TYPES = (
    (
        "ValueSource",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:enum|struct|union|type)[ \t]+"
            r"(?:r#)?ValueSource\b"
        ),
    ),
    (
        "CellValueSource",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:enum|struct|union|type)[ \t]+"
            r"(?:r#)?CellValueSource\b"
        ),
    ),
)
IWA_NUMBERS_WIRE_CELL_VALUE_PARSE_MARKERS = (
    "BncCellView",
    "PreBncCellView",
)
IWA_NUMBERS_WIRE_CELL_VALUE_IMPORT_STATEMENT = re.compile(
    r"(?ms)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?use\b.*?;"
)
IWA_NUMBERS_WIRE_CELL_VALUE_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:litchi_iwa_protos|prost|buffa|litchi_numbers|"
    r"litchi_pages|litchi_keynote)(?![A-Za-z0-9_])"
)
IWA_NUMBERS_WIRE_CELL_VALUE_OWNED_RESULT_TYPES = re.compile(
    r"\b(?:String|Vec|Box|HashMap|BTreeMap|Arc|Rc|Cow|Bytes)\b"
)
IWA_NUMBERS_WIRE_CELL_VALUE_COPY_DERIVE = re.compile(
    r"(?ms)#\s*\[\s*derive\s*\([^]]*\bCopy\b[^]]*\]\s*"
)
IWA_NUMBERS_CELL_VALUE_READER_SOURCES = (NUMBERS_EXTRACTOR_SOURCE,)
IWA_NUMBERS_CELL_VALUE_READER_IMPORT = re.compile(
    r"(?ms)\buse[ \t\r\n]+litchi_numbers_wire[ \t\r\n]*::"
    r"[^;]*\bcell_value\b[^;]*;"
)
IWA_NUMBERS_CELL_VALUE_READER_CALL = re.compile(
    r"\b(?:r#)?decode_cell_value[ \t\r\n]*\("
)
IWA_NUMBERS_CELL_VALUE_READER_LEGACY_TYPES = (
    (
        "BncCellView",
        re.compile(r"\b(?:r#)?BncCellView\b"),
    ),
    (
        "PreBncCellView",
        re.compile(r"\b(?:r#)?PreBncCellView\b"),
    ),
    (
        "StoredValue",
        re.compile(r"\b(?:r#)?StoredValue\b"),
    ),
    (
        "CachedScalar",
        re.compile(r"\b(?:r#)?CachedScalar\b"),
    ),
)
IWA_NUMBERS_CELL_VALUE_READER_LEGACY_METHODS = (
    (
        "cell_type()",
        re.compile(r"\.[ \t\r\n]*cell_type[ \t\r\n]*\("),
    ),
    (
        "stored_value()",
        re.compile(r"\.[ \t\r\n]*stored_value[ \t\r\n]*\("),
    ),
    (
        "cached_scalar()",
        re.compile(r"\.[ \t\r\n]*cached_scalar[ \t\r\n]*\("),
    ),
)
# Table-data-list root/segment selection has one shared topology owner.  The
# wire module coordinates borrowed message candidates, duplicate roots/segments,
# key-range admission, and sorted publication.  Format readers retain their
# decoder calls, aggregate budgets, and typed sidecar converters; those adapter
# responsibilities are intentionally outside this ratchet.
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE = Path(
    "crates/litchi-numbers-wire/src/table_data_list.rs"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_LIB_SOURCE = Path(
    "crates/litchi-numbers-wire/src/lib.rs"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_MODULE_EXPORT = re.compile(
    r"(?m)^[ \t]*pub[ \t]+mod[ \t]+(?:r#)?table_data_list[ \t]*;"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_TYPES = (
    (
        "Message",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?Message\b"
        ),
    ),
    (
        "RootOrSegment",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?RootOrSegment\b"
        ),
    ),
    (
        "KeyRange",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?KeyRange\b"
        ),
    ),
    (
        "EntryBounds",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?EntryBounds\b"
        ),
    ),
    (
        "Candidate",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?Candidate\b"
        ),
    ),
    (
        "CoordinatorIssue",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?CoordinatorIssue\b"
        ),
    ),
    (
        "ListReadPolicy",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?ListReadPolicy\b"
        ),
    ),
    (
        "AllocationTarget",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?AllocationTarget\b"
        ),
    ),
    (
        "OverflowPolicy",
        re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?OverflowPolicy\b"
        ),
    ),
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_TRAIT = re.compile(
    r"(?m)^[ \t]*pub[ \t]+trait[ \t]+(?:r#)?ListDecoder\b"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_FUNCTION = re.compile(
    r"(?m)^[ \t]*pub[ \t]+(?:const[ \t]+)?fn[ \t]+"
    r"(?:r#)?read_list\b"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_BORROWED_MESSAGE = re.compile(
    r"(?ms)^[ \t]*pub[ \t]+struct[ \t]+(?:r#)?Message[ \t]*<[^>]*>"
    r"[^{]*\{[^{}]*pub[ \t]+data[ \t]*:[ \t]*&[ \t]*['\"]?"
    r"[A-Za-z_][A-Za-z0-9_]*[ \t]*\[[ \t]*u8[ \t]*\]"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_READ_LIST_GENERIC = re.compile(
    r"(?ms)^[ \t]*pub[ \t]+fn[ \t]+(?:r#)?read_list\b"
    r".{0,500}\b(?:Root|Segments|Resolve|Decoder)\b"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_IMPORT_STATEMENT = re.compile(
    r"(?ms)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
    r"(?:use|extern[ \t]+crate)\b.*?;"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:litchi_iwa_protos|prost|buffa|"
    r"litchi_numbers|litchi_pages|litchi_keynote|"
    r"litchi_iwa_core)(?![A-Za-z0-9_])"
)
IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_TOPOLOGY_MARKERS = (
    "TABLE_DATA_LIST_MESSAGE_KIND",
    "NATIVE_TABLE_DATA_LIST_MESSAGE_KIND",
    "TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND",
    "DuplicateRoot",
    "DuplicateSegmentPayload",
    "WrongSegmentType",
    "MissingKeyRange",
    "KeyRangeOverflow",
    "EntryOutsideKeyRange",
    "DuplicateEntryKey",
    "EntryLimit",
    "Allocation",
    "sort_unstable_by_key",
)
IWA_NUMBERS_TABLE_DATA_LIST_READER_SOURCES = (NUMBERS_EXTRACTOR_SOURCE,)
IWA_NUMBERS_TABLE_DATA_LIST_READER_IMPORT = re.compile(
    r"\blitchi_numbers_wire[ \t\r\n]*::[ \t\r\n]*"
    r"table_data_list\b"
)
IWA_NUMBERS_TABLE_DATA_LIST_READER_CALL = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)?"
    r"(?:r#)?read_list[ \t\r\n]*\("
)
IWA_NUMBERS_TABLE_DATA_LIST_ROOT_TOPOLOGY_FILTER = re.compile(
    r"\.filter\([ \t\r\n]*\|[ \t\r\n]*message[ \t\r\n]*\|"
    r"[ \t\r\n]*message[ \t\r\n]*\.[ \t\r\n]*"
    r"(?:type_|kind)[ \t\r\n]*==[ \t\r\n]*"
    r"(?:6005|6_005|TABLE_DATA_LIST_MESSAGE_KIND)"
    r"[ \t\r\n]*\|\|[ \t\r\n]*message[ \t\r\n]*\.[ \t\r\n]*"
    r"(?:type_|kind)[ \t\r\n]*==[ \t\r\n]*"
    r"(?:6201|6_201|NATIVE_TABLE_DATA_LIST_MESSAGE_KIND)"
    r"[ \t\r\n]*\)"
)
IWA_NUMBERS_TABLE_DATA_LIST_SEGMENT_TOPOLOGY_FILTER = re.compile(
    r"\.filter\([ \t\r\n]*\|[ \t\r\n]*message[ \t\r\n]*\|"
    r"[ \t\r\n]*message[ \t\r\n]*\.[ \t\r\n]*"
    r"(?:type_|kind)[ \t\r\n]*==[ \t\r\n]*"
    r"(?:6011|6_011|TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND)"
    r"[ \t\r\n]*\)"
)
IWA_NUMBERS_TABLE_DATA_LIST_SELECTION_STATE = re.compile(
    r"\b(?:selected_values|selected_keys|segment_count)\b"
)
# The strict FormulaArchive envelope walk and schema validator live beside the
# shared formula event adapter.  It remains a generated-code boundary:
# ``litchi-numbers-wire`` may inspect the generated-free codec, while format
# readers retain only their owned byte wrapper, decoder calls, report charging,
# and format-specific error mapping.
IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE = Path(
    "crates/litchi-numbers-wire/src/formula_envelope.rs"
)
IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_REQUIRED_MARKERS = (
    "FormulaEnvelopeLimits",
    "AttemptedFormulaEnvelopeCost",
    "FormulaEnvelopeReport",
    "FormulaEnvelopeFailure",
    "FormulaEnvelopeCost",
    "FormulaEnvelopeField",
    "preflight_formula_envelope",
    "required_formula_ast_node_type",
    "require_formula_fields",
    "formula_field_is_repeated",
    "validate_formula_scalar",
    "formula_envelope_field",
)
IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_COMMON_IMPORT = re.compile(
    r"\blitchi_iwa_common[ \t\r\n]*::"
)
IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_PROTOS_IMPORT = re.compile(
    r"\blitchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*"
    r"numbers_formula_codec\b"
)
IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_PUBLIC_ENTRY = re.compile(
    r"(?m)^[ \t]*pub[ \t]+fn[ \t]+"
    r"(?:r#)?preflight_formula_envelope\b"
)
NUMBERS_NAMES_GENERATED_PROTO_MODULES = ("tn", "tsce", "tst", "tswp")
NUMBERS_NAMES_NO_EAGER_PROST_SOURCE_PATTERNS = (
    (
        "prost::Message",
        re.compile(
            r"(?<![A-Za-z0-9_#])prost[ \t\r\n]*::[ \t\r\n]*Message\b"
        ),
    ),
    (
        "generated-message decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?:"
            r"(?:litchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*)?"
            r"(?:"
            + "|".join(NUMBERS_NAMES_GENERATED_PROTO_MODULES)
            + r")[ \t\r\n]*::[ \t\r\n]*"
            r"[A-Za-z_][A-Za-z0-9_]*"
            r"|M"
            r")[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated-message decode helper",
        re.compile(
            r"(?<![A-Za-z0-9_#])decode_message"
            r"(?:[ \t\r\n]*::)?[ \t\r\n]*(?:<|\()"
        ),
    ),
)
NUMBERS_NAMES_SAVE_TOKEN_ROUTE_FORBIDDEN_PATTERNS = (
    (
        "legacy crate::package_metadata route",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:crate|self|super)[ \t\r\n]*::"
            r"[ \t\r\n]*package_metadata\b"
        ),
    ),
    (
        "legacy advance_package_save_token_for_components route",
        re.compile(
            r"(?<![A-Za-z0-9_#])advance_package_save_token_for_components\b"
        ),
    ),
    (
        "legacy set_package_last_object_identifier route",
        re.compile(
            r"(?<![A-Za-z0-9_#])set_package_last_object_identifier\b"
        ),
    ),
    (
        "generated PackageMetadata::decode route",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:"
            r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*"
            r")*(?:r#)?PackageMetadata[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated ComponentInfo::decode route",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:"
            r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*"
            r")*(?:r#)?ComponentInfo[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated Prost Message::decode route",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:r#)?(?:Message|M)"
            r"[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "prost route",
        re.compile(r"(?<![A-Za-z0-9_#])(?:prost|prost_types)\b"),
    ),
)
NUMBERS_NAMES_SAVE_TOKEN_MODULE_IMPORT = re.compile(
    r"^[ \t]*(?:pub[ \t\r\n]+)?use[ \t\r\n]+"
    r"litchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*"
    r"package_metadata_codec\b(?P<tail>[^;]*);",
    re.MULTILINE,
)
NUMBERS_NAMES_SAVE_TOKEN_IWA_USE = re.compile(
    r"^[ \t]*(?:pub[ \t\r\n]+)?use[ \t\r\n]+"
    r"litchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*(?P<path>[^;]+);",
    re.MULTILINE,
)
NUMBERS_NAMES_SAVE_TOKEN_MODULE_ALIAS = re.compile(
    r"\bas[ \t\r\n]+(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*)\b"
)
NUMBERS_NAMES_SAVE_TOKEN_FORBIDDEN_NORMAL_DEPENDENCIES = frozenset(
    {"prost", "prost-types"}
)
NUMBERS_NAMES_SAVE_TOKEN_REQUIRED_NORMAL_DEPENDENCY = "litchi-iwa-protos"
NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE = NUMBERS_SOURCE_ROOT / "sheet" / "order.rs"
NUMBERS_SHEET_ORDER_OWNER_SOURCE = NUMBERS_SOURCE_ROOT / "package" / "sheet_order.rs"
NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT = NUMBERS_SOURCE_ROOT / "package" / "sheet_order"
NUMBERS_SHEET_ORDER_OWNER_HELPER_SOURCES = (
    NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT / "error.rs",
    NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT / "resolve.rs",
    NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT / "rewrite.rs",
)
NUMBERS_SHEET_ORDER_IMPLEMENTATION_SOURCES = (
    NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE,
    NUMBERS_SHEET_ORDER_OWNER_SOURCE,
    *NUMBERS_SHEET_ORDER_OWNER_HELPER_SOURCES,
)
NUMBERS_SHEET_ORDER_EXPORT_SOURCES = (
    NUMBERS_SOURCE_ROOT / "lib.rs",
    NUMBERS_SOURCE_ROOT / "package.rs",
    NUMBERS_SOURCE_ROOT / "sheet.rs",
)
NUMBERS_SHEET_ORDER_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
NUMBERS_SHEET_ORDER_SHORT_NAMES = frozenset(NUMBERS_SHEET_ORDER_CANONICAL_TYPES)
NUMBERS_SHEET_ORDER_PACKAGE_METHODS = (
    "edit_sheet_order",
    "apply_sheet_order",
)
NUMBERS_SHEET_ORDER_FLAT_ALIASES = frozenset(
    "SheetOrder" + suffix for suffix in NUMBERS_SHEET_ORDER_SHORT_NAMES
)
NUMBERS_SHEET_ORDER_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:sheet_order|sheet[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?order)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_SHEET_ORDER_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?sheet_order\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_SHEET_ORDER_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?sheet_order\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_SHEET_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?sheet\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_SHEET_ORDER_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?order\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_SHEET_ORDER_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "DocumentSnapshot",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SheetOrderSnapshot",
        "SheetSnapshot",
        "SnappyStream",
        "SourceCatalog",
    }
)
NUMBERS_SHEET_ORDER_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_SHEET_ORDER_PROTO_ORIGINS = frozenset({"tn", "tsp"})

NUMBERS_TABLE_SORT_SEMANTIC_SOURCE = NUMBERS_SOURCE_ROOT / "table" / "sort.rs"
NUMBERS_TABLE_SORT_TRANSACTION_SOURCE = (
    NUMBERS_SOURCE_ROOT / "table" / "sort" / "transaction.rs"
)
NUMBERS_TABLE_SORT_OWNER_SOURCE = NUMBERS_SOURCE_ROOT / "package" / "table_sort.rs"
NUMBERS_TABLE_SORT_OWNER_HELPER_ROOT = NUMBERS_SOURCE_ROOT / "package" / "table_sort"
NUMBERS_TABLE_SORT_IMPLEMENTATION_SOURCES = (
    NUMBERS_TABLE_SORT_SEMANTIC_SOURCE,
    NUMBERS_TABLE_SORT_TRANSACTION_SOURCE,
    NUMBERS_TABLE_SORT_OWNER_SOURCE,
)
NUMBERS_TABLE_SORT_EXPORT_SOURCES = (
    NUMBERS_SOURCE_ROOT / "lib.rs",
    NUMBERS_SOURCE_ROOT / "package.rs",
    NUMBERS_SOURCE_ROOT / "table.rs",
)
NUMBERS_TABLE_SORT_SEMANTIC_TYPES = (
    "ColumnIndex",
    "Direction",
    "Order",
    "RowRange",
    "Rule",
    "Scope",
)
NUMBERS_TABLE_SORT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
NUMBERS_TABLE_SORT_CANONICAL_TYPES = (
    *NUMBERS_TABLE_SORT_SEMANTIC_TYPES,
    *NUMBERS_TABLE_SORT_TRANSACTION_TYPES,
)
NUMBERS_TABLE_SORT_SHORT_NAMES = frozenset(NUMBERS_TABLE_SORT_CANONICAL_TYPES)
NUMBERS_TABLE_SORT_PACKAGE_METHODS = (
    "table_sort_order",
    "edit_table_sort_order",
    "apply_table_sort_order",
)
NUMBERS_TABLE_SORT_FLAT_ALIAS_PREFIXES = (
    "Sort",
    "SortOrder",
    "TableSort",
    "TableSortOrder",
)
NUMBERS_TABLE_SORT_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in NUMBERS_TABLE_SORT_FLAT_ALIAS_PREFIXES
    for suffix in NUMBERS_TABLE_SORT_TRANSACTION_TYPES
)
NUMBERS_TABLE_SORT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_sort|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?sort)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_NUMBERS_PACKAGE_TABLE_SORT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_sort\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_PACKAGE_TABLE_SORT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_sort\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_SORT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?sort\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_TABLE_SORT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_SORT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_sort_order_codec.rs"
)
NUMBERS_TABLE_SORT_CODEC_ALTERNATE_SOURCE = Path(
    "crates/litchi-iwa-protos/src/table_sort_order_codec.rs"
)
NUMBERS_TABLE_SORT_CODEC_SOURCES = (
    NUMBERS_TABLE_SORT_CODEC_SOURCE,
    NUMBERS_TABLE_SORT_CODEC_ALTERNATE_SOURCE,
)
NUMBERS_TABLE_SORT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
NUMBERS_TABLE_SORT_CODEC_MODULE = "numbers_table_sort_order_codec"
NUMBERS_TABLE_SORT_CODEC_NEUTRAL_MODULE = "table_sort_order_codec"
NUMBERS_TABLE_SORT_CODEC_MODULES = (
    NUMBERS_TABLE_SORT_CODEC_MODULE,
    NUMBERS_TABLE_SORT_CODEC_NEUTRAL_MODULE,
)
NUMBERS_TABLE_SORT_CODEC_REQUIRED_APIS = (
    "SortOrderSnapshot",
    "decode_table_sort_order",
    "decode_table_sort_order_with_report",
    "prepare_table_sort_order_rewrite",
    "PreparedTableSortOrderRewrite",
    "canonical_table_sort_order",
    "rewrite_table_sort_order",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
)
NUMBERS_TABLE_SORT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_sort_order_codec.rs"
)
NUMBERS_TABLE_SORT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_sort_order_codec"
)
NUMBERS_TABLE_SORT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/numbers_table_sort_order.rs"
)
NUMBERS_TABLE_SORT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/numbers_table_sort_order"
)
NUMBERS_TABLE_SORT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "SortRuleArchive",
        "SortRuleReferenceTrackerArchive",
        "TableModelArchive",
        "TableSortOrderArchive",
        "TableSortOrderSnapshot",
    }
)
NUMBERS_TABLE_SORT_WIRE_TYPES = frozenset(
    {
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
NUMBERS_TABLE_SORT_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tst", "tswp"}
)
NUMBERS_TABLE_SORT_CODEC_REQUIRED_MARKERS = (
    "unknown",
    "unknown overlong",
    "report",
    "SORT_ORDER_FIELD",
    "SORT_RULES_FIELD",
)
NUMBERS_TABLE_SORT_OWNER_REQUIRED_MARKERS = {
    "table-model": re.compile(
        r"(?<![A-Za-z0-9])(?:table[_-]?model|TableModel|sort[_-]?order)(?![A-Za-z0-9])"
    ),
    "field44/field45": re.compile(
        r"(?<![A-Za-z0-9])(?:field44|field45|sort_order|sort_rule_reference_tracker)"
        r"(?![A-Za-z0-9])"
    ),
    "budget": re.compile(
        r"(?<![A-Za-z0-9])(?:budget|WireBudget|TransactionBudget|charge_|preflight)"
        r"(?![A-Za-z0-9])"
    ),
    "locality": re.compile(
        r"(?<![A-Za-z0-9])(?:locality|verify_locality|candidate|reopen|same_content)"
        r"(?![A-Za-z0-9])"
    ),
    "atomic/inverse": re.compile(
        r"(?<![A-Za-z0-9])(?:atomic|inverse|conflict|source_bytes|unchanged)"
        r"(?![A-Za-z0-9])"
    ),
}
NUMBERS_TABLE_SORT_RAW_PARAMETER = re.compile(
    r"\b(?:table_id|object_id|model_id|native_id|source_bytes|bytes|row|column)"
    r"\b[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])\b"
)

NUMBERS_TABLE_MOVE_OWNER_SOURCE = (
    NUMBERS_SOURCE_ROOT / "package" / "table_relocation.rs"
)
NUMBERS_TABLE_MOVE_EXPORT_SOURCES = (
    NUMBERS_SOURCE_ROOT / "lib.rs",
    NUMBERS_SOURCE_ROOT / "package.rs",
)
NUMBERS_TABLE_MOVE_PACKAGE_METHOD = "move_table"
NUMBERS_TABLE_MOVE_SELECTOR_TYPES = ("SheetSelector", "TableSelector")
NUMBERS_PACKAGE_TABLE_MOVE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_relocation\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_NUMBERS_PACKAGE_TABLE_MOVE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_relocation\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
NUMBERS_TABLE_MOVE_RAW_PARAMETER = re.compile(
    r"\b(?:[A-Za-z_]+_)?(?:sheet|table|sheet_id|table_id|table_object_id|"
    r"object_id|model_id|native_id|raw_object_id|raw_object_identifier|"
    r"object_identifier|native_object_id)"
    r"\b[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u8|u16|u32|u64|u128|usize|i8|i16|i32|i64|i128|isize)\b"
)

NUMBERS_TABLE_LOCK_IMPLEMENTATION_SOURCES = (
    NUMBERS_SOURCE_ROOT / "package" / "table_lock.rs",
    NUMBERS_SOURCE_ROOT / "table" / "lock.rs",
)
NUMBERS_TABLE_LOCK_EXPORT_SOURCES = (
    NUMBERS_SOURCE_ROOT / "lib.rs",
    NUMBERS_SOURCE_ROOT / "package.rs",
    NUMBERS_SOURCE_ROOT / "table.rs",
)
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

IWA_NUMBERS_CELL_COMMENT_CLEAR_LEGACY_HELPER = "clear_cell_comment_in_package"
RETIRED_IWA_NUMBERS_DOCUMENT_TYPES = (
    "NumbersDocument",
    "NumbersDocumentState",
    "NumbersDocumentStats",
)
RETIRED_IWA_NUMBERS_DOCUMENT_TYPE_SET = frozenset(
    RETIRED_IWA_NUMBERS_DOCUMENT_TYPES
)
NUMBERS_DOCUMENT_PUBLIC_API_SOURCES = (
    NUMBERS_SOURCE_ROOT / "document.rs",
    NUMBERS_SOURCE_ROOT / "lib.rs",
)
NUMBERS_DOCUMENT_PUBLIC_MARKERS = frozenset(
    {
        "DEFAULT_MAX_TEXT_BYTES",
        "Document",
        "DocumentError",
        "DocumentLimits",
        "DocumentReadError",
        "DocumentReadLimitKind",
        "DocumentReadOptions",
        "DocumentResult",
        "DocumentSourceLimitKind",
        "DocumentSourceLimits",
        "DocumentSourceLimitsError",
        "DocumentStats",
        "Error",
        "IoKind",
        "Limits",
        "MAX_MATERIALIZED_CELLS",
        "MAX_SHEETS",
        "MAX_TABLES",
        "ReadError",
        "ReadLimitKind",
        "Result",
        "Stats",
    }
)


def _numbers_names_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in the Numbers names API."""

    if identifier in NUMBERS_NAMES_SAVE_TOKEN_FORBIDDEN_PUBLIC_TYPES:
        return "package-metadata codec type"
    normalized = identifier.lower().replace("-", "_")
    if "save_token" in normalized or "savetoken" in normalized:
        return "save-token state"
    if identifier in NUMBERS_NAMES_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_NAMES_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(
        words[index] in {"archive", "component", "entry", "member"}
        and words[index + 1] in {"name", "names"}
        for index in range(len(words) - 1)
    ):
        return "physical package name"
    return None


def _numbers_names_codec_aliases(source: str) -> frozenset[str]:
    """Return local aliases that import the hidden metadata codec module."""

    aliases: set[str] = set()
    masked_source = _mask_rust_non_code(source)
    for import_match in NUMBERS_NAMES_SAVE_TOKEN_IWA_USE.finditer(masked_source):
        imported_path = import_match.group("path")
        if "package_metadata_codec" not in imported_path:
            continue
        alias_match = re.search(
            r"\bpackage_metadata_codec[ \t\r\n]+as[ \t\r\n]+"
            r"(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*)\b",
            imported_path,
        )
        aliases.add(
            alias_match.group("alias")
            if alias_match is not None
            else "package_metadata_codec"
        )
    return frozenset(aliases)


def _numbers_names_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_NAMES_OWNER_PATH.search(declaration) is not None or any(
        identifier in {"apply_names", "edit_names"} for identifier in identifiers
    )


def _is_numbers_names_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & NUMBERS_NAMES_FLAT_ALIASES) or (
        _numbers_names_owner_declaration(declaration)
    )


def _numbers_sheet_order_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in sheet-order APIs."""

    if identifier in NUMBERS_SHEET_ORDER_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_SHEET_ORDER_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_SHEET_ORDER_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(
        words[index] in {"archive", "component", "entry", "member"}
        and words[index + 1] in {"name", "names"}
        for index in range(len(words) - 1)
    ):
        return "physical package name"
    return None


def _numbers_sheet_order_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_SHEET_ORDER_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_SHEET_ORDER_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_sheet_order_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers & (NUMBERS_SHEET_ORDER_FLAT_ALIASES | {"order", "sheet_order"})
    ) or _numbers_sheet_order_owner_declaration(declaration)


def _numbers_table_sort_public_leak(identifier: str) -> str | None:
    """Classify native/wire vocabulary forbidden in persisted sort APIs."""

    if identifier in NUMBERS_TABLE_SORT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_SORT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_SORT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    if any(
        words[index] in {"archive", "component", "entry", "member"}
        and words[index + 1] in {"name", "names"}
        for index in range(len(words) - 1)
    ):
        return "physical package name"
    return _iwork_public_leak(identifier)


def _numbers_table_sort_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_SORT_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_SORT_PACKAGE_METHODS for identifier in identifiers
    )


def _is_numbers_table_sort_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (NUMBERS_TABLE_SORT_FLAT_ALIASES | {"sort", "table_sort"})
    ) or _numbers_table_sort_owner_declaration(declaration)


def _numbers_table_header_settings_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in header transactions."""

    if identifier == "litchi_iwa_common":
        return None
    if identifier in NUMBERS_TABLE_HEADER_SETTINGS_PHYSICAL_TYPES:
        return "archive/IWA type"
    if (
        identifier == "wire"
        or identifier in NUMBERS_TABLE_HEADER_SETTINGS_WIRE_TYPES
    ):
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(
        words[index] in {"archive", "component", "entry", "member"}
        and words[index + 1] in {"name", "names"}
        for index in range(len(words) - 1)
    ):
        return "physical package name"
    return None


def _numbers_table_header_settings_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_HEADER_SETTINGS_OWNER_PATH.search(
        declaration
    ) is not None or any(
        identifier
        in NUMBERS_TABLE_HEADER_SETTINGS_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_header_settings_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIASES
            | NUMBERS_TABLE_HEADER_SETTINGS_FLAT_SEMANTIC_ALIASES
        )
    ) or _numbers_table_header_settings_owner_declaration(declaration)


def _numbers_table_title_settings_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in title transactions."""

    if identifier in NUMBERS_TABLE_TITLE_SETTINGS_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_TITLE_SETTINGS_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_TITLE_SETTINGS_WIRE_TYPES:
        return "wire type"
    # Settings is canonically shared by the generated-free common semantic crate.
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _numbers_table_appearance_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in appearance APIs."""

    if identifier in NUMBERS_TABLE_APPEARANCE_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_APPEARANCE_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_APPEARANCE_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _numbers_table_appearance_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_APPEARANCE_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_APPEARANCE_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_appearance_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers & NUMBERS_TABLE_APPEARANCE_FLAT_ALIASES
    ) or _numbers_table_appearance_owner_declaration(declaration)


def _numbers_table_cell_pop_up_menu_public_leak(identifier: str) -> str | None:
    """Classify native/control-cell vocabulary forbidden in popup APIs."""

    if identifier in NUMBERS_TABLE_CELL_POP_UP_MENU_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_POP_UP_MENU_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_POP_UP_MENU_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_pop_up_menu_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_POP_UP_MENU_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_pop_up_menu_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers & NUMBERS_TABLE_CELL_POP_UP_MENU_FLAT_ALIASES
    ) or _numbers_table_cell_pop_up_menu_owner_declaration(declaration)


def _numbers_table_cell_number_format_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden by the Number-format facade."""

    if identifier in {"Draft", "Draft", "Draft", "LitchiIwaPackage"}:
        return "monolithic host type"
    if identifier in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_NUMBER_FORMAT_WIRE_TYPES:
        return "wire type"
    # A common semantic helper is allowed to remain a dependency of the
    # focused owner.  The public declarations below still reject every
    # physical value imported from it.
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost", "codec"} for word in words):
        return "protobuf type"
    if identifier in {"Projection", "LitchiIwaProjection"}:
        return "generated type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_number_format_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_number_format_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_CELL_NUMBER_FORMAT_FLAT_ALIASES
            | {"Number", "number", "number_format"}
        )
    ) or any(
        NUMBERS_TABLE_CELL_NUMBER_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(identifier)
        for identifier in identifiers
    ) or _numbers_table_cell_number_format_owner_declaration(declaration)


def _numbers_table_title_settings_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return (
        NUMBERS_TABLE_TITLE_SETTINGS_OWNER_PATH.search(declaration) is not None
        or any(
            identifier in NUMBERS_TABLE_TITLE_SETTINGS_PACKAGE_METHODS
            for identifier in identifiers
        )
    )


def _is_numbers_table_title_settings_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & NUMBERS_TABLE_TITLE_SETTINGS_FLAT_ALIASES) or (
        _numbers_table_title_settings_owner_declaration(declaration)
    )


def _numbers_table_dimension_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in dimension transactions."""

    if identifier in NUMBERS_TABLE_DIMENSION_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_DIMENSION_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_DIMENSION_WIRE_TYPES:
        return "wire type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _is_numbers_table_dimension_common_reexport(declaration: str) -> bool:
    """Recognize only Numbers' canonical common-dimension compatibility path."""

    match = re.fullmatch(
        r"\s*pub\s+use\s+litchi_iwa_common\s*::\s*table\s*::\s*dimension"
        r"\s*::\s*\{(?P<body>[^{}]*)\}\s*;?\s*",
        declaration,
        re.DOTALL,
    )
    if match is None:
        return False
    names = tuple(part.strip() for part in match.group("body").split(",") if part.strip())
    return names == ("Dimension", "Error", "Points", "Size")


def _numbers_table_dimension_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return (
        NUMBERS_TABLE_DIMENSION_OWNER_PATH.search(declaration) is not None
        or any(
            identifier in NUMBERS_TABLE_DIMENSION_PACKAGE_METHOD_SET
            for identifier in identifiers
        )
    )


def _is_numbers_table_dimension_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & NUMBERS_TABLE_DIMENSION_FLAT_ALIASES) or (
        _numbers_table_dimension_owner_declaration(declaration)
    )


def _numbers_table_cells_public_leak(identifier: str) -> str | None:
    """Classify physical vocabulary forbidden in table-cell read APIs."""

    if identifier in NUMBERS_TABLE_CELLS_PROTO_ORIGINS:
        return "protobuf type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELLS_WIRE_TYPES:
        return "wire/BNC type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"bnc", "buffa", "codec", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _numbers_table_cells_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELLS_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELLS_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cells_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & NUMBERS_TABLE_CELLS_FLAT_ALIASES) or (
        _numbers_table_cells_owner_declaration(declaration)
    )


def _numbers_table_cells_mutation_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELLS_MUTATION_OWNER_PATH.search(
        declaration
    ) is not None or any(
        identifier in NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cells_mutation_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & NUMBERS_TABLE_CELLS_FULL_FLAT_ALIASES) or (
        _numbers_table_cells_mutation_owner_declaration(declaration)
    )


def _numbers_formula_public_leak(identifier: str) -> str | None:
    """Classify vocabulary forbidden in the focused Numbers formula API."""

    if identifier in RETIRED_NUMBERS_FORMULA_FACADE_TYPES:
        return "retired formula facade type"
    if identifier == "litchi_iwa_common":
        return "shared iWork formula facade"
    if identifier in NUMBERS_FORMULA_PROTO_ORIGINS:
        return "protobuf type"
    if identifier == "wire" or identifier in NUMBERS_FORMULA_WIRE_TYPES:
        return "wire/BNC type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"bnc", "buffa", "codec", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _numbers_formula_owner_declaration(declaration: str) -> bool:
    """Return whether a declaration routes through the canonical formula module."""

    return NUMBERS_FORMULA_OWNER_PATH.search(declaration) is not None


def _is_numbers_formula_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    """Limit non-formula owner scans to declarations that expose formula API."""

    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    words = {
        word.lower()
        for identifier in identifiers
        for part in identifier.split("_")
        for word in CAMEL_CASE_WORD.findall(part)
    }
    return (
        "formula" in words
        or bool(identifiers & RETIRED_NUMBERS_FORMULA_FACADE_TYPES)
    )


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


def audit_numbers_names_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Enforce the nested, archive-free Numbers names transaction API."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_NAMES_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in NUMBERS_NAMES_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []
    package_export = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        for match in PUBLIC_NUMBERS_PACKAGE_NAMES_MODULE.finditer(package_source):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers names public API exposes duplicate "
                "package::names module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        codec_aliases = _numbers_names_codec_aliases(source)
        masked_source = _mask_rust_non_code(source)
        public_codec_module = re.compile(
            r"^[ \t]*pub(?:[ \t\r\n]*\([^()]*\))?[ \t\r\n]+"
            r"mod[ \t\r\n]+(?:r#)?package_metadata_codec\b",
            re.MULTILINE,
        )
        for module_match in public_codec_module.finditer(masked_source):
            module_line = masked_source.count("\n", 0, module_match.start()) + 1
            violations.append(
                "focused litchi-numbers names public API exposes "
                "package-metadata codec reexport/module alias: "
                f"{path.relative_to(root)}:{module_line}"
            )
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for (
            declaration,
            line_number,
            public_declaration,
            complete_source_scope,
        ) in declarations:
            codec_reexport = False
            if public_declaration:
                codec_reexport = bool(
                    re.search(
                        r"(?<![A-Za-z0-9_:#.])litchi_iwa_protos"
                        r"[ \t\r\n]*::[ \t\r\n]*package_metadata_codec\b",
                        declaration,
                    )
                )
                if not codec_reexport:
                    for alias in codec_aliases:
                        if re.search(
                            rf"(?<![A-Za-z0-9_:#.])(?:r#)?"
                            rf"{re.escape(alias)}"
                            r"(?:[ \t\r\n]*::|[ \t\r\n]*[,;])",
                            declaration,
                        ):
                            codec_reexport = True
                            break
            if codec_reexport:
                violations.append(
                    "focused litchi-numbers names public API exposes "
                    "package-metadata codec reexport/module alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if not _is_numbers_names_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ) and not codec_reexport:
                continue
            owner_declaration = _numbers_names_owner_declaration(declaration)
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and declaration_identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-numbers names public API retains "
                    "root aliases via names glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in NUMBERS_NAMES_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-numbers names public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and declaration_identifiers[:2]
                    in (["pub", "type"], ["pub", "use"])
                    and identifier in NUMBERS_NAMES_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-numbers names public API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_names_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-numbers names public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers names public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_numbers_sheet_order_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the canonical nested, archive-free Numbers sheet-order API."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_SHEET_ORDER_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_SHEET_ORDER_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, NUMBERS_SHEET_ORDER_SHORT_NAMES
    )
    for name in NUMBERS_SHEET_ORDER_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-numbers sheet-order public API is missing "
            f"canonical sheet::order type {name}: "
            f"{NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE}"
        )

    lib_path = root / NUMBERS_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_SHEET_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-numbers sheet-order public API is missing "
            "canonical root sheet module: "
            f"{NUMBERS_SOURCE_ROOT / 'lib.rs'}"
        )

    sheet_path = root / NUMBERS_SOURCE_ROOT / "sheet.rs"
    sheet_source = (
        _mask_rust_non_code(sheet_path.read_text(encoding="utf-8"))
        if sheet_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_SHEET_ORDER_MODULE.search(sheet_source) is None:
        violations.append(
            "focused litchi-numbers sheet-order public API is missing "
            "canonical sheet::order module: "
            f"{NUMBERS_SOURCE_ROOT / 'sheet.rs'}"
        )

    package_export = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        if NUMBERS_PACKAGE_SHEET_ORDER_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-numbers sheet-order public API is missing "
                "private package owner module: "
                f"{package_export.relative_to(root)}"
            )
        for match in PUBLIC_NUMBERS_PACKAGE_SHEET_ORDER_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers sheet-order public API exposes duplicate "
                "package::sheet_order module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-numbers sheet-order public API is missing "
            "private package owner module: "
            f"{NUMBERS_SOURCE_ROOT / 'package.rs'}"
        )

    owner_path = root / NUMBERS_SHEET_ORDER_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-numbers sheet-order public API is missing "
            "private package owner source: "
            f"{NUMBERS_SHEET_ORDER_OWNER_SOURCE}"
        )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for (
            declaration,
            line_number,
            public_declaration,
            complete_source_scope,
        ) in declarations:
            if not _is_numbers_sheet_order_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _numbers_sheet_order_owner_declaration(declaration)
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = declaration_identifiers[:2] in (
                ["pub", "type"],
                ["pub", "use"],
            )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and declaration_identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-numbers sheet-order public API retains "
                    "root aliases via sheet::order glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and public_use_or_type
            ):
                violations.append(
                    "focused litchi-numbers sheet-order public API exposes "
                    "public sheet-order owner alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    public_declaration
                    and identifier in NUMBERS_SHEET_ORDER_FLAT_ALIASES
                ):
                    violations.append(
                        "focused litchi-numbers sheet-order public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in NUMBERS_SHEET_ORDER_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-numbers sheet-order public API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_sheet_order_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-numbers sheet-order public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers sheet-order public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def _numbers_table_sort_owner_present(root: Path) -> bool:
    owner_path = root / NUMBERS_TABLE_SORT_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_SORT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and any((root / path).is_file() for path in NUMBERS_TABLE_SORT_CODEC_SOURCES)
        and NUMBERS_PACKAGE_TABLE_SORT_MODULE.search(package_source) is not None
    )


def audit_numbers_table_sort_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free persisted sort owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_sort_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_SORT_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(NUMBERS_TABLE_SORT_SEMANTIC_TYPES)
    )
    for name in NUMBERS_TABLE_SORT_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-numbers table-sort public API is missing "
                f"canonical table::sort type {name}: {NUMBERS_TABLE_SORT_SEMANTIC_SOURCE}"
            )
    if PUBLIC_NUMBERS_TABLE_SORT_TRANSACTION_MODULE.search(
        _mask_rust_non_code(semantic_source)
    ) is None:
        violations.append(
            "focused litchi-numbers table-sort public API is missing canonical "
            f"table::sort::transaction module: {NUMBERS_TABLE_SORT_SEMANTIC_SOURCE}"
        )

    lib_path = root / NUMBERS_TABLE_SORT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-numbers table-sort public API is missing canonical root "
            f"table module: {NUMBERS_TABLE_SORT_EXPORT_SOURCES[0]}"
        )
    table_path = root / NUMBERS_TABLE_SORT_EXPORT_SOURCES[2]
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_SORT_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-numbers table-sort public API is missing canonical "
            f"table::sort module: {NUMBERS_TABLE_SORT_EXPORT_SOURCES[2]}"
        )

    package_path = root / NUMBERS_TABLE_SORT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    if NUMBERS_PACKAGE_TABLE_SORT_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-numbers table-sort public API is missing private package "
            f"owner module: {NUMBERS_TABLE_SORT_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_NUMBERS_PACKAGE_TABLE_SORT_MODULE.finditer(package_source):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-numbers table-sort public API exposes public "
            f"package::table_sort module: {NUMBERS_TABLE_SORT_EXPORT_SOURCES[1]}:{line_number}"
        )

    owner_path = root / NUMBERS_TABLE_SORT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_exports = _rust_canonical_exports(
        owner_source, frozenset(NUMBERS_TABLE_SORT_TRANSACTION_TYPES)
    )
    for name in NUMBERS_TABLE_SORT_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers table-sort public API is missing canonical "
                f"package type {name}: {NUMBERS_TABLE_SORT_OWNER_SOURCE}"
            )
    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_SORT_PACKAGE_METHODS:
        declaration = owner_methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-numbers table-sort public API is missing "
                f"Package method {method}: {NUMBERS_TABLE_SORT_OWNER_SOURCE}"
            )
            continue
        if method != "apply_table_sort_order":
            for selector in ("SheetSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-numbers table-sort Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{NUMBERS_TABLE_SORT_OWNER_SOURCE}"
                    )
    owner_code = _mask_rust_non_code(owner_source)
    for label, marker in NUMBERS_TABLE_SORT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-numbers table-sort owner is missing "
                f"{label} transaction marker: {NUMBERS_TABLE_SORT_OWNER_SOURCE}"
            )

    codec_path = next(
        (root / path for path in NUMBERS_TABLE_SORT_CODEC_SOURCES if (root / path).is_file()),
        None,
    )
    codec_source = (
        _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        if codec_path is not None
        else ""
    )
    if codec_path is None:
        violations.append(
            "focused litchi-numbers table-sort public API is missing strict hidden "
            f"codec source: {NUMBERS_TABLE_SORT_CODEC_SOURCE}"
        )
    else:
        codec_code = _mask_rust_non_code(codec_source)
        for api in NUMBERS_TABLE_SORT_CODEC_REQUIRED_APIS:
            if re.search(
                rf"\b(?:pub\s+)?(?:fn|struct)\s+{re.escape(api)}\b", codec_code
            ) is None:
                violations.append(
                    "focused litchi-numbers table-sort hidden codec is missing "
                    f"strict API {api}: {codec_path.relative_to(root)}"
                )
        for marker in NUMBERS_TABLE_SORT_CODEC_REQUIRED_MARKERS:
            if re.search(re.escape(marker), codec_source) is None:
                violations.append(
                    "focused litchi-numbers table-sort hidden codec is missing "
                    f"strict {marker} marker: {codec_path.relative_to(root)}"
                )
        codec_modules = (
            NUMBERS_TABLE_SORT_CODEC_MODULES
            if codec_path == root / NUMBERS_TABLE_SORT_CODEC_SOURCE
            else (NUMBERS_TABLE_SORT_CODEC_NEUTRAL_MODULE,)
        )
        codec_lib_path = root / NUMBERS_TABLE_SORT_CODEC_PUBLIC_SOURCE
        codec_lib_source = (
            _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
            if codec_lib_path.is_file()
            else ""
        )
        for codec_module in codec_modules:
            if re.search(
                rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
                rf"pub\s+mod\s+{re.escape(codec_module)}\b",
                codec_lib_source,
            ) is None:
                violations.append(
                    "focused litchi-numbers table-sort public API is missing hidden codec "
                    f"module {codec_module}: {NUMBERS_TABLE_SORT_CODEC_PUBLIC_SOURCE}"
                )

    for target in (
        NUMBERS_TABLE_SORT_CODEC_FUZZ_SOURCE,
        NUMBERS_TABLE_SORT_FUZZ_SOURCE,
    ):
        absolute = root / target
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers table-sort boundary is missing fuzz target: "
                f"{target}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-numbers table-sort fuzz target is missing "
                f"fuzz_target! harness: {target}"
            )
    for corpus in (
        NUMBERS_TABLE_SORT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_SORT_FUZZ_CORPUS,
    ):
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-numbers table-sort boundary is missing fuzz corpus: "
                f"{corpus}"
            )

    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_SORT_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / NUMBERS_TABLE_SORT_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_SORT_EXPORT_SOURCES
        if (root / path).is_file()
    }
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            if not _is_numbers_table_sort_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
            if (
                path in export_sources
                and public_use_or_type
                and _numbers_table_sort_owner_declaration(declaration)
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-numbers table-sort public API retains root aliases "
                    f"via table::sort glob: {path.relative_to(root)}:{line_number}"
                )
            if (
                path in export_sources
                and public_use_or_type
                and _numbers_table_sort_owner_declaration(declaration)
            ):
                violations.append(
                    "focused litchi-numbers table-sort public API exposes public "
                    f"table-sort owner alias: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in NUMBERS_TABLE_SORT_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-numbers table-sort public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    path in export_sources
                    and public_use_or_type
                    and _numbers_table_sort_owner_declaration(declaration)
                    and identifier in NUMBERS_TABLE_SORT_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-numbers table-sort public API retains root alias "
                        f"{identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_sort_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers table-sort public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers table-sort public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )
            for match in NUMBERS_TABLE_SORT_RAW_PARAMETER.finditer(declaration):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                parameter_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers table-sort public API exposes raw "
                    f"parameter {parameter}: {path.relative_to(root)}:{parameter_line}"
                )

    return sorted(set(violations))


def _numbers_table_move_owner_present(root: Path) -> bool:
    """Return whether the private focused Numbers relocation owner is wired."""

    owner_path = root / NUMBERS_TABLE_MOVE_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_MOVE_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and NUMBERS_PACKAGE_TABLE_MOVE_MODULE.search(
        package_source
    ) is not None


def audit_numbers_table_move_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require a private selector-first focused Numbers relocation owner."""

    if not _numbers_table_move_owner_present(root):
        return []

    owner_path = root / NUMBERS_TABLE_MOVE_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_MOVE_EXPORT_SOURCES[1]
    owner_source = _mask_rust_cfg_test_items(
        owner_path.read_text(encoding="utf-8")
    )
    package_source = _mask_rust_cfg_test_items(
        package_path.read_text(encoding="utf-8")
    )
    package_code = _mask_rust_non_code(package_source)
    violations: list[str] = []

    for match in PUBLIC_NUMBERS_PACKAGE_TABLE_MOVE_MODULE.finditer(package_code):
        line_number = package_code.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-numbers table-move package module must remain private: "
            f"{NUMBERS_TABLE_MOVE_EXPORT_SOURCES[1]}:{line_number}"
        )

    methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    method_record = methods.get(NUMBERS_TABLE_MOVE_PACKAGE_METHOD)
    if method_record is None:
        violations.append(
            "focused litchi-numbers table-move public API is missing "
            f"Package::{NUMBERS_TABLE_MOVE_PACKAGE_METHOD}: "
            f"{NUMBERS_TABLE_MOVE_OWNER_SOURCE}"
        )
        return sorted(set(violations))

    declaration, line_number = method_record
    for selector in NUMBERS_TABLE_MOVE_SELECTOR_TYPES:
        if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
            violations.append(
                "focused litchi-numbers table-move Package method "
                f"{NUMBERS_TABLE_MOVE_PACKAGE_METHOD} must accept selector-first "
                f"{selector}: {NUMBERS_TABLE_MOVE_OWNER_SOURCE}:{line_number}"
            )
    for match in NUMBERS_TABLE_MOVE_RAW_PARAMETER.finditer(declaration):
        parameter = re.sub(r"\s+", " ", match.group(0)).strip()
        violations.append(
            "focused litchi-numbers table-move Package method exposes raw "
            f"identifier parameter {parameter}: "
            f"{NUMBERS_TABLE_MOVE_OWNER_SOURCE}:{line_number}"
        )

    return sorted(set(violations))


def audit_numbers_table_header_settings_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the nested, archive-free Numbers header transaction API."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_HEADER_SETTINGS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_HEADER_TRANSACTION_MODULE.search(
        _mask_rust_non_code(semantic_source)
    ) is None:
        violations.append(
            "focused litchi-numbers table-header settings public API is missing "
            "canonical headers::transaction module: "
            f"{NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE}"
        )

    transaction_body = _rust_public_module_body(semantic_source, "transaction")
    external_transaction = root / NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE
    transaction_source = (
        transaction_body
        if transaction_body is not None
        else external_transaction.read_text(encoding="utf-8")
        if external_transaction.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        transaction_source, NUMBERS_TABLE_HEADER_SETTINGS_SHORT_NAMES
    )
    for name in NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-numbers table-header settings public API is missing "
            f"canonical transaction type {name}: "
            f"{NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE}"
        )

    table_path = root / NUMBERS_SOURCE_ROOT / "table.rs"
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_HEADERS_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-numbers table-header settings public API is missing "
            "canonical table::headers module: "
            f"{NUMBERS_SOURCE_ROOT / 'table.rs'}"
        )

    lib_path = root / NUMBERS_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-numbers table-header settings public API is missing "
            "canonical root table module: "
            f"{NUMBERS_SOURCE_ROOT / 'lib.rs'}"
        )

    package_export = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        for match in PUBLIC_NUMBERS_PACKAGE_TABLE_HEADERS_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers table-header settings public API exposes "
                "duplicate package::table_headers module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for (
            declaration,
            line_number,
            public_declaration,
            complete_source_scope,
        ) in declarations:
            if not _is_numbers_table_header_settings_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _numbers_table_header_settings_owner_declaration(
                declaration
            )
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = declaration_identifiers[:2] in (
                ["pub", "type"],
                ["pub", "use"],
            )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and declaration_identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-numbers table-header settings public API "
                    "retains root aliases via table-header glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in (
                    NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIASES
                    | NUMBERS_TABLE_HEADER_SETTINGS_FLAT_SEMANTIC_ALIASES
                ):
                    violations.append(
                        "focused litchi-numbers table-header settings public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in NUMBERS_TABLE_HEADER_SETTINGS_ROOT_ALIASES
                ):
                    violations.append(
                        "focused litchi-numbers table-header settings public API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_header_settings_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-numbers table-header settings public API "
                    f"exposes {reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers table-header settings public API "
                    f"exposes raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_numbers_table_title_settings_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the direct, archive-free Numbers table-title transaction API."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_TITLE_SETTINGS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / NUMBERS_TABLE_TITLE_SETTINGS_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_TITLE_SETTINGS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, NUMBERS_TABLE_TITLE_SETTINGS_SHORT_NAMES
    )
    for name in NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-numbers table-title settings public API is missing "
            f"canonical table::title type {name}: "
            f"{NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE}"
        )

    lib_path = root / NUMBERS_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-numbers table-title settings public API is missing "
            "canonical root table module: "
            f"{NUMBERS_SOURCE_ROOT / 'lib.rs'}"
        )

    table_path = root / NUMBERS_SOURCE_ROOT / "table.rs"
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_TITLE_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-numbers table-title settings public API is missing "
            "canonical table::title module: "
            f"{NUMBERS_SOURCE_ROOT / 'table.rs'}"
        )

    package_export = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        if NUMBERS_PACKAGE_TABLE_TITLE_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-numbers table-title settings public API is missing "
                "private package owner module: "
                f"{package_export.relative_to(root)}"
            )
        for match in PUBLIC_NUMBERS_PACKAGE_TABLE_TITLE_MODULE.finditer(package_source):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers table-title settings public API exposes "
                "duplicate package::table_title module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-numbers table-title settings public API is missing "
            "private package owner module: "
            f"{NUMBERS_SOURCE_ROOT / 'package.rs'}"
        )

    owner_path = root / NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-numbers table-title settings public API is missing "
            "private package owner source: "
            f"{NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE}"
        )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for (
            declaration,
            line_number,
            public_declaration,
            complete_source_scope,
        ) in declarations:
            if not _is_numbers_table_title_settings_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _numbers_table_title_settings_owner_declaration(
                declaration
            )
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = declaration_identifiers[:2] in (
                ["pub", "type"],
                ["pub", "use"],
            )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and declaration_identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-numbers table-title settings public API retains "
                    "root aliases via table::title glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and public_use_or_type
            ):
                violations.append(
                    "focused litchi-numbers table-title settings public API exposes "
                    "public table-title owner alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    public_declaration
                    and identifier in NUMBERS_TABLE_TITLE_SETTINGS_FLAT_ALIASES
                ):
                    violations.append(
                        "focused litchi-numbers table-title settings public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in NUMBERS_TABLE_TITLE_SETTINGS_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-numbers table-title settings public API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_title_settings_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-numbers table-title settings public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers table-title settings public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def _numbers_table_appearance_owner_present(root: Path) -> bool:
    owner_path = root / NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE
    codec_path = root / NUMBERS_TABLE_APPEARANCE_CODEC_SOURCE
    package_path = root / NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and codec_path.is_file()
        and NUMBERS_PACKAGE_TABLE_APPEARANCE_MODULE.search(package_source) is not None
    )


def audit_numbers_table_appearance_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Numbers appearance owner.

    This is intentionally dormant until the private package owner appears.
    The activation token prevents a half-migrated COW/metadata route from
    making the current compatibility host appear retired.
    """

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_appearance_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_APPEARANCE_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(NUMBERS_TABLE_APPEARANCE_SEMANTIC_TYPES)
    )
    for name in NUMBERS_TABLE_APPEARANCE_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-numbers table-appearance public API is missing "
                f"canonical table::appearance type {name}: "
                f"{NUMBERS_TABLE_APPEARANCE_SEMANTIC_SOURCE}"
            )

    table_path = root / NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[2]
    table_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(table_path.read_text(encoding="utf-8"))
        )
        if table_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_APPEARANCE_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-numbers table-appearance public API is missing "
            f"canonical table::appearance module: {NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[2]}"
        )

    package_path = root / NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    if NUMBERS_PACKAGE_TABLE_APPEARANCE_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-numbers table-appearance public API is missing private "
            f"package owner module: {NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_NUMBERS_PACKAGE_TABLE_APPEARANCE_MODULE.finditer(
        package_source
    ):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-numbers table-appearance public API exposes public "
            f"package::table_appearance module: {NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[1]}:{line_number}"
        )

    owner_path = root / NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_code = _mask_rust_non_code(owner_source)
    owner_exports = _rust_canonical_exports(
        owner_source, frozenset(NUMBERS_TABLE_APPEARANCE_TRANSACTION_TYPES)
    )
    for name in NUMBERS_TABLE_APPEARANCE_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers table-appearance public API is missing "
                f"canonical package type {name}: {NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE}"
            )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_APPEARANCE_PACKAGE_METHODS:
        declaration = owner_methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-numbers table-appearance public API is missing "
                f"Package method {method}: {NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
            continue
        if method != "apply_table_appearance" and not re.search(
            r"\bSheetSelector\b", declaration
        ):
            violations.append(
                "focused litchi-numbers table-appearance Package method "
                f"{method} must accept selector-first SheetSelector: "
                f"{NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
        if method != "apply_table_appearance" and not re.search(
            r"\bTableSelector\b", declaration
        ):
            violations.append(
                "focused litchi-numbers table-appearance Package method "
                f"{method} must accept selector-first TableSelector: "
                f"{NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE}"
            )

    for label, marker in NUMBERS_TABLE_APPEARANCE_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-numbers table-appearance owner is missing "
                f"{label} transaction marker: {NUMBERS_TABLE_APPEARANCE_OWNER_SOURCE}"
            )

    codec_path = root / NUMBERS_TABLE_APPEARANCE_CODEC_SOURCE
    codec_source = (
        _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        if codec_path.is_file()
        else ""
    )
    codec_code = _mask_rust_non_code(codec_source)
    # The strict codec documents raw preservation in module docs, including
    # balanced unknown groups. Keep API matching code-only, but let contract
    # markers be proven by non-test source documentation as well.
    codec_marker_source = codec_source
    if not codec_path.is_file():
        violations.append(
            "focused litchi-numbers table-appearance public API is missing strict "
            f"hidden codec source: {NUMBERS_TABLE_APPEARANCE_CODEC_SOURCE}"
        )
    else:
        for api in NUMBERS_TABLE_APPEARANCE_CODEC_REQUIRED_APIS:
            if re.search(rf"\b(?:pub\s+)?(?:fn|struct)\s+{re.escape(api)}\b", codec_code) is None:
                violations.append(
                    "focused litchi-numbers table-appearance hidden codec is missing "
                    f"strict API {api}: {NUMBERS_TABLE_APPEARANCE_CODEC_SOURCE}"
                )
        for marker in NUMBERS_TABLE_APPEARANCE_CODEC_REQUIRED_MARKERS:
            if re.search(re.escape(marker), codec_marker_source) is None:
                violations.append(
                    "focused litchi-numbers table-appearance hidden codec is missing "
                    f"strict {marker} marker: {NUMBERS_TABLE_APPEARANCE_CODEC_SOURCE}"
                )
    codec_lib_path = root / NUMBERS_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE
    codec_lib_source = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_APPEARANCE_CODEC_MODULE)}\b",
        codec_lib_source,
    ) is None:
        violations.append(
            "focused litchi-numbers table-appearance public API is missing hidden "
            f"codec module {NUMBERS_TABLE_APPEARANCE_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE}"
        )

    lib_path = root / NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        )
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+table\b", _mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-numbers table-appearance public API is missing root "
            f"table module: {NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES[0]}"
        )

    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_APPEARANCE_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    if (root / NUMBERS_TABLE_APPEARANCE_OWNER_HELPER_ROOT).is_dir():
        dedicated_sources.update(
            (root / NUMBERS_TABLE_APPEARANCE_OWNER_HELPER_ROOT).rglob("*.rs")
        )
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_APPEARANCE_EXPORT_SOURCES
        if (root / path).is_file()
    }
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_appearance_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "table_appearance"]:
                violations.append(
                    "focused litchi-numbers table-appearance public API exposes "
                    f"duplicate table_appearance module: {path.relative_to(root)}:{line_number}"
                )
            public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
            if (
                public_use_or_type
                and "*" in declaration
                and ("table_appearance" in identifiers or "appearance" in identifiers)
            ):
                violations.append(
                    "focused litchi-numbers table-appearance public API retains "
                    f"root aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in NUMBERS_TABLE_APPEARANCE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-numbers table-appearance public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_appearance_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers table-appearance public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers table-appearance public API exposes raw "
                    f"byte slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def _numbers_table_cell_pop_up_menu_owner_present(root: Path) -> bool:
    owner_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE
    codec_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and codec_path.is_file()
        and NUMBERS_PACKAGE_TABLE_CELL_POP_UP_MENU_MODULE.search(package_source)
        is not None
    )


def audit_numbers_table_cell_pop_up_menu_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Numbers popup-menu owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_pop_up_menu_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_TYPES)
    )
    for name in NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-numbers popup-menu public API is missing "
                f"canonical cell::data_format::pop_up_menu type {name}: "
                f"{NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_SOURCE}"
            )
    semantic_code = _mask_rust_non_code(semantic_source)
    if PUBLIC_NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_MODULE.search(
        semantic_code
    ) is None:
        violations.append(
            "focused litchi-numbers popup-menu public API is missing canonical "
            "cell::data_format::pop_up_menu::transaction module: "
            f"{NUMBERS_TABLE_CELL_POP_UP_MENU_SEMANTIC_SOURCE}"
        )

    data_format_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        )
        if data_format_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_CELL_POP_UP_MENU_MODULE.search(data_format_source) is None:
        violations.append(
            "focused litchi-numbers popup-menu public API is missing canonical "
            "cell::data_format::pop_up_menu module: "
            f"{NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[2]}"
        )

    package_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    if NUMBERS_PACKAGE_TABLE_CELL_POP_UP_MENU_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-numbers popup-menu public API is missing private "
            "package owner module: "
            f"{NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_POP_UP_MENU_MODULE.finditer(
        package_source
    ):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-numbers popup-menu public API exposes public "
            "package::table_cell_pop_up_menu module: "
            f"{NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[1]}:{line_number}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_code = _mask_rust_non_code(owner_source)
    owner_exports = _rust_canonical_exports(
        owner_source, frozenset(NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_TYPES)
    )
    for name in NUMBERS_TABLE_CELL_POP_UP_MENU_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers popup-menu public API is missing "
                f"canonical package type {name}: {NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE}"
            )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_POP_UP_MENU_PACKAGE_METHODS:
        declaration = owner_methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-numbers popup-menu public API is missing "
                f"Package method {method}: {NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE}"
            )
            continue
        if method != "apply_table_cell_pop_up_menu_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-numbers popup-menu Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE}"
                    )

    owner_marker_code = owner_code + "\n" + _iwa_table_cell_format_engine_code(root)
    for label, marker in NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_marker_code) is None:
            violations.append(
                "focused litchi-numbers popup-menu owner is missing "
                f"{label} transaction marker: {NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE}"
            )

    codec_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE
    codec_source = (
        _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        if codec_path.is_file()
        else ""
    )
    codec_code = _mask_rust_non_code(codec_source)
    if not codec_path.is_file():
        violations.append(
            "focused litchi-numbers popup-menu public API is missing strict hidden "
            f"codec source: {NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE}"
        )
    else:
        for api in NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_REQUIRED_APIS:
            if re.search(
                rf"\b(?:pub\s+)?(?:fn|struct)\s+{re.escape(api)}\b",
                codec_code,
            ) is None:
                violations.append(
                    "focused litchi-numbers popup-menu hidden codec is missing "
                    f"strict API {api}: {NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE}"
                )
        for marker in NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_REQUIRED_MARKERS:
            if re.search(re.escape(marker), codec_source) is None:
                violations.append(
                    "focused litchi-numbers popup-menu hidden codec is missing "
                    f"strict {marker} marker: {NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE}"
                )

    codec_lib_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_PUBLIC_SOURCE
    codec_lib_source = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_MODULE)}\b",
        codec_lib_source,
    ) is None:
        violations.append(
            "focused litchi-numbers popup-menu public API is missing hidden codec "
            f"module {NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_PUBLIC_SOURCE}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        )
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+cell\b", lib_source) is None:
        violations.append(
            "focused litchi-numbers popup-menu public API is missing root cell module: "
            f"{NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES[0]}"
        )

    fuzz_targets = (
        NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_FUZZ_SOURCE,
        NUMBERS_TABLE_CELL_POP_UP_MENU_FUZZ_SOURCE,
    )
    fuzz_corpora = (
        NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_POP_UP_MENU_FUZZ_CORPUS,
    )
    for fuzz_path in fuzz_targets:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers popup-menu boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-numbers popup-menu fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in fuzz_corpora:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-numbers popup-menu boundary is missing fuzz corpus: "
                f"{corpus}"
            )

    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_POP_UP_MENU_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_POP_UP_MENU_EXPORT_SOURCES
        if (root / path).is_file()
    }
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_pop_up_menu_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_pop_up_menu" in identifiers
                    or "pop_up_menu" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers popup-menu public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in NUMBERS_TABLE_CELL_POP_UP_MENU_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-numbers popup-menu public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_pop_up_menu_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers popup-menu public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            raw_parameters = NUMBERS_TABLE_CELL_POP_UP_MENU_RAW_PARAMETER.finditer(
                declaration
            )
            for raw_parameter in raw_parameters:
                violations.append(
                    "focused litchi-numbers popup-menu public API exposes raw "
                    f"identifier parameter {raw_parameter.group(0).strip()}: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers popup-menu public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def _numbers_table_cell_number_format_owner_present(root: Path) -> bool:
    """Return whether the complete Number-format owner activates the ratchet."""

    owner_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE
    codec_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    codec_source = (
        _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        if codec_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and codec_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and bool(_mask_rust_non_code(codec_source).strip())
        and NUMBERS_PACKAGE_TABLE_CELL_NUMBER_FORMAT_MODULE.search(package_source)
        is not None
    )


def _numbers_table_cell_number_format_marker_present(
    source: str, marker: str | tuple[str, ...]
) -> bool:
    """Match one marker or an alternative marker family in source text."""

    if isinstance(marker, tuple):
        return any(
            _numbers_table_cell_number_format_marker_present(source, option)
            for option in marker
        )
    return re.search(re.escape(marker), source, re.IGNORECASE) is not None


def _numbers_table_cell_number_format_api_present(
    source: str, api: str | tuple[str, ...]
) -> bool:
    """Recognize a strict codec function/type or a focused reexport."""

    if isinstance(api, tuple):
        return any(
            _numbers_table_cell_number_format_api_present(source, option)
            for option in api
        )
    return re.search(
        rf"\b(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
        rf"(?:async[ \t\r\n]+|const[ \t\r\n]+|unsafe[ \t\r\n]+)*"
        rf"(?:fn|struct|enum|type)[ \t\r\n]+{re.escape(api)}\b",
        source,
    ) is not None or re.search(
        rf"\bpub\s+use[^;]*\b{re.escape(api)}\b",
        source,
    ) is not None


def audit_numbers_table_cell_number_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Number-format codec seam."""

    if not _numbers_table_cell_number_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE
    codec_paths = tuple(
        root / path
        for path in NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    )
    # The neutral Number-format module is a deliberately small wrapper around
    # the existing strict control-cell walker.  Audit the wrapper and the
    # actual implementation/projection together so a reexport cannot hide an
    # eager or unbounded path, while avoiding a duplicate parser requirement.
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)

    for api in NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        expected = (
            " or ".join(api) if isinstance(api, tuple) else api
        )
        violations.append(
            "focused litchi-numbers Number-format hidden codec is missing "
            f"strict API {expected}: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE}"
        )

    for label, markers in NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_REQUIRED_MARKER_GROUPS.items():
        if all(
            _numbers_table_cell_number_format_marker_present(codec_code, marker)
            for marker in markers
        ):
            continue
        violations.append(
            "focused litchi-numbers Number-format hidden codec is missing "
            f"strict {label} marker: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE}"
        )

    # Buffa is the only accepted generated projection.  A normal Prost decode
    # or an owned generated message would make source preservation accidental.
    if re.search(r"\bbuffa\b", codec_code) is None:
        violations.append(
            "focused litchi-numbers Number-format hidden codec is missing Buffa "
            f"projection ingress: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Number-format hidden codec retains a Prost "
            f"production path: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE}"
        )
    if re.search(
        r"NATIVE_NUMBER_FORMAT_TYPE\s*:\s*(?:u\d+\s*)?=\s*256\b|"
        r"NATIVE_NUMBER_FORMAT_TYPE\b[\s\S]{0,200}\b256\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Number-format hidden codec must retain "
            f"native type 256: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE}"
        )

    lazy_matches = [
        match.start()
        for pattern in (
            r"\bdecode_lazy_view\b",
            r"\bdecode_view\b",
            r"\bLazyView\b",
            r"\bvisitor\b",
        )
        for match in re.finditer(pattern, codec_code)
    ]
    preflight_positions = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|scan_number_message|scan_decimal_message|"
            r"scan_display_format|Budget\s*::\s*new)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy_matches and (
        not preflight_positions
        or min(preflight_positions) > min(lazy_matches)
    ):
        violations.append(
            "focused litchi-numbers Number-format hidden codec must preflight "
            "wire before forcing its lazy Buffa view: "
            f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE}"
        )

    # Keep the generated view private while making the strict codec available
    # only as a doc-hidden implementation seam to the focused package.
    codec_lib_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE
    codec_lib_raw = (
        codec_lib_path.read_text(encoding="utf-8")
        if codec_lib_path.is_file()
        else ""
    )
    codec_lib_production = _mask_rust_cfg_test_items(codec_lib_raw)
    codec_lib_code = _mask_rust_non_code(codec_lib_production)
    if not _rust_root_level_matches(
        codec_lib_production,
        NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_HIDDEN_MODULE,
    ):
        violations.append(
            "focused litchi-numbers Number-format public API is missing hidden "
            f"codec module {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULES),
    )
    if not generated_declarations:
        violations.append(
            "focused litchi-numbers Number-format public API is missing private "
            "Buffa generated projection module (one of "
            f"{', '.join(NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULES)}): "
            f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    else:
        for module, visibility, shape, body, line_number in generated_declarations:
            if visibility is not None:
                violations.append(
                    "focused litchi-numbers Number-format generated projection module "
                    f"must remain private: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
                )
            if shape == "invalid" or (
                shape == "inline" and not _rust_generated_module_include_is_valid(body)
            ):
                violations.append(
                    "focused litchi-numbers Number-format generated projection "
                    f"module has invalid include shape: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
                )
    if _rust_root_generated_public_leaks(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_GENERATED_MODULES),
    ):
        violations.append(
            "focused litchi-numbers Number-format public API reexports its Buffa "
            f"generated projection: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_PUBLIC_SOURCE}"
        )

    # Every strict codec has a cfg(test) round-trip/hostile-wire harness.  The
    # exact native acceptance evidence belongs to the Rust test suite; this
    # source ratchet prevents an implementation-only codec from landing.
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,1200}"
        r"#\s*\[\s*test\s*\]",
        _mask_rust_non_code(codec_raw),
    ) is None:
        violations.append(
            "focused litchi-numbers Number-format hidden codec is missing its "
            f"cfg(test) harness: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_SOURCE}"
        )

    return sorted(set(violations))


def audit_numbers_table_cell_number_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Number-format owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_number_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_code = _mask_rust_non_code(semantic_source)
    if (
        re.search(r"\bpub\s+struct\s+(?:r#)?Number\b", semantic_code) is None
        and re.search(r"\bdecimal_format!\s*\(\s*(?:r#)?Number\b", semantic_code)
        is None
    ):
        violations.append(
            "focused litchi-numbers Number-format public API is missing canonical "
            f"cell::data_format::Number: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE}"
        )
    if PUBLIC_NUMBERS_CELL_NUMBER_FORMAT_TRANSACTION_MODULE.search(semantic_code) is None:
        external_transaction = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_SOURCE
        if not external_transaction.is_file():
            violations.append(
                "focused litchi-numbers Number-format public API is missing canonical "
                f"cell::data_format::number::transaction module: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_SEMANTIC_SOURCE}"
            )

    transaction_body = _rust_public_module_body(semantic_source, "transaction")
    transaction_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_SOURCE
    transaction_source = (
        transaction_body
        if transaction_body is not None
        else transaction_path.read_text(encoding="utf-8")
        if transaction_path.is_file()
        else ""
    )
    transaction_exports = _rust_canonical_exports(
        transaction_source,
        frozenset(NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES:
        if name in transaction_exports:
            continue
        violations.append(
            "focused litchi-numbers Number-format public API is missing canonical "
            f"transaction type {name}: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_SOURCE}"
        )

    data_format_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    data_format_code = _mask_rust_non_code(data_format_source)
    if PUBLIC_NUMBERS_CELL_DATA_FORMAT_NUMBER_MODULE.search(data_format_code) is None:
        violations.append(
            "focused litchi-numbers Number-format public API is missing canonical "
            f"cell::data_format::number module: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]}"
        )
    if "Number" not in _rust_canonical_exports(
        data_format_source, frozenset({"Number"})
    ):
        violations.append(
            "focused litchi-numbers Number-format public API is missing canonical "
            f"cell::data_format::Number reexport: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]}"
        )

    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    if NUMBERS_PACKAGE_TABLE_CELL_NUMBER_FORMAT_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-numbers Number-format public API is missing private "
            f"package owner module: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_NUMBER_FORMAT_MODULE.finditer(
        package_code
    ):
        line_number = package_code.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-numbers Number-format public API exposes public "
            f"package::table_cell_number_format module: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]}:{line_number}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+(?:r#)?cell\b", _mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-numbers Number-format public API is missing root cell module: "
            f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[0]}"
        )

    example_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Number-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXAMPLE}"
        )
    else:
        example_source = _mask_rust_cfg_test_items(
            example_path.read_text(encoding="utf-8")
        )
        example_code = _mask_rust_non_code(example_source)
        if not re.search(
            r"(?<![A-Za-z0-9_])(?:edit_table_cell_number_format|"
            r"apply_table_cell_number_format)\b[ \t\r\n]*\(",
            example_code,
        ):
            violations.append(
                "focused litchi-numbers Number-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Number-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = _mask_rust_cfg_test_items(
            integration_path.read_text(encoding="utf-8")
        )
        integration_code = _mask_rust_non_code(integration_source)
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Number-format integration test is missing "
                f"#[test] coverage: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_INTEGRATION_SOURCE}"
            )
        for label, marker in NUMBERS_TABLE_CELL_NUMBER_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_source) is None:
                violations.append(
                    "focused litchi-numbers Number-format integration test is missing "
                    f"{label}: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_INTEGRATION_SOURCE}"
                )

    for fuzz_path, label in (
        (
            NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_FUZZ_SOURCE,
            "codec fuzz target",
        ),
        (NUMBERS_TABLE_CELL_NUMBER_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Number-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Number-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_NUMBER_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_NUMBER_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Number-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Number-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )

    owner_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_helper_root = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    owner_code = _mask_rust_non_code(owner_related_source)
    owner_exports = _rust_canonical_exports(
        owner_related_source,
        frozenset(NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES:
        if name in owner_exports:
            continue
        violations.append(
            "focused litchi-numbers Number-format public API is missing canonical "
            f"package type {name}: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}"
        )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PACKAGE_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Number-format public API is missing "
                f"Package method {method}: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}"
            )
            continue
        declaration, line_number = record
        if method != "apply_table_cell_number_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Number-format Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
                    )
        if method == "table_cell_number_format":
            normalized = re.sub(r"\s+", "", declaration)
            if "Option<Number>" not in normalized:
                violations.append(
                    "focused litchi-numbers Number-format read must return "
                    f"Option<Number>: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
                )
        elif method == "edit_table_cell_number_format":
            if not re.search(r"\bEdit\b", declaration):
                violations.append(
                    "focused litchi-numbers Number-format edit must return "
                    f"Edit: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
                )
        elif not re.search(r"\bPatch\b", declaration):
            violations.append(
                "focused litchi-numbers Number-format apply must accept "
                f"Patch: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
            )
        if method == "apply_table_cell_number_format" and not re.search(
            r"\bCommit\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Number-format apply must return "
                f"Commit: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
            )
        for match in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER.finditer(
            declaration
        ):
            parameter = re.sub(r"\s+", " ", match.group(0)).strip()
            violations.append(
                "focused litchi-numbers Number-format Package method exposes raw "
                f"identifier/coordinate parameter {parameter}: "
                f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
            )
        raw_identifier = NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_ID_PARAMETER.search(
            declaration
        )
        if raw_identifier is not None:
            violations.append(
                "focused litchi-numbers Number-format Package method exposes raw "
                f"identifier parameter {raw_identifier.group(0).strip()}: "
                f"{NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
            )

    edit_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Edit"
        )
    }
    for method in NUMBERS_TABLE_CELL_NUMBER_FORMAT_EDIT_METHODS:
        record = edit_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Number-format edit is missing "
                f"Edit::{method}: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}"
            )
            continue
        declaration, line_number = record
        if method == "set" and not re.search(r"\bNumber\b", declaration):
            violations.append(
                "focused litchi-numbers Number-format Edit::set must accept "
                f"Number: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
            )
        if method == "commit" and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Number-format Edit::commit must return "
                f"Commit: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-numbers Number-format owner is missing "
                f"{label} transaction marker: {NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE}"
            )

    # Inspect only the focused implementation/export scopes.  Other semantic
    # display formats and the shared IWA adapters are intentionally unrelated.
    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_NUMBER_FORMAT_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES
        if (root / path).is_file()
    }
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_number_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (
                ["pub", "use"],
                ["pub", "type"],
            )
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_number_format" in identifiers
                    or "number_format" in identifiers
                    or "number" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers Number-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_NUMBER_FORMAT_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_NUMBER_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Number-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    not dedicated_source
                    and identifier in NUMBERS_TABLE_CELL_NUMBER_FORMAT_TRANSACTION_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers Number-format public API retains "
                        f"root transaction alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_number_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Number-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers Number-format public API exposes raw "
                    f"byte slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )
            for match in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER.finditer(
                declaration
            ):
                container = re.sub(r"\s+", " ", match.group(0)).strip()
                container_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers Number-format public API exposes opaque "
                    f"raw byte container {container}: {path.relative_to(root)}:{container_line}"
                )
            for match in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE.finditer(
                declaration
            ):
                container = re.sub(r"\s+", " ", match.group(0)).strip()
                container_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers Number-format public API exposes opaque "
                    f"raw byte container type {container}: {path.relative_to(root)}:{container_line}"
                )
            for match in NUMBERS_TABLE_CELL_NUMBER_FORMAT_RAW_HELPER.finditer(
                declaration
            ):
                helper = match.group(0)
                helper_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers Number-format public API exposes raw "
                    f"helper {helper}: {path.relative_to(root)}:{helper_line}"
                )
            for match in NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER.finditer(
                declaration
            ):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                parameter_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers Number-format public API exposes typed "
                    f"raw identifier parameter {parameter}: {path.relative_to(root)}:{parameter_line}"
                )
            for match in NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_PARAMETER.finditer(
                declaration
            ):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                parameter_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers Number-format public API exposes raw "
                    f"identifier/coordinate parameter {parameter}: {path.relative_to(root)}:{parameter_line}"
                )

    return sorted(set(violations))


def _numbers_table_cell_text_format_owner_present(root: Path) -> bool:
    """Return whether the focused Text-format owner has activated."""

    owner_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_EXPORT_SOURCES[1]
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        # Keep activation broad enough to audit a nested decoy, then require
        # the actual owner module at the package root in the facade audit.
        and re.search(
            r"\bmod[ \t\r\n]+(?:r#)?table_cell_text_format\b",
            package_source,
        )
        is not None
    )


def _numbers_table_cell_text_format_semantic_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return the dedicated Text leaf or its explicit transition fallback."""

    canonical = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_SOURCE
    if canonical.is_file():
        return (canonical,)
    fallback = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_FALLBACK_SOURCE
    return (fallback,) if fallback.is_file() else ()


def _numbers_table_cell_text_format_transaction_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return sources that may own the focused Text transaction namespace."""

    candidates = (
        NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_SOURCE,
        NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_SOURCE,
    )
    return tuple(root / path for path in candidates if (root / path).is_file())


def _numbers_table_cell_text_format_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden by the Text facade."""

    if identifier in {
        "Draft",
        "Draft",
        "Draft",
        "LitchiIwaPackage",
    }:
        return "monolithic host type"
    if identifier in NUMBERS_TABLE_CELL_TEXT_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_TEXT_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_TEXT_FORMAT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost", "codec"} for word in words):
        return "protobuf type"
    if identifier in {"Projection", "LitchiIwaProjection"}:
        return "generated type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_text_format_owner_declaration(
    declaration: str,
) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_TEXT_FORMAT_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_text_format_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_CELL_TEXT_FORMAT_FLAT_ALIASES
            | {"Text", "text", "text_format"}
        )
    ) or any(
        NUMBERS_TABLE_CELL_TEXT_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(identifier)
        for identifier in identifiers
    ) or _numbers_table_cell_text_format_owner_declaration(declaration)


def audit_numbers_table_cell_text_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Text-format codec seam."""

    if not _numbers_table_cell_text_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE
    dedicated_raw = (
        codec_path.read_text(encoding="utf-8") if codec_path.is_file() else ""
    )
    dedicated_production = _mask_rust_cfg_test_items(dedicated_raw)
    dedicated_code = _mask_rust_non_code(dedicated_production)
    # The focused module is an ownership seam, even when it forwards the
    # actual walker to the shared popup/display core.  A missing or comment-
    # only wrapper must not be able to activate this audit through that shared
    # implementation alone.
    if codec_path.is_file() and not dedicated_code.strip():
        violations.append(
            "focused litchi-numbers Text-format hidden codec wrapper is missing "
            f"or empty: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )
    codec_paths = tuple(
        root / path
        for path in NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    )
    if not codec_path.is_file():
        violations.append(
            "focused litchi-numbers Text-format boundary is missing strict codec "
            f"source: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)

    for api in NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        expected = " or ".join(api) if isinstance(api, tuple) else api
        violations.append(
            "focused litchi-numbers Text-format hidden codec is missing strict "
            f"API {expected}: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )

    for label, markers in NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_REQUIRED_MARKER_GROUPS.items():
        if all(
            _numbers_table_cell_number_format_marker_present(codec_code, marker)
            for marker in markers
        ):
            continue
        violations.append(
            "focused litchi-numbers Text-format hidden codec is missing strict "
            f"{label} marker: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )

    if re.search(
        r"NATIVE_TEXT_FORMAT_TYPE\s*:\s*(?:u\d+\s*)?=\s*260\b|"
        r"(?:TEXT_FORMAT_TYPE|text_format_type)\b[\s\S]{0,200}\b260\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Text-format hidden codec must retain native "
            f"type 260: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Text-format hidden codec retains a Prost "
            f"production path: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )
    if re.search(r"\bbuffa\b", codec_code) is None:
        violations.append(
            "focused litchi-numbers Text-format hidden codec is missing Buffa "
            f"projection ingress: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )

    lazy_matches = [
        match.start()
        for pattern in (
            r"\bdecode_lazy_view\b",
            r"\bdecode_view\b",
            r"\bLazyView\b",
            r"\bvisitor\b",
        )
        for match in re.finditer(pattern, codec_code)
    ]
    preflight_positions = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|scan_text_message|scan_format_message|"
            r"scan_display_format|Budget\s*::\s*new)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy_matches and (
        not preflight_positions or min(preflight_positions) > min(lazy_matches)
    ):
        violations.append(
            "focused litchi-numbers Text-format hidden codec must preflight wire "
            f"before forcing its lazy Buffa view: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )

    codec_lib_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_PUBLIC_SOURCE
    codec_lib_raw = (
        codec_lib_path.read_text(encoding="utf-8")
        if codec_lib_path.is_file()
        else ""
    )
    codec_lib_code = _mask_rust_non_code(
        _mask_rust_cfg_test_items(codec_lib_raw)
    )
    codec_lib_production = _mask_rust_cfg_test_items(codec_lib_raw)
    hidden_module = re.compile(
        rf"(?m)^[ \t]*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_MODULE)}\b"
    )
    if not _rust_root_level_matches(codec_lib_production, hidden_module):
        violations.append(
            "focused litchi-numbers Text-format public API is missing hidden codec "
            f"module {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_PUBLIC_SOURCE}"
        )

    generated_declarations = _rust_root_level_module_declarations(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_GENERATED_MODULES),
    )
    if not generated_declarations:
        violations.append(
            "focused litchi-numbers Text-format public API is missing a private "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    else:
        for module, visibility, shape, body, line_number in generated_declarations:
            if visibility is not None:
                violations.append(
                    "focused litchi-numbers Text-format generated projection module "
                    f"must remain private: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
                )
            if shape == "invalid" or (
                shape == "inline" and not _rust_generated_module_include_is_valid(body)
            ):
                violations.append(
                    "focused litchi-numbers Text-format generated projection module "
                    f"has invalid include shape: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
                )
    if _rust_root_generated_public_leaks(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_GENERATED_MODULES),
    ):
        violations.append(
            "focused litchi-numbers Text-format public API reexports its Buffa "
            f"generated projection: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,4096}"
        r"#\s*\[\s*test\s*\]",
        _mask_rust_non_code(dedicated_raw),
    ) is None:
        violations.append(
            "focused litchi-numbers Text-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_SOURCE}"
        )

    return sorted(set(violations))


def audit_numbers_table_cell_text_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Text-format owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_text_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_paths = _numbers_table_cell_text_format_semantic_sources(root)
    semantic_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in semantic_paths
    )
    semantic_code = _mask_rust_non_code(semantic_source)
    canonical_text_patterns = (
        re.compile(
            r"(?m)^[ \t]*pub[ \t\r\n]+struct[ \t\r\n]+(?:r#)?Text\b"
        ),
        re.compile(r"(?m)^[ \t]*pub[ \t\r\n]+use[^;\n]*\bText\b"),
    )
    if not any(
        _rust_root_level_matches(semantic_source, pattern)
        for pattern in canonical_text_patterns
    ):
        violations.append(
            "focused litchi-numbers Text-format public API is missing canonical "
            f"cell::data_format::Text: {NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_SOURCE}"
        )

    transaction_paths = _numbers_table_cell_text_format_transaction_sources(root)
    inline_transaction = _rust_public_module_body(semantic_source, "transaction")
    external_transaction = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_SOURCE
    transaction_source = (
        inline_transaction
        if inline_transaction is not None
        else _mask_rust_cfg_test_items(
            external_transaction.read_text(encoding="utf-8")
        )
        if external_transaction.is_file()
        else ""
    )
    transaction_code = _mask_rust_non_code(transaction_source)
    transaction_module = re.compile(
        r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    )
    if not (
        _rust_root_level_matches(semantic_source, transaction_module)
        or _rust_root_level_matches(transaction_source, transaction_module)
    ):
        violations.append(
            "focused litchi-numbers Text-format public API is missing the Text "
            f"transaction namespace: {NUMBERS_TABLE_CELL_TEXT_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_names = frozenset(NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES)
    transaction_exports = (
        _rust_canonical_exports(transaction_source, transaction_names)
        if inline_transaction is not None
        else _rust_root_canonical_exports(transaction_source, transaction_names)
    )
    for name in NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers Text-format public API is missing canonical "
                f"transaction type {name}: {NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_SOURCE}"
            )

    data_format_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    data_format_code = _mask_rust_non_code(data_format_source)
    if not (
        _rust_root_level_matches(
            data_format_source, PUBLIC_NUMBERS_CELL_DATA_FORMAT_TEXT_MODULE
        )
        or _rust_root_level_matches(
            data_format_source,
            re.compile(
                r"(?m)^[ \t]*pub[ \t\r\n]+struct[ \t\r\n]+(?:r#)?Text\b"
            ),
        )
    ):
        violations.append(
            "focused litchi-numbers Text-format public API is missing canonical "
            f"cell::data_format::text module: {data_format_path.relative_to(root)}"
        )
    if "Text" not in _rust_root_canonical_exports(
        data_format_source, frozenset({"Text"})
    ):
        violations.append(
            "focused litchi-numbers Text-format public API is missing canonical "
            f"cell::data_format::Text reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    if not _rust_root_level_matches(
        package_source, NUMBERS_PACKAGE_TABLE_CELL_TEXT_FORMAT_MODULE
    ):
        violations.append(
            "focused litchi-numbers Text-format public API is missing private "
            f"package owner module: {package_path.relative_to(root)}"
        )
    public_package = _rust_root_level_matches(
        package_source, PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_TEXT_FORMAT_MODULE
    )
    if public_package:
        line_number = package_code.count("\n", 0, public_package[0].start()) + 1
        violations.append(
            "focused litchi-numbers Text-format public API exposes its public "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if not _rust_root_level_matches(
        lib_source,
        re.compile(r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?cell\b"),
    ):
        violations.append(
            "focused litchi-numbers Text-format public API is missing root cell "
            f"module: {lib_path.relative_to(root)}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_helper_root = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    owner_code = _mask_rust_non_code(owner_related_source)
    owner_exports = _rust_root_canonical_exports(
        owner_related_source,
        frozenset(NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers Text-format public API is missing canonical "
                f"package type {name}: {owner_path.relative_to(root)}"
            )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_TEXT_FORMAT_PACKAGE_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Text-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method != "apply_table_cell_text_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Text-format Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{owner_path.relative_to(root)}:{line_number}"
                    )
        if method == "table_cell_text_format":
            normalized = re.sub(r"\s+", "", declaration)
            if "Option<Text>" not in normalized:
                violations.append(
                    "focused litchi-numbers Text-format read must return "
                    f"Option<Text>: {owner_path.relative_to(root)}:{line_number}"
                )
        elif method == "edit_table_cell_text_format":
            if not re.search(r"\bEdit\b", declaration):
                violations.append(
                    "focused litchi-numbers Text-format edit must return "
                    f"Edit: {owner_path.relative_to(root)}:{line_number}"
                )
        elif not re.search(r"\bPatch\b", declaration):
            violations.append(
                "focused litchi-numbers Text-format apply must accept "
                f"Patch: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "apply_table_cell_text_format" and not re.search(
            r"\bCommit\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Text-format apply must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )
        for pattern in (
            NUMBERS_TABLE_CELL_TEXT_FORMAT_PUBLIC_RAW_PARAMETER,
            NUMBERS_TABLE_CELL_TEXT_FORMAT_RAW_ID_PARAMETER,
        ):
            for match in pattern.finditer(declaration):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                violations.append(
                    "focused litchi-numbers Text-format Package method exposes "
                    f"raw identifier/coordinate parameter {parameter}: "
                    f"{owner_path.relative_to(root)}:{line_number}"
                )

    edit_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Edit"
        )
    }
    for method in NUMBERS_TABLE_CELL_TEXT_FORMAT_EDIT_METHODS:
        record = edit_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Text-format edit is missing "
                f"Edit::{method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method == "set" and not (
            re.search(r"\bText\b", declaration)
            or re.search(r"\bset\s*\(\s*self\s*\)", declaration)
        ):
            violations.append(
                "focused litchi-numbers Text-format Edit::set must be typed with "
                f"Text: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "commit" and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Text-format Edit::commit must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_TEXT_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-numbers Text-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    dedicated_sources = set(semantic_paths) | set(transaction_paths) | {owner_path}
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_TEXT_FORMAT_EXPORT_SOURCES[:3]
        if (root / path).is_file()
    }
    export_sources.update(semantic_paths)
    export_sources.update(transaction_paths)
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_text_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (
                ["pub", "use"],
                ["pub", "type"],
            )
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_text_format" in identifiers
                    or "text_format" in identifiers
                    or "text" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers Text-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_TEXT_FORMAT_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_TEXT_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Text-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    not dedicated_source
                    and identifier in NUMBERS_TABLE_CELL_TEXT_FORMAT_TRANSACTION_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers Text-format public API retains root "
                        f"transaction alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_text_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Text-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for pattern, label in (
                (RUST_BYTE_SLICE, "raw byte slice"),
                (
                    NUMBERS_TABLE_CELL_TEXT_FORMAT_PUBLIC_RAW_CONTAINER,
                    "opaque raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_TEXT_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE,
                    "opaque raw byte container type",
                ),
                (NUMBERS_TABLE_CELL_TEXT_FORMAT_RAW_HELPER, "raw helper"),
                (
                    NUMBERS_TABLE_CELL_TEXT_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_TEXT_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw identifier/coordinate parameter",
                ),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    value_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        "focused litchi-numbers Text-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{value_line}"
                    )

    example_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Text-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_TEXT_FORMAT_EXAMPLE}"
        )
    else:
        example_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(example_path.read_text(encoding="utf-8"))
        )
        if re.search(
            r"\b(?:edit_table_cell_text_format|"
            r"apply_table_cell_text_format)\s*\(",
            example_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Text-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_TEXT_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_TEXT_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Text-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_TEXT_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = integration_path.read_text(encoding="utf-8")
        integration_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_source)
        )
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Text-format integration test is missing "
                f"#[test] coverage: {integration_path}"
            )
        for label, marker in NUMBERS_TABLE_CELL_TEXT_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_code) is None:
                violations.append(
                    "focused litchi-numbers Text-format integration test is missing "
                    f"{label}: {integration_path}"
                )

    for fuzz_path, label in (
        (
            NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_FUZZ_SOURCE,
            "codec fuzz target",
        ),
        (NUMBERS_TABLE_CELL_TEXT_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Text-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Text-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_TEXT_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_TEXT_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Text-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Text-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )

    return sorted(set(violations))


def _numbers_table_cell_date_time_format_owner_present(root: Path) -> bool:
    """Return whether the focused Date & Time-format owner has landed."""

    owner_path = root / NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_OWNER_SOURCE
    semantic_path = root / NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_SEMANTIC_SOURCE
    package_path = root / Path("crates/litchi-numbers/src/package.rs")
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and semantic_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and bool(_mask_rust_non_code(semantic_source).strip())
        and bool(
            _rust_root_level_matches(
                package_source,
                re.compile(
                    r"(?m)^[ \t]*pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+mod"
                    r"[ \t\r\n]+(?:r#)?table_cell_date_time_format\b"
                ),
            )
        )
    )


def _numbers_table_cell_date_time_format_public_leak(
    identifier: str,
) -> str | None:
    """Classify archive, wire, generated, or identifier vocabulary."""

    if identifier == "litchi_iwa_common":
        # The neutral table vocabulary owner.
        return None
    if identifier in NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_RAW_MARKERS:
        return "raw identifier or archive type"
    return _iwork_public_leak(identifier)


def audit_numbers_table_cell_date_time_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Date & Time codec seam."""

    if not _numbers_table_cell_date_time_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_path = root / NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE
    codec_paths = tuple(
        root / path
        for path in NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    )
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)
    if not codec_path.is_file() or not codec_code.strip():
        violations.append(
            "focused litchi-numbers Date & Time-format boundary is missing strict "
            f"codec source: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE}"
        )

    for api in NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        violations.append(
            "focused litchi-numbers Date & Time-format hidden codec is missing "
            f"strict API {api}: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE}"
        )

    for marker, description in (
        ("buffa", "Buffa projection ingress"),
        ("decode_lazy_view", "lazy Buffa view"),
        ("Budget::new", "bounded wire preflight"),
        ("unknown", "unknown-field preservation"),
        ("extend_from_slice", "source-preserving rewrite"),
        ("execution_requirements", "prepared bounded rewrite"),
        ("execute", "prepared bounded execution"),
    ):
        if re.search(re.escape(marker), codec_code, re.IGNORECASE):
            continue
        violations.append(
            "focused litchi-numbers Date & Time-format hidden codec is missing "
            f"{description}: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE}"
        )
    if re.search(
        r"(?:NATIVE_DATE_TIME_FORMAT_TYPE|DATE_TIME_FORMAT_TYPE)\s*"
        r":\s*(?:u\d+\s*)?=\s*261\b|"
        r"(?:NATIVE_DATE_TIME_FORMAT_TYPE|DATE_TIME_FORMAT_TYPE)\b"
        r"[\s\S]{0,200}\b261\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Date & Time-format hidden codec must retain "
            f"native type 261: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Date & Time-format hidden codec retains a "
            f"Prost production path: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE}"
        )

    # ``Budget::new`` is the shared scanner's preflight marker.  Require it to
    # occur before the first lazy-view force so a future wrapper cannot claim
    # lazy decoding while entering the generated view without limits.
    lazy = [
        match.start()
        for match in re.finditer(r"\b(?:decode_lazy_view|decode_view|LazyView)\b", codec_code)
    ]
    preflight = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|Budget\s*::\s*new|scan_date_time_format)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy and (not preflight or min(preflight) > min(lazy)):
        violations.append(
            "focused litchi-numbers Date & Time-format hidden codec must preflight "
            f"wire before forcing its lazy Buffa view: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE}"
        )

    public_path = root / NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_PUBLIC_SOURCE
    public_raw = (
        public_path.read_text(encoding="utf-8") if public_path.is_file() else ""
    )
    public_production = _mask_rust_cfg_test_items(public_raw)
    hidden_module = re.compile(
        rf"(?m)^[ \t]*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_MODULE)}\b"
    )
    if not _rust_root_level_matches(public_production, hidden_module):
        violations.append(
            "focused litchi-numbers Date & Time-format public API is missing hidden "
            f"codec module {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    generated = _rust_root_level_module_declarations(
        public_production,
        frozenset({NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_GENERATED_MODULE}),
    )
    if not generated:
        violations.append(
            "focused litchi-numbers Date & Time-format public API is missing a "
            f"private Buffa generated projection: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    for _module, visibility, shape, body, line_number in generated:
        if visibility is not None:
            violations.append(
                "focused litchi-numbers Date & Time-format generated projection "
                f"must remain private: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "focused litchi-numbers Date & Time-format generated projection "
                f"module has invalid include shape: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        public_production,
        frozenset({NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_GENERATED_MODULE}),
    ):
        violations.append(
            "focused litchi-numbers Date & Time-format public API reexports its "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_PUBLIC_SOURCE}"
        )

    # The dedicated wrapper must carry an explicit hostile-wire test module;
    # shared-core tests alone are not enough to prove the Date & Time seam is
    # exercised after a refactor.
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,4096}"
        r"#\s*\[\s*test\s*\]",
        _mask_rust_non_code(
            codec_path.read_text(encoding="utf-8") if codec_path.is_file() else ""
        ),
    ) is None:
        violations.append(
            "focused litchi-numbers Date & Time-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_CODEC_SOURCE}"
        )
    return sorted(set(violations))


def audit_numbers_table_cell_date_time_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the archive-free Date & Time semantic facade."""

    if not (root / NUMBERS_SOURCE_ROOT).is_dir() or not _numbers_table_cell_date_time_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_SEMANTIC_SOURCE
    semantic_source = _mask_rust_cfg_test_items(
        semantic_path.read_text(encoding="utf-8")
    )
    if not _rust_root_level_matches(
        semantic_source,
        re.compile(r"(?m)^[ \t]*pub\s+struct\s+(?:r#)?DateTime\b"),
    ):
        violations.append(
            "focused litchi-numbers Date & Time-format public API is missing "
            f"canonical cell::data_format::DateTime: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_SEMANTIC_SOURCE}"
        )
    if not _rust_root_level_matches(
        semantic_source,
        re.compile(r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?transaction\b"),
    ):
        violations.append(
            "focused litchi-numbers Date & Time-format public API is missing its "
            f"transaction namespace: {NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_SEMANTIC_SOURCE}"
        )

    data_format_path = root / Path("crates/litchi-numbers/src/cell/data_format.rs")
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    if not _rust_root_level_matches(
        data_format_source,
        re.compile(
            r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?date_time\b[ \t\r\n]*(?:;|\{)"
        ),
    ):
        violations.append(
            "focused litchi-numbers Date & Time-format public API is missing its "
            f"canonical date_time module: {data_format_path.relative_to(root)}"
        )
    if not _rust_root_level_matches(
        data_format_source,
        re.compile(
            r"(?m)^[ \t]*pub\s+use\s+(?:r#)?date_time\s*::\s*"
            r"(?:r#)?DateTime\b"
        ),
    ):
        violations.append(
            "focused litchi-numbers Date & Time-format public API is missing its "
            f"canonical DateTime reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / Path("crates/litchi-numbers/src/package.rs")
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    private_module = re.compile(
        r"(?m)^[ \t]*pub\s*\([ \t]*crate[ \t]*\)\s+mod\s+"
        r"(?:r#)?table_cell_date_time_format\b[ \t\r\n]*(?:;|\{)"
    )
    if not _rust_root_level_matches(package_source, private_module):
        violations.append(
            "focused litchi-numbers Date & Time-format public API is missing its "
            f"private package owner module: {package_path.relative_to(root)}"
        )
    public_module = re.compile(
        r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?table_cell_date_time_format\b"
    )
    public_matches = _rust_root_level_matches(package_source, public_module)
    if public_matches:
        line_number = package_source.count("\n", 0, public_matches[0].start()) + 1
        violations.append(
            "focused litchi-numbers Date & Time-format public API exposes its "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    expected_signatures = {
        "table_cell_date_time_format": (r"Option\s*<\s*DateTime\s*>", False),
        "edit_table_cell_date_time_format": (r"\bEdit\b", False),
        "apply_table_cell_date_time_format": (r"\bPatch\b", True),
    }
    for method in NUMBERS_TABLE_CELL_DATE_TIME_FORMAT_OWNER_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Date & Time-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        signature, needs_commit = expected_signatures[method]
        if not re.search(signature, declaration):
            violations.append(
                "focused litchi-numbers Date & Time-format Package method has an "
                f"unexpected signature {method}: {owner_path.relative_to(root)}:{line_number}"
            )
        if method != "apply_table_cell_date_time_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Date & Time-format Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{owner_path.relative_to(root)}:{line_number}"
                    )
        if needs_commit and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Date & Time-format apply must return Commit: "
                f"{owner_path.relative_to(root)}:{line_number}"
            )

    # Public declarations in the semantic owner may expose only validated
    # values and selectors. Masked declarations make comments/strings unable
    # to satisfy these checks, while private implementation details remain
    # outside the public-declaration scanner.
    for path, source in (
        (semantic_path, semantic_source),
        (owner_path, owner_source),
    ):
        for declaration, line_number in _rust_public_declarations(source):
            for identifier_match in RUST_IDENTIFIER.finditer(declaration):
                identifier = identifier_match.group(1)
                reason = _numbers_table_cell_date_time_format_public_leak(identifier)
                if reason is None:
                    continue
                identifier_line = line_number + declaration.count(
                    "\n", 0, identifier_match.start(1)
                )
                violations.append(
                    "focused litchi-numbers Date & Time-format public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for pattern, label in (
                (RUST_BYTE_SLICE, "raw byte slice"),
                (NUMBERS_TABLE_CELL_NUMBER_FORMAT_PUBLIC_RAW_CONTAINER, "raw byte container"),
                (NUMBERS_TABLE_CELL_NUMBER_FORMAT_TYPED_RAW_ID_PARAMETER, "typed raw identifier parameter"),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    violations.append(
                        "focused litchi-numbers Date & Time-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{line_number}"
                    )
    return sorted(set(violations))


def _numbers_table_cell_duration_format_owner_present(root: Path) -> bool:
    """Return whether the focused Duration-format owner is wired at the root."""

    owner_path = root / NUMBERS_TABLE_CELL_DURATION_FORMAT_OWNER_SOURCE
    semantic_path = root / NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE
    package_path = root / Path("crates/litchi-numbers/src/package.rs")
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and semantic_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and bool(_mask_rust_non_code(semantic_source).strip())
        and bool(
            _rust_root_level_matches(
                package_source,
                re.compile(
                    r"(?m)^[ \t]*pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+mod"
                    r"[ \t\r\n]+(?:r#)?table_cell_duration_format\b"
                ),
            )
        )
    )


def _numbers_table_cell_duration_format_transaction_source(
    root: Path, semantic_source: str
) -> str:
    """Return the root transaction module body or its explicit split source."""

    declarations = _rust_root_level_module_declarations(
        semantic_source, frozenset({"transaction"})
    )
    for name, visibility, shape, body, _line_number in declarations:
        if name != "transaction" or visibility is None or not re.fullmatch(
            r"pub[ \t\r\n]+", visibility
        ):
            continue
        if shape == "inline":
            return body
        if shape == "external":
            external = root / NUMBERS_TABLE_CELL_DURATION_FORMAT_TRANSACTION_SOURCE
            if external.is_file():
                return _mask_rust_cfg_test_items(external.read_text(encoding="utf-8"))
    return ""


def _numbers_table_cell_duration_format_public_leak(
    identifier: str,
) -> str | None:
    """Classify native, archive, wire, generated, or host vocabulary."""

    if identifier == "litchi_iwa_common":
        # The neutral table vocabulary owner.
        return None
    if identifier in NUMBERS_TABLE_CELL_DURATION_FORMAT_RAW_MARKERS:
        return "raw identifier or archive type"
    if identifier in NUMBERS_TABLE_CELL_DURATION_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_DURATION_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_DURATION_FORMAT_WIRE_TYPES:
        return "wire type"
    if identifier in {
        "Draft",
        "Draft",
        "Draft",
        "LitchiIwaPackage",
    }:
        return "monolithic host type"
    if identifier in {"Projection", "LitchiIwaProjection"}:
        return "generated type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    if "generated" in words:
        return "generated type"
    return _iwork_public_leak(identifier)


def audit_numbers_table_cell_duration_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Duration-format codec seam."""

    if not _numbers_table_cell_duration_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_path = root / NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE
    shared_path = root / NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE
    codec_paths = tuple(
        root / path
        for path in NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    )
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)
    if not codec_path.is_file() or not _mask_rust_non_code(
        _mask_rust_cfg_test_items(
            codec_path.read_text(encoding="utf-8")
            if codec_path.is_file()
            else ""
        )
    ).strip():
        violations.append(
            "focused litchi-numbers Duration-format boundary is missing strict "
            f"codec source: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )
    if not shared_path.is_file():
        violations.append(
            "focused litchi-numbers Duration-format boundary is missing its shared "
            f"codec core: {NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE}"
        )

    for api in NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        violations.append(
            "focused litchi-numbers Duration-format hidden codec is missing "
            f"strict API {api}: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )

    # The wrapper may expose the family-specific spelling, but the strict
    # implementation must remain in the shared source-preserving walker.
    shared_code = _mask_rust_non_code(
        _mask_rust_cfg_test_items(
            shared_path.read_text(encoding="utf-8") if shared_path.is_file() else ""
        )
    )
    for api in NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_REQUIRED_APIS:
        if api in {
            "DurationStyle",
            "DurationUnit",
            "DurationFormatSnapshot",
            "DurationFormatWrite",
            "PreparedDurationFormatRewrite",
            "PreparedDurationFormatWrite",
        }:
            continue
        if _numbers_table_cell_number_format_api_present(shared_code, api):
            continue
        violations.append(
            "focused litchi-numbers Duration-format shared codec core is missing "
            f"strict API {api}: {NUMBERS_TABLE_CELL_POP_UP_MENU_CODEC_SOURCE}"
        )

    for marker, description in (
        ("buffa", "Buffa projection ingress"),
        ("decode_lazy_view", "lazy Buffa view"),
        ("Budget::new", "bounded wire preflight"),
        ("unknown", "unknown-field preservation"),
        ("extend_from_slice", "source-preserving rewrite"),
        ("execution_requirements", "prepared bounded rewrite"),
        ("execute", "prepared bounded execution"),
        ("scan_duration_format", "strict Duration scanner"),
        ("buffa_duration_format_parity", "Duration Buffa parity check"),
    ):
        if re.search(re.escape(marker), codec_code, re.IGNORECASE):
            continue
        violations.append(
            "focused litchi-numbers Duration-format hidden codec is missing "
            f"{description}: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )

    if re.search(
        r"\b(?:NATIVE_DURATION_FORMAT_TYPE|DURATION_FORMAT_TYPE)\b"
        r"\s*(?::\s*(?:u\d+\s*)?)?=\s*268\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Duration-format hidden codec must retain "
            f"native type 268: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Duration-format hidden codec retains a "
            f"Prost production path: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )

    # ``Budget::new`` is the shared scanner's preflight marker.  Require it to
    # occur before the first lazy-view force so a wrapper cannot enter the
    # generated view before applying the wire limits.
    lazy = [
        match.start()
        for match in re.finditer(
            r"\b(?:decode_lazy_view|decode_view|LazyView)\b", codec_code
        )
    ]
    preflight = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|Budget\s*::\s*new|scan_duration_format)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy and (not preflight or min(preflight) > min(lazy)):
        violations.append(
            "focused litchi-numbers Duration-format hidden codec must preflight "
            f"wire before forcing its lazy Buffa view: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )

    dedicated_code = _mask_rust_non_code(
        _mask_rust_cfg_test_items(
            codec_path.read_text(encoding="utf-8") if codec_path.is_file() else ""
        )
    )
    if re.search(
        r"\buse\s+crate\s*::\s*numbers_table_cell_pop_up_menu_codec\s+as\s+core\b",
        dedicated_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Duration-format hidden codec must route through "
            f"the shared core: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )
    if codec_path.is_file() and re.search(r"\bcore\s*::", dedicated_code) is None:
        violations.append(
            "focused litchi-numbers Duration-format hidden codec has no shared-core "
            f"calls: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )

    public_path = root / NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_SOURCE
    public_raw = (
        public_path.read_text(encoding="utf-8") if public_path.is_file() else ""
    )
    public_production = _mask_rust_cfg_test_items(public_raw)
    if not _rust_root_level_matches(
        public_production, NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_HIDDEN_MODULE
    ):
        violations.append(
            "focused litchi-numbers Duration-format public API is missing hidden "
            f"codec module {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    generated = _rust_root_level_module_declarations(
        public_production,
        frozenset({NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_GENERATED_MODULE}),
    )
    if not generated:
        violations.append(
            "focused litchi-numbers Duration-format public API is missing a "
            f"private Buffa generated projection: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    for _module, visibility, shape, body, line_number in generated:
        if visibility is not None:
            violations.append(
                "focused litchi-numbers Duration-format generated projection "
                f"must remain private: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline"
            and (
                not _rust_generated_module_include_is_valid(body)
                or re.search(
                    r"buffa[^\"\n]*duration[^\"\n]*|"
                    r"duration[^\"\n]*buffa[^\"\n]*",
                    _mask_rust_comments(body),
                    re.IGNORECASE,
                )
                is None
            )
        ):
            violations.append(
                "focused litchi-numbers Duration-format generated projection "
                f"module has invalid include shape: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        public_production,
        frozenset({NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_GENERATED_MODULE}),
    ):
        violations.append(
            "focused litchi-numbers Duration-format public API reexports its "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_PUBLIC_SOURCE}"
        )

    # The dedicated wrapper must carry an explicit hostile-wire test module;
    # shared-core tests alone are not sufficient evidence for this family
    # boundary.
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,4096}"
        r"#\s*\[\s*test\s*\]",
        _mask_rust_non_code(
            codec_path.read_text(encoding="utf-8") if codec_path.is_file() else ""
        ),
    ) is None:
        violations.append(
            "focused litchi-numbers Duration-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {NUMBERS_TABLE_CELL_DURATION_FORMAT_CODEC_SOURCE}"
        )
    return sorted(set(violations))


def audit_numbers_table_cell_duration_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the archive-free, selector-first Duration semantic facade."""

    if not (
        (root / NUMBERS_SOURCE_ROOT).is_dir()
        and _numbers_table_cell_duration_format_owner_present(root)
    ):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE
    semantic_source = _mask_rust_cfg_test_items(
        semantic_path.read_text(encoding="utf-8")
    )
    if not _rust_root_level_matches(
        semantic_source,
        re.compile(r"(?m)^[ \t]*pub\s+struct\s+(?:r#)?Duration\b"),
    ):
        violations.append(
            "focused litchi-numbers Duration-format public API is missing "
            f"canonical cell::data_format::Duration: {NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE}"
        )
    for semantic_type in NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_TYPES:
        if _rust_root_level_matches(
            semantic_source,
            re.compile(
                rf"(?m)^[ \t]*pub\s+(?:struct|enum|type)\s+(?:r#)?"
                rf"{re.escape(semantic_type)}\b"
            ),
        ):
            continue
        violations.append(
            "focused litchi-numbers Duration-format public API is missing "
            f"semantic type {semantic_type}: {NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE}"
        )

    transaction_declarations = _rust_root_level_module_declarations(
        semantic_source, frozenset({"transaction"})
    )
    transaction_modules = [
        declaration
        for declaration in transaction_declarations
        if declaration[1] is not None
        and re.fullmatch(r"pub[ \t\r\n]+", declaration[1])
    ]
    if not transaction_modules:
        violations.append(
            "focused litchi-numbers Duration-format public API is missing its "
            f"transaction namespace: {NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_source = _numbers_table_cell_duration_format_transaction_source(
        root, semantic_source
    )
    transaction_code = _mask_rust_non_code(transaction_source)
    transaction_exports = _rust_canonical_exports(
        transaction_source,
        frozenset(NUMBERS_TABLE_CELL_DURATION_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_DURATION_FORMAT_TRANSACTION_TYPES:
        if name in transaction_exports:
            continue
        violations.append(
            "focused litchi-numbers Duration-format transaction namespace is missing "
            f"canonical export {name}: {NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE}"
        )
    if not re.search(
        r"\bpub\s+use\s+crate\s*::\s*package\s*::\s*"
        r"table_cell_duration_format\b",
        transaction_code,
    ):
        violations.append(
            "focused litchi-numbers Duration-format transaction namespace must be "
            "source-bound to its private package owner: "
            f"{NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE}"
        )

    # Transaction types belong below the nested namespace; a root flat alias
    # silently recreates the retired monolithic surface.
    semantic_code = _mask_rust_non_code(semantic_source)
    for use_match in _rust_root_level_matches(
        semantic_source,
        re.compile(r"(?m)^[ \t]*pub\s+use\b"),
    ):
        statement_end = _rust_statement_end(semantic_code, use_match.end())
        declaration = semantic_code[use_match.start() : statement_end]
        for name in NUMBERS_TABLE_CELL_DURATION_FORMAT_TRANSACTION_TYPES:
            if not re.search(rf"\b{re.escape(name)}\b", declaration):
                continue
            violations.append(
                "focused litchi-numbers Duration-format public API retains flat "
                f"transaction export {name}: {NUMBERS_TABLE_CELL_DURATION_FORMAT_SEMANTIC_SOURCE}"
            )

    data_format_path = root / Path("crates/litchi-numbers/src/cell/data_format.rs")
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    if not _rust_root_level_matches(
        data_format_source,
        re.compile(
            r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?duration\b[ \t\r\n]*(?:;|\{)"
        ),
    ):
        violations.append(
            "focused litchi-numbers Duration-format public API is missing its "
            f"canonical duration module: {data_format_path.relative_to(root)}"
        )
    if not _rust_root_level_matches(
        data_format_source,
        re.compile(
            r"(?m)^[ \t]*pub\s+use\s+(?:r#)?duration\s*::\s*"
            r"(?:r#)?Duration\b"
        ),
    ):
        violations.append(
            "focused litchi-numbers Duration-format public API is missing its "
            f"canonical Duration reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / Path("crates/litchi-numbers/src/package.rs")
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    private_module = re.compile(
        r"(?m)^[ \t]*pub\s*\([ \t]*crate[ \t]*\)\s+mod\s+"
        r"(?:r#)?table_cell_duration_format\b[ \t\r\n]*(?:;|\{)"
    )
    if not _rust_root_level_matches(package_source, private_module):
        violations.append(
            "focused litchi-numbers Duration-format public API is missing its "
            f"private package owner module: {package_path.relative_to(root)}"
        )
    public_module = re.compile(
        r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?table_cell_duration_format\b"
    )
    public_matches = _rust_root_level_matches(package_source, public_module)
    if public_matches:
        line_number = package_source.count("\n", 0, public_matches[0].start()) + 1
        violations.append(
            "focused litchi-numbers Duration-format public API exposes its "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )
    for declaration, line_number in _rust_public_declarations(package_source):
        if re.search(
            r"\bpub\s+use[^;]*\btable_cell_duration_format\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Duration-format public API exposes its "
                f"package owner path: {package_path.relative_to(root)}:{line_number}"
            )

    owner_path = root / NUMBERS_TABLE_CELL_DURATION_FORMAT_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner_code = _mask_rust_non_code(owner_source)
    for label, marker in NUMBERS_TABLE_CELL_DURATION_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-numbers Duration-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    expected_signatures = {
        "table_cell_duration_format": (r"Option\s*<\s*Duration\s*>", False),
        "edit_table_cell_duration_format": (r"\bEdit\b", False),
        "apply_table_cell_duration_format": (r"\bCommit\b", True),
    }
    for method in NUMBERS_TABLE_CELL_DURATION_FORMAT_OWNER_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Duration-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        signature, needs_commit = expected_signatures[method]
        if not re.search(signature, declaration):
            violations.append(
                "focused litchi-numbers Duration-format Package method has an "
                f"unexpected signature {method}: {owner_path.relative_to(root)}:{line_number}"
            )
        if method != "apply_table_cell_duration_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Duration-format Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{owner_path.relative_to(root)}:{line_number}"
                    )
        elif not re.search(r"&\s*Patch\b", declaration):
            violations.append(
                "focused litchi-numbers Duration-format apply must accept its "
                f"source-bound Patch: {owner_path.relative_to(root)}:{line_number}"
            )
        if needs_commit and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Duration-format apply must return Commit: "
                f"{owner_path.relative_to(root)}:{line_number}"
            )

    # Scan only public declarations.  Private native implementation fields and
    # imports are allowed; public values, transaction exports, and signatures
    # must not expose archive IDs, raw payloads, wire records, or generated
    # projection types.  Rust trivia is masked before every hostile check.
    for path, source in (
        (semantic_path, semantic_source),
        (owner_path, owner_source),
    ):
        for declaration, line_number in _rust_public_declarations(source):
            for identifier_match in RUST_IDENTIFIER.finditer(declaration):
                identifier = identifier_match.group(1)
                reason = _numbers_table_cell_duration_format_public_leak(identifier)
                if reason is None:
                    continue
                identifier_line = line_number + declaration.count(
                    "\n", 0, identifier_match.start(1)
                )
                violations.append(
                    "focused litchi-numbers Duration-format public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for pattern, label in (
                (RUST_BYTE_SLICE, "raw byte slice"),
                (
                    NUMBERS_TABLE_CELL_DURATION_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_DURATION_FORMAT_PUBLIC_RAW_CONTAINER,
                    "raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_DURATION_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    violations.append(
                        "focused litchi-numbers Duration-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{line_number}"
                    )
            if re.match(
                r"^[ \t]*pub(?:\s*\([^()]*\))?\s+(?:async\s+|const\s+|unsafe\s+)*fn\b",
                declaration,
            ):
                for match in NUMBERS_TABLE_CELL_DURATION_FORMAT_PUBLIC_SELECTOR_RAW_PARAMETER.finditer(
                    declaration
                ):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    violations.append(
                        "focused litchi-numbers Duration-format public API exposes "
                        f"selector raw parameter {value}: {path.relative_to(root)}:{line_number}"
                    )
    return sorted(set(violations))


def _numbers_table_cell_custom_format_owner_present(root: Path) -> bool:
    """Return whether the focused document-scoped Custom owner is active."""

    owner_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_SOURCE
    semantic_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE
    package_path = root / NUMBERS_PACKAGE_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    private_module = re.compile(
        r"(?m)^[ \t]*pub\s*\([ \t]*crate[ \t]*\)\s+mod\s+"
        r"(?:r#)?table_cell_custom_format\b"
    )
    return (
        owner_path.is_file()
        and semantic_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and bool(_mask_rust_non_code(semantic_source).strip())
        and bool(_rust_root_level_matches(package_source, private_module))
    )


def _numbers_table_cell_custom_format_public_leak(
    identifier: str,
) -> str | None:
    """Classify physical custom-registry vocabulary forbidden by the facade."""

    if identifier in {"Uuid", "UUID", "ObjectId", "ObjectID", "FormatId", "FormatID"}:
        return "native UUID or raw identifier"
    if identifier in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost", "codec"} for word in words):
        return "protobuf type"
    if identifier in {"Member", "MemberName", "Payload", "Raw", "Archive"}:
        return "native archive member or payload"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_custom_format_owner_declaration(
    declaration: str,
) -> bool:
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_PATH.search(declaration) is not None or bool(
        identifiers
        & (
            set(NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_METHODS)
            | set(NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_TYPES)
            | set(NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TRANSACTION_TYPES)
            | {"custom", "custom_format"}
        )
    )


def _is_numbers_table_cell_custom_format_public_declaration(
    declaration: str,
    *,
    dedicated_source: bool,
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_FLAT_ALIASES
            | set(NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_TYPES)
            | {"Custom", "custom", "custom_format"}
        )
    ) or NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_FLAT_ALIAS_PATTERN.search(
        declaration
    ) is not None or _numbers_table_cell_custom_format_owner_declaration(declaration)


def _numbers_table_cell_custom_format_transaction_source(
    root: Path,
) -> tuple[Path, ...]:
    """Return the inline semantic source or a future split transaction leaf."""

    candidates = (
        NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TRANSACTION_SOURCE,
        Path("crates/litchi-numbers/src/cell/data_format/custom/transaction.rs"),
    )
    return tuple(root / path for path in candidates if (root / path).is_file())


def audit_numbers_table_cell_custom_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Custom-registry codec seam."""

    if not _numbers_table_cell_custom_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE
    codec_raw = codec_path.read_text(encoding="utf-8") if codec_path.is_file() else ""
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)
    if not codec_path.is_file() or not codec_code.strip():
        violations.append(
            "focused litchi-numbers Custom-format boundary is missing strict codec "
            f"source: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE}"
        )

    for api in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        violations.append(
            "focused litchi-numbers Custom-format hidden codec is missing strict "
            f"API {api}: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE}"
        )

    marker_groups = {
        "handwritten wire preflight": (
            ("Budget::new", "preflight"),
            ("scan_custom_format", "scan_custom_registry"),
            ("scan_custom_format_list", "parallel"),
        ),
        "private lazy Buffa projection": (
            ("decode_lazy_view", "LazyView"),
            ("buffa",),
            ("projection", "generated"),
        ),
        "strict hostile-field validation": (
            ("duplicate", "unique"),
            ("noncanonical", "canonical", "wire"),
            ("unknown", "raw"),
        ),
        "UUID/list shape validation": (
            ("uuid", "UUID"),
            ("count", "formats"),
            ("validate_unique_uuids", "unique"),
        ),
        "source-preserving bounded rewrite": (
            ("extend_from_slice", "raw"),
            ("execution_requirements", "Prepared"),
            ("execute", "rewrite"),
            ("retained", "scratch", "allocation"),
        ),
    }
    for label, groups in marker_groups.items():
        if all(
            any(re.search(re.escape(marker), codec_code, re.IGNORECASE) for marker in group)
            for group in groups
        ):
            continue
        violations.append(
            "focused litchi-numbers Custom-format hidden codec is missing strict "
            f"{label} marker: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE}"
        )

    for label, constant, value in (
        (
            "registry message type",
            "CUSTOM_FORMAT_REGISTRY_MESSAGE_TYPE",
            222,
        ),
        ("registry root field", "CUSTOM_FORMAT_REGISTRY_REFERENCE_FIELD", 9),
        ("custom format-list kind", "CUSTOM_FORMAT_LIST_KIND", 6),
        ("custom Number type", "NATIVE_CUSTOM_NUMBER_FORMAT_TYPE", 270),
        ("custom Text type", "NATIVE_CUSTOM_TEXT_FORMAT_TYPE", 271),
        (
            "custom Date & Time type",
            "NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE",
            272,
        ),
    ):
        if re.search(
            rf"\b{re.escape(constant)}\b\s*:\s*[^=;\n]+?=\s*{value}\b",
            codec_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Custom-format hidden codec is missing "
                f"{label} {value}: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE}"
            )

    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Custom-format hidden codec retains a Prost "
            f"production path: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE}"
        )

    lazy_positions = [
        match.start()
        for pattern in (r"\bdecode_lazy_view\b", r"\bLazyView\b")
        for match in re.finditer(pattern, codec_code)
    ]
    preflight_positions = [
        match.start()
        for pattern in (
            r"\bBudget\s*::\s*new\b",
            r"\bpreflight\b",
            r"\bscan_custom_format(?:_list)?\b",
        )
        for match in re.finditer(pattern, codec_code, re.IGNORECASE)
    ]
    if lazy_positions and (
        not preflight_positions or min(preflight_positions) > min(lazy_positions)
    ):
        violations.append(
            "focused litchi-numbers Custom-format hidden codec must preflight wire "
            f"before forcing its lazy Buffa view: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE}"
        )

    public_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE
    public_raw = public_path.read_text(encoding="utf-8") if public_path.is_file() else ""
    public_production = _mask_rust_cfg_test_items(public_raw)
    hidden_module = re.compile(
        rf"(?m)^[ \t]*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_MODULE)}\b"
    )
    if not _rust_root_level_matches(public_production, hidden_module):
        violations.append(
            "focused litchi-numbers Custom-format public API is missing hidden codec "
            f"module {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE}"
        )

    generated = _rust_root_level_module_declarations(
        public_production,
        frozenset({NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_GENERATED_MODULE}),
    )
    if not generated:
        violations.append(
            "focused litchi-numbers Custom-format public API is missing a private "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    for _module, visibility, shape, body, line_number in generated:
        if visibility is not None:
            violations.append(
                "focused litchi-numbers Custom-format generated projection module "
                f"must remain private: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "focused litchi-numbers Custom-format generated projection module "
                f"has invalid include shape: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        public_production,
        frozenset({NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_GENERATED_MODULE}),
    ):
        violations.append(
            "focused litchi-numbers Custom-format public API reexports its Buffa "
            f"generated projection: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_PUBLIC_SOURCE}"
        )

    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,4096}"
        r"#\s*\[\s*test\s*\]",
        _mask_rust_non_code(codec_raw),
    ) is None:
        violations.append(
            "focused litchi-numbers Custom-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_SOURCE}"
        )
    return sorted(set(violations))


def audit_numbers_table_cell_custom_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Custom-format facade."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_custom_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE
    semantic_source = _mask_rust_cfg_test_items(
        semantic_path.read_text(encoding="utf-8")
    )
    semantic_code = _mask_rust_non_code(semantic_source)

    # Root module wiring must be real production declarations at each level;
    # nested/test-only decoys do not activate a focused semantic facade.
    lib_path = root / NUMBERS_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    cell_module_path = root / NUMBERS_SOURCE_ROOT / "cell" / "mod.rs"
    cell_module_source = (
        _mask_rust_cfg_test_items(cell_module_path.read_text(encoding="utf-8"))
        if cell_module_path.is_file()
        else ""
    )
    data_format_path = root / NUMBERS_SOURCE_ROOT / "cell" / "data_format.rs"
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    wiring = (
        (
            lib_source,
            re.compile(r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?cell\b"),
            "root cell module",
            lib_path,
        ),
        (
            cell_module_source,
            re.compile(r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?data_format\b"),
            "cell::data_format module",
            cell_module_path,
        ),
        (
            data_format_source,
            re.compile(r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?custom\b"),
            "cell::data_format::custom module",
            data_format_path,
        ),
    )
    for source, pattern, label, path in wiring:
        if not _rust_root_level_matches(source, pattern):
            violations.append(
                "focused litchi-numbers Custom-format public API is missing "
                f"{label}: {path.relative_to(root)}"
            )
    if "Custom" not in _rust_root_canonical_exports(
        data_format_source, frozenset({"Custom"})
    ):
        violations.append(
            "focused litchi-numbers Custom-format public API is missing canonical "
            f"Custom reexport: {data_format_path.relative_to(root)}"
        )

    if not _rust_root_level_matches(
        semantic_source,
        re.compile(r"(?m)^[ \t]*pub\s+enum\s+(?:r#)?Custom\b"),
    ):
        violations.append(
            "focused litchi-numbers Custom-format public API is missing canonical "
            f"cell::data_format::Custom: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE}"
        )
    variants = _rust_public_enum_variants(semantic_source, "Custom")
    for variant in ("Number", "Text", "DateTime"):
        if variant not in variants:
            violations.append(
                "focused litchi-numbers Custom-format public API is missing "
                f"Custom::{variant}: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE}"
            )

    transaction_sources = _numbers_table_cell_custom_format_transaction_source(root)
    transaction_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in transaction_sources
    )
    inline_transaction = _rust_public_module_body(semantic_source, "transaction")
    transaction_scope = inline_transaction or transaction_source
    if inline_transaction is None and not transaction_source:
        violations.append(
            "focused litchi-numbers Custom-format public API is missing its "
            f"transaction namespace: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE}"
        )
    elif not _rust_root_level_matches(
        semantic_source,
        re.compile(r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?transaction\b"),
    ) and not re.search(
        r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?transaction\b", transaction_source
    ):
        violations.append(
            "focused litchi-numbers Custom-format public API is missing its "
            f"transaction namespace: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        transaction_scope,
        frozenset(NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers Custom-format public API is missing canonical "
                f"transaction reexport {name}: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE}"
            )
    if transaction_scope and re.search(
        r"\bpub\s+use\s+[^;]*package[^;]*table_cell_custom_format",
        transaction_scope,
    ) is None:
        violations.append(
            "focused litchi-numbers Custom-format transaction namespace must reexport "
            f"the private package owner: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_SEMANTIC_SOURCE}"
        )

    package_path = root / NUMBERS_PACKAGE_SOURCE
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    private_module = re.compile(
        r"(?m)^[ \t]*pub\s*\([ \t]*crate[ \t]*\)\s+mod\s+"
        r"(?:r#)?table_cell_custom_format\b"
    )
    if not _rust_root_level_matches(package_source, private_module):
        violations.append(
            "focused litchi-numbers Custom-format public API is missing private "
            f"package owner module: {package_path.relative_to(root)}"
        )
    public_module = re.compile(
        r"(?m)^[ \t]*pub\s+mod\s+(?:r#)?table_cell_custom_format\b"
    )
    public_matches = _rust_root_level_matches(package_source, public_module)
    if public_matches:
        line_number = package_source.count("\n", 0, public_matches[0].start()) + 1
        violations.append(
            "focused litchi-numbers Custom-format public API exposes its public "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )
    for declaration, line_number in _rust_public_declarations(package_source):
        if re.search(
            r"\b(?:table_cell_custom_format|crate\s*::\s*package\s*::\s*"
            r"table_cell_custom_format)\b",
            declaration,
        ) is None:
            continue
        if re.match(r"^[ \t]*pub\s+(?:use|type|const|static)\b", declaration):
            violations.append(
                "focused litchi-numbers Custom-format public API exposes a "
                f"package-owner alias: {package_path.relative_to(root)}:{line_number}"
            )

    owner_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner_helper_root = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    # Custom registry ownership is implemented by the shared private native
    # display-format core.  Include it for topology markers, but never scan
    # it as a public declaration: its raw IDs and member names are private by
    # construction.
    native_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_NATIVE_SOURCE
    native_source = (
        _mask_rust_cfg_test_items(native_path.read_text(encoding="utf-8"))
        if native_path.is_file()
        else ""
    )
    owner_code = _mask_rust_non_code(
        owner_related_source + "\n" + native_source
    )
    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    expected_signatures = {
        "table_cell_custom_format": (r"Option\s*<\s*Custom\s*>", False),
        "edit_table_cell_custom_format": (r"\bEdit\b", False),
        "apply_table_cell_custom_format": (r"\bPatch\b", True),
    }
    for method in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Custom-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        signature, needs_commit = expected_signatures[method]
        if not re.search(signature, declaration):
            violations.append(
                "focused litchi-numbers Custom-format Package method has an "
                f"unexpected signature {method}: {owner_path.relative_to(root)}:{line_number}"
            )
        if re.search(r"&\s*self\b", declaration) is None:
            violations.append(
                "focused litchi-numbers Custom-format Package method "
                f"{method} must retain its Package receiver: "
                f"{owner_path.relative_to(root)}:{line_number}"
            )
        if method != "apply_table_cell_custom_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Custom-format Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{owner_path.relative_to(root)}:{line_number}"
                    )
            selector_positions = [
                declaration.find(selector)
                for selector in ("SheetSelector", "TableSelector", "At")
            ]
            receiver_position = declaration.find("&self")
            if (
                receiver_position >= 0
                and all(position >= 0 for position in selector_positions)
                and (
                    receiver_position > min(selector_positions)
                    or selector_positions != sorted(selector_positions)
                )
            ):
                violations.append(
                    "focused litchi-numbers Custom-format Package method "
                    f"{method} must keep sheet/table/cell selectors first: "
                    f"{owner_path.relative_to(root)}:{line_number}"
                )
        if needs_commit and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Custom-format apply must return Commit: "
                f"{owner_path.relative_to(root)}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-numbers Custom-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    dedicated_sources = set(transaction_sources) | {owner_path}
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / NUMBERS_SOURCE_ROOT / "lib.rs",
        root / NUMBERS_SOURCE_ROOT / "cell" / "mod.rs",
        root / NUMBERS_SOURCE_ROOT / "cell" / "data_format.rs",
    }
    export_sources.update(path for path in export_sources if path.is_file())
    for path in sorted(dedicated_sources | export_sources):
        if not path.is_file():
            continue
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_custom_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
            if (
                public_use_or_type
                and "*" in declaration
                and any(
                    token in identifiers
                    for token in ("custom", "custom_format", "table_cell_custom_format")
                )
            ):
                violations.append(
                    "focused litchi-numbers Custom-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for identifier_match in RUST_IDENTIFIER.finditer(declaration):
                identifier = identifier_match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, identifier_match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Custom-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_custom_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Custom-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for pattern, label in (
                (RUST_BYTE_SLICE, "raw byte slice"),
                (
                    NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_RAW_CONTAINER,
                    "opaque raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_CUSTOM_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_UUID_PARAMETER,
                    "native UUID parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_CUSTOM_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw identifier/member/payload parameter",
                ),
                (NUMBERS_TABLE_CELL_CUSTOM_FORMAT_RAW_HELPER, "raw registry helper"),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    value_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        "focused litchi-numbers Custom-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{value_line}"
                    )

    example_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Custom-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_CUSTOM_FORMAT_EXAMPLE}"
        )
    else:
        example_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(example_path.read_text(encoding="utf-8"))
        )
        if re.search(
            r"\b(?:table_cell_custom_format|edit_table_cell_custom_format|"
            r"apply_table_cell_custom_format)\s*\(",
            example_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Custom-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Custom-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = integration_path.read_text(encoding="utf-8")
        integration_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_source)
        )
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Custom-format integration test is missing "
                f"#[test] coverage: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_SOURCE}"
            )
        for label, marker in NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_code) is None:
                violations.append(
                    "focused litchi-numbers Custom-format integration test is missing "
                    f"{label}: {NUMBERS_TABLE_CELL_CUSTOM_FORMAT_INTEGRATION_SOURCE}"
                )

    for fuzz_path, label in (
        (NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_FUZZ_SOURCE, "codec fuzz target"),
        (NUMBERS_TABLE_CELL_CUSTOM_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Custom-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Custom-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_CUSTOM_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_CUSTOM_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Custom-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Custom-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )

    return sorted(set(violations))


def _numbers_table_cell_percentage_format_codec_paths(
    root: Path,
) -> tuple[Path, ...]:
    """Return the Percentage wrapper and generalized strict codec sources."""

    candidates = (
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_SOURCE,
        NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_SOURCE,
        Path("crates/litchi-iwa-protos/src/numbers_table_cell_pop_up_menu_codec.rs"),
    )
    return tuple(root / path for path in candidates if (root / path).is_file())


def _numbers_table_cell_percentage_format_owner_present(root: Path) -> bool:
    """Return whether the independent Percentage owner has activated."""

    owner_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE.search(package_source)
        is not None
    )


def _numbers_table_cell_percentage_format_semantic_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return the canonical Percentage leaf plus an explicit split fallback."""

    canonical = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE
    if canonical.is_file():
        return (canonical,)
    fallback = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_FALLBACK_SOURCE
    return (fallback,) if fallback.is_file() else ()


def _numbers_table_cell_percentage_format_transaction_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return only sources that can own the Percentage transaction namespace."""

    candidates = (
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_SOURCE,
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE,
    )
    return tuple(root / path for path in candidates if (root / path).is_file())


def _numbers_table_cell_display_format_actual_call(
    source: str, names: tuple[str, ...]
) -> bool:
    """Find a call, rather than a declaration or import-only mention."""

    code = _mask_rust_non_code(source)
    for name in names:
        for match in re.finditer(
            rf"(?<![A-Za-z0-9_#]){re.escape(name)}\s*\(", code
        ):
            prefix = code[max(0, match.start() - 32) : match.start()]
            if re.search(r"\bfn\s*$", prefix):
                continue
            return True
    return False


def _numbers_table_cell_display_format_owner_core_route(
    source: str,
    *,
    forwarding_source: str = "",
) -> tuple[bool, bool]:
    """Return whether one owner imports and calls the shared display core.

    The Number owner predates the generalized display-format module and still
    reaches it through typed re-exports in ``table_cell_control_native``.  A
    plain module-name scan cannot see that transitive route, so recognize only
    the explicit private forwarding seam: the owner must import the control
    module under an alias, that module must re-export the focused Number read
    and rewrite functions from the shared core, and the owner must call those
    re-exported functions through the alias.  ``forwarding_source`` is kept
    optional so Percentage (and small synthetic fixtures) must use a direct
    shared-core import unless they deliberately provide this complete route.
    """

    code = _mask_rust_non_code(source)
    imported = re.search(
        r"\buse\b[^;]*\btable_cell_display_format_native\b", code
    ) is not None
    imported = imported or re.search(
        r"\b(?:use|mod)\b[^;\n]*\bdisplay_format_native\b", code
    ) is not None
    qualified = re.search(
        r"\b(?:table_cell_display_format_native|display_format_native)\s*::",
        code,
    ) is not None

    if not imported and forwarding_source:
        module = re.escape(NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_MODULE)
        aliases = {
            match.group("alias") or NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_MODULE
            for match in re.finditer(
                rf"\buse\b[^;]*\b{module}\b"
                r"(?:[ \t\r\n]+as[ \t\r\n]+(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?",
                code,
            )
        }
        forwarding_code = _mask_rust_non_code(forwarding_source)
        exported: set[str] = set()
        for match in re.finditer(
            r"\bpub(?:[ \t]*\([^()]*\))?[ \t\r\n]+use\b[^;]*"
            # The seam is the shared engine's `control` module,
            # which re-exports its sibling `display` core.
            r"\b(?:table_cell_display_format_native|super\s*::\s*display)\s*::\s*"
            r"\{(?P<items>[^}]*)\}",
            forwarding_code,
        ):
            exported.update(
                identifier.group(1)
                for identifier in RUST_IDENTIFIER.finditer(match.group("items"))
            )
        if (
            aliases
            and exported & NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_READ_APIS
            and exported & NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_REWRITE_APIS
            and any(
                re.search(
                    rf"\b{re.escape(alias)}\s*::\s*(?:"
                    rf"{'|'.join(map(re.escape, sorted(exported & NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_READ_APIS)))}"
                    r")\s*\(",
                    code,
                )
                for alias in aliases
            )
            and any(
                re.search(
                    rf"\b{re.escape(alias)}\s*::\s*(?:"
                    rf"{'|'.join(map(re.escape, sorted(exported & NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_REWRITE_APIS)))}"
                    r")\s*\(",
                    code,
                )
                for alias in aliases
            )
        ):
            imported = True
    return imported or qualified, qualified


def _numbers_table_cell_display_format_function_bodies(
    source: str,
) -> list[tuple[str, str, int]]:
    """Return all function bodies with offsets for duplicate-body checks."""

    code = _mask_rust_non_code(source)
    records: list[tuple[str, str, int]] = []
    for declaration in RUST_FUNCTION_DECLARATION.finditer(code):
        opening = code.find("{", declaration.end())
        if opening < 0:
            continue
        depth = 1
        cursor = opening + 1
        while cursor < len(code) and depth:
            if code[cursor] == "{":
                depth += 1
            elif code[cursor] == "}":
                depth -= 1
            cursor += 1
        if depth:
            continue
        records.append(
            (
                declaration.group(1),
                code[opening + 1 : cursor - 1],
                source.count("\n", 0, declaration.start()) + 1,
            )
        )
    return records


def _numbers_table_cell_display_format_writer_definitions(
    root: Path,
) -> dict[str, list[tuple[Path, int]]]:
    """Index native display-format writer/scanner definitions by name."""

    names = re.compile(
        # Keep the family token mandatory.  The generic control-cell helper
        # is intentionally named ``canonical_format`` and is not a decimal
        # display-format implementation.
        r"^(?:rewrite|canonical|prepare|emit|scan|decode|read|write)_"
        r"(?:display|number|percentage|currency|scientific|fraction)_format(?:_with_budget|_write)?$|"
        r"^(?:number|percentage|currency|scientific|fraction)_from_[A-Za-z0-9_]+$"
    )
    source_root = root / NUMBERS_SOURCE_ROOT / "package"
    files = sorted(source_root.rglob("*.rs")) if source_root.is_dir() else []
    # The shared core lives in `litchi-iwa-table`; index it too
    # so a second implementation beside it stays a duplicate.
    core_path = root / NUMBERS_TABLE_CELL_DISPLAY_FORMAT_NATIVE_SOURCE
    if core_path.is_file() and core_path not in files:
        files.append(core_path)
    indexed: dict[str, list[tuple[Path, int]]] = {}
    for path in files:
        production = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for name, _body, line_number in _numbers_table_cell_display_format_function_bodies(
            production
        ):
            if not names.fullmatch(name):
                continue
            indexed.setdefault(name, []).append((path, line_number))
    return indexed


def _numbers_table_cell_percentage_format_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden by the Percentage facade."""

    if identifier in {
        "Draft",
        "Draft",
        "Draft",
        "LitchiIwaPackage",
    }:
        return "monolithic host type"
    if identifier in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    if identifier in {"Projection", "LitchiIwaProjection"}:
        return "generated type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_percentage_format_owner_declaration(
    declaration: str,
) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_percentage_format_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIASES
            | {"Percentage", "percentage", "percentage_format"}
        )
    ) or any(
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(identifier)
        for identifier in identifiers
    ) or _numbers_table_cell_percentage_format_owner_declaration(declaration)


def _numbers_table_cell_currency_format_owner_present(root: Path) -> bool:
    """Return whether the independent Currency package owner is active."""

    owner_path = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and NUMBERS_PACKAGE_TABLE_CELL_CURRENCY_FORMAT_MODULE.search(package_source)
        is not None
    )


def _numbers_table_cell_currency_format_semantic_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return the canonical Currency leaf or its explicit split fallback."""

    canonical = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_SOURCE
    if canonical.is_file():
        return (canonical,)
    fallback = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_FALLBACK_SOURCE
    return (fallback,) if fallback.is_file() else ()


def _numbers_table_cell_currency_format_transaction_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return sources that can own the Currency transaction namespace."""

    candidates = (
        NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_SOURCE,
        NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_SOURCE,
    )
    return tuple(root / path for path in candidates if (root / path).is_file())


def _numbers_table_cell_currency_format_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden by the Currency facade."""

    if identifier in {
        "Draft",
        "Draft",
        "Draft",
        "LitchiIwaPackage",
    }:
        return "monolithic host type"
    if identifier in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    if identifier in {"Projection", "LitchiIwaProjection"}:
        return "generated type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_currency_format_owner_declaration(
    declaration: str,
) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_currency_format_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FLAT_ALIASES
            | {"Currency", "CurrencyCode", "CurrencyStyle", "currency", "currency_format"}
        )
    ) or any(
        NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(identifier)
        for identifier in identifiers
    ) or _numbers_table_cell_currency_format_owner_declaration(declaration)


def _numbers_table_cell_scientific_format_owner_present(root: Path) -> bool:
    """Return whether the independent Scientific package owner is active."""

    owner_path = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and bool(_mask_rust_non_code(owner_source).strip())
        and NUMBERS_PACKAGE_TABLE_CELL_SCIENTIFIC_FORMAT_MODULE.search(package_source)
        is not None
    )


def _numbers_table_cell_scientific_format_semantic_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return the canonical Scientific leaf or its explicit split fallback."""

    canonical = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_SOURCE
    if canonical.is_file():
        return (canonical,)
    fallback = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_FALLBACK_SOURCE
    return (fallback,) if fallback.is_file() else ()


def _numbers_table_cell_scientific_format_transaction_sources(
    root: Path,
) -> tuple[Path, ...]:
    """Return sources that can own the Scientific transaction namespace."""

    candidates = (
        NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_SOURCE,
        NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_SOURCE,
    )
    return tuple(root / path for path in candidates if (root / path).is_file())


def _numbers_table_cell_scientific_format_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden by the Scientific facade."""

    if identifier in {
        "Draft",
        "Draft",
        "Draft",
        "LitchiIwaPackage",
    }:
        return "monolithic host type"
    if identifier in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    if identifier in {"Projection", "LitchiIwaProjection"}:
        return "generated type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_scientific_format_owner_declaration(
    declaration: str,
) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_scientific_format_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FLAT_ALIASES
            | {"Scientific", "scientific", "scientific_format"}
        )
    ) or any(
        NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(identifier)
        for identifier in identifiers
    ) or _numbers_table_cell_scientific_format_owner_declaration(declaration)


def _numbers_table_cell_display_format_core_source(
    root: Path,
) -> tuple[Path, str]:
    path = root / NUMBERS_TABLE_CELL_DISPLAY_FORMAT_NATIVE_SOURCE
    if not path.is_file():
        return path, ""
    source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
    return path, _mask_rust_non_code(source)


def audit_numbers_table_cell_display_format_shared_core_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one actual shared native display-format graph owner.

    The audit intentionally activates only after the Number and Percentage
    owners exist. This keeps the historical Number-only boundary valid during
    an incremental migration while making every admitted decimal family share
    one scanner/writer graph.
    """

    if not (
        _numbers_table_cell_number_format_owner_present(root)
        and _numbers_table_cell_percentage_format_owner_present(root)
    ):
        return []

    violations: list[str] = []
    core_path, core = _numbers_table_cell_display_format_core_source(root)
    if not core:
        return [
            "focused Numbers decimal owners are missing the private "
            f"shared display-format core: {NUMBERS_TABLE_CELL_DISPLAY_FORMAT_NATIVE_SOURCE}"
        ]

    # The module itself must remain private to the Numbers package.
    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    package_code = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    if re.search(
        r"\bpub\s+mod\s+table_cell_display_format_native\b", package_code
    ) is not None:
        line_number = package_code.count(
            "\n",
            0,
            re.search(
                r"\bpub\s+mod\s+table_cell_display_format_native\b", package_code
            ).start(),
        ) + 1
        violations.append(
            "focused Numbers display-format graph core must remain private: "
            f"{package_path.relative_to(root)}:{line_number}"
        )
    if re.search(
        r"(?m)^\s*pub\s+mod\s+table_cell_display_format_native\b", core
    ) is not None:
        violations.append(
            "focused Numbers display-format graph core must not publicly define "
            f"itself: {core_path.relative_to(root)}"
        )

    # Definitions and calls are deliberately checked separately.  An import
    # is not evidence that either owner routes execution through the core.
    core_definitions = {
        name for name, _body, _line in _numbers_table_cell_display_format_function_bodies(core)
    }
    required_core_definitions = {
        "read": ("read_display_format_with_budget",),
        "rewrite": ("rewrite_display_format",),
        "prepare": ("prepare_display_rewrite", "prepare_display_append"),
    }
    for label, names in required_core_definitions.items():
        if not any(name in core_definitions for name in names):
            violations.append(
                "focused Numbers shared display-format core is missing its "
                f"{label} implementation: {core_path.relative_to(root)}"
            )

    owner_paths = [
        root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE,
        root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_SOURCE,
    ]
    currency_owner_path = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_SOURCE
    if _numbers_table_cell_currency_format_owner_present(root):
        owner_paths.append(currency_owner_path)
    scientific_owner_path = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_SOURCE
    if _numbers_table_cell_scientific_format_owner_present(root):
        owner_paths.append(scientific_owner_path)
    fraction_owner_path = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_SOURCE
    if _numbers_table_cell_fraction_format_owner_present(root):
        owner_paths.append(fraction_owner_path)
    for owner_path in owner_paths:
        owner_source = (
            _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
            if owner_path.is_file()
            else ""
        )
        forwarding_source = ""
        if owner_path == root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_OWNER_SOURCE:
            forwarding_path = root / NUMBERS_TABLE_CELL_DISPLAY_FORMAT_FORWARDING_SOURCE
            if forwarding_path.is_file():
                forwarding_source = _mask_rust_cfg_test_items(
                    forwarding_path.read_text(encoding="utf-8")
                )
        imported, qualified = _numbers_table_cell_display_format_owner_core_route(
            owner_source,
            forwarding_source=forwarding_source,
        )
        owner_code = _mask_rust_non_code(owner_source)
        if not imported:
            violations.append(
                "focused Numbers display-format owner must import the shared "
                f"core: {owner_path.relative_to(root)}"
            )
        for label, names in (
            (
                "read",
                (
                    "read_display_format_with_budget",
                    "read_number_format",
                    "read_number_format_with_budget",
                    "read_percentage_format",
                    "read_percentage_format_with_budget",
                    "read_currency_format",
                    "read_currency_format_with_budget",
                    "read_scientific_format",
                    "read_scientific_format_with_budget",
                    "read_fraction_format",
                    "read_fraction_format_with_budget",
                ),
            ),
            (
                "rewrite",
                (
                    "rewrite_display_format",
                    "rewrite_number_format",
                    "rewrite_percentage_format",
                    "rewrite_currency_format",
                    "rewrite_scientific_format",
                    "rewrite_fraction_format",
                ),
            ),
        ):
            if not _numbers_table_cell_display_format_actual_call(owner_code, names):
                violations.append(
                    "focused Numbers display-format owner must make an actual "
                    f"shared-core {label} call: {owner_path.relative_to(root)}"
                )
        if qualified and not re.search(
            r"\btable_cell_display_format_native\s*::\s*"
            r"(?:read|rewrite)_[A-Za-z0-9_]+\s*\(",
            owner_code,
        ):
            violations.append(
                "focused Numbers display-format owner has a shared-core module "
                f"reference without an executable call: {owner_path.relative_to(root)}"
            )

        for name, _body, line_number in _numbers_table_cell_display_format_function_bodies(
            owner_source
        ):
            if NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CORE_WRITER_DEFINITIONS.search(
                f"fn {name}"
            ) is not None:
                violations.append(
                    "focused Numbers display-format owner must not define a local "
                    f"scanner/writer {name}: {owner_path.relative_to(root)}:{line_number}"
                )

        for name, body, line_number in _numbers_table_cell_display_format_function_bodies(
            owner_source
        ):
            marker_hits = {
                marker
                for marker in NUMBERS_TABLE_CELL_DISPLAY_FORMAT_OWNER_LOCAL_GRAPH_MARKERS
                if re.search(rf"\b{re.escape(marker)}\b", body)
            }
            if len(marker_hits) < 2:
                continue
            violations.append(
                "focused Numbers display-format owner must not duplicate native "
                f"graph traversal in {name}: {owner_path.relative_to(root)}:{line_number}"
            )

    # Exactly one implementation may define each native writer/scanner.  The
    # wrappers in the two focused owners are allowed only when they call the
    # shared core; matching definitions outside that module are violations.
    writer_definitions = _numbers_table_cell_display_format_writer_definitions(root)
    for name, locations in sorted(writer_definitions.items()):
        if len(locations) <= 1:
            continue
        relative = ", ".join(
            f"{path.relative_to(root)}:{line}" for path, line in locations
        )
        violations.append(
            "focused Numbers display-format writer/scanner has duplicate "
            f"implementations {name}: {relative}"
        )
    for name, locations in sorted(writer_definitions.items()):
        for path, line in locations:
            if path == core_path:
                continue
            if path in owner_paths or "table_cell_control_native.rs" in path.parts:
                violations.append(
                    "focused Numbers display-format writer/scanner must be owned "
                    f"by the shared core {name}: {path.relative_to(root)}:{line}"
                )

    return sorted(set(violations))


def audit_numbers_table_cell_percentage_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Percentage-format codec seam."""

    if not _numbers_table_cell_percentage_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_paths = _numbers_table_cell_percentage_format_codec_paths(root)
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)
    codec_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_SOURCE
    if not codec_paths:
        violations.append(
            "focused litchi-numbers Percentage-format boundary is missing strict "
            f"codec source: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_SOURCE}"
        )

    for api in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        expected = " or ".join(api) if isinstance(api, tuple) else api
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec is missing "
            f"strict API {expected}: {codec_path}"
        )

    for label, markers in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_REQUIRED_MARKER_GROUPS.items():
        if all(
            _numbers_table_cell_number_format_marker_present(codec_code, marker)
            for marker in markers
        ):
            continue
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec is missing "
            f"strict {label} marker: {codec_path}"
        )

    # Type 258 must be explicit and must be checked against the decoded
    # discriminator.  Re-exporting the Number constant is not sufficient.
    if re.search(
        r"NATIVE_PERCENTAGE_FORMAT_TYPE\s*:\s*(?:u\d+\s*)?=\s*258\b|"
        r"NATIVE_PERCENTAGE_FORMAT_TYPE\b[\s\S]{0,200}\b258\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec must retain "
            f"native type 258: {codec_path}"
        )
    if re.search(r"\bNATIVE_NUMBER_FORMAT_TYPE\b[\s\S]{0,200}\b256\b", codec_code) is None:
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec must retain the "
            f"Number type 256 alongside type 258: {codec_path}"
        )
    if re.search(
        r"format_type\s*!=\s*(?:expected_format_type|NATIVE_PERCENTAGE_FORMAT_TYPE)|"
        r"expected_format_type\b[\s\S]{0,160}\bformat_type\b",
        codec_code,
        re.IGNORECASE,
    ) is None:
        violations.append(
            "focused litchi-numbers Percentage-format codec must check the "
            f"expected family discriminator: {codec_path}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec retains a "
            f"Prost production path: {codec_path}"
        )

    lazy_matches = [
        match.start()
        for pattern in (
            r"\bdecode_lazy_view\b",
            r"\bdecode_view\b",
            r"\bLazyView\b",
            r"\bvisitor\b",
        )
        for match in re.finditer(pattern, codec_code)
    ]
    preflight_positions = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|scan_percentage_message|scan_decimal_message|"
            r"scan_display_format|Budget\s*::\s*new)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy_matches and (
        not preflight_positions or min(preflight_positions) > min(lazy_matches)
    ):
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec must preflight "
            f"wire before forcing its lazy Buffa view: {codec_path}"
        )
    if re.search(r"\bbuffa\b", codec_code) is None:
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec is missing Buffa "
            f"projection ingress: {codec_path}"
        )

    codec_lib_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_PUBLIC_SOURCE
    codec_lib_raw = (
        codec_lib_path.read_text(encoding="utf-8")
        if codec_lib_path.is_file()
        else ""
    )
    codec_lib_code = _mask_rust_non_code(
        _mask_rust_cfg_test_items(codec_lib_raw)
    )
    codec_lib_production = _mask_rust_cfg_test_items(codec_lib_raw)
    hidden_modules = (
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_MODULE,
        NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_MODULE,
    )
    hidden_module_matches = tuple(
        _rust_root_level_matches(
            codec_lib_production,
            re.compile(
                rf"(?m)^[ \t]*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
                rf"pub\s+mod\s+{re.escape(module)}\b"
            ),
        )
        for module in hidden_modules
    )
    if not any(hidden_module_matches):
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing a "
            f"hidden codec module: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULES),
    )
    if not generated_declarations:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing a "
            f"private Buffa generated projection: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    for module, visibility, shape, body, line_number in generated_declarations:
        if visibility is not None:
            violations.append(
                "focused litchi-numbers Percentage-format generated projection "
                f"must remain private: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "focused litchi-numbers Percentage-format generated projection "
                f"module has invalid include shape: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_GENERATED_MODULES),
    ):
        violations.append(
            "focused litchi-numbers Percentage-format public API reexports its "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,1200}"
        r"#\s*\[\s*test\s*\]",
        codec_raw,
    ) is None:
        violations.append(
            "focused litchi-numbers Percentage-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {codec_path}"
        )
    return sorted(set(violations))


def audit_numbers_table_cell_currency_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Currency-format codec seam."""

    if not _numbers_table_cell_currency_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_paths = tuple(
        root / path
        for path in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    )
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)
    codec_path = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_SOURCE
    if not codec_paths:
        violations.append(
            "focused litchi-numbers Currency-format boundary is missing strict "
            f"codec source: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_SOURCE}"
        )

    for api in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        expected = " or ".join(api) if isinstance(api, tuple) else api
        violations.append(
            "focused litchi-numbers Currency-format hidden codec is missing "
            f"strict API {expected}: {codec_path}"
        )

    for label, markers in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_REQUIRED_MARKER_GROUPS.items():
        if all(
            _numbers_table_cell_number_format_marker_present(codec_code, marker)
            for marker in markers
        ):
            continue
        violations.append(
            "focused litchi-numbers Currency-format hidden codec is missing "
            f"strict {label} marker: {codec_path}"
        )

    if re.search(
        r"NATIVE_CURRENCY_FORMAT_TYPE\s*:\s*(?:u\d+\s*)?=\s*257\b|"
        r"NATIVE_CURRENCY_FORMAT_TYPE\b[\s\S]{0,200}\b257\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Currency-format hidden codec must retain "
            f"native type 257: {codec_path}"
        )
    if re.search(
        r"format_type\s*!=\s*(?:expected_format_type|NATIVE_CURRENCY_FORMAT_TYPE)|"
        r"expected_format_type\b[\s\S]{0,160}\bformat_type\b|"
        r"NATIVE_CURRENCY_FORMAT_TYPE\b[\s\S]{0,200}format_type\b",
        codec_code,
        re.IGNORECASE,
    ) is None:
        violations.append(
            "focused litchi-numbers Currency-format codec must check the "
            f"expected family discriminator: {codec_path}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Currency-format hidden codec retains a "
            f"Prost production path: {codec_path}"
        )
    if re.search(r"\bbuffa\b", codec_code) is None:
        violations.append(
            "focused litchi-numbers Currency-format hidden codec is missing Buffa "
            f"projection ingress: {codec_path}"
        )

    lazy_matches = [
        match.start()
        for pattern in (
            r"\bdecode_lazy_view\b",
            r"\bdecode_view\b",
            r"\bLazyView\b",
            r"\bvisitor\b",
        )
        for match in re.finditer(pattern, codec_code)
    ]
    preflight_positions = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|scan_currency_message|scan_decimal_message|"
            r"scan_display_format|Budget\s*::\s*new)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy_matches and (
        not preflight_positions or min(preflight_positions) > min(lazy_matches)
    ):
        violations.append(
            "focused litchi-numbers Currency-format hidden codec must preflight "
            f"wire before forcing its lazy Buffa view: {codec_path}"
        )

    codec_lib_path = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_PUBLIC_SOURCE
    codec_lib_raw = (
        codec_lib_path.read_text(encoding="utf-8")
        if codec_lib_path.is_file()
        else ""
    )
    codec_lib_code = _mask_rust_non_code(
        _mask_rust_cfg_test_items(codec_lib_raw)
    )
    codec_lib_production = _mask_rust_cfg_test_items(codec_lib_raw)
    hidden_modules = (
        NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_MODULE,
        NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_MODULE,
    )
    hidden_module_matches = tuple(
        _rust_root_level_matches(
            codec_lib_production,
            re.compile(
                rf"(?m)^[ \t]*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
                rf"pub\s+mod\s+{re.escape(module)}\b"
            ),
        )
        for module in hidden_modules
    )
    if not any(hidden_module_matches):
        violations.append(
            "focused litchi-numbers Currency-format public API is missing a "
            f"hidden codec module: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_GENERATED_MODULES),
    )
    if not generated_declarations:
        violations.append(
            "focused litchi-numbers Currency-format public API is missing a "
            f"private Buffa generated projection: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    for module, visibility, shape, body, line_number in generated_declarations:
        if visibility is not None:
            violations.append(
                "focused litchi-numbers Currency-format generated projection "
                f"must remain private: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "focused litchi-numbers Currency-format generated projection "
                f"module has invalid include shape: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_GENERATED_MODULES),
    ):
        violations.append(
            "focused litchi-numbers Currency-format public API reexports its "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,1200}"
        r"#\s*\[\s*test\s*\]",
        codec_raw,
    ) is None:
        violations.append(
            "focused litchi-numbers Currency-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {codec_path}"
        )
    return sorted(set(violations))


def audit_numbers_table_cell_fraction_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Fraction-format codec seam."""

    if not _numbers_table_cell_fraction_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_paths = tuple(
        root / path
        for path in NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    )
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)
    codec_path = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_SOURCE
    if not codec_paths:
        violations.append(
            "focused litchi-numbers Fraction-format boundary is missing strict "
            f"codec source: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_SOURCE}"
        )

    for api in NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        expected = " or ".join(api) if isinstance(api, tuple) else api
        violations.append(
            "focused litchi-numbers Fraction-format hidden codec is missing "
            f"strict API {expected}: {codec_path}"
        )

    for label, markers in NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_REQUIRED_MARKER_GROUPS.items():
        if all(
            _numbers_table_cell_number_format_marker_present(codec_code, marker)
            for marker in markers
        ):
            continue
        violations.append(
            "focused litchi-numbers Fraction-format hidden codec is missing "
            f"strict {label} marker: {codec_path}"
        )

    # Field 20 is a deliberately narrow compatibility escape hatch.  A
    # canonical false is source-authoritative (and therefore must survive a
    # rewrite), while true requests unsupported replacement semantics.  Keep
    # these checks structural so a prose marker or a bare field-name mention
    # cannot satisfy the ratchet.
    if re.search(
        r"(?:requires_fraction_replacement\s*=\s*Some\(\s*value\s*!=\s*0\s*\)"
        r"[\s\S]{0,500}value\s*>\s*1|"
        r"value\s*>\s*1[\s\S]{0,500}"
        r"requires_fraction_replacement\s*=\s*Some\(\s*value\s*!=\s*0\s*\))",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format codec must accept only "
            "canonical false for the field-20 replacement marker: "
            f"{codec_path}"
        )
    if re.search(
        r"(?:requires_fraction_replacement|Some\(true\))[\s\S]{0,320}"
        r"(?:Some\(true\)|value\s*!=\s*0)[\s\S]{0,320}"
        r"(?:return\s+Err|DecodeError::invalid)",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format codec must reject true "
            "field-20 replacement markers: "
            f"{codec_path}"
        )
    if re.search(
        r"view\.requires_fraction_replacement\s*!=\s*"
        r"snapshot\.requires_fraction_replacement|"
        r"requires_fraction_replacement\s*:\s*"
        r"state\.requires_fraction_replacement",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format codec must preserve the "
            "accepted field-20 false marker from source: "
            f"{codec_path}"
        )

    for name, body, line_number in _numbers_table_cell_display_format_function_bodies(
        codec_production
    ):
        if name not in {
            "canonical_fraction_format",
            "prepare_fraction_format_write",
            "prepare_fraction_format_append",
        }:
            continue
        if re.search(
            r"requires_fraction_replacement|FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD",
            body,
        ) is None:
            continue
        violations.append(
            "focused litchi-numbers Fraction-format codec must not synthesize "
            f"field-20 replacement markers in {name}: {codec_path}:{line_number}"
        )

    if re.search(
        r"FORMAT_TYPE_FIELD\s*\|\s*FORMAT_FRACTION_ACCURACY_FIELD",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format codec must restrict ordinary "
            "type-262 fields to the discriminator and fraction accuracy: "
            f"{codec_path}"
        )

    if re.search(
        r"NATIVE_FRACTION_FORMAT_TYPE\s*:\s*(?:u\d+\s*)?=\s*262\b|"
        r"NATIVE_FRACTION_FORMAT_TYPE\b[\s\S]{0,200}\b262\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format hidden codec must retain "
            f"native type 262: {codec_path}"
        )
    if re.search(
        r"format_type\s*!=\s*(?:expected_format_type|NATIVE_FRACTION_FORMAT_TYPE)|"
        r"expected_format_type\b[\s\S]{0,160}\bformat_type\b|"
        r"NATIVE_FRACTION_FORMAT_TYPE\b[\s\S]{0,200}format_type\b",
        codec_code,
        re.IGNORECASE,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format codec must check the "
            f"expected family discriminator: {codec_path}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Fraction-format hidden codec retains a "
            f"Prost production path: {codec_path}"
        )
    if re.search(r"\bbuffa\b", codec_code) is None:
        violations.append(
            "focused litchi-numbers Fraction-format hidden codec is missing Buffa "
            f"projection ingress: {codec_path}"
        )

    lazy_matches = [
        match.start()
        for pattern in (
            r"\bdecode_lazy_view\b",
            r"\bdecode_view\b",
            r"\bLazyView\b",
            r"\bvisitor\b",
        )
        for match in re.finditer(pattern, codec_code)
    ]
    preflight_positions = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|scan_fraction_message|scan_decimal_message|"
            r"scan_display_format|Budget\s*::\s*new)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy_matches and (
        not preflight_positions or min(preflight_positions) > min(lazy_matches)
    ):
        violations.append(
            "focused litchi-numbers Fraction-format hidden codec must preflight "
            f"wire before forcing its lazy Buffa view: {codec_path}"
        )

    codec_lib_path = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_PUBLIC_SOURCE
    codec_lib_raw = (
        codec_lib_path.read_text(encoding="utf-8")
        if codec_lib_path.is_file()
        else ""
    )
    codec_lib_code = _mask_rust_non_code(
        _mask_rust_cfg_test_items(codec_lib_raw)
    )
    codec_lib_production = _mask_rust_cfg_test_items(codec_lib_raw)
    hidden_modules = (
        NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_MODULE,
        NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_MODULE,
    )
    hidden_module_matches = tuple(
        _rust_root_level_matches(
            codec_lib_production,
            re.compile(
                rf"(?m)^[ \t]*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
                rf"pub\s+mod\s+{re.escape(module)}\b"
            ),
        )
        for module in hidden_modules
    )
    if not any(hidden_module_matches):
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing a "
            f"hidden codec module: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_GENERATED_MODULES),
    )
    if not generated_declarations:
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing a "
            f"private Buffa generated projection: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    for module, visibility, shape, body, line_number in generated_declarations:
        if visibility is not None:
            violations.append(
                "focused litchi-numbers Fraction-format generated projection "
                f"must remain private: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "focused litchi-numbers Fraction-format generated projection "
                f"module has invalid include shape: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_GENERATED_MODULES),
    ):
        violations.append(
            "focused litchi-numbers Fraction-format public API reexports its "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,1200}"
        r"#\s*\[\s*test\s*\]",
        codec_raw,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {codec_path}"
        )
    return sorted(set(violations))


def audit_numbers_table_cell_percentage_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Percentage-format owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_percentage_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_paths = _numbers_table_cell_percentage_format_semantic_sources(root)
    semantic_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in semantic_paths
    )
    semantic_code = _mask_rust_non_code(semantic_source)
    if re.search(
        r"\bpub\s+struct\s+(?:r#)?Percentage\b|"
        r"\bpub\s+use[^;]*\bPercentage\b|"
        r"\bdecimal_format!\s*\(\s*(?:r#)?Percentage\b",
        semantic_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing "
            f"canonical cell::data_format::Percentage: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE}"
        )

    transaction_paths = _numbers_table_cell_percentage_format_transaction_sources(root)
    transaction_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in transaction_paths
    )
    transaction_code = _mask_rust_non_code(transaction_source)
    if re.search(
        r"\bpub\s+mod\s+transaction\b", semantic_code
    ) is None and re.search(
        r"\bpub\s+mod\s+transaction\b", transaction_code
    ) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"percentage transaction namespace: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        transaction_source,
        frozenset(NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers Percentage-format public API is missing "
                f"canonical transaction type {name}: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_SOURCE}"
            )

    data_format_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    data_format_code = _mask_rust_non_code(data_format_source)
    if PUBLIC_NUMBERS_CELL_DATA_FORMAT_PERCENTAGE_MODULE.search(data_format_code) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"cell::data_format::percentage module: {data_format_path.relative_to(root)}"
        )
    if "Percentage" not in _rust_canonical_exports(
        data_format_source, frozenset({"Percentage"})
    ):
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"cell::data_format::Percentage reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    if NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"private package owner module: {package_path.relative_to(root)}"
        )
    if PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE.search(package_code) is not None:
        violation = PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE.search(package_code)
        line_number = package_code.count("\n", 0, violation.start()) + 1 if violation else 1
        violations.append(
            "focused litchi-numbers Percentage-format public API exposes its "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+(?:r#)?cell\b", _mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing root "
            f"cell module: {lib_path.relative_to(root)}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_helper_root = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    owner_code = _mask_rust_non_code(owner_related_source)
    owner_exports = _rust_canonical_exports(
        owner_related_source,
        frozenset(NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers Percentage-format public API is missing "
                f"canonical package type {name}: {owner_path.relative_to(root)}"
            )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PACKAGE_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Percentage-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method != "apply_table_cell_percentage_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Percentage-format Package method "
                        f"{method} must accept selector-first {selector}: {owner_path.relative_to(root)}:{line_number}"
                    )
        if method == "table_cell_percentage_format":
            normalized = re.sub(r"\s+", "", declaration)
            if "Option<Percentage>" not in normalized:
                violations.append(
                    "focused litchi-numbers Percentage-format read must return "
                    f"Option<Percentage>: {owner_path.relative_to(root)}:{line_number}"
                )
        elif method == "edit_table_cell_percentage_format":
            if not re.search(r"\bEdit\b", declaration):
                violations.append(
                    "focused litchi-numbers Percentage-format edit must return "
                    f"Edit: {owner_path.relative_to(root)}:{line_number}"
                )
        elif not re.search(r"\bPatch\b", declaration):
            violations.append(
                "focused litchi-numbers Percentage-format apply must accept "
                f"Patch: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "apply_table_cell_percentage_format" and not re.search(
            r"\bCommit\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Percentage-format apply must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )
        for pattern in (
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_PARAMETER,
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_RAW_ID_PARAMETER,
        ):
            for match in pattern.finditer(declaration):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                violations.append(
                    "focused litchi-numbers Percentage-format Package method exposes "
                    f"raw identifier/coordinate parameter {parameter}: {owner_path.relative_to(root)}:{line_number}"
                )

    edit_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Edit"
        )
    }
    for method in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EDIT_METHODS:
        record = edit_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Percentage-format edit is missing "
                f"Edit::{method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method == "set" and not re.search(r"\bPercentage\b", declaration):
            violations.append(
                "focused litchi-numbers Percentage-format Edit::set must accept "
                f"Percentage: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "commit" and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Percentage-format Edit::commit must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_related_source) is None:
            violations.append(
                "focused litchi-numbers Percentage-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    dedicated_sources = {owner_path, *semantic_paths, *transaction_paths}
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in (
            *NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[:3],
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE,
        )
        if (root / path).is_file()
    }
    export_sources.update(semantic_paths)
    export_sources.update(transaction_paths)
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_percentage_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (
                ["pub", "use"],
                ["pub", "type"],
            )
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_percentage_format" in identifiers
                    or "percentage_format" in identifiers
                    or "percentage" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers Percentage-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Percentage-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    not dedicated_source
                    and identifier
                    in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers Percentage-format public API retains "
                        f"root transaction alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_percentage_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Percentage-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for pattern, label in (
                (
                    RUST_BYTE_SLICE,
                    "raw byte slice",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_CONTAINER,
                    "opaque raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE,
                    "opaque raw byte container type",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_RAW_HELPER,
                    "raw helper",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw identifier/coordinate parameter",
                ),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    value_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        "focused litchi-numbers Percentage-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{value_line}"
                    )

    example_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Percentage-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EXAMPLE}"
        )
    else:
        example_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(example_path.read_text(encoding="utf-8"))
        )
        if re.search(
            r"\b(?:edit_table_cell_percentage_format|"
            r"apply_table_cell_percentage_format)\s*\(",
            example_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Percentage-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Percentage-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = integration_path.read_text(encoding="utf-8")
        integration_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_source)
        )
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Percentage-format integration test is missing "
                f"#[test] coverage: {integration_path}"
            )
        for label, marker in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_source) is None:
                violations.append(
                    "focused litchi-numbers Percentage-format integration test is missing "
                    f"{label}: {integration_path}"
                )

    for fuzz_path, label in (
        (
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_FUZZ_SOURCE,
            "codec fuzz target",
        ),
        (NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Percentage-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Percentage-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Percentage-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Percentage-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )
    return sorted(set(violations))


def audit_numbers_table_cell_scientific_format_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the private strict Buffa/lazy Scientific-format codec seam."""

    if not _numbers_table_cell_scientific_format_owner_present(root):
        return []

    violations: list[str] = []
    codec_paths = tuple(
        root / path
        for path in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    )
    codec_raw = "\n".join(
        path.read_text(encoding="utf-8") for path in codec_paths
    )
    codec_production = _mask_rust_cfg_test_items(codec_raw)
    codec_code = _mask_rust_non_code(codec_production)
    codec_path = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_SOURCE
    if not codec_paths:
        violations.append(
            "focused litchi-numbers Scientific-format boundary is missing strict "
            f"codec source: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_SOURCE}"
        )

    for api in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_REQUIRED_APIS:
        if _numbers_table_cell_number_format_api_present(codec_code, api):
            continue
        expected = " or ".join(api) if isinstance(api, tuple) else api
        violations.append(
            "focused litchi-numbers Scientific-format hidden codec is missing "
            f"strict API {expected}: {codec_path}"
        )

    for label, markers in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_REQUIRED_MARKER_GROUPS.items():
        if all(
            _numbers_table_cell_number_format_marker_present(codec_code, marker)
            for marker in markers
        ):
            continue
        violations.append(
            "focused litchi-numbers Scientific-format hidden codec is missing "
            f"strict {label} marker: {codec_path}"
        )

    if re.search(
        r"NATIVE_SCIENTIFIC_FORMAT_TYPE\s*:\s*(?:u\d+\s*)?=\s*259\b|"
        r"NATIVE_SCIENTIFIC_FORMAT_TYPE\b[\s\S]{0,200}\b259\b",
        codec_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Scientific-format hidden codec must retain "
            f"native type 259: {codec_path}"
        )
    if re.search(
        r"format_type\s*!=\s*(?:expected_format_type|NATIVE_SCIENTIFIC_FORMAT_TYPE)|"
        r"expected_format_type\b[\s\S]{0,160}\bformat_type\b|"
        r"NATIVE_SCIENTIFIC_FORMAT_TYPE\b[\s\S]{0,200}format_type\b",
        codec_code,
        re.IGNORECASE,
    ) is None:
        violations.append(
            "focused litchi-numbers Scientific-format codec must check the "
            f"expected family discriminator: {codec_path}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", codec_code) is not None:
        violations.append(
            "focused litchi-numbers Scientific-format hidden codec retains a "
            f"Prost production path: {codec_path}"
        )
    if re.search(r"\bbuffa\b", codec_code) is None:
        violations.append(
            "focused litchi-numbers Scientific-format hidden codec is missing Buffa "
            f"projection ingress: {codec_path}"
        )

    lazy_matches = [
        match.start()
        for pattern in (
            r"\bdecode_lazy_view\b",
            r"\bdecode_view\b",
            r"\bLazyView\b",
            r"\bvisitor\b",
        )
        for match in re.finditer(pattern, codec_code)
    ]
    preflight_positions = [
        match.start()
        for match in re.finditer(
            r"\b(?:preflight|scan_scientific_message|scan_decimal_message|"
            r"scan_display_format|Budget\s*::\s*new)\b",
            codec_code,
            re.IGNORECASE,
        )
    ]
    if lazy_matches and (
        not preflight_positions or min(preflight_positions) > min(lazy_matches)
    ):
        violations.append(
            "focused litchi-numbers Scientific-format hidden codec must preflight "
            f"wire before forcing its lazy Buffa view: {codec_path}"
        )

    codec_lib_path = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_PUBLIC_SOURCE
    codec_lib_raw = (
        codec_lib_path.read_text(encoding="utf-8")
        if codec_lib_path.is_file()
        else ""
    )
    codec_lib_code = _mask_rust_non_code(
        _mask_rust_cfg_test_items(codec_lib_raw)
    )
    codec_lib_production = _mask_rust_cfg_test_items(codec_lib_raw)
    hidden_modules = (
        NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_MODULE,
        NUMBERS_TABLE_CELL_DISPLAY_FORMAT_CODEC_MODULE,
    )
    hidden_module_matches = tuple(
        _rust_root_level_matches(
            codec_lib_production,
            re.compile(
                rf"(?m)^[ \t]*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
                rf"pub\s+mod\s+{re.escape(module)}\b"
            ),
        )
        for module in hidden_modules
    )
    if not any(hidden_module_matches):
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing a "
            f"hidden codec module: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_GENERATED_MODULES),
    )
    if not generated_declarations:
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing a "
            f"private Buffa generated projection: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    for module, visibility, shape, body, line_number in generated_declarations:
        if visibility is not None:
            violations.append(
                "focused litchi-numbers Scientific-format generated projection "
                f"must remain private: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "focused litchi-numbers Scientific-format generated projection "
                f"module has invalid include shape: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        codec_lib_production,
        frozenset(NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_GENERATED_MODULES),
    ):
        violations.append(
            "focused litchi-numbers Scientific-format public API reexports its "
            f"Buffa generated projection: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        r"#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\][\s\S]{0,1200}"
        r"#\s*\[\s*test\s*\]",
        codec_raw,
    ) is None:
        violations.append(
            "focused litchi-numbers Scientific-format hidden codec is missing its "
            f"cfg(test) hostile-wire harness: {codec_path}"
        )
    return sorted(set(violations))


def audit_numbers_table_cell_percentage_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Percentage-format owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_percentage_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_paths = _numbers_table_cell_percentage_format_semantic_sources(root)
    semantic_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in semantic_paths
    )
    semantic_code = _mask_rust_non_code(semantic_source)
    if re.search(
        r"\bpub\s+struct\s+(?:r#)?Percentage\b|"
        r"\bpub\s+use[^;]*\bPercentage\b|"
        r"\bdecimal_format!\s*\(\s*(?:r#)?Percentage\b",
        semantic_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing "
            f"canonical cell::data_format::Percentage: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE}"
        )

    transaction_paths = _numbers_table_cell_percentage_format_transaction_sources(root)
    transaction_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in transaction_paths
    )
    transaction_code = _mask_rust_non_code(transaction_source)
    if re.search(
        r"\bpub\s+mod\s+transaction\b", semantic_code
    ) is None and re.search(
        r"\bpub\s+mod\s+transaction\b", transaction_code
    ) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"percentage transaction namespace: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        transaction_source,
        frozenset(NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers Percentage-format public API is missing "
                f"canonical transaction type {name}: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_SOURCE}"
            )

    data_format_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    data_format_code = _mask_rust_non_code(data_format_source)
    if PUBLIC_NUMBERS_CELL_DATA_FORMAT_PERCENTAGE_MODULE.search(data_format_code) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"cell::data_format::percentage module: {data_format_path.relative_to(root)}"
        )
    if "Percentage" not in _rust_canonical_exports(
        data_format_source, frozenset({"Percentage"})
    ):
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"cell::data_format::Percentage reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    if NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing the "
            f"private package owner module: {package_path.relative_to(root)}"
        )
    if PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE.search(package_code) is not None:
        violation = PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_PERCENTAGE_FORMAT_MODULE.search(package_code)
        line_number = package_code.count("\n", 0, violation.start()) + 1 if violation else 1
        violations.append(
            "focused litchi-numbers Percentage-format public API exposes its "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+(?:r#)?cell\b", _mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-numbers Percentage-format public API is missing root "
            f"cell module: {lib_path.relative_to(root)}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_helper_root = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    owner_code = _mask_rust_non_code(owner_related_source)
    owner_exports = _rust_canonical_exports(
        owner_related_source,
        frozenset(NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers Percentage-format public API is missing "
                f"canonical package type {name}: {owner_path.relative_to(root)}"
            )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PACKAGE_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Percentage-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method != "apply_table_cell_percentage_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Percentage-format Package method "
                        f"{method} must accept selector-first {selector}: {owner_path.relative_to(root)}:{line_number}"
                    )
        if method == "table_cell_percentage_format":
            normalized = re.sub(r"\s+", "", declaration)
            if "Option<Percentage>" not in normalized:
                violations.append(
                    "focused litchi-numbers Percentage-format read must return "
                    f"Option<Percentage>: {owner_path.relative_to(root)}:{line_number}"
                )
        elif method == "edit_table_cell_percentage_format":
            if not re.search(r"\bEdit\b", declaration):
                violations.append(
                    "focused litchi-numbers Percentage-format edit must return "
                    f"Edit: {owner_path.relative_to(root)}:{line_number}"
                )
        elif not re.search(r"\bPatch\b", declaration):
            violations.append(
                "focused litchi-numbers Percentage-format apply must accept "
                f"Patch: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "apply_table_cell_percentage_format" and not re.search(
            r"\bCommit\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Percentage-format apply must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )
        for pattern in (
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_PARAMETER,
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_RAW_ID_PARAMETER,
        ):
            for match in pattern.finditer(declaration):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                violations.append(
                    "focused litchi-numbers Percentage-format Package method exposes "
                    f"raw identifier/coordinate parameter {parameter}: {owner_path.relative_to(root)}:{line_number}"
                )

    edit_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Edit"
        )
    }
    for method in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EDIT_METHODS:
        record = edit_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Percentage-format edit is missing "
                f"Edit::{method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method == "set" and not re.search(r"\bPercentage\b", declaration):
            violations.append(
                "focused litchi-numbers Percentage-format Edit::set must accept "
                f"Percentage: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "commit" and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Percentage-format Edit::commit must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_related_source) is None:
            violations.append(
                "focused litchi-numbers Percentage-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    dedicated_sources = {owner_path, *semantic_paths, *transaction_paths}
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in (
            *NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[:3],
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_SEMANTIC_SOURCE,
        )
        if (root / path).is_file()
    }
    export_sources.update(semantic_paths)
    export_sources.update(transaction_paths)
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_percentage_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (
                ["pub", "use"],
                ["pub", "type"],
            )
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_percentage_format" in identifiers
                    or "percentage_format" in identifiers
                    or "percentage" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers Percentage-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Percentage-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    not dedicated_source
                    and identifier
                    in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TRANSACTION_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers Percentage-format public API retains "
                        f"root transaction alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_percentage_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Percentage-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for pattern, label in (
                (
                    RUST_BYTE_SLICE,
                    "raw byte slice",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_CONTAINER,
                    "opaque raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE,
                    "opaque raw byte container type",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_RAW_HELPER,
                    "raw helper",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw identifier/coordinate parameter",
                ),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    value_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        "focused litchi-numbers Percentage-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{value_line}"
                    )

    example_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Percentage-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EXAMPLE}"
        )
    else:
        example_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(example_path.read_text(encoding="utf-8"))
        )
        if re.search(
            r"\b(?:edit_table_cell_percentage_format|"
            r"apply_table_cell_percentage_format)\s*\(",
            example_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Percentage-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Percentage-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = integration_path.read_text(encoding="utf-8")
        integration_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_source)
        )
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Percentage-format integration test is missing "
                f"#[test] coverage: {integration_path}"
            )
        for label, marker in NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_source) is None:
                violations.append(
                    "focused litchi-numbers Percentage-format integration test is missing "
                    f"{label}: {integration_path}"
                )

    for fuzz_path, label in (
        (
            NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_FUZZ_SOURCE,
            "codec fuzz target",
        ),
        (NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Percentage-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Percentage-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_PERCENTAGE_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Percentage-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Percentage-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )
    return sorted(set(violations))


def audit_numbers_table_cell_currency_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Currency-format owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_currency_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_paths = _numbers_table_cell_currency_format_semantic_sources(root)
    semantic_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in semantic_paths
    )
    semantic_code = _mask_rust_non_code(semantic_source)
    if re.search(
        r"\bpub\s+struct\s+(?:r#)?Currency\b|"
        r"\bpub\s+use[^;]*\bCurrency\b|"
        r"\bdecimal_format!\s*\(\s*(?:r#)?Currency\b",
        semantic_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Currency-format public API is missing "
            f"canonical cell::data_format::Currency: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_SOURCE}"
        )

    transaction_paths = _numbers_table_cell_currency_format_transaction_sources(root)
    transaction_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in transaction_paths
    )
    transaction_code = _mask_rust_non_code(transaction_source)
    if re.search(r"\bpub\s+mod\s+transaction\b", semantic_code) is None and re.search(
        r"\bpub\s+mod\s+transaction\b", transaction_code
    ) is None:
        violations.append(
            "focused litchi-numbers Currency-format public API is missing the "
            f"currency transaction namespace: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        transaction_source,
        frozenset(NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers Currency-format public API is missing "
                f"canonical transaction type {name}: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_SOURCE}"
            )

    data_format_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    data_format_code = _mask_rust_non_code(data_format_source)
    if PUBLIC_NUMBERS_CELL_DATA_FORMAT_CURRENCY_MODULE.search(data_format_code) is None:
        violations.append(
            "focused litchi-numbers Currency-format public API is missing the "
            f"cell::data_format::currency module: {data_format_path.relative_to(root)}"
        )
    if "Currency" not in _rust_canonical_exports(
        data_format_source, frozenset({"Currency"})
    ):
        violations.append(
            "focused litchi-numbers Currency-format public API is missing the "
            f"cell::data_format::Currency reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    if NUMBERS_PACKAGE_TABLE_CELL_CURRENCY_FORMAT_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-numbers Currency-format public API is missing the "
            f"private package owner module: {package_path.relative_to(root)}"
        )
    public_package = PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_CURRENCY_FORMAT_MODULE.search(
        package_code
    )
    if public_package is not None:
        line_number = package_code.count("\n", 0, public_package.start()) + 1
        violations.append(
            "focused litchi-numbers Currency-format public API exposes its "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+(?:r#)?cell\b", _mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-numbers Currency-format public API is missing root "
            f"cell module: {lib_path.relative_to(root)}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_helper_root = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    owner_code = _mask_rust_non_code(owner_related_source)
    owner_exports = _rust_canonical_exports(
        owner_related_source,
        frozenset(NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers Currency-format public API is missing "
                f"canonical package type {name}: {owner_path.relative_to(root)}"
            )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PACKAGE_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Currency-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method != "apply_table_cell_currency_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Currency-format Package method "
                        f"{method} must accept selector-first {selector}: {owner_path.relative_to(root)}:{line_number}"
                    )
        if method == "table_cell_currency_format":
            normalized = re.sub(r"\s+", "", declaration)
            if "Option<Currency>" not in normalized:
                violations.append(
                    "focused litchi-numbers Currency-format read must return "
                    f"Option<Currency>: {owner_path.relative_to(root)}:{line_number}"
                )
        elif method == "edit_table_cell_currency_format":
            if not re.search(r"\bEdit\b", declaration):
                violations.append(
                    "focused litchi-numbers Currency-format edit must return "
                    f"Edit: {owner_path.relative_to(root)}:{line_number}"
                )
        elif not re.search(r"\bPatch\b", declaration):
            violations.append(
                "focused litchi-numbers Currency-format apply must accept "
                f"Patch: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "apply_table_cell_currency_format" and not re.search(
            r"\bCommit\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Currency-format apply must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )
        for pattern in (
            NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PUBLIC_RAW_PARAMETER,
            NUMBERS_TABLE_CELL_CURRENCY_FORMAT_RAW_ID_PARAMETER,
        ):
            for match in pattern.finditer(declaration):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                violations.append(
                    "focused litchi-numbers Currency-format Package method exposes "
                    f"raw identifier/coordinate parameter {parameter}: {owner_path.relative_to(root)}:{line_number}"
                )

    edit_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Edit"
        )
    }
    for method in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_EDIT_METHODS:
        record = edit_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Currency-format edit is missing "
                f"Edit::{method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method == "set" and not re.search(r"\bCurrency\b", declaration):
            violations.append(
                "focused litchi-numbers Currency-format Edit::set must accept "
                f"Currency: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "commit" and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Currency-format Edit::commit must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_related_source) is None:
            violations.append(
                "focused litchi-numbers Currency-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    dedicated_sources = set(semantic_paths) | set(transaction_paths) | {owner_path}
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[:3]
        if (root / path).is_file()
    }
    export_sources.update(semantic_paths)
    export_sources.update(transaction_paths)
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_currency_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (
                ["pub", "use"],
                ["pub", "type"],
            )
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_currency_format" in identifiers
                    or "currency_format" in identifiers
                    or "currency" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers Currency-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Currency-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    not dedicated_source
                    and identifier
                    in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TRANSACTION_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers Currency-format public API retains "
                        f"root transaction alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_currency_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Currency-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for pattern, label in (
                (RUST_BYTE_SLICE, "raw byte slice"),
                (
                    NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PUBLIC_RAW_CONTAINER,
                    "opaque raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE,
                    "opaque raw byte container type",
                ),
                (NUMBERS_TABLE_CELL_CURRENCY_FORMAT_RAW_HELPER, "raw helper"),
                (
                    NUMBERS_TABLE_CELL_CURRENCY_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_CURRENCY_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw identifier/coordinate parameter",
                ),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    value_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        "focused litchi-numbers Currency-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{value_line}"
                    )

    example_path = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Currency-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_CURRENCY_FORMAT_EXAMPLE}"
        )
    else:
        example_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(example_path.read_text(encoding="utf-8"))
        )
        if re.search(
            r"\b(?:edit_table_cell_currency_format|"
            r"apply_table_cell_currency_format)\s*\(",
            example_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Currency-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_CURRENCY_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Currency-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_CURRENCY_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = integration_path.read_text(encoding="utf-8")
        integration_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_source)
        )
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Currency-format integration test is missing "
                f"#[test] coverage: {integration_path}"
            )
        for label, marker in NUMBERS_TABLE_CELL_CURRENCY_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_source) is None:
                violations.append(
                    "focused litchi-numbers Currency-format integration test is missing "
                    f"{label}: {integration_path}"
                )

    for fuzz_path, label in (
        (
            NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_FUZZ_SOURCE,
            "codec fuzz target",
        ),
        (NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Currency-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Currency-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_CURRENCY_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_CURRENCY_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Currency-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Currency-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )
    return sorted(set(violations))


def audit_numbers_table_cell_fraction_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Fraction-format owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_fraction_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_paths = _numbers_table_cell_fraction_format_semantic_sources(root)
    semantic_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in semantic_paths
    )
    semantic_code = _mask_rust_non_code(semantic_source)
    if re.search(
        r"\bpub\s+struct\s+(?:r#)?Fraction\b|"
        r"\bpub\s+use[^;]*\bFraction\b|"
        r"\bdecimal_format!\s*\(\s*(?:r#)?Fraction\b",
        semantic_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing "
            f"canonical cell::data_format::Fraction: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_SOURCE}"
        )
    if re.search(
        r"\bpub\s+(?:struct|enum)\s+(?:r#)?FractionAccuracy\b|"
        r"\bpub\s+use[^;]*\bFractionAccuracy\b",
        semantic_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing "
            f"canonical cell::data_format::FractionAccuracy: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_SOURCE}"
        )

    transaction_paths = _numbers_table_cell_fraction_format_transaction_sources(root)
    transaction_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in transaction_paths
    )
    transaction_code = _mask_rust_non_code(transaction_source)
    if re.search(r"\bpub\s+mod\s+transaction\b", semantic_code) is None and re.search(
        r"\bpub\s+mod\s+transaction\b", transaction_code
    ) is None:
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing the "
            f"fraction transaction namespace: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        transaction_source,
        frozenset(NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers Fraction-format public API is missing "
                f"canonical transaction type {name}: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_SOURCE}"
            )

    data_format_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    data_format_code = _mask_rust_non_code(data_format_source)
    if PUBLIC_NUMBERS_CELL_DATA_FORMAT_FRACTION_MODULE.search(data_format_code) is None:
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing the "
            f"cell::data_format::fraction module: {data_format_path.relative_to(root)}"
        )
    if "Fraction" not in _rust_canonical_exports(
        data_format_source, frozenset({"Fraction"})
    ):
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing the "
            f"cell::data_format::Fraction reexport: {data_format_path.relative_to(root)}"
        )
    if "FractionAccuracy" not in _rust_canonical_exports(
        data_format_source, frozenset({"FractionAccuracy"})
    ):
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing the "
            f"cell::data_format::FractionAccuracy reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    if NUMBERS_PACKAGE_TABLE_CELL_FRACTION_FORMAT_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing the "
            f"private package owner module: {package_path.relative_to(root)}"
        )
    public_package = PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_FRACTION_FORMAT_MODULE.search(
        package_code
    )
    if public_package is not None:
        line_number = package_code.count("\n", 0, public_package.start()) + 1
        violations.append(
            "focused litchi-numbers Fraction-format public API exposes its "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+(?:r#)?cell\b", _mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-numbers Fraction-format public API is missing root "
            f"cell module: {lib_path.relative_to(root)}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_helper_root = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    owner_code = _mask_rust_non_code(owner_related_source)
    owner_exports = _rust_canonical_exports(
        owner_related_source,
        frozenset(NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers Fraction-format public API is missing "
                f"canonical package type {name}: {owner_path.relative_to(root)}"
            )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_FRACTION_FORMAT_PACKAGE_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Fraction-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method != "apply_table_cell_fraction_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Fraction-format Package method "
                        f"{method} must accept selector-first {selector}: {owner_path.relative_to(root)}:{line_number}"
                    )
        if method == "table_cell_fraction_format":
            normalized = re.sub(r"\s+", "", declaration)
            if "Option<Fraction>" not in normalized:
                violations.append(
                    "focused litchi-numbers Fraction-format read must return "
                    f"Option<Fraction>: {owner_path.relative_to(root)}:{line_number}"
                )
        elif method == "edit_table_cell_fraction_format":
            if not re.search(r"\bEdit\b", declaration):
                violations.append(
                    "focused litchi-numbers Fraction-format edit must return "
                    f"Edit: {owner_path.relative_to(root)}:{line_number}"
                )
        elif not re.search(r"\bPatch\b", declaration):
            violations.append(
                "focused litchi-numbers Fraction-format apply must accept "
                f"Patch: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "apply_table_cell_fraction_format" and not re.search(
            r"\bCommit\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Fraction-format apply must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )
        for pattern in (
            NUMBERS_TABLE_CELL_FRACTION_FORMAT_PUBLIC_RAW_PARAMETER,
            NUMBERS_TABLE_CELL_FRACTION_FORMAT_RAW_ID_PARAMETER,
        ):
            for match in pattern.finditer(declaration):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                violations.append(
                    "focused litchi-numbers Fraction-format Package method exposes "
                    f"raw identifier/coordinate parameter {parameter}: {owner_path.relative_to(root)}:{line_number}"
                )

    edit_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Edit"
        )
    }
    for method in NUMBERS_TABLE_CELL_FRACTION_FORMAT_EDIT_METHODS:
        record = edit_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Fraction-format edit is missing "
                f"Edit::{method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method == "set" and not re.search(r"\bFraction\b", declaration):
            violations.append(
                "focused litchi-numbers Fraction-format Edit::set must accept "
                f"Fraction: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "commit" and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Fraction-format Edit::commit must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_FRACTION_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_related_source) is None:
            violations.append(
                "focused litchi-numbers Fraction-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    dedicated_sources = set(semantic_paths) | set(transaction_paths) | {owner_path}
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[:3]
        if (root / path).is_file()
    }
    export_sources.update(semantic_paths)
    export_sources.update(transaction_paths)
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_fraction_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (
                ["pub", "use"],
                ["pub", "type"],
            )
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_fraction_format" in identifiers
                    or "fraction_format" in identifiers
                    or "fraction" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers Fraction-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_FRACTION_FORMAT_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_FRACTION_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Fraction-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    not dedicated_source
                    and identifier
                    in NUMBERS_TABLE_CELL_FRACTION_FORMAT_TRANSACTION_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers Fraction-format public API retains "
                        f"root transaction alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_fraction_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Fraction-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for pattern, label in (
                (RUST_BYTE_SLICE, "raw byte slice"),
                (
                    NUMBERS_TABLE_CELL_FRACTION_FORMAT_PUBLIC_RAW_CONTAINER,
                    "opaque raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_FRACTION_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE,
                    "opaque raw byte container type",
                ),
                (NUMBERS_TABLE_CELL_FRACTION_FORMAT_RAW_HELPER, "raw helper"),
                (
                    NUMBERS_TABLE_CELL_FRACTION_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_FRACTION_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw identifier/coordinate parameter",
                ),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    value_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        "focused litchi-numbers Fraction-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{value_line}"
                    )

    example_path = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Fraction-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_FRACTION_FORMAT_EXAMPLE}"
        )
    else:
        example_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(example_path.read_text(encoding="utf-8"))
        )
        if re.search(
            r"\b(?:edit_table_cell_fraction_format|"
            r"apply_table_cell_fraction_format)\s*\(",
            example_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Fraction-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_FRACTION_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Fraction-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_FRACTION_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = integration_path.read_text(encoding="utf-8")
        integration_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_source)
        )
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Fraction-format integration test is missing "
                f"#[test] coverage: {integration_path}"
            )
        for label, marker in NUMBERS_TABLE_CELL_FRACTION_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_source) is None:
                violations.append(
                    "focused litchi-numbers Fraction-format integration test is missing "
                    f"{label}: {integration_path}"
                )

    for fuzz_path, label in (
        (
            NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_FUZZ_SOURCE,
            "codec fuzz target",
        ),
        (NUMBERS_TABLE_CELL_FRACTION_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Fraction-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Fraction-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_FRACTION_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_FRACTION_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Fraction-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Fraction-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )
    return sorted(set(violations))


def audit_numbers_table_cell_scientific_format_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Scientific-format owner."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_scientific_format_owner_present(root):
        return []

    violations: list[str] = []
    semantic_paths = _numbers_table_cell_scientific_format_semantic_sources(root)
    semantic_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in semantic_paths
    )
    semantic_code = _mask_rust_non_code(semantic_source)
    if re.search(
        r"\bpub\s+struct\s+(?:r#)?Scientific\b|"
        r"\bpub\s+use[^;]*\bScientific\b|"
        r"\bdecimal_format!\s*\(\s*(?:r#)?Scientific\b",
        semantic_code,
    ) is None:
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing "
            f"canonical cell::data_format::Scientific: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_SOURCE}"
        )

    transaction_paths = _numbers_table_cell_scientific_format_transaction_sources(root)
    transaction_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in transaction_paths
    )
    transaction_code = _mask_rust_non_code(transaction_source)
    if re.search(r"\bpub\s+mod\s+transaction\b", semantic_code) is None and re.search(
        r"\bpub\s+mod\s+transaction\b", transaction_code
    ) is None:
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing the "
            f"scientific transaction namespace: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        transaction_source,
        frozenset(NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers Scientific-format public API is missing "
                f"canonical transaction type {name}: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_SOURCE}"
            )

    data_format_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        if data_format_path.is_file()
        else ""
    )
    data_format_code = _mask_rust_non_code(data_format_source)
    if PUBLIC_NUMBERS_CELL_DATA_FORMAT_SCIENTIFIC_MODULE.search(data_format_code) is None:
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing the "
            f"cell::data_format::scientific module: {data_format_path.relative_to(root)}"
        )
    if "Scientific" not in _rust_canonical_exports(
        data_format_source, frozenset({"Scientific"})
    ):
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing the "
            f"cell::data_format::Scientific reexport: {data_format_path.relative_to(root)}"
        )

    package_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    if NUMBERS_PACKAGE_TABLE_CELL_SCIENTIFIC_FORMAT_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing the "
            f"private package owner module: {package_path.relative_to(root)}"
        )
    public_package = PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_SCIENTIFIC_FORMAT_MODULE.search(
        package_code
    )
    if public_package is not None:
        line_number = package_code.count("\n", 0, public_package.start()) + 1
        violations.append(
            "focused litchi-numbers Scientific-format public API exposes its "
            f"package owner module: {package_path.relative_to(root)}:{line_number}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+(?:r#)?cell\b", _mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-numbers Scientific-format public API is missing root "
            f"cell module: {lib_path.relative_to(root)}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_helper_root = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_HELPER_ROOT
    owner_related_paths = {owner_path}
    if owner_helper_root.is_dir():
        owner_related_paths.update(owner_helper_root.rglob("*.rs"))
    owner_related_source = "\n".join(
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    owner_code = _mask_rust_non_code(owner_related_source)
    owner_exports = _rust_canonical_exports(
        owner_related_source,
        frozenset(NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES),
    )
    for name in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers Scientific-format public API is missing "
                f"canonical package type {name}: {owner_path.relative_to(root)}"
            )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PACKAGE_METHODS:
        record = owner_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Scientific-format public API is missing "
                f"Package method {method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method != "apply_table_cell_scientific_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-numbers Scientific-format Package method "
                        f"{method} must accept selector-first {selector}: {owner_path.relative_to(root)}:{line_number}"
                    )
        if method == "table_cell_scientific_format":
            normalized = re.sub(r"\s+", "", declaration)
            if "Option<Scientific>" not in normalized:
                violations.append(
                    "focused litchi-numbers Scientific-format read must return "
                    f"Option<Scientific>: {owner_path.relative_to(root)}:{line_number}"
                )
        elif method == "edit_table_cell_scientific_format":
            if not re.search(r"\bEdit\b", declaration):
                violations.append(
                    "focused litchi-numbers Scientific-format edit must return "
                    f"Edit: {owner_path.relative_to(root)}:{line_number}"
                )
        elif not re.search(r"\bPatch\b", declaration):
            violations.append(
                "focused litchi-numbers Scientific-format apply must accept "
                f"Patch: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "apply_table_cell_scientific_format" and not re.search(
            r"\bCommit\b", declaration
        ):
            violations.append(
                "focused litchi-numbers Scientific-format apply must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )
        for pattern in (
            NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PUBLIC_RAW_PARAMETER,
            NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_RAW_ID_PARAMETER,
        ):
            for match in pattern.finditer(declaration):
                parameter = re.sub(r"\s+", " ", match.group(0)).strip()
                violations.append(
                    "focused litchi-numbers Scientific-format Package method exposes "
                    f"raw identifier/coordinate parameter {parameter}: {owner_path.relative_to(root)}:{line_number}"
                )

    edit_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_related_source, "Edit"
        )
    }
    for method in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_EDIT_METHODS:
        record = edit_methods.get(method)
        if record is None:
            violations.append(
                "focused litchi-numbers Scientific-format edit is missing "
                f"Edit::{method}: {owner_path.relative_to(root)}"
            )
            continue
        declaration, line_number = record
        if method == "set" and not re.search(r"\bScientific\b", declaration):
            violations.append(
                "focused litchi-numbers Scientific-format Edit::set must accept "
                f"Scientific: {owner_path.relative_to(root)}:{line_number}"
            )
        if method == "commit" and not re.search(r"\bCommit\b", declaration):
            violations.append(
                "focused litchi-numbers Scientific-format Edit::commit must return "
                f"Commit: {owner_path.relative_to(root)}:{line_number}"
            )

    for label, marker in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_related_source) is None:
            violations.append(
                "focused litchi-numbers Scientific-format owner is missing "
                f"{label} transaction marker: {owner_path.relative_to(root)}"
            )

    dedicated_sources = set(semantic_paths) | set(transaction_paths) | {owner_path}
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_NUMBER_FORMAT_EXPORT_SOURCES[:3]
        if (root / path).is_file()
    }
    export_sources.update(semantic_paths)
    export_sources.update(transaction_paths)
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_scientific_format_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (
                ["pub", "use"],
                ["pub", "type"],
            )
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_scientific_format" in identifiers
                    or "scientific_format" in identifiers
                    or "scientific" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers Scientific-format public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    identifier in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FLAT_ALIASES
                    or NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FLAT_ALIAS_PATTERN.fullmatch(
                        identifier
                    )
                ):
                    violations.append(
                        "focused litchi-numbers Scientific-format public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    not dedicated_source
                    and identifier
                    in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TRANSACTION_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers Scientific-format public API retains "
                        f"root transaction alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_scientific_format_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers Scientific-format public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for pattern, label in (
                (RUST_BYTE_SLICE, "raw byte slice"),
                (
                    NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PUBLIC_RAW_CONTAINER,
                    "opaque raw byte container",
                ),
                (
                    NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PUBLIC_RAW_BYTE_CONTAINER_TYPE,
                    "opaque raw byte container type",
                ),
                (NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_RAW_HELPER, "raw helper"),
                (
                    NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_TYPED_RAW_ID_PARAMETER,
                    "typed raw identifier parameter",
                ),
                (
                    NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_PUBLIC_RAW_PARAMETER,
                    "raw identifier/coordinate parameter",
                ),
            ):
                for match in pattern.finditer(declaration):
                    value = re.sub(r"\s+", " ", match.group(0)).strip()
                    value_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        "focused litchi-numbers Scientific-format public API exposes "
                        f"{label} {value}: {path.relative_to(root)}:{value_line}"
                    )

    example_path = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_EXAMPLE
    if not example_path.is_file():
        violations.append(
            "focused litchi-numbers Scientific-format boundary is missing example: "
            f"{NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_EXAMPLE}"
        )
    else:
        example_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(example_path.read_text(encoding="utf-8"))
        )
        if re.search(
            r"\b(?:edit_table_cell_scientific_format|"
            r"apply_table_cell_scientific_format)\s*\(",
            example_code,
        ) is None:
            violations.append(
                "focused litchi-numbers Scientific-format example does not exercise "
                f"the selector-first owner: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_EXAMPLE}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_INTEGRATION_SOURCE
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers Scientific-format boundary is missing integration "
            f"test: {NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_INTEGRATION_SOURCE}"
        )
    else:
        integration_source = integration_path.read_text(encoding="utf-8")
        integration_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_source)
        )
        if re.search(r"#\s*\[\s*test\s*\]", integration_code) is None:
            violations.append(
                "focused litchi-numbers Scientific-format integration test is missing "
                f"#[test] coverage: {integration_path}"
            )
        for label, marker in NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_INTEGRATION_MARKERS.items():
            if marker.search(integration_source) is None:
                violations.append(
                    "focused litchi-numbers Scientific-format integration test is missing "
                    f"{label}: {integration_path}"
                )

    for fuzz_path, label in (
        (
            NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_FUZZ_SOURCE,
            "codec fuzz target",
        ),
        (NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FUZZ_SOURCE, "package fuzz target"),
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers Scientific-format boundary is missing "
                f"{label}: {fuzz_path}"
            )
        elif re.search(
            r"\bfuzz_target!\s*\(",
            _mask_rust_non_code(absolute.read_text(encoding="utf-8")),
        ) is None:
            violations.append(
                "focused litchi-numbers Scientific-format fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_SCIENTIFIC_FORMAT_FUZZ_CORPUS,
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers Scientific-format boundary is missing fuzz "
                f"corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.iterdir()):
            violations.append(
                "focused litchi-numbers Scientific-format fuzz corpus must be nonempty: "
                f"{corpus}"
            )
    return sorted(set(violations))


def _numbers_table_cell_control_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden in the generic control facade."""

    if identifier == "litchi_iwa_common":
        # The neutral table vocabulary owner.
        return None
    if identifier in NUMBERS_TABLE_CELL_CONTROL_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in NUMBERS_TABLE_CELL_CONTROL_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in NUMBERS_TABLE_CELL_CONTROL_WIRE_TYPES:
        return "wire type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_comment_reply_public_leak(identifier: str) -> str | None:
    """Classify native vocabulary forbidden in the reply read facade."""

    if identifier in NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_ALLOWED_IDENTIFIERS:
        return None
    if identifier in NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_NATIVE_IDENTIFIERS:
        return "forbidden native comment metadata vocabulary"
    if identifier in NUMBERS_TABLE_CELL_COMMENT_REPLY_PUBLIC_LEAKS:
        return "forbidden native reply vocabulary"
    if identifier.lower() in {
        "bytes",
        "wire",
        "iwa",
        "prost",
        "buffa",
        "rawmessage",
        "storageid",
        "authorid",
        "uuid",
    }:
        return "forbidden native reply vocabulary"
    if any(
        identifier.startswith(prefix)
        for prefix in NUMBERS_TABLE_CELL_COMMENT_REPLY_PUBLIC_LEAK_PREFIXES
    ):
        return "forbidden native reply vocabulary"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_comment_reply_mutation_public_leak(
    identifier: str,
) -> str | None:
    """Classify native vocabulary forbidden in the reply mutation facade."""

    if identifier in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_PUBLIC_LEAKS:
        return "forbidden native reply vocabulary"
    if identifier in {
        "bytes",
        "wire",
        "iwa",
        "prost",
        "buffa",
        "raw_bytes",
        "raw_message",
        "storage_id",
        "reply_storage_id",
        "object_id",
        "object_identifier",
        "usize",
    }:
        return "forbidden native reply vocabulary"
    if any(
        identifier.startswith(prefix)
        for prefix in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_PUBLIC_LEAK_PREFIXES
    ):
        return "forbidden native reply vocabulary"
    return _iwork_public_leak(identifier)


def _numbers_table_cell_control_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return NUMBERS_TABLE_CELL_CONTROL_OWNER_PATH.search(declaration) is not None or any(
        identifier in NUMBERS_TABLE_CELL_CONTROL_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_numbers_table_cell_control_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers & NUMBERS_TABLE_CELL_CONTROL_FLAT_ALIASES
    ) or _numbers_table_cell_control_owner_declaration(declaration)


def _numbers_table_cell_control_owner_present(root: Path) -> bool:
    owner_path = root / NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE
    codec_path = root / NUMBERS_TABLE_CELL_CONTROL_CODEC_SOURCE
    package_path = root / NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and codec_path.is_file()
        and NUMBERS_PACKAGE_TABLE_CELL_CONTROL_MODULE.search(package_source)
        is not None
    )


def audit_numbers_table_cell_control_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the archive-free generic Numbers ``CellControl`` facade."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_control_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_TYPES)
    )
    for name in NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-numbers cell-control public API is missing "
                f"canonical cell::data_format::control type {name}: "
                f"{NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE}"
            )
    semantic_code = _mask_rust_non_code(semantic_source)
    enum_match = re.search(r"\bpub\s+enum\s+CellControl\b", semantic_code)
    if enum_match is not None:
        opening = semantic_code.find("{", enum_match.end())
        closing = semantic_code.find("}", opening + 1) if opening >= 0 else -1
        enum_body = semantic_code[opening + 1 : closing] if closing >= 0 else ""
        for variant in NUMBERS_TABLE_CELL_CONTROL_VARIANTS:
            if re.search(rf"\b{re.escape(variant)}\b", enum_body) is None:
                violations.append(
                    "focused litchi-numbers cell-control public API is missing "
                    f"CellControl::{variant} variant: {NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE}"
                )
    if PUBLIC_NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_MODULE.search(
        semantic_code
    ) is None:
        violations.append(
            "focused litchi-numbers cell-control public API is missing canonical "
            "cell::data_format::control::transaction module: "
            f"{NUMBERS_TABLE_CELL_CONTROL_SEMANTIC_SOURCE}"
        )

    data_format_path = root / NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[2]
    data_format_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(data_format_path.read_text(encoding="utf-8"))
        )
        if data_format_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_CELL_CONTROL_MODULE.search(data_format_source) is None:
        violations.append(
            "focused litchi-numbers cell-control public API is missing canonical "
            f"cell::data_format::control module: {NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[2]}"
        )

    package_path = root / NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    if NUMBERS_PACKAGE_TABLE_CELL_CONTROL_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-numbers cell-control public API is missing private "
            f"package owner module: {NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_CONTROL_MODULE.finditer(
        package_source
    ):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-numbers cell-control public API exposes public "
            f"package::table_cell_control module: {NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[1]}:{line_number}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_related_paths = {
        owner_path,
        *owner_path.parent.glob("table_cell_control*.rs"),
    }
    owner_code = "\n".join(
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for path in sorted(owner_related_paths)
        if path.is_file()
    )
    # The unified owner delegates the legacy Pop-Up Menu graph and its strict
    # metadata selector/edge proof through these package siblings. Include
    # both in the bounded split-graph contract.
    owner_graph_paths = set(owner_related_paths)
    for popup_graph_name in (
        "table_cell_pop_up_menu.rs",
        "table_cell_pop_up_menu_metadata.rs",
    ):
        popup_graph_path = owner_path.parent / popup_graph_name
        if popup_graph_path.is_file():
            owner_graph_paths.add(popup_graph_path)
    owner_graph_code = "\n".join(
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for path in sorted(owner_graph_paths)
        if path.is_file()
    ) + "\n" + _iwa_table_cell_format_engine_code(root)
    owner_exports = _rust_canonical_exports(
        owner_source, frozenset(NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_TYPES)
    )
    for name in NUMBERS_TABLE_CELL_CONTROL_TRANSACTION_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-numbers cell-control public API is missing "
                f"canonical package type {name}: {NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}"
            )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in NUMBERS_TABLE_CELL_CONTROL_PACKAGE_METHODS:
        declaration = owner_methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-numbers cell-control public API is missing "
                f"Package method {method}: {NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}"
            )
            continue
        if method != "apply_table_cell_control_format":
            for selector in ("SheetSelector", "TableSelector", "At"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-numbers cell-control Package method "
                        f"{method} must accept selector-first {selector}: "
                        f"{NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}"
                    )
        if method == "table_cell_control_format" and not re.search(
            r"\bCellControl\b", declaration
        ):
            violations.append(
                "focused litchi-numbers cell-control Package method "
                "table_cell_control_format must expose CellControl: "
                f"{NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}"
            )

    # A second popup graph writer would bypass the generic owner.  Compatibility
    # wrappers may mention the popup spelling, but lifecycle/native writers in
    # this private owner are forbidden once the generic gate is active.
    for pattern in (
        r"\b(?:rewrite|prepare|remove|create|cull)_popup(?:_menu)?\b",
        r"\brewrite_native_popup_menu\b",
        r"\bexisting_popup_model_identifiers\b",
    ):
        match = re.search(pattern, owner_code)
        if match is not None:
            line_number = owner_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers cell-control owner retains duplicate popup "
                f"lifecycle implementation {match.group(0)}: "
                f"{NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}:{line_number}"
            )

    owner_marker_code = owner_code + "\n" + _iwa_table_cell_format_engine_code(root)
    for label, marker in NUMBERS_TABLE_CELL_CONTROL_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_marker_code) is None:
            violations.append(
                "focused litchi-numbers cell-control owner is missing "
                f"{label} transaction marker: {NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}"
            )

    # Keep the split-graph refusal contract as a migration bridge only.  Once
    # the popup split lifecycle contract is complete, requiring the old refusal
    # marker would make a successful multi-member owner fail the generic facade
    # audit.  The dedicated popup lifecycle audit below owns the new positive
    # contract and rejects stale refusal guards itself.
    popup_split_owner_complete = (
        all(
            marker.search(owner_graph_code) is not None
            for marker in NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_OWNER_REQUIRED_MARKERS.values()
        )
        and all(
            marker.search(owner_graph_code) is None
            for marker in NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_FORBIDDEN_MARKERS.values()
        )
    )
    if not popup_split_owner_complete:
        for label, marker in NUMBERS_TABLE_CELL_CONTROL_BOUNDED_OWNER_REQUIRED_MARKERS.items():
            if marker.search(owner_graph_code) is None:
                violations.append(
                    "focused litchi-numbers cell-control bounded cross-component owner is missing "
                    f"{label} marker: {NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}"
                )

    split_test_path = root / NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE
    split_test_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(split_test_path.read_text(encoding="utf-8"))
        )
        if split_test_path.is_file()
        else ""
    )
    if not popup_split_owner_complete:
        for label, marker in NUMBERS_TABLE_CELL_CONTROL_SPLIT_INTEGRATION_MARKERS.items():
            if marker.search(split_test_source) is None:
                violations.append(
                    "focused litchi-numbers cell-control integration is missing "
                    f"{label}: {NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE}"
                )

    codec_path = root / NUMBERS_TABLE_CELL_CONTROL_CODEC_SOURCE
    codec_source = (
        _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        if codec_path.is_file()
        else ""
    )
    codec_code = _mask_rust_non_code(codec_source)
    if not codec_path.is_file():
        violations.append(
            "focused litchi-numbers cell-control public API is missing strict hidden "
            f"codec source: {NUMBERS_TABLE_CELL_CONTROL_CODEC_SOURCE}"
        )
    else:
        for api in NUMBERS_TABLE_CELL_CONTROL_CODEC_REQUIRED_APIS:
            declaration = re.search(
                rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
                codec_code,
            )
            reexport = re.search(
                rf"\bpub\s+use\b[^;]*\b(?:as\s+)?{re.escape(api)}\b",
                codec_code,
            )
            if declaration is None and reexport is None:
                violations.append(
                    "focused litchi-numbers cell-control hidden codec is missing "
                    f"strict API {api}: {NUMBERS_TABLE_CELL_CONTROL_CODEC_SOURCE}"
                )
        for marker in NUMBERS_TABLE_CELL_CONTROL_CODEC_REQUIRED_MARKERS:
            if re.search(re.escape(marker), codec_source) is None:
                violations.append(
                    "focused litchi-numbers cell-control hidden codec is missing "
                    f"strict {marker} marker: {NUMBERS_TABLE_CELL_CONTROL_CODEC_SOURCE}"
                )

    codec_lib_path = root / NUMBERS_TABLE_CELL_CONTROL_CODEC_PUBLIC_SOURCE
    codec_lib_source = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(NUMBERS_TABLE_CELL_CONTROL_CODEC_MODULE)}\b",
        codec_lib_source,
    ) is None:
        violations.append(
            "focused litchi-numbers cell-control public API is missing hidden codec "
            f"module {NUMBERS_TABLE_CELL_CONTROL_CODEC_MODULE}: "
            f"{NUMBERS_TABLE_CELL_CONTROL_CODEC_PUBLIC_SOURCE}"
        )

    lib_path = root / NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        )
        if lib_path.is_file()
        else ""
    )
    if re.search(r"\bpub\s+mod\s+cell\b", lib_source) is None:
        violations.append(
            "focused litchi-numbers cell-control public API is missing root cell module: "
            f"{NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES[0]}"
        )

    for fuzz_path in (
        NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_SOURCE,
        NUMBERS_TABLE_CELL_CONTROL_FUZZ_SOURCE,
    ):
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers cell-control boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-numbers cell-control fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in (
        NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_CORPUS,
        NUMBERS_TABLE_CELL_CONTROL_FUZZ_CORPUS,
    ):
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-numbers cell-control boundary is missing fuzz corpus: "
                f"{corpus}"
            )

    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_CONTROL_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_siblings = (root / NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE).parent.glob(
        "table_cell_control*.rs"
    )
    dedicated_sources.update(path for path in owner_siblings if path.is_file())
    owner_helper_root = root / NUMBERS_TABLE_CELL_CONTROL_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELL_CONTROL_EXPORT_SOURCES
        if (root / path).is_file()
    }
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            if not _is_numbers_table_cell_control_public_declaration(
                declaration, dedicated_source=dedicated_source
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
            if (
                public_use_or_type
                and "*" in declaration
                and (
                    "table_cell_control" in identifiers
                    or "control" in identifiers
                )
            ):
                violations.append(
                    "focused litchi-numbers cell-control public API retains root "
                    f"aliases via glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in NUMBERS_TABLE_CELL_CONTROL_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-numbers cell-control public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cell_control_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers cell-control public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for raw_parameter in NUMBERS_TABLE_CELL_CONTROL_RAW_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-numbers cell-control public API exposes raw "
                    f"identifier parameter {raw_parameter.group(0).strip()}: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers cell-control public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_numbers_table_cell_control_popup_split_lifecycle_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the complete split Pop-Up Menu lifecycle.

    The public surface remains the selector-first ``CellControl`` facade.  A
    popup graph is nevertheless allowed to span several current component
    members, so the boundary must inspect the private popup native/metadata
    siblings as one ownership unit.  This ratchet intentionally checks
    positive lifecycle evidence (member edits, metadata transitions, cull and
    candidate locality) and rejects the temporary split-graph refusal guard.  It
    does not scan the shared adapter or same-named Pages/Keynote code.
    """

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir() or not _numbers_table_cell_control_owner_present(root):
        return []

    owner_path = root / NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE
    graph_paths: set[Path] = {
        owner_path,
        root / NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE,
        *(root / relative for relative in IWA_TABLE_CELL_FORMAT_ENGINE_SOURCES),
    }
    graph_paths.update(owner_path.parent.glob("table_cell_control*.rs"))
    graph_paths.update(owner_path.parent.glob("table_cell_pop_up_menu*.rs"))
    graph_source = "\n".join(
        _mask_rust_non_code(_mask_rust_cfg_test_items(path.read_text(encoding="utf-8")))
        for path in sorted(graph_paths)
        if path.is_file()
    )
    violations: list[str] = []

    for label, marker in NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(graph_source) is None:
            violations.append(
                "focused litchi-numbers cell-control Pop-Up Menu split owner is missing "
                f"{label} marker: {NUMBERS_TABLE_CELL_CONTROL_OWNER_SOURCE}"
            )
    for label, marker in NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_FORBIDDEN_MARKERS.items():
        match = marker.search(graph_source)
        if match is not None:
            line_number = graph_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers cell-control Pop-Up Menu split owner retains "
                f"{label} marker at {NUMBERS_TABLE_CELL_POP_UP_MENU_OWNER_SOURCE}:"
                f"{line_number}"
            )

    integration_path = root / NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE
    integration_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(integration_path.read_text(encoding="utf-8"))
        )
        if integration_path.is_file()
        else ""
    )
    if not integration_path.is_file():
        violations.append(
            "focused litchi-numbers cell-control Pop-Up Menu split integration source is missing: "
            f"{NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE}"
        )
    else:
        for label, marker in NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_INTEGRATION_MARKERS.items():
            if marker.search(integration_source) is None:
                violations.append(
                    "focused litchi-numbers cell-control Pop-Up Menu split integration is missing "
                    f"{label}: {NUMBERS_TABLE_CELL_CONTROL_SPLIT_COMPONENT_SOURCE}"
                )

    fuzz_contracts = (
        (
            NUMBERS_TABLE_CELL_CONTROL_FUZZ_SOURCE,
            "package",
            "package fuzz split lifecycle",
        ),
        (
            NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_SOURCE,
            "codec",
            "codec fuzz prepared replay",
        ),
    )
    for fuzz_path, kind, marker_name in fuzz_contracts:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-numbers cell-control Pop-Up Menu split boundary is missing "
                f"{kind} fuzz target: {fuzz_path}"
            )
            continue
        fuzz_source = _mask_rust_non_code(absolute.read_text(encoding="utf-8"))
        if "fuzz_target!" not in fuzz_source:
            violations.append(
                "focused litchi-numbers cell-control Pop-Up Menu split "
                f"{kind} fuzz target is missing fuzz_target! harness: {fuzz_path}"
            )
        marker = NUMBERS_TABLE_CELL_CONTROL_POPUP_SPLIT_FUZZ_MARKERS[marker_name]
        if marker.search(fuzz_source) is None:
            violations.append(
                "focused litchi-numbers cell-control Pop-Up Menu split "
                f"{kind} fuzz target is missing {marker_name}: {fuzz_path}"
            )

    for corpus, kind in (
        (NUMBERS_TABLE_CELL_CONTROL_FUZZ_CORPUS, "package"),
        (NUMBERS_TABLE_CELL_CONTROL_CODEC_FUZZ_CORPUS, "codec"),
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused litchi-numbers cell-control Pop-Up Menu split boundary is missing "
                f"{kind} fuzz corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.rglob("*")):
            violations.append(
                "focused litchi-numbers cell-control Pop-Up Menu split "
                f"{kind} fuzz corpus is empty: {corpus}"
            )

    return sorted(set(violations))


def audit_numbers_table_dimension_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Numbers sizing transaction API."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_DIMENSION_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / NUMBERS_TABLE_DIMENSION_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_DIMENSION_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES)
    )
    for name in NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-numbers table-dimension public API is missing "
                f"canonical table::dimension type {name}: "
                f"{NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE}"
            )
    masked_semantic = _mask_rust_non_code(semantic_source)
    if PUBLIC_NUMBERS_TABLE_DIMENSION_TRANSACTION_MODULE.search(masked_semantic) is None:
        violations.append(
            "focused litchi-numbers table-dimension public API is missing "
            "canonical table::dimension::transaction module: "
            f"{NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE}"
        )

    transaction_path = root / NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE
    transaction_source = (
        transaction_path.read_text(encoding="utf-8")
        if transaction_path.is_file()
        else ""
    )
    transaction_exports = _rust_canonical_exports(
        transaction_source, NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPE_SET
    )
    for name in NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-numbers table-dimension public API is missing "
                f"canonical table::dimension::transaction type {name}: "
                f"{NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE}"
            )

    lib_path = root / NUMBERS_SOURCE_ROOT / "lib.rs"
    lib_source = (
        lib_path.read_text(encoding="utf-8") if lib_path.is_file() else ""
    )
    root_exports = _rust_canonical_exports(
        lib_source, frozenset(NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES)
    )
    for name in NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES:
        if name not in root_exports:
            violations.append(
                "focused litchi-numbers table-dimension public API is missing "
                f"canonical root semantic type {name}: {NUMBERS_SOURCE_ROOT / 'lib.rs'}"
            )

    table_path = root / NUMBERS_SOURCE_ROOT / "table.rs"
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_DIMENSION_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-numbers table-dimension public API is missing "
            f"canonical table::dimension module: {NUMBERS_SOURCE_ROOT / 'table.rs'}"
        )

    package_export = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        if NUMBERS_PACKAGE_TABLE_DIMENSION_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-numbers table-dimension public API is missing "
                f"private package owner module: {package_export.relative_to(root)}"
            )
        for match in PUBLIC_NUMBERS_PACKAGE_TABLE_DIMENSION_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers table-dimension public API exposes "
                "duplicate package::table_dimension module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-numbers table-dimension public API is missing "
            f"private package owner module: {NUMBERS_SOURCE_ROOT / 'package.rs'}"
        )

    owner_path = root / NUMBERS_TABLE_DIMENSION_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-numbers table-dimension public API is missing "
            f"private package owner source: {NUMBERS_TABLE_DIMENSION_OWNER_SOURCE}"
        )
    else:
        owner_sources = {owner_path}
        if owner_helper_root.is_dir():
            owner_sources.update(owner_helper_root.rglob("*.rs"))
        owner_methods: set[str] = set()
        for path in owner_sources:
            for declaration, _line_number in _rust_public_declarations(
                path.read_text(encoding="utf-8")
            ):
                identifiers = [
                    match.group(1)
                    for match in RUST_IDENTIFIER.finditer(declaration)
                ]
                if identifiers[:2] == ["pub", "fn"] and len(identifiers) > 2:
                    owner_methods.add(identifiers[2])
        for method in NUMBERS_TABLE_DIMENSION_PACKAGE_METHODS:
            if method not in owner_methods:
                violations.append(
                    "focused litchi-numbers table-dimension public API is missing "
                    f"Package method {method}: {NUMBERS_TABLE_DIMENSION_OWNER_SOURCE}"
                )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for declaration, line_number, public_declaration, complete_source_scope in declarations:
            if not _is_numbers_table_dimension_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
            canonical_semantic_reexport = (
                public_declaration
                and identifiers[:2] == ["pub", "use"]
                and path == lib_path
                and "*" not in declaration
                and not re.search(r"\bas\b", declaration)
                and bool(set(identifiers) & set(NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES))
                and not bool(set(identifiers) & NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPE_SET)
            )
            canonical_common_reexport = (
                path == semantic_path
                and _is_numbers_table_dimension_common_reexport(declaration)
            )
            canonical_semantic_reexport = (
                canonical_semantic_reexport or canonical_common_reexport
            )
            canonical_module = (
                public_declaration
                and identifiers[:3] in (["pub", "mod", "dimension"], ["pub", "mod", "transaction"])
            )
            owner_declaration = _numbers_table_dimension_owner_declaration(declaration)
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-numbers table-dimension public API retains "
                    "flat aliases via dimension glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and public_use_or_type
                and not canonical_semantic_reexport
                and not canonical_module
            ):
                violations.append(
                    "focused litchi-numbers table-dimension public API exposes "
                    "public dimension owner alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if canonical_common_reexport and identifier == "litchi_iwa_common":
                    continue
                if public_declaration and identifier in NUMBERS_TABLE_DIMENSION_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-numbers table-dimension public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and public_use_or_type
                    and identifier in NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPE_SET
                    # The shared dimension value error keeps its name beside
                    # the transaction family's own `Error`.
                    and not (canonical_common_reexport and identifier == "Error")
                ):
                    violations.append(
                        "focused litchi-numbers table-dimension public API retains "
                        f"transaction alias outside table::dimension::transaction {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_dimension_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-numbers table-dimension public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count("\n", 0, match.start())
                violations.append(
                    "focused litchi-numbers table-dimension public API exposes "
                    f"raw byte slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_numbers_formula_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Enforce contextual, archive-free Numbers formula authoring vocabulary.

    Formula construction belongs at ``litchi_numbers::formula`` and is consumed
    by the existing selector-first table-cell transaction.  In particular, the
    public focused boundary must never route callers through the old shared-IWA
    formula facade or expose its UUID/table-ID based reference vocabulary.
    """

    semantic_path = root / NUMBERS_FORMULA_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    violations: list[str] = []

    canonical_exports = _rust_canonical_exports(
        semantic_source, NUMBERS_FORMULA_CANONICAL_TYPE_SET
    )
    for name in NUMBERS_FORMULA_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-numbers formula API is missing canonical "
            f"formula type {name}: {NUMBERS_FORMULA_SEMANTIC_SOURCE}"
        )

    transaction_limit_path = root / NUMBERS_FORMULA_TRANSACTION_LIMIT_SOURCE
    transaction_limit_source = (
        transaction_limit_path.read_text(encoding="utf-8")
        if transaction_limit_path.is_file()
        else ""
    )
    limit_variants = _rust_public_enum_variants(transaction_limit_source, "LimitKind")
    for variant in NUMBERS_FORMULA_TRANSACTION_LIMIT_VARIANTS:
        if variant in limit_variants:
            continue
        violations.append(
            "focused litchi-numbers formula transaction API is missing typed "
            f"LimitKind::{variant}: {NUMBERS_FORMULA_TRANSACTION_LIMIT_SOURCE}"
        )

    lib_path = root / NUMBERS_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_FORMULA_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-numbers formula API is missing canonical root "
            f"formula module: {NUMBERS_SOURCE_ROOT / 'lib.rs'}"
        )

    for relative_path in NUMBERS_FORMULA_PUBLIC_API_SOURCES:
        path = root / relative_path
        if not path.is_file():
            continue
        source = path.read_text(encoding="utf-8")
        for declaration, line_number in _rust_public_declarations(source):
            if not _is_numbers_formula_public_declaration(
                declaration,
                dedicated_source=relative_path == NUMBERS_FORMULA_SEMANTIC_SOURCE,
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            owner_declaration = _numbers_formula_owner_declaration(declaration)
            public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
            if (
                relative_path != NUMBERS_FORMULA_SEMANTIC_SOURCE
                and owner_declaration
                and public_use_or_type
                and bool(set(identifiers) & NUMBERS_FORMULA_CANONICAL_TYPE_SET)
            ):
                violations.append(
                    "focused litchi-numbers formula API retains root formula alias: "
                    f"{relative_path}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                reason = _numbers_formula_public_leak(identifier)
                if reason is None:
                    continue
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                violations.append(
                    "focused litchi-numbers formula API exposes "
                    f"{reason} {identifier}: {relative_path}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count("\n", 0, match.start())
                violations.append(
                    "focused litchi-numbers formula API exposes raw byte slice "
                    f"{byte_slice}: {relative_path}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_numbers_table_cells_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Numbers table-cell read API."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_CELLS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELLS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, NUMBERS_TABLE_CELLS_SHORT_NAMES
    )
    for name in NUMBERS_TABLE_CELLS_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-numbers table-cells read API is missing "
            f"canonical table::cells type {name}: "
            f"{NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE}"
        )

    lib_path = root / NUMBERS_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-numbers table-cells read API is missing "
            "canonical root table module: "
            f"{NUMBERS_SOURCE_ROOT / 'lib.rs'}"
        )

    table_path = root / NUMBERS_SOURCE_ROOT / "table.rs"
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_NUMBERS_TABLE_CELLS_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-numbers table-cells read API is missing "
            "canonical table::cells module: "
            f"{NUMBERS_SOURCE_ROOT / 'table.rs'}"
        )

    package_export = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        if NUMBERS_PACKAGE_TABLE_CELLS_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-numbers table-cells read API is missing "
                "private package owner module: "
                f"{package_export.relative_to(root)}"
            )
        for match in PUBLIC_NUMBERS_PACKAGE_TABLE_CELLS_MODULE.finditer(package_source):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers table-cells read API exposes "
                "duplicate package::table_cells module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-numbers table-cells read API is missing "
            "private package owner module: "
            f"{NUMBERS_SOURCE_ROOT / 'package.rs'}"
        )

    owner_path = root / NUMBERS_TABLE_CELLS_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-numbers table-cells read API is missing "
            "private package owner source: "
            f"{NUMBERS_TABLE_CELLS_OWNER_SOURCE}"
        )

    package_methods: set[str] = set()
    if owner_path.is_file():
        for declaration, _line_number in _rust_public_declarations(
            owner_path.read_text(encoding="utf-8")
        ):
            function = RUST_FUNCTION_DECLARATION.search(declaration)
            if function is not None:
                package_methods.add(function.group(1))
    for name in NUMBERS_TABLE_CELLS_PACKAGE_METHODS:
        if name in package_methods:
            continue
        violations.append(
            "focused litchi-numbers table-cells read API is missing "
            f"canonical Package::{name} method: "
            f"{NUMBERS_TABLE_CELLS_OWNER_SOURCE}"
        )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for (
            declaration,
            line_number,
            public_declaration,
            complete_source_scope,
        ) in declarations:
            if not _is_numbers_table_cells_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _numbers_table_cells_owner_declaration(declaration)
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = declaration_identifiers[:2] in (
                ["pub", "type"],
                ["pub", "use"],
            )
            flat_alias_exports = _rust_canonical_exports(
                declaration, NUMBERS_TABLE_CELLS_FLAT_ALIASES
            )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and public_use_or_type
            ):
                violations.append(
                    "focused litchi-numbers table-cells read API exposes "
                    "public table-cells owner alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in flat_alias_exports:
                    violations.append(
                        "focused litchi-numbers table-cells read API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in NUMBERS_TABLE_CELLS_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-numbers table-cells read API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cells_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-numbers table-cells read API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers table-cells read API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_numbers_table_cells_mutation_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the nested, archive-free Numbers table-cell mutation API."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in NUMBERS_TABLE_CELLS_MUTATION_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in NUMBERS_TABLE_CELLS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, frozenset(NUMBERS_TABLE_CELLS_MUTATION_TYPES)
    )
    for name in NUMBERS_TABLE_CELLS_MUTATION_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-numbers table-cells mutation API is missing "
            f"canonical table::cells type {name}: "
            f"{NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE}"
        )

    package_export = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        if NUMBERS_PACKAGE_TABLE_CELL_EDIT_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-numbers table-cells mutation API is missing "
                "private package owner module: "
                f"{package_export.relative_to(root)}"
            )
        for match in PUBLIC_NUMBERS_PACKAGE_TABLE_CELL_EDIT_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers table-cells mutation API exposes "
                "package::table_cell_edit module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-numbers table-cells mutation API is missing "
            "private package owner module: "
            f"{NUMBERS_SOURCE_ROOT / 'package.rs'}"
        )

    owner_path = root / NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-numbers table-cells mutation API is missing "
            "private package owner source: "
            f"{NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE}"
        )

    package_methods: set[str] = set()
    if owner_path.is_file():
        for declaration, _line_number in _rust_public_declarations(
            owner_path.read_text(encoding="utf-8")
        ):
            function = RUST_FUNCTION_DECLARATION.search(declaration)
            if function is not None:
                package_methods.add(function.group(1))
    for name in NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS:
        if name in package_methods:
            continue
        violations.append(
            "focused litchi-numbers table-cells mutation API is missing "
            f"canonical Package::{name} method: "
            f"{NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE}"
        )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = path.read_text(encoding="utf-8")
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for (
            declaration,
            line_number,
            public_declaration,
            complete_source_scope,
        ) in declarations:
            if not _is_numbers_table_cells_mutation_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _numbers_table_cells_mutation_owner_declaration(
                declaration
            )
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = declaration_identifiers[:2] in (
                ["pub", "type"],
                ["pub", "use"],
            )
            flat_alias_exports = _rust_canonical_exports(
                declaration, NUMBERS_TABLE_CELLS_FULL_FLAT_ALIASES
            )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and public_use_or_type
            ):
                violations.append(
                    "focused litchi-numbers table-cells mutation API exposes "
                    "public mutation-owner alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in flat_alias_exports:
                    violations.append(
                        "focused litchi-numbers table-cells mutation API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES
                ):
                    violations.append(
                        "focused litchi-numbers table-cells mutation API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _numbers_table_cells_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-numbers table-cells mutation API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-numbers table-cells mutation API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
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
                    or identifier.startswith("litchi_iwa_")
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


def audit_numbers_document_public_api(root: Path = ROOT) -> list[str]:
    """Keep the focused Numbers reader public API archive-free and semantic."""

    violations: list[str] = []
    for relative in NUMBERS_DOCUMENT_PUBLIC_API_SOURCES:
        path = root / relative
        if not path.is_file():
            continue
        source = path.read_text(encoding="utf-8")
        dedicated_source = relative == NUMBERS_SOURCE_ROOT / "document.rs"
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated_source and not (
                identifiers & NUMBERS_DOCUMENT_PUBLIC_MARKERS
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
                    "focused litchi-numbers document reader public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )

        if not dedicated_source:
            continue
        doc_regions = [
            *re.finditer(r"^[ \t]*//[/!][^\r\n]*", source, re.MULTILINE),
            *re.finditer(r"/\*(?:\*|!)[\s\S]*?\*/", source),
            *re.finditer(r"#\s*\[\s*doc\s*=\s*[^\]]*\]", source),
        ]
        for region in doc_regions:
            for match in RUST_IDENTIFIER.finditer(region.group(0)):
                identifier = match.group(1)
                if not identifier[:1].isupper():
                    continue
                if identifier == "Archive":
                    continue
                reason = _iwork_public_leak(identifier)
                if reason is None:
                    words = [
                        word.lower() for word in CAMEL_CASE_WORD.findall(identifier)
                    ]
                    if "native" in words and any(
                        word in {"id", "identifier", "object"} for word in words
                    ):
                        reason = "native object"
                if reason is None:
                    continue
                offset = region.start() + match.start(1)
                line_number = source.count("\n", 0, offset) + 1
                violations.append(
                    "focused litchi-numbers document reader rustdoc exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_numbers_identity_boundary_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep Numbers duplicate identities on the strict alias boundary.

    Numbers 14.4 can emit an exact physical copy of an object in two
    components after a sheet duplication.  The package may eventually
    coalesce those copies, but an identifier is not a component-local key:
    same-component duplicates and divergent cross-component objects must
    still fail closed.  This source ratchet deliberately does not prescribe a
    component-qualified public identity API.  It only checks the internal
    admission boundary and keeps raw object identifiers out of ``Package``.

    The existing reject-all implementation is valid.  Additional checks are
    activated when the index starts retaining/deduplicating aliases, so a
    future implementation cannot silently turn duplicate rejection into
    first-wins behavior without also proving semantic equality, same-component
    rejection, and physical-object accounting.
    """

    violations: list[str] = []
    alias_mode = False
    index_path = root / NUMBERS_INDEX_SOURCE
    if index_path.is_file():
        raw_index = index_path.read_text(encoding="utf-8")
        index_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_index))

        identifier_comparison = re.compile(
            r"\b(?:identifier|id)\b[ \t\r\n]*(?:==|!=)[\s\S]{0,120}"
            r"\b(?:identifier|id)\b"
        )
        duplicate_regions = [
            index_code[max(0, match.start() - 500) : match.end() + 700]
            for match in identifier_comparison.finditer(index_code)
        ]
        has_duplicate_rejection = any(
            re.search(r"\breturn[ \t\r\n]+Err\b", region) is not None
            for region in duplicate_regions
        ) or re.search(
            r"\b(?:duplicate|same_identifier|duplicate_identifier)\b"
            r"[\s\S]{0,700}\breturn[ \t\r\n]+Err\b",
            index_code,
        ) is not None
        if not has_duplicate_rejection:
            violations.append(
                "focused litchi-numbers index must reject duplicate object identities "
                "unless strict alias admission is proven"
            )

        alias_action = re.compile(
            r"\b(?:continue|retain|dedup|coalesc|or_insert|unique|alias)\w*\b",
            re.IGNORECASE,
        )
        alias_mode = bool(NUMBERS_INDEX_SEMANTIC_ALIAS_WORD.search(index_code)) or any(
            alias_action.search(region) is not None for region in duplicate_regions
        )
        if alias_mode:
            # Alias admission must delegate to the core helper.  Inline
            # ArchiveInfo/messages equality is intentionally insufficient:
            # it can omit retained raw headers and framing widths.
            has_trusted_content_call = any(
                re.search(
                    r"\bsame_content_ignoring_offsets\s*\(",
                    region,
                )
                is not None
                for region in duplicate_regions
            )
            if not has_trusted_content_call:
                violations.append(
                    "focused litchi-numbers index alias admission must call "
                    "same_content_ignoring_offsets"
                )

            component_guard = re.compile(
                r"\bcomponent(?:_index)?\b[\s\S]{0,180}"
                r"(?:==|!=)[\s\S]{0,180}"
                r"\bcomponent(?:_index)?\b[\s\S]{0,500}"
                r"(?:return[ \t\r\n]+Err|continue|reject|InvalidFormat)",
            )
            if component_guard.search(index_code) is None:
                violations.append(
                    "focused litchi-numbers index alias admission must reject "
                    "same-component duplicate identities"
                )

            index_struct = re.search(
                r"\bstruct[ \t\r\n]+Index\b[ \t\r\n]*\{(?P<body>[\s\S]*?)\n\}",
                index_code,
            )
            physical_field = None
            if index_struct is not None:
                physical_field = re.search(
                    r"\b(?P<name>(?:physical|source|archive)"
                    r"[A-Za-z0-9_]*(?:count|objects?))\b",
                    index_struct.group("body"),
                    re.IGNORECASE,
                )
            object_count_method = re.search(
                r"\bfn[ \t\r\n]+object_count\b[^{]*\{(?P<body>[\s\S]*?)\}",
                index_code,
            )
            if physical_field is None or object_count_method is None:
                violations.append(
                    "focused litchi-numbers index alias admission must retain a "
                    "physical object count separate from unique locators"
                )
            elif physical_field.group("name") not in object_count_method.group("body"):
                violations.append(
                    "focused litchi-numbers Package::object_count must report the "
                    "physical count, not unique locator length"
                )
            rebuild_work_method = re.search(
                r"\bfn[ \t\r\n]+rebuild_work\b[^{]*\{(?P<body>[\s\S]*?)\}",
                index_code,
            )
            if rebuild_work_method is None or physical_field is None:
                violations.append(
                    "focused litchi-numbers index rebuild_work must derive sort "
                    "comparison work from the physical object count"
                )
            else:
                rebuild_body = rebuild_work_method.group("body")
                physical_name = physical_field.group("name")
                if physical_name not in rebuild_body or re.search(
                    r"\blookup_work\s*\(", rebuild_body
                ) is not None:
                    violations.append(
                        "focused litchi-numbers index rebuild_work must derive sort "
                        "comparison work from physical count, not unique lookup_work"
                    )
                if re.search(r"\bcomparison_work\s*\(", rebuild_body) is None:
                    violations.append(
                        "focused litchi-numbers index rebuild_work must retain an "
                        "explicit physical comparison-work term"
                    )
            if re.search(
                r"iter_objects\(\)[\s\S]{0,1200}max_objects",
                index_code,
            ) is None:
                violations.append(
                    "focused litchi-numbers index must apply the physical object "
                    "limit before alias coalescing"
                )

    core_path = root / NUMBERS_CORE_ARCHIVE_SOURCE
    if core_path.is_file():
        raw_core = core_path.read_text(encoding="utf-8")
        core_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_core))
        helper_declaration = re.search(
            r"\bfn[ \t\r\n]+same_content_ignoring_offsets\b",
            core_code,
        )
        helper_body = None
        if helper_declaration is not None:
            opening = core_code.find("{", helper_declaration.end())
            if opening >= 0:
                depth = 1
                cursor = opening + 1
                while cursor < len(core_code) and depth:
                    if core_code[cursor] == "{":
                        depth += 1
                    elif core_code[cursor] == "}":
                        depth -= 1
                    cursor += 1
                if depth == 0:
                    helper_body = core_code[opening + 1 : cursor - 1]
        required_fields = (
            "archive_info",
            "messages",
            "header_length",
            "data_length",
            "original_header",
            "original_canonical_header",
        )
        if helper_body is None:
            violations.append(
                "focused litchi-iwa-core archive helper "
                "same_content_ignoring_offsets is missing"
            )
        else:
            resolved_fields = {
                field_name
                for field_name in required_fields
                if re.search(rf"\b{re.escape(field_name)}\b", helper_body)
                is not None
            }
            source_headers_calls = re.findall(
                r"\bsource_headers\s*\(", helper_body
            )
            source_headers_comparison = re.search(
                r"(?:\bself\s*\.\s*source_headers\s*\(\)\s*==\s*"
                r"other\s*\.\s*source_headers\s*\(\)|"
                r"\bother\s*\.\s*source_headers\s*\(\)\s*==\s*"
                r"self\s*\.\s*source_headers\s*\(\))",
                helper_body,
            )
            if source_headers_calls:
                if len(source_headers_calls) < 2:
                    violations.append(
                        "focused litchi-iwa-core archive helper "
                        "same_content_ignoring_offsets must compare both "
                        "source_headers values"
                    )
                elif source_headers_comparison is None:
                    violations.append(
                        "focused litchi-iwa-core archive helper "
                        "same_content_ignoring_offsets must compare self and "
                        "other source_headers values"
                    )
                source_headers_body = _rust_any_function_body(
                    core_code, "source_headers"
                )
                if (
                    source_headers_body is not None
                    and source_headers_comparison is not None
                ):
                    resolved_fields.update(
                        field_name
                        for field_name in (
                            "original_header",
                            "original_canonical_header",
                        )
                        if re.search(
                            rf"\b{re.escape(field_name)}\b", source_headers_body
                        )
                        is not None
                    )
            for field_name in required_fields:
                if field_name not in resolved_fields:
                    violations.append(
                        "focused litchi-iwa-core archive helper "
                        f"same_content_ignoring_offsets omits {field_name}"
                    )
            if re.search(r"\boffset\w*\b", helper_body, re.IGNORECASE) is not None:
                violations.append(
                    "focused litchi-iwa-core archive helper "
                    "same_content_ignoring_offsets compares an offset"
                )
    elif alias_mode:
        violations.append(
            "focused litchi-iwa-core archive helper "
            "same_content_ignoring_offsets is missing"
        )

    package_sources = [root / NUMBERS_PACKAGE_SOURCE]
    package_root = root / NUMBERS_SOURCE_ROOT / "package"
    if package_root.is_dir():
        package_sources.extend(sorted(package_root.rglob("*.rs")))
    for package_path in package_sources:
        if not package_path.is_file():
            continue
        raw_package = package_path.read_text(encoding="utf-8")
        package_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_package))
        for declaration, line_number in _rust_public_declarations(package_code):
            # ``_rust_public_declarations`` also returns public fields inside
            # private or pub(super) structs.  Those fields are not a public
            # Package API; only inspect function declarations for raw-ID
            # parameters and method names.
            if re.search(r"\bfn\b", declaration) is None:
                continue
            method_match = NUMBERS_INDEX_RAW_ID_PUBLIC_METHOD.search(declaration)
            parameter_match = re.search(
                r"\b(?:identifier|object_id|native_id|object_identifier)\b"
                r"[ \t\r\n]*:[ \t\r\n]*(?:u32|u64|NonZeroU(?:32|64))\b",
                declaration,
            )
            if method_match is None and parameter_match is None:
                continue
            leaked_name = (
                method_match.group(0).split()[-1]
                if method_match is not None
                else parameter_match.group(0).split()[0]
            )
            violations.append(
                "focused litchi-numbers Package public API exposes raw object "
                f"identity {leaked_name}: "
                f"{package_path.relative_to(root)}:{line_number}"
            )

    return sorted(set(violations))


def _numbers_names_publication_source_paths(root: Path) -> tuple[Path, ...]:
    """Return the names owner files participating in publication routing."""

    paths = {
        root / relative
        for relative in NUMBERS_NAMES_IMPLEMENTATION_SOURCES
        if (root / relative).is_file()
    }
    helper_root = root / NUMBERS_NAMES_PUBLICATION_HELPER_ROOT
    if helper_root.is_dir():
        paths.update(helper_root.rglob("*.rs"))
    return tuple(sorted(paths))


def _numbers_names_save_token_call_patterns(
    production_sources: Iterable[tuple[Path, str]],
) -> tuple[re.Pattern[str], ...]:
    """Build canonical save-token call patterns from production imports.

    A publication route may import the hidden codec module or the one function
    directly.  The import must originate in ``litchi_iwa_protos``; accepting an
    arbitrary local module called ``package_metadata_codec`` would let a legacy
    host adapter satisfy the ratchet by name alone.
    """

    patterns: list[re.Pattern[str]] = [
        re.compile(
            r"(?<![A-Za-z0-9_:#.])litchi_iwa_protos"
            r"[ \t\r\n]*::[ \t\r\n]*package_metadata_codec"
            r"[ \t\r\n]*::[ \t\r\n]*"
            + re.escape(NUMBERS_NAMES_SAVE_TOKEN_REWRITE)
            + r"[ \t\r\n]*\("
        )
    ]
    module_aliases: set[str] = set()
    bare_function_aliases: set[str] = set()
    bare_function_imported = False
    for _path, source in production_sources:
        masked_source = _mask_rust_non_code(source)
        # Also accept grouped imports such as
        # ``use litchi_iwa_protos::{package_metadata_codec, ...};``.  The
        # direct-module regex below intentionally stays narrow for the common
        # path form, while this pass handles Rust's brace import grammar.
        for use_match in NUMBERS_NAMES_SAVE_TOKEN_IWA_USE.finditer(masked_source):
            imported_path = use_match.group("path")
            module_match = re.search(
                r"(?<![A-Za-z0-9_])package_metadata_codec\b"
                r"(?P<suffix>[ \t\r\n]*(?:::|as)[^,}]*)?",
                imported_path,
            )
            if module_match is None:
                continue
            suffix = module_match.group("suffix") or ""
            alias_match = re.search(
                r"\bas[ \t\r\n]+(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*)\b",
                suffix,
            )
            if alias_match is not None and "{" not in suffix:
                module_aliases.add(alias_match.group("alias"))
            if "::" not in suffix:
                module_aliases.add("package_metadata_codec")
            if NUMBERS_NAMES_SAVE_TOKEN_REWRITE in imported_path:
                bare_function_imported = True
        for import_match in NUMBERS_NAMES_SAVE_TOKEN_MODULE_IMPORT.finditer(
            masked_source
        ):
            tail = import_match.group("tail")
            alias_match = NUMBERS_NAMES_SAVE_TOKEN_MODULE_ALIAS.search(tail)
            if alias_match is not None and "{" not in tail:
                module_aliases.add(alias_match.group("alias"))
            if re.search(
                rf"(?:^|[{{, \t\r\n:])(?:r#)?"
                rf"{re.escape(NUMBERS_NAMES_SAVE_TOKEN_REWRITE)}"
                rf"(?:[ \t\r\n]+as[ \t\r\n]+(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?"
                rf"(?=[}},; \t\r\n]|$)",
                tail,
            ) is not None:
                function_match = re.search(
                    rf"(?:r#)?{re.escape(NUMBERS_NAMES_SAVE_TOKEN_REWRITE)}"
                    rf"(?:[ \t\r\n]+as[ \t\r\n]+"
                    rf"(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?",
                    tail,
                )
                if function_match is not None:
                    bare_function_imported = True
                    if function_match.group("alias"):
                        bare_function_aliases.add(function_match.group("alias"))
            if "*" in tail:
                bare_function_imported = True
        # ``use ...::package_metadata_codec;`` has no ``as`` clause and is
        # consequently handled explicitly above rather than by the alias regex.
        for import_match in NUMBERS_NAMES_SAVE_TOKEN_MODULE_IMPORT.finditer(
            masked_source
        ):
            tail = import_match.group("tail").strip()
            if tail in {"", "as package_metadata_codec"}:
                module_aliases.add("package_metadata_codec")

    for alias in sorted(module_aliases):
        patterns.append(
            re.compile(
                rf"(?<![A-Za-z0-9_:#.]){re.escape(alias)}"
                rf"[ \t\r\n]*::[ \t\r\n]*"
                + re.escape(NUMBERS_NAMES_SAVE_TOKEN_REWRITE)
                + r"[ \t\r\n]*\("
            )
        )
    if bare_function_imported:
        patterns.append(
            re.compile(
                r"(?<![A-Za-z0-9_:#.])(?:r#)?"
                + re.escape(NUMBERS_NAMES_SAVE_TOKEN_REWRITE)
                + r"[ \t\r\n]*\("
            )
        )
    for alias in sorted(bare_function_aliases):
        patterns.append(
            re.compile(
                rf"(?<![A-Za-z0-9_:#.]){re.escape(alias)}"
                r"[ \t\r\n]*\("
            )
        )
    return tuple(patterns)


def _numbers_names_normal_manifest_dependency_lines(
    manifest_path: Path,
) -> tuple[dict[str, int], bool]:
    """Return normal dependency names/lines and whether the canonical owner exists."""

    dependencies: dict[str, int] = {}
    if not manifest_path.is_file():
        return dependencies, False
    section: str | None = None
    for line_number, line in enumerate(
        manifest_path.read_text(encoding="utf-8").splitlines(), start=1
    ):
        header = CARGO_SECTION_HEADER.match(line)
        if header is not None:
            section = header.group(1).strip()
            continue
        normal_dependencies = section == "dependencies" or (
            section is not None
            and section.endswith(".dependencies")
            and not section.endswith(".dev-dependencies")
        )
        if not normal_dependencies:
            continue
        dependency_match = re.match(
            r"^[ \t]*(?:\"(?P<quoted>[A-Za-z0-9_-]+)\"|"
            r"(?P<plain>[A-Za-z0-9_-]+))"
            r"[ \t]*(?:\.[ \t]*workspace)?[ \t]*=",
            line,
        )
        if dependency_match is None:
            continue
        dependency = dependency_match.group("quoted") or dependency_match.group("plain")
        dependencies.setdefault(dependency, line_number)
        package_match = re.search(
            r"\bpackage[ \t]*=[ \t]*\"(?P<package>[A-Za-z0-9_-]+)\"",
            line,
        )
        if package_match is not None:
            dependencies.setdefault(package_match.group("package"), line_number)
    return dependencies, NUMBERS_NAMES_SAVE_TOKEN_REQUIRED_NORMAL_DEPENDENCY in dependencies


def audit_numbers_names_save_token_manifest_topology(root: Path = ROOT) -> list[str]:
    """Keep the names publication owner on the focused codec dependency edge."""

    manifest_path = root / NUMBERS_PACKAGE_MANIFEST
    if not manifest_path.is_file():
        return []
    dependencies, has_codec = _numbers_names_normal_manifest_dependency_lines(
        manifest_path
    )
    violations: list[str] = []
    for dependency in sorted(NUMBERS_NAMES_SAVE_TOKEN_FORBIDDEN_NORMAL_DEPENDENCIES):
        line_number = dependencies.get(dependency)
        if line_number is None:
            continue
        violations.append(
            "focused litchi-numbers names publication manifest retains normal "
            f"{dependency} dependency: {NUMBERS_PACKAGE_MANIFEST}:{line_number}"
        )
    if not has_codec:
        violations.append(
            "focused litchi-numbers names publication manifest is missing normal "
            f"{NUMBERS_NAMES_SAVE_TOKEN_REQUIRED_NORMAL_DEPENDENCY} dependency: "
            f"{NUMBERS_PACKAGE_MANIFEST}"
        )
    return sorted(set(violations))


def audit_numbers_names_save_token_publication_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require name publication to use the bounded metadata save-token seam.

    This is a production call-graph check, not a textual module check.  The
    route starts at ``rewrite_names`` and follows local helper calls across the
    names owner files.  Test-gated builders/oracles are masked item-by-item so
    a cfg(test) decode or legacy helper cannot hide production code that comes
    later in the file, and unreachable helpers remain harmless.
    """

    violations: list[str] = audit_numbers_names_save_token_manifest_topology(root)
    source_paths = _numbers_names_publication_source_paths(root)
    if not source_paths:
        return violations

    production_sources: list[tuple[Path, str]] = []
    production_source_by_path: dict[Path, str] = {}
    functions: dict[str, list[tuple[Path, str, int]]] = {}
    for path in source_paths:
        raw_source = path.read_text(encoding="utf-8")
        production_source = _mask_rust_cfg_test_items(raw_source)
        production_sources.append((path, production_source))
        production_source_by_path[path] = production_source
        for name, (body, body_offset) in _rust_top_level_function_bodies(
            production_source
        ).items():
            functions.setdefault(name, []).append((path, body, body_offset))

    rewrite_records = functions.get(NUMBERS_NAMES_PUBLICATION_FUNCTION, [])
    if not rewrite_records:
        violations.append(
            "focused litchi-numbers names publication route is missing production "
            f"{NUMBERS_NAMES_PUBLICATION_FUNCTION}: "
            f"{NUMBERS_NAMES_PACKAGE_SOURCE}"
        )
        return sorted(set(violations))

    route_patterns = _numbers_names_save_token_call_patterns(production_sources)
    reachable: list[tuple[str, Path, str, int]] = []
    pending = [(NUMBERS_NAMES_PUBLICATION_FUNCTION, *record) for record in rewrite_records]
    visited: set[tuple[str, Path]] = set()
    while pending:
        current_name, path, body, body_offset = pending.pop()
        visit_key = (current_name, path)
        if visit_key in visited:
            continue
        visited.add(visit_key)
        reachable.append((current_name, path, body, body_offset))
        for helper_name, records in functions.items():
            if helper_name == current_name:
                continue
            call = re.compile(
                rf"(?<![A-Za-z0-9_:#.])(?:"
                rf"(?:crate|self|super|names|package)[ \t\r\n]*::[ \t\r\n]*"
                rf")*(?:r#)?{re.escape(helper_name)}"
                r"[ \t\r\n]*\("
            )
            if call.search(body) is None:
                continue
            pending.extend((helper_name, *record) for record in records)

    route_found = False
    for current_name, path, body, body_offset in reachable:
        for pattern in route_patterns:
            if pattern.search(body) is not None:
                route_found = True
                break
        for label, pattern in NUMBERS_NAMES_SAVE_TOKEN_ROUTE_FORBIDDEN_PATTERNS:
            for match in pattern.finditer(body):
                source = production_source_by_path[path]
                line_number = (
                    source.count("\n", 0, body_offset + match.start()) + 1
                )
                helper_suffix = (
                    ""
                    if current_name == NUMBERS_NAMES_PUBLICATION_FUNCTION
                    else f" via helper {current_name}"
                )
                violations.append(
                    "focused litchi-numbers names publication route "
                    f"{label}{helper_suffix}: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    if not route_found:
        path, _body, body_offset = rewrite_records[0]
        production_source = production_source_by_path[path]
        line_number = production_source.count("\n", 0, body_offset) + 1
        violations.append(
            "focused litchi-numbers names publication route does not call "
            f"{NUMBERS_NAMES_SAVE_TOKEN_REWRITE}: "
            f"{path.relative_to(root)}:{line_number}"
        )

    return sorted(set(violations))


def audit_numbers_package_no_eager_prost_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers package ingress on the focused normal dependency edge.

    This is intentionally scoped to ``package.rs`` and the focused
    ``litchi-numbers`` manifest.  Generated Prost fixtures remain available in
    test-only modules, while production metadata/native readers use the normal
    ``litchi-iwa-protos`` owner and cannot regain the ``prost-types`` edge
    without a boundary review.
    The production source slice excludes cfg(test)-gated items individually;
    cfg(test)-gated imports and helpers cannot hide later production code.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_PACKAGE_SOURCE
    if source_path.is_file():
        raw_source = source_path.read_text(encoding="utf-8")
        production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
        for label, pattern in NUMBERS_PACKAGE_NO_EAGER_PROST_SOURCE_PATTERNS:
            for match in pattern.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused litchi-numbers package production source uses "
                    f"{label}: {NUMBERS_PACKAGE_SOURCE}:{line_number}"
                )

    manifest_path = root / NUMBERS_PACKAGE_MANIFEST
    if manifest_path.is_file():
        dependencies, has_codec = _numbers_names_normal_manifest_dependency_lines(
            manifest_path
        )
        for dependency in sorted(
            NUMBERS_NAMES_SAVE_TOKEN_FORBIDDEN_NORMAL_DEPENDENCIES
        ):
            line_number = dependencies.get(dependency)
            if line_number is None:
                continue
            violations.append(
                "focused litchi-numbers Cargo manifest retains normal "
                f"{dependency} dependency: {NUMBERS_PACKAGE_MANIFEST}:{line_number}"
            )
        if not has_codec:
            violations.append(
                "focused litchi-numbers Cargo manifest is missing normal "
                f"{NUMBERS_NAMES_SAVE_TOKEN_REQUIRED_NORMAL_DEPENDENCY} dependency: "
                f"{NUMBERS_PACKAGE_MANIFEST}"
            )

    return sorted(set(violations))


def audit_numbers_extractor_no_eager_rich_text_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers extractor's rich-text ingress archive-free.

    The extractor still has independently migrating generated readers for the
    table-model and cell-storage paths.  This narrow ratchet protects the
    rich-text path that already has a bounded wire projection: a regression to
    ``RichTextPayloadArchive::decode`` would allocate the complete generated
    envelope before its storage-reference and text budgets are applied.  As
    with the package ingress ratchet, canonical Prost fixtures may remain in a
    test-only module.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_EXTRACTOR_SOURCE
    if not source_path.is_file():
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    for label, pattern in NUMBERS_EXTRACTOR_NO_EAGER_RICH_TEXT_SOURCE_PATTERNS:
        for match in pattern.finditer(production_code):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers extractor production source uses "
                f"{label}: {NUMBERS_EXTRACTOR_SOURCE}:{line_number}"
            )

    return sorted(set(violations))


def audit_numbers_extractor_no_eager_tile_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers extractor's tile ingress free of eager Prost reads.

    The extractor retains unrelated generated readers while the tile projection
    migration is completed.  This focused ratchet prevents production code from
    reconstructing complete ``tst::Tile`` or ``tst::TileRowInfo`` messages after
    the bounded tile projection has already validated the payload.  Test-only
    Prost oracles and builders remain in cfg(test)-gated items outside the
    production source slice.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_EXTRACTOR_SOURCE
    if not source_path.is_file():
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    for label, pattern in NUMBERS_EXTRACTOR_NO_EAGER_TILE_SOURCE_PATTERNS:
        for match in pattern.finditer(production_code):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers extractor production source uses "
                f"{label}: {NUMBERS_EXTRACTOR_SOURCE}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_numbers_wire_formula_render_ownership(
    root: Path = ROOT,
) -> list[str]:
    """Keep Numbers formula event rendering in the low-level wire owner.

    ``litchi-numbers-wire`` is the one sanctioned bridge between the generated
    Numbers formula codec and the shared, archive-free expression arena.  The
    host and focused readers continue to retain bounded archive admission and
    decoder report charging, but their semantic event visitor and scalar
    expression walk must be delegated to the wire module.  Missing reader
    files are accepted so the ratchet survives the eventual monolith exit;
    any reader that remains is checked strictly after test-only items and
    non-code have been masked.
    """

    violations: list[str] = []
    wire_path = root / IWA_NUMBERS_WIRE_FORMULA_RENDER_SOURCE
    if not wire_path.is_file():
        return [
            "Numbers wire formula renderer owner is missing: "
            f"{IWA_NUMBERS_WIRE_FORMULA_RENDER_SOURCE}"
        ]

    wire_raw = wire_path.read_text(encoding="utf-8")
    wire_code = _mask_rust_non_code(_mask_rust_cfg_test_items(wire_raw))
    for marker in IWA_NUMBERS_WIRE_FORMULA_RENDER_REQUIRED_MARKERS:
        if re.search(rf"\b{re.escape(marker)}\b", wire_code) is None:
            violations.append(
                "Numbers wire formula renderer is missing shared "
                f"{marker} route: {IWA_NUMBERS_WIRE_FORMULA_RENDER_SOURCE}"
            )
    if IWA_NUMBERS_WIRE_FORMULA_RENDER_COMMON_IMPORT.search(wire_code) is None:
        violations.append(
            "Numbers wire formula renderer is missing the common expression "
            f"arena import: {IWA_NUMBERS_WIRE_FORMULA_RENDER_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_FORMULA_RENDER_PROTOS_IMPORT.search(wire_code) is None:
        violations.append(
            "Numbers wire formula renderer is missing its sanctioned generated "
            "codec import: "
            f"{IWA_NUMBERS_WIRE_FORMULA_RENDER_SOURCE}"
        )

    for relative in IWA_FORMULA_RENDER_READER_SOURCES:
        path = root / relative
        if not path.is_file():
            # The old host reader is allowed to disappear when its callers
            # have moved to the focused package.  A present file must still
            # delegate so a stale copy cannot quietly return later.
            continue
        raw_source = path.read_text(encoding="utf-8")
        production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
        if IWA_FORMULA_RENDER_WIRE_IMPORT.search(production_code) is None:
            violations.append(
                "formula renderer reader is missing the Numbers wire owner "
                f"import: {relative}"
            )
        for label, declaration in IWA_FORMULA_RENDER_LOCAL_VISITOR_DECLARATIONS:
            for match in declaration.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "formula renderer reader redeclares the Numbers wire "
                    f"owner {label}: {relative}:{line_number}"
                )

    return sorted(set(violations))


def audit_iwa_numbers_wire_cell_value_ownership(
    root: Path = ROOT,
) -> list[str]:
    """Keep typed Numbers cell-value interpretation in the wire owner.

    ``cell_value.rs`` owns only the borrowed BNC/pre-BNC value projection.  Its
    public result is copyable scalar data plus unresolved sidecar identifiers,
    so adapters can perform their own bounded table lookup and text retention.
    The compatibility host and focused package reader may keep their error
    mapping and sidecar work, but must delegate the version/type interpretation
    to this module.  A reader file may disappear as the monolith exits; any
    reader that remains is checked after test-only items and non-code have been
    masked.

    The zero-allocation check is intentionally limited to the successful
    result shape.  BNC/pre-BNC parsers may retain owned diagnostic text when
    rejecting malformed input, and the adapter's error wrapper is outside
    this ratchet.
    """

    violations: list[str] = []
    wire_path = root / IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE
    lib_path = root / IWA_NUMBERS_WIRE_CELL_VALUE_LIB_SOURCE

    if not wire_path.is_file():
        return [
            "Numbers wire cell-value owner is missing: "
            f"{IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE}"
        ]

    wire_raw = wire_path.read_text(encoding="utf-8")
    wire_code = _mask_rust_non_code(_mask_rust_cfg_test_items(wire_raw))

    if not lib_path.is_file():
        violations.append(
            "Numbers wire cell-value module is not exported from the crate: "
            f"{IWA_NUMBERS_WIRE_CELL_VALUE_LIB_SOURCE}"
        )
    else:
        lib_raw = lib_path.read_text(encoding="utf-8")
        lib_code = _mask_rust_non_code(_mask_rust_cfg_test_items(lib_raw))
        if IWA_NUMBERS_WIRE_CELL_VALUE_MODULE_EXPORT.search(lib_code) is None:
            violations.append(
                "Numbers wire cell-value owner is missing its public module export: "
                f"{IWA_NUMBERS_WIRE_CELL_VALUE_LIB_SOURCE}"
            )

    if IWA_NUMBERS_WIRE_CELL_VALUE_PUBLIC_DECODER.search(wire_code) is None:
        violations.append(
            "Numbers wire cell-value owner must expose a borrowed typed decoder: "
            f"{IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE}"
        )

    for label, declaration in IWA_NUMBERS_WIRE_CELL_VALUE_PUBLIC_TYPES:
        match = declaration.search(wire_code)
        if match is None:
            violations.append(
                "Numbers wire cell-value owner is missing public "
                f"{label}: {IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE}"
            )
            continue

        # A Copy result makes the successful projection's allocation contract
        # visible at the API boundary.  Search only the nearby derive list so
        # an unrelated type elsewhere in this module cannot satisfy it.
        previous_item_end = wire_code.rfind("}", 0, match.start())
        prefix = wire_code[previous_item_end + 1 : match.start()]
        if IWA_NUMBERS_WIRE_CELL_VALUE_COPY_DERIVE.search(prefix) is None:
            line_number = wire_code.count("\n", 0, match.start()) + 1
            violations.append(
                "Numbers wire cell-value public result must derive Copy: "
                f"{IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE}:{line_number}"
            )

        # Inspect only this declaration's balanced body (or its type-alias
        # statement), rather than the whole module.  Error enums and parser
        # internals are deliberately outside the successful result check.
        opening = wire_code.find("{", match.end())
        semicolon = wire_code.find(";", match.end())
        if opening >= 0 and (semicolon < 0 or opening < semicolon):
            end = _rust_balanced_delimited_end(wire_code, opening)
            declaration_code = (
                wire_code[match.start() : end] if end is not None else ""
            )
        else:
            end = len(wire_code) if semicolon < 0 else semicolon + 1
            declaration_code = wire_code[match.start() : end]
        owned = IWA_NUMBERS_WIRE_CELL_VALUE_OWNED_RESULT_TYPES.search(
            declaration_code
        )
        if owned is not None:
            line_number = wire_code.count("\n", 0, match.start() + owned.start()) + 1
            violations.append(
                "Numbers wire cell-value public result owns allocation-backed "
                f"data ({owned.group()}): {IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE}:"
                f"{line_number}"
            )

    for marker in IWA_NUMBERS_WIRE_CELL_VALUE_PARSE_MARKERS:
        if re.search(
            rf"\b{re.escape(marker)}\b[ \t\r\n]*::[ \t\r\n]*parse\b",
            wire_code,
        ) is None:
            violations.append(
                "Numbers wire cell-value owner is missing its "
                f"{marker} parse route: {IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE}"
            )

    for statement in IWA_NUMBERS_WIRE_CELL_VALUE_IMPORT_STATEMENT.finditer(
        wire_code
    ):
        forbidden = IWA_NUMBERS_WIRE_CELL_VALUE_FORBIDDEN_IMPORT.search(
            statement.group()
        )
        if forbidden is None:
            continue
        line_number = wire_code.count("\n", 0, statement.start()) + 1
        violations.append(
            "Numbers wire cell-value owner imports a format/protobuf crate "
            f"({forbidden.group()}): {IWA_NUMBERS_WIRE_CELL_VALUE_SOURCE}:"
            f"{line_number}"
        )

    for relative in IWA_NUMBERS_CELL_VALUE_READER_SOURCES:
        path = root / relative
        if not path.is_file():
            # Host deletion is the intended monolith-exit state.  A present
            # reader must still delegate so a stale duplicate cannot return.
            continue
        raw_source = path.read_text(encoding="utf-8")
        production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
        if IWA_NUMBERS_CELL_VALUE_READER_IMPORT.search(production_code) is None:
            violations.append(
                "Numbers cell-value reader is missing the shared wire owner "
                f"import: {relative}"
            )
        if IWA_NUMBERS_CELL_VALUE_READER_CALL.search(production_code) is None:
            violations.append(
                "Numbers cell-value reader is missing the shared typed decoder "
                f"call: {relative}"
            )
        for label, pattern in (
            *IWA_NUMBERS_CELL_VALUE_READER_LEGACY_TYPES,
            *IWA_NUMBERS_CELL_VALUE_READER_LEGACY_METHODS,
        ):
            for match in pattern.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "Numbers cell-value reader retains duplicated wire "
                    f"interpretation {label}: {relative}:{line_number}"
                )

    return sorted(set(violations))


def audit_iwa_numbers_wire_table_data_list_ownership(
    root: Path = ROOT,
) -> list[str]:
    """Keep table-data-list topology selection in the shared wire owner.

    The low-level coordinator owns only borrowed root/segment message routing,
    duplicate and key-range checks, bounded merge admission, and final sorted
    publication.  Readers keep their generated-free decoder, aggregate budget,
    and sidecar converter hooks.  Missing reader files are accepted so a host
    can disappear during monolith retirement; a reader that remains must call
    the shared coordinator and must not restore the old root/segment loops.
    """

    violations: list[str] = []
    owner_path = root / IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE
    lib_path = root / IWA_NUMBERS_WIRE_TABLE_DATA_LIST_LIB_SOURCE

    if not owner_path.is_file():
        return [
            "Numbers wire table-data-list coordinator is missing: "
            f"{IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}"
        ]

    owner_raw = owner_path.read_text(encoding="utf-8")
    owner_code = _mask_rust_non_code(_mask_rust_cfg_test_items(owner_raw))

    if not lib_path.is_file():
        violations.append(
            "Numbers wire table-data-list coordinator is not exported from the crate: "
            f"{IWA_NUMBERS_WIRE_TABLE_DATA_LIST_LIB_SOURCE}"
        )
    else:
        lib_raw = lib_path.read_text(encoding="utf-8")
        lib_code = _mask_rust_non_code(_mask_rust_cfg_test_items(lib_raw))
        if IWA_NUMBERS_WIRE_TABLE_DATA_LIST_MODULE_EXPORT.search(lib_code) is None:
            violations.append(
                "Numbers wire table-data-list coordinator is missing its public "
                "module export: "
                f"{IWA_NUMBERS_WIRE_TABLE_DATA_LIST_LIB_SOURCE}"
            )

    for label, declaration in IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_TYPES:
        matches = list(declaration.finditer(owner_code))
        if not matches:
            violations.append(
                "Numbers wire table-data-list coordinator is missing public "
                f"{label}: {IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}"
            )
        elif len(matches) > 1:
            line_number = owner_code.count("\n", 0, matches[1].start()) + 1
            violations.append(
                "Numbers wire table-data-list coordinator redeclares shared "
                f"{label}: {IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}:{line_number}"
            )

    if IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_TRAIT.search(owner_code) is None:
        violations.append(
            "Numbers wire table-data-list coordinator is missing public ListDecoder: "
            f"{IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_FUNCTION.search(owner_code) is None:
        violations.append(
            "Numbers wire table-data-list coordinator is missing public read_list: "
            f"{IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_TABLE_DATA_LIST_BORROWED_MESSAGE.search(owner_code) is None:
        violations.append(
            "Numbers wire table-data-list coordinator must carry borrowed Message "
            f"payloads: {IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_TABLE_DATA_LIST_READ_LIST_GENERIC.search(owner_code) is None:
        violations.append(
            "Numbers wire table-data-list read_list must use generic borrowed "
            f"adapter hooks: {IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}"
        )

    for marker in IWA_NUMBERS_WIRE_TABLE_DATA_LIST_REQUIRED_TOPOLOGY_MARKERS:
        if re.search(rf"\b{re.escape(marker)}\b", owner_code) is None:
            violations.append(
                "Numbers wire table-data-list coordinator is missing its "
                f"{marker} topology route: {IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}"
            )

    for statement in IWA_NUMBERS_WIRE_TABLE_DATA_LIST_IMPORT_STATEMENT.finditer(
        owner_code
    ):
        forbidden = IWA_NUMBERS_WIRE_TABLE_DATA_LIST_FORBIDDEN_IMPORT.search(
            statement.group()
        )
        if forbidden is None:
            continue
        line_number = owner_code.count("\n", 0, statement.start()) + 1
        violations.append(
            "Numbers wire table-data-list coordinator imports a concrete "
            f"format/protobuf/core crate ({forbidden.group()}): "
            f"{IWA_NUMBERS_WIRE_TABLE_DATA_LIST_SOURCE}:{line_number}"
        )

    for relative in IWA_NUMBERS_TABLE_DATA_LIST_READER_SOURCES:
        path = root / relative
        if not path.is_file():
            # A missing host is the intended monolith-exit state.  Do not make
            # the coordinator ratchet block deletion of the old reader.
            continue

        raw_source = path.read_text(encoding="utf-8")
        production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
        if IWA_NUMBERS_TABLE_DATA_LIST_READER_IMPORT.search(production_code) is None:
            violations.append(
                "Numbers table-data-list reader is missing the shared wire owner "
                f"import: {relative}"
            )
        if IWA_NUMBERS_TABLE_DATA_LIST_READER_CALL.search(production_code) is None:
            violations.append(
                "Numbers table-data-list reader is missing the shared read_list "
                f"call: {relative}"
            )

        for pattern, label in (
            (
                IWA_NUMBERS_TABLE_DATA_LIST_ROOT_TOPOLOGY_FILTER,
                "root candidate topology filter",
            ),
            (
                IWA_NUMBERS_TABLE_DATA_LIST_SEGMENT_TOPOLOGY_FILTER,
                "segment payload topology filter",
            ),
            (
                IWA_NUMBERS_TABLE_DATA_LIST_SELECTION_STATE,
                "root/segment selection state",
            ),
        ):
            for match in pattern.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "Numbers table-data-list reader retains duplicated "
                    f"{label}: {relative}:{line_number}"
                )

    return sorted(set(violations))


def audit_iwa_numbers_wire_formula_envelope_ownership(
    root: Path = ROOT,
) -> list[str]:
    """Keep FormulaArchive envelope admission in the low-level wire owner.

    The envelope walk validates native wire types, required fields, repeated
    fields, scalar canonicality, UTF-8, and bounded nested paths before a
    format adapter retains bytes.  It is shared independently of the event
    renderer because the adapters still own their error and aggregate-budget
    mapping.  Missing reader files are accepted for monolith retirement, but
    every reader that remains must import this owner and must not carry a
    second schema table or strict preflight implementation.
    """

    violations: list[str] = []
    envelope_path = root / IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE
    if not envelope_path.is_file():
        return [
            "Numbers wire FormulaArchive envelope owner is missing: "
            f"{IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE}"
        ]

    envelope_raw = envelope_path.read_text(encoding="utf-8")
    envelope_code = _mask_rust_non_code(_mask_rust_cfg_test_items(envelope_raw))
    for marker in IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_REQUIRED_MARKERS:
        if re.search(rf"\b{re.escape(marker)}\b", envelope_code) is None:
            violations.append(
                "Numbers wire FormulaArchive envelope owner is missing "
                f"{marker} route: {IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE}"
            )
    if IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_PUBLIC_ENTRY.search(envelope_code) is None:
        violations.append(
            "Numbers wire FormulaArchive envelope owner must expose its "
            f"preflight entry: {IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_COMMON_IMPORT.search(envelope_code) is None:
        violations.append(
            "Numbers wire FormulaArchive envelope owner is missing its common "
            f"wire import: {IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_PROTOS_IMPORT.search(envelope_code) is None:
        violations.append(
            "Numbers wire FormulaArchive envelope owner is missing its "
            "generated-free codec import: "
            f"{IWA_NUMBERS_WIRE_FORMULA_ENVELOPE_SOURCE}"
        )

    for relative in IWA_FORMULA_RENDER_READER_SOURCES:
        path = root / relative
        if not path.is_file():
            # A deleted host reader is the intended monolith-exit state.  The
            # focused reader receives the same allowance when it is retired.
            continue
        raw_source = path.read_text(encoding="utf-8")
        production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
        if IWA_FORMULA_ENVELOPE_WIRE_IMPORT.search(production_code) is None:
            violations.append(
                "formula renderer reader is missing the Numbers wire "
                f"FormulaArchive envelope import: {relative}"
            )
        for label, declaration in IWA_FORMULA_ENVELOPE_LOCAL_SCHEMA_DECLARATIONS:
            for match in declaration.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "formula renderer reader redeclares the Numbers wire "
                    f"FormulaArchive envelope {label}: {relative}:{line_number}"
                )
        for match in IWA_FORMULA_ENVELOPE_LOCAL_SCHEMA_CONSTANTS.finditer(
            production_code
        ):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "formula renderer reader redeclares the Numbers wire "
                f"FormulaArchive envelope schema constant: {relative}:{line_number}"
            )

    return sorted(set(violations))


def audit_numbers_extractor_no_eager_table_data_list_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers table-data-list ingress free of eager Prost reads.

    The table-data-list and segment payloads have bounded wire projections in
    ``numbers_table_cell_storage_codec``.  Reconstructing either generated
    ``TableDataList`` message in production would bypass those limits after
    projection and retain the complete legacy envelope.  The checker is
    intentionally scoped to the extractor's production slice: test-only
    Prost oracles remain available in cfg(test)-gated items, while
    handwritten ``decode_table_data_list*`` helpers and unrelated generated
    readers are not part of this ratchet.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_EXTRACTOR_SOURCE
    if not source_path.is_file():
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    for label, pattern in NUMBERS_EXTRACTOR_NO_EAGER_TABLE_DATA_LIST_SOURCE_PATTERNS:
        for match in pattern.finditer(production_code):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers extractor production source uses "
                f"{label}: {NUMBERS_EXTRACTOR_SOURCE}:{line_number}"
            )

    return sorted(set(violations))


def _numbers_table_cell_comment_reply_owner_present(root: Path) -> bool:
    """Return whether the production Numbers reply projection has activated."""

    source_root = root / NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return False
    for relative in (
        NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE,
        NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0],
    ):
        path = root / relative
        if not path.is_file():
            continue
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        code = _mask_rust_non_code(production_source)
        if NUMBERS_TABLE_CELL_COMMENT_REPLY_ACTIVATION.search(code) is not None:
            return True
    return False


def audit_numbers_table_cell_comment_reply_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the read-only, selector-first Numbers comment-reply facade.

    The public projection is deliberately archive-free: ``CommentReply`` is
    owned by the Numbers package and is re-exported only as
    ``TableCellCommentReply`` at the crate root.  Its two Package read methods
    use ``SheetSelector``/``TableSelector`` and semantic cell positions, while
    strict comment-storage decoding and a package-wide ownership census stay
    behind the private owner boundary.  This ratchet does not require reply
    mutation.  The legacy raw-ID ``Draft`` reply read/add/set/remove
    methods remain retained mutation/identity compatibility and are
    intentionally neither required nor retired here.
    """

    if not _numbers_table_cell_comment_reply_owner_present(root):
        return []

    violations: list[str] = []
    owner_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_code = _mask_rust_non_code(owner_source)

    def location(path: Path, line_number: int) -> str:
        return f"{path.relative_to(root)}:{line_number}"

    def public_type_surface(source: str, name: str) -> tuple[str, int] | None:
        """Return one complete public type declaration, including its body."""

        code = _mask_rust_non_code(source)
        declaration = re.search(
            rf"(?m)^[ \t]*pub[ \t]+(?:struct|enum|type|union)[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b",
            code,
        )
        if declaration is None:
            return None
        line_number = source.count("\n", 0, declaration.start()) + 1
        opening = code.find("{", declaration.end())
        semicolon = code.find(";", declaration.end())
        if opening < 0 or (semicolon >= 0 and semicolon < opening):
            end = semicolon + 1 if semicolon >= 0 else len(code)
            return code[declaration.start() : end], line_number
        depth = 1
        cursor = opening + 1
        while cursor < len(code) and depth:
            if code[cursor] == "{":
                depth += 1
            elif code[cursor] == "}":
                depth -= 1
            cursor += 1
        end = cursor if depth == 0 else len(code)
        return code[declaration.start() : end], line_number

    def scan_public_surface(
        surface: str,
        path: Path,
        line_number: int,
        *,
        allowed_identifiers: frozenset[str] = frozenset(),
    ) -> None:
        for match in RUST_IDENTIFIER.finditer(surface):
            identifier = match.group(1)
            if identifier in allowed_identifiers:
                continue
            reason = _numbers_table_cell_comment_reply_public_leak(identifier)
            if reason is None:
                continue
            identifier_line = line_number + surface.count("\n", 0, match.start(1))
            violations.append(
                "focused litchi-numbers comment-reply public API exposes "
                f"{reason} {identifier}: {location(path, identifier_line)}"
            )
        for match in RUST_BYTE_SLICE.finditer(surface):
            byte_slice = re.sub(r"\s+", "", match.group(0))
            byte_slice_line = line_number + surface.count("\n", 0, match.start())
            violations.append(
                "focused litchi-numbers comment-reply public API exposes raw byte "
                f"slice {byte_slice}: {location(path, byte_slice_line)}"
            )

    canonical_type = NUMBERS_TABLE_CELL_COMMENT_REPLY_CANONICAL_TYPE
    canonical_exports = _rust_canonical_exports(
        owner_source, frozenset({canonical_type})
    )
    if canonical_type not in canonical_exports:
        violations.append(
            "focused litchi-numbers comment-reply public API is missing "
            f"canonical {canonical_type}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
        )
    type_surface = public_type_surface(owner_source, canonical_type)
    if type_surface is not None:
        scan_public_surface(type_surface[0], owner_path, type_surface[1])

    public_alias = NUMBERS_TABLE_CELL_COMMENT_REPLY_PUBLIC_ALIAS
    lib_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )

    # The text-only projection is the compatibility baseline.  Once the
    # focused owner starts publishing typed author/date metadata, keep that
    # expansion equally narrow: both text-bearing values carry the same
    # optional semantic fields, the value types are named at the crate root,
    # and the author debug representation cannot expose display data.  The
    # activation check is deliberately marker-based so old sources and the
    # existing text-only fixtures remain valid until the owner actually lands
    # the metadata types.
    metadata_activation = any(
        re.search(r"\b(?:CommentAuthor|CommentTimestamp)\b", source)
        is not None
        for source in (owner_code, _mask_rust_non_code(lib_source))
    )
    if metadata_activation:
        metadata_types = NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_TYPES
        metadata_exports = _rust_canonical_exports(owner_source, metadata_types)
        for metadata_type in sorted(metadata_types):
            if metadata_type not in metadata_exports:
                violations.append(
                    "focused litchi-numbers comment metadata is missing canonical "
                    f"type {metadata_type}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
                continue
            metadata_surface = public_type_surface(owner_source, metadata_type)
            if metadata_surface is None:
                violations.append(
                    "focused litchi-numbers comment metadata is missing readable "
                    f"type {metadata_type}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
                continue
            scan_public_surface(
                metadata_surface[0],
                owner_path,
                metadata_surface[1],
                allowed_identifiers=(
                    frozenset({"u64"})
                    if metadata_type == "CommentTimestamp"
                    else frozenset()
                ),
            )

        for value_type in ("Comment", canonical_type):
            value_surface = public_type_surface(owner_source, value_type)
            if value_surface is None:
                violations.append(
                    "focused litchi-numbers comment metadata is missing semantic "
                    f"value type {value_type}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
                continue
            value_code = _mask_rust_non_code(value_surface[0])
            scan_public_surface(value_surface[0], owner_path, value_surface[1])
            for field, field_type in (
                ("timestamp", "CommentTimestamp"),
                ("author", "CommentAuthor"),
            ):
                if re.search(
                    rf"\b{re.escape(field)}\b[ \t\r\n]*:[ \t\r\n]*"
                    rf"Option[ \t\r\n]*<[ \t\r\n]*"
                    rf"{re.escape(field_type)}[ \t\r\n]*>",
                    value_code,
                ) is None:
                    violations.append(
                        "focused litchi-numbers comment metadata value must retain "
                        f"Option<{field_type}> field {field} on {value_type}: "
                        f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                    )

        # Inspect every public inherent method on the metadata-bearing values
        # as well as the metadata types.  The required accessor checks below
        # prove the intended signatures; this pass prevents an additional
        # convenience method from reintroducing a native identity, archive,
        # or wire value through a return or parameter type.
        for value_type in (
            "Comment",
            canonical_type,
            *sorted(NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_TYPES),
        ):
            for _method, declaration, line_number in _rust_public_methods_in_impl(
                owner_source, value_type
            ):
                scan_public_surface(declaration, owner_path, line_number)

        metadata_root_exports = _rust_canonical_exports(
            lib_source,
            frozenset(NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_ROOT_ALIASES.values()),
        )
        for canonical_metadata, public_metadata in sorted(
            NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_ROOT_ALIASES.items()
        ):
            if public_metadata not in metadata_root_exports:
                violations.append(
                    "focused litchi-numbers comment metadata is missing named root "
                    f"alias {public_metadata}: "
                    f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]}"
                )
                continue
            if re.search(
                rf"\b{re.escape(canonical_metadata)}\b[ \t\r\n]+as[ \t\r\n]+"
                rf"{re.escape(public_metadata)}\b",
                lib_source,
            ) is None:
                violations.append(
                    "focused litchi-numbers comment metadata root alias must be a "
                    f"named re-export of {canonical_metadata}: "
                    f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]}"
                )

        required_metadata_methods = {
            "Comment": {
                "timestamp": r"->\s*Option\s*<\s*CommentTimestamp\s*>",
                "author": r"->\s*Option\s*<\s*&\s*CommentAuthor\s*>",
            },
            canonical_type: {
                "timestamp": r"->\s*Option\s*<\s*CommentTimestamp\s*>",
                "author": r"->\s*Option\s*<\s*&\s*CommentAuthor\s*>",
            },
            "CommentAuthor": {
                "display_name": r"->\s*Option\s*<\s*&\s*str\s*>",
                "public_id": r"->\s*Option\s*<\s*&\s*str\s*>",
            },
            "CommentTimestamp": {
                "new": r"\bseconds\s*:\s*f64\b[\s\S]*->\s*Option\s*<\s*Self\s*>",
                "as_f64": r"->\s*f64\b",
            },
        }
        for value_type, methods in required_metadata_methods.items():
            declarations = {
                name: declaration
                for name, declaration, _line_number in _rust_public_methods_in_impl(
                    owner_source, value_type
                )
            }
            for method, return_pattern in methods.items():
                declaration = declarations.get(method)
                if declaration is None or re.search(return_pattern, declaration) is None:
                    violations.append(
                        "focused litchi-numbers comment metadata is missing typed "
                        f"{value_type}::{method} accessor: "
                        f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                    )
                elif declaration:
                    scan_public_surface(
                        declaration,
                        owner_path,
                        next(
                            line_number
                            for name, _declaration, line_number in _rust_public_methods_in_impl(
                                owner_source, value_type
                            )
                            if name == method
                        ),
                    )

        author_derive_debug = re.search(
            r"#[ \t]*\[[ \t]*derive[ \t]*\([^]]*\bDebug\b[^]]*\)[ \t]*\]"
            r"(?:[ \t\r\n]*#[ \t]*\[[^]]*\])*[ \t\r\n]*pub[ \t]+"
            r"(?:struct|enum)[ \t]+(?:r#)?CommentAuthor\b",
            owner_source,
        )
        if author_derive_debug is not None:
            violations.append(
                "focused litchi-numbers CommentAuthor must use redacted custom "
                f"Debug: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )
        author_debug = NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_DEBUG_IMPL.search(
            owner_code
        )
        if author_debug is None:
            violations.append(
                "focused litchi-numbers CommentAuthor is missing redacted custom "
                f"Debug: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )
        else:
            opening = owner_code.find("{", author_debug.end())
            depth = 1
            cursor = opening + 1
            while opening >= 0 and cursor < len(owner_code) and depth:
                if owner_code[cursor] == "{":
                    depth += 1
                elif owner_code[cursor] == "}":
                    depth -= 1
                cursor += 1
            debug_body = owner_code[opening + 1 : cursor - 1] if opening >= 0 and not depth else ""
            debug_body_source = (
                owner_source[opening + 1 : cursor - 1]
                if opening >= 0 and not depth
                else ""
            )
            if (
                NUMBERS_TABLE_CELL_COMMENT_REPLY_DEBUG_REDACTION.search(
                    debug_body_source
                )
                is None
            ):
                violations.append(
                    "focused litchi-numbers CommentAuthor Debug must redact "
                    f"display metadata: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
            if re.search(
                r"\bself\s*\.\s*(?:display_name|public_id)\b", debug_body
            ) is not None:
                violations.append(
                    "focused litchi-numbers CommentAuthor Debug exposes display "
                    f"metadata: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )

    # ADR 0034: the value is published by its focused
    # `cell::comment` module under its canonical name; the prefixed root
    # alias is deleted.
    comment_path = root / NUMBERS_SOURCE_ROOT / "cell" / "comment.rs"
    comment_source = (
        _mask_rust_cfg_test_items(comment_path.read_text(encoding="utf-8"))
        if comment_path.is_file()
        else ""
    )
    if canonical_type not in _rust_canonical_exports(
        comment_source, frozenset({canonical_type})
    ):
        violations.append(
            "focused litchi-numbers comment-reply public API is missing "
            f"canonical {canonical_type} in cell::comment: "
            f"{comment_path.relative_to(root)}"
        )

    canonical_alias_pattern = re.compile(
        rf"\b{re.escape(canonical_type)}\b[ \t\r\n]+as[ \t\r\n]+"
        rf"{re.escape(public_alias)}\b"
    )
    for declaration, line_number in _rust_public_declarations(owner_source):
        identifiers = [
            match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
        ]
        public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
        canonical_type_declaration = re.match(
            rf"^[ \t]*pub[ \t]+(?:struct|enum|type|union)[ \t]+"
            rf"(?:r#)?{re.escape(canonical_type)}\b",
            declaration,
        )
        if (
            public_use_or_type
            and canonical_type in identifiers
            and canonical_type_declaration is None
        ):
            flat_alias = canonical_type
            if identifiers[:2] == ["pub", "type"] and len(identifiers) > 2:
                flat_alias = identifiers[2]
            violations.append(
                "focused litchi-numbers comment-reply public API retains flat "
                f"alias {flat_alias}: {location(owner_path, line_number)}"
            )
        if public_alias in identifiers:
            violations.append(
                "focused litchi-numbers comment-reply public API retains owner "
                f"alias {public_alias}: {location(owner_path, line_number)}"
            )
        for match in re.finditer(
            rf"\b{re.escape(canonical_type)}\b[ \t\r\n]+as[ \t\r\n]+"
            rf"(?P<alias>[A-Za-z_][A-Za-z0-9_]*)",
            declaration,
        ):
            if match.group("alias") != public_alias:
                violations.append(
                    "focused litchi-numbers comment-reply public API retains flat "
                    f"alias {match.group('alias')}: {location(owner_path, line_number)}"
                )
        if public_use_or_type and canonical_type in identifiers:
            scan_public_surface(declaration, owner_path, line_number)

    for declaration, line_number in _rust_public_declarations(lib_source):
        identifiers = [
            match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
        ]
        public_use_or_type = identifiers[:2] in (["pub", "use"], ["pub", "type"])
        if (
            public_use_or_type
            and "*" in declaration
            and (
                "comments" in identifiers
                or canonical_type in identifiers
                or public_alias in identifiers
            )
        ):
            violations.append(
                "focused litchi-numbers comment-reply public API retains root "
                f"aliases via glob: {location(lib_path, line_number)}"
            )
        if public_alias in identifiers and identifiers[:2] == ["pub", "type"]:
            violations.append(
                "focused litchi-numbers comment-reply canonical "
                f"{public_alias} must be a named re-export, not a type alias: "
                f"{location(lib_path, line_number)}"
            )
        if canonical_type in identifiers and identifiers[:2] == ["pub", "use"]:
            if canonical_alias_pattern.search(declaration) is None:
                violations.append(
                    "focused litchi-numbers comment-reply public API retains flat "
                    f"alias {canonical_type}: {location(lib_path, line_number)}"
                )
            for match in re.finditer(
                rf"\b{re.escape(canonical_type)}\b[ \t\r\n]+as[ \t\r\n]+"
                rf"(?P<alias>[A-Za-z_][A-Za-z0-9_]*)",
                declaration,
            ):
                if match.group("alias") != public_alias:
                    violations.append(
                        "focused litchi-numbers comment-reply public API retains flat "
                        f"alias {match.group('alias')}: {location(lib_path, line_number)}"
                    )
        if public_use_or_type and (
            canonical_type in identifiers or public_alias in identifiers
        ):
            scan_public_surface(declaration, lib_path, line_number)

    for name, declaration, line_number in _rust_public_methods_in_impl(
        owner_source, canonical_type
    ):
        scan_public_surface(declaration, owner_path, line_number)

    package_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for name, (declaration, line_number) in package_methods.items():
        if "comment_repl" in name or re.search(r"\bCommentReply\b", declaration):
            scan_public_surface(declaration, owner_path, line_number)
    for method in NUMBERS_TABLE_CELL_COMMENT_REPLY_PACKAGE_METHODS:
        method_record = package_methods.get(method)
        if method_record is None:
            violations.append(
                "focused litchi-numbers comment-reply public API is missing "
                f"Package method {method}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )
            continue
        declaration, line_number = method_record
        scan_public_surface(declaration, owner_path, line_number)
        for selector in ("SheetSelector", "TableSelector"):
            if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                violations.append(
                    "focused litchi-numbers comment-reply Package method "
                    f"{method} must accept selector-first {selector}: "
                    f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
        if method == "table_cell_comment_replies":
            if re.search(r"\bAt\b", declaration) is None:
                violations.append(
                    "focused litchi-numbers comment-reply Package method "
                    f"{method} must accept selector-first At: "
                    f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
        if re.search(r"\bCommentReply\b", declaration) is None:
            violations.append(
                "focused litchi-numbers comment-reply Package method "
                f"{method} must return CommentReply: "
                f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )
        if method == "table_cell_comment_replies_a1":
            if NUMBERS_TABLE_CELL_COMMENT_REPLY_A1_ADDRESS.search(declaration) is None:
                violations.append(
                    "focused litchi-numbers comment-reply Package method "
                    f"{method} must accept an A1 address: "
                    f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
            body = _rust_any_function_body(owner_source, method)
            if body is None:
                violations.append(
                    "focused litchi-numbers comment-reply Package method "
                    f"{method} has no readable routing body: "
                    f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                )
            else:
                body_code = _mask_rust_non_code(body)
                if re.search(
                    r"\btable_cell_comment_replies\b[ \t\r\n]*\(", body_code
                ) is None:
                    violations.append(
                        "focused litchi-numbers comment-reply A1 method must route "
                        "through table_cell_comment_replies: "
                        f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                    )
                if re.search(
                    r"\bAt[ \t\r\n]*::[ \t\r\n]*from_a1\b", body_code
                ) is None:
                    violations.append(
                        "focused litchi-numbers comment-reply A1 method must construct "
                        "At from A1: "
                        f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
                    )

    derive_debug = re.search(
        rf"#[ \t]*\[[ \t]*derive[ \t]*\([^]]*\bDebug\b[^]]*\)[ \t]*\]"
        rf"(?:[ \t\r\n]*#[ \t]*\[[^]]*\])*[ \t\r\n]*pub[ \t]+"
        rf"(?:struct|enum)[ \t]+(?:r#)?{re.escape(canonical_type)}\b",
        owner_source,
    )
    if derive_debug is not None:
        violations.append(
            "focused litchi-numbers comment-reply CommentReply must not derive "
            f"Debug: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}:"
            f"{owner_source.count(chr(10), 0, derive_debug.start()) + 1}"
        )

    debug_match = NUMBERS_TABLE_CELL_COMMENT_REPLY_DEBUG_IMPL.search(owner_code)
    if debug_match is None:
        violations.append(
            "focused litchi-numbers comment-reply CommentReply is missing custom "
            f"Debug implementation: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
        )
    else:
        opening = owner_code.find("{", debug_match.end())
        depth = 1
        cursor = opening + 1
        while opening >= 0 and cursor < len(owner_code) and depth:
            if owner_code[cursor] == "{":
                depth += 1
            elif owner_code[cursor] == "}":
                depth -= 1
            cursor += 1
        if opening < 0 or depth:
            debug_body = ""
        else:
            debug_body = owner_source[opening + 1 : cursor - 1]
        debug_body_code = _mask_rust_non_code(debug_body)
        if NUMBERS_TABLE_CELL_COMMENT_REPLY_DEBUG_REDACTION.search(debug_body) is None:
            violations.append(
                "focused litchi-numbers comment-reply custom Debug must redact "
                f"authored content: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )
        if re.search(
            r"\bself[ \t\r\n]*\.[ \t\r\n]*(?:text|content|body|message)\b",
            debug_body_code,
        ) is not None:
            violations.append(
                "focused litchi-numbers comment-reply custom Debug exposes authored "
                f"content: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )

    graph_paths: set[Path] = {owner_path}
    extractor_path = root / NUMBERS_EXTRACTOR_SOURCE
    if extractor_path.is_file():
        graph_paths.add(extractor_path)
    graph_paths.update(
        path
        for path in owner_path.parent.glob("comments*.rs")
        if path.is_file()
    )
    owner_graph_code = "\n".join(
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for path in sorted(graph_paths)
    )
    codec_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE
    if not codec_path.is_file():
        violations.append(
            "focused litchi-numbers comment-reply public API is missing strict "
            f"comment_storage_codec source: {NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE}"
        )
        codec_code = ""
    else:
        codec_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        )
    for label, marker in NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REQUIRED_MARKERS.items():
        if marker.search(codec_code) is None:
            violations.append(
                "focused litchi-numbers comment-reply owner is missing "
                f"{label} marker: {NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )
    if re.search(
        r"\bdecode_comment_storage_archive_with_visitor\b", owner_graph_code
    ) is None:
        violations.append(
            "focused litchi-numbers comment-reply owner does not call the strict "
            "comment-storage codec visitor: "
            f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
        )
    for label, marker in NUMBERS_TABLE_CELL_COMMENT_REPLY_CENSUS_REQUIRED_MARKERS.items():
        if marker.search(owner_graph_code) is None:
            violations.append(
                "focused litchi-numbers comment-reply owner is missing full "
                f"ownership-census {label} marker: "
                f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_SOURCE}"
            )

    return sorted(set(violations))


def audit_numbers_table_cell_comment_reply_codec_rewrite_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Ratchet the prepared reply-rewrite codec independently of the owner.

    The reply read facade intentionally remains a read-only projection.  This
    separate ratchet is dormant until a production (non-``cfg(test)``)
    ``CommentStorageReplyRewrite`` symbol appears in the strict codec.  Once
    that symbol lands, the codec must expose the complete prepared
    prepare/report/execute contract, document strict raw preservation and
    bounded resource accounting, and carry direct unit coverage.  It does not
    activate or require a Numbers package mutation owner, host migration, or
    fuzz target.
    """

    codec_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE
    if not codec_path.is_file():
        return []

    raw_source = codec_path.read_text(encoding="utf-8")
    production_source = _mask_rust_cfg_test_items(raw_source)
    production_code = _mask_rust_non_code(production_source)
    if (
        NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_ACTIVATION.search(
            production_code
        )
        is None
    ):
        return []

    violations: list[str] = []
    for api in NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_REQUIRED_APIS:
        if re.search(
            rf"\b(?:pub(?:\([^()]*\))?[ \t]+)?"
            rf"(?:fn|struct|enum|type|trait)[ \t]+{re.escape(api)}\b",
            production_code,
        ) is None:
            violations.append(
                "focused Numbers comment-storage codec reply rewrite is missing "
                f"strict API {api}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE}"
            )

    for label, marker in (
        NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_REQUIRED_MARKERS.items()
    ):
        if marker.search(production_source) is None:
            violations.append(
                "focused Numbers comment-storage codec reply rewrite is missing "
                f"{label} marker: {NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE}"
            )

    # Recover only the text masked by the item-level cfg(test) scanner.  This
    # lets the direct-coverage ratchet inspect unit tests without allowing a
    # test-only symbol or decoy to activate the production contract above.
    gated_source = "".join(
        character if character != masked else " "
        for character, masked in zip(raw_source, production_source)
    )
    test_code = _mask_rust_non_code(gated_source)
    for label, marker in (
        NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_TEST_MARKERS.items()
    ):
        if marker.search(test_code) is None:
            violations.append(
                "focused Numbers comment-storage codec reply rewrite is missing "
                f"{label}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_SOURCE}"
            )

    return sorted(set(violations))


def _numbers_table_cell_comment_reply_mutation_owner_present(root: Path) -> bool:
    """Return whether reply mutation is present in production source."""

    source_paths = [
        root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE,
        root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE,
        root / NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0],
    ]
    owner_helper_root = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        source_paths.extend(owner_helper_root.glob("comments*.rs"))
    for path in source_paths:
        if not path.is_file():
            continue
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        if NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ACTIVATION.search(
            _mask_rust_non_code(production_source)
        ) is not None:
            return True
    return False


def audit_numbers_table_cell_comment_reply_mutation_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first reply mutation boundary.

    The plural reply read projection is intentionally independent from this
    ratchet.  Once a production mutation symbol appears, this check requires
    the complete semantic/transaction surface, a private owner, strict
    prepared codec usage, package-wide ownership and metadata proofs, and the
    two fuzz inventories.  All checks are source-only and cfg(test)-items are
    masked before activation or leak inspection.
    """

    if not _numbers_table_cell_comment_reply_mutation_owner_present(root):
        return []

    violations: list[str] = []

    def location(path: Path, line_number: int) -> str:
        return f"{path.relative_to(root)}:{line_number}"

    def read(path: Path) -> str:
        return path.read_text(encoding="utf-8") if path.is_file() else ""

    def production(path: Path) -> str:
        return _mask_rust_cfg_test_items(read(path))

    def code(path: Path) -> str:
        return _mask_rust_non_code(production(path))

    semantic_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE
    semantic_source = production(semantic_path)
    semantic_code = _mask_rust_non_code(semantic_source)
    if not semantic_path.is_file():
        violations.append(
            "focused Numbers comment-reply mutation public API is missing "
            f"semantic cell::comment source: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE}"
        )
    if re.search(r"\bpub\s+mod\s+transaction\b", semantic_code) is None:
        violations.append(
            "focused Numbers comment-reply mutation public API is missing "
            "cell::comment::transaction module: "
            f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE}"
        )
    cell_mod_path = root / NUMBERS_SOURCE_ROOT / "cell" / "mod.rs"
    if re.search(r"\bpub\s+mod\s+comment\b", code(cell_mod_path)) is None:
        violations.append(
            "focused Numbers comment-reply mutation public API is missing "
            f"cell::comment export: {cell_mod_path.relative_to(root)}"
        )
    transaction_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TRANSACTION_SOURCE
    transaction_source = production(transaction_path)
    transaction_code = _mask_rust_non_code(transaction_source)
    if not transaction_path.is_file() and re.search(
        r"\bpub\s+mod\s+transaction\s*\{", semantic_code
    ) is None:
        violations.append(
            "focused Numbers comment-reply mutation public API is missing "
            f"transaction source: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TRANSACTION_SOURCE}"
        )

    owner_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE
    owner_paths: set[Path] = set()
    for candidate in (
        owner_path,
        root / NUMBERS_SOURCE_ROOT / "package" / "table_cell_comment_reply.rs",
    ):
        if candidate.is_file():
            owner_paths.add(candidate)
    owner_parent = owner_path.parent
    if owner_parent.is_dir():
        owner_paths.update(
            path
            for path in owner_parent.glob("comment*.rs")
            if path.is_file()
        )
        owner_paths.update(
            path
            for path in owner_parent.glob("table_cell_comment_reply*.rs")
            if path.is_file()
        )
    helper_root = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        owner_paths.update(
            path for path in helper_root.glob("comments_reply*.rs") if path.is_file()
        )
    # The Numbers reply transaction delegates its native
    # transition to the shared copy-on-write reply engine, which owns the
    # prepared reply-reference codec call.
    shared_engine_root = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_SHARED_ENGINE_ROOT
    if shared_engine_root.is_dir():
        owner_paths.update(
            path for path in shared_engine_root.glob("*.rs") if path.is_file()
        )
    owner_sources = {
        path: production(path) for path in sorted(owner_paths) if path.is_file()
    }
    owner_graph_source = "\n".join(owner_sources.values())
    owner_graph_code = _mask_rust_non_code(owner_graph_source)
    package_path = root / NUMBERS_SOURCE_ROOT / "package.rs"
    if re.search(r"\b(?:pub\s*\([^)]*\)\s*)?mod\s+comments\b", code(package_path)) is None:
        violations.append(
            "focused Numbers comment-reply mutation public API is missing private "
            f"package::comments owner module: {package_path.relative_to(root)}"
        )
    if not owner_paths:
        violations.append(
            "focused Numbers comment-reply mutation public API is missing private "
            f"owner source: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
        )

    semantic_exports = _rust_canonical_exports(
        semantic_source,
        frozenset({"CommentReply"}),
    )
    if "CommentReply" not in semantic_exports:
        violations.append(
            "focused Numbers comment-reply mutation public API is missing canonical "
            f"cell::comment CommentReply: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        transaction_source or semantic_source,
        frozenset(NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TYPES[1:]),
    )
    for name in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TYPES:
        if name == "CommentReplyIndex":
            exports = _rust_canonical_exports(
                semantic_source + "\n" + transaction_source,
                frozenset({name}),
            )
        else:
            exports = transaction_exports
        if name not in exports:
            violations.append(
                "focused Numbers comment-reply mutation public API is missing "
                f"canonical transaction type {name}: "
                f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TRANSACTION_SOURCE}"
            )

    lib_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]
    lib_source = production(lib_path)
    lib_code = _mask_rust_non_code(lib_source)
    if re.search(r"\bpub\s+mod\s+cell\b", lib_code) is None:
        violations.append(
            "focused Numbers comment-reply mutation public API is missing root cell module: "
            f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]}"
        )
    root_exports = _rust_canonical_exports(
        lib_source, frozenset(NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ROOT_ALIASES)
    )
    for alias in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ROOT_ALIASES:
        if alias not in root_exports:
            violations.append(
                "focused Numbers comment-reply mutation public API is missing named root alias "
                f"{alias}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_EXPORT_SOURCES[0]}"
            )

    method_records: dict[str, tuple[str, Path, int]] = {}
    for path, source in owner_sources.items():
        for name, declaration, line_number in _rust_public_methods_in_impl(
            source, "Package"
        ):
            method_records.setdefault(name, (declaration, path, line_number))
    method_result_names = {
        "table_cell_comment_reply": ("CommentReply",),
        "table_cell_comment_reply_a1": ("CommentReply",),
        "edit_table_cell_comment_replies": ("CommentReplyEdit", "Edit"),
        "edit_table_cell_comment_replies_a1": ("CommentReplyEdit", "Edit"),
        "set_table_cell_comment_reply": ("CommentReplyCommit", "Commit"),
        "set_table_cell_comment_reply_a1": ("CommentReplyCommit", "Commit"),
        "remove_table_cell_comment_reply": ("CommentReplyCommit", "Commit"),
        "remove_table_cell_comment_reply_a1": ("CommentReplyCommit", "Commit"),
        "add_table_cell_comment_reply": ("CommentReplyCommit", "Commit"),
        "add_table_cell_comment_reply_a1": ("CommentReplyCommit", "Commit"),
        "apply_table_cell_comment_reply": ("CommentReplyCommit", "Commit"),
    }
    for method in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_PACKAGE_METHODS:
        record = method_records.get(method)
        if record is None:
            violations.append(
                "focused Numbers comment-reply mutation public API is missing "
                f"Package method {method}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
            )
            continue
        declaration, path, line_number = record
        for selector in ("SheetSelector", "TableSelector"):
            if re.search(rf"\b{re.escape(selector)}\b", declaration) is None and method != "apply_table_cell_comment_reply":
                violations.append(
                    "focused Numbers comment-reply mutation Package method "
                    f"{method} must accept selector-first {selector}: {location(path, line_number)}"
                )
        if (
            method != "apply_table_cell_comment_reply"
            and not method.endswith("_a1")
            and re.search(r"\bAt\b", declaration) is None
        ):
            violations.append(
                "focused Numbers comment-reply mutation Package method "
                f"{method} must accept selector-first At: {location(path, line_number)}"
            )
        if method in {"table_cell_comment_reply", "set_table_cell_comment_reply", "remove_table_cell_comment_reply"} and re.search(
            r"\bCommentReplyIndex\b", declaration
        ) is None:
            violations.append(
                "focused Numbers comment-reply mutation Package method "
                f"{method} must accept CommentReplyIndex: {location(path, line_number)}"
            )
        if method.endswith("_a1"):
            if NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_A1_ADDRESS.search(declaration) is None:
                violations.append(
                    "focused Numbers comment-reply mutation A1 method "
                    f"{method} must accept an A1 address: {location(path, line_number)}"
                )
            body = _rust_any_function_body(owner_sources[path], method)
            if body is None or re.search(r"from_a1", _mask_rust_non_code(body)) is None:
                violations.append(
                    "focused Numbers comment-reply mutation A1 method "
                    f"{method} must construct At from A1: {location(path, line_number)}"
                )
        if not any(re.search(rf"\b{re.escape(name)}\b", declaration) for name in method_result_names[method]):
            violations.append(
                "focused Numbers comment-reply mutation Package method "
                f"{method} must expose its typed transaction result: {location(path, line_number)}"
            )

    edit_methods: dict[str, tuple[str, Path, int]] = {}
    for path, source in owner_sources.items():
        for impl_name in ("Edit", "CommentReplyEdit"):
            for name, declaration, line_number in _rust_public_methods_in_impl(
                source, impl_name
            ):
                edit_methods.setdefault(name, (declaration, path, line_number))
    for name in ("set", "remove", "commit"):
        if name not in edit_methods:
            violations.append(
                "focused Numbers comment-reply mutation Edit is missing "
                f"{name} operation: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
            )

    index_methods: dict[str, tuple[str, Path, int]] = {}
    for path, source in {
        **owner_sources,
        semantic_path: semantic_source,
        transaction_path: transaction_source,
    }.items():
        for name, declaration, line_number in _rust_public_methods_in_impl(
            source, "CommentReplyIndex"
        ):
            index_methods.setdefault(name, (declaration, path, line_number))
    for name in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_INDEX_METHODS:
        if name not in index_methods:
            violations.append(
                "focused Numbers comment-reply mutation CommentReplyIndex is missing "
                f"{name} constructor/accessor: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_SEMANTIC_SOURCE}"
            )
        elif name == "try_from_usize" and re.search(
            r"->[ 	\r\n]*(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ 	\r\n]*::[ 	\r\n]*)*"
            r"(?:r#)?Result[ 	\r\n]*<",
            index_methods[name][0],
        ) is None:
            declaration, path, line_number = index_methods[name]
            violations.append(
                "focused Numbers comment-reply mutation CommentReplyIndex "
                f"{name} must return Result: {location(path, line_number)}"
            )

    # Scan the public semantic/transaction/owner/export surface for native
    # vocabulary, raw bytes, flat aliases, and unredacted text-bearing Debug.
    dedicated_paths = set(owner_sources)
    dedicated_paths.update(
        path
        for path in (semantic_path, transaction_path)
        if path.is_file()
    )
    canonical_public_names = set(NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_TYPES)
    canonical_public_names.add("CommentReply")

    def public_type_surface(source: str, name: str) -> tuple[str, int] | None:
        masked = _mask_rust_non_code(source)
        declaration = re.search(
            rf"(?m)^[ \t]*pub[ \t]+(?:struct|enum|type|union)[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b",
            masked,
        )
        if declaration is None:
            return None
        opening = masked.find("{", declaration.end())
        semicolon = masked.find(";", declaration.end())
        if opening < 0 or (semicolon >= 0 and semicolon < opening):
            end = semicolon + 1 if semicolon >= 0 else len(masked)
            return masked[declaration.start() : end], source.count("\n", 0, declaration.start()) + 1
        depth = 1
        cursor = opening + 1
        while cursor < len(masked) and depth:
            if masked[cursor] == "{":
                depth += 1
            elif masked[cursor] == "}":
                depth -= 1
            cursor += 1
        return masked[declaration.start() : cursor], source.count("\n", 0, declaration.start()) + 1

    def scan_surface(
        source: str,
        path: Path,
        surface: str,
        line_number: int,
        *,
        allow_usize: bool = False,
    ) -> None:
        for match in RUST_IDENTIFIER.finditer(surface):
            identifier = match.group(1)
            if allow_usize and identifier == "usize":
                continue
            reason = _numbers_table_cell_comment_reply_mutation_public_leak(identifier)
            if reason is not None:
                violations.append(
                    "focused Numbers comment-reply mutation public API exposes "
                    f"{reason} {identifier}: {location(path, line_number + surface.count(chr(10), 0, match.start(1)))}"
                )
        for match in RUST_BYTE_SLICE.finditer(surface):
            byte_slice = re.sub(r"\s+", "", match.group(0))
            violations.append(
                "focused Numbers comment-reply mutation public API exposes raw byte slice "
                f"{byte_slice}: "
                f"{location(path, line_number + surface.count(chr(10), 0, match.start()))}"
            )

    for path in sorted(dedicated_paths):
        source = production(path)
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:2] == ["pub", "use"] and "*" in declaration and any(
                name in declaration
                for name in ("comment", "CommentReply", "CommentReplyEdit")
            ):
                violations.append(
                    "focused Numbers comment-reply mutation public API retains "
                    f"comment transaction glob: {location(path, line_number)}"
                )
            if identifiers[:2] == ["pub", "type"]:
                alias = identifiers[2] if len(identifiers) > 2 else "unknown"
                targets = set(identifiers[3:])
                if alias not in canonical_public_names and targets & (
                    canonical_public_names | {"TableCellCommentReply"}
                ):
                    violations.append(
                        "focused Numbers comment-reply mutation public API retains flat alias "
                        f"{alias}: {location(path, line_number)}"
                    )
                if path == lib_path and alias in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ROOT_ALIASES:
                    violations.append(
                        "focused Numbers comment-reply mutation root aliases must be named "
                        f"re-exports, not type aliases: {location(path, line_number)}"
                    )
        for name in canonical_public_names:
            # The reply family's short names are defined by the
            # reply owner; the sibling comment family reuses the same names.
            if (
                name in {"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"}
                and not path.name.startswith("comments_reply")
            ):
                continue
            surface = public_type_surface(source, name)
            if surface is None:
                continue
            # Platform-sized counters retained behind a canonical semantic
            # value are not raw-ID/coordinate method leaks. Public Package
            # and index method declarations are scanned separately below,
            # where `usize` remains forbidden except for the checked
            # `try_from_usize` constructor.
            scan_surface(surface[0], path, surface[0], surface[1], allow_usize=True)
            if re.search(
                rf"(?m)^[ \t]*#\s*\[\s*derive\s*\([^]]*\bDebug\b[^]]*\)\s*\]\s*\n"
                rf"[ \t]*pub[ \t]+(?:struct|enum|type|union)[ \t]+(?:r#)?{re.escape(name)}\b",
                source,
            ) is not None and NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_DEBUG_SENSITIVE.search(
                surface[0]
            ) is not None:
                violations.append(
                    "focused Numbers comment-reply mutation text-bearing type must not derive "
                    f"Debug: {location(path, surface[1])}"
                )

    # The crate root and package module contain unrelated public APIs with
    # native implementation vocabulary.  Inspect only declarations that
    # explicitly participate in this comment-reply export, rather than
    # treating those legacy surfaces as part of the new facade.
    for path, source in ((lib_path, lib_source),):
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            relevant = any(
                alias in identifiers
                for alias in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ROOT_ALIASES
            ) or (
                identifiers[:2] in (["pub", "use"], ["pub", "type"])
                and any(
                    name in declaration
                    for name in ("comment", "CommentReply", "TableCellCommentReply")
                )
            )
            if not relevant:
                continue
            scan_surface(declaration, path, declaration, line_number)
            if identifiers[:2] == ["pub", "use"] and "*" in declaration and any(
                name in declaration
                for name in ("comment", "CommentReply", "CommentReplyEdit")
            ):
                violations.append(
                    "focused Numbers comment-reply mutation public API retains "
                    f"comment transaction glob: {location(path, line_number)}"
                )
            if identifiers[:2] == ["pub", "type"]:
                alias = identifiers[2] if len(identifiers) > 2 else "unknown"
                if alias in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_ROOT_ALIASES:
                    violations.append(
                        "focused Numbers comment-reply mutation root aliases must be named "
                        f"re-exports, not type aliases: {location(path, line_number)}"
                    )

    for path, source in owner_sources.items():
        for name, declaration, line_number in _rust_public_methods_in_impl(source, "Package"):
            scan_surface(declaration, path, declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            source, "CommentReplyIndex"
        ):
            scan_surface(
                declaration,
                path,
                declaration,
                line_number,
                allow_usize=name == "try_from_usize",
            )
    for source_path in (semantic_path, transaction_path):
        source = semantic_source if source_path == semantic_path else transaction_source
        if not source:
            continue
        for name, declaration, line_number in _rust_public_methods_in_impl(
            source, "CommentReplyIndex"
        ):
            scan_surface(
                declaration,
                source_path,
                declaration,
                line_number,
                allow_usize=name == "try_from_usize",
            )

    for label, marker in NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_graph_code) is None:
            violations.append(
                "focused Numbers comment-reply mutation owner is missing "
                f"{label} transaction marker: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
            )

    codec_path = root / NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_SOURCE
    codec_source = production(codec_path)
    codec_code = _mask_rust_non_code(codec_source)
    if not codec_path.is_file():
        violations.append(
            "focused Numbers comment-reply mutation owner is missing strict codec source: "
            f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_SOURCE}"
        )
    else:
        for api in NUMBERS_TABLE_CELL_COMMENT_REPLY_CODEC_REWRITE_REQUIRED_APIS:
            if re.search(
                rf"\b(?:pub(?:\([^()]*\))?[ \t]+)?(?:fn|struct|enum|type|trait)"
                rf"[ \t]+{re.escape(api)}\b",
                codec_code,
            ) is None:
                violations.append(
                    "focused Numbers comment-reply mutation strict codec is missing API "
                    f"{api}: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_SOURCE}"
                )
        if re.search(r"\bprepare_comment_storage_reply_rewrite\s*\(", owner_graph_code) is None:
            violations.append(
                "focused Numbers comment-reply mutation owner does not call the prepared "
                f"codec: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
            )

    fuzz_contracts = (
        (
            NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_FUZZ_SOURCE,
            "codec",
            ("prepare_comment_storage_reply_rewrite", "execute", "unknown", "inverse"),
        ),
        (
            NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_FUZZ_SOURCE,
            "package",
            ("add_table_cell_comment_reply", "set_table_cell_comment_reply", "remove_table_cell_comment_reply", "CommentReplyIndex", "inverse", "conflict", "atomic", "limit"),
        ),
    )
    for fuzz_path, kind, markers in fuzz_contracts:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused Numbers comment-reply mutation boundary is missing "
                f"{kind} fuzz target: {fuzz_path}"
            )
        else:
            fuzz_source = _mask_rust_non_code(absolute.read_text(encoding="utf-8"))
            if "fuzz_target!" not in fuzz_source:
                violations.append(
                    "focused Numbers comment-reply mutation "
                    f"{kind} fuzz target is missing fuzz_target! harness: {fuzz_path}"
                )
            for marker in markers:
                if marker not in fuzz_source.lower() and marker.lower() not in fuzz_source.lower():
                    violations.append(
                        "focused Numbers comment-reply mutation "
                        f"{kind} fuzz target is missing {marker}: {fuzz_path}"
                    )
    for corpus, kind in (
        (NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_CODEC_FUZZ_CORPUS, "codec"),
        (NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_FUZZ_CORPUS, "package"),
    ):
        absolute = root / corpus
        if not absolute.is_dir():
            violations.append(
                "focused Numbers comment-reply mutation boundary is missing "
                f"{kind} fuzz corpus: {corpus}"
            )
        elif not any(path.is_file() for path in absolute.rglob("*")):
            violations.append(
                "focused Numbers comment-reply mutation "
                f"{kind} fuzz corpus is empty: {corpus}"
            )

    debug_source = owner_graph_source + "\n" + semantic_source
    debug_code = _mask_rust_non_code(debug_source)
    debug_match = NUMBERS_TABLE_CELL_COMMENT_REPLY_DEBUG_IMPL.search(debug_code)
    if debug_match is None:
        violations.append(
            "focused Numbers comment-reply mutation CommentReply is missing custom "
            f"Debug redaction: {NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
        )
    else:
        opening = debug_code.find("{", debug_match.end())
        cursor = opening + 1 if opening >= 0 else -1
        depth = 1 if opening >= 0 else 0
        while cursor >= 0 and cursor < len(debug_code) and depth:
            if debug_code[cursor] == "{":
                depth += 1
            elif debug_code[cursor] == "}":
                depth -= 1
            cursor += 1
        if cursor < 0 or depth:
            cursor = len(debug_code)
        debug_region = debug_source[debug_match.start() : cursor]
        debug_region_code = debug_code[debug_match.start() : cursor]
        if NUMBERS_TABLE_CELL_COMMENT_REPLY_DEBUG_REDACTION.search(
            debug_region
        ) is None:
            violations.append(
                "focused Numbers comment-reply mutation Debug must redact authored content: "
                f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
            )
        else:
            # A typed CommentAuthor has its own redacted Debug implementation.
            # Once that metadata value is present, including the author value
            # in the CommentReply debug struct is safe; authored text and all
            # other content-bearing fields remain forbidden.  Keep the old
            # ``author`` rejection for text-only/synthetic mutation owners.
            author_debug_present = (
                NUMBERS_TABLE_CELL_COMMENT_REPLY_METADATA_DEBUG_IMPL.search(
                    debug_code
                )
                is not None
            )
            forbidden_fields = "text|content|body|message"
            if not author_debug_present:
                forbidden_fields += "|author"
            if re.search(
                rf"\bself\s*\.\s*(?:{forbidden_fields})\b",
                debug_region_code,
            ) is not None:
                violations.append(
                    "focused Numbers comment-reply mutation Debug exposes authored content: "
                    f"{NUMBERS_TABLE_CELL_COMMENT_REPLY_MUTATION_OWNER_SOURCE}"
                )

    return sorted(set(violations))


def audit_numbers_comment_clear_metadata_prerequisite_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep comment clears gated on the bounded metadata primitives.

    A changed comment clear remains unsupported until its owner reaches one
    atomic metadata publication that removes the exact comment-graph
    registrations and advances the selected current-component save tokens.
    Once that combined seam is reachable, admission is intentionally narrow:
    the route must retain root-only, uniqueness/reply, and global-ownership
    checks, and its semantic public boundary must not expose native IDs, wire
    types, or generated-Prost materialization.  The reject-all route remains a
    valid state while that implementation is absent.
    """

    violations: list[str] = []
    codec_path = root / PACKAGE_METADATA_CODEC_SOURCE
    if not codec_path.is_file():
        return [
            "focused Numbers comment-clear metadata prerequisite is missing "
            f"codec source: {PACKAGE_METADATA_CODEC_SOURCE}"
        ]

    raw_codec = codec_path.read_text(encoding="utf-8")
    codec_source = _mask_rust_cfg_test_items(raw_codec)
    codec_code = _mask_rust_non_code(codec_source)

    for label, pattern in PACKAGE_METADATA_CODEC_NO_EAGER_PATTERNS:
        for match in pattern.finditer(codec_code):
            line_number = codec_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused PackageMetadata removal/save-token codec uses "
                f"{label}: {PACKAGE_METADATA_CODEC_SOURCE}:{line_number}"
            )

    def public_function_signature(name: str) -> str | None:
        match = re.search(
            rf"(?<![A-Za-z0-9_#])pub(?:[ \t\r\n]+\([^()]*\))?"
            rf"[ \t\r\n]+fn[ \t\r\n]+{re.escape(name)}\b",
            codec_code,
        )
        if match is None:
            return None
        opening = codec_code.find("{", match.end())
        if opening < 0:
            return codec_code[match.start() : match.end()]
        return codec_code[match.start() : opening]

    required_functions = (
        (
            PACKAGE_METADATA_REMOVE_FUNCTION,
            PACKAGE_METADATA_REMOVE_BATCH,
            "removal",
        ),
        (
            PACKAGE_METADATA_SAVE_TOKEN_FUNCTION,
            PACKAGE_METADATA_SAVE_TOKEN_BATCH,
            "save-token",
        ),
        (
            PACKAGE_METADATA_COMBINED_FUNCTION,
            PACKAGE_METADATA_COMBINED_BATCH,
            "combined removal/save-token",
        ),
    )
    for function_name, batch_name, description in required_functions:
        signature = public_function_signature(function_name)
        if signature is None:
            violations.append(
                "focused PackageMetadata codec is missing typed "
                f"{description} API {function_name}: {PACKAGE_METADATA_CODEC_SOURCE}"
            )
            continue
        if not re.search(rf"\b{re.escape(batch_name)}\b", signature):
            line_number = codec_code.count("\n", 0, codec_code.find(signature)) + 1
            violations.append(
                "focused PackageMetadata codec "
                f"{function_name} must accept {batch_name}: "
                f"{PACKAGE_METADATA_CODEC_SOURCE}:{line_number}"
            )

    for type_name in (
        PACKAGE_METADATA_SELECTOR,
        PACKAGE_METADATA_REMOVE_BATCH,
        PACKAGE_METADATA_SAVE_TOKEN_BATCH,
        PACKAGE_METADATA_COMBINED_BATCH,
    ):
        if re.search(
            rf"(?<![A-Za-z0-9_#])(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
            rf"(?:struct|enum|type)[ \t\r\n]+{re.escape(type_name)}\b",
            codec_code,
        ) is not None:
            continue
        violations.append(
            "focused PackageMetadata codec is missing typed "
            f"{type_name}: {PACKAGE_METADATA_CODEC_SOURCE}"
        )

    combined_batch = re.search(
        rf"(?<![A-Za-z0-9_#])pub[ \t\r\n]+struct[ \t\r\n]+"
        rf"{re.escape(PACKAGE_METADATA_COMBINED_BATCH)}\b[\s\S]*?\}}",
        codec_code,
    )
    if combined_batch is None or not all(
        re.search(rf"\b{re.escape(batch_name)}\b", combined_batch.group(0))
        for batch_name in (
            PACKAGE_METADATA_REMOVE_BATCH,
            PACKAGE_METADATA_SAVE_TOKEN_BATCH,
        )
    ):
        violations.append(
            "focused PackageMetadata combined batch must compose typed removal "
            f"and save-token batches: {PACKAGE_METADATA_CODEC_SOURCE}"
        )

    comment_path = root / NUMBERS_COMMENT_SOURCE
    if not comment_path.is_file():
        return sorted(set(violations))
    raw_comments = comment_path.read_text(encoding="utf-8")
    comment_source = _mask_rust_cfg_test_items(raw_comments)
    comment_code = _mask_rust_non_code(comment_source)

    function_bodies: dict[str, str] = {}
    for declaration in RUST_FUNCTION_DECLARATION.finditer(comment_code):
        opening = comment_code.find("{", declaration.end())
        if opening < 0:
            continue
        depth = 1
        cursor = opening + 1
        while cursor < len(comment_code) and depth:
            if comment_code[cursor] == "{":
                depth += 1
            elif comment_code[cursor] == "}":
                depth -= 1
            cursor += 1
        if depth == 0:
            function_bodies.setdefault(
                declaration.group(1), comment_code[opening + 1 : cursor - 1]
            )

    def reachable_bodies(function_name: str) -> list[tuple[str, str]]:
        """Return the local helper bodies reachable from one owner function."""

        reachable: list[tuple[str, str]] = []
        pending = [function_name]
        visited: set[str] = set()
        while pending:
            current = pending.pop()
            if current in visited:
                continue
            body = function_bodies.get(current)
            if body is None:
                continue
            visited.add(current)
            reachable.append((current, body))
            for call in PACKAGE_METADATA_BARE_CALL.finditer(body):
                called = call.group("name").removeprefix("r#")
                if called in function_bodies and called not in visited:
                    pending.append(called)
        return reachable

    clear_body = function_bodies.get(NUMBERS_COMMENT_CLEAR_FUNCTION)
    if clear_body is None:
        violations.append(
            "focused Numbers comment owner is missing production "
            f"{NUMBERS_COMMENT_CLEAR_FUNCTION}: {NUMBERS_COMMENT_SOURCE}"
        )
    else:
        if re.search(r"\.(?:clear)[ \t\r\n]*\(", clear_body) is None:
            violations.append(
                "focused Numbers comment clear must retain its typed staging "
                f"clear path: {NUMBERS_COMMENT_SOURCE}"
            )

    prepared_markers = (
        r"\bprepare_metadata_token_entry[ \t\r\n]*\(",
        rf"\b{re.escape(PACKAGE_METADATA_SAVE_TOKEN_FUNCTION)}[ \t\r\n]*\(",
        r"\bmetadata_entry\b",
        r"\bEntryEdit[ \t\r\n]*::[ \t\r\n]*new[ \t\r\n]*\(",
    )
    commit_body = function_bodies.get(NUMBERS_COMMENT_COMMIT_FUNCTION)
    if commit_body is None:
        violations.append(
            "focused Numbers comment owner is missing production "
            f"{NUMBERS_COMMENT_COMMIT_FUNCTION}: {NUMBERS_COMMENT_SOURCE}"
        )
    else:
        reachable = reachable_bodies(NUMBERS_COMMENT_COMMIT_FUNCTION)
        reachable_code = "\n".join(body for _name, body in reachable)
        combined_call = re.search(
            rf"\b{re.escape(PACKAGE_METADATA_COMBINED_FUNCTION)}"
            r"[ \t\r\n]*\(",
            reachable_code,
        )
        prepared_route = re.search(r"\bprepare_root_clear[ \t\r\n]*\(", reachable_code)
        prepared_route_complete = prepared_route is not None and all(
            re.search(marker, reachable_code) is not None for marker in prepared_markers
        )
        if prepared_route is not None and not prepared_route_complete:
            violations.append(
                "focused Numbers prepared comment clear must rewrite and publish "
                f"the metadata token artifact: {NUMBERS_COMMENT_SOURCE}"
            )
        has_combined_route = combined_call is not None or prepared_route_complete

        unsupported = commit_body.find(NUMBERS_COMMENT_CLEAR_UNSUPPORTED_ERROR)
        rewrite = commit_body.find("rewrite_existing")
        if not has_combined_route and unsupported < 0:
            violations.append(
                "focused Numbers comment changed clear must fail with typed "
                f"{NUMBERS_COMMENT_CLEAR_UNSUPPORTED_ERROR} before publication: "
                f"{NUMBERS_COMMENT_SOURCE}"
            )
        elif not has_combined_route and rewrite >= 0 and unsupported > rewrite:
            violations.append(
                "focused Numbers comment changed clear must reject before its "
                f"rewrite path: {NUMBERS_COMMENT_SOURCE}"
            )

        if has_combined_route:
            # The combined seam is the only permitted metadata route.  A
            # removal-only or token-only call can otherwise make the native
            # clear appear atomic while leaving the other side of the graph
            # stale.
            forbidden_standalone = (
                (PACKAGE_METADATA_REMOVE_FUNCTION,)
                if prepared_route_complete
                else (PACKAGE_METADATA_REMOVE_FUNCTION, PACKAGE_METADATA_SAVE_TOKEN_FUNCTION)
            )
            for function_name in forbidden_standalone:
                if re.search(
                    rf"\b{re.escape(function_name)}\b[ \t\r\n]*\(",
                    reachable_code,
                ) is not None:
                    violations.append(
                        "focused Numbers comment clear must use only the combined "
                        f"PackageMetadata route, found {function_name}: "
                        f"{NUMBERS_COMMENT_SOURCE}"
                    )

            # These checks encode the narrow root-only admission policy.  Each
            # guard must fail with UnsupportedDependency nearby; a mere field
            # mention is not enough to admit a destructive graph mutation.
            required_guards = (
                (
                    "root-only/segment ownership",
                    re.compile(r"\bEntryOwner[ \t\r\n]*::[ \t\r\n]*Root\b"),
                    True,
                ),
                (
                    "shared entry uniqueness",
                    re.compile(
                        r"\b(?:[A-Za-z_][A-Za-z0-9_]*\s*\.\s*)*refcount\b[\s\S]{0,100}"
                        r"(?:!=|==|<|>)[\s\r\n]*1\b"
                    ),
                    True,
                ),
                (
                    "single storage ownership",
                    re.compile(
                        r"\bstorage_occurrences\b[\s\S]{0,100}"
                        r"(?:!=|==|<|>)[\s\r\n]*1\b"
                    ),
                    True,
                ),
                (
                    "reply ownership",
                    re.compile(
                        r"\breplies\b[\s\r\n]{0,100}"
                        r"(?:!=|==|<|>)[\s\r\n]*0\b"
                    ),
                    True,
                ),
                (
                    "unsupported entry owner",
                    re.compile(r"\bsupports_text_rewrite[ \t\r\n]*\("),
                    True,
                ),
                (
                    "global comment ownership",
                    re.compile(r"\b(?:prove_global_comment_ownership|global_comment_ownership)\b"),
                    False,
                ),
                (
                    "deletion-grade archive reference census",
                    re.compile(
                        r"\b(?:inspect_references_with_policy_and_limits|"
                        r"prove_archive_reference_ownership)\b"
                    ),
                    False,
                ),
                (
                    "opaque archive-owner rejection",
                    re.compile(r"\bArchiveReferencePolicy\s*::\s*RejectUnknownMetadata\b"),
                    False,
                ),
            )
            for label, pattern, requires_unsupported in required_guards:
                guard_match = pattern.search(reachable_code)
                if guard_match is None:
                    violations.append(
                        "focused Numbers combined comment clear is missing its "
                        f"{label} guard: {NUMBERS_COMMENT_SOURCE}"
                    )
                    continue
                guard_region = reachable_code[
                    max(0, guard_match.start() - 520) : guard_match.end() + 520
                ]
                if (
                    "return Err" not in guard_region
                    or (
                        requires_unsupported
                        and NUMBERS_COMMENT_CLEAR_UNSUPPORTED_ERROR not in guard_region
                    )
                ):
                    if not requires_unsupported and "map_err" in guard_region:
                        continue
                    guard_requirement = (
                        f"for {label} with {NUMBERS_COMMENT_CLEAR_UNSUPPORTED_ERROR}: "
                        if requires_unsupported
                        else f"for {label}: "
                    )
                    violations.append(
                        "focused Numbers combined comment clear must fail closed "
                        + guard_requirement
                        + f"{NUMBERS_COMMENT_SOURCE}"
                    )

            # Keep the semantic owner free of low-level public leakage and
            # eager generated reads.  Internal raw surgery remains private;
            # only public declarations and the reachable clear route are
            # inspected here.
            public_forbidden = re.compile(
                r"\b(?:RawMessage|WireView|WireLimits|WireResourceLimit|"
                r"Archive(?:Object)?|ObjectIdentifier|NativeId|native_id|"
                r"ComponentSelector|RemovalBatch|SaveTokenBatch|"
                r"RemovalSaveTokenBatch|prost|Message)\b"
            )
            for declaration, line_number in _rust_public_declarations(comment_source):
                if public_forbidden.search(declaration) is None:
                    continue
                violations.append(
                    "focused Numbers comment public API exposes a raw/wire/Prost "
                    f"type: {NUMBERS_COMMENT_SOURCE}:{line_number}"
                )
            if re.search(r"\b(?:prost|prost_types)\b", comment_code) is not None:
                violations.append(
                    "focused Numbers combined comment clear retains a Prost "
                    f"dependency marker: {NUMBERS_COMMENT_SOURCE}"
                )
            for label, pattern in PACKAGE_METADATA_NO_EAGER_DECODE_PATTERNS[:8]:
                if pattern.search(reachable_code) is None:
                    continue
                violations.append(
                    "focused Numbers combined comment clear reaches "
                    f"{label}: {NUMBERS_COMMENT_SOURCE}"
                )

    # In the reject-all state, no metadata primitive may be reachable from this
    # focused owner.  This prevents a partial clear from removing a cell
    # pointer while leaving stale UUID/external registrations.
    reachable_code = "\n".join(
        body for _name, body in reachable_bodies(NUMBERS_COMMENT_COMMIT_FUNCTION)
    )
    has_combined_route = re.search(
        rf"\b{re.escape(PACKAGE_METADATA_COMBINED_FUNCTION)}"
        r"[ \t\r\n]*\(",
        reachable_code,
    ) is not None or (
        re.search(r"\bprepare_root_clear[ \t\r\n]*\(", reachable_code) is not None
        and all(re.search(marker, reachable_code) is not None for marker in prepared_markers)
    )
    for function_name in (
        PACKAGE_METADATA_REMOVE_FUNCTION,
        PACKAGE_METADATA_SAVE_TOKEN_FUNCTION,
        PACKAGE_METADATA_COMBINED_FUNCTION,
    ):
        if has_combined_route:
            # A complete combined route may share private helpers/imports with
            # replacement and inspection paths.  The reachable clear graph was
            # checked above to forbid standalone removal/token publication.
            continue
        if re.search(
            rf"\b{re.escape(function_name)}\b[ \t\r\n]*\(", comment_code
        ) is None:
            continue
        line_number = comment_code.count(
            "\n", 0, comment_code.find(function_name)
        ) + 1
        violations.append(
            "focused Numbers comment owner reaches PackageMetadata "
            f"{function_name} before the combined clear route is complete: "
            f"{NUMBERS_COMMENT_SOURCE}:{line_number}"
        )

    return sorted(set(violations))


def audit_numbers_extractor_no_eager_comment_storage_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers comment-storage ingress free of eager Prost reads.

    Comment storage now has a bounded borrowed projection in the Numbers
    extractor's production path.  Reconstructing the generated
    ``tsd::CommentStorageArchive`` would materialize the complete legacy
    envelope, including the unbounded replies vector, before the strict
    comment budgets are applied.  The audit is scoped to this extractor and
    its production slice: canonical generated fixtures remain available in
    cfg(test)-gated items.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_EXTRACTOR_SOURCE
    if not source_path.is_file():
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    for label, pattern in NUMBERS_EXTRACTOR_NO_EAGER_COMMENT_STORAGE_SOURCE_PATTERNS:
        for match in pattern.finditer(production_code):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers extractor production source uses "
                f"{label}: {NUMBERS_EXTRACTOR_SOURCE}:{line_number}"
            )

    return sorted(set(violations))


def audit_numbers_extractor_no_eager_formula_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers formula sidecar free of generated archive reads.

    Formula entries are strictly preflighted and retained as bounded source
    bytes before the private lazy projection/event stream renders a referenced
    cell.  Reintroducing ``FormulaArchive::decode`` in the extractor's
    production source would materialize the generated repeated AST graph and
    bypass the production ownership boundary.  Test-only generated builders
    and decodes remain available as individually ``cfg(test)``-gated
    differential fixture material.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_EXTRACTOR_SOURCE
    if not source_path.is_file():
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    for label, pattern in NUMBERS_EXTRACTOR_NO_EAGER_FORMULA_SOURCE_PATTERNS:
        for match in pattern.finditer(production_code):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers extractor production source uses "
                f"{label}: {NUMBERS_EXTRACTOR_SOURCE}:{line_number}"
            )

    return sorted(set(violations))


def audit_numbers_names_package_no_eager_prost_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers names owner free of generated Prost reads.

    The names transaction is a production ingress just like the package
    reader.  Its wire projections must go through the bounded Buffa codecs;
    generated archive ``decode`` calls are retained only in an explicitly
    test-gated module for differential fixtures.  The source slice is
    therefore excludes cfg(test)-gated items from the production source
    slice, matching the package ingress ratchet above.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_NAMES_PACKAGE_SOURCE
    if not source_path.is_file():
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    for label, pattern in NUMBERS_NAMES_NO_EAGER_PROST_SOURCE_PATTERNS:
        for match in pattern.finditer(production_code):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers names production source uses "
                f"{label}: {NUMBERS_NAMES_PACKAGE_SOURCE}:{line_number}"
            )

    return sorted(set(violations))


def audit_numbers_chart_arrangement_owner_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the selector-first Numbers chart-arrangement package owner."""

    return _audit_focused_chart_arrangement_owner_source_topology(
        root,
        label="litchi-numbers chart-arrangement",
        owner_source=NUMBERS_CHART_ARRANGEMENT_OWNER_SOURCE,
        package_source=NUMBERS_PACKAGE_SOURCE,
        # ADR 0034: the family is published through the
        # focused `sheet::chart::arrangement` module with short names.
        public_source=NUMBERS_SOURCE_ROOT / "sheet" / "chart.rs",
        module_name=NUMBERS_CHART_ARRANGEMENT_PACKAGE_MODULE,
        methods=NUMBERS_CHART_ARRANGEMENT_PACKAGE_METHODS,
        selector_types={
            "sheet_chart_arrangement": NUMBERS_CHART_ARRANGEMENT_SELECTOR_TYPES,
            "sheet_chart_arrangements": ("SheetSelector",),
            "edit_sheet_chart_arrangement": NUMBERS_CHART_ARRANGEMENT_SELECTOR_TYPES,
            "apply_sheet_chart_arrangement": (),
        },
        method_markers={
            "sheet_chart_arrangement": ("select_chart_with_budget",),
            "sheet_chart_arrangements": ("discover_sheet_charts",),
            "edit_sheet_chart_arrangement": ("Edit::new",),
            "apply_sheet_chart_arrangement": (
                "select_chart_with_budget",
                "verify_locality",
            ),
        },
        owner_markers=NUMBERS_CHART_ARRANGEMENT_OWNER_MARKERS,
        lazy_codec_markers=NUMBERS_CHART_ARRANGEMENT_LAZY_CODEC_MARKERS,
        canonical_types=(
            "Commit",
            "Diagnostics",
            "Edit",
            "Error",
            "LimitKind",
            "Patch",
        ),
    )

# Native merge formulas are decoded once by the low-level Numbers wire crate.
# The wire owner may depend on generated schema definitions, but focused
# format facades must consume this semantic report instead of exposing a
# generated AST or recreating the merge walk.
IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE = Path(
    "crates/litchi-numbers-wire/src/table_merges.rs"
)
IWA_NUMBERS_WIRE_TABLE_MERGES_REQUIRED_TYPES = frozenset(
    {"AttemptedCost", "MergeRead", "MergeReadError", "ReadLimits", "ReadReport"}
)
IWA_NUMBERS_WIRE_TABLE_MERGES_REQUIRED_FUNCTIONS = frozenset(
    {"read_table_merges"}
)
IWA_NUMBERS_WIRE_TABLE_MERGES_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_merges\b"
)
IWA_NUMBERS_WIRE_TABLE_MERGES_COMMON_IMPORT = re.compile(
    r"\blitchi_iwa_common[ \t\r\n]*::"
)
IWA_NUMBERS_WIRE_TABLE_MERGES_WIRE_LIMITS = re.compile(r"\bWireLimits\b")
IWA_NUMBERS_WIRE_TABLE_MERGES_REGION = re.compile(r"\bRegion\b")
IWA_NUMBERS_WIRE_TABLE_MERGES_FORMAT_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"litchi_numbers|litchi_pages|litchi_keynote"
    r")(?![A-Za-z0-9_])"
)
IWA_NUMBERS_WIRE_TABLE_MERGES_LOCAL_DECLARATIONS = (
    (
        "ReadLimits",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+(?:r#)?ReadLimits\b"
        ),
    ),
    (
        "ReadReport",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+(?:r#)?ReadReport\b"
        ),
    ),
    (
        "MergeRead",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+(?:r#)?MergeRead\b"
        ),
    ),
    (
        "AttemptedCost",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+(?:r#)?AttemptedCost\b"
        ),
    ),
    (
        "MergeReadError",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+(?:r#)?MergeReadError\b"
        ),
    ),
)
IWA_NUMBERS_WIRE_TABLE_MERGES_LOCAL_FUNCTIONS = (
    (
        "read_table_merges",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?read_table_merges\b"
        ),
    ),
)
NUMBERS_TABLE_MERGE_READER_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_merges/reader.rs"
)
NUMBERS_TABLE_MERGE_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?table_merges\b[ \t]*(?:;|\{)"
)
NUMBERS_TABLE_MERGE_PUBLIC_METHOD = "table_merges"
NUMBERS_TABLE_MERGE_SELECTOR_TYPES = frozenset({"SheetSelector", "TableSelector"})
NUMBERS_TABLE_MERGE_FORBIDDEN_PUBLIC_BNC = re.compile(
    r"(?<![A-Za-z0-9_])(?:BncCell|BncCellView|PreBncCellView|"
    r"StoredValue|CachedScalar)(?![A-Za-z0-9_])"
)
# The package-level Numbers merge owner keeps the shared geometry/budget seam,
# while this optional private child owns ZIP ingress and rooted metadata
# selection.  Keep its source-facing surface deliberately small: callers may
# construct a reader from a path or immutable package bytes, then query with
# typed selectors.  The child may use the borrowed name codec, but generated
# protobuf/BNC values and byte-returning escape hatches stay below this seam.
NUMBERS_TABLE_MERGE_READER_ALLOWED_METHODS = frozenset(
    {
        "open",
        "open_with_options",
        "from_bytes",
        "from_bytes_with_options",
        "from_shared_bytes",
        "from_shared_bytes_with_options",
        "table_merges",
    }
)
NUMBERS_TABLE_MERGE_READER_PROTO_TOKEN = re.compile(
    r"(?<![A-Za-z0-9_])litchi_iwa_protos(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_MERGE_READER_FORBIDDEN_TYPES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"FormulaArchive|AstNodeArchive|AstNodeType|MergeRegionMapArchive|"
    r"DocumentArchive|SheetArchive|FormBasedSheetArchive|"
    r"TableInfoArchive|TableModelArchive|"
    r"BncCell|BncCellView|PreBncCellView|StoredValue|CachedScalar|"
    r"numbers_formula_codec|numbers_table_cell_storage_codec|"
    r"prost|prost_types|buffa"
    r")(?![A-Za-z0-9_])"
)
NUMBERS_TABLE_MERGE_READER_USE_STATEMENT = re.compile(
    r"(?ms)^[ \t]*(?:pub(?:[ \t]*\([^()\r\n]*\))?[ \t]+)?"
    r"use\b.*?;"
)
NUMBERS_TABLE_MERGE_READER_ALLOWED_PROTO_IMPORT = re.compile(
    r"(?m)^[ \t]*use[ \t]+litchi_iwa_protos[ \t]*::[ \t]*"
    r"numbers_names_codec[ \t]*;"
)
IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_MODULE = re.compile(
    r"(?m)^[ \t]*mod[ \t]+(?:r#)?writer[ \t]*;"
)
IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_EXPORT = re.compile(
    r"(?ms)^[ \t]*pub[ \t]+use[ \t]+writer[ \t]*::[ \t]*\{"
    r"[^{};]*\bMergeWrite\b[^{};]*\brewrite_table_merges\b[^{};]*\}[ \t]*;"
)
IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_PROTO_IMPORT = re.compile(
    r"(?m)^[ \t]*use[ \t]+litchi_iwa_protos[ \t]*::[ \t]*"
    r"table_merge_formula_codec[ \t]*"
    r"(?:::[ \t]*\{[^{};]*\})?[ \t]*;"
)
IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_CODEC_CALL = re.compile(
    r"\b(?:merge_formula_codec|table_merge_formula_codec)[ \t\r\n]*::"
    r"[ \t\r\n]*encode_merge_formula[ \t\r\n]*\("
)
IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:prost|prost_types|"
    r"litchi_numbers|litchi_pages|litchi_keynote)(?![A-Za-z0-9_])"
)
IWA_NUMBERS_DEPENDENCY_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_dependency_codec.rs"
)
IWA_NUMBERS_DEPENDENCY_REMOVAL_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_dependency_codec/removal.rs"
)
IWA_NUMBERS_DEPENDENCY_REMOVAL_MODULE = "removal"
IWA_NUMBERS_DEPENDENCY_INBOUND_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_cell_dependency_codec/inbound.rs"
)
IWA_NUMBERS_DEPENDENCY_INBOUND_MODULE = "inbound"
IWA_NUMBERS_DEPENDENCY_INBOUND_TYPES = (
    "FormulaDependencyFact",
    "FormulaOwnerInternalDependencyFact",
    "FormulaOwnerInternalDependencyKind",
    "FormulaOwnerUuidDependencyFact",
    "FormulaOwnerUuidDependencyKind",
)
IWA_NUMBERS_DEPENDENCY_INBOUND_PARENT_FUNCTIONS = (
    "decode_formula_owner_dependency_facts_with_visitor",
    "decode_formula_owner_dependency_facts",
    "decode_formula_owner_dependency_facts_with_report",
)
IWA_NUMBERS_DEPENDENCY_INBOUND_MARKER_GROUPS = {
    "volatile dependency facts": (
        "decode_volatile_dependencies_in",
        "VolatileGeometry",
    ),
    "spanning dependency facts": (
        "decode_spanning_dependencies_in",
        "SpanningColumn",
        "SpanningRow",
    ),
    "whole-owner dependency facts": (
        "decode_whole_owner_dependencies_in",
        "WholeOwner",
    ),
    "UUID dependency facts": (
        "decode_uuid_references_in",
        "TableUuidReference",
        "OwnerUuid",
    ),
    "typed visitor dispatch": (
        "visit_formula_owner_internal_dependency",
        "visit_formula_owner_uuid_dependency",
        "FormulaDependencyFact",
    ),
    "coordinate validation": (
        "decode_cell_coordinate_in",
        "decode_range_coordinate_in",
        "coord_set",
        "range_context",
    ),
    "bounded borrowed traversal": (
        "budget.message",
        "child_depth",
        "checked_add",
        "set_once",
        "DecodeError::invalid",
    ),
}
IWA_NUMBERS_DEPENDENCY_REMOVAL_PUBLIC_PREPARED = re.compile(
    r"(?m)^\s*pub\s+struct\s+(?:Prepared[A-Za-z0-9_]*|"
    r"[A-Za-z0-9_]*(?:Removal|Rewrite)Plan|[A-Za-z0-9_]*Plan)\b"
)
IWA_NUMBERS_DEPENDENCY_REMOVAL_PREPARE_FUNCTION = re.compile(
    r"(?m)^\s*pub\s+fn\s+(?:prepare|execute|rewrite)_[A-Za-z0-9_]*"
    r"(?:dependency|formula|owner|tile|remov|prun)[A-Za-z0-9_]*\b",
    re.IGNORECASE,
)
IWA_NUMBERS_DEPENDENCY_REMOVAL_MARKER_GROUPS = {
    "formula engine owner removal": (
        "CalculationEngine",
        "calculation_engine",
        "formula_engine",
        "formula_owner",
        "owner_id_map",
        "number_of_formulas",
    ),
    "dependency tile pruning": (
        "cell_record_tile",
        "cell_record_tiles",
        "range_precedents_tile",
        "tiled",
        "tile",
        "prune",
        "remove",
    ),
    "prepared bounded rewrite": (
        "Prepared",
        "CalculationEngineOwnerRemovalPlan",
        "FormulaDependencyRewriteRequirements",
        "prepare_calculation_engine_owner_removal",
        "prepare_formula_owner_cell_edges_removal",
        "prepare_cell_record_tile_edges_removal",
        "execute_calculation_engine_owner_removal",
        "execute_formula_owner_cell_edges_removal",
        "execute_cell_record_tile_edges_removal",
        "rewrite_calculation_engine_owner_removal",
        "rewrite_formula_owner_cell_edges",
        "rewrite_cell_record_tile_edges",
    ),
    "Buffa lazy source preservation": (
        "buffa",
        "Buffa",
        "decode_lazy_view",
        "LazyView",
        "source",
        "unknown",
        "candidate",
    ),
    "checked accounting": (
        "checked_add",
        "checked_sub",
        "try_reserve",
        "limit",
        "budget",
        "bounded",
    ),
}

IWA_NUMBERS_WIRE_TABLE_CELLS_SOURCE = Path(
    "crates/litchi-numbers-wire/src/table_cells.rs"
)
IWA_NUMBERS_WIRE_TABLE_CELLS_LIB_SOURCE = Path(
    "crates/litchi-numbers-wire/src/lib.rs"
)
IWA_NUMBERS_WIRE_TABLE_CELLS_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_cells\b"
)
IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_TYPES = frozenset(
    {"TableDimensions", "TileReference", "CellSource", "TableCellIssue", "TableCellReadReport"}
)
IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_TRAITS = frozenset(
    {"TableCellReadBudget", "CellValueSink"}
)
IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_IMPORTS = (
    re.compile(r"\blitchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*numbers_table_cell_storage_codec\b"),
    re.compile(r"\bcrate[ \t\r\n]*::[ \t\r\n]*cell_value\b"),
)
IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_MARKERS = (
    "decode_tile_with_visitor",
    "decode_cell_value",
    "TableCellIssue",
)

IWA_NUMBERS_WIRE_TABLE_SIDECARS_SOURCE = Path(
    "crates/litchi-numbers-wire/src/table_sidecars.rs"
)
IWA_NUMBERS_WIRE_TABLE_SIDECARS_LIB_SOURCE = IWA_NUMBERS_WIRE_TABLE_CELLS_LIB_SOURCE
IWA_NUMBERS_WIRE_TABLE_SIDECARS_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_sidecars\b"
)
IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_TYPES = frozenset(
    {"SidecarKind", "SidecarReference", "SidecarValue", "SidecarList", "SidecarTables"}
)
IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_TRAITS = frozenset(
    {"SidecarReadBudget", "CellSidecarResolver"}
)
IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_FUNCTIONS = frozenset(
    {"read_sidecar_list", "read_comment_storage"}
)
IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_IMPORTS = (
    re.compile(r"\btable_data_list[ \t\r\n]*::[ \t\r\n]*read_list_with_decoder\b"),
    re.compile(r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*table[ \t\r\n]*::[ \t\r\n]*read\b"),
)
IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_MARKERS = (
    "read_list_with_decoder",
    "materialize_cell",
    "CellSidecarResolver",
)
IWA_NUMBERS_WIRE_TABLE_SIDECARS_PEER_IMPORT = IWA_NUMBERS_WIRE_TABLE_CELLS_PEER_IMPORT

# Formula-owner identity and category labels are another low-level Numbers
# wire seam.  The owner may use the generated GroupNode codec, but it must
# publish borrowed/scalar projections and keep concrete format crates out of
# its dependency surface.  Reader migration is checked separately once a
# focused package actually consumes these APIs.
IWA_NUMBERS_WIRE_FORMULA_NAMES_SOURCE = Path(
    "crates/litchi-numbers-wire/src/formula_names.rs"
)
IWA_NUMBERS_WIRE_FORMULA_NAMES_LIB_SOURCE = IWA_NUMBERS_WIRE_TABLE_CELLS_LIB_SOURCE
IWA_NUMBERS_WIRE_FORMULA_NAMES_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?formula_names\b"
)
IWA_NUMBERS_WIRE_FORMULA_NAMES_REQUIRED_TYPES = frozenset(
    {
        "ReadLimits",
        "ReadReport",
        "FormulaOwnerDependencies",
        "FormulaOwnerReadError",
        "CategoryReadLimits",
        "CategoryLabel",
        "CategoryName",
        "CategoryRead",
        "CategoryReadError",
    }
)
IWA_NUMBERS_WIRE_FORMULA_NAMES_REQUIRED_FUNCTIONS = frozenset(
    {"read_formula_owner_dependencies", "read_formula_category_names"}
)
IWA_NUMBERS_WIRE_FORMULA_NAMES_REQUIRED_MARKERS = (
    "preflight_wire_tree",
    "group_node_category_codec",
    "WireDescent",
    "read_formula_owner_dependencies",
    "read_formula_category_names",
)
IWA_NUMBERS_WIRE_FORMULA_NAMES_PEER_IMPORT = IWA_NUMBERS_WIRE_TABLE_CELLS_PEER_IMPORT


def audit_iwa_numbers_wire_table_merges_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep native table-merge decoding in one bounded wire owner.

    ``litchi-numbers-wire`` is allowed to inspect the generated Numbers schema
    because it is the immediate low-level adapter.  Its exported result is a
    compact semantic report over common ``Region`` values.  Format packages
    must call this report and keep package identity, selectors, and typed
    error mapping at their own boundary.
    """

    owner_path = root / IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE
    focused_paths = (
        root / Path("crates/litchi-pages/src/package/body_table_merges.rs"),
        root / Path("crates/litchi-keynote/src/package/slide_table_merges.rs"),
    )
    if not owner_path.is_file() and not any(path.is_file() for path in focused_paths):
        return []

    violations: list[str] = []
    if not owner_path.is_file():
        violations.append(
            "Numbers wire table-merge owner is missing: "
            f"{IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
        )
        return violations

    raw_source = owner_path.read_text(encoding="utf-8")
    code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    library_path = root / Path("crates/litchi-numbers-wire/src/lib.rs")
    if not library_path.is_file():
        violations.append(
            "Numbers wire table-merge owner is missing its crate export: "
            "crates/litchi-numbers-wire/src/lib.rs"
        )
    else:
        library_code = _mask_rust_non_code(library_path.read_text(encoding="utf-8"))
        if IWA_NUMBERS_WIRE_TABLE_MERGES_MODULE.search(library_code) is None:
            violations.append(
                "Numbers wire table-merge owner is missing public table_merges module: "
                "crates/litchi-numbers-wire/src/lib.rs"
            )

    for name in sorted(IWA_NUMBERS_WIRE_TABLE_MERGES_REQUIRED_TYPES):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(code) is None:
            violations.append(
                "Numbers wire table-merge owner is missing public "
                f"{name}: {IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
            )
    for name in sorted(IWA_NUMBERS_WIRE_TABLE_MERGES_REQUIRED_FUNCTIONS):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:const[ \t]+)?fn[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(code) is None:
            violations.append(
                "Numbers wire table-merge owner is missing public "
                f"{name}: {IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
            )

    for name, declaration in IWA_NUMBERS_WIRE_TABLE_MERGES_LOCAL_DECLARATIONS:
        if len(declaration.findall(code)) > 1:
            violations.append(
                "Numbers wire table-merge owner redeclares shared type "
                f"{name}: {IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
            )
    for name, declaration in IWA_NUMBERS_WIRE_TABLE_MERGES_LOCAL_FUNCTIONS:
        if len(declaration.findall(code)) > 1:
            violations.append(
                "Numbers wire table-merge owner redeclares shared function "
                f"{name}: {IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
            )

    if IWA_NUMBERS_WIRE_TABLE_MERGES_COMMON_IMPORT.search(code) is None:
        violations.append(
            "Numbers wire table-merge owner is missing its common import: "
            f"{IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_TABLE_MERGES_WIRE_LIMITS.search(code) is None:
        violations.append(
            "Numbers wire table-merge owner is missing bounded WireLimits: "
            f"{IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
        )
    if IWA_NUMBERS_WIRE_TABLE_MERGES_REGION.search(code) is None:
        violations.append(
            "Numbers wire table-merge owner is missing common Region values: "
            f"{IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}"
        )

    for match in IWA_NUMBERS_WIRE_TABLE_MERGES_FORMAT_IMPORT.finditer(code):
        line_number = code.count("\n", 0, match.start()) + 1
        violations.append(
            "Numbers wire table-merge owner imports a concrete format owner "
            f"{match.group(0)}: {IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE}:{line_number}"
        )

    return sorted(set(violations))


def audit_numbers_table_merge_reader_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the optional Numbers merge-reader child at the ingress seam.

    ``table_merges.rs`` owns the archive-free geometry boundary.  When the
    private ``reader.rs`` child exists, it may expose only path/byte ingress
    constructors and the typed-selector query, plus the borrowed Numbers name
    codec import needed for metadata selection.  This audit is intentionally
    dormant while the child is absent so the migration can land in stages.
    """

    reader_path = root / NUMBERS_TABLE_MERGE_READER_SOURCE
    if not reader_path.is_file():
        return []

    owner_path = root / NUMBERS_TABLE_MERGE_SOURCE
    violations: list[str] = []
    reader_source = reader_path.read_text(encoding="utf-8")
    production_source = _mask_rust_cfg_test_items(reader_source)
    code = _mask_rust_non_code(production_source)
    relative_reader = NUMBERS_TABLE_MERGE_READER_SOURCE

    if not owner_path.is_file():
        violations.append(
            "Numbers table-merge reader child has no package owner: "
            f"{NUMBERS_TABLE_MERGE_SOURCE}"
        )
    else:
        owner_source = owner_path.read_text(encoding="utf-8")
        module_declarations = _rust_root_level_module_declarations(
            owner_source, frozenset({"reader"})
        )
        if not module_declarations:
            violations.append(
                "Numbers table-merge reader child is not wired from its owner: "
                f"{NUMBERS_TABLE_MERGE_SOURCE}"
            )
        elif any(visibility is not None for _name, visibility, *_rest in module_declarations):
            violations.append(
                "Numbers table-merge reader child module must remain private: "
                f"{NUMBERS_TABLE_MERGE_SOURCE}"
            )

    # The only physical-protobuf import admitted above the wire reader is the
    # exact borrowed Numbers name codec.  Inspect use statements separately so
    # a broad/aliased litchi_iwa_protos import cannot hide an owned projection.
    allowed_import_spans: list[tuple[int, int]] = []
    for statement in NUMBERS_TABLE_MERGE_READER_USE_STATEMENT.finditer(code):
        if not NUMBERS_TABLE_MERGE_READER_PROTO_TOKEN.search(statement.group(0)):
            continue
        if NUMBERS_TABLE_MERGE_READER_ALLOWED_PROTO_IMPORT.fullmatch(
            statement.group(0).strip()
        ) is None:
            line_number = code.count("\n", 0, statement.start()) + 1
            violations.append(
                "Numbers table-merge reader must use only the exact "
                "numbers_names_codec import: "
                f"{relative_reader}:{line_number}"
            )
        else:
            allowed_import_spans.append(statement.span())

    for match in NUMBERS_TABLE_MERGE_READER_PROTO_TOKEN.finditer(code):
        if any(start <= match.start() < end for start, end in allowed_import_spans):
            continue
        line_number = code.count("\n", 0, match.start()) + 1
        violations.append(
            "Numbers table-merge reader contains an unapproved litchi_iwa_protos "
            f"path: {relative_reader}:{line_number}"
        )

    for match in NUMBERS_TABLE_MERGE_READER_FORBIDDEN_TYPES.finditer(code):
        line_number = code.count("\n", 0, match.start()) + 1
        violations.append(
            "Numbers table-merge reader contains generated/owned protobuf or "
            f"BNC type {match.group(0)}: {relative_reader}:{line_number}"
        )

    merge_reader_struct = _rust_named_struct_body(production_source, "MergeReader")
    if merge_reader_struct is not None:
        struct_body, body_offset = merge_reader_struct
        for match in NUMBERS_TABLE_MERGE_READER_PUBLIC_FIELD.finditer(struct_body):
            line_number = code.count("\n", 0, body_offset + match.start()) + 1
            violations.append(
                "Numbers table-merge reader exposes a public MergeReader field: "
                f"{relative_reader}:{line_number}"
            )

    public_declarations = _rust_public_declarations(production_source)
    merge_reader_method_lines = {
        line
        for _name, _declaration, line in _rust_public_methods_in_impl(
            production_source, "MergeReader"
        )
    }
    for declaration, line_number in public_declarations:
        function = NUMBERS_TABLE_MERGE_READER_FUNCTION_NAME.search(declaration)
        if function is None:
            if re.search(r"\bstruct[ \t]+(?:r#)?MergeReader\b", declaration):
                continue
            violations.append(
                "Numbers table-merge reader exposes an unsupported public "
                f"declaration: {relative_reader}:{line_number}"
            )
            continue

        name = function.group("name")
        if line_number not in merge_reader_method_lines:
            violations.append(
                "Numbers table-merge reader public method must be an associated "
                f"MergeReader method ({name}): {relative_reader}:{line_number}"
            )
        arrow = declaration.find("->")
        return_type = declaration[arrow + 2 :] if arrow >= 0 else ""
        if (
            RUST_BYTE_SLICE.search(return_type) is not None
            or NUMBERS_TABLE_MERGE_READER_BYTE_RETURN.search(return_type) is not None
        ):
            violations.append(
                "Numbers table-merge reader exposes raw bytes from public method "
                f"{name}: {relative_reader}:{line_number}"
            )

        raw_id = IWA_FOCUSED_TABLE_MERGES_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            violations.append(
                "Numbers table-merge reader exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: {relative_reader}:{line_number}"
            )

        if name not in NUMBERS_TABLE_MERGE_READER_ALLOWED_METHODS:
            violations.append(
                "Numbers table-merge reader exposes unsupported public method "
                f"{name}: {relative_reader}:{line_number}"
            )
            continue

        if not re.search(r"\bResult[ \t]*<[ \t]*Self\b", return_type):
            if name != "table_merges":
                violations.append(
                    "Numbers table-merge reader source ingress method must return "
                    f"Result<Self, _> ({name}): {relative_reader}:{line_number}"
                )
        if name in {"open", "open_with_options"}:
            if re.search(r"\bAsRef[ \t]*<[ \t]*Path[ \t]*>", declaration) is None:
                violations.append(
                    "Numbers table-merge reader path ingress must accept "
                    f"AsRef<Path> ({name}): {relative_reader}:{line_number}"
                )
        elif name in {"from_bytes", "from_bytes_with_options"}:
            if RUST_BYTE_SLICE.search(declaration) is None:
                violations.append(
                    "Numbers table-merge reader byte ingress must accept &[u8] "
                    f"({name}): {relative_reader}:{line_number}"
                )
        elif name in {"from_shared_bytes", "from_shared_bytes_with_options"}:
            if re.search(
                r"\bArc[ \t]*<[ \t]*\[[ \t]*u8[ \t]*\][ \t]*>", declaration
            ) is None:
                violations.append(
                    "Numbers table-merge reader shared ingress must accept "
                    f"Arc<[u8]> ({name}): {relative_reader}:{line_number}"
                )
        elif name == "table_merges":
            if not re.search(r"\bSheetSelector\b", declaration):
                violations.append(
                    "Numbers table-merge reader table_merges must use "
                    f"SheetSelector: {relative_reader}:{line_number}"
                )
            if not re.search(r"\bTableSelector\b", declaration):
                violations.append(
                    "Numbers table-merge reader table_merges must use "
                    f"TableSelector: {relative_reader}:{line_number}"
                )
            if re.search(r"\bResult[ \t]*<[ \t]*Vec[ \t]*<[ \t]*Region\b", return_type) is None:
                violations.append(
                    "Numbers table-merge reader table_merges must return "
                    f"semantic Region values: {relative_reader}:{line_number}"
                )

    return sorted(set(violations))


def audit_numbers_table_merge_source_topology(root: Path = ROOT) -> list[str]:
    """Keep Numbers merged-cell reads selector-first and archive-free.

    The package owns sheet/table selection and delegates native merge storage
    to the bounded Numbers wire reader.  Its old raw-ID compatibility reader
    remains in the migration host; this guard covers only the new focused
    package owner and keeps BNC/pre-BNC views private to the wire layer.
    """

    violations = _audit_focused_table_merge_source_topology(
        root,
        format_name="Numbers",
        owner_path=NUMBERS_TABLE_MERGE_SOURCE,
        package_path=NUMBERS_TABLE_MERGE_PACKAGE_SOURCE,
        module_pattern=NUMBERS_TABLE_MERGE_MODULE,
        table_module_path=Path("crates/litchi-numbers/src/table.rs"),
        table_merge_path=Path("crates/litchi-numbers/src/table/merge.rs"),
        table_module_pattern=PAGES_TABLE_MERGE_COMMON_TABLE_MODULE,
        public_method=NUMBERS_TABLE_MERGE_PUBLIC_METHOD,
        selector_types=NUMBERS_TABLE_MERGE_SELECTOR_TYPES,
        forbidden_public_types=NUMBERS_TABLE_MERGE_FORBIDDEN_PUBLIC_BNC,
    )
    violations.extend(audit_numbers_table_merge_reader_source_topology(root))
    return sorted(set(violations))


def audit_numbers_table_merge_transaction_source_topology(
    root: Path = ROOT,
) -> list[str]:
    return _audit_table_merge_transaction_source_topology(root, IWA_TABLE_MERGE_TRANSACTION_SPECS[0])


def audit_iwa_numbers_wire_table_merge_writer_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the shared merge writer bounded, source-preserving, and neutral."""

    writer_path = IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_SOURCE
    writer_absolute = root / writer_path
    owner_path = IWA_NUMBERS_WIRE_TABLE_MERGES_SOURCE
    owner_absolute = root / owner_path
    if not writer_absolute.is_file():
        if owner_absolute.is_file() and re.search(
            r"(?m)^\s*mod\s+writer\s*;|\brewrite_table_merges\b",
            _mask_rust_non_code(owner_absolute.read_text(encoding="utf-8")),
        ):
            return [f"Numbers wire table-merge writer is missing: {writer_path}"]
        return []

    violations: list[str] = []
    writer_source = _mask_rust_cfg_test_items(
        writer_absolute.read_text(encoding="utf-8")
    )
    code = _mask_rust_non_code(writer_source)
    if not re.search(r"(?m)^\s*pub\s+struct\s+MergeWrite\b", code):
        violations.append(f"Numbers wire table-merge writer is missing public MergeWrite: {writer_path}")
    if not re.search(r"(?m)^\s*pub\s+fn\s+rewrite_table_merges\b", code):
        violations.append(f"Numbers wire table-merge writer is missing public rewrite_table_merges: {writer_path}")

    merge_write = _rust_named_struct_body(writer_source, "MergeWrite")
    if merge_write is not None:
        body, _offset = merge_write
        for field in ("data", "report", "changed"):
            if re.search(rf"(?m)^\s*pub\s+{re.escape(field)}\s*:", body) is None:
                violations.append(
                    f"Numbers wire table-merge writer MergeWrite is missing public field "
                    f"{field}: {writer_path}"
                )

    marker_groups = {
        "strict readback": ("read_table_merges", "finish_candidate"),
        "source-preserving walk": ("gather_source", "gather_owner", "gather_store", "sink.raw"),
        "desired-state validation": (
            "validate_desired_regions",
            "validate_bounds_and_set",
            "same_region_set",
        ),
        "owner lifecycle": (
            ("RemoveOwner", "RootAction::Remove"),
            ("ReplaceOwner", "RootAction::Replace"),
            ("AppendOwner", "RootAction::Append"),
            ("next_index",),
        ),
        "bounded output": (
            "Sink",
            "try_reserve_exact",
            "checked_add",
            "max_output",
            "LimitKind::OutputBytes",
        ),
        "formula-pair encoding": ("encode_merge_formula", "encode_pair", "encode_owner_id"),
    }
    for label, markers in marker_groups.items():
        if not all(
            _table_merge_transaction_has_marker(
                code,
                marker if isinstance(marker, tuple) else (marker,),
            )
            for marker in markers
        ):
            violations.append(
                f"Numbers wire table-merge writer is missing {label} marker: "
                f"{writer_path}"
            )

    if owner_absolute.is_file():
        owner_source = _mask_rust_cfg_test_items(
            owner_absolute.read_text(encoding="utf-8")
        )
        owner_code = _mask_rust_non_code(owner_source)
        if IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_MODULE.search(owner_code) is None:
            violations.append(
                f"Numbers wire table-merge writer is missing its private module wiring: "
                f"{owner_path}"
            )
        if re.search(r"(?m)^\s*pub\s+mod\s+writer\b", owner_code):
            violations.append(
                f"Numbers wire table-merge writer module must remain private: {owner_path}"
            )
        if not IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_EXPORT.search(owner_code):
            violations.append(
                f"Numbers wire table-merge writer is missing its checked re-export: {owner_path}"
            )
    else:
        violations.append(f"Numbers wire table-merge owner is missing: {owner_path}")

    use_statement = re.compile(
        r"(?ms)^[ \t]*(?:pub(?:\([^()\r\n]*\))?[ \t]+)?use\b.*?;"
    )
    for statement in use_statement.finditer(code):
        text = statement.group(0)
        if IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_FORBIDDEN_IMPORT.search(text):
            line_number = code.count("\n", 0, statement.start()) + 1
            violations.append(
                f"Numbers wire table-merge writer imports a concrete format/protobuf "
                f"owner: {writer_path}:{line_number}"
            )
        if "litchi_iwa_protos" in text and not IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_PROTO_IMPORT.fullmatch(text.strip()):
            line_number = code.count("\n", 0, statement.start()) + 1
            violations.append(
                f"Numbers wire table-merge writer may import only the narrow formula "
                f"codec: {writer_path}:{line_number}"
            )

    for pattern, label in (
        (re.compile(r"\b(?:prost|prost_types|encode_to_vec|Message::decode)\b"), "eager/prost codec"),
        (re.compile(r"\b(?:TableModelArchive|TableInfoArchive|FormulaArchive|AstNodeArchive)\b"), "generated schema type"),
    ):
        match = pattern.search(code)
        if match is not None:
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                f"Numbers wire table-merge writer retains a {label}: "
                f"{writer_path}:{line_number}"
            )

    # The writer may account for the codec's report and translate its errors,
    # but formula bytes must be produced by the narrow Buffa owner.  Check the
    # call itself so a local helper with the same name cannot silently restore
    # the handwritten TSCE builder.
    codec_call = IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_CODEC_CALL.search(code)
    if codec_call is None:
        violations.append(
            "Numbers wire table-merge writer must delegate formula encoding to "
            f"{IWA_TABLE_MERGE_FORMULA_CODEC_MODULE}: {writer_path}"
        )
    else:
        formula_helper = _rust_top_level_function_bodies(code).get(
            "encode_merge_formula"
        )
        if formula_helper is not None and IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_CODEC_CALL.search(
            formula_helper[0]
        ) is None:
            line_number = code.count("\n", 0, formula_helper[1]) + 1
            violations.append(
                "Numbers wire table-merge writer's formula helper must delegate "
                f"to {IWA_TABLE_MERGE_FORMULA_CODEC_MODULE}: "
                f"{writer_path}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_numbers_table_cell_dependency_removal_source_topology(
    root: Path = ROOT,
) -> list[str]:
    return _audit_iwa_component_removal_codec_source_topology(
        root,
        parent_path=IWA_NUMBERS_DEPENDENCY_CODEC_SOURCE,
        child_path=IWA_NUMBERS_DEPENDENCY_REMOVAL_SOURCE,
        module_name=IWA_NUMBERS_DEPENDENCY_REMOVAL_MODULE,
        label="Numbers dependency-removal",
        prepared_pattern=IWA_NUMBERS_DEPENDENCY_REMOVAL_PUBLIC_PREPARED,
        prepare_pattern=IWA_NUMBERS_DEPENDENCY_REMOVAL_PREPARE_FUNCTION,
        marker_groups=IWA_NUMBERS_DEPENDENCY_REMOVAL_MARKER_GROUPS,
        build_marker='println!("cargo:rerun-if-changed=src/numbers_table_cell_dependency_codec.rs");',
        guard_marker='include_str!("numbers_table_cell_dependency_codec/removal.rs")',
    )


def audit_iwa_numbers_table_cell_dependency_inbound_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep expanded inbound dependency facts in one private Buffa child."""

    parent_path = IWA_NUMBERS_DEPENDENCY_CODEC_SOURCE
    child_path = IWA_NUMBERS_DEPENDENCY_INBOUND_SOURCE
    parent_absolute = root / parent_path
    child_absolute = root / child_path
    parent_source = (
        parent_absolute.read_text(encoding="utf-8")
        if parent_absolute.is_file()
        else ""
    )
    parent_modules = _rust_root_level_module_declarations(
        parent_source, frozenset({IWA_NUMBERS_DEPENDENCY_INBOUND_MODULE})
    )
    # The inbound visitor is being staged independently of the removal child;
    # do not report a missing child until its private module declaration lands.
    if not child_absolute.is_file() and not parent_modules:
        return []

    violations: list[str] = []
    if not parent_absolute.is_file():
        violations.append(f"Numbers dependency inbound codec parent is missing: {parent_path}")
        return violations
    if not child_absolute.is_file():
        violations.append(f"Numbers dependency inbound codec child is missing: {child_path}")
        return violations

    parent_code = _mask_rust_non_code(parent_source)
    child_source = _mask_rust_cfg_test_items(child_absolute.read_text(encoding="utf-8"))
    child_code = _mask_rust_non_code(child_source)

    if not parent_modules:
        violations.append(
            "Numbers dependency inbound codec parent is missing private child module "
            f"{IWA_NUMBERS_DEPENDENCY_INBOUND_MODULE}: {parent_path}"
        )
    else:
        for _name, visibility, shape, _body, line in parent_modules:
            if visibility is not None and visibility.strip() == "pub":
                violations.append(
                    "Numbers dependency inbound codec child module must remain private: "
                    f"{parent_path}:{line}"
                )
            if shape == "inline":
                violations.append(
                    "Numbers dependency inbound codec child must use a dedicated source: "
                    f"{parent_path}:{line}"
                )

    for name in IWA_NUMBERS_DEPENDENCY_INBOUND_TYPES:
        if not _rust_canonical_exports(child_source, frozenset({name})):
            violations.append(
                "Numbers dependency inbound codec child is missing public typed fact "
                f"{name}: {child_path}"
            )
        if name not in _rust_canonical_exports(
            parent_source, frozenset({name})
        ):
            violations.append(
                "Numbers dependency inbound codec parent does not re-export typed fact "
                f"{name}: {parent_path}"
            )

    for function_name in IWA_NUMBERS_DEPENDENCY_INBOUND_PARENT_FUNCTIONS:
        if re.search(
            rf"(?m)^\s*pub\s+fn\s+{re.escape(function_name)}\b", parent_code
        ) is None:
            violations.append(
                "Numbers dependency inbound codec parent is missing public fact API "
                f"{function_name}: {parent_path}"
            )

    for marker_label, markers in IWA_NUMBERS_DEPENDENCY_INBOUND_MARKER_GROUPS.items():
        if not _boundary_marker_present(child_code, markers):
            violations.append(
                "Numbers dependency inbound codec child is missing "
                f"{marker_label} marker: {child_path}"
            )

    if re.search(
        r"\b(?:prost|prost_types)\b|\b(?:Message|RawMessage)::(?:decode|encode)",
        child_code,
    ):
        violations.append(f"Numbers dependency inbound codec child must not use Prost: {child_path}")
    for declaration, line in _rust_public_declarations(child_source):
        physical = IWA_DEPENDENCY_CODEC_PUBLIC_PHYSICAL.search(declaration)
        if physical is not None:
            violations.append(
                "Numbers dependency inbound codec child exposes a generated or owned "
                f"wire type {physical.group()}: {child_path}:{line}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                "Numbers dependency inbound codec child exposes raw bytes: "
                f"{child_path}:{line}"
            )

    build_absolute = root / IWA_CODEC_BUILD_SOURCE
    build_marker = 'println!("cargo:rerun-if-changed=src/numbers_table_cell_dependency_codec/inbound.rs");'
    if not build_absolute.is_file():
        violations.append(
            "Numbers dependency inbound codec build source is missing: "
            f"{IWA_CODEC_BUILD_SOURCE}"
        )
    elif build_marker not in build_absolute.read_text(encoding="utf-8"):
        violations.append(
            "Numbers dependency inbound codec build is missing its rerun marker: "
            f"{IWA_CODEC_BUILD_SOURCE}"
        )

    guard_absolute = root / IWA_CODEC_PRODUCTION_GUARD_SOURCE
    guard_marker = 'include_str!("numbers_table_cell_dependency_codec/inbound.rs")'
    if not guard_absolute.is_file():
        violations.append(
            "Numbers dependency inbound codec production inventory source is missing: "
            f"{IWA_CODEC_PRODUCTION_GUARD_SOURCE}"
        )
    elif guard_marker not in guard_absolute.read_text(encoding="utf-8"):
        violations.append(
            "Numbers dependency inbound codec child is missing production inventory "
            f"registration: {IWA_CODEC_PRODUCTION_GUARD_SOURCE}"
        )

    return sorted(set(violations))


def _audit_iwa_numbers_wire_table_owner(
    root: Path,
    *,
    source_path: Path,
    lib_path: Path,
    module_pattern: re.Pattern[str],
    required_types: frozenset[str],
    required_traits: frozenset[str],
    required_functions: frozenset[str],
    required_imports: tuple[re.Pattern[str], ...],
    required_markers: tuple[str, ...],
    peer_import: re.Pattern[str],
    label: str,
) -> list[str]:
    """Audit one low-level Numbers-wire table reader owner."""

    owner_path = root / source_path
    lib_absolute = root / lib_path
    if not owner_path.is_file() and not lib_absolute.is_file():
        return []
    if not owner_path.is_file():
        return [f"Numbers wire {label} owner is missing: {source_path}"]

    violations: list[str] = []
    raw_source = owner_path.read_text(encoding="utf-8")
    code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    if not lib_absolute.is_file():
        violations.append(
            f"Numbers wire {label} owner is missing its crate export: {lib_path}"
        )
    else:
        lib_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(lib_absolute.read_text(encoding="utf-8"))
        )
        if module_pattern.search(lib_code) is None:
            violations.append(
                f"Numbers wire {label} owner is missing its public module export: "
                f"{lib_path}"
            )

    for name in sorted(required_types):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(code) is None:
            violations.append(
                f"Numbers wire {label} owner is missing public {name}: {source_path}"
            )
    for name in sorted(required_traits):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+trait[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(code) is None:
            violations.append(
                f"Numbers wire {label} owner is missing public {name}: {source_path}"
            )
    for name in sorted(required_functions):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:const[ \t]+)?fn[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(code) is None:
            violations.append(
                f"Numbers wire {label} owner is missing public {name}: {source_path}"
            )
    for pattern in required_imports:
        if pattern.search(code) is None:
            violations.append(
                f"Numbers wire {label} owner is missing a required shared import: "
                f"{source_path}"
            )
    for marker in required_markers:
        if re.search(rf"\b{re.escape(marker)}\b", code) is None:
            violations.append(
                f"Numbers wire {label} owner is missing its {marker} route: "
                f"{source_path}"
            )

    for statement in IWA_COMMON_TABLE_READ_IMPORT_STATEMENT.finditer(code):
        concrete_peer = peer_import.search(statement.group())
        if concrete_peer is None:
            continue
        line_number = code.count("\n", 0, statement.start() + concrete_peer.start()) + 1
        violations.append(
            f"Numbers wire {label} owner imports a concrete format peer "
            f"{concrete_peer.group()}: {source_path}:{line_number}"
        )
    return sorted(set(violations))


def audit_iwa_numbers_wire_table_cells_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep selected table-cell storage traversal in the Numbers wire owner."""

    return _audit_iwa_numbers_wire_table_owner(
        root,
        source_path=IWA_NUMBERS_WIRE_TABLE_CELLS_SOURCE,
        lib_path=IWA_NUMBERS_WIRE_TABLE_CELLS_LIB_SOURCE,
        module_pattern=IWA_NUMBERS_WIRE_TABLE_CELLS_MODULE,
        required_types=IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_TYPES,
        required_traits=IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_TRAITS,
        required_functions=frozenset({"read_table_cells"}),
        required_imports=IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_IMPORTS,
        required_markers=IWA_NUMBERS_WIRE_TABLE_CELLS_REQUIRED_MARKERS,
        peer_import=IWA_NUMBERS_WIRE_TABLE_CELLS_PEER_IMPORT,
        label="table-cell",
    )


def audit_iwa_numbers_wire_table_sidecars_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep selected table sidecar traversal and conversion in wire."""

    return _audit_iwa_numbers_wire_table_owner(
        root,
        source_path=IWA_NUMBERS_WIRE_TABLE_SIDECARS_SOURCE,
        lib_path=IWA_NUMBERS_WIRE_TABLE_SIDECARS_LIB_SOURCE,
        module_pattern=IWA_NUMBERS_WIRE_TABLE_SIDECARS_MODULE,
        required_types=IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_TYPES,
        required_traits=IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_TRAITS,
        required_functions=IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_FUNCTIONS,
        required_imports=IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_IMPORTS,
        required_markers=IWA_NUMBERS_WIRE_TABLE_SIDECARS_REQUIRED_MARKERS,
        peer_import=IWA_NUMBERS_WIRE_TABLE_SIDECARS_PEER_IMPORT,
        label="table-sidecar",
    )


def audit_iwa_numbers_wire_formula_names_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep formula-owner identity and category labels in Numbers wire."""

    return _audit_iwa_numbers_wire_table_owner(
        root,
        source_path=IWA_NUMBERS_WIRE_FORMULA_NAMES_SOURCE,
        lib_path=IWA_NUMBERS_WIRE_FORMULA_NAMES_LIB_SOURCE,
        module_pattern=IWA_NUMBERS_WIRE_FORMULA_NAMES_MODULE,
        required_types=IWA_NUMBERS_WIRE_FORMULA_NAMES_REQUIRED_TYPES,
        required_traits=frozenset(),
        required_functions=IWA_NUMBERS_WIRE_FORMULA_NAMES_REQUIRED_FUNCTIONS,
        required_imports=(
            re.compile(
                r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*wire\b"
            ),
            re.compile(
                r"\blitchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*"
                r"group_node_category_codec\b"
            ),
        ),
        required_markers=IWA_NUMBERS_WIRE_FORMULA_NAMES_REQUIRED_MARKERS,
        peer_import=IWA_NUMBERS_WIRE_FORMULA_NAMES_PEER_IMPORT,
        label="formula-name",
    )
