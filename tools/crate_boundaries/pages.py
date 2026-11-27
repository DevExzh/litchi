"""Dependency and public API checks for Pages."""

from __future__ import annotations

from .common import (
    CAMEL_CASE_WORD,
    CARGO_PROST_DEPENDENCY,
    CARGO_SECTION_HEADER,
    IWA_TABLE_HIDDEN_AXES_ENGINE_SOURCE,
    IWA_TABLE_MERGE_TRANSACTION_PUBLIC_FIELD,
    IWA_TABLE_MERGE_TRANSACTION_SPECS,
    PAGES_BODY_CHART_ARRANGEMENT_OWNER_SOURCE,
    PAGES_BODY_CHART_ARRANGEMENT_PACKAGE_MODULE,
    PAGES_PACKAGE_SOURCE,
    PAGES_PUBLIC_USE_STATEMENT,
    PAGES_SOURCE_ROOT,
    PAGES_TABLE_MERGE_COMMON_TABLE_MODULE,
    PAGES_TABLE_MERGE_PACKAGE_SOURCE,
    PAGES_TABLE_MERGE_SOURCE,
    PAGES_TRANSACTION_FAMILIES,
    PAGES_TRANSACTION_FAMILY_BY_OWNER,
    Path,
    ROOT,
    RUST_BYTE_SLICE,
    RUST_IDENTIFIER,
    _audit_focused_chart_arrangement_owner_source_topology,
    _audit_focused_table_cells_source_topology,
    _audit_focused_table_merge_reader_source_topology,
    _audit_focused_table_merge_source_topology,
    _audit_table_merge_transaction_source_topology,
    _boundary_marker_present,
    _iwork_public_leak,
    _mask_rust_cfg_test_items,
    _mask_rust_comments,
    _mask_rust_non_code,
    _pages_masked_production_code,
    _pages_public_module_scope,
    _pages_retired_name_violations,
    _pages_transaction_family_violations,
    _pages_transaction_legacy_names,
    _rust_canonical_exports,
    _rust_function_declarations,
    _rust_impl_headers,
    _rust_named_struct_body,
    _rust_public_declarations,
    _rust_public_method_bodies,
    _rust_public_methods_in_impl,
    _rust_root_level_module_declarations,
    re,
)


PAGES_FOOTNOTE_TEXT_OWNER_SOURCE = PAGES_SOURCE_ROOT / "package" / "footnote_text.rs"
PAGES_FOOTNOTE_TEXT_PACKAGE_METHOD = "edit_body_footnote_text"
RETIRED_IWA_PAGES_FOOTNOTE_LIFECYCLE_METHODS = (
    "insert_body_footnote",
    "remove_body_footnote",
)
RETIRED_IWA_PAGES_FOOTNOTE_FACADE_HELPERS = (
    "focused_body_footnote_package",
    "focused_body_footnote_limits",
    "publish_focused_body_footnote_commit",
)
RETIRED_IWA_PAGES_FOOTNOTE_FACADE_HELPER_SET = frozenset(
    RETIRED_IWA_PAGES_FOOTNOTE_FACADE_HELPERS
)
PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_SOURCE = (
    PAGES_SOURCE_ROOT / "footnote" / "body.rs"
)
PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "body_footnote.rs"
)
PAGES_FOOTNOTE_LIFECYCLE_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "body_footnote"
)
PAGES_FOOTNOTE_LIFECYCLE_IMPLEMENTATION_SOURCES = (
    PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE,
    PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_SOURCE,
)
PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "footnote.rs",
)
PAGES_FOOTNOTE_LIFECYCLE_SELECTOR_SOURCE = (
    PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_SOURCE
)
PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_TYPES = (
    "Footnote",
    "Position",
    "Selector",
)
# ADR 0034: canonical short names defined by the owner and re-exported
# only by ``litchi_pages::footnote::body::transaction``.  The deleted prefixed
# spellings stay focused markers and alias targets; the ``body_footnote`` owner
# path marks
# the canonical re-export.
PAGES_FOOTNOTE_LIFECYCLE_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Snapshot",
)
PAGES_FOOTNOTE_LIFECYCLE_LEGACY_CANONICAL_TYPES = (
    "BodyFootnoteEdit",
    "BodyFootnotePatch",
    "BodyFootnoteCommit",
    "BodyFootnoteDiagnostics",
    "BodyFootnoteError",
    "BodyFootnoteLimitKind",
)
PAGES_FOOTNOTE_LIFECYCLE_SHORT_NAMES = frozenset(
    PAGES_FOOTNOTE_LIFECYCLE_LEGACY_CANONICAL_TYPES
    + PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_TYPES
)
PAGES_FOOTNOTE_LIFECYCLE_PACKAGE_METHODS = (
    "edit_body_footnote",
    "insert_body_footnote",
    "apply_body_footnote",
)
PAGES_FOOTNOTE_LIFECYCLE_FLAT_METHODS = frozenset(
    {
        "remove_body_footnote",
        "insert_footnote",
        "remove_footnote",
        "apply_footnote",
        "edit_footnote",
    }
)
PAGES_FOOTNOTE_LIFECYCLE_FLAT_ALIASES = frozenset(
    {
        "FootnoteEdit",
        "FootnotePatch",
        "FootnoteCommit",
        "FootnoteDiagnostics",
        "FootnoteError",
        "FootnoteLimitKind",
        "BodyFootnoteInsert",
        "BodyFootnoteRemove",
        "BodyFootnoteInsertCommit",
        "BodyFootnoteRemoveCommit",
        "BodyFootnoteTransaction",
    }
)
PAGES_FOOTNOTE_LIFECYCLE_ALIAS_TARGETS = frozenset(
    PAGES_FOOTNOTE_LIFECYCLE_LEGACY_CANONICAL_TYPES
    + PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_TYPES
)
PAGES_FOOTNOTE_LIFECYCLE_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?body_footnote"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_PACKAGE_FOOTNOTE_LIFECYCLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body_footnote\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_FOOTNOTE_LIFECYCLE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_footnote\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_FOOTNOTE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?footnote\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_FOOTNOTE_BODY_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_FOOTNOTE_LIFECYCLE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "DrawableArchive",
        "EntryEdit",
        "ExactArtifacts",
        "FootnoteGraph",
        "FootnoteObjectIds",
        "FootnoteStorageProjection",
        "HostPackage",
        "NativeFootnote",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
    }
)
PAGES_FOOTNOTE_LIFECYCLE_WIRE_TYPES = frozenset(
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
PAGES_FOOTNOTE_LIFECYCLE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tsd", "tsp", "tswp"}
)
# The body-footnote spelling is the public semantic contract used by the
# dormant ratchet.  Keep the shorter aliases for callers of the boundary tool
# that use the wording from the migration notes.
PAGES_BODY_FOOTNOTE_SEMANTIC_SOURCE = PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_SOURCE
PAGES_BODY_FOOTNOTE_OWNER_SOURCE = PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE
PAGES_BODY_FOOTNOTE_OWNER_HELPER_ROOT = PAGES_FOOTNOTE_LIFECYCLE_OWNER_HELPER_ROOT
PAGES_BODY_FOOTNOTE_IMPLEMENTATION_SOURCES = (
    PAGES_FOOTNOTE_LIFECYCLE_IMPLEMENTATION_SOURCES
)
PAGES_BODY_FOOTNOTE_EXPORT_SOURCES = PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES
PAGES_BODY_FOOTNOTE_SELECTOR_SOURCE = PAGES_FOOTNOTE_LIFECYCLE_SELECTOR_SOURCE
PAGES_BODY_FOOTNOTE_SEMANTIC_TYPES = PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_TYPES
PAGES_BODY_FOOTNOTE_CANONICAL_TYPES = PAGES_FOOTNOTE_LIFECYCLE_CANONICAL_TYPES
PAGES_BODY_FOOTNOTE_PACKAGE_METHODS = PAGES_FOOTNOTE_LIFECYCLE_PACKAGE_METHODS
RETIRED_IWA_PAGES_BODY_FOOTNOTE_METHODS = (
    RETIRED_IWA_PAGES_FOOTNOTE_LIFECYCLE_METHODS
)
PAGES_DOCUMENT_PUBLIC_API_SOURCES = (
    PAGES_SOURCE_ROOT / "document.rs",
    PAGES_SOURCE_ROOT / "lib.rs",
)
PAGES_DOCUMENT_DEDICATED_PUBLIC_API_SOURCES = frozenset(
    {PAGES_SOURCE_ROOT / "document.rs"}
)
# ADR 0034 replaced the ``DocumentSourceLimits`` /
# ``ReadError`` ingress family with the root ``OpenOptions``, ``SourceLimits``
# and ``SemanticLimits`` profiles (defined in the private ``limits`` module and
# re-exported from ``lib.rs``).
PAGES_DOCUMENT_PUBLIC_MARKERS = frozenset(
    {
        "Body",
        "DEFAULT_MAX_TEXT_BYTES",
        "Document",
        "DocumentStats",
        "Error",
        "LimitKind",
        "MAX_BODY_STORAGES",
        "MAX_SECTIONS",
        "OpenOptions",
        "Result",
        "Root",
        "SemanticLimitKind",
        "SemanticLimits",
        "SemanticLimitsBuilder",
        "SourceLimitKind",
        "SourceLimits",
        "SourceLimitsBuilder",
    }
)
# Root aliases and ingress types deleted by the iWork rename that have no
# prefix+short spelling.  ``MetadataEdit*`` and ``SectionParagraphs*`` were
# never published but are banned with the other prefixed spellings so the
# families stay uniform.
PAGES_TRANSACTION_RETIRED_EXTRA_NAMES = frozenset(
    {
        "BodyTableHeaderSettings",
        "BodyTableLockState",
        "BodyTableName",
        "BodyTableNameValueError",
        "DocumentReadOptions",
        "DocumentSourceLimitKind",
        "DocumentSourceLimits",
        "DocumentSourceLimitsError",
        "IoKind",
        "ReadError",
        "ReadLimitKind",
        "SemanticLimitsError",
    }
)


PAGES_TRANSACTION_RETIRED_NAMES = frozenset(
    name
    for family in PAGES_TRANSACTION_FAMILIES
    for name in _pages_transaction_legacy_names(family[0])
) | PAGES_TRANSACTION_RETIRED_EXTRA_NAMES
PAGES_BODY_CHART_ARRANGEMENT_PACKAGE_METHODS = (
    "body_chart_arrangement",
    "body_chart_arrangements",
    "edit_body_chart_arrangement",
    "apply_body_chart_arrangement",
)
PAGES_BODY_CHART_ARRANGEMENT_SELECTOR_TYPES = ("BodyChartSelector",)
PAGES_BODY_CHART_ARRANGEMENT_OWNER_MARKERS = (
    "resolve_target",
    "resolve_targets",
    "verify_locality",
)
PAGES_BODY_CHART_ARRANGEMENT_LAZY_CODEC_MARKERS = (
    "chart_arrangement_codec",
    "decode_chart_arrangement_with_report",
    "prepare_chart_arrangement_rewrite",
    "execution_requirements",
    "execute",
)
PAGES_PACKAGE_MANIFEST = Path("crates/litchi-pages/Cargo.toml")


PAGES_PACKAGE_TEST_MODULE = re.compile(
    r"^[ \t]*#[ \t]*\[[ \t]*cfg[ \t]*\([ \t]*test[ \t]*\)[ \t]*\]",
    re.MULTILINE,
)
PAGES_PACKAGE_NO_EAGER_PROST_SOURCE_PATTERNS = (
    (
        "prost::Message",
        re.compile(
            r"(?<![A-Za-z0-9_#])prost[ \t\r\n]*::[ \t\r\n]*Message\b"
        ),
    ),
    (
        "direct generated tswp",
        re.compile(
            r"(?<![A-Za-z0-9_#])"
            r"(?:litchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*"
            r"(?:tswp\b|\{[^;]*\btswp\b)|tswp[ \t\r\n]*::)"
        ),
    ),
    (
        "StorageArchive::decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])StorageArchive[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
    (
        "litchi_iwa_text_wire::from_archive",
        re.compile(
            r"(?<![A-Za-z0-9_#])litchi_iwa_text_wire[ \t\r\n]*::"
            r"[ \t\r\n]*from_archive\b"
        ),
    ),
)
# The package preserves exact source bytes internally, but the supported
# output boundary is the streaming ``Package::write_to`` API.  Keep the
# forbidden names explicit so a public method, alias, or root re-export cannot
# silently recreate the old byte-oriented surface.
PAGES_PACKAGE_OUTPUT_EXPORT_SOURCE = PAGES_SOURCE_ROOT / "lib.rs"
PAGES_PACKAGE_OUTPUT_FORBIDDEN_NAMES = frozenset(
    {"source_bytes", "from_archive_bytes"}
)
PAGES_PACKAGE_OUTPUT_METHOD = "write_to"
PAGES_PACKAGE_OUTPUT_ERROR = "WriteError"
PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "page_layout.rs"
)
PAGES_PAGE_LAYOUT_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "page_layout.rs"
PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCES = (
    PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCE,
    PAGES_PAGE_LAYOUT_SEMANTIC_SOURCE,
)
PAGES_PAGE_LAYOUT_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
)
# ADR 0034 deleted these prefixed spellings (the owner now defines the
# short names, published by ``page_layout::transaction``).  They stay markers
# so a reintroduced root export is still inspected; the ``page_layout`` path
# test in ``_is_pages_page_layout_public_declaration`` marks the canonical
# re-export, and ``audit_pages_transaction_family_source_topology`` bans them.
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
PAGES_TABLE_LOCK_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "table" / "lock.rs"
PAGES_TABLE_LOCK_OWNER_SOURCE = PAGES_SOURCE_ROOT / "package" / "table_lock.rs"
PAGES_TABLE_LOCK_OWNER_HELPER_ROOT = PAGES_SOURCE_ROOT / "package" / "table_lock"
PAGES_TABLE_LOCK_PACKAGE_EXPORT_SOURCE = PAGES_SOURCE_ROOT / "package.rs"
PAGES_TABLE_LOCK_IMPLEMENTATION_SOURCES = (
    PAGES_TABLE_LOCK_OWNER_SOURCE,
    PAGES_TABLE_LOCK_SEMANTIC_SOURCE,
)
PAGES_TABLE_LOCK_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_TABLE_LOCK_PACKAGE_EXPORT_SOURCE,
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
# ADR 0034: ``table::lock`` publishes ``State`` plus the owner's canonical
# short transaction names; the ``BodyTableLockState`` root alias and the
# prefixed ``BodyTableLock*`` spellings are deleted (rejected by the shared
# family helper).  The legacy spellings remain focused markers and alias
# targets because the short names are too common to locate a table-lock
# declaration outside the owner; the ``table_lock``/``table::lock`` owner path
# marks the canonical re-export.
PAGES_TABLE_LOCK_SEMANTIC_TYPES = ("State",)
PAGES_TABLE_LOCK_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
PAGES_TABLE_LOCK_LEGACY_CANONICAL_TYPES = (
    "BodyTableLockEdit",
    "BodyTableLockPatch",
    "BodyTableLockCommit",
    "BodyTableLockDiagnostics",
    "BodyTableLockError",
    "BodyTableLockLimitKind",
)
PAGES_TABLE_LOCK_VALUE_TYPE = "State"
PAGES_TABLE_LOCK_SHORT_NAMES = frozenset(PAGES_TABLE_LOCK_LEGACY_CANONICAL_TYPES)
PAGES_TABLE_LOCK_PUBLIC_NAMES = (
    PAGES_TABLE_LOCK_SHORT_NAMES | frozenset(PAGES_TABLE_LOCK_SEMANTIC_TYPES)
)
PAGES_TABLE_LOCK_PACKAGE_METHODS = (
    "body_table_lock",
    "edit_body_table_lock",
    "apply_body_table_lock",
)
PAGES_TABLE_LOCK_FLAT_METHODS = frozenset(
    {
        "table_lock",
        "edit_table_lock",
        "apply_table_lock",
    }
)
PAGES_TABLE_LOCK_FLAT_ALIASES = frozenset(
    {
        "TableLock",
        "PagesTableLock",
        "TableLockEdit",
        "TableLockPatch",
        "TableLockCommit",
        "TableLockDiagnostics",
        "TableLockError",
        "TableLockLimitKind",
        "TableLockState",
        "TableSelector",
    }
)
# Keep the export-scope matcher focused.  Short semantic names such as
# ``State`` and ``Error`` occur throughout ``package.rs``; only the
# application-owned spellings (or the compatibility aliases above) identify
# a table-lock declaration outside its dedicated owner source.
PAGES_TABLE_LOCK_FOCUSED_MARKERS = frozenset(
    PAGES_TABLE_LOCK_SHORT_NAMES | {"BodyTableLockState"}
)
PAGES_TABLE_LOCK_ALIAS_TARGETS = frozenset(
    PAGES_TABLE_LOCK_LEGACY_CANONICAL_TYPES + ("BodyTableLockState",)
)
PAGES_TABLE_LOCK_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:table_lock|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?lock)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_TABLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_LOCK_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?lock\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_TABLE_LOCK_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table_lock\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_TABLE_LOCK_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?table_lock\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_TABLE_LOCK_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "DrawableArchive",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableInfoArchive",
        "TableLockSnapshot",
    }
)
PAGES_TABLE_LOCK_WIRE_TYPES = frozenset(
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
PAGES_TABLE_LOCK_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tsd", "tsp", "tst", "tswp"}
)
PAGES_TABLE_TITLE_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "table" / "title.rs"
PAGES_TABLE_TITLE_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "body_table_title.rs"
)
PAGES_TABLE_TITLE_OWNER_HELPER_ROOT = PAGES_SOURCE_ROOT / "package" / "body_table_title"
PAGES_TABLE_TITLE_IMPLEMENTATION_SOURCES = (
    PAGES_TABLE_TITLE_OWNER_SOURCE,
    PAGES_TABLE_TITLE_SEMANTIC_SOURCE,
)
PAGES_TABLE_TITLE_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
PAGES_TABLE_TITLE_SELECTOR_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_TABLE_TITLE_SEMANTIC_TYPES = ("Settings",)
# ADR 0034: the owner defines these canonical short names and only
# ``litchi_pages::table::title`` re-exports them (checked by the shared
# family helper).  The deleted prefixed spellings stay focused markers
# and alias targets: the short names are too common in ``lib.rs`` and
# ``package.rs`` to locate a declaration of this family, while the owner
# path marks the canonical re-export.
PAGES_TABLE_TITLE_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
PAGES_TABLE_TITLE_LEGACY_CANONICAL_TYPES = (
    "BodyTableTitleEdit",
    "BodyTableTitlePatch",
    "BodyTableTitleCommit",
    "BodyTableTitleDiagnostics",
    "BodyTableTitleError",
    "BodyTableTitleLimitKind",
)
PAGES_TABLE_TITLE_SHORT_NAMES = frozenset(
    PAGES_TABLE_TITLE_LEGACY_CANONICAL_TYPES + PAGES_TABLE_TITLE_SEMANTIC_TYPES
)
PAGES_TABLE_TITLE_PUBLIC_NAMES = (
    PAGES_TABLE_TITLE_SHORT_NAMES | {"BodyTableTitleSettings"}
)
PAGES_TABLE_TITLE_PACKAGE_METHODS = (
    "body_table_title_settings",
    "edit_body_table_title",
    "apply_body_table_title",
)
PAGES_TABLE_TITLE_FLAT_METHODS = frozenset(
    {"table_title", "edit_table_title", "apply_table_title"}
)
PAGES_TABLE_TITLE_FLAT_ALIASES = frozenset(
    {
        "TitleSettings",
        "TableTitleSettings",
        "TableTitleEdit",
        "TableTitlePatch",
        "TableTitleCommit",
        "TableTitleDiagnostics",
        "TableTitleError",
        "TableTitleLimitKind",
        "TitleEdit",
        "TitlePatch",
        "TitleCommit",
        "TitleDiagnostics",
        "TitleError",
        "TitleLimitKind",
    }
)
PAGES_TABLE_TITLE_ALIAS_TARGETS = frozenset(
    PAGES_TABLE_TITLE_LEGACY_CANONICAL_TYPES + ("Settings",)
)
PAGES_TABLE_TITLE_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:body_table_title|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?title)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_TABLE_TITLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?title\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_TABLE_TITLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body_table_title\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_TABLE_TITLE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_table_title\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_TABLE_TITLE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "DrawableArchive",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableInfoArchive",
        "TableTitleArchive",
        "TableTitleSnapshot",
    }
)
PAGES_TABLE_TITLE_WIRE_TYPES = frozenset(
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
PAGES_TABLE_TITLE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tsd", "tsp", "tst", "tswp"}
)
RETIRED_IWA_PAGES_TABLE_NAME_HELPER = "rename_table_in_package"
PAGES_TABLE_NAME_OWNER_SOURCE = PAGES_SOURCE_ROOT / "package" / "body_table_name.rs"
PAGES_TABLE_NAME_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
PAGES_TABLE_NAME_SELECTOR_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_PACKAGE_TABLE_NAME_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_table_name\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_TABLE_NAME_PACKAGE_METHODS = (
    "body_table_name",
    "edit_body_table_name",
    "apply_body_table_name",
)
PAGES_TABLE_HEADERS_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "table" / "headers.rs"
PAGES_TABLE_HEADERS_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "body_table_headers.rs"
)
PAGES_TABLE_HEADERS_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "body_table_headers"
)
PAGES_TABLE_HEADERS_IMPLEMENTATION_SOURCES = (
    PAGES_TABLE_HEADERS_OWNER_SOURCE,
    PAGES_TABLE_HEADERS_SEMANTIC_SOURCE,
)
PAGES_TABLE_HEADERS_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
PAGES_TABLE_HEADERS_SELECTOR_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_TABLE_HEADERS_SEMANTIC_TYPES = ("Count", "Error", "Settings")
# ADR 0034: the owner defines these canonical short names and only
# ``litchi_pages::table::headers::transaction`` re-exports them (checked by the shared
# family helper).  The deleted prefixed spellings stay focused markers
# and alias targets: the short names are too common in ``lib.rs`` and
# ``package.rs`` to locate a declaration of this family, while the owner
# path marks the canonical re-export.
PAGES_TABLE_HEADERS_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "InvalidReason",
    "LimitKind",
)
PAGES_TABLE_HEADERS_LEGACY_CANONICAL_TYPES = (
    "BodyTableHeaderSettingsEdit",
    "BodyTableHeaderSettingsPatch",
    "BodyTableHeaderSettingsCommit",
    "BodyTableHeaderSettingsDiagnostics",
    "BodyTableHeaderSettingsError",
    "BodyTableHeaderSettingsInvalidReason",
    "BodyTableHeaderSettingsLimitKind",
)
PAGES_TABLE_HEADERS_SHORT_NAMES = frozenset(
    PAGES_TABLE_HEADERS_LEGACY_CANONICAL_TYPES + PAGES_TABLE_HEADERS_SEMANTIC_TYPES
)
PAGES_TABLE_HEADERS_FOCUSED_MARKERS = frozenset(
    PAGES_TABLE_HEADERS_LEGACY_CANONICAL_TYPES + ("Count", "Settings", "BodyTableHeaderSettings")
)
PAGES_TABLE_HEADERS_PUBLIC_NAMES = (
    PAGES_TABLE_HEADERS_SHORT_NAMES | {"BodyTableHeaderSettings"}
)
PAGES_TABLE_HEADERS_PACKAGE_METHODS = (
    "body_table_header_settings",
    "edit_body_table_header_settings",
    "apply_body_table_header_settings",
)
PAGES_TABLE_HEADERS_FLAT_METHODS = frozenset(
    {"table_header_settings", "edit_table_headers", "apply_table_headers"}
)
PAGES_TABLE_HEADERS_FLAT_ALIASES = frozenset(
    {
        "HeaderSettings",
        "TableHeaderSettings",
        "TableHeaders",
        "TableHeadersEdit",
        "TableHeadersPatch",
        "TableHeadersCommit",
        "TableHeadersDiagnostics",
        "TableHeadersError",
        "TableHeadersLimitKind",
        "HeaderEdit",
        "HeaderPatch",
        "HeaderCommit",
        "HeaderDiagnostics",
        "HeaderError",
        "HeaderLimitKind",
    }
)
PAGES_TABLE_HEADERS_ALIAS_TARGETS = frozenset(
    PAGES_TABLE_HEADERS_LEGACY_CANONICAL_TYPES + ("Count", "Settings")
)
PAGES_TABLE_HEADERS_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:body_table_headers|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?headers)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_TABLE_HEADERS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?headers\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_TABLE_HEADERS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body_table_headers\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_TABLE_HEADERS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_table_headers\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_TABLE_HEADERS_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "DrawableArchive",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableHeaderSettingsArchive",
        "TableHeaderSettingsSnapshot",
        "TableHeadersArchive",
    }
)
PAGES_TABLE_HEADERS_WIRE_TYPES = frozenset(
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
PAGES_TABLE_HEADERS_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tsd", "tsp", "tst", "tswp"}
)
# Body-table dimensions are the remaining layout value family in the Pages
# host.  Keep the semantic value archive-free and give the package owner its
# own namespace so a future root glob cannot accidentally publish the private
# transaction implementation.
PAGES_TABLE_DIMENSION_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "table" / "dimension.rs"
PAGES_TABLE_DIMENSION_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "body_table_dimension.rs"
)
PAGES_TABLE_DIMENSION_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "body_table_dimension"
)
PAGES_TABLE_DIMENSION_IMPLEMENTATION_SOURCES = (
    PAGES_TABLE_DIMENSION_OWNER_SOURCE,
    PAGES_TABLE_DIMENSION_SEMANTIC_SOURCE,
)
PAGES_TABLE_DIMENSION_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
PAGES_TABLE_DIMENSION_SELECTOR_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_TABLE_DIMENSION_SEMANTIC_TYPES = ("Dimension", "Points", "Size")
# ADR 0034: the owner defines these canonical short names and only
# ``litchi_pages::table::dimension::transaction`` re-exports them (checked by the shared
# family helper).  The deleted prefixed spellings stay focused markers
# and alias targets: the short names are too common in ``lib.rs`` and
# ``package.rs`` to locate a declaration of this family, while the owner
# path marks the canonical re-export.
PAGES_TABLE_DIMENSION_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
PAGES_TABLE_DIMENSION_LEGACY_CANONICAL_TYPES = (
    "BodyTableDimensionEdit",
    "BodyTableDimensionPatch",
    "BodyTableDimensionCommit",
    "BodyTableDimensionDiagnostics",
    "BodyTableDimensionError",
    "BodyTableDimensionLimitKind",
)
PAGES_TABLE_DIMENSION_SHORT_NAMES = frozenset(
    PAGES_TABLE_DIMENSION_LEGACY_CANONICAL_TYPES + PAGES_TABLE_DIMENSION_SEMANTIC_TYPES
)
PAGES_TABLE_DIMENSION_PUBLIC_NAMES = PAGES_TABLE_DIMENSION_SHORT_NAMES
PAGES_TABLE_DIMENSION_PACKAGE_METHODS = (
    "body_table_dimension_size",
    "edit_body_table_dimension_size",
    "apply_body_table_dimension_size",
)
PAGES_TABLE_DIMENSION_FLAT_METHODS = frozenset(
    {"table_dimension_size", "edit_table_dimension_size", "apply_table_dimension_size"}
)
PAGES_TABLE_DIMENSION_FLAT_ALIASES = frozenset(
    {
        "DimensionEdit",
        "DimensionPatch",
        "DimensionCommit",
        "DimensionDiagnostics",
        "DimensionError",
        "DimensionLimitKind",
        "TableDimensionEdit",
        "TableDimensionPatch",
        "TableDimensionCommit",
        "TableDimensionDiagnostics",
        "TableDimensionError",
        "TableDimensionLimitKind",
        "TableDimensionSize",
        "TableDimensionSizeEdit",
        "TableDimensionSizePatch",
        "TableDimensionSizeCommit",
        "TableDimensionSizeDiagnostics",
        "TableDimensionSizeError",
        "TableDimensionSizeLimitKind",
    }
)
PAGES_TABLE_DIMENSION_ALIAS_TARGETS = frozenset(
    PAGES_TABLE_DIMENSION_LEGACY_CANONICAL_TYPES + PAGES_TABLE_DIMENSION_SEMANTIC_TYPES
)
PAGES_TABLE_DIMENSION_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:body_table_dimension|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?dimension)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_TABLE_DIMENSION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?dimension\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_TABLE_DIMENSION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body_table_dimension\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_TABLE_DIMENSION_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_table_dimension\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_TABLE_DIMENSION_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "DrawableArchive",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "TableDimensionArchive",
        "TableDimensionSnapshot",
        "TableModelArchive",
        "TSTHeaderStorageBucket",
    }
)
PAGES_TABLE_DIMENSION_WIRE_TYPES = frozenset(
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
PAGES_TABLE_DIMENSION_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tst", "tsp", "tswp"}
)

PAGES_TABLE_SORT_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "table" / "sort.rs"
PAGES_TABLE_SORT_OWNER_SOURCE = PAGES_SOURCE_ROOT / "package" / "body_table_sort.rs"
PAGES_TABLE_SORT_OWNER_HELPER_ROOT = PAGES_SOURCE_ROOT / "package" / "body_table_sort"
PAGES_TABLE_SORT_IMPLEMENTATION_SOURCES = (
    PAGES_TABLE_SORT_SEMANTIC_SOURCE,
    PAGES_TABLE_SORT_OWNER_SOURCE,
)
PAGES_TABLE_SORT_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
PAGES_TABLE_SORT_SELECTOR_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_TABLE_SORT_SEMANTIC_TYPES = (
    "ColumnIndex",
    "Direction",
    "Order",
    "RowRange",
    "Rule",
    "Scope",
)
PAGES_TABLE_SORT_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
# ADR 0034: the owner defines these canonical short names and only
# ``litchi_pages::table::sort::transaction`` re-exports them (checked by the shared
# family helper).  The deleted prefixed spellings stay focused markers
# and alias targets: the short names are too common in ``lib.rs`` and
# ``package.rs`` to locate a declaration of this family, while the owner
# path marks the canonical re-export.
PAGES_TABLE_SORT_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
PAGES_TABLE_SORT_LEGACY_CANONICAL_TYPES = (
    "BodyTableSortEdit",
    "BodyTableSortPatch",
    "BodyTableSortCommit",
    "BodyTableSortDiagnostics",
    "BodyTableSortError",
    "BodyTableSortLimitKind",
)
PAGES_TABLE_SORT_SHORT_NAMES = frozenset(
    PAGES_TABLE_SORT_SEMANTIC_TYPES + PAGES_TABLE_SORT_TRANSACTION_TYPES
)
PAGES_TABLE_SORT_PUBLIC_NAMES = frozenset(
    PAGES_TABLE_SORT_SEMANTIC_TYPES + PAGES_TABLE_SORT_LEGACY_CANONICAL_TYPES
)
PAGES_TABLE_SORT_PACKAGE_METHODS = (
    "body_table_sort_order",
    "edit_body_table_sort_order",
    "apply_body_table_sort_order",
)
PAGES_TABLE_SORT_FLAT_METHODS = frozenset(
    {"table_sort_order", "edit_table_sort_order", "apply_table_sort_order"}
)
PAGES_TABLE_SORT_FLAT_ALIASES = frozenset(
    {
        "SortEdit",
        "SortPatch",
        "SortCommit",
        "SortDiagnostics",
        "SortError",
        "SortLimitKind",
        "SortPath",
        "TableSortEdit",
        "TableSortPatch",
        "TableSortCommit",
        "TableSortDiagnostics",
        "TableSortError",
        "TableSortLimitKind",
        "TableSortPath",
        "TableSortOrder",
        "TableSortOrderEdit",
        "TableSortOrderPatch",
        "TableSortOrderCommit",
        "TableSortOrderDiagnostics",
        "TableSortOrderError",
        "TableSortOrderLimitKind",
        "TableSortOrderPath",
    }
)
PAGES_TABLE_SORT_ALIAS_TARGETS = frozenset(
    PAGES_TABLE_SORT_LEGACY_CANONICAL_TYPES + PAGES_TABLE_SORT_SEMANTIC_TYPES
)
PAGES_TABLE_SORT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:body_table_sort|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?sort)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_PACKAGE_TABLE_SORT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body_table_sort\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_TABLE_SORT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_table_sort\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_SORT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?sort\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_SORT_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
# The format-neutral public spelling is an inline alias in lib.rs; the strict
# implementation remains the already-fuzzed Numbers source so Pages cannot
# fork the field-44 wire policy.
PAGES_TABLE_SORT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_sort_order_codec.rs"
)
PAGES_TABLE_SORT_CODEC_PUBLIC_SOURCE = Path("crates/litchi-iwa-protos/src/lib.rs")
PAGES_TABLE_SORT_CODEC_MODULE = "table_sort_order_codec"
PAGES_TABLE_SORT_CODEC_REQUIRED_APIS = (
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
PAGES_TABLE_SORT_CODEC_REQUIRED_MARKERS = (
    "unknown",
    "unknown overlong",
    "report",
    "SORT_ORDER_FIELD",
    "SORT_RULES_FIELD",
)
PAGES_TABLE_SORT_CODEC_FUZZ_SOURCE = Path(
    "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_sort_order_codec.rs"
)
PAGES_TABLE_SORT_CODEC_FUZZ_CORPUS = Path(
    "crates/litchi-iwa-protos/fuzz/corpus/numbers_table_sort_order_codec"
)
PAGES_TABLE_SORT_FUZZ_SOURCE = Path(
    "crates/litchi/fuzz/fuzz_targets/pages_body_table_sort_order.rs"
)
PAGES_TABLE_SORT_FUZZ_CORPUS = Path(
    "crates/litchi/fuzz/corpus/pages_body_table_sort_order"
)
PAGES_TABLE_SORT_PHYSICAL_TYPES = frozenset(
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
PAGES_TABLE_SORT_WIRE_TYPES = frozenset(
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
PAGES_TABLE_SORT_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tn", "tsp", "tst", "tswp"}
)
PAGES_TABLE_SORT_OWNER_REQUIRED_MARKERS = {
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
    "reassembly": re.compile(
        r"(?<![A-Za-z0-9])(?:reassembl\w*|zip|snappy|compress)(?![A-Za-z0-9])"
    ),
    "atomic/inverse": re.compile(
        r"(?<![A-Za-z0-9])(?:atomic|inverse|conflict|source_bytes|unchanged)"
        r"(?![A-Za-z0-9])"
    ),
}
PAGES_TABLE_SORT_RAW_PARAMETER = re.compile(
    r"\b(?:table_id|object_id|model_id|native_id|source_bytes|bytes|row|column)"
    r"\b[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])\b"
)

PAGES_HEADER_FOOTER_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "header_footer.rs"
PAGES_HEADER_FOOTER_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "header_footer_text.rs"
)
PAGES_HEADER_FOOTER_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "header_footer_text"
)
PAGES_HEADER_FOOTER_IMPLEMENTATION_SOURCES = (
    PAGES_HEADER_FOOTER_OWNER_SOURCE,
    PAGES_HEADER_FOOTER_SEMANTIC_SOURCE,
)
PAGES_HEADER_FOOTER_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_HEADER_FOOTER_SEMANTIC_SOURCE,
)
PAGES_HEADER_FOOTER_SELECTOR_SOURCE = PAGES_HEADER_FOOTER_SEMANTIC_SOURCE
PAGES_HEADER_FOOTER_SEMANTIC_TYPES = ("HeaderFooter", "Template", "Kind")
PAGES_HEADER_FOOTER_SELECTOR_TYPES = ("HeaderFooterSelector",)
# ADR 0034: the owner defines the canonical short names and only
# ``litchi_pages::header_footer::text`` re-exports them.  The deleted prefixed
# spellings stay in the focused-marker and alias-target sets below, so a
# declaration that reintroduces one outside the owner is still inspected; the
# short names are too common in ``lib.rs``/``package.rs`` to serve as markers
# (the ``header_footer_text`` owner path marks the canonical re-export).
PAGES_HEADER_FOOTER_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
PAGES_HEADER_FOOTER_LEGACY_CANONICAL_TYPES = (
    "HeaderFooterTextEdit",
    "HeaderFooterTextPatch",
    "HeaderFooterTextCommit",
    "HeaderFooterTextDiagnostics",
    "HeaderFooterTextError",
    "HeaderFooterTextLimitKind",
)
PAGES_HEADER_FOOTER_SHORT_NAMES = frozenset(
    PAGES_HEADER_FOOTER_LEGACY_CANONICAL_TYPES
    + PAGES_HEADER_FOOTER_SEMANTIC_TYPES
    + PAGES_HEADER_FOOTER_SELECTOR_TYPES
)
PAGES_HEADER_FOOTER_PUBLIC_NAMES = PAGES_HEADER_FOOTER_SHORT_NAMES
PAGES_HEADER_FOOTER_PACKAGE_METHODS = (
    "header_footers",
    "edit_header_footer_text",
    "apply_header_footer_text",
)
PAGES_HEADER_FOOTER_FLAT_METHODS = frozenset(
    {
        "header_footer_text",
        "edit_header_footer",
        "apply_header_footer",
        "replace_header_footer_text",
        "set_header_footer_text",
        "clear_header_footer",
    }
)
PAGES_HEADER_FOOTER_FLAT_ALIASES = frozenset(
    {
        "HeaderFooterText",
        "HeaderFooterEdit",
        "HeaderFooterPatch",
        "HeaderFooterCommit",
        "HeaderFooterDiagnostics",
        "HeaderFooterError",
        "HeaderFooterLimitKind",
        "HeaderFooterRegion",
        "HeaderFooterSlot",
        "HeaderFooterInfo",
        "PagesHeaderFooterInfo",
        "TextStorageId",
    }
)
PAGES_HEADER_FOOTER_ALIAS_TARGETS = frozenset(
    PAGES_HEADER_FOOTER_LEGACY_CANONICAL_TYPES
    + PAGES_HEADER_FOOTER_SEMANTIC_TYPES
    + PAGES_HEADER_FOOTER_SELECTOR_TYPES
)
PAGES_HEADER_FOOTER_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:header_footer_text|header_footer|"
    r"header[ \t\r\n]*footer)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_HEADER_FOOTER_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?header_footer\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_HEADER_FOOTER_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?header_footer_text\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_HEADER_FOOTER_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?header_footer_text\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_HEADER_FOOTER_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HeaderFooterArchive",
        "HeaderFooterLocation",
        "HeaderFooterSnapshot",
        "HostPackage",
        "RawMessage",
        "Resolved",
        "SectionArchive",
        "SectionTemplateArchive",
        "SnappyStream",
        "SourceCatalog",
        "StorageArchive",
        "TextStorage",
        "TextStorageId",
        "PagesHeaderFooterInfo",
    }
)
PAGES_HEADER_FOOTER_WIRE_TYPES = frozenset(
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
PAGES_HEADER_FOOTER_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tsd", "tsp", "tswp"}
)

PAGES_SECTION_TEXT_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_SECTION_TEXT_OWNER_SOURCE = PAGES_SOURCE_ROOT / "package" / "section_text.rs"
PAGES_SECTION_TEXT_OWNER_HELPER_ROOT = PAGES_SOURCE_ROOT / "package" / "section_text"
PAGES_SECTION_TEXT_IMPLEMENTATION_SOURCES = (
    PAGES_SECTION_TEXT_SEMANTIC_SOURCE,
    PAGES_SECTION_TEXT_OWNER_SOURCE,
)
PAGES_SECTION_TEXT_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SECTION_TEXT_SEMANTIC_SOURCE,
)
PAGES_SECTION_TEXT_SELECTOR_SOURCE = PAGES_SECTION_TEXT_SEMANTIC_SOURCE
PAGES_SECTION_TEXT_SEMANTIC_TYPES = ("SectionSelector",)
# ADR 0034: canonical short names defined by the owner and re-exported
# only by ``litchi_pages::section::text``.  The deleted prefixed spellings
# remain focused markers/alias targets (they are also flat aliases below); the
# ``section_text`` owner path marks the canonical re-export.
PAGES_SECTION_TEXT_CANONICAL_TYPES = (
    "Commit",
    "Diagnostics",
    "Edit",
    "Error",
    "LimitKind",
    "Patch",
)
PAGES_SECTION_TEXT_LEGACY_CANONICAL_TYPES = (
    "SectionTextCommit",
    "SectionTextDiagnostics",
    "SectionTextEdit",
    "SectionTextError",
    "SectionTextLimitKind",
    "SectionTextPatch",
)
PAGES_SECTION_TEXT_SHORT_NAMES = frozenset(
    PAGES_SECTION_TEXT_LEGACY_CANONICAL_TYPES + PAGES_SECTION_TEXT_SEMANTIC_TYPES
)
PAGES_SECTION_TEXT_PUBLIC_NAMES = PAGES_SECTION_TEXT_SHORT_NAMES
PAGES_SECTION_TEXT_ALLOWED_PUBLIC_IDENTIFIERS = frozenset(
    {
        "AmbiguousSelector",
        "Allocation",
        "BodySectionCount",
        "DependentContent",
        "FootnoteAnchorReplacement",
        "InvalidSource",
        "LimitExceeded",
        "NameNotFound",
        "ObjectMarkerReplacement",
        "OperationAlreadyStaged",
        "PatchConflict",
        "PositionNotFound",
        "SectionBreakReplacement",
        "SpanOutOfBounds",
        "SurrogateBoundary",
        "UnsupportedSource",
        "Verification",
        "WireBytes",
        "WireFields",
        "WireNesting",
        "WireWork",
    }
)
PAGES_SECTION_TEXT_PACKAGE_METHODS = (
    "section_text",
    "edit_section_text",
    "set_section_text",
    "clear_section_text",
    "apply_section_text",
)
PAGES_SECTION_TEXT_FLAT_METHODS = frozenset(
    {
        "text",
        "body_text",
        "edit_text",
        "set_text",
        "clear_text",
        "replace_text",
        "apply_text",
    }
)
PAGES_SECTION_TEXT_FLAT_ALIASES = frozenset(
    {
        "SectionText",
        "SectionBodyText",
        "SectionTextValue",
        "SectionTextRegion",
        "SectionTextInfo",
        "SectionTextEdit",
        "SectionTextPatch",
        "SectionTextCommit",
        "SectionTextDiagnostics",
        "SectionTextError",
        "SectionTextLimitKind",
        "PagesSectionText",
        "PagesSectionTextEdit",
        "PagesSectionTextPatch",
        "PagesSectionTextCommit",
        "PagesSectionTextDiagnostics",
        "PagesSectionTextError",
        "PagesSectionTextLimitKind",
        "TextStorageId",
    }
)
PAGES_SECTION_TEXT_ALIAS_TARGETS = frozenset(
    PAGES_SECTION_TEXT_LEGACY_CANONICAL_TYPES + PAGES_SECTION_TEXT_SEMANTIC_TYPES
)
PAGES_SECTION_TEXT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:section_text|section[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?text)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_SECTION_TEXT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?selector\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_SECTION_TEXT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?section_text\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_SECTION_TEXT_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?section_text\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_SECTION_TEXT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "NativeSectionReference",
        "RawMessage",
        "Resolved",
        "SectionArchive",
        "SectionTextArchive",
        "SectionTextSnapshot",
        "SnappyStream",
        "SourceCatalog",
        "Storage",
        "StorageArchive",
        "StorageWireLimitsError",
        "TextStorage",
        "TextStorageId",
    }
)
PAGES_SECTION_TEXT_WIRE_TYPES = frozenset(
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
PAGES_SECTION_TEXT_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tp", "tsd", "tsp", "tswp"}
)
PAGES_SECTION_TEXT_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|section|storage|text|native|archive|message|component|"
    r"entry|metadata|package|uuid)[A-Za-z_]*(?:id|identifier))"
    r"[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)

PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE = PAGES_SOURCE_ROOT / "table" / "appearance.rs"
PAGES_TABLE_APPEARANCE_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "body_table_appearance.rs"
)
PAGES_TABLE_APPEARANCE_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "body_table_appearance"
)
PAGES_TABLE_APPEARANCE_IMPLEMENTATION_SOURCES = (
    PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE,
    PAGES_TABLE_APPEARANCE_OWNER_SOURCE,
)
PAGES_TABLE_APPEARANCE_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
PAGES_TABLE_APPEARANCE_SELECTOR_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES = (
    "Appearance",
    "Banding",
    "GridlineVisibility",
    "Gridlines",
    "RowSizing",
)
PAGES_TABLE_APPEARANCE_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
# ADR 0034: the owner defines the canonical short names (identical to the
# ``transaction`` re-export set) and only
# ``litchi_pages::table::appearance::transaction`` publishes them.  The
# deleted prefixed spellings stay focused markers and alias targets.
PAGES_TABLE_APPEARANCE_CANONICAL_TYPES = PAGES_TABLE_APPEARANCE_TRANSACTION_TYPES
PAGES_TABLE_APPEARANCE_LEGACY_CANONICAL_TYPES = (
    "BodyTableAppearanceEdit",
    "BodyTableAppearancePatch",
    "BodyTableAppearanceCommit",
    "BodyTableAppearanceDiagnostics",
    "BodyTableAppearanceError",
    "BodyTableAppearanceLimitKind",
    "BodyTableAppearancePath",
)
PAGES_TABLE_APPEARANCE_SHORT_NAMES = frozenset(
    PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES
    + PAGES_TABLE_APPEARANCE_TRANSACTION_TYPES
)
PAGES_TABLE_APPEARANCE_PUBLIC_NAMES = frozenset(
    PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES
    + PAGES_TABLE_APPEARANCE_LEGACY_CANONICAL_TYPES
    + ("BodyTableSelector",)
)
# Public names that may legitimately be the target spelling of a ``pub use``
# rename (none of the renamed transaction names may be renamed again).
PAGES_TABLE_APPEARANCE_ALIAS_EXEMPT_NAMES = frozenset(
    PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES + ("BodyTableSelector",)
)
PAGES_TABLE_APPEARANCE_PACKAGE_METHODS = (
    "body_table_appearance",
    "edit_body_table_appearance",
    "apply_body_table_appearance",
)
PAGES_TABLE_APPEARANCE_EDIT_METHODS = frozenset({"set", "commit"})
PAGES_TABLE_APPEARANCE_FLAT_ALIASES = frozenset(
    {
        "TableAppearance",
        "TableAppearanceCommit",
        "TableAppearanceDiagnostics",
        "TableAppearanceEdit",
        "TableAppearanceError",
        "TableAppearanceLimitKind",
        "TableAppearancePatch",
        "TableAppearancePath",
        "AppearanceCommit",
        "AppearanceDiagnostics",
        "AppearanceEdit",
        "AppearanceError",
        "AppearanceLimitKind",
        "AppearancePatch",
        "AppearancePath",
    }
)
PAGES_TABLE_APPEARANCE_ALIAS_TARGETS = frozenset(
    PAGES_TABLE_APPEARANCE_LEGACY_CANONICAL_TYPES + PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES
)
PAGES_TABLE_APPEARANCE_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:body_table_appearance|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?appearance)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_PACKAGE_TABLE_APPEARANCE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body_table_appearance\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_TABLE_APPEARANCE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_table_appearance\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_APPEARANCE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?appearance\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?table\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_APPEARANCE_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_TABLE_APPEARANCE_PHYSICAL_TYPES = frozenset(
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
        "TableModelArchive",
        "TableStyleArchive",
        "TableStylePresetArchive",
        "TableStyleNetworkArchive",
        "TableStylePropertiesArchive",
        "StylesheetArchive",
    }
)
PAGES_TABLE_APPEARANCE_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeOptions",
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
PAGES_TABLE_APPEARANCE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tst", "tsp", "tswp", "litchi_iwa_protos"}
)
PAGES_TABLE_APPEARANCE_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/table_appearance_codec.rs"
)
PAGES_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE = Path("crates/litchi-iwa-protos/src/lib.rs")
PAGES_TABLE_APPEARANCE_CODEC_MODULE = "table_appearance_codec"
PAGES_TABLE_APPEARANCE_CODEC_REQUIRED_APIS = (
    "DecodeOptions",
    "DecodeError",
    "AppearanceSnapshot",
    "TableModelSnapshot",
    "PreparedRewrite",
    "PreparedTableModelStyleRewrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "decode_table_model_with_report",
    "prepare_table_model_style_rewrite",
    "canonical_table_style_variation",
    "append_stylesheet_style",
    "resolve_table_style_appearance",
    "decode_table_style_preset_with_report",
    "decode_table_style_network_with_report",
)
PAGES_TABLE_APPEARANCE_CODEC_REQUIRED_MARKER_GROUPS = {
    "canonical unknown preservation": ("scan_group_fields", "find_group_end", "copy_from_slice"),
    "prepared rewrite contract": ("prepare_table_model_style_rewrite", "execution_requirements", "execute"),
    "bounded decode axes": ("max_input_bytes", "max_output_bytes", "max_fields", "max_work_bytes", "max_allocations"),
}
PAGES_TABLE_APPEARANCE_TEST_SOURCES = (
    Path("crates/litchi-pages/tests/body_table_appearance.rs"),
)
PAGES_TABLE_APPEARANCE_FUZZ_SOURCES = (
    Path("crates/litchi-iwa-protos/fuzz/fuzz_targets/table_appearance.rs"),
    Path("crates/litchi/fuzz/fuzz_targets/pages_body_table_appearance.rs"),
)
PAGES_TABLE_APPEARANCE_FUZZ_CORPORA = (
    Path("crates/litchi-iwa-protos/fuzz/corpus/table_appearance"),
    Path("crates/litchi/fuzz/corpus/pages_body_table_appearance"),
)
PAGES_TABLE_APPEARANCE_OWNER_MARKER_GROUPS = {
    "aggregate transaction budget": ("metadata_options", "WireLimits", "effective_archive_limits", "snappy_limits"),
    "bounded resource axes": ("max_references", "max_iwa_stream_bytes", "effective_archive_limits", "snappy_limits", "try_reserve_exact"),
    "strict prepared codec": ("appearance_codec", "prepare_table_model_style_rewrite", "execution_requirements", "exact_limits"),
    "copy-on-write style graph": ("new_style", "replace_model_in_archive", "append_style_object", "replace_stylesheet_in_archive"),
    "metadata ownership": ("MetadataFacts", "inspect_package_metadata_with_visitor", "uuid_additions", "SaveTokenBatch", "external_additions"),
    "prepared execution accounting": ("model_requirements", "stylesheet_requirements", "execution_requirements", "execute"),
    "package reassembly": ("prepare_reassembly_with_deletions", "EntryEdit", "execute"),
    # ADR 0034: the owner's patch is the canonical short ``Patch``.
    "exact artifacts and inverse": ("impl Patch", "source_fingerprint", "inverse", "is_noop", "PatchConflict"),
    "candidate reopen and locality": ("candidate", "Package::from_source_with_options", "verify_candidate_locality", "full_reparse_performed"),
    "preview preservation": ("deleted_previews", "root_preview_deletions"),
}
# A Pages owner may use a more specific private spelling for the same seam
# (for example ``MetadataSelectorSet`` instead of the generic ``MetadataFacts``
# name, or ``from_source_catalog`` instead of the convenience constructor).
# Keep the public ratchet strict while recognizing those equivalent private
# implementations.  Test scaffolds generated from the canonical groups still
# exercise the missing-marker branches directly.
PAGES_TABLE_APPEARANCE_OWNER_MARKER_ALTERNATIVES = {
    "strict prepared codec": (
        "codec",
        "prepare_table_model_style_rewrite",
        "execution_requirements",
        "exact_limits",
    ),
    "metadata ownership": (
        "MetadataSelectorSet",
        "inspect_package_metadata_with_visitor",
        "uuid_additions",
        "SaveTokenBatch",
        "external_additions",
    ),
    "prepared execution accounting": (
        "model_plan",
        "stylesheet_plan",
        "execution_requirements",
        "execute",
    ),
    "candidate reopen and locality": (
        "candidate",
        "from_source_catalog",
        "verify_candidate_locality",
        "full_reparse_performed",
    ),
}
PAGES_TABLE_APPEARANCE_FORBIDDEN_OWNER_PATTERNS = (
    (
        "row/cell movement",
        re.compile(
            r"\b(?:move|reorder|sort)_?(?:rows?|cells?)\b|"
            r"\b(?:rows?|cells?)_(?:move|reorder|sort)\b",
            re.I,
        ),
    ),
    (
        "formula/tile mutation",
        re.compile(
            r"\b(?:cell|cells?|formula|formulas?|tile|tiles?|data_store|data_stores?)_"
            r"(?:rewrite|update|move|mutate|reorder|sort)\b|"
            r"\b(?:rewrite|update|move|mutate|reorder|sort)_"
            r"(?:cell|cells?|formula|formulas?|tile|tiles?|data_store|data_stores?)\b",
            re.I,
        ),
    ),
    (
        "preview deletion",
        re.compile(
            r"\b(?:delete|remove|drop|rewrite)_?(?:root_)?previews?\b|"
            r"\b(?:root_)?previews?_(?:delete|remove|drop|rewrite)\b",
            re.I,
        ),
    ),
    (
        "one-shot style rewrite",
        re.compile(r"\b(?:rewrite_table_model_style|append_stylesheet_style)\s*\("),
    ),
)
PAGES_TABLE_APPEARANCE_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)

# Body-table spelling is also useful to callers which use the longer
# body-table name.  Keep these aliases as constants, not duplicate inventories.
PAGES_BODY_TABLE_APPEARANCE_SEMANTIC_SOURCE = PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE
PAGES_BODY_TABLE_APPEARANCE_OWNER_SOURCE = PAGES_TABLE_APPEARANCE_OWNER_SOURCE
PAGES_BODY_TABLE_APPEARANCE_EXPORT_SOURCES = PAGES_TABLE_APPEARANCE_EXPORT_SOURCES
PAGES_BODY_TABLE_APPEARANCE_PACKAGE_METHODS = PAGES_TABLE_APPEARANCE_PACKAGE_METHODS
PAGES_BODY_TABLE_APPEARANCE_CANONICAL_TYPES = PAGES_TABLE_APPEARANCE_CANONICAL_TYPES

IWA_PAGES_README_TABLE_HIDDEN_AXES_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:pages|pages_editor|Draft|editor)"
        r"[ \t\r\n]*(?:\.|::)[ \t\r\n]*(?:r#)?"
        r"(?P<method>table_hidden_axes|set_table_hidden_axes)\b"
        r"[ \t\r\n]*\(",
    ),
)
# The focused package must remain a terminal semantic owner.  A new
# ``Draft`` call or a direct shared-helper call in its
# owner would turn an unsupported focused operation into an implicit host
# fallback.  These patterns are applied after comments, literals, and
# cfg(test) items are masked.
PAGES_TABLE_HIDDEN_AXES_FOCUSED_FALLBACK_PATTERNS = (
    re.compile(r"\blitchi_pages[ \t\r\n]*::"),
    re.compile(r"\b(?:Draft|PagesPackage|FocusedPagesPackage)\b"),
    re.compile(
        r"(?<![A-Za-z0-9_])(?:crate[ \t\r\n]*::[ \t\r\n]*)?"
        r"(?:table_hidden_axes|set_table_hidden_axes)[ \t\r\n]*\("
    ),
)

PAGES_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE = (
    PAGES_SOURCE_ROOT / "table" / "hidden_axes.rs"
)
PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "body_table_hidden_axes.rs"
)
PAGES_TABLE_HIDDEN_AXES_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "body_table_hidden_axes"
)
PAGES_TABLE_HIDDEN_AXES_IMPLEMENTATION_SOURCES = (
    PAGES_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE,
    PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE,
)
PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "table" / "mod.rs",
)
PAGES_TABLE_HIDDEN_AXES_SELECTOR_SOURCE = PAGES_SOURCE_ROOT / "selector.rs"
PAGES_TABLE_HIDDEN_AXES_SEMANTIC_TYPES = ("AxisIndex", "HiddenAxes")
PAGES_TABLE_HIDDEN_AXES_TRANSACTION_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
# ADR 0034: the owner defines the canonical short names (identical to the
# ``transaction`` re-export set) and only
# ``litchi_pages::table::hidden_axes::transaction`` publishes them.  The
# deleted prefixed spellings stay focused markers and alias targets.
PAGES_TABLE_HIDDEN_AXES_CANONICAL_TYPES = PAGES_TABLE_HIDDEN_AXES_TRANSACTION_TYPES
PAGES_TABLE_HIDDEN_AXES_LEGACY_CANONICAL_TYPES = (
    "BodyTableHiddenAxesEdit",
    "BodyTableHiddenAxesPatch",
    "BodyTableHiddenAxesCommit",
    "BodyTableHiddenAxesDiagnostics",
    "BodyTableHiddenAxesError",
    "BodyTableHiddenAxesLimitKind",
    "BodyTableHiddenAxesPath",
)
PAGES_TABLE_HIDDEN_AXES_SHORT_NAMES = frozenset(
    PAGES_TABLE_HIDDEN_AXES_SEMANTIC_TYPES
    + PAGES_TABLE_HIDDEN_AXES_TRANSACTION_TYPES
)
PAGES_TABLE_HIDDEN_AXES_PUBLIC_NAMES = frozenset(
    PAGES_TABLE_HIDDEN_AXES_SEMANTIC_TYPES
    + PAGES_TABLE_HIDDEN_AXES_LEGACY_CANONICAL_TYPES
    + ("BodyTableSelector",)
)
# Public names that may legitimately be the target spelling of a rename
# (none of the renamed transaction names may be renamed again).
PAGES_TABLE_HIDDEN_AXES_ALIAS_EXEMPT_NAMES = frozenset(
    PAGES_TABLE_HIDDEN_AXES_SEMANTIC_TYPES + ("BodyTableSelector",)
)
PAGES_TABLE_HIDDEN_AXES_PACKAGE_METHODS = (
    "body_table_hidden_axes",
    "edit_body_table_hidden_axes",
    "apply_body_table_hidden_axes",
)
PAGES_TABLE_HIDDEN_AXES_EDIT_METHODS = frozenset({"set", "clear", "reset", "commit"})
PAGES_TABLE_HIDDEN_AXES_FLAT_METHODS = frozenset(
    {
        "table_hidden_axes",
        "set_table_hidden_axes",
        "edit_table_hidden_axes",
        "apply_table_hidden_axes",
        "clear_table_hidden_axes",
        "reset_table_hidden_axes",
    }
)
PAGES_TABLE_HIDDEN_AXES_FLAT_ALIASES = frozenset(
    {
        "BodyTableAxisIndex",
        "BodyTableHiddenAxes",
        "TableHiddenAxes",
        "PagesTableHiddenAxes",
        "HiddenAxesEdit",
        "HiddenAxesPatch",
        "HiddenAxesCommit",
        "HiddenAxesDiagnostics",
        "HiddenAxesError",
        "HiddenAxesLimitKind",
        "HiddenAxesPath",
        "TableHiddenAxesEdit",
        "TableHiddenAxesPatch",
        "TableHiddenAxesCommit",
        "TableHiddenAxesDiagnostics",
        "TableHiddenAxesError",
        "TableHiddenAxesLimitKind",
        "TableHiddenAxesPath",
        "PagesTableHiddenAxesEdit",
        "PagesTableHiddenAxesPatch",
        "PagesTableHiddenAxesCommit",
        "PagesTableHiddenAxesDiagnostics",
        "PagesTableHiddenAxesError",
        "PagesTableHiddenAxesLimitKind",
        "PagesTableHiddenAxesPath",
    }
)
PAGES_TABLE_HIDDEN_AXES_ALIAS_TARGETS = frozenset(
    PAGES_TABLE_HIDDEN_AXES_LEGACY_CANONICAL_TYPES
    + PAGES_TABLE_HIDDEN_AXES_SEMANTIC_TYPES
)
PAGES_TABLE_HIDDEN_AXES_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:body_table_hidden_axes|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?hidden_axes)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_PACKAGE_TABLE_HIDDEN_AXES_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?body_table_hidden_axes\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_TABLE_HIDDEN_AXES_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?body_table_hidden_axes\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_HIDDEN_AXES_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?hidden_axes\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_TABLE_HIDDEN_AXES_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_TABLE_HIDDEN_AXES_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HiddenStateExtentArchive",
        "HiddenStatesArchive",
        "HiddenStatesOwnerArchive",
        "HostPackage",
        "RawMessage",
        "Resolved",
        "RowOrColumnState",
        "SnappyStream",
        "SourceCatalog",
        "TableInfoArchive",
        "TableModelArchive",
    }
)
PAGES_TABLE_HIDDEN_AXES_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "PreparedRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteOutput",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
PAGES_TABLE_HIDDEN_AXES_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "litchi_iwa_protos", "tsp", "tst", "tsd", "tswp"}
)
PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/pages_hidden_state_codec.rs"
)
PAGES_TABLE_HIDDEN_AXES_CODEC_PUBLIC_SOURCE = Path("crates/litchi-iwa-protos/src/lib.rs")
PAGES_TABLE_HIDDEN_AXES_CODEC_MODULE = "pages_hidden_state_codec"
PAGES_TABLE_HIDDEN_AXES_CODEC_ALIAS_MODULE = "pages_hidden_axes_codec"
PAGES_TABLE_HIDDEN_AXES_CODEC_ALIAS_ROUTE = re.compile(
    rf"(?<![A-Za-z0-9_#])(?:mod|use|pub[ \t\r\n]+use)[ \t\r\n]+"
    rf"(?:r#)?{re.escape(PAGES_TABLE_HIDDEN_AXES_CODEC_ALIAS_MODULE)}\b"
    rf"|\bas[ \t\r\n]+(?:r#)?{re.escape(PAGES_TABLE_HIDDEN_AXES_CODEC_ALIAS_MODULE)}\b"
)
PAGES_TABLE_HIDDEN_AXES_CODEC_PROJECTION_SOURCE = Path(
    "crates/litchi-iwa-protos/src/buffa-projections/TSTPagesHiddenStateProjection.proto"
)
PAGES_TABLE_HIDDEN_AXES_CODEC_GENERATED_MODULE = (
    "buffa_pages_hidden_state_generated"
)
PAGES_TABLE_HIDDEN_AXES_CODEC_GENERATED_INCLUDE = (
    "/buffa-pages-hidden-state/iwa_pages_hidden_state_buffa_protos.rs"
)
# Pages' body table UID map is the 6267 native message.  6005 is a Numbers
# table-data list and must never become a hidden-axis ownership shortcut.
PAGES_TABLE_HIDDEN_AXES_UID_MAP_MESSAGE_TYPE = 6267
PAGES_TABLE_HIDDEN_AXES_WRONG_UID_MAP_MESSAGE_TYPE = 6005
PAGES_TABLE_HIDDEN_AXES_LEGACY_UID_MAP_MESSAGE_TYPE = 6200
# Keep the native route proof close to the feature inventory.  The focused
# owner and codec are the only places allowed to mention a UID-map routing
# ID; a test can still exercise the wrong type through the adversarial
# integration fixture below.
PAGES_TABLE_HIDDEN_AXES_UID_MAP_OWNER_MARKER = re.compile(
    r"\bUID_MAP_MESSAGE_TYPE\b[ \t\r\n]*:[ \t\r\n]*u32[ \t\r\n]*=[ \t\r\n]*6_?267\b"
)
PAGES_TABLE_HIDDEN_AXES_UID_MAP_CODEC_MARKER = re.compile(
    r"\bCOLUMN_ROW_UID_MAP_MESSAGE_TYPE\b[ \t\r\n]*:[ \t\r\n]*u32[ \t\r\n]*=[ \t\r\n]*6_?267\b"
)
PAGES_TABLE_HIDDEN_AXES_LEGACY_OWNER_MARKER = re.compile(
    r"\bLEGACY_UID_MAP_MESSAGE_TYPE\b[ \t\r\n]*:[ \t\r\n]*u32[ \t\r\n]*=[ \t\r\n]*6_?200\b"
)
PAGES_TABLE_HIDDEN_AXES_LEGACY_CODEC_MARKER = re.compile(
    r"\bLEGACY_COLUMN_ROW_UID_MAP_MESSAGE_TYPE\b[ \t\r\n]*:[ \t\r\n]*u32[ \t\r\n]*=[ \t\r\n]*6_?200\b"
)
PAGES_TABLE_HIDDEN_AXES_FORBIDDEN_UID_MAP_LITERAL = re.compile(
    r"(?<![A-Za-z0-9_])6_?005(?![A-Za-z0-9_])"
)
PAGES_TABLE_HIDDEN_AXES_LEGACY_OPT_IN_MARKER = re.compile(
    r"\ballow_legacy\b"
)
# The owner never creates a hidden-state graph: an absent hidden-state owner
# may be read as empty, but a changed non-empty edit must fail rather than
# inventing a graph, IDs, or metadata objects.  Keep the check on the known
# construction helpers rather than banning ordinary collection ``insert`` calls
# used to validate an existing owner.
PAGES_TABLE_HIDDEN_AXES_FORBIDDEN_CREATION_MARKERS = (
    re.compile(r"\bNewIds\b"),
    re.compile(r"\ballocate_ids\b"),
    re.compile(r"\bowner_to_proto\b"),
    re.compile(r"\binsert_object\b"),
)
PAGES_TABLE_HIDDEN_AXES_CODEC_REQUIRED_APIS = (
    "DecodeOptions",
    "DecodeLimit",
    "DecodeError",
    "DecodeReport",
    "UuidSnapshot",
    "ReferenceSnapshot",
    "AxisDirection",
    "RowOrColumnStateSnapshot",
    "HiddenStateExtentSnapshot",
    "HiddenStatesSnapshot",
    "HiddenStatesOwnerSnapshot",
    "TableInfoSnapshot",
    "TableModelSnapshot",
    "RewriteExecutionLimits",
    "RewriteExecutionRequirements",
    "RewriteOutput",
    "PreparedRewrite",
    "decode_table_info",
    "decode_table_info_with_report",
    "decode_table_model",
    "decode_table_model_with_report",
    "decode_hidden_states_owner",
    "decode_hidden_states_owner_with_report",
    "decode_hidden_state_extent",
    "decode_hidden_state_extent_with_report",
    "decode_row_or_column_state",
    "decode_row_or_column_state_with_report",
    "prepare_table_info_rewrite",
    "prepare_table_model_rewrite",
    "prepare_hidden_states_owner_rewrite",
    "prepare_hidden_state_extent_rewrite",
    "prepare_row_or_column_state_rewrite",
    "rewrite_table_info",
    "rewrite_table_model",
    "rewrite_hidden_states_owner",
    "rewrite_hidden_state_extent",
    "rewrite_row_or_column_state",
)
PAGES_TABLE_HIDDEN_AXES_CODEC_REQUIRED_MARKERS = (
    "unknown",
    "overlong",
    "preflight",
    "canonical",
    "duplicate",
    "decode_lazy_view",
    "execution_requirements",
    "execute",
    "max_input_bytes",
    "max_output_bytes",
    "max_fields",
    "max_work_bytes",
    "max_allocations",
    "max_retained_bytes",
    "max_scratch_bytes",
)
PAGES_TABLE_HIDDEN_AXES_CODEC_FORBIDDEN_PATTERNS = (
    re.compile(r"\bprost(?:_types)?\b"),
    re.compile(r"\b(?:encode_to_vec|to_owned_message|try_encode|merge_from)\b"),
    re.compile(r"\b(?:Message|prost::Message)[ \t\r\n]*::[ \t\r\n]*decode\b"),
)
PAGES_TABLE_HIDDEN_AXES_PROJECTION_MARKERS = (
    'syntax = "proto2";',
    "package LitchiIwaPagesHiddenStateProjection;",
    "message Uuid",
    "message Reference",
    "message TableInfoArchive",
    "message TableModelArchive",
    "message HiddenStatesOwnerArchive",
    "message HiddenStatesArchive",
    "message HiddenStateExtentArchive",
    "message RowOrColumnState",
    "hidden_states_owner",
    "hidden_states_uuid",
    "row_or_column_direction",
    "user_hidden",
    "filtered",
    "pivot_hidden",
)
PAGES_TABLE_HIDDEN_AXES_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
PAGES_TABLE_HIDDEN_AXES_TEST_SOURCES = (
    Path("crates/litchi-pages/tests/body_table_hidden_axes.rs"),
    Path("crates/litchi-pages/tests/body_table_hidden_axes_concurrency.rs"),
)
PAGES_TABLE_HIDDEN_AXES_FUZZ_SOURCES = (
    Path("crates/litchi-pages/fuzz/fuzz_targets/pages_body_table_hidden_axes.rs"),
)
PAGES_TABLE_HIDDEN_AXES_FUZZ_CORPORA = (
    Path("crates/litchi-pages/fuzz/corpus/pages_body_table_hidden_axes"),
)


# Longer body terminology aliases keep callers and boundary tests consistent
# with the public owner names without maintaining duplicate inventories.
PAGES_BODY_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE = PAGES_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE
PAGES_BODY_TABLE_HIDDEN_AXES_OWNER_SOURCE = PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE
PAGES_BODY_TABLE_HIDDEN_AXES_EXPORT_SOURCES = PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES
PAGES_BODY_TABLE_HIDDEN_AXES_PACKAGE_METHODS = PAGES_TABLE_HIDDEN_AXES_PACKAGE_METHODS
PAGES_BODY_TABLE_HIDDEN_AXES_CANONICAL_TYPES = PAGES_TABLE_HIDDEN_AXES_CANONICAL_TYPES

RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_MODULES = (
    "document_options",
    "footnote_settings",
)
PAGES_DOCUMENT_SETTINGS_IMPLEMENTATION_SOURCES = (
    PAGES_SOURCE_ROOT / "document_settings.rs",
    PAGES_SOURCE_ROOT / "package" / "document_settings.rs",
)
PAGES_DOCUMENT_SETTINGS_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
)
PAGES_DOCUMENT_SETTINGS_PUBLIC_MARKERS = frozenset(
    {
        "DocumentSettings",
        "DocumentSettingsCommit",
        "DocumentSettingsDiagnostics",
        "DocumentSettingsEdit",
        "DocumentSettingsError",
        "DocumentSettingsLimitKind",
        "DocumentSettingsPatch",
    }
)
PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE = (
    PAGES_SOURCE_ROOT / "section" / "settings.rs"
)
PAGES_SECTION_SETTINGS_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "section_settings.rs"
)
PAGES_SECTION_SETTINGS_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "section_settings"
)
PAGES_SECTION_SETTINGS_IMPLEMENTATION_SOURCES = (
    PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE,
    PAGES_SECTION_SETTINGS_OWNER_SOURCE,
)
PAGES_SECTION_SETTINGS_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "section.rs",
)
PAGES_SECTION_SETTINGS_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
    "DependencyKind",
)
PAGES_SECTION_SETTINGS_VALUE_TYPE = "Settings"
PAGES_SECTION_SETTINGS_SHORT_NAMES = frozenset(
    PAGES_SECTION_SETTINGS_CANONICAL_TYPES
)
PAGES_SECTION_SETTINGS_PUBLIC_NAMES = (
    PAGES_SECTION_SETTINGS_SHORT_NAMES | {PAGES_SECTION_SETTINGS_VALUE_TYPE}
)
PAGES_SECTION_SETTINGS_PACKAGE_METHODS = (
    "section_settings",
    "edit_section_settings",
    "apply_section_settings",
)
PAGES_SECTION_SETTINGS_FLAT_ALIASES = frozenset(
    {"SectionSettings", "PagesSectionSettings"}
    | {
        prefix + suffix
        for prefix in ("SectionSettings", "PagesSectionSettings")
        for suffix in PAGES_SECTION_SETTINGS_SHORT_NAMES
        if suffix != "Settings"
    }
)
PAGES_SECTION_SETTINGS_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:section_settings|section[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?settings)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_SECTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?section\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_SECTION_SETTINGS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?settings\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_SECTION_SETTINGS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?section_settings\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_SECTION_SETTINGS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?section_settings\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_SECTION_SETTINGS_PHYSICAL_TYPES = frozenset(
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
        "SectionArchive",
        "SectionSettingsArchive",
        "SectionSettingsSnapshot",
    }
)
PAGES_SECTION_SETTINGS_WIRE_TYPES = frozenset(
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
PAGES_SECTION_SETTINGS_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tp", "tsp", "tswp"}
)
RETIRED_IWA_PAGES_SECTION_BACKGROUND_MODULES = (
    "section_background",
    "section_settings",
)
PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE = (
    PAGES_SOURCE_ROOT / "section" / "background.rs"
)
PAGES_SECTION_BACKGROUND_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "section_background.rs"
)
PAGES_SECTION_BACKGROUND_OWNER_HELPER_ROOT = (
    PAGES_SOURCE_ROOT / "package" / "section_background"
)
PAGES_SECTION_BACKGROUND_IMPLEMENTATION_SOURCES = (
    PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE,
    PAGES_SECTION_BACKGROUND_OWNER_SOURCE,
)
PAGES_SECTION_BACKGROUND_EXPORT_SOURCES = (
    PAGES_SOURCE_ROOT / "lib.rs",
    PAGES_SOURCE_ROOT / "package.rs",
    PAGES_SOURCE_ROOT / "section.rs",
)
PAGES_SECTION_BACKGROUND_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
PAGES_SECTION_BACKGROUND_SHORT_NAMES = frozenset(
    PAGES_SECTION_BACKGROUND_CANONICAL_TYPES
)
PAGES_SECTION_BACKGROUND_VALUE_TYPE = "Background"
PAGES_SECTION_BACKGROUND_PUBLIC_NAMES = (
    PAGES_SECTION_BACKGROUND_SHORT_NAMES | {PAGES_SECTION_BACKGROUND_VALUE_TYPE}
)
PAGES_SECTION_BACKGROUND_PACKAGE_METHODS = (
    "section_background",
    "edit_section_background",
    "apply_section_background",
)
PAGES_SECTION_BACKGROUND_EDIT_METHODS = (
    "background",
    "set_solid",
    "clear",
    "commit",
)
PAGES_SECTION_BACKGROUND_FLAT_ALIASES = frozenset(
    {"SectionBackground", "PagesSectionBackground"}
    | {
        prefix + suffix
        for prefix in ("SectionBackground", "PagesSectionBackground")
        for suffix in PAGES_SECTION_BACKGROUND_SHORT_NAMES
    }
)
PAGES_SECTION_BACKGROUND_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:section_background|section[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?background)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
PUBLIC_PAGES_SECTION_BACKGROUND_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?background\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_PAGES_PACKAGE_SECTION_BACKGROUND_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?section_background\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_PACKAGE_SECTION_BACKGROUND_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?section_background\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PAGES_SECTION_BACKGROUND_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "Opaque",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SectionArchive",
        "SectionBackgroundArchive",
        "SectionBackgroundSnapshot",
        "SnappyStream",
        "SourceCatalog",
    }
)
PAGES_SECTION_BACKGROUND_WIRE_TYPES = frozenset(
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
PAGES_SECTION_BACKGROUND_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "tp", "tsd", "tsp", "tswp"}
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


def _pages_table_lock_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in Pages table locks."""

    if identifier in PAGES_TABLE_LOCK_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_TABLE_LOCK_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_TABLE_LOCK_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_table_lock_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_TABLE_LOCK_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_TABLE_LOCK_PACKAGE_METHODS for identifier in identifiers
    )


def _pages_table_title_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in Pages table titles."""

    if identifier in PAGES_TABLE_TITLE_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_TABLE_TITLE_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_TABLE_TITLE_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_table_title_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_TABLE_TITLE_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_TABLE_TITLE_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_table_title_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers & (PAGES_TABLE_TITLE_FLAT_ALIASES | PAGES_TABLE_TITLE_PUBLIC_NAMES)
    ) or _pages_table_title_owner_declaration(declaration)


def _pages_table_headers_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in Pages table headers."""

    if identifier in PAGES_TABLE_HEADERS_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_TABLE_HEADERS_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_TABLE_HEADERS_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_table_headers_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_TABLE_HEADERS_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_TABLE_HEADERS_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_table_headers_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers & (PAGES_TABLE_HEADERS_FLAT_ALIASES | PAGES_TABLE_HEADERS_FOCUSED_MARKERS)
    ) or _pages_table_headers_owner_declaration(declaration)


def _pages_table_dimension_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in Pages dimensions."""

    if identifier in PAGES_TABLE_DIMENSION_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_TABLE_DIMENSION_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_TABLE_DIMENSION_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_table_sort_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in Pages sort APIs."""

    if identifier in PAGES_TABLE_SORT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_TABLE_SORT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_TABLE_SORT_WIRE_TYPES:
        return "wire type"
    if identifier == "litchi_iwa_common":
        return None
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_table_sort_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_TABLE_SORT_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_TABLE_SORT_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_table_sort_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers & (PAGES_TABLE_SORT_FLAT_ALIASES | PAGES_TABLE_SORT_PUBLIC_NAMES)
    ) or _pages_table_sort_owner_declaration(declaration)


def _pages_table_appearance_public_leak(identifier: str) -> str | None:
    """Classify physical/wire vocabulary forbidden in Pages appearance APIs."""

    if identifier == "litchi_iwa_common":
        return None
    if identifier in PAGES_TABLE_APPEARANCE_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_TABLE_APPEARANCE_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_TABLE_APPEARANCE_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_table_appearance_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_TABLE_APPEARANCE_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_TABLE_APPEARANCE_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_table_appearance_public_declaration(
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
            set(PAGES_TABLE_APPEARANCE_FLAT_ALIASES)
            | set(PAGES_TABLE_APPEARANCE_PUBLIC_NAMES)
        )
    ) or _pages_table_appearance_owner_declaration(declaration)


def _pages_table_appearance_owner_present(root: Path) -> bool:
    """Activate the body-table appearance audit only after its owner is wired."""

    owner_path = root / PAGES_TABLE_APPEARANCE_OWNER_SOURCE
    package_path = root / PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and (
        PAGES_PACKAGE_TABLE_APPEARANCE_MODULE.search(package_source) is not None
    )


def _pages_table_hidden_axes_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in hidden-axis facades."""

    if identifier in PAGES_TABLE_HIDDEN_AXES_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_TABLE_HIDDEN_AXES_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_TABLE_HIDDEN_AXES_WIRE_TYPES:
        return "wire type"
    # The semantic module intentionally re-exports the shared archive-free
    # value implementation.  Its crate path is not a physical leak; the
    # package/codec paths above are the only native boundary this ratchet
    # permits.
    if identifier == "litchi_iwa_common":
        return None
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_table_hidden_axes_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_TABLE_HIDDEN_AXES_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_TABLE_HIDDEN_AXES_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_table_hidden_axes_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    if bool(
        identifiers
        & (
            PAGES_TABLE_HIDDEN_AXES_FLAT_ALIASES
            | PAGES_TABLE_HIDDEN_AXES_PUBLIC_NAMES
        )
    ) or _pages_table_hidden_axes_owner_declaration(declaration):
        return True
    # A rogue public re-export need not use one of the canonical names (for
    # example ``pub use litchi_iwa_protos::TableInfoArchive``).  Treat any
    # recognized native/wire vocabulary or raw byte/ID parameter as evidence
    # that the declaration must be inspected, so it cannot evade the facade
    # leak ratchet by choosing a new alias.
    # Limit the broad fallback to hidden-axis-specific wire/projection names.
    # ``package.rs`` is also an export source and legitimately publishes the
    # generic archive ingress (`Package::from_bytes`, `Archive`, etc.); those
    # declarations must not be mistaken for this table owner.  A direct
    # `litchi_iwa_protos`/Prost/generated projection route remains fail-closed.
    native_identifiers = (
        identifiers
        & (
            PAGES_TABLE_HIDDEN_AXES_PROTO_ORIGINS
            | PAGES_TABLE_HIDDEN_AXES_WIRE_TYPES
            | (
                PAGES_TABLE_HIDDEN_AXES_PHYSICAL_TYPES
                - {
                    "Archive",
                    "ArchiveObject",
                    "ComponentCatalog",
                    "EntryEdit",
                    "ExactArtifacts",
                    "HostPackage",
                    "RawMessage",
                    "Resolved",
                    "SnappyStream",
                    "SourceCatalog",
                }
            )
        )
    )
    generated_or_codec = any(
        any(word in identifier.lower() for word in ("buffa", "prost", "generated"))
        for identifier in identifiers
    )
    if native_identifiers or generated_or_codec:
        return True
    hidden_context = bool(
        identifiers
        & (
            PAGES_TABLE_HIDDEN_AXES_FLAT_ALIASES
            | PAGES_TABLE_HIDDEN_AXES_PUBLIC_NAMES
            | set(PAGES_TABLE_HIDDEN_AXES_PACKAGE_METHODS)
        )
    )
    return (
        hidden_context
        and (
            PAGES_TABLE_HIDDEN_AXES_PUBLIC_RAW_ID_PARAMETER.search(declaration)
            is not None
            or RUST_BYTE_SLICE.search(declaration) is not None
        )
    )


def _pages_table_hidden_axes_owner_present(root: Path) -> bool:
    """Activate the hidden-axis ratchets only after the private owner is wired."""

    owner_path = root / PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE
    package_path = root / PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and (
        PAGES_PACKAGE_TABLE_HIDDEN_AXES_MODULE.search(package_source) is not None
    )


def _pages_footnote_lifecycle_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in body-footnote APIs."""

    if identifier in PAGES_FOOTNOTE_LIFECYCLE_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_FOOTNOTE_LIFECYCLE_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_FOOTNOTE_LIFECYCLE_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_footnote_lifecycle_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_FOOTNOTE_LIFECYCLE_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_FOOTNOTE_LIFECYCLE_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_pages_footnote_lifecycle_public_declaration(
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
            PAGES_FOOTNOTE_LIFECYCLE_FLAT_ALIASES
            | PAGES_FOOTNOTE_LIFECYCLE_SHORT_NAMES
        )
    ) or _pages_footnote_lifecycle_owner_declaration(declaration)


def _pages_table_dimension_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_TABLE_DIMENSION_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_TABLE_DIMENSION_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_table_dimension_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (PAGES_TABLE_DIMENSION_FLAT_ALIASES | PAGES_TABLE_DIMENSION_PUBLIC_NAMES)
    ) or _pages_table_dimension_owner_declaration(declaration)


def _pages_header_footer_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in Pages header/footer APIs."""

    if identifier in PAGES_HEADER_FOOTER_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_HEADER_FOOTER_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_HEADER_FOOTER_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_header_footer_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_HEADER_FOOTER_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_HEADER_FOOTER_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_header_footer_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (PAGES_HEADER_FOOTER_FLAT_ALIASES | PAGES_HEADER_FOOTER_PUBLIC_NAMES)
    ) or _pages_header_footer_owner_declaration(declaration)


def _pages_section_text_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in section-text APIs."""

    if identifier in PAGES_SECTION_TEXT_ALLOWED_PUBLIC_IDENTIFIERS:
        return None
    if identifier in PAGES_SECTION_TEXT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_SECTION_TEXT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_SECTION_TEXT_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_section_text_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_SECTION_TEXT_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_SECTION_TEXT_PACKAGE_METHODS for identifier in identifiers
    )


def _is_pages_section_text_public_declaration(
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
            PAGES_SECTION_TEXT_FLAT_ALIASES
            | PAGES_SECTION_TEXT_PUBLIC_NAMES
            | set(PAGES_SECTION_TEXT_PACKAGE_METHODS)
            | PAGES_SECTION_TEXT_FLAT_METHODS
        )
    ) or _pages_section_text_owner_declaration(declaration)


def _is_pages_table_lock_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return (
        bool(identifiers & PAGES_TABLE_LOCK_FLAT_ALIASES)
        or bool(identifiers & PAGES_TABLE_LOCK_FOCUSED_MARKERS)
        or _pages_table_lock_owner_declaration(declaration)
    )


def _is_pages_document_settings_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & PAGES_DOCUMENT_SETTINGS_PUBLIC_MARKERS) or any(
        "document_settings" in identifier.lower() for identifier in identifiers
    )


def _pages_section_settings_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in section settings."""

    if identifier in PAGES_SECTION_SETTINGS_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_SECTION_SETTINGS_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_SECTION_SETTINGS_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_section_settings_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_SECTION_SETTINGS_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_SECTION_SETTINGS_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_pages_section_settings_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return (
        bool(identifiers & PAGES_SECTION_SETTINGS_FLAT_ALIASES)
        or (
            "settings" in identifiers
            and bool(identifiers & PAGES_SECTION_SETTINGS_PUBLIC_NAMES)
        )
        or _pages_section_settings_owner_declaration(declaration)
    )


def _pages_section_background_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in section backgrounds."""

    if identifier in PAGES_SECTION_BACKGROUND_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in PAGES_SECTION_BACKGROUND_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in PAGES_SECTION_BACKGROUND_WIRE_TYPES:
        return "wire type"
    reason = _iwork_public_leak(identifier)
    if reason is not None:
        return reason
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
    return None


def _pages_section_background_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return PAGES_SECTION_BACKGROUND_OWNER_PATH.search(declaration) is not None or any(
        identifier in PAGES_SECTION_BACKGROUND_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_pages_section_background_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return (
        bool(identifiers & PAGES_SECTION_BACKGROUND_FLAT_ALIASES)
        or (
            "background" in identifiers
            and bool(identifiers & PAGES_SECTION_BACKGROUND_PUBLIC_NAMES)
        )
        or _pages_section_background_owner_declaration(declaration)
    )


def _pages_table_name_owner_present(root: Path) -> bool:
    """Activate body-table name retirement only after its package owner lands."""

    owner_path = root / PAGES_TABLE_NAME_OWNER_SOURCE
    package_path = root / PAGES_TABLE_NAME_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and (
        PAGES_PACKAGE_TABLE_NAME_MODULE.search(package_source) is not None
    )


def audit_pages_table_name_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Require the selector-first Pages package name seam once it lands."""

    if not _pages_table_name_owner_present(root):
        return []

    violations: list[str] = []
    owner_path = root / PAGES_TABLE_NAME_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(
        owner_path.read_text(encoding="utf-8")
    )
    methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in PAGES_TABLE_NAME_PACKAGE_METHODS:
        declaration = methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-pages table-name public API is missing Package "
                f"method {method}: {PAGES_TABLE_NAME_OWNER_SOURCE}"
            )
            continue
        if method != "apply_body_table_name" and "BodyTableSelector" not in declaration:
            violations.append(
                "focused litchi-pages table-name Package method "
                f"{method} must accept selector-first BodyTableSelector: "
                f"{PAGES_TABLE_NAME_OWNER_SOURCE}"
            )
    if "rename_table" in methods:
        violations.append(
            "focused litchi-pages table-name public API retains flat Package "
            f"method rename_table: {PAGES_TABLE_NAME_OWNER_SOURCE}"
        )

    selector_path = root / PAGES_TABLE_NAME_SELECTOR_SOURCE
    selector_source = (
        _mask_rust_non_code(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    if "BodyTableSelector" not in _rust_canonical_exports(
        selector_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-name public API is missing canonical "
            f"BodyTableSelector: {PAGES_TABLE_NAME_SELECTOR_SOURCE}"
        )

    lib_path = root / PAGES_TABLE_NAME_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if "BodyTableSelector" not in _rust_canonical_exports(
        lib_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-name public API is missing root "
            f"BodyTableSelector re-export: {PAGES_TABLE_NAME_EXPORT_SOURCES[0]}"
        )
    return sorted(set(violations))


def _pages_table_sort_owner_present(root: Path) -> bool:
    owner_path = root / PAGES_TABLE_SORT_OWNER_SOURCE
    package_path = root / PAGES_TABLE_SORT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and (
        PAGES_PACKAGE_TABLE_SORT_MODULE.search(package_source) is not None
    )


def audit_pages_table_sort_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Enforce the selector-first, archive-free Pages table-sort owner."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir() or not _pages_table_sort_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / PAGES_TABLE_SORT_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(PAGES_TABLE_SORT_SEMANTIC_TYPES)
    )
    for name in PAGES_TABLE_SORT_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-pages table-sort public API is missing "
                f"semantic table::sort type {name}: {PAGES_TABLE_SORT_SEMANTIC_SOURCE}"
            )
    if PUBLIC_PAGES_TABLE_SORT_TRANSACTION_MODULE.search(
        _mask_rust_non_code(semantic_source)
    ) is None:
        violations.append(
            "focused litchi-pages table-sort public API is missing canonical "
            f"table::sort::transaction module: {PAGES_TABLE_SORT_SEMANTIC_SOURCE}"
        )

    selector_path = root / PAGES_TABLE_SORT_SELECTOR_SOURCE
    selector_source = (
        _mask_rust_non_code(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    if "BodyTableSelector" not in _rust_canonical_exports(
        selector_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-sort public API is missing canonical "
            f"BodyTableSelector: {PAGES_TABLE_SORT_SELECTOR_SOURCE}"
        )

    owner_path = root / PAGES_TABLE_SORT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_exports = _rust_canonical_exports(
        owner_source, frozenset(PAGES_TABLE_SORT_CANONICAL_TYPES)
    )
    for name in PAGES_TABLE_SORT_CANONICAL_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-pages table-sort public API is missing "
                f"canonical package type {name}: {PAGES_TABLE_SORT_OWNER_SOURCE}"
            )
    violations.extend(
        _pages_transaction_family_violations(
            root, "body_table_sort", "focused litchi-pages table-sort public API"
        )
    )
    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in PAGES_TABLE_SORT_PACKAGE_METHODS:
        declaration = owner_methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-pages table-sort public API is missing "
                f"Package method {method}: {PAGES_TABLE_SORT_OWNER_SOURCE}"
            )
            continue
        if method != "apply_body_table_sort_order":
            if "BodyTableSelector" not in declaration:
                violations.append(
                    "focused litchi-pages table-sort Package method "
                    f"{method} must accept selector-first BodyTableSelector: "
                    f"{PAGES_TABLE_SORT_OWNER_SOURCE}"
                )
    for method in sorted(PAGES_TABLE_SORT_FLAT_METHODS & owner_methods.keys()):
        violations.append(
            "focused litchi-pages table-sort public API retains flat Package method "
            f"{method}: {PAGES_TABLE_SORT_OWNER_SOURCE}"
        )
    owner_code = _mask_rust_non_code(owner_source)
    for label, marker in PAGES_TABLE_SORT_OWNER_REQUIRED_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-pages table-sort owner is missing "
                f"{label} transaction marker: {PAGES_TABLE_SORT_OWNER_SOURCE}"
            )

    lib_path = root / PAGES_TABLE_SORT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages table-sort public API is missing canonical root "
            f"table module: {PAGES_TABLE_SORT_EXPORT_SOURCES[0]}"
        )
    if "BodyTableSelector" not in _rust_canonical_exports(
        lib_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-sort public API is missing root "
            f"BodyTableSelector re-export: {PAGES_TABLE_SORT_EXPORT_SOURCES[0]}"
        )
    table_path = root / PAGES_TABLE_SORT_EXPORT_SOURCES[2]
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_SORT_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-pages table-sort public API is missing canonical "
            f"table::sort module: {PAGES_TABLE_SORT_EXPORT_SOURCES[2]}"
        )
    package_path = root / PAGES_TABLE_SORT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    if PAGES_PACKAGE_TABLE_SORT_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-pages table-sort public API is missing private package "
            f"owner module: {PAGES_TABLE_SORT_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_PAGES_PACKAGE_TABLE_SORT_MODULE.finditer(package_source):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-pages table-sort public API exposes duplicate "
            f"package::body_table_sort module: {PAGES_TABLE_SORT_EXPORT_SOURCES[1]}:{line_number}"
        )

    codec_path = root / PAGES_TABLE_SORT_CODEC_SOURCE
    if not codec_path.is_file():
        violations.append(
            "focused litchi-pages table-sort public API is missing strict hidden "
            f"codec source: {PAGES_TABLE_SORT_CODEC_SOURCE}"
        )
    else:
        codec_source = _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        codec_code = _mask_rust_non_code(codec_source)
        for api in PAGES_TABLE_SORT_CODEC_REQUIRED_APIS:
            if re.search(
                rf"\b(?:pub\s+)?(?:fn|struct|type)\s+{re.escape(api)}\b",
                codec_code,
            ) is None:
                violations.append(
                    "focused litchi-pages table-sort hidden codec is missing "
                    f"strict API {api}: {codec_path.relative_to(root)}"
                )
        for marker in PAGES_TABLE_SORT_CODEC_REQUIRED_MARKERS:
            if re.search(re.escape(marker), codec_source) is None:
                violations.append(
                    "focused litchi-pages table-sort hidden codec is missing "
                    f"strict {marker} marker: {codec_path.relative_to(root)}"
                )
        codec_lib_path = root / PAGES_TABLE_SORT_CODEC_PUBLIC_SOURCE
        codec_lib_source = (
            _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
            if codec_lib_path.is_file()
            else ""
        )
        if re.search(
            rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
            rf"pub\s+mod\s+{re.escape(PAGES_TABLE_SORT_CODEC_MODULE)}\b",
            codec_lib_source,
        ) is None:
            violations.append(
                "focused litchi-pages table-sort public API is missing hidden codec "
                f"module {PAGES_TABLE_SORT_CODEC_MODULE}: {PAGES_TABLE_SORT_CODEC_PUBLIC_SOURCE}"
            )

    for target in (
        PAGES_TABLE_SORT_CODEC_FUZZ_SOURCE,
        PAGES_TABLE_SORT_FUZZ_SOURCE,
    ):
        absolute = root / target
        if not absolute.is_file():
            violations.append(
                "focused litchi-pages table-sort boundary is missing fuzz target: "
                f"{target}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-pages table-sort fuzz target is missing "
                f"fuzz_target! harness: {target}"
            )
    for corpus in (PAGES_TABLE_SORT_CODEC_FUZZ_CORPUS, PAGES_TABLE_SORT_FUZZ_CORPUS):
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-pages table-sort boundary is missing fuzz corpus: "
                f"{corpus}"
            )

    dedicated_sources = {
        root / path
        for path in PAGES_TABLE_SORT_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_TABLE_SORT_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_TABLE_SORT_EXPORT_SOURCES
        if (root / path).is_file()
    }

    # Scan the whole facade tree.  In particular, cfg(test) is masked item by
    # item so a production sibling cannot launder a raw-ID or glob route.
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "body_table_sort"]:
                violations.append(
                    "focused litchi-pages table-sort public API exposes duplicate "
                    f"body_table_sort module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if {"body_table_sort", "sort", "package"} & set(identifiers):
                    violations.append(
                        "focused litchi-pages table-sort public API retains root "
                        f"aliases via table::sort glob: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_SORT_ALIAS_TARGETS
                    or "body_table_sort" in target_identifiers
                    or "sort" in target_identifiers
                ) and alias and alias != target:
                    # ADR 0034: no rename of the owner's short names and no
                    # prefixed root alias.
                    violations.append(
                        "focused litchi-pages table-sort public API retains "
                        f"alternate alias {alias} for {target}: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_SORT_ALIAS_TARGETS
                    or "body_table_sort" in target_identifiers
                    or "sort" in target_identifiers
                ) and alias != target:
                    violations.append(
                        "focused litchi-pages table-sort public API retains "
                        f"alternate alias {alias} for {target}: {path.relative_to(root)}:{line_number}"
                    )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for method, line_number in _rust_function_declarations(source):
            if method not in PAGES_TABLE_HIDDEN_AXES_FLAT_METHODS:
                continue
            violations.append(
                "focused litchi-pages body-table hidden-axes public API retains flat function "
                f"{method}: {path.relative_to(root)}:{line_number}"
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
        for declaration, line_number, public_declaration, complete_scope in declarations:
            if not _is_pages_table_sort_public_declaration(
                declaration, dedicated_source=complete_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages table-sort public API retains root aliases "
                    f"via table::sort glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_TABLE_SORT_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages table-sort public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_table_sort_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-pages table-sort public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages table-sort public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

            if public_declaration and PAGES_TABLE_SORT_RAW_PARAMETER.search(declaration):
                match = PAGES_TABLE_SORT_RAW_PARAMETER.search(declaration)
                assert match is not None
                violations.append(
                    "focused litchi-pages table-sort public API exposes raw identifier/bytes "
                    f"{match.group(0)}: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _pages_header_footer_owner_present(root: Path) -> bool:
    return (root / PAGES_HEADER_FOOTER_OWNER_SOURCE).is_file()


def audit_pages_table_appearance_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Pages appearance owner."""

    if not _pages_table_appearance_owner_present(root):
        return []

    owner_path = root / PAGES_TABLE_APPEARANCE_OWNER_SOURCE
    semantic_path = root / PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE
    selector_path = root / PAGES_TABLE_APPEARANCE_SELECTOR_SOURCE
    package_path = root / PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[1]
    lib_path = root / PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[0]
    table_path = root / PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[2]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path, table_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    package_code = code[package_path]
    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+body_table_appearance\b",
        package_code + code[lib_path],
    ):
        violations.append(
            "focused litchi-pages body-table appearance owner module must remain private: "
            f"{PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[1]}"
        )
    if PAGES_PACKAGE_TABLE_APPEARANCE_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-pages body-table appearance owner module is missing: "
            f"{PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[1]}"
        )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path], frozenset(PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES)
    )
    for name in PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-pages body-table appearance semantic API is missing "
                f"{name}: {PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE}"
            )
    if PUBLIC_PAGES_TABLE_APPEARANCE_TRANSACTION_MODULE.search(
        code[semantic_path]
    ) is None:
        violations.append(
            "focused litchi-pages body-table appearance semantic API is missing "
            f"table::appearance::transaction: {PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], frozenset(PAGES_TABLE_APPEARANCE_TRANSACTION_TYPES)
    )
    for name in PAGES_TABLE_APPEARANCE_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-pages body-table appearance transaction API is missing "
                f"{name}: {PAGES_TABLE_APPEARANCE_SEMANTIC_SOURCE}"
            )

    selector_exports = _rust_canonical_exports(
        sources[selector_path], frozenset({"BodyTableSelector"})
    )
    if "BodyTableSelector" not in selector_exports:
        violations.append(
            "focused litchi-pages body-table appearance public API is missing canonical "
            f"BodyTableSelector: {PAGES_TABLE_APPEARANCE_SELECTOR_SOURCE}"
        )
    if "BodyTableSelector" not in _rust_canonical_exports(
        sources[lib_path], frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages body-table appearance public API is missing root "
            f"BodyTableSelector re-export: {PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[0]}"
        )

    owner_exports = _rust_canonical_exports(
        sources[owner_path], frozenset(PAGES_TABLE_APPEARANCE_CANONICAL_TYPES)
    )
    for name in PAGES_TABLE_APPEARANCE_CANONICAL_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-pages body-table appearance public API is missing canonical "
                f"type {name}: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
    violations.extend(
        _pages_transaction_family_violations(
            root,
            "body_table_appearance",
            "focused litchi-pages body-table appearance public API",
        )
    )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for method in PAGES_TABLE_APPEARANCE_PACKAGE_METHODS:
        declaration = owner_methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-pages body-table appearance Package method is missing "
                f"{method}: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
            continue
        if method != "apply_body_table_appearance" and "BodyTableSelector" not in declaration:
            violations.append(
                "focused litchi-pages body-table appearance Package method "
                f"{method} must accept selector-first BodyTableSelector: "
                f"{PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
    apply_declaration = owner_methods.get("apply_body_table_appearance")
    if apply_declaration is not None and re.search(r"\bPatch\b", apply_declaration) is None:
        violations.append(
            "focused litchi-pages body-table appearance apply method must accept "
            f"Patch: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
        )
    edit_methods = {
        name
        for edit_type in ("BodyTableAppearanceEdit", "Edit")
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], edit_type
        )
    }
    for name in sorted(PAGES_TABLE_APPEARANCE_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-pages body-table appearance edit is missing "
            f"{name}: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
        )
    for name in ("clear", "reset"):
        if name in edit_methods:
            violations.append(
                "focused litchi-pages body-table appearance edit must not expose "
                f"{name}: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
    flat_methods = {"table_appearance", "edit_table_appearance", "apply_table_appearance"}
    for method in sorted(flat_methods & owner_methods.keys()):
        violations.append(
            "focused litchi-pages body-table appearance public API retains flat Package method "
            f"{method}: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
        )

    if PUBLIC_PAGES_TABLE_MODULE.search(code[lib_path]) is None:
        violations.append(
            "focused litchi-pages body-table appearance public API is missing canonical root "
            f"table module: {PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[0]}"
        )
    if PUBLIC_PAGES_TABLE_APPEARANCE_MODULE.search(code[table_path]) is None:
        violations.append(
            "focused litchi-pages body-table appearance public API is missing canonical "
            f"table::appearance module: {PAGES_TABLE_APPEARANCE_EXPORT_SOURCES[2]}"
        )

    codec_path = root / PAGES_TABLE_APPEARANCE_CODEC_SOURCE
    if not codec_path.is_file():
        violations.append(
            "focused litchi-pages body-table appearance public API is missing strict hidden "
            f"codec source: {PAGES_TABLE_APPEARANCE_CODEC_SOURCE}"
        )
    else:
        codec_raw = codec_path.read_text(encoding="utf-8")
        codec_source = _mask_rust_cfg_test_items(codec_raw)
        codec_code = _mask_rust_non_code(codec_source)
        for api in PAGES_TABLE_APPEARANCE_CODEC_REQUIRED_APIS:
            if re.search(
                rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
                codec_code,
            ) is None:
                violations.append(
                    "focused litchi-pages body-table appearance hidden codec is missing "
                    f"strict API {api}: {PAGES_TABLE_APPEARANCE_CODEC_SOURCE}"
                )
        for label, markers in PAGES_TABLE_APPEARANCE_CODEC_REQUIRED_MARKER_GROUPS.items():
            if not all(marker in codec_code for marker in markers):
                violations.append(
                    "focused litchi-pages body-table appearance hidden codec is missing "
                    f"{label} marker: {PAGES_TABLE_APPEARANCE_CODEC_SOURCE}"
                )
        codec_lib_path = root / PAGES_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE
        codec_lib_source = (
            _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
            if codec_lib_path.is_file()
            else ""
        )
        if re.search(
            rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
            rf"pub\s+mod\s+{re.escape(PAGES_TABLE_APPEARANCE_CODEC_MODULE)}\b",
            codec_lib_source,
        ) is None:
            violations.append(
                "focused litchi-pages body-table appearance public API is missing hidden "
                f"codec module {PAGES_TABLE_APPEARANCE_CODEC_MODULE}: {PAGES_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE}"
            )
        if re.search(r"(?m)#\s*\[\s*cfg\(\s*test\s*\)\s*\]", codec_raw) is None or re.search(
            r"(?m)#\s*\[\s*test\s*\]", codec_raw
        ) is None:
            violations.append(
                "focused litchi-pages body-table appearance hidden codec is missing cfg(test) #[test] coverage: "
                f"{PAGES_TABLE_APPEARANCE_CODEC_SOURCE}"
            )

    for test_path in PAGES_TABLE_APPEARANCE_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-pages body-table appearance boundary is missing integration test: "
                f"{test_path}"
            )
        else:
            test_code = _mask_rust_non_code(absolute.read_text(encoding="utf-8"))
            if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
                violations.append(
                    "focused litchi-pages body-table appearance integration test is missing #[test] coverage: "
                    f"{test_path}"
                )
            if not re.search(
                r"\b(?:body_table_appearance|edit_body_table_appearance|apply_body_table_appearance)\b",
                test_code,
            ):
                violations.append(
                    "focused litchi-pages body-table appearance integration test does not exercise the owner: "
                    f"{test_path}"
                )
    for fuzz_path in PAGES_TABLE_APPEARANCE_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-pages body-table appearance boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-pages body-table appearance fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in PAGES_TABLE_APPEARANCE_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-pages body-table appearance boundary is missing fuzz corpus: "
                f"{corpus}"
            )

    dedicated_sources = {
        root / path
        for path in PAGES_TABLE_APPEARANCE_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_TABLE_APPEARANCE_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_TABLE_APPEARANCE_EXPORT_SOURCES
        if (root / path).is_file()
    }
    facade_names = (
        set(PAGES_TABLE_APPEARANCE_LEGACY_CANONICAL_TYPES)
        | set(PAGES_TABLE_APPEARANCE_SEMANTIC_TYPES)
        | {"BodyTableSelector"}
        | set(PAGES_TABLE_APPEARANCE_PACKAGE_METHODS)
        | set(PAGES_TABLE_APPEARANCE_FLAT_ALIASES)
    )
    for path in sorted(dedicated_sources | export_sources):
        dedicated = path in dedicated_sources
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        declarations = [
            (declaration, line_number, True, dedicated)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for declaration, line_number, public_declaration, complete_scope in declarations:
            if not _is_pages_table_appearance_public_declaration(
                declaration, dedicated_source=complete_scope
            ):
                continue
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if public_declaration and "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-pages body-table appearance public API retains a glob re-export: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for identifier in sorted(identifiers):
                reason = _pages_table_appearance_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-pages body-table appearance public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if public_declaration and identifier in PAGES_TABLE_APPEARANCE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages body-table appearance public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in PAGES_TABLE_APPEARANCE_PUBLIC_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-pages body-table appearance public API exposes raw parameter "
                    f"{match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-pages body-table appearance public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{line_number}"
                )

    # Catch public aliases and duplicate module routes in sibling facade files.
    for path in sorted((root / PAGES_SOURCE_ROOT).rglob("*.rs")):
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "body_table_appearance"] and path != lib_path:
                violations.append(
                    "focused litchi-pages body-table appearance public API exposes duplicate "
                    f"body_table_appearance module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if {"body_table_appearance", "appearance", "package"} & set(identifiers):
                    violations.append(
                        "focused litchi-pages body-table appearance public API retains root aliases via glob: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target = identifiers[alias_index - 1] if alias_index > 2 else ""
                alias = identifiers[alias_index + 1] if alias_index + 1 < len(identifiers) else ""
                if target in PAGES_TABLE_APPEARANCE_ALIAS_TARGETS and alias and alias != target:
                    if alias in PAGES_TABLE_APPEARANCE_ALIAS_EXEMPT_NAMES:
                        continue
                    violations.append(
                        "focused litchi-pages body-table appearance public API retains alternate alias "
                        f"{alias} for {target}: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target = identifiers[-1]
                if target in PAGES_TABLE_APPEARANCE_ALIAS_TARGETS and alias != target and alias not in PAGES_TABLE_APPEARANCE_ALIAS_EXEMPT_NAMES:
                    violations.append(
                        "focused litchi-pages body-table appearance public API retains alternate alias "
                        f"{alias} for {target}: {path.relative_to(root)}:{line_number}"
                    )
    return sorted(set(violations))


def audit_pages_table_hidden_axes_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Pages hidden-axis owner."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir() or not _pages_table_hidden_axes_owner_present(root):
        return []

    owner_path = root / PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE
    semantic_path = root / PAGES_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE
    selector_path = root / PAGES_TABLE_HIDDEN_AXES_SELECTOR_SOURCE
    lib_path = root / PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[0]
    package_path = root / PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[1]
    table_path = root / PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[2]
    paths = (owner_path, semantic_path, selector_path, lib_path, package_path, table_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    package_code = code[package_path]
    package_raw = sources[package_path]
    lib_code = code[lib_path]
    if PUBLIC_PAGES_PACKAGE_TABLE_HIDDEN_AXES_MODULE.search(
        package_raw + sources[lib_path]
    ):
        violations.append(
            "focused litchi-pages body-table hidden-axes owner module must remain private: "
            f"{PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[1]}"
        )
    if PAGES_PACKAGE_TABLE_HIDDEN_AXES_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes owner module is missing: "
            f"{PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[1]}"
        )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path], frozenset(PAGES_TABLE_HIDDEN_AXES_SEMANTIC_TYPES)
    )
    for name in PAGES_TABLE_HIDDEN_AXES_SEMANTIC_TYPES:
        if name not in semantic_exports:
            violations.append(
                "focused litchi-pages body-table hidden-axes semantic API is missing "
                f"{name}: {PAGES_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE}"
            )
    if PUBLIC_PAGES_TABLE_HIDDEN_AXES_TRANSACTION_MODULE.search(
        code[semantic_path]
    ) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes semantic API is missing "
            f"table::hidden_axes::transaction: {PAGES_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], frozenset(PAGES_TABLE_HIDDEN_AXES_TRANSACTION_TYPES)
    )
    for name in PAGES_TABLE_HIDDEN_AXES_TRANSACTION_TYPES:
        if name not in transaction_exports:
            violations.append(
                "focused litchi-pages body-table hidden-axes transaction API is missing "
                f"{name}: {PAGES_TABLE_HIDDEN_AXES_SEMANTIC_SOURCE}"
            )

    if "BodyTableSelector" not in _rust_canonical_exports(
        sources[selector_path], frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages body-table hidden-axes public API is missing canonical "
            f"BodyTableSelector: {PAGES_TABLE_HIDDEN_AXES_SELECTOR_SOURCE}"
        )
    if "BodyTableSelector" not in _rust_canonical_exports(
        sources[lib_path], frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages body-table hidden-axes public API is missing root "
            f"BodyTableSelector re-export: {PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[0]}"
        )

    owner_exports = _rust_canonical_exports(
        sources[owner_path], frozenset(PAGES_TABLE_HIDDEN_AXES_CANONICAL_TYPES)
    )
    for name in PAGES_TABLE_HIDDEN_AXES_CANONICAL_TYPES:
        if name not in owner_exports:
            violations.append(
                "focused litchi-pages body-table hidden-axes public API is missing canonical "
                f"type {name}: {PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
            )
    violations.extend(
        _pages_transaction_family_violations(
            root,
            "body_table_hidden_axes",
            "focused litchi-pages body-table hidden-axes public API",
        )
    )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for method in PAGES_TABLE_HIDDEN_AXES_PACKAGE_METHODS:
        declaration = owner_methods.get(method)
        if declaration is None:
            violations.append(
                "focused litchi-pages body-table hidden-axes Package method is missing "
                f"{method}: {PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
            )
            continue
        if method != "apply_body_table_hidden_axes" and "BodyTableSelector" not in declaration:
            violations.append(
                "focused litchi-pages body-table hidden-axes Package method "
                f"{method} must accept selector-first BodyTableSelector: "
                f"{PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
            )
    apply_declaration = owner_methods.get("apply_body_table_hidden_axes")
    if apply_declaration is not None and re.search(r"\bPatch\b", apply_declaration) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes apply method must accept "
            f"Patch: {PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
        )

    edit_methods = {
        name
        for edit_type in ("BodyTableHiddenAxesEdit", "Edit")
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], edit_type
        )
    }
    for name in sorted(PAGES_TABLE_HIDDEN_AXES_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-pages body-table hidden-axes edit is missing "
            f"{name}: {PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
        )
    for method in sorted(PAGES_TABLE_HIDDEN_AXES_FLAT_METHODS & owner_methods.keys()):
        violations.append(
            "focused litchi-pages body-table hidden-axes public API retains flat Package "
            f"method {method}: {PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
        )

    if PUBLIC_PAGES_TABLE_MODULE.search(lib_code) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes public API is missing canonical root "
            f"table module: {PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[0]}"
        )
    if PUBLIC_PAGES_TABLE_HIDDEN_AXES_MODULE.search(code[table_path]) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes public API is missing canonical "
            f"table::hidden_axes module: {PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES[2]}"
        )

    # The focused owner must not recover unsupported source-built edits by
    # delegating to the migration host.  Keep this executable-source check
    # separate from the public declaration leak scan below: a private helper
    # can still create an implicit fallback without changing the facade type
    # signatures.
    for path in (owner_path, semantic_path):
        for pattern in PAGES_TABLE_HIDDEN_AXES_FOCUSED_FALLBACK_PATTERNS:
            match = pattern.search(code[path])
            if match is None:
                continue
            line_number = code[path].count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages body-table hidden-axes owner must not fallback "
                "through the migration host: "
                f"{path.relative_to(root)}:{line_number}"
            )

    # Keep the Buffa sidecar wholly behind the protos crate's hidden module.
    # The generated module must be private and use the expected OUT_DIR file;
    # downstream Pages code may import the strict codec, never generated
    # archive types or Prost's eager message API.
    owner_code = code[owner_path]
    engine_path = root / IWA_TABLE_HIDDEN_AXES_ENGINE_SOURCE
    if engine_path.is_file():
        owner_code += "\n" + _mask_rust_non_code(
            _mask_rust_cfg_test_items(engine_path.read_text(encoding="utf-8"))
        )
    if PAGES_TABLE_HIDDEN_AXES_UID_MAP_OWNER_MARKER.search(owner_code) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes owner is missing the "
            f"canonical UID-map message type {PAGES_TABLE_HIDDEN_AXES_UID_MAP_MESSAGE_TYPE}: "
            f"{PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
        )
    if PAGES_TABLE_HIDDEN_AXES_LEGACY_OWNER_MARKER.search(owner_code) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes owner is missing its explicit "
            f"legacy UID-map marker {PAGES_TABLE_HIDDEN_AXES_LEGACY_UID_MAP_MESSAGE_TYPE}: "
            f"{PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
        )
    if PAGES_TABLE_HIDDEN_AXES_LEGACY_OPT_IN_MARKER.search(owner_code) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes owner must qualify legacy "
            "UID-map support through an explicit allow_legacy path: "
            f"{PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}"
        )
    for pattern in PAGES_TABLE_HIDDEN_AXES_FORBIDDEN_CREATION_MARKERS:
        match = pattern.search(owner_code)
        if match is None:
            continue
        line_number = owner_code.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-pages body-table hidden-axes owner retains forbidden "
            f"owner-creation helper {match.group(0)}: "
            f"{PAGES_TABLE_HIDDEN_AXES_OWNER_SOURCE}:{line_number}"
        )

    codec_path = root / PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE
    codec_lib_path = root / PAGES_TABLE_HIDDEN_AXES_CODEC_PUBLIC_SOURCE
    codec_code = ""
    if not codec_path.is_file():
        violations.append(
            "focused litchi-pages body-table hidden-axes public API is missing strict hidden "
            f"codec source: {PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}"
        )
    else:
        codec_raw = codec_path.read_text(encoding="utf-8")
        codec_source = _mask_rust_cfg_test_items(codec_raw)
        codec_code = _mask_rust_non_code(codec_source)
        if PAGES_TABLE_HIDDEN_AXES_UID_MAP_CODEC_MARKER.search(codec_code) is None:
            violations.append(
                "focused litchi-pages body-table hidden-axes codec is missing the "
                f"canonical UID-map message type {PAGES_TABLE_HIDDEN_AXES_UID_MAP_MESSAGE_TYPE}: "
                f"{PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}"
            )
        if PAGES_TABLE_HIDDEN_AXES_LEGACY_CODEC_MARKER.search(codec_code) is None:
            violations.append(
                "focused litchi-pages body-table hidden-axes codec is missing its explicit "
                f"legacy UID-map marker {PAGES_TABLE_HIDDEN_AXES_LEGACY_UID_MAP_MESSAGE_TYPE}: "
                f"{PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}"
            )
        if PAGES_TABLE_HIDDEN_AXES_LEGACY_OPT_IN_MARKER.search(codec_code) is None:
            violations.append(
                "focused litchi-pages body-table hidden-axes codec must qualify legacy "
                "UID-map support through an explicit allow_legacy path: "
                f"{PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}"
            )
        for api in PAGES_TABLE_HIDDEN_AXES_CODEC_REQUIRED_APIS:
            if re.search(
                rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
                codec_code,
            ) is None:
                violations.append(
                    "focused litchi-pages body-table hidden-axes hidden codec is missing "
                    f"strict API {api}: {PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}"
                )
        # Keep cfg(test) names visible for behavior markers (the overlong
        # unknown-field case is intentionally named in the codec tests), but
        # still mask comments so prose cannot satisfy the ratchet.
        codec_markers = _mask_rust_comments(codec_raw)
        for marker in PAGES_TABLE_HIDDEN_AXES_CODEC_REQUIRED_MARKERS:
            if marker not in codec_markers:
                violations.append(
                    "focused litchi-pages body-table hidden-axes hidden codec is missing "
                    f"strict marker {marker}: {PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}"
                )
        for pattern in PAGES_TABLE_HIDDEN_AXES_CODEC_FORBIDDEN_PATTERNS:
            match = pattern.search(codec_code)
            if match is None:
                continue
            line_number = codec_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages body-table hidden-axes hidden codec retains forbidden "
                f"generated/Prost encoding {match.group(0).strip()}: "
                f"{PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}:{line_number}"
            )
        if re.search(r"(?m)#\s*\[\s*cfg\(\s*test\s*\)\s*\]", codec_raw) is None or re.search(
            r"(?m)#\s*\[\s*test\s*\]", codec_raw
        ) is None:
            violations.append(
                "focused litchi-pages body-table hidden-axes hidden codec is missing cfg(test) #[test] coverage: "
                f"{PAGES_TABLE_HIDDEN_AXES_CODEC_SOURCE}"
            )
    for path, source_code in (
        (owner_path, owner_code),
        (codec_path, codec_code),
    ):
        if not source_code:
            continue
        for match in PAGES_TABLE_HIDDEN_AXES_FORBIDDEN_UID_MAP_LITERAL.finditer(
            source_code
        ):
            line_number = source_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages body-table hidden-axes source uses the forbidden "
                f"Numbers table-data UID-map message type {PAGES_TABLE_HIDDEN_AXES_WRONG_UID_MAP_MESSAGE_TYPE}: "
                f"{path.relative_to(root)}:{line_number}"
            )

    if not codec_lib_path.is_file():
        codec_lib_source = ""
    else:
        codec_lib_source = _mask_rust_cfg_test_items(
            codec_lib_path.read_text(encoding="utf-8")
        )
    codec_lib_code = _mask_rust_non_code(codec_lib_source)
    alias_route = PAGES_TABLE_HIDDEN_AXES_CODEC_ALIAS_ROUTE.search(codec_lib_code)
    if alias_route is not None:
        line_number = codec_lib_code.count("\n", 0, alias_route.start()) + 1
        violations.append(
            "focused litchi-pages body-table hidden-axes codec retains its retired "
            f"alias module {PAGES_TABLE_HIDDEN_AXES_CODEC_ALIAS_MODULE}: "
            f"{PAGES_TABLE_HIDDEN_AXES_CODEC_PUBLIC_SOURCE}:{line_number}"
        )
    hidden_module_pattern = re.compile(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(PAGES_TABLE_HIDDEN_AXES_CODEC_MODULE)}\b"
    )
    if hidden_module_pattern.search(codec_lib_source) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes public API is missing hidden "
            f"codec module {PAGES_TABLE_HIDDEN_AXES_CODEC_MODULE}: "
            f"{PAGES_TABLE_HIDDEN_AXES_CODEC_PUBLIC_SOURCE}"
        )
    generated_module = re.compile(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"mod\s+{re.escape(PAGES_TABLE_HIDDEN_AXES_CODEC_GENERATED_MODULE)}\b"
    )
    if generated_module.search(codec_lib_source) is None:
        violations.append(
            "focused litchi-pages body-table hidden-axes Buffa generated module must remain private: "
            f"{PAGES_TABLE_HIDDEN_AXES_CODEC_PUBLIC_SOURCE}"
        )
    if PAGES_TABLE_HIDDEN_AXES_CODEC_GENERATED_INCLUDE not in codec_lib_source:
        violations.append(
            "focused litchi-pages body-table hidden-axes Buffa generated module has the wrong include: "
            f"{PAGES_TABLE_HIDDEN_AXES_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        rf"(?m)^\s*pub\s+mod\s+{re.escape(PAGES_TABLE_HIDDEN_AXES_CODEC_GENERATED_MODULE)}\b",
        codec_lib_source,
    ):
        violations.append(
            "focused litchi-pages body-table hidden-axes Buffa generated module must not be public: "
            f"{PAGES_TABLE_HIDDEN_AXES_CODEC_PUBLIC_SOURCE}"
        )

    projection_path = root / PAGES_TABLE_HIDDEN_AXES_CODEC_PROJECTION_SOURCE
    if not projection_path.is_file():
        violations.append(
            "focused litchi-pages body-table hidden-axes boundary is missing Buffa projection: "
            f"{PAGES_TABLE_HIDDEN_AXES_CODEC_PROJECTION_SOURCE}"
        )
    else:
        projection_raw = projection_path.read_text(encoding="utf-8")
        projection_code = _mask_rust_comments(projection_raw)
        for marker in PAGES_TABLE_HIDDEN_AXES_PROJECTION_MARKERS:
            if marker not in projection_code:
                violations.append(
                    "focused litchi-pages body-table hidden-axes Buffa projection is missing "
                    f"marker {marker}: {PAGES_TABLE_HIDDEN_AXES_CODEC_PROJECTION_SOURCE}"
                )
        repeated = re.search(r"\brepeated\b", projection_code)
        if repeated is not None:
            line_number = projection_code.count("\n", 0, repeated.start()) + 1
            violations.append(
                "focused litchi-pages body-table hidden-axes Buffa projection must not expose repeated generated fields: "
                f"{PAGES_TABLE_HIDDEN_AXES_CODEC_PROJECTION_SOURCE}:{line_number}"
            )

    # Positive owner tests/fuzz are part of the focused seam.  Keep these
    # checks here (rather than a broad workspace scan) so shared Numbers and
    # Keynote compatibility fixtures do not become accidental Pages routes.
    for test_path in PAGES_TABLE_HIDDEN_AXES_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-pages body-table hidden-axes boundary is missing integration test: "
                f"{test_path}"
            )
        else:
            test_code = _mask_rust_non_code(absolute.read_text(encoding="utf-8"))
            if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
                violations.append(
                    "focused litchi-pages body-table hidden-axes integration test is missing #[test] coverage: "
                    f"{test_path}"
                )
            if not re.search(
                r"\b(?:body_table_hidden_axes|edit_body_table_hidden_axes|apply_body_table_hidden_axes)\b",
                test_code,
            ):
                violations.append(
                    "focused litchi-pages body-table hidden-axes integration test does not exercise the owner: "
                    f"{test_path}"
                )
            if test_path == PAGES_TABLE_HIDDEN_AXES_TEST_SOURCES[0]:
                expected_uid_map = str(PAGES_TABLE_HIDDEN_AXES_UID_MAP_MESSAGE_TYPE)
                wrong_uid_map = str(PAGES_TABLE_HIDDEN_AXES_WRONG_UID_MAP_MESSAGE_TYPE)
                if not re.search(
                    rf"\b(?:MESSAGE_TYPE|message_type|uid_map)[A-Za-z0-9_]*\b"
                    rf"[^;\n]*\b{re.escape(expected_uid_map[:1])}_?"
                    rf"{re.escape(expected_uid_map[1:])}\b",
                    test_code,
                    re.IGNORECASE,
                ):
                    violations.append(
                        "focused litchi-pages body-table hidden-axes integration test must use "
                        f"UID-map message type {PAGES_TABLE_HIDDEN_AXES_UID_MAP_MESSAGE_TYPE}: {test_path}"
                    )
                if re.search(
                    rf"\b(?:MESSAGE_TYPE|message_type|uid_map)[A-Za-z0-9_]*\b"
                    rf"[^;\n]*\b{re.escape(wrong_uid_map[:1])}_?"
                    rf"{re.escape(wrong_uid_map[1:])}\b",
                    test_code,
                    re.IGNORECASE,
                ):
                    violations.append(
                        "focused litchi-pages body-table hidden-axes integration test must not use "
                        f"Numbers table-data message type {PAGES_TABLE_HIDDEN_AXES_WRONG_UID_MAP_MESSAGE_TYPE}: {test_path}"
                    )
    for fuzz_path in PAGES_TABLE_HIDDEN_AXES_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-pages body-table hidden-axes boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-pages body-table hidden-axes fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in PAGES_TABLE_HIDDEN_AXES_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-pages body-table hidden-axes boundary is missing fuzz corpus: "
                f"{corpus}"
            )

    dedicated_sources = {
        root / path
        for path in PAGES_TABLE_HIDDEN_AXES_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_TABLE_HIDDEN_AXES_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_TABLE_HIDDEN_AXES_EXPORT_SOURCES
        if (root / path).is_file()
    }
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(source)
            )
        for declaration, line_number, public_declaration, complete_scope in declarations:
            if not _is_pages_table_hidden_axes_public_declaration(
                declaration, dedicated_source=complete_scope
            ):
                continue
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if public_declaration and "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-pages body-table hidden-axes public API retains a glob re-export: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for identifier in sorted(identifiers):
                reason = _pages_table_hidden_axes_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-pages body-table hidden-axes public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if public_declaration and identifier in PAGES_TABLE_HIDDEN_AXES_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages body-table hidden-axes public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in PAGES_TABLE_HIDDEN_AXES_PUBLIC_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-pages body-table hidden-axes public API exposes raw parameter "
                    f"{match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-pages body-table hidden-axes public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{line_number}"
                )

    # A second route in a sibling module defeats the package boundary even if
    # the canonical package/table exports above look correct.
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                identifiers[:3] == ["pub", "mod", "body_table_hidden_axes"]
                or (
                    identifiers[:3] == ["pub", "mod", "hidden_axes"]
                    and path != table_path
                )
            ):
                violations.append(
                    "focused litchi-pages body-table hidden-axes public API exposes duplicate module: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if {"body_table_hidden_axes", "hidden_axes", "package"} & set(identifiers):
                    violations.append(
                        "focused litchi-pages body-table hidden-axes public API retains root aliases via glob: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                target = target_identifiers[-1] if target_identifiers else ""
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                if (
                    target in PAGES_TABLE_HIDDEN_AXES_ALIAS_TARGETS
                    or "body_table_hidden_axes" in target_identifiers
                    or "hidden_axes" in target_identifiers
                ) and alias and alias != target:
                    if alias in PAGES_TABLE_HIDDEN_AXES_ALIAS_EXEMPT_NAMES:
                        continue
                    violations.append(
                        "focused litchi-pages body-table hidden-axes public API retains alternate alias "
                        f"{alias} for {target}: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_HIDDEN_AXES_ALIAS_TARGETS
                    or "body_table_hidden_axes" in target_identifiers
                    or "hidden_axes" in target_identifiers
                ) and alias != target and alias not in PAGES_TABLE_HIDDEN_AXES_ALIAS_EXEMPT_NAMES:
                    violations.append(
                        "focused litchi-pages body-table hidden-axes public API retains alternate alias "
                        f"{alias} for {target}: {path.relative_to(root)}:{line_number}"
                    )

    return sorted(set(violations))


# Keep the longer body-table spelling available to callers that use the owner
# terminology from the public API.
audit_pages_body_table_hidden_axes_facade_source_topology = (
    audit_pages_table_hidden_axes_facade_source_topology
)


def audit_pages_table_appearance_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one bounded, prepared Pages appearance transaction."""

    if not _pages_table_appearance_owner_present(root):
        return []
    owner_path = root / PAGES_TABLE_APPEARANCE_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner = _mask_rust_non_code(owner_source)
    violations: list[str] = []
    for label, markers in PAGES_TABLE_APPEARANCE_OWNER_MARKER_GROUPS.items():
        alternatives = PAGES_TABLE_APPEARANCE_OWNER_MARKER_ALTERNATIVES.get(label)
        has_markers = all(marker in owner for marker in markers)
        if not has_markers and alternatives is not None:
            has_markers = all(marker in owner for marker in alternatives)
        if not has_markers:
            violations.append(
                "focused litchi-pages body-table appearance owner is missing "
                f"{label} transaction marker: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
    if re.search(
        r"deleted_previews\s*:\s*(?:usize\s*=\s*)?0|deleted_previews\s*=\s*0",
        owner,
    ) is None:
        violations.append(
            "focused litchi-pages body-table appearance owner must preserve previews "
            f"with deleted_previews = 0: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}"
        )
    for label, pattern in PAGES_TABLE_APPEARANCE_FORBIDDEN_OWNER_PATTERNS:
        match = pattern.search(owner)
        if match is not None:
            line_number = owner.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages body-table appearance owner must not perform "
                f"{label}: {PAGES_TABLE_APPEARANCE_OWNER_SOURCE}:{line_number}"
            )
    return sorted(set(violations))


audit_pages_body_table_appearance_facade_source_topology = (
    audit_pages_table_appearance_facade_source_topology
)
audit_pages_body_table_appearance_resource_source_topology = (
    audit_pages_table_appearance_resource_source_topology
)


def _pages_section_text_owner_present(root: Path) -> bool:
    """Return whether the focused Pages section-text owner has landed."""

    return (root / PAGES_SECTION_TEXT_OWNER_SOURCE).is_file()


def _pages_footnote_lifecycle_owner_present(root: Path) -> bool:
    """Return whether the private focused body-footnote owner has landed.

    A package-module declaration can be introduced while an owner is being
    assembled.  Treat the dedicated source file as the activation token so a
    half-landed migration does not turn the dormant ratchet into repository
    noise; once that file exists the private-module and complete-contract
    checks fail closed.
    """

    return (root / PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE).is_file()


def audit_pages_footnote_text_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the focused Pages package footnote-text edit route."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    owner_path = root / PAGES_FOOTNOTE_TEXT_OWNER_SOURCE
    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    methods = {
        name for name, _declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    if PAGES_FOOTNOTE_TEXT_PACKAGE_METHOD in methods:
        return []
    return [
        "focused litchi-pages footnote-text public API is missing Package method "
        f"{PAGES_FOOTNOTE_TEXT_PACKAGE_METHOD}: {PAGES_FOOTNOTE_TEXT_OWNER_SOURCE}"
    ]


def audit_pages_footnote_lifecycle_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the dormant selector-first body-footnote lifecycle contract."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir() or not _pages_footnote_lifecycle_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source,
        frozenset(PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_TYPES),
    )
    for name in PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_TYPES:
        if name in semantic_exports:
            continue
        violations.append(
            "focused litchi-pages body-footnote public API is missing semantic "
            f"footnote::body type {name}: "
            f"{PAGES_FOOTNOTE_LIFECYCLE_SEMANTIC_SOURCE}"
        )

    selector_source = semantic_source
    for name in ("Selector", "Position"):
        if name in _rust_canonical_exports(selector_source, frozenset({name})):
            continue
        violations.append(
            "focused litchi-pages body-footnote public API is missing semantic "
            f"selector/{name}: {PAGES_FOOTNOTE_LIFECYCLE_SELECTOR_SOURCE}"
        )

    owner_path = root / PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE
    package_path = root / PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    if PAGES_PACKAGE_FOOTNOTE_LIFECYCLE_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-pages body-footnote public API is missing private "
            "package owner module: "
            f"{PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_PAGES_PACKAGE_FOOTNOTE_LIFECYCLE_MODULE.finditer(
        package_source
    ):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-pages body-footnote public API exposes public "
            "package::body_footnote module: "
            f"{PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[1]}:{line_number}"
        )
    if not owner_path.is_file():
        violations.append(
            "focused litchi-pages body-footnote public API is missing private "
            f"package owner source: {PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE}"
        )

    dedicated_sources = {
        root / path
        for path in PAGES_FOOTNOTE_LIFECYCLE_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_FOOTNOTE_LIFECYCLE_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES
        if (root / path).is_file()
    }

    owner_parts = [
        _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for path in sorted(dedicated_sources)
        if path != semantic_path and path.is_file()
    ]
    owner_source = "\n".join(owner_parts)
    canonical_exports = _rust_canonical_exports(
        owner_source,
        frozenset(PAGES_FOOTNOTE_LIFECYCLE_CANONICAL_TYPES),
    )
    for name in PAGES_FOOTNOTE_LIFECYCLE_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages body-footnote public API is missing canonical "
            f"package type {name}: {PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE}"
        )

    owner_methods = {
        name
        for source in owner_parts
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            source, "Package"
        )
    }
    for method in PAGES_FOOTNOTE_LIFECYCLE_PACKAGE_METHODS:
        if method in owner_methods:
            continue
        violations.append(
            "focused litchi-pages body-footnote public API is missing Package "
            f"method {method}: {PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE}"
        )
    for method in sorted(PAGES_FOOTNOTE_LIFECYCLE_FLAT_METHODS & owner_methods):
        violations.append(
            "focused litchi-pages body-footnote public API retains flat Package "
            f"method {method}: {PAGES_FOOTNOTE_LIFECYCLE_OWNER_SOURCE}"
        )

    lib_path = root / PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_FOOTNOTE_MODULE.search(_mask_rust_non_code(lib_source)) is None:
        violations.append(
            "focused litchi-pages body-footnote public API is missing canonical "
            f"root footnote module: {PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[0]}"
        )
    violations.extend(
        _pages_transaction_family_violations(
            root,
            "body_footnote",
            "focused litchi-pages body-footnote public API",
        )
    )

    footnote_path = root / PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[2]
    footnote_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(footnote_path.read_text(encoding="utf-8"))
        )
        if footnote_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_FOOTNOTE_BODY_MODULE.search(footnote_source) is None:
        violations.append(
            "focused litchi-pages body-footnote public API is missing canonical "
            f"footnote::body module: {PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[2]}"
        )

    if package_path.is_file():
        package_code = _mask_rust_comments(package_source)
        if PAGES_PACKAGE_FOOTNOTE_LIFECYCLE_MODULE.search(package_code) is None:
            violations.append(
                "focused litchi-pages body-footnote public API is missing private "
                f"package owner module: {PAGES_FOOTNOTE_LIFECYCLE_EXPORT_SOURCES[1]}"
            )

    # Inspect the complete focused source tree after masking cfg(test) items.
    # This catches a new sibling's alias or public module without treating
    # source-built fixtures as production API.
    for path in sorted(source_root.rglob("*.rs")):
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "body_footnote"]:
                violations.append(
                    "focused litchi-pages body-footnote public API exposes duplicate "
                    f"body_footnote module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if {"body_footnote", "footnote", "package"} & set(identifiers):
                    violations.append(
                        "focused litchi-pages body-footnote public API retains root "
                        "aliases via owner glob: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_FOOTNOTE_LIFECYCLE_ALIAS_TARGETS
                    or "body_footnote" in target_identifiers
                    or ("footnote" in target_identifiers and "body" in target_identifiers)
                ) and alias and alias != target:
                    if path == semantic_path and alias in {
                        "Commit",
                        "Diagnostics",
                        "Edit",
                        "Error",
                        "LimitKind",
                        "Patch",
                    }:
                        continue
                    violations.append(
                        "focused litchi-pages body-footnote public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_FOOTNOTE_LIFECYCLE_ALIAS_TARGETS
                    or "body_footnote" in target_identifiers
                    or ("footnote" in target_identifiers and "body" in target_identifiers)
                ) and alias != target:
                    violations.append(
                        "focused litchi-pages body-footnote public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(
                production_source
            )
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(production_source)
            )
        for (
            declaration,
            line_number,
            public_declaration,
            complete_source_scope,
        ) in declarations:
            if not _is_pages_footnote_lifecycle_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages body-footnote public API retains root "
                    f"aliases via owner glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_FOOTNOTE_LIFECYCLE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages body-footnote public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_footnote_lifecycle_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-pages body-footnote public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages body-footnote public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_document_public_api(root: Path = ROOT) -> list[str]:
    """Keep the focused Pages reader public API archive-free and semantic."""

    violations: list[str] = []
    for relative in PAGES_DOCUMENT_PUBLIC_API_SOURCES:
        path = root / relative
        if not path.is_file():
            continue
        source = path.read_text(encoding="utf-8")
        dedicated_source = relative in PAGES_DOCUMENT_DEDICATED_PUBLIC_API_SOURCES
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated_source and not (
                identifiers & PAGES_DOCUMENT_PUBLIC_MARKERS
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
                    "focused litchi-pages document reader public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )

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
                reason = _iwork_public_leak(identifier)
                if reason is None and identifier[:1].isupper():
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
                    "focused litchi-pages document reader rustdoc exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_pages_package_output_api_source_topology(root: Path = ROOT) -> list[str]:
    """Enforce the focused Pages package output boundary.

    Pages retains the exact native source internally for preservation, but the
    supported package API publishes bytes only through ``Package::write_to``.
    This audit deliberately scans the whole Pages source tree after masking
    each ``cfg(test)`` item individually: a test fixture may keep private
    byte-oriented helpers, but it cannot hide a later production declaration
    or re-export.
    """

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        raw_source = path.read_text(encoding="utf-8")
        production_source = _mask_rust_cfg_test_items(raw_source)
        for declaration, line_number in _rust_public_declarations(production_source):
            for forbidden in sorted(PAGES_PACKAGE_OUTPUT_FORBIDDEN_NAMES):
                identifier = next(
                    (
                        match
                        for match in RUST_IDENTIFIER.finditer(declaration)
                        if match.group(1) == forbidden
                    ),
                    None,
                )
                if identifier is None:
                    continue
                identifier_line = line_number + declaration.count(
                    "\n", 0, identifier.start(1)
                )
                violations.append(
                    "focused litchi-pages package output API retains public or "
                    f"re-exported {forbidden}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )

    package_path = root / PAGES_PACKAGE_SOURCE
    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    methods = _rust_public_methods_in_impl(package_source, "Package")
    write_methods = [
        (declaration, line_number)
        for name, declaration, line_number in methods
        if name == PAGES_PACKAGE_OUTPUT_METHOD
    ]
    if not write_methods:
        violations.append(
            "focused litchi-pages package output API is missing production "
            f"Package::{PAGES_PACKAGE_OUTPUT_METHOD}: {PAGES_PACKAGE_SOURCE}"
        )
    elif not any(
        re.search(
            rf"->[\s\S]*\b{re.escape(PAGES_PACKAGE_OUTPUT_ERROR)}\b",
            declaration,
        )
        for declaration, _line_number in write_methods
    ):
        line_number = write_methods[0][1]
        violations.append(
            "focused litchi-pages package output API Package::write_to must "
            f"return {PAGES_PACKAGE_OUTPUT_ERROR}: "
            f"{PAGES_PACKAGE_SOURCE}:{line_number}"
        )

    public_error = False
    for declaration, _line_number in _rust_public_declarations(package_source):
        if re.search(
            rf"\bpub[ \t\r\n]+(?:struct|enum|type)[ \t\r\n]+"
            rf"(?:r#)?{re.escape(PAGES_PACKAGE_OUTPUT_ERROR)}\b",
            declaration,
        ):
            public_error = True
            break
    if not public_error:
        violations.append(
            "focused litchi-pages package output API is missing public "
            f"{PAGES_PACKAGE_OUTPUT_ERROR}: {PAGES_PACKAGE_SOURCE}"
        )

    export_path = root / PAGES_PACKAGE_OUTPUT_EXPORT_SOURCE
    export_source = (
        _mask_rust_cfg_test_items(export_path.read_text(encoding="utf-8"))
        if export_path.is_file()
        else ""
    )
    if PAGES_PACKAGE_OUTPUT_ERROR not in _rust_canonical_exports(
        export_source, frozenset({PAGES_PACKAGE_OUTPUT_ERROR})
    ):
        violations.append(
            "focused litchi-pages package output API is missing root "
            f"{PAGES_PACKAGE_OUTPUT_ERROR} export: "
            f"{PAGES_PACKAGE_OUTPUT_EXPORT_SOURCE}"
        )

    return sorted(set(violations))


def audit_pages_package_no_eager_prost_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep focused Pages production ingress archive-free and Prost-free."""

    violations: list[str] = []

    source_path = root / PAGES_PACKAGE_SOURCE
    if source_path.is_file():
        raw_source = source_path.read_text(encoding="utf-8")
        production_source = _mask_rust_cfg_test_items(raw_source)
        production_code = _mask_rust_non_code(production_source)
        for label, pattern in PAGES_PACKAGE_NO_EAGER_PROST_SOURCE_PATTERNS:
            for match in pattern.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused litchi-pages package production source uses "
                    f"{label}: {PAGES_PACKAGE_SOURCE}:{line_number}"
                )

    manifest_path = root / PAGES_PACKAGE_MANIFEST
    if manifest_path.is_file():
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
            if not normal_dependencies or CARGO_PROST_DEPENDENCY.match(line) is None:
                continue
            violations.append(
                "focused litchi-pages Cargo manifest retains normal prost "
                f"dependency: {PAGES_PACKAGE_MANIFEST}:{line_number}"
            )

    return sorted(set(violations))


def audit_pages_body_chart_arrangement_owner_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the selector-first Pages body-chart arrangement owner."""

    return _audit_focused_chart_arrangement_owner_source_topology(
        root,
        label="litchi-pages body-chart arrangement",
        owner_source=PAGES_BODY_CHART_ARRANGEMENT_OWNER_SOURCE,
        package_source=PAGES_PACKAGE_SOURCE,
        public_source=PAGES_SOURCE_ROOT / "lib.rs",
        module_name=PAGES_BODY_CHART_ARRANGEMENT_PACKAGE_MODULE,
        methods=PAGES_BODY_CHART_ARRANGEMENT_PACKAGE_METHODS,
        selector_types={
            "body_chart_arrangement": PAGES_BODY_CHART_ARRANGEMENT_SELECTOR_TYPES,
            "body_chart_arrangements": (),
            "edit_body_chart_arrangement": PAGES_BODY_CHART_ARRANGEMENT_SELECTOR_TYPES,
            "apply_body_chart_arrangement": (),
        },
        method_markers={
            "body_chart_arrangement": ("resolve_target",),
            "body_chart_arrangements": ("resolve_targets",),
            "edit_body_chart_arrangement": ("resolve_target",),
            "apply_body_chart_arrangement": ("resolve_target", "verify_locality"),
        },
        owner_markers=PAGES_BODY_CHART_ARRANGEMENT_OWNER_MARKERS,
        lazy_codec_markers=PAGES_BODY_CHART_ARRANGEMENT_LAZY_CODEC_MARKERS,
        # ADR 0034: the canonical short names are published only by
        # ``litchi_pages::chart::arrangement``; the shared family check below
        # verifies the owner definitions and that exact re-export.
        canonical_types=(),
        owner_module_visibility="pub(crate)",
    ) + _pages_transaction_family_violations(
        root,
        "body_chart_arrangement",
        "focused litchi-pages body-chart arrangement",
    )


def audit_pages_page_layout_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Reject physical identifiers and implementation types from the layout facade."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
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
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages page-layout public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_table_lock_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Enforce the nested, archive-free Pages table-lock owner API."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    dedicated_sources = {
        root / path
        for path in PAGES_TABLE_LOCK_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in PAGES_TABLE_LOCK_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / PAGES_TABLE_LOCK_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(PAGES_TABLE_LOCK_SEMANTIC_TYPES)
    )
    for name in PAGES_TABLE_LOCK_SEMANTIC_TYPES:
        if name in semantic_exports:
            continue
        violations.append(
            "focused litchi-pages table-lock public API is missing "
            f"canonical table::lock type {name}: {PAGES_TABLE_LOCK_SEMANTIC_SOURCE}"
        )

    package_path = root / PAGES_TABLE_LOCK_PACKAGE_EXPORT_SOURCE
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    owner_path = root / PAGES_TABLE_LOCK_OWNER_SOURCE
    owner_present = owner_path.is_file() or (
        PAGES_PACKAGE_TABLE_LOCK_MODULE.search(package_source) is not None
    )
    if owner_present:
        owner_helper_root = root / PAGES_TABLE_LOCK_OWNER_HELPER_ROOT
        if owner_helper_root.is_dir():
            dedicated_sources.update(owner_helper_root.rglob("*.rs"))
        if owner_path.is_file():
            dedicated_sources.add(owner_path)
        export_sources.add(package_path)
        canonical_exports = _rust_canonical_exports(
            owner_path.read_text(encoding="utf-8") if owner_path.is_file() else "",
            frozenset(PAGES_TABLE_LOCK_CANONICAL_TYPES),
        )
    else:
        canonical_exports = frozenset()
    for name in PAGES_TABLE_LOCK_CANONICAL_TYPES:
        if not owner_present or name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages table-lock public API is missing "
            f"canonical package table-lock type {name}: {PAGES_TABLE_LOCK_OWNER_SOURCE}"
        )
    if owner_present:
        violations.extend(
            _pages_transaction_family_violations(
                root, "table_lock", "focused litchi-pages table-lock public API"
            )
        )

    lib_path = root / PAGES_TABLE_LOCK_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages table-lock public API is missing canonical "
            f"root table module: {PAGES_TABLE_LOCK_EXPORT_SOURCES[0]}"
        )

    table_path = root / PAGES_TABLE_LOCK_EXPORT_SOURCES[2]
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_LOCK_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-pages table-lock public API is missing "
            "canonical table::lock module: "
            f"{PAGES_TABLE_LOCK_EXPORT_SOURCES[2]}"
        )
    if table_path.is_file():
        for declaration, line_number in _rust_public_declarations(table_source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                identifiers[:2] == ["pub", "use"]
                and "lock" in identifiers
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages table-lock public API retains root aliases "
                    "via table::lock glob: "
                    f"{PAGES_TABLE_LOCK_EXPORT_SOURCES[2]}:{line_number}"
                )

    if owner_present:
        if PAGES_PACKAGE_TABLE_LOCK_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-pages table-lock public API is missing "
                f"private package owner module: {PAGES_TABLE_LOCK_PACKAGE_EXPORT_SOURCE}"
            )
        for match in PUBLIC_PAGES_PACKAGE_TABLE_LOCK_MODULE.finditer(package_source):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages table-lock public API exposes duplicate "
                "package::table_lock module: "
                f"{PAGES_TABLE_LOCK_PACKAGE_EXPORT_SOURCE}:{line_number}"
            )
        if not owner_path.is_file():
            violations.append(
                "focused litchi-pages table-lock public API is missing private "
                f"package owner source: {PAGES_TABLE_LOCK_OWNER_SOURCE}"
            )
        owner_source = (
            owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
        )
        owner_methods = {
            name
            for declaration, _line_number in _rust_public_declarations(owner_source)
            for name, _nested_line in _rust_function_declarations(declaration)
        }
        for method in PAGES_TABLE_LOCK_PACKAGE_METHODS:
            if method in owner_methods:
                continue
            violations.append(
                "focused litchi-pages table-lock public API is missing Package "
                f"method {method}: {PAGES_TABLE_LOCK_OWNER_SOURCE}"
            )
        for method in sorted(PAGES_TABLE_LOCK_FLAT_METHODS & owner_methods):
            violations.append(
                "focused litchi-pages table-lock public API retains flat Package "
                f"method {method}: {PAGES_TABLE_LOCK_OWNER_SOURCE}"
            )

    # Search the whole facade, not only today's three export files. A new
    # sibling module must not create an alias that can later be globbed into
    # the crate root without crossing this ratchet.
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "table_lock"]:
                violations.append(
                    "focused litchi-pages table-lock public API exposes duplicate "
                    f"table_lock module: {path.relative_to(root)}:{line_number}"
                )

            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                wildcard_owner = (
                    "package" in identifiers
                    or "table_lock" in identifiers
                    or (path == table_path and "lock" in identifiers)
                )
                if wildcard_owner:
                    violations.append(
                        "focused litchi-pages table-lock public API retains root "
                        f"aliases via owner glob: {path.relative_to(root)}:{line_number}"
                    )

            if "as" in identifiers and identifiers[:2] == ["pub", "use"]:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                table_lock_target = bool(
                    set(target_identifiers) & PAGES_TABLE_LOCK_ALIAS_TARGETS
                ) or ("lock" in target_identifiers and "State" in target_identifiers)
                target = target_identifiers[-1] if target_identifiers else ""
                if table_lock_target and alias and alias != target:
                    violations.append(
                        "focused litchi-pages table-lock public API retains alternate "
                        f"alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                table_lock_target = bool(
                    set(target_identifiers) & PAGES_TABLE_LOCK_ALIAS_TARGETS
                ) or ("lock" in target_identifiers and "State" in target_identifiers)
                # The ``BodyTableLockState`` alias is deleted: any type alias
                # of ``State`` in ``table::lock`` is now an alternate route.
                if path == semantic_path and "State" in target_identifiers:
                    table_lock_target = True
                if table_lock_target:
                    target = target_identifiers[-1]
                    if alias != target:
                        violations.append(
                            "focused litchi-pages table-lock public API retains "
                            f"alternate alias {alias} for {target}: "
                            f"{path.relative_to(root)}:{line_number}"
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
            if not _is_pages_table_lock_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                path == semantic_path
                and tuple(identifiers)
                == ("pub", "use", "litchi_iwa_common", "table", "lock", "State")
            ):
                continue
            public_use_or_type = identifiers[:2] in (["pub", "type"], ["pub", "use"])
            if (
                public_declaration
                and public_use_or_type
                and "*" in declaration
                and path in export_sources
            ):
                violations.append(
                    "focused litchi-pages table-lock public API retains root aliases "
                    "via table::lock glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_TABLE_LOCK_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages table-lock public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_table_lock_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-pages table-lock public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages table-lock public API exposes raw "
                    f"byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_table_title_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Enforce the selector-first, archive-free Pages table-title API."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    dedicated_sources = {
        root / path
        for path in PAGES_TABLE_TITLE_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_TABLE_TITLE_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_TABLE_TITLE_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / PAGES_TABLE_TITLE_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8") if semantic_path.is_file() else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(PAGES_TABLE_TITLE_SEMANTIC_TYPES)
    )
    for name in PAGES_TABLE_TITLE_SEMANTIC_TYPES:
        if name in semantic_exports:
            continue
        violations.append(
            "focused litchi-pages table-title public API is missing "
            f"semantic table::title type {name}: "
            f"{PAGES_TABLE_TITLE_SEMANTIC_SOURCE}"
        )

    selector_path = root / PAGES_TABLE_TITLE_SELECTOR_SOURCE
    selector_source = (
        _mask_rust_non_code(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    if "BodyTableSelector" not in _rust_canonical_exports(
        selector_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-title public API is missing canonical "
            f"BodyTableSelector: {PAGES_TABLE_TITLE_SELECTOR_SOURCE}"
        )

    owner_path = root / PAGES_TABLE_TITLE_OWNER_SOURCE
    owner_present = owner_path.is_file() or any(
        PAGES_PACKAGE_TABLE_TITLE_MODULE.search(
            _mask_rust_non_code(
                path.read_text(encoding="utf-8")
            )
        )
        for path in (root / PAGES_TABLE_TITLE_EXPORT_SOURCES[1],)
        if path.is_file()
    )
    if owner_present:
        if owner_path.is_file():
            dedicated_sources.add(owner_path)
            export_sources.add(root / PAGES_TABLE_TITLE_EXPORT_SOURCES[1])
        canonical_exports = _rust_canonical_exports(
            owner_path.read_text(encoding="utf-8") if owner_path.is_file() else "",
            frozenset(PAGES_TABLE_TITLE_CANONICAL_TYPES),
        )
    else:
        canonical_exports = frozenset()
    for name in PAGES_TABLE_TITLE_CANONICAL_TYPES:
        if owner_present and name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages table-title public API is missing "
            f"canonical package type {name}: {PAGES_TABLE_TITLE_OWNER_SOURCE}"
        )
    violations.extend(
        _pages_transaction_family_violations(
            root, "body_table_title", "focused litchi-pages table-title public API"
        )
    )
    if owner_present:
        owner_source = (
            owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
        )
        owner_methods = {
            name
            for declaration, _line_number in _rust_public_declarations(owner_source)
            for name, _nested_line in _rust_function_declarations(declaration)
        }
        for method in PAGES_TABLE_TITLE_PACKAGE_METHODS:
            if method in owner_methods:
                continue
            violations.append(
                "focused litchi-pages table-title public API is missing Package "
                f"method {method}: {PAGES_TABLE_TITLE_OWNER_SOURCE}"
            )
        for method in sorted(PAGES_TABLE_TITLE_FLAT_METHODS & owner_methods):
            violations.append(
                "focused litchi-pages table-title public API retains flat Package "
                f"method {method}: {PAGES_TABLE_TITLE_OWNER_SOURCE}"
            )

    lib_path = root / PAGES_TABLE_TITLE_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages table-title public API is missing canonical "
            f"root table module: {PAGES_TABLE_TITLE_EXPORT_SOURCES[0]}"
        )
    if "BodyTableSelector" not in _rust_canonical_exports(
        lib_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-title public API is missing root "
            f"BodyTableSelector re-export: {PAGES_TABLE_TITLE_EXPORT_SOURCES[0]}"
        )

    table_path = root / PAGES_TABLE_TITLE_EXPORT_SOURCES[2]
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_TITLE_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-pages table-title public API is missing canonical "
            "table::title module: "
            f"{PAGES_TABLE_TITLE_EXPORT_SOURCES[2]}"
        )

    package_path = root / PAGES_TABLE_TITLE_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    if owner_present and PAGES_PACKAGE_TABLE_TITLE_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-pages table-title public API is missing private "
            f"package owner module: {PAGES_TABLE_TITLE_EXPORT_SOURCES[1]}"
        )
    if owner_present:
        for match in PUBLIC_PAGES_PACKAGE_TABLE_TITLE_MODULE.finditer(package_source):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages table-title public API exposes duplicate "
                "package::body_table_title module: "
                f"{PAGES_TABLE_TITLE_EXPORT_SOURCES[1]}:{line_number}"
            )

    # The complete facade tree is scanned so a newly added sibling module
    # cannot introduce a second export path or an alias hidden from the
    # established root/package/table files. cfg(test) items are deliberately
    # masked one item at a time; production declarations after test helpers
    # remain in the audit slice.
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "body_table_title"]:
                violations.append(
                    "focused litchi-pages table-title public API exposes duplicate "
                    f"body_table_title module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                owner_glob = (
                    "body_table_title" in identifiers
                    or "title" in identifiers
                    or "package" in identifiers
                )
                if owner_glob:
                    glob_label = (
                        "table::title glob" if path == table_path else "owner glob"
                    )
                    violations.append(
                        "focused litchi-pages table-title public API retains root "
                        f"aliases via {glob_label}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_TITLE_ALIAS_TARGETS
                    or "body_table_title" in target_identifiers
                ) and alias and alias != target:
                    # ADR 0034: ``table::title`` re-exports the owner's
                    # short names directly, so no rename is permitted here.
                    violations.append(
                        "focused litchi-pages table-title public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_TITLE_ALIAS_TARGETS
                    or "body_table_title" in target_identifiers
                ) and alias != target:
                    violations.append(
                        "focused litchi-pages table-title public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
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
            if not _is_pages_table_title_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages table-title public API retains root "
                    f"aliases via table::title glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_TABLE_TITLE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages table-title public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_table_title_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-pages table-title public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages table-title public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_table_headers_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Enforce the selector-first, archive-free Pages table-header API."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    dedicated_sources = {
        root / path
        for path in PAGES_TABLE_HEADERS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_TABLE_HEADERS_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_TABLE_HEADERS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / PAGES_TABLE_HEADERS_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(PAGES_TABLE_HEADERS_SEMANTIC_TYPES)
    )
    for name in PAGES_TABLE_HEADERS_SEMANTIC_TYPES:
        if name in semantic_exports:
            continue
        violations.append(
            "focused litchi-pages table-headers public API is missing "
            f"semantic table::headers type {name}: "
            f"{PAGES_TABLE_HEADERS_SEMANTIC_SOURCE}"
        )

    selector_path = root / PAGES_TABLE_HEADERS_SELECTOR_SOURCE
    selector_source = (
        _mask_rust_non_code(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    if "BodyTableSelector" not in _rust_canonical_exports(
        selector_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-headers public API is missing canonical "
            f"BodyTableSelector: {PAGES_TABLE_HEADERS_SELECTOR_SOURCE}"
        )

    package_path = root / PAGES_TABLE_HEADERS_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    owner_path = root / PAGES_TABLE_HEADERS_OWNER_SOURCE
    owner_present = owner_path.is_file() or (
        PAGES_PACKAGE_TABLE_HEADERS_MODULE.search(package_source) is not None
    )
    if owner_present:
        if helper_root.is_dir():
            dedicated_sources.update(helper_root.rglob("*.rs"))
        if owner_path.is_file():
            dedicated_sources.add(owner_path)
        export_sources.add(package_path)
        canonical_exports = _rust_canonical_exports(
            owner_path.read_text(encoding="utf-8") if owner_path.is_file() else "",
            frozenset(PAGES_TABLE_HEADERS_CANONICAL_TYPES),
        )
    else:
        canonical_exports = frozenset()
    for name in PAGES_TABLE_HEADERS_CANONICAL_TYPES:
        if owner_present and name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages table-headers public API is missing "
            f"canonical package type {name}: {PAGES_TABLE_HEADERS_OWNER_SOURCE}"
        )
    violations.extend(
        _pages_transaction_family_violations(
            root, "body_table_headers", "focused litchi-pages table-headers public API"
        )
    )

    lib_path = root / PAGES_TABLE_HEADERS_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages table-headers public API is missing canonical "
            f"root table module: {PAGES_TABLE_HEADERS_EXPORT_SOURCES[0]}"
        )
    if "BodyTableSelector" not in _rust_canonical_exports(
        lib_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-headers public API is missing root "
            f"BodyTableSelector re-export: {PAGES_TABLE_HEADERS_EXPORT_SOURCES[0]}"
        )

    table_path = root / PAGES_TABLE_HEADERS_EXPORT_SOURCES[2]
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_HEADERS_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-pages table-headers public API is missing canonical "
            "table::headers module: "
            f"{PAGES_TABLE_HEADERS_EXPORT_SOURCES[2]}"
        )

    if owner_present:
        if PAGES_PACKAGE_TABLE_HEADERS_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-pages table-headers public API is missing private "
                f"package owner module: {PAGES_TABLE_HEADERS_EXPORT_SOURCES[1]}"
            )
        for match in PUBLIC_PAGES_PACKAGE_TABLE_HEADERS_MODULE.finditer(package_source):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages table-headers public API exposes duplicate "
                "package::body_table_headers module: "
                f"{PAGES_TABLE_HEADERS_EXPORT_SOURCES[1]}:{line_number}"
            )
        if not owner_path.is_file():
            violations.append(
                "focused litchi-pages table-headers public API is missing private "
                f"package owner source: {PAGES_TABLE_HEADERS_OWNER_SOURCE}"
            )
        owner_source = owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
        owner_methods = {
            name
            for declaration, _line_number in _rust_public_declarations(owner_source)
            for name, _nested_line in _rust_function_declarations(declaration)
        }
        for method in PAGES_TABLE_HEADERS_PACKAGE_METHODS:
            if method in owner_methods:
                continue
            violations.append(
                "focused litchi-pages table-headers public API is missing Package "
                f"method {method}: {PAGES_TABLE_HEADERS_OWNER_SOURCE}"
            )
        for method in sorted(PAGES_TABLE_HEADERS_FLAT_METHODS & owner_methods):
            violations.append(
                "focused litchi-pages table-headers public API retains flat Package "
                f"method {method}: {PAGES_TABLE_HEADERS_OWNER_SOURCE}"
            )

    # Scan every facade file so a new sibling module cannot hide a second
    # public route.  cfg(test) declarations are masked item-by-item.
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "body_table_headers"]:
                violations.append(
                    "focused litchi-pages table-headers public API exposes duplicate "
                    f"body_table_headers module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if (
                    "body_table_headers" in identifiers
                    or "headers" in identifiers
                    or "package" in identifiers
                ):
                    glob_label = (
                        "table::headers glob" if path == table_path else "owner glob"
                    )
                    violations.append(
                        "focused litchi-pages table-headers public API retains root "
                        f"aliases via {glob_label}: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_HEADERS_ALIAS_TARGETS
                    or "body_table_headers" in target_identifiers
                    or "headers" in target_identifiers
                ) and alias and alias != target:
                    # ADR 0034 deleted the ``BodyTableHeaderSettings`` and
                    # prefixed aliases; the short names are re-exported as-is.
                    violations.append(
                        "focused litchi-pages table-headers public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_HEADERS_ALIAS_TARGETS
                    or "body_table_headers" in target_identifiers
                    or "headers" in target_identifiers
                ) and alias != target:
                    violations.append(
                        "focused litchi-pages table-headers public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
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
            if not _is_pages_table_headers_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages table-headers public API retains root "
                    f"aliases via table::headers glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_TABLE_HEADERS_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages table-headers public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_table_headers_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-pages table-headers public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages table-headers public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_table_dimension_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Pages table-dimension owner.

    The owner is deliberately feature-gated by its private package source.
    This lets the migration inventory land before the compatibility host is
    retired, while still making the complete facade contract fail closed as
    soon as the owner is introduced.
    """

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    owner_path = root / PAGES_TABLE_DIMENSION_OWNER_SOURCE
    package_path = root / PAGES_TABLE_DIMENSION_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    owner_present = owner_path.is_file() or (
        PAGES_PACKAGE_TABLE_DIMENSION_MODULE.search(package_source) is not None
    )
    if not owner_present:
        return []

    dedicated_sources = {
        root / path
        for path in PAGES_TABLE_DIMENSION_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_TABLE_DIMENSION_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_TABLE_DIMENSION_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / PAGES_TABLE_DIMENSION_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source, frozenset(PAGES_TABLE_DIMENSION_SEMANTIC_TYPES)
    )
    for name in PAGES_TABLE_DIMENSION_SEMANTIC_TYPES:
        if name in semantic_exports:
            continue
        violations.append(
            "focused litchi-pages table-dimension public API is missing "
            f"semantic table::dimension type {name}: "
            f"{PAGES_TABLE_DIMENSION_SEMANTIC_SOURCE}"
        )

    selector_path = root / PAGES_TABLE_DIMENSION_SELECTOR_SOURCE
    selector_source = (
        _mask_rust_non_code(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    if "BodyTableSelector" not in _rust_canonical_exports(
        selector_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-dimension public API is missing canonical "
            f"BodyTableSelector: {PAGES_TABLE_DIMENSION_SELECTOR_SOURCE}"
        )

    owner_source = owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
    canonical_exports = _rust_canonical_exports(
        owner_source, frozenset(PAGES_TABLE_DIMENSION_CANONICAL_TYPES)
    )
    for name in PAGES_TABLE_DIMENSION_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages table-dimension public API is missing "
            f"canonical package type {name}: {PAGES_TABLE_DIMENSION_OWNER_SOURCE}"
        )
    violations.extend(
        _pages_transaction_family_violations(
            root, "body_table_dimension", "focused litchi-pages table-dimension public API"
        )
    )

    owner_methods = {
        name
        for declaration, _line_number in _rust_public_declarations(owner_source)
        for name, _nested_line in _rust_function_declarations(declaration)
    }
    for method in PAGES_TABLE_DIMENSION_PACKAGE_METHODS:
        if method in owner_methods:
            continue
        violations.append(
            "focused litchi-pages table-dimension public API is missing Package "
            f"method {method}: {PAGES_TABLE_DIMENSION_OWNER_SOURCE}"
        )
    for method in sorted(PAGES_TABLE_DIMENSION_FLAT_METHODS & owner_methods):
        violations.append(
            "focused litchi-pages table-dimension public API retains flat Package "
            f"method {method}: {PAGES_TABLE_DIMENSION_OWNER_SOURCE}"
        )

    lib_path = root / PAGES_TABLE_DIMENSION_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages table-dimension public API is missing canonical "
            f"root table module: {PAGES_TABLE_DIMENSION_EXPORT_SOURCES[0]}"
        )
    if "BodyTableSelector" not in _rust_canonical_exports(
        lib_source, frozenset({"BodyTableSelector"})
    ):
        violations.append(
            "focused litchi-pages table-dimension public API is missing root "
            f"BodyTableSelector re-export: {PAGES_TABLE_DIMENSION_EXPORT_SOURCES[0]}"
        )

    table_path = root / PAGES_TABLE_DIMENSION_EXPORT_SOURCES[2]
    table_source = (
        _mask_rust_non_code(table_path.read_text(encoding="utf-8"))
        if table_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_TABLE_DIMENSION_MODULE.search(table_source) is None:
        violations.append(
            "focused litchi-pages table-dimension public API is missing canonical "
            f"table::dimension module: {PAGES_TABLE_DIMENSION_EXPORT_SOURCES[2]}"
        )
    if PAGES_PACKAGE_TABLE_DIMENSION_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-pages table-dimension public API is missing private "
            f"package owner module: {PAGES_TABLE_DIMENSION_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_PAGES_PACKAGE_TABLE_DIMENSION_MODULE.finditer(package_source):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-pages table-dimension public API exposes duplicate "
            f"package::body_table_dimension module: {PAGES_TABLE_DIMENSION_EXPORT_SOURCES[1]}:{line_number}"
        )

    # Scan every facade file after masking cfg(test) items individually.  A
    # production sibling cannot launder an alternate alias through a glob.
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "body_table_dimension"]:
                violations.append(
                    "focused litchi-pages table-dimension public API exposes duplicate "
                    f"body_table_dimension module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if {
                    "body_table_dimension",
                    "dimension",
                    "package",
                } & set(identifiers):
                    glob_label = (
                        "table::dimension glob" if path == table_path else "owner glob"
                    )
                    violations.append(
                        "focused litchi-pages table-dimension public API retains root "
                        f"aliases via {glob_label}: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_DIMENSION_ALIAS_TARGETS
                    or "body_table_dimension" in target_identifiers
                    or "dimension" in target_identifiers
                ) and alias and alias != target:
                    # ADR 0034: no rename of the owner's short names.
                    violations.append(
                        "focused litchi-pages table-dimension public API retains "
                        f"alternate alias {alias} for {target}: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_TABLE_DIMENSION_ALIAS_TARGETS
                    or "body_table_dimension" in target_identifiers
                    or "dimension" in target_identifiers
                ) and alias != target:
                    violations.append(
                        "focused litchi-pages table-dimension public API retains "
                        f"alternate alias {alias} for {target}: {path.relative_to(root)}:{line_number}"
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
            if not _is_pages_table_dimension_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages table-dimension public API retains root "
                    f"aliases via table::dimension glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_TABLE_DIMENSION_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages table-dimension public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_table_dimension_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-pages table-dimension public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages table-dimension public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_header_footer_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Pages header/footer owner.

    This ratchet is dormant until ``package/header_footer_text.rs`` exists.
    The existing ``header_footer::{Template, Kind}`` role vocabulary remains
    useful during that migration, but it does not by itself constitute a
    package owner.  Once activated, all public transaction names, selector
    exports, private-module boundaries, alias routes, and raw physical/wire
    leaks are checked across the complete Pages source tree.
    """

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir() or not _pages_header_footer_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / PAGES_HEADER_FOOTER_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    semantic_exports = _rust_canonical_exports(
        semantic_source,
        frozenset(PAGES_HEADER_FOOTER_SEMANTIC_TYPES),
    )
    for name in PAGES_HEADER_FOOTER_SEMANTIC_TYPES:
        if name in semantic_exports:
            continue
        violations.append(
            "focused litchi-pages header/footer public API is missing semantic "
            f"header_footer type {name}: {PAGES_HEADER_FOOTER_SEMANTIC_SOURCE}"
        )

    selector_path = root / PAGES_HEADER_FOOTER_SELECTOR_SOURCE
    selector_source = (
        _mask_rust_cfg_test_items(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    for name in PAGES_HEADER_FOOTER_SELECTOR_TYPES:
        if name in _rust_canonical_exports(selector_source, frozenset({name})):
            continue
        violations.append(
            "focused litchi-pages header/footer public API is missing canonical "
            f"{name}: {PAGES_HEADER_FOOTER_SELECTOR_SOURCE}"
        )

    owner_path = root / PAGES_HEADER_FOOTER_OWNER_SOURCE
    package_path = root / PAGES_HEADER_FOOTER_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    if PAGES_PACKAGE_HEADER_FOOTER_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-pages header/footer public API is missing private "
            f"package owner module: {PAGES_HEADER_FOOTER_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_PAGES_PACKAGE_HEADER_FOOTER_MODULE.finditer(package_source):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-pages header/footer public API exposes public "
            f"package::header_footer_text module: {PAGES_HEADER_FOOTER_EXPORT_SOURCES[1]}:{line_number}"
        )

    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        owner_source, frozenset(PAGES_HEADER_FOOTER_CANONICAL_TYPES)
    )
    for name in PAGES_HEADER_FOOTER_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages header/footer public API is missing canonical "
            f"package type {name}: {PAGES_HEADER_FOOTER_OWNER_SOURCE}"
        )

    owner_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in PAGES_HEADER_FOOTER_PACKAGE_METHODS:
        if method in owner_methods:
            continue
        violations.append(
            "focused litchi-pages header/footer public API is missing Package "
            f"method {method}: {PAGES_HEADER_FOOTER_OWNER_SOURCE}"
        )
    for method in sorted(PAGES_HEADER_FOOTER_FLAT_METHODS & owner_methods):
        violations.append(
            "focused litchi-pages header/footer public API retains flat Package "
            f"method {method}: {PAGES_HEADER_FOOTER_OWNER_SOURCE}"
        )

    lib_path = root / PAGES_HEADER_FOOTER_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        )
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_HEADER_FOOTER_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages header/footer public API is missing canonical "
            f"root header_footer module: {PAGES_HEADER_FOOTER_EXPORT_SOURCES[0]}"
        )
    violations.extend(
        _pages_transaction_family_violations(
            root,
            "header_footer_text",
            "focused litchi-pages header/footer public API",
        )
    )
    if "HeaderFooterSelector" not in _rust_canonical_exports(
        lib_source, frozenset({"HeaderFooterSelector"})
    ):
        violations.append(
            "focused litchi-pages header/footer public API is missing root "
            f"HeaderFooterSelector re-export: {PAGES_HEADER_FOOTER_EXPORT_SOURCES[0]}"
        )

    if PUBLIC_PAGES_HEADER_FOOTER_MODULE.search(
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(
                semantic_path.read_text(encoding="utf-8")
                if semantic_path.is_file()
                else ""
            )
        )
    ) is not None:
        # A semantic source defining a second public module is an alternate
        # route, not the canonical root module.
        violations.append(
            "focused litchi-pages header/footer semantic source exposes duplicate "
            f"header_footer module: {PAGES_HEADER_FOOTER_SEMANTIC_SOURCE}"
        )

    dedicated_sources = {
        root / path
        for path in PAGES_HEADER_FOOTER_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_HEADER_FOOTER_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_HEADER_FOOTER_EXPORT_SOURCES
        if (root / path).is_file()
    }

    # Every sibling is checked after cfg(test) item masking.  This catches a
    # second public alias/module even when it is declared after a fixture.
    root_header_module_seen = False
    for path in sorted(source_root.rglob("*.rs")):
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                identifiers[:3] == ["pub", "mod", "header_footer"]
                and (path != lib_path or root_header_module_seen)
            ):
                violations.append(
                    "focused litchi-pages header/footer public API exposes duplicate "
                    f"header_footer module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:3] == ["pub", "mod", "header_footer"] and path == lib_path:
                root_header_module_seen = True
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if {"header_footer_text", "header_footer", "package", "selector"} & set(
                    identifiers
                ):
                    violations.append(
                        "focused litchi-pages header/footer public API retains root "
                        f"aliases via owner glob: {path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_HEADER_FOOTER_ALIAS_TARGETS
                    or "header_footer_text" in target_identifiers
                    or "header_footer" in target_identifiers
                ) and alias and alias != target:
                    violations.append(
                        "focused litchi-pages header/footer public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_HEADER_FOOTER_ALIAS_TARGETS
                    or "header_footer_text" in target_identifiers
                    or "header_footer" in target_identifiers
                ) and alias != target:
                    violations.append(
                        "focused litchi-pages header/footer public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(production_source)
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(production_source)
            )
        for declaration, line_number, public_declaration, complete_scope in declarations:
            if not _is_pages_header_footer_public_declaration(
                declaration, dedicated_source=complete_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages header/footer public API retains root "
                    f"aliases via owner glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_HEADER_FOOTER_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages header/footer public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_header_footer_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-pages header/footer public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages header/footer public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_section_text_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Pages section-text owner."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir() or not _pages_section_text_owner_present(root):
        return []

    violations: list[str] = []
    semantic_path = root / PAGES_SECTION_TEXT_SEMANTIC_SOURCE
    semantic_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        )
        if semantic_path.is_file()
        else ""
    )
    for name in PAGES_SECTION_TEXT_SEMANTIC_TYPES:
        if name in _rust_canonical_exports(semantic_source, frozenset({name})):
            continue
        violations.append(
            "focused litchi-pages section-text public API is missing semantic "
            f"selector type {name}: {PAGES_SECTION_TEXT_SEMANTIC_SOURCE}"
        )

    owner_path = root / PAGES_SECTION_TEXT_OWNER_SOURCE
    owner_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        )
        if owner_path.is_file()
        else ""
    )
    owner_exports = _rust_canonical_exports(
        owner_source, frozenset(PAGES_SECTION_TEXT_CANONICAL_TYPES)
    )
    for name in PAGES_SECTION_TEXT_CANONICAL_TYPES:
        if name in owner_exports:
            continue
        violations.append(
            "focused litchi-pages section-text public API is missing canonical "
            f"package type {name}: {PAGES_SECTION_TEXT_OWNER_SOURCE}"
        )

    owner_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for method in PAGES_SECTION_TEXT_PACKAGE_METHODS:
        if method not in owner_methods:
            violations.append(
                "focused litchi-pages section-text public API is missing Package "
                f"method {method}: {PAGES_SECTION_TEXT_OWNER_SOURCE}"
            )
    for method in sorted(PAGES_SECTION_TEXT_FLAT_METHODS & owner_methods):
        violations.append(
            "focused litchi-pages section-text public API retains flat Package "
            f"method {method}: {PAGES_SECTION_TEXT_OWNER_SOURCE}"
        )

    lib_path = root / PAGES_SECTION_TEXT_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        )
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_SECTION_TEXT_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages section-text public API is missing canonical "
            f"root selector module: {PAGES_SECTION_TEXT_EXPORT_SOURCES[0]}"
        )
    if "SectionSelector" not in _rust_canonical_exports(
        lib_source, frozenset({"SectionSelector"})
    ):
        violations.append(
            "focused litchi-pages section-text public API is missing root "
            f"SectionSelector re-export: {PAGES_SECTION_TEXT_EXPORT_SOURCES[0]}"
        )
    violations.extend(
        _pages_transaction_family_violations(
            root,
            "section_text",
            "focused litchi-pages section-text public API",
        )
    )

    package_path = root / PAGES_SECTION_TEXT_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    if PAGES_PACKAGE_SECTION_TEXT_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-pages section-text public API is missing private "
            f"package owner module: {PAGES_SECTION_TEXT_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_PAGES_PACKAGE_SECTION_TEXT_MODULE.finditer(package_source):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-pages section-text public API exposes public "
            f"package::section_text module: {PAGES_SECTION_TEXT_EXPORT_SOURCES[1]}:{line_number}"
        )

    # Inspect the complete focused source tree after masking each cfg(test)
    # item.  This prevents a sibling module or root glob from laundering the
    # selector or section-text implementation into a second public route.
    root_selector_module_seen = False
    for path in sorted(source_root.rglob("*.rs")):
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        for declaration, line_number in _rust_public_declarations(production_source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if identifiers[:3] == ["pub", "mod", "selector"]:
                if path != lib_path or root_selector_module_seen:
                    violations.append(
                        "focused litchi-pages section-text public API exposes duplicate "
                        f"selector module: {path.relative_to(root)}:{line_number}"
                    )
                if path == lib_path:
                    root_selector_module_seen = True
            if identifiers[:3] == ["pub", "mod", "section_text"]:
                violations.append(
                    "focused litchi-pages section-text public API exposes duplicate "
                    f"section_text module: {path.relative_to(root)}:{line_number}"
                )
            if identifiers[:2] == ["pub", "use"] and "*" in declaration:
                if {"section_text", "selector", "package"} & set(identifiers):
                    violations.append(
                        "focused litchi-pages section-text public API retains root "
                        "aliases via owner glob: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "use"] and "as" in identifiers:
                alias_index = identifiers.index("as")
                target_identifiers = identifiers[2:alias_index]
                alias = (
                    identifiers[alias_index + 1]
                    if alias_index + 1 < len(identifiers)
                    else ""
                )
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_SECTION_TEXT_ALIAS_TARGETS
                    or "section_text" in target_identifiers
                    or "selector" in target_identifiers
                ) and alias and alias != target:
                    violations.append(
                        "focused litchi-pages section-text public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )
            if identifiers[:2] == ["pub", "type"] and len(identifiers) >= 4:
                alias = identifiers[2]
                target_identifiers = identifiers[3:]
                target = target_identifiers[-1] if target_identifiers else ""
                if (
                    target in PAGES_SECTION_TEXT_ALIAS_TARGETS
                    or "section_text" in target_identifiers
                    or "selector" in target_identifiers
                ) and alias != target:
                    violations.append(
                        "focused litchi-pages section-text public API retains "
                        f"alternate alias {alias} for {target}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    dedicated_sources = {
        root / path
        for path in PAGES_SECTION_TEXT_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    helper_root = root / PAGES_SECTION_TEXT_OWNER_HELPER_ROOT
    if helper_root.is_dir():
        dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_SECTION_TEXT_EXPORT_SOURCES
        if (root / path).is_file()
    }
    for path in sorted(dedicated_sources | export_sources):
        dedicated_source = path in dedicated_sources
        production_source = _mask_rust_cfg_test_items(
            path.read_text(encoding="utf-8")
        )
        declarations = [
            (declaration, line_number, True, dedicated_source)
            for declaration, line_number in _rust_public_declarations(
                production_source
            )
        ]
        if dedicated_source:
            declarations.extend(
                (declaration, line_number, False, False)
                for declaration, line_number in _rust_impl_headers(production_source)
            )
        for declaration, line_number, public_declaration, complete_scope in declarations:
            if not _is_pages_section_text_public_declaration(
                declaration, dedicated_source=complete_scope
            ):
                continue
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages section-text public API retains root "
                    f"aliases via owner glob: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_SECTION_TEXT_FLAT_ALIASES:
                    # The canonical transaction names are public by design in
                    # the focused owner; only alternate/legacy spellings are
                    # flat aliases.
                    if identifier not in PAGES_SECTION_TEXT_CANONICAL_TYPES:
                        violations.append(
                            "focused litchi-pages section-text public API retains flat "
                            f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                        )
                reason = _pages_section_text_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-pages section-text public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
            if "fn" in identifiers and re.search(
                r"\:\s*(?:&\s*)?u64\b", declaration
            ):
                violations.append(
                    "focused litchi-pages section-text public API exposes raw u64: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in PAGES_SECTION_TEXT_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-pages section-text public API exposes raw identifier "
                    f"parameter {match.group(0).strip()}: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages section-text public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_document_settings_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Reject physical details from the combined document-settings facade."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in PAGES_DOCUMENT_SETTINGS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in PAGES_DOCUMENT_SETTINGS_EXPORT_SOURCES
        if (root / path).is_file()
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
            if not _is_pages_document_settings_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in PAGES_DOCUMENT_SETTINGS_PUBLIC_MARKERS:
                    violations.append(
                        "focused litchi-pages document-settings public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _iwork_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-pages document-settings public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages document-settings public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_section_settings_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the nested, archive-free Pages section-settings owner API."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    dedicated_sources = {
        root / path
        for path in PAGES_SECTION_SETTINGS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / PAGES_SECTION_SETTINGS_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_SECTION_SETTINGS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, PAGES_SECTION_SETTINGS_SHORT_NAMES
    )
    for name in PAGES_SECTION_SETTINGS_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages section-settings public API is missing "
            f"canonical section::settings type {name}: "
            f"{PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE}"
        )
    if PAGES_SECTION_SETTINGS_VALUE_TYPE in _rust_canonical_exports(
        semantic_source, frozenset({PAGES_SECTION_SETTINGS_VALUE_TYPE})
    ):
        violations.append(
            "focused litchi-pages section-settings public API duplicates canonical "
            "section::Settings inside section::settings: "
            f"{PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE}"
        )

    lib_path = root / PAGES_SECTION_SETTINGS_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_SECTION_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages section-settings public API is missing "
            f"canonical root section module: {PAGES_SECTION_SETTINGS_EXPORT_SOURCES[0]}"
        )

    section_path = root / PAGES_SECTION_SETTINGS_EXPORT_SOURCES[2]
    section_source = (
        _mask_rust_non_code(section_path.read_text(encoding="utf-8"))
        if section_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_SECTION_SETTINGS_MODULE.search(section_source) is None:
        violations.append(
            "focused litchi-pages section-settings public API is missing "
            "canonical section::settings module: "
            f"{PAGES_SECTION_SETTINGS_EXPORT_SOURCES[2]}"
        )
    if PAGES_SECTION_SETTINGS_VALUE_TYPE not in _rust_canonical_exports(
        section_source, frozenset({PAGES_SECTION_SETTINGS_VALUE_TYPE})
    ):
        violations.append(
            "focused litchi-pages section-settings public API is missing canonical "
            f"section::Settings: {PAGES_SECTION_SETTINGS_EXPORT_SOURCES[2]}"
        )

    package_path = root / PAGES_SECTION_SETTINGS_EXPORT_SOURCES[1]
    if package_path.is_file():
        package_source = _mask_rust_non_code(
            package_path.read_text(encoding="utf-8")
        )
        if PAGES_PACKAGE_SECTION_SETTINGS_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-pages section-settings public API is missing "
                f"private package owner module: {PAGES_SECTION_SETTINGS_EXPORT_SOURCES[1]}"
            )
        for match in PUBLIC_PAGES_PACKAGE_SECTION_SETTINGS_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages section-settings public API exposes duplicate "
                "package::section_settings module: "
                f"{PAGES_SECTION_SETTINGS_EXPORT_SOURCES[1]}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-pages section-settings public API is missing "
            f"private package owner module: {PAGES_SECTION_SETTINGS_EXPORT_SOURCES[1]}"
        )

    owner_path = root / PAGES_SECTION_SETTINGS_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-pages section-settings public API is missing private "
            f"package owner source: {PAGES_SECTION_SETTINGS_OWNER_SOURCE}"
        )
    owner_source = owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
    owner_methods = {
        name
        for declaration, _line_number in _rust_public_declarations(owner_source)
        for name, _nested_line in _rust_function_declarations(declaration)
    }
    for method in PAGES_SECTION_SETTINGS_PACKAGE_METHODS:
        if method in owner_methods:
            continue
        violations.append(
            "focused litchi-pages section-settings public API is missing Package "
            f"method {method}: {PAGES_SECTION_SETTINGS_OWNER_SOURCE}"
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
            if not _is_pages_section_settings_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _pages_section_settings_owner_declaration(declaration)
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "type"], ["pub", "use"])
            section_local_owner_alias = (
                public_declaration
                and path == section_path
                and public_use_or_type
                and "settings" in identifiers
                and bool(set(identifiers) & PAGES_SECTION_SETTINGS_PUBLIC_NAMES)
            )
            if (
                public_declaration
                and path in export_sources
                and (owner_declaration or section_local_owner_alias)
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages section-settings public API retains "
                    "root aliases via section::settings glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if (
                public_declaration
                and path in export_sources
                and (owner_declaration or section_local_owner_alias)
                and public_use_or_type
            ):
                violations.append(
                    "focused litchi-pages section-settings public API exposes public "
                    f"section-settings owner alias: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if public_declaration and identifier in PAGES_SECTION_SETTINGS_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages section-settings public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and (owner_declaration or section_local_owner_alias)
                    and public_use_or_type
                    and identifier in PAGES_SECTION_SETTINGS_PUBLIC_NAMES
                ):
                    violations.append(
                        "focused litchi-pages section-settings public API retains root "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_section_settings_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-pages section-settings public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-pages section-settings public API exposes raw "
                    f"byte slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_pages_section_background_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the nested, archive-free Pages section-background API."""

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in PAGES_SECTION_BACKGROUND_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / PAGES_SECTION_BACKGROUND_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in PAGES_SECTION_BACKGROUND_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8") if semantic_path.is_file() else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, PAGES_SECTION_BACKGROUND_SHORT_NAMES
    )
    for name in PAGES_SECTION_BACKGROUND_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-pages section-background public API is missing "
            f"canonical section::background type {name}: "
            f"{PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE}"
        )
    if PAGES_SECTION_BACKGROUND_VALUE_TYPE in _rust_canonical_exports(
        semantic_source, frozenset({PAGES_SECTION_BACKGROUND_VALUE_TYPE})
    ):
        violations.append(
            "focused litchi-pages section-background public API duplicates canonical "
            "section::Background inside section::background: "
            f"{PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE}"
        )

    lib_path = root / PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[0]
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_SECTION_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-pages section-background public API is missing "
            f"canonical root section module: {PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[0]}"
        )

    section_path = root / PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[2]
    section_source = (
        _mask_rust_non_code(section_path.read_text(encoding="utf-8"))
        if section_path.is_file()
        else ""
    )
    if PUBLIC_PAGES_SECTION_BACKGROUND_MODULE.search(section_source) is None:
        violations.append(
            "focused litchi-pages section-background public API is missing "
            "canonical section::background module: "
            f"{PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[2]}"
        )
    if PAGES_SECTION_BACKGROUND_VALUE_TYPE not in _rust_canonical_exports(
        section_source, frozenset({PAGES_SECTION_BACKGROUND_VALUE_TYPE})
    ):
        violations.append(
            "focused litchi-pages section-background public API is missing canonical "
            f"section::Background: {PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[2]}"
        )

    package_path = root / PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[1]
    if package_path.is_file():
        package_source = _mask_rust_non_code(
            package_path.read_text(encoding="utf-8")
        )
        if PAGES_PACKAGE_SECTION_BACKGROUND_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-pages section-background public API is missing "
                "private package owner module: "
                f"{PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[1]}"
            )
        for match in PUBLIC_PAGES_PACKAGE_SECTION_BACKGROUND_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-pages section-background public API exposes duplicate "
                "package::section_background module: "
                f"{PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[1]}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-pages section-background public API is missing "
            "private package owner module: "
            f"{PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[1]}"
        )

    owner_path = root / PAGES_SECTION_BACKGROUND_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-pages section-background public API is missing private "
            f"package owner source: {PAGES_SECTION_BACKGROUND_OWNER_SOURCE}"
        )
    owner_source = owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
    owner_methods = {
        name
        for declaration, _line_number in _rust_public_declarations(owner_source)
        for name, _nested_line in _rust_function_declarations(declaration)
    }
    for method in PAGES_SECTION_BACKGROUND_PACKAGE_METHODS:
        if method in owner_methods:
            continue
        violations.append(
            "focused litchi-pages section-background public API is missing Package "
            f"method {method}: {PAGES_SECTION_BACKGROUND_OWNER_SOURCE}"
        )

    edit_methods = {
        name
        for implementation_path in dedicated_sources
        for declaration, _line_number in _rust_public_declarations(
            implementation_path.read_text(encoding="utf-8")
        )
        for name, _nested_line in _rust_function_declarations(declaration)
    }
    for method in PAGES_SECTION_BACKGROUND_EDIT_METHODS:
        if method in edit_methods:
            continue
        violations.append(
            "focused litchi-pages section-background public API is missing "
            f"canonical Edit::{method} method: {PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE}"
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
            if not _is_pages_section_background_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _pages_section_background_owner_declaration(declaration)
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "type"], ["pub", "use"])
            section_local_owner_alias = (
                public_declaration
                and path == section_path
                and public_use_or_type
                and "background" in identifiers
                and bool(set(identifiers) & PAGES_SECTION_BACKGROUND_PUBLIC_NAMES)
            )
            if (
                public_declaration
                and path in export_sources
                and (owner_declaration or section_local_owner_alias)
                and identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-pages section-background public API retains "
                    "root aliases via section::background glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if (
                public_declaration
                and path in export_sources
                and (owner_declaration or section_local_owner_alias)
                and public_use_or_type
            ):
                violations.append(
                    "focused litchi-pages section-background public API exposes public "
                    f"section-background owner alias: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count("\n", 0, match.start(1))
                if public_declaration and identifier in PAGES_SECTION_BACKGROUND_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-pages section-background public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and (owner_declaration or section_local_owner_alias)
                    and public_use_or_type
                    and identifier in PAGES_SECTION_BACKGROUND_PUBLIC_NAMES
                ):
                    violations.append(
                        "focused litchi-pages section-background public API retains root "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _pages_section_background_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-pages section-background public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count("\n", 0, match.start())
                violations.append(
                    "focused litchi-pages section-background public API exposes raw "
                    f"byte slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))
NEUTRAL_TABLE_PAGES_SEMANTIC_TABLE_READ_IMPORT = re.compile(
    r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*table"
    r"[ \t\r\n]*::[ \t\r\n]*read\b[^;]*\b(?:r#)?TableRead\b"
)
NEUTRAL_TABLE_PAGES_SEMANTIC_TABLE_FIELD = re.compile(
    r"(?m)^[ \t]*(?:(?:pub(?:[ \t]*\([^()\r\n]*\))?)[ \t]+)?"
    r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*:[ \t\r\n]*"
    r"(?:litchi_iwa_common[ \t\r\n]*::[ \t\r\n]*table"
    r"[ \t\r\n]*::[ \t\r\n]*read[ \t\r\n]*::[ \t\r\n]*)?"
    r"(?:r#)?TableRead\b"
)
PAGES_TABLE_MERGE_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?body_table_merges\b[ \t]*(?:;|\{)"
)
PAGES_TABLE_MERGE_PUBLIC_METHOD = "body_table_merges"
PAGES_TABLE_MERGE_SELECTOR_TYPES = frozenset({"BodyTableSelector"})
# Pages and Keynote expose the same focused merge-reader contract as Numbers.
# Keep each child path explicit so the source-ingress implementation remains
# private to its package owner while the format crate re-exports only the
# checked semantic reader.
PAGES_TABLE_MERGE_READER_SOURCE = Path(
    "crates/litchi-pages/src/package/body_table_merges/reader.rs"
)

# Pages table deletion is a package-owned structural transaction.  Keep the
# physical graph/archive/metadata/formula/text proof in five private children so a
# semantic caller can only select by ``BodyTableSelector`` and receive the
# archive-free snapshot/transaction vocabulary.  The audit is dormant until
# the owner source appears: package re-exports may land before the implementation
# while the migration is staged in several commits.
IWA_PAGES_BODY_TABLE_DELETION_OWNER_SOURCE = Path(
    "crates/litchi-pages/src/package/body_table_deletion.rs"
)
IWA_PAGES_BODY_TABLE_DELETION_CHILD_ROOT = Path(
    "crates/litchi-pages/src/package/body_table_deletion"
)
IWA_PAGES_BODY_TABLE_DELETION_CHILD_MODULES = (
    "graph",
    "archive",
    "metadata",
    "formula",
    "text",
)
IWA_PAGES_BODY_TABLE_DELETION_CHILD_SOURCES = tuple(
    IWA_PAGES_BODY_TABLE_DELETION_CHILD_ROOT / f"{name}.rs"
    for name in IWA_PAGES_BODY_TABLE_DELETION_CHILD_MODULES
)
IWA_PAGES_BODY_TABLE_DELETION_PACKAGE_SOURCE = Path(
    "crates/litchi-pages/src/package.rs"
)
IWA_PAGES_BODY_TABLE_DELETION_LIBRARY_SOURCE = Path(
    "crates/litchi-pages/src/lib.rs"
)
# ADR 0034: the owner defines these canonical short names and only
# ``litchi_pages::table::deletion`` re-exports them; the removed table is the
# ``table::catalog::Snapshot`` value.
IWA_PAGES_BODY_TABLE_DELETION_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
# ``Package::edit_body_table_deletion`` starts the scoped
# edit, ``Edit::remove`` stages the selector and ``Edit::commit`` resolves it
# before any staging or publication helper runs.
IWA_PAGES_BODY_TABLE_DELETION_PACKAGE_METHODS = (
    "edit_body_table_deletion",
    "apply_body_table_deletion",
)
IWA_PAGES_BODY_TABLE_DELETION_CHILD_MARKER_GROUPS = {
    "graph selector and ownership proof": (
        "BodyTableTarget",
        "validate_body_table_target",
        "discover_table_graph",
        "selected_position",
        "selected_roots",
        "OwnershipRole",
        "prove_inbound_references",
        "ownership",
    ),
    "graph references and bounded accounting": (
        "references",
        "inbound",
        "shared",
        "private",
        "checked_add",
        "checked_mul",
        "Budget",
        "Limits",
    ),
    "archive source-preserving removal": (
        "Archive",
        "archive",
        "SourceCatalog",
        "source_preserving",
        "source-preserving",
        "remove_object",
        "remove_message",
        "candidate",
    ),
    "archive verification": (
        "verify",
        "verification",
        "reopen",
        "absence",
        "checked",
        "exact",
        "validate_plan",
        "validate_canonical_object_framing",
        "has_live_records",
    ),
    "metadata component registration": (
        "package_metadata_codec",
        "component_removal",
        "ComponentRemoval",
        "component registration",
        "registration",
    ),
    "metadata identity/reference cleanup": (
        "uuid",
        "UUID",
        "external_reference",
        "external-reference",
        "remove_component_registration",
        "object_uuid",
    ),
    "formula dependency pruning": (
        "numbers_table_cell_dependency_codec",
        "dependency_wire",
        "formula_engine",
        "CalculationEngine",
        "formula_owner",
        "owner",
    ),
    "formula tile pruning and counters": (
        "tile",
        "tiled",
        "prune",
        "remove_formula",
        "number_of_formulas",
        "checked_sub",
    ),
    "text-wire delegation": (
        "litchi_iwa_text_wire",
        "text_wire",
        "rewrite_storage_text",
        "prepare_storage_text_rewrite",
        "StorageText",
    ),
    "UTF-16 selected-anchor proof": (
        "utf16_index",
        "utf16_len",
        "UTF_16",
        "UTF-16",
        "anchor",
        "character_index",
        "selected_anchor",
        "TextSpan",
        "Position",
    ),
    "bounded text accounting": (
        "try_reserve",
        "checked_add",
        "checked_sub",
        "WireBudget",
        "budget",
        "limit",
        "max",
    ),
}
IWA_PAGES_BODY_TABLE_DELETION_OWNER_MARKER_GROUPS = {
    "selector admission": (
        "BodyTableSelector",
        "remove_body_table",
        "resolve_body_table",
        "select_body_table",
        "body_table_catalog",
    ),
    "semantic deletion patch": (
        "Patch",
        "Snapshot",
        "removed_table",
        "staged",
        "deletion",
    ),
    "bounded graph/rewrite accounting": (
        "Budget",
        "Limits",
        "limit",
        "checked_add",
        "checked_mul",
        "try_reserve",
    ),
    "candidate reassembly and readback": (
        "reassemble",
        "prepare_reassembly",
        "candidate",
        "reopen",
        "from_bytes",
        "verify",
        "readback",
    ),
    "exact-source conflict and inverse": (
        "PatchConflict",
        "source_fingerprint",
        "source_fingerprint",
        "exact_source",
        "inverse",
    ),
}
IWA_PAGES_BODY_TABLE_DELETION_LEGACY_FALLBACK = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"table_cell_merges_in_package|regions_in_package|"
    r"legacy_table_deletion|fallback_table_deletion|"
    r"delete_table_in_host|remove_table_in_host"
    r")(?![A-Za-z0-9_])[ \t\r\n]*\("
)
IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_RAW_ID = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|table_id|table_identifier|"
    r"model_id|model_identifier|native_id|object_id|drawable_id|"
    r"component_id|archive_id|message_id|storage_id|uuid_id|"
    r"source_bytes|package_bytes|raw_bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*"
    r"(?:Option[ \t\r\n]*<[ \t\r\n]*)?"
    r"(?:u8|u16|u32|u64|u128|usize|i8|i16|i32|i64|i128|isize)\b"
)
IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_BYTES = re.compile(
    r"&[ \t\r\n]*\[[ \t\r\n]*u8\s*\]|"
    r"\b(?:Vec|Box|Arc)[ \t\r\n]*<[ \t\r\n]*(?:u8|\[[ \t\r\n]*u8\s*\])"
    r"|\b(?:Bytes|ByteString)\b"
)
IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_PHYSICAL = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"Archive|ArchiveObject|ComponentCatalog|EntryEdit|ExactArtifacts|"
    r"OwnedExactArtifacts|RawMessage|SnappyStream|SourceCatalog|SharedBytes|"
    r"PhysicalSource|Document|Body|Show|Slide|Sheet|TableModelArchive|"
    r"TableInfoArchive|FormulaArchive|AstNodeArchive|AstNodeType|"
    r"WireView|WireLimits|DecodeOptions|RawMessage|"
    r"litchi_iwa_(?:archive|core|protos)|prost|prost_types|buffa|"
    r"(?:[A-Za-z_][A-Za-z0-9_]*Generated|[A-Za-z_][A-Za-z0-9_]*ArchiveView)"
    r")(?![A-Za-z0-9_])"
)

PAGES_TABLE_CELLS_SOURCE = Path("crates/litchi-pages/src/package/table_cells.rs")
PAGES_TABLE_CELLS_PACKAGE_SOURCE = Path("crates/litchi-pages/src/package.rs")
PAGES_TABLE_CELLS_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?table_cells\b[ \t]*(?:;|\{)"
)


def audit_pages_table_merge_source_topology(root: Path = ROOT) -> list[str]:
    """Keep Pages merged-cell reads selector-first and archive-free."""

    violations = _audit_focused_table_merge_source_topology(
        root,
        format_name="Pages",
        owner_path=PAGES_TABLE_MERGE_SOURCE,
        package_path=PAGES_TABLE_MERGE_PACKAGE_SOURCE,
        module_pattern=PAGES_TABLE_MERGE_MODULE,
        table_module_path=Path("crates/litchi-pages/src/table/mod.rs"),
        table_merge_path=Path("crates/litchi-pages/src/table/merge.rs"),
        table_module_pattern=PAGES_TABLE_MERGE_COMMON_TABLE_MODULE,
        public_method=PAGES_TABLE_MERGE_PUBLIC_METHOD,
        selector_types=PAGES_TABLE_MERGE_SELECTOR_TYPES,
    )
    violations.extend(audit_pages_table_merge_reader_source_topology(root))
    return sorted(set(violations))


def audit_pages_table_merge_reader_source_topology(root: Path = ROOT) -> list[str]:
    """Keep Pages' focused merge-reader child private and ergonomic."""

    return _audit_focused_table_merge_reader_source_topology(
        root,
        format_name="Pages",
        owner_path=PAGES_TABLE_MERGE_SOURCE,
        reader_path=PAGES_TABLE_MERGE_READER_SOURCE,
        query_method="body_table_merges",
        selector_types=PAGES_TABLE_MERGE_SELECTOR_TYPES,
        limits_type="Limits",
        options_type="OpenOptions",
    )


def audit_pages_table_merge_transaction_source_topology(
    root: Path = ROOT,
) -> list[str]:
    return _audit_table_merge_transaction_source_topology(root, IWA_TABLE_MERGE_TRANSACTION_SPECS[1])


def _audit_pages_body_table_deletion_public_surface(
    root: Path,
    sources: tuple[tuple[Path, str], ...],
    package_source: str,
    library_source: str,
) -> list[str]:
    """Reject native values from the Pages deletion transaction's public API."""

    violations: list[str] = []
    canonical = IWA_PAGES_BODY_TABLE_DELETION_CANONICAL_TYPES

    for path, source in sources:
        production = _mask_rust_cfg_test_items(source)
        for name in sorted(canonical):
            struct = _rust_named_struct_body(production, name)
            if struct is None:
                continue
            body, offset = struct
            for match in IWA_TABLE_MERGE_TRANSACTION_PUBLIC_FIELD.finditer(body):
                line = production.count("\n", 0, offset + match.start()) + 1
                violations.append(
                    "Pages body-table deletion transaction exposes a public field: "
                    f"{path}:{line}"
                )

    for path, source in sources:
        production = _mask_rust_cfg_test_items(source)
        for type_name in (*canonical, "Package"):
            for method_name, declaration, line in _rust_public_methods_in_impl(
                production, type_name
            ):
                raw_id = IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_RAW_ID.search(
                    declaration
                )
                if raw_id is not None:
                    violations.append(
                        "Pages body-table deletion public method exposes a raw ID "
                        f"parameter {raw_id.group(0).strip()}: {path}:{line}"
                    )
                if IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_BYTES.search(declaration):
                    violations.append(
                        "Pages body-table deletion public method exposes raw bytes: "
                        f"{path}:{line}"
                    )
                physical = IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_PHYSICAL.search(
                    declaration
                )
                if physical is not None:
                    violations.append(
                        "Pages body-table deletion public method exposes a native, "
                        f"generated, or wire type {physical.group()}: {path}:{line}"
                    )

    # Public re-exports are checked independently from the owner methods: a
    # future facade alias must not reintroduce a physical type through a
    # declaration that never appears in the transaction child itself.  After
    # ADR 0034 the family is published only by ``table::deletion``; the
    # short names are too common to filter ``package.rs``/``lib.rs``
    # declarations, so the scan covers that public module's scope.
    public_file, public_scope, _ = _pages_public_module_scope(
        root, PAGES_TRANSACTION_FAMILY_BY_OWNER["body_table_deletion"][2]
    )
    facade_sources = (
        []
        if public_file is None or public_scope is None
        else [(public_file.relative_to(root), public_scope)]
    )
    for path, source in facade_sources:
        for declaration, line in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not identifiers & canonical:
                continue
            raw_id = IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_RAW_ID.search(declaration)
            if raw_id is not None:
                violations.append(
                    "Pages body-table deletion facade exposes a raw ID parameter "
                    f"{raw_id.group(0).strip()}: {path}:{line}"
                )
            if IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_BYTES.search(declaration):
                violations.append(
                    "Pages body-table deletion facade exposes raw bytes: "
                    f"{path}:{line}"
                )
            physical = IWA_PAGES_BODY_TABLE_DELETION_PUBLIC_PHYSICAL.search(
                declaration
            )
            if physical is not None:
                violations.append(
                    "Pages body-table deletion facade exposes a native, generated, "
                    f"or wire type {physical.group()}: {path}:{line}"
                )

    return violations


def audit_pages_body_table_deletion_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep Pages table deletion selector-first and physically modular."""

    owner_path = IWA_PAGES_BODY_TABLE_DELETION_OWNER_SOURCE
    owner_absolute = root / owner_path
    # The package declaration may be staged before the owner implementation.
    # Waiting for the owner file keeps the checker useful during that split
    # while still failing closed as soon as the transaction is present.
    if not owner_absolute.is_file():
        return []

    violations: list[str] = []
    owner_source = _mask_rust_cfg_test_items(owner_absolute.read_text(encoding="utf-8"))
    owner_code = _mask_rust_non_code(owner_source)
    sources: list[tuple[Path, str]] = [(owner_path, owner_source)]

    modules = _rust_root_level_module_declarations(
        owner_source,
        frozenset(IWA_PAGES_BODY_TABLE_DELETION_CHILD_MODULES),
    )
    for child_name in IWA_PAGES_BODY_TABLE_DELETION_CHILD_MODULES:
        matches = [item for item in modules if item[0] == child_name]
        if not matches:
            violations.append(
                "Pages body-table deletion owner is missing private child module "
                f"{child_name}: {owner_path}"
            )
            continue
        for _name, visibility, shape, _body, line in matches:
            if visibility is not None and visibility.strip() == "pub":
                violations.append(
                    "Pages body-table deletion child module must remain private: "
                    f"{owner_path}:{line}"
                )
            if shape == "inline":
                violations.append(
                    "Pages body-table deletion child must use a dedicated source: "
                    f"{owner_path}:{line}"
                )

    for child_path in IWA_PAGES_BODY_TABLE_DELETION_CHILD_SOURCES:
        absolute = root / child_path
        if not absolute.is_file():
            violations.append(
                "Pages body-table deletion owner is missing child source: "
                f"{child_path}"
            )
            continue
        child_source = _mask_rust_cfg_test_items(absolute.read_text(encoding="utf-8"))
        sources.append((child_path, child_source))

    package_absolute = root / IWA_PAGES_BODY_TABLE_DELETION_PACKAGE_SOURCE
    library_absolute = root / IWA_PAGES_BODY_TABLE_DELETION_LIBRARY_SOURCE
    package_source = (
        _mask_rust_cfg_test_items(package_absolute.read_text(encoding="utf-8"))
        if package_absolute.is_file()
        else ""
    )
    library_source = (
        _mask_rust_cfg_test_items(library_absolute.read_text(encoding="utf-8"))
        if library_absolute.is_file()
        else ""
    )
    if not package_source:
        violations.append(
            "Pages body-table deletion package facade is missing: "
            f"{IWA_PAGES_BODY_TABLE_DELETION_PACKAGE_SOURCE}"
        )
    if not library_source:
        violations.append(
            "Pages body-table deletion crate facade is missing: "
            f"{IWA_PAGES_BODY_TABLE_DELETION_LIBRARY_SOURCE}"
        )

    package_modules = _rust_root_level_module_declarations(
        package_source, frozenset({"body_table_deletion"})
    )
    if not package_modules:
        violations.append(
            "Pages body-table deletion package module is missing: "
            f"{IWA_PAGES_BODY_TABLE_DELETION_PACKAGE_SOURCE}"
        )
    elif any(
        visibility is not None and visibility.strip() == "pub"
        for _name, visibility, _shape, _body, _line in package_modules
    ):
        violations.append(
            "Pages body-table deletion package module must remain private: "
            f"{IWA_PAGES_BODY_TABLE_DELETION_PACKAGE_SOURCE}"
        )

    canonical = IWA_PAGES_BODY_TABLE_DELETION_CANONICAL_TYPES
    for name in sorted(canonical):
        if name not in _rust_canonical_exports(owner_source, canonical):
            violations.append(
                "Pages body-table deletion owner is missing canonical type "
                f"{name}: {owner_path}"
            )
    # The package and crate-root facades no longer re-export the family; the
    # shared helper checks ``table::deletion`` and the deleted spellings.
    violations.extend(
        _pages_transaction_family_violations(
            root, "body_table_deletion", "Pages body-table deletion"
        )
    )

    all_source = "\n".join(source for _path, source in sources)
    all_code = _mask_rust_non_code(all_source)
    for child_name, child_path in zip(
        IWA_PAGES_BODY_TABLE_DELETION_CHILD_MODULES,
        IWA_PAGES_BODY_TABLE_DELETION_CHILD_SOURCES,
    ):
        child_absolute = root / child_path
        if not child_absolute.is_file():
            continue
        child = _mask_rust_non_code(
            _mask_rust_cfg_test_items(child_absolute.read_text(encoding="utf-8"))
        )
        # The graph/archive split has a few established naming families.  A
        # child must still prove every concern assigned to it, so moving a
        # helper between private files cannot silently remove an entire proof.
        if child_name == "graph":
            groups = (
                ("graph selector and ownership proof",),
                ("graph references and bounded accounting",),
            )
        elif child_name == "archive":
            groups = (
                ("archive source-preserving removal",),
                ("archive verification",),
            )
        elif child_name == "metadata":
            groups = (
                ("metadata component registration",),
                ("metadata identity/reference cleanup",),
            )
        elif child_name == "formula":
            groups = (
                ("formula dependency pruning",),
                ("formula tile pruning and counters",),
            )
        else:
            groups = (
                ("text-wire delegation",),
                ("UTF-16 selected-anchor proof",),
                ("bounded text accounting",),
            )
        for group_name_tuple in groups:
            group_name = group_name_tuple[0]
            if not _boundary_marker_present(
                child, IWA_PAGES_BODY_TABLE_DELETION_CHILD_MARKER_GROUPS[group_name]
            ):
                violations.append(
                    "Pages body-table deletion child is missing "
                    f"{group_name} marker: {child_path}"
                )

    for label, markers in IWA_PAGES_BODY_TABLE_DELETION_OWNER_MARKER_GROUPS.items():
        if not _boundary_marker_present(all_code, markers):
            violations.append(
                "Pages body-table deletion transaction is missing "
                f"{label} marker: {owner_path}"
            )

    # Keep the public operation's admission fence ahead of any staging or
    # publication helper.  The patch-apply route admits its exact source patch
    # before the same physical rewrite machinery is reached.
    package_method_records: dict[str, list[tuple[str, str, int, Path]]] = {
        name: [] for name in IWA_PAGES_BODY_TABLE_DELETION_PACKAGE_METHODS
    }
    commit_records: list[tuple[str, str, int, Path]] = []
    stage_records: list[tuple[str, str, int, Path]] = []
    for path, source in sources:
        for method_name in package_method_records:
            for declaration, body, line in _rust_public_method_bodies(
                source, method_name, type_name="Package"
            ):
                package_method_records[method_name].append(
                    (declaration, body, line, path)
                )
        for declaration, body, line in _rust_public_method_bodies(
            source, "commit", type_name="Edit"
        ):
            commit_records.append((declaration, body, line, path))
        for declaration, body, line in _rust_public_method_bodies(
            source, "remove", type_name="Edit"
        ):
            stage_records.append((declaration, body, line, path))

    if not package_method_records["edit_body_table_deletion"]:
        violations.append(
            "Pages body-table deletion Package method is missing "
            f"edit_body_table_deletion: {owner_path}"
        )
    if not stage_records:
        violations.append(
            f"Pages body-table deletion Edit method is missing remove: {owner_path}"
        )
    for declaration, _body, line, path in stage_records:
        if re.search(r"\bBodyTableSelector\b", declaration) is None:
            violations.append(
                "Pages body-table deletion Edit::remove must accept BodyTableSelector: "
                f"{path}:{line}"
            )
    if not commit_records:
        violations.append(
            f"Pages body-table deletion Edit method is missing commit: {owner_path}"
        )
    for declaration, body, line, path in commit_records:
        if re.search(r"\bResult\s*<\s*Commit\s*,\s*Error\s*>", declaration) is None:
            violations.append(
                "Pages body-table deletion Edit::commit must return the typed deletion "
                f"result: {path}:{line}"
            )
        body_code = _mask_rust_non_code(body)
        body_opening = body_code.find("{")
        if body_opening >= 0:
            body_code = body_code[body_opening + 1 :]
        admission = re.search(
            r"\b(?:select_target|resolve_body_table|select_body_table|"
            r"body_table_catalog)\b",
            body_code,
        )
        staging = re.search(
            r"\b(?:Patch|stage|staged|reassemble|rewrite|"
            r"candidate)\b",
            body_code,
        )
        if admission is None:
            violations.append(
                "Pages body-table deletion Edit::commit must resolve its typed selector "
                f"before staging: {path}:{line}"
            )
        elif staging is not None and staging.start() < admission.start():
            violations.append(
                "Pages body-table deletion Edit::commit stages deletion before selector "
                f"admission: {path}:{line}"
            )
        if IWA_PAGES_BODY_TABLE_DELETION_LEGACY_FALLBACK.search(body_code):
            violations.append(
                "Pages body-table deletion Edit::commit retains a legacy/catchall "
                f"deletion fallback: {path}:{line}"
            )
        if re.search(r"\bErr\s*\(\s*_\s*\)\s*=>", body_code):
            violations.append(
                "Pages body-table deletion Edit::commit uses a post-admission catchall "
                f"error arm: {path}:{line}"
            )

    apply_records = package_method_records["apply_body_table_deletion"]
    if not apply_records:
        violations.append(
            "Pages body-table deletion Package method is missing "
            "apply_body_table_deletion: "
            f"{owner_path}"
        )
    for declaration, body, line, path in apply_records:
        if re.search(r"\bPatch\b", declaration) is None:
            violations.append(
                "Pages apply_body_table_deletion must accept table::deletion::Patch: "
                f"{path}:{line}"
            )
        if re.search(r"\bCommit\b", declaration) is None:
            violations.append(
                "Pages apply_body_table_deletion must return table::deletion::Commit: "
                f"{path}:{line}"
            )
        body_code = _mask_rust_non_code(body)
        body_opening = body_code.find("{")
        if body_opening >= 0:
            body_code = body_code[body_opening + 1 :]
        admission = re.search(
            r"\b(?:PatchConflict|source_fingerprint|target_fingerprint|"
            r"exact_source|authoriz|validate_patch|Patch)\b",
            body_code,
        )
        staging = re.search(
            r"\b(?:stage|staged|reassemble|rewrite|candidate)\b",
            body_code,
        )
        if admission is None:
            violations.append(
                "Pages apply_body_table_deletion must validate its exact patch before "
                f"staging: {path}:{line}"
            )
        elif staging is not None and staging.start() < admission.start():
            violations.append(
                "Pages apply_body_table_deletion stages deletion before patch admission: "
                f"{path}:{line}"
            )
        if IWA_PAGES_BODY_TABLE_DELETION_LEGACY_FALLBACK.search(body_code):
            violations.append(
                "Pages apply_body_table_deletion retains a legacy/catchall deletion "
                f"fallback: {path}:{line}"
            )
        if re.search(r"\bErr\s*\(\s*_\s*\)\s*=>", body_code):
            violations.append(
                "Pages apply_body_table_deletion uses a post-admission catchall error "
                f"arm: {path}:{line}"
            )

    commit_methods = {
        method_name
        for path, source in sources
        for method_name, _declaration, _line in _rust_public_methods_in_impl(
            _mask_rust_cfg_test_items(source), "Commit"
        )
    }
    for required in ("package", "into_package", "patch", "diagnostics", "removed_table"):
        if required not in commit_methods:
            violations.append(
                "Pages table::deletion::Commit is missing public "
                f"{required}: {owner_path}"
            )

    removed_table_records = [
        (path, declaration, line)
        for path, source in sources
        for method_name, declaration, line in _rust_public_methods_in_impl(
            _mask_rust_cfg_test_items(source), "Commit"
        )
        if method_name == "removed_table"
    ]
    if removed_table_records and not any(
        re.search(r"\bSnapshot\b", declaration)
        for _path, declaration, _line in removed_table_records
    ):
        path, _declaration, line = removed_table_records[0]
        violations.append(
            "Pages table::deletion::Commit::removed_table must return semantic "
            f"table::catalog::Snapshot: {path}:{line}"
        )

    violations.extend(
        _audit_pages_body_table_deletion_public_surface(
            root, tuple(sources), package_source, library_source
        )
    )
    return sorted(set(violations))


def audit_pages_table_cells_source_topology(root: Path = ROOT) -> list[str]:
    """Keep Pages body-table cell reads selector-first and semantic."""

    return _audit_focused_table_cells_source_topology(
        root,
        format_name="Pages",
        owner_path=PAGES_TABLE_CELLS_SOURCE,
        package_path=PAGES_TABLE_CELLS_PACKAGE_SOURCE,
        module_pattern=PAGES_TABLE_CELLS_MODULE,
        public_method="body_table_cells",
    )


def audit_pages_transaction_family_source_topology(root: Path = ROOT) -> list[str]:
    """Keep the ADR 0034 Pages transaction modules canonical.

    Every family in ``PAGES_TRANSACTION_FAMILIES`` is checked by
    ``_pages_transaction_family_violations``.  Crate-wide, this audit also
    rejects:

    * a nominal definition (any visibility) or a ``pub use`` of a deleted
      prefixed spelling or retired root alias (``BodyTableLockState``,
      ``BodyTableHeaderSettings``, ``BodyTableName``, the
      ``DocumentSourceLimits``/``ReadError`` family, ...);
    * a second public path to a family's canonical names: any other
      ``pub use`` that names the owner module and imports a canonical short
      name, an alias or a glob.  Private and ``pub(crate)`` imports (for
      example ``use ...::{Error as BodyChartArrangementError}`` inside another
      owner) are not public surface and are allowed;
    * raw byte-slice or archive/wire/protobuf identifiers in public
      declarations of the owner sources.
    """

    source_root = root / PAGES_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    violations: list[str] = []
    for owner, *_ in PAGES_TRANSACTION_FAMILIES:
        violations.extend(_pages_transaction_family_violations(root, owner))

    violations.extend(
        _pages_retired_name_violations(
            root,
            PAGES_TRANSACTION_RETIRED_EXTRA_NAMES,
            "Pages retired alias/ingress family",
        )
    )

    owners = {family[0]: frozenset(family[3]) for family in PAGES_TRANSACTION_FAMILIES}
    reexports: dict[str, int] = {owner: 0 for owner in owners}
    for path in sorted(source_root.rglob("*.rs")):
        relative = path.relative_to(root)
        code = _pages_masked_production_code(path)
        for statement in PAGES_PUBLIC_USE_STATEMENT.finditer(code):
            text = statement.group(0)
            line_number = code.count("\n", 0, statement.start()) + 1
            identifiers = [match.group(1) for match in RUST_IDENTIFIER.finditer(text)]
            for owner, shorts in owners.items():
                if owner not in identifiers:
                    continue
                tail = text[text.find(owner) + len(owner):]
                tail_identifiers = set(
                    match.group(1) for match in RUST_IDENTIFIER.finditer(tail)
                )
                if "*" in tail:
                    violations.append(
                        f"Pages transaction owner {owner} is glob re-exported: "
                        f"{relative}:{line_number}"
                    )
                if tail_identifiers & shorts:
                    reexports[owner] += 1
                for alias in re.finditer(
                    r"(?:r#)?([A-Za-z_][A-Za-z0-9_]*)[ \t\r\n]+as[ \t\r\n]+"
                    r"(?:r#)?([A-Za-z_][A-Za-z0-9_]*)",
                    tail,
                ):
                    if alias.group(1) != alias.group(2):
                        violations.append(
                            f"Pages transaction owner {owner} has alternate public alias "
                            f"{alias.group(2)} for {alias.group(1)}: {relative}:{line_number}"
                        )
    for owner, count in sorted(reexports.items()):
        if count > 1:
            violations.append(
                f"Pages transaction owner {owner} has {count} public re-export paths; "
                f"only litchi_pages::{'::'.join(PAGES_TRANSACTION_FAMILY_BY_OWNER[owner][2])} "
                f"may publish it: {PAGES_SOURCE_ROOT}"
            )

    for owner in owners:
        owner_relative = PAGES_SOURCE_ROOT / "package" / f"{owner}.rs"
        owner_path = root / owner_relative
        if not owner_path.is_file():
            continue
        source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        lines = _mask_rust_non_code(source).split("\n")
        for declaration, line_number in _rust_public_declarations(source):
            if re.search(r"(?<![A-Za-z0-9_])pub[ \t\r\n]*\(", declaration[:16]):
                continue
            # ``#[doc(hidden)]`` items are the unstable internal-source seams
            # (``__decode_*``); they are audited by their own focused gates.
            attribute_index = line_number - 2
            hidden = False
            while attribute_index >= 0 and lines[attribute_index].strip().startswith("#"):
                if re.search(r"\bdoc\s*\(\s*hidden\s*\)", lines[attribute_index]):
                    hidden = True
                attribute_index -= 1
            if hidden:
                continue
            for match in RUST_IDENTIFIER.finditer(declaration):
                reason = _iwork_public_leak(match.group(1))
                # Counts such as ``removed_objects`` and error variants such
                # as ``ObjectMarkerReplacement`` are semantic vocabulary, not
                # native handles; the focused per-family gates classify them.
                if reason is None or reason == "native object":
                    continue
                identifier_line = line_number + declaration.count("\n", 0, match.start(1))
                violations.append(
                    f"Pages {owner} transaction public API exposes {reason} "
                    f"{match.group(1)}: {owner_relative}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                identifier_line = line_number + declaration.count("\n", 0, match.start())
                violations.append(
                    f"Pages {owner} transaction public API exposes raw byte slice: "
                    f"{owner_relative}:{identifier_line}"
                )
    return sorted(set(violations))
