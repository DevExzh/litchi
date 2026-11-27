"""Dependency and public API checks for Keynote."""

from __future__ import annotations

from .common import (
    CAMEL_CASE_WORD,
    CARGO_PROST_DEPENDENCY,
    CARGO_SECTION_HEADER,
    IWA_TABLE_MERGE_TRANSACTION_SPECS,
    KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES,
    KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE,
    KEYNOTE_CHART_ARRANGEMENT_PACKAGE_MODULE,
    KEYNOTE_SOURCE_ROOT,
    KEYNOTE_TABLE_MERGE_PACKAGE_SOURCE,
    KEYNOTE_TABLE_MERGE_SOURCE,
    PAGES_TABLE_MERGE_COMMON_TABLE_MODULE,
    PAGES_TABLE_MERGE_FOCUSED_CALL,
    PAGES_TABLE_MERGE_FORBIDDEN_MODULES,
    Path,
    ROOT,
    RUST_BYTE_SLICE,
    RUST_IDENTIFIER,
    _audit_focused_table_cells_source_topology,
    _audit_focused_table_merge_reader_source_topology,
    _audit_focused_table_merge_source_topology,
    _audit_table_merge_transaction_source_topology,
    _iwork_public_leak,
    _keynote_chart_arrangement_owner_present,
    _mask_rust_cfg_test_items,
    _mask_rust_non_code,
    _proto_message_body,
    _rust_any_function_body,
    _rust_canonical_exports,
    _rust_function_declarations,
    _rust_impl_headers,
    _rust_named_function_body,
    _rust_named_struct_body,
    _rust_public_declarations,
    _rust_public_enum_variants,
    _rust_public_methods_in_impl,
    field,
    re,
)


RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_MODULES = ("show_settings",)
KEYNOTE_PACKAGE_MANIFEST = Path("crates/litchi-keynote/Cargo.toml")
KEYNOTE_DOCUMENT_PUBLIC_API_SOURCES = (KEYNOTE_SOURCE_ROOT / "document.rs",)
# These are semantic resource ceilings, not native-object vocabulary. Keep the
# allowlist exact so similarly named physical identifiers remain forbidden.
KEYNOTE_DOCUMENT_SEMANTIC_LIMIT_PARAMETERS = frozenset(
    {
        "max_objects",
        "max_slides",
        "max_references",
        "max_text_storages",
        "max_text_fragments",
        "max_text_bytes",
    }
)
# These files are included only from `#[cfg(test)]` modules in their parent
# source files. Keep the production audits from treating those test-only
# module bodies as reachable crate items when walking the source tree.
KEYNOTE_TEST_ONLY_SOURCE_NAMES = frozenset(
    {"perf_tests.rs", "verification_tests.rs"}
)
KEYNOTE_GENERATED_PROTO_MODULES = (
    "kn",
    "knsos",
    "tn",
    "tnsos",
    "tp",
    "tpsos",
    "tsa",
    "tsasos",
    "tsd",
    "tsdsos",
    "tsk",
    "tsp",
    "tss",
    "tsssos",
    "tswp",
    "tswpsos",
)
KEYNOTE_NO_EAGER_PROST_SOURCE_PATTERNS = (
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
            + "|".join(KEYNOTE_GENERATED_PROTO_MODULES)
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
# The focused package owns selector-first chart-title transactions.  Keep its
# production source from calling the compatibility host's raw-ID mutation
# methods while allowing the selector-suffixed near names used by adapters.
KEYNOTE_CHART_TITLE_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?P<method>set_slide_chart_title|"
    r"remove_slide_chart_title)(?![A-Za-z0-9_])[ \t\r\n]*\("
)

# Only the Keynote chart-legend visibility bit lives in the focused package
# owner. Keep this seam independent from the remaining legacy legend
# style operations (fill, frame, font, stroke, and shadow), and from the
# shared Pages/Numbers chart-options helper.
KEYNOTE_CHART_LEGEND_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_chart_legend.rs"
)
# ADR 0034: focused Keynote families use the unprefixed
# transaction vocabulary inside their own modules.  These generic names no
# longer identify one family, so the facade-leak filters of the crate-root and
# package seams ignore them and look for family-specific names instead.
KEYNOTE_FAMILY_GENERIC_NAMES = frozenset(
    {"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path", "InvalidReason", "Data"}
)
KEYNOTE_CHART_LEGEND_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_CHART_LEGEND_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_LEGEND_FAMILY_MODULE = 'legend'
KEYNOTE_CHART_LEGEND_PACKAGE_MODULE = "slide_chart_legend"
KEYNOTE_CHART_LEGEND_PACKAGE_METHODS = frozenset(
    {
        "slide_chart_legend_visible",
        "edit_slide_chart_legend",
        "apply_slide_chart_legend",
    }
)
KEYNOTE_CHART_LEGEND_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_CHART_LEGEND_RAW_ID_TOKEN = re.compile(
    r"\b(?:native_id|object_id|drawable_object_id|chart_object_id|identifier)\b"
)

# The Keynote chart Arrange-panel state lives behind the focused package
# owner.  Pages and Numbers have their own focused owners as well, so
# the former shared host shell is retired across all three formats.
KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_ARRANGEMENT_COMMON_SOURCE = Path(
    "crates/litchi-iwa-common/src/chart/arrangement.rs"
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_CHART_ARRANGEMENT_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_ARRANGEMENT_FAMILY_MODULE = 'arrangement'
KEYNOTE_CHART_ARRANGEMENT_PACKAGE_METHODS = frozenset(
    {
        "slide_chart_arrangement",
        "slide_chart_arrangements",
        "edit_slide_chart_arrangement",
        "apply_slide_chart_arrangement",
    }
)
KEYNOTE_CHART_ARRANGEMENT_HOST_TYPED_METHODS = frozenset(
    {
        "slide_chart_arrangement_by_selector",
        "set_slide_chart_arrangement_by_selector",
    }
)
KEYNOTE_CHART_ARRANGEMENT_EDIT_METHODS = frozenset(
    {"before", "after", "set", "set_locked", "set_constrain_proportions", "commit"}
)
KEYNOTE_CHART_ARRANGEMENT_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_TYPES = frozenset({"ChartArrangement"})
KEYNOTE_CHART_ARRANGEMENT_COMMON_METHODS = frozenset(
    {"new", "locked", "constrain_proportions", "with_locked", "with_constrain_proportions"}
)
KEYNOTE_CHART_ARRANGEMENT_SELECTOR_TYPES = frozenset(
    {"ChartSelector", "SlideSelector"}
)
KEYNOTE_CHART_ARRANGEMENT_FLAT_ALIASES = frozenset(
    {
        "Arrangement",
        "ArrangementCommit",
        "ArrangementDiagnostics",
        "ArrangementEdit",
        "ArrangementError",
        "ArrangementLimitKind",
        "ArrangementPatch",
        "ChartArrangementSnapshot",
        "ChartArrangementWrite",
    }
)
KEYNOTE_CHART_ARRANGEMENT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "ChartArrangementArchive",
        "ChartArrangementSnapshot",
        "ChartArrangementWrite",
    }
)
KEYNOTE_CHART_ARRANGEMENT_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "PreparedChartArrangementRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteReport",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_CHART_ARRANGEMENT_PROTO_ORIGINS = frozenset(
    {
        "buffa",
        "prost",
        "prost_types",
        "kn",
        "tsch",
        "tsp",
        "tsd",
        "litchi_iwa_protos",
    }
)
KEYNOTE_CHART_ARRANGEMENT_PUBLIC_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|"
    r"drawable_object_id|chart_object_id|component_id|member_id|archive_id|"
    r"message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_CHART_ARRANGEMENT_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|chart|native|component|archive|message|"
    r"resource|entry|metadata|package|uuid)[A-Za-z_]*(?:id|identifier))"
    r"[ \t\r\n]*:[ \t\r\n]*u64\b"
)
KEYNOTE_CHART_ARRANGEMENT_LEGACY_METHODS = frozenset(
    {"slide_chart_arrangement", "set_slide_chart_arrangement"}
)
RETIRED_IWA_KEYNOTE_CHART_ARRANGEMENT_HELPERS = ("chart_selector_for_drawable",)
RETIRED_IWA_KEYNOTE_CHART_ARRANGEMENT_HELPER_SET = frozenset(
    RETIRED_IWA_KEYNOTE_CHART_ARRANGEMENT_HELPERS
)
KEYNOTE_CHART_ARRANGEMENT_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?P<method>slide_chart_arrangement|"
    r"set_slide_chart_arrangement)(?![A-Za-z0-9_])[ \t\r\n]*\("
)
KEYNOTE_CHART_ARRANGEMENT_MONOLITH_PATH = re.compile(
    r"\b(?:crate|super)\s*::\s*charts\b"
)

KEYNOTE_CHART_CAPTION_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?P<method>set_slide_chart_caption|"
    r"remove_slide_chart_caption)(?![A-Za-z0-9_])[ \t\r\n]*\("
)
KEYNOTE_CHART_CAPTION_EDIT_METHODS = frozenset({"set", "clear", "commit"})
KEYNOTE_CHART_CAPTION_GRAPH_OPERATION = "rewrite_chart_caption_operation"
KEYNOTE_CHART_CAPTION_REFERENCE_TRANSITION = "patch_chart_caption_edge"
KEYNOTE_CHART_CAPTION_REFERENCE_TRANSITIONS = frozenset(
    {KEYNOTE_CHART_CAPTION_REFERENCE_TRANSITION, "patch_caption_edge"}
)
KEYNOTE_CHART_CAPTION_REFERENCE_TRANSITION_MARKERS = frozenset(
    {
        "expected_identifier",
        "replacement_identifier",
        "validate_selected_message_metadata",
        "rewrite_chart_caption_with_report",
        "replace_message_transitioning_object_references_preserving_header",
    }
)
KEYNOTE_CHART_CAPTION_SAVE_TOKEN_REWRITE_MARKERS = frozenset(
    {
        "prepare_package_metadata_additions_and_save_tokens",
        "prepare_package_metadata_save_tokens",
        "rewrite_package_metadata_additions_and_save_tokens",
        "rewrite_package_metadata_save_tokens",
    }
)
KEYNOTE_CHART_CAPTION_OBJECT_FRAMING_MARKERS = frozenset(
    {"validate_canonical_object_framing"}
)
KEYNOTE_CHART_CAPTION_CROSS_COMPONENT_MARKERS = frozenset(
    {"cross_component", "cross-component", "CrossComponent", "cross component"}
)
KEYNOTE_CHART_CAPTION_EXTERNAL_REFERENCE_MARKERS = frozenset(
    {
        "external_reference",
        "external_references",
        "ExternalReference",
        "ExternalReferenceDescriptor",
        "visit_external_reference",
        "external_ref",
    }
)
# The package owner is allowed to use physical/archive helpers privately, but
# every graph transition must run under one transaction-local aggregate budget.
# Keep these markers semantic rather than tying the ratchet to a concrete
# implementation type: a future owner may call it ``ChartCaptionBudget`` or
# ``CaptionPackageBudget`` while retaining the same accounting contract.
KEYNOTE_CHART_CAPTION_PACKAGE_BUDGET_TYPE = re.compile(
    r"(?m)^[ \t]*(?:pub[ \t]*\([ \t]*super[ \t]*\)[ \t]+)?struct[ \t]+"
    r"(?P<name>(?:Chart|Caption|Package)[A-Za-z0-9_]*Budget)\b"
)
KEYNOTE_CHART_CAPTION_CODEC_REPORT_CHARGE = re.compile(
    r"\b(?:budget|package_budget|aggregate_budget|transaction_budget)"
    r"[ \t]*(?:\.[ \t]*|::[ \t]*)"
    r"(?:charge|account)[A-Za-z0-9_]*(?:codec|wire|graph|metadata)[A-Za-z0-9_]*report\b"
    r"|\b(?:charge|account)_[A-Za-z0-9_]*(?:codec|wire|graph|metadata)"
    r"[A-Za-z0-9_]*report\b"
)
KEYNOTE_CHART_CAPTION_RESOURCE_CHARGES = {
    "archive": re.compile(
        r"\b(?:charge|account|reserve)[A-Za-z0-9_]*(?:archive|iwa)"
        r"[A-Za-z0-9_]*\b|\b(?:archive|iwa)[A-Za-z0-9_]*"
        r"(?:charge|account|reserve)[A-Za-z0-9_]*\b"
    ),
    "snappy": re.compile(
        r"\b(?:charge|account|reserve)[A-Za-z0-9_]*"
        r"(?:snappy|compress|decompress)[A-Za-z0-9_]*\b|\b(?:snappy|compress|decompress)"
        r"[A-Za-z0-9_]*(?:charge|account|reserve)[A-Za-z0-9_]*\b"
    ),
    "zip": re.compile(
        r"\b(?:charge|account|reserve)[A-Za-z0-9_]*(?:zip|reassembl|entry)[A-Za-z0-9_]*\b"
        r"|\b(?:zip|reassembl|entry)[A-Za-z0-9_]*(?:charge|account|reserve)[A-Za-z0-9_]*\b"
    ),
    "reopen": re.compile(
        r"\b(?:charge|account|reserve)[A-Za-z0-9_]*(?:reopen|candidate|verify|reparse)"
        r"[A-Za-z0-9_]*\b|\b(?:reopen|candidate|verify|reparse)[A-Za-z0-9_]*"
        r"(?:charge|account|reserve)[A-Za-z0-9_]*\b"
    ),
    "artifact": re.compile(
        r"\b(?:charge|account|reserve)[A-Za-z0-9_]*(?:artifact|exact)"
        r"[A-Za-z0-9_]*\b|\b(?:artifact|exact)[A-Za-z0-9_]*"
        r"(?:charge|account|reserve)[A-Za-z0-9_]*\b"
    ),
}
KEYNOTE_CHART_CAPTION_RAW_RESOURCE_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|caption|storage|reference|placement|style|chart|"
    r"resource|native|archive|message|component|entry|metadata|package|uuid)"
    r"[A-Za-z_]*(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_CHART_CAPTION_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|caption|storage|reference|placement|style|chart)"
    r"[A-Za-z_]*(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*u64\b"
)
KEYNOTE_CHART_CAPTION_FLAT_ALIASES = frozenset(
    {
        "Caption",
        "CaptionCommit",
        "CaptionDiagnostics",
        "CaptionEdit",
        "CaptionError",
        "CaptionLimitKind",
        "CaptionPatch",
        "ChartCaption",
        "ChartCaptionSnapshot",
        "ChartCaptionWrite",
    }
)
KEYNOTE_CHART_CAPTION_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "ChartCaptionSnapshot",
        "ChartCaptionWrite",
    }
)
KEYNOTE_CHART_CAPTION_WIRE_TYPES = frozenset(
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
KEYNOTE_CHART_CAPTION_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tsp", "tsd", "litchi_iwa_protos"}
)
KEYNOTE_CHART_CAPTION_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_chart_caption.rs"
)
KEYNOTE_CHART_CAPTION_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_CHART_CAPTION_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_CAPTION_FAMILY_MODULE = 'caption'
KEYNOTE_CHART_CAPTION_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_CHART_CAPTION_PACKAGE_METHODS = frozenset(
    {
        "slide_chart_caption",
        "edit_slide_chart_caption",
        "apply_slide_chart_caption",
    }
)

# Keynote chart-axis-title reads and mutations live behind the same
# selector-first package boundary as chart titles/captions.  Keep this gate
# independent from the Pages/Numbers native axis helpers: those editors still
# share the low-level chart-axis wire implementation until their own owners
# land.  Every audit below is dormant until both the focused owner source and
# its private package module declaration exist.
KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_AXIS_TITLE_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "selector.rs"
KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_chart_axis_title.rs"
)
KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_CHART_AXIS_TITLE_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_AXIS_TITLE_FAMILY_MODULE = 'axis_title'
KEYNOTE_CHART_AXIS_TITLE_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_TYPES = frozenset({"Axis"})
KEYNOTE_CHART_AXIS_TITLE_SELECTOR_TYPES = frozenset(
    {"ChartSelector", "SlideSelector"}
)
KEYNOTE_CHART_AXIS_TITLE_PACKAGE_METHODS = frozenset(
    {
        "slide_chart_axis_title",
        "edit_slide_chart_axis_title",
        "apply_slide_chart_axis_title",
    }
)
KEYNOTE_CHART_AXIS_TITLE_EDIT_METHODS = frozenset(
    {"before", "after", "set", "clear", "commit"}
)
KEYNOTE_CHART_AXIS_TITLE_FLAT_ALIASES = frozenset(
    {
        "AxisTitle",
        "AxisTitleCommit",
        "AxisTitleDiagnostics",
        "AxisTitleEdit",
        "AxisTitleError",
        "AxisTitleLimitKind",
        "AxisTitlePatch",
        "AxisTitleSnapshot",
        "AxisTitleWrite",
        "ChartAxisTitle",
        "ChartAxisTitleSnapshot",
        "ChartAxisTitleWrite",
    }
)
KEYNOTE_CHART_AXIS_TITLE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "ChartAxisTitleArchive",
        "ChartAxisTitleSnapshot",
        "ChartAxisTitleWrite",
    }
)
KEYNOTE_CHART_AXIS_TITLE_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "PreparedAxisTitleRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteReport",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_CHART_AXIS_TITLE_PROTO_ORIGINS = frozenset(
    {
        "buffa",
        "prost",
        "prost_types",
        "kn",
        "tsch",
        "tsp",
        "tsd",
        "tswp",
        "litchi_iwa_protos",
    }
)
KEYNOTE_CHART_AXIS_TITLE_PUBLIC_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|"
    r"drawable_object_id|chart_object_id|axis_id|component_id|member_id|"
    r"archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_CHART_AXIS_TITLE_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?P<method>set_slide_chart_axis_title|"
    r"remove_slide_chart_axis_title)(?![A-Za-z0-9_])[ \t\r\n]*\("
)

# The low-level codec is intentionally a private, format-neutral seam.  The
# build-script provenance check fixes the exact ingress/rewrite spelling, so
# keep those names explicit here as a second source-boundary ratchet.
KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_chart_axis_title_codec.rs"
)
KEYNOTE_CHART_AXIS_TITLE_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_CHART_AXIS_TITLE_CODEC_MODULE = "keynote_chart_axis_title_codec"
KEYNOTE_CHART_AXIS_TITLE_CODEC_REQUIRED_APIS = (
    "DecodeError",
    "DecodeLimit",
    "DecodeOptions",
    "AxisTitleSnapshot",
    "AxisTitleWrite",
    "DecodeReport",
    "RewriteReport",
    "decode_axis_titles",
    "decode_axis_titles_with_report",
    "prepare_axis_title_rewrite",
    "rewrite_axis_title",
    "AxisTitleKind",
)
KEYNOTE_CHART_AXIS_TITLE_CODEC_REQUIRED_MARKER_GROUPS = {
    "field-13-to-16 axis projection": (
        "CHART_AXIS_CATEGORY_VISIBLE_FIELD: u32 = 13",
        "CHART_AXIS_VALUE_VISIBLE_FIELD: u32 = 14",
        "CHART_AXIS_CATEGORY_TEXT_FIELD: u32 = 15",
        "CHART_AXIS_VALUE_TEXT_FIELD: u32 = 16",
    ),
    "strict canonical ingress": (
        "next_strict_field",
        "require_canonical_bool",
        "duplicate",
        "noncanonical",
    ),
    "unknown/raw preservation": (
        "unknown",
        "raw",
        "extend_from_slice",
    ),
    "finite prepared rewrite": (
        "MAX_RECURSION_LIMIT",
        "preflight_axis_titles",
        "execution_requirements",
        "execute",
    ),
    "typed axis projection": ("visible_title(self, kind: AxisTitleKind)",),
}
KEYNOTE_CHART_AXIS_TITLE_CODEC_PREPARED_TYPES = frozenset(
    {
        "PreparedAxisTitleRewrite",
        "RewriteExecutionRequirements",
        "RewriteExecutionLimits",
    }
)
KEYNOTE_CHART_AXIS_TITLE_CODEC_FLOW_MARKERS = {
    "decode report": re.compile(r"\bdecode_axis_titles_with_report\s*\("),
    "prepared rewrite": re.compile(r"\bprepare_axis_title_rewrite\s*\("),
    "execution requirements": re.compile(r"\bexecution_requirements\s*\("),
    "single prepared execute": re.compile(
        r"\b(?:prepared|rewrite)[A-Za-z0-9_]*[ \t\r\n]*\.[ \t\r\n]*execute\s*\("
    ),
}
KEYNOTE_CHART_AXIS_TITLE_PACKAGE_BUDGET_TYPE = re.compile(
    r"(?m)^[ \t]*(?:pub[ \t]*\([ \t]*super[ \t]*\)[ \t]+)?struct[ \t]+"
    r"(?P<name>(?:ChartAxisTitle|AxisTitle|Chart|Axis|Package|Transaction)"
    r"[A-Za-z0-9_]*Budget)\b"
)
KEYNOTE_CHART_AXIS_TITLE_PACKAGE_MARKER_GROUPS = {
    "aggregate transaction budget": (
        "Budget",
        "budget",
        "charge",
    ),
    "strict prepared codec": (
        "keynote_chart_axis_title_codec",
        "prepare_axis_title_rewrite",
        "execution_requirements",
        "execute",
    ),
    "exact artifacts/fingerprint/inverse/conflict": (
        "ExactArtifacts",
        "source_fingerprint",
        "inverse",
        "PatchConflict",
    ),
    "candidate reopen/locality": (
        "candidate",
        "reopen",
        "locality",
    ),
    "root preview deletion": (
        "root_preview_deletions",
        "deleted_previews",
    ),
    "source authority": (
        "authority",
        "ChartSelector",
        "Axis",
    ),
}
KEYNOTE_CHART_AXIS_TITLE_SOURCE_AUTHORITY_MARKER = re.compile(
    r"\b(?:authority|authorizes_source|source_authority|selection_authority)\b"
)
KEYNOTE_CHART_AXIS_TITLE_TEST_SOURCES = (
    Path("crates/litchi-keynote/tests/slide_chart_axis_title.rs"),
)
KEYNOTE_CHART_AXIS_TITLE_FUZZ_SOURCES = (
    Path(
        "crates/litchi-iwa-protos/fuzz/fuzz_targets/keynote_chart_axis_title_codec.rs"
    ),
    Path("crates/litchi/fuzz/fuzz_targets/keynote_chart_axis_title.rs"),
)
KEYNOTE_CHART_AXIS_TITLE_FUZZ_CORPORA = (
    Path("crates/litchi-iwa-protos/fuzz/corpus/keynote_chart_axis_title"),
    Path("crates/litchi/fuzz/corpus/keynote_chart_axis_title"),
)

# Keynote movie-caption replacement is the next focused package owner.  Keep
# this ratchet deliberately narrower than the chart-caption graph audit:
# phase 1 owns the selector-first existing-caption transaction, while the
# compatibility host may retain its title/graph creation/removal methods.
KEYNOTE_MOVIE_CAPTION_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "movie.rs"
KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_movie_caption.rs"
)
KEYNOTE_MOVIE_CAPTION_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_MOVIE_CAPTION_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "movie.rs"
KEYNOTE_MOVIE_CAPTION_FAMILY_MODULE = 'caption'
KEYNOTE_MOVIE_CAPTION_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_MOVIE_CAPTION_SELECTOR_TYPES = frozenset({"MovieSelector"})
KEYNOTE_MOVIE_CAPTION_PACKAGE_METHODS = frozenset(
    {
        "slide_movie_caption",
        "edit_slide_movie_caption",
        "apply_slide_movie_caption",
    }
)
KEYNOTE_MOVIE_CAPTION_EDIT_METHODS = frozenset({"set", "clear", "commit"})
KEYNOTE_MOVIE_CAPTION_FLAT_ALIASES = frozenset(
    {
        "Caption",
        "CaptionCommit",
        "CaptionDiagnostics",
        "CaptionEdit",
        "CaptionError",
        "CaptionLimitKind",
        "CaptionPatch",
        "MovieCaption",
        "MovieCaptionCommit",
        "MovieCaptionDiagnostics",
        "MovieCaptionEdit",
        "MovieCaptionError",
        "MovieCaptionLimitKind",
        "MovieCaptionPatch",
        "MovieCaptionSnapshot",
        "MovieCaptionWrite",
    }
)
KEYNOTE_MOVIE_CAPTION_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "MovieCaptionSnapshot",
        "MovieCaptionWrite",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_MOVIE_CAPTION_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_MOVIE_CAPTION_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tsp", "tsd", "litchi_iwa_protos"}
)
KEYNOTE_MOVIE_CAPTION_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|caption|storage|reference|placement|style|"
    r"native|archive|message|component|entry|metadata|package|uuid)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_MOVIE_CAPTION_GRAPH_OPERATION = "rewrite_movie_caption_operation"
KEYNOTE_MOVIE_CAPTION_GRAPH_OPERATIONS = frozenset(
    {
        KEYNOTE_MOVIE_CAPTION_GRAPH_OPERATION,
        "rewrite_caption_graph_operation_with_budget",
    }
)
KEYNOTE_MOVIE_CAPTION_REFERENCE_TRANSITION = "patch_movie_caption_edge"
KEYNOTE_MOVIE_CAPTION_REFERENCE_TRANSITIONS = frozenset(
    {KEYNOTE_MOVIE_CAPTION_REFERENCE_TRANSITION, "patch_caption_edge"}
)
KEYNOTE_MOVIE_CAPTION_REFERENCE_TRANSITION_MARKERS = frozenset(
    {
        "expected_identifier",
        "replacement_identifier",
        "validate_selected_message_metadata",
        "rewrite_movie_caption_with_report",
        "replace_message_transitioning_object_references_preserving_header",
    }
)
KEYNOTE_MOVIE_CAPTION_SAVE_TOKEN_REWRITE_MARKERS = frozenset(
    {
        "prepare_package_metadata_additions_and_save_tokens",
        "prepare_package_metadata_save_tokens",
        "rewrite_package_metadata_additions_and_save_tokens",
        "rewrite_package_metadata_save_tokens",
        "advance_package_save_token",
    }
)
KEYNOTE_MOVIE_CAPTION_METADATA_GRAPH_MARKERS = frozenset(
    {
        "add_component_object_uuids",
        "remove_component_object_uuids",
        "set_package_last_object_identifier",
        "ObjectUuidAddition",
        "ObjectUuidRemoval",
        "AdditionSaveTokenBatch",
        "RemovalSaveTokenBatch",
        "ExternalReferenceAddition",
        "ExternalReferenceRemoval",
        "DataReferenceOwnerRemoval",
    }
)
KEYNOTE_MOVIE_CAPTION_GRAPH_MARKERS = frozenset(
    {
        "CaptionObjectIds",
        "CaptionThemeStyle",
        "DrawableCaptionKind",
        "DrawableCaptionSlot",
        "IWorkThemeArchive",
        "ObjectGraph",
        "ArchiveObject",
        "RawMessage",
        "WireView",
        "DecodeOptions",
        "MovieCaptionWrite",
        "decode_movie_caption_identifier",
        "rewrite_movie_caption",
        "caption_objects",
        "drawable_caption_slot",
        "replace_object_reference",
        "standin_caption_object",
        "insert_slide_movie_caption",
        "insert_slide_movie_caption_standin",
        "replace_slide_movie_caption_reference",
        "next_object_identifier",
        "keynote_movie_caption_codec",
        "pages_movie_caption_codec",
        "litchi_iwa_protos",
        "prost",
        "prost_types",
        "kn",
        "tsa",
        "tsd",
        "tsp",
    }
)
KEYNOTE_MOVIE_CAPTION_OBJECT_FRAMING_MARKERS = frozenset(
    {"validate_canonical_object_framing"}
)
KEYNOTE_MOVIE_CAPTION_CROSS_COMPONENT_MARKERS = frozenset(
    {"cross_component", "cross-component", "CrossComponent", "cross component"}
)
KEYNOTE_MOVIE_CAPTION_EXTERNAL_REFERENCE_MARKERS = frozenset(
    {
        "external_reference",
        "external_references",
        "ExternalReference",
        "ExternalReferenceDescriptor",
        "visit_external_reference",
        "external_ref",
    }
)
# The movie-caption owner must account graph creation/removal and the existing
# text path under one private transaction budget.  This deliberately accepts
# the established ``CaptionPackageBudget``/``MovieCaptionBudget`` naming family
# without exposing the type through the public facade.
KEYNOTE_MOVIE_CAPTION_PACKAGE_BUDGET_TYPE = re.compile(
    r"(?m)^[ \t]*(?:pub[ \t]*\([ \t]*super[ \t]*\)[ \t]+)?struct[ \t]+"
    r"(?P<name>(?:Movie|Caption|Package)[A-Za-z0-9_]*Budget)\b"
)

# The *title* edge of a Keynote movie lives in the same focused
# package owner as the caption edge.  Keep this vocabulary separate from the
# caption ratchet: a title implementation must prove its own semantic
# facade, dual-edge codec, graph transition, and host retirement rather than
# inheriting a caption-only marker by name.
KEYNOTE_MOVIE_TITLE_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "movie.rs"
KEYNOTE_MOVIE_TITLE_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_movie_title.rs"
)
KEYNOTE_MOVIE_TITLE_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_MOVIE_TITLE_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "movie.rs"
KEYNOTE_MOVIE_TITLE_FAMILY_MODULE = 'title'
KEYNOTE_MOVIE_TITLE_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_MOVIE_TITLE_SELECTOR_TYPES = frozenset({"MovieSelector"})
KEYNOTE_MOVIE_TITLE_PACKAGE_METHODS = frozenset(
    {
        "slide_movie_title",
        "edit_slide_movie_title",
        "apply_slide_movie_title",
    }
)
KEYNOTE_MOVIE_TITLE_EDIT_METHODS = frozenset({"set", "clear", "commit"})
KEYNOTE_MOVIE_TITLE_FLAT_ALIASES = frozenset(
    {
        "Title",
        "TitleCommit",
        "TitleDiagnostics",
        "TitleEdit",
        "TitleError",
        "TitleLimitKind",
        "TitlePatch",
        "MovieTitle",
        "MovieTitleCommit",
        "MovieTitleDiagnostics",
        "MovieTitleEdit",
        "MovieTitleError",
        "MovieTitleLimitKind",
        "MovieTitlePatch",
        "MovieTitleSnapshot",
        "MovieTitleWrite",
    }
)
KEYNOTE_MOVIE_TITLE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "MovieTitleSnapshot",
        "MovieTitleWrite",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_MOVIE_TITLE_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_MOVIE_TITLE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tsa", "tsp", "tsd", "litchi_iwa_protos"}
)
KEYNOTE_MOVIE_TITLE_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|title|caption|storage|reference|placement|style|"
    r"native|archive|message|component|entry|metadata|package|uuid)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_MOVIE_TITLE_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_movie_caption_codec.rs"
)
KEYNOTE_MOVIE_TITLE_GRAPH_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_chart_caption_graph_codec.rs"
)
KEYNOTE_MOVIE_TITLE_CODEC_MODULE = "keynote_movie_caption_codec"
KEYNOTE_MOVIE_TITLE_GRAPH_CODEC_MODULE = "keynote_chart_caption_graph_codec"
KEYNOTE_MOVIE_TITLE_CODEC_FUNCTIONS = frozenset(
    {
        "decode_movie_caption",
        "decode_movie_caption_with_report",
        "rewrite_movie_caption_with_report",
        "encode_caption_graph_with_kind_with_report",
    }
)
KEYNOTE_MOVIE_TITLE_CODEC_TYPES = frozenset(
    {
        "MovieCaptionWrite",
        "MovieCaptionSnapshot",
        "CaptionGraphKind",
        "CaptionGraphWrite",
    }
)
KEYNOTE_MOVIE_TITLE_DUAL_EDGE_MARKERS = (
    frozenset(
        {
            "MovieCaptionWrite::title",
            "MovieCaptionSlot::Title",
            "MOVIE_TITLE_FIELD",
            "MOVIE_SUPER_FIELD",
            "TITLE_FIELD",
            "DRAWABLE_TITLE_FIELD",
            "title_identifier",
        }
    ),
    frozenset(
        {
            "CaptionGraphKind::Title",
            "CaptionOrTitleKind::Title",
            "DrawableCaptionKind::Title",
            "CAPTION_INFO_MESSAGE_TYPE",
            "CAPTION_INFO_FIELD",
            "TITLE_INFO_FIELD",
            "CaptionInfoArchive",
            "CaptionInfo",
            "caption_info",
            "TITLE_KIND",
        }
    ),
)
KEYNOTE_MOVIE_TITLE_GRAPH_OPERATION_NAMES = frozenset(
    {
        "rewrite_movie_title_operation",
        "rewrite_movie_title_graph_operation",
        "rewrite_caption_graph_operation_with_budget",
    }
)
KEYNOTE_MOVIE_TITLE_REFERENCE_TRANSITIONS = frozenset(
    {"patch_movie_title_edge", "patch_title_edge", "patch_caption_edge"}
)
KEYNOTE_MOVIE_TITLE_GRAPH_MARKERS = frozenset(
    {
        "MovieTitleGraph",
        "MovieTitleObjectIds",
        "MovieTitleSlot",
        "CaptionObjectIds",
        "CaptionThemeStyle",
        "DrawableCaptionKind::Title",
        "title_objects",
        "movie_title_slot",
        "standin_title_object",
        "insert_slide_movie_title",
        "insert_slide_movie_title_standin",
        "replace_slide_movie_title_reference",
        "replace_movie_title_reference",
        "next_object_identifier",
        "keynote_movie_caption_codec",
        "keynote_chart_caption_graph_codec",
        "keynote_movie_caption_codec",
    }
)
KEYNOTE_MOVIE_TITLE_METADATA_MARKERS = frozenset(
    {
        "prepare_package_metadata_additions_and_save_tokens",
        "prepare_package_metadata_save_tokens",
        "rewrite_package_metadata_additions_and_save_tokens",
        "rewrite_package_metadata_save_tokens",
        "AdditionSaveTokenBatch",
        "ObjectUuidAddition",
        "ExternalReferenceAddition",
        "add_component_object_uuids",
        "set_package_last_object_identifier",
    }
)
KEYNOTE_MOVIE_TITLE_RESOURCE_CHARGE_PATTERNS = {
    category: re.compile(
        rf"\b(?:charge|account|reserve)[A-Za-z0-9_]*{category}"
        rf"|\b{category}[A-Za-z0-9_]*\b[^;{{}}\n]{{0,180}}"
        r"\b(?:charge|account|reserve)[A-Za-z0-9_]*\b",
        re.I,
    )
    for category in ("archive", "snappy", "zip", "reopen", "artifact")
}
KEYNOTE_MOVIE_TITLE_CODEC_REPORT_CHARGE = re.compile(
    r"\b(?:charge|account|reserve)[A-Za-z0-9_]*(?:codec|report|wire)"
    r"|\b(?:codec|report|wire)[A-Za-z0-9_]*\b[^;{}\n]{0,180}"
    r"\b(?:charge|account|reserve)[A-Za-z0-9_]*\b",
    re.I,
)
KEYNOTE_MOVIE_TITLE_INVERSE_LOCALITY_MARKERS = frozenset(
    {
        "ExactArtifacts",
        "verify_caption_graph_transition",
        "verify_movie_title",
        "verify_title_candidate",
        "verify_candidate",
        "locality",
        "inverse",
        "PatchConflict",
    }
)
IWA_KEYNOTE_MOVIE_TITLE_CODEC_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:litchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*)?"
    r"(?:keynote_movie_title_codec|movie_title_codec)[ \t\r\n]*::[ \t\r\n]*"
    r"(?:decode|rewrite|validate)[A-Za-z0-9_]*\b"
)
# The scalar playback edge of a file-backed Keynote movie lives in the
# package crate.  This is intentionally separate from the title/caption
# graph ratchets above: playback changes only selected MovieArchive scalar
# fields and must not acquire graph/Metadata ownership by implication.
KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "media.rs"
KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_movie_playback.rs"
)
KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_MOVIE_PLAYBACK_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "movie.rs"
KEYNOTE_MOVIE_PLAYBACK_FAMILY_MODULE = 'playback'
KEYNOTE_MOVIE_PLAYBACK_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_TYPES = frozenset(
    {"MediaLoopMode", "MediaPlaybackSettings", "MediaVolume", "TimeField"}
)
KEYNOTE_MOVIE_PLAYBACK_SELECTOR_TYPES = frozenset({"MovieSelector", "SlideSelector"})
KEYNOTE_MOVIE_PLAYBACK_PACKAGE_METHODS = frozenset(
    {
        "slide_movie_playback_settings",
        "edit_slide_movie_playback_settings",
        "apply_slide_movie_playback_settings",
    }
)
KEYNOTE_MOVIE_PLAYBACK_EDIT_METHODS = frozenset({"set", "commit"})
KEYNOTE_MOVIE_PLAYBACK_FLAT_ALIASES = frozenset(
    {
        "Playback",
        "PlaybackSettings",
        "MoviePlayback",
        "MoviePlaybackSettings",
        "MoviePlaybackSnapshot",
        "MoviePlaybackWrite",
        "MoviePlaybackEdit",
        "MoviePlaybackPatch",
        "MoviePlaybackCommit",
        "MoviePlaybackDiagnostics",
        "MoviePlaybackError",
        "MoviePlaybackLimitKind",
        "PlaybackSettingsEdit",
        "PlaybackSettingsPatch",
        "PlaybackSettingsCommit",
        "PlaybackSettingsDiagnostics",
        "PlaybackSettingsError",
        "PlaybackSettingsLimitKind",
        "PlaybackEdit",
        "PlaybackPatch",
        "PlaybackCommit",
        "PlaybackDiagnostics",
        "PlaybackError",
        "PlaybackLimitKind",
    }
)
KEYNOTE_MOVIE_PLAYBACK_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "MoviePlaybackSnapshot",
        "MoviePlaybackWrite",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_MOVIE_PLAYBACK_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_MOVIE_PLAYBACK_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tsp", "tsd", "litchi_iwa_protos"}
)
KEYNOTE_MOVIE_PLAYBACK_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|playback|storage|reference|native|"
    r"archive|message|component|entry|metadata|package|uuid)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_MOVIE_PLAYBACK_TRANSACTION_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub[ \t]+)?mod[ \t]+(?:r#)?transaction\b"
)
KEYNOTE_MOVIE_PLAYBACK_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/movie_playback_codec.rs"
)
KEYNOTE_MOVIE_PLAYBACK_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_MOVIE_PLAYBACK_CODEC_MODULE = "movie_playback_codec"
KEYNOTE_MOVIE_PLAYBACK_CODEC_FUNCTIONS = frozenset(
    {
        "decode_movie_playback_with_report",
        "prepare_movie_playback_rewrite",
        "rewrite_movie_playback",
    }
)
KEYNOTE_MOVIE_PLAYBACK_CODEC_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "MoviePlaybackSnapshot",
        "MoviePlaybackWrite",
        "PreparedMoviePlaybackRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
    }
)
KEYNOTE_MOVIE_PLAYBACK_CODEC_MARKER_GROUPS = {
    "private Buffa/lazy ingress": (
        "buffa",
        "Buffa",
        "lazy",
        "Lazy",
        "visit_with",
        "sidecar",
    ),
    "unknown/raw preservation": (
        "unknown",
        "Unknown",
        "overlong",
        "balanced",
        "raw_fields",
    ),
    "resource/preflight accounting": (
        "max_input",
        "max_output",
        "max_fields",
        "max_work",
        "max_depth",
        "max_references",
        "try_reserve",
        "RewriteExecutionRequirements",
    ),
}
KEYNOTE_MOVIE_PLAYBACK_FUZZ_SOURCES = (
    Path("crates/litchi-iwa-protos/fuzz/fuzz_targets/movie_playback_codec.rs"),
    Path("crates/litchi/fuzz/fuzz_targets/keynote_movie_playback.rs"),
)
KEYNOTE_MOVIE_PLAYBACK_FUZZ_CORPORA = (
    Path("crates/litchi-iwa-protos/fuzz/corpus/movie_playback_codec"),
    Path("crates/litchi/fuzz/corpus/keynote_movie_playback"),
)
KEYNOTE_MOVIE_PLAYBACK_PACKAGE_MARKER_GROUPS = {
    "aggregate transaction budget": ("TransactionBudget", "transaction_budget", "budget"),
    "preflight/candidate": ("preflight", "candidate", "reopen", "locality"),
    "exact inverse/artifacts": ("ExactArtifacts", "inverse", "PatchConflict", "same_content"),
    "codec report accounting": ("codec_report", "charge_codec", "wire", "fields", "work"),
    "physical archive accounting": ("archive", "snappy", "Snappy", "zip", "ZIP", "output"),
}
IWA_KEYNOTE_MOVIE_PLAYBACK_TYPED_METHODS = frozenset(
    {
        "slide_movie_playback_settings_by_selector",
        "set_slide_movie_playback_settings_by_selector",
    }
)

# The compatibility Keynote media readers are retired: the focused package
# owns semantic movie/audio reads.  The old reader methods and
# projection types must disappear from production and test-only source alike;
# only the private graph fixture helpers (which use distinct names) remain
# valid.  Scan the complete Keynote host tree plus integration/examples so an
# old call cannot survive in a fixture while the public implementation is
# removed.
IWA_KEYNOTE_MEDIA_READER_METHODS = frozenset(
    {"slide_movies", "slide_audio", "slide_media_infos"}
)
IWA_KEYNOTE_MEDIA_READER_TYPES = frozenset(
    {"KeynoteSlideMovieInfo", "KeynoteSlideAudioInfo"}
)


# Keynote movie geometry lives behind a selector-first package facade.
# Keep this separate from playback/title/caption and from generic image, shape,
# or audio geometry. Geometry writes own preview invalidation as an explicit
# transition concern.
KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "media.rs"
KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_movie_geometry.rs"
)
KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_MOVIE_GEOMETRY_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "movie.rs"
KEYNOTE_MOVIE_GEOMETRY_FAMILY_MODULE = 'geometry'
KEYNOTE_MOVIE_GEOMETRY_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_TYPES = frozenset({"MovieGeometry"})
KEYNOTE_MOVIE_GEOMETRY_SELECTOR_TYPES = frozenset({"MovieSelector", "SlideSelector"})
KEYNOTE_MOVIE_GEOMETRY_PACKAGE_METHODS = frozenset(
    {
        "slide_movie_geometry",
        "edit_slide_movie_geometry",
        "apply_slide_movie_geometry",
    }
)
KEYNOTE_MOVIE_GEOMETRY_EDIT_METHODS = frozenset(
    {"set", "restore_original_size", "set_transform", "flip", "commit"}
)
KEYNOTE_MOVIE_GEOMETRY_FLAT_ALIASES = frozenset(
    {
        "Geometry",
        "GeometryCommit",
        "GeometryDiagnostics",
        "GeometryEdit",
        "GeometryError",
        "GeometryLimitKind",
        "GeometryPatch",
        "MovieGeometryCommit",
        "MovieGeometryDiagnostics",
        "MovieGeometryEdit",
        "MovieGeometryError",
        "MovieGeometryLimitKind",
        "MovieGeometryPatch",
        "MovieGeometrySnapshot",
        "MovieGeometryWrite",
    }
)
KEYNOTE_MOVIE_GEOMETRY_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "MovieGeometrySnapshot",
        "MovieGeometryWrite",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_MOVIE_GEOMETRY_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_MOVIE_GEOMETRY_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tsp", "tsd", "litchi_iwa_protos"}
)
KEYNOTE_MOVIE_GEOMETRY_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|geometry|storage|reference|placement|"
    r"native|archive|message|component|entry|metadata|package|uuid)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_MOVIE_GEOMETRY_TRANSACTION_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub[ \t]+)?mod[ \t]+(?:r#)?transaction\b"
)
KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_movie_geometry_codec.rs"
)
KEYNOTE_MOVIE_GEOMETRY_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_MOVIE_GEOMETRY_CODEC_MODULE = "keynote_movie_geometry_codec"
KEYNOTE_MOVIE_GEOMETRY_CODEC_FUNCTIONS = frozenset(
    {
        "decode_movie_geometry_with_report",
        "prepare_movie_geometry_rewrite",
        "rewrite_movie_geometry",
    }
)
KEYNOTE_MOVIE_GEOMETRY_CODEC_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "MovieGeometrySnapshot",
        "MovieGeometryWrite",
        "PreparedMovieGeometryRewrite",
        "MoviePositionWrite",
        "PreparedMoviePositionRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
    }
)
KEYNOTE_MOVIE_GEOMETRY_CODEC_MARKER_GROUPS = {
    "private Buffa/lazy ingress": (
        "buffa",
        "Buffa",
        "lazy",
        "Lazy",
        "visit_with",
        "sidecar",
    ),
    "unknown/raw preservation": (
        "unknown",
        "Unknown",
        "overlong",
        "balanced",
        "raw_fields",
    ),
    "resource/preflight accounting": (
        "max_input",
        "max_output",
        "max_fields",
        "max_work",
        "max_depth",
        "max_references",
        "try_reserve",
        "RewriteExecutionRequirements",
    ),
}
KEYNOTE_MOVIE_GEOMETRY_FUZZ_SOURCES = (
    Path("crates/litchi-iwa-protos/fuzz/fuzz_targets/keynote_movie_geometry_codec.rs"),
    Path("crates/litchi/fuzz/fuzz_targets/keynote_slide_movie_geometry.rs"),
)
KEYNOTE_MOVIE_GEOMETRY_FUZZ_CORPORA = (
    Path("crates/litchi-iwa-protos/fuzz/corpus/keynote_movie_geometry_codec"),
    Path("crates/litchi/fuzz/corpus/keynote_slide_movie_geometry"),
)
KEYNOTE_MOVIE_GEOMETRY_PACKAGE_MARKER_GROUPS = {
    "aggregate transaction budget": ("TransactionBudget", "transaction_budget", "budget"),
    "preflight/candidate": ("preflight", "candidate", "reopen", "locality"),
    "exact inverse/artifacts": ("ExactArtifacts", "inverse", "PatchConflict", "same_content"),
    "codec report accounting": ("codec_report", "charge_codec", "wire", "fields", "work"),
    "physical archive accounting": ("archive", "snappy", "Snappy", "zip", "ZIP", "output"),
    "preview invalidation": (
        "preview",
        "previews",
        "invalidate_preview",
        "invalidated_previews",
        "deleted_previews",
    ),
}
# This is the completion seam for movie geometry.  The geometry owner
# intentionally stopped at position/size and therefore still permits a
# compatibility implementation for native Arrange angle/flip/flags.  Keep that
# historical ratchet unchanged, but give the eventual expanded owner an
# explicit activation token.  This avoids silently treating a partial owner as
# complete while making the retirement contract executable once the semantic
# and codec surfaces land together.
KEYNOTE_MOVIE_GEOMETRY_COMPLETE_ACTIVATION = re.compile(
    r"(?<![A-Za-z0-9_])(?:MOVIE_GEOMETRY_(?:COMPLETE|FULL_OWNERSHIP)|"
    r"MovieGeometryCapabilities::(?:Complete|AngleFlagsFlip))(?![A-Za-z0-9_])"
)
KEYNOTE_MOVIE_GEOMETRY_COMPLETE_CAPABILITY_MARKERS = {
    "angle": re.compile(
        r"(?m)^[ \t]*(?:pub(?:\([^)]*\))?[ \t]+)?(?:const[ \t]+)?fn[ \t]+"
        r"(?:angle|rotation|radians)\b|\b(?:MOVIE_GEOMETRY|GEOMETRY)_ANGLE_FIELD\b|"
        r"\b(?:MovieGeometry|MovieGeometryWrite)::(?:angle|rotation|radians)\b"
    ),
    "flags": re.compile(
        r"(?m)^[ \t]*(?:pub(?:\([^)]*\))?[ \t]+)?(?:const[ \t]+)?fn[ \t]+"
        r"(?:flags|geometry_flags)\b|\b(?:MOVIE_GEOMETRY|GEOMETRY)_FLAGS_FIELD\b|"
        r"\b(?:MovieGeometry|MovieGeometryWrite)::(?:flags|geometry_flags)\b"
    ),
    "flip": re.compile(
        r"(?m)^[ \t]*pub(?:\([^)]*\))?[ \t]+(?:enum|struct)[ \t]+"
        r"(?:Movie)?Flip(?:Axis|Mode)?\b|"
        r"^[ \t]*(?:pub(?:\([^)]*\))?[ \t]+)?(?:const[ \t]+)?fn[ \t]+"
        r"(?:flip|flip_axis|horizontal_flip|vertical_flip)\b|"
        r"\b(?:MOVIE_GEOMETRY|GEOMETRY)_FLIP(?:_[A-Z0-9_]+)?_FIELD\b"
    ),
}
KEYNOTE_MOVIE_GEOMETRY_COMPLETE_CODEC_FLOW_MARKERS = {
    "prepared rewrite": re.compile(
        r"\b(?:keynote_movie_geometry_codec[ \t\r\n]*::[ \t\r\n]*)?"
        r"prepare_movie_(?:geometry|transform)_rewrite\s*\("
    ),
    "codec report": re.compile(
        r"\b(?:prepare_report|codec_report)\s*\("
    ),
    "execution requirements": re.compile(
        r"\bexecution_requirements\s*\("
    ),
    "single prepared execute": re.compile(
        r"\b(?:prepared|rewrite)[A-Za-z0-9_]*[ \t\r\n]*\.[ \t\r\n]*"
        r"execute\s*\("
    ),
    "residual finite options": re.compile(
        r"\b(?:residual|remaining_[A-Za-z0-9_]*|"
        r"with_max_(?:input|output|fields|work|depth|allocations|retained|scratch))\b"
    ),
}
KEYNOTE_MOVIE_GEOMETRY_COMPLETE_DIRECT_REWRITE_CALL = re.compile(
    r"\b(?:keynote_movie_geometry_codec[ \t\r\n]*::[ \t\r\n]*)?"
    r"rewrite_movie_(?:geometry|transform)\s*\("
)
IWA_KEYNOTE_MOVIE_GEOMETRY_COMPLETE_RAW_HELPER = re.compile(
    # ``set_movie_geometry`` and the patch helpers are shared physical
    # executors used by audio/properties and duplicate-object lifecycle paths;
    # they are not a focused movie-geometry fallback.  Keep only the
    # movie-specific compatibility helpers here.
    r"(?<![A-Za-z0-9_])(?:try_set_file_movie_geometry_with_package|"
    r"flip_drawable_geometry|restore_drawable_original_size|file_movie_selector)"
    r"(?![A-Za-z0-9_])"
)
KEYNOTE_MOVIE_GEOMETRY_COMPLETE_CODEC_CAPABILITY_MARKERS = {
    "codec angle field": re.compile(
        r"\b(?:GEOMETRY|MOVIE_GEOMETRY)_ANGLE_FIELD\b|"
        r"\b(?:angle_degrees|set_angle_degrees|clear_angle_degrees)\b"
    ),
    "codec flags field": re.compile(
        r"\b(?:GEOMETRY|MOVIE_GEOMETRY)_FLAGS_FIELD\b|"
        r"\b(?:geometry_flags|set_flags|clear_flags)\b"
    ),
    "codec transform rewrite": re.compile(
        r"\b(?:prepare_movie_transform_rewrite|"
        r"PreparedMovieTransformRewrite|MovieTransformWrite|TransformField)\b"
    ),
}
KEYNOTE_MOVIE_GEOMETRY_COMPLETE_OWNER_CAPABILITY_MARKERS = {
    "transform read": re.compile(
        r"\bslide_movie_transform(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]*\("
    ),
    "transform edit": re.compile(r"\bset_transform\s*\("),
    "flip edit": re.compile(r"\bflip\s*\("),
}
IWA_KEYNOTE_MOVIE_GEOMETRY_LEGACY_METHODS = frozenset(
    {
        "slide_movie_geometry",
        "set_slide_movie_geometry",
        "restore_slide_movie_original_size",
        "flip_slide_movie",
    }
)

# The expanded geometry owner is admitted only with a native-authored source
# oracle, an accepted native candidate, and checked-in integration tests that
# exercise the all-media source order.  The source oracle records Arrange's
# persisted angle/flip/original-size fields; the focused candidate is the
# operation-level save/reopen receipt.
KEYNOTE_MOVIE_GEOMETRY_NATIVE_TEST_SOURCE = Path(
    "crates/litchi-keynote/tests/slide_movie_geometry_native.rs"
)
KEYNOTE_MOVIE_GEOMETRY_NATIVE_EVIDENCE = (
    (
        Path("test-data/iwork/keynote/slide-movie-geometry-source-native.key"),
        "native_arrange_oracle_retains_original_size_rotation_and_reflection",
    ),
    (
        Path("test-data/iwork/keynote/slide-movie-geometry-focused-native.key"),
        "native_geometry_saved_candidates_remain_editable_and_reversible",
    ),
    (
        Path("test-data/iwork/keynote/slide-movie-geometry-vertical-focused-native.key"),
        "native_geometry_saved_candidates_remain_editable_and_reversible",
    ),
)

# The existing host audio editor still exposes the raw-ID position reader and
# writer.  Keep their retirement dormant until the focused owner is actually
# wired into ``Package``: source presence alone is not an ownership handoff.
# The owner intentionally reuses the movie-geometry wire projection because
# audio controls carry the same ``MovieArchive.super.geometry.position`` edge;
# this boundary therefore checks the position-specific codec seam without
# inventing a second generated or eager protobuf surface.
KEYNOTE_AUDIO_POSITION_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "media.rs"
KEYNOTE_AUDIO_POSITION_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_audio_position.rs"
)
KEYNOTE_AUDIO_POSITION_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_AUDIO_POSITION_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "audio.rs"
KEYNOTE_AUDIO_POSITION_FAMILY_MODULE = 'position'
KEYNOTE_AUDIO_POSITION_CODEC_SOURCE = KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE
KEYNOTE_AUDIO_POSITION_CODEC_PUBLIC_SOURCE = KEYNOTE_MOVIE_GEOMETRY_CODEC_PUBLIC_SOURCE
KEYNOTE_AUDIO_POSITION_CODEC_MODULE = KEYNOTE_MOVIE_GEOMETRY_CODEC_MODULE
KEYNOTE_AUDIO_POSITION_OWNER_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?slide_audio_position[ \t]*;"
)
KEYNOTE_AUDIO_POSITION_PUBLIC_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?![ \t]*\([ \t]*crate[ \t]*\))(?:\([^()]*\))?[ \t]+mod[ \t]+"
    r"(?:r#)?slide_audio_position\b"
)
KEYNOTE_AUDIO_POSITION_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_AUDIO_POSITION_SELECTOR_TYPES = frozenset({"SlideSelector", "MovieSelector"})
KEYNOTE_AUDIO_POSITION_PACKAGE_METHODS = frozenset(
    {
        "slide_audio_position",
        "edit_slide_audio_position",
        "apply_slide_audio_position",
    }
)
KEYNOTE_AUDIO_POSITION_EDIT_METHODS = frozenset({"set", "commit"})
KEYNOTE_AUDIO_POSITION_CODEC_FUNCTIONS = frozenset(
    {"decode_movie_position_with_report", "prepare_movie_position_rewrite"}
)
KEYNOTE_AUDIO_POSITION_CODEC_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeOptions",
        "DecodeReport",
        "MovieGeometrySnapshot",
        "MovieGeometryWrite",
        "PreparedMovieGeometryRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
    }
)
KEYNOTE_AUDIO_POSITION_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "Entry",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_AUDIO_POSITION_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "PreparedMovieGeometryRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteReport",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_AUDIO_POSITION_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|audio|native|archive|component|"
    r"message|resource|entry|metadata|package|uuid)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|u32|usize|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_AUDIO_POSITION_CODEC_MARKER_GROUPS = {
    "lazy Buffa ingress": (
        "buffa::DecodeOptions",
        "decode_lazy_view",
    ),
    "strict source accounting": (
        "max_message_bytes",
        "max_fields",
        "max_work_bytes",
        ("unknown", "raw"),
    ),
    "prepared source-preserving rewrite": (
        "prepare_movie_position_rewrite",
        "output_bytes",
        "source",
        "execute",
    ),
}
KEYNOTE_AUDIO_POSITION_TEST_SOURCE = Path(
    "crates/litchi-keynote/tests/slide_audio_position.rs"
)
KEYNOTE_AUDIO_POSITION_FUZZ_SOURCE = Path(
    "crates/litchi-keynote/fuzz/fuzz_targets/keynote_slide_audio_position.rs"
)
KEYNOTE_AUDIO_POSITION_FUZZ_CORPUS = Path(
    "crates/litchi-keynote/fuzz/corpus/keynote_slide_audio_position"
)
KEYNOTE_AUDIO_POSITION_TEST_MARKERS = (
    "#[test]",
    "slide_audio_position",
    "edit_slide_audio_position",
    "apply_slide_audio_position",
    "inverse",
    "preview",
    "atomic",
    "limit",
)
KEYNOTE_AUDIO_POSITION_OWNER_MARKER_GROUPS = {
    "aggregate transaction budget": (
        re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*Budget\b"),
        re.compile(r"\b(?:for_package|for_source|new)[ \t\r\n]*\("),
    ),
    "selector resolution": (
        re.compile(r"\b(?:SlideSelector|MovieSelector)\b"),
        re.compile(r"\b(?:select|resolve|position)"),
    ),
    "candidate verification": (
        re.compile(r"\b(?:candidate|reopen|readback|verify|verification)\b"),
    ),
    "exact source and inverse": (
        re.compile(r"\b(?:ExactArtifacts|source_fingerprint|inverse|PatchConflict)\b"),
    ),
    "preview invalidation": (
        re.compile(r"\b(?:preview|previews|invalidate_preview|deleted_previews)\b"),
    ),
}


# The next focused Keynote media slice moves the four scalar drawable
# properties shared by file movies and audio controls behind one selector-first
# Package owner.  Keep this inventory independent from geometry and playback:
# properties do not require a complete positive movie size, and the lock bit is
# itself an editable property.  The owner is intentionally dormant until the
# private Package module, semantic value, and neutral codec are all present.
KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "media" / "properties.rs"
)
KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_media_properties.rs"
)
KEYNOTE_MEDIA_PROPERTIES_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_MEDIA_PROPERTIES_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "media" / "properties.rs"
KEYNOTE_MEDIA_PROPERTIES_FAMILY_MODULE = None
KEYNOTE_MEDIA_PROPERTIES_OWNER_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?slide_media_properties[ \t]*;"
)
KEYNOTE_MEDIA_PROPERTIES_PUBLIC_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?![ \t]*\([ \t]*crate[ \t]*\))(?:\([^()]*\))?[ \t]+mod[ \t]+"
    r"(?:r#)?slide_media_properties\b"
)
KEYNOTE_MEDIA_PROPERTIES_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_TYPES = frozenset({"MediaProperties"})
KEYNOTE_MEDIA_PROPERTIES_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "MovieSelector"}
)
KEYNOTE_MEDIA_PROPERTIES_PACKAGE_METHODS = frozenset(
    {
        "slide_media_properties",
        "edit_slide_media_properties",
        "apply_slide_media_properties",
    }
)
KEYNOTE_MEDIA_PROPERTIES_EDIT_METHODS = frozenset(
    {"set", "commit"}
)
KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_media_properties_codec.rs"
)
KEYNOTE_MEDIA_PROPERTIES_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_MEDIA_PROPERTIES_CODEC_MODULE = "keynote_media_properties_codec"
KEYNOTE_MEDIA_PROPERTIES_CODEC_GENERATED_MODULE = (
    "buffa_keynote_media_properties_generated"
)
KEYNOTE_MEDIA_PROPERTIES_CODEC_FUNCTIONS = frozenset(
    {
        "decode_movie_properties",
        "decode_movie_properties_with_report",
        "prepare_movie_properties_rewrite",
        "rewrite_movie_properties",
    }
)
KEYNOTE_MEDIA_PROPERTIES_CODEC_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeOptions",
        "DecodeReport",
        "MoviePropertiesSnapshot",
        "MoviePropertiesWrite",
        "PreparedMoviePropertiesRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
    }
)
KEYNOTE_MEDIA_PROPERTIES_CODEC_MARKER_GROUPS = {
    "lazy Buffa ingress": (
        ("buffa::DecodeOptions", "buffa", "decode_lazy_view", "decode_view"),
    ),
    "strict canonical ingress": (
        ("preflight", "parse_message", "require_known_framing"),
        ("duplicate",),
        ("noncanonical", "non-canonical"),
    ),
    "unknown/raw preservation": (
        ("unknown", "Unknown"),
        ("raw", "Raw"),
        ("preserv",),
    ),
    "bounded prepared rewrite": (
        ("max_message_bytes", "max_input"),
        ("max_fields",),
        ("max_work_bytes", "max_work"),
        ("execution_requirements", "RewriteExecutionRequirements"),
        ("execute",),
    ),
    "drawable property fields": (
        ("hyperlink", "hyperlink_url"),
        ("locked",),
        ("aspect_ratio_locked", "aspect"),
        ("accessibility_description", "accessibility"),
    ),
}
KEYNOTE_MEDIA_PROPERTIES_CODEC_EAGER_PATTERNS = (
    re.compile(r"\bprost(?:_types)?\b"),
    re.compile(r"\bMessage\s*::\s*(?:decode|encode)\s*\("),
    re.compile(r"\b(?:to_owned_message|try_encode_to_vec|encode_to_vec)\s*\("),
)
KEYNOTE_MEDIA_PROPERTIES_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "Entry",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_MEDIA_PROPERTIES_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "PreparedMoviePropertiesRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteReport",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_MEDIA_PROPERTIES_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "litchi_iwa_protos"}
)
KEYNOTE_MEDIA_PROPERTIES_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|audio|native|archive|component|"
    r"message|resource|entry|metadata|package|uuid)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|u32|usize|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_MEDIA_PROPERTIES_OWNER_MARKER_GROUPS = {
    "aggregate transaction budget": (
        re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*Budget\b"),
        re.compile(r"\b(?:for_package|for_source|new)[ \t\r\n]*\("),
    ),
    "selector resolution": (
        re.compile(r"\b(?:SlideSelector|MovieSelector)\b"),
        re.compile(r"\b(?:select|resolve|movie|media)[A-Za-z_]*"),
    ),
    "candidate verification": (
        re.compile(r"\b(?:candidate|reopen|readback|verify|verification)\b"),
    ),
    "exact source and inverse": (
        re.compile(r"\b(?:ExactArtifacts|source_fingerprint|inverse|PatchConflict)\b"),
    ),
    "property locality and preservation": (
        re.compile(r"(?:locality|unchanged|unknown|raw|preserv)"),
    ),
    "property decode accounting": (
        re.compile(r"\bbudget_property_report\b"),
    ),
    "property rewrite accounting": (
        re.compile(r"\bbudget_property_requirements\b"),
    ),
    "prepublication accounting": (
        re.compile(
            r"\b(?:codec_report|charge_codec|wire|fields|work|archive|snappy|"
            r"zip|ZIP|output|preview)[A-Za-z_]*\b"
        ),
    ),
}


# Replacement of already-materialized Keynote slide movie/audio data lives
# in one selector-first package owner.  The owner deliberately covers
# the shared media payload edge (movie content, movie poster, and audio
# content) without owning movie/audio creation, duplication, deletion,
# playback, geometry, title/caption, or generic ``Draft::replace_media``
# compatibility.  Keep this inventory independent from the playback and
# geometry ratchets: a payload replacement must prove both the nested
# ``TSP.DataReference`` edge and the PackageMetadata digest/length closure.
KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_media_replacement.rs"
)
KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_MEDIA_DATA_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "media.rs"
KEYNOTE_SLIDE_MEDIA_DATA_FAMILY_MODULE = 'data'
KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "media.rs"
KEYNOTE_SLIDE_MEDIA_DATA_PACKAGE_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub\(crate\)[ \t]+)?mod[ \t]+(?:r#)?slide_media_replacement\s*;"
)
KEYNOTE_SLIDE_MEDIA_DATA_PUBLIC_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?![ \t]*\([ \t]*crate[ \t]*\))(?:\([^()]*\))?[ \t]+mod[ \t]+"
    r"(?:r#)?slide_media_replacement\b"
)
KEYNOTE_SLIDE_MEDIA_DATA_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_TYPES = frozenset(
    {"MediaPart", "Data"}
)
KEYNOTE_SLIDE_MEDIA_DATA_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "MovieSelector"}
)
KEYNOTE_SLIDE_MEDIA_DATA_PACKAGE_METHODS = frozenset(
    {
        "slide_media_data",
        "edit_slide_media_data",
        "apply_slide_media_data",
    }
)
KEYNOTE_SLIDE_MEDIA_DATA_EDIT_METHODS = frozenset(
    {"before", "after", "part", "set", "commit"}
)
KEYNOTE_SLIDE_MEDIA_DATA_FLAT_ALIASES = frozenset(
    {
        "MediaData",
        "MediaDataCommit",
        "MediaDataDiagnostics",
        "MediaDataEdit",
        "MediaDataError",
        "MediaDataLimitKind",
        "MediaDataPatch",
        "SlideMediaReplacement",
        "SlideMediaReplacementCommit",
        "SlideMediaReplacementDiagnostics",
        "SlideMediaReplacementEdit",
        "SlideMediaReplacementError",
        "SlideMediaReplacementLimitKind",
        "SlideMediaReplacementPatch",
    }
)
KEYNOTE_SLIDE_MEDIA_DATA_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "MediaRecord",
        "MetadataVisitor",
        "OwnerRecord",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_SLIDE_MEDIA_DATA_WIRE_TYPES = frozenset(
    {
        "ComponentSnapshot",
        "ComponentSelector",
        "DataInfoAddition",
        "DataInfoContentReplacement",
        "DataInfoRemoval",
        "DataInfoSnapshot",
        "DataMetadataMapEntry",
        "DataMetadataMapSource",
        "DataMetadataMapVisitor",
        "DataReferenceOwnerAddition",
        "DataReferenceOwnerCountUpdate",
        "DataReferenceOwnerRemoval",
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "MediaRewriteBatch",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "OwnerSnapshot",
        "RewriteError",
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
KEYNOTE_SLIDE_MEDIA_DATA_PROTO_ORIGINS = frozenset(
    {
        "buffa",
        "prost",
        "prost_types",
        "kn",
        "tsa",
        "tsd",
        "tsp",
        "litchi_iwa_protos",
    }
)
KEYNOTE_SLIDE_MEDIA_DATA_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|media|data|poster|audio|component|"
    r"archive|message|resource|entry|metadata|package|uuid)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|u32|usize|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES = (
    Path("crates/litchi-iwa-protos/src/keynote_media_codec.rs"),
    Path("crates/litchi-iwa-protos/src/package_metadata_media_codec.rs"),
)
KEYNOTE_SLIDE_MEDIA_DATA_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_SLIDE_MEDIA_DATA_CODEC_APIS = {
    "keynote_media_codec": (
        "decode_data_reference",
        "decode_data_reference_with_report",
        "DataReferenceSnapshot",
        "DecodeOptions",
        "DecodeReport",
    ),
    "package_metadata_media_codec": (
        "visit_package_metadata_media",
        "prepare_package_metadata_media_rewrite",
        "rewrite_package_metadata_media",
        "MediaRewriteBatch",
        "DataInfoContentReplacement",
        "DataInfoRemoval",
        "DataReferenceOwnerRemoval",
        "DataMetadataMapSource",
        "DataMetadataMapEntry",
        "DataMetadataMapVisitor",
        "RewriteExecutionRequirements",
    ),
}
KEYNOTE_SLIDE_MEDIA_DATA_CODEC_MARKERS = {
    "keynote media lazy Buffa ingress": (
        "buffa",
        "decode_lazy_view",
        "DataReferenceLazyView",
    ),
    "metadata media lazy Buffa ingress": (
        "buffa",
        "decode_lazy_view",
        "PackageMetadataMedia",
    ),
    "metadata source-preserving rewrite": (
        "unknown",
        "raw",
        "extend_from_slice",
        "prepared",
    ),
    "finite media resource accounting": (
        "max_message_bytes",
        "max_fields",
        "max_work_bytes",
        "try_reserve",
    ),
}
# The physical clone/remap primitive stays in the archive core.  The focused
# Keynote owner may use it through private adapters; the implementation may
# live in archive.rs or a private archive child module as the core is split.
KEYNOTE_SLIDE_MEDIA_DATA_CORE_SOURCE_ROOT = Path("crates/litchi-iwa-core/src")
KEYNOTE_SLIDE_MEDIA_DATA_CORE_CLONE_METHODS = (
    "clone_with_identity_remap",
)
KEYNOTE_SLIDE_MEDIA_DATA_CORE_CLONE_FORBIDDEN_IMPORTS = re.compile(
    r"(?<![A-Za-z0-9_])(?:buffa|prost|prost_types|litchi_iwa_protos|"
    r"(?:kn|tsa|tsd|tsp|tn|tp|tswp)(?:sos)?)[ \t\r\n]*::"
)
KEYNOTE_SLIDE_MEDIA_DATA_METADATA_MAP_APIS = (
    "DataMetadataMapSource",
    "DataMetadataMapEntry",
    "DataMetadataMapVisitor",
)
KEYNOTE_SLIDE_MEDIA_DATA_METADATA_PUBLIC_COLLECTION = re.compile(
    r"(?m)^[ \t]*pub(?:\([^()]*\))?[ \t]+(?:struct|enum|type|trait|use)"
    r"[^\n;{}]*(?:HashMap|BTreeMap|DataMetadataMap)(?!Source|Entry|Visitor)"
)
KEYNOTE_SLIDE_MEDIA_DATA_OWNER_MARKER_GROUPS = {
    "selector resolution": ("SlideSelector", "MovieSelector", "select_media"),
    "media and metadata closure": (
        "keynote_media_codec",
        "package_metadata_media_codec",
        "DataInfo",
        "digest",
        "materialized_length",
        "owner",
    ),
    "bounded candidate publication": (
        "WireLimits",
        "preflight",
        "candidate",
        "reopen",
        "Verification",
        "locality",
    ),
    "exact reversible artifacts": (
        "ExactArtifacts",
        "PatchConflict",
        "inverse",
        "source_fingerprint",
    ),
    "preview invalidation": (
        "preview",
        "deleted_previews",
        "root_previews_absent",
    ),
}
# The focused owner keeps its transaction ledger and physical-closure proof in
# private child modules. Keep these paths explicit so a flat owner cannot
# silently absorb the budget/closure implementation again as the migration
# settles.
KEYNOTE_SLIDE_MEDIA_DATA_CHILD_ROOT = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_media_replacement"
)
KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE = (
    KEYNOTE_SLIDE_MEDIA_DATA_CHILD_ROOT / "budget.rs"
)
KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE = (
    KEYNOTE_SLIDE_MEDIA_DATA_CHILD_ROOT / "closure.rs"
)
KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?budget[ \t]*;"
)
KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?closure[ \t]*;"
)
KEYNOTE_SLIDE_MEDIA_DATA_PUBLIC_CHILD_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t]+mod[ \t]+"
    r"(?:r#)?(?:budget|closure)\b"
)
KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_TYPE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?struct[ \t]+"
    r"(?:r#)?MediaBudget\b"
)
KEYNOTE_SLIDE_MEDIA_DATA_PUBLIC_BUDGET_TYPE = re.compile(
    r"(?m)^[ \t]*pub[ \t]+struct[ \t]+(?:r#)?MediaBudget\b"
)
KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_FACTORY = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
    r"(?:r#)?for_package\b[^{};]*\([^{};]*\bpackage\b[^{};]*\)"
)
KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_ROOT_CALL = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*[ \t]*::[ \t]*)?"
    r"MediaBudget[ \t]*::[ \t]*for_package[ \t\r\n]*\("
)
KEYNOTE_SLIDE_MEDIA_DATA_MUTABLE_BUDGET_PARAMETER = re.compile(
    r"&[ \t]*mut[ \t]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t]*::[ \t]*)?"
    r"MediaBudget\b"
)
KEYNOTE_SLIDE_MEDIA_DATA_PREPARE_METADATA = re.compile(
    r"\bprepare_package_metadata_media_(?:rewrite|content_replacements?)[ \t\r\n]*\("
)
KEYNOTE_SLIDE_MEDIA_DATA_METADATA_EXECUTION_REQUIREMENTS = re.compile(
    r"\bexecution_requirements[ \t\r\n]*\("
)
KEYNOTE_SLIDE_MEDIA_DATA_PREPARED_EXECUTE = re.compile(
    r"\.[ \t\r\n]*execute[ \t\r\n]*\("
)
KEYNOTE_SLIDE_MEDIA_DATA_FOR_SOURCE_RESET = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)?"
    r"DecodeOptions[ \t\r\n]*::[ \t\r\n]*for_source[ \t\r\n]*\("
)
KEYNOTE_SLIDE_MEDIA_DATA_DIRECT_METADATA_REWRITE = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)?"
    r"rewrite_package_metadata_media_content_replacement[ \t\r\n]*\("
)
KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_HELPER = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
    r"(?:r#)?validate_selected_media_closure\b"
)
KEYNOTE_SLIDE_MEDIA_DATA_LIMIT_VARIANTS = frozenset(
    {
        "InputBytes",
        "OutputBytes",
        "Entries",
        "EntryBytes",
        "TotalBytes",
        "Slides",
        "References",
        "MediaBytes",
        "WireFields",
        "WireNesting",
        "WireWork",
        "Allocations",
    }
)

IWA_KEYNOTE_SLIDE_MEDIA_LIFECYCLE_RETIRED_TYPES = frozenset(
    {"RemovedKeynoteSlideMovie", "RemovedKeynoteSlideAudio"}
)

# The next lifecycle seam owns movie/audio duplication and removal in the
# focused Keynote package.  Keep this inventory separate from the existing
# replacement owner: replacement already crossed its host-retirement gate,
# while lifecycle graph mutation must remain dormant until its own package
# module is wired.  The codec/projection checks are independent because their
# neutral source can be reviewed before the package owner is activated.
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_media_lifecycle.rs"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_media_lifecycle"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_SOURCES = (
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "comment_graph.rs",
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "comment_removal.rs",
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "graph.rs",
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "identifier_watermark.rs",
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "metadata.rs",
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "budget.rs",
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "clone_payload.rs",
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / "node_cache.rs",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "media.rs"
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_FAMILY_MODULE = 'lifecycle'
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PACKAGE_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?slide_media_lifecycle[ \t]*;"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PUBLIC_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?![ \t]*\([ \t]*crate[ \t]*\))(?:\([^()]*\))?[ \t]+mod[ \t]+"
    r"(?:r#)?slide_media_lifecycle\b"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PUBLIC_CHILD_MODULE = re.compile(
    r"(?m)^[ \t]*(?P<visibility>pub(?:\([^()\n]*\))?)[ \t]+mod[ \t]+"
    r"(?P<name>(?:r#)?(?:comment_graph|comment_removal|graph|identifier_watermark|metadata|budget|clone_payload|node_cache))\b"
)
# These two adapters are intentionally visible to the containing ``package``
# module because the focused drawable-comment owner reuses their bounded
# budget and same-component comment graph.  The exact restricted visibility is
# part of the boundary: widening either module, or exposing any other lifecycle
# child, remains an audit failure.
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SHARED_CHILD_MODULES = frozenset(
    {"budget", "comment_graph"}
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SHARED_CHILD_VISIBILITY = "pub(in crate::package)"
# The lifecycle is a scoped family. ``Package`` starts the
# edit and replays its patch; the selector-first verbs stage on ``Edit``.
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PACKAGE_METHODS = frozenset(
    {
        "edit_slide_media_lifecycle",
        "apply_slide_media_lifecycle",
    }
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_TYPED_PACKAGE_METHODS = frozenset(
    {
        "duplicate",
        "remove",
        "duplicate_movie",
        "duplicate_audio",
        "remove_movie",
        "remove_audio",
    }
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PATCH_METHODS = frozenset({"apply_slide_media_lifecycle"})
# The functions that run a lifecycle transaction (one aggregate budget each):
# the staged edit's commit and the exact-source patch replay.
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OPERATIONS = frozenset({"commit", "apply_slide_media_lifecycle"})
# These names are intentionally accepted in both the short and prefixed
# forms.  Existing Keynote owners use the prefixed form, while the lifecycle
# module may keep a concise internal semantic spelling before re-exporting it.
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CANONICAL_TYPE_ALTERNATIVES = {
    "commit": frozenset({"LifecycleCommit", "Commit"}),
    "patch": frozenset({"LifecyclePatch", "Patch"}),
    "error": frozenset({"LifecycleError", "Error"}),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "MovieSelector"}
)
# These are semantic counters on the committed patch/diagnostics facade.  The
# generic vocabulary classifier quite reasonably treats ``objects`` as a
# native graph noun, but these accessors report counts and never expose an
# object or identifier value.
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SEMANTIC_IDENTIFIERS = frozenset(
    {"created_objects", "removed_objects"}
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "RawMessage",
        "SourceCatalog",
        "WireView",
    }
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_WIRE_TYPES = frozenset(
    {
        "BuildChunkLifecycleEdit",
        "BuildLifecycleEdit",
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "IdentifierRewrite",
        "Reference",
        "RewriteReport",
        "SlideLifecycleEdit",
        "SlideLifecycleSnapshot",
        "Uuid",
        "UuidRewrite",
    }
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:object|drawable|movie|media|data|poster|audio|component|"
    r"archive|message|resource|entry|metadata|package|build|chunk)[A-Za-z_]*"
    r"(?:id|identifier))[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|u32|usize|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)

# The direct-drawable comment route is a separate semantic owner from media
# lifecycle.  Keep its inventory explicit while the host still retains the
# compatibility CRUD surface: this ratchet reviews the new Package API and
# its private graph/transaction children, but intentionally performs no host
# retirement.  A future native/reply acceptance pass may add that retirement
# audit without weakening this source-ownership boundary.
KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_drawable_comments.rs"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_drawable_comments"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_SOURCES = (
    KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT / "engine.rs",
    KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT / "graph.rs",
    KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT / "metadata.rs",
    KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT / "metadata" / "identities.rs",
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "comment.rs"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_DRAWABLE_COMMENT_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "comment.rs"
KEYNOTE_SLIDE_DRAWABLE_COMMENT_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_DRAWABLE_COMMENT_PACKAGE_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub\(crate\)[ \t]+)?mod[ \t]+(?:r#)?slide_drawable_comments[ \t]*;"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_PUBLIC_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?![ \t]*\([ \t]*crate[ \t]*\))(?:\([^()]*\))?[ \t]+mod[ \t]+"
    r"(?:r#)?slide_drawable_comments\b"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_PUBLIC_CHILD_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?:\([^()]*\))?[ \t]+mod[ \t]+"
    r"(?:r#)?(?:engine|graph|metadata)\b"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_PRIVATE_IDENTITY_MODULE = re.compile(
    r"(?m)^[ \t]*mod[ \t]+(?:r#)?identities[ \t]*;"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_PUBLIC_IDENTITY_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?:\([^()]*\))?[ \t]+mod[ \t]+"
    r"(?:r#)?identities\b"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_PACKAGE_METHODS = frozenset(
    {
        "slide_drawables",
        "slide_drawable_comment",
        "slide_drawable_comment_replies",
        "edit_slide_drawable_comment",
        "apply_slide_drawable_comment",
    }
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_EDIT_METHODS = frozenset(
    {
        "set",
        "clear",
        "add_reply",
        "set_reply",
        "remove_reply",
        "commit",
    }
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_CANONICAL_TYPES = frozenset(
    {
        "DrawableSummary",
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_SEMANTIC_TYPES = frozenset(
    {"Comment", "CommentAuthor", "CommentTimestamp", "DrawableKind", "Reply"}
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "DrawableSelector", "ReplySelector", "Position"}
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "EntryEdit",
        "ExactArtifacts",
        "LifecycleBudget",
        "MessageInfo",
        "PhysicalSource",
        "RawMessage",
        "Selection",
        "SnappyStream",
        "SourceCatalog",
        "ThreadSnapshot",
        "WireLimits",
        "WireView",
    }
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_WIRE_TYPES = frozenset(
    {
        "AnnotationAuthorSnapshot",
        "AnnotationAuthorStorageSnapshot",
        "CommentStorageSnapshot",
        "DecodeError",
        "DecodeOptions",
        "DecodeReport",
        "EncodeError",
        "EncodeOptions",
        "RawMessage",
        "ReferenceRecord",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "UuidSnapshot",
    }
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|"
    r"[A-Za-z_]*(?:author|comment|drawable|message|object|reference|reply|"
    r"storage|thread|uuid)[A-Za-z_]*(?:id|identifier))[ \t\r\n]*"
    r":[ \t\r\n]*(?:u64|u32|usize|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_MARKER_GROUPS = {
    "selector-first graph": (
        "SlideSelector",
        "DrawableSelector",
        "select_drawable",
        "inventory_drawables",
    ),
    "semantic comment/reply operations": (
        "Comment",
        "Reply",
        "set_reply",
        "remove_reply",
        "add_reply",
        "clear",
    ),
    "shared bounded transaction ledger": (
        "LifecycleBudget",
        "budget_for",
        "LimitExceeded",
        "WireLimits",
        "try_reserve",
    ),
    "lazy source-preserving codec route": (
        "comment_storage_codec",
        "annotation_author_codec",
        "WireView",
        "unknown",
        "source",
    ),
    "exact candidate publication": (
        "ExactArtifacts",
        "candidate",
        "Verification",
        "inverse",
        "PatchConflict",
    ),
    "metadata author ownership": (
        "package_metadata_codec",
        "MetadataPlan",
        "AuthorRegistrationPlan",
        "GENERATED_AUTHOR_PUBLIC_ID",
    ),
}
KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_FIXTURE = Path(
    "test-data/iwork/keynote/drawable-comments-source-native.key"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_TEST_SOURCE = Path(
    "crates/litchi-keynote/tests/slide_drawable_comments.rs"
)
KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_TEST_NAME = "permanent_native_fixture_is_pinned"
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_media_lifecycle_codec.rs"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_CHILD_ROOT = Path(
    "crates/litchi-iwa-protos/src/keynote_media_lifecycle_codec"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE = (
    KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_CHILD_ROOT / "node_cache.rs"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_MODULE = "keynote_media_lifecycle_codec"
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE = Path(
    "crates/litchi-iwa-protos/src/buffa-projections/KNMediaLifecycleArchive.proto"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_APIS = (
    "DecodeError",
    "DecodeLimit",
    "DecodeOptions",
    "DecodeReport",
    "Uuid",
    "Reference",
    "UuidSnapshot",
    "SlideLifecycleSnapshot",
    "BuildLifecycleSnapshot",
    "BuildChunkLifecycleSnapshot",
    "IdentifierRewrite",
    "UuidRewrite",
    "SlideLifecycleEdit",
    "BuildLifecycleEdit",
    "BuildChunkLifecycleEdit",
    "RewriteReport",
    "PreparedSlideLifecycleRewrite",
    "decode_slide_lifecycle",
    "decode_slide_lifecycle_with_report",
    "decode_build",
    "decode_build_with_report",
    "decode_build_chunk",
    "decode_build_chunk_with_report",
    "prepare_slide_lifecycle_rewrite",
    "rewrite_slide_lifecycle",
    "rewrite_slide_lifecycle_with_report",
    "rewrite_build",
    "rewrite_build_with_report",
    "rewrite_build_chunk",
    "rewrite_build_chunk_with_report",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_MARKER_GROUPS = {
    "lazy Buffa ingress": (
        "buffa::DecodeOptions",
        "decode_lazy_view",
        "buffa_keynote_media_lifecycle_generated",
    ),
    "strict source walk": (
        "validate_input",
        ("canonical", "noncanonical"),
        ("unknown", "raw"),
    ),
    "finite resource accounting": (
        "max_message_bytes",
        "max_output_bytes",
        "max_fields",
        "max_work_bytes",
        "try_reserve",
    ),
    "source-preserving rewrite": (
        "prepare_slide_lifecycle_rewrite",
        "output_bytes",
        "readback",
        "source",
    ),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_APIS = (
    "SlideNodeBuildCacheSnapshot",
    "SlideNodeBuildCacheEdit",
    "PreparedSlideNodeBuildCacheRewrite",
    "decode_slide_node_build_cache",
    "decode_slide_node_build_cache_with_report",
    "prepare_slide_node_build_cache_rewrite",
    "rewrite_slide_node_build_cache",
    "rewrite_slide_node_build_cache_with_report",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_MARKER_GROUPS = {
    "lazy Buffa sidecar": (
        "projection::SlideNodeBuildCacheArchiveLazyView",
        "buffa()",
        "decode_lazy_view",
    ),
    "strict source walk": (
        "validate_input",
        "canonical_bool",
        "expected_values",
        "Presence",
        ("unknown", "raw"),
    ),
    "source-preserving rewrite": (
        "source",
        "readback",
        ("try_reserve", "reserve_output"),
        "output_bytes",
        ("noncanonical", "invalid"),
    ),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/comment_storage_codec/lifecycle.rs"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_PARENT_SOURCE = Path(
    "crates/litchi-iwa-protos/src/comment_storage_codec.rs"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_TEST_SOURCE = Path(
    "crates/litchi-iwa-protos/src/comment_storage_codec/lifecycle_tests.rs"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_APIS = (
    "CommentStorageLifecycleRewrite",
    "PreparedCommentStorageLifecycleRewrite",
    "prepare_comment_storage_lifecycle_rewrite",
    "rewrite_comment_storage_lifecycle",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_MARKER_GROUPS = {
    "source-witnessed UUID rewrite": (
        "expecting_fingerprint",
        "expecting_storage_uuid",
        "replacing_storage_uuid",
        "storage_uuid",
    ),
    "prepared execution accounting": (
        "prepare_report",
        "execution_requirements",
        "RewriteExecutionRequirements",
        "RewriteExecutionLimits",
        "execute",
    ),
    "streamed reply preservation": (
        "reply_remaps",
        "scan_comment_storage_raw",
        "emit_lifecycle_rewrite",
        "replies",
        "RawField",
    ),
    "bounded source rewrite": (
        "max_message_bytes",
        "output_bytes",
        "scratch_bytes",
        "retained_bytes",
        "try_reserve_exact",
    ),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_TEST_MARKERS = (
    "#[test]",
    "prepare_comment_storage_lifecycle_rewrite",
    "execution_requirements",
    "execute",
    "replacing_storage_uuid",
    "unknown",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_MESSAGES = {
    "Reference": {"identifier": 1, "deprecated_type": 2, "deprecated_is_external": 3},
    "UUID": {"lower": 1, "upper": 2},
    "BuildChunkIdentifierArchive": {"build_id": 1, "build_chunk_id": 2},
    "BuildArchive": {"drawable": 1},
    "BuildChunkArchive": {"build": 1, "build_chunk_identifier": 7, "build_id": 8},
    "SlideNodeBuildCacheArchive": {
        "build_event_count": 15,
        "build_event_count_cache_version": 26,
        "build_event_count_is_up_to_date": 22,
        "has_explicit_builds": 20,
        "has_explicit_builds_cache_version": 27,
        "has_explicit_builds_is_up_to_date": 23,
    },
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_BUDGET_MARKERS = (
    ("aggregate budget", re.compile(r"\b(?:Media|Lifecycle|Package|Transaction)[A-Za-z0-9_]*Budget\b")),
    ("budget factory", re.compile(r"\b(?:for_package|for_source|new)[ \t\r\n]*\(")),
    ("fallible allocation", re.compile(r"\b(?:try_reserve|checked_add|checked_mul|allocation)\b")),
    ("candidate validation", re.compile(r"\b(?:candidate|reopen|readback|verify|verification|locality)(?:\b|_)")),
    ("exact inverse", re.compile(r"\b(?:ExactArtifacts|source_fingerprint|inverse|PatchConflict)\b")),
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_BUDGET_FACTORY_CALL = re.compile(
    r"\b[A-Za-z_][A-Za-z0-9_]*Budget\s*::\s*"
    r"(?:for_package|for_source|new)\s*\("
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CLONE_PAYLOAD_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:[A-Za-z_][A-Za-z0-9_]*\s*::\s*)*"
    r"remap_clone_payload_with_budget\s*\("
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_LEGACY_CLONE_PAYLOAD_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:[A-Za-z_][A-Za-z0-9_]*\s*::\s*)*"
    r"remap_clone_payload\s*\("
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_HEADER_REWRITE_HELPER = re.compile(
    r"\breplace_slide_message_with_lifecycle_refs\b"
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CORE_HEADER_RESERVATION = re.compile(
    r"\breserve_core_header_work\s*\("
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_ALLOCATION_PLAN = re.compile(
    r"\bcharge_allocation_plan\s*\("
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_NODE_CACHE_APIS = (
    "prepare_slide_node_build_cache",
    "validate_candidate_node_cache",
    "validate_candidate_package_node_cache",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_NODE_CACHE_MARKER_GROUPS = {
    "neutral cache codec": (
        "lifecycle_codec::node_cache::",
        "decode_slide_node_build_cache_with_report",
        "SlideNodeBuildCacheEdit",
        ("prepare_slide_node_build_cache_rewrite", "rewrite_slide_node_build_cache_with_report"),
    ),
    "bounded publication": (
        "LifecycleBudget",
        "reserve_core_header_work",
        "charge_node_cache_decode_report",
        "charge_node_cache_rewrite_report",
    ),
    "unique supported message": (
        "unique_slide_node_message",
        "message_infos.len() != object.messages.len()",
        "message.type_ != SLIDE_NODE_MESSAGE_TYPE",
        "index.replace(position).is_some()",
    ),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_NODE_CACHE_REQUIRED_TYPES = (
    "SLIDE_NODE_MESSAGE_TYPE",
    "InvalidSource",
    "Verification",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_GRAPH_APIS = (
    "CommentStorageIdentity",
    "CommentAuthorDependency",
    "CommentStorageLocation",
    "CommentGraphPlan",
    "CommentGraphPayloadPolicy",
    "CommentGraphComponentPolicy",
    "plan_comment_graph",
    "plan_comment_graph_cross_component",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_GRAPH_MARKER_GROUPS = {
    "strict versus package-scope policy": (
        "CommentGraphPayloadPolicy",
        "Strict",
        "PreserveRootExtensions",
        "CommentGraphComponentPolicy",
        "SameComponent",
        "Package",
        "plan_comment_graph_cross_component",
        "reject_unknown_root_fields",
    ),
    "per-node storage locations": (
        "CommentStorageLocation",
        "storage_locations",
        "storage_component",
    ),
    "strict comment-storage visitor": (
        "comment_storage_codec",
        "decode_comment_storage_archive_with_visitor",
        "CommentStorageVisitor",
        "ReplyCollector",
    ),
    "same-component closure": (
        "storage_ids",
        "storage_identities",
        "author_ids",
        "author_dependencies",
        "binary_search",
        "InvalidSource",
    ),
    "storage identity and UUID": (
        "CommentStorageIdentity",
        "storage_uuid",
        "root_storage_uuid",
        "uuid",
    ),
    "bounded graph budget": (
        "LifecycleBudget",
        "charge_entries",
        "charge_references",
        "charge_allocations",
        "try_reserve",
    ),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_REMOVAL_APIS = (
    "CommentRemovalPlan",
    "plan_comment_removal",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_REMOVAL_MARKER_GROUPS = {
    "same-component source admission": (
        "component_name",
        "selected_object_ids",
        "comment_graph",
        "archive_limits",
        "InvalidSource",
    ),
    "strict shared-budget census": (
        "LifecycleBudget",
        "reserve_core_header_inspection_work",
        "strict_reference_census",
        "budget.charge_entries",
        ("budget.charge_references", "budget.charge_wire_fields"),
        "budget.charge_wire_work",
        "budget.charge_allocations",
    ),
    "source-bound comment retention": (
        "retained_comment_storage_ids",
        "source_component",
        "comment_edges",
        "selected_comment_author_edges",
        "pending",
        "processed",
        "lower_bound_source",
        "insert_sorted_unique",
    ),
    "author retention diagnostics": (
        "author_dependencies",
        "external_author_ids",
        "unused_external_author_ids",
        "removed_object_ids",
    ),
    "known payload admission": (
        "validate_known_payload_relationships",
        "reserve_core_header_inspection_work",
        "strict_reference_census",
    ),
    "movie comment payload/header relationship": (
        "MOVIE_MESSAGE_TYPE",
        "direct_drawable_comment",
        "validate_movie_payload_relationship",
        "object_references",
        "count",
        "comment_storage_ids",
    ),
    "comment storage payload/header relationship": (
        "COMMENT_STORAGE_MESSAGE_TYPE",
        "MessageInfo",
        "message.data",
        ("comment_storage_codec", "super::comment_graph"),
        "author",
        "validate_comment_storage_payload_relationship",
    ),
    "known comment target resolution": (
        "validate_comment_storage_target",
        "validate_annotation_author_target",
        "validate_single_message_type",
        "object_with_component",
        "actual_component",
        "component_name",
    ),
    "exact target message shape": (
        "archive_info.identifier",
        "messages.len()",
        "message_infos.len()",
        "message.type_",
        "info.type_",
        "info.length",
        "message.data.len()",
    ),
    "neutral storage facts target closure": (
        "let facts",
        "facts.author_identifier",
        "facts.reply_identifiers",
        "validate_comment_storage_relationship",
    ),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_IDENTIFIER_WATERMARK_APIS = (
    "plan_release",
    "verify_release",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_IDENTIFIER_WATERMARK_MARKER_GROUPS = {
    "source-bounded release plan": (
        "source",
        "removed_ids",
        "expected_last_identifier",
        "LifecycleBudget",
        "removed_ids.binary_search",
        "source.state.source.components",
        "component.archive().objects",
        "budget.charge_entries",
        "budget.charge_wire_work",
        "maximum_remaining",
        "Ok(Some(maximum_remaining))",
    ),
    "candidate release verification": (
        "candidate",
        "new_last_identifier",
        "candidate.state.source.components",
        "component.archive().objects",
        "budget.charge_entries",
        "budget.charge_wire_work",
        "maximum_remaining != new_last_identifier",
        "Error::Verification",
    ),
}
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_METADATA_IDENTIFIER_MARKERS = (
    "metadata_identifier_maximum",
    "fn observe_identifier",
    "self.metadata_identifier_maximum",
    ".max(identifier)",
)
KEYNOTE_SLIDE_MEDIA_LIFECYCLE_METADATA_IDENTIFIER_CALLBACKS = {
    "component identifier observation": re.compile(
        r"\bfn\s+visit_component\b[\s\S]{0,900}?"
        r"\bobserve_identifier\s*\(\s*component\.identifier\s*\(\s*\)\s*\)"
    ),
    "object UUID identifier observation": re.compile(
        r"\bfn\s+visit_object_uuid\b[\s\S]{0,900}?"
        r"\bobserve_identifier\s*\(\s*binding\.object_identifier\s*\(\s*\)\s*\)"
    ),
    "external object identifier observation": re.compile(
        r"\bfn\s+visit_external_reference\b[\s\S]{0,1200}?"
        r"\bobserve_identifier\s*\(\s*identifier\s*\)"
    ),
    "owner object identifier observation": re.compile(
        r"\bfn\s+visit_data_reference_owner\b[\s\S]{0,900}?"
        r"\bobserve_identifier\s*\(\s*owner\.object_identifier\s*\(\s*\)\s*\)"
    ),
    "ambiguous identifier observation": re.compile(
        r"\bfn\s+visit_ambiguous_object_identifier\b[\s\S]{0,900}?"
        r"\bobserve_identifier\s*\(\s*identifier\s*\)"
    ),
    "metadata-map object identifier observation": re.compile(
        r"\bfn\s+visit_data_metadata_map\b[\s\S]{0,900}?"
        r"\bobserve_identifier\s*\(\s*object_identifier\s*\)"
    ),
}

# Existing canonical Keynote slide-table title settings live behind a
# selector-first package facade.  The strict Buffa projections remain shared
# iWork codecs; this ratchet owns only the semantic facade and the bounded
# Keynote transaction.
KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "title.rs"
)
KEYNOTE_SLIDE_TABLE_TITLE_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table.rs"
KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_title.rs"
)
KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_TITLE_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "title.rs"
KEYNOTE_SLIDE_TABLE_TITLE_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_TITLE_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_TYPES = frozenset({"Settings"})
KEYNOTE_SLIDE_TABLE_TITLE_SELECTOR_TYPES = frozenset({"SlideSelector", "TableSelector"})
KEYNOTE_SLIDE_TABLE_TITLE_PACKAGE_METHODS = frozenset(
    {
        "slide_table_title_settings",
        "edit_slide_table_title",
        "apply_slide_table_title",
    }
)
KEYNOTE_SLIDE_TABLE_TITLE_EDIT_METHODS = frozenset({"set", "commit"})
KEYNOTE_SLIDE_TABLE_TITLE_FLAT_ALIASES = frozenset(
    {
        "TableTitleCommit",
        "TableTitleDiagnostics",
        "TableTitleEdit",
        "TableTitleError",
        "TableTitleLimitKind",
        "TableTitlePatch",
        "TitleCommit",
        "TitleDiagnostics",
        "TitleEdit",
        "TitleError",
        "TitleLimitKind",
        "TitlePatch",
    }
)
KEYNOTE_SLIDE_TABLE_TITLE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_SLIDE_TABLE_TITLE_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_SLIDE_TABLE_TITLE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tsp", "tsd", "tst", "litchi_iwa_protos"}
)
KEYNOTE_SLIDE_TABLE_TITLE_PACKAGE_MARKER_GROUPS = {
    "aggregate transaction budget": ("TitleBudget",),
    "strict Buffa projections": ("numbers_table_title_codec", "table_info_codec"),
    "raw-preserving rewrite": (
        "patch_nested_fields_batched",
        "NestedFieldReplacement",
    ),
    "exact artifacts and inverse": ("ExactArtifacts", "inverse"),
    "prepared reassembly": ("prepare_reassembly_with_deletions",),
    "preview invalidation": ("root_preview_deletions",),
    "candidate reopen": ("candidate.validate",),
    "object/member locality": ("verify_locality",),
}

# Keynote's physical row-reordering executor stays in the legacy host; only the
# persisted field-44 sort configuration lives behind a selector-first package
# facade. The compatibility host's normal physical entry points are likewise
# selector-first; its older raw-ID methods remain declarations only for source
# compatibility. This is deliberately a separate ratchet from the title owner
# above: a table sort edit must not accidentally become a row/tile/UID,
# metadata, or focused physical ``Sort Now`` migration.
KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "sort.rs"
)
KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table.rs"
KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_sort_order.rs"
)
KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_SORT_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "sort.rs"
KEYNOTE_SLIDE_TABLE_SORT_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_SORT_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
        "Path",
    }
)
KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_TYPES = frozenset(
    {"ColumnIndex", "Direction", "Order", "RowRange", "Rule", "Scope"}
)
KEYNOTE_SLIDE_TABLE_SORT_TRANSACTION_TYPES = frozenset(
    {"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"}
)
KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "TableSelector"}
)
KEYNOTE_SLIDE_TABLE_SORT_PACKAGE_METHODS = frozenset(
    {
        "slide_table_sort_order",
        "edit_slide_table_sort_order",
        "apply_slide_table_sort_order",
    }
)
KEYNOTE_SLIDE_TABLE_SORT_EDIT_METHODS = frozenset(
    {"order", "set", "clear", "reset", "commit"}
)
KEYNOTE_SLIDE_TABLE_SORT_FLAT_ALIASES = frozenset(
    {
        "SortCommit",
        "SortDiagnostics",
        "SortEdit",
        "SortError",
        "SortLimitKind",
        "SortPatch",
        "SortPath",
        "SortOrder",
        "SortOrderCommit",
        "SortOrderDiagnostics",
        "SortOrderEdit",
        "SortOrderError",
        "SortOrderLimitKind",
        "SortOrderPatch",
        "SortOrderPath",
        "TableSortCommit",
        "TableSortDiagnostics",
        "TableSortEdit",
        "TableSortError",
        "TableSortLimitKind",
        "TableSortOrder",
        "TableSortOrderCommit",
        "TableSortOrderDiagnostics",
        "TableSortOrderEdit",
        "TableSortOrderError",
        "TableSortOrderLimitKind",
        "TableSortOrderPatch",
        "TableSortOrderPath",
    }
)
KEYNOTE_SLIDE_TABLE_SORT_PHYSICAL_TYPES = frozenset(
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
        "TableInfoArchive",
        "TableModelArchive",
        "TableSortOrderArchive",
        "TableSortOrderSnapshot",
    }
)
KEYNOTE_SLIDE_TABLE_SORT_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_SLIDE_TABLE_SORT_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tst", "tsp", "tsd"}
)
KEYNOTE_SLIDE_TABLE_SORT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_sort_order_codec.rs"
)
KEYNOTE_SLIDE_TABLE_SORT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_SLIDE_TABLE_SORT_CODEC_MODULES = (
    "numbers_table_sort_order_codec",
    "table_sort_order_codec",
)
KEYNOTE_SLIDE_TABLE_SORT_CODEC_REQUIRED_APIS = (
    "SortOrderSnapshot",
    "decode_table_model_sort_order_with_report",
    "prepare_table_model_sort_order_rewrite",
    "PreparedTableSortOrderRewrite",
    "PreparedTableModelSortOrderRewrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "canonical_table_sort_order",
    "rewrite_table_model_sort_order",
)
KEYNOTE_SLIDE_TABLE_SORT_SHARED_CORE_APIS = (
    "select_table",
    "same_target",
    "verify_locality",
)
KEYNOTE_SLIDE_TABLE_SORT_PACKAGE_MARKER_GROUPS = {
    "aggregate transaction budget": (
        "Budget",
        "residual",
        "budget",
    ),
    "strict neutral codec": (
        "sort_order_codec",
        "decode_table_model_sort_order_with_report",
        "prepare_table_model_sort_order_rewrite",
    ),
    "prepared codec requirements": (
        "execution_requirements",
        "report()",
        "execute",
    ),
    "field-44/field-45 authority": (
        "TABLE_MODEL_MESSAGE_TYPE",
        "sort_order",
    ),
    "raw-preserving rewrite": (
        "replace_message_preserving_header",
    ),
    "exact artifacts and inverse": (
        "ExactArtifacts",
        "inverse",
        "is_noop",
    ),
    "prepared reassembly": ("prepare_reassembly",),
    "candidate reopen": ("candidate", "reopen", "validate"),
    "object/member locality": ("verify_locality", "same_target"),
    "preview preservation": (
        "deleted_previews",
    ),
}

# The persisted field-44 owner above is deliberately not the physical
# "Sort Now" owner. Keep the eventual hand-off behind an explicit,
# production-only activation token so a partial implementation cannot silently
# retire the compatibility executor. The physical owner is intentionally
# stricter than the persisted owner: it must prove every row-affine structure
# it admits before moving a single tile row.
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "physical_sort.rs"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_physical_sort.rs"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "physical_sort.rs"
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
        "Path",
    }
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "TableSelector"}
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PACKAGE_METHODS = frozenset(
    {
        "execute_slide_table_sort_order",
        "execute_slide_table_sort_order_to_rows",
        "apply_slide_table_physical_sort",
    }
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EDIT_METHODS = frozenset(
    {"rows", "order", "commit"}
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FLAT_ALIASES = frozenset(
    {
        "PhysicalSortCommit",
        "PhysicalSortDiagnostics",
        "PhysicalSortEdit",
        "PhysicalSortError",
        "PhysicalSortLimitKind",
        "PhysicalSortPatch",
        "PhysicalSortPath",
        "TablePhysicalSortCommit",
        "TablePhysicalSortDiagnostics",
        "TablePhysicalSortEdit",
        "TablePhysicalSortError",
        "TablePhysicalSortLimitKind",
        "TablePhysicalSortPatch",
        "TablePhysicalSortPath",
    }
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ColumnRowUIDMapArchive",
        "ComponentCatalog",
        "DataStoreArchive",
        "EntryEdit",
        "ExactArtifacts",
        "HeaderStorageArchive",
        "HeaderStorageBucketArchive",
        "HiddenStatesArchive",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "Resolved",
        "SnappyStream",
        "SourceCatalog",
        "StrokeSidecarArchive",
        "TableDataListArchive",
        "TableInfoArchive",
        "TableModelArchive",
        "TileArchive",
        "TileRowInfoArchive",
        "TileStorageArchive",
    }
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_WIRE_TYPES = frozenset(
    {
        # The hidden codec's borrowed snapshots/plans are implementation
        # values too. Listing them explicitly prevents a public re-export
        # from smuggling low-level objects through an otherwise innocuous
        # package alias; the stable facade exposes only the transaction
        # wrappers below the package boundary.
        "BncCellView",
        "BuffaDecodeOptions",
        "Budget",
        "ColumnRowUidMapSnapshot",
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "Field",
        "FieldSpan",
        "HeaderRecord",
        "HeaderSnapshot",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "ParseItem",
        "PreparedColumnRowUidMapRewrite",
        "PreparedHeaderRowsRewrite",
        "PreparedTableModelPhysicalSortRewrite",
        "PreparedTablePhysicalSortRewrite",
        "PreparedTileRowsRewrite",
        "RawWireView",
        "RewriteReport",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteRequirements",
        "RowMove",
        "RowUidPermutation",
        "StoredValue",
        "TileRowInfoSnapshot",
        "TileRowRecord",
        "TileSnapshot",
        "PhysicalTableSortSnapshot",
        "LitchiIwaTablePhysicalSortProjection",
        "TablePhysicalSortSnapshot",
        "Uuid",
        "Value",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PROTO_ORIGINS = frozenset(
    {
        "Buffa",
        "buffa",
        "Prost",
        "ProstTypes",
        "prost",
        "prost_types",
        "kn",
        "tst",
        "tsp",
        "tsd",
        "litchi_iwa_protos",
    }
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_ACTIVATION = re.compile(
    r"(?<![A-Za-z0-9_])KEYNOTE_PHYSICAL_SORT_OWNER_ACTIVE(?![A-Za-z0-9_])"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_physical_sort_codec.rs"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_MODULES = (
    "numbers_table_physical_sort_codec",
    "table_physical_sort_codec",
)
# Keep codec API checks tolerant of a neutral table_* alias while still
# requiring the prepared, lazy rewrite lifecycle. The first spelling is the
# canonical one used by the focused owner and by the boundary fixtures.
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_REQUIRED_APIS = {
    "snapshot": (
        "TileSnapshot",
        "HeaderSnapshot",
        "ColumnRowUidMapSnapshot",
        "PhysicalTableSortSnapshot",
        "TablePhysicalSortSnapshot",
    ),
    "uid-map decode": ("decode_column_row_uid_map",),
    "tile-row plan": ("plan_tile_rows_rewrite",),
    "header-row plan": ("plan_header_storage_bucket_rows",),
    "uid-map plan": ("plan_column_row_uid_map_rewrite",),
    "tile-row execute": ("execute_tile_rows_rewrite",),
    "header-row execute": ("execute_header_storage_bucket_rows",),
    "uid-map execute": ("execute_column_row_uid_map_rewrite",),
    "decode": (
        "decode_tile_for_rewrite",
        "decode_header_bucket_for_rewrite",
        "decode_table_physical_sort_with_report",
        "decode_table_model_physical_sort_with_report",
    ),
    "prepared rewrite": (
        "plan_tile_rows_rewrite",
        "plan_header_storage_bucket_rows",
        "plan_column_row_uid_map_rewrite",
        "prepare_table_physical_sort_rewrite",
        "prepare_table_model_physical_sort_rewrite",
    ),
    "prepared type": (
        "PreparedTileRowsRewrite",
        "PreparedHeaderRowsRewrite",
        "PreparedColumnRowUidMapRewrite",
        "PreparedTablePhysicalSortRewrite",
        "PreparedTableModelPhysicalSortRewrite",
    ),
    "execution requirements": (
        "RewriteExecutionRequirements",
        "RewriteRequirements",
    ),
    "execution limits": ("RewriteExecutionLimits", "DecodeOptions"),
}
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_FLOW_MARKERS = {
    "borrowed/lazy projection": re.compile(
        r"\b(?:Buffa|buffa|WireView|BncCellView|decode_lazy_view|lazy|borrowed)\b"
    ),
    "strict admission": re.compile(
        r"\b(?:preflight|validate|strict|reject|Unsupported)\w*\b"
    ),
    "prepared execution": re.compile(
        r"\b(?:execution_requirements|requirements|exact_limits|"
        r"execute(?:_[A-Za-z0-9_]+)?)\s*\("
    ),
    "raw-preserving rewrite": re.compile(
        r"\b(?:replace_message_preserving|NestedFieldReplacement|"
        r"preserv(?:e|ing)_unknown|unknown_fields|replace_nested_varint|"
        r"assemble_(?:tile|header|uid)|FieldSpan)\w*\b"
    ),
}
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SHARED_EXECUTOR_SOURCE = Path(
    "crates/litchi-iwa-table/src/sort.rs"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SHARED_EXECUTOR_MARKERS = (
    re.compile(r"\bpub\s+fn\s+execute\b"),
    re.compile(r"\bsame_tile_only\b"),
    re.compile(r"\bTILE_ROWS_FIELD\b"),
    re.compile(r"\bBUCKET_HEADERS_FIELD\b"),
    re.compile(r"\bUID_ROW_UID_FOR_INDEX_FIELD\b"),
    re.compile(r"\bworking\.publish\s*\("),
)
# Owner markers that remain once the rewrite is delegated: the
# owner still verifies the reopened candidate and checks row and header
# placement against its own scan.
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_DELEGATED_GROUPS = {
    "candidate verification": (
        re.compile(r"\b(?:candidate|reopen)\b"),
        re.compile(r"\b(?:validate|verify)\w*\s*\("),
    ),
    "tile row affinity": (
        re.compile(r"\b(?:tile_row_index|numrows)\b"),
        re.compile(r"\bdestination_by_source\b"),
    ),
    "sparse header affinity": (
        re.compile(
            r"\b(?:row_header_indices|header_index|header_indices|sparse_headers)\b"
        ),
    ),
}
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_MARKER_GROUPS = {
    "selector-first exact admission": (
        re.compile(r"\bselect_table\s*\("),
        re.compile(r"\b(?:SlideSelector|TableSelector)\b"),
        re.compile(r"\b(?:exact_source|source_fingerprint|same_target)\b"),
    ),
    "shared bounded budget": (
        re.compile(r"\b(?:Budget|TransactionBudget|PhysicalSortBudget)\b"),
        re.compile(r"\b(?:try_reserve|reserve_exact|charge_|max_rows)\w*\b"),
        re.compile(r"\b(?:PhysicalRows|PhysicalColumns|KeyCells)\b"),
    ),
    "exact patch and inverse": (
        re.compile(r"\bExactArtifacts\b"),
        re.compile(r"\binverse\s*\("),
        re.compile(r"\bPatchConflict\b"),
        re.compile(r"\b(?:source|target)_fingerprint\b"),
    ),
    "candidate verification": (
        re.compile(r"\b(?:candidate|reopen)\b"),
        re.compile(r"\b(?:validate|verify)\w*\s*\("),
        re.compile(r"\b(?:prepare_reassembly|reassembly)\b"),
    ),
    "exact locality allowlist": (
        # The physical owner uses a distinct helper because its allowlist
        # admits several changed component members.  Keep the generic helper
        # spelling accepted for synthetic boundary fixtures while recognizing
        # the canonical physical spelling as well.
        re.compile(r"\bverify_(?:physical_)?locality\s*\("),
        re.compile(
            r"\b(?:LocalityAllowlist|locality|allowlist|allowed_members|"
            r"touched_members|touched_components)\w*\b"
        ),
        re.compile(
            r"\b(?:with_model_component|changed_components|same_target|"
            r"exact_target)\b"
        ),
    ),
    "preview invalidation/preservation": (
        # Changed physical rows invalidate the root previews, while an exact
        # no-op must report that no preview was deleted. Require both sides of
        # that contract in the owner rather than accepting a stale zero-only
        # diagnostic.
        re.compile(r"\broot_preview_deletions\b"),
        re.compile(r"\bdeleted_previews\b"),
    ),
    "tile row affinity": (
        re.compile(r"\b(?:TileStorage|TileArchive|TileRowInfo(?:Snapshot)?)\b"),
        re.compile(r"\b(?:rowInfos|row_infos|tile_row_index|numrows)\b"),
        re.compile(r"\b(?:row_coverage|row_permutation|permutation)\w*\b"),
    ),
    "sparse header affinity": (
        re.compile(
            r"\b(?:HeaderStorage|HeaderStorageBucket|Header(?:Snapshot)?|"
            r"RowMove)\b"
        ),
        re.compile(
            r"\b(?:row_header_indices|header_index|header_indices|sparse_headers)\b"
        ),
        re.compile(r"\b(?:bucket|bucket_index)\w*\b"),
    ),
    "row UID bijection": (
        re.compile(
            r"\b(?:ColumnRowUIDMap|row_uid|uid_for_index|index_for_uid)\w*\b"
        ),
        re.compile(r"\b(?:bijection|inverse_consistent|uid_map)\w*\b"),
        re.compile(r"\b(?:row_count|number_of_rows|body_rows)\b"),
    ),
    "border/stroke affinity": (
        # Stroke sidecars are validated by the shared physical admission
        # routine before the owner receives its topology.  Accept that
        # canonical delegation alongside the historical explicit sidecar
        # vocabulary; the remaining markers still require the topology-based
        # locality fence and mutable-component census in the owner.
        re.compile(
            r"\b(?:StrokeSidecar|StrokeLayer|StrokeRun|stroke_sidecar|"
            r"admit_physical_table)\b"
        ),
        re.compile(
            r"\b(?:row_layer|top_row|bottom_row|stroke_rows|"
            r"ensure_component_local|with_physical_topology)\w*\b"
        ),
        re.compile(
            r"\b(?:run|runs|candidate\.component|component\.as_ref|"
            r"mutable_component_count)\w*\b"
        ),
    ),
    "hidden-axis affinity": (
        re.compile(
            r"\b(?:HiddenStates|hidden_states|user_hidden|hiding_state)\w*\b"
        ),
        re.compile(r"\b(?:row_uid|column_uid|positional)\w*\b"),
        re.compile(r"\b(?:restore|snapshot|hidden)\w*\b"),
    ),
    "explicit dependency refusals": (
        re.compile(r"\b(?:UnsupportedDependency|unsupported_dependency)\b"),
        re.compile(
            r"\b(?:formula|formula_error|merge|conditional|filtered|"
            r"pivot|spill|group)\w*\b"
        ),
        re.compile(r"\b(?:reject|refuse|unsupported)\w*\b"),
    ),
}
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FORBIDDEN_OWNER_PATTERNS = (
    (
        "eager Prost/generated decode",
        re.compile(
            r"\bprost\s*::\s*Message\b|"
            r"\b(?:kn|tst|tsp|tsd)\s*::\s*[A-Za-z_][A-Za-z0-9_]*\s*::\s*decode\s*\("
        ),
    ),
    (
        "direct legacy Numbers writer",
        re.compile(
            r"(?<![A-Za-z0-9_])(?:crate\s*::\s*numbers\s*::\s*editor\s*::\s*)?"
            r"apply_table_sort_order(?:_to_rows)?_in_package\s*\("
        ),
    ),
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_LEGACY_WRITER_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:crate\s*::\s*numbers\s*::\s*editor\s*::\s*)?"
    r"apply_table_sort_order(?:_to_rows)?_in_package\s*\("
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_LEGACY_WRITER_IMPORT = re.compile(
    r"(?m)\buse\s+[^;\n]*\b(?:numbers\s*::\s*editor\s*::\s*)?"
    r"apply_table_sort_order(?:_to_rows)?_in_package\b\s+as\s+"
    r"(?P<alias>[A-Za-z_][A-Za-z0-9_]*)"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PUBLIC_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes|"
    r"raw_message|raw_bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_TEST_SOURCES = (
    Path("crates/litchi-keynote/tests/slide_table_physical_sort.rs"),
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FUZZ_SOURCES = (
    Path(
        "crates/litchi-iwa-protos/fuzz/fuzz_targets/keynote_table_physical_sort_codec.rs"
    ),
    Path("crates/litchi/fuzz/fuzz_targets/keynote_slide_table_physical_sort.rs"),
)
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FUZZ_CORPORA = (
    Path("crates/litchi-iwa-protos/fuzz/corpus/keynote_table_physical_sort_codec"),
    Path("crates/litchi/fuzz/corpus/keynote_slide_table_physical_sort"),
)
# A short-lived neutral spelling was used by the first physical-sort handoff.
# Accept it only as an explicit compatibility location while the focused
# Keynote target remains the canonical evidence path above. This keeps a
# partially migrated checkout auditable without weakening the requirement for
# two independent fuzz harnesses and non-empty corpora.
KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FUZZ_CODEC_ALTERNATES = (
    (
        Path(
            "crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_physical_sort_codec.rs"
        ),
        Path("crates/litchi-iwa-protos/fuzz/corpus/numbers_table_physical_sort_codec"),
    ),
)
KEYNOTE_SLIDE_TABLE_SORT_FORBIDDEN_OWNER_PATTERNS = (
    (
        "local SortBudget",
        re.compile(r"\bSortBudget\b"),
    ),
    (
        "local ArchiveReferenceVisitor",
        re.compile(r"\bArchiveReferenceVisitor\b"),
    ),
    (
        "residual limit grant via .max(1)",
        re.compile(
            r"(?:\bsaturating_sub\s*\([^;{}]*?\)|"
            r"\b(?:residual|remaining_[A-Za-z0-9_]*)\b[^;{}]*?|"
            r"\bfn\s+(?:residual|remaining_[A-Za-z0-9_]*)[A-Za-z0-9_]*\b"
            r"[^{}]*\{[^{}]*?)"
            r"\s*\.\s*max\s*\(\s*1\s*,?\s*\)"
        ),
    ),
    (
        "saturating reference counter",
        re.compile(
            r"\b(?:self\s*\.\s*)?(?:"
            r"[A-Za-z_][A-Za-z0-9_]*_(?:reference|references|edge|edges|census)"
            r"|(?:reference|references|edge|edges|census))"
            r"[A-Za-z0-9_]*"
            r"\s*=\s*[^;{}]*?\.\s*saturating_add\s*\("
        ),
    ),
    (
        "direct table_info_codec graph selection",
        re.compile(r"\btable_info_codec\b"),
    ),
)
KEYNOTE_SLIDE_TABLE_SORT_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:slide_table_sort_order|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?sort)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
KEYNOTE_SLIDE_TABLE_SORT_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)

# Persisted Keynote slide-table lock state is exposed through the same
# selector-first package boundary as title and sort.  The native editor lock
# module is retired.
KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "lock.rs"
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table.rs"
KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_lock_state.rs"
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_LOCK_STATE_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "lock.rs"
KEYNOTE_SLIDE_TABLE_LOCK_STATE_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_LOCK_STATE_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
        "Path",
    }
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_TYPES = frozenset({"State"})
KEYNOTE_SLIDE_TABLE_LOCK_STATE_TRANSACTION_TYPES = frozenset(
    {"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"}
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "TableSelector"}
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_PACKAGE_METHODS = frozenset(
    {
        "slide_table_lock_state",
        "edit_slide_table_lock_state",
        "apply_slide_table_lock_state",
    }
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_EDIT_METHODS = frozenset(
    {"before", "after", "state", "set", "set_state", "lock", "unlock", "commit"}
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_FLAT_ALIASES = frozenset(
    {
        "TableLockStateCommit",
        "TableLockStateDiagnostics",
        "TableLockStateEdit",
        "TableLockStateError",
        "TableLockStateLimitKind",
        "TableLockStatePatch",
        "TableLockStatePath",
        "LockStateCommit",
        "LockStateDiagnostics",
        "LockStateEdit",
        "LockStateError",
        "LockStateLimitKind",
        "LockStatePatch",
        "LockStatePath",
    }
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_SLIDE_TABLE_LOCK_STATE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tst", "tsp", "tsd"}
)
KEYNOTE_SLIDE_TABLE_LOCK_STATE_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)

# Persisted Keynote slide-table row/column dimensions live behind a
# selector-first package owner. Keep this ratchet separate from the physical
# geometry executor: persisted dimension settings must not retain raw native
# identifiers, generated archives, or a public IWA editor surface.
KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "dimension.rs"
)
KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table.rs"
KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_dimension.rs"
)
KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_DIMENSION_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "dimension.rs"
KEYNOTE_SLIDE_TABLE_DIMENSION_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_DIMENSION_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
        "Path",
    }
)
KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_TYPES = frozenset(
    {"Dimension", "Error", "Points", "Size"}
)
KEYNOTE_SLIDE_TABLE_DIMENSION_TRANSACTION_TYPES = frozenset(
    {"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"}
)
KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "TableSelector"}
)
KEYNOTE_SLIDE_TABLE_DIMENSION_PACKAGE_METHODS = frozenset(
    {
        "slide_table_dimension_size",
        "edit_slide_table_dimension_size",
        "apply_slide_table_dimension_size",
    }
)
KEYNOTE_SLIDE_TABLE_DIMENSION_EDIT_METHODS = frozenset(
    {"size", "set", "set_points", "reset", "commit"}
)
KEYNOTE_SLIDE_TABLE_DIMENSION_FLAT_ALIASES = frozenset(
    {
        "DimensionCommit",
        "DimensionDiagnostics",
        "DimensionEdit",
        "DimensionError",
        "DimensionLimitKind",
        "DimensionPatch",
        "DimensionPath",
        "DimensionSizeCommit",
        "DimensionSizeDiagnostics",
        "DimensionSizeEdit",
        "DimensionSizeError",
        "DimensionSizeLimitKind",
        "DimensionSizePatch",
        "DimensionSizePath",
        "TableDimensionCommit",
        "TableDimensionDiagnostics",
        "TableDimensionEdit",
        "TableDimensionError",
        "TableDimensionLimitKind",
        "TableDimensionPatch",
        "TableDimensionPath",
        "TableDimensionSizeCommit",
        "TableDimensionSizeDiagnostics",
        "TableDimensionSizeEdit",
        "TableDimensionSizeError",
        "TableDimensionSizeLimitKind",
        "TableDimensionSizePatch",
        "TableDimensionSizePath",
    }
)
KEYNOTE_SLIDE_TABLE_DIMENSION_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ArchiveReferenceKind",
        "ArchiveReferenceOccurrence",
        "ArchiveReferencePolicy",
        "ArchiveReferenceScope",
        "ArchiveReferenceVisitor",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "FieldType",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "TableDimensionArchive",
        "TableDimensionSnapshot",
        "TableInfoArchive",
        "TableModelArchive",
        "TableModelSnapshot",
    }
)
KEYNOTE_SLIDE_TABLE_DIMENSION_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_SLIDE_TABLE_DIMENSION_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tst", "tsp", "tsd", "tswp"}
)
KEYNOTE_SLIDE_TABLE_DIMENSION_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:slide_table_dimension|table[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?dimension)(?=[ \t\r\n]*(?:::|as\b|;|=))"
)

# Persisted Keynote slide-table display names live behind the same
# selector-first package boundary as the adjacent table configuration owners.
# A name is semantic UTF-8 text; its owner may not expose the native model
# identifier, archive/protobuf values, or a public raw-wire escape hatch.
KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "name.rs"
)
KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table.rs"
KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_name.rs"
)
KEYNOTE_SLIDE_TABLE_CORE_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_core.rs"
)
KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_NAME_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "name.rs"
KEYNOTE_SLIDE_TABLE_NAME_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_NAME_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
        "Path",
    }
)
KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_TYPES = frozenset({"Error", "Name"})
KEYNOTE_SLIDE_TABLE_NAME_TRANSACTION_TYPES = frozenset(
    {"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"}
)
KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "TableSelector"}
)
KEYNOTE_SLIDE_TABLE_NAME_PACKAGE_METHODS = frozenset(
    {
        "slide_table_name",
        "edit_slide_table_name",
        "apply_slide_table_name",
    }
)
KEYNOTE_SLIDE_TABLE_NAME_EDIT_METHODS = frozenset(
    {"path", "before", "name", "set", "set_name", "commit"}
)
KEYNOTE_SLIDE_TABLE_NAME_FLAT_ALIASES = frozenset(
    {
        "TableName",
        "TableNameCommit",
        "TableNameDiagnostics",
        "TableNameEdit",
        "TableNameError",
        "TableNameLimitKind",
        "TableNamePatch",
        "TableNamePath",
        "NameCommit",
        "NameDiagnostics",
        "NameEdit",
        "NameError",
        "NameLimitKind",
        "NamePatch",
        "NamePath",
    }
)
KEYNOTE_SLIDE_TABLE_NAME_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ArchiveReferenceKind",
        "ArchiveReferenceOccurrence",
        "ArchiveReferencePolicy",
        "ArchiveReferenceScope",
        "ArchiveReferenceVisitor",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "FieldType",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "TableInfoArchive",
        "TableModelArchive",
        "TableModelSnapshot",
    }
)
KEYNOTE_SLIDE_TABLE_NAME_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeOptions",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "TableModelNameRewriteLimits",
        "TableModelNameRewriteOutput",
        "TableModelNameRewriteReport",
        "TableModelNameRewriteRequirements",
        "TableModelNameWrite",
        "PreparedTableModelNameRewrite",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_SLIDE_TABLE_NAME_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tst", "tsp", "tsd", "tswp", "litchi_iwa_protos"}
)
KEYNOTE_SLIDE_TABLE_NAME_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/table_model_discovery_codec.rs"
)
KEYNOTE_SLIDE_TABLE_NAME_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_SLIDE_TABLE_NAME_CODEC_MODULE = "table_model_discovery_codec"
KEYNOTE_SLIDE_TABLE_NAME_CODEC_REQUIRED_APIS = (
    "DecodeError",
    "DecodeOptions",
    "TableModelSnapshot",
    "TableModelNameWrite",
    "TableModelNameRewriteReport",
    "TableModelNameRewriteRequirements",
    "TableModelNameRewriteLimits",
    "TableModelNameRewriteOutput",
    "PreparedTableModelNameRewrite",
    "decode_table_model_with_report",
    "prepare_table_model_name_rewrite",
    "table_model_source_fingerprint",
)
KEYNOTE_SLIDE_TABLE_NAME_CODEC_REQUIRED_MARKER_GROUPS = {
    "field-8 name framing": ("TABLE_NAME_FIELD", "table_name"),
    "canonical duplicate/wire checks": (
        "known_field",
        "duplicate(",
        "noncanonical",
        "require_wire",
    ),
    "unknown/group preservation": ("skip_unknown", "nested_wire", "extend_from_slice"),
    "prepared execution contract": (
        "prepare_table_model_name_rewrite",
        "execution_requirements",
        "execute",
    ),
    "readback/source fingerprint": ("readback", "source_fingerprint"),
}
KEYNOTE_SLIDE_TABLE_NAME_PACKAGE_MARKER_GROUPS = {
    "aggregate transaction budget": ("Budget", "residual", "budget"),
    "prepared field-8 codec": (
        "table_model_discovery_codec",
        "prepare_table_model_name_rewrite",
        "execution_requirements",
        "execute",
    ),
    "field-8 authority": ("TABLE_MODEL_MESSAGE_TYPE", "table_name"),
    "prepared execution accounting": (
        "codec_report",
        "requirements",
        "exact_limits",
    ),
    "exact artifacts and inverse": (
        "ExactArtifacts",
        "source_fingerprint",
        "inverse",
        "is_noop",
        "PatchConflict",
    ),
    "prepared reassembly": (
        "prepare_reassembly",
        "EntryEdit",
        "execute",
    ),
    "candidate reopen and authority": (
        "candidate",
        "reopen",
        "validate",
        "authority",
    ),
    "object/member locality": ("verify_locality", "same_content"),
    "preview invalidation": ("root_preview_deletions", "deleted_previews"),
}
KEYNOTE_SLIDE_TABLE_NAME_TEST_SOURCES = (
    Path("crates/litchi-keynote/tests/slide_table_name.rs"),
)
KEYNOTE_SLIDE_TABLE_NAME_NOOP_PREVIEW_ZERO = re.compile(
    r"\bis_noop\s*\(\)[\s\S]{0,600}?"
    r"\bdeleted_previews\s*\(\)\s*,\s*0\b"
)
KEYNOTE_SLIDE_TABLE_NAME_CHANGED_PREVIEW_REPORT = re.compile(
    r"\bchanged\s*\(\)[\s\S]{0,800}?"
    r"\bdeleted_previews\s*\(\)\s*,\s*(?!0\b)"
    r"[A-Za-z_][A-Za-z0-9_]*(?:\s*\.\s*[A-Za-z_][A-Za-z0-9_]*\s*\(\))?"
)
KEYNOTE_SLIDE_TABLE_NAME_FUZZ_SOURCES = (
    Path("crates/litchi/fuzz/fuzz_targets/keynote_slide_table_name.rs"),
)
KEYNOTE_SLIDE_TABLE_NAME_FUZZ_CORPORA = (
    Path("crates/litchi/fuzz/corpus/keynote_slide_table_name"),
)


# The slide-table header owner owns only the seven persisted
# header/footer/freeze/repetition scalar fields of a Keynote slide table.  It
# is deliberately separate from the legacy editor: the latter still knows how
# to move rows/cells and therefore must not become an accidental dependency of
# this scalar rewrite.
KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "headers.rs"
)
KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table.rs"
KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_headers.rs"
)
KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_HEADERS_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "headers.rs"
KEYNOTE_SLIDE_TABLE_HEADERS_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_HEADERS_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "InvalidReason",
        "LimitKind",
        "Patch",
        "Path",
    }
)
KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_TYPES = frozenset(
    {"Count", "Error", "Settings"}
)
KEYNOTE_SLIDE_TABLE_HEADERS_TRANSACTION_TYPES = frozenset(
    {"Edit", "Patch", "Commit", "Diagnostics", "Error", "InvalidReason", "LimitKind", "Path"}
)
KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "TableSelector"}
)
KEYNOTE_SLIDE_TABLE_HEADERS_PACKAGE_METHODS = frozenset(
    {
        "slide_table_header_settings",
        "edit_slide_table_headers",
        "apply_slide_table_headers",
    }
)
KEYNOTE_SLIDE_TABLE_HEADERS_EDIT_METHODS = frozenset(
    {"slide_position", "table_position", "path", "before", "after", "settings", "set", "commit"}
)
KEYNOTE_SLIDE_TABLE_HEADERS_FLAT_ALIASES = frozenset(
    {
        "TableHeaderCommit",
        "TableHeaderDiagnostics",
        "TableHeaderEdit",
        "TableHeaderError",
        "TableHeaderInvalidReason",
        "TableHeaderLimitKind",
        "TableHeaderPatch",
        "TableHeaderPath",
        "HeaderCommit",
        "HeaderDiagnostics",
        "HeaderEdit",
        "HeaderError",
        "HeaderInvalidReason",
        "HeaderLimitKind",
        "HeaderPatch",
        "HeaderPath",
        "HeaderSettings",
        "HeaderSettingsCommit",
        "HeaderSettingsDiagnostics",
        "HeaderSettingsEdit",
        "HeaderSettingsError",
        "HeaderSettingsInvalidReason",
        "HeaderSettingsLimitKind",
        "HeaderSettingsPatch",
        "HeaderSettingsPath",
        "SlideTableHeaders",
        "SlideTableHeadersCommit",
        "SlideTableHeadersDiagnostics",
        "SlideTableHeadersEdit",
        "SlideTableHeadersError",
        "SlideTableHeadersInvalidReason",
        "SlideTableHeadersLimitKind",
        "SlideTableHeadersPatch",
        "SlideTableHeadersPath",
    }
)
KEYNOTE_SLIDE_TABLE_HEADERS_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ArchiveReferenceKind",
        "ArchiveReferenceOccurrence",
        "ArchiveReferencePolicy",
        "ArchiveReferenceScope",
        "ArchiveReferenceVisitor",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "FieldType",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "TableHeaderSettingsSnapshot",
        "TableHeaderSettingsWrite",
    }
)
KEYNOTE_SLIDE_TABLE_HEADERS_WIRE_TYPES = frozenset(
    {
        "DecodeError",
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
KEYNOTE_SLIDE_TABLE_HEADERS_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tst", "tsd", "tsp", "litchi_iwa_protos"}
)
KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/numbers_table_header_settings_codec.rs"
)
KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_MODULES = (
    "numbers_table_header_settings_codec",
    "table_header_settings_codec",
)
KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_REQUIRED_APIS = (
    "DecodeOptions",
    "DecodeError",
    "TableHeaderSettingsSnapshot",
    "TableHeaderSettingsWrite",
    "RewriteReport",
    "RewriteOutput",
    "PreparedTableHeaderSettingsRewrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
    "decode_table_header_settings",
    "prepare_table_header_settings_rewrite",
    "rewrite_table_header_settings",
    "rewrite_table_header_settings_with_report",
)
KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_REQUIRED_MARKER_GROUPS = {
    "strict known-field envelope": (
        "HEADER_ROWS_FIELD",
        "HEADER_COLUMNS_FIELD",
        "FOOTER_ROWS_FIELD",
        "HEADER_ROWS_FROZEN_FIELD",
        "HEADER_COLUMNS_FROZEN_FIELD",
        "REPEATING_HEADER_ROWS_FIELD",
        "REPEATING_HEADER_COLUMNS_FIELD",
    ),
    "canonical duplicate/wire checks": (
        "known_field",
        "duplicate(",
        "noncanonical",
        "WireTypeMismatch",
    ),
    "unknown and group policy": (
        "strict_unknown",
        "unknown",
        "group",
        "preserv",
    ),
    "prepared execution contract": (
        "check_execution_limits",
        "prepare_table_header_settings_rewrite",
        "execution_requirements",
        "execute",
    ),
    "bounded allocation axes": (
        "with_max_output_bytes",
        "with_max_allocations",
        "with_max_retained_bytes",
        "with_max_scratch_bytes",
    ),
    "readback parity": (
        "readback",
        "optional_write()",
        "source",
    ),
}
KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_MARKER_GROUPS = {
    "aggregate transaction budget": (
        "HeaderBudget",
        "TransactionWork",
        "residual",
        "budget",
    ),
    "selector-first strict codec": (
        "header_codec",
        "decode_header",
        "prepare_table_header_settings_rewrite",
        "execution_requirements",
        "exact_limits",
    ),
    "seven-field canonical model rewrite": (
        "TABLE_MODEL_MESSAGE_TYPE",
        "6_001",
        "rewrite_headers",
    ),
    "prepared execution accounting": (
        "execution_requirements",
        "exact_limits",
        "execute",
        "codec_report",
    ),
    "exact artifacts and inverse": (
        "ExactArtifacts",
        "inverse",
        "is_noop",
        "PatchConflict",
    ),
    "candidate reopen and locality": (
        "candidate.validate",
        "reopen",
        "verify_locality",
        "same_selection",
    ),
    "preview preservation": (
        "deleted_previews: 0",
    ),
}
KEYNOTE_SLIDE_TABLE_HEADERS_FORBIDDEN_OWNER_PATTERNS = (
    (
        "row movement",
        re.compile(
            r"\b(?:move|reorder|sort)_?(?:rows?|cells?)\b|"
            r"\b(?:rows?|cells?)_(?:move|reorder|sort)\b",
            re.I,
        ),
    ),
    (
        "cell/formula mutation",
        re.compile(
            r"\b(?:cell|cells?|formula|formulas?|tile|tiles?|data_store|data_stores?)_"
            r"(?:rewrite|update|move|mutate|reorder|sort)\b|"
            r"\b(?:rewrite|update|move|mutate|reorder|sort)_"
            r"(?:cell|cells?|formula|formulas?|tile|tiles?|data_store|data_stores?)\b",
            re.I,
        ),
    ),
    (
        "metadata mutation",
        re.compile(
            r"\b(?:set|advance|update|rewrite|add)_[A-Za-z0-9_]*"
            r"(?:metadata|uuid|save_token|object_identifier)\b|"
            r"\b(?:metadata|uuid|save_token|object_identifier)_"
            r"(?:set|advance|update|rewrite|add)\b",
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
)
KEYNOTE_SLIDE_TABLE_HEADERS_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_SLIDE_TABLE_HEADERS_TEST_SOURCES = (
    Path("crates/litchi-keynote/tests/slide_table_headers.rs"),
)
KEYNOTE_SLIDE_TABLE_HEADERS_FUZZ_SOURCES = (
    Path("crates/litchi-iwa-protos/fuzz/fuzz_targets/numbers_table_header_settings_codec.rs"),
    Path("crates/litchi/fuzz/fuzz_targets/keynote_slide_table_headers.rs"),
)
KEYNOTE_SLIDE_TABLE_HEADERS_FUZZ_CORPORA = (
    Path("crates/litchi-iwa-protos/fuzz/corpus/numbers_table_header_settings_codec"),
    Path("crates/litchi/fuzz/corpus/keynote_slide_table_headers"),
)

# Complete Keynote slide-table appearance ownership lives behind the
# selector-first package facade.  The existing shared table-appearance codec
# remains the strict wire seam; this inventory gates the package owner and
# its host retirement independently from the listing-only ratchet.
KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "table" / "appearance.rs"
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table.rs"
KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_table_appearance.rs"
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_SLIDE_TABLE_APPEARANCE_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "table" / "appearance.rs"
KEYNOTE_SLIDE_TABLE_APPEARANCE_FAMILY_MODULE = 'transaction'
KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
        "Path",
    }
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_TYPES = frozenset(
    {"Appearance", "Banding", "GridlineVisibility", "Gridlines", "RowSizing"}
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_TRANSACTION_TYPES = frozenset(
    {"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"}
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_TYPES = frozenset(
    {"SlideSelector", "TableSelector"}
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_PACKAGE_METHODS = frozenset(
    {
        "slide_table_appearance",
        "edit_slide_table_appearance",
        "apply_slide_table_appearance",
    }
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_EDIT_METHODS = frozenset(
    {"path", "appearance", "set", "commit"}
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_FLAT_ALIASES = frozenset(
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
KEYNOTE_SLIDE_TABLE_APPEARANCE_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ArchiveReferenceKind",
        "ArchiveReferenceOccurrence",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "FieldType",
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
        "TableStyleVariationWrite",
    }
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_WIRE_TYPES = frozenset(
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
KEYNOTE_SLIDE_TABLE_APPEARANCE_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tst", "tsd", "tsp", "litchi_iwa_protos"}
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/table_appearance_codec.rs"
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_MODULES = ("table_appearance_codec",)
KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_REQUIRED_APIS = (
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
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_REQUIRED_MARKER_GROUPS = {
    "canonical unknown preservation": (
        # The production pass is intentionally code-only.  These are the
        # actual raw-field/group traversal seams used by the codec; comments
        # and cfg(test) examples are not sufficient proof on their own.
        "scan_group_fields",
        "find_group_end",
        "copy_from_slice",
    ),
    "prepared rewrite contract": (
        "prepare_table_model_style_rewrite",
        "execution_requirements",
        "execute",
    ),
    "bounded decode axes": (
        "max_input_bytes",
        "max_output_bytes",
        "max_fields",
        "max_work_bytes",
        "max_allocations",
    ),
}
KEYNOTE_SLIDE_TABLE_APPEARANCE_TEST_SOURCES = (
    Path("crates/litchi-keynote/tests/slide_table_appearance.rs"),
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_FUZZ_SOURCES = (
    Path("crates/litchi-iwa-protos/fuzz/fuzz_targets/table_appearance.rs"),
    Path("crates/litchi/fuzz/fuzz_targets/keynote_slide_table_appearance.rs"),
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_FUZZ_CORPORA = (
    Path("crates/litchi-iwa-protos/fuzz/corpus/table_appearance"),
    Path("crates/litchi/fuzz/corpus/keynote_slide_table_appearance"),
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_MARKER_GROUPS = {
    "aggregate transaction budget": (
        "metadata_options",
        "WireLimits",
        "effective_archive_limits",
        "snappy_limits",
    ),
    "bounded resource axes": (
        "max_references",
        "max_iwa_stream_bytes",
        "effective_archive_limits",
        "snappy_limits",
        "try_reserve_exact",
    ),
    "strict prepared codec": (
        "appearance_codec",
        "prepare_table_model_style_rewrite",
        "execution_requirements",
        "exact_limits",
    ),
    "copy-on-write style graph": (
        "new_style",
        "replace_model_in_archive",
        "append_style_object",
        "replace_stylesheet_in_archive",
    ),
    "metadata ownership": (
        "MetadataFacts",
        "inspect_package_metadata_with_visitor",
        "uuid_additions",
        "SaveTokenBatch",
        "external_additions",
    ),
    "prepared execution accounting": (
        "model_requirements",
        "stylesheet_requirements",
        "execution_requirements",
        "execute",
    ),
    "package reassembly": (
        "prepare_reassembly_with_deletions",
        "EntryEdit",
        "execute",
    ),
    "exact artifacts and inverse": (
        "Patch",
        "source_fingerprint",
        "inverse",
        "is_noop",
        "PatchConflict",
    ),
    "candidate reopen and locality": (
        "candidate",
        "Package::from_source_with_options",
        "verify_candidate_locality",
        "full_reparse_performed",
    ),
    "preview preservation": (
        "deleted_previews",
        "root_preview_deletions",
    ),
}
KEYNOTE_SLIDE_TABLE_APPEARANCE_FORBIDDEN_OWNER_PATTERNS = (
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
        re.compile(
            r"\b(?:rewrite_table_model_style|append_stylesheet_style)\s*\("
        ),
    ),
)
KEYNOTE_SLIDE_TABLE_APPEARANCE_PUBLIC_RAW_ID_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|model_id|"
    r"component_id|member_id|archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)


IWA_KEYNOTE_SLIDE_TABLE_CONFIG_PACKAGE_RECEIVERS = frozenset(
    {
        "package",
        "keynote_package",
        "focused_package",
        "focused_keynote",
        "sort_package",
        "title_package",
        "lock_package",
        "name_package",
        "table_name_package",
    }
)


KEYNOTE_SHOW_SETTINGS_IMPLEMENTATION_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "show.rs",
    KEYNOTE_SOURCE_ROOT / "package" / "show_settings.rs",
)
KEYNOTE_SHOW_SETTINGS_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "lib.rs",
    KEYNOTE_SOURCE_ROOT / "package.rs",
)
KEYNOTE_SHOW_SETTINGS_FLAT_ALIASES = frozenset(
    {
        "ShowSettings",
        "ShowSettingsCommit",
        "ShowSettingsDiagnostics",
        "ShowSettingsEdit",
        "ShowSettingsError",
        "ShowSettingsLimitKind",
        "ShowSettingsPatch",
    }
)
KEYNOTE_SHOW_SETTINGS_FLAT_SEMANTIC_ALIASES = frozenset(
    {"Mode", "Settings", "Show", "Size"}
)
KEYNOTE_SHOW_SETTINGS_SHORT_NAMES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Mode",
        "Patch",
        "Settings",
        "Show",
        "Size",
    }
)
KEYNOTE_SHOW_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:show|show_settings)"
    r"[ \t\r\n]*::"
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_MODULES = ("soundtrack",)


# Build-order mutation is now owned by the selector-first Keynote package. The
# migration host must not restore either raw-ID operation or teach callers a
# host-side ``move`` command.
RETIRED_IWA_KEYNOTE_BUILD_ORDER_METHODS = (
    "move_slide_build",
    "reorder_slide_builds",
)
RETIRED_IWA_KEYNOTE_BUILD_ORDER_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_BUILD_ORDER_METHODS
)
KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "soundtrack.rs"
KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "soundtrack_settings.rs"
)
KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_ROOT = (
    KEYNOTE_SOURCE_ROOT / "package" / "soundtrack_settings"
)
KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_SOURCES = (
    KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_ROOT / "media.rs",
    KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_ROOT / "rewrite.rs",
)
KEYNOTE_SOUNDTRACK_SETTINGS_IMPLEMENTATION_SOURCES = (
    KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE,
    KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE,
    *KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_SOURCES,
)
KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "lib.rs",
    KEYNOTE_SOURCE_ROOT / "package.rs",
)
KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES = (
    "Mode",
    "Settings",
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
KEYNOTE_SOUNDTRACK_SETTINGS_SHORT_NAMES = frozenset(
    KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES
)
KEYNOTE_SOUNDTRACK_SETTINGS_PACKAGE_METHODS = (
    "soundtrack_settings",
    "edit_soundtrack_settings",
    "apply_soundtrack_settings",
)
KEYNOTE_SOUNDTRACK_SETTINGS_FLAT_ALIASES = frozenset(
    {
        "Soundtrack",
        "SoundtrackMode",
        "SoundtrackSettings",
        "SoundtrackEdit",
        "SoundtrackPatch",
        "SoundtrackCommit",
        "SoundtrackDiagnostics",
        "SoundtrackError",
        "SoundtrackLimitKind",
        "SoundtrackSettingsEdit",
        "SoundtrackSettingsPatch",
        "SoundtrackSettingsCommit",
        "SoundtrackSettingsDiagnostics",
        "SoundtrackSettingsError",
        "SoundtrackSettingsLimitKind",
    }
)
KEYNOTE_SOUNDTRACK_SETTINGS_FORBIDDEN_PUBLIC_MEMBERS = frozenset(
    {"set_soundtrack_settings"}
)
KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:soundtrack|soundtrack_settings)"
    r"[ \t\r\n]*::"
)
PUBLIC_KEYNOTE_PACKAGE_SOUNDTRACK_SETTINGS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?soundtrack_settings\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_PACKAGE_SOUNDTRACK_SETTINGS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?soundtrack_settings\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_SOUNDTRACK_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?soundtrack\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_SOUNDTRACK_TRANSACTION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transaction\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_SOUNDTRACK_SETTINGS_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "ReferenceSnapshot",
        "Resolved",
        "SnappyStream",
        "SoundtrackRecord",
        "SoundtrackSnapshot",
        "SoundtrackSettingsSnapshot",
        "SourceCatalog",
    }
)
KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_TYPES = frozenset(
    {
        "DecodeLimitKind",
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
KEYNOTE_SOUNDTRACK_SETTINGS_PROTO_ORIGINS = frozenset({"kn", "tsp"})
KEYNOTE_SOUNDTRACK_SETTINGS_MEDIA_TOPOLOGY_NAMES = frozenset(
    {
        "DataReference",
        "EmbeddedMediaAsset",
        "KeynoteSoundtrackItemInfo",
        "MediaAssetId",
        "data_reference",
        "data_references",
        "media",
        "media_items",
        "movie_media",
        "payload",
        "payloads",
    }
)

# Soundtrack item media lifecycle is moving into a focused Keynote owner.  The
# Keep the paths and names explicit so a similarly named helper cannot satisfy
# the migration merely by importing the old host.
KEYNOTE_SOUNDTRACK_ITEMS_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "soundtrack" / "items.rs"
)
KEYNOTE_SOUNDTRACK_ITEMS_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "soundtrack_items.rs"
)
KEYNOTE_SOUNDTRACK_ITEMS_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "lib.rs",
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "soundtrack.rs",
)
KEYNOTE_SOUNDTRACK_ITEMS_CANONICAL_TYPES = (
    "Item",
    "ItemHandle",
    "ItemSelector",
    "AudioSource",
    "OperationKind",
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
KEYNOTE_SOUNDTRACK_ITEMS_SHORT_NAMES = frozenset(
    KEYNOTE_SOUNDTRACK_ITEMS_CANONICAL_TYPES
)
KEYNOTE_SOUNDTRACK_ITEMS_PACKAGE_METHODS = (
    "soundtrack_items",
    "edit_soundtrack_items",
    "apply_soundtrack_items",
)
KEYNOTE_SOUNDTRACK_ITEMS_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?items\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_PACKAGE_SOUNDTRACK_ITEMS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?soundtrack_items\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_SOUNDTRACK_ITEMS_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_SOUNDTRACK_ITEMS_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
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
KEYNOTE_SOUNDTRACK_ITEMS_PROTO_ORIGINS = frozenset(
    {"buffa", "prost", "prost_types", "kn", "tsa", "tsk", "tsp"}
)
KEYNOTE_SOUNDTRACK_ITEMS_TOPOLOGY_NAMES = frozenset(
    {
        "ComponentDataReference",
        "ComponentInfo",
        "DataInfo",
        "DataReference",
        "EmbeddedMediaAsset",
        "FieldInfo",
        "KeynoteSoundtrackItemInfo",
        "MediaAssetId",
        "PackageMetadata",
        "data_reference",
        "data_references",
        "metadata",
        "media",
        "media_items",
        "movie_media",
        "payload",
        "payloads",
    }
)
# AudioSource is a semantic content value, so its immutable payload accessor
# is intentionally allowed.  Package bytes, wire buffers, and media topology
# remain forbidden from the public focused API.
KEYNOTE_SOUNDTRACK_ITEMS_ALLOWED_RAW_SLICE_METHODS = frozenset(
    {"bytes", "as_ref", "from_bytes"}
)
RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_MODULES = (
    "transition",
    "transition_lifecycle",
)
KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "transition.rs",
    KEYNOTE_SOURCE_ROOT / "package" / "slide_transition.rs",
)
KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "lib.rs",
    KEYNOTE_SOURCE_ROOT / "package.rs",
)
KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
KEYNOTE_SLIDE_TRANSITION_SEMANTIC_TYPES = (
    "Acceleration",
    "AccelerationKind",
    "AnimationParameters",
    "CustomParameters",
    "Direction",
    "Effect",
    "MosaicType",
    "Settings",
    "SettingsBuilder",
    "TextDelivery",
    "TextDeliveryKind",
    "TimingCurveSlot",
)
KEYNOTE_SLIDE_TRANSITION_SHORT_NAMES = frozenset(
    KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
)
KEYNOTE_SLIDE_TRANSITION_FLAT_ALIAS_PREFIXES = (
    "SlideTransition",
    "Transition",
)
KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in KEYNOTE_SLIDE_TRANSITION_FLAT_ALIAS_PREFIXES
    for suffix in KEYNOTE_SLIDE_TRANSITION_SHORT_NAMES
)
# Opaque wire validation is an adapter seam. It must not become a second
# public transition operation or expose wire-policy details through the
# archive-free semantic facade.
KEYNOTE_SLIDE_TRANSITION_FORBIDDEN_PUBLIC_MEMBERS = frozenset(
    {"validate_opaque_transition_settings"}
)
KEYNOTE_SLIDE_TRANSITION_ROOT_ALIASES = frozenset(
    KEYNOTE_SLIDE_TRANSITION_SEMANTIC_TYPES
    + KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
)
KEYNOTE_SLIDE_TRANSITION_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:transition|slide_transition)"
    r"[ \t\r\n]*::"
)
PUBLIC_KEYNOTE_PACKAGE_SLIDE_TRANSITION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?slide_transition\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_TRANSITION_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?transition\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_SLIDE_TRANSITION_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "EntryEdit",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "TransitionSettingsSnapshot",
    }
)
KEYNOTE_SLIDE_TRANSITION_WIRE_TYPES = frozenset(
    {
        "WireDescent",
        "WireError",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_SLIDE_TRANSITION_SEMANTIC_IDENTIFIER_NAMES = frozenset(
    {
        "MAX_IDENTIFIER_BYTES",
        "from_identifier",
        "identifier",
        "identifiers",
    }
)
KEYNOTE_SLIDE_TRANSITION_SEMANTIC_OBJECT_NAMES = frozenset(
    {
        "BY_OBJECT_DELIVERY",
        "ByObject",
        "magic_move_fade_unmatched_objects",
        "set_magic_move_fade_unmatched_objects",
    }
)
KEYNOTE_SLIDE_TRANSITION_SEMANTIC_OPAQUE_PAYLOAD_MEMBERS = frozenset(
    {
        "color_payload",
        "set_color_payload",
        "set_timing_curve_payload",
        "timing_curve_payload",
        "timing_curve_payloads",
    }
)

RETIRED_IWA_KEYNOTE_SLIDE_BACKGROUND_MODULES = (
    "slide_background",
    "slide_background_color",
    "slide_background_gradient_wire",
)
RETIRED_IWA_KEYNOTE_SLIDE_BACKGROUND_MODULE_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SLIDE_BACKGROUND_MODULES
)
IWA_KEYNOTE_SLIDE_BACKGROUND_ORACLE_MODULE = re.compile(
    r"^[ \t]*#[ \t]*\[[ \t]*cfg[ \t]*\([ \t]*test[ \t]*\)[ \t]*\]"
    r"[ \t\r\n]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(?P<module>slide_background_color|"
    r"slide_background_gradient_wire)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)

# These are ownership markers, not a ban on using the focused package. Keep
# the list explicit so the adapter can continue to call semantic package
# methods while any accidental return of the old generated/wire/style graph
# implementation is reported at the exact source line.
IWA_KEYNOTE_SLIDE_BACKGROUND_ADAPTER_GENERATED_MARKERS = frozenset(
    {
        "kn",
        "tsd",
        "tsp",
        "tss",
        "tswp",
        "prost",
        "prost_types",
        "litchi_iwa_protos",
        "Message",
        "ArchiveObject",
        "RawMessage",
        "HostPackage",
        "SlideArchive",
        "SlideStyleArchive",
        "SlideStylePropertiesArchive",
        "StyleArchive",
        "StylesheetArchive",
        "FillArchive",
        "GradientArchive",
        "Color",
        "Reference",
        "ObjectGraph",
        "decode_type",
        "decode_message",
        "message_data_type",
        "archive_name",
    }
)
IWA_KEYNOTE_SLIDE_BACKGROUND_ADAPTER_WIRE_MARKERS = frozenset(
    {
        "WireView",
        "WireLimits",
        "WireFieldView",
        "parse_wire_field",
        "parse_wire_fields",
        "optional_length_delimited_payload",
        "required_length_delimited_payload",
        "repeated_length_delimited_payloads",
        "patch_length_delimited_field",
        "patch_fixed32_field",
        "patch_fixed64_field",
        "patch_varint_field",
        "patch_nested_length_delimited_field",
        "patch_nested_fixed32_field",
        "patch_nested_fixed64_field",
        "patch_nested_varint_field",
        "remove_repeated_length_delimited_field_where",
        "rewrite_repeated_length_delimited_fields",
        "transform_length_delimited_field",
        "transform_length_delimited_fields_at_path",
    }
)
IWA_KEYNOTE_SLIDE_BACKGROUND_ADAPTER_STYLE_MARKERS = frozenset(
    {
        "style_id",
        "stylesheet_id",
        "SLIDE_STYLE_MESSAGE_TYPE",
        "STYLESHEET_MESSAGE_TYPE",
        "is_collapsible_background_variation",
        "style_is_exclusive",
        "patch_slide_style_reference",
        "patch_stylesheet",
        "rewrite_stylesheet_data",
        "update_package_metadata",
        "ensure_slide_style_external_reference",
        "component_identifier_for_entry",
        "next_object_identifier",
        "set_package_last_object_identifier",
        "clone_style_object",
        "replace_style_data",
        "release_package_identifier_suffix",
        "remove_background_field_metadata",
        "new_style_object",
        "encode_background_fill",
        "gradient_to_fill",
        "native_color_space",
        "color_to_native",
        "color_from_native",
        "gradient_from_fill",
        "opaque_background",
        "has_exact_fields",
        "update_archive",
    }
)
IWA_KEYNOTE_SLIDE_BACKGROUND_ADAPTER_MARKER_GROUPS = (
    (
        "generated protobuf ownership marker",
        IWA_KEYNOTE_SLIDE_BACKGROUND_ADAPTER_GENERATED_MARKERS,
    ),
    ("wire ownership marker", IWA_KEYNOTE_SLIDE_BACKGROUND_ADAPTER_WIRE_MARKERS),
    (
        "style-registry ownership marker",
        IWA_KEYNOTE_SLIDE_BACKGROUND_ADAPTER_STYLE_MARKERS,
    ),
)

KEYNOTE_SLIDE_BACKGROUND_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_background.rs"
)
KEYNOTE_SLIDE_BACKGROUND_IMPLEMENTATION_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "background.rs",
    KEYNOTE_SLIDE_BACKGROUND_OWNER_SOURCE,
)
KEYNOTE_SLIDE_BACKGROUND_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "lib.rs",
    KEYNOTE_SOURCE_ROOT / "package.rs",
)
KEYNOTE_SLIDE_BACKGROUND_PACKAGE_METHODS = (
    "slide_background",
    "slide_background_override",
    "edit_slide_background",
    "apply_slide_background",
)
KEYNOTE_SLIDE_BACKGROUND_EDIT_METHODS = ("set", "clear")
KEYNOTE_PACKAGE_SLIDE_BACKGROUND_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?slide_background\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_PACKAGE_SLIDE_BACKGROUND_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?slide_background\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
RETIRED_IWA_KEYNOTE_SLIDE_DELETE_MODULES = ("slide_delete",)
KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "slide" / "delete.rs"
KEYNOTE_SLIDE_DELETE_OWNER_SOURCE = KEYNOTE_SOURCE_ROOT / "package" / "slide_delete.rs"
KEYNOTE_SLIDE_DELETE_IMPLEMENTATION_SOURCES = (
    KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE,
    KEYNOTE_SLIDE_DELETE_OWNER_SOURCE,
)
KEYNOTE_SLIDE_DELETE_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "lib.rs",
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "slide.rs",
)
KEYNOTE_SLIDE_DELETE_CANONICAL_TYPES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
    "Path",
)
KEYNOTE_SLIDE_DELETE_SHORT_NAMES = frozenset(KEYNOTE_SLIDE_DELETE_CANONICAL_TYPES)
KEYNOTE_SLIDE_DELETE_PACKAGE_METHODS = (
    "edit_slide_deletion",
    "apply_slide_deletion",
)
KEYNOTE_SLIDE_DELETE_EDIT_METHODS = ("remove_slide",)
KEYNOTE_SLIDE_DELETE_FLAT_ALIAS_PREFIXES = (
    "SlideDelete",
    "SlideDeletion",
)
KEYNOTE_SLIDE_DELETE_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in KEYNOTE_SLIDE_DELETE_FLAT_ALIAS_PREFIXES
    for suffix in KEYNOTE_SLIDE_DELETE_SHORT_NAMES
)
KEYNOTE_SLIDE_DELETE_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:slide_delete|slide[ \t\r\n]*::"
    r"[ \t\r\n]*(?:r#)?delete)"
    r"(?=[ \t\r\n]*(?:::|as\b|;|=))"
)
KEYNOTE_PACKAGE_SLIDE_DELETE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?slide_delete\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_PACKAGE_SLIDE_DELETE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?slide_delete\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_SLIDE_DELETE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?delete\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_SLIDE_DELETE_PHYSICAL_TYPES = frozenset(
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
        "SlideArchive",
        "SlideDeletionSnapshot",
        "SlideNodeArchive",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_SLIDE_DELETE_WIRE_TYPES = frozenset(
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
KEYNOTE_SLIDE_DELETE_PROTO_ORIGINS = frozenset(
    {"kn", "prost", "prost_types", "tsp", "tswp"}
)
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_MODULES = (
    "placeholder_visibility",
)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_MODULES = ("slide_number",)
KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "slide" / "placeholder.rs"
)
KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_placeholder_visibility.rs"
)
KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_placeholder_visibility"
)
KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_SOURCES = (
    KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT / "errors.rs",
    KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT / "resolve.rs",
    KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT / "rewrite.rs",
    KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT / "slide_number.rs",
    KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT / "verification.rs",
)
KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_PREVIEW_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_preview.rs"
)
KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_ROOT = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_preview"
)
KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_SOURCES = (
    KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_ROOT / "slide_number.rs",
)
KEYNOTE_PLACEHOLDER_VISIBILITY_IMPLEMENTATION_SOURCES = (
    KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE,
    KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE,
    *KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_SOURCES,
    *KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_SOURCES,
)
KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "lib.rs",
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "slide.rs",
)
KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES = (
    "Kind",
    "State",
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
    "Error",
    "LimitKind",
)
KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_KINDS = (
    "Title",
    "Body",
    "SlideNumber",
)
KEYNOTE_PLACEHOLDER_VISIBILITY_SHORT_NAMES = frozenset(
    KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES
)
KEYNOTE_PLACEHOLDER_VISIBILITY_PACKAGE_METHODS = (
    "slide_placeholder_visibility",
    "edit_slide_placeholder_visibility",
    "apply_slide_placeholder_visibility",
)
KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIAS_PREFIXES = (
    "Placeholder",
    "PlaceholderVisibility",
    "SlidePlaceholder",
    "SlidePlaceholderVisibility",
    "SlideTextPlaceholder",
    "SlideNumber",
    "SlideNumberPlaceholder",
    "SlideNumberVisibility",
    "SlideNumberPlaceholderVisibility",
)
KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES = frozenset(
    prefix + suffix
    for prefix in KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIAS_PREFIXES
    for suffix in KEYNOTE_PLACEHOLDER_VISIBILITY_SHORT_NAMES
)
KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_SEMANTIC_ALIASES = frozenset(
    {"SlideNumberPlaceholder", "SlideNumberVisibility"}
)
KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_NUMBER_PUBLIC_MEMBERS = frozenset(
    {
        "apply_slide_number_visibility",
        "edit_slide_number_visibility",
        "hide_slide_number",
        "is_slide_number_visible",
        "set_slide_number_visible",
        "show_slide_number",
        "slide_number_visibility",
    }
)
KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:placeholder|slide_placeholder_visibility)"
    r"[ \t\r\n]*::"
)
PUBLIC_KEYNOTE_PACKAGE_PLACEHOLDER_VISIBILITY_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?slide_placeholder_visibility\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_PACKAGE_PLACEHOLDER_VISIBILITY_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?slide_placeholder_visibility\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_SLIDE_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?slide\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_PLACEHOLDER_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?placeholder\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
PUBLIC_KEYNOTE_SLIDE_NUMBER_HELPER_MODULE = re.compile(
    r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?slide_number\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
KEYNOTE_PLACEHOLDER_VISIBILITY_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "NodeVisibilitySnapshot",
        "PhysicalSource",
        "PlaceholderOwnerSnapshot",
        "PlaceholderTextOwnerSnapshot",
        "PlaceholderVisibilitySnapshot",
        "RawMessage",
        "ReferenceSnapshot",
        "Resolved",
        "SlideOwnerSnapshot",
        "SlideNodeSnapshot",
        "SlideNumberSnapshot",
        "SlideNumberVisibilitySnapshot",
        "SnappyStream",
        "SourceCatalog",
    }
)
KEYNOTE_PLACEHOLDER_VISIBILITY_PROTO_ORIGINS = frozenset({"kn", "tsp"})
KEYNOTE_PLACEHOLDER_VISIBILITY_WIRE_TYPES = frozenset(
    {
        "DecodeLimitKind",
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


KEYNOTE_MEDIA_CREATION_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_media_creation_codec.rs"
)
KEYNOTE_MEDIA_CREATION_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE = Path(
    "crates/litchi-iwa-protos/src/buffa-projections/"
    "TSDKeynoteMediaCreationArchive.proto"
)
KEYNOTE_MEDIA_CREATION_BUILD_SOURCE = Path("crates/litchi-iwa-protos/build.rs")
KEYNOTE_MEDIA_CREATION_CODEC_MODULE = "keynote_media_creation_codec"
KEYNOTE_MEDIA_CREATION_GENERATED_MODULE = (
    "buffa_keynote_media_creation_generated"
)
KEYNOTE_MEDIA_CREATION_GENERATED_INCLUDE = (
    "/buffa-keynote-media-creation/"
    "iwa_keynote_media_creation_buffa_protos.rs"
)
KEYNOTE_MEDIA_CREATION_CODEC_REQUIRED_MARKERS = (
    "pub struct MediaArchiveWrite",
    "pub enum MediaKind",
    "pub struct EncodeOptions",
    "pub struct EncodeReport",
    "pub struct EncodeOutput",
    "pub enum EncodeError",
    "pub fn encode_media_archive(",
    "pub fn encode_media_archive_with_report(",
    "try_encoded_len",
    "try_encode_bounded",
    "try_reserve_exact",
    "fn preflight(",
    "buffa_keynote_media_creation_generated::",
)
KEYNOTE_MEDIA_CREATION_CODEC_FORBIDDEN_OPERATIONS = (
    "prost::",
    "prost_types",
    "encode_to_vec",
    "try_encode_to_vec",
    "to_owned_message",
    "OwnedView",
    "crate::tsd::",
    "ArchiveObject",
    "HostPackage",
    "RawMessage",
)
KEYNOTE_MEDIA_CREATION_PROJECTION_FIELDS = {
    "Reference": {
        "identifier": ("required", "uint64", 1),
    },
    "DataReference": {
        "identifier": ("required", "uint64", 1),
    },
    "Point": {
        "x": ("required", "float", 1),
        "y": ("required", "float", 2),
    },
    "Size": {
        "width": ("required", "float", 1),
        "height": ("required", "float", 2),
    },
    "GeometryArchive": {
        "position": ("optional", "Point", 1),
        "size": ("optional", "Size", 2),
        "flags": ("optional", "uint32", 3),
        "angle": ("optional", "float", 4),
    },
    "ExteriorTextWrapArchive": {
        "type": ("optional", "uint32", 1),
        "direction": ("optional", "uint32", 2),
        "fit_type": ("optional", "uint32", 3),
        "margin": ("optional", "float", 4),
        "alpha_threshold": ("optional", "float", 5),
        "is_html_wrap": ("optional", "bool", 6),
    },
    "DrawableArchive": {
        "geometry": ("optional", "GeometryArchive", 1),
        "parent": ("optional", "Reference", 2),
        "exterior_text_wrap": ("optional", "ExteriorTextWrapArchive", 3),
        "locked": ("optional", "bool", 5),
        "aspect_ratio_locked": ("optional", "bool", 7),
        "title": ("optional", "Reference", 10),
        "caption": ("optional", "Reference", 11),
        "title_hidden": ("optional", "bool", 12),
        "caption_hidden": ("optional", "bool", 13),
    },
    "MovieArchive": {
        "super": ("required", "DrawableArchive", 1),
        "movie_data": ("optional", "DataReference", 14),
        "start_time": ("optional", "float", 3),
        "end_time": ("optional", "float", 4),
        "poster_time": ("optional", "float", 5),
        # The canonical field is an enum (wire-compatible with int32).  The
        # projection keeps the raw scalar so the writer never invents a
        # generated enum conversion at this boundary.
        "loop_option": ("optional", "int32", 24),
        "volume": ("optional", "float", 7),
        "audio_only": ("optional", "bool", 9),
        "streaming": ("optional", "bool", 18),
        "plays_across_slides": ("optional", "bool", 28),
        "poster_image_data": ("optional", "DataReference", 15),
        "poster_image_generated_with_alpha_support": ("optional", "bool", 23),
        "flags": ("optional", "uint32", 13),
        "style": ("optional", "Reference", 19),
        "original_size": ("optional", "Size", 20),
        "natural_size": ("optional", "Size", 21),
    },
}

# The focused Package owner and its private children must stay on neutral
# codecs.  The package crate still needs archive and source-catalog primitives
# for publication, so this guard targets only generated schema imports/types
# and direct Prost reads rather than banning the physical transaction layer.
KEYNOTE_SLIDE_AUDIO_CREATION_OWNER_SOURCE = Path(
    "crates/litchi-keynote/src/package/slide_audio_creation.rs"
)
KEYNOTE_SLIDE_AUDIO_CREATION_CHILD_ROOT = Path(
    "crates/litchi-keynote/src/package/slide_audio_creation"
)
KEYNOTE_SLIDE_AUDIO_CREATION_GENERATED_IMPORT = re.compile(
    r"(?s)\blitchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*(?:"
    r"(?:r#)?(?:kn|tn|tp|tsa|tsd|tsp|tss|tswp|"
    r"knsos|tnsos|tpsos|tsasos|tsdsos|tsssos|tswpsos)\b|"
    r"\{[^}]*\b(?:kn|tn|tp|tsa|tsd|tsp|tss|tswp|"
    r"knsos|tnsos|tpsos|tsasos|tsdsos|tsssos|tswpsos)\b"
    r")"
)
KEYNOTE_SLIDE_AUDIO_CREATION_GENERATED_MODULE = re.compile(
    r"(?<![A-Za-z0-9_#])(?:kn|tn|tp|tsa|tsd|tsp|tss|tswp|"
    r"knsos|tnsos|tpsos|tsasos|tsdsos|tsssos|tswpsos)"
    r"[ \t\r\n]*::"
)
KEYNOTE_SLIDE_AUDIO_CREATION_GENERATED_TYPE = re.compile(
    r"(?<![A-Za-z0-9_#])(?:DocumentArchive|ShowArchive|"
    r"StylesheetArchive|MovieArchive|DrawableArchive)\b"
)
KEYNOTE_SLIDE_AUDIO_CREATION_FORBIDDEN_SOURCE_PATTERNS = (
    (
        "prost import",
        re.compile(r"(?<![A-Za-z0-9_#])prost(?:_types)?[ \t\r\n]*::"),
    ),
    (
        "generated protobuf import",
        KEYNOTE_SLIDE_AUDIO_CREATION_GENERATED_IMPORT,
    ),
    (
        "generated protobuf module",
        KEYNOTE_SLIDE_AUDIO_CREATION_GENERATED_MODULE,
    ),
    (
        "generated archive type",
        KEYNOTE_SLIDE_AUDIO_CREATION_GENERATED_TYPE,
    ),
    (
        "generated protobuf decode",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?:Message|[A-Za-z_][A-Za-z0-9_]*)"
            r"[ \t\r\n]*::[ \t\r\n]*decode[ \t\r\n]*\("
        ),
    ),
    (
        "eager protobuf encoding",
        re.compile(
            r"\b(?:encode_to_vec|try_encode_to_vec|to_owned_message)"
            r"[ \t\r\n]*\("
        ),
    ),
)

# File-backed movie creation shares the audio transaction's private planning
# owner.  Keep a separate operation ratchet so a future ``CreationOptions::Movie``
# branch cannot silently delegate back to the migration host or materialize a
# generated schema value.  The standalone movie owner and shared audio owner
# are both scanned because the latter contains the reusable graph/metadata
# helpers; the operation marker below activates the audit.
KEYNOTE_SLIDE_MOVIE_CREATION_OWNER_SOURCE = Path(
    "crates/litchi-keynote/src/package/slide_movie_creation.rs"
)
KEYNOTE_SLIDE_MOVIE_CREATION_SHARED_OWNER_SOURCE = KEYNOTE_SLIDE_AUDIO_CREATION_OWNER_SOURCE
KEYNOTE_SLIDE_MOVIE_CREATION_CHILD_ROOT = KEYNOTE_SLIDE_AUDIO_CREATION_CHILD_ROOT
KEYNOTE_SLIDE_MOVIE_CREATION_NATIVE_FIXTURE = Path(
    "test-data/iwork/keynote/slide-movie-creation-focused-native.key"
)
KEYNOTE_SLIDE_MOVIE_CREATION_FRESH_NATIVE_FIXTURE = Path(
    "test-data/iwork/keynote/slide-movie-creation-fresh-focused-native.key"
)
KEYNOTE_SLIDE_MOVIE_CREATION_NATIVE_EVIDENCE = (
    (
        KEYNOTE_SLIDE_MOVIE_CREATION_NATIVE_FIXTURE,
        "native_saved_file_movie_creation_reuse_candidate_is_stable",
    ),
    (
        KEYNOTE_SLIDE_MOVIE_CREATION_FRESH_NATIVE_FIXTURE,
        "native_saved_file_movie_creation_fresh_candidate_is_stable",
    ),
)
KEYNOTE_SLIDE_MOVIE_CREATION_NATIVE_TESTS = tuple(
    test_name for _, test_name in KEYNOTE_SLIDE_MOVIE_CREATION_NATIVE_EVIDENCE
)

# The focused owner may use archive/source-catalog primitives to publish the
# transaction.  These patterns therefore match only generated schemas,
# eager protobuf operations, and explicit calls back into the legacy editor.
KEYNOTE_SLIDE_MOVIE_CREATION_FORBIDDEN_SOURCE_PATTERNS = (
    *KEYNOTE_SLIDE_AUDIO_CREATION_FORBIDDEN_SOURCE_PATTERNS,
    (
        "legacy editor delegation",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?:Draft|HostPackage)"
            r"(?:[ \t\r\n]*::|[ \t\r\n]*\.)|"
            r"(?<![A-Za-z0-9_#])(?:editor|host|legacy_editor)"
            r"[ \t\r\n]*\.[ \t\r\n]*add_slide_movie\b"
        ),
    ),
)
KEYNOTE_SLIDE_MOVIE_CREATION_ROUTE = re.compile(
    r"(?<![A-Za-z0-9_#])(?:CreationOptions|MediaCreationOptions)"
    r"[ \t\r\n]*::[ \t\r\n]*Movie\b|"
    r"(?<![A-Za-z0-9_#])(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t\r\n]+(?:add|create)_slide_movie\b|"
    r"(?<![A-Za-z0-9_#])make_movie_objects\b|"
    r"(?<![A-Za-z0-9_#])SlideMovieCreation(?:Commit|Patch|Error|Diagnostics)\b"
)
KEYNOTE_SLIDE_MOVIE_CREATION_CODEC_ROUTE = re.compile(
    r"(?<![A-Za-z0-9_#])(?:litchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*)?"
    r"(?:keynote_media_creation_codec|media_creation_codec)"
    r"[ \t\r\n]*::[ \t\r\n]*"
    r"encode_media_archive(?:_with_report)?[ \t\r\n]*\("
)
KEYNOTE_SLIDE_MOVIE_CREATION_RAW_API_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_#])(?:object_id|native_id|archive_id|component_id|"
    r"drawable_id|movie_id|raw_id)[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u64|u32|usize|Option[ \t\r\n]*<[ \t\r\n]*u64[ \t\r\n]*>)"
)
KEYNOTE_SLIDE_MOVIE_CREATION_RAW_API_TYPE = re.compile(
    r"(?<![A-Za-z0-9_#])(?:ArchiveObject|RawMessage|Draft|HostPackage)\b"
)


def _keynote_show_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return KEYNOTE_SHOW_OWNER_PATH.search(declaration) is not None or any(
        "show_settings" in identifier.lower() for identifier in identifiers
    )


def _is_keynote_show_settings_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    identifier_set = set(identifiers)
    return bool(
        identifier_set & KEYNOTE_SHOW_SETTINGS_FLAT_ALIASES
    ) or _keynote_show_owner_declaration(declaration)


def _keynote_soundtrack_settings_public_leak(identifier: str) -> str | None:
    """Classify physical vocabulary in the focused soundtrack-settings API."""

    if identifier in KEYNOTE_SOUNDTRACK_SETTINGS_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in KEYNOTE_SOUNDTRACK_SETTINGS_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_TYPES:
        return "wire type"
    if identifier in KEYNOTE_SOUNDTRACK_SETTINGS_MEDIA_TOPOLOGY_NAMES:
        return "soundtrack media topology"
    return _iwork_public_leak(identifier)


def _keynote_soundtrack_settings_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_PATH.search(
        declaration
    ) is not None or any(
        identifier in KEYNOTE_SOUNDTRACK_SETTINGS_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_keynote_soundtrack_settings_public_declaration(
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
            KEYNOTE_SOUNDTRACK_SETTINGS_FLAT_ALIASES
            | KEYNOTE_SOUNDTRACK_SETTINGS_FORBIDDEN_PUBLIC_MEMBERS
            | {"soundtrack", "soundtrack_settings"}
        )
    ) or _keynote_soundtrack_settings_owner_declaration(declaration)


def _keynote_slide_transition_public_leak(
    identifier: str, *, semantic_source: bool
) -> str | None:
    """Classify implementation vocabulary in the focused transition API."""

    if (
        semantic_source
        and identifier
        in (
            KEYNOTE_SLIDE_TRANSITION_SEMANTIC_IDENTIFIER_NAMES
            | KEYNOTE_SLIDE_TRANSITION_SEMANTIC_OBJECT_NAMES
        )
    ):
        return None
    if identifier in KEYNOTE_SLIDE_TRANSITION_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in KEYNOTE_SLIDE_TRANSITION_WIRE_TYPES:
        return "wire type"
    return _iwork_public_leak(identifier)


def _keynote_slide_transition_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return KEYNOTE_SLIDE_TRANSITION_OWNER_PATH.search(declaration) is not None or any(
        identifier
        in {"apply_slide_transition", "edit_slide_transition", "slide_transition"}
        for identifier in identifiers
    )


def _is_keynote_slide_transition_public_declaration(
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
            KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES
            | KEYNOTE_SLIDE_TRANSITION_FORBIDDEN_PUBLIC_MEMBERS
        )
    ) or _keynote_slide_transition_owner_declaration(declaration)


def _keynote_slide_delete_public_leak(identifier: str) -> str | None:
    """Classify physical vocabulary forbidden in slide-deletion APIs."""

    if identifier in KEYNOTE_SLIDE_DELETE_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in KEYNOTE_SLIDE_DELETE_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in KEYNOTE_SLIDE_DELETE_WIRE_TYPES:
        return "wire type"
    return _iwork_public_leak(identifier)


def _keynote_slide_delete_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return KEYNOTE_SLIDE_DELETE_OWNER_PATH.search(declaration) is not None or any(
        identifier
        in (
            set(KEYNOTE_SLIDE_DELETE_PACKAGE_METHODS)
            | set(KEYNOTE_SLIDE_DELETE_EDIT_METHODS)
        )
        for identifier in identifiers
    )


def _is_keynote_slide_delete_public_declaration(
    declaration: str, *, dedicated_source: bool
) -> bool:
    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(identifiers & KEYNOTE_SLIDE_DELETE_FLAT_ALIASES) or (
        _keynote_slide_delete_owner_declaration(declaration)
    )


def _keynote_placeholder_visibility_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary in placeholder visibility APIs."""

    if identifier in KEYNOTE_PLACEHOLDER_VISIBILITY_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in KEYNOTE_PLACEHOLDER_VISIBILITY_PHYSICAL_TYPES:
        return "archive/IWA type"
    if (
        identifier == "wire"
        or identifier in KEYNOTE_PLACEHOLDER_VISIBILITY_WIRE_TYPES
    ):
        return "wire type"
    return _iwork_public_leak(identifier)


def _keynote_placeholder_visibility_owner_declaration(declaration: str) -> bool:
    identifiers = [
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    ]
    return KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_PATH.search(
        declaration
    ) is not None or any(
        identifier in KEYNOTE_PLACEHOLDER_VISIBILITY_PACKAGE_METHODS
        for identifier in identifiers
    )


def _is_keynote_placeholder_visibility_public_declaration(
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
            KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES
            | KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_SEMANTIC_ALIASES
            | KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_NUMBER_PUBLIC_MEMBERS
            | {"slide_number"}
        )
    ) or _keynote_placeholder_visibility_owner_declaration(declaration)


def _keynote_family_view(
    root: Path, path: Path, sources: dict, module: str | None
) -> str:
    """Return the Rust scope that publishes one focused Keynote family.

    ADR 0034 moved every Keynote transaction family into a
    focused public module (``chart::arrangement``,
    ``slide::table::sort::transaction``, ...).  ``path`` is either the private
    owner source (returned unchanged) or the semantic module file; for the
    latter the body of the inline ``pub mod <module>`` is returned, or the
    whole file when the family is re-exported at its top level.
    """

    if path in sources:
        return sources[path]
    absolute = path if path.is_absolute() else root / path
    source = (
        _mask_rust_cfg_test_items(absolute.read_text(encoding="utf-8"))
        if absolute.is_file()
        else ""
    )
    if module is None:
        return source
    match = re.search(
        rf"(?m)^[ \t]*pub[ \t]+mod[ \t]+{re.escape(module)}[ \t]*\{{", source
    )
    if match is None:
        return ""
    depth, cursor = 1, match.end()
    while cursor < len(source) and depth:
        if source[cursor] == "{":
            depth += 1
        elif source[cursor] == "}":
            depth -= 1
        cursor += 1
    return source[match.end() : cursor - 1]


def audit_keynote_show_settings_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the nested, archive-free Keynote show-settings facade."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in KEYNOTE_SHOW_SETTINGS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in KEYNOTE_SHOW_SETTINGS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []
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
            if not _is_keynote_show_settings_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            show_owner_declaration = _keynote_show_owner_declaration(declaration)
            declaration_identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            if (
                public_declaration
                and path in export_sources
                and show_owner_declaration
                and declaration_identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-keynote show-settings public API retains "
                    "flat semantic aliases via show glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    public_declaration
                    and identifier in KEYNOTE_SHOW_SETTINGS_FLAT_ALIASES
                ):
                    violations.append(
                        "focused litchi-keynote show-settings public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and show_owner_declaration
                    and declaration_identifiers[:2]
                    in (["pub", "type"], ["pub", "use"])
                    and identifier in KEYNOTE_SHOW_SETTINGS_FLAT_SEMANTIC_ALIASES
                ):
                    violations.append(
                        "focused litchi-keynote show-settings public API "
                        f"retains flat semantic alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _iwork_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-keynote show-settings public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-keynote show-settings public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_keynote_soundtrack_settings_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the direct, archive-free Keynote soundtrack-settings API."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in KEYNOTE_SOUNDTRACK_SETTINGS_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, KEYNOTE_SOUNDTRACK_SETTINGS_SHORT_NAMES
    )
    for name in KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-keynote soundtrack settings public API is missing "
            f"canonical soundtrack type {name}: "
            f"{KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE}"
        )

    semantic_code = _mask_rust_non_code(semantic_source)
    for match in PUBLIC_KEYNOTE_SOUNDTRACK_TRANSACTION_MODULE.finditer(semantic_code):
        line_number = semantic_code.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-keynote soundtrack settings public API exposes "
            "duplicate soundtrack::transaction module: "
            f"{KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE}:{line_number}"
        )

    lib_export = root / KEYNOTE_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_export.read_text(encoding="utf-8"))
        if lib_export.is_file()
        else ""
    )
    if PUBLIC_KEYNOTE_SOUNDTRACK_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-keynote soundtrack settings public API is missing "
            "canonical root soundtrack module: "
            f"{KEYNOTE_SOURCE_ROOT / 'lib.rs'}"
        )

    package_export = root / KEYNOTE_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        if KEYNOTE_PACKAGE_SOUNDTRACK_SETTINGS_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-keynote soundtrack settings public API is missing "
                "private package owner module: "
                f"{package_export.relative_to(root)}"
            )
        for match in PUBLIC_KEYNOTE_PACKAGE_SOUNDTRACK_SETTINGS_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote soundtrack settings public API exposes "
                "duplicate package::soundtrack_settings module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-keynote soundtrack settings public API is missing "
            "private package owner module: "
            f"{KEYNOTE_SOURCE_ROOT / 'package.rs'}"
        )

    owner_path = root / KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-keynote soundtrack settings public API is missing "
            "private package owner source: "
            f"{KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE}"
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
            if not _is_keynote_soundtrack_settings_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _keynote_soundtrack_settings_owner_declaration(
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
                    "focused litchi-keynote soundtrack settings public API "
                    "retains root aliases via soundtrack glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if (
                public_declaration
                and path in export_sources
                and public_use_or_type
                and any(
                    identifier in {"soundtrack", "soundtrack_settings"}
                    for identifier in declaration_identifiers
                )
            ):
                violations.append(
                    "focused litchi-keynote soundtrack settings public API exposes "
                    "public soundtrack owner alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    public_declaration
                    and identifier in KEYNOTE_SOUNDTRACK_SETTINGS_FLAT_ALIASES
                ):
                    violations.append(
                        "focused litchi-keynote soundtrack settings public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in KEYNOTE_SOUNDTRACK_SETTINGS_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-keynote soundtrack settings public API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and identifier
                    in KEYNOTE_SOUNDTRACK_SETTINGS_FORBIDDEN_PUBLIC_MEMBERS
                ):
                    violations.append(
                        "focused litchi-keynote soundtrack settings public API "
                        f"retains host-style public member {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _keynote_soundtrack_settings_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-keynote soundtrack settings public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-keynote soundtrack settings public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def _keynote_soundtrack_items_public_leak(identifier: str) -> str | None:
    """Classify physical vocabulary forbidden in soundtrack item values."""

    if identifier in KEYNOTE_SOUNDTRACK_ITEMS_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in KEYNOTE_SOUNDTRACK_ITEMS_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in KEYNOTE_SOUNDTRACK_ITEMS_WIRE_TYPES:
        return "wire type"
    if identifier in KEYNOTE_SOUNDTRACK_ITEMS_TOPOLOGY_NAMES:
        return "soundtrack media topology"
    return _iwork_public_leak(identifier)


def audit_keynote_soundtrack_items_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the focused semantic soundtrack-item owner."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    violations: list[str] = []
    semantic_path = root / KEYNOTE_SOUNDTRACK_ITEMS_SEMANTIC_SOURCE
    owner_path = root / KEYNOTE_SOUNDTRACK_ITEMS_OWNER_SOURCE
    package_path = root / (KEYNOTE_SOURCE_ROOT / "package.rs")
    soundtrack_path = root / (KEYNOTE_SOURCE_ROOT / "soundtrack.rs")
    lib_path = root / (KEYNOTE_SOURCE_ROOT / "lib.rs")

    if not semantic_path.is_file():
        violations.append(
            "focused litchi-keynote soundtrack items public API is missing "
            f"semantic source: {KEYNOTE_SOUNDTRACK_ITEMS_SEMANTIC_SOURCE}"
        )
        return violations
    if not owner_path.is_file():
        violations.append(
            "focused litchi-keynote soundtrack items public API is missing "
            f"private package owner source: {KEYNOTE_SOUNDTRACK_ITEMS_OWNER_SOURCE}"
        )
    if not package_path.is_file():
        violations.append(
            "focused litchi-keynote soundtrack items public API is missing "
            f"package owner module source: {KEYNOTE_SOURCE_ROOT / 'package.rs'}"
        )

    semantic_source = semantic_path.read_text(encoding="utf-8")
    semantic_exports = _rust_canonical_exports(
        semantic_source, KEYNOTE_SOUNDTRACK_ITEMS_SHORT_NAMES
    )
    for name in KEYNOTE_SOUNDTRACK_ITEMS_CANONICAL_TYPES:
        if name in semantic_exports:
            continue
        violations.append(
            "focused litchi-keynote soundtrack items public API is missing "
            f"canonical soundtrack item type {name}: "
            f"{KEYNOTE_SOUNDTRACK_ITEMS_SEMANTIC_SOURCE}"
        )

    if soundtrack_path.is_file():
        soundtrack_source = _mask_rust_non_code(
            soundtrack_path.read_text(encoding="utf-8")
        )
        if KEYNOTE_SOUNDTRACK_ITEMS_MODULE.search(soundtrack_source) is None:
            violations.append(
                "focused litchi-keynote soundtrack items public API is missing "
                "public soundtrack::items module: "
                f"{KEYNOTE_SOURCE_ROOT / 'soundtrack.rs'}"
            )
    else:
        violations.append(
            "focused litchi-keynote soundtrack items public API is missing "
            f"semantic soundtrack module source: {KEYNOTE_SOURCE_ROOT / 'soundtrack.rs'}"
        )

    if package_path.is_file():
        package_source = _mask_rust_non_code(
            package_path.read_text(encoding="utf-8")
        )
        if KEYNOTE_PACKAGE_SOUNDTRACK_ITEMS_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-keynote soundtrack items public API is missing "
                "private package soundtrack_items module: "
                f"{KEYNOTE_SOURCE_ROOT / 'package.rs'}"
            )
        public_module = re.compile(
            r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?soundtrack_items\b",
            re.MULTILINE,
        )
        for match in public_module.finditer(package_source):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote soundtrack items public API exposes "
                "package::soundtrack_items module: "
                f"{KEYNOTE_SOURCE_ROOT / 'package.rs'}:{line_number}"
            )

    if lib_path.is_file():
        lib_source = _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if re.search(
            r"^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?soundtrack\b",
            lib_source,
            re.MULTILINE,
        ) is None:
            violations.append(
                "focused litchi-keynote soundtrack items public API is missing "
                f"root soundtrack module: {KEYNOTE_SOURCE_ROOT / 'lib.rs'}"
            )

    implementation_sources = [semantic_path]
    if owner_path.is_file():
        implementation_sources.append(owner_path)
    declarations: list[tuple[Path, str, int]] = []
    for path in implementation_sources:
        source = path.read_text(encoding="utf-8")
        declarations.extend(
            (path, declaration, line_number)
            for declaration, line_number in _rust_public_declarations(source)
        )
    for path, declaration, line_number in declarations:
        for match in RUST_IDENTIFIER.finditer(declaration):
            identifier = match.group(1)
            reason = _keynote_soundtrack_items_public_leak(identifier)
            if reason is None:
                continue
            identifier_line = line_number + declaration.count(
                "\n", 0, match.start(1)
            )
            violations.append(
                "focused litchi-keynote soundtrack items public API exposes "
                f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
            )
        for match in RUST_BYTE_SLICE.finditer(declaration):
            method = re.search(
                r"\bfn[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)\b",
                declaration,
            )
            if method is not None and method.group(1) in (
                KEYNOTE_SOUNDTRACK_ITEMS_ALLOWED_RAW_SLICE_METHODS
            ):
                continue
            byte_slice = re.sub(r"\s+", "", match.group(0))
            byte_slice_line = line_number + declaration.count(
                "\n", 0, match.start()
            )
            violations.append(
                "focused litchi-keynote soundtrack items public API exposes "
                f"raw byte slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
            )

    method_sources = [package_path]
    if owner_path.is_file():
        method_sources.append(owner_path)
    methods: set[str] = set()
    for path in method_sources:
        methods.update(
            name for name, _line_number in _rust_function_declarations(
                path.read_text(encoding="utf-8")
            )
        )
    for method in KEYNOTE_SOUNDTRACK_ITEMS_PACKAGE_METHODS:
        if method in methods:
            continue
        violations.append(
            "focused litchi-keynote soundtrack items public API is missing "
            f"Package::{method}: {KEYNOTE_SOURCE_ROOT / 'package.rs'}"
        )

    return sorted(set(violations))


def audit_keynote_slide_transition_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the canonical nested, archive-free slide-transition API."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    export_sources = {
        root / path
        for path in KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES
        if (root / path).is_file()
    }
    semantic_source = root / KEYNOTE_SOURCE_ROOT / "transition.rs"
    violations: list[str] = []

    canonical_exports = (
        _rust_canonical_exports(
            semantic_source.read_text(encoding="utf-8"),
            KEYNOTE_SLIDE_TRANSITION_SHORT_NAMES,
        )
        if semantic_source.is_file()
        else frozenset()
    )
    for name in KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-keynote slide-transition public API is missing "
            f"canonical transition type {name}: "
            f"{KEYNOTE_SOURCE_ROOT / 'transition.rs'}"
        )

    lib_export = root / KEYNOTE_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_export.read_text(encoding="utf-8"))
        if lib_export.is_file()
        else ""
    )
    if PUBLIC_KEYNOTE_TRANSITION_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-keynote slide-transition public API is missing "
            "canonical root transition module: "
            f"{KEYNOTE_SOURCE_ROOT / 'lib.rs'}"
        )

    package_export = root / KEYNOTE_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        for match in PUBLIC_KEYNOTE_PACKAGE_SLIDE_TRANSITION_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote slide-transition public API exposes duplicate "
                "package::slide_transition module: "
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
            if not _is_keynote_slide_transition_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _keynote_slide_transition_owner_declaration(
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
                    "focused litchi-keynote slide-transition public API retains "
                    "root aliases via transition glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    public_declaration
                    and identifier in KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES
                ):
                    violations.append(
                        "focused litchi-keynote slide-transition public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and identifier
                    in KEYNOTE_SLIDE_TRANSITION_FORBIDDEN_PUBLIC_MEMBERS
                ):
                    violations.append(
                        "focused litchi-keynote slide-transition public API "
                        f"retains adapter-only public member {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in KEYNOTE_SLIDE_TRANSITION_ROOT_ALIASES
                ):
                    violations.append(
                        "focused litchi-keynote slide-transition public API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _keynote_slide_transition_public_leak(
                    identifier, semantic_source=path == semantic_source
                )
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-keynote slide-transition public API exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            allow_opaque_payload = path == semantic_source and bool(
                set(declaration_identifiers)
                & KEYNOTE_SLIDE_TRANSITION_SEMANTIC_OPAQUE_PAYLOAD_MEMBERS
            )
            if allow_opaque_payload:
                continue
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-keynote slide-transition public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_keynote_slide_background_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce selector-first focused Keynote slide-background ownership."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    violations: list[str] = []
    package_path = root / KEYNOTE_SLIDE_BACKGROUND_EXPORT_SOURCES[1]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    if not package_path.is_file():
        violations.append(
            "focused litchi-keynote slide-background public API is missing "
            f"package owner module: {KEYNOTE_SLIDE_BACKGROUND_EXPORT_SOURCES[1]}"
        )
    elif KEYNOTE_PACKAGE_SLIDE_BACKGROUND_MODULE.search(package_source) is None:
        violations.append(
            "focused litchi-keynote slide-background public API is missing "
            f"private package owner module: {KEYNOTE_SLIDE_BACKGROUND_EXPORT_SOURCES[1]}"
        )
    for match in PUBLIC_KEYNOTE_PACKAGE_SLIDE_BACKGROUND_MODULE.finditer(
        package_source
    ):
        line_number = package_source.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-keynote slide-background public API exposes duplicate "
            "package::slide_background module: "
            f"{KEYNOTE_SLIDE_BACKGROUND_EXPORT_SOURCES[1]}:{line_number}"
        )

    owner_path = root / KEYNOTE_SLIDE_BACKGROUND_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-keynote slide-background public API is missing private "
            f"package owner source: {KEYNOTE_SLIDE_BACKGROUND_OWNER_SOURCE}"
        )
        owner_source = ""
    else:
        owner_source = owner_path.read_text(encoding="utf-8")

    owner_public_methods = {
        name
        for declaration, _line_number in _rust_public_declarations(owner_source)
        for name, _nested_line in _rust_function_declarations(declaration)
    }
    for method in KEYNOTE_SLIDE_BACKGROUND_PACKAGE_METHODS:
        if method in owner_public_methods:
            continue
        violations.append(
            "focused litchi-keynote slide-background public API is missing Package "
            f"method {method}: {KEYNOTE_SLIDE_BACKGROUND_OWNER_SOURCE}"
        )

    for method in KEYNOTE_SLIDE_BACKGROUND_EDIT_METHODS:
        if method in owner_public_methods:
            continue
        violations.append(
            "focused litchi-keynote slide-background public API is missing "
            f"Edit method {method}: {KEYNOTE_SLIDE_BACKGROUND_OWNER_SOURCE}"
        )

    # Keep the owner and semantic value source free of generated Prost reads;
    # cfg(test) differential helpers remain legal through item masking.
    for path in (
        root / KEYNOTE_SLIDE_BACKGROUND_IMPLEMENTATION_SOURCES[0],
        owner_path,
    ):
        if not path.is_file():
            continue
        raw_source = path.read_text(encoding="utf-8")
        production_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(raw_source)
        )
        for label, pattern in KEYNOTE_NO_EAGER_PROST_SOURCE_PATTERNS:
            for match in pattern.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused litchi-keynote slide-background production source uses "
                    f"{label}: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_keynote_slide_delete_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the nested, selector-first, archive-free slide-delete API."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in KEYNOTE_SLIDE_DELETE_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    for helper_root in (
        root / KEYNOTE_SOURCE_ROOT / "slide" / "delete",
        root / KEYNOTE_SOURCE_ROOT / "package" / "slide_delete",
    ):
        if helper_root.is_dir():
            dedicated_sources.update(helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in KEYNOTE_SLIDE_DELETE_EXPORT_SOURCES
        if (root / path).is_file()
    }
    semantic_path = root / KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE
    owner_path = root / KEYNOTE_SLIDE_DELETE_OWNER_SOURCE
    violations: list[str] = []

    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source, KEYNOTE_SLIDE_DELETE_SHORT_NAMES
    )
    for name in KEYNOTE_SLIDE_DELETE_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-keynote slide-delete public API is missing "
            f"canonical slide::delete type {name}: "
            f"{KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE}"
        )

    lib_path = root / KEYNOTE_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    if PUBLIC_KEYNOTE_SLIDE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-keynote slide-delete public API is missing "
            f"canonical root slide module: {KEYNOTE_SOURCE_ROOT / 'lib.rs'}"
        )

    slide_path = root / KEYNOTE_SOURCE_ROOT / "slide.rs"
    slide_source = (
        _mask_rust_non_code(slide_path.read_text(encoding="utf-8"))
        if slide_path.is_file()
        else ""
    )
    if PUBLIC_KEYNOTE_SLIDE_DELETE_MODULE.search(slide_source) is None:
        violations.append(
            "focused litchi-keynote slide-delete public API is missing "
            f"canonical slide::delete module: {KEYNOTE_SOURCE_ROOT / 'slide.rs'}"
        )

    package_path = root / KEYNOTE_SOURCE_ROOT / "package.rs"
    if package_path.is_file():
        package_source = _mask_rust_non_code(
            package_path.read_text(encoding="utf-8")
        )
        if KEYNOTE_PACKAGE_SLIDE_DELETE_MODULE.search(package_source) is None:
            violations.append(
                "focused litchi-keynote slide-delete public API is missing "
                f"private package owner module: {package_path.relative_to(root)}"
            )
        for match in PUBLIC_KEYNOTE_PACKAGE_SLIDE_DELETE_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote slide-delete public API exposes duplicate "
                "package::slide_delete module: "
                f"{package_path.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-keynote slide-delete public API is missing "
            f"private package owner module: {KEYNOTE_SOURCE_ROOT / 'package.rs'}"
        )

    if not owner_path.is_file():
        violations.append(
            "focused litchi-keynote slide-delete public API is missing private "
            f"package owner source: {KEYNOTE_SLIDE_DELETE_OWNER_SOURCE}"
        )

    owner_methods = {
        name
        for declaration, _line_number in _rust_public_declarations(
            owner_path.read_text(encoding="utf-8") if owner_path.is_file() else ""
        )
        for name, _nested_line in _rust_function_declarations(declaration)
    }
    for name in KEYNOTE_SLIDE_DELETE_PACKAGE_METHODS:
        if name in owner_methods:
            continue
        violations.append(
            "focused litchi-keynote slide-delete public API is missing canonical "
            f"Package::{name} method: {KEYNOTE_SLIDE_DELETE_OWNER_SOURCE}"
        )

    edit_methods = {
        name
        for implementation_path in dedicated_sources
        for declaration, _line_number in _rust_public_declarations(
            implementation_path.read_text(encoding="utf-8")
        )
        for name, _nested_line in _rust_function_declarations(declaration)
    }
    for name in KEYNOTE_SLIDE_DELETE_EDIT_METHODS:
        if name in edit_methods:
            continue
        violations.append(
            "focused litchi-keynote slide-delete public API is missing canonical "
            f"Edit::{name} method: {KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE}"
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
            if not _is_keynote_slide_delete_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _keynote_slide_delete_owner_declaration(declaration)
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            public_use_or_type = identifiers[:2] in (["pub", "type"], ["pub", "use"])
            flat_alias_exports = _rust_canonical_exports(
                declaration, KEYNOTE_SLIDE_DELETE_FLAT_ALIASES
            )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and public_use_or_type
            ):
                violations.append(
                    "focused litchi-keynote slide-delete public API exposes public "
                    f"slide-delete owner alias: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if identifier in flat_alias_exports:
                    violations.append(
                        "focused litchi-keynote slide-delete public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in KEYNOTE_SLIDE_DELETE_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-keynote slide-delete public API retains root "
                        f"alias {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                reason = _keynote_slide_delete_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-keynote slide-delete public API exposes "
                    f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-keynote slide-delete public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def audit_keynote_placeholder_visibility_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the nested, archive-free Keynote placeholder visibility API."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    dedicated_sources = {
        root / path
        for path in KEYNOTE_PLACEHOLDER_VISIBILITY_IMPLEMENTATION_SOURCES
        if (root / path).is_file()
    }
    owner_helper_root = root / KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT
    if owner_helper_root.is_dir():
        dedicated_sources.update(owner_helper_root.rglob("*.rs"))
    preview_helper_root = root / KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_ROOT
    if preview_helper_root.is_dir():
        dedicated_sources.update(preview_helper_root.rglob("*.rs"))
    export_sources = {
        root / path
        for path in KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES
        if (root / path).is_file()
    }
    violations: list[str] = []

    semantic_path = root / KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE
    semantic_source = (
        semantic_path.read_text(encoding="utf-8")
        if semantic_path.is_file()
        else ""
    )
    canonical_exports = _rust_canonical_exports(
        semantic_source,
        KEYNOTE_PLACEHOLDER_VISIBILITY_SHORT_NAMES,
    )
    for name in KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES:
        if name in canonical_exports:
            continue
        violations.append(
            "focused litchi-keynote placeholder visibility public API is missing "
            f"canonical slide::placeholder type {name}: "
            f"{KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE}"
        )
    canonical_kinds = _rust_public_enum_variants(semantic_source, "Kind")
    if "Kind" in canonical_exports:
        for name in KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_KINDS:
            if name in canonical_kinds:
                continue
            violations.append(
                "focused litchi-keynote placeholder visibility public API is missing "
                f"canonical placeholder kind {name}: "
                f"{KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE}"
            )

    lib_export = root / KEYNOTE_SOURCE_ROOT / "lib.rs"
    lib_source = (
        _mask_rust_non_code(lib_export.read_text(encoding="utf-8"))
        if lib_export.is_file()
        else ""
    )
    if PUBLIC_KEYNOTE_SLIDE_MODULE.search(lib_source) is None:
        violations.append(
            "focused litchi-keynote placeholder visibility public API is missing "
            "canonical root slide module: "
            f"{KEYNOTE_SOURCE_ROOT / 'lib.rs'}"
        )

    slide_export = root / KEYNOTE_SOURCE_ROOT / "slide.rs"
    slide_source = (
        _mask_rust_non_code(slide_export.read_text(encoding="utf-8"))
        if slide_export.is_file()
        else ""
    )
    if PUBLIC_KEYNOTE_PLACEHOLDER_MODULE.search(slide_source) is None:
        violations.append(
            "focused litchi-keynote placeholder visibility public API is missing "
            "canonical slide::placeholder module: "
            f"{KEYNOTE_SOURCE_ROOT / 'slide.rs'}"
        )

    package_export = root / KEYNOTE_SOURCE_ROOT / "package.rs"
    if package_export.is_file():
        package_source = _mask_rust_non_code(
            package_export.read_text(encoding="utf-8")
        )
        if (
            KEYNOTE_PACKAGE_PLACEHOLDER_VISIBILITY_MODULE.search(package_source)
            is None
        ):
            violations.append(
                "focused litchi-keynote placeholder visibility public API is missing "
                "private package owner module: "
                f"{package_export.relative_to(root)}"
            )
        for match in PUBLIC_KEYNOTE_PACKAGE_PLACEHOLDER_VISIBILITY_MODULE.finditer(
            package_source
        ):
            line_number = package_source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote placeholder visibility public API exposes "
                "duplicate package::slide_placeholder_visibility module: "
                f"{package_export.relative_to(root)}:{line_number}"
            )
    else:
        violations.append(
            "focused litchi-keynote placeholder visibility public API is missing "
            "private package owner module: "
            f"{KEYNOTE_SOURCE_ROOT / 'package.rs'}"
        )

    owner_path = root / KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE
    if not owner_path.is_file():
        violations.append(
            "focused litchi-keynote placeholder visibility public API is missing "
            "private package owner source: "
            f"{KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE}"
        )

    for module_source in (
        KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE,
        KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE,
        KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_PREVIEW_SOURCE,
        *KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES,
    ):
        module_path = root / module_source
        if not module_path.is_file():
            continue
        module_text = _mask_rust_non_code(
            module_path.read_text(encoding="utf-8")
        )
        for match in PUBLIC_KEYNOTE_SLIDE_NUMBER_HELPER_MODULE.finditer(
            module_text
        ):
            line_number = module_text.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote placeholder visibility public API exposes "
                "public slide-number helper module: "
                f"{module_source}:{line_number}"
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
            if not _is_keynote_placeholder_visibility_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
                continue
            owner_declaration = _keynote_placeholder_visibility_owner_declaration(
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
                and path in export_sources | {semantic_path}
                and public_use_or_type
                and "slide_number" in declaration_identifiers
            ):
                violations.append(
                    "focused litchi-keynote placeholder visibility public API exposes "
                    "public slide-number owner alias: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if (
                public_declaration
                and path in export_sources
                and owner_declaration
                and declaration_identifiers[:2] == ["pub", "use"]
                and "*" in declaration
            ):
                violations.append(
                    "focused litchi-keynote placeholder visibility public API "
                    "retains root aliases via slide::placeholder glob: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                if (
                    public_declaration
                    and identifier in KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES
                ):
                    violations.append(
                        "focused litchi-keynote placeholder visibility public API "
                        f"retains flat alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and identifier
                    in KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_SEMANTIC_ALIASES
                ):
                    violations.append(
                        "focused litchi-keynote placeholder visibility public API "
                        f"retains flat semantic alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and identifier
                    in KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_NUMBER_PUBLIC_MEMBERS
                ):
                    violations.append(
                        "focused litchi-keynote placeholder visibility public API "
                        f"retains slide-number-specific public member {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                if (
                    public_declaration
                    and path in export_sources
                    and owner_declaration
                    and public_use_or_type
                    and identifier in KEYNOTE_PLACEHOLDER_VISIBILITY_SHORT_NAMES
                ):
                    violations.append(
                        "focused litchi-keynote placeholder visibility public API "
                        f"retains root alias {identifier}: "
                        f"{path.relative_to(root)}:{identifier_line}"
                    )
                reason = _keynote_placeholder_visibility_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-keynote placeholder visibility public API "
                    f"exposes {reason} {identifier}: "
                    f"{path.relative_to(root)}:{identifier_line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                byte_slice_line = line_number + declaration.count(
                    "\n", 0, match.start()
                )
                violations.append(
                    "focused litchi-keynote placeholder visibility public API "
                    f"exposes raw byte slice {byte_slice}: "
                    f"{path.relative_to(root)}:{byte_slice_line}"
                )

    return sorted(set(violations))


def _keynote_document_public_leak(identifier: str) -> str | None:
    """Classify physical vocabulary forbidden by the archive-free reader."""

    if identifier in KEYNOTE_DOCUMENT_SEMANTIC_LIMIT_PARAMETERS:
        return None
    if identifier == "wire" or identifier.startswith("Wire") or identifier.endswith("Wire"):
        return "wire type"
    return _iwork_public_leak(identifier)


def audit_keynote_document_public_api(root: Path = ROOT) -> list[str]:
    """Keep the focused Keynote document reader archive-free and semantic."""

    violations: list[str] = []
    for relative in KEYNOTE_DOCUMENT_PUBLIC_API_SOURCES:
        path = root / relative
        if not path.is_file():
            continue
        source = path.read_text(encoding="utf-8")
        for declaration, line_number in _rust_public_declarations(source):
            for match in RUST_IDENTIFIER.finditer(declaration):
                identifier = match.group(1)
                reason = _keynote_document_public_leak(identifier)
                if reason is None:
                    continue
                identifier_line = line_number + declaration.count(
                    "\n", 0, match.start(1)
                )
                violations.append(
                    "focused litchi-keynote document reader public API exposes "
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
                if not identifier[:1].isupper() or identifier == "Archive":
                    continue
                reason = _keynote_document_public_leak(identifier)
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
                    "focused litchi-keynote document reader rustdoc exposes "
                    f"{reason} {identifier}: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _keynote_slide_table_appearance_owner_present(root: Path) -> bool:
    """Return whether the complete slide-table appearance owner is wired."""

    owner_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_appearance\s*;",
        package_source,
    ) is not None


def _keynote_slide_table_appearance_public_leak(identifier: str) -> str | None:
    """Classify physical/wire vocabulary forbidden by the appearance facade."""

    if identifier == "litchi_iwa_common":
        return None
    if identifier in KEYNOTE_SLIDE_TABLE_APPEARANCE_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in KEYNOTE_SLIDE_TABLE_APPEARANCE_PHYSICAL_TYPES:
        return "archive/IWA type"
    if identifier == "wire" or identifier in KEYNOTE_SLIDE_TABLE_APPEARANCE_WIRE_TYPES:
        return "wire type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(word in {"buffa", "prost"} for word in words):
        return "protobuf type"
    return _iwork_public_leak(identifier)


def _keynote_slide_table_appearance_public_declaration(
    declaration: str,
    *,
    dedicated_source: bool,
) -> bool:
    """Select appearance declarations without inspecting unrelated APIs."""

    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    return bool(
        identifiers
        & (
            (KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES - KEYNOTE_FAMILY_GENERIC_NAMES)
            | KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_TYPES
            | KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_TYPES
            | KEYNOTE_SLIDE_TABLE_APPEARANCE_PACKAGE_METHODS
            | KEYNOTE_SLIDE_TABLE_APPEARANCE_FLAT_ALIASES
        )
    ) or (
        "appearance" in identifiers
        and ("table" in identifiers or "slide" in identifiers)
    )


def _keynote_slide_table_appearance_check_codec(root: Path) -> list[str]:
    """Require the shared strict, prepared appearance codec and its seams."""

    violations: list[str] = []
    codec_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_SOURCE
    if not codec_path.is_file():
        return [
            "focused litchi-keynote slide-table appearance public API is missing "
            f"strict neutral codec source: {KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_SOURCE}"
        ]
    codec_raw_source = codec_path.read_text(encoding="utf-8")
    codec_source = _mask_rust_cfg_test_items(codec_raw_source)
    codec_code = _mask_rust_non_code(codec_source)
    for api in KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_REQUIRED_APIS:
        if re.search(
            rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
            codec_code,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-table appearance hidden codec is missing "
                f"strict API {api}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not all(marker in codec_code for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table appearance hidden codec is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_SOURCE}"
            )
    codec_lib_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE
    codec_lib_source = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    for codec_module in KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_MODULES:
        if re.search(
            rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
            rf"pub\s+mod\s+{re.escape(codec_module)}\b",
            codec_lib_source,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-table appearance public API is missing hidden "
                f"codec module {codec_module}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_PUBLIC_SOURCE}"
            )
    if re.search(
        r"(?m)#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\]",
        codec_raw_source,
    ) is None or re.search(r"(?m)#\s*\[\s*test\s*\]", codec_raw_source) is None:
        violations.append(
            "focused litchi-keynote slide-table appearance hidden codec is missing cfg(test) #[test] coverage: "
            f"{KEYNOTE_SLIDE_TABLE_APPEARANCE_CODEC_SOURCE}"
        )
    for test_path in KEYNOTE_SLIDE_TABLE_APPEARANCE_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote slide-table appearance boundary is missing integration test: "
                f"{test_path}"
            )
            continue
        test_source = absolute.read_text(encoding="utf-8")
        test_code = _mask_rust_non_code(test_source)
        if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
            violations.append(
                "focused litchi-keynote slide-table appearance integration test is missing #[test] coverage: "
                f"{test_path}"
            )
        if not re.search(
            r"\b(?:slide_table_appearance|edit_slide_table_appearance|"
            r"apply_slide_table_appearance)\b",
            test_code,
        ):
            violations.append(
                "focused litchi-keynote slide-table appearance integration test does not exercise the owner: "
                f"{test_path}"
            )
    for fuzz_path in KEYNOTE_SLIDE_TABLE_APPEARANCE_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote slide-table appearance boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote slide-table appearance fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in KEYNOTE_SLIDE_TABLE_APPEARANCE_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-keynote slide-table appearance boundary is missing fuzz corpus: "
                f"{corpus}"
            )
    return sorted(set(violations))


def audit_keynote_slide_table_appearance_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Keynote appearance facade."""

    if not _keynote_slide_table_appearance_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
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
        r"(?m)^pub[ \t]+mod[ \t]+slide_table_appearance\b",
        package_code + code[lib_path],
    ):
        violations.append(
            "focused litchi-keynote slide-table appearance owner module must remain private: "
            f"{KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_appearance\s*;",
        package_code,
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table appearance owner module is missing: "
            f"{KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_APPEARANCE_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_APPEARANCE_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-table appearance public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_TYPES,
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_TYPES - semantic_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table appearance semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_SOURCE}"
        )
    if re.search(r"(?m)^pub[ \t]+mod[ \t]+transaction\b", code[semantic_path]) is None:
        violations.append(
            "focused litchi-keynote slide-table appearance semantic API is missing "
            f"slide::table::appearance::transaction: {KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_SLIDE_TABLE_APPEARANCE_TRANSACTION_TYPES
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_APPEARANCE_TRANSACTION_TYPES - transaction_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table appearance transaction API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_TYPES,
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_TYPES - selector_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table appearance public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_EXPORT_SOURCES[1]}"
        )
    if re.search(
        r"impl[ \t\r\n]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*)?"
        r"TableSelector\b[\s\S]*?\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+"
        r"(?:index|position)\b",
        code[selector_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table appearance selector must expose checked "
            f"TableSelector::index/position: {KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_SOURCE}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_APPEARANCE_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-table appearance Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_table_appearance":
            for selector in ("SlideSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-table appearance Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE}"
                    )
    apply_declaration = owner_methods.get("apply_slide_table_appearance")
    if apply_declaration is not None and "Patch" not in apply_declaration:
        violations.append(
            "focused litchi-keynote slide-table appearance apply method must accept "
            f"Patch: {KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE}"
        )

    edit_methods = {
        name
        for edit_type in ("Edit", "Edit")
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], edit_type
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_APPEARANCE_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote slide-table appearance edit is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_SLIDE_TABLE_APPEARANCE_CANONICAL_TYPES
        | KEYNOTE_SLIDE_TABLE_APPEARANCE_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_TABLE_APPEARANCE_SELECTOR_TYPES
        | KEYNOTE_SLIDE_TABLE_APPEARANCE_PACKAGE_METHODS
        | KEYNOTE_SLIDE_TABLE_APPEARANCE_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path}
    for path in paths:
        dedicated = path in dedicated_sources
        for declaration, line_number in _rust_public_declarations(sources[path]):
            if not _keynote_slide_table_appearance_public_declaration(
                declaration, dedicated_source=dedicated
            ):
                continue
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-table appearance public API retains a glob re-export: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for identifier in sorted(identifiers):
                reason = _keynote_slide_table_appearance_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-table appearance public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_APPEARANCE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-table appearance public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in KEYNOTE_SLIDE_TABLE_APPEARANCE_PUBLIC_RAW_ID_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote slide-table appearance public API exposes raw parameter "
                    f"{match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote slide-table appearance public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{line_number}"
                )

    violations.extend(_keynote_slide_table_appearance_check_codec(root))
    return sorted(set(violations))


def audit_keynote_slide_table_appearance_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one bounded style transaction and exact package preservation."""

    if not _keynote_slide_table_appearance_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner = _mask_rust_non_code(owner_source)
    violations: list[str] = []
    for label, markers in KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_MARKER_GROUPS.items():
        if not all(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table appearance owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE}"
            )
    if re.search(
        r"deleted_previews\s*:\s*(?:usize\s*=\s*)?0|deleted_previews\s*=\s*0",
        owner,
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table appearance owner must preserve previews "
            f"with deleted_previews = 0: {KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE}"
        )
    for label, pattern in KEYNOTE_SLIDE_TABLE_APPEARANCE_FORBIDDEN_OWNER_PATTERNS:
        match = pattern.search(owner)
        if match is not None:
            line_number = owner.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote slide-table appearance owner must not perform "
                f"{label}: {KEYNOTE_SLIDE_TABLE_APPEARANCE_OWNER_SOURCE}:{line_number}"
            )
    return sorted(set(violations))


def _audit_keynote_media_creation_projection(root: Path) -> list[str]:
    """Keep the fresh-media Buffa projection closed and wire-compatible."""

    projection_path = root / KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE
    if not projection_path.is_file():
        return [
            "focused Keynote media-creation codec is missing private Buffa "
            f"projection: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
        ]

    raw = projection_path.read_text(encoding="utf-8")
    projection = re.sub(r"//[^\n]*|/\*.*?\*/", "", raw, flags=re.DOTALL)
    violations: list[str] = []
    for marker in ('syntax = "proto2"', "package LitchiIwaProjection"):
        if marker not in projection:
            violations.append(
                "focused Keynote media-creation projection is missing marker "
                f"{marker}: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
            )
    if "repeated " in projection:
        violations.append(
            "focused Keynote media-creation projection must not materialize "
            f"repeated fields: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
        )
    if "extensions " in projection or "extend " in projection:
        violations.append(
            "focused Keynote media-creation projection must not widen the "
            f"native extension surface: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
        )
    if len(raw.encode("utf-8")) > 4 * 1024:
        violations.append(
            "focused Keynote media-creation projection exceeds its 4 KiB source "
            f"budget: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
        )

    message_names = set(
        re.findall(r"\bmessage\s+([A-Za-z_][A-Za-z0-9_]*)\s*\{", projection)
    )
    expected_names = set(KEYNOTE_MEDIA_CREATION_PROJECTION_FIELDS)
    for unexpected in sorted(message_names - expected_names):
        violations.append(
            "focused Keynote media-creation projection has unexpected message "
            f"{unexpected}: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
        )
    for missing in sorted(expected_names - message_names):
        violations.append(
            "focused Keynote media-creation projection is missing message "
            f"{missing}: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
        )

    field_pattern = re.compile(
        r"(?m)^\s*(?P<label>required|optional|repeated)\s+"
        r"(?P<type>[.A-Za-z_][.A-Za-z0-9_.]*)\s+"
        r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)\s*=\s*"
        r"(?P<number>[0-9]+)\s*(?:\[[^]]*\])?\s*;"
    )
    for message, expected_fields in KEYNOTE_MEDIA_CREATION_PROJECTION_FIELDS.items():
        body = _proto_message_body(projection, message)
        if body is None:
            continue
        matches = list(field_pattern.finditer(body))
        actual: dict[str, tuple[str, str, int]] = {}
        for field in matches:
            field_type = field.group("type").rsplit(".", 1)[-1]
            actual[field.group("name")] = (
                field.group("label"),
                field_type,
                int(field.group("number")),
            )
        if len(actual) != len(matches):
            violations.append(
                "focused Keynote media-creation projection repeats a field "
                f"name in {message}: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
            )
        for field, expected in expected_fields.items():
            if actual.get(field) != expected:
                violations.append(
                    "focused Keynote media-creation projection has incorrect or "
                    f"missing field {message}.{field}={expected[2]}: "
                    f"{KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
                )
        for unexpected in sorted(actual.keys() - expected_fields.keys()):
            violations.append(
                "focused Keynote media-creation projection has unexpected field "
                f"{message}.{unexpected}: {KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE}"
            )

    return sorted(set(violations))


def audit_keynote_media_creation_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the fresh-media codec, schema, and Buffa build route private."""

    codec_path = root / KEYNOTE_MEDIA_CREATION_CODEC_SOURCE
    projection_path = root / KEYNOTE_MEDIA_CREATION_PROJECTION_SOURCE
    if not codec_path.is_file():
        return []

    violations: list[str] = []
    if not codec_path.is_file():
        violations.append(
            "focused Keynote media-creation codec source is missing: "
            f"{KEYNOTE_MEDIA_CREATION_CODEC_SOURCE}"
        )
    else:
        codec_source = _mask_rust_cfg_test_items(
            codec_path.read_text(encoding="utf-8")
        )
        codec = _mask_rust_non_code(codec_source)
        for marker in KEYNOTE_MEDIA_CREATION_CODEC_REQUIRED_MARKERS:
            if marker not in codec:
                violations.append(
                    "focused Keynote media-creation codec is missing marker "
                    f"{marker}: {KEYNOTE_MEDIA_CREATION_CODEC_SOURCE}"
                )
        for operation in KEYNOTE_MEDIA_CREATION_CODEC_FORBIDDEN_OPERATIONS:
            if operation in codec:
                violations.append(
                    "focused Keynote media-creation codec retains forbidden "
                    f"operation {operation}: {KEYNOTE_MEDIA_CREATION_CODEC_SOURCE}"
                )

    violations.extend(_audit_keynote_media_creation_projection(root))

    build_path = root / KEYNOTE_MEDIA_CREATION_BUILD_SOURCE
    if not build_path.is_file():
        violations.append(
            "focused Keynote media-creation Buffa build source is missing: "
            f"{KEYNOTE_MEDIA_CREATION_BUILD_SOURCE}"
        )
    else:
        build = build_path.read_text(encoding="utf-8")
        route = re.search(
            r"(?s)\blet\s+buffa_keynote_media_creation_out_directory\b.*?"
            r"\benforce_keynote_media_creation_projection_budget\s*\(\s*"
            r"&\s*buffa_keynote_media_creation_out_directory\s*,?\s*\)",
            build,
        )
        if route is None:
            violations.append(
                "focused Keynote media-creation Buffa build is missing its "
                f"dedicated route: {KEYNOTE_MEDIA_CREATION_BUILD_SOURCE}"
            )
            build_focus = ""
        else:
            build_focus = route.group(0)
        for marker in (
            "TSDKeynoteMediaCreationArchive.proto",
            "buffa-keynote-media-creation",
            "iwa_keynote_media_creation_buffa_protos.rs",
            ".includes(&[buffa_projection_directory])",
            ".generate_views(true)",
            ".lazy_views(true)",
            ".preserve_unknown_fields(false)",
            ".generate_json(false)",
            ".generate_text(false)",
            ".reflect_mode(buffa_build::ReflectMode::Off)",
            ".idiomatic_field_names(true)",
            ".compile()",
            "enforce_keynote_media_creation_projection_budget",
        ):
            if marker not in build_focus:
                violations.append(
                    "focused Keynote media-creation Buffa build is missing marker "
                    f"{marker}: {KEYNOTE_MEDIA_CREATION_BUILD_SOURCE}"
                )
        for marker in (
            "fn enforce_keynote_media_creation_projection_provenance(",
            "enforce_keynote_media_creation_projection_provenance(\n"
            "        proto_directory,",
        ):
            if marker not in build:
                violations.append(
                    "focused Keynote media-creation build is missing provenance "
                    f"marker {marker}: {KEYNOTE_MEDIA_CREATION_BUILD_SOURCE}"
                )

    public_path = root / KEYNOTE_MEDIA_CREATION_CODEC_PUBLIC_SOURCE
    if not public_path.is_file():
        violations.append(
            "focused Keynote media-creation generated module owner is missing: "
            f"{KEYNOTE_MEDIA_CREATION_CODEC_PUBLIC_SOURCE}"
        )
    else:
        public_source = _mask_rust_cfg_test_items(
            public_path.read_text(encoding="utf-8")
        )
        if re.search(
            rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
            rf"pub\s+mod\s+{re.escape(KEYNOTE_MEDIA_CREATION_CODEC_MODULE)}\b",
            public_source,
        ) is None:
            violations.append(
                "focused Keynote media-creation codec is missing its hidden "
                f"public module {KEYNOTE_MEDIA_CREATION_CODEC_MODULE}: "
                f"{KEYNOTE_MEDIA_CREATION_CODEC_PUBLIC_SOURCE}"
            )
        if re.search(
            rf"(?ms)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\].{{0,260}}?"
            rf"^\s*mod\s+{re.escape(KEYNOTE_MEDIA_CREATION_GENERATED_MODULE)}"
            rf"\s*\{{",
            public_source,
        ) is None:
            violations.append(
                "focused Keynote media-creation generated Buffa view must remain "
                f"private: {KEYNOTE_MEDIA_CREATION_CODEC_PUBLIC_SOURCE}"
            )
        if KEYNOTE_MEDIA_CREATION_GENERATED_INCLUDE not in public_source:
            violations.append(
                "focused Keynote media-creation generated Buffa include is "
                f"missing: {KEYNOTE_MEDIA_CREATION_CODEC_PUBLIC_SOURCE}"
            )

    return sorted(set(violations))


def audit_keynote_slide_audio_creation_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the focused slide-audio Package owner generated-free.

    ``slide_audio_creation`` still manipulates physical archive records, so a
    broad ban on ``litchi_iwa_core`` or archive types would reject its valid
    transaction layer.  This source-local scan blocks only direct generated
    schema imports/uses, Prost traits, generated archive values, and eager
    protobuf reads/encodes in the owner and its private planning children.
    """

    owner_path = root / KEYNOTE_SLIDE_AUDIO_CREATION_OWNER_SOURCE
    child_root = root / KEYNOTE_SLIDE_AUDIO_CREATION_CHILD_ROOT
    paths = ([owner_path] if owner_path.is_file() else []) + (
        sorted(child_root.rglob("*.rs")) if child_root.is_dir() else []
    )
    if not paths:
        return []

    violations: list[str] = []
    for path in paths:
        if path.name in KEYNOTE_TEST_ONLY_SOURCE_NAMES:
            continue
        raw_source = path.read_text(encoding="utf-8")
        production_source = _mask_rust_cfg_test_items(raw_source)
        production_code = _mask_rust_non_code(production_source)
        for label, pattern in KEYNOTE_SLIDE_AUDIO_CREATION_FORBIDDEN_SOURCE_PATTERNS:
            for match in pattern.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused litchi-keynote slide-audio creation production "
                    f"source uses {label}: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _keynote_slide_movie_creation_owner_paths(root: Path) -> list[Path]:
    """Return the shared movie-creation owner and its private Rust children."""

    owner = root / KEYNOTE_SLIDE_MOVIE_CREATION_OWNER_SOURCE
    shared_owner = root / KEYNOTE_SLIDE_MOVIE_CREATION_SHARED_OWNER_SOURCE
    child_root = root / KEYNOTE_SLIDE_MOVIE_CREATION_CHILD_ROOT
    paths = [
        path
        for path in (owner, shared_owner)
        if path.is_file()
    ] + (
        sorted(child_root.rglob("*.rs")) if child_root.is_dir() else []
    )
    return list(dict.fromkeys(paths))


def _keynote_slide_movie_creation_route_source(root: Path) -> str:
    """Return masked focused sources when the shared owner has a movie route."""

    paths = _keynote_slide_movie_creation_owner_paths(root)
    if not paths:
        return ""
    source = "\n".join(
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for path in paths
        if path.name not in KEYNOTE_TEST_ONLY_SOURCE_NAMES
    )
    return source if KEYNOTE_SLIDE_MOVIE_CREATION_ROUTE.search(source) else ""


def audit_keynote_slide_movie_creation_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the focused file-movie route typed and neutral.

    The shared audio owner is allowed to contain physical archive transaction
    code.  This operation-specific audit activates only after an explicit
    ``CreationOptions::Movie`` (or equivalent typed enum) route appears, then
    requires the bounded neutral media codec and rejects generated schemas,
    eager protobuf operations, and direct legacy-editor delegation.  Raw
    native identifiers are checked only on the public movie-creation function
    signature so private archive bookkeeping remains available to the owner.
    """

    route_source = _keynote_slide_movie_creation_route_source(root)
    if not route_source:
        return []

    violations: list[str] = []
    if KEYNOTE_SLIDE_MOVIE_CREATION_CODEC_ROUTE.search(route_source) is None:
        violations.append(
            "focused litchi-keynote file-movie creation must route payloads through "
            f"the bounded neutral media codec: {KEYNOTE_SLIDE_MOVIE_CREATION_OWNER_SOURCE}"
        )
    if re.search(r"\b(?:try_encoded_len|try_encode_bounded|try_reserve_exact)\b", route_source) is None:
        violations.append(
            "focused litchi-keynote file-movie creation is missing bounded encoding "
            f"preflight: {KEYNOTE_SLIDE_MOVIE_CREATION_OWNER_SOURCE}"
        )

    for path in _keynote_slide_movie_creation_owner_paths(root):
        if path.name in KEYNOTE_TEST_ONLY_SOURCE_NAMES:
            continue
        raw_source = path.read_text(encoding="utf-8")
        production_source = _mask_rust_cfg_test_items(raw_source)
        production_code = _mask_rust_non_code(production_source)
        for label, pattern in KEYNOTE_SLIDE_MOVIE_CREATION_FORBIDDEN_SOURCE_PATTERNS:
            for match in pattern.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused litchi-keynote file-movie creation production source "
                    f"uses {label}: {path.relative_to(root)}:{line_number}"
                )

        # Keep raw-ID rejection local to a semantic entry point.  The shared
        # owner necessarily uses u64 references while staging the archive.
        function = _rust_named_function_body(production_source, "create_slide_movie")
        if function is not None:
            body, body_offset = function
            signature = production_code[:body_offset]
            signature_start = signature.rfind("fn create_slide_movie")
            if signature_start >= 0:
                signature = signature[max(0, signature_start - 240) :]
                for match in KEYNOTE_SLIDE_MOVIE_CREATION_RAW_API_PARAMETER.finditer(
                    signature
                ):
                    line_number = production_code.count(
                        "\n", 0, body_offset - len(signature) + match.start()
                    ) + 1
                    violations.append(
                        "focused litchi-keynote file-movie creation public API exposes "
                        f"a raw identifier parameter: {path.relative_to(root)}:{line_number}"
                    )
                for match in KEYNOTE_SLIDE_MOVIE_CREATION_RAW_API_TYPE.finditer(
                    signature
                ):
                    line_number = production_code.count(
                        "\n", 0, body_offset - len(signature) + match.start()
                    ) + 1
                    violations.append(
                        "focused litchi-keynote file-movie creation public API exposes "
                        f"a raw object type {match.group(0)}: "
                        f"{path.relative_to(root)}:{line_number}"
                    )

    return sorted(set(violations))


def audit_keynote_package_no_eager_prost_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep focused Keynote production ingress free of generated Prost reads."""

    violations: list[str] = []

    source_root = root / KEYNOTE_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            if path.name in KEYNOTE_TEST_ONLY_SOURCE_NAMES:
                continue
            raw_source = path.read_text(encoding="utf-8")
            production_source = _mask_rust_cfg_test_items(raw_source)
            production_code = _mask_rust_non_code(production_source)
            for label, pattern in KEYNOTE_NO_EAGER_PROST_SOURCE_PATTERNS:
                for match in pattern.finditer(production_code):
                    line_number = production_code.count("\n", 0, match.start()) + 1
                    violations.append(
                        "focused litchi-keynote production source uses "
                        f"{label}: {path.relative_to(root)}:{line_number}"
                    )

    manifest_path = root / KEYNOTE_PACKAGE_MANIFEST
    if manifest_path.is_file():
        section: str | None = None
        for line_number, line in enumerate(
            manifest_path.read_text(encoding="utf-8").splitlines(), start=1
        ):
            header = CARGO_SECTION_HEADER.match(line)
            if header is not None:
                section = header.group(1).strip()
                continue
            if CARGO_PROST_DEPENDENCY.match(line) is None:
                continue
            is_dev_dependency = section == "dev-dependencies" or (
                section is not None and section.endswith(".dev-dependencies")
            )
            if is_dev_dependency:
                continue
            violations.append(
                "focused litchi-keynote Cargo manifest retains normal prost "
                f"dependency: {KEYNOTE_PACKAGE_MANIFEST}:{line_number}"
            )

    return sorted(set(violations))


def audit_keynote_chart_title_legacy_calls(root: Path = ROOT) -> list[str]:
    """Keep focused chart-title code off the legacy host mutation calls."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []

    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        raw_source = path.read_text(encoding="utf-8")
        source = _mask_rust_non_code(raw_source)
        for match in KEYNOTE_CHART_TITLE_LEGACY_CALL.finditer(source):
            line_number = source.count("\n", 0, match.start("method")) + 1
            violations.append(
                "focused litchi-keynote chart-title source retains legacy call "
                f"{match.group('method')}: {path.relative_to(root)}:{line_number}"
            )

    return sorted(set(violations))


def _keynote_chart_legend_boundary_active(root: Path) -> bool:
    """Return whether the focused owner crossed its package-module seam."""

    owner_path = root / KEYNOTE_CHART_LEGEND_OWNER_SOURCE
    package_path = root / KEYNOTE_CHART_LEGEND_EXPORT_SOURCES[0]
    if not owner_path.is_file() or not package_path.is_file():
        return False
    package_source = _mask_rust_non_code(
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
    )
    return re.search(
        rf"(?m)^(?:pub(?:\([^()]*\))?\s+)?mod\s+"
        rf"{re.escape(KEYNOTE_CHART_LEGEND_PACKAGE_MODULE)}\s*;",
        package_source,
    ) is not None


def audit_keynote_chart_legend_visibility_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the selector-first, generated-free focused legend facade.

    This audit remains dormant until the owner and private package module are
    wired together, so historical checkouts continue to pass while a
    migration is being staged. Once active, it checks only public Rust
    declarations and explicit re-exports; private graph identifiers are
    intentionally outside this boundary scan.
    """

    if not _keynote_chart_legend_boundary_active(root):
        return []

    violations: list[str] = []
    owner_path = root / KEYNOTE_CHART_LEGEND_OWNER_SOURCE
    package_path, lib_path = (
        root / path for path in KEYNOTE_CHART_LEGEND_EXPORT_SOURCES
    )

    if not owner_path.is_file():
        violations.append(
            "focused litchi-keynote chart-legend visibility public API is missing "
            f"private package owner source: {KEYNOTE_CHART_LEGEND_OWNER_SOURCE}"
        )
    if not package_path.is_file():
        violations.append(
            "focused litchi-keynote chart-legend visibility public API is missing "
            f"package export source: {KEYNOTE_CHART_LEGEND_EXPORT_SOURCES[0]}"
        )
    if not lib_path.is_file():
        violations.append(
            "focused litchi-keynote chart-legend visibility public API is missing "
            f"crate export source: {KEYNOTE_CHART_LEGEND_EXPORT_SOURCES[1]}"
        )

    package_source = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source)
    module_matches = list(
        re.finditer(
            rf"(?m)^(?P<indent>\s*)(?P<public>pub(?:\([^()]*\))?\s+)?"
            rf"mod\s+{re.escape(KEYNOTE_CHART_LEGEND_PACKAGE_MODULE)}\s*;",
            package_code,
        )
    )
    if not module_matches:
        violations.append(
            "focused litchi-keynote chart-legend visibility public API is missing "
            f"private package owner module: {KEYNOTE_CHART_LEGEND_EXPORT_SOURCES[0]}"
        )
    elif any(
        match.group("public") and match.group("public").strip() != "pub(crate)"
        for match in module_matches
    ):
        line_number = package_code.count("\n", 0, module_matches[0].start()) + 1
        violations.append(
            "focused litchi-keynote chart-legend visibility owner module must remain "
            f"private: {KEYNOTE_CHART_LEGEND_EXPORT_SOURCES[0]}:{line_number}"
        )

    owner_source = (
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        if owner_path.is_file()
        else ""
    )
    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for name in sorted(KEYNOTE_CHART_LEGEND_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote chart-legend visibility Package method is "
                f"missing {name}: {KEYNOTE_CHART_LEGEND_OWNER_SOURCE}"
            )
            continue
        if name == "apply_slide_chart_legend":
            required = "Patch"
            if re.search(rf"\b{re.escape(required)}\b", declaration) is None:
                violations.append(
                    "focused litchi-keynote chart-legend visibility apply method must "
                    f"accept {required}: {KEYNOTE_CHART_LEGEND_OWNER_SOURCE}"
                )
        else:
            for selector in ("SlideSelector", "ChartSelector"):
                if re.search(rf"\b{selector}\b", declaration) is None:
                    violations.append(
                        "focused litchi-keynote chart-legend visibility Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_CHART_LEGEND_OWNER_SOURCE}"
                    )
        if re.search(r"\bu64\b", declaration):
            violations.append(
                "focused litchi-keynote chart-legend visibility Package method must "
                f"not expose u64: {name}: {KEYNOTE_CHART_LEGEND_OWNER_SOURCE}"
            )
        raw_id = KEYNOTE_CHART_LEGEND_RAW_ID_TOKEN.search(declaration)
        if raw_id is not None:
            violations.append(
                "focused litchi-keynote chart-legend visibility Package method must "
                f"not expose raw identifier {raw_id.group(0)}: "
                f"{name}: {KEYNOTE_CHART_LEGEND_OWNER_SOURCE}"
            )

    owner_declarations = _rust_public_declarations(owner_source)
    for declaration, line_number in owner_declarations:
        for match in KEYNOTE_CHART_LEGEND_RAW_ID_TOKEN.finditer(declaration):
            violations.append(
                "focused litchi-keynote chart-legend visibility public API exposes "
                f"raw identifier {match.group(0)}: "
                f"{KEYNOTE_CHART_LEGEND_OWNER_SOURCE}:{line_number}"
            )
        for match in RUST_BYTE_SLICE.finditer(declaration):
            raw_slice = re.sub(r"\s+", "", match.group(0))
            violations.append(
                "focused litchi-keynote chart-legend visibility public API exposes "
                f"raw byte slice {raw_slice}: "
                f"{KEYNOTE_CHART_LEGEND_OWNER_SOURCE}:{line_number}"
            )
        if re.search(
            r"\b(?:prost|prost_types|buffa|kn\w*|tsch?\w*|tsd\w*|"
            r"Archive|RawMessage|HostPackage|ComponentCatalog|WireView)\b",
            declaration,
        ):
            violations.append(
                "focused litchi-keynote chart-legend visibility public API exposes "
                f"generated/physical type: {KEYNOTE_CHART_LEGEND_OWNER_SOURCE}:"
                f"{line_number}"
            )

    family_path = root / KEYNOTE_CHART_LEGEND_FAMILY_SOURCE
    family_exports = _rust_canonical_exports(
        _keynote_family_view(root, family_path, {}, KEYNOTE_CHART_LEGEND_FAMILY_MODULE),
        KEYNOTE_CHART_LEGEND_CANONICAL_TYPES,
    )
    for name in sorted(KEYNOTE_CHART_LEGEND_CANONICAL_TYPES - family_exports):
        violations.append(
            "focused litchi-keynote chart-legend visibility public API is missing "
            f"canonical type {name}: {family_path.relative_to(root)}"
        )
    for path in (package_path, lib_path, family_path):
        if not path.is_file():
            continue
        source = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not (
                "slide_chart_legend" in declaration
                or identifiers
                & (KEYNOTE_CHART_LEGEND_CANONICAL_TYPES - KEYNOTE_FAMILY_GENERIC_NAMES)
            ):
                continue
            for match in KEYNOTE_CHART_LEGEND_RAW_ID_TOKEN.finditer(declaration):
                violations.append(
                    "focused litchi-keynote chart-legend visibility public API exposes "
                    f"raw identifier {match.group(0)}: {path.relative_to(root)}:"
                    f"{line_number}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                raw_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote chart-legend visibility public API exposes "
                    f"raw byte slice {raw_slice}: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if re.search(
                r"\b(?:prost|prost_types|buffa|kn\w*|tsch?\w*|tsd\w*|"
                r"Archive|RawMessage|HostPackage|ComponentCatalog|WireView)\b",
                declaration,
            ):
                violations.append(
                    "focused litchi-keynote chart-legend visibility public API exposes "
                    f"generated/physical type: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_keynote_chart_arrangement_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require a selector-first Keynote chart-arrangement package facade."""

    if not _keynote_chart_arrangement_owner_present(root):
        return []

    owner_path = root / KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE
    semantic_path = root / KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_SOURCE
    common_path = root / KEYNOTE_CHART_ARRANGEMENT_COMMON_SOURCE
    selector_path = root / KEYNOTE_SOURCE_ROOT / "selector.rs"
    package_path, lib_path = (
        root / path for path in KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES
    )
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    common_source = (
        _mask_rust_cfg_test_items(common_path.read_text(encoding="utf-8"))
        if common_path.is_file()
        else ""
    )
    common_code = _mask_rust_non_code(common_source)
    if not common_path.is_file():
        violations.append(
            "focused litchi-keynote chart-arrangement semantic API is missing its "
            f"common value source: {KEYNOTE_CHART_ARRANGEMENT_COMMON_SOURCE}"
        )
    elif re.search(
        r"(?m)^\s*pub\s+struct\s+ChartArrangement\b", common_code
    ) is None:
        violations.append(
            "focused litchi-keynote chart-arrangement common value is missing "
            f"ChartArrangement: {KEYNOTE_CHART_ARRANGEMENT_COMMON_SOURCE}"
        )
    common_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            common_source, "ChartArrangement"
        )
    }
    for name in sorted(KEYNOTE_CHART_ARRANGEMENT_COMMON_METHODS - common_methods):
        violations.append(
            "focused litchi-keynote chart-arrangement common value is missing "
            f"API method {name}: {KEYNOTE_CHART_ARRANGEMENT_COMMON_SOURCE}"
        )

    for path, label in (
        (owner_path, "private package owner source"),
        (semantic_path, "semantic source"),
        (selector_path, "selector source"),
        (package_path, "package export source"),
        (lib_path, "crate export source"),
    ):
        if not path.is_file():
            violations.append(
                "focused litchi-keynote chart-arrangement facade is missing "
                f"{label}: {path.relative_to(root)}"
            )

    package_code = code[package_path]
    module_matches = list(
        re.finditer(
            rf"(?m)^(?P<indent>\s*)(?P<public>pub(?:\([^()]*\))?\s+)?"
            rf"mod\s+{re.escape(KEYNOTE_CHART_ARRANGEMENT_PACKAGE_MODULE)}\s*;",
            package_code,
        )
    )
    if not module_matches:
        violations.append(
            "focused litchi-keynote chart-arrangement facade is missing private "
            f"owner module: {KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES[0]}"
        )
    elif any(
        match.group("public") and match.group("public").strip() != "pub(crate)"
        for match in module_matches
    ):
        line_number = package_code.count("\n", 0, module_matches[0].start()) + 1
        violations.append(
            "focused litchi-keynote chart-arrangement owner module must remain "
            f"private: {KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES[0]}:{line_number}"
        )

    semantic_code = code[semantic_path]
    common_reexport = re.compile(
        r"\s*pub\s+use\s+litchi_iwa_common\s*::\s*chart\s*::\s*"
        r"arrangement\s*::\s*ChartArrangement\s*;\s*",
        re.DOTALL,
    )
    common_arrangement_use = re.compile(
        r"(?m)^\s*pub\s+use\s+litchi_iwa_common\s*::\s*chart\s*::\s*"
        r"arrangement\b[^;]*;"
    )
    exact_common_reexports = [
        declaration
        for declaration, _line_number in _rust_public_declarations(
            sources[semantic_path]
        )
        if common_reexport.fullmatch(declaration)
    ]
    if not exact_common_reexports:
        violations.append(
            "focused litchi-keynote chart-arrangement semantic API is missing the "
            "exact common ChartArrangement re-export: "
            f"{KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_SOURCE}"
        )
    for match in common_arrangement_use.finditer(semantic_code):
        declaration = match.group(0)
        if common_reexport.fullmatch(declaration):
            continue
        line_number = semantic_code.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-keynote chart-arrangement semantic API must use only "
            "the exact common ChartArrangement re-export: "
            f"{KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_SOURCE}:{line_number}"
        )
    semantic_exports = _rust_canonical_exports(
        sources[lib_path], KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_TYPES
    )
    for name in sorted(KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_TYPES - semantic_exports):
        violations.append(
            "focused litchi-keynote chart-arrangement semantic API is missing root "
            f"export {name}: {KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES[1]}"
        )
    selector_exports = _rust_canonical_exports(
        sources[lib_path], KEYNOTE_CHART_ARRANGEMENT_SELECTOR_TYPES
    )
    for name in sorted(KEYNOTE_CHART_ARRANGEMENT_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote chart-arrangement public API is missing selector "
            f"{name}: {KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES[1]}"
        )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_CHART_ARRANGEMENT_PACKAGE_METHODS):
        record = owner_methods.get(name)
        if record is None:
            violations.append(
                "focused litchi-keynote chart-arrangement Package method is missing "
                f"{name}: {KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}"
            )
            continue
        declaration, line_number = record
        if name == "apply_slide_chart_arrangement":
            if re.search(r"\bPatch\b", declaration) is None:
                violations.append(
                    "focused litchi-keynote chart-arrangement apply method must accept "
                    f"Patch: {KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}:"
                    f"{line_number}"
                )
        elif name == "slide_chart_arrangements":
            if re.search(r"\bSlideSelector\b", declaration) is None:
                violations.append(
                    "focused litchi-keynote chart-arrangement batch method must accept "
                    f"selector-first SlideSelector: {KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}:"
                    f"{line_number}"
                )
            if re.search(r"\bChartArrangement\b", declaration) is None:
                violations.append(
                    "focused litchi-keynote chart-arrangement batch method must expose "
                    f"ChartArrangement state: {KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}:"
                    f"{line_number}"
                )
        else:
            for selector in sorted(KEYNOTE_CHART_ARRANGEMENT_SELECTOR_TYPES):
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                    violations.append(
                        "focused litchi-keynote chart-arrangement Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}:{line_number}"
                    )
            if (
                name == "slide_chart_arrangement"
                and re.search(r"\bChartArrangement\b", declaration) is None
            ):
                violations.append(
                    "focused litchi-keynote chart-arrangement Package method must "
                    f"expose ChartArrangement state: {name}: "
                    f"{KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}:{line_number}"
                )
        if re.search(r"\bu64\b", declaration):
            violations.append(
                "focused litchi-keynote chart-arrangement Package method must not "
                f"expose u64: {name}: {KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}:{line_number}"
            )
        raw_identifier = KEYNOTE_CHART_ARRANGEMENT_RAW_ID_PARAMETER.search(declaration)
        if raw_identifier is not None:
            violations.append(
                "focused litchi-keynote chart-arrangement Package method exposes raw "
                f"identifier {raw_identifier.group(0).strip()}: {name}: "
                f"{KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}:{line_number}"
            )

    edit_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(KEYNOTE_CHART_ARRANGEMENT_EDIT_METHODS):
        if name not in edit_methods:
            violations.append(
                "focused litchi-keynote chart-arrangement edit is missing "
                f"{name}: {KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE}"
            )

    for name in sorted(KEYNOTE_CHART_ARRANGEMENT_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_CHART_ARRANGEMENT_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_CHART_ARRANGEMENT_FAMILY_MODULE), KEYNOTE_CHART_ARRANGEMENT_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote chart-arrangement public API is missing "
                    f"canonical type {name}: {path.relative_to(root)}"
                )
    facade_names = (
        KEYNOTE_CHART_ARRANGEMENT_CANONICAL_TYPES
        | KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_TYPES
        | KEYNOTE_CHART_ARRANGEMENT_SELECTOR_TYPES
        | KEYNOTE_CHART_ARRANGEMENT_PACKAGE_METHODS
        | KEYNOTE_CHART_ARRANGEMENT_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path, selector_path}
    for path in paths:
        dedicated = path in dedicated_sources
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if path == semantic_path and not identifiers.intersection(
                KEYNOTE_CHART_ARRANGEMENT_SEMANTIC_TYPES
                | KEYNOTE_CHART_ARRANGEMENT_FLAT_ALIASES
            ):
                continue
            if path == selector_path and not identifiers.intersection(
                KEYNOTE_CHART_ARRANGEMENT_SELECTOR_TYPES
            ):
                continue
            if not dedicated and not identifiers.intersection(facade_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            canonical_common_reexport = (
                path == semantic_path and common_reexport.fullmatch(declaration)
            )
            for identifier in sorted(identifiers):
                if canonical_common_reexport and identifier == "litchi_iwa_common":
                    continue
                if identifier in KEYNOTE_CHART_ARRANGEMENT_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_CHART_ARRANGEMENT_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_CHART_ARRANGEMENT_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote chart-arrangement public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_CHART_ARRANGEMENT_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote chart-arrangement public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote chart-arrangement public API exposes raw "
                    f"byte slice {byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_CHART_ARRANGEMENT_PUBLIC_RAW_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote chart-arrangement public API exposes raw "
                    f"parameter {match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote chart-arrangement public API retains a "
                    f"glob re-export: {path.relative_to(root)}:{line_number}"
                )
            if KEYNOTE_CHART_ARRANGEMENT_MONOLITH_PATH.search(declaration):
                violations.append(
                    "focused litchi-keynote chart-arrangement public API retains a "
                    f"monolith chart re-export: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _keynote_chart_arrangement_deprecated_legacy_methods(
    source: str,
) -> frozenset[str]:
    """Return raw-ID compatibility methods carrying a direct deprecation attr."""

    code = _mask_rust_non_code(source)
    declaration = re.compile(
        r"(?<![A-Za-z0-9_#])pub(?:[ \t]*\([^()]*\))?[ \t\r\n]+fn[ \t\r\n]+"
        r"(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
    )
    attributes = re.compile(r"#[ \t]*\[(?P<body>[^\]]*)\]")
    deprecated: set[str] = set()
    for match in declaration.finditer(code):
        name = match.group("name")
        if name not in KEYNOTE_CHART_ARRANGEMENT_LEGACY_METHODS:
            continue
        prefix = code[: match.start()]
        for attribute in attributes.finditer(prefix):
            # Only attributes in the contiguous item attribute block may
            # describe this function.  This prevents a deprecated unrelated
            # item from satisfying a later compatibility shell.
            suffix = prefix[attribute.end() :]
            if not re.fullmatch(
                r"[ \t\r\n]*(?:#[ \t]*\[[^\]]*\][ \t\r\n]*)*",
                suffix,
            ):
                continue
            if re.match(r"[ \t\r\n]*deprecated\b", attribute.group("body")):
                deprecated.add(name)
                break
    return frozenset(deprecated)


def _keynote_chart_axis_title_owner_present(root: Path) -> bool:
    """Return whether the axis-title owner crossed its wiring seam."""

    owner_path = root / KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE
    package_path = root / KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_chart_axis_title\s*;",
        package_source,
    ) is not None


def audit_keynote_chart_axis_title_legacy_calls(root: Path = ROOT) -> list[str]:
    """Keep focused package source off retired Keynote axis-title calls."""

    if not _keynote_chart_axis_title_owner_present(root):
        return []
    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for match in KEYNOTE_CHART_AXIS_TITLE_LEGACY_CALL.finditer(source):
            line_number = source.count("\n", 0, match.start("method")) + 1
            violations.append(
                "focused litchi-keynote chart-axis-title source retains legacy call "
                f"{match.group('method')}: {path.relative_to(root)}:{line_number}"
            )
    return sorted(set(violations))


def _keynote_chart_axis_title_check_codec(root: Path) -> list[str]:
    """Require the private neutral projection and prepared rewrite seam."""

    codec_path = root / KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE
    if not codec_path.is_file():
        return [
            "focused litchi-keynote chart-axis-title public API is missing strict neutral "
            f"codec source: {KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
        ]
    raw_source = codec_path.read_text(encoding="utf-8")
    source = _mask_rust_cfg_test_items(raw_source)
    code = _mask_rust_non_code(source)
    violations: list[str] = []
    for api in KEYNOTE_CHART_AXIS_TITLE_CODEC_REQUIRED_APIS:
        if re.search(
            rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
            code,
        ) is None:
            violations.append(
                "focused litchi-keynote chart-axis-title hidden codec is missing strict API "
                f"{api}: {KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
            )
    for type_name in sorted(KEYNOTE_CHART_AXIS_TITLE_CODEC_PREPARED_TYPES):
        if re.search(
            rf"\b(?:pub\s+)?(?:struct|enum|type)\s+{re.escape(type_name)}\b",
            code,
        ) is None:
            violations.append(
                "focused litchi-keynote chart-axis-title hidden codec is missing prepared "
                f"type {type_name}: {KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_CHART_AXIS_TITLE_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not all(marker in code for marker in markers):
            violations.append(
                "focused litchi-keynote chart-axis-title hidden codec is missing "
                f"{label} marker: {KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
            )
    if any(marker in code for marker in ("prost", "HostPackage", "chart_graph")):
        violations.append(
            "focused litchi-keynote chart-axis-title hidden codec must remain neutral "
            f"and generated-free: {KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
        )

    codec_lib_path = root / KEYNOTE_CHART_AXIS_TITLE_CODEC_PUBLIC_SOURCE
    codec_lib = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(KEYNOTE_CHART_AXIS_TITLE_CODEC_MODULE)}\b",
        codec_lib,
    ) is None:
        violations.append(
            "focused litchi-keynote chart-axis-title public API is missing hidden codec "
            f"module {KEYNOTE_CHART_AXIS_TITLE_CODEC_MODULE}: "
            f"{KEYNOTE_CHART_AXIS_TITLE_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(r"(?m)#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\]", raw_source) is None:
        violations.append(
            "focused litchi-keynote chart-axis-title hidden codec is missing cfg(test) coverage: "
            f"{KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
        )
    elif re.search(r"(?m)#\s*\[\s*test\s*\]", raw_source) is None:
        violations.append(
            "focused litchi-keynote chart-axis-title hidden codec is missing #[test] coverage: "
            f"{KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
        )
    return violations


def audit_keynote_chart_axis_title_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the typed Axis/ChartSelector package owner and neutral codec."""

    if not _keynote_chart_axis_title_owner_present(root):
        return []
    owner_path = root / KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE
    semantic_path = root / KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_CHART_AXIS_TITLE_SELECTOR_SOURCE
    package_path = root / KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []
    package_code = code[package_path]
    package_library_code = _mask_rust_non_code(
        sources[package_path] + sources[lib_path]
    )
    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+slide_chart_axis_title\b",
        package_library_code,
    ):
        violations.append(
            "focused litchi-keynote chart-axis-title owner module must remain private: "
            f"{KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_chart_axis_title\s*;",
        package_code,
    ) is None:
        violations.append(
            "focused litchi-keynote chart-axis-title owner module is missing: "
            f"{KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_CHART_AXIS_TITLE_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_CHART_AXIS_TITLE_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_CHART_AXIS_TITLE_FAMILY_MODULE), KEYNOTE_CHART_AXIS_TITLE_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote chart-axis-title public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_TYPES
    )
    for name in sorted(KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_TYPES - semantic_exports):
        violations.append(
            "focused litchi-keynote chart-axis-title semantic API is missing "
            f"{name}: {KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_SOURCE}"
        )
    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[semantic_path] + sources[lib_path],
        KEYNOTE_CHART_AXIS_TITLE_SELECTOR_TYPES,
    )
    for name in sorted(KEYNOTE_CHART_AXIS_TITLE_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote chart-axis-title public API is missing selector "
            f"{name}: {KEYNOTE_CHART_AXIS_TITLE_EXPORT_SOURCES[1]}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_CHART_AXIS_TITLE_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote chart-axis-title Package method is missing "
                f"{name}: {KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_chart_axis_title":
            for selector in ("SlideSelector", "ChartSelector", "Axis"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote chart-axis-title Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
                    )
        elif "Patch" not in declaration:
            violations.append(
                "focused litchi-keynote chart-axis-title apply method must accept "
                f"Patch: {KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
            )

    edit_methods = {
        name
        for name, _declaration, _line in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(KEYNOTE_CHART_AXIS_TITLE_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote chart-axis-title edit is missing "
            f"{name}: {KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_CHART_AXIS_TITLE_CANONICAL_TYPES
        | KEYNOTE_CHART_AXIS_TITLE_SEMANTIC_TYPES
        | KEYNOTE_CHART_AXIS_TITLE_SELECTOR_TYPES
        | KEYNOTE_CHART_AXIS_TITLE_PACKAGE_METHODS
        | KEYNOTE_CHART_AXIS_TITLE_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path, selector_path}
    for path in paths:
        dedicated = path in dedicated_sources
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not identifiers.intersection(facade_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_CHART_AXIS_TITLE_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_CHART_AXIS_TITLE_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif (
                    identifier == "wire"
                    or identifier in KEYNOTE_CHART_AXIS_TITLE_WIRE_TYPES
                ):
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote chart-axis-title public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_CHART_AXIS_TITLE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote chart-axis-title public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote chart-axis-title public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_CHART_AXIS_TITLE_PUBLIC_RAW_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote chart-axis-title public API exposes raw "
                    f"parameter {match.group(0).strip()}: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote chart-axis-title public API retains a glob "
                    f"re-export: {path.relative_to(root)}:{line_number}"
                )

    owner_code = code[owner_path]
    for label, marker in KEYNOTE_CHART_AXIS_TITLE_CODEC_FLOW_MARKERS.items():
        if marker.search(owner_code) is None:
            violations.append(
                "focused litchi-keynote chart-axis-title owner is missing strict prepared "
                f"{label} marker: {KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
            )
    violations.extend(_keynote_chart_axis_title_check_codec(root))
    return sorted(set(violations))


def audit_keynote_chart_axis_title_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one bounded transaction budget and source-authoritative graph work."""

    if not _keynote_chart_axis_title_owner_present(root):
        return []
    owner_path = root / KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    violations: list[str] = []
    budgets = list(KEYNOTE_CHART_AXIS_TITLE_PACKAGE_BUDGET_TYPE.finditer(owner))
    if len(budgets) != 1:
        violations.append(
            "focused litchi-keynote chart-axis-title owner must define exactly one "
            f"aggregate transaction budget (found {len(budgets)}): "
            f"{KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
        )
    for label, markers in KEYNOTE_CHART_AXIS_TITLE_PACKAGE_MARKER_GROUPS.items():
        if label == "source authority":
            present = KEYNOTE_CHART_AXIS_TITLE_SOURCE_AUTHORITY_MARKER.search(owner) is not None and all(
                marker in owner for marker in markers if marker != "authority"
            )
        else:
            present = all(marker in owner for marker in markers)
        if not present:
            violations.append(
                "focused litchi-keynote chart-axis-title owner is missing "
                f"{label} marker: {KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
            )
    return sorted(set(violations))


def audit_keynote_chart_axis_title_completion_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require exact completion artifacts and a single prepared execution."""

    if not _keynote_chart_axis_title_owner_present(root):
        return []
    owner_path = root / KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    violations: list[str] = []
    for label, marker in KEYNOTE_CHART_AXIS_TITLE_CODEC_FLOW_MARKERS.items():
        if marker.search(owner) is None:
            violations.append(
                "focused litchi-keynote chart-axis-title completion is missing "
                f"{label}: {KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
            )
    execute_calls = list(
        re.finditer(r"\.[ \t\r\n]*execute\s*\(", owner)
    )
    if len(execute_calls) != 1:
        violations.append(
            "focused litchi-keynote chart-axis-title completion must have exactly one "
            f"prepared execute call (found {len(execute_calls)}): "
            f"{KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}"
        )
    direct_rewrite_patterns = [
        re.compile(
            r"\b(?:keynote_chart_axis_title_codec\s*::\s*)?"
            r"rewrite_axis_title\s*\("
        )
    ]
    # The owner commonly aliases the neutral function on import to keep the
    # axis-title domain visible at call sites. Resolve those aliases before
    # scanning so a direct rewrite cannot hide behind a compatibility name.
    for import_match in re.finditer(
        r"\brewrite_axis_title\b[ \t\r\n]+as[ \t\r\n]+"
        r"(?P<alias>[A-Za-z_][A-Za-z0-9_]*)\b",
        owner,
    ):
        direct_rewrite_patterns.append(
            re.compile(rf"\b{re.escape(import_match.group('alias'))}\s*\(")
        )
    for direct_rewrite in direct_rewrite_patterns:
        for match in direct_rewrite.finditer(owner):
            line_start = owner.rfind("\n", 0, match.start()) + 1
            line_end = owner.find("\n", match.end())
            if line_end < 0:
                line_end = len(owner)
            if re.search(r"\bfn\s+rewrite_axis_title\b", owner[line_start:line_end]):
                continue
            line_number = owner.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote chart-axis-title completion must execute the prepared "
                f"codec, not direct rewrite_axis_title: {KEYNOTE_CHART_AXIS_TITLE_OWNER_SOURCE}:{line_number}"
            )

    codec_path = root / KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE
    if codec_path.is_file():
        codec_raw = codec_path.read_text(encoding="utf-8")
        codec = _mask_rust_non_code(_mask_rust_cfg_test_items(codec_raw))
        if re.search(r"\bexecute\s*\(", codec) is None:
            violations.append(
                "focused litchi-keynote chart-axis-title completion is missing codec execute "
                f"implementation: {KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
            )
    else:
        violations.append(
            "focused litchi-keynote chart-axis-title completion is missing codec source: "
            f"{KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
        )

    # Keep the unit-test and fuzz seams tied to the owner activation.  A
    # cfg(test) decoy in the owner cannot satisfy the integration requirement.
    codec_raw = (
        codec_path.read_text(encoding="utf-8") if codec_path.is_file() else ""
    )
    if re.search(r"(?m)#\s*\[\s*test\s*\]", codec_raw) is None:
        violations.append(
            "focused litchi-keynote chart-axis-title completion is missing codec #[test] coverage: "
            f"{KEYNOTE_CHART_AXIS_TITLE_CODEC_SOURCE}"
        )
    for test_path in KEYNOTE_CHART_AXIS_TITLE_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote chart-axis-title boundary is missing integration test: "
                f"{test_path}"
            )
            continue
        test_code = _mask_rust_non_code(absolute.read_text(encoding="utf-8"))
        if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
            violations.append(
                "focused litchi-keynote chart-axis-title integration test is missing #[test] coverage: "
                f"{test_path}"
            )
        if not re.search(
            r"\b(?:slide_chart_axis_title|edit_slide_chart_axis_title|"
            r"apply_slide_chart_axis_title)\b",
            test_code,
        ):
            violations.append(
                "focused litchi-keynote chart-axis-title integration test does not exercise owner: "
                f"{test_path}"
            )
        if "deleted_previews" not in test_code:
            violations.append(
                "focused litchi-keynote chart-axis-title integration test must report root "
                f"preview deletions: {test_path}"
            )

    for fuzz_path in KEYNOTE_CHART_AXIS_TITLE_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote chart-axis-title boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote chart-axis-title fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in KEYNOTE_CHART_AXIS_TITLE_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-keynote chart-axis-title boundary is missing fuzz corpus: "
                f"{corpus}"
            )
    return sorted(set(violations))


# The complete Keynote value-axis scalar surface lives behind one
# selector-first package owner.  Bounds, steps, and scale are intentionally
# audited as one aggregate: three independently staged transactions could
# each pass a local check while still losing source spans, violating locality,
# or reopening the package three times.  The gate is dormant until both the
# private owner source and its package-module wiring exist, which keeps a
# partially staged migration usable while the focused crate is being built.
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMMON_SOURCE = Path(
    "crates/litchi-iwa-common/src/chart/axis/settings.rs"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMMON_AXIS_SOURCE = Path(
    "crates/litchi-iwa-common/src/chart/axis.rs"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SEMANTIC_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SELECTOR_SOURCE = KEYNOTE_SOURCE_ROOT / "selector.rs"
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_chart_value_axis.rs"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
# ADR 0034: the family's public home is its focused semantic
# module, which re-exports the unprefixed transaction names.
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FAMILY_SOURCE = KEYNOTE_SOURCE_ROOT / "chart.rs"
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FAMILY_MODULE = 'value_axis'
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CANONICAL_TYPES = frozenset(
    {
        "Commit",
        "Diagnostics",
        "Edit",
        "Error",
        "LimitKind",
        "Patch",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SEMANTIC_TYPES = frozenset(
    {
        "ValueAxisSettings",
        "Bounds",
        "Bound",
        "Steps",
        "MajorStepCount",
        "MinorStepCount",
        "Scale",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SELECTOR_TYPES = frozenset(
    {"ChartSelector", "SlideSelector"}
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PACKAGE_METHODS = frozenset(
    {
        "slide_chart_value_axis_settings",
        "edit_slide_chart_value_axis_settings",
        "apply_slide_chart_value_axis_settings",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EDIT_METHODS = frozenset(
    {"before", "after", "set", "set_bounds", "set_steps", "set_scale", "commit"}
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FLAT_ALIASES = frozenset(
    {
        "ValueAxisSettingsSnapshot",
        "ValueAxisSettingsWrite",
        "AxisValueSettingsSnapshot",
        "AxisValueSettingsWrite",
        "ChartValueAxis",
        "ChartValueAxisSnapshot",
        "ChartValueAxisWrite",
        "ValueAxisCommit",
        "ValueAxisDiagnostics",
        "ValueAxisEdit",
        "ValueAxisError",
        "ValueAxisLimitKind",
        "ValueAxisPatch",
        "ValueAxisSettingsCommit",
        "ValueAxisSettingsDiagnostics",
        "ValueAxisSettingsEdit",
        "ValueAxisSettingsError",
        "ValueAxisSettingsLimitKind",
        "ValueAxisSettingsPatch",
        "ChartValueAxisSettingsCommit",
        "ChartValueAxisSettingsDiagnostics",
        "ChartValueAxisSettingsEdit",
        "ChartValueAxisSettingsError",
        "ChartValueAxisSettingsLimitKind",
        "ChartValueAxisSettingsPatch",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMPATIBILITY_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t]+mod[ \t]+"
    r"(?:r#)?(?:raw|wire|archive|protobuf|proto|generated|compat|legacy)\b"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMPATIBILITY_REEXPORT = re.compile(
    r"(?m)^[ \t]*pub[ \t]+use\b[^;\n]*\b"
    r"(?:raw|wire|archive|protobuf|proto|generated|compat|legacy)\b"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PHYSICAL_TYPES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ComponentCatalog",
        "EntryEdit",
        "ExactArtifacts",
        "HostPackage",
        "PhysicalSource",
        "RawMessage",
        "SnappyStream",
        "SourceCatalog",
        "ValueAxisArchive",
        "ValueAxisSnapshot",
        "ValueAxisWrite",
        "AxisValueSettings",
        "AxisValueSettingsSnapshot",
        "AxisValueSettingsWrite",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_WIRE_TYPES = frozenset(
    {
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "AxisValueBound",
        "AxisValueBounds",
        "AxisValueSettings",
        "AxisValueSettingsSnapshot",
        "AxisValueSettingsWrite",
        "AxisValueSteps",
        "ValueAxisSemanticError",
        "NestedFieldEdit",
        "NestedFieldReplacement",
        "PreparedAxisValueSettingsRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
        "RewriteReport",
        "WireDescent",
        "WireError",
        "WireFieldView",
        "WireLimits",
        "WireResourceLimit",
        "WireView",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PROTO_ORIGINS = frozenset(
    {
        "buffa",
        "prost",
        "prost_types",
        "kn",
        "tsch",
        "tsp",
        "tsd",
        "litchi_iwa_protos",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PUBLIC_RAW_PARAMETER = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|native_id|object_id|"
    r"drawable_object_id|chart_object_id|axis_id|component_id|member_id|"
    r"archive_id|message_id|uuid|source_bytes|bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*(?:u64|u32|usize|i64|i32|&\s*\[\s*u8\s*\])"
    r"(?=$|[^A-Za-z0-9_])"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_LEGACY_METHODS = frozenset(
    {
        "slide_chart_value_axis_bounds",
        "set_slide_chart_value_axis_bounds",
        "slide_chart_value_axis_steps",
        "set_slide_chart_value_axis_steps",
        "slide_chart_value_axis_scale",
        "set_slide_chart_value_axis_scale",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_LEGACY_METHOD_DECLARATION = re.compile(
    r"(?<![A-Za-z0-9_#])(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"(?:unsafe[ \t\r\n]+|async[ \t\r\n]+|const[ \t\r\n]+)*"
    r"fn[ \t\r\n]+(?P<name>slide_chart_value_axis_bounds|"
    r"set_slide_chart_value_axis_bounds|slide_chart_value_axis_steps|"
    r"set_slide_chart_value_axis_steps|slide_chart_value_axis_scale|"
    r"set_slide_chart_value_axis_scale)\b"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_LEGACY_CALL = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?P<method>slide_chart_value_axis_bounds|"
    r"set_slide_chart_value_axis_bounds|slide_chart_value_axis_steps|"
    r"set_slide_chart_value_axis_steps|slide_chart_value_axis_scale|"
    r"set_slide_chart_value_axis_scale)(?![A-Za-z0-9_])[ \t\r\n]*\("
)
RETIRED_IWA_KEYNOTE_CHART_AXIS_VALUE_SETTINGS_TESTS = (
    "scratch_presentation_supports_native_chart_value_axis_bounds_crud",
    "scratch_presentation_supports_native_chart_value_axis_steps_crud",
    "scratch_presentation_supports_native_chart_value_axis_scale_crud",
)
RETIRED_IWA_KEYNOTE_CHART_AXIS_VALUE_SETTINGS_MODULES = (
    "axis_bounds",
    "axis_steps",
    "axis_scale",
)

KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_chart_axis_value_settings_codec.rs"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_MODULE = (
    "keynote_chart_axis_value_settings_codec"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PROJECTION_SOURCE = Path(
    "crates/litchi-iwa-protos/src/buffa-projections/TSCHChartAxisValueSettingsArchive.proto"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_REQUIRED_APIS = (
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "AxisValueSettingsSnapshot",
        "AxisValueSettingsWrite",
        "DecodeReport",
        "RewriteReport",
        "decode_axis_value_settings",
        "decode_axis_value_settings_with_report",
        "prepare_axis_value_settings_rewrite",
        "rewrite_axis_value_settings",
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_PREPARED_TYPES = frozenset(
    {
        "PreparedAxisValueSettingsRewrite",
        "RewriteExecutionRequirements",
        "RewriteExecutionLimits",
    }
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_REQUIRED_MARKER_GROUPS = {
    "value-axis scalar fields": (
        re.compile(r"(?i)major"),
        re.compile(r"(?i)minor"),
        re.compile(r"(?i)scale"),
        re.compile(r"(?i)maximum"),
        re.compile(r"(?i)minimum"),
    ),
    "native field numbers": (
        re.compile(r"\bu32\s*=\s*5\b"),
        re.compile(r"\bu32\s*=\s*6\b"),
        re.compile(r"\bu32\s*=\s*8\b"),
        re.compile(r"\bu32\s*=\s*17\b"),
        re.compile(r"\bu32\s*=\s*18\b"),
    ),
    "strict canonical ingress": (
        "preflight",
        "next_strict_field",
        "duplicate",
        ("wrong_wire", "WireTypeMismatch", "wrong wire", "wire type"),
        ("noncanonical", "NonCanonical", "non-canonical"),
    ),
    "finite bounds and checked steps": (
        "finite",
        ("NaN", "nan", "nonfinite", "NonFinite"),
        ("infinite", "Infinity", "nonfinite", "NonFinite"),
        "negative",
        ("overflow", "checked_"),
    ),
    "unknown/raw preservation": (
        "unknown",
        "raw",
        "extend_from_slice",
        "span",
    ),
    "lazy Buffa projection": (
        "buffa_keynote_chart_axis_value_settings_generated",
        ("decode_lazy_view", "decode_view"),
    ),
    "prepared rewrite": (
        "execution_requirements",
        "execute",
    ),
}
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_FLOW_MARKERS = {
    "decode report": re.compile(
        r"\bdecode_axis_value_settings"
        r"_with_report\s*\("
    ),
    "prepared rewrite": re.compile(
        r"\bprepare_axis_value_settings_rewrite\s*\("
    ),
    "execution requirements": re.compile(r"\bexecution_requirements\s*\("),
    "single prepared execute": re.compile(
        r"\b(?:prepared|rewrite)[A-Za-z0-9_]*[ \t\r\n]*\.[ \t\r\n]*execute\s*\("
    ),
}
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PACKAGE_BUDGET_TYPE = re.compile(
    r"(?m)^[ \t]*(?:pub[ \t]*\([ \t]*super[ \t]*\)[ \t]+)?struct[ \t]+"
    r"(?P<name>(?:Chart|ValueAxis|Axis|Package|Transaction)[A-Za-z0-9_]*Budget)\b"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PACKAGE_MARKER_GROUPS = {
    "aggregate transaction budget": ("Budget", "budget", "charge"),
    "strict prepared codec": (
        "keynote_chart_axis_value_settings_codec",
        "execution_requirements",
        "execute",
    ),
    "exact artifacts/fingerprint/inverse/conflict": (
        "ExactArtifacts",
        "source_fingerprint",
        "inverse",
        "PatchConflict",
    ),
    "candidate reopen/locality": ("candidate", "reopen", "locality"),
    "root preview deletion": ("root_preview_deletions", "deleted_previews"),
    "source authority": ("authority", "ChartSelector", "ValueAxisSettings"),
}
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SOURCE_AUTHORITY_MARKER = re.compile(
    r"\b(?:authority|authorizes_source|source_authority|selection_authority)\b"
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_TEST_SOURCES = (
    Path("crates/litchi-keynote/tests/slide_chart_value_axis.rs"),
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FUZZ_SOURCES = (
    Path(
        "crates/litchi-iwa-protos/fuzz/fuzz_targets/"
        "keynote_chart_axis_value_settings_codec.rs"
    ),
    Path("crates/litchi/fuzz/fuzz_targets/keynote_chart_value_axis.rs"),
)
KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FUZZ_CORPORA = (
    Path(
        "crates/litchi-iwa-protos/fuzz/corpus/"
        "keynote_chart_axis_value_settings_codec"
    ),
    Path("crates/litchi/fuzz/corpus/keynote_chart_value_axis"),
)


def _keynote_chart_axis_value_settings_owner_present(root: Path) -> bool:
    """Return whether the value-axis owner crossed its wiring seam."""

    owner_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE
    package_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_chart_value_axis\s*;",
        package_source,
    ) is not None


def _keynote_chart_axis_value_settings_public_source(
    path: Path,
) -> str:
    """Read a production Rust source with test-only items masked."""

    return _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))


def _keynote_chart_axis_value_settings_has_marker(
    source: str, marker: str | re.Pattern[str] | tuple[str, ...]
) -> bool:
    """Match a required source marker, accepting spelling alternatives."""

    if isinstance(marker, tuple):
        return any(
            _keynote_chart_axis_value_settings_has_marker(source, item)
            for item in marker
        )
    if isinstance(marker, re.Pattern):
        return marker.search(source) is not None
    return marker in source


def _keynote_chart_axis_value_settings_marker_group_present(
    source: str, markers: tuple[str | re.Pattern[str] | tuple[str, ...], ...]
) -> bool:
    return all(
        _keynote_chart_axis_value_settings_has_marker(source, marker)
        for marker in markers
    )


def _keynote_chart_axis_value_settings_focused_package_call(
    source: str, match: re.Match[str]
) -> bool:
    """Recognize a selector-only compatibility call into the focused owner."""

    prefix = source[max(0, match.start() - 360) : match.start()]
    cursor = match.end()
    depth = 1
    while cursor < len(source) and depth:
        if source[cursor] == "(":
            depth += 1
        elif source[cursor] == ")":
            depth -= 1
        cursor += 1
    suffix = source[match.end() : cursor - 1 if depth == 0 else match.end() + 480]
    if re.search(
        r"(?:focused_chart_axis_value_settings_package|"
        r"litchi_keynote[ \t\r\n]*::[ \t\r\n]*Package|"
        r"KeynotePackage|package)",
        prefix,
    ) is None:
        return False
    if re.search(r"\b(?:selector|chart_selector)\b", suffix) is None:
        return False
    if re.search(r"\bValueAxisSettings\b", suffix) is None:
        return False
    return re.search(
        r"\b(?:drawable_object_id|chart_object_id|native_id|object_id|"
        r"axis_id|identifier|chart_position)\b",
        suffix,
    ) is None


def audit_keynote_chart_axis_value_settings_legacy_calls(
    root: Path = ROOT,
) -> list[str]:
    """Keep focused Keynote value-axis code off retired host methods."""

    if not _keynote_chart_axis_value_settings_owner_present(root):
        return []
    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for match in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_LEGACY_CALL.finditer(source):
            line_number = source.count("\n", 0, match.start("method")) + 1
            violations.append(
                "focused litchi-keynote chart value-axis settings source retains "
                f"legacy call {match.group('method')}: {path.relative_to(root)}:{line_number}"
            )
    return sorted(set(violations))


def _keynote_chart_axis_value_settings_check_projection(
    root: Path,
) -> list[str]:
    """Require the bounded private Buffa projection and its provenance gate."""

    violations: list[str] = []
    projection_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PROJECTION_SOURCE
    projection = (
        projection_path.read_text(encoding="utf-8")
        if projection_path.is_file()
        else ""
    )
    if not projection:
        violations.append(
            "focused Keynote chart value-axis settings codec is missing private "
            f"Buffa projection: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PROJECTION_SOURCE}"
        )
    else:
        projection_markers = (
            "syntax = \"proto2\"",
            "message ChartAxisValueSettingsArchive",
            "tschchartaxisvaluenumberofmajorgridlines",
            "tschchartaxisvaluenumberofminorgridlines",
            "tschchartaxisvaluescale",
            "tschchartaxisdefaultusermax",
            "tschchartaxisdefaultusermin",
            "message ChartsNsNumberDoubleArchive",
            "optional double number_archive",
        )
        for marker in projection_markers:
            if marker not in projection:
                violations.append(
                    "focused Keynote chart value-axis settings private projection is "
                    f"missing marker {marker}: "
                    f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PROJECTION_SOURCE}"
                )
        # Keep field numbers pinned to the native archive contract.  Matching
        # the declaration name and number together prevents a generic ``= 5``
        # comment from satisfying the projection check.
        fields = (
            ("tschchartaxisvaluenumberofdecades", 4),
            ("tschchartaxisvaluenumberofmajorgridlines", 5),
            ("tschchartaxisvaluenumberofminorgridlines", 6),
            ("tschchartaxisvaluescale", 8),
            ("tschchartaxisdefaultusermax", 17),
            ("tschchartaxisdefaultusermin", 18),
        )
        for name, number in fields:
            if re.search(
                rf"\b(?:optional|repeated)[^\n;]*\b{re.escape(name)}\b"
                rf"[^\n;]*=[ \t]*{number}\b",
                projection,
            ) is None:
                violations.append(
                    "focused Keynote chart value-axis settings private projection has "
                    f"incorrect/missing field {name}={number}: "
                    f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PROJECTION_SOURCE}"
                )

    build_path = root / Path("crates/litchi-iwa-protos/build.rs")
    build = build_path.read_text(encoding="utf-8") if build_path.is_file() else ""
    build_markers = (
        "enforce_keynote_chart_axis_value_settings_projection_provenance",
        "TSCHChartAxisValueSettingsArchive.proto",
        "generate_views(true)",
        "lazy_views(true)",
        "preserve_unknown_fields(false)",
        "include_file(\"iwa_keynote_chart_axis_value_settings_buffa_protos.rs\")",
    )
    for marker in build_markers:
        if marker not in build:
            violations.append(
                "focused Keynote chart value-axis settings Buffa build is missing "
                f"marker {marker}: crates/litchi-iwa-protos/build.rs"
            )

    public_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_PUBLIC_SOURCE
    public_source = (
        public_path.read_text(encoding="utf-8") if public_path.is_file() else ""
    )
    if re.search(
        r"(?ms)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\].{0,240}?"
        r"^\s*mod\s+buffa_keynote_chart_axis_value_settings_generated\b",
        public_source,
    ) is None:
        violations.append(
            "focused Keynote chart value-axis settings generated Buffa projection must "
            "remain private: crates/litchi-iwa-protos/src/lib.rs"
        )
    if (
        "buffa-keynote-chart-axis-value-settings"
        not in public_source
        or "iwa_keynote_chart_axis_value_settings_buffa_protos.rs" not in public_source
    ):
        violations.append(
            "focused Keynote chart value-axis settings generated projection include is "
            "missing: crates/litchi-iwa-protos/src/lib.rs"
        )
    return violations


def _keynote_chart_axis_value_settings_check_codec(root: Path) -> list[str]:
    """Require a strict borrowed codec without eager generated decoding."""

    codec_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE
    if not codec_path.is_file():
        return [
            "focused litchi-keynote chart value-axis settings public API is missing "
            f"strict neutral codec source: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
        ]
    raw_source = codec_path.read_text(encoding="utf-8")
    source = _mask_rust_cfg_test_items(raw_source)
    code = _mask_rust_non_code(source)
    violations: list[str] = []
    for api in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_REQUIRED_APIS:
        if re.search(
            rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
            code,
        ) is None:
            violations.append(
                "focused litchi-keynote chart value-axis settings hidden codec is "
                f"missing strict API {api}: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
            )
    for type_name in sorted(KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_PREPARED_TYPES):
        if re.search(
            rf"\b(?:pub\s+)?(?:struct|enum|type)\s+{re.escape(type_name)}\b",
            code,
        ) is None:
            violations.append(
                "focused litchi-keynote chart value-axis settings hidden codec is "
                f"missing prepared type {type_name}: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
            )

    for label, markers in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not _keynote_chart_axis_value_settings_marker_group_present(code, markers):
            violations.append(
                "focused litchi-keynote chart value-axis settings hidden codec is "
                f"missing {label} marker: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
            )

    # ``prost`` and eager generated Message decode/encode are forbidden in
    # production.  The private Buffa lazy-view module is the only generated
    # ingress permitted, and raw source spans remain authoritative.
    eager_patterns = (
        re.compile(r"\bprost(?:_types)?\b"),
        re.compile(r"\bMessage\s*::\s*decode\s*\("),
        re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\s*::\s*decode\s*\("),
        re.compile(r"\b(?:encode_to_vec|encode_into|merge_from)\s*\("),
        re.compile(r"\.[ \t\r\n]*encode_to_vec\s*\("),
        re.compile(r"\.[ \t\r\n]*encode\s*\("),
        re.compile(r"\b(?:to_owned_message|try_encode)\s*\("),
    )
    for pattern in eager_patterns:
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote chart value-axis settings hidden codec must not "
                f"use eager generated decode/encode {match.group(0).strip()}: "
                f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}:{line_number}"
            )
    for marker in ("HostPackage", "chart_graph"):
        if marker in code:
            violations.append(
                "focused litchi-keynote chart value-axis settings hidden codec must remain "
                f"neutral and generated-free ({marker}): {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
            )

    codec_lib_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_PUBLIC_SOURCE
    codec_lib = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_MODULE)}\b",
        codec_lib,
    ) is None:
        violations.append(
            "focused litchi-keynote chart value-axis settings public API is missing hidden "
            f"codec module {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_MODULE}: "
            f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(r"(?m)#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\]", raw_source) is None:
        violations.append(
            "focused litchi-keynote chart value-axis settings hidden codec is missing "
            f"cfg(test) coverage: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
        )
    elif re.search(r"(?m)#\s*\[\s*test\s*\]", raw_source) is None:
        violations.append(
            "focused litchi-keynote chart value-axis settings hidden codec is missing "
            f"#[test] coverage: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
        )
    violations.extend(_keynote_chart_axis_value_settings_check_projection(root))
    return sorted(set(violations))


def audit_keynote_chart_axis_value_settings_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the archive-free aggregate Package owner and semantic exports."""

    if not _keynote_chart_axis_value_settings_owner_present(root):
        return []
    owner_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE
    semantic_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SELECTOR_SOURCE
    package_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EXPORT_SOURCES[1]
    common_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMMON_SOURCE
    common_axis_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMMON_AXIS_SOURCE
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _keynote_chart_axis_value_settings_public_source(path)
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    # Keep raw/wire/archive compatibility namespaces out of the focused
    # public surface.  The neutral codec's hidden module lives in
    # ``litchi-iwa-protos`` and is audited separately, so this scan is limited
    # to the Keynote semantic/package export files above.
    for path in paths:
        for pattern, label in (
            (KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMPATIBILITY_MODULE, "module"),
            (KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMPATIBILITY_REEXPORT, "re-export"),
        ):
            for match in pattern.finditer(code[path]):
                line_number = code[path].count("\n", 0, match.start()) + 1
                violations.append(
                    "focused litchi-keynote chart value-axis settings public API retains "
                    f"compatibility {label}: {path.relative_to(root)}:{line_number}"
                )

    common_source = common_path.read_text(encoding="utf-8") if common_path.is_file() else ""
    common_axis_source = (
        common_axis_path.read_text(encoding="utf-8")
        if common_axis_path.is_file()
        else ""
    )
    if not re.search(r"\bpub\s+struct\s+ValueAxisSettings\b", common_source):
        violations.append(
            "focused Keynote chart value-axis settings semantic API is missing common "
            f"ValueAxisSettings: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMMON_SOURCE}"
        )
    if "ValueAxisSettings" not in common_axis_source:
        violations.append(
            "focused Keynote chart value-axis settings common axis module is missing "
            f"ValueAxisSettings export: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_COMMON_AXIS_SOURCE}"
        )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SEMANTIC_TYPES
    )
    for name in sorted(
        KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SEMANTIC_TYPES - semantic_exports
    ):
        violations.append(
            "focused litchi-keynote chart value-axis settings semantic API is missing "
            f"{name}: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[semantic_path] + sources[lib_path],
        KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SELECTOR_TYPES,
    )
    for name in sorted(
        KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SELECTOR_TYPES - selector_exports
    ):
        violations.append(
            "focused litchi-keynote chart value-axis settings public API is missing "
            f"selector {name}: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EXPORT_SOURCES[1]}"
        )

    package_code = code[package_path]
    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+slide_chart_value_axis\b",
        package_code,
    ):
        violations.append(
            "focused litchi-keynote chart value-axis settings owner module must remain "
            f"private: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_chart_value_axis\s*;",
        package_code,
    ) is None:
        violations.append(
            "focused litchi-keynote chart value-axis settings owner module is missing: "
            f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FAMILY_MODULE), KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote chart value-axis settings public API is "
                    f"missing canonical type {name}: {path.relative_to(root)}"
                )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote chart value-axis settings Package method is "
                f"missing {name}: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_chart_value_axis_settings":
            for selector in ("SlideSelector", "ChartSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote chart value-axis settings Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
                    )
            if (
                name == "slide_chart_value_axis_settings"
                and "ValueAxisSettings" not in declaration
            ):
                violations.append(
                    "focused litchi-keynote chart value-axis settings Package method "
                    f"{name} must expose ValueAxisSettings: "
                    f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
                )
        elif "Patch" not in declaration:
            violations.append(
                "focused litchi-keynote chart value-axis settings apply method must accept "
                f"Patch: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
            )

    edit_methods = {
        name
        for name, _declaration, _line in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(KEYNOTE_CHART_AXIS_VALUE_SETTINGS_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote chart value-axis settings edit is missing "
            f"{name}: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CANONICAL_TYPES
        | KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SEMANTIC_TYPES
        | KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SELECTOR_TYPES
        | KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PACKAGE_METHODS
        | KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path, selector_path}
    for path in paths:
        dedicated = path in dedicated_sources
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not identifiers.intersection(facade_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif (
                    identifier == "wire"
                    or identifier in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_WIRE_TYPES
                ):
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote chart value-axis settings public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote chart value-axis settings public API retains "
                        f"flat alias {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote chart value-axis settings public API exposes raw "
                    f"byte slice {byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PUBLIC_RAW_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote chart value-axis settings public API exposes raw "
                    f"parameter {match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote chart value-axis settings public API retains a "
                    f"glob re-export: {path.relative_to(root)}:{line_number}"
                )

    owner_code = code[owner_path]
    codec_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE
    codec_code = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        )
        if codec_path.is_file()
        else ""
    )
    # The package owns graph selection and the codec owns the prepared wire
    # execution.  Audit the concatenated production slices so an adapter may
    # keep the one execution helper private to the codec without duplicating
    # a rewrite transaction in the package owner.
    prepared_flow = owner_code + "\n" + codec_code
    for label, marker in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_FLOW_MARKERS.items():
        if marker.search(prepared_flow) is None:
            violations.append(
                "focused litchi-keynote chart value-axis settings owner is missing strict "
                f"prepared {label} marker: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
            )
    violations.extend(_keynote_chart_axis_value_settings_check_codec(root))
    return sorted(set(violations))


def audit_keynote_chart_axis_value_settings_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one aggregate bounded transaction over graph and codec work."""

    if not _keynote_chart_axis_value_settings_owner_present(root):
        return []
    owner_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    codec_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE
    codec = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        )
        if codec_path.is_file()
        else ""
    )
    prepared_source = owner + "\n" + codec
    violations: list[str] = []
    budgets = list(KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PACKAGE_BUDGET_TYPE.finditer(owner))
    if len(budgets) != 1:
        violations.append(
            "focused litchi-keynote chart value-axis settings owner must define exactly one "
            f"aggregate transaction budget (found {len(budgets)}): "
            f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
        )
    for label, markers in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_PACKAGE_MARKER_GROUPS.items():
        if label == "source authority":
            present = (
                KEYNOTE_CHART_AXIS_VALUE_SETTINGS_SOURCE_AUTHORITY_MARKER.search(owner)
                is not None
                and all(marker in owner for marker in markers if marker != "authority")
            )
        else:
            present = all(marker in prepared_source for marker in markers)
        if not present:
            violations.append(
                "focused litchi-keynote chart value-axis settings owner is missing "
                f"{label} marker: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
            )
    if any(
        operation in owner
        for operation in ("saturating_add", "saturating_sub", "saturating_mul")
    ):
        violations.append(
            "focused litchi-keynote chart value-axis settings owner must use checked "
            f"resource accounting: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def audit_keynote_chart_axis_value_settings_completion_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require integration, low/high fuzz, and one prepared execution."""

    if not _keynote_chart_axis_value_settings_owner_present(root):
        return []
    owner_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    codec_path = root / KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE
    codec = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        )
        if codec_path.is_file()
        else ""
    )
    prepared_source = owner + "\n" + codec
    violations: list[str] = []
    for label, marker in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_FLOW_MARKERS.items():
        if marker.search(prepared_source) is None:
            violations.append(
                "focused litchi-keynote chart value-axis settings completion is missing "
                f"{label}: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
            )
    # The Package owner must issue one prepared execution.  The neutral codec
    # also exposes a convenience rewrite function, which performs its own
    # prepared execution; counting that implementation here would incorrectly
    # couple two separate ownership layers and reject a valid facade.
    execute_calls = list(re.finditer(r"\.[ \t\r\n]*execute\s*\(", owner))
    if len(execute_calls) != 1:
        violations.append(
            "focused litchi-keynote chart value-axis settings completion must have exactly "
            f"one prepared execute call (found {len(execute_calls)}): "
            f"{KEYNOTE_CHART_AXIS_VALUE_SETTINGS_OWNER_SOURCE}"
        )

    codec_raw = codec_path.read_text(encoding="utf-8") if codec_path.is_file() else ""
    if codec_path.is_file() and re.search(r"\bexecute\s*\(", codec) is None:
        violations.append(
            "focused litchi-keynote chart value-axis settings completion is missing codec "
            f"execute implementation: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
        )
    elif not codec_path.is_file():
        violations.append(
            "focused litchi-keynote chart value-axis settings completion is missing codec "
            f"source: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
        )
    if re.search(r"(?m)#\s*\[\s*test\s*\]", codec_raw) is None:
        violations.append(
            "focused litchi-keynote chart value-axis settings completion is missing codec "
            f"#[test] coverage: {KEYNOTE_CHART_AXIS_VALUE_SETTINGS_CODEC_SOURCE}"
        )

    for test_path in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote chart value-axis settings boundary is missing "
                f"integration test: {test_path}"
            )
            continue
        test_code = _mask_rust_non_code(absolute.read_text(encoding="utf-8"))
        if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
            violations.append(
                "focused litchi-keynote chart value-axis settings integration test is "
                f"missing #[test] coverage: {test_path}"
            )
        for marker in (
            "slide_chart_value_axis_settings",
            "edit_slide_chart_value_axis_settings",
            "apply_slide_chart_value_axis_settings",
            "ValueAxisSettings",
            "Bounds",
            "Steps",
            "Scale",
        ):
            if marker not in test_code:
                violations.append(
                    "focused litchi-keynote chart value-axis settings integration test is "
                    f"missing {marker} coverage: {test_path}"
                )
        for marker in ("deleted_previews", "reopen", "inverse", "locality"):
            if marker not in test_code:
                violations.append(
                    "focused litchi-keynote chart value-axis settings integration test is "
                    f"missing {marker} verification: {test_path}"
                )

    for fuzz_path in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote chart value-axis settings boundary is missing "
                f"fuzz target: {fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote chart value-axis settings fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus in KEYNOTE_CHART_AXIS_VALUE_SETTINGS_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-keynote chart value-axis settings boundary is missing "
                f"fuzz corpus: {corpus}"
            )
    return sorted(set(violations))


def audit_keynote_movie_title_legacy_calls(root: Path = ROOT) -> list[str]:
    """Keep focused Keynote movie-title code off legacy host mutations."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    violations: list[str] = []
    legacy_call = re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?P<method>set_slide_movie_title|"
        r"remove_slide_movie_title)(?![A-Za-z0-9_])[ \t\r\n]*\("
    )
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for match in legacy_call.finditer(source):
            line_start = source.rfind("\n", 0, match.start()) + 1
            line_end = source.find("\n", match.end())
            if line_end < 0:
                line_end = len(source)
            line = source[line_start:line_end]
            if re.search(r"\bfn[ \t\r\n]+(?:set|remove)_slide_movie_title\b", line):
                continue
            line_number = source.count("\n", 0, match.start("method")) + 1
            violations.append(
                "focused litchi-keynote movie-title source retains legacy call "
                f"{match.group('method')}: {path.relative_to(root)}:{line_number}"
            )
    return sorted(set(violations))


def audit_keynote_chart_caption_legacy_calls(root: Path = ROOT) -> list[str]:
    """Keep focused chart-caption code off raw-ID host mutation calls."""

    source_root = root / KEYNOTE_SOURCE_ROOT
    if not source_root.is_dir():
        return []
    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        source = _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for match in KEYNOTE_CHART_CAPTION_LEGACY_CALL.finditer(source):
            line_number = source.count("\n", 0, match.start("method")) + 1
            violations.append(
                "focused litchi-keynote chart-caption source retains legacy call "
                f"{match.group('method')}: {path.relative_to(root)}:{line_number}"
            )
    return sorted(set(violations))


def audit_keynote_chart_caption_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the selector-first, archive-free chart-caption package owner.

    The package may keep physical parsing and wire helpers private, but none
    of that vocabulary may cross its public declarations or re-export lists.
    This distinction matters: scanning the whole implementation would reject
    the owner for doing its job, while scanning only methods misses public
    fields and ``pub use`` aliases that expose the same physical graph.
    """

    owner_path = root / KEYNOTE_CHART_CAPTION_OWNER_SOURCE
    package_path = root / KEYNOTE_CHART_CAPTION_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_CHART_CAPTION_EXPORT_SOURCES[1]
    violations: list[str] = []
    if not owner_path.is_file():
        return [
            "focused litchi-keynote chart-caption owner source is missing: "
            f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        ]
    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package)
    package_library_code = _mask_rust_non_code(package + library)
    if re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_chart_caption[ \t]*;", package_code) is None:
        violations.append(
            "focused litchi-keynote chart-caption owner module is missing: "
            f"{KEYNOTE_CHART_CAPTION_EXPORT_SOURCES[0]}"
        )
    if re.search(r"(?m)^pub[ \t]+mod[ \t]+slide_chart_caption\b", package_library_code):
        violations.append(
            "focused litchi-keynote chart-caption owner module must remain private: "
            f"{KEYNOTE_CHART_CAPTION_EXPORT_SOURCES[0]}"
        )
    canonical_names = frozenset(KEYNOTE_CHART_CAPTION_CANONICAL_TYPES)
    owner_exports = _rust_canonical_exports(owner, canonical_names)
    package_exports = _rust_canonical_exports(
        package, canonical_names
    )
    lib_exports = _rust_canonical_exports(
        _keynote_family_view(
            root, root / KEYNOTE_CHART_CAPTION_FAMILY_SOURCE, {}, KEYNOTE_CHART_CAPTION_FAMILY_MODULE
        ),
        canonical_names,
    )
    for name in sorted(KEYNOTE_CHART_CAPTION_CANONICAL_TYPES):
        if name not in owner_exports or name not in lib_exports:
            violations.append(
                "focused litchi-keynote chart-caption public API is missing canonical type "
                f"{name}: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
            )
    owner_code = _mask_rust_non_code(owner)

    def function_bodies(code: str) -> list[tuple[str, str, int]]:
        """Return production Rust function bodies with source offsets."""

        bodies: list[tuple[str, str, int]] = []
        for declaration in re.finditer(
            r"(?<![A-Za-z0-9_#])fn[ \t\r\n]+(?:r#)?"
            r"([A-Za-z_][A-Za-z0-9_]*)\b",
            code,
        ):
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
            bodies.append(
                (
                    declaration.group(1),
                    code[opening + 1 : cursor - 1],
                    declaration.start(),
                )
            )
        return bodies

    owner_functions = function_bodies(owner_code)
    owner_function_map = {name: body for name, body, _offset in owner_functions}

    def function_signature(name: str) -> str:
        declaration = re.search(
            rf"(?<![A-Za-z0-9_#])fn[ \t\r\n]+(?:r#)?{re.escape(name)}\b",
            owner_code,
        )
        if declaration is None:
            return ""
        opening = owner_code.find("{", declaration.end())
        return owner_code[declaration.start() : opening] if opening >= 0 else ""

    def called_function_names(body: str) -> set[str]:
        return {
            match.group(1)
            for match in re.finditer(
                r"(?<![A-Za-z0-9_#])(?:self[ \t\r\n]*\.[ \t\r\n]*)?"
                r"([A-Za-z_][A-Za-z0-9_]*)[ \t\r\n]*\(",
                body,
            )
        }

    def reachable_bodies(seed_names: set[str]) -> list[str]:
        reachable: list[str] = []
        visited: set[str] = set()
        pending = list(seed_names)
        while pending:
            name = pending.pop()
            if name in visited:
                continue
            visited.add(name)
            body = owner_function_map.get(name)
            if body is None:
                continue
            # Retain the helper name as an accounting marker. Resource charge
            # helpers deliberately keep their semantic category in the name,
            # while their bodies may operate only on generic byte counters.
            reachable.append(f"fn {name}\n{body}")
            pending.extend(called_function_names(body) - visited)
        return reachable

    budget_declarations = list(
        KEYNOTE_CHART_CAPTION_PACKAGE_BUDGET_TYPE.finditer(owner_code)
    )
    if not budget_declarations:
        violations.append(
            "focused litchi-keynote chart-caption owner is missing one private aggregate "
            f"package budget: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )
        budget_name = None
    else:
        if len(budget_declarations) != 1:
            violations.append(
                "focused litchi-keynote chart-caption owner must define exactly one private "
                f"aggregate package budget: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
            )
        budget_name = budget_declarations[0].group("name")

    def require_budget_thread(
        name: str,
        label: str,
        *,
        require_parameter: bool = True,
    ) -> str | None:
        body = owner_function_map.get(name)
        if body is None:
            violations.append(
                "focused litchi-keynote chart-caption transaction is missing "
                f"{label} function {name}: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
            )
            return None
        signature = function_signature(name)
        if require_parameter and (
            budget_name is None
            or budget_name not in signature
            or not re.search(
                r"\b(?:budget|package_budget|aggregate_budget|transaction_budget)\b",
                signature,
            )
        ):
            violations.append(
                "focused litchi-keynote chart-caption "
                f"{label} must thread the private aggregate package budget: "
                f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
            )
        if not re.search(
            r"\b(?:budget|package_budget|aggregate_budget|transaction_budget)\b",
            body,
        ):
            violations.append(
                "focused litchi-keynote chart-caption "
                f"{label} must use the threaded aggregate package budget: "
                f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
            )
        return body

    commit_body = require_budget_thread(
        "commit", "commit", require_parameter=False
    )
    apply_body = require_budget_thread(
        "apply_slide_chart_caption", "apply", require_parameter=False
    )
    if commit_body is not None and not re.search(
        r"\brewrite_chart_caption_operation\b[^;{}]*\bbudget\b",
        commit_body,
        re.DOTALL,
    ):
        violations.append(
            "focused litchi-keynote chart-caption commit must pass its one aggregate "
            f"package budget into the graph/existing rewrite: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )
    if apply_body is not None and not re.search(
        r"\bverify_caption_candidate\b[^;{}]*\bbudget\b",
        apply_body,
        re.DOTALL,
    ):
        violations.append(
            "focused litchi-keynote chart-caption apply must pass its one aggregate "
            f"package budget into candidate verification: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )

    def has_budget_charge(body: str, category: str) -> bool:
        if KEYNOTE_CHART_CAPTION_RESOURCE_CHARGES[category].search(body):
            return True
        # Permit a generic charge method when the resource category is named
        # in its argument.  This keeps the ratchet independent of whether the
        # owner uses charge_archive_work or charge_bytes(archive_bytes, ...).
        return re.search(
            rf"(?is)\b(?:budget|package_budget|aggregate_budget|transaction_budget)\b"
            rf".{{0,120}}\b(?:charge|account|reserve)\w*\b.{{0,120}}\b{category}\w*\b"
            rf"|\b{category}\w*\b.{{0,120}}\b(?:charge|account|reserve)\w*\b",
            body,
        ) is not None

    def require_transaction_charges(
        path_label: str,
        seed_names: set[str],
        *,
        require_codec_report: bool = True,
        categories: tuple[str, ...] = ("archive", "snappy", "zip", "reopen", "artifact"),
    ) -> None:
        path_bodies = reachable_bodies(seed_names)
        path_blob = "\n".join(path_bodies)
        if require_codec_report and not KEYNOTE_CHART_CAPTION_CODEC_REPORT_CHARGE.search(path_blob):
            violations.append(
                "focused litchi-keynote chart-caption "
                f"{path_label} must charge nested codec reports through the aggregate "
                f"package budget: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
            )
        for category in categories:
            if not has_budget_charge(path_blob, category):
                violations.append(
                    "focused litchi-keynote chart-caption "
                    f"{path_label} must charge prepublication {category} work through "
                    f"the aggregate package budget: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
                )

    for name, label in (
        ("rewrite_chart_caption_operation", "graph create/remove"),
    ):
        if name in owner_function_map:
            require_budget_thread(name, label)
            require_transaction_charges(
                label,
                {name},
                categories=("archive", "snappy", "zip", "reopen"),
            )
    if apply_body is not None:
        require_transaction_charges(
            "apply",
            {"apply_slide_chart_caption"},
            require_codec_report=False,
            categories=("reopen", "artifact"),
        )
    if commit_body is not None and not has_budget_charge(
        "\n".join(reachable_bodies({"commit"})), "artifact"
    ):
        violations.append(
            "focused litchi-keynote chart-caption commit must charge exact-artifact work "
            "through the aggregate package budget: "
            f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )
    operation = next(
        (
            (body, offset)
            for name, body, offset in owner_functions
            if name == KEYNOTE_CHART_CAPTION_GRAPH_OPERATION
        ),
        None,
    )
    if operation is None:
        violations.append(
            "focused litchi-keynote chart-caption graph operation helper is missing "
            f"{KEYNOTE_CHART_CAPTION_GRAPH_OPERATION}: "
            f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )
    else:
        operation_body, operation_offset = operation
        operation_path = "\n".join(
            [operation_body, *reachable_bodies({KEYNOTE_CHART_CAPTION_GRAPH_OPERATION})]
        )
        for marker in sorted(KEYNOTE_CHART_CAPTION_OBJECT_FRAMING_MARKERS):
            if marker in operation_path:
                continue
            line_number = owner_code.count("\n", 0, operation_offset) + 1
            violations.append(
                "focused litchi-keynote chart-caption graph create/remove must validate "
                f"canonical object framing ({marker}): "
                f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}:{line_number}"
            )

        existing_branch = re.search(
            r"\bif[ \t\r\n]+!creating[ \t\r\n]*&&[ \t\r\n]*!removing\b",
            operation_body,
        )
        if existing_branch is None:
            line_number = owner_code.count("\n", 0, operation_offset) + 1
            violations.append(
                "focused litchi-keynote chart-caption graph operation must expose "
                "the changed existing-caption path: "
                f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}:{line_number}"
            )
        else:
            branch_opening = operation_body.find("{", existing_branch.end())
            branch_depth = 1
            branch_cursor = branch_opening + 1 if branch_opening >= 0 else len(operation_body)
            while branch_opening >= 0 and branch_cursor < len(operation_body) and branch_depth:
                if operation_body[branch_cursor] == "{":
                    branch_depth += 1
                elif operation_body[branch_cursor] == "}":
                    branch_depth -= 1
                branch_cursor += 1
            branch_body = (
                operation_body[branch_opening + 1 : branch_cursor - 1]
                if branch_depth == 0
                else ""
            )
            called_helpers = {
                match.group(1)
                for match in re.finditer(
                    r"\b([A-Za-z_][A-Za-z0-9_]*)[ \t\r\n]*\(",
                    branch_body,
                )
            }
            reachable_branch_bodies = [branch_body]
            reachable_branch_bodies.extend(reachable_bodies(called_helpers))
            existing_path_blob = "\n".join(reachable_branch_bodies)
            if not re.search(
                r"\b(?:budget|package_budget|aggregate_budget|transaction_budget)\b",
                existing_path_blob,
            ):
                violations.append(
                    "focused litchi-keynote chart-caption existing-caption rewrite must "
                    "thread the aggregate package budget: "
                    f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
                )
            if not any(
                marker in body
                for body in reachable_branch_bodies
                for marker in KEYNOTE_CHART_CAPTION_SAVE_TOKEN_REWRITE_MARKERS
            ):
                line_number = owner_code.count(
                    "\n", 0, operation_offset + existing_branch.start()
                ) + 1
                violations.append(
                    "focused litchi-keynote chart-caption existing-caption rewrite must "
                    "advance Metadata save tokens: "
                    f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}:{line_number}"
                )
            if not KEYNOTE_CHART_CAPTION_CODEC_REPORT_CHARGE.search(existing_path_blob):
                violations.append(
                    "focused litchi-keynote chart-caption existing-caption rewrite must "
                    "charge nested codec reports through the aggregate package budget: "
                    f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
                )
            for category in ("archive", "snappy", "zip", "reopen"):
                if not has_budget_charge(existing_path_blob, category):
                    violations.append(
                        "focused litchi-keynote chart-caption existing-caption rewrite "
                        f"must charge prepublication {category} work through the aggregate "
                        f"package budget: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
                    )

    transition = next(
        (
            (body, offset)
            for name, body, offset in owner_functions
            if name in KEYNOTE_CHART_CAPTION_REFERENCE_TRANSITIONS
        ),
        None,
    )
    if transition is None:
        violations.append(
            "focused litchi-keynote chart-caption source-authoritative reference "
            f"transition helper is missing {KEYNOTE_CHART_CAPTION_REFERENCE_TRANSITION}: "
            f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )
    else:
        transition_body, transition_offset = transition
        for marker in sorted(KEYNOTE_CHART_CAPTION_REFERENCE_TRANSITION_MARKERS):
            if marker in transition_body:
                continue
            line_number = owner_code.count("\n", 0, transition_offset) + 1
            violations.append(
                "focused litchi-keynote chart-caption reference transition must retain "
                f"source-authoritative marker {marker}: "
                f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}:{line_number}"
            )

    if any(marker in owner_code for marker in KEYNOTE_CHART_CAPTION_CROSS_COMPONENT_MARKERS):
        if not any(
            marker in owner_code
            for marker in KEYNOTE_CHART_CAPTION_EXTERNAL_REFERENCE_MARKERS
        ):
            violations.append(
                "focused litchi-keynote chart-caption cross-component graph dependency "
                "must carry external-reference attribution: "
                f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
            )

    for name, body, offset in owner_functions:
        plain_replace = re.search(
            r"\.replace_message_preserving_header(?!_)[ \t\r\n]*\(", body
        )
        if plain_replace is None:
            continue
        reference_mutation = re.search(
            r"(?:object_references|field_infos|data_references)\b", body
        )
        if reference_mutation is None or reference_mutation.start() > plain_replace.start():
            continue
        line_number = owner_code.count("\n", 0, offset) + 1
        violations.append(
            "focused litchi-keynote chart-caption MessageInfo reference mutation must "
            "use the source-authoritative transitioning replacement helper: "
            f"{KEYNOTE_CHART_CAPTION_OWNER_SOURCE}:{line_number}"
        )

    method_names = {
        match.group(1)
        for match in re.finditer(
            r"(?<![A-Za-z0-9_#])pub[ \t\r\n]+fn[ \t\r\n]+"
            r"([A-Za-z_][A-Za-z0-9_]*)\b",
            owner_code,
        )
    }
    for name in sorted(KEYNOTE_CHART_CAPTION_PACKAGE_METHODS - method_names):
        violations.append(
            "focused litchi-keynote chart-caption Package method is missing "
            f"{name}: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )

    # Require the complete edit lifecycle on Edit itself.  A
    # similarly named Package helper must not satisfy this check accidentally.
    edit_impl = re.search(
        r"(?<![A-Za-z0-9_#])impl(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]+"
        r"(?:'[^ ]+[ \t\r\n]+)?Edit\b",
        owner_code,
    )
    if edit_impl is None:
        edit_names: set[str] = set()
    else:
        opening = owner_code.find("{", edit_impl.end())
        depth = 1
        cursor = opening + 1 if opening >= 0 else len(owner_code)
        while opening >= 0 and cursor < len(owner_code) and depth:
            if owner_code[cursor] == "{":
                depth += 1
            elif owner_code[cursor] == "}":
                depth -= 1
            cursor += 1
        edit_body = owner_code[opening + 1 : cursor - 1] if depth == 0 else ""
        edit_names = {
            match.group(1)
            for match in re.finditer(
                r"(?<![A-Za-z0-9_#])pub[ \t\r\n]+fn[ \t\r\n]+"
                r"([A-Za-z_][A-Za-z0-9_]*)\b",
                edit_body,
            )
        }
    for name in sorted(KEYNOTE_CHART_CAPTION_EDIT_METHODS - edit_names):
        violations.append(
            "focused litchi-keynote chart-caption Edit method is missing "
            f"{name}: {KEYNOTE_CHART_CAPTION_OWNER_SOURCE}"
        )

    # Inspect every public declaration, not just methods.  This catches
    # public struct fields, aliases, and re-exports while leaving private
    # package implementation imports free to use the archive/wire layers.
    for source, source_path in (
        (owner, owner_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if source_path != owner_path and not (
                identifiers
                & (
                    (canonical_names - KEYNOTE_FAMILY_GENERIC_NAMES)
                    | KEYNOTE_CHART_CAPTION_FLAT_ALIASES
                    | KEYNOTE_CHART_CAPTION_PACKAGE_METHODS
                    | {"slide_chart_caption"}
                )
            ):
                # package.rs/lib.rs contain many unrelated public archive
                # helpers.  Only declarations that can participate in this
                # caption facade are in scope; the owner itself is scanned in
                # full, including its public fields.
                continue
            for identifier in sorted(identifiers):
                reason = None
                if identifier in KEYNOTE_CHART_CAPTION_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_CHART_CAPTION_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_CHART_CAPTION_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is None:
                    continue
                violations.append(
                    "focused litchi-keynote chart-caption public API exposes "
                    f"{reason} {identifier}: {source_path.relative_to(root)}:{line_number}"
                )

            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote chart-caption public API exposes "
                    f"raw byte slice {byte_slice}: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )

            for match in KEYNOTE_CHART_CAPTION_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote chart-caption public API exposes "
                    f"raw identifier parameter {match.group(0).strip()}: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_CHART_CAPTION_RAW_RESOURCE_ID_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote chart-caption public API exposes "
                    f"raw resource/native identifier parameter {match.group(0).strip()}: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )

            if "pub use" in declaration or re.search(r"\bpub[ \t]+type\b", declaration):
                aliases = {
                    match.group(1)
                    for match in RUST_IDENTIFIER.finditer(declaration)
                    if match.group(1) in KEYNOTE_CHART_CAPTION_FLAT_ALIASES
                }
                for alias in sorted(aliases):
                    if alias in canonical_names:
                        continue
                    violations.append(
                        "focused litchi-keynote chart-caption public API retains flat alias "
                        f"{alias}: {source_path.relative_to(root)}:{line_number}"
                    )
                for match in re.finditer(
                    r"(?<![A-Za-z0-9_#])(?:r#)?([A-Za-z_][A-Za-z0-9_]*)"
                    r"[ \t\r\n]+as[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)\b",
                    declaration,
                ):
                    original, public_name = match.groups()
                    if original in canonical_names and public_name != original:
                        violations.append(
                            "focused litchi-keynote chart-caption public API retains "
                            f"alternate alias {public_name} for {original}: "
                            f"{source_path.relative_to(root)}:{line_number}"
                        )
                type_alias = re.search(
                    r"\bpub[ \t]+type[ \t]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)"
                    r"(?:[ \t]*<[^;=]*>)?[ \t]*=[ \t]*(?:[A-Za-z_][A-Za-z0-9_]*::)*"
                    r"(?:r#)?([A-Za-z_][A-Za-z0-9_]*)\b",
                    declaration,
                )
                if type_alias is not None:
                    alias, target = type_alias.groups()
                    if target in canonical_names and alias != target:
                        violations.append(
                            "focused litchi-keynote chart-caption public API retains "
                            f"alternate alias {alias} for {target}: "
                            f"{source_path.relative_to(root)}:{line_number}"
                        )

    return sorted(set(violations))


def audit_keynote_movie_caption_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the selector-first Keynote movie-caption package boundary.

    Movie graph creation/removal remains a compatibility-host concern in the
    first owner slice.  The package boundary is nevertheless checked as a
    complete semantic surface: a typed movie selector, a private package
    module, canonical transaction types, and the three package operations.
    Physical archive/wire types and native identifiers may remain inside the
    private owner, but cannot cross its public declarations or re-exports.
    """

    owner_path = root / KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE
    selector_path = root / KEYNOTE_MOVIE_CAPTION_SELECTOR_SOURCE
    package_path = root / KEYNOTE_MOVIE_CAPTION_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_MOVIE_CAPTION_EXPORT_SOURCES[1]
    if not owner_path.is_file():
        return [
            "focused litchi-keynote movie-caption owner source is missing: "
            f"{KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        ]

    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    selector = (
        _mask_rust_cfg_test_items(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    violations: list[str] = []
    package_code = _mask_rust_non_code(package)
    package_library_code = _mask_rust_non_code(package + library)
    if re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_movie_caption[ \t]*;", package_code) is None:
        violations.append(
            "focused litchi-keynote movie-caption owner module is missing: "
            f"{KEYNOTE_MOVIE_CAPTION_EXPORT_SOURCES[0]}"
        )
    if re.search(r"(?m)^pub[ \t]+mod[ \t]+slide_movie_caption\b", package_library_code):
        violations.append(
            "focused litchi-keynote movie-caption owner module must remain private: "
            f"{KEYNOTE_MOVIE_CAPTION_EXPORT_SOURCES[0]}"
        )

    canonical_names = frozenset(KEYNOTE_MOVIE_CAPTION_CANONICAL_TYPES)
    owner_exports = _rust_canonical_exports(owner, canonical_names)
    package_exports = _rust_canonical_exports(package, canonical_names)
    lib_exports = _rust_canonical_exports(
        _keynote_family_view(
            root, root / KEYNOTE_MOVIE_CAPTION_FAMILY_SOURCE, {}, KEYNOTE_MOVIE_CAPTION_FAMILY_MODULE
        ),
        canonical_names,
    )
    for name in sorted(canonical_names):
        if name not in owner_exports or name not in lib_exports:
            violations.append(
                "focused litchi-keynote movie-caption public API is missing canonical type "
                f"{name}: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )

    selector_names = frozenset(KEYNOTE_MOVIE_CAPTION_SELECTOR_TYPES)
    selector_exports = _rust_canonical_exports(selector, selector_names)
    lib_selector_exports = _rust_canonical_exports(library, selector_names)
    if "MovieSelector" not in selector_exports:
        violations.append(
            "focused litchi-keynote movie-caption selector source is missing canonical "
            f"MovieSelector: {KEYNOTE_MOVIE_CAPTION_SELECTOR_SOURCE}"
        )
    if "MovieSelector" not in lib_selector_exports:
        violations.append(
            "focused litchi-keynote movie-caption public API is missing canonical "
            f"MovieSelector re-export: {KEYNOTE_MOVIE_CAPTION_EXPORT_SOURCES[1]}"
        )

    owner_code = _mask_rust_non_code(owner)
    method_names = {
        match.group(1)
        for match in re.finditer(
            r"(?<![A-Za-z0-9_#])pub[ \t\r\n]+fn[ \t\r\n]+"
            r"([A-Za-z_][A-Za-z0-9_]*)\b",
            owner_code,
        )
    }
    for name in sorted(KEYNOTE_MOVIE_CAPTION_PACKAGE_METHODS - method_names):
        violations.append(
            "focused litchi-keynote movie-caption Package method is missing "
            f"{name}: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )

    # Keep the edit lifecycle typed as well.  This prevents a public Package
    # setter from becoming the only mutation path while clear/apply semantics
    # silently fall back to a raw graph helper.
    edit_impl = re.search(
        r"(?<![A-Za-z0-9_#])impl(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]+"
        r"(?:'[^ ]+[ \t\r\n]+)?Edit\b",
        owner_code,
    )
    if edit_impl is None:
        edit_names: set[str] = set()
    else:
        opening = owner_code.find("{", edit_impl.end())
        depth = 1
        cursor = opening + 1 if opening >= 0 else len(owner_code)
        while opening >= 0 and cursor < len(owner_code) and depth:
            if owner_code[cursor] == "{":
                depth += 1
            elif owner_code[cursor] == "}":
                depth -= 1
            cursor += 1
        edit_body = owner_code[opening + 1 : cursor - 1] if depth == 0 else ""
        edit_names = {
            match.group(1)
            for match in re.finditer(
                r"(?<![A-Za-z0-9_#])pub[ \t\r\n]+fn[ \t\r\n]+"
                r"([A-Za-z_][A-Za-z0-9_]*)\b",
                edit_body,
            )
        }
    for name in sorted(KEYNOTE_MOVIE_CAPTION_EDIT_METHODS - edit_names):
        violations.append(
            "focused litchi-keynote movie-caption Edit method is missing "
            f"{name}: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )

    # Public declarations in package.rs/lib.rs are filtered to this facade;
    # the owner and selector files are dedicated sources and are scanned in
    # full.  cfg(test) items are already masked item-by-item above.
    facade_names = (
        canonical_names
        | selector_names
        | KEYNOTE_MOVIE_CAPTION_PACKAGE_METHODS
        | KEYNOTE_MOVIE_CAPTION_FLAT_ALIASES
        | {"Edit"}
    )
    for source, source_path in (
        (owner, owner_path),
        (selector, selector_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = source_path in {owner_path, selector_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_MOVIE_CAPTION_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_MOVIE_CAPTION_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_MOVIE_CAPTION_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote movie-caption public API exposes "
                        f"{reason} {identifier}: {source_path.relative_to(root)}:{line_number}"
                    )

            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote movie-caption public API exposes raw byte slice "
                    f"{byte_slice}: {source_path.relative_to(root)}:{line_number}"
                )

            for match in KEYNOTE_MOVIE_CAPTION_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote movie-caption public API exposes raw identifier "
                    f"parameter {match.group(0).strip()}: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )

            for alias in sorted(identifiers & KEYNOTE_MOVIE_CAPTION_FLAT_ALIASES):
                if alias in canonical_names:
                    continue
                violations.append(
                    "focused litchi-keynote movie-caption public API retains flat alias "
                    f"{alias}: {source_path.relative_to(root)}:{line_number}"
                )

            if "pub use" in declaration or re.search(r"\bpub[ \t]+type\b", declaration):
                aliases = {
                    match.group(1)
                    for match in RUST_IDENTIFIER.finditer(declaration)
                    if match.group(1) in KEYNOTE_MOVIE_CAPTION_FLAT_ALIASES
                }
                for alias in sorted(aliases):
                    if alias in canonical_names:
                        continue
                    violations.append(
                        "focused litchi-keynote movie-caption public API retains flat alias "
                        f"{alias}: {source_path.relative_to(root)}:{line_number}"
                    )
                for match in re.finditer(
                    r"(?<![A-Za-z0-9_#])(?:r#)?([A-Za-z_][A-Za-z0-9_]*)"
                    r"[ \t\r\n]+as[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)\b",
                    declaration,
                ):
                    original, public_name = match.groups()
                    if (
                        original in canonical_names | selector_names
                        and public_name != original
                    ):
                        violations.append(
                            "focused litchi-keynote movie-caption public API retains "
                            f"alternate alias {public_name} for {original}: "
                            f"{source_path.relative_to(root)}:{line_number}"
                        )

    return sorted(set(violations))


def audit_keynote_movie_title_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the selector-first, archive-free movie-title package owner."""

    owner_path = root / KEYNOTE_MOVIE_TITLE_OWNER_SOURCE
    selector_path = root / KEYNOTE_MOVIE_TITLE_SELECTOR_SOURCE
    package_path = root / KEYNOTE_MOVIE_TITLE_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_MOVIE_TITLE_EXPORT_SOURCES[1]
    if not owner_path.is_file():
        return [
            "focused litchi-keynote movie-title owner source is missing: "
            f"{KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        ]

    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    selector = (
        _mask_rust_cfg_test_items(selector_path.read_text(encoding="utf-8"))
        if selector_path.is_file()
        else ""
    )
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    violations: list[str] = []
    package_code = _mask_rust_non_code(package)
    package_library_code = _mask_rust_non_code(package + library)
    if re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_movie_title[ \t]*;", package_code) is None:
        violations.append(
            "focused litchi-keynote movie-title owner module is missing: "
            f"{KEYNOTE_MOVIE_TITLE_EXPORT_SOURCES[0]}"
        )
    if re.search(r"(?m)^pub[ \t]+mod[ \t]+slide_movie_title\b", package_library_code):
        violations.append(
            "focused litchi-keynote movie-title owner module must remain private: "
            f"{KEYNOTE_MOVIE_TITLE_EXPORT_SOURCES[0]}"
        )

    canonical_names = frozenset(KEYNOTE_MOVIE_TITLE_CANONICAL_TYPES)
    owner_exports = _rust_canonical_exports(_mask_rust_non_code(owner), canonical_names)
    package_exports = _rust_canonical_exports(package, canonical_names)
    library_exports = _rust_canonical_exports(
        _keynote_family_view(
            root, root / KEYNOTE_MOVIE_TITLE_FAMILY_SOURCE, {}, KEYNOTE_MOVIE_TITLE_FAMILY_MODULE
        ),
        canonical_names,
    )
    for name in sorted(canonical_names):
        if name not in owner_exports or name not in library_exports:
            violations.append(
                "focused litchi-keynote movie-title public API is missing canonical type "
                f"{name}: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
            )

    selector_names = frozenset(KEYNOTE_MOVIE_TITLE_SELECTOR_TYPES)
    if "MovieSelector" not in _rust_canonical_exports(selector, selector_names):
        violations.append(
            "focused litchi-keynote movie-title selector source is missing canonical "
            f"MovieSelector: {KEYNOTE_MOVIE_TITLE_SELECTOR_SOURCE}"
        )
    if "MovieSelector" not in _rust_canonical_exports(library, selector_names):
        violations.append(
            "focused litchi-keynote movie-title public API is missing canonical "
            f"MovieSelector re-export: {KEYNOTE_MOVIE_TITLE_EXPORT_SOURCES[1]}"
        )

    owner_code = _mask_rust_non_code(owner)
    method_names = {
        match.group(1)
        for match in re.finditer(
            r"(?<![A-Za-z0-9_#])pub[ \t\r\n]+fn[ \t\r\n]+"
            r"([A-Za-z_][A-Za-z0-9_]*)\b",
            owner_code,
        )
    }
    for name in sorted(KEYNOTE_MOVIE_TITLE_PACKAGE_METHODS - method_names):
        violations.append(
            "focused litchi-keynote movie-title Package method is missing "
            f"{name}: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )

    edit_impl = re.search(
        r"(?<![A-Za-z0-9_#])impl(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]+"
        r"(?:'[^ ]+[ \t\r\n]+)?Edit\b",
        owner_code,
    )
    edit_body = ""
    if edit_impl is not None:
        opening = owner_code.find("{", edit_impl.end())
        depth = 1
        cursor = opening + 1 if opening >= 0 else len(owner_code)
        while opening >= 0 and cursor < len(owner_code) and depth:
            if owner_code[cursor] == "{":
                depth += 1
            elif owner_code[cursor] == "}":
                depth -= 1
            cursor += 1
        if depth == 0:
            edit_body = owner_code[opening + 1 : cursor - 1]
    edit_names = {
        match.group(1)
        for match in re.finditer(
            r"(?<![A-Za-z0-9_#])pub[ \t\r\n]+fn[ \t\r\n]+"
            r"([A-Za-z_][A-Za-z0-9_]*)\b",
            edit_body,
        )
    }
    for name in sorted(KEYNOTE_MOVIE_TITLE_EDIT_METHODS - edit_names):
        violations.append(
            "focused litchi-keynote movie-title Edit method is missing "
            f"{name}: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )

    facade_names = (
        canonical_names
        | selector_names
        | KEYNOTE_MOVIE_TITLE_PACKAGE_METHODS
        | KEYNOTE_MOVIE_TITLE_FLAT_ALIASES
        | {"Edit"}
    )
    for source, source_path in (
        (owner, owner_path),
        (selector, selector_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = source_path in {owner_path, selector_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_MOVIE_TITLE_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_MOVIE_TITLE_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_MOVIE_TITLE_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote movie-title public API exposes "
                        f"{reason} {identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote movie-title public API exposes raw byte slice "
                    f"{byte_slice}: {source_path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_MOVIE_TITLE_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote movie-title public API exposes raw identifier "
                    f"{match.group(0).strip()}: {source_path.relative_to(root)}:{line_number}"
                )
            for alias in sorted(identifiers & KEYNOTE_MOVIE_TITLE_FLAT_ALIASES):
                if alias in canonical_names:
                    continue
                violations.append(
                    "focused litchi-keynote movie-title public API retains flat alias "
                    f"{alias}: {source_path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_keynote_movie_title_codec_source_topology(root: Path = ROOT) -> list[str]:
    """Require a strict source-preserving codec for both movie-title edges."""

    owner_path = root / KEYNOTE_MOVIE_TITLE_OWNER_SOURCE
    if not owner_path.is_file():
        return []
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    codec_paths = (
        root / KEYNOTE_MOVIE_TITLE_CODEC_SOURCE,
        root / KEYNOTE_MOVIE_TITLE_GRAPH_CODEC_SOURCE,
    )
    violations: list[str] = []
    missing = [path for path in codec_paths if not path.is_file()]
    if missing:
        return [
            "focused litchi-keynote movie-title owner is missing strict dual-edge codec: "
            + ", ".join(str(path.relative_to(root)) for path in missing)
        ]
    shared_path = root / KEYNOTE_CHART_CAPTION_OWNER_SOURCE
    shared = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(shared_path.read_text(encoding="utf-8"))
        )
        if shared_path.is_file()
        else ""
    )
    codec_parts = [
        _mask_rust_non_code(_mask_rust_cfg_test_items(path.read_text(encoding="utf-8")))
        for path in codec_paths
    ]
    blob = "\n".join((owner, shared, *codec_parts))
    for function in sorted(KEYNOTE_MOVIE_TITLE_CODEC_FUNCTIONS):
        if not re.search(rf"(?<![A-Za-z0-9_]){re.escape(function)}\b", blob):
            violations.append(
                "focused litchi-keynote movie-title codec is missing strict operation "
                f"{function}: {KEYNOTE_MOVIE_TITLE_CODEC_SOURCE}"
            )
    if not any(
        re.search(rf"(?<![A-Za-z0-9_]){re.escape(name)}\b", blob)
        for name in KEYNOTE_MOVIE_TITLE_CODEC_TYPES
    ):
        violations.append(
            "focused litchi-keynote movie-title codec is missing a typed source-preserving "
            f"write/snapshot seam: {KEYNOTE_MOVIE_TITLE_CODEC_SOURCE}"
        )
    for edge_index, markers in enumerate(KEYNOTE_MOVIE_TITLE_DUAL_EDGE_MARKERS, start=1):
        if not any(
            re.search(rf"(?<![A-Za-z0-9_]){re.escape(marker)}(?![A-Za-z0-9_])", blob)
            for marker in markers
        ):
            violations.append(
                "focused litchi-keynote movie-title codec must validate dual-edge "
                f"graph edge {edge_index}: {KEYNOTE_MOVIE_TITLE_CODEC_SOURCE}"
            )
    if KEYNOTE_MOVIE_TITLE_CODEC_MODULE not in blob:
        violations.append(
            "focused litchi-keynote movie-title owner must route through hidden "
            f"{KEYNOTE_MOVIE_TITLE_CODEC_MODULE}: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if KEYNOTE_MOVIE_TITLE_GRAPH_CODEC_MODULE not in blob:
        violations.append(
            "focused litchi-keynote movie-title owner must route through hidden "
            f"{KEYNOTE_MOVIE_TITLE_GRAPH_CODEC_MODULE}: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if re.search(
        r"\bprost[ \t\r\n]*::[ \t\r\n]*Message\b|::decode\s*\(",
        "\n".join(codec_parts),
    ):
        violations.append(
            "focused litchi-keynote movie-title codec retains eager/generated decode: "
            f"{KEYNOTE_MOVIE_TITLE_CODEC_SOURCE}"
        )
    return sorted(set(violations))


def audit_keynote_movie_title_lifecycle_source_topology(root: Path = ROOT) -> list[str]:
    """Require title graph creation/removal, metadata, and exact transaction proof."""

    owner_path = root / KEYNOTE_MOVIE_TITLE_OWNER_SOURCE
    if not owner_path.is_file():
        return []
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    shared_path = root / KEYNOTE_CHART_CAPTION_OWNER_SOURCE
    shared = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(shared_path.read_text(encoding="utf-8"))
        )
        if shared_path.is_file()
        else ""
    )
    blob = f"{owner}\n{shared}"
    violations: list[str] = []
    budget_matches = list(KEYNOTE_MOVIE_CAPTION_PACKAGE_BUDGET_TYPE.finditer(blob))
    if not budget_matches:
        violations.append(
            "focused litchi-keynote movie-title owner is missing one private aggregate "
            f"CaptionBudget: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    operation_names = KEYNOTE_MOVIE_TITLE_GRAPH_OPERATION_NAMES | {
        name
        for name in re.findall(
            r"(?<![A-Za-z0-9_])fn[ \t\r\n]+(rewrite_[A-Za-z0-9_]+)[ \t\r\n]*\(", blob
        )
        if "movie" in name.lower() and "title" in name.lower()
    }
    operation_present = next(
        (name for name in operation_names if re.search(rf"\b{re.escape(name)}\b", owner)),
        None,
    )
    if operation_present is None and not re.search(
        r"rewrite_caption_graph_operation_with_budget", owner
    ):
        violations.append(
            "focused litchi-keynote movie-title graph operation helper is missing: "
            f"{KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if not re.search(r"\b(?:creating|creating_title|create)\b", blob):
        violations.append(
            "focused litchi-keynote movie-title lifecycle must expose creation path: "
            f"{KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if not re.search(r"\b(?:removing|removing_title|remove)\b", blob):
        violations.append(
            "focused litchi-keynote movie-title lifecycle must expose removal path: "
            f"{KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if not any(marker in blob for marker in KEYNOTE_MOVIE_TITLE_GRAPH_MARKERS):
        violations.append(
            "focused litchi-keynote movie-title lifecycle must retain title graph markers: "
            f"{KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if not any(marker in blob for marker in KEYNOTE_MOVIE_TITLE_METADATA_MARKERS):
        violations.append(
            "focused litchi-keynote movie-title lifecycle must update Metadata UUID/save-token "
            f"registry: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if not any(marker in blob for marker in KEYNOTE_MOVIE_TITLE_REFERENCE_TRANSITIONS):
        violations.append(
            "focused litchi-keynote movie-title lifecycle must use a source-authoritative "
            f"dual-edge reference transition: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if not any(marker in blob for marker in KEYNOTE_MOVIE_TITLE_INVERSE_LOCALITY_MARKERS):
        violations.append(
            "focused litchi-keynote movie-title lifecycle must verify exact inverse/locality "
            f"artifacts: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    if not KEYNOTE_CHART_CAPTION_CODEC_REPORT_CHARGE.search(blob):
        violations.append(
            "focused litchi-keynote movie-title lifecycle must charge strict codec reports "
            f"through CaptionBudget: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
        )
    for category, pattern in KEYNOTE_CHART_CAPTION_RESOURCE_CHARGES.items():
        if not pattern.search(blob):
            violations.append(
                "focused litchi-keynote movie-title lifecycle must charge prepublication "
                f"{category} work through CaptionBudget: {KEYNOTE_MOVIE_TITLE_OWNER_SOURCE}"
            )
    return sorted(set(violations))


def audit_keynote_movie_caption_lifecycle_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the complete private movie-caption graph transition.

    The facade audit proves that semantic types do not leak physical values;
    this companion ratchet proves that ``set``/``clear`` are not merely
    routed to the old host.  Creation and removal must share one bounded
    operation that updates the graph registry/Metadata and verifies the exact
    candidate.  Unsupported malformed or unproven *dependencies* may still
    fail closed, but a blanket absent-caption/clear branch is not a lifecycle
    owner.
    """

    owner_path = root / KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE
    if not owner_path.is_file():
        return [
            "focused litchi-keynote movie-caption owner source is missing: "
            f"{KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        ]
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    shared_path = root / KEYNOTE_CHART_CAPTION_OWNER_SOURCE
    shared = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(shared_path.read_text(encoding="utf-8"))
        )
        if shared_path.is_file()
        else ""
    )
    # Movie lifecycle orchestration deliberately reuses the private drawable-
    # generic caption graph/budget seam owned beside the chart transaction.
    # Audit both private package sources as one reachable implementation while
    # keeping the movie facade itself in the dedicated owner file.
    lifecycle_source = f"{shared}\n{owner}"
    declaration = re.compile(
        r"(?<![A-Za-z0-9_#])(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
        r"(?:unsafe[ \t\r\n]+|async[ \t\r\n]+|const[ \t\r\n]+)*"
        r"fn[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)\b"
    )
    functions: dict[str, tuple[str, str, int]] = {}
    for match in declaration.finditer(lifecycle_source):
        opening = lifecycle_source.find("{", match.end())
        if opening < 0:
            continue
        depth = 1
        cursor = opening + 1
        while cursor < len(lifecycle_source) and depth:
            if lifecycle_source[cursor] == "{":
                depth += 1
            elif lifecycle_source[cursor] == "}":
                depth -= 1
            cursor += 1
        if depth:
            continue
        functions[match.group(1)] = (
            lifecycle_source[match.start() : opening],
            lifecycle_source[opening + 1 : cursor - 1],
            match.start(),
        )

    violations: list[str] = []
    budget_declarations = list(
        KEYNOTE_MOVIE_CAPTION_PACKAGE_BUDGET_TYPE.finditer(lifecycle_source)
    )
    if not budget_declarations:
        violations.append(
            "focused litchi-keynote movie-caption owner is missing one private aggregate "
            f"package budget: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )
        budget_name: str | None = None
    else:
        if len(budget_declarations) != 1:
            violations.append(
                "focused litchi-keynote movie-caption owner must define exactly one private "
                f"aggregate package budget: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )
        budget_name = budget_declarations[0].group("name")

    def require_budget(name: str, label: str, *, signature_required: bool = True) -> str | None:
        function = functions.get(name)
        if function is None:
            violations.append(
                "focused litchi-keynote movie-caption lifecycle is missing "
                f"{label} function {name}: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )
            return None
        signature, body, _offset = function
        if signature_required and (
            budget_name is None
            or budget_name not in signature
            or not re.search(
                r"\b(?:budget|package_budget|aggregate_budget|transaction_budget)\b",
                signature,
            )
        ):
            violations.append(
                "focused litchi-keynote movie-caption lifecycle "
                f"{label} must thread the private aggregate package budget: "
                f"{KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )
        if not re.search(
            r"\b(?:budget|package_budget|aggregate_budget|transaction_budget)\b",
            body,
        ):
            violations.append(
                "focused litchi-keynote movie-caption lifecycle "
                f"{label} must use the aggregate package budget: "
                f"{KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )
        return body

    commit_body = require_budget("commit", "commit", signature_required=False)
    apply_body = require_budget(
        "apply_slide_movie_caption", "apply", signature_required=False
    )

    operation_name = None
    for name in functions:
        if name in KEYNOTE_MOVIE_CAPTION_GRAPH_OPERATIONS or re.search(
            r"rewrite_[A-Za-z0-9_]*movie[A-Za-z0-9_]*caption"
            r"(?:_graph|_lifecycle|_operation)?$",
            name,
        ):
            operation_name = name
            break
    if operation_name is None:
        violations.append(
            "focused litchi-keynote movie-caption graph operation helper is missing "
            f"{KEYNOTE_MOVIE_CAPTION_GRAPH_OPERATION}: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )
        operation_body = ""
        operation_offset = 0
    else:
        operation_body = functions[operation_name][1]
        operation_offset = functions[operation_name][2]
        for marker in ("creating", "removing"):
            if not re.search(rf"\b{re.escape(marker)}\b", operation_body):
                line_number = lifecycle_source.count("\n", 0, operation_offset) + 1
                violations.append(
                    "focused litchi-keynote movie-caption graph operation must expose "
                    f"the {marker} path: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}:{line_number}"
                )
        if not re.search(r"\b(?:creating|removing)\b", operation_body):
            # Keep the diagnostic above useful if an implementation chose an
            # enum/action spelling rather than booleans.
            violations.append(
                "focused litchi-keynote movie-caption graph operation must distinguish "
                f"create/remove transitions: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )

    lifecycle_blob = "\n".join(
        [lifecycle_source, operation_body, commit_body or "", apply_body or ""]
    )
    if not any(marker in lifecycle_blob for marker in KEYNOTE_MOVIE_CAPTION_GRAPH_MARKERS):
        violations.append(
            "focused litchi-keynote movie-caption lifecycle must retain a strict graph "
            f"transition marker: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )
    if not any(
        marker in lifecycle_blob for marker in KEYNOTE_MOVIE_CAPTION_METADATA_GRAPH_MARKERS
    ):
        violations.append(
            "focused litchi-keynote movie-caption lifecycle must update the Metadata/UUID "
            f"registry through a strict transition: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )
    if not any(marker in lifecycle_blob for marker in KEYNOTE_MOVIE_CAPTION_SAVE_TOKEN_REWRITE_MARKERS):
        violations.append(
            "focused litchi-keynote movie-caption lifecycle must advance Metadata save tokens: "
            f"{KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )
    if not any(marker in lifecycle_blob for marker in KEYNOTE_MOVIE_CAPTION_OBJECT_FRAMING_MARKERS):
        line_number = lifecycle_source.count("\n", 0, operation_offset) + 1
        violations.append(
            "focused litchi-keynote movie-caption lifecycle must validate canonical object "
            f"framing: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}:{line_number}"
        )
    if not any(marker in lifecycle_blob for marker in KEYNOTE_MOVIE_CAPTION_REFERENCE_TRANSITION_MARKERS):
        violations.append(
            "focused litchi-keynote movie-caption lifecycle must use a source-authoritative "
            f"caption reference transition: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )

    if any(marker in lifecycle_blob for marker in KEYNOTE_MOVIE_CAPTION_CROSS_COMPONENT_MARKERS):
        if not any(
            marker in lifecycle_blob
            for marker in KEYNOTE_MOVIE_CAPTION_EXTERNAL_REFERENCE_MARKERS
        ):
            violations.append(
                "focused litchi-keynote movie-caption cross-component graph dependency "
                f"must carry external-reference attribution: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )

    if not KEYNOTE_CHART_CAPTION_CODEC_REPORT_CHARGE.search(lifecycle_blob):
        violations.append(
            "focused litchi-keynote movie-caption lifecycle must charge nested codec reports "
            f"through the aggregate package budget: {KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )
    for category in ("archive", "snappy", "zip", "reopen", "artifact"):
        pattern = KEYNOTE_CHART_CAPTION_RESOURCE_CHARGES[category]
        if not pattern.search(lifecycle_blob):
            violations.append(
                "focused litchi-keynote movie-caption lifecycle must charge prepublication "
                f"{category} work through the aggregate package budget: "
                f"{KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
            )

    # A lifecycle owner may reject an unproven *dependency*, but it must not
    # retain the old blanket stand-in/clear fallback branch.
    if re.search(
        r"(?:storage_identifier|caption_info_identifier)\s*\.\s*is_none\(\)"
        r"[^;\n]{0,220}UnsupportedDependency|"
        r"after\s*\.\s*is_none\(\)[^;\n]{0,220}UnsupportedDependency",
        owner,
        re.DOTALL,
    ):
        violations.append(
            "focused litchi-keynote movie-caption lifecycle retains a blanket "
            "absent/clear UnsupportedDependency branch: "
            f"{KEYNOTE_MOVIE_CAPTION_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _keynote_movie_playback_owner_present(root: Path) -> bool:
    """Return whether the playback owner has crossed the activation seam."""

    owner_path = root / KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE
    codec_path = root / KEYNOTE_MOVIE_PLAYBACK_CODEC_SOURCE
    package_path = root / KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and codec_path.is_file()
        and re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_movie_playback\s*;", package_source)
        is not None
    )


def audit_keynote_movie_playback_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free movie-playback boundary.

    The audit is intentionally dormant until the private owner, neutral codec,
    and package wiring exist together.  This avoids turning the still-live
    Draft compatibility surface into a false migration failure while
    the owner is being assembled.
    """

    if not _keynote_movie_playback_owner_present(root):
        return []

    owner_path = root / KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE
    semantic_path = root / KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_SOURCE
    package_path = root / KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[1]
    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    semantic = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    violations: list[str] = []
    owner_code = _mask_rust_non_code(owner)
    semantic_code = _mask_rust_non_code(semantic)
    package_code = _mask_rust_non_code(package)
    library_code = _mask_rust_non_code(library)
    package_library_code = _mask_rust_non_code(package + library)

    if re.search(r"(?m)^pub[ \t]+mod[ \t]+slide_movie_playback\b", package_library_code):
        violations.append(
            "focused litchi-keynote movie-playback owner module must remain private: "
            f"{KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[0]}"
        )
    if re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_movie_playback\s*;", package_code) is None:
        violations.append(
            "focused litchi-keynote movie-playback owner module is missing: "
            f"{KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[0]}"
        )

    family_path = root / KEYNOTE_MOVIE_PLAYBACK_FAMILY_SOURCE
    family = _keynote_family_view(root, family_path, {}, KEYNOTE_MOVIE_PLAYBACK_FAMILY_MODULE)
    for name in sorted(KEYNOTE_MOVIE_PLAYBACK_CANONICAL_TYPES):
        for source, path in ((owner, owner_path), (family, family_path)):
            if name not in _rust_canonical_exports(source, KEYNOTE_MOVIE_PLAYBACK_CANONICAL_TYPES):
                violations.append(
                    "focused litchi-keynote movie-playback public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )
    for name in sorted(KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_TYPES):
        if name not in _rust_canonical_exports(semantic, KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_TYPES):
            violations.append(
                "focused litchi-keynote movie-playback semantic API is missing "
                f"{name}: {KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_SOURCE}"
            )
    if not family.strip():
        violations.append(
            "focused litchi-keynote movie-playback semantic API is missing public transaction module: "
            f"{KEYNOTE_MOVIE_PLAYBACK_FAMILY_SOURCE}"
        )
    for name in sorted(KEYNOTE_MOVIE_PLAYBACK_SELECTOR_TYPES):
        if name not in _rust_canonical_exports(semantic + library, KEYNOTE_MOVIE_PLAYBACK_SELECTOR_TYPES):
            violations.append(
                "focused litchi-keynote movie-playback public API is missing selector "
                f"{name}: {KEYNOTE_MOVIE_PLAYBACK_EXPORT_SOURCES[1]}"
            )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Package")
    }
    for name in sorted(KEYNOTE_MOVIE_PLAYBACK_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote movie-playback Package method is missing "
                f"{name}: {KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_movie_playback_settings":
            for selector in ("SlideSelector", "MovieSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote movie-playback Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE}"
                    )

    edit_impl = re.search(
        r"(?<![A-Za-z0-9_#])impl(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]+"
        r"(?:'[^ ]+[ \t\r\n]+)?Edit\b",
        owner_code,
    )
    edit_body = owner_code if edit_impl is None else owner_code[edit_impl.end() :]
    for name in sorted(KEYNOTE_MOVIE_PLAYBACK_EDIT_METHODS):
        if not re.search(rf"\bpub[ \t]+fn[ \t]+{re.escape(name)}\b", edit_body):
            violations.append(
                "focused litchi-keynote movie-playback edit is missing "
                f"{name}: {KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE}"
            )

    facade_names = (
        KEYNOTE_MOVIE_PLAYBACK_CANONICAL_TYPES
        | KEYNOTE_MOVIE_PLAYBACK_SEMANTIC_TYPES
        | KEYNOTE_MOVIE_PLAYBACK_SELECTOR_TYPES
        | KEYNOTE_MOVIE_PLAYBACK_PACKAGE_METHODS
        | KEYNOTE_MOVIE_PLAYBACK_FLAT_ALIASES
        | {"Edit"}
    )
    for source, source_path in (
        (owner, owner_path),
        (semantic, semantic_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = source_path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)}
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_MOVIE_PLAYBACK_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_MOVIE_PLAYBACK_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_MOVIE_PLAYBACK_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote movie-playback public API exposes "
                        f"{reason} {identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_MOVIE_PLAYBACK_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote movie-playback public API retains flat alias "
                        f"{identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote movie-playback public API exposes raw byte slice "
                    f"{byte_slice}: {source_path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_MOVIE_PLAYBACK_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote movie-playback public API exposes raw identifier "
                    f"{match.group(0).strip()}: {source_path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote movie-playback public API retains a glob re-export: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )

    codec_path = root / KEYNOTE_MOVIE_PLAYBACK_CODEC_SOURCE
    codec = _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
    codec_code = _mask_rust_non_code(codec)
    for function in sorted(KEYNOTE_MOVIE_PLAYBACK_CODEC_FUNCTIONS):
        if re.search(rf"\b(?:pub[ \t]+)?fn[ \t]+{re.escape(function)}\b", codec_code) is None:
            violations.append(
                "focused litchi-keynote movie-playback hidden codec is missing strict API "
                f"{function}: {KEYNOTE_MOVIE_PLAYBACK_CODEC_SOURCE}"
            )
    for type_name in sorted(KEYNOTE_MOVIE_PLAYBACK_CODEC_TYPES):
        if re.search(
            rf"\b(?:pub[ \t]+)?(?:struct|enum|type)[ \t]+{re.escape(type_name)}\b",
            codec_code,
        ) is None:
            violations.append(
                "focused litchi-keynote movie-playback hidden codec is missing strict type "
                f"{type_name}: {KEYNOTE_MOVIE_PLAYBACK_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_MOVIE_PLAYBACK_CODEC_MARKER_GROUPS.items():
        if not any(marker in codec for marker in markers):
            violations.append(
                "focused litchi-keynote movie-playback hidden codec is missing "
                f"{label} marker: {KEYNOTE_MOVIE_PLAYBACK_CODEC_SOURCE}"
            )
    codec_lib = (
        _mask_rust_cfg_test_items(
            (root / KEYNOTE_MOVIE_PLAYBACK_CODEC_PUBLIC_SOURCE).read_text(encoding="utf-8")
        )
        if (root / KEYNOTE_MOVIE_PLAYBACK_CODEC_PUBLIC_SOURCE).is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(KEYNOTE_MOVIE_PLAYBACK_CODEC_MODULE)}\b",
        codec_lib,
    ) is None:
        violations.append(
            "focused litchi-keynote movie-playback public API is missing hidden codec module "
            f"{KEYNOTE_MOVIE_PLAYBACK_CODEC_MODULE}: {KEYNOTE_MOVIE_PLAYBACK_CODEC_PUBLIC_SOURCE}"
        )

    for fuzz_path in KEYNOTE_MOVIE_PLAYBACK_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote movie-playback boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote movie-playback fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in KEYNOTE_MOVIE_PLAYBACK_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-keynote movie-playback boundary is missing fuzz corpus: "
                f"{corpus}"
            )
    return sorted(set(violations))


def audit_keynote_movie_playback_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require aggregate resource/prepublication markers for playback edits."""

    if not _keynote_movie_playback_owner_present(root):
        return []
    owner_path = root / KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    violations: list[str] = []
    for label, markers in KEYNOTE_MOVIE_PLAYBACK_PACKAGE_MARKER_GROUPS.items():
        if not any(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote movie-playback owner is missing "
                f"{label} marker: {KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE}"
            )
    if "movie_playback_codec" not in owner:
        violations.append(
            "focused litchi-keynote movie-playback owner must route through hidden movie_playback_codec: "
            f"{KEYNOTE_MOVIE_PLAYBACK_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _keynote_slide_media_data_owner_present(root: Path) -> bool:
    """Return whether the slide-media replacement owner is wired and active."""

    owner_path = root / KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[0]
    if not owner_path.is_file() or not package_path.is_file():
        return False
    package = _mask_rust_non_code(
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
    )
    return KEYNOTE_SLIDE_MEDIA_DATA_PACKAGE_MODULE.search(package) is not None


def audit_keynote_slide_media_data_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free slide-media data facade.

    This audit stays dormant until the focused module is wired into
    ``Package``.  It then checks the public semantic transaction surface and
    scans its declarations for native identifiers, archive objects, generated
    messages, and wire values.  Borrowed media bytes are intentional semantic
    payloads for this operation and are therefore not treated as physical
    archive leakage.
    """

    if not _keynote_slide_media_data_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_SOURCE
    package_path = root / KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[1]
    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    semantic = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package = _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
    library = _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
    owner_code = _mask_rust_non_code(owner)
    semantic_code = _mask_rust_non_code(semantic)
    package_code = _mask_rust_non_code(package)
    library_code = _mask_rust_non_code(library)
    package_library_code = _mask_rust_non_code(package + library)
    violations: list[str] = []

    if KEYNOTE_SLIDE_MEDIA_DATA_PUBLIC_MODULE.search(package_library_code):
        violations.append(
            "focused litchi-keynote slide-media data owner module must remain private: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[0]}"
        )
    if KEYNOTE_SLIDE_MEDIA_DATA_PACKAGE_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-keynote slide-media data owner module is missing: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[0]}"
        )

    family_path = root / KEYNOTE_SLIDE_MEDIA_DATA_FAMILY_SOURCE
    family = _keynote_family_view(root, family_path, {}, KEYNOTE_SLIDE_MEDIA_DATA_FAMILY_MODULE)
    for name in sorted(KEYNOTE_SLIDE_MEDIA_DATA_CANONICAL_TYPES):
        for source, path in (
            (owner, owner_path),
            (family, family_path),
        ):
            if name not in _rust_canonical_exports(
                source, KEYNOTE_SLIDE_MEDIA_DATA_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-media data public API is missing "
                    f"canonical type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        owner + semantic + family,
        KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_TYPES - semantic_exports):
        violations.append(
            "focused litchi-keynote slide-media data semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        semantic + package + library,
        KEYNOTE_SLIDE_MEDIA_DATA_SELECTOR_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_MEDIA_DATA_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote slide-media data public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES[1]}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Package")
    }
    for name in sorted(KEYNOTE_SLIDE_MEDIA_DATA_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-media data Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_media_data":
            for selector in ("SlideSelector", "MovieSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-media data Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
                    )
        if name == "slide_media_data" and "MediaPart" not in declaration:
            violations.append(
                "focused litchi-keynote slide-media data read must select a typed "
                f"MediaPart: {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )

    edit_impl = re.search(
        r"(?<![A-Za-z0-9_#])impl(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]+"
        r"(?:'[^ ]+[ \t\r\n]+)?Edit\b",
        owner_code,
    )
    edit_body = owner_code if edit_impl is None else owner_code[edit_impl.end() :]
    for name in sorted(KEYNOTE_SLIDE_MEDIA_DATA_EDIT_METHODS):
        if not re.search(
            rf"\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+{re.escape(name)}\b",
            edit_body,
        ):
            violations.append(
                "focused litchi-keynote slide-media data edit is missing "
                f"{name}: {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )

    facade_names = (
        KEYNOTE_SLIDE_MEDIA_DATA_CANONICAL_TYPES
        | KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_MEDIA_DATA_SELECTOR_TYPES
        | KEYNOTE_SLIDE_MEDIA_DATA_PACKAGE_METHODS
        | KEYNOTE_SLIDE_MEDIA_DATA_FLAT_ALIASES
        | {"Edit", "MediaPart"}
    )
    for source, source_path in (
        (owner, owner_path),
        (semantic, semantic_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = source_path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_SLIDE_MEDIA_DATA_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_SLIDE_MEDIA_DATA_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_SLIDE_MEDIA_DATA_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-media data public API exposes "
                        f"{reason} {identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_MEDIA_DATA_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-media data public API retains flat alias "
                        f"{identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-media data public API retains a glob re-export: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_MEDIA_DATA_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote slide-media data public API exposes raw identifier "
                    f"{match.group(0).strip()}: {source_path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_keynote_slide_media_data_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the neutral lazy Buffa media and metadata codec seams."""

    if not _keynote_slide_media_data_owner_present(root):
        return []

    violations: list[str] = []
    for codec_name, required in KEYNOTE_SLIDE_MEDIA_DATA_CODEC_APIS.items():
        codec_path = root / Path("crates/litchi-iwa-protos/src") / f"{codec_name}.rs"
        if not codec_path.is_file():
            violations.append(
                "focused litchi-keynote slide-media data codec is missing: "
                f"{codec_path.relative_to(root)}"
            )
            continue
        source = _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        code = _mask_rust_non_code(source)
        for name in required:
            if re.search(
                rf"\b(?:pub[ \t]+)?(?:fn|struct|enum|type|trait)[ \t]+"
                rf"{re.escape(name)}\b",
                code,
            ) is None:
                violations.append(
                    "focused litchi-keynote slide-media data codec is missing "
                    f"{codec_name} API {name}: {codec_path.relative_to(root)}"
                )

    for label, markers in KEYNOTE_SLIDE_MEDIA_DATA_CODEC_MARKERS.items():
        if label.startswith("keynote"):
            codec_path = root / KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES[0]
        elif label.startswith("metadata"):
            codec_path = root / KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES[1]
        else:
            codec_path = root / KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES[1]
        if not codec_path.is_file():
            continue
        source = _mask_rust_non_code(
            _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        )
        if not any(marker in source for marker in markers):
            violations.append(
                "focused litchi-keynote slide-media data codec is missing "
                f"{label} marker: {codec_path.relative_to(root)}"
            )

    public_source_path = root / KEYNOTE_SLIDE_MEDIA_DATA_CODEC_PUBLIC_SOURCE
    if public_source_path.is_file():
        public_source = _mask_rust_cfg_test_items(
            public_source_path.read_text(encoding="utf-8")
        )
        for codec_name in KEYNOTE_SLIDE_MEDIA_DATA_CODEC_APIS:
            if re.search(
                rf"(?m)^\s*pub\s+mod\s+{re.escape(codec_name)}\s*;",
                public_source,
            ) is None:
                violations.append(
                    "focused litchi-keynote slide-media data codec is not exposed through "
                    f"the neutral proto crate: {codec_name}: "
                    f"{KEYNOTE_SLIDE_MEDIA_DATA_CODEC_PUBLIC_SOURCE}"
                )

    generated_decode = re.compile(
        r"\b(?:[A-Za-z_][A-Za-z0-9_]*::)?"
        r"(?:DataReference|PackageMetadata|DataInfo|ComponentInfo)"
        r"\s*::\s*decode\s*\(|\b(?:prost|buffa)\s*::\s*Message\b"
    )
    for codec_path in KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES:
        absolute = root / codec_path
        if not absolute.is_file():
            continue
        source = _mask_rust_non_code(
            _mask_rust_cfg_test_items(absolute.read_text(encoding="utf-8"))
        )
        for match in generated_decode.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote slide-media data codec eagerly decodes a generated "
                f"message {match.group(0).strip()}: {codec_path}:{line_number}"
            )
    return sorted(set(violations))


def audit_keynote_slide_media_data_core_boundary_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep identity-remap clone implementation in the archive core.

    The compatibility host may route through the neutral core helper while
    the migration is in progress.  This ratchet therefore checks ownership of
    definitions, the core helper's dependency direction, and public facade
    exposure without forbidding a legitimate compatibility call site.
    """

    if not _keynote_slide_media_data_owner_present(root):
        return []

    violations: list[str] = []

    def masked_file(path: Path) -> str:
        if not path.is_file():
            return ""
        return _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )

    core_root = root / KEYNOTE_SLIDE_MEDIA_DATA_CORE_SOURCE_ROOT
    core_files = (
        sorted(core_root.rglob("*.rs")) if core_root.is_dir() else []
    )

    definition_pattern = {
        method: re.compile(
            rf"\b(?:pub(?:\s*\([^)]*\))?\s+)?fn\s+"
            rf"(?:r#)?{re.escape(method)}\b"
        )
        for method in KEYNOTE_SLIDE_MEDIA_DATA_CORE_CLONE_METHODS
    }
    definitions: dict[str, list[tuple[Path, str]]] = {
        method: [] for method in KEYNOTE_SLIDE_MEDIA_DATA_CORE_CLONE_METHODS
    }
    for path in core_files:
        code = masked_file(path)
        for method, pattern in definition_pattern.items():
            if pattern.search(code) is not None:
                definitions[method].append((path, code))

    # The audit is dormant for fixtures that predate the clone migration.  A
    # physical helper, once introduced, must not leave a second owner.
    if not any(definitions.values()):
        return []
    for method, entries in definitions.items():
        if not entries:
            violations.append(
                "focused litchi-keynote slide-media clone boundary is missing "
                f"core helper {method}: {KEYNOTE_SLIDE_MEDIA_DATA_CORE_SOURCE_ROOT}"
            )
            continue
        for path, code in entries:
            body = _rust_any_function_body(code, method)
            if body is None:
                continue
            forbidden = KEYNOTE_SLIDE_MEDIA_DATA_CORE_CLONE_FORBIDDEN_IMPORTS.search(
                body
            )
            if forbidden is not None:
                violations.append(
                    "focused litchi-keynote slide-media clone boundary core helper "
                    "must not depend on generated/protobuf types: "
                    f"{forbidden.group(0).strip()}: {path.relative_to(root)}"
                )

    # Definitions belong to the archive core; inspect declarations only.
    for source_root in (root / Path("crates/litchi-iwa-protos/src"),):
        if not source_root.is_dir():
            continue
        for path in sorted(source_root.rglob("*.rs")):
            code = masked_file(path)
            for method, pattern in definition_pattern.items():
                if pattern.search(code) is not None:
                    line_number = code.count("\n", 0, pattern.search(code).start()) + 1
                    violations.append(
                        "focused litchi-keynote slide-media clone helper is "
                        "implemented outside the archive core: "
                        f"{method}: {path.relative_to(root)}:{line_number}"
                    )

    # A public semantic declaration must not publish the physical helper or
    # any of its raw archive vocabulary. The dedicated facade audit performs
    # the broader type classification; this exact check catches a helper
    # re-export even when its source name is aliased.
    public_paths = (
        KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE,
        KEYNOTE_SLIDE_MEDIA_DATA_SEMANTIC_SOURCE,
        *KEYNOTE_SLIDE_MEDIA_DATA_EXPORT_SOURCES,
    )
    clone_name = re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?clone_with_identity_remap(?![A-Za-z0-9_])"
    )
    for relative in public_paths:
        path = root / relative
        source = masked_file(path)
        if not source:
            continue
        for declaration, line_number in _rust_public_declarations(source):
            if clone_name.search(declaration) is not None:
                violations.append(
                    "focused litchi-keynote slide-media public API exposes the "
                    f"physical clone helper: {relative}:{line_number}"
                )

    return sorted(set(violations))


def audit_keynote_slide_media_data_metadata_boundary_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require a bounded map witness for atomic owner/DataInfo removal.

    The neutral codec may expose a borrowed DataMetadataMapSource to a private
    format adapter, but it must validate the external payload with the same
    finite limits as PackageMetadata and carry owner/DataInfo removals through
    the prepared transaction. This blocks a direct map materialization or a
    one-shot removal path from silently bypassing the final-owner proof.
    """

    if not _keynote_slide_media_data_owner_present(root):
        return []
    codec_path = root / KEYNOTE_SLIDE_MEDIA_DATA_CODEC_SOURCES[1]
    if not codec_path.is_file():
        return [
            "focused litchi-keynote slide-media metadata boundary is missing "
            f"the neutral codec: {codec_path.relative_to(root)}"
        ]

    source = _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
    code = _mask_rust_non_code(source)
    violations: list[str] = []

    # Keep this focused ratchet dormant for older metadata codecs.  Once the
    # DataInfo removal seam exists, the neutral borrowed witness is required.
    if re.search(r"\b(?:pub\s+)?struct\s+DataInfoRemoval\b", code) is None:
        return []

    for name in KEYNOTE_SLIDE_MEDIA_DATA_METADATA_MAP_APIS:
        if re.search(
            rf"\b(?:pub[ \t]+)?(?:struct|enum|trait|type)[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b",
            code,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-media metadata boundary is missing "
                f"map witness API {name}: {codec_path.relative_to(root)}"
            )

    map_struct = _rust_named_struct_body(code, "DataMetadataMapSource")
    if map_struct is None:
        violations.append(
            "focused litchi-keynote slide-media metadata map witness must remain a "
            f"named borrowed source type: {codec_path.relative_to(root)}"
        )
    else:
        map_body, _offset = map_struct
        for marker in ("payload", "fields", "work_bytes", "entries", "scratch_bytes"):
            if not re.search(rf"\b{re.escape(marker)}\b", map_body):
                violations.append(
                    "focused litchi-keynote slide-media metadata map witness omits "
                    f"{marker}: {codec_path.relative_to(root)}"
                )
        if re.search(
            r"\bpayload\s*:\s*&\s*(?:'[A-Za-z_][A-Za-z0-9_]*\s*)?\[\s*u8\s*\]",
            map_body,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-media metadata map witness must borrow "
                f"its payload: {codec_path.relative_to(root)}"
            )

    batch_struct = _rust_named_struct_body(code, "MediaRewriteBatch")
    if batch_struct is None:
        violations.append(
            "focused litchi-keynote slide-media metadata boundary is missing the "
            f"atomic rewrite batch: {codec_path.relative_to(root)}"
        )
    else:
        batch_body, _offset = batch_struct
        for marker in ("data_removals", "owner_removals", "data_metadata_map_source"):
            if not re.search(rf"\b{re.escape(marker)}\b", batch_body):
                violations.append(
                    "focused litchi-keynote slide-media metadata batch must carry "
                    f"{marker}: {codec_path.relative_to(root)}"
                )

    validation = _rust_named_function_body(code, "validate_batch")
    if validation is None:
        violations.append(
            "focused litchi-keynote slide-media metadata boundary is missing "
            f"batch validation: {codec_path.relative_to(root)}"
        )
    else:
        validation_body, _offset = validation
        removal_loop = re.search(
            r"\bfor\b[^\n{]*\bbatch\.data_removals\b", validation_body
        )
        removal_region = ""
        if removal_loop is not None:
            removal_start = removal_loop.start()
            removal_end = re.search(
                r"\n\s*for\s+[^\n]*\bbatch\.(?:data_replacements|"
                r"owner_(?:additions|removals|updates))\b",
                validation_body[removal_start:],
            )
            end = (
                removal_start + removal_end.start()
                if removal_end is not None
                else len(validation_body)
            )
            removal_region = validation_body[removal_start:end]
        if re.search(
            r"\b(?:count_data_owners|owner_removals|owner_matches|"
            r"remaining|final|after|data_references_removed_after_owner_batch)\w*\b",
            removal_region,
            re.IGNORECASE,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-media metadata DataInfo removal must "
                "prove final-owner state or conservatively refuse: "
                f"{codec_path.relative_to(root)}"
            )
        if "data_metadata_map_source" not in removal_region.lower():
            violations.append(
                "focused litchi-keynote slide-media metadata DataInfo removal must "
                "honor the bounded map witness dependency: "
                f"{codec_path.relative_to(root)}"
            )

    # A public map-shaped collection is a second, unbounded ownership route.
    # The three borrowed witness names above are the only permitted public map
    # vocabulary inside this neutral module.
    for match in KEYNOTE_SLIDE_MEDIA_DATA_METADATA_PUBLIC_COLLECTION.finditer(code):
        line_number = code.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-keynote slide-media metadata boundary exposes an "
            f"unbounded/public map collection: {codec_path.relative_to(root)}:{line_number}"
        )

    return sorted(set(violations))


def audit_keynote_slide_media_data_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require bounded metadata closure and exact publication markers."""

    if not _keynote_slide_media_data_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    violations: list[str] = []
    for label, markers in KEYNOTE_SLIDE_MEDIA_DATA_OWNER_MARKER_GROUPS.items():
        if not any(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote slide-media data owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )
    return sorted(set(violations))


def audit_keynote_slide_media_data_transaction_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one bounded ledger through the slide-media transaction graph.

    The facade/resource audits prove vocabulary and marker coverage. This
    companion ratchet protects the implementation seam that makes those
    claims true: one private MediaBudget is created at each public operation
    root, mutable budget state is threaded through physical and closure
    helpers, metadata uses the prepared rewrite contract, and the dedicated
    closure child is actually reachable. The checks intentionally inspect
    function shape and call flow rather than requiring incidental helper names
    beyond the two explicit child seams.
    """

    if not _keynote_slide_media_data_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE
    budget_path = root / KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE
    closure_path = root / KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE
    owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner_code = _mask_rust_non_code(owner_source)
    violations: list[str] = []

    def masked_file(path: Path) -> str:
        if not path.is_file():
            return ""
        return _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )

    budget_code = masked_file(budget_path)
    closure_code = masked_file(closure_path)

    if not budget_path.is_file():
        violations.append(
            "focused litchi-keynote slide-media transaction is missing private budget child: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE}"
        )
    if not closure_path.is_file():
        violations.append(
            "focused litchi-keynote slide-media transaction is missing private closure child: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE}"
        )
    if KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_MODULE.search(owner_code) is None:
        violations.append(
            "focused litchi-keynote slide-media owner must wire private budget child: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
        )
    if KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_MODULE.search(owner_code) is None:
        violations.append(
            "focused litchi-keynote slide-media owner must wire private closure child: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
        )
    if KEYNOTE_SLIDE_MEDIA_DATA_PUBLIC_CHILD_MODULE.search(owner_code):
        violations.append(
            "focused litchi-keynote slide-media budget/closure children must remain private: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
        )

    if budget_code:
        if KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_TYPE.search(budget_code) is None:
            violations.append(
                "focused litchi-keynote slide-media budget child is missing private MediaBudget: "
                f"{KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE}"
            )
        if KEYNOTE_SLIDE_MEDIA_DATA_PUBLIC_BUDGET_TYPE.search(budget_code):
            violations.append(
                "focused litchi-keynote slide-media MediaBudget must remain private: "
                f"{KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE}"
            )
        if KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_FACTORY.search(budget_code) is None:
            violations.append(
                "focused litchi-keynote slide-media budget child is missing the "
                "root for_package factory: "
                f"{KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE}"
            )
        for variant in sorted(KEYNOTE_SLIDE_MEDIA_DATA_LIMIT_VARIANTS):
            if not re.search(
                rf"\bLimitKind\s*::\s*{re.escape(variant)}\b",
                budget_code,
            ):
                violations.append(
                    "focused litchi-keynote slide-media budget child leaves limit "
                    f"variant unused {variant}: {KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_SOURCE}"
                )

    # Small source-shape parser used only for this owner graph. It preserves
    # signatures and bodies while avoiding assumptions about helper names.
    function_declaration = re.compile(
        r"(?<![A-Za-z0-9_#])(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
        r"(?:unsafe[ \t\r\n]+|async[ \t\r\n]+|const[ \t\r\n]+)*"
        r"fn[ \t\r\n]+(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
    )

    def function_records(code: str) -> list[tuple[str, str, str]]:
        records: list[tuple[str, str, str]] = []
        for match in function_declaration.finditer(code):
            opening = code.find("{", match.end())
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
                    match.group("name"),
                    code[match.start() : opening],
                    code[opening + 1 : cursor - 1],
                )
            )
        return records

    def borrowed_budget_receiver_methods(code: str) -> set[str]:
        """Find receiver methods that retain the operation's mutable ledger.

        A Buffa visitor callback cannot add a budget argument to its trait
        method. Permit that shape only when the named receiver struct stores
        exactly one ``&mut MediaBudget`` field and a function receiving
        ``&mut MediaBudget`` constructs it with that same borrowed variable.
        This keeps ``&mut self`` from becoming a general budget exemption.
        """

        struct_names = {
            match.group(1)
            for match in re.finditer(
                r"\bstruct[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)\b", code
            )
        }
        if not struct_names:
            return set()

        borrowed_field = re.compile(
            r"\bbudget\s*:\s*&\s*(?:'[A-Za-z_][A-Za-z0-9_]*\s*)?"
            r"mut\s+MediaBudget\b"
        )
        literal_start = {
            name: re.compile(
                rf"\b{re.escape(name)}(?:\s*<[^{{}}]*>)?\s*\{{"
            )
            for name in struct_names
        }

        def balanced_body(source: str, opening: int) -> str | None:
            depth = 1
            cursor = opening + 1
            while cursor < len(source) and depth:
                if source[cursor] == "{":
                    depth += 1
                elif source[cursor] == "}":
                    depth -= 1
                cursor += 1
            if depth:
                return None
            return source[opening + 1 : cursor - 1]

        borrowed_structs: set[str] = set()
        for name in struct_names:
            struct = _rust_named_struct_body(code, name)
            if struct is None:
                continue
            struct_body, _offset = struct
            if len(borrowed_field.findall(struct_body)) != 1:
                continue

            # Require the literal inside a function with an explicit mutable
            # MediaBudget parameter. This rejects owned or copied visitors.
            constructed_from_ledger = False
            for _function_name, signature, function_body in function_records(code):
                if KEYNOTE_SLIDE_MEDIA_DATA_MUTABLE_BUDGET_PARAMETER.search(
                    signature
                ) is None:
                    continue
                for literal in literal_start[name].finditer(function_body):
                    opening = function_body.find("{", literal.start(), literal.end())
                    if opening < 0:
                        continue
                    body = balanced_body(function_body, opening)
                    if body is None:
                        continue
                    if re.search(
                        r"(?:\bbudget\s*:\s*budget\b(?=\s*(?:,|$))|"
                        r"(?<![A-Za-z0-9_*])budget\s*(?:,|$))",
                        body,
                    ):
                        constructed_from_ledger = True
                        break
                if constructed_from_ledger:
                    break
            if constructed_from_ledger:
                borrowed_structs.add(name)

        if not borrowed_structs:
            return set()

        allowed: set[str] = set()
        for impl in re.finditer(r"\bimpl\b[^{}]*\{", code):
            opening = code.find("{", impl.start(), impl.end())
            if opening < 0:
                continue
            body = balanced_body(code, opening)
            if body is None:
                continue
            header = code[impl.start() : opening]
            if not any(
                re.search(rf"\b{re.escape(name)}\b", header)
                for name in borrowed_structs
            ):
                continue
            for method in function_declaration.finditer(body):
                method_opening = body.find("{", method.end())
                if method_opening < 0:
                    continue
                signature = body[method.start() : method_opening]
                if re.search(r"&[ \t]*mut[ \t]+self\b", signature):
                    allowed.add(method.group("name"))
        return allowed

    owner_functions = function_records(owner_code)
    closure_functions = function_records(closure_code)
    owner_by_name: dict[str, tuple[str, str]] = {
        name: (signature, body) for name, signature, body in owner_functions
    }

    function_call = re.compile(
        r"(?<![A-Za-z0-9_])(?P<qualifier>(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*)"
        r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)[ \t\r\n]*\("
    )

    def called_functions(body: str) -> set[str]:
        called: set[str] = set()
        for match in function_call.finditer(body):
            name = match.group("name")
            if name not in owner_by_name:
                continue
            qualifier = match.group("qualifier").replace(" ", "").replace("\t", "")
            if qualifier:
                qualifier_type = qualifier.rstrip(":").split("::")[-1]
                if qualifier_type not in {"Self", "Edit"}:
                    continue
            called.add(name)
        return called

    def reachable_functions(seed: str) -> dict[str, tuple[str, str]]:
        reachable: dict[str, tuple[str, str]] = {}
        pending = [seed]
        while pending:
            name = pending.pop()
            if name in reachable or name not in owner_by_name:
                continue
            signature, body = owner_by_name[name]
            reachable[name] = (signature, body)
            pending.extend(called_functions(body) - reachable.keys())
        return reachable

    # Each public operation root owns exactly one ledger. A public method may
    # delegate its root construction to one immediate private operation helper
    # (for example edit -> Edit::new); the reachable call graph
    # must still contain exactly one factory call. set is also a public staging
    # root because it bounds a caller-provided allocation.
    budget_roots = set(KEYNOTE_SLIDE_MEDIA_DATA_PACKAGE_METHODS) | {"commit", "set"}
    factory_owner_names: set[str] = set()
    for name in sorted(budget_roots):
        record = owner_by_name.get(name)
        if record is None:
            continue
        _signature, body = record
        reachable = reachable_functions(name)
        factory_owners = {
            helper_name
            for helper_name, (_helper_signature, helper_body) in reachable.items()
            if KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_ROOT_CALL.search(helper_body)
        }
        factory_count = sum(
            len(KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_ROOT_CALL.findall(helper_body))
            for _helper_name, (_helper_signature, helper_body) in reachable.items()
        )
        if factory_count != 1:
            violations.append(
                "focused litchi-keynote slide-media transaction root must create exactly one "
                f"MediaBudget::for_package ledger ({name}): {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )
        factory_owner_names.update(factory_owners)
        ledger_owner = next(iter(factory_owners), None)
        if ledger_owner is not None and name != "set":
            _ledger_signature, ledger_body = reachable[ledger_owner]
            if not re.search(r"&[ \t]*mut[ \t]+budget\b", ledger_body):
                violations.append(
                    "focused litchi-keynote slide-media transaction root must thread its mutable "
                    f"budget ledger ({name} via {ledger_owner}): {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
                )
        elif ledger_owner is None and name != "set":
            violations.append(
                "focused litchi-keynote slide-media transaction root must thread its mutable "
                f"budget ledger ({name}): {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )

    # A helper that touches the ledger must receive the same mutable instance;
    # a second factory call inside the graph would reset aggregate limits.
    receiver_budget_methods = {
        KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE: borrowed_budget_receiver_methods(
            owner_code
        ),
        KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE: borrowed_budget_receiver_methods(
            closure_code
        ),
    }
    for source_label, records in (
        (KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE, owner_functions),
        (KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE, closure_functions),
    ):
        for name, signature, body in records:
            if name in budget_roots:
                continue
            if name in factory_owner_names:
                continue
            uses_budget = bool(
                re.search(r"\bMediaBudget\b|\bbudget\b", signature + "\n" + body)
            )
            if not uses_budget:
                continue
            if KEYNOTE_SLIDE_MEDIA_DATA_BUDGET_ROOT_CALL.search(body):
                violations.append(
                    "focused litchi-keynote slide-media helper must not reset the "
                    f"root MediaBudget ({name}): {source_label}"
                )
            if name in receiver_budget_methods[source_label]:
                continue
            if KEYNOTE_SLIDE_MEDIA_DATA_MUTABLE_BUDGET_PARAMETER.search(signature) is None:
                violations.append(
                    "focused litchi-keynote slide-media helper must thread &mut MediaBudget "
                    f"({name}): {source_label}"
                )

    # A prepared metadata rewrite is required to charge and execute the exact
    # planned allocation. The old one-shot content-replacement helper would
    # silently recreate codec limits and is therefore a tombstoned edge.
    prepared_bodies = [
        body
        for _name, _signature, body in owner_functions
        if KEYNOTE_SLIDE_MEDIA_DATA_PREPARE_METADATA.search(body)
    ]
    if not prepared_bodies:
        violations.append(
            "focused litchi-keynote slide-media transaction must prepare metadata rewrite: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
        )
    else:
        prepared_body = "\n".join(prepared_bodies)
        if KEYNOTE_SLIDE_MEDIA_DATA_METADATA_EXECUTION_REQUIREMENTS.search(
            prepared_body
        ) is None:
            violations.append(
                "focused litchi-keynote slide-media metadata rewrite must consume "
                "execution_requirements: "
                f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )
        if KEYNOTE_SLIDE_MEDIA_DATA_PREPARED_EXECUTE.search(prepared_body) is None:
            violations.append(
                "focused litchi-keynote slide-media metadata rewrite must execute the "
                "prepared plan: "
                f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )
        if not re.search(r"\bmetadata_requirements\b", prepared_body):
            violations.append(
                "focused litchi-keynote slide-media metadata rewrite must charge prepared "
                "requirements through MediaBudget: "
                f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
            )
    direct_rewrite = KEYNOTE_SLIDE_MEDIA_DATA_DIRECT_METADATA_REWRITE.search(owner_code)
    if direct_rewrite is not None:
        line_number = owner_code.count("\n", 0, direct_rewrite.start()) + 1
        violations.append(
            "focused litchi-keynote slide-media owner must retire one-shot metadata "
            f"rewrite: {KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}:{line_number}"
        )

    reset = KEYNOTE_SLIDE_MEDIA_DATA_FOR_SOURCE_RESET.search(
        owner_code + "\n" + closure_code
    )
    if reset is not None:
        violations.append(
            "focused litchi-keynote slide-media transaction must not reset codec limits "
            "with DecodeOptions::for_source: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
        )

    # The closure child must be reachable from the owner, and it receives the
    # same ledger as the selector/metadata path.
    if KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_HELPER.search(closure_code) is None:
        violations.append(
            "focused litchi-keynote slide-media closure child is missing its validation helper: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE}"
        )
    closure_record = next(
        (
            (signature, body)
            for name, signature, body in closure_functions
            if name == "validate_selected_media_closure"
        ),
        None,
    )
    if closure_record is not None and KEYNOTE_SLIDE_MEDIA_DATA_MUTABLE_BUDGET_PARAMETER.search(
        closure_record[0]
    ) is None:
        violations.append(
            "focused litchi-keynote slide-media closure helper must thread &mut MediaBudget: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_CLOSURE_SOURCE}"
        )
    if not any(
        re.search(r"\bvalidate_selected_media_closure[ \t\r\n]*\(", body)
        for _name, _signature, body in owner_functions
    ):
        violations.append(
            "focused litchi-keynote slide-media owner must wire the private closure helper: "
            f"{KEYNOTE_SLIDE_MEDIA_DATA_OWNER_SOURCE}"
        )

    return sorted(set(violations))


def _keynote_slide_media_lifecycle_owner_present(root: Path) -> bool:
    """Return whether the focused duplicate/remove owner crossed its seam."""

    owner = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE
    package = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_EXPORT_SOURCES[0]
    if not owner.is_file() or not package.is_file():
        return False
    source = _mask_rust_non_code(
        _mask_rust_cfg_test_items(package.read_text(encoding="utf-8"))
    )
    return KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PACKAGE_MODULE.search(source) is not None


def _keynote_lifecycle_marker_present(
    source: str, marker: str | re.Pattern[str] | tuple[str, ...]
) -> bool:
    """Match one marker or one of a tuple of spelling alternatives."""

    if isinstance(marker, tuple):
        return any(_keynote_lifecycle_marker_present(source, item) for item in marker)
    if isinstance(marker, re.Pattern):
        return marker.search(source) is not None
    return marker in source


def audit_keynote_slide_media_lifecycle_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the focused lifecycle facade semantic and selector-first.

    The check activates only after ``Package`` wires the private owner.  Its
    declarations may then expose selectors, package values, and exact patches,
    but they may not expose archive objects, generated wire values, raw bytes,
    or native numeric identifiers.
    """

    if not _keynote_slide_media_lifecycle_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE
    package_path, lib_path = (
        root / path for path in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_EXPORT_SOURCES
    )
    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    owner_code = _mask_rust_non_code(owner)
    package_code = _mask_rust_non_code(package)
    library_code = _mask_rust_non_code(library)
    violations: list[str] = []

    if KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PUBLIC_MODULE.search(package + library):
        violations.append(
            "focused litchi-keynote slide-media lifecycle owner module must remain private: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_EXPORT_SOURCES[0]}"
        )

    declarations = dict(
        (name, declaration)
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Package")
    )
    edit_declarations = dict(
        (name, declaration)
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Edit")
    )
    for name in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PACKAGE_METHODS):
        declaration = declarations.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-media lifecycle Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
            continue
    for name in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_TYPED_PACKAGE_METHODS):
        declaration = edit_declarations.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-media lifecycle Edit method is missing "
                f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
            continue

    for name in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_TYPED_PACKAGE_METHODS):
        declaration = edit_declarations.get(name)
        if declaration is None:
            continue
        for selector in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SELECTOR_TYPES):
            if not re.search(rf"\b{re.escape(selector)}\b", declaration):
                violations.append(
                    "focused litchi-keynote slide-media lifecycle Package method must "
                    f"accept selector {selector} ({name}): "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
    for name in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PATCH_METHODS):
        declaration = declarations.get(name)
        if declaration is None:
            continue
        if not re.search(r"\b(?:LifecyclePatch|Patch)\b", declaration):
            violations.append(
                "focused litchi-keynote slide-media lifecycle apply method must accept "
                f"its typed patch ({name}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
    staged_declarations = [
        declarations.get(name) for name in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PACKAGE_METHODS)
    ] + [
        edit_declarations.get(name)
        for name in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_TYPED_PACKAGE_METHODS)
    ]
    for declaration in staged_declarations:
        if declaration is None:
            continue
        if re.search(r"\b(?:u64|u32|usize)\b|&\s*\[\s*u8\s*\]|Vec\s*<\s*u8\s*>", declaration):
            violations.append(
                "focused litchi-keynote slide-media lifecycle Package method exposes a raw "
                f"wire value: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        for match in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_RAW_ID_PARAMETER.finditer(declaration):
            violations.append(
                "focused litchi-keynote slide-media lifecycle Package method exposes raw "
                f"identifier {match.group(0).strip()}: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )

    all_public = owner + "\n" + package + "\n" + library
    exported_code = owner_code + "\n" + package_code + "\n" + library_code
    for role, alternatives in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CANONICAL_TYPE_ALTERNATIVES.items():
        if not any(
            re.search(
                rf"\b(?:pub(?:\([^()]*\))?\s+)?(?:struct|enum|type)\s+"
                rf"{re.escape(name)}\b",
                _mask_rust_non_code(all_public),
            )
            for name in alternatives
        ):
            violations.append(
                "focused litchi-keynote slide-media lifecycle public API is missing "
                f"canonical {role} type: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )

    facade_names = (
        set(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PACKAGE_METHODS)
        | KEYNOTE_SLIDE_MEDIA_LIFECYCLE_TYPED_PACKAGE_METHODS
        | KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PATCH_METHODS
        | KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SELECTOR_TYPES
        | set().union(*KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CANONICAL_TYPE_ALTERNATIVES.values())
    )
    physical = KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PHYSICAL_TYPES
    wire = KEYNOTE_SLIDE_MEDIA_LIFECYCLE_WIRE_TYPES
    for source, source_path in (
        (owner, owner_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = source_path == owner_path
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SEMANTIC_IDENTIFIERS:
                    continue
                if identifier in physical:
                    reason = "archive/IWA type"
                elif identifier in wire or identifier == "wire":
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-media lifecycle public API exposes "
                        f"{reason} {identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-media lifecycle public API retains a glob "
                    f"re-export: {source_path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote slide-media lifecycle public API exposes raw "
                    f"identifier {match.group(0).strip()}: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )
            if re.search(r"&\s*\[\s*u8\s*\]|Vec\s*<\s*u8\s*>", declaration):
                violations.append(
                    "focused litchi-keynote slide-media lifecycle public API exposes raw "
                    f"bytes: {source_path.relative_to(root)}:{line_number}"
                )

    # The owner may use low-level types privately, but public declarations
    # must still route through the exact source-bound patch contract.
    if not re.search(r"\b(?:patch|source_fingerprint|ExactArtifacts|inverse)\b", exported_code):
        violations.append(
            "focused litchi-keynote slide-media lifecycle facade is missing an exact-source "
            f"patch/inverse marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _audit_keynote_slide_media_lifecycle_comment_codec(
    root: Path,
) -> list[str]:
    """Keep the selected comment graph rewrite neutral and source-bound.

    The child is activated only once the hidden parent module is wired.  This
    lets the neutral codec land before the Keynote owner while still making a
    missing child or test seam fail as soon as the production module claims it.
    """

    parent_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_PARENT_SOURCE
    parent = parent_path.read_text(encoding="utf-8") if parent_path.is_file() else ""
    if re.search(
        r"(?ms)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\].{0,180}?"
        r"^\s*pub\s+mod\s+lifecycle\s*;",
        parent,
    ) is None:
        return []

    child_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_SOURCE
    test_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_TEST_SOURCE
    violations: list[str] = []
    if not child_path.is_file():
        return [
            "focused Keynote media lifecycle comment codec is missing its private lifecycle "
            f"child: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_SOURCE}"
        ]

    raw = child_path.read_text(encoding="utf-8")
    production = _mask_rust_cfg_test_items(raw)
    code = _mask_rust_non_code(production)
    for name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_APIS:
        kind = "fn" if name.startswith(("prepare_", "rewrite_")) else ""
        pattern = (
            rf"\b(?:pub\s+)?fn\s+{re.escape(name)}\b"
            if kind == "fn"
            else rf"\b(?:pub\s+)?(?:struct|enum|type|trait)\s+{re.escape(name)}\b"
        )
        if re.search(pattern, code) is None:
            violations.append(
                "focused Keynote media lifecycle comment codec is missing strict API "
                f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_MARKER_GROUPS.items():
        if not all(_keynote_lifecycle_marker_present(code, marker) for marker in markers):
            violations.append(
                "focused Keynote media lifecycle comment codec is missing "
                f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_SOURCE}"
            )

    for pattern in (
        re.compile(r"\bprost(?:_types)?\b"),
        re.compile(r"\bMessage\s*::\s*(?:decode|encode)\s*\("),
        re.compile(r"\b(?:to_owned_message|try_encode_to_vec|encode_to_vec)\s*\("),
    ):
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused Keynote media lifecycle comment codec must remain lazy/source-preserving; "
                f"eager generated operation {match.group(0).strip()}: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_SOURCE}:{line_number}"
            )
    for marker in ("HostPackage", "ArchiveObject", "SourceCatalog", "RawMessage"):
        if re.search(rf"\b{re.escape(marker)}\b", code):
            violations.append(
                "focused Keynote media lifecycle comment codec must remain neutral and archive-free "
                f"({marker}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_SOURCE}"
            )

    if not test_path.is_file():
        violations.append(
            "focused Keynote media lifecycle comment codec is missing direct lifecycle tests: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_TEST_SOURCE}"
        )
    else:
        test_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(test_path.read_text(encoding="utf-8"))
        )
        # The helper is itself a cfg(test) file, so retain the test attributes
        # from the raw source for this coverage-only assertion.
        test_raw = test_path.read_text(encoding="utf-8")
        for marker in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_TEST_MARKERS:
            if marker not in test_raw and marker not in test_code:
                violations.append(
                    "focused Keynote media lifecycle comment codec tests are missing "
                    f"{marker}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_CODEC_TEST_SOURCE}"
                )
    return sorted(set(violations))


def audit_keynote_slide_media_lifecycle_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the neutral bounded lazy lifecycle codec and private projection."""

    codec_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE
    if not codec_path.is_file():
        return []
    raw = codec_path.read_text(encoding="utf-8")
    source = _mask_rust_cfg_test_items(raw)
    code = _mask_rust_non_code(source)
    violations: list[str] = []
    for name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_APIS:
        kind = "fn" if name.startswith(("decode_", "prepare_", "rewrite_")) else ""
        pattern = (
            rf"\b(?:pub\s+)?fn\s+{re.escape(name)}\b"
            if kind == "fn"
            else rf"\b(?:pub\s+)?(?:struct|enum|type|trait)\s+{re.escape(name)}\b"
        )
        if re.search(pattern, code) is None:
            violations.append(
                "focused Keynote media lifecycle codec is missing strict API "
                f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE}"
            )

    for label, markers in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_MARKER_GROUPS.items():
        if not all(_keynote_lifecycle_marker_present(code, marker) for marker in markers):
            violations.append(
                "focused Keynote media lifecycle codec is missing "
                f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE}"
            )

    eager_patterns = (
        re.compile(r"\bprost(?:_types)?\b"),
        re.compile(r"\bMessage\s*::\s*(?:decode|encode)\s*\("),
        re.compile(r"\b(?:to_owned_message|try_encode_to_vec|encode_to_vec)\s*\("),
    )
    for pattern in eager_patterns:
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused Keynote media lifecycle codec must remain lazy/source-preserving; "
                f"eager generated operation {match.group(0).strip()}: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE}:{line_number}"
            )
    for marker in ("HostPackage", "ArchiveObject", "SourceCatalog"):
        if re.search(rf"\b{re.escape(marker)}\b", code):
            violations.append(
                "focused Keynote media lifecycle codec must remain neutral and archive-free "
                f"({marker}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE}"
            )

    public_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_PUBLIC_SOURCE
    public_source = (
        _mask_rust_cfg_test_items(public_path.read_text(encoding="utf-8"))
        if public_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\S]{{0,240}}?"
        rf"^\s*pub\s+mod\s+{re.escape(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_MODULE)}\s*;",
        public_source,
        re.MULTILINE,
    ) is None:
        violations.append(
            "focused Keynote media lifecycle codec is missing hidden public module "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_MODULE}: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_PUBLIC_SOURCE}"
        )
    child_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE
    child = (
        _mask_rust_cfg_test_items(child_path.read_text(encoding="utf-8"))
        if child_path.is_file()
        else ""
    )
    child_code = _mask_rust_non_code(child)
    if child_path.is_file():
        for name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_APIS:
            kind = "fn" if name.startswith(("decode_", "prepare_", "rewrite_")) else ""
            pattern = (
                rf"\b(?:pub\s+)?fn\s+{re.escape(name)}\b"
                if kind == "fn"
                else rf"\b(?:pub\s+)?(?:struct|enum|type|trait)\s+{re.escape(name)}\b"
            )
            if re.search(pattern, child_code) is None:
                violations.append(
                    "focused Keynote media lifecycle node-cache codec is missing strict API "
                    f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE}"
                )
        for label, markers in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_MARKER_GROUPS.items():
            if not all(_keynote_lifecycle_marker_present(child_code, marker) for marker in markers):
                violations.append(
                    "focused Keynote media lifecycle node-cache codec is missing "
                    f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE}"
                )
        for pattern in eager_patterns:
            for match in pattern.finditer(child_code):
                line_number = child_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused Keynote media lifecycle codec must remain lazy/source-preserving; "
                    f"eager generated operation {match.group(0).strip()}: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE}:{line_number}"
                )
        for marker in ("HostPackage", "ArchiveObject", "SourceCatalog"):
            if re.search(rf"\b{re.escape(marker)}\b", child_code):
                violations.append(
                    "focused Keynote media lifecycle codec must remain neutral and archive-free "
                    f"({marker}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE}"
                )
    else:
        violations.append(
            "focused Keynote media lifecycle codec is missing private node-cache child: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_NODE_CACHE_SOURCE}"
        )
    if re.search(
        r"(?ms)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\].{0,180}?"
        r"^\s*pub\s+mod\s+node_cache\s*;",
        source,
    ) is None:
        violations.append(
            "focused Keynote media lifecycle codec is missing its hidden node-cache module: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_SOURCE}"
        )
    projection = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE
    violations.extend(_audit_keynote_media_lifecycle_projection(root, projection))
    violations.extend(_audit_keynote_slide_media_lifecycle_comment_codec(root))
    return sorted(set(violations))


def _audit_keynote_media_lifecycle_projection(
    root: Path, projection_path: Path
) -> list[str]:
    """Check native field numbers and the private Buffa build seam."""

    violations: list[str] = []
    raw = projection_path.read_text(encoding="utf-8") if projection_path.is_file() else ""
    projection = re.sub(r"//[^\n]*|/\*.*?\*/", "", raw, flags=re.DOTALL)
    if not projection:
        return [
            "focused Keynote media lifecycle codec is missing private Buffa projection: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE}"
        ]
    for marker in ("syntax = \"proto2\"", "package LitchiIwaProjection"):
        if marker not in projection:
            violations.append(
                "focused Keynote media lifecycle private projection is missing marker "
                f"{marker}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE}"
            )
    repeated = re.search(r"\brepeated\b", projection)
    if repeated is not None:
        line_number = projection.count("\n", 0, repeated.start()) + 1
        violations.append(
            "focused Keynote media lifecycle private projection must not expose generated "
            f"repeated views: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE}:{line_number}"
        )
    for message, fields in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_MESSAGES.items():
        if re.search(rf"\bmessage\s+{re.escape(message)}\s*\{{", projection) is None:
            violations.append(
                "focused Keynote media lifecycle private projection is missing message "
                f"{message}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE}"
            )
            continue
        for field, number in fields.items():
            if re.search(
                rf"\b(?:optional|repeated|required)\s+[A-Za-z_][A-Za-z0-9_.]*\s+"
                rf"{re.escape(field)}\s*=\s*{number}\b",
                projection,
            ) is None:
                violations.append(
                    "focused Keynote media lifecycle private projection has incorrect or "
                    f"missing field {message}.{field}={number}: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PROJECTION_SOURCE}"
                )

    build_path = root / Path("crates/litchi-iwa-protos/build.rs")
    build = build_path.read_text(encoding="utf-8") if build_path.is_file() else ""
    build_route = re.search(
        r"(?s)\blet\s+buffa_keynote_media_lifecycle_out_directory\b.*?"
        r"\benforce_keynote_media_lifecycle_projection_budget\s*\(\s*"
        r"&\s*buffa_keynote_media_lifecycle_out_directory\s*,?\s*\)",
        build,
    )
    if build_route is None:
        violations.append(
            "focused Keynote media lifecycle Buffa build is missing its dedicated route: "
            "crates/litchi-iwa-protos/build.rs"
        )
        build_focus = ""
    else:
        build_focus = build_route.group(0)
    build_markers = (
        "KNMediaLifecycleArchive.proto",
        "buffa-keynote-media-lifecycle",
        "iwa_keynote_media_lifecycle_buffa_protos.rs",
        ".generate_views(true)",
        ".lazy_views(true)",
        ".preserve_unknown_fields(false)",
        "enforce_keynote_media_lifecycle_projection_budget",
    )
    for marker in build_markers:
        if marker not in build_focus:
            violations.append(
                "focused Keynote media lifecycle Buffa build is missing marker "
                f"{marker}: crates/litchi-iwa-protos/build.rs"
            )
    budget_function = re.search(
        r"(?s)\bfn\s+enforce_keynote_media_lifecycle_projection_budget\b.*?"
        r"(?=\nfn\s|\Z)",
        build,
    )
    budget_source = budget_function.group(0) if budget_function is not None else ""
    for marker in (
        "RepeatedView",
        "LazyRepeatedView",
        "const EXPECTED_FILES: usize = 5",
        "const MAX_GENERATED_BYTES: u64 = 192 * 1024",
        "repeated_views != 0",
        "lazy_repeated_views != 0",
    ):
        if marker not in budget_source:
            violations.append(
                "focused Keynote media lifecycle Buffa budget is missing marker "
                f"{marker}: crates/litchi-iwa-protos/build.rs"
            )
    public_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_PUBLIC_SOURCE
    public = public_path.read_text(encoding="utf-8") if public_path.is_file() else ""
    if re.search(
        r"(?ms)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\].{0,260}?"
        r"^\s*mod\s+buffa_keynote_media_lifecycle_generated\b",
        public,
    ) is None:
        violations.append(
            "focused Keynote media lifecycle generated Buffa projection must remain private: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_PUBLIC_SOURCE}"
        )
    if (
        "buffa-keynote-media-lifecycle" not in public
        or "iwa_keynote_media_lifecycle_buffa_protos.rs" not in public
    ):
        violations.append(
            "focused Keynote media lifecycle generated projection include is missing: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CODEC_PUBLIC_SOURCE}"
        )
    return violations


def audit_keynote_slide_media_lifecycle_transaction_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one bounded, source-authoritative lifecycle transaction."""

    if not _keynote_slide_media_lifecycle_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE
    owner_raw = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner = _mask_rust_non_code(owner_raw)
    violations: list[str] = []
    child_sources_by_name: dict[str, str] = {}
    for relative in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_SOURCES:
        path = root / relative
        if path.is_file():
            child_sources_by_name[path.stem] = _mask_rust_non_code(
                _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
            )
    child_sources = list(child_sources_by_name.values())
    lifecycle_code = owner + "\n" + "\n".join(child_sources)

    module_names = {
        match.group(1)
        for match in re.finditer(
            r"(?m)^\s*(?:pub(?:\([^()]*\))?\s+)?mod\s+([A-Za-z_][A-Za-z0-9_]*)\s*;",
            owner,
        )
    }
    for child in (
        "comment_graph",
        "comment_removal",
        "graph",
        "identifier_watermark",
        "metadata",
        "budget",
        "clone_payload",
        "node_cache",
    ):
        if child not in module_names:
            violations.append(
                "focused Keynote media lifecycle owner must wire private "
                f"{child} child: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        child_path = root / KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / f"{child}.rs"
        if not child_path.is_file():
            violations.append(
                "focused Keynote media lifecycle owner is missing private child "
                f"{child}: {child_path.relative_to(root)}"
            )
    if KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PUBLIC_MODULE.search(owner):
        violations.append(
            "focused Keynote media lifecycle child module must remain private: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
        )
    for match in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_PUBLIC_CHILD_MODULE.finditer(owner):
        visibility = match.group("visibility")
        name = match.group("name").removeprefix("r#")
        if (
            visibility == KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SHARED_CHILD_VISIBILITY
            and name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_SHARED_CHILD_MODULES
        ):
            continue
        violations.append(
            "focused Keynote media lifecycle child modules must remain private: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
        )

    # The graph owns source-preserving slide-header replacement and the
    # production payload remapper.  The latter accepts the operation's
    # charging callback; keeping the old cfg(test) convenience wrapper out of
    # this production seam prevents a clone from escaping the transaction's
    # aggregate ledger.
    graph_code = child_sources_by_name.get("graph", "")
    if graph_code:
        def graph_function_body(name: str) -> str:
            declaration = re.search(
                rf"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
                rf"(?:unsafe[ \t]+|async[ \t]+|const[ \t]+)*fn[ \t]+"
                rf"{re.escape(name)}\b",
                graph_code,
            )
            if declaration is None:
                return ""
            opening = graph_code.find("{", declaration.end())
            if opening < 0:
                return ""
            depth = 1
            cursor = opening + 1
            while cursor < len(graph_code) and depth:
                if graph_code[cursor] == "{":
                    depth += 1
                elif graph_code[cursor] == "}":
                    depth -= 1
                cursor += 1
            return graph_code[opening + 1 : cursor - 1] if depth == 0 else ""

        def graph_call_records(
            body: str, name: str
        ) -> list[tuple[re.Match[str], str]]:
            records: list[tuple[re.Match[str], str]] = []
            for call in re.finditer(rf"\b{re.escape(name)}\s*\(", body):
                opening = body.find("(", call.start(), call.end())
                depth = 1
                cursor = opening + 1
                while cursor < len(body) and depth:
                    if body[cursor] == "(":
                        depth += 1
                    elif body[cursor] == ")":
                        depth -= 1
                    cursor += 1
                arguments = body[opening + 1 : cursor - 1] if depth == 0 else ""
                records.append((call, arguments))
            return records

        if KEYNOTE_SLIDE_MEDIA_LIFECYCLE_HEADER_REWRITE_HELPER.search(graph_code) is None:
            violations.append(
                "focused Keynote media lifecycle graph is missing its source-preserving "
                "slide-header helper: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        if KEYNOTE_SLIDE_MEDIA_LIFECYCLE_LEGACY_CLONE_PAYLOAD_CALL.search(graph_code):
            violations.append(
                "focused Keynote media lifecycle graph must not call the cfg(test) clone "
                "payload wrapper: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        clone_call = KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CLONE_PAYLOAD_CALL.search(graph_code)
        if clone_call is None:
            violations.append(
                "focused Keynote media lifecycle graph is missing its budgeted clone "
                "payload route: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        else:
            opening = graph_code.find("(", clone_call.start(), clone_call.end())
            arguments = ""
            if opening >= 0:
                depth = 1
                cursor = opening + 1
                while cursor < len(graph_code) and depth:
                    if graph_code[cursor] == "(":
                        depth += 1
                    elif graph_code[cursor] == ")":
                        depth -= 1
                    cursor += 1
                if depth == 0:
                    arguments = graph_code[opening + 1 : cursor - 1]
            if not re.search(r"&\s*mut\s+charge\b", arguments):
                violations.append(
                    "focused Keynote media lifecycle graph must pass its shared budget "
                    "charge callback to the clone payload route: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )
            charge_closure = re.search(
                r"\blet\s+mut\s+charge\b(?P<body>[\s\S]{0,900})",
                graph_code,
            )
            if charge_closure is None or not re.search(
                r"\bbudget\s*\.\s*charge_[A-Za-z0-9_]*\s*\(",
                charge_closure.group("body"),
            ):
                violations.append(
                    "focused Keynote media lifecycle graph must charge clone payload "
                    "allocations through the operation budget: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )

        def require_core_budget_wrapper(
            name: str, arguments_pattern: str, label: str
        ) -> str:
            body = graph_function_body(name)
            if not body:
                violations.append(
                    "focused Keynote media lifecycle graph is missing its bounded core-header "
                    f"{label} wrapper: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )
                return ""
            inner_calls = graph_call_records(body, "reserve_core_header_work_inner")
            if not any(re.search(arguments_pattern, arguments) for _call, arguments in inner_calls):
                violations.append(
                    "focused Keynote media lifecycle core-header "
                    f"{label} wrapper must delegate source, limits, and the shared budget "
                    f"to reserve_core_header_work_inner: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )
            if re.search(r"\bcharge_(?:allocation_plan|wire_work)\b", body):
                violations.append(
                    "focused Keynote media lifecycle core-header wrappers must leave atomic "
                    "allocation and wire charging to the shared inner helper: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )
            return body

        require_core_budget_wrapper(
            "reserve_core_header_inspection_work",
            r"\bsource\b\s*,\s*0\s*,\s*\blimits\b\s*,\s*\bbudget\b\s*,\s*false\b",
            "inspection",
        )
        require_core_budget_wrapper(
            "reserve_core_header_work",
            r"\bsource\b\s*,\s*\badditional_references\b\s*,\s*\blimits\b\s*,\s*\bbudget\b\s*,\s*true\b",
            "mutation",
        )
        inner_body = graph_function_body("reserve_core_header_work_inner")
        if not inner_body:
            violations.append(
                "focused Keynote media lifecycle graph is missing its shared core-header "
                f"reservation inner helper: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        else:
            inner_allocation_calls = graph_call_records(
                inner_body, "charge_allocation_plan"
            )
            if not any(
                re.search(r"\bbytes\b\s*,\s*\bevents\b", arguments)
                for _call, arguments in inner_allocation_calls
            ):
                violations.append(
                    "focused Keynote media lifecycle core-header inner helper must charge "
                    "one atomic byte/event allocation plan: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )
            if not re.search(
                r"\bbudget\s*\.\s*charge_wire_work\s*\(\s*bytes\s*\)",
                inner_body,
            ):
                violations.append(
                    "focused Keynote media lifecycle core-header inner helper must charge "
                    "its bounded wire work through the shared budget: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )

        def require_core_reservation(
            body: str, operation: str, core_call: str
        ) -> None:
            reservations = graph_call_records(body, "reserve_core_header_work")
            core_calls = graph_call_records(body, core_call)
            if len(core_calls) != 1 or not reservations:
                violations.append(
                    "focused Keynote media lifecycle graph must reserve core-header "
                    f"work before {operation}: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )
                return
            all_core_calls = [
                (call, name)
                for name in (
                    "clone_with_identity_remap",
                    "replace_message_pruning_object_references_preserving_header",
                    "replace_message_transitioning_object_references_preserving_header",
                )
                for call, _arguments in graph_call_records(body, name)
            ]
            if len(reservations) != len(all_core_calls):
                violations.append(
                    "focused Keynote media lifecycle graph must reserve core-header work "
                    f"once before every core publication ({operation}): "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )
                return
            preceding = [
                arguments
                for call, arguments in reservations
                if call.start() < core_calls[0][0].start()
            ]
            last_reservation = next(
                (
                    call
                    for call, _arguments in reversed(reservations)
                    if call.start() < core_calls[0][0].start()
                ),
                None,
            )
            intervening_core_call = last_reservation is not None and any(
                last_reservation.start() < call.start() < core_calls[0][0].start()
                for call, _name in all_core_calls
            )
            if (
                not preceding
                or not re.search(r"\bbudget\b", preceding[-1])
                or intervening_core_call
            ):
                violations.append(
                    "focused Keynote media lifecycle graph must route the core-header "
                    f"reservation budget before {operation}: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
                )

        require_core_reservation(
            graph_function_body("clone_object"),
            "clone core publication",
            "clone_with_identity_remap",
        )
        header_body = graph_function_body("replace_slide_message_with_lifecycle_refs")
        require_core_reservation(
            header_body,
            "prune core publication",
            "replace_message_pruning_object_references_preserving_header",
        )
        require_core_reservation(
            header_body,
            "transition core publication",
            "replace_message_transitioning_object_references_preserving_header",
        )

    budget_code = child_sources_by_name.get("budget", "")
    if budget_code:
        if not re.search(r"\bfn\s+charge_allocation_plan\b", budget_code):
            violations.append(
                "focused Keynote media lifecycle budget is missing its atomic allocation "
                "plan method: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'budget.rs'}"
            )
        for marker in (
            "observed_bytes",
            "observed_events",
            "self.allocation_bytes = observed_bytes",
            "self.allocations = observed_events",
        ):
            if marker not in budget_code:
                violations.append(
                    "focused Keynote media lifecycle budget must atomically check and commit "
                    f"allocation bytes/events ({marker}): "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'budget.rs'}"
                )
        if not re.search(
            r"\bcharge_allocations\b[\s\S]{0,320}?"
            r"\bcharge_allocation_plan\s*\(\s*amount\s*,\s*1\s*\)",
            budget_code,
        ):
            violations.append(
                "focused Keynote media lifecycle budget must route single allocation charges "
                "through its atomic allocation plan: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'budget.rs'}"
            )

    node_cache_code = child_sources_by_name.get("node_cache", "")
    if node_cache_code:
        for name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_NODE_CACHE_APIS:
            if re.search(rf"\bfn\s+{re.escape(name)}\b", node_cache_code) is None:
                violations.append(
                    "focused Keynote media lifecycle node-cache child is missing helper "
                    f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'node_cache.rs'}"
                )
        for label, markers in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_NODE_CACHE_MARKER_GROUPS.items():
            if not all(_keynote_lifecycle_marker_present(node_cache_code, marker) for marker in markers):
                violations.append(
                    "focused Keynote media lifecycle node-cache child is missing "
                    f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'node_cache.rs'}"
                )
        for marker in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_NODE_CACHE_REQUIRED_TYPES:
            if not re.search(rf"\b{re.escape(marker)}\b", node_cache_code):
                violations.append(
                    "focused Keynote media lifecycle node-cache child is missing validation "
                    f"marker {marker}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'node_cache.rs'}"
                )
        for helper in (
            "prepare_slide_node_build_cache",
            "validate_candidate_package_node_cache",
        ):
            call = re.search(
                rf"\bnode_cache\s*::\s*{re.escape(helper)}\s*\(", owner
            )
            if call is None:
                violations.append(
                    "focused Keynote media lifecycle transaction must route its node-cache "
                    f"helper {helper} through the owner: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
                continue
            call_window = owner[call.start() : call.start() + 900]
            if not re.search(r"\bbudget\b", call_window):
                violations.append(
                    "focused Keynote media lifecycle transaction must route its shared budget "
                    f"through node-cache helper {helper}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )

    comment_graph_code = child_sources_by_name.get("comment_graph", "")
    if comment_graph_code:
        for name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_GRAPH_APIS:
            kind = "fn" if name.startswith("plan_") else ""
            pattern = (
                rf"\b(?:pub(?:\([^()]*\))?\s+)?fn\s+{re.escape(name)}\b"
                if kind == "fn"
                else rf"\b(?:pub(?:\([^()]*\))?\s+)?(?:struct|enum|type|trait)\s+{re.escape(name)}\b"
            )
            if re.search(pattern, comment_graph_code) is None:
                violations.append(
                    "focused Keynote media lifecycle comment graph is missing strict API "
                    f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_graph.rs'}"
                )
        for label, markers in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_GRAPH_MARKER_GROUPS.items():
            if not all(_keynote_lifecycle_marker_present(comment_graph_code, marker) for marker in markers):
                violations.append(
                    "focused Keynote media lifecycle comment graph is missing "
                    f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_graph.rs'}"
                )

        graph_code = child_sources_by_name.get("graph", "")
        strict_body = _rust_any_function_body(comment_graph_code, "plan_comment_graph") or ""
        if (
            not strict_body
            or "CommentGraphPayloadPolicy::Strict" not in strict_body
            or "CommentGraphComponentPolicy::SameComponent" not in strict_body
            or "plan_comment_graph_with_policy" not in strict_body
        ):
            violations.append(
                "focused Keynote media lifecycle comment graph must keep its clone wrapper strict: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_graph.rs'}"
            )
        cross_component_body = (
            _rust_any_function_body(comment_graph_code, "plan_comment_graph_cross_component")
            or ""
        )
        if (
            not cross_component_body
            or "CommentGraphPayloadPolicy::PreserveRootExtensions" not in cross_component_body
            or "CommentGraphComponentPolicy::Package" not in cross_component_body
            or "plan_comment_graph_with_policy" not in cross_component_body
        ):
            violations.append(
                "focused Keynote media lifecycle comment graph must keep its package-scope "
                "cross-component wrapper explicit: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_graph.rs'}"
            )
        policy_body = _rust_any_function_body(comment_graph_code, "plan_comment_graph_with_policy") or ""
        if not re.search(
            r"matches!?\s*\(\s*component_policy\s*,\s*"
            r"CommentGraphComponentPolicy::SameComponent\s*\)",
            policy_body,
        ):
            violations.append(
                "focused Keynote media lifecycle comment graph must keep same-component "
                "checks inside its policy helper: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_graph.rs'}"
            )
        if not re.search(
            r"storage_component\s*!=\s*component_name", policy_body
        ):
            violations.append(
                "focused Keynote media lifecycle comment graph must validate each strict "
                "storage node against its selected component: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_graph.rs'}"
            )
        if not re.search(
            r"\bplan_comment_graph\s*\([^)]*\bbudget\b",
            graph_code,
            re.DOTALL,
        ):
            violations.append(
                "focused Keynote media lifecycle graph must route the shared budget through "
                f"plan_comment_graph: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        if re.search(r"\bplan_comment_graph_cross_component\s*\(", graph_code):
            violations.append(
                "focused Keynote media lifecycle graph must not use the package-scope "
                "cross-component comment plan: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        private_graph_body = graph_function_body("private_graph")
        if not re.search(
            r"comment_graph[\s\S]{0,520}?\bUnsupportedComment\b",
            private_graph_body,
        ):
            violations.append(
                "focused Keynote media lifecycle graph must keep unsupported caption/comment "
                "edges fail-closed: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        clone_routes = graph_call_records(graph_code, "rewrite_comment_payload")
        if not clone_routes or not any(
            re.search(r"\bbudget\b", arguments)
            for _call, arguments in clone_routes
        ):
            violations.append(
                "focused Keynote media lifecycle graph must route comment cloning through its "
                f"shared budget: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        author_dependency_marker = re.compile(
            r"\bauthor_ids\b[\s\S]{0,420}?\bcontinue\b"
        )
        if author_dependency_marker.search(graph_code) is None:
            violations.append(
                "focused Keynote media lifecycle graph is missing shared author dependency "
                "skip semantics: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'graph.rs'}"
            )
        if re.search(
            r"\b(?:storage_uuid|root_storage_uuid|storage_identities)\b",
            comment_graph_code,
        ) is None:
            violations.append(
                "focused Keynote media lifecycle comment graph is missing source storage UUID "
                "witness marker: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_graph.rs'}"
            )

    comment_removal_code = child_sources_by_name.get("comment_removal", "")
    if "comment_removal" in module_names:
        for name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_REMOVAL_APIS:
            kind = "fn" if name.startswith("plan_") else ""
            pattern = (
                rf"\b(?:pub(?:\([^()]*\))?\s+)?fn\s+{re.escape(name)}\b"
                if kind == "fn"
                else rf"\b(?:pub(?:\([^()]*\))?\s+)?(?:struct|enum|type|trait)\s+{re.escape(name)}\b"
            )
            if re.search(pattern, comment_removal_code) is None:
                violations.append(
                    "focused Keynote media lifecycle comment removal is missing strict API "
                    f"{name}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
                )
        for label, markers in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_COMMENT_REMOVAL_MARKER_GROUPS.items():
            if not all(
                _keynote_lifecycle_marker_present(comment_removal_code, marker)
                for marker in markers
            ):
                violations.append(
                    "focused Keynote media lifecycle comment removal is missing "
                    f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
                )

        known_payload_call = comment_removal_code.find(
            "validate_known_payload_relationships("
        )
        census_call = comment_removal_code.find("strict_reference_census(")
        if (
            known_payload_call < 0
            or census_call < 0
            or known_payload_call > census_call
        ):
            violations.append(
                "focused Keynote media lifecycle comment removal must validate known Movie "
                "and CommentStorage payload/header relationships before its strict census: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )

        def comment_function_body(name: str) -> str:
            declaration = re.search(
                rf"(?m)^\s*(?:pub(?:\([^()]*\))?\s+)?"
                rf"(?:unsafe\s+|async\s+|const\s+)*fn\s+"
                rf"{re.escape(name)}\b",
                comment_removal_code,
            )
            if declaration is None:
                return ""
            opening = comment_removal_code.find("{", declaration.end())
            if opening < 0:
                return ""
            depth = 1
            cursor = opening + 1
            while cursor < len(comment_removal_code) and depth:
                if comment_removal_code[cursor] == "{":
                    depth += 1
                elif comment_removal_code[cursor] == "}":
                    depth -= 1
                cursor += 1
            return (
                comment_removal_code[opening + 1 : cursor - 1]
                if depth == 0
                else ""
            )

        def comment_call_arguments(body: str, name: str) -> list[str]:
            arguments: list[str] = []
            for call in re.finditer(rf"\b{re.escape(name)}\s*\(", body):
                opening = body.find("(", call.start(), call.end())
                if opening < 0:
                    continue
                depth = 1
                cursor = opening + 1
                while cursor < len(body) and depth:
                    if body[cursor] == "(":
                        depth += 1
                    elif body[cursor] == ")":
                        depth -= 1
                    cursor += 1
                if depth == 0:
                    arguments.append(body[opening + 1 : cursor - 1])
            return arguments

        # The payload admission call is source-bound: resolving a numeric
        # reference without the package and owning component would allow a
        # same-ID object from another component to satisfy the check.
        removal_body = comment_function_body("plan_comment_removal")
        known_payload_arguments = comment_call_arguments(
            removal_body, "validate_known_payload_relationships"
        )
        if not any(
            re.search(r"\bpackage\b", arguments)
            and re.search(r"\bcomponent\.name\s*\(\s*\)", arguments)
            for arguments in known_payload_arguments
        ):
            violations.append(
                "focused Keynote media lifecycle known payload admission must pass the "
                "package and owning component for target resolution: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )

        movie_body = comment_function_body("validate_movie_payload_relationship")
        movie_target_arguments = comment_call_arguments(
            movie_body, "validate_comment_storage_target"
        )
        if not any(
            re.search(r"\bpackage\b", arguments)
            and re.search(r"\bcomponent_name\b", arguments)
            and re.search(r"\b(?:identifier|direct_comment)\b", arguments)
            for arguments in movie_target_arguments
        ):
            violations.append(
                "focused Keynote media lifecycle Movie validation must resolve each comment "
                "target against its package and owning component: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )

        storage_body = comment_function_body("validate_comment_storage_relationship")
        storage_target_arguments = comment_call_arguments(
            storage_body, "validate_comment_storage_target"
        )
        author_target_arguments = comment_call_arguments(
            storage_body, "validate_annotation_author_target"
        )
        if not re.search(r"\blet\s+facts\s*=", storage_body):
            violations.append(
                "focused Keynote media lifecycle CommentStorage validation must retain the "
                "neutral StorageFacts result: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )
        if not re.search(r"\bfacts\.author_identifier\b", storage_body) or not any(
            re.search(r"\bpackage\b", arguments)
            and re.search(r"\bauthor_identifier\b", arguments)
            for arguments in author_target_arguments
        ):
            violations.append(
                "focused Keynote media lifecycle CommentStorage validation must resolve its "
                "author target through the package: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )
        if not re.search(r"\bfacts\.reply_identifiers\b", storage_body) or not any(
            re.search(r"\bpackage\b", arguments)
            and re.search(r"\bcomponent_name\b", arguments)
            and re.search(r"\breply_identifier\b", arguments)
            for arguments in storage_target_arguments
        ):
            violations.append(
                "focused Keynote media lifecycle CommentStorage validation must resolve each "
                "reply target against its package and owning component: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )

        storage_target_body = comment_function_body("validate_comment_storage_target")
        if not re.search(
            r"\bobject_with_component\s*\(", storage_target_body
        ) or not re.search(
            r"\bactual_component\s*!=\s*component_name\b", storage_target_body
        ) or not re.search(
            r"\bvalidate_single_message_type\s*\(", storage_target_body
        ) or not re.search(
            r"\bCOMMENT_STORAGE_MESSAGE_TYPE\b", storage_target_body
        ):
            violations.append(
                "focused Keynote media lifecycle storage target validation must enforce the "
                "same-component CommentStorage object type: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )

        author_target_body = comment_function_body("validate_annotation_author_target")
        if not re.search(
            r"\bobject_with_component\s*\(", author_target_body
        ) or not re.search(
            r"\bvalidate_single_message_type\s*\([^)]*\b212\b", author_target_body
        ):
            violations.append(
                "focused Keynote media lifecycle author target validation must resolve an "
                "exact annotation-author object: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )

        single_target_body = comment_function_body("validate_single_message_type")
        for marker in (
            "object.archive_info.identifier",
            "object.messages.len()",
            "object.archive_info.message_infos.len()",
            "message.type_",
            "info.type_",
            "info.length",
            "message.data.len()",
        ):
            if marker not in single_target_body:
                violations.append(
                    "focused Keynote media lifecycle target validation must verify the exact "
                    f"singleton MessageInfo/message type and length ({marker}): "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
                )

        removal_call = re.search(
            r"\bcomment_removal\s*::\s*plan_comment_removal\s*\(", owner
        )
        removal_arguments = ""
        if removal_call is not None:
            opening = owner.find("(", removal_call.start(), removal_call.end())
            depth = 1
            cursor = opening + 1
            while cursor < len(owner) and depth:
                if owner[cursor] == "(":
                    depth += 1
                elif owner[cursor] == ")":
                    depth -= 1
                cursor += 1
            if depth == 0:
                removal_arguments = owner[opening + 1 : cursor - 1]
        if removal_call is None or not re.search(r"\bbudget\b", removal_arguments):
            violations.append(
                "focused Keynote media lifecycle removal must route plan_comment_removal "
                "through the shared budget: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        else:
            removal_window = (
                owner[max(0, removal_call.start() - 720) : removal_call.start() + 720]
            )
            if not re.search(r"\bLifecycleAction::Remove\b", removal_window):
                violations.append(
                    "focused Keynote media lifecycle removal plan must be gated by the "
                    f"remove action: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
            if not re.search(r"\bremoved_object_ids\b", owner):
                violations.append(
                    "focused Keynote media lifecycle removal must filter physical source IDs "
                    f"through CommentRemovalPlan: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
            if not re.search(
                r"\b(?:unused_external_author_ids|external_author_ids)\b", owner
            ):
                violations.append(
                    "focused Keynote media lifecycle removal must preserve author diagnostics "
                    f"through the owner: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
            if not re.search(
                r"\b(?:with_external_reference_removals|external_reference_removals)\b",
                lifecycle_code,
            ):
                violations.append(
                    "focused Keynote media lifecycle removal must route exact external-edge "
                    f"removals through metadata: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )

        # Physical author objects are retained by native Keynote.  A removal
        # planner may report unused author dependencies, but it must not turn
        # those diagnostics into the effective object-removal list.
        if re.search(
            r"for[^{}\n]*\b(?:external_author_ids|unused_external_author_ids)\b"
            r"[\s\S]{0,520}?\b(?:push_vec|insert_sorted_unique)\s*\(\s*&mut\s*removed_object_ids",
            comment_removal_code,
        ):
            violations.append(
                "focused Keynote media lifecycle comment removal must never delete author "
                f"objects: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'comment_removal.rs'}"
            )

    identifier_watermark_code = child_sources_by_name.get("identifier_watermark", "")
    if "identifier_watermark" in module_names:
        for name in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_IDENTIFIER_WATERMARK_APIS:
            if re.search(rf"\bfn\s+{re.escape(name)}\b", identifier_watermark_code) is None:
                violations.append(
                    "focused Keynote media lifecycle identifier watermark is missing strict "
                    f"API {name}: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'identifier_watermark.rs'}"
                )
        for label, markers in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_IDENTIFIER_WATERMARK_MARKER_GROUPS.items():
            if not all(
                _keynote_lifecycle_marker_present(identifier_watermark_code, marker)
                for marker in markers
            ):
                violations.append(
                    "focused Keynote media lifecycle identifier watermark is missing "
                    f"{label} marker: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'identifier_watermark.rs'}"
                )

        def owner_call_arguments(name: str) -> list[tuple[int, str]]:
            calls: list[tuple[int, str]] = []
            for call in re.finditer(rf"\b{re.escape(name)}\s*\(", owner):
                opening = owner.find("(", call.start(), call.end())
                if opening < 0:
                    continue
                depth = 1
                cursor = opening + 1
                while cursor < len(owner) and depth:
                    if owner[cursor] == "(":
                        depth += 1
                    elif owner[cursor] == ")":
                        depth -= 1
                    cursor += 1
                if depth == 0:
                    calls.append((call.start(), owner[opening + 1 : cursor - 1]))
            return calls

        release_calls = owner_call_arguments("identifier_watermark::plan_release")
        if not release_calls:
            violations.append(
                "focused Keynote media lifecycle removal must route the effective source IDs "
                "through identifier_watermark::plan_release: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        else:
            release_position, release_arguments = release_calls[0]
            if not all(
                re.search(pattern, release_arguments)
                for pattern in (
                    r"\bsource\b",
                    r"\bsource_ids\b",
                    r"\bsnapshot\.last_identifier\s*\(\s*\)",
                    r"\bbudget\b",
                )
            ):
                violations.append(
                    "focused Keynote media lifecycle release planning must use source_ids, the "
                    "source watermark, and the shared budget: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
            release_window = owner[max(0, release_position - 720) : release_position + 720]
            if not re.search(r"\bLifecycleAction::Remove\b", release_window):
                violations.append(
                    "focused Keynote media lifecycle release planning must remain remove-only: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
            release_prefix = owner[max(0, release_position - 300) : release_position + 260]
            if not re.search(
                r"if\s+action\s*==\s*LifecycleAction::Remove\s*\{"
                r"[\s\S]{0,260}identifier_watermark::plan_release",
                release_prefix,
            ):
                violations.append(
                    "focused Keynote media lifecycle release planning must be gated by the "
                    "remove action directly: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
            source_assignment = re.search(
                r"\blet\s+source_ids\s*=\s*(?:removal_plan|plan)"
                r"[\s\S]{0,420}?\bremoved_object_ids\b",
                owner,
            )
            if source_assignment is None or source_assignment.start() > release_position:
                violations.append(
                    "focused Keynote media lifecycle release planning must follow the filtered "
                    "effective removed-object set: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )

        verify_calls = owner_call_arguments("identifier_watermark::verify_release")
        if not verify_calls:
            violations.append(
                "focused Keynote media lifecycle candidate verification must route through "
                "identifier_watermark::verify_release: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        else:
            verify_position, verify_arguments = verify_calls[0]
            if not all(
                re.search(pattern, verify_arguments)
                for pattern in (
                    r"\bcandidate\b",
                    r"\bnew_last_identifier\b",
                    r"\bbudget\b",
                )
            ):
                violations.append(
                    "focused Keynote media lifecycle watermark verification must use the "
                    "candidate, released watermark, and shared budget: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
            verify_window = owner[max(0, verify_position - 420) : verify_position + 420]
            if not re.search(
                r"if\s+let\s+Some\s*\(\s*new_last_identifier\s*\)\s*=\s*"
                r"released_watermark",
                verify_window,
            ):
                violations.append(
                    "focused Keynote media lifecycle watermark verification must be gated by "
                    "the source release proof: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )

        identity_watermark_call = re.search(
            r"\bwith_new_last_identifier\s*\(\s*new_last_identifier\s*\)", owner
        )
        if identity_watermark_call is None:
            violations.append(
                "focused Keynote media lifecycle removal must attach the released watermark "
                "to its IdentityBatch: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        else:
            identity_window = owner[
                max(0, identity_watermark_call.start() - 720) : identity_watermark_call.end() + 180
            ]
            if not re.search(r"\bLifecycleAction::Remove\b", identity_window) or not re.search(
                r"\breleased_watermark\b", identity_window
            ):
                violations.append(
                    "focused Keynote media lifecycle released watermark must be attached only "
                    "to the remove IdentityBatch: "
                    f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )

    metadata_code = child_sources_by_name.get("metadata", "")
    for marker in (
        "batch.is_removal_transition()",
        "batch.new_last_identifier()",
        "with_new_last_object_identifier",
    ):
        if marker not in metadata_code:
            violations.append(
                "focused Keynote media lifecycle metadata adapter is missing optional removal "
                f"watermark propagation marker {marker}: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'metadata.rs'}"
            )

    metadata_identifier_code = metadata_code
    identity_start = metadata_identifier_code.find("struct IdentityCollector")
    identity_end = metadata_identifier_code.find(
        "struct NoopIdentityVisitor", identity_start + 1
    )
    identity_code = (
        metadata_identifier_code[identity_start:identity_end]
        if identity_start >= 0 and identity_end > identity_start
        else ""
    )
    for marker in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_METADATA_IDENTIFIER_MARKERS:
        if marker not in metadata_identifier_code:
            violations.append(
                "focused Keynote media lifecycle metadata identifier census is missing "
                f"marker {marker}: "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'metadata.rs'}"
            )
    for label, marker in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_METADATA_IDENTIFIER_CALLBACKS.items():
        if marker.search(identity_code) is None:
            violations.append(
                "focused Keynote media lifecycle metadata identifier census is missing "
                f"{label}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'metadata.rs'}"
            )
    for forbidden in (
        r"\bfn\s+visit_data_reference\b",
        r"\bobserve_identifier\s*\([^)]*\bdata_identifier\b",
        r"\bobserve_identifier\s*\([^)]*\btarget_component_identifier\b",
        r"\bobserve_identifier\s*\([^)]*\bsource\s*\(\s*\)\s*\.\s*identifier\b",
    ):
        if re.search(forbidden, identity_code):
            violations.append(
                "focused Keynote media lifecycle metadata identifier census must observe only "
                f"host object identifiers ({forbidden}): "
                f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_CHILD_ROOT / 'metadata.rs'}"
            )

    def owner_function_body(name: str) -> str:
        declaration = re.search(
            rf"(?m)^\s*(?:pub(?:\([^()]*\))?\s+)?"
            rf"(?:unsafe\s+|async\s+|const\s+)*fn\s+{re.escape(name)}\b",
            owner,
        )
        if declaration is None:
            return ""
        opening = owner.find("{", declaration.end())
        if opening < 0:
            return ""
        depth = 1
        cursor = opening + 1
        while cursor < len(owner) and depth:
            if owner[cursor] == "{":
                depth += 1
            elif owner[cursor] == "}":
                depth -= 1
            cursor += 1
        return owner[opening + 1 : cursor - 1] if depth == 0 else ""

    allocator_body = owner_function_body("allocate_clone_identities")
    allocator_markers = (
        r"\bglobal_max\s*=\s*snapshot\s*\.\s*last_identifier\s*\(\s*\)\s*\.\s*max\s*\(\s*"
        r"snapshot\s*\.\s*metadata_identifier_maximum\s*\(\s*\)\s*\)",
        r"\bpackage\.state\.source\.components\s*\(\s*\)\s*\.\s*iter\s*\(\s*\)",
        r"\bcomponent\.archive\s*\(\s*\)\.objects\b",
        r"\bobject\.archive_info\.identifier\b",
        r"\bglobal_max\s*=\s*global_max\.max\s*\(\s*identifier\s*\)",
        r"\bbudget\.charge_entries\s*\(\s*object_count\s*\)",
    )
    for marker in allocator_markers:
        if re.search(marker, allocator_body) is None:
            violations.append(
                "focused Keynote media lifecycle clone allocation is missing its metadata and "
                f"physical identifier census marker {marker}: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
    census_position = allocator_body.find("budget.charge_entries(object_count)")
    result_vector_position = allocator_body.find("let mut remap")
    if (
        census_position < 0
        or result_vector_position < 0
        or census_position > result_vector_position
    ):
        violations.append(
            "focused Keynote media lifecycle clone allocation must charge its full identifier "
            "census before growing result vectors: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
        )

    budget_names = {
        match.group(1)
        for match in re.finditer(
            r"(?m)^\s*(?:pub\s*\([^)]*\)\s+)?struct\s+"
            r"([A-Za-z_][A-Za-z0-9_]*Budget)\b",
            lifecycle_code,
        )
    }
    if len(budget_names) != 1:
        violations.append(
            "focused Keynote media lifecycle owner must define exactly one private aggregate "
            f"budget: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
        )
    for label, marker in KEYNOTE_SLIDE_MEDIA_LIFECYCLE_BUDGET_MARKERS:
        if marker.search(lifecycle_code) is None:
            violations.append(
                "focused Keynote media lifecycle transaction is missing "
                f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )

    function_declaration = re.compile(
        r"(?<![A-Za-z0-9_#])(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
        r"(?:unsafe[ \t\r\n]+|async[ \t\r\n]+|const[ \t\r\n]+)*"
        r"fn[ \t\r\n]+(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
    )

    def function_records(code: str) -> dict[str, tuple[str, str]]:
        records: dict[str, tuple[str, str]] = {}
        for match in function_declaration.finditer(code):
            opening = code.find("{", match.end())
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
            if depth == 0:
                records[match.group("name")] = (
                    code[match.start() : opening],
                    code[opening + 1 : cursor - 1],
                )
        return records

    functions = function_records(owner)

    def reachable_functions(root_name: str) -> dict[str, tuple[str, str]]:
        """Follow private owner helpers so one root ledger may be delegated."""

        reachable: dict[str, tuple[str, str]] = {}
        pending = [root_name]
        while pending:
            name = pending.pop()
            if name in reachable:
                continue
            record = functions.get(name)
            if record is None:
                continue
            reachable[name] = record
            _signature, body = record
            pending.extend(
                helper
                for helper in functions
                if helper not in reachable
                and re.search(rf"\b(?:r#)?{re.escape(helper)}\s*\(", body)
            )
        return reachable

    for name in sorted(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OPERATIONS):
        signature_body = functions.get(name)
        if signature_body is None:
            continue
        signature, body = signature_body
        reachable = reachable_functions(name)
        reachable_source = "\n".join(
            signature_item + "\n" + body_item
            for signature_item, body_item in reachable.values()
        )
        if not re.search(r"\b(?:budget|ledger)\b", reachable_source):
            violations.append(
                "focused Keynote media lifecycle operation must thread its aggregate budget "
                f"({name}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        if sum(
            len(KEYNOTE_SLIDE_MEDIA_LIFECYCLE_BUDGET_FACTORY_CALL.findall(helper_body))
            for _helper_signature, helper_body in reachable.values()
        ) != 1:
            violations.append(
                "focused Keynote media lifecycle operation must create exactly one aggregate "
                f"budget ({name}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
        for route in (
            "select_media",
            "rewrite_lifecycle",
            "verify_candidate_selection",
            "clone_object",
            "replace_slide_message_with_lifecycle_refs",
        ):
            route_call = re.search(
                rf"\b(?:[A-Za-z_][A-Za-z0-9_]*::)?{route}\s*\((?P<arguments>[^;{{}}]*)\)",
                reachable_source,
            )
            if route_call is not None and not re.search(
                r"\b(?:&\s*mut\s+)?budget\b", route_call.group("arguments")
            ):
                violations.append(
                    "focused Keynote media lifecycle operation must route its shared budget "
                    f"through {route} ({name}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
                )
        if not re.search(
            r"\b(?:select_media|private_graph|clone_object|rewrite_slide|"
            r"rewrite_build|rewrite_build_chunk|replace_slide_message|metadata|"
            r"uuid|UUID)(?:\b|_)",
            reachable_source,
            re.IGNORECASE,
        ) and not re.search(
            r"\b(?:verify_candidate_selection|verify_zip_locality|candidate)(?:\b|_)",
            reachable_source,
            re.IGNORECASE,
        ):
            violations.append(
                "focused Keynote media lifecycle operation must traverse its graph/metadata "
                f"closure ({name}): {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )

    lifecycle_blob = lifecycle_code
    for label, marker in (
        ("selector resolution", re.compile(r"\b(?:SlideSelector|MovieSelector|select_media)\b")),
        (
            "identity remapping",
            re.compile(
                r"\b(?:clone_with_identity_remap|remap_clone_payload_with_budget|"
                r"IdentifierRewrite|UuidRewrite)(?:\b|_)"
            ),
        ),
        ("metadata transition", re.compile(r"\b(?:MetadataSnapshot|IdentityBatch|MediaBatch|package_metadata)\b")),
        ("atomic publication", re.compile(r"\b(?:candidate|reopen|write_to|reassembly|readback|verify)\b")),
        (
            "typed selector kind validation",
            re.compile(
                r"(?s)(?=.*\bKindMismatch\b)(?=.*\bexpected_kind\b)"
                r"(?=.*\bMovieKind\b)(?=.*\bexpected\b)(?=.*\bactual\b)"
            ),
        ),
        (
            "source and target media counts",
            re.compile(
                r"(?s)(?=.*\bcount_slide_media\b)(?=.*\bsource_media_count\b)"
                r"(?=.*\btarget_media_count\b)(?=.*\bactual_count\b)"
            ),
        ),
        ("unsupported source validation", re.compile(r"\bUnsupportedSource\b")),
    ):
        if marker.search(lifecycle_blob) is None:
            violations.append(
                "focused Keynote media lifecycle transaction is missing "
                f"{label} marker: {KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
            )
    count_start = owner.find("fn count_slide_media")
    count_end = owner.find("\nfn unique_payload", count_start + 1) if count_start >= 0 else -1
    count_code = owner[count_start : count_end if count_end >= 0 else len(owner)]
    if not re.search(
        r"(?s)(?=.*\bMOVIE_MESSAGE_TYPE\b)(?=.*\bfind\s*\()"
        r"(?=.*\bcontinue\b)(?=.*\bchecked_add\s*\(\s*1\s*\))",
        count_code,
    ):
        violations.append(
            "focused Keynote media lifecycle count must skip unsupported drawable siblings "
            "without shifting source-order positions: "
            f"{KEYNOTE_SLIDE_MEDIA_LIFECYCLE_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _keynote_slide_drawable_comment_owner_present(root: Path) -> bool:
    """Return whether the selector-first drawable-comment owner is wired."""

    owner_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_EXPORT_SOURCES[0]
    if not owner_path.is_file() or not package_path.is_file():
        return False
    package = _mask_rust_non_code(
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
    )
    return KEYNOTE_SLIDE_DRAWABLE_COMMENT_PACKAGE_MODULE.search(package) is not None


def _keynote_slide_drawable_comment_owner_sources(root: Path) -> list[Path]:
    """Return the public owner and private graph/metadata implementation files."""

    owner = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE
    children = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT
    paths = [owner] if owner.is_file() else []
    if children.is_dir():
        # The metadata owner has a private child directory.  Keep the entire
        # owner closure in the source inventory so a new nested wire helper
        # cannot silently evade the public-leak and marker ratchets.
        paths.extend(sorted(children.rglob("*.rs")))
    return paths


def audit_keynote_slide_drawable_comment_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep direct drawable comments semantic while host CRUD remains pending."""

    if not _keynote_slide_drawable_comment_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_SEMANTIC_SOURCE
    package_path, lib_path = (
        root / path for path in KEYNOTE_SLIDE_DRAWABLE_COMMENT_EXPORT_SOURCES
    )
    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    semantic = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    violations: list[str] = []
    owner_code = _mask_rust_non_code(owner)
    semantic_code = _mask_rust_non_code(semantic)
    package_code = _mask_rust_non_code(package)
    library_code = _mask_rust_non_code(library)
    all_owner_code = "\n".join(
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        for path in _keynote_slide_drawable_comment_owner_sources(root)
        if path.is_file()
    )

    if KEYNOTE_SLIDE_DRAWABLE_COMMENT_PUBLIC_MODULE.search(package + library):
        violations.append(
            "focused Keynote drawable-comment owner modules must remain private: "
            f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_EXPORT_SOURCES[0]}"
        )
    if KEYNOTE_SLIDE_DRAWABLE_COMMENT_PACKAGE_MODULE.search(package) is None:
        violations.append(
            "focused Keynote drawable-comment Package module is missing: "
            f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_EXPORT_SOURCES[0]}"
        )
    for child in ("engine", "graph", "metadata"):
        if re.search(
            rf"(?m)^\s*mod\s+{re.escape(child)}\s*;", owner_code
        ) is None:
            violations.append(
                "focused Keynote drawable-comment owner is missing private child "
                f"{child}: {KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
            )
    if KEYNOTE_SLIDE_DRAWABLE_COMMENT_PUBLIC_CHILD_MODULE.search(owner_code):
        violations.append(
            "focused Keynote drawable-comment child modules must remain private: "
            f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
        )
    for relative in KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_SOURCES:
        if not (root / relative).is_file():
            violations.append(
                "focused Keynote drawable-comment owner is missing child source: "
                f"{relative}"
            )

    metadata_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT / "metadata.rs"
    metadata = (
        _mask_rust_cfg_test_items(metadata_path.read_text(encoding="utf-8"))
        if metadata_path.is_file()
        else ""
    )
    metadata_code = _mask_rust_non_code(metadata)
    if metadata_path.is_file() and (
        KEYNOTE_SLIDE_DRAWABLE_COMMENT_PRIVATE_IDENTITY_MODULE.search(metadata_code)
        is None
    ):
        violations.append(
            "focused Keynote drawable-comment metadata is missing its private identities child: "
            f"{metadata_path.relative_to(root)}"
        )
    if KEYNOTE_SLIDE_DRAWABLE_COMMENT_PUBLIC_IDENTITY_MODULE.search(metadata_code):
        violations.append(
            "focused Keynote drawable-comment metadata identities child must remain private: "
            f"{metadata_path.relative_to(root)}"
        )

    family_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_FAMILY_SOURCE
    family = _keynote_family_view(
        root, family_path, {}, KEYNOTE_SLIDE_DRAWABLE_COMMENT_FAMILY_MODULE
    )
    for name in sorted(KEYNOTE_SLIDE_DRAWABLE_COMMENT_CANONICAL_TYPES):
        for source, path in (
            (owner, owner_path),
            (family, family_path),
        ):
            if name not in _rust_canonical_exports(
                source, KEYNOTE_SLIDE_DRAWABLE_COMMENT_CANONICAL_TYPES
            ):
                violations.append(
                    "focused Keynote drawable-comment public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )
    for name in sorted(KEYNOTE_SLIDE_DRAWABLE_COMMENT_SEMANTIC_TYPES):
        if name not in _rust_canonical_exports(
            semantic + library, KEYNOTE_SLIDE_DRAWABLE_COMMENT_SEMANTIC_TYPES
        ):
            violations.append(
                "focused Keynote drawable-comment semantic API is missing "
                f"{name}: {KEYNOTE_SLIDE_DRAWABLE_COMMENT_SEMANTIC_SOURCE}"
            )
    for name in sorted(KEYNOTE_SLIDE_DRAWABLE_COMMENT_SELECTOR_TYPES):
        if name not in _rust_canonical_exports(
            library, KEYNOTE_SLIDE_DRAWABLE_COMMENT_SELECTOR_TYPES
        ):
            violations.append(
                "focused Keynote drawable-comment public API is missing selector "
                f"{name}: {KEYNOTE_SLIDE_DRAWABLE_COMMENT_EXPORT_SOURCES[1]}"
            )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Package")
    }
    for name in sorted(KEYNOTE_SLIDE_DRAWABLE_COMMENT_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused Keynote drawable-comment Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_drawable_comment":
            selectors = (
                ("SlideSelector",)
                if name == "slide_drawables"
                else ("SlideSelector", "DrawableSelector")
            )
            for selector in selectors:
                if not re.search(rf"\b{re.escape(selector)}\b", declaration):
                    violations.append(
                        "focused Keynote drawable-comment Package method must accept "
                        f"selector-first {selector} ({name}): "
                        f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
                    )
        if name == "apply_slide_drawable_comment" and not re.search(
            r"\bPatch\b", declaration
        ):
            violations.append(
                "focused Keynote drawable-comment apply method must accept its typed patch: "
                f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
            )
        if re.search(r"\b(?:Archive|RawMessage|WireView|WireLimits|u64)\b", declaration):
            if "fingerprint" not in declaration:
                violations.append(
                    "focused Keynote drawable-comment Package method exposes a raw wire value: "
                    f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
                )
        for match in KEYNOTE_SLIDE_DRAWABLE_COMMENT_RAW_ID_PARAMETER.finditer(
            declaration
        ):
            violations.append(
                "focused Keynote drawable-comment Package method exposes raw identifier "
                f"{match.group(0).strip()}: {KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
            )

    edit_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(
            owner, "Edit"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_DRAWABLE_COMMENT_EDIT_METHODS):
        if name not in edit_methods:
            violations.append(
                "focused Keynote drawable-comment edit is missing "
                f"{name}: {KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
            )

    facade_names = (
        KEYNOTE_SLIDE_DRAWABLE_COMMENT_CANONICAL_TYPES
        | KEYNOTE_SLIDE_DRAWABLE_COMMENT_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_DRAWABLE_COMMENT_SELECTOR_TYPES
        | KEYNOTE_SLIDE_DRAWABLE_COMMENT_PACKAGE_METHODS
        | KEYNOTE_SLIDE_DRAWABLE_COMMENT_EDIT_METHODS
    )
    for source, path in (
        (owner, owner_path),
        (semantic, semantic_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            # Private fields are intentionally physical inside the owner. Only
            # inspect a public item header, where a leaked type is observable.
            header = declaration.split("{", 1)[0]
            header_identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(header)
            }
            for identifier in sorted(header_identifiers):
                if identifier in KEYNOTE_SLIDE_DRAWABLE_COMMENT_PHYSICAL_TYPES:
                    violations.append(
                        "focused Keynote drawable-comment public API exposes archive/IWA type "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
                elif identifier in KEYNOTE_SLIDE_DRAWABLE_COMMENT_WIRE_TYPES:
                    violations.append(
                        "focused Keynote drawable-comment public API exposes wire type "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
                else:
                    reason = _iwork_public_leak(identifier)
                    if reason is not None and identifier not in {
                        # The semantic author value carries an opaque public
                        # string identifier, not a native object identity.
                        "public_id",
                        "SourceCatalog",
                        "ExactArtifacts",
                    }:
                        violations.append(
                            "focused Keynote drawable-comment public API exposes "
                            f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                        )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused Keynote drawable-comment public API retains a glob re-export: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    for label, markers in KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_MARKER_GROUPS.items():
        if not any(marker in all_owner_code for marker in markers):
            violations.append(
                "focused Keynote drawable-comment owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_DRAWABLE_COMMENT_OWNER_SOURCE}"
            )

    # Drawable-comment CRUD uses the package-scope graph wrapper.  The
    # lifecycle owner remains strict and same-component; keeping the call
    # site inside the focused read function prevents either policy from being
    # widened by an unrelated helper or import.
    drawable_graph_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_CHILD_ROOT / "graph.rs"
    drawable_graph_source = (
        _mask_rust_cfg_test_items(drawable_graph_path.read_text(encoding="utf-8"))
        if drawable_graph_path.is_file()
        else ""
    )
    drawable_graph_code = _mask_rust_non_code(drawable_graph_source)
    read_graph_body = _rust_any_function_body(drawable_graph_code, "read_comment_graph") or ""
    if not re.search(r"\bplan_comment_graph_cross_component\s*\(", read_graph_body):
        violations.append(
            "focused Keynote drawable-comment read graph must use the package-scope "
            "cross-component comment plan: "
            f"{drawable_graph_path.relative_to(root)}"
        )
    if re.search(r"\bplan_comment_graph\s*\(", read_graph_body):
        violations.append(
            "focused Keynote drawable-comment read graph must not use the strict "
            "same-component comment plan: "
            f"{drawable_graph_path.relative_to(root)}"
        )

    # The package-scope plan retains a physical CommentStorageLocation for
    # every node.  The focused reader must consume that evidence before it
    # decodes or exposes a node, rather than treating a package-global lookup
    # as proof of the selected node's identity.
    read_node_body = _rust_any_function_body(drawable_graph_code, "read_storage_node") or ""
    for label, pattern in (
        (
            "per-node storage-component accessor",
            r"plan\s*\.\s*storage_component\s*\(\s*identifier\s*\)",
        ),
        (
            "per-node component equality check",
            r"component_name\s*!=\s*expected_component",
        ),
        (
            "per-node archive identifier check",
            r"object\.archive_info\.identifier\s*!=\s*Some\s*\(\s*identifier\s*\)",
        ),
    ):
        if not re.search(pattern, read_node_body):
            violations.append(
                "focused Keynote drawable-comment reader is missing "
                f"{label}: {drawable_graph_path.relative_to(root)}"
            )

    # Slide-owned membership remains the selector authority even though a
    # supported drawable may reside in a different archive component.  Keep
    # source order and duplicate rejection local to slide-context resolution,
    # and resolve the selected drawable's actual component at selection time.
    slide_context_body = _rust_any_function_body(drawable_graph_code, "resolve_slide_context") or ""
    for label, pattern in (
        ("slide-owned field", r"SLIDE_OWNED_DRAWABLES_FIELD"),
        ("source-order enumeration", r"enumerate\s*\(\)"),
        ("unique drawable membership", r"seen\s*\.\s*insert\s*\(\s*identifier\s*\)"),
        ("source-order drawable append", r"drawables\s*\.\s*push\s*\(\s*\(\s*position\s*,\s*identifier\s*\)"),
    ):
        if not re.search(pattern, slide_context_body):
            violations.append(
                "focused Keynote drawable-comment selector is missing "
                f"{label} proof: {drawable_graph_path.relative_to(root)}"
            )
    select_body = _rust_any_function_body(drawable_graph_code, "select_drawable") or ""
    for label, pattern in (
        (
            "selected drawable component lookup",
            r"object_with_component\s*\(\s*resolved\.identifier\s*\)",
        ),
        ("selected component storage", r"component_name\s*=\s*copy_arc\s*\("),
    ):
        if not re.search(pattern, select_body):
            violations.append(
                "focused Keynote drawable-comment selector is missing "
                f"{label}: {drawable_graph_path.relative_to(root)}"
            )

    # Direct-user counts inspect every archive component and the complete
    # ArchiveInfo census.  This preserves copy-on-write when an opaque or
    # header-only owner is discovered alongside the selected route.
    direct_users_body = _rust_any_function_body(drawable_graph_code, "global_direct_users") or ""
    for label, pattern in (
        ("complete header census", r"census_core_header_root_reference\s*\("),
        ("all-component scan", r"package\.state\.source\.components\s*\(\s*\)"),
        ("distinct owner count", r"users\s*=\s*users\.checked_add\s*\("),
        ("self-owner exclusion", r"object_identifier\s*!=\s*root_identifier"),
        ("zero-owner rejection", r"if\s+users\s*==\s*0"),
    ):
        if not re.search(pattern, direct_users_body):
            violations.append(
                "focused Keynote drawable-comment graph is missing "
                f"{label}: {drawable_graph_path.relative_to(root)}"
            )

    # This is a source-oracle receipt only. Candidate CRUD/reply acceptance
    # remains a separate native gate and must not be implied by this check.
    native_test_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_TEST_SOURCE
    native_fixture_path = root / KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_FIXTURE
    if not native_fixture_path.is_file():
        violations.append(
            "focused Keynote drawable-comment owner is missing native source fixture: "
            f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_FIXTURE}"
        )
    elif native_fixture_path.stat().st_size == 0:
        violations.append(
            "focused Keynote drawable-comment native source fixture is empty: "
            f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_FIXTURE}"
        )
    if not native_test_path.is_file():
        violations.append(
            "focused Keynote drawable-comment owner is missing native source test: "
            f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_TEST_SOURCE}"
        )
    else:
        native_test = native_test_path.read_text(encoding="utf-8")
        if re.search(
            rf"\binclude_bytes!\s*\([^)]*"
            rf"{re.escape(KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_FIXTURE.as_posix())}",
            native_test,
        ) is None:
            violations.append(
                "focused Keynote drawable-comment source test must include its native fixture: "
                f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_TEST_SOURCE}"
            )
        if re.search(
            rf"(?m)^\s*#\s*\[\s*test\s*\]\s*\n\s*fn\s+"
            rf"{re.escape(KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_TEST_NAME)}\b",
            _mask_rust_non_code(native_test),
        ) is None:
            violations.append(
                "focused Keynote drawable-comment source test is missing its fixture pin: "
                f"{KEYNOTE_SLIDE_DRAWABLE_COMMENT_NATIVE_TEST_SOURCE}"
            )
    return sorted(set(violations))


def _keynote_movie_geometry_owner_present(root: Path) -> bool:
    """Return whether the movie-geometry owner has crossed its activation seam."""

    owner_path = root / KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE
    codec_path = root / KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE
    package_path = root / KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and codec_path.is_file()
        and re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_movie_geometry\s*;", package_source)
        is not None
    )


def audit_keynote_movie_geometry_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free movie-geometry boundary."""

    if not _keynote_movie_geometry_owner_present(root):
        return []

    owner_path = root / KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE
    semantic_path = root / KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_SOURCE
    package_path = root / KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[1]
    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    semantic = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    violations: list[str] = []
    owner_code = _mask_rust_non_code(owner)
    semantic_code = _mask_rust_non_code(semantic)
    package_code = _mask_rust_non_code(package)
    library_code = _mask_rust_non_code(library)
    package_library_code = _mask_rust_non_code(package + library)

    if re.search(r"(?m)^pub[ \t]+mod[ \t]+slide_movie_geometry\b", package_library_code):
        violations.append(
            "focused litchi-keynote movie-geometry owner module must remain private: "
            f"{KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[0]}"
        )
    if re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_movie_geometry\s*;", package_code) is None:
        violations.append(
            "focused litchi-keynote movie-geometry owner module is missing: "
            f"{KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[0]}"
        )

    family_path = root / KEYNOTE_MOVIE_GEOMETRY_FAMILY_SOURCE
    family = _keynote_family_view(root, family_path, {}, KEYNOTE_MOVIE_GEOMETRY_FAMILY_MODULE)
    for name in sorted(KEYNOTE_MOVIE_GEOMETRY_CANONICAL_TYPES):
        for source, path in ((owner, owner_path), (family, family_path)):
            if name not in _rust_canonical_exports(source, KEYNOTE_MOVIE_GEOMETRY_CANONICAL_TYPES):
                violations.append(
                    "focused litchi-keynote movie-geometry public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )
    for name in sorted(KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_TYPES):
        if name not in _rust_canonical_exports(semantic + library, KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_TYPES):
            violations.append(
                "focused litchi-keynote movie-geometry semantic API is missing "
                f"{name}: {KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_SOURCE}"
            )
    if re.search(r"(?m)^\s*pub\s+mod\s+geometry\s*\{", semantic_code) is None:
        violations.append(
            "focused litchi-keynote movie-geometry semantic API is missing geometry module: "
            f"{KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_SOURCE}"
        )
    # ADR 0034: the transaction family moved from
    # `slide::media::geometry::transaction` to `slide::movie::geometry`.
    if not family.strip():
        violations.append(
            "focused litchi-keynote movie-geometry semantic API is missing public transaction module: "
            f"{KEYNOTE_MOVIE_GEOMETRY_FAMILY_SOURCE}"
        )
    for name in sorted(KEYNOTE_MOVIE_GEOMETRY_SELECTOR_TYPES):
        if name not in _rust_canonical_exports(semantic + library, KEYNOTE_MOVIE_GEOMETRY_SELECTOR_TYPES):
            violations.append(
                "focused litchi-keynote movie-geometry public API is missing selector "
                f"{name}: {KEYNOTE_MOVIE_GEOMETRY_EXPORT_SOURCES[1]}"
            )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Package")
    }
    for name in sorted(KEYNOTE_MOVIE_GEOMETRY_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote movie-geometry Package method is missing "
                f"{name}: {KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_movie_geometry":
            for selector in ("SlideSelector", "MovieSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote movie-geometry Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
                    )

    edit_impl = re.search(
        r"(?<![A-Za-z0-9_#])impl(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]+"
        r"(?:'[^ ]+[ \t\r\n]+)?Edit\b",
        owner_code,
    )
    edit_body = owner_code if edit_impl is None else owner_code[edit_impl.end() :]
    for name in sorted(KEYNOTE_MOVIE_GEOMETRY_EDIT_METHODS):
        if not re.search(rf"\bpub[ \t]+fn[ \t]+{re.escape(name)}\b", edit_body):
            violations.append(
                "focused litchi-keynote movie-geometry edit is missing "
                f"{name}: {KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
            )

    facade_names = (
        KEYNOTE_MOVIE_GEOMETRY_CANONICAL_TYPES
        | KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_TYPES
        | KEYNOTE_MOVIE_GEOMETRY_SELECTOR_TYPES
        | KEYNOTE_MOVIE_GEOMETRY_PACKAGE_METHODS
        | KEYNOTE_MOVIE_GEOMETRY_FLAT_ALIASES
        | {"Edit"}
    )
    for source, source_path in (
        (owner, owner_path),
        (semantic, semantic_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = source_path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)}
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_MOVIE_GEOMETRY_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_MOVIE_GEOMETRY_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_MOVIE_GEOMETRY_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote movie-geometry public API exposes "
                        f"{reason} {identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_MOVIE_GEOMETRY_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote movie-geometry public API retains flat alias "
                        f"{identifier}: {source_path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote movie-geometry public API exposes raw byte slice "
                    f"{byte_slice}: {source_path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_MOVIE_GEOMETRY_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote movie-geometry public API exposes raw identifier "
                    f"{match.group(0).strip()}: {source_path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote movie-geometry public API retains a glob re-export: "
                    f"{source_path.relative_to(root)}:{line_number}"
                )

    codec_path = root / KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE
    codec = _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
    codec_code = _mask_rust_non_code(codec)
    for function in sorted(KEYNOTE_MOVIE_GEOMETRY_CODEC_FUNCTIONS):
        if re.search(rf"\b(?:pub[ \t]+)?fn[ \t]+{re.escape(function)}\b", codec_code) is None:
            violations.append(
                "focused litchi-keynote movie-geometry hidden codec is missing strict API "
                f"{function}: {KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE}"
            )
    for type_name in sorted(KEYNOTE_MOVIE_GEOMETRY_CODEC_TYPES):
        if re.search(
            rf"\b(?:pub[ \t]+)?(?:struct|enum|type)[ \t]+{re.escape(type_name)}\b",
            codec_code,
        ) is None:
            violations.append(
                "focused litchi-keynote movie-geometry hidden codec is missing strict type "
                f"{type_name}: {KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_MOVIE_GEOMETRY_CODEC_MARKER_GROUPS.items():
        if not any(marker in codec for marker in markers):
            violations.append(
                "focused litchi-keynote movie-geometry hidden codec is missing "
                f"{label} marker: {KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE}"
            )
    codec_lib_path = root / KEYNOTE_MOVIE_GEOMETRY_CODEC_PUBLIC_SOURCE
    codec_lib = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(KEYNOTE_MOVIE_GEOMETRY_CODEC_MODULE)}\b",
        codec_lib,
    ) is None:
        violations.append(
            "focused litchi-keynote movie-geometry public API is missing hidden codec module "
            f"{KEYNOTE_MOVIE_GEOMETRY_CODEC_MODULE}: {KEYNOTE_MOVIE_GEOMETRY_CODEC_PUBLIC_SOURCE}"
        )

    for fuzz_path in KEYNOTE_MOVIE_GEOMETRY_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote movie-geometry boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote movie-geometry fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in KEYNOTE_MOVIE_GEOMETRY_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-keynote movie-geometry boundary is missing fuzz corpus: "
                f"{corpus}"
            )
    return sorted(set(violations))


def audit_keynote_movie_geometry_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require resource, candidate, locality, and preview invalidation markers."""

    if not _keynote_movie_geometry_owner_present(root):
        return []
    owner_path = root / KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    violations: list[str] = []
    for label, markers in KEYNOTE_MOVIE_GEOMETRY_PACKAGE_MARKER_GROUPS.items():
        if not any(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote movie-geometry owner is missing "
                f"{label} marker: {KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
            )
    if "keynote_movie_geometry_codec" not in owner:
        violations.append(
            "focused litchi-keynote movie-geometry owner must route through hidden "
            f"keynote_movie_geometry_codec: {KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _audit_keynote_movie_geometry_native_evidence(root: Path) -> list[str]:
    """Require native source and accepted candidate geometry receipts."""

    test_path = root / KEYNOTE_MOVIE_GEOMETRY_NATIVE_TEST_SOURCE
    if not test_path.is_file():
        return [
            "focused Keynote movie-geometry completion is missing its native "
            f"integration test: {KEYNOTE_MOVIE_GEOMETRY_NATIVE_TEST_SOURCE}"
        ]

    raw_test = test_path.read_text(encoding="utf-8")
    code_test = _mask_rust_non_code(raw_test)
    violations: list[str] = []
    for fixture, test_name in KEYNOTE_MOVIE_GEOMETRY_NATIVE_EVIDENCE:
        fixture_path = root / fixture
        if not fixture_path.is_file():
            violations.append(
                "focused Keynote movie-geometry completion is missing native "
                f"fixture: {fixture}"
            )
        elif fixture_path.stat().st_size == 0:
            violations.append(
                "focused Keynote movie-geometry native fixture is empty: "
                f"{fixture}"
            )

        if re.search(
            rf"\binclude_bytes!\s*\([^)]*{re.escape(fixture.as_posix())}",
            raw_test,
        ) is None:
            violations.append(
                "focused Keynote movie-geometry native integration test must "
                f"include fixture {fixture}: {KEYNOTE_MOVIE_GEOMETRY_NATIVE_TEST_SOURCE}"
            )
        if re.search(
            rf"(?m)^\s*#\s*\[\s*test\s*\]\s*\n\s*fn\s+"
            rf"{re.escape(test_name)}\b",
            code_test,
        ) is None:
            violations.append(
                "focused Keynote movie-geometry native integration test is "
                f"missing {test_name}: {KEYNOTE_MOVIE_GEOMETRY_NATIVE_TEST_SOURCE}"
            )

    # Keep the acceptance test tied to the expanded typed operation.  The
    # source-order index catches a regression that silently filters audio or
    # other MovieArchive siblings before resolving a file movie.
    required_markers = {
        "typed source-order selector": r"\bMovieSelector::index\s*\(\s*4\s*\)",
        "typed original-size restore": r"\.restore_original_size\s*\(",
        "typed reflection operation": r"\.flip\s*\(",
        "exact inverse": r"\.inverse\s*\(\)",
    }
    for label, marker in required_markers.items():
        if re.search(marker, code_test) is None:
            violations.append(
                "focused Keynote movie-geometry native integration test is missing "
                f"{label}: {KEYNOTE_MOVIE_GEOMETRY_NATIVE_TEST_SOURCE}"
            )
    return sorted(set(violations))


def audit_keynote_movie_geometry_completion_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Retire the last raw movie-geometry compatibility seam when activated.

    The geometry owner deliberately owns only position and displayed size.  The
    compatibility host therefore still owns native angle, flip, and flags.  The
    expanded owner uses an explicit ``MOVIE_GEOMETRY_COMPLETE``/capability
    token as a handoff: until the semantic value, strict codec, and host bridge
    all advertise the expanded surface this audit is dormant.  Once activated,
    every remaining raw declaration, identifier fallback, and native rewrite
    helper is a violation; the old compatibility-source exemption does not
    apply to the completed owner.
    """

    owner_path = root / KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE
    codec_path = root / KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE
    semantic_path = root / KEYNOTE_MOVIE_GEOMETRY_SEMANTIC_SOURCE
    if not (
        _keynote_movie_geometry_owner_present(root)
        and owner_path.is_file()
        and codec_path.is_file()
    ):
        return []

    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    if KEYNOTE_MOVIE_GEOMETRY_COMPLETE_ACTIVATION.search(owner) is None:
        return []

    codec = _mask_rust_non_code(
        _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
    )
    semantic = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        )
        if semantic_path.is_file()
        else ""
    )
    violations: list[str] = _audit_keynote_movie_geometry_native_evidence(root)
    feature_source = f"{semantic}\n{owner}\n{codec}"

    for label, marker in KEYNOTE_MOVIE_GEOMETRY_COMPLETE_CAPABILITY_MARKERS.items():
        if marker.search(feature_source) is None:
            violations.append(
                "focused litchi-keynote movie-geometry completion is missing "
                f"{label} semantic/codec capability marker: "
                f"{KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
            )

    for label, marker in KEYNOTE_MOVIE_GEOMETRY_COMPLETE_CODEC_CAPABILITY_MARKERS.items():
        if marker.search(codec) is None:
            violations.append(
                "focused litchi-keynote movie-geometry completion is missing "
                f"{label}: {KEYNOTE_MOVIE_GEOMETRY_CODEC_SOURCE}"
            )

    for label, marker in KEYNOTE_MOVIE_GEOMETRY_COMPLETE_OWNER_CAPABILITY_MARKERS.items():
        if marker.search(owner) is None:
            violations.append(
                "focused litchi-keynote movie-geometry completion is missing "
                f"{label} owner marker: {KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
            )

    for label, marker in KEYNOTE_MOVIE_GEOMETRY_COMPLETE_CODEC_FLOW_MARKERS.items():
        if marker.search(owner) is None:
            violations.append(
                "focused litchi-keynote movie-geometry completion is missing "
                f"{label} owner marker: {KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}"
            )

    for match in KEYNOTE_MOVIE_GEOMETRY_COMPLETE_DIRECT_REWRITE_CALL.finditer(owner):
        line_start = owner.rfind("\n", 0, match.start()) + 1
        line_end = owner.find("\n", match.end())
        if line_end < 0:
            line_end = len(owner)
        line = owner[line_start:line_end]
        if re.search(r"\bfn[ \t]+rewrite_movie_(?:geometry|transform)\b", line):
            continue
        line_number = owner.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-keynote movie-geometry completion must execute the "
            f"prepared codec, not direct {match.group(0).strip()}: "
            f"{KEYNOTE_MOVIE_GEOMETRY_OWNER_SOURCE}:{line_number}"
        )

    return sorted(set(violations))


def _keynote_audio_position_owner_present(root: Path) -> bool:
    """Return whether the focused audio-position owner crossed its wiring seam."""

    owner_path = root / KEYNOTE_AUDIO_POSITION_OWNER_SOURCE
    codec_path = root / KEYNOTE_AUDIO_POSITION_CODEC_SOURCE
    package_path = root / KEYNOTE_AUDIO_POSITION_EXPORT_SOURCES[0]
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
        and KEYNOTE_AUDIO_POSITION_OWNER_MODULE.search(package_source) is not None
    )


def _keynote_audio_position_marker_present(
    source: str, marker: str | re.Pattern[str] | tuple[str, ...]
) -> bool:
    """Match one structural owner/codec marker or an accepted alternative."""

    if isinstance(marker, tuple):
        return any(_keynote_audio_position_marker_present(source, item) for item in marker)
    if isinstance(marker, re.Pattern):
        return marker.search(source) is not None
    return marker in source


def audit_keynote_audio_position_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free audio-position facade."""

    if not _keynote_audio_position_owner_present(root):
        return []

    owner_path = root / KEYNOTE_AUDIO_POSITION_OWNER_SOURCE
    semantic_path = root / KEYNOTE_AUDIO_POSITION_SEMANTIC_SOURCE
    package_path, lib_path = (
        root / path for path in KEYNOTE_AUDIO_POSITION_EXPORT_SOURCES
    )
    owner = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    semantic = (
        _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        if semantic_path.is_file()
        else ""
    )
    package = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    library = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    owner_code = _mask_rust_non_code(owner)
    semantic_code = _mask_rust_non_code(semantic)
    package_code = _mask_rust_non_code(package)
    library_code = _mask_rust_non_code(library)
    violations: list[str] = []

    if KEYNOTE_AUDIO_POSITION_PUBLIC_MODULE.search(package + library):
        violations.append(
            "focused litchi-keynote audio-position owner module must remain private: "
            f"{KEYNOTE_AUDIO_POSITION_EXPORT_SOURCES[0]}"
        )
    if KEYNOTE_AUDIO_POSITION_OWNER_MODULE.search(package) is None:
        violations.append(
            "focused litchi-keynote audio-position owner module is missing: "
            f"{KEYNOTE_AUDIO_POSITION_EXPORT_SOURCES[0]}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Package")
    }
    for name in sorted(KEYNOTE_AUDIO_POSITION_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote audio-position Package method is missing "
                f"{name}: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_audio_position":
            for selector in sorted(KEYNOTE_AUDIO_POSITION_SELECTOR_TYPES):
                if not re.search(rf"\b{re.escape(selector)}\b", declaration):
                    violations.append(
                        "focused litchi-keynote audio-position Package method must accept "
                        f"selector-first {selector} ({name}): {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
                    )
            if name == "slide_audio_position" and "Point" not in declaration:
                violations.append(
                    "focused litchi-keynote audio-position Package method must use the "
                    f"semantic Point value ({name}): {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
                )
        elif "Patch" not in declaration:
            violations.append(
                "focused litchi-keynote audio-position apply method must accept "
                f"Patch: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
            )
        if re.search(r"\b(?:u64|u32|usize)\b|&\s*\[\s*u8\s*\]|Vec\s*<\s*u8\s*>", declaration):
            violations.append(
                "focused litchi-keynote audio-position Package method exposes a raw value: "
                f"{KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
            )
        for match in KEYNOTE_AUDIO_POSITION_RAW_PARAMETER.finditer(declaration):
            violations.append(
                "focused litchi-keynote audio-position Package method exposes raw identifier "
                f"{match.group(0).strip()}: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
            )

    for name in sorted(KEYNOTE_AUDIO_POSITION_EDIT_METHODS):
        if not re.search(
            rf"\bpub[ \t]+fn[ \t]+{re.escape(name)}\b", owner_code
        ):
            violations.append(
                "focused litchi-keynote audio-position edit is missing "
                f"{name}: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
            )

    for path, source in (
        (owner_path, owner),
        (semantic_path, semantic),
        (package_path, package),
        (lib_path, library),
    ):
        if not source:
            continue
        dedicated = path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not (
                identifiers
                & (
                    (KEYNOTE_AUDIO_POSITION_CANONICAL_TYPES - KEYNOTE_FAMILY_GENERIC_NAMES)
                    | KEYNOTE_AUDIO_POSITION_SELECTOR_TYPES
                    | KEYNOTE_AUDIO_POSITION_PACKAGE_METHODS
                    | {"Point"}
                )
            ):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_AUDIO_POSITION_WIRE_TYPES:
                    reason = "wire type"
                elif identifier in KEYNOTE_AUDIO_POSITION_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote audio-position public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote audio-position public API retains a glob "
                    f"re-export: {path.relative_to(root)}:{line_number}"
                )
            if RUST_BYTE_SLICE.search(declaration):
                violations.append(
                    "focused litchi-keynote audio-position public API exposes raw bytes: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_AUDIO_POSITION_RAW_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote audio-position public API exposes raw identifier "
                    f"{match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )

    for path, source, names in (
        (owner_path, owner, KEYNOTE_AUDIO_POSITION_CANONICAL_TYPES),
        (
            root / KEYNOTE_AUDIO_POSITION_FAMILY_SOURCE,
            _keynote_family_view(
                root,
                root / KEYNOTE_AUDIO_POSITION_FAMILY_SOURCE,
                {},
                KEYNOTE_AUDIO_POSITION_FAMILY_MODULE,
            ),
            KEYNOTE_AUDIO_POSITION_CANONICAL_TYPES,
        ),
    ):
        exported = _rust_canonical_exports(source, names)
        for name in sorted(names - exported):
            violations.append(
                "focused litchi-keynote audio-position public API is missing canonical "
                f"type {name}: {path.relative_to(root)}"
            )
    if not re.search(r"\bpub[ \t]+struct[ \t]+Point\b", semantic_code):
        violations.append(
            "focused litchi-keynote audio-position semantic API is missing Point: "
            f"{KEYNOTE_AUDIO_POSITION_SEMANTIC_SOURCE}"
        )
    if not re.search(r"\b(?:source_fingerprint|ExactArtifacts|inverse|PatchConflict)\b", owner_code):
        violations.append(
            "focused litchi-keynote audio-position facade is missing exact-source patch/inverse "
            f"markers: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def audit_keynote_audio_position_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the position edge to use the existing lazy geometry codec seam."""

    if not _keynote_audio_position_owner_present(root):
        return []
    codec_path = root / KEYNOTE_AUDIO_POSITION_CODEC_SOURCE
    codec_raw = codec_path.read_text(encoding="utf-8")
    codec = _mask_rust_cfg_test_items(codec_raw)
    code = _mask_rust_non_code(codec)
    violations: list[str] = []
    for function in sorted(KEYNOTE_AUDIO_POSITION_CODEC_FUNCTIONS):
        if re.search(rf"\b(?:pub[ \t]+)?fn[ \t]+{re.escape(function)}\b", code) is None:
            violations.append(
                "focused Keynote audio-position codec is missing strict API "
                f"{function}: {KEYNOTE_AUDIO_POSITION_CODEC_SOURCE}"
            )
    for type_name in sorted(KEYNOTE_AUDIO_POSITION_CODEC_TYPES):
        if re.search(
            rf"\b(?:pub[ \t]+)?(?:struct|enum|type)[ \t]+{re.escape(type_name)}\b",
            code,
        ) is None:
            violations.append(
                "focused Keynote audio-position codec is missing strict type "
                f"{type_name}: {KEYNOTE_AUDIO_POSITION_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_AUDIO_POSITION_CODEC_MARKER_GROUPS.items():
        if not all(_keynote_audio_position_marker_present(code, marker) for marker in markers):
            violations.append(
                "focused Keynote audio-position codec is missing "
                f"{label}: {KEYNOTE_AUDIO_POSITION_CODEC_SOURCE}"
            )
    for pattern in (
        re.compile(r"\bprost(?:_types)?\b"),
        re.compile(r"\bMessage\s*::\s*(?:decode|encode)\s*\("),
        re.compile(r"\b(?:to_owned_message|try_encode_to_vec|encode_to_vec)\s*\("),
    ):
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused Keynote audio-position codec must remain lazy/source-preserving; "
                f"eager generated operation {match.group(0).strip()}: "
                f"{KEYNOTE_AUDIO_POSITION_CODEC_SOURCE}:{line_number}"
            )

    public_path = root / KEYNOTE_AUDIO_POSITION_CODEC_PUBLIC_SOURCE
    public_source = (
        _mask_rust_cfg_test_items(public_path.read_text(encoding="utf-8"))
        if public_path.is_file()
        else ""
    )
    if re.search(
        rf"\bpub\s+mod\s+{re.escape(KEYNOTE_AUDIO_POSITION_CODEC_MODULE)}\b",
        public_source,
    ) is None:
        violations.append(
            "focused Keynote audio-position codec is missing its hidden geometry codec module: "
            f"{KEYNOTE_AUDIO_POSITION_CODEC_PUBLIC_SOURCE}"
        )
    return sorted(set(violations))


def audit_keynote_audio_position_completion_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require focused integration and adversarial fuzz admission coverage."""

    if not _keynote_audio_position_owner_present(root):
        return []
    violations: list[str] = []
    test_path = root / KEYNOTE_AUDIO_POSITION_TEST_SOURCE
    if not test_path.is_file():
        violations.append(
            "focused Keynote audio-position boundary is missing integration tests: "
            f"{KEYNOTE_AUDIO_POSITION_TEST_SOURCE}"
        )
    else:
        test_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(test_path.read_text(encoding="utf-8"))
        )
        for marker in KEYNOTE_AUDIO_POSITION_TEST_MARKERS:
            if marker not in test_code and marker not in test_path.read_text(encoding="utf-8"):
                violations.append(
                    "focused Keynote audio-position integration tests are missing "
                    f"{marker}: {KEYNOTE_AUDIO_POSITION_TEST_SOURCE}"
                )

    fuzz_path = root / KEYNOTE_AUDIO_POSITION_FUZZ_SOURCE
    if not fuzz_path.is_file():
        violations.append(
            "focused Keynote audio-position boundary is missing fuzz target: "
            f"{KEYNOTE_AUDIO_POSITION_FUZZ_SOURCE}"
        )
    elif "fuzz_target!" not in fuzz_path.read_text(encoding="utf-8"):
        violations.append(
            "focused Keynote audio-position fuzz target is missing fuzz_target! harness: "
            f"{KEYNOTE_AUDIO_POSITION_FUZZ_SOURCE}"
        )
    if not (root / KEYNOTE_AUDIO_POSITION_FUZZ_CORPUS).is_dir():
        violations.append(
            "focused Keynote audio-position boundary is missing fuzz corpus: "
            f"{KEYNOTE_AUDIO_POSITION_FUZZ_CORPUS}"
        )
    return sorted(set(violations))


def audit_keynote_audio_position_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one bounded, source-bound transaction for audio-position edits."""

    if not _keynote_audio_position_owner_present(root):
        return []
    owner_path = root / KEYNOTE_AUDIO_POSITION_OWNER_SOURCE
    codec_path = root / KEYNOTE_AUDIO_POSITION_CODEC_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    codec = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        )
        if codec_path.is_file()
        else ""
    )
    prepared = owner + "\n" + codec
    violations: list[str] = []
    budget_declarations = re.findall(
        r"(?m)^\s*(?:pub(?:\([^()]*\))?\s+)?struct\s+"
        r"[A-Za-z_][A-Za-z0-9_]*Budget\b",
        owner,
    )
    shared_budget_import = re.search(
        r"\b(?:use\s+[^;\n]*::\s*)?GeometryBudget\b", owner
    ) is not None
    if len(budget_declarations) > 1 or (
        not budget_declarations and not shared_budget_import
    ):
        violations.append(
            "focused litchi-keynote audio-position owner must use exactly one aggregate "
            "transaction budget (local or shared GeometryBudget): "
            f"{KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
        )
    for label, markers in KEYNOTE_AUDIO_POSITION_OWNER_MARKER_GROUPS.items():
        if not all(_keynote_audio_position_marker_present(prepared, marker) for marker in markers):
            violations.append(
                "focused litchi-keynote audio-position owner is missing "
                f"{label}: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
            )
    if not (
        "decode_movie_position_with_report" in owner
        or "select_audio_with_budget" in owner
    ):
        violations.append(
            "focused litchi-keynote audio-position owner must route position reads through "
            f"the strict codec/selection helper: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
        )
    if "prepare_movie_position_rewrite" not in owner:
        violations.append(
            "focused litchi-keynote audio-position owner must route writes through a prepared "
            f"codec rewrite: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
        )
    if len(re.findall(r"\.[ \t\r\n]*execute\s*\(", owner)) != 1:
        violations.append(
            "focused litchi-keynote audio-position owner must execute exactly one prepared "
            f"rewrite: {KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
        )
    if re.search(r"\bsaturating_(?:add|sub|mul)\b", owner):
        violations.append(
            "focused litchi-keynote audio-position owner must use checked resource accounting: "
            f"{KEYNOTE_AUDIO_POSITION_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _keynote_media_properties_owner_present(root: Path) -> bool:
    """Return whether the focused media-properties owner crossed its seam."""

    owner_path = root / KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE
    codec_path = root / KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE
    package_path = root / KEYNOTE_MEDIA_PROPERTIES_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return (
        owner_path.is_file()
        and codec_path.is_file()
        and KEYNOTE_MEDIA_PROPERTIES_OWNER_MODULE.search(package_source) is not None
    )


def audit_keynote_media_properties_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the archive-free selector-first media-properties facade.

    This audit deliberately stays dormant until the focused owner, private
    Package module, and neutral codec exist together.  A partially copied
    source file must not turn the still-supported compatibility host into a
    migration failure, while a claimed handoff receives the same strict
    public-surface checks as the preceding Keynote verticals.
    """

    if not _keynote_media_properties_owner_present(root):
        return []

    owner_path = root / KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE
    semantic_path = root / KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_SOURCE
    package_path, lib_path = (
        root / path for path in KEYNOTE_MEDIA_PROPERTIES_EXPORT_SOURCES
    )
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in (owner_path, semantic_path, package_path, lib_path)
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    owner = sources[owner_path]
    owner_code = code[owner_path]
    package = sources[package_path]
    package_code = code[package_path]
    library = sources[lib_path]
    semantic = sources[semantic_path]
    semantic_code = code[semantic_path]
    violations: list[str] = []

    if KEYNOTE_MEDIA_PROPERTIES_PUBLIC_MODULE.search(package_code + code[lib_path]):
        violations.append(
            "focused litchi-keynote media-properties owner module must remain private: "
            f"{KEYNOTE_MEDIA_PROPERTIES_EXPORT_SOURCES[0]}"
        )
    if KEYNOTE_MEDIA_PROPERTIES_OWNER_MODULE.search(package_code) is None:
        violations.append(
            "focused litchi-keynote media-properties owner module is missing: "
            f"{KEYNOTE_MEDIA_PROPERTIES_EXPORT_SOURCES[0]}"
        )

    family_path = root / KEYNOTE_MEDIA_PROPERTIES_FAMILY_SOURCE
    family = _keynote_family_view(
        root, family_path, {}, KEYNOTE_MEDIA_PROPERTIES_FAMILY_MODULE
    )
    for name in sorted(KEYNOTE_MEDIA_PROPERTIES_CANONICAL_TYPES):
        for source, path in (
            (owner, owner_path),
            (family, family_path),
        ):
            if name not in _rust_canonical_exports(
                source, KEYNOTE_MEDIA_PROPERTIES_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote media-properties public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        semantic, KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_TYPES
    )
    if not semantic_exports:
        violations.append(
            "focused litchi-keynote media-properties semantic API is missing MediaProperties: "
            f"{KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_SOURCE}"
        )
    media_properties_struct = _rust_named_struct_body(semantic, "MediaProperties")
    if media_properties_struct is None:
        violations.append(
            "focused litchi-keynote media-properties semantic value must be a private-field "
            f"struct: {KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_SOURCE}"
        )
    else:
        media_properties_body, _offset = media_properties_struct
        for field_name in (
            "hyperlink_url",
            "locked",
            "aspect_ratio_locked",
            "accessibility_description",
        ):
            if re.search(rf"\b{re.escape(field_name)}\b", media_properties_body) is None:
                violations.append(
                    "focused litchi-keynote media-properties semantic value is missing field "
                    f"{field_name}: {KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_SOURCE}"
                )
        if not re.search(r"\bOption\s*<\s*String\s*>", media_properties_body):
            violations.append(
                "focused litchi-keynote media-properties semantic value must retain optional "
                f"string presence: {KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_SOURCE}"
            )
        if not re.search(r"\bOption\s*<\s*bool\s*>", media_properties_body):
            violations.append(
                "focused litchi-keynote media-properties semantic value must retain optional "
                f"boolean presence: {KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_SOURCE}"
            )

    selector_exports = _rust_canonical_exports(
        library, KEYNOTE_MEDIA_PROPERTIES_SELECTOR_TYPES
    )
    for selector in sorted(KEYNOTE_MEDIA_PROPERTIES_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote media-properties public API is missing selector "
            f"{selector}: {KEYNOTE_MEDIA_PROPERTIES_EXPORT_SOURCES[1]}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(owner, "Package")
    }
    for name in sorted(KEYNOTE_MEDIA_PROPERTIES_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote media-properties Package method is missing "
                f"{name}: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_media_properties":
            for selector in sorted(KEYNOTE_MEDIA_PROPERTIES_SELECTOR_TYPES):
                if not re.search(rf"\b{re.escape(selector)}\b", declaration):
                    violations.append(
                        "focused litchi-keynote media-properties Package method must accept "
                        f"selector-first {selector} ({name}): "
                        f"{KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
                    )
        if re.search(r"\b(?:u64|u32|usize)\b|&\s*\[\s*u8\s*\]", declaration):
            violations.append(
                "focused litchi-keynote media-properties Package method exposes a raw wire "
                f"value: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
            )
        for match in KEYNOTE_MEDIA_PROPERTIES_RAW_PARAMETER.finditer(declaration):
            violations.append(
                "focused litchi-keynote media-properties Package method exposes raw identifier "
                f"{match.group(0).strip()}: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
            )
    apply_declaration = owner_methods.get("apply_slide_media_properties")
    if apply_declaration is not None and not re.search(
        r"\b(?:Patch|MediaPropertiesPatch)\b", apply_declaration
    ):
        violations.append(
            "focused litchi-keynote media-properties apply method must accept its typed patch: "
            f"{KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )

    edit_methods = {
        name
        for name, _declaration, _line in _rust_public_methods_in_impl(
            owner, "Edit"
        )
    }
    for name in sorted(KEYNOTE_MEDIA_PROPERTIES_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote media-properties edit is missing "
            f"{name}: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_MEDIA_PROPERTIES_CANONICAL_TYPES
        | KEYNOTE_MEDIA_PROPERTIES_SEMANTIC_TYPES
        | KEYNOTE_MEDIA_PROPERTIES_SELECTOR_TYPES
        | KEYNOTE_MEDIA_PROPERTIES_PACKAGE_METHODS
    )
    for source, path in (
        (owner, owner_path),
        (semantic, semantic_path),
        (package, package_path),
        (library, lib_path),
    ):
        if not source:
            continue
        dedicated = path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not (identifiers & (facade_names - KEYNOTE_FAMILY_GENERIC_NAMES)):
                continue
            for identifier in sorted(identifiers):
                if identifier in KEYNOTE_MEDIA_PROPERTIES_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_MEDIA_PROPERTIES_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_MEDIA_PROPERTIES_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote media-properties public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote media-properties public API retains a glob "
                    f"re-export: {path.relative_to(root)}:{line_number}"
                )
            if RUST_BYTE_SLICE.search(declaration):
                violations.append(
                    "focused litchi-keynote media-properties public API exposes raw bytes: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_MEDIA_PROPERTIES_RAW_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote media-properties public API exposes raw identifier "
                    f"{match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def audit_keynote_media_properties_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the neutral lazy Buffa media-properties codec and private view."""

    if not _keynote_media_properties_owner_present(root):
        return []

    codec_path = root / KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE
    raw = codec_path.read_text(encoding="utf-8")
    source = _mask_rust_cfg_test_items(raw)
    code = _mask_rust_non_code(source)
    violations: list[str] = []
    for function in sorted(KEYNOTE_MEDIA_PROPERTIES_CODEC_FUNCTIONS):
        if re.search(
            rf"\b(?:pub[ \t]+)?fn[ \t]+{re.escape(function)}\b", code
        ) is None:
            violations.append(
                "focused Keynote media-properties codec is missing strict API "
                f"{function}: {KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE}"
            )
    for type_name in sorted(KEYNOTE_MEDIA_PROPERTIES_CODEC_TYPES):
        if re.search(
            rf"\b(?:pub[ \t]+)?(?:struct|enum|type|trait)[ \t]+"
            rf"{re.escape(type_name)}\b",
            code,
        ) is None:
            violations.append(
                "focused Keynote media-properties codec is missing strict type "
                f"{type_name}: {KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_MEDIA_PROPERTIES_CODEC_MARKER_GROUPS.items():
        if not all(_keynote_lifecycle_marker_present(code, marker) for marker in markers):
            violations.append(
                "focused Keynote media-properties codec is missing "
                f"{label}: {KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE}"
            )
    for pattern in KEYNOTE_MEDIA_PROPERTIES_CODEC_EAGER_PATTERNS:
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused Keynote media-properties codec must remain lazy/source-preserving; "
                f"eager generated operation {match.group(0).strip()}: "
                f"{KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE}:{line_number}"
            )
    for marker in KEYNOTE_MEDIA_PROPERTIES_PHYSICAL_TYPES:
        if re.search(rf"\b{re.escape(marker)}\b", code):
            violations.append(
                "focused Keynote media-properties codec must remain neutral and archive-free "
                f"({marker}): {KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE}"
            )

    public_path = root / KEYNOTE_MEDIA_PROPERTIES_CODEC_PUBLIC_SOURCE
    public_source = (
        _mask_rust_cfg_test_items(public_path.read_text(encoding="utf-8"))
        if public_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(KEYNOTE_MEDIA_PROPERTIES_CODEC_MODULE)}\b",
        public_source,
    ) is None:
        violations.append(
            "focused Keynote media-properties codec is missing its hidden public module "
            f"{KEYNOTE_MEDIA_PROPERTIES_CODEC_MODULE}: "
            f"{KEYNOTE_MEDIA_PROPERTIES_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        rf"(?ms)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\].{{0,260}}?"
        rf"^\s*mod\s+{re.escape(KEYNOTE_MEDIA_PROPERTIES_CODEC_GENERATED_MODULE)}\s*\{{",
        public_source,
    ) is None:
        violations.append(
            "focused Keynote media-properties generated Buffa view must remain private: "
            f"{KEYNOTE_MEDIA_PROPERTIES_CODEC_PUBLIC_SOURCE}"
        )
    for marker in (
        "buffa-keynote-media-properties",
        "iwa_keynote_media_properties_buffa_protos.rs",
    ):
        if marker not in public_source:
            violations.append(
                "focused Keynote media-properties generated Buffa include is missing "
                f"{marker}: {KEYNOTE_MEDIA_PROPERTIES_CODEC_PUBLIC_SOURCE}"
            )
    return sorted(set(violations))


def audit_keynote_media_properties_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one exact-source, aggregate-budget properties transaction."""

    if not _keynote_media_properties_owner_present(root):
        return []
    owner_path = root / KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE
    codec_path = root / KEYNOTE_MEDIA_PROPERTIES_CODEC_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    codec = _mask_rust_non_code(
        _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
    )
    prepared = owner + "\n" + codec
    violations: list[str] = []
    budget_declarations = re.findall(
        r"(?m)^\s*(?:pub(?:\([^()]*\))?[ \t]+)?struct\s+"
        r"[A-Za-z_][A-Za-z0-9_]*Budget\b",
        owner,
    )
    imported_budget = re.search(
        r"\b(?:Media|Properties|Geometry|Transaction)Budget\b", owner
    ) is not None
    if len(budget_declarations) > 1 or (not budget_declarations and not imported_budget):
        violations.append(
            "focused litchi-keynote media-properties owner must use exactly one aggregate "
            f"transaction budget: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )
    for label, markers in KEYNOTE_MEDIA_PROPERTIES_OWNER_MARKER_GROUPS.items():
        if not all(_keynote_lifecycle_marker_present(prepared, marker) for marker in markers):
            violations.append(
                "focused litchi-keynote media-properties owner is missing "
                f"{label}: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
            )
    if "decode_movie_properties_with_report" not in owner:
        violations.append(
            "focused litchi-keynote media-properties owner must route reads through the "
            f"strict codec report: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )
    if "prepare_movie_properties_rewrite" not in owner:
        violations.append(
            "focused litchi-keynote media-properties owner must route writes through a "
            f"prepared codec rewrite: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )
    if "execution_requirements" not in prepared:
        violations.append(
            "focused litchi-keynote media-properties transaction must consume prepared "
            f"execution requirements: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )
    if len(re.findall(r"\.[ \t\r\n]*execute\s*\(", owner)) != 1:
        violations.append(
            "focused litchi-keynote media-properties owner must execute exactly one prepared "
            f"rewrite: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )
    if re.search(r"\bsaturating_(?:add|sub|mul)\b", owner):
        violations.append(
            "focused litchi-keynote media-properties owner must use checked resource "
            f"accounting: {KEYNOTE_MEDIA_PROPERTIES_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _keynote_slide_table_title_owner_present(root: Path) -> bool:
    """Return whether the slide-table title owner is active."""

    owner_path = root / KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_table_title\s*;", package_source
    ) is not None


def audit_keynote_slide_table_title_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free slide-table title boundary."""

    if not _keynote_slide_table_title_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_TITLE_SELECTOR_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+slide_table_title\b",
        code[package_path] + code[lib_path],
    ):
        violations.append(
            "focused litchi-keynote slide-table title owner module must remain private: "
            f"{KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[0]}"
        )
    if re.search(r"(?m)^(?:pub\(crate\)[ \t]+)?mod[ \t]+slide_table_title\s*;", code[package_path]) is None:
        violations.append(
            "focused litchi-keynote slide-table title owner module is missing: "
            f"{KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_SLIDE_TABLE_TITLE_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_TITLE_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_TITLE_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_TITLE_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-table title public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_TYPES - semantic_exports):
        violations.append(
            "focused litchi-keynote slide-table title semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_SOURCE}"
        )
    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_TITLE_SELECTOR_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_TITLE_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote slide-table title public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_TABLE_TITLE_EXPORT_SOURCES[1]}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_TITLE_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-table title Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_table_title":
            for selector in ("SlideSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-table title Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE}"
                    )

    edit_impl = re.search(
        r"(?<![A-Za-z0-9_#])impl(?:[ \t\r\n]*<[^>{}]*>)?[ \t\r\n]+"
        r"(?:'[^ ]+[ \t\r\n]+)?Edit\b",
        code[owner_path],
    )
    edit_body = code[owner_path] if edit_impl is None else code[owner_path][edit_impl.end() :]
    for name in sorted(KEYNOTE_SLIDE_TABLE_TITLE_EDIT_METHODS):
        if not re.search(rf"\bpub[ \t]+fn[ \t]+{re.escape(name)}\b", edit_body):
            violations.append(
                "focused litchi-keynote slide-table title edit is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE}"
            )

    facade_names = (
        KEYNOTE_SLIDE_TABLE_TITLE_CANONICAL_TYPES
        | KEYNOTE_SLIDE_TABLE_TITLE_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_TABLE_TITLE_SELECTOR_TYPES
        | KEYNOTE_SLIDE_TABLE_TITLE_PACKAGE_METHODS
        | KEYNOTE_SLIDE_TABLE_TITLE_FLAT_ALIASES
    )
    for path in paths:
        dedicated = path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)}
            if not dedicated and not identifiers.intersection(facade_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_SLIDE_TABLE_TITLE_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_SLIDE_TABLE_TITLE_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_SLIDE_TABLE_TITLE_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-table title public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_TITLE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-table title public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote slide-table title public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_MOVIE_GEOMETRY_RAW_ID_PARAMETER.finditer(declaration):
                violations.append(
                    "focused litchi-keynote slide-table title public API exposes raw identifier "
                    f"{match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-table title public API retains a glob re-export: "
                    f"{path.relative_to(root)}:{line_number}"
                )
    return sorted(set(violations))


def audit_keynote_slide_table_title_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require bounded projection, exact artifacts, locality, and reopen markers."""

    if not _keynote_slide_table_title_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    violations: list[str] = []
    for label, markers in KEYNOTE_SLIDE_TABLE_TITLE_PACKAGE_MARKER_GROUPS.items():
        if not all(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table title owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_TITLE_OWNER_SOURCE}"
            )
    return sorted(set(violations))


def _keynote_slide_table_sort_owner_present(root: Path) -> bool:
    """Return whether the persisted-sort owner is active."""

    owner_path = root / KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_sort_order\s*;",
        package_source,
    ) is not None


def audit_keynote_slide_table_sort_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Keynote field-44 facade."""

    if not _keynote_slide_table_sort_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+slide_table_sort_order\b",
        code[package_path] + code[lib_path],
    ):
        violations.append(
            "focused litchi-keynote slide-table sort owner module must remain private: "
            f"{KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_sort_order\s*;",
        code[package_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table sort owner module is missing: "
            f"{KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_SLIDE_TABLE_SORT_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_SORT_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_SORT_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_SORT_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-table sort public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_TYPES - semantic_exports):
        violations.append(
            "focused litchi-keynote slide-table sort semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_SOURCE}"
        )
    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+transaction\b",
        code[semantic_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table sort public API is missing "
            f"slide::table::sort::transaction: {KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_SLIDE_TABLE_SORT_TRANSACTION_TYPES
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_SORT_TRANSACTION_TYPES - transaction_exports):
        violations.append(
            "focused litchi-keynote slide-table sort transaction API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote slide-table sort public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_TABLE_SORT_EXPORT_SOURCES[1]}"
        )
    if re.search(
        r"impl[ \t\r\n]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*)?"
        r"TableSelector\b[\s\S]*?\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+"
        r"(?:index|position)\b",
        code[selector_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table sort selector must expose checked "
            f"TableSelector::index/position: {KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_SOURCE}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_SORT_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-table sort Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_table_sort_order":
            for selector in ("SlideSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-table sort Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
                    )
        if "RowRange" in declaration:
            violations.append(
                "focused litchi-keynote slide-table sort Package method must not accept "
                f"legacy RowRange: {KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
            )
    apply_declaration = owner_methods.get("apply_slide_table_sort_order")
    if apply_declaration is not None and "Patch" not in apply_declaration:
        violations.append(
            "focused litchi-keynote slide-table sort apply method must accept "
            f"Patch: {KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
        )
    for method_name, declaration in owner_methods.items():
        if "RowRange" in declaration:
            violations.append(
                "focused litchi-keynote slide-table sort Package method must not accept "
                f"legacy RowRange ({method_name}): {KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
            )

    edit_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_SORT_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote slide-table sort edit is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_SLIDE_TABLE_SORT_CANONICAL_TYPES
        | KEYNOTE_SLIDE_TABLE_SORT_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_TYPES
        | KEYNOTE_SLIDE_TABLE_SORT_PACKAGE_METHODS
        | KEYNOTE_SLIDE_TABLE_SORT_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path}
    for path in paths:
        dedicated = path in dedicated_sources
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not identifiers.intersection(facade_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_SLIDE_TABLE_SORT_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_SLIDE_TABLE_SORT_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_SLIDE_TABLE_SORT_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-table sort public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_SORT_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-table sort public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote slide-table sort public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_TABLE_SORT_PUBLIC_RAW_ID_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote slide-table sort public API exposes raw "
                    f"parameter {match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-table sort public API retains a glob re-export: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    codec_path = root / KEYNOTE_SLIDE_TABLE_SORT_CODEC_SOURCE
    if not codec_path.is_file():
        violations.append(
            "focused litchi-keynote slide-table sort public API is missing strict neutral "
            f"codec source: {KEYNOTE_SLIDE_TABLE_SORT_CODEC_SOURCE}"
        )
    else:
        codec_source = _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        codec_code = _mask_rust_non_code(codec_source)
        for api in KEYNOTE_SLIDE_TABLE_SORT_CODEC_REQUIRED_APIS:
            if re.search(
                rf"\b(?:pub\s+)?(?:fn|struct|type)\s+{re.escape(api)}\b",
                codec_code,
            ) is None:
                violations.append(
                    "focused litchi-keynote slide-table sort hidden codec is missing "
                    f"strict API {api}: {codec_path.relative_to(root)}"
                )
        codec_lib_path = root / KEYNOTE_SLIDE_TABLE_SORT_CODEC_PUBLIC_SOURCE
        codec_lib_source = (
            _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
            if codec_lib_path.is_file()
            else ""
        )
        for codec_module in KEYNOTE_SLIDE_TABLE_SORT_CODEC_MODULES:
            if re.search(
                rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
                rf"pub\s+mod\s+{re.escape(codec_module)}\b",
                codec_lib_source,
            ) is None:
                violations.append(
                    "focused litchi-keynote slide-table sort public API is missing hidden codec "
                    f"module {codec_module}: {KEYNOTE_SLIDE_TABLE_SORT_CODEC_PUBLIC_SOURCE}"
                )

    return sorted(set(violations))


def _audit_keynote_slide_table_sort_shared_core_boundary(
    root: Path,
    owner_path: Path,
    owner: str,
) -> list[str]:
    """Keep graph authority and transaction accounting in ``slide_table_core``.

    The persisted sort owner is intentionally a small field-44 transaction
    facade.  Its selection, identity proof, locality verification, and
    aggregate resource ledger belong to the private shared graph boundary so
    later table-property owners cannot grow subtly different admission rules.
    ``owner`` has already had comments, literals, and test-only items masked;
    all patterns below therefore operate on production Rust tokens while
    retaining source offsets for useful diagnostics.
    """

    owner_label = owner_path.relative_to(root)
    violations: list[str] = []
    shared_core_names = {"slide_table_core"}
    for alias_match in re.finditer(
        r"\buse\b[^;]*\bslide_table_core\s+as\s+"
        r"([A-Za-z_][A-Za-z0-9_]*)\s*;",
        owner,
    ):
        shared_core_names.add(alias_match.group(1))
    qualified_owner = "|".join(
        re.escape(name) for name in sorted(shared_core_names)
    )
    for api in KEYNOTE_SLIDE_TABLE_SORT_SHARED_CORE_APIS:
        imported = re.search(
            rf"\buse\b[^;]*\bslide_table_core\s*::[^;]*\b"
            rf"{re.escape(api)}\b",
            owner,
        )
        qualified = re.search(
            rf"(?<![A-Za-z0-9_])(?:{qualified_owner})\s*::\s*"
            rf"{re.escape(api)}\s*\(",
            owner,
        )
        called = re.search(
            rf"(?<![A-Za-z0-9_:]){re.escape(api)}\s*\(",
            owner,
        )
        if called is not None:
            declaration_prefix = owner[max(0, called.start() - 24) : called.start()]
            if re.search(r"\bfn\s*$", declaration_prefix):
                called = None
        if qualified is None and (imported is None or called is None):
            violations.append(
                "focused litchi-keynote slide-table sort owner must route "
                f"{api} through slide_table_core: "
                f"{owner_label}"
            )

    for label, pattern in KEYNOTE_SLIDE_TABLE_SORT_FORBIDDEN_OWNER_PATTERNS:
        match = pattern.search(owner)
        if match is None:
            continue
        line_number = owner.count("\n", 0, match.start()) + 1
        violations.append(
            "focused litchi-keynote slide-table sort owner must not retain "
            f"{label}; use slide_table_core: "
            f"{owner_label}:{line_number}"
        )

    return sorted(set(violations))


def audit_keynote_slide_table_sort_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require one bounded field-44 transaction and exact preview policy."""

    if not _keynote_slide_table_sort_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    violations = _audit_keynote_slide_table_sort_shared_core_boundary(
        root, owner_path, owner
    )
    for label, markers in KEYNOTE_SLIDE_TABLE_SORT_PACKAGE_MARKER_GROUPS.items():
        if not all(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table sort owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
            )
    if re.search(
        r"deleted_previews\s*:\s*(?:usize\s*=\s*)?0|deleted_previews\s*=\s*0",
        owner,
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table sort owner must preserve previews "
            f"with deleted_previews = 0: {KEYNOTE_SLIDE_TABLE_SORT_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _keynote_slide_table_physical_sort_owner_present(root: Path) -> bool:
    """Return whether the physical Keynote Sort Now hand-off is active.

    The source file and private module declaration are intentionally not
    enough to activate this gate.  A production-only token makes activation
    explicit and keeps an incomplete owner dormant while the compatibility
    executor is still being migrated.
    """

    owner_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EXPORT_SOURCES[0]
    if not owner_path.is_file() or not package_path.is_file():
        return False
    owner = _mask_rust_non_code(
        _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    )
    package = _mask_rust_non_code(
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
    )
    return (
        KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_ACTIVATION.search(owner) is not None
        and re.search(
            r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
            r"mod[ \t]+slide_table_physical_sort\s*;",
            package,
        )
        is not None
    )


def _keynote_slide_table_physical_sort_masked_sources(
    root: Path,
) -> dict[Path, str]:
    """Read the physical-sort source set with cfg(test) items removed."""

    paths = (
        root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE,
        root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SEMANTIC_SOURCE,
        root / KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_SOURCE,
        *(root / path for path in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EXPORT_SOURCES),
    )
    return {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }


def _keynote_slide_table_physical_sort_public_leak(
    identifier: str,
) -> str | None:
    """Classify an identifier that must stay private to the physical owner."""

    # The archive-free stable permutation planner is the one shared semantic
    # dependency intentionally allowed in the public semantic module.
    if identifier == "litchi_iwa_common":
        return None
    if identifier in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PROTO_ORIGINS:
        return "protobuf type"
    if identifier in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PHYSICAL_TYPES:
        return "archive/IWA type"
    if (
        identifier == "wire"
        or identifier in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_WIRE_TYPES
    ):
        return "wire type"
    return _iwork_public_leak(identifier)


def audit_keynote_slide_table_physical_sort_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first physical Keynote Sort Now facade.

    This audit is dormant until the explicit production activation token is
    present. Once active, it checks the complete public surface separately
    from the persisted field-44 sort-order owner: the physical owner must
    expose selectors and typed transactions, while all archives, wire views,
    generated messages, IDs, and byte slices remain private.
    """

    if not _keynote_slide_table_physical_sort_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_SOURCE
    package_path, lib_path = (
        root / path for path in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EXPORT_SOURCES
    )
    sources = _keynote_slide_table_physical_sort_masked_sources(root)
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    for path, label in (
        (owner_path, "private package owner source"),
        (semantic_path, "semantic source"),
        (selector_path, "selector source"),
        (package_path, "package export source"),
        (lib_path, "crate export source"),
    ):
        if not path.is_file():
            violations.append(
                "focused litchi-keynote physical slide-table sort is missing "
                f"{label}: {path.relative_to(root)}"
            )

    package_code = code[package_path]
    module_matches = list(
        re.finditer(
            r"(?m)^(?P<indent>\s*)(?P<public>pub(?:\([^()]*\))?\s+)?"
            r"mod\s+slide_table_physical_sort\s*;",
            package_code,
        )
    )
    if not module_matches:
        violations.append(
            "focused litchi-keynote physical slide-table sort is missing private "
            f"owner module: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EXPORT_SOURCES[0]}"
        )
    elif any(
        match.group("public") and match.group("public").strip() != "pub(crate)"
        for match in module_matches
    ):
        line_number = package_code.count("\n", 0, module_matches[0].start()) + 1
        violations.append(
            "focused litchi-keynote physical slide-table sort owner module must remain "
            f"private: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EXPORT_SOURCES[0]}:{line_number}"
        )

    owner_source = sources[owner_path]
    semantic_source = sources[semantic_path]
    package_source = sources[package_path]
    lib_source = sources[lib_path]
    owner_code = code[owner_path]
    semantic_code = code[semantic_path]

    for name in sorted(KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FAMILY_SOURCE):
            # The value-free Diagnostics/Error/LimitKind/Path vocabulary is
            # owned by the semantic module; the package owner defines only the
            # package-bearing Edit/Patch/Commit.
            if path == owner_path and name not in {"Edit", "Patch", "Commit"}:
                continue
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote physical slide-table sort public API is "
                    f"missing canonical type {name}: {path.relative_to(root)}"
                )

    if re.search(r"(?m)^\s*pub\s+mod\s+transaction\b", semantic_code) is None:
        violations.append(
            "focused litchi-keynote physical slide-table sort semantic API is missing "
            f"transaction module: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SEMANTIC_SOURCE}"
        )
    for marker in ("Order", "Rule", "Scope", "RowRange"):
        if re.search(rf"\b{marker}\b", semantic_code) is None:
            violations.append(
                "focused litchi-keynote physical slide-table sort semantic API is "
                f"missing {marker}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SEMANTIC_SOURCE}"
            )

    for name in sorted(KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SELECTOR_TYPES):
        if name not in _rust_canonical_exports(
            sources[selector_path] + sources[lib_path],
            KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SELECTOR_TYPES,
        ):
            violations.append(
                "focused litchi-keynote physical slide-table sort public API is "
                f"missing selector {name}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EXPORT_SOURCES[1]}"
            )
    if re.search(
        r"impl[ \t\r\n]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*)?"
        r"TableSelector\b[\s\S]*?\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+"
        r"(?:index|position)\b",
        code[selector_path],
    ) is None:
        violations.append(
            "focused litchi-keynote physical slide-table sort selector must expose "
            f"checked TableSelector::index/position: {KEYNOTE_SLIDE_TABLE_SORT_SELECTOR_SOURCE}"
        )

    owner_methods = {
        name: (declaration, line_number)
        for name, declaration, line_number in _rust_public_methods_in_impl(
            owner_source, "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PACKAGE_METHODS):
        record = owner_methods.get(name)
        if record is None:
            violations.append(
                "focused litchi-keynote physical slide-table sort Package method is "
                f"missing {name}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}"
            )
            continue
        declaration, line_number = record
        if name == "apply_slide_table_physical_sort":
            if "Patch" not in declaration:
                violations.append(
                    "focused litchi-keynote physical slide-table sort apply method must "
                    "accept Patch: "
                    f"{KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}:{line_number}"
                )
        else:
            for selector in ("SlideSelector", "TableSelector"):
                if re.search(rf"\b{selector}\b", declaration) is None:
                    violations.append(
                        "focused litchi-keynote physical slide-table sort Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}:{line_number}"
                    )
            if (
                name == "execute_slide_table_sort_order_to_rows"
                and re.search(r"\bRowRange\b", declaration) is None
            ):
                violations.append(
                    "focused litchi-keynote physical slide-table sort selected-row "
                    "method must accept RowRange: "
                    f"{KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}:{line_number}"
                )
        raw_identifier = KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PUBLIC_RAW_PARAMETER.search(
            declaration
        )
        if raw_identifier is not None:
            violations.append(
                "focused litchi-keynote physical slide-table sort Package method exposes "
                f"raw parameter {raw_identifier.group(0).strip()}: "
                f"{KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}:{line_number}"
            )
        if re.search(r"\bu64\b|&\s*\[\s*u8\s*\]", declaration):
            violations.append(
                "focused litchi-keynote physical slide-table sort Package method exposes "
                f"raw physical type: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}:{line_number}"
            )

    edit_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            owner_source, "Edit"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote physical slide-table sort edit is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CANONICAL_TYPES
        | KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SELECTOR_TYPES
        | KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PACKAGE_METHODS
        | KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FLAT_ALIASES
        | {"Order", "Rule", "Scope", "RowRange", "transaction"}
    )
    for path in (owner_path, semantic_path, selector_path, package_path, lib_path):
        dedicated = path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            # A broad package/lib re-export can hide every low-level symbol
            # behind `*`, leaving no canonical identifier for the ordinary
            # facade filter to find. Force declarations that mention this
            # physical-sort path or any protobuf/wire origin through the leak
            # scanner before applying the narrower unrelated-item filter.
            physical_reexport = (
                "pub use" in declaration
                and re.search(
                    r"\b(?:slide_table_physical_sort|physical_sort|"
                    r"numbers_table_physical_sort_codec|table_physical_sort_codec|"
                    r"keynote_table_physical_sort_codec|litchi_iwa_protos|"
                    r"buffa|prost(?:_types)?)\b",
                    declaration,
                )
                is not None
            )
            if not dedicated and not physical_reexport and not identifiers.intersection(
                facade_names - KEYNOTE_FAMILY_GENERIC_NAMES
            ):
                continue
            for identifier in sorted(identifiers):
                reason = _keynote_slide_table_physical_sort_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote physical slide-table sort public API "
                        f"exposes {reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote physical slide-table sort public API "
                        f"retains flat alias {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote physical slide-table sort public API exposes "
                    f"raw byte slice {byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_PUBLIC_RAW_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote physical slide-table sort public API exposes "
                    f"raw parameter {match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote physical slide-table sort public API retains "
                    f"a glob re-export: {path.relative_to(root)}:{line_number}"
                )

    # A physical owner may use the infrastructure crates privately, but it
    # cannot fall back to generated Prost decoding. Keep this scan source-local
    # and production-only so test fixtures can retain adversarial decoys.
    for path, source in ((owner_path, owner_source), (semantic_path, semantic_source)):
        masked = _mask_rust_non_code(source)
        for label, pattern in KEYNOTE_NO_EAGER_PROST_SOURCE_PATTERNS:
            for match in pattern.finditer(masked):
                line_number = masked.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused litchi-keynote physical slide-table sort owner retains "
                    f"{label}: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _keynote_slide_table_physical_sort_manifest_dependencies(
    root: Path,
) -> tuple[dict[str, int], dict[str, int]]:
    """Return normal and dev dependency names with source line numbers."""

    path = root / KEYNOTE_PACKAGE_MANIFEST
    if not path.is_file():
        return {}, {}
    normal: dict[str, int] = {}
    dev: dict[str, int] = {}
    section: str | None = None
    for line_number, line in enumerate(
        path.read_text(encoding="utf-8").splitlines(), start=1
    ):
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            section = stripped[1:-1]
            continue
        is_normal_section = section == "dependencies" or (
            section.startswith("target.") and section.endswith(".dependencies")
        )
        is_dev_section = section == "dev-dependencies" or (
            section.startswith("target.") and section.endswith(".dev-dependencies")
        )
        if not (is_normal_section or is_dev_section):
            continue
        match = re.match(r"([A-Za-z0-9_-]+)\s*=", stripped)
        if match is None:
            continue
        # Cargo accepts underscores and hyphens interchangeably in package
        # keys; compare the canonical hyphenated spelling so an alias cannot
        # bypass the dependency policy.
        dependency = match.group(1).replace("_", "-")
        (normal if is_normal_section else dev)[dependency] = line_number
    return normal, dev


def _audit_keynote_slide_table_physical_sort_codec(
    root: Path,
) -> list[str]:
    """Require the hidden lazy Buffa projection/rewrite and fuzz harness."""

    violations: list[str] = []
    codec_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE
    if not codec_path.is_file():
        violations.append(
            "focused litchi-keynote physical slide-table sort is missing strict "
            f"codec source: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE}"
        )
        return violations

    raw_codec = codec_path.read_text(encoding="utf-8")
    codec_source = _mask_rust_cfg_test_items(raw_codec)
    codec_code = _mask_rust_non_code(codec_source)
    for label, names in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_REQUIRED_APIS.items():
        if any(
            re.search(
                rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(name)}\b",
                codec_code,
            )
            is not None
            for name in names
        ):
            continue
        violations.append(
            "focused litchi-keynote physical slide-table sort hidden codec is missing "
            f"{label}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE}"
        )
    for label, marker in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_FLOW_MARKERS.items():
        if marker.search(codec_code) is None:
            violations.append(
                "focused litchi-keynote physical slide-table sort hidden codec is "
                f"missing {label}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE}"
            )

    lib_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_PUBLIC_SOURCE
    lib_source = (
        _mask_rust_cfg_test_items(lib_path.read_text(encoding="utf-8"))
        if lib_path.is_file()
        else ""
    )
    for module in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_MODULES:
        if re.search(
            rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
            rf"pub\s+mod\s+{re.escape(module)}\b",
            lib_source,
        ) is None:
            violations.append(
                "focused litchi-keynote physical slide-table sort is missing hidden "
                f"codec module {module}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_PUBLIC_SOURCE}"
            )

    if re.search(r"(?m)#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\]", raw_codec) is None:
        violations.append(
            "focused litchi-keynote physical slide-table sort hidden codec is missing "
            f"cfg(test) coverage: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE}"
        )
    elif re.search(r"(?m)#\s*\[\s*test\s*\]", raw_codec) is None:
        violations.append(
            "focused litchi-keynote physical slide-table sort hidden codec is missing "
            f"#[test] coverage: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE}"
        )

    for label, pattern in KEYNOTE_NO_EAGER_PROST_SOURCE_PATTERNS:
        for match in pattern.finditer(codec_code):
            line_number = codec_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote physical slide-table sort codec retains "
                f"{label}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE}:{line_number}"
            )

    for fuzz_index, fuzz_path in enumerate(
        KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FUZZ_SOURCES
    ):
        absolute = root / fuzz_path
        # The codec target briefly used a neutral Numbers spelling during the
        # handoff. Prefer the canonical Keynote target whenever it exists;
        # otherwise accept one explicitly listed compatibility target so an
        # otherwise complete checkout does not fail solely on that rename.
        if not absolute.is_file() and fuzz_index == 0:
            for alternate_target, _alternate_corpus in (
                KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FUZZ_CODEC_ALTERNATES
            ):
                candidate = root / alternate_target
                if candidate.is_file():
                    absolute = candidate
                    break
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote physical slide-table sort is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote physical slide-table sort fuzz target is missing "
                f"fuzz_target! harness: {fuzz_path}"
            )
    for corpus_index, corpus in enumerate(
        KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FUZZ_CORPORA
    ):
        absolute = root / corpus
        if not absolute.is_dir() and corpus_index == 0:
            for _alternate_target, alternate_corpus in (
                KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FUZZ_CODEC_ALTERNATES
            ):
                candidate = root / alternate_corpus
                if candidate.is_dir():
                    absolute = candidate
                    break
        if not absolute.is_dir():
            violations.append(
                "focused litchi-keynote physical slide-table sort is missing fuzz corpus: "
                f"{corpus}"
            )
        elif not any(absolute.iterdir()):
            violations.append(
                "focused litchi-keynote physical slide-table sort fuzz corpus is empty: "
                f"{corpus}"
            )
    return violations


def audit_keynote_slide_table_physical_sort_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require full physical ownership, evidence, and dependency gates."""

    if not _keynote_slide_table_physical_sort_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SEMANTIC_SOURCE
    codec_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_CODEC_SOURCE
    owner = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        )
        if owner_path.is_file()
        else ""
    )
    semantic = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(semantic_path.read_text(encoding="utf-8"))
        )
        if semantic_path.is_file()
        else ""
    )
    codec = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        )
        if codec_path.is_file()
        else ""
    )
    violations: list[str] = []

    # The permutation and the tile/header/UID rewrite (and
    # their reassembly) moved to the shared executor `litchi_iwa_table::sort`,
    # called with Keynote's same-tile, preview-dropping hooks. The owner keeps
    # admission and candidate verification; the rewrite-affinity groups are
    # then proven by the shared executor's source instead.
    shared_executor_path = root / KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SHARED_EXECUTOR_SOURCE
    shared_executor = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(shared_executor_path.read_text(encoding="utf-8"))
        )
        if shared_executor_path.is_file()
        else ""
    )
    delegates = (
        re.search(r"\bsort_engine::execute\s*\(", owner) is not None
        and re.search(r"\binvalidating_previews_same_tile\s*\(", owner) is not None
        and all(
            marker.search(shared_executor) is not None
            for marker in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_SHARED_EXECUTOR_MARKERS
        )
    )
    for label, markers in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_MARKER_GROUPS.items():
        if delegates and label in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_DELEGATED_GROUPS:
            markers = KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_DELEGATED_GROUPS[label]
        # The row-affine admission proof belongs to the package owner. Do not
        # let a generic codec helper or semantic type name satisfy an owner
        # marker; otherwise removing an admitted structure from the actual
        # transaction could silently pass because the codec mentions it.
        if not all(marker.search(owner) is not None for marker in markers):
            violations.append(
                "focused litchi-keynote physical slide-table sort owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}"
            )

    for label, pattern in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_FORBIDDEN_OWNER_PATTERNS:
        match = pattern.search(owner)
        if match is not None:
            line_number = owner.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote physical slide-table sort owner must not retain "
                f"{label}: {KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_OWNER_SOURCE}:{line_number}"
            )

    normal, dev = _keynote_slide_table_physical_sort_manifest_dependencies(root)
    manifest_path = root / KEYNOTE_PACKAGE_MANIFEST
    if not manifest_path.is_file():
        violations.append(
            "focused litchi-keynote physical slide-table sort is missing package "
            f"manifest: {KEYNOTE_PACKAGE_MANIFEST}"
        )
    else:
        for dependency in ("litchi-numbers", "prost", "prost-types"):
            if dependency in normal:
                violations.append(
                    "focused litchi-keynote physical slide-table sort must not add "
                    f"normal dependency {dependency}: {KEYNOTE_PACKAGE_MANIFEST}:{normal[dependency]}"
                )
        for dependency in ("litchi-iwa-common", "litchi-iwa-protos"):
            if dependency not in normal:
                violations.append(
                    "focused litchi-keynote physical slide-table sort requires normal "
                    f"dependency {dependency}: {KEYNOTE_PACKAGE_MANIFEST}"
                )
        # A dev-only Prost dependency is acceptable for legacy/generated
        # fixture tests; it must never leak into production ownership.
        _ = dev

    violations.extend(_audit_keynote_slide_table_physical_sort_codec(root))

    for test_path in KEYNOTE_SLIDE_TABLE_PHYSICAL_SORT_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote physical slide-table sort is missing integration "
                f"test: {test_path}"
            )
            continue
        raw_test = absolute.read_text(encoding="utf-8")
        test_source = _mask_rust_cfg_test_items(raw_test)
        test_code = _mask_rust_non_code(test_source)
        if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
            violations.append(
                "focused litchi-keynote physical slide-table sort integration test is "
                f"missing #[test] coverage: {test_path}"
            )
        for marker in (
            "execute_slide_table_sort_order",
            "apply_slide_table_physical_sort",
            "inverse",
            "PatchConflict",
            "reopen",
        ):
            if re.search(rf"\b{re.escape(marker)}\b", test_code, re.IGNORECASE) is None:
                violations.append(
                    "focused litchi-keynote physical slide-table sort integration test "
                    f"is missing {marker} evidence: {test_path}"
                )
        if re.search(r"(?i)\b(?:native|keynote|sort\s+now)\b", test_code) is None:
            violations.append(
                "focused litchi-keynote physical slide-table sort integration test is "
                f"missing native Keynote evidence: {test_path}"
            )

    return sorted(set(violations))


def _keynote_slide_table_lock_state_owner_present(root: Path) -> bool:
    """Return whether the persisted lock owner is active."""

    owner_path = root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_lock_state\s*;",
        package_source,
    ) is not None


def audit_keynote_slide_table_lock_state_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Keynote lock-state facade."""

    if not _keynote_slide_table_lock_state_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+slide_table_lock_state\b",
        code[package_path] + code[lib_path],
    ):
        violations.append(
            "focused litchi-keynote slide-table lock owner module must remain private: "
            f"{KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_lock_state\s*;",
        code[package_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table lock owner module is missing: "
            f"{KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_SLIDE_TABLE_LOCK_STATE_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_LOCK_STATE_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_LOCK_STATE_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_LOCK_STATE_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-table lock public API is missing "
                    f"canonical type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_TYPES,
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_TYPES - semantic_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table lock semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_SOURCE}"
        )
    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+transaction\b",
        code[semantic_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table lock public API is missing "
            f"slide::table::lock::transaction: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_SLIDE_TABLE_LOCK_STATE_TRANSACTION_TYPES
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_LOCK_STATE_TRANSACTION_TYPES - transaction_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table lock transaction API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_TYPES,
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_TYPES - selector_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table lock public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_EXPORT_SOURCES[1]}"
        )
    if re.search(
        r"impl[ \t\r\n]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*)?"
        r"TableSelector\b[\s\S]*?\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+"
        r"(?:index|position)\b",
        code[selector_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table lock selector must expose checked "
            f"TableSelector::index/position: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_SOURCE}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_LOCK_STATE_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-table lock Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_table_lock_state":
            for selector in ("SlideSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-table lock Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE}"
                    )
        if name == "apply_slide_table_lock_state" and (
            "Patch" not in declaration
        ):
            violations.append(
                "focused litchi-keynote slide-table lock apply method must accept "
                f"Patch: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE}"
            )

    edit_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_LOCK_STATE_EDIT_METHODS - edit_methods
    ):
        violations.append(
            "focused litchi-keynote slide-table lock edit is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_LOCK_STATE_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_SLIDE_TABLE_LOCK_STATE_CANONICAL_TYPES
        | KEYNOTE_SLIDE_TABLE_LOCK_STATE_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_TABLE_LOCK_STATE_SELECTOR_TYPES
        | KEYNOTE_SLIDE_TABLE_LOCK_STATE_PACKAGE_METHODS
        | KEYNOTE_SLIDE_TABLE_LOCK_STATE_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path}
    for path in paths:
        dedicated = path in dedicated_sources
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not identifiers.intersection(facade_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_SLIDE_TABLE_LOCK_STATE_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_SLIDE_TABLE_LOCK_STATE_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_SLIDE_TABLE_LOCK_STATE_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    identifier_line = line_number + declaration.count(
                        "\n", 0, declaration.find(identifier)
                    )
                    violations.append(
                        "focused litchi-keynote slide-table lock public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{identifier_line}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_LOCK_STATE_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-table lock public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote slide-table lock public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_TABLE_LOCK_STATE_PUBLIC_RAW_ID_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote slide-table lock public API exposes raw parameter "
                    f"{match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-table lock public API retains a glob "
                    f"re-export: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _keynote_slide_table_dimension_owner_present(root: Path) -> bool:
    """Return whether the persisted-dimension owner is active."""

    owner_path = root / KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_dimension\s*;",
        package_source,
    ) is not None


def audit_keynote_slide_table_dimension_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Keynote dimension facade."""

    if not _keynote_slide_table_dimension_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+slide_table_dimension\b",
        code[package_path] + code[lib_path],
    ):
        violations.append(
            "focused litchi-keynote slide-table dimension owner module must remain private: "
            f"{KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_dimension\s*;",
        code[package_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table dimension owner module is missing: "
            f"{KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_SLIDE_TABLE_DIMENSION_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_DIMENSION_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_DIMENSION_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_DIMENSION_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-table dimension public API is missing "
                    f"canonical type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_TYPES,
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_TYPES - semantic_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table dimension semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_SOURCE}"
        )
    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+transaction\b",
        code[semantic_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table dimension semantic API is missing "
            f"slide::table::dimension::transaction: {KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_SLIDE_TABLE_DIMENSION_TRANSACTION_TYPES
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_DIMENSION_TRANSACTION_TYPES - transaction_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table dimension transaction API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_TYPES,
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_TYPES - selector_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table dimension public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_TABLE_DIMENSION_EXPORT_SOURCES[1]}"
        )
    if re.search(
        r"impl[ \t\r\n]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*)?"
        r"TableSelector\b[\s\S]*?\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+"
        r"(?:index|position)\b",
        code[selector_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table dimension selector must expose checked "
            f"TableSelector::index/position: {KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_SOURCE}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_DIMENSION_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-table dimension Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_table_dimension_size":
            for selector in ("SlideSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-table dimension Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE}"
                    )
            if not re.search(r"\bDimension\b", declaration):
                violations.append(
                    "focused litchi-keynote slide-table dimension Package method "
                    f"{name} must accept Dimension: {KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE}"
                )
        elif "Patch" not in declaration:
            violations.append(
                "focused litchi-keynote slide-table dimension apply method must accept "
                f"Patch: {KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE}"
            )

    edit_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_DIMENSION_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote slide-table dimension edit is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_DIMENSION_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_SLIDE_TABLE_DIMENSION_CANONICAL_TYPES
        | KEYNOTE_SLIDE_TABLE_DIMENSION_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_TABLE_DIMENSION_SELECTOR_TYPES
        | KEYNOTE_SLIDE_TABLE_DIMENSION_PACKAGE_METHODS
        | KEYNOTE_SLIDE_TABLE_DIMENSION_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path}
    for path in paths:
        dedicated = path in dedicated_sources
        # ``Error`` is also a ubiquitous nested identifier in unrelated
        # package errors (for example ``ReadError::Archive``). It remains a
        # required semantic export, but alone must not select every public
        # package error declaration for the facade leak scan.
        selection_names = facade_names if dedicated else facade_names - {"Error"}
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not identifiers.intersection(selection_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-table dimension public API retains a glob "
                    f"re-export: {path.relative_to(root)}:{line_number}"
                )
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_SLIDE_TABLE_DIMENSION_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_SLIDE_TABLE_DIMENSION_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif (
                    identifier == "wire"
                    or identifier in KEYNOTE_SLIDE_TABLE_DIMENSION_WIRE_TYPES
                ):
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-table dimension public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_DIMENSION_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-table dimension public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote slide-table dimension public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_TABLE_DIMENSION_PUBLIC_RAW_ID_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote slide-table dimension public API exposes raw "
                    f"parameter {match.group(0).strip()}: {path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _keynote_slide_table_name_owner_present(root: Path) -> bool:
    """Return whether the slide-table-name owner is active."""

    owner_path = root / KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:(?:pub(?:[ \t]*\([ \t]*crate[ \t]*\))?)[ \t]+)?"
        r"mod[ \t]+slide_table_name\s*;",
        package_source,
    ) is not None


def _keynote_slide_table_name_check_codec(root: Path) -> list[str]:
    """Require the existing strict field-8 discovery/name codec seam."""

    violations: list[str] = []
    codec_path = root / KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE
    if not codec_path.is_file():
        return [
            "focused litchi-keynote slide-table name public API is missing strict neutral "
            f"codec source: {KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE}"
        ]

    codec_raw_source = codec_path.read_text(encoding="utf-8")
    codec_source = _mask_rust_cfg_test_items(codec_raw_source)
    codec_code = _mask_rust_non_code(codec_source)
    for api in KEYNOTE_SLIDE_TABLE_NAME_CODEC_REQUIRED_APIS:
        if re.search(
            rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
            codec_code,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-table name hidden codec is missing strict API "
                f"{api}: {KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_SLIDE_TABLE_NAME_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not all(marker in codec_code for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table name hidden codec is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE}"
            )

    codec_lib_path = root / KEYNOTE_SLIDE_TABLE_NAME_CODEC_PUBLIC_SOURCE
    codec_lib_source = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    if re.search(
        rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
        rf"pub\s+mod\s+{re.escape(KEYNOTE_SLIDE_TABLE_NAME_CODEC_MODULE)}\b",
        codec_lib_source,
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table name public API is missing hidden codec "
            f"module {KEYNOTE_SLIDE_TABLE_NAME_CODEC_MODULE}: "
            f"{KEYNOTE_SLIDE_TABLE_NAME_CODEC_PUBLIC_SOURCE}"
        )

    # The neutral codec already owns malformed-wire fuzzing.  The focused
    # Keynote owner additionally needs an integration lifecycle and a bounded
    # command corpus once its gate activates.
    if re.search(r"(?m)#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\]", codec_raw_source) is None:
        violations.append(
            "focused litchi-keynote slide-table name hidden codec is missing cfg(test) coverage: "
            f"{KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE}"
        )
    elif re.search(r"(?m)#\s*\[\s*test\s*\]", codec_raw_source) is None:
        violations.append(
            "focused litchi-keynote slide-table name hidden codec is missing #[test] coverage: "
            f"{KEYNOTE_SLIDE_TABLE_NAME_CODEC_SOURCE}"
        )

    for test_path in KEYNOTE_SLIDE_TABLE_NAME_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote slide-table name boundary is missing integration test: "
                f"{test_path}"
            )
            continue
        test_source = absolute.read_text(encoding="utf-8")
        test_code = _mask_rust_non_code(test_source)
        if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
            violations.append(
                "focused litchi-keynote slide-table name integration test is missing #[test] coverage: "
                f"{test_path}"
            )
        if not re.search(
            r"\b(?:slide_table_name|edit_slide_table_name|apply_slide_table_name)\b",
            test_code,
        ):
            violations.append(
                "focused litchi-keynote slide-table name integration test does not exercise the owner: "
                f"{test_path}"
            )
        if KEYNOTE_SLIDE_TABLE_NAME_NOOP_PREVIEW_ZERO.search(test_code) is None:
            violations.append(
                "focused litchi-keynote slide-table name integration test must prove "
                "exact no-op deleted_previews == 0: "
                f"{test_path}"
            )
        if KEYNOTE_SLIDE_TABLE_NAME_CHANGED_PREVIEW_REPORT.search(test_code) is None:
            violations.append(
                "focused litchi-keynote slide-table name integration test must report "
                "root preview deletions for a changed rename: "
                f"{test_path}"
            )

    for fuzz_path in KEYNOTE_SLIDE_TABLE_NAME_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote slide-table name boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote slide-table name fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in KEYNOTE_SLIDE_TABLE_NAME_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-keynote slide-table name boundary is missing fuzz corpus: "
                f"{corpus}"
            )
    return violations


def audit_keynote_slide_table_name_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free Keynote table-name facade."""

    if not _keynote_slide_table_name_owner_present(root):
        return []

    owner_path = root / KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
    sources = {
        path: _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        if path.is_file()
        else ""
        for path in paths
    }
    code = {path: _mask_rust_non_code(source) for path, source in sources.items()}
    violations: list[str] = []

    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+slide_table_name\b",
        code[package_path] + code[lib_path],
    ):
        violations.append(
            "focused litchi-keynote slide-table name owner module must remain private: "
            f"{KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_name\s*;",
        code[package_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table name owner module is missing: "
            f"{KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_SLIDE_TABLE_NAME_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_NAME_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_NAME_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_NAME_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-table name public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_TYPES - semantic_exports):
        violations.append(
            "focused litchi-keynote slide-table name semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_SOURCE}"
        )
    if re.search(
        r"(?m)^pub[ \t]+mod[ \t]+transaction\b",
        code[semantic_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table name semantic API is missing "
            f"slide::table::name::transaction: {KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_SLIDE_TABLE_NAME_TRANSACTION_TYPES
    )
    for name in sorted(
        KEYNOTE_SLIDE_TABLE_NAME_TRANSACTION_TYPES - transaction_exports
    ):
        violations.append(
            "focused litchi-keynote slide-table name transaction API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote slide-table name public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_TABLE_NAME_EXPORT_SOURCES[1]}"
        )
    if re.search(
        r"impl[ \t\r\n]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*)?"
        r"TableSelector\b[\s\S]*?\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+"
        r"(?:index|position)\b",
        code[selector_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table name selector must expose checked "
            f"TableSelector::index/position: {KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_SOURCE}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_NAME_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-table name Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_table_name":
            for selector in ("SlideSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-table name Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE}"
                    )
        elif "Patch" not in declaration:
            violations.append(
                "focused litchi-keynote slide-table name apply method must accept "
                f"Patch: {KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE}"
            )

    edit_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_NAME_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote slide-table name edit is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE}"
        )

    facade_names = (
        KEYNOTE_SLIDE_TABLE_NAME_CANONICAL_TYPES
        | KEYNOTE_SLIDE_TABLE_NAME_SEMANTIC_TYPES
        | KEYNOTE_SLIDE_TABLE_NAME_SELECTOR_TYPES
        | KEYNOTE_SLIDE_TABLE_NAME_PACKAGE_METHODS
        | KEYNOTE_SLIDE_TABLE_NAME_FLAT_ALIASES
    )
    dedicated_sources = {owner_path, semantic_path}
    for path in paths:
        dedicated = path in dedicated_sources
        # ``Error`` is ubiquitous in unrelated package errors.  It remains a
        # required semantic export, but alone must not select every public
        # declaration in the package/lib facade leak scan.
        selection_names = facade_names if dedicated else facade_names - {"Error"}
        for declaration, line_number in _rust_public_declarations(sources[path]):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not dedicated and not identifiers.intersection(selection_names - KEYNOTE_FAMILY_GENERIC_NAMES):
                continue
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-table name public API retains a glob "
                    f"re-export: {path.relative_to(root)}:{line_number}"
                )
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_SLIDE_TABLE_NAME_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_SLIDE_TABLE_NAME_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif (
                    identifier == "wire"
                    or identifier in KEYNOTE_SLIDE_TABLE_NAME_WIRE_TYPES
                ):
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-table name public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_NAME_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-table name public API retains flat "
                        f"alias {identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote slide-table name public API exposes raw byte "
                    f"slice {byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_TABLE_NAME_PUBLIC_RAW_ID_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote slide-table name public API exposes raw "
                    f"parameter {match.group(0).strip()}: "
                    f"{path.relative_to(root)}:{line_number}"
                )

    violations.extend(_keynote_slide_table_name_check_codec(root))
    return sorted(set(violations))


def audit_keynote_slide_table_name_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require bounded field-8 rewriting, authority, and exact locality."""

    if not _keynote_slide_table_name_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(
        owner_path.read_text(encoding="utf-8")
    )
    owner = _mask_rust_non_code(owner_source)
    core_path = root / KEYNOTE_SLIDE_TABLE_CORE_SOURCE
    if core_path.is_file():
        core_source = _mask_rust_cfg_test_items(
            core_path.read_text(encoding="utf-8")
        )
        owner = f"{owner}\n{_mask_rust_non_code(core_source)}"
    violations: list[str] = []
    for label, markers in KEYNOTE_SLIDE_TABLE_NAME_PACKAGE_MARKER_GROUPS.items():
        if not all(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table name owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE}"
            )
    for label, pattern in (
        (
            "one-shot table-model name rewrite",
            re.compile(r"\brewrite_table_model_name(?:_with_report)?\s*\("),
        ),
        (
            "legacy Numbers name helper",
            re.compile(r"\brename_table_in_package\s*\("),
        ),
    ):
        match = pattern.search(owner)
        if match is not None:
            line_number = owner.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote slide-table name owner must not perform "
                f"{label}: {KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE}:{line_number}"
            )
    if re.search(
        r"deleted_previews\s*:\s*(?:usize\s*=\s*)?0|deleted_previews\s*=\s*0",
        owner,
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table name owner must preserve previews "
            f"with deleted_previews = 0: {KEYNOTE_SLIDE_TABLE_NAME_OWNER_SOURCE}"
        )
    return sorted(set(violations))


def _keynote_slide_table_headers_owner_present(root: Path) -> bool:
    """Return whether the slide-table-header owner is active.

    The owner gate intentionally stops at the package module declaration, as
    with the adjacent persisted-sort owner.  Until that wiring exists, an
    entirely absent or staged owner leaves this migration baseline dormant.
    """

    owner_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_headers\s*;",
        package_source,
    ) is not None


def _keynote_slide_table_headers_public_declaration(
    declaration: str,
    *,
    dedicated_source: bool,
) -> bool:
    """Select facade declarations without inspecting unrelated crate APIs."""

    if dedicated_source:
        return True
    identifiers = {
        match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
    }
    if (
        "slide_table_headers" in identifiers
        or ("package" in identifiers and "*" in declaration)
        or ("headers" in identifiers and "*" in declaration)
    ):
        return True
    return bool(
        identifiers
        & (
            (KEYNOTE_SLIDE_TABLE_HEADERS_CANONICAL_TYPES - KEYNOTE_FAMILY_GENERIC_NAMES)
            | KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_TYPES
            | KEYNOTE_SLIDE_TABLE_HEADERS_PACKAGE_METHODS
            | KEYNOTE_SLIDE_TABLE_HEADERS_FLAT_ALIASES
        )
    )


def _keynote_slide_table_headers_check_public_leaks(
    root: Path,
    sources: dict[Path, str],
    paths: tuple[Path, ...],
) -> list[str]:
    """Return public physical/wire/raw-ID leaks from the focused facade."""

    violations: list[str] = []
    owner_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_SOURCE
    for path in paths:
        dedicated = path in {owner_path, semantic_path}
        for declaration, line_number in _rust_public_declarations(sources[path]):
            if not _keynote_slide_table_headers_public_declaration(
                declaration, dedicated_source=dedicated
            ):
                continue
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            for identifier in sorted(identifiers):
                if identifier == "litchi_iwa_common":
                    reason = None
                elif identifier in KEYNOTE_SLIDE_TABLE_HEADERS_PROTO_ORIGINS:
                    reason = "protobuf type"
                elif identifier in KEYNOTE_SLIDE_TABLE_HEADERS_PHYSICAL_TYPES:
                    reason = "archive/IWA type"
                elif identifier == "wire" or identifier in KEYNOTE_SLIDE_TABLE_HEADERS_WIRE_TYPES:
                    reason = "wire type"
                else:
                    reason = _iwork_public_leak(identifier)
                if reason is not None:
                    violations.append(
                        "focused litchi-keynote slide-table headers public API exposes "
                        f"{reason} {identifier}: {path.relative_to(root)}:{line_number}"
                    )
                if identifier in KEYNOTE_SLIDE_TABLE_HEADERS_FLAT_ALIASES:
                    violations.append(
                        "focused litchi-keynote slide-table headers public API retains flat alias "
                        f"{identifier}: {path.relative_to(root)}:{line_number}"
                    )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                byte_slice = re.sub(r"\s+", "", match.group(0))
                violations.append(
                    "focused litchi-keynote slide-table headers public API exposes raw byte slice "
                    f"{byte_slice}: {path.relative_to(root)}:{line_number}"
                )
            for match in KEYNOTE_SLIDE_TABLE_HEADERS_PUBLIC_RAW_ID_PARAMETER.finditer(
                declaration
            ):
                violations.append(
                    "focused litchi-keynote slide-table headers public API exposes raw "
                    f"parameter {match.group(0).strip()}: "
                    f"{path.relative_to(root)}:{line_number}"
                )
            if "pub use" in declaration and "*" in declaration:
                violations.append(
                    "focused litchi-keynote slide-table headers public API retains a glob re-export: "
                    f"{path.relative_to(root)}:{line_number}"
                )
    return violations


def _keynote_slide_table_headers_check_codec(
    root: Path,
) -> list[str]:
    """Require the shared strict, prepared seven-field codec seam."""

    violations: list[str] = []
    codec_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE
    if not codec_path.is_file():
        return [
            "focused litchi-keynote slide-table headers public API is missing strict neutral "
            f"codec source: {KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE}"
        ]
    codec_raw_source = codec_path.read_text(encoding="utf-8")
    codec_source = _mask_rust_cfg_test_items(codec_raw_source)
    codec_code = _mask_rust_non_code(codec_source)
    for api in KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_REQUIRED_APIS:
        if re.search(
            rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
            codec_code,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-table headers hidden codec is missing strict API "
                f"{api}: {KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE}"
            )
    for label, markers in KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not all(marker in codec_code for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table headers hidden codec is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE}"
            )
    codec_lib_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_PUBLIC_SOURCE
    codec_lib_source = (
        _mask_rust_cfg_test_items(codec_lib_path.read_text(encoding="utf-8"))
        if codec_lib_path.is_file()
        else ""
    )
    for codec_module in KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_MODULES:
        if re.search(
            rf"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
            rf"pub\s+mod\s+{re.escape(codec_module)}\b",
            codec_lib_source,
        ) is None:
            violations.append(
                "focused litchi-keynote slide-table headers public API is missing hidden codec "
                f"module {codec_module}: {KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_PUBLIC_SOURCE}"
            )

    # A strict owner needs both a unit-test seam and a fuzz target/corpus.  The
    # codec's own cfg(test) module is the low-level seam, while the
    # package integration test proves the selector-first owner is exercised.
    # Both are required once the owner gate activates.
    if re.search(r"(?m)#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\]", codec_raw_source) is None:
        violations.append(
            "focused litchi-keynote slide-table headers hidden codec is missing cfg(test) coverage: "
            f"{KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE}"
        )
    elif re.search(r"(?m)#\s*\[\s*test\s*\]", codec_raw_source) is None:
        violations.append(
            "focused litchi-keynote slide-table headers hidden codec is missing #[test] coverage: "
            f"{KEYNOTE_SLIDE_TABLE_HEADERS_CODEC_SOURCE}"
        )
    for test_path in KEYNOTE_SLIDE_TABLE_HEADERS_TEST_SOURCES:
        absolute = root / test_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote slide-table headers boundary is missing integration test: "
                f"{test_path}"
            )
            continue
        # Integration tests are normally compiled as test crates and do not
        # carry a surrounding cfg(test) item.  Keep the raw test attribute and
        # owner call visible so a cfg-only decoy cannot satisfy this seam.
        test_source = absolute.read_text(encoding="utf-8")
        test_code = _mask_rust_non_code(test_source)
        if re.search(r"(?m)#\s*\[\s*test\s*\]", test_code) is None:
            violations.append(
                "focused litchi-keynote slide-table headers integration test is missing #[test] coverage: "
                f"{test_path}"
            )
        if not re.search(
            r"\b(?:slide_table_header_settings|edit_slide_table_headers|"
            r"apply_slide_table_headers)\b",
            test_code,
        ):
            violations.append(
                "focused litchi-keynote slide-table headers integration test does not exercise the owner: "
                f"{test_path}"
            )
    for fuzz_path in KEYNOTE_SLIDE_TABLE_HEADERS_FUZZ_SOURCES:
        absolute = root / fuzz_path
        if not absolute.is_file():
            violations.append(
                "focused litchi-keynote slide-table headers boundary is missing fuzz target: "
                f"{fuzz_path}"
            )
        elif "fuzz_target!" not in absolute.read_text(encoding="utf-8"):
            violations.append(
                "focused litchi-keynote slide-table headers fuzz target is missing fuzz_target! harness: "
                f"{fuzz_path}"
            )
    for corpus in KEYNOTE_SLIDE_TABLE_HEADERS_FUZZ_CORPORA:
        if not (root / corpus).is_dir():
            violations.append(
                "focused litchi-keynote slide-table headers boundary is missing fuzz corpus: "
                f"{corpus}"
            )
    return violations


def audit_keynote_slide_table_headers_facade_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Enforce the selector-first, archive-free seven-field header facade."""

    if not _keynote_slide_table_headers_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE
    semantic_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_SOURCE
    selector_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_SOURCE
    package_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[0]
    lib_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[1]
    paths = (owner_path, semantic_path, selector_path, package_path, lib_path)
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
        r"(?m)^pub[ \t]+mod[ \t]+slide_table_headers\b",
        code[package_path] + code[lib_path],
    ):
        violations.append(
            "focused litchi-keynote slide-table headers owner module must remain private: "
            f"{KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[0]}"
        )
    if re.search(
        r"(?m)^(?:pub[ \t]*\([ \t]*crate[ \t]*\)[ \t]+)?"
        r"mod[ \t]+slide_table_headers\s*;",
        package_code,
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table headers owner module is missing: "
            f"{KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[0]}"
        )

    for name in sorted(KEYNOTE_SLIDE_TABLE_HEADERS_CANONICAL_TYPES):
        for path in (owner_path, root / KEYNOTE_SLIDE_TABLE_HEADERS_FAMILY_SOURCE):
            if name not in _rust_canonical_exports(
                _keynote_family_view(root, path, sources, KEYNOTE_SLIDE_TABLE_HEADERS_FAMILY_MODULE), KEYNOTE_SLIDE_TABLE_HEADERS_CANONICAL_TYPES
            ):
                violations.append(
                    "focused litchi-keynote slide-table headers public API is missing canonical "
                    f"type {name}: {path.relative_to(root)}"
                )

    semantic_exports = _rust_canonical_exports(
        sources[semantic_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_TYPES - semantic_exports):
        violations.append(
            "focused litchi-keynote slide-table headers semantic API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_SOURCE}"
        )
    if re.search(r"(?m)^pub[ \t]+mod[ \t]+transaction\b", code[semantic_path]) is None:
        violations.append(
            "focused litchi-keynote slide-table headers semantic API is missing "
            f"slide::table::headers::transaction: {KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_SOURCE}"
        )
    transaction_exports = _rust_canonical_exports(
        sources[semantic_path], KEYNOTE_SLIDE_TABLE_HEADERS_TRANSACTION_TYPES
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_HEADERS_TRANSACTION_TYPES - transaction_exports):
        violations.append(
            "focused litchi-keynote slide-table headers transaction API is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_HEADERS_SEMANTIC_SOURCE}"
        )

    selector_exports = _rust_canonical_exports(
        sources[selector_path] + sources[lib_path],
        KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_TYPES,
    )
    for name in sorted(KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_TYPES - selector_exports):
        violations.append(
            "focused litchi-keynote slide-table headers public API is missing selector "
            f"{name}: {KEYNOTE_SLIDE_TABLE_HEADERS_EXPORT_SOURCES[1]}"
        )
    if re.search(
        r"impl[ \t\r\n]+(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*)?"
        r"TableSelector\b[\s\S]*?\bpub[ \t]+(?:const[ \t]+)?fn[ \t]+"
        r"(?:index|position)\b",
        code[selector_path],
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table headers selector must expose checked "
            f"TableSelector::index/position: {KEYNOTE_SLIDE_TABLE_HEADERS_SELECTOR_SOURCE}"
        )

    owner_methods = {
        name: declaration
        for name, declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Package"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_HEADERS_PACKAGE_METHODS):
        declaration = owner_methods.get(name)
        if declaration is None:
            violations.append(
                "focused litchi-keynote slide-table headers Package method is missing "
                f"{name}: {KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE}"
            )
            continue
        if name != "apply_slide_table_headers":
            for selector in ("SlideSelector", "TableSelector"):
                if not re.search(rf"\b{selector}\b", declaration):
                    violations.append(
                        "focused litchi-keynote slide-table headers Package method "
                        f"{name} must accept selector-first {selector}: "
                        f"{KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE}"
                    )
    apply_declaration = owner_methods.get("apply_slide_table_headers")
    if apply_declaration is not None and "Patch" not in apply_declaration:
        violations.append(
            "focused litchi-keynote slide-table headers apply method must accept "
            f"Patch: {KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE}"
        )

    edit_methods = {
        name
        for name, _declaration, _line_number in _rust_public_methods_in_impl(
            sources[owner_path], "Edit"
        )
    }
    for name in sorted(KEYNOTE_SLIDE_TABLE_HEADERS_EDIT_METHODS - edit_methods):
        violations.append(
            "focused litchi-keynote slide-table headers edit is missing "
            f"{name}: {KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE}"
        )

    violations.extend(
        _keynote_slide_table_headers_check_public_leaks(root, sources, paths)
    )
    violations.extend(_keynote_slide_table_headers_check_codec(root))
    return sorted(set(violations))


def audit_keynote_slide_table_headers_resource_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require bounded one-model rewriting and exact preservation policy."""

    if not _keynote_slide_table_headers_owner_present(root):
        return []
    owner_path = root / KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE
    owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
    owner = _mask_rust_non_code(owner_source)
    violations: list[str] = []
    for label, markers in KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_MARKER_GROUPS.items():
        if not all(marker in owner for marker in markers):
            violations.append(
                "focused litchi-keynote slide-table headers owner is missing "
                f"{label} marker: {KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE}"
            )
    for label, pattern in KEYNOTE_SLIDE_TABLE_HEADERS_FORBIDDEN_OWNER_PATTERNS:
        match = pattern.search(owner)
        if match is not None:
            line_number = owner.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-keynote slide-table headers owner must not perform "
                f"{label}: {KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE}:{line_number}"
            )
    if re.search(
        r"deleted_previews\s*:\s*(?:usize\s*=\s*)?0|deleted_previews\s*=\s*0",
        owner,
    ) is None:
        violations.append(
            "focused litchi-keynote slide-table headers owner must preserve previews "
            f"with deleted_previews = 0: {KEYNOTE_SLIDE_TABLE_HEADERS_OWNER_SOURCE}"
        )
    return sorted(set(violations))
KEYNOTE_TABLE_MERGE_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?slide_table_merges\b[ \t]*(?:;|\{)"
)
KEYNOTE_TABLE_MERGE_COMMON_TABLE_MODULE = PAGES_TABLE_MERGE_COMMON_TABLE_MODULE
KEYNOTE_TABLE_MERGE_FOCUSED_CALL = PAGES_TABLE_MERGE_FOCUSED_CALL
KEYNOTE_TABLE_MERGE_PUBLIC_METHOD = "slide_table_merges"
KEYNOTE_TABLE_MERGE_SELECTOR_TYPES = frozenset({"SlideSelector", "TableSelector"})
KEYNOTE_TABLE_MERGE_READER_SOURCE = Path(
    "crates/litchi-keynote/src/package/slide_table_merges/reader.rs"
)
KEYNOTE_TABLE_MERGE_FORBIDDEN_MODULES = PAGES_TABLE_MERGE_FORBIDDEN_MODULES
KEYNOTE_TABLE_CELLS_SOURCE = Path(
    "crates/litchi-keynote/src/package/slide_table_cells.rs"
)
KEYNOTE_TABLE_CELLS_PACKAGE_SOURCE = Path("crates/litchi-keynote/src/package.rs")
KEYNOTE_TABLE_CELLS_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod[ \t\r\n]+"
    r"(?:r#)?slide_table_cells\b[ \t]*(?:;|\{)"
)


def audit_keynote_table_merge_source_topology(root: Path = ROOT) -> list[str]:
    """Keep Keynote merged-cell reads selector-first and archive-free."""

    violations = _audit_focused_table_merge_source_topology(
        root,
        format_name="Keynote",
        owner_path=KEYNOTE_TABLE_MERGE_SOURCE,
        package_path=KEYNOTE_TABLE_MERGE_PACKAGE_SOURCE,
        module_pattern=KEYNOTE_TABLE_MERGE_MODULE,
        table_module_path=Path("crates/litchi-keynote/src/slide/table.rs"),
        table_merge_path=Path("crates/litchi-keynote/src/slide/table/merge.rs"),
        table_module_pattern=KEYNOTE_TABLE_MERGE_COMMON_TABLE_MODULE,
        public_method=KEYNOTE_TABLE_MERGE_PUBLIC_METHOD,
        selector_types=KEYNOTE_TABLE_MERGE_SELECTOR_TYPES,
    )
    violations.extend(audit_keynote_table_merge_reader_source_topology(root))
    return sorted(set(violations))


def audit_keynote_table_merge_reader_source_topology(root: Path = ROOT) -> list[str]:
    """Keep Keynote's focused merge-reader child private and ergonomic."""

    return _audit_focused_table_merge_reader_source_topology(
        root,
        format_name="Keynote",
        owner_path=KEYNOTE_TABLE_MERGE_SOURCE,
        reader_path=KEYNOTE_TABLE_MERGE_READER_SOURCE,
        query_method="slide_table_merges",
        selector_types=KEYNOTE_TABLE_MERGE_SELECTOR_TYPES,
        limits_type="OpenOptions",
        options_type="OpenOptions",
    )


def audit_keynote_table_merge_transaction_source_topology(
    root: Path = ROOT,
) -> list[str]:
    return _audit_table_merge_transaction_source_topology(root, IWA_TABLE_MERGE_TRANSACTION_SPECS[2])


def audit_keynote_table_cells_source_topology(root: Path = ROOT) -> list[str]:
    """Keep Keynote slide-table cell reads selector-first and semantic."""

    return _audit_focused_table_cells_source_topology(
        root,
        format_name="Keynote",
        owner_path=KEYNOTE_TABLE_CELLS_SOURCE,
        package_path=KEYNOTE_TABLE_CELLS_PACKAGE_SOURCE,
        module_pattern=KEYNOTE_TABLE_CELLS_MODULE,
        public_method="slide_table_cells",
    )
