"""Shared rules and source inspection for workspace dependency checks."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import tomllib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parents[2]

# The archive-free cell data-format
# values moved from `litchi-numbers` to `litchi-iwa-common::table`, and each
# `litchi-numbers` module now re-exports them beside its transactions. The
# focused Numbers audits below read a data-format module as its semantic
# owner; reading it as the module plus its moved value definitions keeps those
# audits checking the same declarations at their canonical home.
_NUMBERS_DATA_FORMAT_SOURCE = ("crates", "litchi-numbers", "src", "cell")
_COMMON_DATA_FORMAT_SOURCE = ("crates", "litchi-iwa-common", "src", "table", "cell")
_PATH_READ_TEXT = Path.read_text


def _moved_data_format_counterpart(path: Path) -> Path | None:
    parts = path.parts
    for index in range(len(parts) - len(_NUMBERS_DATA_FORMAT_SOURCE)):
        window = parts[index : index + len(_NUMBERS_DATA_FORMAT_SOURCE)]
        if window != _NUMBERS_DATA_FORMAT_SOURCE:
            continue
        rest = parts[index + len(_NUMBERS_DATA_FORMAT_SOURCE) :]
        if not rest or rest[0] not in {"data_format", "data_format.rs"}:
            return None
        return Path(*parts[:index], *_COMMON_DATA_FORMAT_SOURCE, *rest)
    return None


def _read_text_with_moved_data_formats(self: Path, *args, **kwargs) -> str:
    text = _PATH_READ_TEXT(self, *args, **kwargs)
    counterpart = _moved_data_format_counterpart(self)
    if counterpart is not None and counterpart.is_file():
        text = f"{text}\n{_PATH_READ_TEXT(counterpart, *args, **kwargs)}"
    return text


Path.read_text = _read_text_with_moved_data_formats  # type: ignore[method-assign]
DEFAULT_POLICY = ROOT / "tools" / "crate_boundaries.json"

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
IWA_PROTOS_FACADE_SOURCE = Path("crates/litchi-iwa-protos/src/lib.rs")
RETIRED_IWA_PROTOS_ROOT_ALIASES = frozenset(
    {
        "knsos",
        "tnsos",
        "tpsos",
        "tsasos",
        "tschsos",
        "tsck",
        "tscksos",
        "tsdsos",
        "tsssos",
        "tstsos",
        "tswpsos",
    }
)
IWA_PROTOS_PUBLIC_USE = re.compile(
    r"^[ \t]*pub(?![ \t\r\n]*\()[ \t\r\n]+use[ \t\r\n]+"
    r"(?P<tail>.*?);",
    re.MULTILINE | re.DOTALL,
)
IWA_PROTOS_EXTERNAL_MODULE_DECLARATION = re.compile(
    r"^[ \t]*pub(?![ \t\r\n]*\()[ \t\r\n]+mod[ \t\r\n]+"
    r"(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b",
    re.MULTILINE,
)
RETIRED_IWA_BUNDLE_ARCHIVE_CONSTRUCTORS = (
    "from_archive_bytes",
    "from_archive_bytes_with_limits",
)
RETIRED_IWA_BUNDLE_ARCHIVE_CONSTRUCTOR_SET = frozenset(
    RETIRED_IWA_BUNDLE_ARCHIVE_CONSTRUCTORS
)
# The umbrella's iWork examples are semantic-reader examples. Keep the
# compatibility host's old editor examples from becoming an accidental
# recommendation for callers, even though the legacy crate itself remains
# available for source compatibility.
IWORK_EXAMPLE_SOURCES = (
    Path("crates/litchi/examples/read_iwork.rs"),
)
IWORK_EXAMPLE_LEGACY_API = re.compile(
    r"(?<![A-Za-z0-9_])(?:Draft|HostPackage)(?![A-Za-z0-9_])"
)
IWORK_EXAMPLE_NATIVE_ID_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:id|object_id|native_id)[ \t\r\n]*\("
)
KEYNOTE_SOURCE_ROOT = Path("crates/litchi-keynote/src")
KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE = (
    KEYNOTE_SOURCE_ROOT / "package" / "slide_chart_arrangement.rs"
)
KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES = (
    KEYNOTE_SOURCE_ROOT / "package.rs",
    KEYNOTE_SOURCE_ROOT / "lib.rs",
)
KEYNOTE_CHART_ARRANGEMENT_PACKAGE_MODULE = "slide_chart_arrangement"
CHART_ARRANGEMENT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/chart_arrangement_codec.rs"
)
RETIRED_KEYNOTE_CHART_ARRANGEMENT_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/keynote_chart_arrangement_codec.rs"
)
RETIRED_KEYNOTE_CHART_ARRANGEMENT_CODEC_MODULE = "keynote_chart_arrangement_codec"
CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
CHART_ARRANGEMENT_CODEC_MODULE = "chart_arrangement_codec"
CHART_ARRANGEMENT_CODEC_GENERATED_MODULE = (
    "buffa_chart_arrangement_generated"
)
CHART_ARRANGEMENT_CODEC_REQUIRED_APIS = (
    "DecodeError",
    "DecodeOptions",
    "ChartArrangementSnapshot",
    "ChartArrangementWrite",
    "DecodeReport",
    "RewriteReport",
    "decode_chart_arrangement",
    "decode_chart_arrangement_with_report",
    "prepare_chart_arrangement_rewrite",
    "rewrite_chart_arrangement",
    "PreparedChartArrangementRewrite",
    "RewriteExecutionRequirements",
    "RewriteExecutionLimits",
)
CHART_ARRANGEMENT_CODEC_REQUIRED_MARKER_GROUPS = {
    "lazy Buffa projection": (
        CHART_ARRANGEMENT_CODEC_GENERATED_MODULE,
        ("decode_lazy_view", "decode_view"),
        "buffa",
    ),
    "strict canonical ingress": (
        "preflight",
        "duplicate",
        ("noncanonical", "non-canonical"),
    ),
    "unknown/raw preservation": ("unknown", "raw", "extend_from_slice"),
    "bounded prepared rewrite": (
        "MAX_RECURSION_LIMIT",
        "execution_requirements",
        "execute",
    ),
    "arrangement fields": ("locked", "aspect_ratio", "constrain_proportions"),
}
CHART_ARRANGEMENT_CODEC_HIDDEN_MODULE = re.compile(
    rf"(?m)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
    rf"pub\s+mod\s+{re.escape(CHART_ARRANGEMENT_CODEC_MODULE)}\b"
)
CHART_ARRANGEMENT_CODEC_PUBLIC_GENERATED_MODULE = re.compile(
    rf"(?m)^\s*pub\s+mod\s+{re.escape(CHART_ARRANGEMENT_CODEC_GENERATED_MODULE)}\b"
)

# Keynote no longer provides the raw borrowed-payload compatibility ingress;
# the shared playback ratchet below treats this name as a forbidden symbol.
IWA_KEYNOTE_MOVIE_PLAYBACK_RETIRED_SEAM = "__decode_movie_playback_payload"
IWA_KEYNOTE_MOVIE_PLAYBACK_RETIRED_SEAM_REFERENCE = re.compile(
    rf"(?<![A-Za-z0-9_#]){re.escape(IWA_KEYNOTE_MOVIE_PLAYBACK_RETIRED_SEAM)}"
    r"(?![A-Za-z0-9_])"
)

IWA_SHARED_MEDIA_PLAYBACK_ROOT_MODULE = re.compile(
    r"(?m)^[ \t]*pub(?:\([^()]*\))?[ \t]+mod[ \t]+(?:r#)?media_playback\b"
)
IWA_MEDIA_PLAYBACK_OWNER_SOURCES = {
    "Keynote": Path("crates/litchi-keynote/src/package/slide_movie_playback.rs"),
    "Pages": Path("crates/litchi-pages/src/package/movie_playback.rs"),
}
IWA_MEDIA_PLAYBACK_OWNER_PACKAGE_SOURCES = {
    "Keynote": Path("crates/litchi-keynote/src/package.rs"),
    "Pages": Path("crates/litchi-pages/src/package.rs"),
}
IWA_MEDIA_PLAYBACK_OWNER_MODULES = {
    "Keynote": re.compile(
        r"(?m)^[ \t]*(?:pub\(crate\)[ \t]+)?mod[ \t]+(?:r#)?slide_movie_playback\s*;"
    ),
    "Pages": re.compile(r"(?m)^[ \t]*mod[ \t]+(?:r#)?movie_playback\s*;"),
}
IWA_MEDIA_PLAYBACK_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/movie_playback_codec.rs"
)
IWA_MEDIA_PLAYBACK_GENERATED_LEAK = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*::)?(?:MovieArchive|MovieArchiveMessage)\s*::\s*decode\s*\("
    r"|\b(?:prost|buffa)\s*::\s*Message\b"
    r"|\b(?:tsd|tsp)\s*::\s*MovieArchive\b"
)
IWA_SHARED_IMAGE_ADJUSTMENTS_ROOT_MODULE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?mod[ \t]+"
    r"(?:r#)?image_adjustments\b"
)
IWA_IMAGE_ADJUSTMENTS_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/image_adjustments_codec.rs"
)
IWA_IMAGE_ADJUSTMENTS_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
IWA_IMAGE_ADJUSTMENTS_CODEC_MODULE = re.compile(
    r"(?ms)^\s*#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\]\s*\n"
    r"\s*pub\s+mod\s+(?:r#)?image_adjustments_codec\s*;"
)
IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES = {
    "Keynote": Path("crates/litchi-keynote/src/package/image_adjustments.rs"),
    "Pages": Path("crates/litchi-pages/src/package/image_adjustments.rs"),
    "Numbers": Path("crates/litchi-numbers/src/package/image_adjustments.rs"),
}
IWA_IMAGE_ADJUSTMENTS_OWNER_PACKAGE_SOURCES = {
    "Keynote": Path("crates/litchi-keynote/src/package.rs"),
    "Pages": Path("crates/litchi-pages/src/package.rs"),
    "Numbers": Path("crates/litchi-numbers/src/package.rs"),
}
IWA_IMAGE_ADJUSTMENTS_OWNER_EXPORT_SOURCES = {
    "Keynote": Path("crates/litchi-keynote/src/lib.rs"),
    "Pages": Path("crates/litchi-pages/src/lib.rs"),
    "Numbers": Path("crates/litchi-numbers/src/shape/image.rs"),
}
IWA_IMAGE_ADJUSTMENTS_OWNER_MODULES = {
    # All three packages expose selector-first semantic image APIs by default.
    ecosystem: re.compile(r"(?m)^\s*mod\s+(?:r#)?image_adjustments\s*;")
    for ecosystem in ("Numbers",)
} | {
    # ADR 0034: the Pages and Keynote owners are crate-visible so their
    # focused family modules can re-export the short names; ``pub mod``
    # remains rejected.
    ecosystem: re.compile(
        r"(?m)^\s*pub\s*\(\s*crate\s*\)\s*mod\s+(?:r#)?image_adjustments\s*;"
    )
    for ecosystem in ("Keynote", "Pages")
}
# Ecosystems whose semantic image-adjustment types are published through an
# ADR 0034 transaction module (checked by the shared Pages family helper)
# rather than through the crate root.
IWA_IMAGE_ADJUSTMENTS_TRANSACTION_FAMILY_OWNERS = {"Pages": "image_adjustments"}
IWA_IMAGE_ADJUSTMENTS_RETIRED_WRITE_SEAM = (
    "__rewrite_image_adjustments_payload"
)
IWA_IMAGE_ADJUSTMENTS_RETIRED_WRITE_ECOSYSTEMS = frozenset(
    IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES
)
IWA_IMAGE_ADJUSTMENTS_SEMANTIC_METHODS = {
    "Keynote": (
        "slide_image_adjustments",
        "edit_slide_image_adjustments",
        "apply_slide_image_adjustments",
    ),
    "Pages": (
        "body_image_adjustments",
        "edit_body_image_adjustments",
        "apply_body_image_adjustments",
    ),
    "Numbers": (
        "sheet_image_adjustments",
        "edit_sheet_image_adjustments",
        "apply_sheet_image_adjustments",
    ),
}
IWA_IMAGE_ADJUSTMENTS_SEMANTIC_SELECTORS = {
    "Keynote": ("SlideSelector", "ImageSelector"),
    "Pages": ("ImageSelector",),
    "Numbers": ("SheetSelector", "ImageSelector"),
}
IWA_IMAGE_ADJUSTMENTS_SEMANTIC_TYPES = {
    "Keynote": (
        "Error",
        "Edit",
        "Commit",
        "Patch",
        "Diagnostics",
    ),
    # ADR 0034 short names, published as ``litchi_pages::image::adjustments``.
    "Pages": (
        "Error",
        "Edit",
        "Commit",
        "Patch",
        "Diagnostics",
    ),
    # ADR 0034: short names in `shape::image::adjustments`.
    "Numbers": (
        "Error",
        "Edit",
        "Commit",
        "Patch",
        "Diagnostics",
    ),
}
IWA_DOC_HIDDEN_ATTRIBUTE = re.compile(
    r"#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\]"
)
IWA_IMAGE_ADJUSTMENTS_RETIRED_WRITE_CALLS = {
    ecosystem: re.compile(
        rf"\blitchi_{ecosystem.lower()}\s*::\s*"
        rf"{re.escape(IWA_IMAGE_ADJUSTMENTS_RETIRED_WRITE_SEAM)}\s*\("
    )
    for ecosystem in IWA_IMAGE_ADJUSTMENTS_RETIRED_WRITE_ECOSYSTEMS
}
IWA_IMAGE_ADJUSTMENTS_CODEC_REQUIRED_MARKERS = (
    "DecodeOptions",
    "ImageAdjustmentsSnapshot",
    "ImageAdjustmentsWrite",
    "decode_image_adjustments",
    "rewrite_image_adjustments",
)
IWA_IMAGE_ADJUSTMENTS_CODEC_LAZY_MARKERS = (
    "buffa",
    "decode_view",
    "WireView",
    "visit",
)
IWA_IMAGE_ADJUSTMENTS_CODEC_DECODE_CALL = re.compile(
    r"\bdecode_image_adjustments(?:_with_report)?\s*\("
)
IWA_IMAGE_ADJUSTMENTS_CODEC_GENERATED_LEAK = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*::)?ImageAdjustmentsArchive\s*::\s*decode\s*\("
    r"|\b(?:[A-Za-z_][A-Za-z0-9_]*::)?ImageArchive\s*::\s*decode\s*\("
    r"|\b(?:prost|buffa)\s*::\s*Message\b"
)
IWA_IMAGE_ADJUSTMENTS_GENERATED_TYPE_LEAK = re.compile(
    r"\b(?:tsd|tsp)\s*::\s*(?:ImageArchive|ImageAdjustmentsArchive)\b"
)
IWA_IMAGE_ADJUSTMENTS_HOST_LEGACY_HELPER = re.compile(
    # Pages keeps a focused editor-local ``image_adjustments`` adapter under
    # ``editor``; only the crate-root module names the retired shared owner.
    r"\bcrate\s*::\s*image_adjustments\b"
)

# Annotation authors are a neutral wire codec consumed by the Keynote owner.
# Keep the generated Buffa projection private and the source-preserving codec
# inventory independent of any one format owner so Numbers/Pages may reuse it.
ANNOTATION_AUTHOR_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/annotation_author_codec.rs"
)
ANNOTATION_AUTHOR_CODEC_PROJECTION_SOURCE = Path(
    "crates/litchi-iwa-protos/src/buffa-projections/TSKAnnotationAuthorProjection.proto"
)
ANNOTATION_AUTHOR_CODEC_PUBLIC_SOURCE = Path("crates/litchi-iwa-protos/src/lib.rs")
ANNOTATION_AUTHOR_CODEC_BUILD_SOURCE = Path("crates/litchi-iwa-protos/build.rs")
ANNOTATION_AUTHOR_CODEC_GENERATED_MODULE = "buffa_annotation_author_generated"
ANNOTATION_AUTHOR_CODEC_APIS = frozenset(
    {
        "decode_annotation_author",
        "decode_annotation_author_with_report",
        "decode_annotation_author_storage",
        "decode_annotation_author_storage_with_report",
        "encode_annotation_author",
        "encode_annotation_author_with_report",
        "encode_annotation_author_storage",
        "encode_annotation_author_storage_with_report",
        "prepare_annotation_author_storage_rewrite",
        "rewrite_annotation_author_storage",
    }
)
ANNOTATION_AUTHOR_CODEC_TYPES = frozenset(
    {
        "AnnotationAuthorSnapshot",
        "AnnotationAuthorStorageSnapshot",
        "AnnotationAuthorWrite",
        "AnnotationAuthorStorageWrite",
        "AuthorColorWrite",
        "DecodeError",
        "DecodeLimit",
        "DecodeOptions",
        "DecodeReport",
        "EncodeError",
        "EncodeLimit",
        "EncodeOptions",
        "EncodeReport",
        "PreparedAnnotationAuthorStorageRewrite",
        "RewriteExecutionLimits",
        "RewriteExecutionRequirements",
    }
)
ANNOTATION_AUTHOR_CODEC_MARKER_GROUPS = {
    "lazy Buffa view ingress": (
        "buffa",
        "decode_lazy_view",
        "LazyView",
        "ViewEncode",
    ),
    "strict author/storage field validation": (
        "AUTHOR_NAME_FIELD",
        "AUTHOR_PUBLIC_ID_FIELD",
        "STORAGE_AUTHOR_FIELD",
        "duplicate",
        "require_wire",
    ),
    "unknown-field/source preservation": (
        "unknown",
        "raw",
        "source",
        "preserv",
        "source_order",
    ),
    "bounded decode/encode accounting": (
        "max_message_bytes",
        "max_fields",
        "max_work_bytes",
        "max_allocations",
        "try_reserve_exact",
        "RewriteExecutionRequirements",
    ),
}
ANNOTATION_AUTHOR_CODEC_FORBIDDEN_PATTERNS = (
    (
        "prost import",
        re.compile(r"(?<![A-Za-z0-9_#])prost(?:_types)?[ \t\r\n]*::"),
    ),
    (
        "eager generated decode",
        re.compile(
            r"\b(?:Message|[A-Za-z_][A-Za-z0-9_]*Archive)[ \t\r\n]*::"
            r"[ \t\r\n]*(?:decode|encode_to_vec)\b"
        ),
    ),
)
# The format-neutral cell-format engine. The Numbers popup and
# cell-control audits read its markers together with the Numbers facades.
IWA_TABLE_CELL_FORMAT_ENGINE_SOURCES = (
    Path("crates/litchi-iwa-table/src/cell_format.rs"),
    Path("crates/litchi-iwa-table/src/cell_format/budget.rs"),
    Path("crates/litchi-iwa-table/src/cell_format/control.rs"),
    Path("crates/litchi-iwa-table/src/cell_format/display.rs"),
    Path("crates/litchi-iwa-table/src/cell_format/pop_up_menu.rs"),
    Path("crates/litchi-iwa-table/src/cell_format/pop_up_menu_native.rs"),
    Path("crates/litchi-iwa-table/src/cell_format/pop_up_menu_metadata.rs"),
)


def _iwa_table_cell_format_engine_code(root: Path) -> str:
    """Masked code of the shared cell-format engine."""

    return "\n".join(
        _mask_rust_non_code(_mask_rust_cfg_test_items(path.read_text(encoding="utf-8")))
        for path in (root / relative for relative in IWA_TABLE_CELL_FORMAT_ENGINE_SOURCES)
        if path.is_file()
    )
NUMBERS_SOURCE_ROOT = Path("crates/litchi-numbers/src")
NUMBERS_PACKAGE_SOURCE = NUMBERS_SOURCE_ROOT / "package.rs"
NUMBERS_CHART_ARRANGEMENT_OWNER_SOURCE = (
    NUMBERS_SOURCE_ROOT / "package" / "chart_arrangement.rs"
)
NUMBERS_CHART_ARRANGEMENT_PACKAGE_MODULE = "chart_arrangement"
NUMBERS_EXTRACTOR_SOURCE = NUMBERS_SOURCE_ROOT / "package" / "extractor.rs"

PACKAGE_METADATA_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/package_metadata_codec.rs"
)
PACKAGE_METADATA_REMOVE_FUNCTION = "remove_package_metadata"
PACKAGE_METADATA_SAVE_TOKEN_FUNCTION = "rewrite_package_metadata_save_tokens"
PACKAGE_METADATA_COMBINED_FUNCTION = (
    "rewrite_package_metadata_removals_and_save_tokens"
)
PACKAGE_METADATA_REMOVE_BATCH = "RemovalBatch"
PACKAGE_METADATA_SAVE_TOKEN_BATCH = "SaveTokenBatch"
PACKAGE_METADATA_COMBINED_BATCH = "RemovalSaveTokenBatch"
PACKAGE_METADATA_SELECTOR = "ComponentSelector"
PACKAGE_METADATA_CODEC_NO_EAGER_PATTERNS = (
    (
        "generated PackageMetadata::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:"
            r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*"
            r")*(?:r#)?PackageMetadata[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated ComponentInfo::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:"
            r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*"
            r")*(?:r#)?ComponentInfo[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated ObjectUuidMapEntry::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:"
            r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*"
            r")*(?:r#)?ObjectUuidMapEntry[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated Prost Message::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:r#)?(?:Message|M)"
            r"[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "prost dependency marker",
        re.compile(r"(?<![A-Za-z0-9_#])(?:prost|prost_types)\b"),
    ),
)
PACKAGE_METADATA_BARE_CALL = re.compile(
    r"(?<![A-Za-z0-9_:#.])(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)"
    r"[ \t\r\n]*\("
)
PACKAGE_METADATA_NO_EAGER_DECODE_PATTERNS = (
    (
        "PackageMetadata::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])"
            r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
            r"(?:r#)?PackageMetadata[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "ComponentInfo::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])"
            r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
            r"(?:r#)?ComponentInfo[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "ObjectUuidMapEntry::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])"
            r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
            r"(?:r#)?ObjectUuidMapEntry[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "UFCS generated PackageMetadata decode",
        re.compile(
            r"<[\s\S]*?(?:r#)?PackageMetadata[\s\S]*?>"
            r"[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "UFCS generated ComponentInfo decode",
        re.compile(
            r"<[\s\S]*?(?:r#)?ComponentInfo[\s\S]*?>"
            r"[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "UFCS generated ObjectUuidMapEntry decode",
        re.compile(
            r"<[\s\S]*?(?:r#)?ObjectUuidMapEntry[\s\S]*?>"
            r"[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated Prost Message::decode",
        re.compile(
            r"(?<![A-Za-z0-9_:#.])(?:r#)?Message[ \t\r\n]*::"
            r"[ \t\r\n]*decode\b"
        ),
    ),
    (
        "qualified Prost Message::decode",
        re.compile(
            r"(?<![A-Za-z0-9_])(?:r#)?prost[ \t\r\n]*::[ \t\r\n]*"
            r"(?:r#)?Message[ \t\r\n]*::[ \t\r\n]*decode\b"
        ),
    ),
    (
        "generated PackageMetadata collection materialization",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?!(?:r#)?(?:self|visitor)\b)"
            r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
            r"[ \t\r\n]*\.[ \t\r\n]*"
            r"(?:components|versioned_components|data_metadata_map)\b"
        ),
    ),
    (
        "generated ComponentInfo collection materialization",
        re.compile(
            r"(?<![A-Za-z0-9_#])(?:[A-Za-z_][A-Za-z0-9_]*|r#[A-Za-z_][A-Za-z0-9_]*)"
            r"[ \t\r\n]*\.[ \t\r\n]*"
            r"(?:object_uuid_map_entries|external_references|"
            r"versioned_external_references)\b"
        ),
    ),
    (
        "generated metadata collection type",
        re.compile(
            r"(?<![A-Za-z0-9_])(?:::)?"
            r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
            r"(?:r#)?(?:Vec|HashSet|BTreeSet)"
            r"[ \t\r\n]*<[\s\S]*?(?:PackageMetadata|ComponentInfo|"
            r"ObjectUuidMapEntry)\b"
        ),
    ),
)
IWA_FORMULA_RENDER_WIRE_IMPORT = re.compile(
    r"\blitchi_numbers_wire[ \t\r\n]*::[ \t\r\n]*formula_render\b"
)
IWA_FORMULA_RENDER_LOCAL_VISITOR_DECLARATIONS = (
    (
        "CompatibilityFormulaVisitor",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?CompatibilityFormulaVisitor\b"
        ),
    ),
    (
        "FormulaRenderCodecVisitor",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?FormulaRenderCodecVisitor\b"
        ),
    ),
    (
        "FormulaRenderVisitor implementation",
        re.compile(
            r"(?ms)^[ \t]*impl\b[^{};]*\bFormulaRenderVisitor\b"
        ),
    ),
    (
        "render_scalar_formula_nodes helper",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?render_scalar_formula_nodes\b"
        ),
    ),
)
IWA_FORMULA_ENVELOPE_WIRE_IMPORT = re.compile(
    r"\blitchi_numbers_wire[ \t\r\n]*::[ \t\r\n]*"
    r"formula_envelope\b"
)
IWA_FORMULA_ENVELOPE_LOCAL_SCHEMA_DECLARATIONS = (
    (
        "FormulaEnvelopeCost",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?FormulaEnvelopeCost\b"
        ),
    ),
    (
        "FormulaEnvelopeField",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?FormulaEnvelopeField\b"
        ),
    ),
    (
        "FormulaScalar",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?FormulaScalar\b"
        ),
    ),
    (
        "preflight_formula_envelope",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?preflight_formula_envelope\b"
        ),
    ),
    (
        "required_formula_ast_node_type",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?required_formula_ast_node_type\b"
        ),
    ),
    (
        "require_formula_fields",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?require_formula_fields\b"
        ),
    ),
    (
        "formula_field_is_repeated",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?formula_field_is_repeated\b"
        ),
    ),
    (
        "validate_formula_scalar",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?validate_formula_scalar\b"
        ),
    ),
    (
        "formula_envelope_field",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            r"(?:r#)?formula_envelope_field\b"
        ),
    ),
)
IWA_FORMULA_ENVELOPE_LOCAL_SCHEMA_CONSTANTS = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?const[ \t]+"
    r"(?:r#)?(?:FORMULA_REQUIRED_[A-Za-z0-9_]+|FORMULA_ENVELOPE_UNKNOWN)\b"
)
# The allocation-conscious expression arena lives outside both format
# readers.  The readers retain only their format-specific budget adapters,
# while the archive-free common renderer owns the handles, nodes, parts, and
# iterative output walk.  Keep this audit separate from the older generated-
# free FormulaArchive ratchet above: a reader may still mention the codec's
# FormulaNode type without owning a second arena.
IWA_COMMON_FORMULA_RENDER_SOURCE = Path(
    "crates/litchi-iwa-common/src/formula/render.rs"
)
IWA_COMMON_FORMULA_RENDER_REQUIRED_DECLARATIONS = (
    (
        "FormulaRenderBudget trait",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?trait[ \t]+"
            r"(?:r#)?FormulaRenderBudget\b"
        ),
    ),
    (
        "FormulaExpr handle",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?struct[ \t]+"
            r"(?:r#)?FormulaExpr\b"
        ),
    ),
    (
        "FormulaRenderer arena",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?struct[ \t]+"
            r"(?:r#)?FormulaRenderer\b"
        ),
    ),
    (
        "FormulaRenderer::render",
        re.compile(
            r"\b(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+(?:r#)?render\b"
        ),
    ),
)
IWA_COMMON_FORMULA_RENDER_IMPORT_STATEMENT = re.compile(
    r"(?ms)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
    r"(?:use|extern[ \t]+crate)\b.*?;"
)
IWA_COMMON_FORMULA_RENDER_FORBIDDEN_IMPORT_ROOTS = (
    "litchi_iwa_protos",
    "litchi_iwa_core",
    "litchi_iwa_archive",
    "litchi_core",
    "litchi_numbers",
    "litchi_pages",
    "litchi_keynote",
    "prost",
    "prost_types",
    "buffa",
    "tsce",
    "tst",
    "tn",
    "tsp",
    "tswp",
    "numbers_formula_codec",
    "raw",
    "protobuf",
    "proto",
    "archive",
    "wire",
)
IWA_COMMON_FORMULA_RENDER_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    + "|".join(
        re.escape(root) for root in IWA_COMMON_FORMULA_RENDER_FORBIDDEN_IMPORT_ROOTS
    )
    + r")(?![A-Za-z0-9_])"
)
IWA_FORMULA_RENDER_READER_SOURCES = (NUMBERS_EXTRACTOR_SOURCE,)
IWA_FORMULA_RENDER_COMMON_IMPORT = re.compile(
    r"\b(?:litchi_iwa_common[ \t\r\n]*::[ \t\r\n]*formula"
    r"[ \t\r\n]*::[ \t\r\n]*render|"
    r"litchi_numbers_wire[ \t\r\n]*::[ \t\r\n]*formula_render)\b"
)
IWA_FORMULA_RENDER_LOCAL_ARENA_DECLARATIONS = (
    (
        "FormulaExpr",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            r"(?:struct|enum|union|type)[ \t]+"
            r"(?:r#)?FormulaExpr\b"
        ),
    ),
    (
        "FormulaPart",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?(?:struct|enum|union)[ \t]+"
            r"(?:r#)?FormulaPart\b"
        ),
    ),
    (
        "FormulaRenderer",
        re.compile(
            r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?(?:struct|enum|union)[ \t]+"
            r"(?:r#)?FormulaRenderer\b"
        ),
    ),
)
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
PAGES_SOURCE_ROOT = Path("crates/litchi-pages/src")
PAGES_PACKAGE_SOURCE = PAGES_SOURCE_ROOT / "package.rs"

# ADR 0034: every Pages transaction family lives in a
# ``pub(crate) mod <owner>;`` under ``package/`` that defines the canonical
# short names (``Edit``, ``Patch``, ``Commit``, ``Diagnostics``, ``Error``,
# ``LimitKind`` plus a few family extras), and exactly one public module
# re-exports them with ``pub use crate::package::<owner>::{...};``.  The crate
# root re-exports none of them, and the former prefixed spellings are deleted
# rather than aliased.
#
# Each row is ``(owner module, legacy prefix, public module path, short names,
# legacy-name overrides)``.  The legacy name of a short name is
# ``prefix + short`` unless the override map names it.
PAGES_TRANSACTION_SIX = ("Commit", "Diagnostics", "Edit", "Error", "LimitKind", "Patch")
# A family may also export its typed refusal reasons.
PAGES_TRANSACTION_OPTIONAL_EXPORTS = frozenset({"Refusal"})
PAGES_TRANSACTION_FAMILIES: tuple[
    tuple[str, str, tuple[str, ...], tuple[str, ...], dict[str, str]], ...
] = (
    ("body_chart_arrangement", "BodyChartArrangement", ("chart", "arrangement"), PAGES_TRANSACTION_SIX, {}),
    ("body_chart_data", "BodyChartData", ("chart", "data"), PAGES_TRANSACTION_SIX, {}),
    ("body_chart_metadata", "BodyChartMetadata", ("chart", "metadata"), ("Error", "LimitKind"), {}),
    ("body_footnote", "BodyFootnote", ("footnote", "body", "transaction"), PAGES_TRANSACTION_SIX + ("Snapshot",), {}),
    ("body_table_appearance", "BodyTableAppearance", ("table", "appearance", "transaction"), PAGES_TRANSACTION_SIX + ("Path",), {}),
    (
        "body_table_catalog",
        "BodyTableCatalog",
        ("table", "catalog"),
        ("Catalog", "Error", "LimitKind", "Snapshot"),
        {"Catalog": "BodyTableCatalog", "Snapshot": "BodyTableSnapshot"},
    ),
    ("body_table_deletion", "BodyTableDeletion", ("table", "deletion"), PAGES_TRANSACTION_SIX, {}),
    ("body_table_dimension", "BodyTableDimension", ("table", "dimension", "transaction"), PAGES_TRANSACTION_SIX, {}),
    ("body_table_headers", "BodyTableHeaderSettings", ("table", "headers", "transaction"), PAGES_TRANSACTION_SIX + ("InvalidReason",), {}),
    ("body_table_hidden_axes", "BodyTableHiddenAxes", ("table", "hidden_axes", "transaction"), PAGES_TRANSACTION_SIX + ("Path",), {}),
    ("body_table_merges", "BodyTableMerges", ("table", "merge", "transaction"), PAGES_TRANSACTION_SIX, {}),
    ("body_table_name", "BodyTableName", ("table", "name", "transaction"), PAGES_TRANSACTION_SIX + ("Path",), {}),
    ("body_table_sort", "BodyTableSort", ("table", "sort", "transaction"), PAGES_TRANSACTION_SIX + ("Path",), {}),
    ("body_table_title", "BodyTableTitle", ("table", "title"), PAGES_TRANSACTION_SIX, {}),
    ("drawable_order", "BodyDrawableOrder", ("drawable", "order"), PAGES_TRANSACTION_SIX, {}),
    ("footnote_text", "FootnoteText", ("footnote", "text"), PAGES_TRANSACTION_SIX, {}),
    ("header_footer_text", "HeaderFooterText", ("header_footer", "text"), PAGES_TRANSACTION_SIX, {}),
    ("image_adjustments", "BodyImageAdjustments", ("image", "adjustments"), PAGES_TRANSACTION_SIX, {}),
    ("metadata_edit", "MetadataEdit", ("metadata",), PAGES_TRANSACTION_SIX, {}),
    ("page_layout", "PageLayout", ("page_layout", "transaction"), PAGES_TRANSACTION_SIX, {}),
    ("section_name", "SectionName", ("section", "name"), PAGES_TRANSACTION_SIX, {}),
    ("section_pagination", "SectionPagination", ("section", "pagination", "transaction"), PAGES_TRANSACTION_SIX, {}),
    ("section_paragraphs", "SectionParagraphs", ("section", "paragraphs"), PAGES_TRANSACTION_SIX, {}),
    ("section_text", "SectionText", ("section", "text"), PAGES_TRANSACTION_SIX, {}),
    ("table_cells", "BodyTableCells", ("table", "cells"), ("Error", "LimitKind"), {}),
    ("table_lock", "BodyTableLock", ("table", "lock"), PAGES_TRANSACTION_SIX, {}),
)
PAGES_TRANSACTION_FAMILY_BY_OWNER = {
    family[0]: family for family in PAGES_TRANSACTION_FAMILIES
}
# A family may use a shared litchi-core vocabulary item when it carries no
# format-specific state.  This is an explicit source contract, not a missing
# family exemption: the owner must re-export the exact shared item, and the
# shared definition is checked below for the required API.
PAGES_TRANSACTION_SHARED_TYPES = {
    "metadata_edit": frozenset({"Diagnostics"}),
}


# Deleted root aliases that belonged to one family.
PAGES_TRANSACTION_FAMILY_RETIRED_ALIASES = {
    "body_table_headers": frozenset({"BodyTableHeaderSettings"}),
    "body_table_name": frozenset({"BodyTableName", "BodyTableNameValueError"}),
    "table_lock": frozenset({"BodyTableLockState"}),
}


def _pages_transaction_legacy_names(owner: str) -> frozenset[str]:
    """Return the deleted prefixed spellings and aliases of one renamed family."""

    _, prefix, _, shorts, overrides = PAGES_TRANSACTION_FAMILY_BY_OWNER[owner]
    return frozenset(
        overrides.get(short, prefix + short) for short in shorts
    ) | PAGES_TRANSACTION_FAMILY_RETIRED_ALIASES.get(owner, frozenset())
PAGES_BODY_CHART_ARRANGEMENT_OWNER_SOURCE = (
    PAGES_SOURCE_ROOT / "package" / "body_chart_arrangement.rs"
)
PAGES_BODY_CHART_ARRANGEMENT_PACKAGE_MODULE = "body_chart_arrangement"
IWA_DRAWABLE_CONTAINER_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/drawable_container_codec.rs"
)
IWA_DRAWABLE_CONTAINER_CODEC_PUBLIC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/lib.rs"
)
IWA_DRAWABLE_CONTAINER_PROJECTION_SOURCE = Path(
    "crates/litchi-iwa-protos/src/buffa-projections/TSDDrawableContainerArchive.proto"
)
IWA_DRAWABLE_CONTAINER_BUILD_SOURCE = Path("crates/litchi-iwa-protos/build.rs")
IWA_DRAWABLE_CONTAINER_GENERATED_MODULE = "buffa_drawable_container_generated"
IWA_DRAWABLE_CONTAINER_GENERATED_INCLUDE = (
    "/buffa-drawable-container/iwa_drawable_container_buffa_protos.rs"
)
IWA_DRAWABLE_CONTAINER_FORBIDDEN_GENERATED_DECODE = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)?"
    r"(?:ContainerArchive|GroupArchive)[ \t\r\n]*::[ \t\r\n]*decode[ \t\r\n]*\(|"
    r"\b(?:decode_type|decode_message)[ \t\r\n]*(?:::<[^;{}]+>)?[ \t\r\n]*\(\s*"
)
IWA_DRAWABLE_CONTAINER_PROST_DECODE = re.compile(
    r"\b(?:prost[ \t\r\n]*::[ \t\r\n]*)?Message[ \t\r\n]*::[ \t\r\n]*decode[ \t\r\n]*\("
)
IWA_DRAWABLE_CONTAINER_CODEC_FORBIDDEN_OPERATIONS = (
    "to_owned_message",
    "encode_to_vec",
    "try_encode",
    ".encode(",
)
IWA_DRAWABLE_CONTAINER_CODEC_REQUIRED_MARKERS = (
    "const CONTAINER_PARENT_FIELD: u32 = 2;",
    "const CONTAINER_CHILDREN_FIELD: u32 = 3;",
    "const GROUP_SUPER_FIELD: u32 = 1;",
    "const GROUP_CHILDREN_FIELD: u32 = 2;",
    "pub struct ReferenceSnapshot",
    "pub fn decode_container_references(",
    "pub fn decode_group_references(",
    "fn preflight_reference(",
    "decode_lazy_view",
    "crate::buffa_drawable_container_generated::",
    IWA_DRAWABLE_CONTAINER_GENERATED_MODULE,
)
IWA_DRAWABLE_CONTAINER_PROJECTION_FIELDS = {
    "ContainerArchive": {
        "parent": ("Reference", 2),
        "children": ("Reference", 3),
    },
    "GroupArchive": {
        "super": ("DrawableArchive", 1),
        "children": ("Reference", 2),
    },
}

# Failure-atomic iWork package publication is centralized in the archive
# crate.  Keep the owner and the three concrete package boundaries explicit so
# a format crate cannot quietly grow a second tempfile/rename/sync policy.  The
# owner module is intentionally public: the focused format crates and the
# compatibility host call its typed publication seam directly.
IWORK_ARCHIVE_PUBLICATION_FILE = Path(
    "crates/litchi-iwa-archive/src/publication.rs"
)
IWORK_ARCHIVE_PUBLICATION_MODULE_SOURCE = Path("crates/litchi-iwa-archive/src/lib.rs")
IWORK_ARCHIVE_PUBLICATION_MODULE = "publication"
IWORK_ARCHIVE_PUBLICATION_ERROR = "Error"
IWORK_ARCHIVE_PUBLICATION_HELPER = "replace_with"

IWORK_FOCUSED_PACKAGE_SOURCE_ROOTS = (
    ("litchi-pages", Path("crates/litchi-pages/src")),
    ("litchi-numbers", Path("crates/litchi-numbers/src")),
    ("litchi-keynote", Path("crates/litchi-keynote/src")),
)
IWORK_FOCUSED_PACKAGE_SOURCES = tuple(
    source_root / "package.rs"
    for _name, source_root in IWORK_FOCUSED_PACKAGE_SOURCE_ROOTS
)
IWORK_FOCUSED_PACKAGE_SAVE_SOURCES = tuple(
    source_root / "package" / "save.rs"
    for _name, source_root in IWORK_FOCUSED_PACKAGE_SOURCE_ROOTS
)
IWORK_FOCUSED_PACKAGE_EXPORT_SOURCES = tuple(
    source_root / "lib.rs"
    for _name, source_root in IWORK_FOCUSED_PACKAGE_SOURCE_ROOTS
)

# Short aliases keep callers of the boundary checker aligned with the
# migration language used in the ADRs while the explicit names above remain
# the canonical inventory.
IWORK_ARCHIVE_PUBLICATION_SOURCE = IWORK_ARCHIVE_PUBLICATION_FILE
IWORK_PACKAGE_SAVE_SOURCES = IWORK_FOCUSED_PACKAGE_SAVE_SOURCES
CARGO_SECTION_HEADER = re.compile(r"^[ \t]*\[([^\]]+)\][ \t]*(?:#.*)?$")
CARGO_PROST_DEPENDENCY = re.compile(
    r"^[ \t]*(?:prost|\"prost\")(?:[ \t]*\.[ \t]*workspace)?[ \t]*="
)
# The hidden-axis engine (graph proof, owner creation,
# UID-map rewrite) lives in `litchi-iwa-table`; the Pages owner is its facade.
IWA_TABLE_HIDDEN_AXES_ENGINE_SOURCE = Path("crates/litchi-iwa-table/src/hidden_axes.rs")
CAMEL_CASE_WORD = re.compile(r"[A-Z]+(?=[A-Z][a-z]|$)|[A-Z]?[a-z]+|[0-9]+")
RUST_BYTE_SLICE = re.compile(
    r"&[ \t\r\n]*(?:'[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]+)?"
    r"(?:mut[ \t\r\n]+)?\[[ \t\r\n]*u8[ \t\r\n]*\]"
)
FACADE_DEFAULT_FEATURE = "default"
FACADE_ALL_FEATURE = "all"
FACADE_SOURCE_ROOT = Path("crates/litchi/src")
# The umbrella keeps package-preserving owners under ``pages``, ``keynote``,
# and ``numbers`` while exposing a deliberately smaller immutable reader under
# each ``semantic`` child module.  Keep this allowlist source-local: the leaf
# rustdoc gate checks resolved types, while this gate protects the umbrella's
# re-export shape and prevents a package alias from being hidden behind a
# semantic namespace.
SEMANTIC_FACADE_REQUIRED_EXPORTS = {
    # ADR 0034: the Pages reader returns the leaf's
    # ``PackageError`` and is configured by ``OpenOptions`` with the
    # ``SourceLimits``/``SemanticLimits`` profiles.  The retired
    # ``DocumentSourceLimits``/``ReadError`` family is listed in
    # ``SEMANTIC_FACADE_OWNER_FORBIDDEN_NAMES`` below.
    "pages": frozenset(
        {
            "Body",
            "Document",
            "DocumentStats",
            "Error",
            "LimitKind",
            "OpenOptions",
            "PackageError",
            "Position",
            "Result",
            "Root",
            "Section",
            "SectionSelector",
            "SectionType",
            "SelectorError",
            "SelectorResult",
            "SemanticLimitKind",
            "SemanticLimits",
            "SourceLimitKind",
            "SourceLimits",
            "TextPosition",
            "TextSpan",
        }
    ),
    # ADR 0034: the deck projection is `Presentation`, its
    # reader errors live in `presentation`, and the ingress envelope is the
    # crate `OpenOptions` with checked source and semantic limits.
    "keynote": frozenset(
        {
            "Build",
            "LimitKind",
            "OpenOptions",
            "Presentation",
            "SemanticLimitKind",
            "SemanticLimits",
            "SourceLimitKind",
            "SourceLimits",
            "presentation",
            "Error",
            "Mode",
            "Position",
            "Result",
            "Seconds",
            "Settings",
            "Show",
            "Size",
            "Slide",
            "SlideSelector",
            "SlideSelectorError",
            "SlideSelectorResult",
            "TextPosition",
            "TextSpan",
            "Transition",
        }
    ),
    # ADR 0034.
    "numbers": frozenset(
        {
            "Workbook",
            "ReadError",
            "OpenOptions",
            "SourceLimitKind",
            "SourceLimits",
            "Stats",
            "Error",
            "SemanticLimitKind",
            "SemanticLimits",
            "ReadLimitKind",
            "Result",
            "Sheet",
            "SheetSelector",
            "Table",
            "TableSelector",
            "TableSelectorError",
        }
    ),
}
SEMANTIC_FACADE_FORBIDDEN_NAMES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ArchiveView",
        "ComponentId",
        "Generated",
        "HostPackage",
        "NativeId",
        "ObjectId",
        "Package",
        "PackageError",
        "PackageResult",
        "RawId",
        "RawMessage",
        "SourceCatalog",
        "Wire",
        "WireDescent",
        "WireFieldView",
        "WireLimits",
        "WireView",
    }
)
# Per-owner exceptions to ``SEMANTIC_FACADE_FORBIDDEN_NAMES``.  The Pages
# readers report typed wrong-format, limit and malformed-input failures through
# the leaf ``PackageError``; ``Package`` itself and ``PackageResult`` remain
# forbidden.
SEMANTIC_FACADE_OWNER_ALLOWED_NAMES = {
    "pages": frozenset({"PackageError"}),
}
# Per-owner names that must not reappear in a semantic facade.  For Pages these
# are the ingress types deleted by ADR 0034.
SEMANTIC_FACADE_OWNER_FORBIDDEN_NAMES = {
    "pages": frozenset(
        {
            "DocumentReadOptions",
            "DocumentSourceLimitKind",
            "DocumentSourceLimits",
            "DocumentSourceLimitsError",
            "IoKind",
            "ReadError",
            "ReadLimitKind",
            "SemanticLimitsError",
        }
    ),
}
# Reader helpers in a semantic facade may delegate to the leaf package ingress
# inside a function body (``litchi_pages::Package::open(path).map(...)``).
# Only that exact call shape is removed before the forbidden-name scan, so a
# ``Package`` export, alias or signature is still reported.
SEMANTIC_FACADE_OWNER_PACKAGE_INGRESS_CALL = {
    "pages": re.compile(
        r"\blitchi_pages[ \t\r\n]*::[ \t\r\n]*Package[ \t\r\n]*::"
        r"[ \t\r\n]*(?:open|from_bytes)[ \t\r\n]*\("
    ),
}
SEMANTIC_FACADE_FORBIDDEN_PATH = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"(?:litchi_iwa_(?:archive|common|core|detect|index|package|protos|"
    r"structured|text_wire)|litchi_numbers_wire|buffa|prost|prost_types)"
    r"[ \t\r\n]*::[ \t\r\n]*[A-Za-z_]"
    r"[A-Za-z0-9_]*"
    r"|(?:litchi_pages|litchi_keynote|litchi_numbers)[ \t\r\n]*::"
    r"[ \t\r\n]*(?:package|archive|wire|protos?)(?:[ \t\r\n]*::|\b)"
    r")",
)
SEMANTIC_FACADE_GLOB_REEXPORT = re.compile(
    r"\bpub[ \t\r\n]+use[\s\S]*?\*[ \t\r\n]*;",
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
    if migration_raw:
        # A workspace with no migration debt may retire its last migration
        # host.  While debt remains, however, an empty host set would make the
        # debt unowned and therefore silently unenforceable.
        if not migration_hosts:
            raise PolicyError(
                "migration_hosts must be non-empty while migration_debt exists"
            )
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
    non_host_migration_debt = sorted(
        item.edge for item in migration if item.edge.dependent not in migration_hosts
    )
    if non_host_migration_debt:
        raise PolicyError(
            "migration debt dependents must be migration hosts: "
            + ", ".join(edge.display() for edge in non_host_migration_debt)
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


def audit_iwa_protos_root_alias_source_topology(root: Path = ROOT) -> list[str]:
    """Keep retired generated-schema aliases out of the public crate root."""

    path = root / IWA_PROTOS_FACADE_SOURCE
    if not path.is_file():
        return []

    production_source = _mask_rust_cfg_test_items(
        path.read_text(encoding="utf-8")
    )
    source = _mask_rust_non_code(production_source)
    violations: list[str] = []
    for declaration in IWA_PROTOS_PUBLIC_USE.finditer(source):
        tail = declaration.group("tail")
        tokens = set(re.findall(r"\b[A-Za-z_][A-Za-z0-9_]*\b", tail))
        aliases = RETIRED_IWA_PROTOS_ROOT_ALIASES.intersection(
            tokens
        )
        if "*" in tail and "generated" in tokens:
            aliases = RETIRED_IWA_PROTOS_ROOT_ALIASES
        if not aliases:
            continue
        line_number = source.count("\n", 0, declaration.start()) + 1
        violations.append(
            "retired litchi-iwa-protos generated root aliases returned "
            f"({', '.join(sorted(aliases))}): "
            f"{IWA_PROTOS_FACADE_SOURCE}:{line_number}"
        )
    for declaration in IWA_PROTOS_EXTERNAL_MODULE_DECLARATION.finditer(source):
        alias = declaration.group("name")
        if alias not in RETIRED_IWA_PROTOS_ROOT_ALIASES:
            continue
        line_number = source.count("\n", 0, declaration.start()) + 1
        violations.append(
            "retired litchi-iwa-protos generated root alias returned as a "
            f"public module ({alias}): {IWA_PROTOS_FACADE_SOURCE}:{line_number}"
        )

    return sorted(set(violations))


def audit_litchi_semantic_facade_source_topology(root: Path = ROOT) -> list[str]:
    """Keep the umbrella's semantic iWork namespaces package-free.

    The leaf rustdoc gate follows resolved public types.  This complementary
    source check covers the umbrella's local re-export shape, where a broad
    glob or a package alias could otherwise hide behind ``pages::semantic``,
    ``keynote::semantic``, or ``numbers::semantic`` without changing a leaf's
    own rustdoc graph.
    """

    path = root / FACADE_SOURCE_ROOT / "lib.rs"
    if not path.is_file():
        return []

    source = path.read_text(encoding="utf-8")
    code = _mask_rust_non_code(source)
    violations: list[str] = []

    def balanced_body(opening: int) -> tuple[str, int] | None:
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

    for owner, required in sorted(SEMANTIC_FACADE_REQUIRED_EXPORTS.items()):
        owner_match = re.search(
            rf"(?m)^\s*pub[ \t]+mod[ \t]+{re.escape(owner)}\b",
            code,
        )
        if owner_match is None:
            # Feature-gated modules are absent from a source-only fixture or
            # an intentionally reduced facade; the normal manifest/features
            # audit reports missing owners separately.
            continue
        owner_opening = code.find("{", owner_match.end())
        owner_semicolon = code.find(";", owner_match.end())
        if owner_opening < 0 or (
            owner_semicolon >= 0 and owner_semicolon < owner_opening
        ):
            violations.append(
                f"litchi {owner} semantic facade must be an inline module: "
                f"{path.relative_to(root)}:{source.count(chr(10), 0, owner_match.start()) + 1}"
            )
            continue
        owner_result = balanced_body(owner_opening)
        if owner_result is None:
            violations.append(
                f"litchi {owner} semantic facade has an unbalanced module: "
                f"{path.relative_to(root)}:{source.count(chr(10), 0, owner_match.start()) + 1}"
            )
            continue
        owner_body, owner_body_offset = owner_result
        semantic_match = re.search(
            r"(?m)^\s*pub[ \t]+mod[ \t]+semantic\b",
            owner_body,
        )
        if semantic_match is None:
            violations.append(
                f"litchi {owner} semantic facade is missing `semantic`: "
                f"{path.relative_to(root)}:{source.count(chr(10), 0, owner_match.start()) + 1}"
            )
            continue
        semantic_opening = owner_body.find("{", semantic_match.end())
        semantic_semicolon = owner_body.find(";", semantic_match.end())
        semantic_line = source.count(
            chr(10),
            0,
            owner_body_offset + semantic_match.start(),
        ) + 1
        if semantic_opening < 0 or (
            semantic_semicolon >= 0 and semantic_semicolon < semantic_opening
        ):
            violations.append(
                f"litchi {owner} semantic facade must be an inline module: "
                f"{path.relative_to(root)}:{semantic_line}"
            )
            continue

        # Balance relative to the owner body; use the same scanner with an
        # offset into the full masked source to catch nested braces correctly.
        semantic_start = owner_body_offset + semantic_opening
        semantic_result = balanced_body(semantic_start)
        if semantic_result is None:
            violations.append(
                f"litchi {owner} semantic facade has an unbalanced module: "
                f"{path.relative_to(root)}:{semantic_line}"
            )
            continue
        semantic_body, semantic_body_offset = semantic_result
        ingress_call = SEMANTIC_FACADE_OWNER_PACKAGE_INGRESS_CALL.get(owner)
        scanned_body = (
            semantic_body
            if ingress_call is None
            else ingress_call.sub(lambda match: " " * len(match.group(0)), semantic_body)
        )
        identifiers = {
            match.group(1) for match in RUST_IDENTIFIER.finditer(scanned_body)
        }
        missing = required - identifiers
        if missing:
            violations.append(
                f"litchi {owner} semantic facade is missing exports: "
                + ", ".join(sorted(missing))
                + f": {path.relative_to(root)}:{semantic_line}"
            )

        forbidden = (
            SEMANTIC_FACADE_FORBIDDEN_NAMES
            - SEMANTIC_FACADE_OWNER_ALLOWED_NAMES.get(owner, frozenset())
        ) | SEMANTIC_FACADE_OWNER_FORBIDDEN_NAMES.get(owner, frozenset())
        for identifier in sorted(identifiers & forbidden):
            violations.append(
                f"litchi {owner} semantic facade exposes forbidden `{identifier}`: "
                f"{path.relative_to(root)}:{semantic_line}"
            )
        for match in SEMANTIC_FACADE_FORBIDDEN_PATH.finditer(semantic_body):
            line = source.count(
                chr(10),
                0,
                semantic_body_offset + match.start(),
            ) + 1
            violations.append(
                f"litchi {owner} semantic facade exposes physical path: "
                f"{path.relative_to(root)}:{line}"
            )
        for match in SEMANTIC_FACADE_GLOB_REEXPORT.finditer(semantic_body):
            line = source.count(
                chr(10),
                0,
                semantic_body_offset + match.start(),
            ) + 1
            violations.append(
                f"litchi {owner} semantic facade exposes a glob re-export: "
                f"{path.relative_to(root)}:{line}"
            )

    return sorted(set(violations))


def audit_iwork_example_source_topology(root: Path = ROOT) -> list[str]:
    """Keep user-facing iWork examples on semantic, selector-first APIs.

    Workspace examples are user-facing documentation and must not teach
    callers to use the retired editor-host vocabulary or print native object
    handles.  The
    aggregate ``litchi`` example is checked by the same narrow gate so a
    future example cannot accidentally bypass the facade.
    """

    violations: list[str] = []
    for relative in IWORK_EXAMPLE_SOURCES:
        path = root / relative
        if not path.is_file():
            continue
        source = _mask_rust_non_code(path.read_text(encoding="utf-8"))
        for pattern, label in (
            (IWORK_EXAMPLE_LEGACY_API, "legacy editor-host API"),
            (IWORK_EXAMPLE_NATIVE_ID_CALL, "native object identifier call"),
        ):
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start()) + 1
                violations.append(
                    f"iWork example exposes {label}: "
                    f"{relative}:{line_number}"
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


def _mask_rust_comments(source: str) -> str:
    """Mask Rust comments while retaining ordinary string-token contents.

    A small number of source ratchets need to inspect a command literal (for
    example, an example's ``"move"`` operation) rather than only identifiers.
    Reusing :func:`_mask_rust_non_code` would hide that literal, while scanning
    the raw source would let comments or raw-string documentation satisfy the
    ratchet.  Preserve ordinary quoted tokens, mask comments and raw strings,
    and keep offsets/newlines stable for diagnostics.
    """

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
            cursor = quoted_string_end(cursor)
            continue
        cursor += 1

    return "".join(masked)


RUST_CFG_TEST_ATTRIBUTE = re.compile(
    r"^[ \t]*#[ \t]*\[[ \t]*cfg[ \t]*\([ \t]*test[ \t]*\)[ \t]*\]",
    re.MULTILINE,
)
# A source ratchet must distinguish an item that is test-only from one that
# merely has a test configuration as one possible branch.  In particular,
# ``cfg(any(test, feature = "oracle"))`` is production-reachable when the
# feature is enabled and must stay visible to the production audit, while
# ``cfg(all(test, feature = "oracle"))`` is safe to mask.  Keep the old exact
# matcher above for callers that need its historical spelling; the item
# masker below uses this balanced scanner so nested cfg expressions cannot
# bypass it by adding whitespace or a second predicate.
RUST_CFG_ATTRIBUTE_START = re.compile(
    r"^[ \t]*#[ \t]*\[[ \t]*cfg\b",
    re.MULTILINE,
)
# Keep the cfg(test) item recognizer deliberately broader than the public
# declaration scanner.  A test-only item can be crate-visible, an ABI
# qualified function, an extern block, or a macro.  Masking only ``pub fn``
# would leave a production ratchet blind to a decoy hidden behind, for
# example, ``#[cfg(test)] pub unsafe extern "C" fn ...`` or
# ``#[cfg(test)] macro_rules! ...``.  The input is already passed through
# ``_mask_rust_non_code``; the quoted ABI therefore appears as whitespace and
# cannot make this recognizer inspect string contents.
RUST_CFG_TEST_ITEM = re.compile(
    r"(?:pub(?:[ \t\r\n]*\([^()\r\n]*\))?[ \t\r\n]+)?"
    r"(?:(?:unsafe|async|const|default)[ \t\r\n]+)*"
    r"(?:(?:extern(?:[ \t\r\n]+\"[^\"]*\")?)[ \t\r\n]+)?"
    r"(?P<kind>use|fn|mod|impl|struct|enum|trait|type|const|static|union|"
    r"extern|macro_rules|macro)\b"
)


def _cfg_expression_requires_test(expression: str) -> bool:
    """Return whether a cfg expression can only be true under ``test``.

    The parser intentionally models only the boolean cfg meta-expression
    operators.  Values such as ``feature = "oracle"`` are non-test atoms;
    their string payload is irrelevant.  For a conjunction, one operand that
    requires test is enough.  For a disjunction, every alternative must
    require test.  Negation never proves a test-only item.  Unknown/malformed
    forms are conservative and remain visible to the production scan.
    """

    tokens = re.findall(r"[A-Za-z_][A-Za-z0-9_]*|[(),=]", expression)
    cursor = 0

    def parse_term() -> bool:
        nonlocal cursor
        if cursor >= len(tokens):
            return False
        atom = tokens[cursor]
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", atom):
            cursor += 1
            return False
        cursor += 1

        if cursor < len(tokens) and tokens[cursor] == "(":
            cursor += 1
            values: list[bool] = []
            malformed = False
            while cursor < len(tokens) and tokens[cursor] != ")":
                before = cursor
                values.append(parse_term())
                if cursor == before:
                    malformed = True
                    cursor += 1
                if cursor < len(tokens) and tokens[cursor] == ",":
                    cursor += 1
                elif cursor < len(tokens) and tokens[cursor] != ")":
                    malformed = True
                    # Make forward progress without treating the remainder
                    # as a nested production item.
                    cursor += 1
            if cursor >= len(tokens) or tokens[cursor] != ")":
                return False
            cursor += 1
            if malformed or not values:
                return False
            if atom == "all":
                return any(values)
            if atom == "any":
                return all(values)
            # ``not(test)`` is production-reachable outside tests.  Unknown
            # operators are likewise not safe to hide.
            return False

        if cursor < len(tokens) and tokens[cursor] == "=":
            # Consume a cfg key/value atom (usually ``feature = \"...\"``).
            # The value is intentionally opaque to this test-only classifier.
            cursor += 1
            while cursor < len(tokens) and tokens[cursor] not in {",", ")"}:
                cursor += 1
            return False
        return atom == "test"

    result = parse_term()
    # An unconsumed token sequence means this is not a form we can prove
    # test-only.  This also keeps invalid/composite syntax fail-closed.
    return result and cursor == len(tokens)


def _rust_cfg_test_only_attribute_spans(source: str) -> tuple[tuple[int, int], ...]:
    """Return balanced ``cfg`` attributes whose condition requires ``test``."""

    code = _mask_rust_non_code(source)
    spans: list[tuple[int, int]] = []
    for candidate in RUST_CFG_ATTRIBUTE_START.finditer(code):
        opening = code.find("[", candidate.start(), candidate.end())
        if opening < 0:
            continue
        bracket_depth = 1
        cursor = opening + 1
        while cursor < len(code) and bracket_depth:
            if code[cursor] == "[":
                bracket_depth += 1
            elif code[cursor] == "]":
                bracket_depth -= 1
            cursor += 1
        if bracket_depth:
            continue
        closing = cursor
        paren_open = code.find("(", candidate.end(), closing)
        if paren_open < 0:
            continue
        paren_depth = 1
        paren_cursor = paren_open + 1
        while paren_cursor < closing and paren_depth:
            if code[paren_cursor] == "(":
                paren_depth += 1
            elif code[paren_cursor] == ")":
                paren_depth -= 1
            paren_cursor += 1
        if paren_depth:
            continue
        expression = code[paren_open + 1 : paren_cursor - 1]
        if _cfg_expression_requires_test(expression):
            spans.append((candidate.start(), closing))
    return tuple(spans)


def _mask_rust_cfg_test_items(source: str) -> str:
    """Mask individual ``cfg(test)`` Rust items without hiding later production.

    Some large modules interleave test-only imports and reference helpers with
    production items. Truncating at the first ``cfg(test)`` attribute therefore
    creates a ratchet bypass. This small item scanner preserves offsets and
    newlines while blanking each gated item, including its body when present.
    """

    code = _mask_rust_non_code(source)
    masked = list(source)

    def blank(start: int, end: int) -> None:
        for offset in range(start, end):
            if masked[offset] != "\n":
                masked[offset] = " "

    def skip_whitespace(cursor: int) -> int:
        while cursor < len(code) and code[cursor].isspace():
            cursor += 1
        return cursor

    def skip_attribute(cursor: int) -> int | None:
        if cursor >= len(code) or code[cursor] != "#":
            return None
        cursor += 1
        cursor = skip_whitespace(cursor)
        if cursor >= len(code) or code[cursor] != "[":
            return None
        depth = 1
        cursor += 1
        while cursor < len(code) and depth:
            if code[cursor] == "[":
                depth += 1
            elif code[cursor] == "]":
                depth -= 1
            cursor += 1
        return cursor if depth == 0 else len(code)

    def item_end(cursor: int, kind: str) -> int:
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
            elif parentheses == 0 and brackets == 0:
                if character == ";":
                    return cursor + 1
                if character == "{":
                    depth = 1
                    cursor += 1
                    while cursor < len(code) and depth:
                        if code[cursor] == "{":
                            depth += 1
                        elif code[cursor] == "}":
                            depth -= 1
                        cursor += 1
                    after = skip_whitespace(cursor)
                    if (
                        kind in {"const", "static"}
                        and after < len(code)
                        and code[after] == ";"
                    ):
                        return after + 1
                    return cursor
            cursor += 1
        return len(code)

    for attribute_start, attribute_end in _rust_cfg_test_only_attribute_spans(
        source
    ):
        if all(
            character in {" ", "\n"}
            for character in masked[attribute_start:attribute_end]
        ):
            continue
        cursor = skip_whitespace(attribute_end)
        while True:
            next_cursor = skip_attribute(cursor)
            if next_cursor is None:
                break
            cursor = skip_whitespace(next_cursor)
        item = RUST_CFG_TEST_ITEM.match(code, cursor)
        if item is None:
            # Leave an unrecognized item visible so a future production marker
            # cannot be hidden by an item shape this small scanner does not yet
            # understand.
            blank(attribute_start, attribute_end)
            continue
        blank(attribute_start, item_end(item.end(), item.group("kind")))

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


def _rust_mask_nested_function_items(source: str) -> str:
    """Mask nested Rust function items while preserving offsets and newlines."""

    code = _mask_rust_non_code(source)
    masked = list(code)

    def blank(start: int, end: int) -> None:
        for offset in range(start, end):
            if masked[offset] != "\n":
                masked[offset] = " "

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
        if depth == 0:
            blank(declaration.start(), cursor)

    return "".join(masked)


def _rust_top_level_function_bodies(source: str) -> dict[str, tuple[str, int]]:
    """Return module-level Rust function bodies and their source offsets."""

    code = _mask_rust_non_code(source)
    functions: dict[str, tuple[str, int]] = {}
    for declaration in RUST_FUNCTION_DECLARATION.finditer(code):
        if code.count("{", 0, declaration.start()) != code.count("}", 0, declaration.start()):
            continue
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
        name = declaration.group(1)
        functions.setdefault(
            name,
            (
                _rust_mask_nested_function_items(code[opening + 1 : cursor - 1]),
                opening + 1,
            ),
        )
    return functions


def _rust_named_function_body(source: str, name: str) -> tuple[str, int] | None:
    """Return one top-level Rust function body and its source offset."""

    return _rust_top_level_function_bodies(source).get(name)


def _rust_any_function_body(source: str, name: str) -> str | None:
    """Return one Rust function body, including methods inside ``impl`` blocks.

    The older named-body helper deliberately limits itself to module-level
    functions.  A few compatibility ratchets need to inspect an inherent
    method's routing body as well, so use the masked source only for finding
    braces and return the corresponding original source slice.
    """

    code = _mask_rust_non_code(source)
    declaration = re.search(
        rf"\bfn[ \t\r\n]+(?:r#)?{re.escape(name)}\b", code
    )
    if declaration is None:
        return None
    opening = code.find("{", declaration.end())
    if opening < 0:
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
    return source[opening + 1 : cursor - 1]


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


def _rust_brace_depth_at(code: str, offset: int) -> int:
    """Return the lexical Rust brace depth immediately before ``offset``."""

    return code.count("{", 0, offset) - code.count("}", 0, offset)


def _rust_balanced_delimited_end(
    code: str, opening: int, opener: str = "{", closer: str = "}"
) -> int | None:
    """Return the exclusive end of one balanced delimiter pair."""

    if opening < 0 or opening >= len(code) or code[opening] != opener:
        return None
    depth = 1
    cursor = opening + 1
    while cursor < len(code) and depth:
        if code[cursor] == opener:
            depth += 1
        elif code[cursor] == closer:
            depth -= 1
        cursor += 1
    return cursor if depth == 0 else None


def _rust_root_level_matches(
    source: str, pattern: re.Pattern[str]
) -> tuple[re.Match[str], ...]:
    """Find pattern matches at the crate/module root, ignoring Rust trivia."""

    code = _mask_rust_non_code(source)
    return tuple(
        match
        for match in pattern.finditer(code)
        if _rust_brace_depth_at(code, match.start()) == 0
    )


def _rust_root_level_module_declarations(
    source: str, names: frozenset[str]
) -> tuple[tuple[str, str | None, str, str, int], ...]:
    """Return root module declarations and their ``;``/inline-body shape.

    The returned tuple is ``(name, visibility, shape, body, line)``.  ``shape``
    is ``external`` for ``mod name;``, ``inline`` for a balanced module body,
    and ``invalid`` when the declaration is truncated or has another token
    after its name.  Nested decoy modules are deliberately excluded.
    """

    if not names:
        return ()
    code = _mask_rust_non_code(source)
    pattern = re.compile(
        r"(?m)^[ \t]*(?P<visibility>pub(?:[ \t]*\([^()\r\n]*\))?[ \t\r\n]+)?"
        r"mod[ \t\r\n]+(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
    )
    declarations: list[tuple[str, str | None, str, str, int]] = []
    for match in pattern.finditer(code):
        name = match.group("name")
        if name not in names or _rust_brace_depth_at(code, match.start()) != 0:
            continue
        cursor = match.end()
        while cursor < len(code) and code[cursor].isspace():
            cursor += 1
        shape = "invalid"
        body = ""
        if cursor < len(code) and code[cursor] == ";":
            shape = "external"
        elif cursor < len(code) and code[cursor] == "{":
            end = _rust_balanced_delimited_end(code, cursor)
            if end is not None:
                shape = "inline"
                body = source[cursor + 1 : end - 1]
        declarations.append(
            (
                name,
                match.group("visibility"),
                shape,
                body,
                code.count("\n", 0, match.start()) + 1,
            )
        )
    return tuple(declarations)


def _rust_generated_module_include_is_valid(body: str) -> bool:
    """Check that an inline generated module includes one OUT_DIR Buffa file."""

    code = _mask_rust_non_code(body)
    includes = list(re.finditer(r"\binclude!\s*\(", code))
    if len(includes) != 1:
        return False
    include = includes[0]
    end = _rust_balanced_delimited_end(
        code, code.find("(", include.start(), include.end()), "(", ")"
    )
    if end is None:
        return False
    # Inspect the actual macro call after comments are removed.  Ordinary
    # strings are retained here so the generated path can be validated, while
    # a comment/string-only fake include was already excluded by ``code``.
    call = _mask_rust_comments(body[include.start() : end])
    return re.fullmatch(
        r"\s*include!\s*\(\s*concat!\s*\(\s*env!\s*\(\s*\"OUT_DIR\"\s*\)"
        r"\s*,\s*\"[^\"]*buffa[^\"]*\"\s*\)\s*\)\s*",
        call,
        re.IGNORECASE | re.DOTALL,
    ) is not None


def _rust_statement_end(code: str, start: int) -> int:
    """Find a top-level semicolon or item-body opening after ``start``."""

    parentheses = 0
    brackets = 0
    cursor = start
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
                return cursor + 1
            if character == "{":
                return cursor
        cursor += 1
    return len(code)


def _rust_root_generated_public_leaks(
    source: str, modules: frozenset[str]
) -> tuple[int, ...]:
    """Return root-line numbers exposing generated modules or aliases publicly."""

    if not modules:
        return ()
    code = _mask_rust_non_code(source)
    module_pattern = re.compile(
        r"(?<![A-Za-z0-9_#])(?:" + "|".join(
            re.escape(module) for module in sorted(modules, key=len, reverse=True)
        ) + r")(?![A-Za-z0-9_#])"
    )
    aliases: set[str] = set()
    leaking_lines: set[int] = set()
    use_pattern = re.compile(
        r"(?m)^[ \t]*(?P<visibility>pub(?:[ \t]*\([^()\r\n]*\))?[ \t\r\n]+)?"
        r"use\b"
    )
    public_item_pattern = re.compile(
        r"(?m)^[ \t]*pub(?:[ \t]*\([^()\r\n]*\))?[ \t\r\n]+"
        r"(?P<kind>type|struct|enum|trait|union|const|static|fn)\b"
    )

    for match in use_pattern.finditer(code):
        if _rust_brace_depth_at(code, match.start()) != 0:
            continue
        end = _rust_statement_end(code, match.end())
        declaration = code[match.start() : end]
        direct = module_pattern.search(declaration) is not None
        for alias in re.finditer(
            r"\b(?:r#)?[A-Za-z_][A-Za-z0-9_]*\b[ \t\r\n]+as[ \t\r\n]+"
            r"(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b",
            declaration,
        ):
            prefix = declaration[: alias.start("name")]
            if module_pattern.search(prefix) is not None:
                aliases.add(alias.group("name"))
        public = match.group("visibility") is not None
        if public and (direct or any(
            re.search(
                rf"(?<![A-Za-z0-9_#]){re.escape(alias)}(?![A-Za-z0-9_#])",
                declaration,
            )
            for alias in aliases
        )):
            leaking_lines.add(code.count("\n", 0, match.start()) + 1)

    for match in public_item_pattern.finditer(code):
        if _rust_brace_depth_at(code, match.start()) != 0:
            continue
        end = _rust_statement_end(code, match.end())
        if match.group("kind") in {"struct", "enum", "trait", "union"}:
            balanced_end = _rust_balanced_delimited_end(code, end)
            if balanced_end is not None:
                end = balanced_end
        declaration = code[match.start() : end]
        if module_pattern.search(declaration) is not None or any(
            re.search(
                rf"(?<![A-Za-z0-9_#]){re.escape(alias)}(?![A-Za-z0-9_#])",
                declaration,
            )
            for alias in aliases
        ):
            leaking_lines.add(code.count("\n", 0, match.start()) + 1)
    return tuple(sorted(leaking_lines))


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


def _rust_public_methods_in_impl(
    source: str, type_name: str
) -> list[tuple[str, str, int]]:
    """Return public inherent-method declarations for one Rust type.

    The existing declaration helpers intentionally do not retain item
    ownership, which is enough for most facade scans but would let a free
    ``write_to`` satisfy a ``Package::write_to`` requirement.  This narrow
    helper records only public methods inside an inherent ``impl Type`` block,
    preserving the source line and the declaration signature for return-type
    checks.  Callers should mask ``cfg(test)`` items before invoking it.
    """

    code = _mask_rust_non_code(source)
    impl_ranges: list[tuple[int, int]] = []
    for implementation in RUST_IMPL_DECLARATION.finditer(code):
        if code.count("{", 0, implementation.start()) != code.count(
            "}", 0, implementation.start()
        ):
            continue
        parentheses = 0
        brackets = 0
        opening = -1
        cursor = implementation.end()
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
                opening = cursor
                break
            cursor += 1
        if opening < 0:
            continue
        header = re.sub(r"\s+", " ", code[implementation.start() : opening]).strip()
        if re.match(
            rf"^impl(?:\s*<[^>{{}}]*>)?\s+{re.escape(type_name)}"
            rf"(?:\s*<[^>{{}}]*>)?(?:\s+where .*)?$",
            header,
        ) is None:
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
            impl_ranges.append((opening + 1, cursor - 1))

    method = re.compile(
        r"(?<![A-Za-z0-9_#])pub(?![ \t\r\n]*\()[ \t\r\n]+"
        r"(?:(?:unsafe|async|const)[ \t\r\n]+)*fn[ \t\r\n]+"
        r"(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
    )
    methods: list[tuple[str, str, int]] = []
    for match in method.finditer(code):
        if not any(start <= match.start() < end for start, end in impl_ranges):
            continue
        parentheses = 0
        brackets = 0
        cursor = match.end()
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
            elif not parentheses and not brackets and character in "{;":
                break
            cursor += 1
        methods.append(
            (
                match.group("name"),
                code[match.start() : cursor],
                source.count("\n", 0, match.start()) + 1,
            )
        )
    return methods


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


def _rust_public_enum_variants(source: str, name: str) -> frozenset[str]:
    """Return top-level variants of one unrestricted public enum."""

    code = _mask_rust_non_code(source)
    declaration_pattern = re.compile(
        rf"(?<![A-Za-z0-9_#])pub[ \t\r\n]+enum[ \t\r\n]+"
        rf"(?:r#)?{re.escape(name)}\b",
    )
    declaration = next(
        (
            candidate
            for candidate in declaration_pattern.finditer(code)
            if code.count("{", 0, candidate.start())
            == code.count("}", 0, candidate.start())
        ),
        None,
    )
    if declaration is None:
        return frozenset()
    opening = code.find("{", declaration.end())
    if opening < 0:
        return frozenset()
    depth = 1
    cursor = opening + 1
    while cursor < len(code) and depth:
        if code[cursor] == "{":
            depth += 1
        elif code[cursor] == "}":
            depth -= 1
        cursor += 1
    if depth:
        return frozenset()

    body = code[opening + 1 : cursor - 1]
    variants: set[str] = set()
    start = 0
    parentheses = 0
    brackets = 0
    braces = 0
    for index, character in enumerate(body + ","):
        if character == "(":
            parentheses += 1
        elif character == ")" and parentheses:
            parentheses -= 1
        elif character == "[":
            brackets += 1
        elif character == "]" and brackets:
            brackets -= 1
        elif character == "{":
            braces += 1
        elif character == "}" and braces:
            braces -= 1
        elif character == "," and not (parentheses or brackets or braces):
            segment = body[start:index]
            start = index + 1
            while True:
                attribute = re.match(r"^[ \t\r\n]*#\[[^\]]*\]", segment)
                if attribute is None:
                    break
                segment = segment[attribute.end() :]
            variant = RUST_IDENTIFIER.search(segment)
            if variant is not None:
                variants.add(variant.group(1))
    return frozenset(variants)


def _rust_public_module_body(source: str, name: str) -> str | None:
    """Return one unrestricted public inline module body, ignoring trivia."""

    code = _mask_rust_non_code(source)
    declaration = re.search(
        rf"(?<![A-Za-z0-9_#])pub[ \t\r\n]+mod[ \t\r\n]+"
        rf"(?:r#)?{re.escape(name)}\b",
        code,
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
    return code[opening + 1 : cursor - 1]


def _iwork_public_leak(identifier: str) -> str | None:
    """Classify physical vocabulary forbidden in focused iWork facades."""

    if identifier.startswith("litchi_iwa_"):
        return "archive/IWA type"
    if identifier in {"buffa", "prost", "prost_types"}:
        return "protobuf type"
    if identifier in {"HostPackage", "SourceCatalog"}:
        return "archive/IWA type"
    words: list[str] = []
    for part in identifier.split("_"):
        words.extend(word.lower() for word in CAMEL_CASE_WORD.findall(part))
    if any(
        words[index] == "source" and words[index + 1] in {"byte", "bytes"}
        for index in range(len(words) - 1)
    ):
        return "raw source bytes"
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
    if "generated" in words:
        return "generated type"
    return None


def _rust_canonical_exports(
    source: str, names: frozenset[str]
) -> frozenset[str]:
    """Return exact names publicly defined or reexported in one Rust scope."""

    exported: set[str] = set()
    for declaration, _ in _rust_public_declarations(source):
        identifiers = [
            match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
        ]
        if identifiers[:2] == ["pub", "use"]:
            code = declaration
            for alias in re.finditer(
                r"(?<![A-Za-z0-9_#])(?:r#)?([A-Za-z_][A-Za-z0-9_]*)"
                r"[ \t\r\n]+as[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)",
                code,
            ):
                original, public_name = alias.groups()
                if public_name in names:
                    exported.add(public_name)
                if original in names:
                    code = (
                        code[: alias.start(1)]
                        + (" " * len(original))
                        + code[alias.end(1) :]
                    )
            exported.update(
                match.group(1)
                for match in RUST_IDENTIFIER.finditer(code)
                if match.group(1) in names
            )
            continue
        item = re.match(
            r"^[ \t]*pub[ \t\r\n]+(?:struct|enum|type|trait|union)"
            r"[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)",
            declaration,
        )
        if item is not None and item.group(1) in names:
            exported.add(item.group(1))
    return frozenset(exported)


def _rust_root_canonical_exports(
    source: str, names: frozenset[str]
) -> frozenset[str]:
    """Return canonical public names declared/reexported at the root scope."""

    exported: set[str] = set()
    for name in names:
        escaped = re.escape(name)
        patterns = (
            re.compile(
                rf"(?m)^[ \t]*pub[ \t\r\n]+(?:struct|enum|type|trait|union)"
                rf"[ \t\r\n]+(?:r#)?{escaped}\b"
            ),
            re.compile(
                rf"(?m)^[ \t]*pub[ \t\r\n]+use[^;\n]*"
                rf"(?:r#)?{escaped}\b"
            ),
        )
        if any(_rust_root_level_matches(source, pattern) for pattern in patterns):
            exported.add(name)
    return frozenset(exported)


def _rust_attribute_block_before(source: str, offset: int) -> str:
    """Return the contiguous Rust attribute block immediately before an item."""

    lines = source[:offset].splitlines(keepends=True)
    index = len(lines) - 1
    while index >= 0 and not lines[index].strip():
        index -= 1
    if index < 0:
        return ""
    stripped = lines[index].strip()
    if not (stripped.startswith("#[") or stripped.endswith("]")):
        return ""

    end = index + 1
    start = index
    while index >= 0:
        stripped = lines[index].strip()
        if stripped.startswith("#["):
            start = index
            index -= 1
            while index >= 0 and not lines[index].strip():
                index -= 1
            if index >= 0 and lines[index].strip().startswith("#["):
                continue
            break
        index -= 1
    return "".join(lines[start:end])


def audit_iwa_common_formula_render_ownership(
    root: Path = ROOT,
) -> list[str]:
    """Keep the shared formula arena independent and singular.

    The archive-free common module is the only owner of the expression handles
    and renderer arena.  Readers may reach it directly while the migration is
    staged, or through the low-level ``litchi-numbers-wire`` formula adapter;
    the companion wire audit checks that the latter owns the event visitor.
    This audit deliberately checks declarations and import boundaries rather
    than private field names, so a storage optimization can evolve without
    weakening the ownership rule.
    """

    violations: list[str] = []
    common_path = root / IWA_COMMON_FORMULA_RENDER_SOURCE
    if not common_path.is_file():
        return [
            "shared formula renderer owner is missing: "
            f"{IWA_COMMON_FORMULA_RENDER_SOURCE}"
        ]

    common_raw = common_path.read_text(encoding="utf-8")
    common_code = _mask_rust_non_code(common_raw)
    for label, declaration in IWA_COMMON_FORMULA_RENDER_REQUIRED_DECLARATIONS:
        if declaration.search(common_code) is None:
            violations.append(
                "shared formula renderer owner is missing "
                f"{label}: {IWA_COMMON_FORMULA_RENDER_SOURCE}"
            )

    for statement in IWA_COMMON_FORMULA_RENDER_IMPORT_STATEMENT.finditer(common_code):
        forbidden = IWA_COMMON_FORMULA_RENDER_FORBIDDEN_IMPORT.search(statement.group())
        if forbidden is None:
            continue
        line_number = (
            common_code.count("\n", 0, forbidden.start() + statement.start()) + 1
        )
        violations.append(
            "shared formula renderer imports a format/protobuf owner "
            f"{forbidden.group()}: {IWA_COMMON_FORMULA_RENDER_SOURCE}:{line_number}"
        )

    for relative in IWA_FORMULA_RENDER_READER_SOURCES:
        path = root / relative
        if not path.is_file():
            # A retired reader is an accepted end state once its remaining
            # format peer has moved to the common arena.  Keep auditing files
            # that still exist so reintroduced local arenas cannot bypass the
            # ownership ratchet.
            continue
        raw_source = path.read_text(encoding="utf-8")
        production_code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
        if IWA_FORMULA_RENDER_COMMON_IMPORT.search(production_code) is None:
            violations.append(
                "formula renderer reader is missing the shared arena import: "
                f"{relative}"
            )
        for label, declaration in IWA_FORMULA_RENDER_LOCAL_ARENA_DECLARATIONS:
            for match in declaration.finditer(production_code):
                line_number = production_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "formula renderer reader redeclares shared arena "
                    f"{label}: {relative}:{line_number}"
                )

    return sorted(set(violations))


# The migration host's feature and hidden seams, now removed. A leaf
# that needs one of these again gets a `pub(crate)` item or a glossary-checked
# public API instead.
IWA_RETIRED_HOST_FEATURE = "internal-iwork-source"
IWA_RETIRED_HOST_SEAMS = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"__CatalogBodyTableAppearanceSource|__CatalogTableAppearanceSource|"
    r"__catalog_body_table_appearance|__catalog_body_table_style_edges|"
    r"__catalog_table_appearance|__catalog_table_style_edges|"
    r"__clear_table_sort_order_from_bytes_for_compatibility|"
    r"__compatibility_tables_from_prepared_source|"
    r"__decode_image_adjustments_payload|__decode_movie_playback_payload|"
    r"__decode_slide_drawable_projection|__decode_slide_node_projection|"
    r"__from_archive|__from_prepared_source|"
    r"__from_prepared_source_with_semantic_limits|__from_shared_archives|"
    r"__from_shared_catalog|__from_shared_source_with_options|"
    r"__invalidate_slide_preview|__is_litchi_source_built_compatibility|"
    r"__is_valid_pages_text_storage|__move_table_from_bytes_for_compatibility|"
    r"__movie_playback_settings|__pages_body_storage_discovery|"
    r"__presentation_from_prepared_source|__rewrite_movie_playback_payload|"
    r"__rewrite_movie_playback_settings|"
    r"__semantic_document_from_prepared_source|"
    r"__set_table_sort_order_from_bytes_for_compatibility|"
    r"__slide_table_appearances|__source_owner|"
    r"__table_appearance_from_source_built|"
    r"__table_cell_comment_from_bytes_for_compatibility(?:_with_options)?|"
    r"__table_cell_comment_replies_from_bytes_for_compatibility(?:_with_options)?|"
    r"__table_sort_order_from_bytes_for_compatibility|"
    r"__transition_settings_from_source|"
    r"OwnedSourceCatalog|SourceBuiltAppearancePayload"
    r")(?![A-Za-z0-9_])"
)


def audit_iwa_retired_host_seams(root: Path = ROOT) -> list[str]:
    """Keep the deleted migration host's feature and hidden seams out of crates."""

    crates = root / "crates"
    if not crates.is_dir():
        return []
    violations: list[str] = []
    for manifest in sorted(crates.glob("*/Cargo.toml")):
        if IWA_RETIRED_HOST_FEATURE in manifest.read_text(encoding="utf-8"):
            violations.append(
                f"retired iWork host feature {IWA_RETIRED_HOST_FEATURE}: "
                f"{manifest.relative_to(root)}"
            )
    for source_path in sorted(crates.glob("*/**/*.rs")):
        source = source_path.read_text(encoding="utf-8", errors="replace")
        if IWA_RETIRED_HOST_FEATURE in source:
            violations.append(
                f"retired iWork host feature {IWA_RETIRED_HOST_FEATURE}: "
                f"{source_path.relative_to(root)}"
            )
        for match in IWA_RETIRED_HOST_SEAMS.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                f"retired iWork host seam {match.group(0)}: "
                f"{source_path.relative_to(root)}:{line_number}"
            )
    return violations


def audit_iwa_shared_media_playback_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Retire the shared playback host after focused owners are wired.

    The physical archive lookup and transaction staging remain in the narrow
    format adapters.  Scalar playback projection and source-preserving rewrite
    belong to the corresponding Numbers, Pages, or Keynote package, where the
    strict Buffa codec is available.  This audit follows only those concrete
    seams and their known host adapters; it does not reject unrelated
    ``MovieArchive`` graph materialization used for geometry or creation.
    """

    owner_paths = {
        ecosystem: root / path
        for ecosystem, path in IWA_MEDIA_PLAYBACK_OWNER_SOURCES.items()
    }
    package_paths = {
        ecosystem: root / path
        for ecosystem, path in IWA_MEDIA_PLAYBACK_OWNER_PACKAGE_SOURCES.items()
    }
    codec_path = root / IWA_MEDIA_PLAYBACK_CODEC_SOURCE
    if not codec_path.is_file() or not all(path.is_file() for path in owner_paths.values()):
        return []

    violations: list[str] = []
    owner_sources: dict[str, str] = {}
    owner_codes: dict[str, str] = {}
    owner_raw_codes: dict[str, str] = {}
    for ecosystem, path in owner_paths.items():
        raw_source = path.read_text(encoding="utf-8")
        source = _mask_rust_cfg_test_items(raw_source)
        owner_sources[ecosystem] = source
        owner_codes[ecosystem] = _mask_rust_non_code(source)
        owner_raw_codes[ecosystem] = _mask_rust_non_code(raw_source)

    for ecosystem, owner_path in owner_paths.items():
        owner_code = owner_codes[ecosystem]

        package_source = (
            _mask_rust_cfg_test_items(package_paths[ecosystem].read_text(encoding="utf-8"))
            if package_paths[ecosystem].is_file()
            else ""
        )
        package_code = _mask_rust_comments(package_source)
        if IWA_MEDIA_PLAYBACK_OWNER_MODULES[ecosystem].search(package_code) is None:
            violations.append(
                f"focused litchi-{ecosystem.lower()} movie-playback owner module is missing: "
                f"{IWA_MEDIA_PLAYBACK_OWNER_PACKAGE_SOURCES[ecosystem]}"
            )

        if ecosystem == "Keynote":
            seam_sources = [
                (owner_raw_codes[ecosystem], owner_paths[ecosystem]),
            ]
            if package_paths[ecosystem].is_file():
                seam_sources.append(
                    (
                        _mask_rust_non_code(
                            package_paths[ecosystem].read_text(encoding="utf-8")
                        ),
                        package_paths[ecosystem],
                    )
                )
            for seam_source, seam_path in seam_sources:
                for match in IWA_KEYNOTE_MOVIE_PLAYBACK_RETIRED_SEAM_REFERENCE.finditer(
                    seam_source
                ):
                    line_number = seam_source.count("\n", 0, match.start()) + 1
                    violations.append(
                        "retired focused litchi-keynote movie-playback hidden seam "
                        f"{IWA_KEYNOTE_MOVIE_PLAYBACK_RETIRED_SEAM}: "
                        f"{seam_path.relative_to(root)}:{line_number}"
                    )

        if "movie_playback_codec" not in owner_code:
            violations.append(
                f"focused litchi-{ecosystem.lower()} movie-playback owner must use "
                f"movie_playback_codec: {owner_paths[ecosystem]}"
            )
        if re.search(r"\b(?:decode|rewrite)_movie_playback(?:\b|_)", owner_code) is None:
            violations.append(
                f"focused litchi-{ecosystem.lower()} movie-playback owner is missing "
                f"strict codec routing: {owner_paths[ecosystem]}"
            )
        for match in IWA_MEDIA_PLAYBACK_GENERATED_LEAK.finditer(owner_code):
            line_number = owner_code.count("\n", 0, match.start()) + 1
            violations.append(
                f"focused litchi-{ecosystem.lower()} movie-playback owner retains "
                f"generated playback decoding {match.group(0).strip()}: "
                f"{owner_paths[ecosystem]}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_shared_image_adjustments_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Retire the shared image-adjustment host after focused owners land.

    The host retains archive lookup and transaction staging, while each
    format package projects the three semantic inspector controls through the
    neutral bounded Buffa codec.  The source payload remains authoritative, so
    this audit checks the concrete seams and rejects generated archive ingress
    without banning unrelated image-graph work elsewhere in the editor.
    """

    owner_paths = {
        ecosystem: root / path
        for ecosystem, path in IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES.items()
    }
    if not (root / IWA_IMAGE_ADJUSTMENTS_CODEC_SOURCE).is_file() or not all(
        path.is_file() for path in owner_paths.values()
    ):
        return []

    violations: list[str] = []

    def source_for(path: Path) -> str:
        if not path.is_file():
            return ""
        return _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))

    def code_for(path: Path) -> str:
        return _mask_rust_non_code(source_for(path))

    for ecosystem, owner_path in owner_paths.items():
        owner_source = source_for(owner_path)
        owner_code = _mask_rust_non_code(owner_source)
        owner_methods = {
            name: declaration
            for name, declaration, _line_number in _rust_public_methods_in_impl(
                owner_source, "Package"
            )
        }
        methods = IWA_IMAGE_ADJUSTMENTS_SEMANTIC_METHODS[ecosystem]
        selectors = IWA_IMAGE_ADJUSTMENTS_SEMANTIC_SELECTORS[ecosystem]
        semantic_error = IWA_IMAGE_ADJUSTMENTS_SEMANTIC_TYPES[ecosystem][0]
        for method in methods:
            declaration = owner_methods.get(method)
            if declaration is None:
                violations.append(
                    f"focused litchi-{ecosystem.lower()} image-adjustments public API "
                    f"is missing Package method {method}: "
                    f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}"
                )
                continue
            if method != methods[-1]:
                for selector in selectors:
                    if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                        violations.append(
                            f"focused litchi-{ecosystem.lower()} image-adjustments "
                            f"Package method {method} must use selector-first "
                            f"{selector}: {IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}"
                        )
            if re.search(rf"\b{re.escape(semantic_error)}\b", declaration) is None:
                violations.append(
                    f"focused litchi-{ecosystem.lower()} image-adjustments Package "
                    f"method {method} must return its semantic error type: "
                    f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}"
                )

        semantic_types = frozenset(IWA_IMAGE_ADJUSTMENTS_SEMANTIC_TYPES[ecosystem])
        owner_exports = _rust_canonical_exports(owner_source, semantic_types)
        for type_name in sorted(semantic_types - owner_exports):
            violations.append(
                f"focused litchi-{ecosystem.lower()} image-adjustments owner is missing "
                f"semantic type {type_name}: "
                f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}"
            )

        # Keep raw bytes and generated archive vocabulary out of the
        # selector-first API.
        for declaration, line_number in _rust_public_declarations(owner_source):
            if RUST_BYTE_SLICE.search(declaration) is not None:
                violations.append(
                    f"focused litchi-{ecosystem.lower()} image-adjustments public API "
                    "exposes a raw byte slice: "
                    f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}:{line_number}"
                )
            if re.search(r"\b(?:tsd|tsp)::", declaration):
                violations.append(
                    f"focused litchi-{ecosystem.lower()} image-adjustments public API "
                    "exposes generated types: "
                    f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}:{line_number}"
                )

        required_owner_markers = (
            "image_adjustments_codec",
            "ImageAdjustment::new",
            "WireLimits",
            "DecodeOptions",
        )
        for marker in required_owner_markers:
            if re.search(rf"\b{re.escape(marker)}\b", owner_code) is None:
                violations.append(
                    f"focused litchi-{ecosystem.lower()} image-adjustments owner is missing "
                    f"bounded codec/semantic marker {marker}: "
                    f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}"
                )
        required_codec_calls = [
            ("decode_image_adjustments", IWA_IMAGE_ADJUSTMENTS_CODEC_DECODE_CALL)
        ]
        for marker, pattern in required_codec_calls:
            if pattern.search(owner_code) is None:
                violations.append(
                    f"focused litchi-{ecosystem.lower()} image-adjustments owner is missing "
                    f"bounded codec/semantic marker {marker}: "
                    f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}"
                )

        for leak in (
            IWA_IMAGE_ADJUSTMENTS_CODEC_GENERATED_LEAK,
            IWA_IMAGE_ADJUSTMENTS_GENERATED_TYPE_LEAK,
        ):
            for match in leak.finditer(owner_code):
                line_number = owner_code.count("\n", 0, match.start()) + 1
                violations.append(
                    f"focused litchi-{ecosystem.lower()} image-adjustments owner retains "
                    f"generated archive ingress {match.group(0).strip()}: "
                    f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}:{line_number}"
                )

        package_path = root / IWA_IMAGE_ADJUSTMENTS_OWNER_PACKAGE_SOURCES[ecosystem]
        package_source = source_for(package_path)
        package_code = _mask_rust_comments(package_source)
        if IWA_IMAGE_ADJUSTMENTS_OWNER_MODULES[ecosystem].search(package_code) is None:
            violations.append(
                f"focused litchi-{ecosystem.lower()} image-adjustments owner module is "
                "missing its default semantic declaration: "
                f"{IWA_IMAGE_ADJUSTMENTS_OWNER_PACKAGE_SOURCES[ecosystem]}"
            )

        export_sources = [package_source]
        export_path = root / IWA_IMAGE_ADJUSTMENTS_OWNER_EXPORT_SOURCES[ecosystem]
        if export_path.is_file():
            export_sources.append(_mask_rust_comments(source_for(export_path)))
        export_sources[0] = _mask_rust_comments(export_sources[0])
        semantic_types = frozenset(IWA_IMAGE_ADJUSTMENTS_SEMANTIC_TYPES[ecosystem])
        family_owner = IWA_IMAGE_ADJUSTMENTS_TRANSACTION_FAMILY_OWNERS.get(ecosystem)
        if family_owner is not None:
            violations.extend(
                _pages_transaction_family_violations(
                    root,
                    family_owner,
                    f"focused litchi-{ecosystem.lower()} image-adjustments",
                )
            )
        elif not any(
            IWA_IMAGE_ADJUSTMENTS_SEMANTIC_TYPES[ecosystem][0]
            in _rust_canonical_exports(source, semantic_types)
            for source in export_sources
        ):
            violations.append(
                f"focused litchi-{ecosystem.lower()} image-adjustments public facade "
                "is missing its semantic error export: "
                f"{IWA_IMAGE_ADJUSTMENTS_OWNER_EXPORT_SOURCES[ecosystem]}"
            )
        if ecosystem in IWA_IMAGE_ADJUSTMENTS_RETIRED_WRITE_ECOSYSTEMS:
            retired_write = IWA_IMAGE_ADJUSTMENTS_RETIRED_WRITE_SEAM
            for name, line_number in _rust_function_declarations(owner_source):
                if name == retired_write:
                    violations.append(
                        f"focused litchi-{ecosystem.lower()} image-adjustments "
                        "write bridge is retired: "
                        f"{IWA_IMAGE_ADJUSTMENTS_OWNER_SOURCES[ecosystem]}:{line_number}"
                    )
            for export_path, export_source in (
                (package_path, package_source),
                (
                    export_path,
                    _mask_rust_comments(source_for(export_path)),
                ),
            ):
                for match in re.finditer(
                    rf"(?m)^\s*pub\s+use[^;\n]*\b{re.escape(retired_write)}\b",
                    _mask_rust_non_code(export_source),
                ):
                    line_number = export_source.count("\n", 0, match.start()) + 1
                    violations.append(
                        f"focused litchi-{ecosystem.lower()} image-adjustments "
                        "write bridge re-export is retired: "
                        f"{export_path}:{line_number}"
                    )

    codec_path = root / IWA_IMAGE_ADJUSTMENTS_CODEC_SOURCE
    codec_code = code_for(codec_path)
    for marker in IWA_IMAGE_ADJUSTMENTS_CODEC_REQUIRED_MARKERS:
        if re.search(rf"\b{re.escape(marker)}\b", codec_code) is None:
            violations.append(
                "neutral image-adjustments codec is missing required bounded marker "
                f"{marker}: {IWA_IMAGE_ADJUSTMENTS_CODEC_SOURCE}"
            )
    if not any(
        re.search(rf"\b{re.escape(marker)}\b", codec_code)
        for marker in IWA_IMAGE_ADJUSTMENTS_CODEC_LAZY_MARKERS
    ):
        violations.append(
            "neutral image-adjustments codec must retain a private Buffa/lazy "
            f"projection path: {IWA_IMAGE_ADJUSTMENTS_CODEC_SOURCE}"
        )
    for leak in (
        IWA_IMAGE_ADJUSTMENTS_CODEC_GENERATED_LEAK,
        IWA_IMAGE_ADJUSTMENTS_GENERATED_TYPE_LEAK,
    ):
        for match in leak.finditer(codec_code):
            line_number = codec_code.count("\n", 0, match.start()) + 1
            violations.append(
                "neutral image-adjustments codec retains generated archive ingress "
                f"{match.group(0).strip()}: {IWA_IMAGE_ADJUSTMENTS_CODEC_SOURCE}:{line_number}"
            )

    codec_public_path = root / IWA_IMAGE_ADJUSTMENTS_CODEC_PUBLIC_SOURCE
    codec_public_code = _mask_rust_comments(source_for(codec_public_path))
    if IWA_IMAGE_ADJUSTMENTS_CODEC_MODULE.search(codec_public_code) is None:
        violations.append(
            "neutral image-adjustments codec is missing its hidden public module "
            f"declaration: {IWA_IMAGE_ADJUSTMENTS_CODEC_PUBLIC_SOURCE}"
        )

    return sorted(set(violations))


def audit_iwork_atomic_publication(root: Path = ROOT) -> list[str]:
    """Enforce the single durable publication boundary for iWork packages.

    The archive crate owns the physical temporary-file, replacement, and
    durability policy.  The three focused package crates retain their
    streaming Package::write_to method and expose a typed SaveError; their
    Package::save methods may either call the archive helper directly or
    delegate through a private package::save helper.

    This audit is dormant until the archive owner/module is introduced.  That
    keeps an older checkout useful while making the migration ratchet fail
    closed as soon as either half of the new boundary appears.
    """

    def production_source(path: Path) -> str:
        if not path.is_file():
            return ""
        return _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))

    publication_path = root / IWORK_ARCHIVE_PUBLICATION_FILE
    publication_mod_path = root / Path(
        "crates/litchi-iwa-archive/src/publication/mod.rs"
    )
    owner_path = (
        publication_path
        if publication_path.is_file()
        else publication_mod_path
        if publication_mod_path.is_file()
        else None
    )
    module_path = root / IWORK_ARCHIVE_PUBLICATION_MODULE_SOURCE
    module_source = production_source(module_path)
    module_code = _mask_rust_non_code(module_source)
    module_pattern = re.compile(
        rf"(?m)^[ \t]*(?:pub(?:[ \t\r\n]*\([^()]*\))?[ \t\r\n]+)?"
        rf"mod[ \t\r\n]+(?:r#)?{re.escape(IWORK_ARCHIVE_PUBLICATION_MODULE)}\b"
    )
    public_module_pattern = re.compile(
        rf"(?m)^[ \t]*pub(?![ \t\r\n]*\([^()]*\))[ \t\r\n]+mod"
        rf"[ \t\r\n]+(?:r#)?{re.escape(IWORK_ARCHIVE_PUBLICATION_MODULE)}\b"
    )
    has_module = module_pattern.search(module_code) is not None

    # Do not make the historical checkout fail until the new source/module
    # seam starts landing.  Once either appears, all required pieces are
    # checked, including a missing sibling owner or module declaration.
    if owner_path is None and not has_module:
        return []

    violations: list[str] = []
    if owner_path is None:
        violations.append(
            "iWork archive publication owner source is missing: "
            f"{IWORK_ARCHIVE_PUBLICATION_FILE}"
        )
    if not has_module:
        violations.append(
            "iWork archive publication module is missing from the archive crate: "
            f"{IWORK_ARCHIVE_PUBLICATION_MODULE_SOURCE}"
        )
    elif public_module_pattern.search(module_code) is None:
        violations.append(
            "iWork archive publication module must be public for focused package "
            "save call-through: "
            f"{IWORK_ARCHIVE_PUBLICATION_MODULE_SOURCE}"
        )

    if owner_path is not None:
        owner_source = production_source(owner_path)
        owner_code = _mask_rust_non_code(owner_source)
        public_error = re.compile(
            rf"(?m)^[ \t]*pub(?![ \t\r\n]*\([^()]*\))[ \t\r\n]+"
            rf"(?:struct|enum|type)[ \t\r\n]+"
            rf"(?:r#)?{re.escape(IWORK_ARCHIVE_PUBLICATION_ERROR)}\b"
        )
        if public_error.search(owner_code) is None:
            violations.append(
                "iWork archive publication owner is missing public typed "
                f"{IWORK_ARCHIVE_PUBLICATION_ERROR}: "
                f"{IWORK_ARCHIVE_PUBLICATION_FILE}"
            )
        public_helper = re.compile(
            rf"(?m)^[ \t]*pub(?![ \t\r\n]*\([^()]*\))[ \t\r\n]+fn"
            rf"[ \t\r\n]+(?:r#)?{re.escape(IWORK_ARCHIVE_PUBLICATION_HELPER)}\b"
        )
        if public_helper.search(owner_code) is None:
            violations.append(
                "iWork archive publication owner is missing public "
                f"{IWORK_ARCHIVE_PUBLICATION_HELPER} helper: "
                f"{IWORK_ARCHIVE_PUBLICATION_FILE}"
            )

    publication_call = re.compile(
        rf"(?<![A-Za-z0-9_#])(?:r#)?{re.escape(IWORK_ARCHIVE_PUBLICATION_HELPER)}"
        r"[ \t\r\n]*(?:::?[ \t\r\n]*<[^;{}()]*>)?"
        r"[ \t\r\n]*\("
    )
    publication_path_reference = re.compile(
        rf"(?<![A-Za-z0-9_#])(?:r#)?(?:litchi_iwa_archive|crate)"
        rf"[ \t\r\n]*::[ \t\r\n]*(?:r#)?"
        rf"{re.escape(IWORK_ARCHIVE_PUBLICATION_MODULE)}\b"
        rf"|(?<![A-Za-z0-9_#])(?:r#)?"
        rf"{re.escape(IWORK_ARCHIVE_PUBLICATION_MODULE)}"
        rf"[ \t\r\n]*::[ \t\r\n]*(?:r#)?"
        rf"(?:{re.escape(IWORK_ARCHIVE_PUBLICATION_HELPER)}\b|"
        rf"\{{[^;}}]*\b{re.escape(IWORK_ARCHIVE_PUBLICATION_HELPER)}\b)"
    )

    def routes_through_publication(source: str) -> bool:
        code = _mask_rust_non_code(source)
        if publication_call.search(code) is None:
            return False
        if publication_path_reference.search(code) is not None:
            return True
        # Tiny boundary fixtures sometimes omit imports because the owner
        # module itself is supplied by the fixture.  Treat an unqualified
        # helper call as the intended seam unless the concrete crate defines
        # a competing helper with the same name.
        local_helper = re.compile(
            rf"(?m)^[ \t]*(?:pub(?:[ \t\r\n]*\([^()]*\))?[ \t\r\n]+)?fn"
            rf"[ \t\r\n]+(?:r#)?{re.escape(IWORK_ARCHIVE_PUBLICATION_HELPER)}\b"
        )
        return local_helper.search(code) is None

    forbidden_publication_patterns = (
        (
            "NamedTempFile/tempfile builder",
            re.compile(
                r"(?<![A-Za-z0-9_#])(?:NamedTempFile|tempfile[ \t\r\n]*::"
                r"[ \t\r\n]*(?:Builder|tempfile_in))\b"
            ),
        ),
        (
            "fs::rename",
            re.compile(
                r"(?<![A-Za-z0-9_#])(?:std[ \t\r\n]*::[ \t\r\n]*)?"
                r"fs[ \t\r\n]*::[ \t\r\n]*(?:rename\b|"
                r"\{[^}]*\brename\b[^}]*\})"
                r"|(?<![A-Za-z0-9_#])rename[ \t\r\n]*\("
            ),
        ),
        (
            "sync publication",
            re.compile(
                r"(?<![A-Za-z0-9_#])(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*\."
                r"[ \t\r\n]*)?sync_(?:all|data)[ \t\r\n]*\("
            ),
        ),
        (
            "persist publication",
            re.compile(
                r"(?<![A-Za-z0-9_#])(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*\."
                r"[ \t\r\n]*)?persist(?:_noclobber|_in)?"
                r"[ \t\r\n]*\("
            ),
        ),
    )

    def report_forbidden(path: Path, label: str, code: str) -> None:
        relative = path.relative_to(root)
        for description, pattern in forbidden_publication_patterns:
            for match in pattern.finditer(code):
                line_number = code.count("\n", 0, match.start()) + 1
                label_prefix = f"focused {label}"
                violations.append(
                    f"{label_prefix} production source recreates direct "
                    f"{description} publication: {relative}:{line_number}"
                )

    for package_name, source_root in IWORK_FOCUSED_PACKAGE_SOURCE_ROOTS:
        absolute_root = root / source_root
        if not absolute_root.is_dir():
            violations.append(
                f"focused {package_name} package source root is missing: "
                f"{source_root}"
            )
            continue

        production_sources: list[str] = []
        for path in sorted(absolute_root.rglob("*.rs")):
            source = production_source(path)
            production_sources.append(source)
            report_forbidden(path, package_name, _mask_rust_non_code(source))
        package_source = "\n".join(production_sources)

        methods = _rust_public_methods_in_impl(package_source, "Package")
        saves = [
            (declaration, line_number)
            for name, declaration, line_number in methods
            if name == "save"
        ]
        writes = [
            (declaration, line_number)
            for name, declaration, line_number in methods
            if name == "write_to"
        ]
        if not saves:
            violations.append(
                f"focused {package_name} Package::save is missing a production "
                f"method: {source_root / 'package.rs'}"
            )
        else:
            if not any("SaveError" in declaration for declaration, _ in saves):
                violations.append(
                    f"focused {package_name} Package::save must return typed "
                    f"SaveError: {source_root / 'package.rs'}:{saves[0][1]}"
                )
            method_body = _rust_any_function_body(package_source, "save")
            direct_route = (
                method_body is not None
                and publication_call.search(_mask_rust_non_code(method_body))
                is not None
                and routes_through_publication(package_source)
            )
            helper_route = (
                method_body is not None
                and re.search(
                    r"\b(?:save|publish|publication)[ \t\r\n]*::[ \t\r\n]*"
                    r"(?:save|publish|replace_with)\b",
                    _mask_rust_non_code(method_body),
                )
                is not None
                and routes_through_publication(package_source)
            )
            if not (direct_route or helper_route):
                violations.append(
                    f"focused {package_name} Package::save must route through "
                    f"litchi-iwa-archive::{IWORK_ARCHIVE_PUBLICATION_MODULE}::"
                    f"{IWORK_ARCHIVE_PUBLICATION_HELPER}: "
                    f"{source_root / 'package.rs'}:{saves[0][1]}"
                )

        if not writes:
            violations.append(
                f"focused {package_name} Package::write_to must be preserved as a "
                f"production method: {source_root / 'package.rs'}"
            )
        elif not any("WriteError" in declaration for declaration, _ in writes):
            violations.append(
                f"focused {package_name} Package::write_to must retain its "
                f"WriteError surface: {source_root / 'package.rs'}:{writes[0][1]}"
            )

        public_save_error = any(
            re.search(
                r"\bpub(?![ \t\r\n]*\([^()]*\))[ \t\r\n]+"
                r"(?:struct|enum|type)[ \t\r\n]+(?:r#)?SaveError\b",
                declaration,
            )
            is not None
            for declaration, _line_number in _rust_public_declarations(package_source)
        )
        if not public_save_error:
            violations.append(
                f"focused {package_name} package is missing public SaveError: "
                f"{source_root / 'package.rs'}"
            )

        library_path = root / (source_root / "lib.rs")
        library_source = production_source(library_path)
        if "SaveError" not in _rust_canonical_exports(
            library_source, frozenset({"SaveError"})
        ):
            violations.append(
                f"focused {package_name} package is missing root SaveError export: "
                f"{source_root / 'lib.rs'}"
            )

    return sorted(set(violations))


def _audit_iwa_drawable_container_projection(
    root: Path,
) -> list[str]:
    """Check the checked-in TSD container/group projection field ownership."""

    projection_path = root / IWA_DRAWABLE_CONTAINER_PROJECTION_SOURCE
    if not projection_path.is_file():
        return [
            "focused TSD drawable-container codec is missing private Buffa "
            f"projection: {IWA_DRAWABLE_CONTAINER_PROJECTION_SOURCE}"
        ]
    raw = projection_path.read_text(encoding="utf-8")
    projection = re.sub(r"//[^\n]*|/\*.*?\*/", "", raw, flags=re.DOTALL)
    violations: list[str] = []
    for marker in ('syntax = "proto2"', "package LitchiIwaProjection"):
        if marker not in projection:
            violations.append(
                "focused TSD drawable-container projection is missing marker "
                f"{marker}: {IWA_DRAWABLE_CONTAINER_PROJECTION_SOURCE}"
            )
    for message, fields in IWA_DRAWABLE_CONTAINER_PROJECTION_FIELDS.items():
        message_match = re.search(
            rf"(?s)\bmessage\s+{re.escape(message)}\s*\{{(?P<body>.*?)\}}",
            projection,
        )
        if message_match is None:
            violations.append(
                "focused TSD drawable-container projection is missing message "
                f"{message}: {IWA_DRAWABLE_CONTAINER_PROJECTION_SOURCE}"
            )
            continue
        body = message_match.group("body")
        for field, (field_type, number) in fields.items():
            if re.search(
                rf"\b(?:optional|repeated|required)\s+"
                rf"\.?(?:[A-Za-z_][A-Za-z0-9_.]*\.)?"
                rf"{re.escape(field_type)}\s+{re.escape(field)}\s*=\s*{number}\b",
                body,
            ) is None:
                violations.append(
                    "focused TSD drawable-container projection has incorrect or "
                    f"missing field {message}.{field}={number}: "
                    f"{IWA_DRAWABLE_CONTAINER_PROJECTION_SOURCE}"
                )
    return violations


def audit_iwa_drawable_container_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the TSD container/group codec and Buffa build ownership private."""

    codec_path = root / IWA_DRAWABLE_CONTAINER_CODEC_SOURCE
    projection_path = root / IWA_DRAWABLE_CONTAINER_PROJECTION_SOURCE
    build_path = root / IWA_DRAWABLE_CONTAINER_BUILD_SOURCE
    public_path = root / IWA_DRAWABLE_CONTAINER_CODEC_PUBLIC_SOURCE
    if not (codec_path.is_file() or projection_path.is_file()):
        return []

    violations: list[str] = []
    if not codec_path.is_file():
        violations.append(
            "focused TSD drawable-container codec source is missing: "
            f"{IWA_DRAWABLE_CONTAINER_CODEC_SOURCE}"
        )
    else:
        raw_codec = _mask_rust_cfg_test_items(codec_path.read_text(encoding="utf-8"))
        codec = _mask_rust_non_code(raw_codec)
        for marker in IWA_DRAWABLE_CONTAINER_CODEC_REQUIRED_MARKERS:
            if marker not in codec:
                violations.append(
                    "focused TSD drawable-container codec is missing marker "
                    f"{marker}: {IWA_DRAWABLE_CONTAINER_CODEC_SOURCE}"
                )
        for operation in IWA_DRAWABLE_CONTAINER_CODEC_FORBIDDEN_OPERATIONS:
            if operation in codec:
                violations.append(
                    "focused TSD drawable-container codec retains eager/encoding "
                    f"operation {operation}: {IWA_DRAWABLE_CONTAINER_CODEC_SOURCE}"
                )
        for pattern, description in (
            (
                IWA_DRAWABLE_CONTAINER_FORBIDDEN_GENERATED_DECODE,
                "generated archive decode",
            ),
            (IWA_DRAWABLE_CONTAINER_PROST_DECODE, "generic generated Message decode"),
        ):
            match = pattern.search(codec)
            if match is not None:
                line = codec.count("\n", 0, match.start()) + 1
                violations.append(
                    "focused TSD drawable-container codec retains "
                    f"{description}: {IWA_DRAWABLE_CONTAINER_CODEC_SOURCE}:{line}"
                )

    violations.extend(_audit_iwa_drawable_container_projection(root))

    if not build_path.is_file():
        violations.append(
            "focused TSD drawable-container Buffa build source is missing: "
            f"{IWA_DRAWABLE_CONTAINER_BUILD_SOURCE}"
        )
    else:
        build = build_path.read_text(encoding="utf-8")
        route = re.search(
            r"(?s)\blet\s+buffa_drawable_container_out_directory\b.*?"
            r"\benforce_drawable_container_projection_budget\s*\(\s*"
            r"&\s*buffa_drawable_container_out_directory\s*,?\s*\)",
            build,
        )
        if route is None:
            violations.append(
                "focused TSD drawable-container Buffa build is missing its "
                f"dedicated route: {IWA_DRAWABLE_CONTAINER_BUILD_SOURCE}"
            )
            build_focus = ""
        else:
            build_focus = route.group(0)
        for marker in (
            "TSDDrawableContainerArchive.proto",
            "buffa-drawable-container",
            "iwa_drawable_container_buffa_protos.rs",
            ".includes(&[buffa_projection_directory])",
            ".generate_views(true)",
            ".lazy_views(true)",
            ".preserve_unknown_fields(false)",
            ".compile()",
            "enforce_drawable_container_projection_budget",
        ):
            if marker not in build_focus:
                violations.append(
                    "focused TSD drawable-container Buffa build is missing marker "
                    f"{marker}: {IWA_DRAWABLE_CONTAINER_BUILD_SOURCE}"
                )
        for marker in (
            "fn enforce_drawable_container_projection_provenance(",
            "enforce_drawable_container_projection_provenance(proto_directory, "
            "buffa_projection_directory)?",
        ):
            if marker not in build:
                violations.append(
                    "focused TSD drawable-container build is missing provenance "
                    f"marker {marker}: {IWA_DRAWABLE_CONTAINER_BUILD_SOURCE}"
                )

    if not public_path.is_file():
        violations.append(
            "focused TSD drawable-container generated module owner is missing: "
            f"{IWA_DRAWABLE_CONTAINER_CODEC_PUBLIC_SOURCE}"
        )
    else:
        public = public_path.read_text(encoding="utf-8")
        if re.search(
            rf"(?ms)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\].{{0,260}}?"
            rf"^\s*mod\s+{re.escape(IWA_DRAWABLE_CONTAINER_GENERATED_MODULE)}\b",
            public,
        ) is None:
            violations.append(
                "focused TSD drawable-container generated Buffa projection must "
                f"remain private: {IWA_DRAWABLE_CONTAINER_CODEC_PUBLIC_SOURCE}"
            )
        if IWA_DRAWABLE_CONTAINER_GENERATED_INCLUDE not in public:
            violations.append(
                "focused TSD drawable-container generated projection include is "
                f"missing: {IWA_DRAWABLE_CONTAINER_CODEC_PUBLIC_SOURCE}"
            )

    return sorted(set(violations))


def _proto_message_body(source: str, name: str) -> str | None:
    """Return one brace-balanced Protocol Buffer message body.

    The media-creation projection has a nested enum inside ``MovieArchive``;
    a non-greedy regular expression would stop at that enum's closing brace
    and silently leave the outer fields unaudited.  This tiny scanner is
    intentionally limited to the comment-stripped schema used by the boundary
    checker and does not attempt to be a full protobuf parser.
    """

    match = re.search(rf"\bmessage\s+{re.escape(name)}\s*\{{", source)
    if match is None:
        return None
    opening = source.find("{", match.start(), match.end())
    if opening < 0:
        return None
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


def _keynote_chart_arrangement_owner_present(root: Path) -> bool:
    """Return whether the focused chart-arrangement owner is wired."""

    owner_path = root / KEYNOTE_CHART_ARRANGEMENT_OWNER_SOURCE
    package_path = root / KEYNOTE_CHART_ARRANGEMENT_EXPORT_SOURCES[0]
    package_source = (
        _mask_rust_non_code(
            _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        )
        if package_path.is_file()
        else ""
    )
    return owner_path.is_file() and re.search(
        rf"(?m)^(?:pub(?:[ \t]*\([^()]*\))?[ \t]+)?mod[ \t]+"
        rf"{re.escape(KEYNOTE_CHART_ARRANGEMENT_PACKAGE_MODULE)}\s*;",
        package_source,
    ) is not None


def _chart_arrangement_all_owners_present(root: Path) -> bool:
    """Return whether Keynote, Numbers, and Pages owners are all wired."""

    if not _keynote_chart_arrangement_owner_present(root):
        return False
    for owner_path, package_path, module_name in (
        (
            NUMBERS_CHART_ARRANGEMENT_OWNER_SOURCE,
            NUMBERS_PACKAGE_SOURCE,
            NUMBERS_CHART_ARRANGEMENT_PACKAGE_MODULE,
        ),
        (
            PAGES_BODY_CHART_ARRANGEMENT_OWNER_SOURCE,
            PAGES_PACKAGE_SOURCE,
            PAGES_BODY_CHART_ARRANGEMENT_PACKAGE_MODULE,
        ),
    ):
        package = root / package_path
        if not (root / owner_path).is_file() or not package.is_file():
            return False
        source = _mask_rust_cfg_test_items(package.read_text(encoding="utf-8"))
        if not _rust_root_level_module_declarations(
            source,
            frozenset({module_name}),
        ):
            return False
    return True


def _chart_arrangement_check_codec(root: Path) -> list[str]:
    """Require the neutral private, Buffa-backed chart-arrangement boundary."""

    codec_path = root / CHART_ARRANGEMENT_CODEC_SOURCE
    retired_codec_path = root / RETIRED_KEYNOTE_CHART_ARRANGEMENT_CODEC_SOURCE
    public_path = root / CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE
    violations: list[str] = []
    if retired_codec_path.is_file():
        violations.append(
            "retired Keynote-specific chart-arrangement codec source must remain "
            f"deleted: {RETIRED_KEYNOTE_CHART_ARRANGEMENT_CODEC_SOURCE}"
        )
    if not codec_path.is_file():
        return [
            "neutral chart-arrangement codec source is missing: "
            f"{CHART_ARRANGEMENT_CODEC_SOURCE}"
        ]

    raw_codec = codec_path.read_text(encoding="utf-8")
    codec = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_codec))
    codec_attributes = _mask_rust_non_code(raw_codec)

    for api in CHART_ARRANGEMENT_CODEC_REQUIRED_APIS:
        if re.search(
            rf"\b(?:pub\s+)?(?:fn|struct|enum|type)\s+{re.escape(api)}\b",
            codec,
        ) is None:
            violations.append(
                "neutral chart-arrangement codec is missing strict API "
                f"{api}: {CHART_ARRANGEMENT_CODEC_SOURCE}"
            )

    def has_marker(marker: str | tuple[str, ...]) -> bool:
        if isinstance(marker, tuple):
            return any(has_marker(item) for item in marker)
        return marker in codec

    for label, markers in CHART_ARRANGEMENT_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not all(has_marker(marker) for marker in markers):
            violations.append(
                "neutral chart-arrangement codec is missing "
                f"{label} marker: {CHART_ARRANGEMENT_CODEC_SOURCE}"
            )

    if re.search(r"\b(?:prost|prost_types)\b", codec) is not None:
        violations.append(
            "neutral chart-arrangement codec must use Buffa rather "
            f"than Prost: {CHART_ARRANGEMENT_CODEC_SOURCE}"
        )
    for declaration, line_number in _rust_public_declarations(codec):
        if re.search(
            rf"\b{re.escape(CHART_ARRANGEMENT_CODEC_GENERATED_MODULE)}\b"
            r"|\bprojection\s*::",
            declaration,
        ):
            violations.append(
                "neutral chart-arrangement codec must not expose Buffa "
                f"generated types: {CHART_ARRANGEMENT_CODEC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        codec,
        frozenset({CHART_ARRANGEMENT_CODEC_GENERATED_MODULE}),
    ):
        violations.append(
            "neutral chart-arrangement codec must not expose Buffa "
            f"generated aliases: {CHART_ARRANGEMENT_CODEC_SOURCE}"
        )

    if not public_path.is_file():
        violations.append(
            "neutral chart-arrangement codec public module source is "
            f"missing: {CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE}"
        )
        return sorted(set(violations))

    raw_public = public_path.read_text(encoding="utf-8")
    public = _mask_rust_cfg_test_items(raw_public)
    public_code = _mask_rust_non_code(public)
    if not _rust_root_level_matches(public, CHART_ARRANGEMENT_CODEC_HIDDEN_MODULE):
        violations.append(
            "neutral chart-arrangement codec module must be hidden: "
            f"{CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE}"
        )
    retired_module_declarations = _rust_root_level_module_declarations(
        public,
        frozenset({RETIRED_KEYNOTE_CHART_ARRANGEMENT_CODEC_MODULE}),
    )
    for _, _, _, _, line_number in retired_module_declarations:
        violations.append(
            "retired Keynote-specific chart-arrangement codec module must remain "
            f"deleted: {CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE}:{line_number}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        public,
        frozenset({CHART_ARRANGEMENT_CODEC_GENERATED_MODULE}),
    )
    if not generated_declarations:
        violations.append(
            "neutral chart-arrangement codec is missing its private "
            f"Buffa generated module: {CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE}"
        )
    for module, visibility, shape, body, line_number in generated_declarations:
        if visibility is not None:
            violations.append(
                "neutral chart-arrangement codec generated module must "
                f"remain private: {CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "neutral chart-arrangement codec generated module has "
                f"invalid include shape: {CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        public,
        frozenset({CHART_ARRANGEMENT_CODEC_GENERATED_MODULE}),
    ):
        violations.append(
            "neutral chart-arrangement codec must not re-export Buffa "
            f"generated types: {CHART_ARRANGEMENT_CODEC_PUBLIC_SOURCE}"
        )
    if re.search(
        r"(?m)#\s*\[\s*cfg\s*\(\s*test\s*\)\s*\]", codec_attributes
    ) is None:
        violations.append(
            "neutral chart-arrangement codec is missing cfg(test) "
            f"coverage: {CHART_ARRANGEMENT_CODEC_SOURCE}"
        )
    elif re.search(r"(?m)#\s*\[\s*test\s*\]", codec_attributes) is None:
        violations.append(
            "neutral chart-arrangement codec is missing #[test] "
            f"coverage: {CHART_ARRANGEMENT_CODEC_SOURCE}"
        )
    return sorted(set(violations))


def audit_chart_arrangement_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Require the neutral chart-arrangement codec and private Buffa view."""

    if not _chart_arrangement_all_owners_present(root):
        return []
    return _chart_arrangement_check_codec(root)


def _audit_focused_chart_arrangement_owner_source_topology(
    root: Path,
    *,
    label: str,
    owner_source: Path,
    package_source: Path,
    public_source: Path,
    module_name: str,
    methods: tuple[str, ...],
    selector_types: dict[str, tuple[str, ...]],
    method_markers: dict[str, tuple[str, ...]],
    owner_markers: tuple[str, ...],
    lazy_codec_markers: tuple[str, ...],
    canonical_types: tuple[str, ...],
    owner_module_visibility: str | None = None,
) -> list[str]:
    """Audit one focused chart-arrangement package owner.

    ``owner_module_visibility`` is the exact visibility the package owner
    module must carry: ``None`` for a private ``mod``, ``"pub(crate)"`` for
    the ADR 0034 Pages owners.  A ``pub mod`` owner is always rejected.

    The ratchet is deliberately limited to the public package seam and a few
    stable ownership markers.  Native graph identifiers and wire helpers may
    occur in the private implementation, but the selector-first methods must
    admit semantic selectors, route through bounded discovery, and verify the
    selected candidate locally before publication.
    """

    owner_path = root / owner_source
    package_path = root / package_source
    public_path = root / public_source
    violations: list[str] = []
    if not owner_path.is_file():
        return [
            f"focused {label} owner source is missing: {owner_source}"
        ]
    if not package_path.is_file():
        violations.append(
            f"focused {label} package source is missing: {package_source}"
        )
    if not public_path.is_file():
        violations.append(
            f"focused {label} crate export source is missing: {public_source}"
        )

    owner_source_text = _mask_rust_cfg_test_items(
        owner_path.read_text(encoding="utf-8")
    )
    owner_code = _mask_rust_non_code(owner_source_text)
    package_source_text = (
        _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
        if package_path.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(package_source_text)
    public_source_text = (
        _mask_rust_cfg_test_items(public_path.read_text(encoding="utf-8"))
        if public_path.is_file()
        else ""
    )
    public_code = _mask_rust_non_code(public_source_text)

    module_declarations = _rust_root_level_module_declarations(
        package_source_text,
        frozenset({module_name}),
    )
    if not module_declarations:
        violations.append(
            f"focused {label} package owner module is missing: {package_source}"
        )
    for _, visibility, _, _, line_number in module_declarations:
        normalized = None if visibility is None else re.sub(r"\s+", "", visibility)
        if normalized != owner_module_visibility:
            violations.append(
                f"focused {label} package owner module must remain private: "
                f"{package_source}:{line_number}"
            )

    public_declarations = _rust_public_declarations(owner_source_text)
    method_declarations: dict[str, tuple[str, int]] = {}
    for name in methods:
        for declaration, line_number in public_declarations:
            if re.search(
                rf"\bfn[ \t\r\n]+(?:r#)?{re.escape(name)}\b",
                declaration,
            ):
                method_declarations[name] = (declaration, line_number)
                break
        declaration_record = method_declarations.get(name)
        if declaration_record is None:
            violations.append(
                f"focused {label} package API is missing public method {name}: "
                f"{owner_source}"
            )
            continue
        declaration, line_number = declaration_record
        expected_selectors = selector_types.get(name, ())
        for selector in expected_selectors:
            if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                violations.append(
                    f"focused {label} selector API {name} must accept {selector}: "
                    f"{owner_source}:{line_number}"
                )
        if re.search(
            r"&\s*\[\s*u8\s*\]|\b(?:u64|u32|i64|i32)\b",
            declaration,
        ):
            violations.append(
                f"focused {label} public method {name} exposes a raw native/wire "
                f"parameter: {owner_source}:{line_number}"
            )
        body = _rust_any_function_body(owner_source_text, name)
        if body is None:
            violations.append(
                f"focused {label} package API method {name} has no inspectable body: "
                f"{owner_source}:{line_number}"
            )
            continue
        for marker in method_markers.get(name, ()):
            if re.search(rf"\b{re.escape(marker)}\b", body) is None:
                violations.append(
                    f"focused {label} package API method {name} is missing "
                    f"ownership marker {marker}: {owner_source}:{line_number}"
                )

    for marker in owner_markers:
        if re.search(rf"\b{re.escape(marker)}\b", owner_code) is None:
            violations.append(
                f"focused {label} owner is missing admission/locality marker "
                f"{marker}: {owner_source}"
            )
    for marker in lazy_codec_markers:
        if re.search(rf"\b{re.escape(marker)}\b", owner_code) is None:
            violations.append(
                f"focused {label} owner is missing lazy Buffa codec marker "
                f"{marker}: {owner_source}"
            )
    if re.search(r"\bPackage\s*::\s*from_bytes\s*\(", owner_code):
        violations.append(
            f"focused {label} owner must not reparse through Package::from_bytes: "
            f"{owner_source}"
        )
    if re.search(r"\b(?:prost|prost_types)\b", owner_code):
        violations.append(
            f"focused {label} owner must use the neutral Buffa codec rather than Prost: "
            f"{owner_source}"
        )

    if public_path.is_file():
        for name in canonical_types:
            if re.search(rf"\b{re.escape(name)}\b", public_code) is None:
                violations.append(
                    f"focused {label} crate export is missing {name}: {public_source}"
                )
        if re.search(r"\bChartArrangement\b", public_code) is None:
            violations.append(
                f"focused {label} crate export is missing semantic ChartArrangement: "
                f"{public_source}"
            )

    return sorted(set(violations))


def audit_annotation_author_codec_source_topology(root: Path = ROOT) -> list[str]:
    """Keep the neutral annotation-author codec lazy, bounded, and private-wire."""

    source_path = root / ANNOTATION_AUTHOR_CODEC_SOURCE
    if not source_path.is_file():
        return []

    raw_source = source_path.read_text(encoding="utf-8")
    source = _mask_rust_cfg_test_items(raw_source)
    code = _mask_rust_non_code(source)
    violations: list[str] = []

    for api in sorted(ANNOTATION_AUTHOR_CODEC_APIS):
        if re.search(
            rf"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
            rf"(?:r#)?{re.escape(api)}\b",
            code,
        ) is None:
            violations.append(
                "annotation-author codec is missing strict API "
                f"{api}: {ANNOTATION_AUTHOR_CODEC_SOURCE}"
            )
    for type_name in sorted(ANNOTATION_AUTHOR_CODEC_TYPES):
        if re.search(
            rf"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
            rf"(?:struct|enum|type)[ \t]+(?:r#)?{re.escape(type_name)}\b",
            code,
        ) is None:
            violations.append(
                "annotation-author codec is missing strict type "
                f"{type_name}: {ANNOTATION_AUTHOR_CODEC_SOURCE}"
            )

    for label, markers in ANNOTATION_AUTHOR_CODEC_MARKER_GROUPS.items():
        if not any(marker in raw_source for marker in markers):
            violations.append(
                "annotation-author codec is missing "
                f"{label} marker: {ANNOTATION_AUTHOR_CODEC_SOURCE}"
            )

    for label, pattern in ANNOTATION_AUTHOR_CODEC_FORBIDDEN_PATTERNS:
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                "annotation-author codec production source uses "
                f"{label}: {ANNOTATION_AUTHOR_CODEC_SOURCE}:{line_number}"
            )

    projection_path = root / ANNOTATION_AUTHOR_CODEC_PROJECTION_SOURCE
    if not projection_path.is_file():
        violations.append(
            "annotation-author codec is missing its Buffa projection: "
            f"{ANNOTATION_AUTHOR_CODEC_PROJECTION_SOURCE}"
        )
    else:
        # This is a protobuf projection, not Rust. Keep its literal syntax
        # visible instead of passing it through the Rust string masker.
        projection = projection_path.read_text(encoding="utf-8")
        for marker in (
            'syntax = "proto2"',
            "message Color",
            "message Reference",
            "message AnnotationAuthorArchive",
            "message AnnotationAuthorStorageArchive",
            "required uint64 identifier",
            "optional string public_id",
        ):
            if marker not in projection:
                violations.append(
                    "annotation-author Buffa projection is missing "
                    f"{marker}: {ANNOTATION_AUTHOR_CODEC_PROJECTION_SOURCE}"
                )

    public_path = root / ANNOTATION_AUTHOR_CODEC_PUBLIC_SOURCE
    if not public_path.is_file():
        violations.append(
            "annotation-author codec is missing its public module source: "
            f"{ANNOTATION_AUTHOR_CODEC_PUBLIC_SOURCE}"
        )
    else:
        public = public_path.read_text(encoding="utf-8")
        public_code = _mask_rust_non_code(public)
        if re.search(
            rf"(?m)^\s*pub\s+mod\s+{re.escape('annotation_author_codec')}\s*;",
            public_code,
        ) is None:
            violations.append(
                "annotation-author codec module is missing its hidden public ingress: "
                f"{ANNOTATION_AUTHOR_CODEC_PUBLIC_SOURCE}"
            )
        if re.search(
            rf"(?m)^\s*pub\s+mod\s+{re.escape(ANNOTATION_AUTHOR_CODEC_GENERATED_MODULE)}\b",
            public_code,
        ) is not None:
            violations.append(
                "annotation-author generated Buffa module must remain private: "
                f"{ANNOTATION_AUTHOR_CODEC_PUBLIC_SOURCE}"
            )
        if re.search(
            rf"(?m)^\s*mod\s+{re.escape(ANNOTATION_AUTHOR_CODEC_GENERATED_MODULE)}\s*\{{",
            public_code,
        ) is None:
            violations.append(
                "annotation-author generated Buffa module is missing: "
                f"{ANNOTATION_AUTHOR_CODEC_PUBLIC_SOURCE}"
            )
        if "/buffa-annotation-author/iwa_annotation_author_buffa_protos.rs" not in public:
            violations.append(
                "annotation-author generated Buffa include is missing: "
                f"{ANNOTATION_AUTHOR_CODEC_PUBLIC_SOURCE}"
            )

    build_path = root / ANNOTATION_AUTHOR_CODEC_BUILD_SOURCE
    if not build_path.is_file():
        violations.append(
            "annotation-author codec is missing build-script provenance: "
            f"{ANNOTATION_AUTHOR_CODEC_BUILD_SOURCE}"
        )
    else:
        build = build_path.read_text(encoding="utf-8")
        for marker in (
            "src/annotation_author_codec.rs",
            "TSKAnnotationAuthorProjection.proto",
            "enforce_annotation_author_projection_budget",
            "buffa-annotation-author",
        ):
            if marker not in build:
                violations.append(
                    "annotation-author build script is missing "
                    f"{marker}: {ANNOTATION_AUTHOR_CODEC_BUILD_SOURCE}"
                )
        if re.search(
            rf"(?s){re.escape('src/annotation_author_codec.rs')}.*?"
            rf"{re.escape(ANNOTATION_AUTHOR_CODEC_GENERATED_MODULE)}",
            build,
        ) is None:
            violations.append(
                "annotation-author production ingress is not tied to its private generated module: "
                f"{ANNOTATION_AUTHOR_CODEC_BUILD_SOURCE}"
            )

    if not re.search(r"(?m)^\s*#\s*\[test\b", raw_source):
        violations.append(
            "annotation-author codec is missing direct unit coverage: "
            f"{ANNOTATION_AUTHOR_CODEC_SOURCE}"
        )
    return sorted(set(violations))


def _repository_cargo_manifests(root: Path) -> tuple[Path, ...]:
    """Return every repository Cargo.toml outside generated metadata trees."""

    root = root.resolve()
    manifests: set[Path] = set()
    for path in root.rglob("Cargo.toml"):
        relative = path.relative_to(root)
        if any(part in {".git", "target"} for part in relative.parts):
            continue
        if path.is_file():
            manifests.add(path.resolve())
    return tuple(sorted(manifests))


def _cargo_dependency_tables(
    document: dict[str, Any],
) -> Iterable[tuple[str, dict[str, Any]]]:
    """Yield dependency tables Cargo can use to introduce package edges."""

    for kind in ("dependencies", "dev-dependencies", "build-dependencies"):
        table = document.get(kind)
        if isinstance(table, dict):
            yield kind, table

    targets = document.get("target")
    if isinstance(targets, dict):
        for target_name in sorted(targets):
            target = targets[target_name]
            if not isinstance(target, dict):
                continue
            for kind in ("dependencies", "dev-dependencies", "build-dependencies"):
                table = target.get(kind)
                if isinstance(table, dict):
                    yield f"target.{target_name}.{kind}", table

    workspace = document.get("workspace")
    if isinstance(workspace, dict):
        table = workspace.get("dependencies")
        if isinstance(table, dict):
            yield "workspace.dependencies", table

    patch = document.get("patch")
    if isinstance(patch, dict):
        for source_name in sorted(patch):
            table = patch[source_name]
            if isinstance(table, dict):
                yield f"patch.{source_name}", table

    replace = document.get("replace")
    if isinstance(replace, dict):
        yield "replace", replace


def _cargo_path_package_name(
    manifest: Path,
    path_value: Any,
    cache: dict[Path, str | None],
) -> str | None:
    """Resolve a path dependency's package name for aliases without `package`."""

    if not isinstance(path_value, str) or not path_value or "$" in path_value:
        return None
    target = Path(path_value)
    path_component = target.name
    if path_component == "Cargo.toml":
        path_component = target.parent.name
    if not target.is_absolute():
        target = manifest.parent / target
    try:
        target = target.resolve()
    except OSError:
        return None
    if target.is_dir():
        target /= "Cargo.toml"
    if target in cache:
        return cache[target]

    package_name: str | None = None
    try:
        document = tomllib.loads(target.read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError):
        # A removed path target may leave only the alias behind.  The final
        # path component is still useful evidence, while other missing path
        # targets remain Cargo errors.
        if path_component:
            package_name = path_component
    else:
        package = document.get("package")
        if isinstance(package, dict) and isinstance(package.get("name"), str):
            package_name = package["name"]
    cache[target] = package_name
    return package_name


def _cargo_dependency_candidates(
    manifest: Path,
    table_name: str,
    alias: str,
    specification: Any,
    path_cache: dict[Path, str | None],
) -> frozenset[str]:
    """Collect package names represented by one manifest dependency entry."""

    if table_name == "replace":
        candidates = {alias.split(":", 1)[0] if ":" in alias else alias}
    elif isinstance(specification, dict) and "package" in specification:
        # Cargo's `package` field names the real package; the table key is
        # only a local dependency alias and must not create a false edge.
        package_name = specification.get("package")
        candidates = {package_name} if isinstance(package_name, str) else {alias}
    else:
        candidates = {alias}
    if isinstance(specification, dict) and "package" not in specification:
        path_name = _cargo_path_package_name(
            manifest, specification.get("path"), path_cache
        )
        if path_name is not None:
            candidates.add(path_name)
    return frozenset(
        variant
        for candidate in candidates
        for variant in (
            candidate,
            candidate.replace("_", "-"),
            candidate.replace("-", "_"),
        )
    )


def _manifest_relative_path(path: Path, root: Path) -> str:
    try:
        return str(path.resolve().relative_to(root.resolve()))
    except ValueError:
        return str(path)


def audit_manifest_dependency_inventory(
    root: Path = ROOT,
    policy: Policy | None = None,
) -> list[str]:
    """Reject retired migration-host edges in every repository Cargo manifest.

    Cargo metadata only covers workspace members.  Standalone fuzz crates,
    excluded tools, and root/workspace manifests still influence dependency
    resolution, so inspect their TOML tables directly and fail closed on
    unreadable or malformed manifests.
    """

    migration_hosts = policy.migration_hosts if policy is not None else frozenset()
    violations: list[str] = []
    path_cache: dict[Path, str | None] = {}
    for manifest in _repository_cargo_manifests(root):
        try:
            document = tomllib.loads(manifest.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError) as error:
            violations.append(
                "Cargo manifest cannot be parsed: "
                f"{_manifest_relative_path(manifest, root)} ({error})"
            )
            continue

        for table_name, table in _cargo_dependency_tables(document):
            for alias, specification in table.items():
                if not isinstance(alias, str):
                    continue
                offending = sorted(
                    _cargo_dependency_candidates(
                        manifest,
                        table_name,
                        alias,
                        specification,
                        path_cache,
                    )
                    & migration_hosts
                )
                for package_name in offending:
                    violations.append(
                        "retired migration-host manifest dependency outside policy: "
                        f"{_manifest_relative_path(manifest, root)} "
                        f"[{table_name}] alias={alias}, package={package_name}"
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


IWA_TABLE_CELL_BORDERS_COMMON_SOURCE = Path(
    "crates/litchi-iwa-common/src/table/cell.rs"
)
IWA_TABLE_CELL_BORDERS_LEGACY_PATH = re.compile(
    r"\b(?:crate::numbers::editor::table::cell|super::table::cell)::Borders\b"
)
IWA_TABLE_CELL_BORDERS_COMMON_IMPORT = re.compile(
    r"\b(?:pub[ \t]+)?use[ \t]+litchi_iwa_common::table::cell::"
    r"(?:Borders\b|\{[^;\n]*\bBorders\b)"
)

# Keep the neutral sparse table owner below concrete iWork formats; facade
# behavior remains covered by cross-crate tests.
NEUTRAL_TABLE_MANIFEST = Path("crates/litchi-iwa-common/Cargo.toml")
NEUTRAL_TABLE_SOURCE_ROOT = Path("crates/litchi-iwa-common/src/table")
NEUTRAL_TABLE_MODULE_SOURCE = NEUTRAL_TABLE_SOURCE_ROOT / "mod.rs"
NEUTRAL_TABLE_CELL_MODULE_SOURCE = NEUTRAL_TABLE_SOURCE_ROOT / "cell.rs"
NEUTRAL_TABLE_VALUE_SOURCE = NEUTRAL_TABLE_SOURCE_ROOT / "cell" / "value.rs"
NEUTRAL_TABLE_COORDINATE_SOURCE = NEUTRAL_TABLE_SOURCE_ROOT / "coordinate.rs"
NEUTRAL_TABLE_MODEL_SOURCE = NEUTRAL_TABLE_SOURCE_ROOT / "model.rs"
NEUTRAL_TABLE_REQUIRED_EXPORTS = {
    NEUTRAL_TABLE_VALUE_SOURCE: frozenset({"FiniteF64", "FiniteF64Error", "Type", "Update", "Value"}),
    NEUTRAL_TABLE_COORDINATE_SOURCE: frozenset({"AddressError", "At", "Area", "Error", "Result"}),
    NEUTRAL_TABLE_MODEL_SOURCE: frozenset({"Builder", "Cell", "Dimensions", "Error", "Grid", "Table", "View"}),
}
NEUTRAL_TABLE_REQUIRED_MODULES = {
    NEUTRAL_TABLE_MODULE_SOURCE: frozenset({"cell", "coordinate", "model"}),
    NEUTRAL_TABLE_CELL_MODULE_SOURCE: frozenset({"value"}),
}
NEUTRAL_TABLE_FORBIDDEN_DEPENDENCIES = frozenset(
    {"buffa", "prost", "prost-types"}
    | {
        f"litchi-{name}"
        for name in (
            "iwa", "iwa-archive", "iwa-core", "iwa-detect", "iwa-graph",
            "iwa-index", "iwa-package", "iwa-protos", "iwa-text-wire",
            "keynote", "numbers", "numbers-wire", "pages",
        )
    }
)
NEUTRAL_TABLE_FORBIDDEN_IMPORTS = frozenset(
    {"archive", "proto", "protobuf", "wire"}
    | {name.replace("-", "_") for name in NEUTRAL_TABLE_FORBIDDEN_DEPENDENCIES}
)
NEUTRAL_TABLE_FORBIDDEN_PATH = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    + "|".join(
        re.escape(name)
        for name in sorted(NEUTRAL_TABLE_FORBIDDEN_IMPORTS, key=len, reverse=True)
    )
    + r")(?![A-Za-z0-9_])[ \t\r\n]*::"
)
NEUTRAL_TABLE_FORBIDDEN_USE = re.compile(
    r"(?m)^[ \t]*(?:pub(?:[ \t]*\([^()\r\n]*\))?[ \t]+)?use\b"
)
NEUTRAL_TABLE_EXTERN_CRATE = re.compile(
    r"(?m)^[ \t]*extern[ \t]+crate[ \t]+(?:r#)?"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
)
NEUTRAL_TABLE_PUBLIC_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t]+mod[ \t]+(?:r#)?"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
)

# Merged-cell geometry is a neutral table concern.  Keep the checker focused
# on the public declaration/import seam so the common owner can change its
# private representation without turning the ratchet into a body snapshot.
IWA_COMMON_TABLE_MERGE_SOURCE = NEUTRAL_TABLE_SOURCE_ROOT / "merge.rs"
IWA_COMMON_TABLE_MERGE_REQUIRED_TYPES = frozenset(
    {"AnchorRelocation", "Axis", "Deletion", "Error", "Region", "Result"}
)
IWA_COMMON_TABLE_MERGE_REQUIRED_FUNCTIONS = frozenset(
    {"after_deletion", "after_insertion", "anchor_relocation_after_deletion"}
)
IWA_COMMON_TABLE_MERGE_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    + "|".join(
        re.escape(name)
        for name in (
            "litchi_iwa_archive",
            "litchi_iwa_core",
            "litchi_iwa_protos",
            "litchi_numbers",
            "litchi_pages",
            "litchi_keynote",
            "buffa",
            "prost",
            "prost_types",
            "archive",
            "protobuf",
            "proto",
            "wire",
        )
    )
    + r")(?![A-Za-z0-9_])"
)
IWA_COMMON_TABLE_MERGE_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?merge\b"
)
IWA_COMMON_TABLE_MERGE_REEXPORT = re.compile(
    r"(?ms)^[ \t]*pub[ \t\r\n]+use[ \t\r\n]+"
    r"litchi_iwa_common[ \t\r\n]*::[ \t\r\n]*table"
    r"[ \t\r\n]*::[ \t\r\n]*merge\b.*?;"
)
IWA_COMMON_TABLE_MERGE_LOCAL_DECLARATION = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
    r"(?:struct|enum|union|type)[ \t]+(?:r#)?"
    r"(?:AnchorRelocation|Axis|Deletion|Region)\b"
)
IWA_COMMON_TABLE_MERGE_LOCAL_FUNCTION = re.compile(
    r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+(?:r#)?"
    r"(?:after_deletion|after_insertion|anchor_relocation_after_deletion)\b"
)
IWA_FOCUSED_TABLE_MERGES_WIRE_IMPORT = re.compile(
    r"\blitchi_numbers_wire[ \t\r\n]*::[ \t\r\n]*table_merges\b"
)
IWA_FOCUSED_TABLE_MERGES_COMMON_IMPORT = re.compile(
    r"(?:\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*(?:"
    r"table[ \t\r\n]*::[ \t\r\n]*merge\b|"
    r"\{[^{};]*\btable[ \t\r\n]*::[ \t\r\n]*merge\b"
    r")|\bcrate[ \t\r\n]*::[ \t\r\n]*table[ \t\r\n]*::"
    r"[ \t\r\n]*merge\b)"
)
IWA_FOCUSED_TABLE_MERGES_GENERATED_AST = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"FormulaArchive|AstNodeArchive|AstNodeType|MergeRegionMapArchive|"
    r"numbers_formula_codec|litchi_iwa_protos|prost::Message|"
    r"prost_types::"
    r")(?![A-Za-z0-9_])"
)
IWA_FOCUSED_TABLE_MERGES_RAW_ID_PARAMETER = re.compile(
    r"\b(?:table|model|drawable|object|native|slide|sheet)?"
    r"_?(?:id|identifier)\b[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u8|u16|u32|u64|u128|usize|i8|i16|i32|i64|i128|isize)\b"
)
IWA_FOCUSED_TABLE_MERGES_API_METHODS = (
    "body_table_merges",
    "slide_table_merges",
)
PAGES_TABLE_MERGE_SOURCE = Path(
    "crates/litchi-pages/src/package/body_table_merges.rs"
)
KEYNOTE_TABLE_MERGE_SOURCE = Path(
    "crates/litchi-keynote/src/package/slide_table_merges.rs"
)
PAGES_TABLE_MERGE_PACKAGE_SOURCE = Path("crates/litchi-pages/src/package.rs")
KEYNOTE_TABLE_MERGE_PACKAGE_SOURCE = Path("crates/litchi-keynote/src/package.rs")
PAGES_TABLE_MERGE_COMMON_TABLE_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t]+mod[ \t]+(?:r#)?merge\b"
)
PAGES_TABLE_MERGE_FOCUSED_CALL = re.compile(
    r"\b(?:read_table_merges|table_merges::[ \t\r\n]*read_table_merges)\b"
)
# Numbers uses the same focused merge reader shape, but its package owns both
# selectors.  Keep the native BNC vocabulary private to the wire adapter even
# while the package retains its raw-ID compatibility reader elsewhere.
NUMBERS_TABLE_MERGE_SOURCE = Path(
    "crates/litchi-numbers/src/package/table_merges.rs"
)
NUMBERS_TABLE_MERGE_PACKAGE_SOURCE = Path("crates/litchi-numbers/src/package.rs")
NUMBERS_TABLE_MERGE_READER_FUNCTION_NAME = re.compile(
    r"\bfn[ \t\r\n]+(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
)
NUMBERS_TABLE_MERGE_READER_BYTE_RETURN = re.compile(
    r"(?<![A-Za-z0-9_])(?:Vec|Box|Arc)[ \t]*<[ \t]*(?:u8|\[[ \t]*u8[ \t]*\])"
    r"|\b(?:Bytes|ByteString)\b"
)
NUMBERS_TABLE_MERGE_READER_PUBLIC_FIELD = re.compile(
    r"(?:^|,)[ \t\r\n]*pub(?:[ \t]*\([^()\r\n]*\))?[ \t]+"
    r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*:"
)
FOCUSED_TABLE_MERGE_READER_ALLOWED_METHODS = frozenset(
    {
        "open",
        "open_with_options",
        "from_bytes",
        "from_bytes_with_options",
        "from_shared_bytes",
        "from_shared_bytes_with_options",
        "table_merges",
        "body_table_merges",
        "slide_table_merges",
        "read_options",
        "options",
    }
)
FOCUSED_TABLE_MERGE_READER_FORBIDDEN_TYPES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"FormulaArchive|AstNodeArchive|AstNodeType|MergeRegionMapArchive|"
    r"DocumentArchive|BodyArchive|ShowArchive|SlideArchive|SheetArchive|"
    r"TableInfoArchive|TableModelArchive|"
    r"BncCell|BncCellView|PreBncCellView|StoredValue|CachedScalar|"
    r"numbers_formula_codec|numbers_table_cell_storage_codec|"
    r"litchi_iwa_protos|prost|prost_types|buffa"
    r")(?![A-Za-z0-9_])"
)
FOCUSED_TABLE_MERGE_READER_ARC_BYTES = re.compile(
    r"\bArc[ \t]*<[ \t]*\[[ \t]*u8[ \t]*\][ \t]*>"
)
FOCUSED_TABLE_MERGE_READER_PUBLIC_PHYSICAL_TYPES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"Archive|ArchiveObject|ComponentCatalog|RawMessage|SourceCatalog|"
    r"Package|Document|Body|Show|Slide|Sheet"
    r")(?![A-Za-z0-9_])"
)
FOCUSED_TABLE_MERGE_READER_PATH_INPUT = re.compile(
    r"\bAsRef[ \t]*<[ \t]*Path[ \t]*>"
)
FOCUSED_TABLE_MERGE_READER_BYTE_INPUT = re.compile(
    r"&[ \t\r\n]*(?:'[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]+)?"
    r"(?:mut[ \t\r\n]+)?\[[ \t]*u8[ \t]*\]"
)
FOCUSED_TABLE_MERGE_READER_MODEL_REFERENCE = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"crate[ \t\r\n]*::[ \t\r\n]*Document|"
    r"Document|Body|Show|Slide|Sheet|"
    r"(?:full|semantic)[ \t_]*(?:document|model)|"
    r"Document[ \t_]*(?:snapshot|model)"
    r")(?![A-Za-z0-9_])",
)
PAGES_TABLE_MERGE_FORBIDDEN_MODULES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"litchi_iwa_protos|prost|buffa|"
    r"litchi_numbers|litchi_keynote|FormulaArchive|AstNodeArchive|AstNodeType|"
    r"MergeRegionMapArchive"
    r")(?![A-Za-z0-9_])"
)

# Focused merge edits keep the public lifecycle semantic and selector-first;
# their private transaction children may own archive/source proof and the
# candidate reassembly machinery.  Keep the three paths explicit so a future
# inline helper cannot silently become a second transaction owner.
IWA_TABLE_MERGE_TRANSACTION_SPECS = (
    {
        "format_name": "Numbers",
        "owner_path": NUMBERS_TABLE_MERGE_SOURCE,
        "transaction_path": Path(
            "crates/litchi-numbers/src/package/table_merges/transaction.rs"
        ),
        "package_path": NUMBERS_TABLE_MERGE_PACKAGE_SOURCE,
        "library_path": Path("crates/litchi-numbers/src/lib.rs"),
        "module_name": "table_merges",
        "transaction_module": "transaction",
        # ADR 0034: the family carries short names and is
        # published through one focused module instead of prefixed aliases.
        "prefix": "",
        "focused_module": Path("crates/litchi-numbers/src/table/merge.rs"),
        "edit_method": "edit_table_merges",
        "apply_method": "apply_table_merges",
        "selector_types": frozenset({"SheetSelector", "TableSelector"}),
        "admission_markers": ("select_semantic_positions", "select_table"),
        # Numbers keeps edit entry O(1): the owned selectors are admitted by
        # the constructor, while the rooted target is resolved on the first
        # staged/read operation.  Every operation that can stage or publish
        # must therefore retain the private loader guard.
        "deferred_admission": {
            "helper": "ensure_loaded",
            "stage_methods": ("regions", "merge", "unmerge", "commit"),
        },
    },
    {
        "format_name": "Pages",
        "owner_path": PAGES_TABLE_MERGE_SOURCE,
        "transaction_path": Path(
            "crates/litchi-pages/src/package/body_table_merges/transaction.rs"
        ),
        "package_path": PAGES_TABLE_MERGE_PACKAGE_SOURCE,
        "library_path": Path("crates/litchi-pages/src/lib.rs"),
        "module_name": "body_table_merges",
        "transaction_module": "transaction",
        "prefix": "BodyTableMerges",
        # ADR 0034: canonical short names published only by
        # ``litchi_pages::table::merge::transaction``.
        "type_names": {
            suffix: suffix for suffix in ("Edit", "Patch", "Commit", "Diagnostics")
        },
        "type_display_prefix": "table::merge::transaction::",
        "transaction_family_owner": "body_table_merges",
        "edit_method": "edit_body_table_merges",
        "apply_method": "apply_body_table_merges",
        "selector_types": frozenset({"BodyTableSelector"}),
        "admission_markers": (
            "resolve_body_table_with_budget",
            "select_body_table",
            "resolve_body_table",
            # The lazy Edit resolver delegates through the owned typed
            # selector; this is the target-resolution marker in that helper.
            "resolve_with_budget",
        ),
        # Pages also defers rooted selection until the first operation.  The
        # private resolver performs the typed target lookup and readback; all
        # staging/publication methods must call it before touching state.
        "deferred_admission": {
            "helper": "resolve",
            "stage_methods": ("regions", "merge", "unmerge", "commit"),
        },
    },
    {
        "format_name": "Keynote",
        "owner_path": KEYNOTE_TABLE_MERGE_SOURCE,
        "transaction_path": Path(
            "crates/litchi-keynote/src/package/slide_table_merges/transaction.rs"
        ),
        "package_path": KEYNOTE_TABLE_MERGE_PACKAGE_SOURCE,
        # ADR 0034: the family's public home is
        # `slide::table::merge::transaction` with unprefixed names.
        "package_export_path": Path("crates/litchi-keynote/src/slide/table/merge.rs"),
        "library_path": Path("crates/litchi-keynote/src/slide/table/merge.rs"),
        "module_name": "slide_table_merges",
        "transaction_module": "transaction",
        "prefix": "",
        "edit_method": "edit_slide_table_merges",
        "apply_method": "apply_slide_table_merges",
        "selector_types": frozenset({"SlideSelector", "TableSelector"}),
        "admission_markers": ("select_merges", "select_table", "select_slide"),
    },
)
IWA_TABLE_MERGE_TRANSACTION_CANONICAL_SUFFIXES = (
    "Edit",
    "Patch",
    "Commit",
    "Diagnostics",
)
IWA_TABLE_MERGE_TRANSACTION_PUBLIC_FIELD = re.compile(
    r"(?m)^[ \t]*pub(?![ \t]*\([^()\r\n]*\))[ \t]+"
    r"(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*:"
)
IWA_TABLE_MERGE_TRANSACTION_PUBLIC_RAW_ID = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?(?:id|identifier|table_id|table_identifier|"
    r"model_id|model_identifier|native_id|object_id|drawable_id|"
    r"component_id|archive_id|message_id|sheet_id|sheet_identifier|"
    r"slide_id|slide_identifier|source_bytes|package_bytes|raw_bytes)"
    r"[ \t\r\n]*:[ \t\r\n]*"
    r"(?:Option[ \t\r\n]*<[ \t\r\n]*)?"
    r"(?:u8|u16|u32|u64|u128|usize|i8|i16|i32|i64|i128|isize)\b"
)
IWA_TABLE_MERGE_TRANSACTION_PUBLIC_BYTE_RETURN = re.compile(
    r"(?:Vec|Box|Arc)[ \t\r\n]*<[ \t\r\n]*(?:u8|\[[ \t\r\n]*u8\s*\])"
    r"|\b(?:Bytes|ByteString)\b"
)
IWA_TABLE_MERGE_TRANSACTION_FORBIDDEN_PUBLIC_TYPES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"Archive|ArchiveObject|ComponentCatalog|EntryEdit|ExactArtifacts|"
    r"OwnedExactArtifacts|RawMessage|SnappyStream|SourceCatalog|SharedBytes|"
    r"PhysicalSource|Document|Body|Show|Slide|Sheet|TableModelArchive|"
    r"TableInfoArchive|FormulaArchive|AstNodeArchive|AstNodeType|"
    r"MergeRegionMapArchive|litchi_iwa_protos|prost|prost_types|buffa|"
    r"BncCell|BncCellView|PreBncCellView|StoredValue|CachedScalar|"
    r"numbers_formula_codec|numbers_table_cell_storage_codec"
    r")(?![A-Za-z0-9_])"
)
IWA_TABLE_MERGE_TRANSACTION_LEGACY_MUTATION = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"merge_table_cells_in_package|unmerge_table_cells_in_package|"
    r"table_cell_merges_in_package|regions_in_package|"
    r"legacy_table_merge|fallback_table_merge"
    r")(?![A-Za-z0-9_])[ \t\r\n]*\("
)

# The Numbers wire writer is the only raw-byte merge rewrite surface.  Its
# byte-returning result is deliberately allowed here; this audit instead
# keeps the writer source-preserving, bounded, and format-neutral.
IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_SOURCE = Path(
    "crates/litchi-numbers-wire/src/table_merges/writer.rs"
)

# The narrow Buffa formula encoder is optional while its source migration is
# in flight.  Once either the module or its writer import appears, the audit
# becomes strict and prevents a handwritten TSCE schema from returning.
IWA_TABLE_MERGE_FORMULA_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/table_merge_formula_codec.rs"
)
IWA_TABLE_MERGE_FORMULA_CODEC_MODULE = "table_merge_formula_codec"
IWA_TABLE_MERGE_FORMULA_CODEC_REQUIRED_TYPES = (
    "EncodeOptions",
    "EncodeReport",
    "EncodeLimit",
    "EncodeError",
)
IWA_TABLE_MERGE_FORMULA_CODEC_REQUIRED_FUNCTIONS = (
    "encode_merge_formula",
)
IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_SOURCE = Path(
    "crates/litchi-iwa-protos/build.rs"
)
IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_RERUN_MARKER = (
    'println!("cargo:rerun-if-changed=src/table_merge_formula_codec.rs");'
)
IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_INVENTORY = re.compile(
    r"(?s)\(\s*"
    r'"src/table_merge_formula_codec\.rs"\s*,\s*'
    r'"crate::buffa_formula_generated::"\s*,\s*'
    r'"mod buffa_formula_generated \{"\s*,?\s*\)'
)
IWA_TABLE_MERGE_FORMULA_CODEC_HIDDEN_MODULE = re.compile(
    rf"(?m)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
    rf"pub\s+mod\s+{re.escape(IWA_TABLE_MERGE_FORMULA_CODEC_MODULE)}\b"
)
IWA_TABLE_MERGE_FORMULA_CODEC_ENCODER = re.compile(
    r"(?m)^\s*pub\s+fn\s+(?:[A-Za-z_][A-Za-z0-9_]*_)?"
    r"(?:merge|table_merge|formula)[A-Za-z0-9_]*encode[A-Za-z0-9_]*\b|"
    r"^\s*pub\s+fn\s+encode[A-Za-z0-9_]*(?:merge|table_merge|formula)"
    r"[A-Za-z0-9_]*\b"
)

# The metadata and dependency codecs keep their raw-wire inputs private to a
# hidden protos facade, but their structural removals are still separate child
# modules.  These audits activate as soon as the child declaration or source
# appears and require Buffa lazy validation, bounded planning, and exact
# candidate verification without assuming one concrete helper spelling.
IWA_PACKAGE_METADATA_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/package_metadata_codec.rs"
)
IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_SOURCE = Path(
    "crates/litchi-iwa-protos/src/package_metadata_codec/component_removal.rs"
)
IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_MODULE = "component_removal"
IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_BUILD_MARKER = (
    'println!("cargo:rerun-if-changed=src/package_metadata_codec.rs");'
)
IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_PUBLIC_PREPARED = re.compile(
    r"(?m)^\s*pub\s+struct\s+Prepared[A-Za-z0-9_]*"
    r"(?:Component|component)[A-Za-z0-9_]*(?:Removal|removal)[A-Za-z0-9_]*\b"
)
IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_PREPARE_FUNCTION = re.compile(
    r"(?m)^\s*pub\s+fn\s+prepare_[A-Za-z0-9_]*"
    r"(?:component|Component)[A-Za-z0-9_]*(?:removal|Removal)[A-Za-z0-9_]*\b"
)
IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_MARKER_GROUPS = {
    "whole component registration": (
        "ComponentRemovalBatch",
        "component registration",
        "current-component",
        "component_removal",
        "component_header",
    ),
    "prepared execution": (
        "Prepared",
        "prepare_",
        "execute",
        "rewrite_",
    ),
    "Buffa lazy view": (
        "buffa",
        "Buffa",
        "decode_lazy_view",
        "LazyView",
    ),
    "source-preserving candidate": (
        "source",
        "source-preserving",
        "unknown_fields",
        "candidate",
        "Verification",
        "validate_candidate",
    ),
    "bounded accounting": (
        "RewriteExecutionRequirements",
        "preflight_execution",
        "try_reserve",
        "checked_add",
        "Budget",
        "max_output",
    ),
}
IWA_DEPENDENCY_CODEC_PUBLIC_PHYSICAL = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"prost|prost_types|OwnedView|RawMessage|ArchiveObject|"
    r"(?:[A-Za-z_][A-Za-z0-9_]*Generated|[A-Za-z_][A-Za-z0-9_]*ArchiveView)"
    r")(?![A-Za-z0-9_])"
)
IWA_CODEC_PRODUCTION_GUARD_SOURCE = Path(
    "crates/litchi-iwa-protos/src/production_codec_guard.rs"
)
IWA_CODEC_BUILD_SOURCE = Path("crates/litchi-iwa-protos/build.rs")


# Selected table-cell reads share the same three-layer boundary as merged
# geometry: the common crate publishes only archive-free semantic results, the
# Numbers wire crate owns borrowed storage/sidecar traversal, and each focused
# format adapter proves its graph selection while delegating both walks through
# one aggregate budget.  Keep these checks at the public/source seam.  Private
# adapter helpers are intentionally allowed to mention generated snapshots and
# native identifiers; the public selector-first method must not leak them.
IWA_COMMON_TABLE_READ_SOURCE = NEUTRAL_TABLE_SOURCE_ROOT / "read.rs"
IWA_COMMON_TABLE_READ_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?read\b"
)
IWA_COMMON_TABLE_READ_PUBLIC_TYPES = frozenset(
    {"CellComment", "Comment", "CommentAuthor", "CommentReply", "CommentTimestamp", "TableRead", "Builder"}
)
IWA_COMMON_TABLE_READ_IMPORT_STATEMENT = re.compile(
    r"(?ms)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
    r"(?:use|extern[ \t]+crate)\b.*?;"
)
IWA_COMMON_TABLE_READ_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"litchi_iwa_(?:archive|core|protos|package|structured|text_wire)|"
    r"litchi_numbers(?:_wire)?|litchi_pages|litchi_keynote|"
    r"buffa|prost(?:_types)?|archive|protobuf|proto|wire"
    r")(?![A-Za-z0-9_])"
)
IWA_TABLE_READ_PUBLIC_PHYSICAL_TYPES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"Archive(?:Object)?|RawMessage|HostPackage|Native(?:Id|Object)|"
    r"(?:Object|Package|Table|DataStore|Tile|Sidecar)Id|"
    r"(?:TableModel|DataStore|Native|Tile|Sidecar)Snapshot|"
    r"DecodeOptions|Wire(?:View|Limits|FieldView|Descent)|"
    r"(?:Formula|Ast)Archive|prost(?:_types)?|buffa"
    r")(?![A-Za-z0-9_])"
)
IWA_TABLE_READ_RAW_ID_PARAMETER = re.compile(
    r"\b(?:table|model|drawable|object|native|slide|sheet|tile|storage|"
    r"sidecar|package|data_store|cell|row|column)?"
    r"_?(?:id|identifier)\b[ \t\r\n]*:[ \t\r\n]*"
    r"(?:u8|u16|u32|u64|u128|usize|i8|i16|i32|i64|i128|isize)\b"
)
IWA_NUMBERS_WIRE_TABLE_CELLS_PEER_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:litchi_numbers|litchi_pages|litchi_keynote)"
    r"(?![A-Za-z0-9_])"
)
IWA_FOCUSED_TABLE_CELLS_COMMON_IMPORT = re.compile(
    r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*table[ \t\r\n]*::"
    r"[ \t\r\n]*read\b"
)
IWA_FOCUSED_TABLE_CELLS_WIRE_IMPORT = re.compile(
    r"\blitchi_numbers_wire[ \t\r\n]*::[ \t\r\n]*table_cells\b"
)
IWA_FOCUSED_TABLE_CELLS_SIDECAR_IMPORT = re.compile(
    r"\blitchi_numbers_wire[ \t\r\n]*::[ \t\r\n]*table_sidecars\b"
)
IWA_FOCUSED_TABLE_CELLS_WIRE_CALL = re.compile(r"\b(?:[A-Za-z_][A-Za-z0-9_]*::)?read_table_cells\s*\(")
IWA_FOCUSED_TABLE_CELLS_SIDECAR_CALL = re.compile(r"\b(?:[A-Za-z_][A-Za-z0-9_]*::)?read_sidecar_list\s*\(")
IWA_FOCUSED_TABLE_CELLS_COMMON_CONSTRUCTOR = re.compile(
    r"\bTableRead[ \t\r\n]*::[ \t\r\n]*"
    r"(?:builder|new|try_from_(?:owned_parts|comments|parts))"
    r"[ \t\r\n]*\("
)
IWA_FOCUSED_TABLE_CELLS_PEER_IMPORT = IWA_NUMBERS_WIRE_TABLE_CELLS_PEER_IMPORT
IWA_FOCUSED_TABLE_CELLS_GENERATED_PUBLIC = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"RawMessage|TableModelSnapshot|DataStoreSnapshot|NativeMessages|"
    r"TileReference|SidecarReference|Sidecar(?:List|Tables|Value)|"
    r"DecodeOptions|Wire(?:View|Limits|FieldView|Descent)|"
    r"FormulaArchive|AstNode(?:Archive|Type)|numbers_[A-Za-z0-9_]+_codec|"
    r"litchi_iwa_protos|prost(?:_types)?|buffa"
    r")(?![A-Za-z0-9_])"
)
IWA_FOCUSED_TABLE_CELLS_BUDGET_IMPL = re.compile(
    r"(?m)^\s*impl\s+(?P<trait>TableCellReadBudget|SidecarReadBudget)"
    r"\s+for\s+(?P<owner>[A-Za-z_][A-Za-z0-9_]*)\b"
)

# Chart metadata is a read-only semantic value shared by all three concrete
# format crates.  Keep the common value archive-free, the strict Buffa
# projection private to the protos crate, and the focused package methods
# selector-first.  This audit deliberately checks source/import/public-seam
# invariants rather than implementation helper names so the host can be
# removed later without weakening the format boundary.
IWA_COMMON_CHART_METADATA_SOURCE = Path(
    "crates/litchi-iwa-common/src/chart/metadata.rs"
)
IWA_COMMON_CHART_METADATA_MODULE_SOURCE = Path(
    "crates/litchi-iwa-common/src/chart.rs"
)
IWA_COMMON_CHART_METADATA_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?metadata\b"
)
IWA_COMMON_CHART_METADATA_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"litchi_iwa_(?:archive|core|protos|structured|text_wire)|"
    r"litchi_numbers(?:_wire)?|litchi_pages|litchi_keynote|"
    r"buffa|prost(?:_types)?|archive|protobuf|proto|wire"
    r")(?![A-Za-z0-9_])"
)
IWA_CHART_METADATA_PUBLIC_PHYSICAL_TYPES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"Archive(?:Object)?|RawMessage|HostPackage|Native(?:Id|Object)|"
    r"(?:Object|Package|Chart|Drawable|Message|Component)Id|"
    r"(?:Chart|Drawable|Message|Component)Archive|"
    r"DecodeOptions|DecodeReport|Wire(?:View|Limits|FieldView|Descent)|"
    r"FormulaArchive|prost(?:_types)?|buffa|litchi_iwa_protos"
    r")(?![A-Za-z0-9_])"
)
IWA_CHART_METADATA_RAW_ID_PARAMETER = re.compile(
    r"\b(?:table|model|drawable|object|native|slide|sheet|chart|"
    r"component|archive|message|source|package|metadata)?"
    r"_?(?:id|identifier)\b[ \t\r\n]*:[ \t\r\n]*"
    r"(?:NonZero(?:U|I)(?:8|16|32|64|128)|u8|u16|u32|u64|u128|"
    r"usize|i8|i16|i32|i64|i128|isize|Option[ \t\r\n]*<[^>]*\b(?:u|i)[0-9]+\b)"
    r"\b"
)

IWA_CHART_METADATA_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/chart_metadata_codec.rs"
)
IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE = IWA_PROTOS_FACADE_SOURCE
IWA_CHART_METADATA_CODEC_MODULE = "chart_metadata_codec"
IWA_CHART_METADATA_CODEC_GENERATED_MODULE = "buffa_chart_metadata_generated"
IWA_CHART_METADATA_CODEC_HIDDEN_MODULE = re.compile(
    rf"(?m)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
    rf"pub\s+mod\s+{re.escape(IWA_CHART_METADATA_CODEC_MODULE)}\b"
)
IWA_CHART_METADATA_CODEC_REQUIRED_TYPES = (
    "ChartMetadataSnapshot",
    "DecodeOptions",
    "DecodeReport",
)
IWA_CHART_METADATA_CODEC_REQUIRED_MARKER_GROUPS = {
    "Buffa lazy ingress": ("buffa", "decode_lazy_view"),
    "modern and legacy decode routes": (
        "decode_modern",
        "decode_modern_with_report",
        "decode_legacy",
        "decode_legacy_with_report",
    ),
    "strict preflight": ("preflight_modern", "preflight_legacy"),
}
IWA_CHART_METADATA_CODEC_PROST = re.compile(
    r"(?<![A-Za-z0-9_])prost(?:_types)?(?![A-Za-z0-9_])"
)

IWA_CHART_METADATA_FOCUSED_OWNERS = {
    "Pages": Path("crates/litchi-pages/src/package/body_chart_metadata.rs"),
    "Numbers": Path("crates/litchi-numbers/src/package/chart_metadata.rs"),
    "Keynote": Path("crates/litchi-keynote/src/package/slide_chart_metadata.rs"),
}
IWA_CHART_METADATA_FOCUSED_PACKAGES = {
    "Pages": Path("crates/litchi-pages/src/package.rs"),
    "Numbers": Path("crates/litchi-numbers/src/package.rs"),
    "Keynote": Path("crates/litchi-keynote/src/package.rs"),
}
IWA_CHART_METADATA_FOCUSED_MODULES = {
    "Pages": "body_chart_metadata",
    "Numbers": "chart_metadata",
    "Keynote": "slide_chart_metadata",
}
IWA_CHART_METADATA_FOCUSED_METHODS = {
    "Pages": "body_chart_metadata",
    "Numbers": "sheet_chart_metadata",
    "Keynote": "slide_chart_metadata",
}
IWA_CHART_METADATA_COMMON_IMPORT = re.compile(
    r"(?:\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*chart"
    r"[ \t\r\n]*::[ \t\r\n]*metadata\b|"
    r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*chart"
    r"[ \t\r\n]*::[ \t\r\n]*\{[^{};]*\bmetadata\b)"
)
IWA_CHART_METADATA_CODEC_IMPORT = re.compile(
    r"\blitchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*chart_metadata_codec\b"
)
IWA_CHART_METADATA_CODEC_CALL = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
    r"chart_metadata_codec[ \t\r\n]*::"
)
IWA_CHART_METADATA_FOCUSED_MODULE = re.compile(
    r"(?m)^[ \t]*(?P<visibility>pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t]+(?:r#)?(?P<name>body_chart_metadata|"
    r"chart_metadata|sheet_chart_metadata|slide_chart_metadata)\b[ \t]*(?:;|\{)"
)
IWA_CHART_METADATA_FOCUSED_FORBIDDEN_PROST = re.compile(
    r"(?<![A-Za-z0-9_])(?:prost|prost_types|"
    r"(?:Chart|ChartInfo|ChartDrawable|ChartGrid)Archive[ \t\r\n]*::[ \t\r\n]*decode)"
    r"(?![A-Za-z0-9_])"
)
IWA_CHART_METADATA_HOST_LOCAL_TYPE = re.compile(
    r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|type)[ \t]+"
    r"(?:r#)?ChartMetadata\b"
)
IWA_CHART_METADATA_HOST_OLD_IMPORT = re.compile(
    r"\b(?:crate|self|super)[ \t\r\n]*::[ \t\r\n]*protobuf"
    r"[ \t\r\n]*::[ \t\r\n]*tsch\b|\bprost[ \t\r\n]*::[ \t\r\n]*Message\b"
)
IWA_CHART_METADATA_HOST_GENERATED_DECODE = re.compile(
    r"\b(?:Chart|ChartInfo|ChartDrawable|ChartGrid)Archive[ \t\r\n]*::"
    r"[ \t\r\n]*decode\s*\("
)
IWA_CHART_METADATA_HOST_CODEC_CALL = re.compile(
    r"\b(?:[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
    r"chart_metadata_codec[ \t\r\n]*::"
)

# Chart data follows the same ownership split as chart metadata, but its
# common value is a validated rectangular grid.  Keep this ratchet separate
# from metadata: a metadata-only reader must not accidentally make the data
# surface look migrated, and a future removal of the compatibility host must
# not require an otherwise obsolete host shim to remain in the tree.
IWA_COMMON_CHART_DATA_SOURCE = Path(
    "crates/litchi-iwa-common/src/chart/data.rs"
)
IWA_COMMON_CHART_DATA_MODULE_SOURCE = Path(
    "crates/litchi-iwa-common/src/chart.rs"
)
IWA_COMMON_CHART_DATA_MODULE = re.compile(
    r"(?m)^[ \t]*pub[ \t\r\n]+mod[ \t\r\n]+(?:r#)?data\b"
)
IWA_COMMON_CHART_DATA_FORBIDDEN_IMPORT = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"litchi_iwa_(?:archive|core|protos|structured|text_wire)|"
    r"litchi_numbers(?:_wire)?|litchi_pages|litchi_keynote|"
    r"buffa|prost(?:_types)?|archive|protobuf|proto|wire"
    r")(?![A-Za-z0-9_])"
)
IWA_CHART_DATA_PUBLIC_PHYSICAL_TYPES = re.compile(
    r"(?<![A-Za-z0-9_])(?:"
    r"Archive(?:Object)?|RawMessage|HostPackage|Native(?:Id|Object)|"
    r"(?:Object|Package|Chart|Drawable|Message|Component)Id|"
    r"(?:Chart|Drawable|Message|Component|Grid|Value)Archive|"
    r"ChartDataSnapshot|DecodeError|DecodeLimit|LabelList|ValueList|"
    r"DecodeOptions|DecodeReport|Wire(?:View|Limits|FieldView|Descent)|"
    r"FormulaArchive|prost(?:_types)?|buffa|litchi_iwa_protos"
    r")(?![A-Za-z0-9_])"
)
IWA_CHART_DATA_RAW_ID_PARAMETER = re.compile(
    r"\b(?:table|model|drawable|object|native|slide|sheet|chart|"
    r"component|archive|message|source|package|metadata|grid|row|column)?"
    r"_?(?:id|identifier)\b[ \t\r\n]*:[ \t\r\n]*"
    r"(?:NonZero(?:U|I)(?:8|16|32|64|128)|u8|u16|u32|u64|u128|"
    r"usize|i8|i16|i32|i64|i128|isize|Option[ \t\r\n]*<[^>]*\b(?:u|i)[0-9]+\b)"
    r"\b"
)

IWA_CHART_DATA_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/chart_data_codec.rs"
)
IWA_CHART_DATA_CODEC_PROJECTION_SOURCE = Path(
    "crates/litchi-iwa-protos/src/buffa-projections/TSCHChartDataArchive.proto"
)
IWA_CHART_DATA_CODEC_PUBLIC_SOURCE = IWA_PROTOS_FACADE_SOURCE
IWA_CHART_DATA_CODEC_MODULE = "chart_data_codec"
IWA_CHART_DATA_CODEC_GENERATED_MODULE = "buffa_chart_data_generated"
IWA_CHART_DATA_CODEC_HIDDEN_MODULE = re.compile(
    rf"(?m)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
    rf"pub\s+mod\s+{re.escape(IWA_CHART_DATA_CODEC_MODULE)}\b"
)
IWA_CHART_DATA_CODEC_REQUIRED_TYPES = (
    "ChartDataSnapshot",
    "DecodeOptions",
    "DecodeReport",
)
IWA_CHART_DATA_CODEC_REQUIRED_MARKER_GROUPS = {
    "Buffa lazy ingress": ("buffa", "decode_lazy_view"),
    "chart-data decode route": (
        ("decode_chart_data", "decode_modern", "decode_grid"),
    ),
    "strict preflight": (("preflight_modern", "preflight_grid", "preflight"),),
}
IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_TYPES = (
    "RewriteError",
    "RewriteExecutionRequirements",
)
IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_TYPE_GROUPS = {
    # Keep the prepared-plan name flexible.  The request itself may be a
    # borrowed rectangular slice or a common semantic value, so the boundary
    # must not force one private input representation.
    "prepared rewrite type": (
        "PreparedChartDataRewrite",
        "PreparedChartDataEdit",
        "PreparedRewrite",
    ),
}
IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_FUNCTIONS = (
    "prepare_chart_data_rewrite",
)
IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_MARKER_GROUPS = {
    "source-preserving rewrite preparation": (
        "prepare_chart_data_rewrite",
        "RewriteExecutionRequirements",
    ),
    "exact rewrite accounting": (
        "execution_requirements",
        "RewriteError",
    ),
    # A focused chart-data writer may implement the wire walk in a private
    # helper or a nested rewrite module.  A generated Buffa encoder is one
    # valid implementation; a manual bounded emitter is also valid when the
    # prepared plan performs an authoritative private Buffa lazy readback.
    "bounded Buffa production validation": (
        ("ViewEncode", "try_encoded_len", "try_encode_bounded"),
        ("decode_lazy_view", "BuffaDecodeOptions"),
    ),
}
IWA_CHART_DATA_CODEC_PROST = re.compile(
    r"(?<![A-Za-z0-9_])prost(?:_types)?(?![A-Za-z0-9_])"
)

IWA_CHART_DATA_FOCUSED_OWNERS = {
    "Pages": Path("crates/litchi-pages/src/package/body_chart_data.rs"),
    "Numbers": Path("crates/litchi-numbers/src/package/chart_data.rs"),
    "Keynote": Path("crates/litchi-keynote/src/package/slide_chart_data.rs"),
}
IWA_CHART_DATA_FOCUSED_PACKAGES = {
    "Pages": Path("crates/litchi-pages/src/package.rs"),
    "Numbers": Path("crates/litchi-numbers/src/package.rs"),
    "Keynote": Path("crates/litchi-keynote/src/package.rs"),
}
IWA_CHART_DATA_FOCUSED_MODULES = {
    "Pages": "body_chart_data",
    "Numbers": "chart_data",
    "Keynote": "slide_chart_data",
}
IWA_CHART_DATA_FOCUSED_METHODS = {
    "Pages": "body_chart_data",
    "Numbers": "sheet_chart_data",
    "Keynote": "slide_chart_data",
}
IWA_CHART_DATA_FOCUSED_EDIT_METHODS = {
    "Pages": "edit_body_chart_data",
    "Numbers": "edit_sheet_chart_data",
    "Keynote": "edit_slide_chart_data",
}
IWA_CHART_DATA_FOCUSED_APPLY_METHODS = {
    "Pages": "apply_body_chart_data",
    "Numbers": "apply_sheet_chart_data",
    "Keynote": "apply_slide_chart_data",
}
IWA_CHART_DATA_COMMON_IMPORT = re.compile(
    r"(?:"
    r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*chart"
    r"[ \t\r\n]*::[ \t\r\n]*(?:data|ChartData)\b|"
    r"\blitchi_iwa_common[ \t\r\n]*::[ \t\r\n]*chart"
    r"[ \t\r\n]*::[ \t\r\n]*\{[^{};]*\b(?:data|ChartData)\b)"
)
IWA_CHART_DATA_CODEC_IMPORT = re.compile(
    r"\blitchi_iwa_protos[ \t\r\n]*::[ \t\r\n]*chart_data_codec\b"
)
IWA_CHART_DATA_FOCUSED_MODULE = re.compile(
    r"(?m)^[ \t]*(?P<visibility>pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t]+(?:r#)?(?P<name>body_chart_data|chart_data|"
    r"sheet_chart_data|slide_chart_data)\b[ \t]*(?:;|\{)"
)
IWA_CHART_DATA_FOCUSED_FORBIDDEN_PROST = re.compile(
    r"(?<![A-Za-z0-9_])(?:prost|prost_types|"
    r"(?:Chart|ChartInfo|ChartDrawable|ChartGrid|Grid|Value)Archive"
    r"[ \t\r\n]*::[ \t\r\n]*decode)(?![A-Za-z0-9_])"
)

IWA_CHART_DATA_HOST_LOCAL_TYPE = re.compile(
    r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|type)[ \t]+"
    r"(?:r#)?ChartData\b"
)
IWA_CHART_DATA_HOST_COMMON_REEXPORT = re.compile(
    r"(?:\bpub[ \t]+use[ \t]+litchi_iwa_common[ \t\r\n]*::"
    r"[ \t\r\n]*chart[ \t\r\n]*::[ \t\r\n]*"
    r"(?:data[ \t\r\n]*::[ \t\r\n]*)?ChartData\b|"
    r"\bpub[ \t]+use[ \t]+litchi_iwa_common[ \t\r\n]*::"
    r"[ \t\r\n]*chart[ \t\r\n]*::[ \t\r\n]*\{[^{};]*\bChartData\b)"
)
IWA_CHART_DATA_HOST_LOCAL_REEXPORT = re.compile(
    r"\bpub[ \t]+use[ \t]+(?:crate|self|super)[ \t\r\n]*::"
    r"[ \t\r\n]*charts?[ \t\r\n]*::[ \t\r\n]*data"
    r"[ \t\r\n]*::[ \t\r\n]*ChartData\b"
)
IWA_CHART_DATA_HOST_COMPAT_REEXPORT = re.compile(
    r"\bpub[ \t]+use[ \t]+(?:r#)?data[ \t\r\n]*::[ \t\r\n]*"
    r"(?:r#)?ChartData\b"
)

# Fresh chart-grid authoring lives outside the compatibility host's
# generated model.  The read/rewrite chart-data codec above intentionally
# remains a separate owner: fresh grids have no source payload to rewrite and
# therefore need a bounded, borrowed encoder that emits the native wire
# representation directly.  Keep this boundary source-local so unrelated
# chart-property mutations in the host can continue to use their compatibility
# graph while the authoring seam is migrated.
IWA_CHART_GRID_CREATION_CODEC_SOURCE = Path(
    "crates/litchi-iwa-protos/src/chart_grid_creation_codec.rs"
)
IWA_CHART_GRID_CREATION_CODEC_PUBLIC_SOURCE = IWA_PROTOS_FACADE_SOURCE
IWA_CHART_GRID_CREATION_CODEC_MODULE = "chart_grid_creation_codec"
IWA_CHART_GRID_CREATION_CODEC_REQUIRED_TYPES = (
    "ChartGridCreationRequest",
    "PreparedChartGridCreation",
    "ExecutionRequirements",
    "ExecutionLimits",
    "EncodeError",
    "EncodeOutput",
)
IWA_CHART_GRID_CREATION_CODEC_REQUIRED_FUNCTIONS = (
    "prepare_chart_grid_creation",
    "encode_chart_grid",
)
IWA_CHART_GRID_CREATION_CODEC_REQUIRED_MARKER_GROUPS = {
    # The request must borrow the caller's semantic labels and values.  The
    # exact field layout is intentionally left to the Rust implementation;
    # these markers prove that the API does not take an eager generated grid.
    "borrowed semantic input": (("row_labels", "column_labels", "Option<f64>"),),
    "borrowed lifetime": (
        ("&'a",),
        ("&'source",),
        ("&[",),
        ("borrowed",),
    ),
    # A production encoder must preflight output and use the Buffa view seam;
    # a plain `Vec<u8>` builder is not sufficient evidence of this boundary.
    "bounded Buffa encoding": (
        "buffa",
        ("ViewEncode", "try_encode_bounded", "encode_bounded"),
        ("try_encoded_len", "encoded_len"),
    ),
    # Grid identity generation is part of the shared owner.  Keeping these
    # markers here prevents a host helper from silently rebuilding row/column
    # identity maps after the generated archive is removed from the host.
    "deterministic row-column identity": (
        ("ID_MAP_ROW_FIELD", "ID_MAP_COLUMN_FIELD", "seed"),
        ("row_id_map", "column_id_map", "seed"),
        ("seed", "identifier", "uuid"),
    ),
    "fallible shape validation": (
        ("plan_request", "row_labels", "column_labels"),
        ("validate", ("row_names", "row_labels"), ("column_names", "column_labels")),
    ),
}
IWA_CHART_GRID_CREATION_CODEC_PROST = re.compile(
    r"(?<![A-Za-z0-9_])prost(?:_types)?(?![A-Za-z0-9_])"
)
IWA_CHART_GRID_CREATION_CODEC_GENERATED_CONSTRUCTION = re.compile(
    r"\b(?:ChartGridArchive|GridRow|GridValue|ChartGridRowColumnIdMap)\s*\{"
)
IWA_CHART_GRID_CREATION_CODEC_GENERATED_OPERATION = re.compile(
    r"\b(?:ChartGridArchive|GridRow|GridValue|ChartGridRowColumnIdMap)\s*::\s*"
    r"(?:decode|encode|encode_to_vec)\b"
)
IWA_CHART_GRID_CREATION_CODEC_GENERATED_REFERENCE = re.compile(
    r"\b(?:ChartGridArchive|GridRow|GridValue|ChartGridRowColumnIdMap)\b"
)
IWA_CHART_GRID_CREATION_CODEC_EAGER_GENERATED_VECTOR = re.compile(
    r"\b(?:Vec|SmallVec|ArrayVec)\s*<[^>]*\b"
    r"(?:ChartGridArchive|GridRow|GridValue|ChartGridRowColumnIdMap)\b"
)
IWA_CHART_GRID_CREATION_CODEC_HIDDEN_MODULE = re.compile(
    rf"(?m)#\s*\[\s*doc\s*\(\s*hidden\s*\)\s*\][\s\r\n]*"
    rf"pub\s+mod\s+{re.escape(IWA_CHART_GRID_CREATION_CODEC_MODULE)}\b"
)
IWA_CHART_GRID_CREATION_FULL_REPLACE_METHOD = re.compile(
    r"\bfn\s+(?:set_body_chart_data_full_replace|"
    r"set_sheet_chart_data_full_replace|set_slide_chart_data_full_replace)\b"
)


def audit_iwa_table_cell_borders_source_topology(root: Path = ROOT) -> list[str]:
    """Keep the cell-border value in common and the old path compatibility-only."""

    violations: list[str] = []

    common_path = root / IWA_TABLE_CELL_BORDERS_COMMON_SOURCE
    if not common_path.is_file():
        violations.append(
            "common table-cell Borders owner is missing: "
            f"{IWA_TABLE_CELL_BORDERS_COMMON_SOURCE}"
        )
    else:
        common_source = _mask_rust_non_code(common_path.read_text(encoding="utf-8"))
        if re.search(r"\bpub[ \t]+struct[ \t]+Borders\b", common_source) is None:
            violations.append(
                "common table-cell Borders owner is missing its canonical struct: "
                f"{IWA_TABLE_CELL_BORDERS_COMMON_SOURCE}"
            )

    return sorted(set(violations))


def audit_neutral_table_model_source_topology(root: Path = ROOT) -> list[str]:
    """Protect neutral table ownership and the Numbers/Pages source boundary."""

    common_root = root / NEUTRAL_TABLE_SOURCE_ROOT
    if not common_root.is_dir():
        return []
    violations: list[str] = []

    manifest_path = root / NEUTRAL_TABLE_MANIFEST
    if not manifest_path.is_file():
        violations.append(f"neutral table model manifest is missing: {NEUTRAL_TABLE_MANIFEST}")
    else:
        try:
            manifest = tomllib.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError) as error:
            violations.append(f"neutral table model manifest cannot be parsed: {error}")
        else:
            cache: dict[Path, str | None] = {}
            for table_name, dependencies in _cargo_dependency_tables(manifest):
                for alias, specification in dependencies.items():
                    for candidate in _cargo_dependency_candidates(
                        manifest_path, table_name, alias, specification, cache
                    ):
                        if candidate.replace("_", "-") in NEUTRAL_TABLE_FORBIDDEN_DEPENDENCIES:
                            violations.append(
                                "neutral table model depends on concrete format/archive "
                                f"or wire crate {candidate}: {NEUTRAL_TABLE_MANIFEST} "
                                f"[{table_name}] alias={alias}"
                            )

    for relative_path, required in NEUTRAL_TABLE_REQUIRED_MODULES.items():
        path = root / relative_path
        if not path.is_file():
            violations.append(f"neutral table module source is missing: {relative_path}")
            continue
        source = _mask_rust_non_code(path.read_text(encoding="utf-8"))
        declared = {
            match.group("name") for match in NEUTRAL_TABLE_PUBLIC_MODULE.finditer(source)
        }
        for name in sorted(required - declared):
            violations.append(f"neutral table module is missing public {name}: {relative_path}")

    for relative_path, required in NEUTRAL_TABLE_REQUIRED_EXPORTS.items():
        path = root / relative_path
        if not path.is_file():
            violations.append(f"neutral table semantic owner source is missing: {relative_path}")
            continue
        source = path.read_text(encoding="utf-8")
        missing = required - _rust_canonical_exports(source, required)
        for name in sorted(missing):
            violations.append(f"neutral table semantic owner is missing public {name}: {relative_path}")
        for declaration, line_number in _rust_public_declarations(source):
            for match in RUST_IDENTIFIER.finditer(declaration):
                reason = _iwork_public_leak(match.group(1))
                if reason is None:
                    continue
                line = line_number + declaration.count("\n", 0, match.start(1))
                violations.append(
                    f"neutral table semantic owner exposes {reason} {match.group(1)}: "
                    f"{relative_path}:{line}"
                )
            for match in RUST_BYTE_SLICE.finditer(declaration):
                line = line_number + declaration.count("\n", 0, match.start())
                violations.append(
                    f"neutral table semantic owner exposes raw byte slice: "
                    f"{relative_path}:{line}"
                )

    # Mask trivia and apply the import ratchet to every table sidecar.
    for path in sorted(common_root.rglob("*.rs")):
        code = _mask_rust_non_code(path.read_text(encoding="utf-8"))
        for match in NEUTRAL_TABLE_FORBIDDEN_PATH.finditer(code):
            line = code.count("\n", 0, match.start()) + 1
            violations.append(
                f"neutral table source imports forbidden path: {path.relative_to(root)}:{line}"
            )
        for match in NEUTRAL_TABLE_FORBIDDEN_USE.finditer(code):
            end = code.find(";", match.end())
            statement = code[match.start() : len(code) if end < 0 else end]
            for identifier in RUST_IDENTIFIER.finditer(statement):
                if identifier.group(1) not in NEUTRAL_TABLE_FORBIDDEN_IMPORTS:
                    continue
                line = code.count("\n", 0, match.start() + identifier.start(1)) + 1
                violations.append(
                    f"neutral table source imports forbidden crate/module "
                    f"{identifier.group(1)}: {path.relative_to(root)}:{line}"
                )
        for match in NEUTRAL_TABLE_EXTERN_CRATE.finditer(code):
            if match.group("name") not in NEUTRAL_TABLE_FORBIDDEN_IMPORTS:
                continue
            line = code.count("\n", 0, match.start()) + 1
            violations.append(
                f"neutral table source declares forbidden external crate "
                f"{match.group('name')}: {path.relative_to(root)}:{line}"
            )

    # Numbers may wrap and alias the common model, but cannot re-define it.
    numbers_definition_checks = (
        (
            Path("crates/litchi-numbers/src/cell/mod.rs"),
            frozenset({"FiniteF64", "FiniteF64Error", "Type", "Update", "Value"}),
        ),
        (
            Path("crates/litchi-numbers/src/table/coordinate.rs"),
            frozenset({"AddressError", "At", "Area", "Error", "Result"}),
        ),
        (
            Path("crates/litchi-numbers/src/table.rs"),
            frozenset({"Cell", "Dimensions", "Error", "Grid", "View"}),
        ),
    )
    local_definition = re.compile(
        r"(?m)^\s*pub\s+(?:struct|enum|trait|union)\s+"
        r"(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
    )
    for relative_path, forbidden in numbers_definition_checks:
        path = root / relative_path
        if not path.is_file():
            continue
        source = _mask_rust_non_code(path.read_text(encoding="utf-8"))
        for match in local_definition.finditer(source):
            if match.group("name") not in forbidden:
                continue
            line = source.count("\n", 0, match.start()) + 1
            violations.append(
                f"Numbers facade re-defines neutral table type {match.group('name')}: "
                f"{relative_path}:{line}"
            )

    return sorted(set(violations))


def audit_iwa_common_table_merge_source_topology(root: Path = ROOT) -> list[str]:
    """Keep merged-cell geometry neutral and singular.

    The common table module owns checked coordinate geometry and its pure
    topology transforms. Numbers keeps only an import-path compatibility
    re-export while Pages and Keynote adapters consume the same owner. This
    audit is dormant until the owner or a focused merge adapter exists, so a
    staged source deletion does not become a false failure during monolith
    retirement.
    """

    common_path = root / IWA_COMMON_TABLE_MERGE_SOURCE
    module_path = root / NEUTRAL_TABLE_MODULE_SOURCE
    numbers_path = root / Path("crates/litchi-numbers/src/table/merge.rs")
    focused_paths = (
        root / Path("crates/litchi-pages/src/package/body_table_merges.rs"),
        root / Path("crates/litchi-keynote/src/package/slide_table_merges.rs"),
    )
    if not common_path.is_file() and not any(
        path.is_file() for path in (numbers_path, *focused_paths)
    ):
        return []

    violations: list[str] = []
    if not common_path.is_file():
        violations.append(
            "common merged-cell geometry owner is missing: "
            f"{IWA_COMMON_TABLE_MERGE_SOURCE}"
        )
        return violations

    common_raw = common_path.read_text(encoding="utf-8")
    common_code = _mask_rust_non_code(_mask_rust_cfg_test_items(common_raw))
    if not module_path.is_file():
        violations.append(
            "common table module is missing its merge export: "
            f"{NEUTRAL_TABLE_MODULE_SOURCE}"
        )
    else:
        module_code = _mask_rust_non_code(module_path.read_text(encoding="utf-8"))
        if IWA_COMMON_TABLE_MERGE_MODULE.search(module_code) is None:
            violations.append(
                "common table module is missing public merge module: "
                f"{NEUTRAL_TABLE_MODULE_SOURCE}"
            )

    for name in sorted(IWA_COMMON_TABLE_MERGE_REQUIRED_TYPES):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(common_code) is None:
            violations.append(
                "common merged-cell geometry owner is missing public "
                f"{name}: {IWA_COMMON_TABLE_MERGE_SOURCE}"
            )
    for name in sorted(IWA_COMMON_TABLE_MERGE_REQUIRED_FUNCTIONS):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:const[ \t]+)?fn[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(common_code) is None:
            violations.append(
                "common merged-cell geometry owner is missing public "
                f"{name}: {IWA_COMMON_TABLE_MERGE_SOURCE}"
            )

    import_statement = re.compile(
        r"(?ms)^[ \t]*(?:pub[ \t]+)?(?:use|extern[ \t]+crate)\b.*?;"
    )
    for statement in import_statement.finditer(common_code):
        match = IWA_COMMON_TABLE_MERGE_FORBIDDEN_IMPORT.search(statement.group(0))
        if match is None:
            continue
        line_number = common_code.count("\n", 0, statement.start() + match.start()) + 1
        violations.append(
            "common merged-cell geometry imports a format/archive/wire owner "
            f"{match.group(0)}: {IWA_COMMON_TABLE_MERGE_SOURCE}:{line_number}"
        )
    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(common_raw)
    ):
        if RUST_BYTE_SLICE.search(declaration) is None:
            continue
        violations.append(
            "common merged-cell geometry exposes raw bytes: "
            f"{IWA_COMMON_TABLE_MERGE_SOURCE}:{line_number}"
        )

    if numbers_path.is_file():
        numbers_raw = numbers_path.read_text(encoding="utf-8")
        numbers_code = _mask_rust_non_code(_mask_rust_cfg_test_items(numbers_raw))
        if IWA_COMMON_TABLE_MERGE_REEXPORT.search(numbers_code) is None:
            violations.append(
                "Numbers merged-cell geometry compatibility module must re-export "
                f"the common owner: {numbers_path.relative_to(root)}"
            )
        for pattern, label in (
            (IWA_COMMON_TABLE_MERGE_LOCAL_DECLARATION, "type"),
            (IWA_COMMON_TABLE_MERGE_LOCAL_FUNCTION, "function"),
        ):
            for match in pattern.finditer(numbers_code):
                line_number = numbers_code.count("\n", 0, match.start()) + 1
                violations.append(
                    "Numbers merged-cell geometry compatibility module redeclares "
                    f"common {label}: {numbers_path.relative_to(root)}:{line_number}"
                )

    return sorted(set(violations))


def _rust_public_method_bodies(
    source: str, method_name: str, *, type_name: str = "Package"
) -> list[tuple[str, str, int]]:
    """Return declarations and bodies for one public inherent method."""

    masked_source = _mask_rust_cfg_test_items(source)
    code = _mask_rust_non_code(masked_source)
    methods = _rust_public_methods_in_impl(masked_source, type_name)
    if not methods:
        return []

    method_pattern = re.compile(
        r"(?m)^[ \t]*pub(?![ \t\r\n]*\()[ \t\r\n]+"
        r"(?:(?:unsafe|async|const)[ \t\r\n]+)*fn[ \t\r\n]+"
        rf"(?:r#)?{re.escape(method_name)}\b"
    )
    expected_lines = {line for name, _declaration, line in methods if name == method_name}
    evidence: list[tuple[str, str, int]] = []
    for match in method_pattern.finditer(code):
        line_number = source.count("\n", 0, match.start()) + 1
        if line_number not in expected_lines:
            continue
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
            elif not parentheses and not brackets and character in "{;":
                break
            cursor += 1
        if cursor >= len(code) or code[cursor] != "{":
            continue
        end = _rust_balanced_delimited_end(code, cursor)
        if end is None:
            continue
        evidence.append((code[match.start() : cursor], code[match.start() : end], line_number))
    return evidence


def _rust_method_bodies(
    source: str, method_name: str, *, type_name: str | None = None
) -> list[tuple[str, str, int]]:
    """Return declarations and bodies for public or private inherent methods.

    The table-merge boundary has a deliberate lazy-entry path: its private
    resolver is called by each staging/publication method rather than by the
    package constructor.  This narrow sibling of
    :func:`_rust_public_method_bodies` lets the audit inspect that resolver
    without treating a marker in an unrelated helper as admission evidence.
    """

    masked_source = _mask_rust_cfg_test_items(source)
    code = _mask_rust_non_code(masked_source)
    impl_ranges: list[tuple[int, int]] = []
    if type_name is not None:
        for implementation in RUST_IMPL_DECLARATION.finditer(code):
            if code.count("{", 0, implementation.start()) != code.count(
                "}", 0, implementation.start()
            ):
                continue
            parentheses = 0
            brackets = 0
            opening = -1
            cursor = implementation.end()
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
                    opening = cursor
                    break
                cursor += 1
            if opening < 0:
                continue
            header = re.sub(r"\s+", " ", code[implementation.start() : opening]).strip()
            if re.match(
                rf"^impl(?:\s*<[^>{{}}]*>)?\s+{re.escape(type_name)}"
                rf"(?:\s*<[^>{{}}]*>)?(?:\s+where .*)?$",
                header,
            ) is None:
                continue
            end = _rust_balanced_delimited_end(code, opening)
            if end is not None:
                impl_ranges.append((opening + 1, end - 1))

    method_pattern = re.compile(
        r"(?m)^[ \t]*(?:(?:pub(?:\([^()\r\n]*\))?)[ \t]+)?"
        r"(?:(?:unsafe|async|const)[ \t]+)*fn[ \t\r\n]+"
        rf"(?:r#)?{re.escape(method_name)}\b"
    )
    evidence: list[tuple[str, str, int]] = []
    for match in method_pattern.finditer(code):
        if type_name is not None and not any(
            start <= match.start() < end for start, end in impl_ranges
        ):
            continue
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
            elif not parentheses and not brackets and character in "{;":
                break
            cursor += 1
        if cursor >= len(code) or code[cursor] != "{":
            continue
        end = _rust_balanced_delimited_end(code, cursor)
        if end is None:
            continue
        evidence.append(
            (code[match.start() : cursor], code[match.start() : end],
             code.count("\n", 0, match.start()) + 1)
        )
    return evidence


def _table_merge_deferred_admission_is_checked(
    transaction_source: str, spec: dict
) -> bool:
    """Check a lazy edit's private resolver and every guarded entry point."""

    deferred = spec.get("deferred_admission")
    if deferred is None:
        return False
    helper = deferred["helper"]
    edit_type = spec.get("type_names", {}).get(
        "Edit", f"{spec.get('prefix', '')}Edit"
    )
    helper_methods = _rust_method_bodies(
        transaction_source, helper, type_name=edit_type
    )
    if not helper_methods:
        return False
    helper_has_selector = any(
        _table_merge_transaction_has_marker(
            _mask_rust_non_code(body), spec["admission_markers"]
        )
        for _declaration, body, _line in helper_methods
    )
    if not helper_has_selector:
        return False
    call = re.compile(
        rf"\bself\s*\.\s*{re.escape(helper)}\s*\(\s*\)\s*\?\s*;"
    )
    for method_name in deferred["stage_methods"]:
        methods = _rust_method_bodies(
            transaction_source, method_name, type_name=edit_type
        )
        checked = False
        for _declaration, body, _line in methods:
            body_code = _mask_rust_non_code(body)
            opening = body_code.find("{")
            if opening < 0:
                continue
            body_code = body_code[opening + 1 :]
            match = call.search(body_code)
            if match is None:
                continue
            # The resolver must propagate its typed error, and it must be the
            # first self operation in the entry point.  This prevents a
            # discarded-error decoy or staging before target admission.
            first_self = re.search(r"\bself\s*\.", body_code)
            if first_self is None or match.start() != first_self.start():
                continue
            checked = True
            break
        if not checked:
            return False
    return True


def _audit_focused_table_merge_source_topology(
    root: Path,
    *,
    format_name: str,
    owner_path: Path,
    package_path: Path,
    module_pattern: re.Pattern[str],
    table_module_path: Path,
    table_merge_path: Path,
    table_module_pattern: re.Pattern[str],
    public_method: str,
    selector_types: frozenset[str],
    forbidden_public_types: re.Pattern[str] | None = None,
) -> list[str]:
    """Protect a selector-first focused merged-cell reader."""

    owner_absolute = root / owner_path
    package_absolute = root / package_path
    package_source = (
        package_absolute.read_text(encoding="utf-8")
        if package_absolute.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(_mask_rust_cfg_test_items(package_source))
    module_match = module_pattern.search(package_code)

    implementation_sources: list[tuple[Path, str]] = []
    if package_absolute.is_file():
        implementation_sources.append((package_path, package_source))
    package_directory = package_absolute.parent / package_absolute.stem
    if package_directory.is_dir():
        implementation_sources.extend(
            (path.relative_to(root), path.read_text(encoding="utf-8"))
            for path in sorted(package_directory.glob("*.rs"))
        )
    method_evidence = [
        (path, declaration, body, line)
        for path, source in implementation_sources
        for declaration, body, line in _rust_public_method_bodies(source, public_method)
    ]

    # Keep the audit dormant while an entire focused API is absent. A package
    # module declaration is itself an intentional source-topology claim, so a
    # half-removed owner still reports a useful failure.
    if not owner_absolute.is_file() and module_match is None and not method_evidence:
        return []

    violations: list[str] = []
    if not owner_absolute.is_file():
        violations.append(
            f"{format_name} table-merge focused owner is missing: {owner_path}"
        )
    elif module_match is None:
        violations.append(
            f"{format_name} table-merge package is missing its focused module: "
            f"{package_path}"
        )
    elif re.search(
        rf"(?m)^[ \t]*pub[ \t]+mod[ \t]+(?:r#)?"
        rf"{re.escape(public_method.removesuffix('_merges'))}_merges\b",
        package_code,
    ):
        # The exact module patterns below are intentionally private or
        # ``pub(crate)``. A public module would expose the physical adapter.
        violations.append(
            f"{format_name} table-merge focused module must remain private: "
            f"{package_path}"
        )

    focused_source = (
        owner_absolute.read_text(encoding="utf-8")
        if owner_absolute.is_file()
        else package_source
    )
    focused_raw = focused_source
    focused_code = _mask_rust_non_code(_mask_rust_cfg_test_items(focused_source))
    relative_owner = owner_path if owner_absolute.is_file() else package_path

    if IWA_FOCUSED_TABLE_MERGES_WIRE_IMPORT.search(focused_code) is None:
        violations.append(
            f"{format_name} table-merge focused owner must consume numbers-wire "
            f"table_merges: {relative_owner}"
        )
    if IWA_FOCUSED_TABLE_MERGES_COMMON_IMPORT.search(focused_code) is None:
        violations.append(
            f"{format_name} table-merge focused owner must import common merge "
            f"geometry: {relative_owner}"
        )
    if PAGES_TABLE_MERGE_FOCUSED_CALL.search(focused_code) is None:
        violations.append(
            f"{format_name} table-merge focused owner must call the shared reader: "
            f"{relative_owner}"
        )

    for pattern, label in (
        (PAGES_TABLE_MERGE_FORBIDDEN_MODULES, "format/archive or generated AST import"),
        (IWA_FOCUSED_TABLE_MERGES_GENERATED_AST, "generated AST projection"),
    ):
        for match in pattern.finditer(focused_code):
            line_number = focused_code.count("\n", 0, match.start()) + 1
            violations.append(
                f"{format_name} table-merge focused owner contains {label} "
                f"{match.group(0)}: {relative_owner}:{line_number}"
            )

    focused_production = _mask_rust_cfg_test_items(focused_raw)
    for declaration, line_number in _rust_public_declarations(focused_production):
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                f"{format_name} table-merge focused owner exposes raw bytes: "
                f"{relative_owner}:{line_number}"
            )
        raw_id = IWA_FOCUSED_TABLE_MERGES_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            violations.append(
                f"{format_name} table-merge focused owner exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: {relative_owner}:{line_number}"
            )
        if forbidden_public_types is not None:
            for match in forbidden_public_types.finditer(declaration):
                type_line = line_number + declaration.count("\n", 0, match.start())
                violations.append(
                    f"{format_name} table-merge focused owner exposes a private "
                    f"wire/BNC type {match.group(0)}: "
                    f"{relative_owner}:{type_line}"
                )

    duplicate_type = re.compile(
        r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?"
        r"(?:struct|enum|union|type)[ \t]+(?:r#)?"
        r"(?:Region|ReadLimits|ReadReport|MergeRead|MergeReadError|AttemptedCost)\b"
    )
    duplicate_function = re.compile(
        r"(?m)^[ \t]*(?:pub(?:\([^()]*\))?[ \t]+)?fn[ \t]+"
        r"(?:r#)?read_table_merges\b"
    )
    for pattern, label in (
        (duplicate_type, "semantic type"),
        (duplicate_function, "wire reader"),
    ):
        for match in pattern.finditer(focused_code):
            line_number = focused_code.count("\n", 0, match.start()) + 1
            violations.append(
                f"{format_name} table-merge focused owner redeclares a shared {label}: "
                f"{relative_owner}:{line_number}"
            )

    if owner_absolute.is_file():
        table_module_absolute = root / table_module_path
        table_merge_absolute = root / table_merge_path
        if not table_module_absolute.is_file():
            violations.append(
                f"{format_name} table semantic module is missing its merge export: "
                f"{table_module_path}"
            )
        else:
            table_module_source = table_module_absolute.read_text(encoding="utf-8")
            table_module_code = _mask_rust_non_code(table_module_source)
            if table_module_pattern.search(table_module_code) is None:
                violations.append(
                    f"{format_name} table semantic module is missing public merge "
                    f"module: {table_module_path}"
                )
        if not table_merge_absolute.is_file():
            violations.append(
                f"{format_name} table semantic merge wrapper is missing: "
                f"{table_merge_path}"
            )
        else:
            table_merge_source = table_merge_absolute.read_text(encoding="utf-8")
            table_merge_code = _mask_rust_non_code(table_merge_source)
            wrapper_code = table_merge_code
            if table_module_absolute.is_file():
                wrapper_code += "\n" + _mask_rust_non_code(
                    table_module_absolute.read_text(encoding="utf-8")
                )
            if IWA_FOCUSED_TABLE_MERGES_COMMON_IMPORT.search(wrapper_code) is None:
                violations.append(
                    f"{format_name} table semantic merge wrapper must re-export "
                    f"common geometry: {table_merge_path}"
                )
            for pattern, label in (
                (PAGES_TABLE_MERGE_FORBIDDEN_MODULES, "forbidden physical import"),
                (IWA_FOCUSED_TABLE_MERGES_GENERATED_AST, "generated AST projection"),
            ):
                for match in pattern.finditer(table_merge_code):
                    line_number = table_merge_code.count("\n", 0, match.start()) + 1
                    violations.append(
                        f"{format_name} table semantic merge wrapper contains "
                        f"{label} {match.group(0)}: {table_merge_path}:{line_number}"
                    )

    if not method_evidence:
        violations.append(
            f"{format_name} table-merge focused owner is missing public "
            f"Package::{public_method}: {owner_path}"
        )
    else:
        for path, declaration, body, line_number in method_evidence:
            if re.search(r"\bRegion\b", declaration) is None:
                violations.append(
                    f"{format_name} Package::{public_method} must return common Region "
                    f"values: {path}:{line_number}"
                )
            missing_selectors = sorted(
                selector
                for selector in selector_types
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None
            )
            for selector in missing_selectors:
                violations.append(
                    f"{format_name} Package::{public_method} must resolve typed "
                    f"{selector}: {path}:{line_number}"
                )
            if RUST_BYTE_SLICE.search(declaration) is not None:
                violations.append(
                    f"{format_name} Package::{public_method} exposes raw bytes: "
                    f"{path}:{line_number}"
                )
            raw_id = IWA_FOCUSED_TABLE_MERGES_RAW_ID_PARAMETER.search(declaration)
            if raw_id is not None:
                violations.append(
                    f"{format_name} Package::{public_method} exposes a raw ID parameter "
                    f"{raw_id.group(0).strip()}: {path}:{line_number}"
                )
            if forbidden_public_types is not None:
                for match in forbidden_public_types.finditer(declaration):
                    type_line = line_number + declaration.count(
                        "\n", 0, match.start()
                    )
                    violations.append(
                        f"{format_name} Package::{public_method} exposes a private "
                        f"wire/BNC type {match.group(0)}: {path}:{type_line}"
                    )
            generated = IWA_FOCUSED_TABLE_MERGES_GENERATED_AST.search(body)
            if generated is not None:
                violations.append(
                    f"{format_name} Package::{public_method} contains a generated AST "
                    f"projection {generated.group(0)}: {path}:{line_number}"
                )

    return sorted(set(violations))


def _audit_focused_table_merge_reader_source_topology(
    root: Path,
    *,
    format_name: str,
    owner_path: Path,
    reader_path: Path,
    query_method: str,
    selector_types: frozenset[str],
    limits_type: str,
    options_type: str | None = None,
) -> list[str]:
    """Keep one focused merge reader source-facing and selector-first.

    The package-level geometry ratchet remains responsible for ``Package``;
    this helper owns only the private reader child's public constructors and
    query.
    """

    owner_absolute = root / owner_path
    reader_absolute = root / reader_path
    owner_source = (
        owner_absolute.read_text(encoding="utf-8")
        if owner_absolute.is_file()
        else ""
    )
    owner_reader_modules = (
        _rust_root_level_module_declarations(owner_source, frozenset({"reader"}))
        if owner_source
        else ()
    )
    if not reader_absolute.is_file():
        if owner_reader_modules or (
            owner_source
            and _rust_named_struct_body(owner_source, "MergeReader") is not None
        ):
            return [
                f"{format_name} table-merge reader child is missing: {reader_path}"
            ]
        return []

    violations: list[str] = []
    if reader_absolute.is_file():
        if not owner_absolute.is_file():
            violations.append(
                f"{format_name} table-merge reader child has no package owner: "
                f"{owner_path}"
            )
        else:
            module_declarations = owner_reader_modules
            if not module_declarations:
                violations.append(
                    f"{format_name} table-merge reader child is not wired from its "
                    f"private owner: {owner_path}"
                )
            elif any(
                visibility is not None
                for _name, visibility, *_rest in module_declarations
            ):
                violations.append(
                    f"{format_name} table-merge reader child module must remain "
                    f"private: {owner_path}"
                )

        reader_source = reader_absolute.read_text(encoding="utf-8")
        relative_reader = reader_path

    production_source = _mask_rust_cfg_test_items(reader_source)
    code = _mask_rust_non_code(production_source)

    struct = _rust_named_struct_body(production_source, "MergeReader")
    if struct is None:
        violations.append(
            f"{format_name} table-merge reader is missing its MergeReader state: "
            f"{relative_reader}"
        )
    else:
        struct_body, body_offset = struct
        if re.search(
            r"(?m)^[ \t]*pub[ \t]+struct[ \t]+(?:r#)?MergeReader\b",
            code,
        ) is None:
            violations.append(
                f"{format_name} table-merge reader MergeReader must be public: "
                f"{relative_reader}"
            )
        for match in NUMBERS_TABLE_MERGE_READER_PUBLIC_FIELD.finditer(struct_body):
            line_number = code.count("\n", 0, body_offset + match.start()) + 1
            violations.append(
                f"{format_name} table-merge reader exposes a public MergeReader "
                f"field: {relative_reader}:{line_number}"
            )

    # Generated/owned protobuf values, the complete semantic document model,
    # and byte-returning escape hatches must stay below this boundary.  Byte
    # parameters are checked separately and are allowed only for constructors.
    for pattern, label in (
        (
            FOCUSED_TABLE_MERGE_READER_FORBIDDEN_TYPES,
            "generated/owned protobuf, native, or semantic model type",
        ),
        (
            FOCUSED_TABLE_MERGE_READER_MODEL_REFERENCE,
            "full document model reference",
        ),
    ):
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                f"{format_name} table-merge reader contains {label} "
                f"{match.group(0)}: {relative_reader}:{line_number}"
            )

    methods = _rust_public_methods_in_impl(production_source, "MergeReader")
    method_lines = {line for _name, _declaration, line in methods}
    method_names = {name for name, _declaration, _line in methods}
    if "open" not in method_names:
        violations.append(
            f"{format_name} table-merge reader is missing path ingress: "
            f"{relative_reader}"
        )
    if "from_bytes" not in method_names:
        violations.append(
            f"{format_name} table-merge reader is missing byte ingress: "
            f"{relative_reader}"
        )
    if query_method not in method_names:
        violations.append(
            f"{format_name} table-merge reader is missing its typed query: "
            f"{relative_reader}"
        )

    # Every public item in this child belongs to the reader.
    for declaration, line_number in _rust_public_declarations(production_source):
        if re.search(r"\bstruct[ \t]+(?:r#)?MergeReader\b", declaration):
            continue
        function = NUMBERS_TABLE_MERGE_READER_FUNCTION_NAME.search(declaration)
        if function is None:
            violations.append(
                f"{format_name} table-merge reader exposes an unsupported public "
                f"declaration: {relative_reader}:{line_number}"
            )
            continue
        if line_number not in method_lines:
            violations.append(
                f"{format_name} table-merge reader public method must be an "
                f"associated MergeReader method ({function.group('name')}): "
                f"{relative_reader}:{line_number}"
            )

    for name, declaration, line_number in methods:
        arrow = declaration.find("->")
        return_type = declaration[arrow + 2 :] if arrow >= 0 else ""
        if (
            RUST_BYTE_SLICE.search(return_type) is not None
            or NUMBERS_TABLE_MERGE_READER_BYTE_RETURN.search(return_type) is not None
            or FOCUSED_TABLE_MERGE_READER_ARC_BYTES.search(return_type) is not None
        ):
            violations.append(
                f"{format_name} table-merge reader exposes raw bytes from public "
                f"method {name}: {relative_reader}:{line_number}"
            )

        raw_id = IWA_FOCUSED_TABLE_MERGES_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            violations.append(
                f"{format_name} table-merge reader exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: {relative_reader}:{line_number}"
            )

        physical_type = FOCUSED_TABLE_MERGE_READER_PUBLIC_PHYSICAL_TYPES.search(
            declaration
        )
        if physical_type is not None:
            violations.append(
                f"{format_name} table-merge reader exposes a physical or "
                f"semantic type {physical_type.group(0)} from public method "
                f"{name}: {relative_reader}:{line_number}"
            )

        if name not in FOCUSED_TABLE_MERGE_READER_ALLOWED_METHODS:
            violations.append(
                f"{format_name} table-merge reader exposes unsupported public "
                f"method {name}: {relative_reader}:{line_number}"
            )
            continue

        if name.endswith("_with_limits") and re.search(
            rf"\b{re.escape(limits_type)}\b", declaration
        ) is None:
            violations.append(
                f"{format_name} table-merge reader limits ingress must retain "
                f"{limits_type} ({name}): {relative_reader}:{line_number}"
            )
        if name.endswith("_with_options") and options_type is not None and re.search(
            rf"\b{re.escape(options_type)}\b", declaration
        ) is None:
            violations.append(
                f"{format_name} table-merge reader options ingress must retain "
                f"{options_type} ({name}): {relative_reader}:{line_number}"
            )

        if name in {
            "open",
            "open_with_limits",
            "open_with_options",
        }:
            if FOCUSED_TABLE_MERGE_READER_PATH_INPUT.search(declaration) is None:
                violations.append(
                    f"{format_name} table-merge reader path ingress must accept "
                    f"AsRef<Path> ({name}): {relative_reader}:{line_number}"
                )
        elif name in {"from_bytes", "from_bytes_with_limits", "from_bytes_with_options"}:
            if FOCUSED_TABLE_MERGE_READER_BYTE_INPUT.search(declaration) is None:
                violations.append(
                    f"{format_name} table-merge reader byte ingress must accept "
                    f"&[u8] ({name}): {relative_reader}:{line_number}"
                )
        elif name in {
            "from_shared_bytes",
            "from_shared_bytes_with_limits",
            "from_shared_bytes_with_options",
        }:
            if FOCUSED_TABLE_MERGE_READER_ARC_BYTES.search(declaration) is None:
                violations.append(
                    f"{format_name} table-merge reader shared ingress must accept "
                    f"Arc<[u8]> ({name}): {relative_reader}:{line_number}"
                )
        elif name == query_method:
            missing_selectors = sorted(
                selector
                for selector in selector_types
                if re.search(rf"\b{re.escape(selector)}\b", declaration) is None
            )
            for selector in missing_selectors:
                violations.append(
                    f"{format_name} table-merge reader query must use typed "
                    f"{selector}: {relative_reader}:{line_number}"
                )
            if re.search(r"\bRegion\b", declaration) is None:
                violations.append(
                    f"{format_name} table-merge reader query must return semantic "
                f"Region values: {relative_reader}:{line_number}"
                )

        if name not in {query_method, "read_options", "options"} and re.search(
            r"\bResult[ \t]*<[ \t]*Self\b", return_type
        ) is None:
            violations.append(
                f"{format_name} table-merge reader ingress method must return "
                f"Result<Self, _> ({name}): {relative_reader}:{line_number}"
            )

    return sorted(set(violations))


def _table_merge_transaction_sources(
    root: Path, spec: dict
) -> tuple[list[tuple[Path, str]], list[str]]:
    """Load one focused merge transaction owner and its private child."""

    owner_path = spec["owner_path"]
    transaction_path = spec["transaction_path"]
    owner_absolute = root / owner_path
    transaction_absolute = root / transaction_path
    violations: list[str] = []
    sources: list[tuple[Path, str]] = []

    if not owner_absolute.is_file():
        violations.append(
            f"focused {spec['format_name']} table-merge owner is missing: "
            f"{owner_path}"
        )
        return sources, violations

    owner_source = _mask_rust_cfg_test_items(
        owner_absolute.read_text(encoding="utf-8")
    )
    sources.append((owner_path, owner_source))
    declarations = _rust_root_level_module_declarations(
        owner_source, frozenset({spec["transaction_module"]})
    )
    if not declarations:
        violations.append(
            f"focused {spec['format_name']} table-merge transaction child is not "
            f"wired from its private owner: {owner_path}"
        )
    else:
        for _name, visibility, shape, _body, line_number in declarations:
            if visibility is not None and visibility.strip() == "pub":
                violations.append(
                    f"focused {spec['format_name']} table-merge transaction child "
                    f"module must remain private: {owner_path}:{line_number}"
                )
            if shape == "inline":
                violations.append(
                    f"focused {spec['format_name']} table-merge transaction must "
                    f"use a dedicated child source: {owner_path}:{line_number}"
                )

    if not transaction_absolute.is_file():
        violations.append(
            f"focused {spec['format_name']} table-merge transaction child is "
            f"missing: {transaction_path}"
        )
    else:
        transaction_source = _mask_rust_cfg_test_items(
            transaction_absolute.read_text(encoding="utf-8")
        )
        sources.append((transaction_path, transaction_source))

    return sources, violations


def _table_merge_transaction_has_marker(
    code: str, markers: tuple[str, ...] | frozenset[str]
) -> bool:
    """Return whether one semantic marker or any allowed alternative exists."""

    return any(
        re.search(rf"(?<![A-Za-z0-9_]){re.escape(marker)}(?![A-Za-z0-9_])", code)
        is not None
        for marker in markers
    )


def _audit_table_merge_transaction_source_topology(
    root: Path, spec: dict
) -> list[str]:
    """Audit a focused merge transaction without inspecting private wire fields."""

    sources, violations = _table_merge_transaction_sources(root, spec)
    if not sources:
        return sorted(set(violations))

    format_name = spec["format_name"]
    prefix = spec["prefix"]
    # ADR 0034 Pages owners define the canonical short names (``Edit``,
    # ``Patch``, ...); other formats still use ``<prefix><suffix>``.
    type_names = {
        suffix: spec.get("type_names", {}).get(suffix, f"{prefix}{suffix}")
        for suffix in IWA_TABLE_MERGE_TRANSACTION_CANONICAL_SUFFIXES
    }
    display_prefix = spec.get("type_display_prefix", "")
    type_labels = {
        suffix: f"{display_prefix}{name}" for suffix, name in type_names.items()
    }
    canonical_names = frozenset(type_names.values())
    family_owner = spec.get("transaction_family_owner")
    owner_path = spec["owner_path"]
    owner_source = sources[0][1]
    owner_code = _mask_rust_non_code(owner_source)

    # Keep the transaction child private at the package seam. ``pub(crate)``
    # remains an internal implementation visibility; an unrestricted ``pub``
    # module would make the physical child a supported API by accident.
    package_path = spec["package_path"]
    package_absolute = root / package_path
    package_source = (
        _mask_rust_cfg_test_items(package_absolute.read_text(encoding="utf-8"))
        if package_absolute.is_file()
        else ""
    )
    if not package_source:
        violations.append(
            f"focused {format_name} table-merge package export source is missing: "
            f"{package_path}"
        )
    else:
        package_modules = _rust_root_level_module_declarations(
            package_source, frozenset({spec["module_name"]})
        )
        if not package_modules:
            violations.append(
                f"focused {format_name} table-merge package module is missing: "
                f"{package_path}"
            )
        elif any(
            visibility is not None and visibility.strip() == "pub"
            for _name, visibility, _shape, _body, _line in package_modules
        ):
            violations.append(
                f"focused {format_name} table-merge package module must remain "
                f"private: {package_path}"
            )

    library_path = spec["library_path"]
    library_absolute = root / library_path
    library_source = (
        _mask_rust_cfg_test_items(library_absolute.read_text(encoding="utf-8"))
        if library_absolute.is_file()
        else ""
    )
    if not library_source:
        violations.append(
            f"focused {format_name} table-merge public facade source is missing: "
            f"{library_path}"
        )

    transaction_source = "\n".join(source for _path, source in sources)
    transaction_code = _mask_rust_non_code(transaction_source)

    # The owner must publish the four semantic lifecycle types through its
    # private child. Package and crate-root facades repeat that exact set,
    # unless the family is published through one focused public module.
    focused_module = spec.get("focused_module")
    if focused_module is not None:
        focused_absolute = root / focused_module
        focused_source = (
            _mask_rust_cfg_test_items(focused_absolute.read_text(encoding="utf-8"))
            if focused_absolute.is_file()
            else ""
        )
        focused_exports = _rust_canonical_exports(focused_source, canonical_names)
        for name in sorted(canonical_names - focused_exports):
            violations.append(
                f"focused {format_name} table-merge focused module is missing "
                f"canonical export {name}: {focused_module}"
            )
        package_source = ""
        library_source = ""
    owner_exports = (
        canonical_names
        if focused_module is not None
        else _rust_canonical_exports(owner_source, canonical_names)
    )
    for name in sorted(canonical_names - owner_exports):
        violations.append(
            f"focused {format_name} table-merge owner is missing canonical "
            f"transaction type {name}: {owner_path}"
        )
    if sources:
        child_path, child_source = sources[-1]
        child_exports = _rust_canonical_exports(child_source, canonical_names)
        for name in sorted(canonical_names - child_exports):
            violations.append(
                f"focused {format_name} table-merge transaction child is missing "
                f"canonical type {name}: {child_path}"
            )
    facade_sources = [(package_path, package_source), (library_path, library_source)]
    if family_owner is not None:
        # Only the family's public module re-exports the short names;
        # the package and crate-root facades must not.  The shared helper
        # checks the owner visibility, the exact re-export and the deleted
        # prefixed spellings.
        violations.extend(
            _pages_transaction_family_violations(
                root, family_owner, f"focused {format_name} table-merge"
            )
        )
        public_file, public_scope, _ = _pages_public_module_scope(
            root, PAGES_TRANSACTION_FAMILY_BY_OWNER[family_owner][2]
        )
        facade_sources = (
            []
            if public_file is None or public_scope is None
            else [(public_file.relative_to(root), public_scope)]
        )
    elif package_source:
        # A format whose focused family module is the public home (ADR 0034)
        # checks that module instead of the package
        # facade.
        export_path = spec.get("package_export_path", package_path)
        export_absolute = root / export_path
        export_source = (
            _mask_rust_cfg_test_items(export_absolute.read_text(encoding="utf-8"))
            if export_path != package_path and export_absolute.is_file()
            else package_source
        )
        package_exports = _rust_canonical_exports(export_source, canonical_names)
        for name in sorted(canonical_names - package_exports):
            violations.append(
                f"focused {format_name} table-merge package facade is missing "
                f"canonical export {name}: {export_path}"
            )
    if library_source and family_owner is None:
        library_exports = _rust_canonical_exports(library_source, canonical_names)
        for name in sorted(canonical_names - library_exports):
            violations.append(
                f"focused {format_name} table-merge crate facade is missing "
                f"canonical export {name}: {library_path}"
            )

    # A public field would expose proof state or an owned native representation
    # and make the semantic lifecycle impossible to evolve safely.
    for path, source in sources:
        production = _mask_rust_cfg_test_items(source)
        for name in sorted(canonical_names):
            struct = _rust_named_struct_body(production, name)
            if struct is None:
                continue
            body, offset = struct
            for match in IWA_TABLE_MERGE_TRANSACTION_PUBLIC_FIELD.finditer(body):
                line_number = production.count(
                    "\n", 0, offset + match.start()
                ) + 1
                violations.append(
                    f"focused {format_name} table-merge transaction exposes a "
                    f"public field: {path}:{line_number}"
                )

    # Only public declarations cross this check. Private fields and helpers
    # may retain archive/source objects, while public signatures stay semantic.
    public_methods: list[tuple[Path, str, str, int]] = []
    for path, source in sources:
        production = _mask_rust_cfg_test_items(source)
        for type_name in (*canonical_names, "Package"):
            for method_name, declaration, line_number in _rust_public_methods_in_impl(
                production, type_name
            ):
                public_methods.append((path, method_name, declaration, line_number))

    for path, method_name, declaration, line_number in public_methods:
        raw_id = IWA_TABLE_MERGE_TRANSACTION_PUBLIC_RAW_ID.search(declaration)
        if raw_id is not None:
            violations.append(
                f"focused {format_name} table-merge transaction public method "
                f"{method_name} exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: {path}:{line_number}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                f"focused {format_name} table-merge transaction public method "
                f"{method_name} exposes raw bytes: {path}:{line_number}"
            )
        arrow = declaration.find("->")
        return_type = declaration[arrow + 2 :] if arrow >= 0 else ""
        if IWA_TABLE_MERGE_TRANSACTION_PUBLIC_BYTE_RETURN.search(return_type):
            violations.append(
                f"focused {format_name} table-merge transaction public method "
                f"{method_name} returns raw bytes: {path}:{line_number}"
            )
        physical = IWA_TABLE_MERGE_TRANSACTION_FORBIDDEN_PUBLIC_TYPES.search(
            declaration
        )
        if physical is not None:
            violations.append(
                f"focused {format_name} table-merge transaction public method "
                f"{method_name} exposes a private physical/model type "
                f"{physical.group()}: {path}:{line_number}"
            )

    # Facade re-exports are public declarations too. Restrict this scan to
    # declarations carrying one of the canonical names so unrelated package
    # APIs remain outside the merge ratchet.
    for path, source in facade_sources:
        if not source:
            continue
        for declaration, line_number in _rust_public_declarations(source):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            }
            if not identifiers & canonical_names:
                continue
            physical = IWA_TABLE_MERGE_TRANSACTION_FORBIDDEN_PUBLIC_TYPES.search(
                declaration
            )
            if physical is not None:
                violations.append(
                    f"focused {format_name} table-merge public facade exposes a "
                    f"private physical/model type {physical.group()}: "
                    f"{path}:{line_number}"
                )
            if RUST_BYTE_SLICE.search(declaration) is not None:
                violations.append(
                    f"focused {format_name} table-merge public facade exposes raw "
                    f"bytes: {path}:{line_number}"
                )
            raw_id = IWA_TABLE_MERGE_TRANSACTION_PUBLIC_RAW_ID.search(declaration)
            if raw_id is not None:
                violations.append(
                    f"focused {format_name} table-merge public facade exposes a raw "
                    f"ID parameter {raw_id.group(0).strip()}: {path}:{line_number}"
                )

    # Package operations are intentionally checked by their owning impl, so a
    # free helper with the same name cannot satisfy the public contract.
    package_methods = {
        method_name: (declaration, line_number, path)
        for path, method_name, declaration, line_number in public_methods
        if method_name in {spec["edit_method"], spec["apply_method"]}
    }
    edit_evidence: list[tuple[str, str, int, Path]] = []
    apply_evidence: list[tuple[str, str, int, Path]] = []
    for path, source in sources:
        for evidence_declaration, body, line_number in _rust_public_method_bodies(
            source, spec["edit_method"], type_name="Package"
        ):
            edit_evidence.append((evidence_declaration, body, line_number, path))
        for evidence_declaration, body, line_number in _rust_public_method_bodies(
            source, spec["apply_method"], type_name="Package"
        ):
            apply_evidence.append((evidence_declaration, body, line_number, path))
    deferred_admission_checked = _table_merge_deferred_admission_is_checked(
        transaction_source, spec
    )

    if spec["edit_method"] not in package_methods:
        violations.append(
            f"focused {format_name} table-merge Package method is missing "
            f"{spec['edit_method']}: {owner_path}"
        )
    else:
        declaration, _line_number, path = package_methods[spec["edit_method"]]
        if re.search(rf"\b{re.escape(type_names['Edit'])}\b", declaration) is None:
            violations.append(
                f"focused {format_name} {spec['edit_method']} must return the "
                f"typed {type_labels['Edit']}: {path}"
            )
        for selector in sorted(spec["selector_types"]):
            if re.search(rf"\b{re.escape(selector)}\b", declaration) is None:
                violations.append(
                    f"focused {format_name} {spec['edit_method']} must accept "
                    f"typed {selector}: {path}"
                )
        if not edit_evidence:
            violations.append(
                f"focused {format_name} {spec['edit_method']} has no public Package "
                f"body: {path}"
            )
        else:
            for _declaration, body, line_number, body_path in edit_evidence:
                body_code = _mask_rust_non_code(body)
                if not (
                    _table_merge_transaction_has_marker(
                        body_code, spec["admission_markers"]
                    )
                    or deferred_admission_checked
                ):
                    violations.append(
                        f"focused {format_name} {spec['edit_method']} must resolve "
                        f"its typed selector before staging: {body_path}:{line_number}"
                    )
                if IWA_TABLE_MERGE_TRANSACTION_LEGACY_MUTATION.search(body_code):
                    violations.append(
                        f"focused {format_name} {spec['edit_method']} retains a "
                        f"legacy/catchall mutation fallback: {body_path}:{line_number}"
                    )
                if re.search(r"\bErr\s*\(\s*_\s*\)\s*=>", body_code):
                    violations.append(
                        f"focused {format_name} {spec['edit_method']} uses a "
                        f"post-admission catchall error arm: {body_path}:{line_number}"
                    )

    if spec["apply_method"] not in package_methods:
        violations.append(
            f"focused {format_name} table-merge Package method is missing "
            f"{spec['apply_method']}: {owner_path}"
        )
    else:
        declaration, _line_number, path = package_methods[spec["apply_method"]]
        if re.search(rf"\b{re.escape(type_names['Patch'])}\b", declaration) is None:
            violations.append(
                f"focused {format_name} {spec['apply_method']} must accept the "
                f"typed {type_labels['Patch']}: {path}"
            )
        if re.search(rf"\b{re.escape(type_names['Commit'])}\b", declaration) is None:
            violations.append(
                f"focused {format_name} {spec['apply_method']} must return the "
                f"typed {type_labels['Commit']}: {path}"
            )
        if not apply_evidence:
            violations.append(
                f"focused {format_name} {spec['apply_method']} has no public Package "
                f"body: {path}"
            )
        else:
            for _declaration, body, line_number, body_path in apply_evidence:
                body_code = _mask_rust_non_code(body)
                if IWA_TABLE_MERGE_TRANSACTION_LEGACY_MUTATION.search(body_code):
                    violations.append(
                        f"focused {format_name} {spec['apply_method']} retains a "
                        f"legacy/catchall mutation fallback: {body_path}:{line_number}"
                    )
                if re.search(r"\bErr\s*\(\s*_\s*\)\s*=>", body_code):
                    violations.append(
                        f"focused {format_name} {spec['apply_method']} uses a "
                        f"post-admission catchall error arm: {body_path}:{line_number}"
                    )

    edit_methods = {
        method_name for method_name, _declaration, _line in sum(
            [
                _rust_public_methods_in_impl(
                    _mask_rust_cfg_test_items(source), type_names["Edit"]
                )
                for _path, source in sources
            ],
            [],
        )
    }
    edit_signatures = {
        method_name: declaration
        for _path, source in sources
        for method_name, declaration, _line in _rust_public_methods_in_impl(
            _mask_rust_cfg_test_items(source), type_names["Edit"]
        )
    }
    for required in ("merge", "unmerge", "commit"):
        if required not in edit_methods:
            violations.append(
                f"focused {format_name} {type_labels['Edit']} is missing public {required}: "
                f"{spec['transaction_path']}"
            )
        elif required in {"merge", "unmerge"} and re.search(
            r"\bRegion\b", edit_signatures[required]
        ) is None:
            violations.append(
                f"focused {format_name} {type_labels['Edit']}::{required} must use common "
                f"Region values: {spec['transaction_path']}"
            )
    patch_methods = {
        method_name for method_name, _declaration, _line in sum(
            [
                _rust_public_methods_in_impl(
                    _mask_rust_cfg_test_items(source), type_names["Patch"]
                )
                for _path, source in sources
            ],
            [],
        )
    }
    for required in ("before", "after", "is_noop", "inverse"):
        if required not in patch_methods:
            violations.append(
                f"focused {format_name} {type_labels['Patch']} is missing public {required}: "
                f"{spec['transaction_path']}"
            )
    commit_methods = {
        method_name for method_name, _declaration, _line in sum(
            [
                _rust_public_methods_in_impl(
                    _mask_rust_cfg_test_items(source), type_names["Commit"]
                )
                for _path, source in sources
            ],
            [],
        )
    }
    for required in ("package", "into_package", "patch", "diagnostics"):
        if required not in commit_methods:
            violations.append(
                f"focused {format_name} {type_labels['Commit']} is missing public {required}: "
                f"{spec['transaction_path']}"
            )
    diagnostics_methods = {
        method_name for method_name, _declaration, _line in sum(
            [
                _rust_public_methods_in_impl(
                    _mask_rust_cfg_test_items(source), type_names["Diagnostics"]
                )
                for _path, source in sources
            ],
            [],
        )
    }
    if "changed" not in diagnostics_methods:
        violations.append(
            f"focused {format_name} {type_labels['Diagnostics']} is missing public changed: "
            f"{spec['transaction_path']}"
        )

    marker_groups = {
        "source-preserving writer": (
            # The derived entry point takes the new owner UUID from the
            # table UUID instead of ambient randomness.
            ("rewrite_table_merges", "rewrite_table_merges_derived"),
            (
                "replace_message_preserving_header",
                "rewrite_model",
                "rewrite_merges",
            ),
        ),
        "bounded accounting": (
            ("Budget", "WireBudget", "transaction_budget", "ReadLimits", "WireLimits"),
            ("try_reserve", "try_reserve_exact"),
            ("checked_add", "checked_mul"),
        ),
        "candidate verification": (
            (
                "reopen_candidate",
                "reopen_patch",
                "parse_candidate",
                "from_source_owner_with_options",
                "from_source_catalog",
            ),
            ("read_regions", "read_regions_at", "read_table_merges", "select_merges"),
            ("verify_locality", "verify_exact_locality"),
        ),
        "exact-source patch semantics": (
            ("source_is_exact", "authorizes_source", "authorizes_owner"),
            ("is_noop", "unchanged", "changed: false"),
            ("inverse", "PatchConflict"),
        ),
        "strict semantic validation": (
            ("validate_region", "end_row", "end_column"),
            ("normalize_regions",),
            ("OverlappingRegion", "overlaps"),
        ),
        "bounded reassembly": (
            (
                "prepare_reassembly",
                "reassemble_to_bytes",
                "execution_requirements",
                "reassembly",
            ),
        ),
    }
    for label, groups in marker_groups.items():
        if not all(_table_merge_transaction_has_marker(transaction_code, group) for group in groups):
            violations.append(
                f"focused {format_name} table-merge transaction is missing "
                f"{label} marker: {spec['transaction_path']}"
            )

    return sorted(set(violations))


def audit_iwa_table_merge_formula_codec_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Guard the optional narrow Buffa formula encoder once it is wired."""

    codec_path = IWA_TABLE_MERGE_FORMULA_CODEC_SOURCE
    codec_absolute = root / codec_path
    facade_absolute = root / IWA_PROTOS_FACADE_SOURCE
    writer_absolute = root / IWA_NUMBERS_WIRE_TABLE_MERGE_WRITER_SOURCE
    facade_source = (
        _mask_rust_cfg_test_items(facade_absolute.read_text(encoding="utf-8"))
        if facade_absolute.is_file()
        else ""
    )
    writer_source = (
        _mask_rust_cfg_test_items(writer_absolute.read_text(encoding="utf-8"))
        if writer_absolute.is_file()
        else ""
    )
    wired = bool(
        codec_absolute.is_file()
        or IWA_TABLE_MERGE_FORMULA_CODEC_MODULE in _mask_rust_non_code(facade_source)
        or IWA_TABLE_MERGE_FORMULA_CODEC_MODULE in _mask_rust_non_code(writer_source)
    )
    if not wired:
        return []

    violations: list[str] = []
    if not codec_absolute.is_file():
        violations.append(
            f"table-merge formula codec source is missing: {codec_path}"
        )
        return violations

    build_absolute = root / IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_SOURCE
    if not build_absolute.is_file():
        violations.append(
            "table-merge formula codec build inventory source is missing: "
            f"{IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_SOURCE}"
        )
    else:
        build_source = build_absolute.read_text(encoding="utf-8")
        if IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_RERUN_MARKER not in build_source:
            violations.append(
                "table-merge formula codec build is missing its rerun inventory "
                f"marker: {IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_SOURCE}"
            )
        if IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_INVENTORY.search(build_source) is None:
            violations.append(
                "table-merge formula codec build is missing its private Buffa "
                "production-ingress inventory entry: "
                f"{IWA_TABLE_MERGE_FORMULA_CODEC_BUILD_SOURCE}"
            )

    codec_source = _mask_rust_cfg_test_items(
        codec_absolute.read_text(encoding="utf-8")
    )
    code = _mask_rust_non_code(codec_source)
    if not re.search(r"\bbuffa\b", code):
        violations.append(f"table-merge formula codec must use Buffa: {codec_path}")
    if re.search(r"\b(?:prost|prost_types)\b", code):
        violations.append(f"table-merge formula codec must not use Prost: {codec_path}")
    for type_name in IWA_TABLE_MERGE_FORMULA_CODEC_REQUIRED_TYPES:
        if re.search(
            rf"(?m)^\s*pub\s+(?:struct|enum|type)\s+"
            rf"{re.escape(type_name)}\b",
            code,
        ) is None:
            violations.append(
                f"table-merge formula codec is missing public {type_name}: "
                f"{codec_path}"
            )
    for function_name in IWA_TABLE_MERGE_FORMULA_CODEC_REQUIRED_FUNCTIONS:
        if re.search(
            rf"(?m)^\s*pub\s+fn\s+{re.escape(function_name)}\b",
            code,
        ) is None:
            violations.append(
                f"table-merge formula codec is missing narrow public encoder "
                f"{function_name}: {codec_path}"
            )
    for marker_group in (
        ("decode_lazy_view", "LazyMessageView"),
        ("try_encode_to_vec", "try_encoded_len", "ViewEncode", "encode_bounded"),
        ("preflight", "DecodeOptions", "EncodeOptions"),
    ):
        if not _table_merge_transaction_has_marker(code, marker_group):
            violations.append(
                f"table-merge formula codec is missing Buffa/lazy marker "
                f"{marker_group[0]}: {codec_path}"
            )
    if re.search(r"\b(?:FormulaArchive|AstNodeArchive|ASTNodeArrayArchive)\b", code):
        # The codec may mention schema names in private comments/tests, but
        # production source must not expose generated values as its API.
        for declaration, line_number in _rust_public_declarations(codec_source):
            if re.search(r"\b(?:FormulaArchive|AstNodeArchive|ASTNodeArrayArchive)\b", declaration):
                violations.append(
                    f"table-merge formula codec exposes a generated TSCE value: "
                    f"{codec_path}:{line_number}"
                )
    if IWA_TABLE_MERGE_FORMULA_CODEC_HIDDEN_MODULE.search(facade_source) is None:
        violations.append(
            f"table-merge formula codec module must be doc(hidden): "
            f"{IWA_PROTOS_FACADE_SOURCE}"
        )
    if re.search(r"(?m)^\s*pub\s+mod\s+(?:buffa_)?[A-Za-z0-9_]*generated\b", facade_source):
        violations.append(
            f"table-merge formula codec must keep generated Buffa modules private: "
            f"{IWA_PROTOS_FACADE_SOURCE}"
        )
    return sorted(set(violations))


def _boundary_marker_present(code: str, markers: Iterable[str]) -> bool:
    """Return whether one token or phrase marker occurs in masked Rust code."""

    for marker in markers:
        if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", marker):
            if re.search(
                rf"(?<![A-Za-z0-9_]){re.escape(marker)}(?![A-Za-z0-9_])",
                code,
            ):
                return True
        elif marker in code:
            return True
    return False


def _audit_iwa_component_removal_codec_source_topology(
    root: Path,
    *,
    parent_path: Path,
    child_path: Path,
    module_name: str,
    label: str,
    prepared_pattern: re.Pattern[str],
    prepare_pattern: re.Pattern[str],
    marker_groups: dict[str, tuple[str, ...]],
    build_marker: str,
    guard_marker: str,
) -> list[str]:
    """Audit one private source-preserving component/dependency codec child."""

    parent_absolute = root / parent_path
    child_absolute = root / child_path
    # A parent module declaration can land before its child during a staged
    # migration.  Only activate once the child source or directory is present;
    # an existing parent directory still fails closed on a missing child.
    if not child_absolute.is_file() and not child_absolute.parent.is_dir():
        return []
    violations: list[str] = []
    if not parent_absolute.is_file():
        return [f"{label} codec parent source is missing: {parent_path}"]
    if not child_absolute.is_file():
        return [f"{label} codec child source is missing: {child_path}"]

    parent_source = _mask_rust_cfg_test_items(parent_absolute.read_text(encoding="utf-8"))
    parent_code = _mask_rust_non_code(parent_source)
    child_source = _mask_rust_cfg_test_items(child_absolute.read_text(encoding="utf-8"))
    child_code = _mask_rust_non_code(child_source)

    modules = _rust_root_level_module_declarations(parent_source, frozenset({module_name}))
    if not modules:
        violations.append(
            f"{label} codec parent is missing private child module {module_name}: "
            f"{parent_path}"
        )
    else:
        for _name, visibility, shape, _body, line in modules:
            if visibility is not None and visibility.strip() == "pub":
                violations.append(
                    f"{label} codec child module must remain private: {parent_path}:{line}"
                )
            if shape == "inline":
                violations.append(
                    f"{label} codec child must use a dedicated source: {parent_path}:{line}"
                )

    if prepared_pattern.search(child_code) is None:
        violations.append(
            f"{label} codec child is missing its prepared removal plan: {child_path}"
        )
    if prepare_pattern.search(child_code) is None:
        violations.append(
            f"{label} codec child is missing its narrow prepare/rewrite entry point: "
            f"{child_path}"
        )

    for marker_label, markers in marker_groups.items():
        if not _boundary_marker_present(child_code, markers):
            violations.append(
                f"{label} codec child is missing {marker_label} marker: {child_path}"
            )

    if re.search(r"\b(?:prost|prost_types)\b|\b(?:Message|RawMessage)::(?:decode|encode)", child_code):
        violations.append(f"{label} codec child must not use Prost: {child_path}")
    for declaration, line in _rust_public_declarations(child_source):
        physical = IWA_DEPENDENCY_CODEC_PUBLIC_PHYSICAL.search(declaration)
        if physical is not None:
            violations.append(
                f"{label} codec child exposes a generated or owned wire type "
                f"{physical.group()}: {child_path}:{line}"
            )

    for name in (
        "PreparedPackageMetadataComponentRemovalRewrite"
        if "package_metadata" in str(parent_path)
        else "CalculationEngineOwnerRemovalPlan",
    ):
        # Parent re-export checks remain name-flexible for the dependency
        # child, but ensure at least the prepared public type crosses the
        # hidden facade rather than becoming an unreachable private helper.
        if name in child_code and name not in _rust_canonical_exports(
            parent_source, frozenset({name})
        ):
            violations.append(
                f"{label} codec parent does not re-export prepared type {name}: "
                f"{parent_path}"
            )

    build_absolute = root / IWA_CODEC_BUILD_SOURCE
    if not build_absolute.is_file():
        violations.append(
            f"{label} codec build source is missing: {IWA_CODEC_BUILD_SOURCE}"
        )
    else:
        build_source = build_absolute.read_text(encoding="utf-8")
        if build_marker not in build_source:
            violations.append(
                f"{label} codec build is missing its rerun marker: "
                f"{IWA_CODEC_BUILD_SOURCE}"
            )

    guard_absolute = root / IWA_CODEC_PRODUCTION_GUARD_SOURCE
    if not guard_absolute.is_file():
        violations.append(
            f"{label} codec production inventory source is missing: "
            f"{IWA_CODEC_PRODUCTION_GUARD_SOURCE}"
        )
    elif guard_marker not in guard_absolute.read_text(encoding="utf-8"):
        violations.append(
            f"{label} codec child is missing production inventory registration: "
            f"{IWA_CODEC_PRODUCTION_GUARD_SOURCE}"
        )

    return sorted(set(violations))


def audit_iwa_package_metadata_component_removal_source_topology(
    root: Path = ROOT,
) -> list[str]:
    return _audit_iwa_component_removal_codec_source_topology(
        root,
        parent_path=IWA_PACKAGE_METADATA_CODEC_SOURCE,
        child_path=IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_SOURCE,
        module_name=IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_MODULE,
        label="PackageMetadata component-removal",
        prepared_pattern=IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_PUBLIC_PREPARED,
        prepare_pattern=IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_PREPARE_FUNCTION,
        marker_groups=IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_MARKER_GROUPS,
        build_marker=IWA_PACKAGE_METADATA_COMPONENT_REMOVAL_BUILD_MARKER,
        guard_marker='include_str!("package_metadata_codec/component_removal.rs")',
    )


def audit_iwa_common_table_read_source_topology(root: Path = ROOT) -> list[str]:
    """Keep the archive-free table-read vocabulary in the common crate."""

    source_path = root / IWA_COMMON_TABLE_READ_SOURCE
    module_path = root / NEUTRAL_TABLE_MODULE_SOURCE
    if not source_path.is_file() and not module_path.is_file():
        return []

    violations: list[str] = []
    if not source_path.is_file():
        violations.append(
            "common table-read semantic owner is missing: "
            f"{IWA_COMMON_TABLE_READ_SOURCE}"
        )
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    if not module_path.is_file():
        violations.append(
            "common table module is missing its table-read export: "
            f"{NEUTRAL_TABLE_MODULE_SOURCE}"
        )
    else:
        module_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(module_path.read_text(encoding="utf-8"))
        )
        if IWA_COMMON_TABLE_READ_MODULE.search(module_code) is None:
            violations.append(
                "common table module is missing public table-read module: "
                f"{NEUTRAL_TABLE_MODULE_SOURCE}"
            )

    for name in sorted(IWA_COMMON_TABLE_READ_PUBLIC_TYPES):
        declaration = re.compile(
            r"(?m)^[ \t]*pub[ \t]+(?:struct|enum|union|type)[ \t]+"
            rf"(?:r#)?{re.escape(name)}\b"
        )
        if declaration.search(code) is None:
            violations.append(
                "common table-read semantic owner is missing public "
                f"{name}: {IWA_COMMON_TABLE_READ_SOURCE}"
            )

    for statement in IWA_COMMON_TABLE_READ_IMPORT_STATEMENT.finditer(code):
        forbidden = IWA_COMMON_TABLE_READ_FORBIDDEN_IMPORT.search(statement.group())
        if forbidden is None:
            continue
        line_number = code.count("\n", 0, statement.start() + forbidden.start()) + 1
        violations.append(
            "common table-read semantic owner imports a concrete format, wire, "
            f"or protobuf crate ({forbidden.group()}): "
            f"{IWA_COMMON_TABLE_READ_SOURCE}:{line_number}"
        )

    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(raw_source)
    ):
        physical = IWA_TABLE_READ_PUBLIC_PHYSICAL_TYPES.search(declaration)
        if physical is not None:
            physical_line = line_number + declaration.count("\n", 0, physical.start())
            violations.append(
                "common table-read semantic owner exposes a physical or generated "
                f"type {physical.group()}: {IWA_COMMON_TABLE_READ_SOURCE}:{physical_line}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                "common table-read semantic owner exposes raw bytes: "
                f"{IWA_COMMON_TABLE_READ_SOURCE}:{line_number}"
            )
        raw_id = IWA_TABLE_READ_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            raw_id_line = line_number + declaration.count("\n", 0, raw_id.start())
            violations.append(
                "common table-read semantic owner exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: "
                f"{IWA_COMMON_TABLE_READ_SOURCE}:{raw_id_line}"
            )

    return sorted(set(violations))


def _audit_focused_table_cells_source_topology(
    root: Path,
    *,
    format_name: str,
    owner_path: Path,
    package_path: Path,
    module_pattern: re.Pattern[str],
    public_method: str,
) -> list[str]:
    """Protect one focused selector-first table-cell semantic adapter."""

    owner_absolute = root / owner_path
    package_absolute = root / package_path
    package_source = (
        package_absolute.read_text(encoding="utf-8")
        if package_absolute.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(_mask_rust_cfg_test_items(package_source))
    module_match = module_pattern.search(package_code)

    implementation_sources: list[tuple[Path, str]] = []
    if package_absolute.is_file():
        implementation_sources.append((package_path, package_source))
    package_directory = package_absolute.parent / package_absolute.stem
    if package_directory.is_dir():
        implementation_sources.extend(
            (path.relative_to(root), path.read_text(encoding="utf-8"))
            for path in sorted(package_directory.glob("*.rs"))
        )
    method_evidence = [
        (path, declaration, body, line)
        for path, source in implementation_sources
        for declaration, body, line in _rust_public_method_bodies(source, public_method)
    ]

    # A completely absent focused API is the expected end state when a host
    # package is retired.  A package module, owner file, or public method is a
    # partial migration and therefore remains subject to the ratchet.
    if not owner_absolute.is_file() and module_match is None and not method_evidence:
        return []

    violations: list[str] = []
    if not owner_absolute.is_file():
        violations.append(
            f"{format_name} table-cell focused owner is missing: {owner_path}"
        )
    if module_match is None:
        violations.append(
            f"{format_name} table-cell package is missing its focused module: "
            f"{package_path}"
        )
    module_name = "slide_table_cells" if "Keynote" == format_name else "table_cells"
    if re.search(
        rf"(?m)^[ \t]*pub[ \t]+mod[ \t]+(?:r#)?{re.escape(module_name)}\b",
        package_code,
    ):
        violations.append(
            f"{format_name} table-cell focused module must remain private: "
            f"{package_path}"
        )

    focused_source = (
        owner_absolute.read_text(encoding="utf-8")
        if owner_absolute.is_file()
        else package_source
    )
    focused_code = _mask_rust_non_code(_mask_rust_cfg_test_items(focused_source))
    relative_owner = owner_path if owner_absolute.is_file() else package_path

    for pattern, description in (
        (IWA_FOCUSED_TABLE_CELLS_COMMON_IMPORT, "common table-read"),
        (IWA_FOCUSED_TABLE_CELLS_WIRE_IMPORT, "Numbers wire table-cell"),
        (IWA_FOCUSED_TABLE_CELLS_SIDECAR_IMPORT, "Numbers wire table-sidecar"),
    ):
        if pattern.search(focused_code) is None:
            violations.append(
                f"{format_name} table-cell focused owner is missing its "
                f"{description} import: {relative_owner}"
            )
    for pattern, description in (
        (IWA_FOCUSED_TABLE_CELLS_WIRE_CALL, "read_table_cells"),
        (IWA_FOCUSED_TABLE_CELLS_SIDECAR_CALL, "read_sidecar_list"),
        (IWA_FOCUSED_TABLE_CELLS_COMMON_CONSTRUCTOR, "a TableRead constructor"),
    ):
        if pattern.search(focused_code) is None:
            violations.append(
                f"{format_name} table-cell focused owner must delegate through "
                f"{description}: {relative_owner}"
            )

    for statement in IWA_COMMON_TABLE_READ_IMPORT_STATEMENT.finditer(focused_code):
        peer = IWA_FOCUSED_TABLE_CELLS_PEER_IMPORT.search(statement.group())
        if peer is None:
            continue
        line_number = focused_code.count("\n", 0, statement.start() + peer.start()) + 1
        violations.append(
            f"{format_name} table-cell focused owner imports a concrete format peer "
            f"{peer.group()}: {relative_owner}:{line_number}"
        )

    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(focused_source)
    ):
        generated = IWA_FOCUSED_TABLE_CELLS_GENERATED_PUBLIC.search(declaration)
        if generated is not None:
            generated_line = line_number + declaration.count("\n", 0, generated.start())
            violations.append(
                f"{format_name} table-cell focused public API exposes generated or "
                f"wire type {generated.group()}: {relative_owner}:{generated_line}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                f"{format_name} table-cell focused public API exposes raw bytes: "
                f"{relative_owner}:{line_number}"
            )
        raw_id = IWA_TABLE_READ_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            raw_id_line = line_number + declaration.count("\n", 0, raw_id.start())
            violations.append(
                f"{format_name} table-cell focused public API exposes a raw ID "
                f"parameter {raw_id.group(0).strip()}: {relative_owner}:{raw_id_line}"
            )

    budget_owners: dict[str, list[str]] = {"TableCellReadBudget": [], "SidecarReadBudget": []}
    budget_impls = list(IWA_FOCUSED_TABLE_CELLS_BUDGET_IMPL.finditer(focused_code))
    for match in budget_impls:
        budget_owners[match.group("trait")].append(match.group("owner"))
    for trait, owners in budget_owners.items():
        if not owners:
            violations.append(
                f"{format_name} table-cell focused owner is missing its {trait} "
                f"implementation: {relative_owner}"
            )
        elif len(owners) > 1:
            violations.append(
                f"{format_name} table-cell focused owner has multiple {trait} "
                f"implementations: {relative_owner}"
            )
    present_owners = {owners[0] for owners in budget_owners.values() if len(owners) == 1}
    if len(present_owners) == 2:
        violations.append(
            f"{format_name} table-cell focused owner uses separate aggregate budget "
            f"types ({', '.join(sorted(present_owners))}): {relative_owner}"
        )

    if not method_evidence:
        violations.append(
            f"{format_name} table-cell focused owner is missing public "
            f"Package::{public_method}: {owner_path}"
        )
    else:
        for path, declaration, _body, line_number in method_evidence:
            if re.search(r"\bTableRead\b", declaration) is None:
                violations.append(
                    f"{format_name} Package::{public_method} must return common "
                    f"TableRead values: {path}:{line_number}"
                )
            if RUST_BYTE_SLICE.search(declaration) is not None:
                violations.append(
                    f"{format_name} Package::{public_method} exposes raw bytes: "
                    f"{path}:{line_number}"
                )
            raw_id = IWA_TABLE_READ_RAW_ID_PARAMETER.search(declaration)
            if raw_id is not None:
                violations.append(
                    f"{format_name} Package::{public_method} exposes a raw ID "
                    f"parameter {raw_id.group(0).strip()}: {path}:{line_number}"
                )

    return sorted(set(violations))


def _chart_metadata_boundary_active(root: Path) -> bool:
    """Return whether the chart-metadata migration has a production owner."""

    if (root / IWA_CHART_METADATA_CODEC_SOURCE).is_file():
        return True
    codec_public = root / IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE
    if codec_public.is_file():
        codec_source = _mask_rust_cfg_test_items(
            codec_public.read_text(encoding="utf-8")
        )
        if re.search(
            rf"(?m)^\s*(?:pub\s+)?mod\s+"
            rf"(?:r#)?{re.escape(IWA_CHART_METADATA_CODEC_MODULE)}\b",
            codec_source,
        ):
            return True
    for format_name, owner_path in IWA_CHART_METADATA_FOCUSED_OWNERS.items():
        if (root / owner_path).is_file():
            return True
        package_path = root / IWA_CHART_METADATA_FOCUSED_PACKAGES[format_name]
        if not package_path.is_file():
            continue
        package_source = _mask_rust_cfg_test_items(
            package_path.read_text(encoding="utf-8")
        )
        module_name = IWA_CHART_METADATA_FOCUSED_MODULES[format_name]
        if re.search(
            rf"(?m)^\s*(?:pub(?:\([^()]*\))?\s+)?mod\s+"
            rf"(?:r#)?{re.escape(module_name)}\b",
            package_source,
        ):
            return True
    return False


def _chart_metadata_external_public_declaration(declaration: str) -> bool:
    """Return whether a Rust declaration is externally public."""

    return re.match(r"\s*pub(?!\s*\()", declaration) is not None


def _chart_metadata_check_common_source(root: Path) -> list[str]:
    """Check the archive-free common chart metadata value."""

    source_path = root / IWA_COMMON_CHART_METADATA_SOURCE
    module_path = root / IWA_COMMON_CHART_METADATA_MODULE_SOURCE
    violations: list[str] = []
    if not source_path.is_file():
        return [
            "common chart-metadata semantic owner is missing: "
            f"{IWA_COMMON_CHART_METADATA_SOURCE}"
        ]

    raw_source = source_path.read_text(encoding="utf-8")
    code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    if not module_path.is_file():
        violations.append(
            "common chart-metadata module export source is missing: "
            f"{IWA_COMMON_CHART_METADATA_MODULE_SOURCE}"
        )
    else:
        module_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(module_path.read_text(encoding="utf-8"))
        )
        if IWA_COMMON_CHART_METADATA_MODULE.search(module_code) is None:
            violations.append(
                "common chart module is missing public chart-metadata export: "
                f"{IWA_COMMON_CHART_METADATA_MODULE_SOURCE}"
            )

    if re.search(r"(?m)^\s*pub\s+struct\s+ChartMetadata\b", code) is None:
        violations.append(
            "common chart-metadata semantic owner is missing public ChartMetadata: "
            f"{IWA_COMMON_CHART_METADATA_SOURCE}"
        )
    if re.search(r"\bpub\s+fn\s+from_owned\b", code) is None:
        violations.append(
            "common chart-metadata semantic owner is missing its owned constructor: "
            f"{IWA_COMMON_CHART_METADATA_SOURCE}"
        )

    import_statement = re.compile(
        r"(?ms)^[ \t]*(?:pub(?:\([^()\r\n]*\))?[ \t]+)?"
        r"(?:use|extern[ \t]+crate)\b.*?;"
    )
    for statement in import_statement.finditer(code):
        forbidden = IWA_COMMON_CHART_METADATA_FORBIDDEN_IMPORT.search(statement.group())
        if forbidden is None:
            continue
        line_number = code.count("\n", 0, statement.start() + forbidden.start()) + 1
        violations.append(
            "common chart-metadata semantic owner imports a concrete format, wire, "
            f"or protobuf crate ({forbidden.group()}): "
            f"{IWA_COMMON_CHART_METADATA_SOURCE}:{line_number}"
        )

    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(raw_source)
    ):
        if not _chart_metadata_external_public_declaration(declaration):
            continue
        physical = IWA_CHART_METADATA_PUBLIC_PHYSICAL_TYPES.search(declaration)
        if physical is not None:
            physical_line = line_number + declaration.count("\n", 0, physical.start())
            violations.append(
                "common chart-metadata semantic owner exposes a physical or generated "
                f"type {physical.group()}: {IWA_COMMON_CHART_METADATA_SOURCE}:{physical_line}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                "common chart-metadata semantic owner exposes raw bytes: "
                f"{IWA_COMMON_CHART_METADATA_SOURCE}:{line_number}"
            )
        raw_id = IWA_CHART_METADATA_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            raw_id_line = line_number + declaration.count("\n", 0, raw_id.start())
            violations.append(
                "common chart-metadata semantic owner exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: "
                f"{IWA_COMMON_CHART_METADATA_SOURCE}:{raw_id_line}"
            )
    return sorted(set(violations))


def _chart_metadata_check_codec_source(root: Path) -> list[str]:
    """Check the private Buffa chart-metadata projection and codec seam."""

    codec_path = root / IWA_CHART_METADATA_CODEC_SOURCE
    public_path = root / IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE
    violations: list[str] = []
    if not codec_path.is_file():
        return [
            "neutral chart-metadata codec source is missing: "
            f"{IWA_CHART_METADATA_CODEC_SOURCE}"
        ]

    raw_codec = codec_path.read_text(encoding="utf-8")
    codec = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_codec))
    for name in IWA_CHART_METADATA_CODEC_REQUIRED_TYPES:
        if re.search(
            rf"\b(?:pub\s+)?(?:struct|enum|type)\s+{re.escape(name)}\b",
            codec,
        ) is None:
            violations.append(
                "neutral chart-metadata codec is missing strict API "
                f"{name}: {IWA_CHART_METADATA_CODEC_SOURCE}"
            )

    def marker_present(marker: str | tuple[str, ...]) -> bool:
        if isinstance(marker, tuple):
            return any(marker_present(item) for item in marker)
        return re.search(rf"\b{re.escape(marker)}\b", codec) is not None

    for label, markers in IWA_CHART_METADATA_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not all(marker_present(marker) for marker in markers):
            violations.append(
                "neutral chart-metadata codec is missing "
                f"{label} marker: {IWA_CHART_METADATA_CODEC_SOURCE}"
            )

    if IWA_CHART_METADATA_CODEC_PROST.search(codec) is not None:
        violations.append(
            "neutral chart-metadata codec must use Buffa rather than Prost: "
            f"{IWA_CHART_METADATA_CODEC_SOURCE}"
        )
    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(raw_codec)
    ):
        if not _chart_metadata_external_public_declaration(declaration):
            continue
        if re.search(
            rf"\b{re.escape(IWA_CHART_METADATA_CODEC_GENERATED_MODULE)}\b|"
            r"\bprojection\s*::",
            declaration,
        ):
            violations.append(
                "neutral chart-metadata codec must not expose Buffa generated "
                f"types: {IWA_CHART_METADATA_CODEC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        codec,
        frozenset({IWA_CHART_METADATA_CODEC_GENERATED_MODULE}),
    ):
        violations.append(
            "neutral chart-metadata codec must not expose Buffa generated aliases: "
            f"{IWA_CHART_METADATA_CODEC_SOURCE}"
        )

    if not public_path.is_file():
        violations.append(
            "neutral chart-metadata codec public module source is missing: "
            f"{IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE}"
        )
        return sorted(set(violations))

    raw_public = public_path.read_text(encoding="utf-8")
    public = _mask_rust_cfg_test_items(raw_public)
    if _rust_root_level_matches(public, IWA_CHART_METADATA_CODEC_HIDDEN_MODULE) == ():
        violations.append(
            "neutral chart-metadata codec module must be hidden: "
            f"{IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        public,
        frozenset({IWA_CHART_METADATA_CODEC_GENERATED_MODULE}),
    )
    if not generated_declarations:
        violations.append(
            "neutral chart-metadata codec is missing its private Buffa generated "
            f"module: {IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE}"
        )
    for _, visibility, shape, body, line_number in generated_declarations:
        if visibility is not None:
            violations.append(
                "neutral chart-metadata codec generated module must remain private: "
                f"{IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "neutral chart-metadata codec generated module has invalid include "
                f"shape: {IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        public,
        frozenset({IWA_CHART_METADATA_CODEC_GENERATED_MODULE}),
    ):
        violations.append(
            "neutral chart-metadata codec must not re-export Buffa generated types: "
            f"{IWA_CHART_METADATA_CODEC_PUBLIC_SOURCE}"
        )
    return sorted(set(violations))


def _audit_focused_chart_metadata_owner(
    root: Path,
    *,
    format_name: str,
    owner_path: Path,
    package_path: Path,
    module_name: str,
    public_method: str,
) -> list[str]:
    """Check one selector-first chart metadata package owner."""

    owner_absolute = root / owner_path
    package_absolute = root / package_path
    package_source = (
        package_absolute.read_text(encoding="utf-8")
        if package_absolute.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(_mask_rust_cfg_test_items(package_source))
    module_match = re.search(
        rf"(?m)^[ \t]*(?P<visibility>pub(?:\([^()]*\))?[ \t]+)?"
        rf"mod[ \t]+(?:r#)?{re.escape(module_name)}\b[ \t]*(?:;|\{{)",
        package_code,
    )
    owner_source = (
        owner_absolute.read_text(encoding="utf-8")
        if owner_absolute.is_file()
        else package_source
    )
    owner_code = _mask_rust_non_code(_mask_rust_cfg_test_items(owner_source))
    method_evidence = _rust_public_method_bodies(owner_source, public_method)
    if not owner_absolute.is_file() and module_match is None and not method_evidence:
        return []

    violations: list[str] = []
    if not owner_absolute.is_file():
        violations.append(
            f"focused {format_name} chart-metadata owner is missing: {owner_path}"
        )
    if module_match is None:
        violations.append(
            f"focused {format_name} chart-metadata package is missing its private "
            f"module: {package_path}"
        )
    elif module_match.group("visibility") and not module_match.group("visibility").startswith(
        "pub("
    ):
        line_number = package_code.count("\n", 0, module_match.start()) + 1
        violations.append(
            f"focused {format_name} chart-metadata module must remain private: "
            f"{package_path}:{line_number}"
        )

    if IWA_CHART_METADATA_COMMON_IMPORT.search(owner_code) is None:
        violations.append(
            f"focused {format_name} chart-metadata owner is missing its common "
            f"ChartMetadata import: {owner_path}"
        )
    if "chart_metadata_codec" not in owner_code:
        violations.append(
            f"focused {format_name} chart-metadata owner must delegate to the shared "
            f"codec: {owner_path}"
        )
    if IWA_CHART_METADATA_FOCUSED_FORBIDDEN_PROST.search(owner_code) is not None:
        violations.append(
            f"focused {format_name} chart-metadata owner retains production Prost or "
            f"generated metadata decoding: {owner_path}"
        )

    relative_owner = owner_path if owner_absolute.is_file() else package_path
    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(owner_source)
    ):
        if not _chart_metadata_external_public_declaration(declaration):
            continue
        physical = IWA_CHART_METADATA_PUBLIC_PHYSICAL_TYPES.search(declaration)
        if physical is not None:
            physical_line = line_number + declaration.count("\n", 0, physical.start())
            violations.append(
                f"focused {format_name} chart-metadata public API exposes generated or "
                f"wire type {physical.group()}: {relative_owner}:{physical_line}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                f"focused {format_name} chart-metadata public API exposes raw bytes: "
                f"{relative_owner}:{line_number}"
            )
        raw_id = IWA_CHART_METADATA_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            raw_id_line = line_number + declaration.count("\n", 0, raw_id.start())
            violations.append(
                f"focused {format_name} chart-metadata public API exposes a raw ID "
                f"parameter {raw_id.group(0).strip()}: {relative_owner}:{raw_id_line}"
            )

    if not method_evidence:
        violations.append(
            f"focused {format_name} chart-metadata owner is missing public "
            f"Package::{public_method}: {owner_path}"
        )
    else:
        for declaration, _body, line_number in method_evidence:
            if re.search(r"\bChartMetadata\b", declaration) is None:
                violations.append(
                    f"focused {format_name} Package::{public_method} must return common "
                    f"ChartMetadata values: {owner_path}:{line_number}"
                )
            if re.search(r"\b[A-Za-z_][A-Za-z0-9_]*Selector\b", declaration) is None:
                violations.append(
                    f"focused {format_name} Package::{public_method} must accept a "
                    f"semantic selector: {owner_path}:{line_number}"
                )
            if RUST_BYTE_SLICE.search(declaration) is not None:
                violations.append(
                    f"focused {format_name} Package::{public_method} exposes raw bytes: "
                    f"{owner_path}:{line_number}"
                )
            raw_id = IWA_CHART_METADATA_RAW_ID_PARAMETER.search(declaration)
            if raw_id is not None:
                violations.append(
                    f"focused {format_name} Package::{public_method} exposes a raw ID "
                    f"parameter {raw_id.group(0).strip()}: {owner_path}:{line_number}"
                )

    return sorted(set(violations))


def audit_iwa_chart_metadata_source_topology(root: Path = ROOT) -> list[str]:
    """Keep chart metadata common, private-wire, and focused seams separated."""

    if not _chart_metadata_boundary_active(root):
        return []

    violations = _chart_metadata_check_common_source(root)
    violations.extend(_chart_metadata_check_codec_source(root))
    for format_name, owner_path in IWA_CHART_METADATA_FOCUSED_OWNERS.items():
        violations.extend(
            _audit_focused_chart_metadata_owner(
                root,
                format_name=format_name,
                owner_path=owner_path,
                package_path=IWA_CHART_METADATA_FOCUSED_PACKAGES[format_name],
                module_name=IWA_CHART_METADATA_FOCUSED_MODULES[format_name],
                public_method=IWA_CHART_METADATA_FOCUSED_METHODS[format_name],
            )
        )

    return sorted(set(violations))


def _chart_data_boundary_active(root: Path) -> bool:
    """Return whether the chart-data migration has a production owner."""

    if (root / IWA_CHART_DATA_CODEC_SOURCE).is_file():
        return True
    codec_public = root / IWA_CHART_DATA_CODEC_PUBLIC_SOURCE
    if codec_public.is_file():
        codec_source = _mask_rust_cfg_test_items(
            codec_public.read_text(encoding="utf-8")
        )
        if re.search(
            rf"(?m)^\s*(?:pub\s+)?mod\s+"
            rf"(?:r#)?{re.escape(IWA_CHART_DATA_CODEC_MODULE)}\b",
            codec_source,
        ):
            return True
    for format_name, owner_path in IWA_CHART_DATA_FOCUSED_OWNERS.items():
        if (root / owner_path).is_file():
            return True
        package_path = root / IWA_CHART_DATA_FOCUSED_PACKAGES[format_name]
        if not package_path.is_file():
            continue
        package_source = _mask_rust_cfg_test_items(
            package_path.read_text(encoding="utf-8")
        )
        module_name = IWA_CHART_DATA_FOCUSED_MODULES[format_name]
        if re.search(
            rf"(?m)^\s*(?:pub(?:\([^()]*\))?\s+)?mod\s+"
            rf"(?:r#)?{re.escape(module_name)}\b",
            package_source,
        ):
            return True
    return False


def _chart_data_check_common_source(root: Path) -> list[str]:
    """Check the archive-free common chart data value."""

    source_path = root / IWA_COMMON_CHART_DATA_SOURCE
    module_path = root / IWA_COMMON_CHART_DATA_MODULE_SOURCE
    violations: list[str] = []
    if not source_path.is_file():
        return [
            "common chart-data semantic owner is missing: "
            f"{IWA_COMMON_CHART_DATA_SOURCE}"
        ]

    raw_source = source_path.read_text(encoding="utf-8")
    code = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_source))
    if not module_path.is_file():
        violations.append(
            "common chart-data module export source is missing: "
            f"{IWA_COMMON_CHART_DATA_MODULE_SOURCE}"
        )
    else:
        module_code = _mask_rust_non_code(
            _mask_rust_cfg_test_items(module_path.read_text(encoding="utf-8"))
        )
        if IWA_COMMON_CHART_DATA_MODULE.search(module_code) is None:
            violations.append(
                "common chart module is missing public chart-data export: "
                f"{IWA_COMMON_CHART_DATA_MODULE_SOURCE}"
            )

    if re.search(r"(?m)^\s*pub\s+struct\s+ChartData\b", code) is None:
        violations.append(
            "common chart-data semantic owner is missing public ChartData: "
            f"{IWA_COMMON_CHART_DATA_SOURCE}"
        )
    if re.search(r"(?m)^\s*pub\s+(?:enum|struct)\s+DataError\b", code) is None:
        violations.append(
            "common chart-data semantic owner is missing typed DataError: "
            f"{IWA_COMMON_CHART_DATA_SOURCE}"
        )
    for marker, label in (
        (r"\bpub\s+fn\s+new\b", "constructor"),
        (r"\bpub\s+fn\s+row_names\b", "row-label accessor"),
        (r"\bpub\s+fn\s+column_names\b", "column-label accessor"),
        (r"\bpub\s+fn\s+values\b", "value accessor"),
    ):
        if re.search(marker, code) is None:
            violations.append(
                "common chart-data semantic owner is missing its "
                f"{label}: {IWA_COMMON_CHART_DATA_SOURCE}"
            )

    import_statement = re.compile(
        r"(?ms)^[ \t]*(?:pub(?:\([^()\r\n]*\))?[ \t]+)?"
        r"(?:use|extern[ \t]+crate)\b.*?;"
    )
    for statement in import_statement.finditer(code):
        forbidden = IWA_COMMON_CHART_DATA_FORBIDDEN_IMPORT.search(statement.group())
        if forbidden is None:
            continue
        line_number = code.count("\n", 0, statement.start() + forbidden.start()) + 1
        violations.append(
            "common chart-data semantic owner imports a concrete format, wire, "
            f"or protobuf crate ({forbidden.group()}): "
            f"{IWA_COMMON_CHART_DATA_SOURCE}:{line_number}"
        )

    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(raw_source)
    ):
        if not _chart_metadata_external_public_declaration(declaration):
            continue
        physical = IWA_CHART_DATA_PUBLIC_PHYSICAL_TYPES.search(declaration)
        if physical is not None:
            physical_line = line_number + declaration.count(
                "\n", 0, physical.start()
            )
            violations.append(
                "common chart-data semantic owner exposes a physical or generated "
                f"type {physical.group()}: {IWA_COMMON_CHART_DATA_SOURCE}:{physical_line}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                "common chart-data semantic owner exposes raw bytes: "
                f"{IWA_COMMON_CHART_DATA_SOURCE}:{line_number}"
            )
        raw_id = IWA_CHART_DATA_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            raw_id_line = line_number + declaration.count("\n", 0, raw_id.start())
            violations.append(
                "common chart-data semantic owner exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: "
                f"{IWA_COMMON_CHART_DATA_SOURCE}:{raw_id_line}"
            )
    return sorted(set(violations))


def _chart_data_check_codec_source(root: Path) -> list[str]:
    """Check the private Buffa chart-data projection and codec seam."""

    codec_path = root / IWA_CHART_DATA_CODEC_SOURCE
    public_path = root / IWA_CHART_DATA_CODEC_PUBLIC_SOURCE
    violations: list[str] = []
    if not codec_path.is_file():
        return [
            "neutral chart-data codec source is missing: "
            f"{IWA_CHART_DATA_CODEC_SOURCE}"
        ]

    projection_path = root / IWA_CHART_DATA_CODEC_PROJECTION_SOURCE
    if not projection_path.is_file():
        violations.append(
            "neutral chart-data codec is missing its private Buffa projection: "
            f"{projection_path.relative_to(root)}"
        )

    raw_codec = codec_path.read_text(encoding="utf-8")
    codec = _mask_rust_non_code(_mask_rust_cfg_test_items(raw_codec))
    for name in IWA_CHART_DATA_CODEC_REQUIRED_TYPES:
        if re.search(
            rf"\b(?:pub\s+)?(?:struct|enum|type)\s+{re.escape(name)}\b",
            codec,
        ) is None:
            violations.append(
                "neutral chart-data codec is missing strict API "
                f"{name}: {IWA_CHART_DATA_CODEC_SOURCE}"
            )

    def marker_present(marker: str | tuple[str, ...]) -> bool:
        if isinstance(marker, tuple):
            return any(marker_present(item) for item in marker)
        return re.search(rf"\b{re.escape(marker)}\b", codec) is not None

    for label, markers in IWA_CHART_DATA_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not all(marker_present(marker) for marker in markers):
            violations.append(
                "neutral chart-data codec is missing "
                f"{label} marker: {IWA_CHART_DATA_CODEC_SOURCE}"
            )

    if IWA_CHART_DATA_CODEC_PROST.search(codec) is not None:
        violations.append(
            "neutral chart-data codec must use Buffa rather than Prost: "
            f"{IWA_CHART_DATA_CODEC_SOURCE}"
        )
    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(raw_codec)
    ):
        if not _chart_metadata_external_public_declaration(declaration):
            continue
        if re.search(
            rf"\b{re.escape(IWA_CHART_DATA_CODEC_GENERATED_MODULE)}\b|"
            r"\bprojection\s*::",
            declaration,
        ):
            violations.append(
                "neutral chart-data codec must not expose Buffa generated "
                f"types: {IWA_CHART_DATA_CODEC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        codec, frozenset({IWA_CHART_DATA_CODEC_GENERATED_MODULE})
    ):
        violations.append(
            "neutral chart-data codec must not expose Buffa generated aliases: "
            f"{IWA_CHART_DATA_CODEC_SOURCE}"
        )

    if not public_path.is_file():
        violations.append(
            "neutral chart-data codec public module source is missing: "
            f"{IWA_CHART_DATA_CODEC_PUBLIC_SOURCE}"
        )
        return sorted(set(violations))

    raw_public = public_path.read_text(encoding="utf-8")
    public = _mask_rust_cfg_test_items(raw_public)
    if _rust_root_level_matches(public, IWA_CHART_DATA_CODEC_HIDDEN_MODULE) == ():
        violations.append(
            "neutral chart-data codec module must be hidden: "
            f"{IWA_CHART_DATA_CODEC_PUBLIC_SOURCE}"
        )
    generated_declarations = _rust_root_level_module_declarations(
        public, frozenset({IWA_CHART_DATA_CODEC_GENERATED_MODULE})
    )
    if not generated_declarations:
        violations.append(
            "neutral chart-data codec is missing its private Buffa generated "
            f"module: {IWA_CHART_DATA_CODEC_PUBLIC_SOURCE}"
        )
    for _, visibility, shape, body, line_number in generated_declarations:
        if visibility is not None:
            violations.append(
                "neutral chart-data codec generated module must remain private: "
                f"{IWA_CHART_DATA_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
        if shape == "invalid" or (
            shape == "inline" and not _rust_generated_module_include_is_valid(body)
        ):
            violations.append(
                "neutral chart-data codec generated module has invalid include "
                f"shape: {IWA_CHART_DATA_CODEC_PUBLIC_SOURCE}:{line_number}"
            )
    if _rust_root_generated_public_leaks(
        public, frozenset({IWA_CHART_DATA_CODEC_GENERATED_MODULE})
    ):
        violations.append(
            "neutral chart-data codec must not re-export Buffa generated types: "
            f"{IWA_CHART_DATA_CODEC_PUBLIC_SOURCE}"
        )
    return sorted(set(violations))


def _chart_data_codec_production_sources(root: Path) -> tuple[tuple[Path, str], ...]:
    """Return chart-data codec production sources, including rewrite modules.

    The decoder started life as one source file and may grow a private
    ``chart_data_codec/`` implementation directory as the writer becomes
    larger.  Keep the boundary about the production seam rather than forcing
    that implementation layout.  Test-only sources are deliberately excluded
    because differential Prost fixtures are allowed there.
    """

    codec_path = root / IWA_CHART_DATA_CODEC_SOURCE
    if not codec_path.is_file():
        return ()
    paths = [codec_path]
    module_directory = codec_path.with_suffix("")
    if module_directory.is_dir():
        paths.extend(
            path
            for path in sorted(module_directory.glob("*.rs"))
            if path.name not in {"tests.rs", "test.rs"}
            and not path.stem.endswith("_tests")
        )
    return tuple(
        (
            path,
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8")),
        )
        for path in paths
    )


def _chart_data_check_codec_rewrite_source(root: Path) -> list[str]:
    """Require a bounded Buffa/source-preserving chart-data write seam."""

    sources = _chart_data_codec_production_sources(root)
    if not sources:
        return []
    production = "\n".join(source for _path, source in sources)
    code = _mask_rust_non_code(production)
    violations: list[str] = []

    if IWA_CHART_DATA_CODEC_PROST.search(code) is not None:
        violations.append(
            "neutral chart-data rewrite codec must use Buffa rather than Prost: "
            f"{IWA_CHART_DATA_CODEC_SOURCE}"
        )

    for name in IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_TYPES:
        if re.search(
            rf"\b(?:pub(?:\([^()]*\))?[ \t]+)?"
            rf"(?:struct|enum|type)\s+{re.escape(name)}\b",
            code,
        ) is None:
            violations.append(
                "neutral chart-data rewrite codec is missing strict API "
                f"{name}: {IWA_CHART_DATA_CODEC_SOURCE}"
            )

    for label, alternatives in IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_TYPE_GROUPS.items():
        if not any(
            re.search(
                rf"\b(?:pub(?:\([^()]*\))?[ \t]+)?"
                rf"(?:struct|enum|type)\s+{re.escape(name)}\b",
                code,
            )
            for name in alternatives
        ):
            violations.append(
                "neutral chart-data rewrite codec is missing "
                f"{label}: {IWA_CHART_DATA_CODEC_SOURCE}"
            )

    for name in IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_FUNCTIONS:
        if re.search(
            rf"\b(?:pub(?:\([^()]*\))?[ \t]+)?fn\s+"
            rf"{re.escape(name)}\b",
            code,
        ) is None:
            violations.append(
                "neutral chart-data rewrite codec is missing strict API "
                f"{name}: {IWA_CHART_DATA_CODEC_SOURCE}"
            )

    def marker_present(marker: str | tuple[str, ...]) -> bool:
        if isinstance(marker, tuple):
            return all(
                re.search(rf"\b{re.escape(item)}\b", code) is not None
                for item in marker
            )
        return re.search(rf"\b{re.escape(marker)}\b", code) is not None

    for label, markers in IWA_CHART_DATA_CODEC_REWRITE_REQUIRED_MARKER_GROUPS.items():
        alternatives = bool(markers) and isinstance(markers[0], tuple)
        present = (
            any(marker_present(option) for option in markers)
            if alternatives
            else all(marker_present(marker) for marker in markers)
        )
        if not present:
            violations.append(
                "neutral chart-data rewrite codec is missing "
                f"{label}: {IWA_CHART_DATA_CODEC_SOURCE}"
            )

    return sorted(set(violations))


def _rust_method_declaration(source: str, method_name: str) -> str | None:
    """Return one public method declaration without descending into its body."""

    masked = _mask_rust_cfg_test_items(source)
    code = _mask_rust_non_code(masked)
    pattern = re.compile(
        r"(?m)^[ \t]*pub(?![ \t\r\n]*\()[ \t\r\n]+"
        r"(?:(?:unsafe|async|const)[ \t\r\n]+)*fn[ \t\r\n]+"
        rf"(?:r#)?{re.escape(method_name)}\b"
    )
    match = pattern.search(code)
    if match is None:
        return None
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
        elif not parentheses and not brackets and character in "{;":
            break
        cursor += 1
    if cursor >= len(code) or code[cursor] != "{":
        return None
    return code[match.start() : cursor]


def _focused_chart_data_write_sources(
    root: Path, owner_path: Path
) -> tuple[tuple[Path, str], ...]:
    """Return one focused chart-data owner and its narrowly named siblings."""

    absolute = root / owner_path
    if not absolute.is_file():
        return ()
    paths = [absolute]
    paths.extend(
        path
        for path in sorted(absolute.parent.glob(f"{absolute.stem}*.rs"))
        if path != absolute
    )
    module_directory = absolute.parent / absolute.stem
    if module_directory.is_dir():
        paths.extend(sorted(module_directory.glob("*.rs")))
    unique_paths = tuple(dict.fromkeys(paths))
    return tuple(
        (path, path.read_text(encoding="utf-8")) for path in unique_paths
    )


def _audit_focused_chart_data_write_owner(
    root: Path,
    *,
    format_name: str,
    owner_path: Path,
    edit_method: str,
    apply_method: str,
) -> list[str]:
    """Check one focused package's semantic chart-data transaction seam."""

    sources = _focused_chart_data_write_sources(root, owner_path)
    if not sources:
        return []
    source = "\n".join(source for _path, source in sources)
    code = _mask_rust_non_code(_mask_rust_cfg_test_items(source))
    violations: list[str] = []

    edit_declaration = next(
        (
            declaration
            for _path, implementation in sources
            if (declaration := _rust_method_declaration(implementation, edit_method))
            is not None
        ),
        None,
    )
    apply_declaration = next(
        (
            declaration
            for _path, implementation in sources
            if (declaration := _rust_method_declaration(implementation, apply_method))
            is not None
        ),
        None,
    )
    if edit_declaration is None:
        violations.append(
            f"focused {format_name} chart-data owner is missing public "
            f"Package::{edit_method}: {owner_path}"
        )
    if apply_declaration is None:
        violations.append(
            f"focused {format_name} chart-data owner is missing public "
            f"Package::{apply_method}: {owner_path}"
        )

    if edit_declaration is not None:
        if re.search(r"\b[A-Za-z_][A-Za-z0-9_]*Selector\b", edit_declaration) is None:
            violations.append(
                f"focused {format_name} {edit_method} must accept a semantic selector: "
                f"{owner_path}"
            )
        if RUST_BYTE_SLICE.search(edit_declaration) is not None:
            violations.append(
                f"focused {format_name} {edit_method} exposes raw bytes: {owner_path}"
            )
        raw_id = IWA_CHART_DATA_RAW_ID_PARAMETER.search(edit_declaration)
        if raw_id is not None:
            violations.append(
                f"focused {format_name} {edit_method} exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: {owner_path}"
            )

    if apply_declaration is not None:
        if RUST_BYTE_SLICE.search(apply_declaration) is not None:
            violations.append(
                f"focused {format_name} {apply_method} exposes raw bytes: {owner_path}"
            )
        raw_id = IWA_CHART_DATA_RAW_ID_PARAMETER.search(apply_declaration)
        if raw_id is not None:
            violations.append(
                f"focused {format_name} {apply_method} exposes a raw ID parameter "
                f"{raw_id.group(0).strip()}: {owner_path}"
            )

    if re.search(r"\bChartData\b", code) is None:
        violations.append(
            f"focused {format_name} chart-data transaction is missing common ChartData: "
            f"{owner_path}"
        )
    if IWA_CHART_DATA_CODEC_IMPORT.search(code) is None:
        violations.append(
            f"focused {format_name} chart-data transaction is missing shared codec import: "
            f"{owner_path}"
        )
    if re.search(r"\bprepare_chart_data_rewrite\b", code) is None:
        violations.append(
            f"focused {format_name} chart-data transaction must prepare the shared "
            f"source-preserving rewrite: {owner_path}"
        )

    return sorted(set(violations))


def audit_iwa_chart_data_write_source_topology(root: Path = ROOT) -> list[str]:
    """Keep focused numeric chart-data writes selector-first and source-preserving."""

    if not _chart_data_boundary_active(root):
        return []

    violations = _chart_data_check_codec_rewrite_source(root)
    for format_name, owner_path in IWA_CHART_DATA_FOCUSED_OWNERS.items():
        violations.extend(
            _audit_focused_chart_data_write_owner(
                root,
                format_name=format_name,
                owner_path=owner_path,
                edit_method=IWA_CHART_DATA_FOCUSED_EDIT_METHODS[format_name],
                apply_method=IWA_CHART_DATA_FOCUSED_APPLY_METHODS[format_name],
            )
        )
    return sorted(set(violations))


def _chart_grid_creation_marker_present(code: str, marker: str | tuple[str, ...]) -> bool:
    """Match an identifier marker or a punctuation-bearing Rust marker."""

    if isinstance(marker, tuple):
        return any(_chart_grid_creation_marker_present(code, item) for item in marker)
    if not marker:
        return True
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", marker) is None:
        return marker in code
    return re.search(rf"\b{re.escape(marker)}\b", code) is not None


def _chart_grid_creation_marker_group_present(
    code: str, markers: tuple[str | tuple[str, ...], ...]
) -> bool:
    """Match required markers with nested tuples representing alternatives."""

    if not markers:
        return True
    if all(isinstance(marker, tuple) for marker in markers):
        return any(
            all(_chart_grid_creation_marker_present(code, item) for item in alternative)
            for alternative in markers
        )
    return all(_chart_grid_creation_marker_present(code, marker) for marker in markers)


def _chart_grid_creation_check_codec_source(root: Path) -> list[str]:
    """Require a private, bounded Buffa encoder for fresh chart grids."""

    codec_path = root / IWA_CHART_GRID_CREATION_CODEC_SOURCE
    if not codec_path.is_file():
        return [
            "neutral chart-grid creation codec source is missing: "
            f"{IWA_CHART_GRID_CREATION_CODEC_SOURCE}"
        ]

    raw_source = codec_path.read_text(encoding="utf-8")
    production_source = _mask_rust_cfg_test_items(raw_source)
    code = _mask_rust_non_code(production_source)
    violations: list[str] = []

    for name in IWA_CHART_GRID_CREATION_CODEC_REQUIRED_TYPES:
        if re.search(
            rf"\bpub[ \t]+"
            rf"(?:struct|enum|type)\s+{re.escape(name)}\b",
            code,
        ) is None:
            violations.append(
                "neutral chart-grid creation codec is missing strict API "
                f"{name}: {IWA_CHART_GRID_CREATION_CODEC_SOURCE}"
            )

    for name in IWA_CHART_GRID_CREATION_CODEC_REQUIRED_FUNCTIONS:
        if re.search(
            rf"\bpub[ \t]+fn\s+"
            rf"{re.escape(name)}\b",
            code,
        ) is None:
            violations.append(
                "neutral chart-grid creation codec is missing strict API "
                f"{name}: {IWA_CHART_GRID_CREATION_CODEC_SOURCE}"
            )

    for label, markers in IWA_CHART_GRID_CREATION_CODEC_REQUIRED_MARKER_GROUPS.items():
        if not _chart_grid_creation_marker_group_present(code, markers):
            violations.append(
                "neutral chart-grid creation codec is missing "
                f"{label} marker: {IWA_CHART_GRID_CREATION_CODEC_SOURCE}"
            )

    if IWA_CHART_GRID_CREATION_CODEC_PROST.search(code) is not None:
        violations.append(
            "neutral chart-grid creation codec must use Buffa rather than Prost: "
            f"{IWA_CHART_GRID_CREATION_CODEC_SOURCE}"
        )

    for pattern, message in (
        (
            IWA_CHART_GRID_CREATION_CODEC_GENERATED_CONSTRUCTION,
            "generated ChartGridArchive/GridRow/GridValue construction",
        ),
        (
            IWA_CHART_GRID_CREATION_CODEC_GENERATED_OPERATION,
            "generated chart-grid encode/decode operation",
        ),
        (
            IWA_CHART_GRID_CREATION_CODEC_EAGER_GENERATED_VECTOR,
            "eager vector of generated chart-grid values",
        ),
    ):
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            violations.append(
                "neutral chart-grid creation codec retains "
                f"{message}: {IWA_CHART_GRID_CREATION_CODEC_SOURCE}:{line_number}"
            )

    for match in IWA_CHART_GRID_CREATION_CODEC_GENERATED_REFERENCE.finditer(code):
        line_number = code.count("\n", 0, match.start()) + 1
        violations.append(
            "neutral chart-grid creation codec retains a generated chart-grid "
            f"type reference {match.group(0)}: "
            f"{IWA_CHART_GRID_CREATION_CODEC_SOURCE}:{line_number}"
        )

    for declaration, line_number in _rust_public_declarations(production_source):
        if re.search(
            r"\b(?:ChartGridArchive|GridRow|GridValue|"
            r"ChartGridRowColumnIdMap)\b|"
            r"\b(?:projection|buffa_[A-Za-z0-9_]*generated)\s*::",
            declaration,
        ):
            violations.append(
                "neutral chart-grid creation codec must not expose generated "
                f"types: {IWA_CHART_GRID_CREATION_CODEC_SOURCE}:{line_number}"
            )

    public_path = root / IWA_CHART_GRID_CREATION_CODEC_PUBLIC_SOURCE
    if not public_path.is_file():
        violations.append(
            "neutral chart-grid creation codec public module source is missing: "
            f"{IWA_CHART_GRID_CREATION_CODEC_PUBLIC_SOURCE}"
        )
    else:
        public_source = _mask_rust_cfg_test_items(
            public_path.read_text(encoding="utf-8")
        )
        if IWA_CHART_GRID_CREATION_CODEC_HIDDEN_MODULE.search(public_source) is None:
            violations.append(
                "neutral chart-grid creation codec module must be hidden: "
                f"{IWA_CHART_GRID_CREATION_CODEC_PUBLIC_SOURCE}"
            )

    return sorted(set(violations))


def audit_iwa_chart_grid_creation_source_topology(root: Path = ROOT) -> list[str]:
    """Keep fresh chart-grid authoring on the bounded shared wire seam."""

    if not (root / IWA_CHART_GRID_CREATION_CODEC_SOURCE).is_file():
        return []

    violations = _chart_grid_creation_check_codec_source(root)
    return sorted(set(violations))


def _audit_focused_chart_data_owner(
    root: Path,
    *,
    format_name: str,
    owner_path: Path,
    package_path: Path,
    module_name: str,
    public_method: str,
) -> list[str]:
    """Check one selector-first chart data package owner."""

    owner_absolute = root / owner_path
    package_absolute = root / package_path
    package_source = (
        package_absolute.read_text(encoding="utf-8")
        if package_absolute.is_file()
        else ""
    )
    package_code = _mask_rust_non_code(_mask_rust_cfg_test_items(package_source))
    module_match = re.search(
        rf"(?m)^[ \t]*(?P<visibility>pub(?:\([^()]*\))?[ \t]+)?"
        rf"mod[ \t]+(?:r#)?{re.escape(module_name)}\b[ \t]*(?:;|\{{)",
        package_code,
    )
    owner_source = (
        owner_absolute.read_text(encoding="utf-8")
        if owner_absolute.is_file()
        else package_source
    )
    owner_code = _mask_rust_non_code(_mask_rust_cfg_test_items(owner_source))
    method_evidence = _rust_public_method_bodies(owner_source, public_method)
    if not owner_absolute.is_file() and module_match is None and not method_evidence:
        return []

    violations: list[str] = []
    if not owner_absolute.is_file():
        violations.append(
            f"focused {format_name} chart-data owner is missing: {owner_path}"
        )
    if module_match is None:
        violations.append(
            f"focused {format_name} chart-data package is missing its private "
            f"module: {package_path}"
        )
    elif module_match.group("visibility") and not module_match.group("visibility").startswith(
        "pub("
    ):
        line_number = package_code.count("\n", 0, module_match.start()) + 1
        violations.append(
            f"focused {format_name} chart-data module must remain private: "
            f"{package_path}:{line_number}"
        )

    if IWA_CHART_DATA_COMMON_IMPORT.search(owner_code) is None:
        violations.append(
            f"focused {format_name} chart-data owner is missing its common "
            f"ChartData import: {owner_path}"
        )
    if IWA_CHART_DATA_CODEC_IMPORT.search(owner_code) is None:
        violations.append(
            f"focused {format_name} chart-data owner must delegate to the shared "
            f"codec: {owner_path}"
        )
    if IWA_CHART_DATA_FOCUSED_FORBIDDEN_PROST.search(owner_code) is not None:
        violations.append(
            f"focused {format_name} chart-data owner retains production Prost or "
            f"generated grid decoding: {owner_path}"
        )

    relative_owner = owner_path if owner_absolute.is_file() else package_path
    for declaration, line_number in _rust_public_declarations(
        _mask_rust_cfg_test_items(owner_source)
    ):
        if not _chart_metadata_external_public_declaration(declaration):
            continue
        physical = IWA_CHART_DATA_PUBLIC_PHYSICAL_TYPES.search(declaration)
        if physical is not None:
            physical_line = line_number + declaration.count(
                "\n", 0, physical.start()
            )
            violations.append(
                f"focused {format_name} chart-data public API exposes generated or "
                f"wire type {physical.group()}: {relative_owner}:{physical_line}"
            )
        if RUST_BYTE_SLICE.search(declaration) is not None:
            violations.append(
                f"focused {format_name} chart-data public API exposes raw bytes: "
                f"{relative_owner}:{line_number}"
            )
        raw_id = IWA_CHART_DATA_RAW_ID_PARAMETER.search(declaration)
        if raw_id is not None:
            raw_id_line = line_number + declaration.count("\n", 0, raw_id.start())
            violations.append(
                f"focused {format_name} chart-data public API exposes a raw ID "
                f"parameter {raw_id.group(0).strip()}: {relative_owner}:{raw_id_line}"
            )

    if not method_evidence:
        violations.append(
            f"focused {format_name} chart-data owner is missing public "
            f"Package::{public_method}: {owner_path}"
        )
    else:
        for declaration, _body, line_number in method_evidence:
            if re.search(r"\bChartData\b", declaration) is None:
                violations.append(
                    f"focused {format_name} Package::{public_method} must return common "
                    f"ChartData values: {owner_path}:{line_number}"
                )
            if re.search(r"\b[A-Za-z_][A-Za-z0-9_]*Selector\b", declaration) is None:
                violations.append(
                    f"focused {format_name} Package::{public_method} must accept a "
                    f"semantic selector: {owner_path}:{line_number}"
                )
            if RUST_BYTE_SLICE.search(declaration) is not None:
                violations.append(
                    f"focused {format_name} Package::{public_method} exposes raw bytes: "
                    f"{owner_path}:{line_number}"
                )
            raw_id = IWA_CHART_DATA_RAW_ID_PARAMETER.search(declaration)
            if raw_id is not None:
                violations.append(
                    f"focused {format_name} Package::{public_method} exposes a raw ID "
                    f"parameter {raw_id.group(0).strip()}: {owner_path}:{line_number}"
                )

    return sorted(set(violations))


def audit_iwa_chart_data_source_topology(root: Path = ROOT) -> list[str]:
    """Keep chart data common, private-wire, and focused seams separated."""

    if not _chart_data_boundary_active(root):
        return []

    violations = _chart_data_check_common_source(root)
    violations.extend(_chart_data_check_codec_source(root))
    for format_name, owner_path in IWA_CHART_DATA_FOCUSED_OWNERS.items():
        violations.extend(
            _audit_focused_chart_data_owner(
                root,
                format_name=format_name,
                owner_path=owner_path,
                package_path=IWA_CHART_DATA_FOCUSED_PACKAGES[format_name],
                module_name=IWA_CHART_DATA_FOCUSED_MODULES[format_name],
                public_method=IWA_CHART_DATA_FOCUSED_METHODS[format_name],
            )
        )

    return sorted(set(violations))


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
    if not items:
        lines.append("  none")
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


PAGES_TRANSACTION_REEXPORT = re.compile(
    r"(?<![A-Za-z0-9_])pub[ \t\r\n]+use[ \t\r\n]+(?:crate[ \t\r\n]*::[ \t\r\n]*)?"
    r"package[ \t\r\n]*::[ \t\r\n]*(?:r#)?(?P<owner>[A-Za-z_][A-Za-z0-9_]*)"
    r"[ \t\r\n]*::[ \t\r\n]*(?P<items>\{[^{}]*\}|\*|(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
    r"(?:[ \t\r\n]+as[ \t\r\n]+(?:r#)?[A-Za-z_][A-Za-z0-9_]*)?)[ \t\r\n]*;"
)
PAGES_TRANSACTION_SHARED_TYPE_REEXPORTS = {
    "Diagnostics": re.compile(
        r"(?m)^[ \t]*pub[ \t]+use[ \t]+litchi_core[ \t]*::[ \t]*"
        r"edit[ \t]*::[ \t]*Diagnostics[ \t]*;"
    ),
}
PAGES_TRANSACTION_SHARED_TYPE_ALIASES = {
    "Diagnostics": re.compile(
        r"(?ms)\bpub\s+use\s+litchi_core\s*::\s*edit\s*::\s*"
        r"(?:Diagnostics\s+as\s+(?P<direct>[A-Za-z_][A-Za-z0-9_]*)"
        r"|\{\s*[^;{}]*?\bDiagnostics\s+as\s+"
        r"(?P<group>[A-Za-z_][A-Za-z0-9_]*)[^;{}]*\})\s*;"
    ),
}
PAGES_TRANSACTION_SHARED_TYPE_SOURCES = {
    "Diagnostics": Path("crates/litchi-core/src/edit.rs"),
}
PAGES_PUBLIC_USE_STATEMENT = re.compile(
    r"(?<![A-Za-z0-9_])pub[ \t\r\n]+use\b[^;]*;"
)
PAGES_NOMINAL_DEFINITION = re.compile(
    r"(?<![A-Za-z0-9_#])(?:struct|enum|union|trait|type)[ \t\r\n]+"
    r"(?:r#)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
)


def _pages_shared_type_violations(
    root: Path, name: str, title: str
) -> list[str]:
    """Validate one shared canonical type used by a Pages family."""

    source_path = root / PAGES_TRANSACTION_SHARED_TYPE_SOURCES[name]
    if not source_path.is_file():
        return [
            f"{title} shared canonical type {name} source is missing: "
            f"{PAGES_TRANSACTION_SHARED_TYPE_SOURCES[name]}"
        ]
    source = _mask_rust_cfg_test_items(source_path.read_text(encoding="utf-8"))
    code = _mask_rust_non_code(source)
    if re.search(
        rf"(?m)^\s*pub\s+struct\s+{re.escape(name)}\b", code
    ) is None:
        return [
            f"{title} shared canonical type {name} is not a public struct: "
            f"{PAGES_TRANSACTION_SHARED_TYPE_SOURCES[name]}"
        ]
    if name == "Diagnostics" and re.search(
        r"\bpub\s+const\s+fn\s+changed\s*\(\s*&\s*self\s*\)", code
    ) is None:
        return [
            f"{title} shared Diagnostics is missing its public changed accessor: "
            f"{PAGES_TRANSACTION_SHARED_TYPE_SOURCES[name]}"
        ]
    return []


_PAGES_MASKED_SOURCE_CACHE: dict[tuple[str, int, int], str] = {}


def _pages_masked_production_code(path: Path) -> str:
    """Return cfg(test)- and comment-masked code, cached by file identity."""

    stat = path.stat()
    key = (str(path), stat.st_mtime_ns, stat.st_size)
    cached = _PAGES_MASKED_SOURCE_CACHE.get(key)
    if cached is None:
        cached = _mask_rust_non_code(
            _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        )
        _PAGES_MASKED_SOURCE_CACHE[key] = cached
    return cached


def _pages_retired_name_violations(
    root: Path, names: frozenset[str], title: str
) -> list[str]:
    """Reject a definition or ``pub use`` of a retired Pages spelling.

    Any nominal definition (``struct``/``enum``/``union``/``trait``/``type``
    at any visibility) is reported, as is any exactly-``pub`` ``use`` that
    names the spelling.  Private and ``pub(crate)`` imports such as
    ``use super::body_chart_arrangement::Error as BodyChartArrangementError``
    are implementation detail and remain allowed.
    """

    source_root = root / PAGES_SOURCE_ROOT
    if not names or not source_root.is_dir():
        return []
    violations: list[str] = []
    for path in sorted(source_root.rglob("*.rs")):
        code = _pages_masked_production_code(path)
        if not any(name in code for name in names):
            continue
        relative = path.relative_to(root)
        for match in PAGES_NOMINAL_DEFINITION.finditer(code):
            if match.group("name") in names:
                line_number = code.count("\n", 0, match.start()) + 1
                violations.append(
                    f"{title} defines retired name {match.group('name')}: "
                    f"{relative}:{line_number}"
                )
        for statement in PAGES_PUBLIC_USE_STATEMENT.finditer(code):
            identifiers = {
                match.group(1) for match in RUST_IDENTIFIER.finditer(statement.group(0))
            }
            for name in sorted(identifiers & names):
                line_number = code.count("\n", 0, statement.start()) + 1
                violations.append(
                    f"{title} publicly re-exports retired name {name}: "
                    f"{relative}:{line_number}"
                )
    return violations


def _pages_reexport_items(items: str) -> tuple[frozenset[str], tuple[str, ...], bool]:
    """Split one ``use`` item list into public names, aliases and a glob flag."""

    names: set[str] = set()
    aliases: list[str] = []
    glob = False
    for item in items.strip().strip("{}").split(","):
        tokens = item.replace("r#", "").split()
        if not tokens:
            continue
        if tokens == ["*"]:
            glob = True
            continue
        if len(tokens) == 3 and tokens[1] == "as":
            if tokens[0] != tokens[2]:
                aliases.append(f"{tokens[0]} as {tokens[2]}")
            names.add(tokens[2])
            continue
        names.add(tokens[0])
    return frozenset(names), tuple(aliases), glob


def _pages_public_module_scope(
    root: Path, segments: tuple[str, ...]
) -> tuple[Path | None, str | None, list[str]]:
    """Resolve one ``litchi_pages::a::b`` public module to its source scope.

    Returns the file holding the innermost module, that module's source text
    (an inline body or a whole file) and any visibility violations on the
    way.  Every segment must be declared exactly ``pub mod``.
    """

    source_root = root / PAGES_SOURCE_ROOT
    path = source_root / "lib.rs"
    public_path = "::".join(segments)
    if not path.is_file():
        return None, None, [f"Pages crate root is missing: {PAGES_SOURCE_ROOT / 'lib.rs'}"]
    scope = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
    directory = source_root
    violations: list[str] = []
    for index, segment in enumerate(segments):
        prefix = "::".join(segments[: index + 1])
        declarations = _rust_root_level_module_declarations(scope, frozenset({segment}))
        if not declarations:
            return path, None, violations + [
                f"Pages transaction module {public_path} is missing public module "
                f"{prefix}: {path.relative_to(root)}"
            ]
        _, visibility, shape, body, line_number = declarations[0]
        if visibility is None or re.sub(r"\s+", "", visibility) != "pub":
            violations.append(
                f"Pages transaction module {prefix} must be declared `pub mod`: "
                f"{path.relative_to(root)}:{line_number}"
            )
        if shape == "inline":
            scope = body
        elif shape == "external":
            candidates = (directory / f"{segment}.rs", directory / segment / "mod.rs")
            found = next((candidate for candidate in candidates if candidate.is_file()), None)
            if found is None:
                return path, None, violations + [
                    f"Pages transaction module {prefix} has no source file: "
                    f"{path.relative_to(root)}:{line_number}"
                ]
            path = found
            scope = _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
        else:
            return path, None, violations + [
                f"Pages transaction module {prefix} has an invalid declaration: "
                f"{path.relative_to(root)}:{line_number}"
            ]
        directory = directory / segment
    return path, scope, violations


def _pages_transaction_family_violations(
    root: Path, owner: str, label: str | None = None
) -> list[str]:
    """Check one renamed transaction family's owner, definitions and public module.

    * ``package.rs`` declares the owner exactly ``pub(crate) mod <owner>;``
      (a ``pub mod`` owner or a private one is reported);
    * the owner (or a child source under ``package/<owner>/``) defines every
      canonical short name and the owner makes it reachable;
    * the family's public module re-exports exactly the canonical short
      names with one ``pub use crate::package::<owner>::{...};`` and no
      alias or glob.

    Alternate paths, legacy spellings and globs elsewhere in the crate are
    checked once by ``audit_pages_transaction_family_source_topology``.
    """

    _, _, segments, shorts, _ = PAGES_TRANSACTION_FAMILY_BY_OWNER[owner]
    title = label or f"Pages {owner} transaction family"
    public_path = "::".join(segments)
    expected = frozenset(shorts)
    violations: list[str] = []

    package_path = root / PAGES_PACKAGE_SOURCE
    if not package_path.is_file():
        return [f"{title} package source is missing: {PAGES_PACKAGE_SOURCE}"]
    package_source = _mask_rust_cfg_test_items(package_path.read_text(encoding="utf-8"))
    declarations = _rust_root_level_module_declarations(package_source, frozenset({owner}))
    if not declarations:
        violations.append(
            f"{title} owner module is missing `pub(crate) mod {owner};`: "
            f"{PAGES_PACKAGE_SOURCE}"
        )
    for _, visibility, shape, _, line_number in declarations:
        normalized = re.sub(r"\s+", "", visibility or "")
        if normalized != "pub(crate)":
            violations.append(
                f"{title} owner module must be declared `pub(crate) mod {owner};` "
                f"(found `{normalized or 'private'} mod`): "
                f"{PAGES_PACKAGE_SOURCE}:{line_number}"
            )
        if shape != "external":
            violations.append(
                f"{title} owner module must live in its own source file: "
                f"{PAGES_PACKAGE_SOURCE}:{line_number}"
            )

    owner_relative = PAGES_SOURCE_ROOT / "package" / f"{owner}.rs"
    owner_path = root / owner_relative
    if not owner_path.is_file():
        violations.append(f"{title} owner source is missing: {owner_relative}")
    else:
        owner_source = _mask_rust_cfg_test_items(owner_path.read_text(encoding="utf-8"))
        reachable = _rust_canonical_exports(owner_source, expected)
        shared_types = PAGES_TRANSACTION_SHARED_TYPES.get(owner, frozenset())
        defined: set[str] = set()
        family_sources = [owner_source]
        helper_root = root / PAGES_SOURCE_ROOT / "package" / owner
        if helper_root.is_dir():
            family_sources.extend(
                _mask_rust_cfg_test_items(path.read_text(encoding="utf-8"))
                for path in sorted(helper_root.rglob("*.rs"))
            )
        for family_source in family_sources:
            code = _mask_rust_non_code(family_source)
            for match in PAGES_NOMINAL_DEFINITION.finditer(code):
                if match.group("name") in expected:
                    defined.add(match.group("name"))
        for name in shorts:
            if name in shared_types:
                if name in defined:
                    violations.append(
                        f"{title} must not define shared canonical type {name}; "
                        f"re-export litchi_core::edit::{name}: {owner_relative}"
                    )
                reexport = PAGES_TRANSACTION_SHARED_TYPE_REEXPORTS.get(name)
                if reexport is None or reexport.search(owner_source) is None:
                    violations.append(
                        f"{title} must re-export shared litchi_core::edit::{name}: "
                        f"{owner_relative}"
                    )
                else:
                    alias_pattern = PAGES_TRANSACTION_SHARED_TYPE_ALIASES.get(name)
                    if alias_pattern is not None:
                        owner_code = _mask_rust_non_code(owner_source)
                        for alias_match in alias_pattern.finditer(owner_code):
                            alias = alias_match.group("direct") or alias_match.group("group")
                            violations.append(
                                f"{title} has alternate public alias {alias} for "
                                f"shared litchi_core::edit::{name}: {owner_relative}"
                            )
                    violations.extend(_pages_shared_type_violations(root, name, title))
                    defined.add(name)
                continue
            if name not in defined:
                violations.append(
                    f"{title} owner is missing canonical type {name}: {owner_relative}"
                )
            elif name not in reachable:
                violations.append(
                    f"{title} owner does not publish canonical type {name}: "
                    f"{owner_relative}"
                )

    violations.extend(
        _pages_retired_name_violations(root, _pages_transaction_legacy_names(owner), title)
    )

    public_file, scope, module_violations = _pages_public_module_scope(root, segments)
    violations.extend(f"{title}: {item}" for item in module_violations)
    if public_file is None or scope is None:
        return violations
    public_relative = public_file.relative_to(root)
    code = _mask_rust_non_code(scope)
    exported: set[str] = set()
    statements = 0
    for match in PAGES_TRANSACTION_REEXPORT.finditer(code):
        if match.group("owner") != owner or _rust_brace_depth_at(code, match.start()):
            continue
        statements += 1
        names, aliases, glob = _pages_reexport_items(match.group("items"))
        exported.update(names)
        for alias in aliases:
            violations.append(
                f"{title} public module {public_path} re-exports alias {alias}: "
                f"{public_relative}"
            )
        if glob:
            violations.append(
                f"{title} public module {public_path} glob re-exports its owner: "
                f"{public_relative}"
            )
    if not statements:
        violations.append(
            f"{title} public module {public_path} is missing "
            f"`pub use crate::package::{owner}::{{...}};`: {public_relative}"
        )
    elif statements > 1:
        violations.append(
            f"{title} public module {public_path} re-exports its owner more than once: "
            f"{public_relative}"
        )
    for name in sorted(expected - exported):
        if statements:
            violations.append(
                f"{title} public module {public_path} is missing canonical export "
                f"{name}: {public_relative}"
            )
    for name in sorted(exported - expected - PAGES_TRANSACTION_OPTIONAL_EXPORTS):
        violations.append(
            f"{title} public module {public_path} exports non-canonical {name}: "
            f"{public_relative}"
        )
    return violations
