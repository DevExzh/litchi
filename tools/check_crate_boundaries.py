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
IWA_KEYNOTE_EDITOR_SOURCE = IWA_KEYNOTE_SOURCE_ROOT / "editor.rs"
RETIRED_IWA_KEYNOTE_METHODS = (
    "set_slide_name",
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
RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_METHODS = (
    "show_settings",
    "set_show_settings",
)
RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_METHODS
)
RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_SOURCE = (
    IWA_KEYNOTE_SOURCE_ROOT / "editor" / "show_settings.rs"
)
RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_MODULES = ("show_settings",)
RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_EXAMPLE = Path(
    "crates/litchi-iwa/examples/edit_keynote_show.rs"
)
IWA_KEYNOTE_README = Path("crates/litchi-iwa/README.md")
RETIRED_IWA_KEYNOTE_SLIDE_NAME_EXAMPLE = Path(
    "crates/litchi-iwa/examples/rename_keynote_slide.rs"
)
IWA_KEYNOTE_README_SLIDE_NAME_CALL = re.compile(
    r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
    r"[ \t\r\n]*\.[ \t\r\n]*(?P<method>set_slide_name)"
    r"\b[ \t\r\n]*\("
)
RETIRED_IWA_KEYNOTE_DOCUMENT_SOURCE = IWA_KEYNOTE_SOURCE_ROOT / "document.rs"
RETIRED_IWA_KEYNOTE_DOCUMENT_TYPES = (
    "KeynoteDocument",
    "KeynoteDocumentState",
    "KeynoteDocumentStats",
)
RETIRED_IWA_KEYNOTE_DOCUMENT_TYPE_SET = frozenset(
    RETIRED_IWA_KEYNOTE_DOCUMENT_TYPES
)
IWA_KEYNOTE_MODULE_SOURCE = IWA_KEYNOTE_SOURCE_ROOT / "mod.rs"
IWA_KEYNOTE_DOCUMENT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(document)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_KEYNOTE_DOCUMENT_LOCAL_REEXPORT = re.compile(
    r"^[ \t]*pub(?:\([^()]*\))?[ \t\r\n]+use[ \t\r\n]+"
    r"(?:(?:r#)?self[ \t\r\n]*::[ \t\r\n]*)?"
    r"(?:r#)?(?P<module>document)\b",
    re.MULTILINE,
)
IWA_KEYNOTE_DOCUMENT_CALLER_ROOTS = (
    Path("crates/litchi-iwa/src"),
    Path("crates/litchi-iwa/tests"),
    Path("crates/litchi-iwa/examples"),
)
KEYNOTE_SOURCE_ROOT = Path("crates/litchi-keynote/src")
KEYNOTE_PACKAGE_MANIFEST = Path("crates/litchi-keynote/Cargo.toml")
# `perf_tests.rs` is included only from a `#[cfg(test)]` module in its parent
# source file. Keep the production audit from treating that test-only module
# body as a reachable crate item when walking the source tree.
KEYNOTE_TEST_ONLY_SOURCE_NAMES = frozenset({"perf_tests.rs"})
KEYNOTE_PRODUCTION_TEST_MODULE = re.compile(
    r"^[ \t]*#[ \t]*\[[ \t]*cfg[ \t]*\([ \t]*test[ \t]*\)[ \t]*\]"
    r"[ \t\r\n]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?mod\b",
    re.MULTILINE,
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
IWA_KEYNOTE_SHOW_SETTINGS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(show_settings)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_KEYNOTE_README_SHOW_SETTINGS_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:keynote|editor)[ \t\r\n]*\."
        r"[ \t\r\n]*(?:r#)?(?P<method>show_settings|set_show_settings)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?KeynoteEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>show_settings|set_show_settings)\b"
        r"[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_show_settings)\b[ \t\r\n]*\("
    ),
)
KEYNOTE_SHOW_OWNER_PATH = re.compile(
    r"(?<![A-Za-z0-9_#])(?:r#)?(?:show|show_settings)"
    r"[ \t\r\n]*::"
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_METHODS = (
    "soundtrack_settings",
    "set_soundtrack_settings",
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_METHODS
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_SOURCE = (
    IWA_KEYNOTE_SOURCE_ROOT / "editor" / "soundtrack.rs"
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_MODULES = ("soundtrack",)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_EXAMPLE = Path(
    "crates/litchi-iwa/examples/edit_keynote_soundtrack.rs"
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_TESTS = (
    "soundtrack_settings_are_typed_transactional_and_wire_exact",
    "soundtrack_settings_handle_absent_and_malformed_objects_transactionally",
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_TEST_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_TESTS
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_METHODS = (
    "patch_soundtrack_wire",
)
RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_METHODS
)
IWA_KEYNOTE_SOUNDTRACK_WIRE_SOURCE = (
    IWA_KEYNOTE_SOURCE_ROOT / "editor" / "soundtrack_wire.rs"
)
IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLER_SOURCES = (
    Path("crates/litchi-iwa/examples/inspect_keynote_structure.rs"),
)
IWA_KEYNOTE_SOUNDTRACK_SETTINGS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(soundtrack)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:keynote|editor)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>soundtrack_settings|set_soundtrack_settings)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?KeynoteEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>soundtrack_settings|set_soundtrack_settings)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_soundtrack_settings)\b[ \t\r\n]*\("
    ),
)
IWA_KEYNOTE_README_SOUNDTRACK_SETTINGS_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>edit_keynote_soundtrack)"
    r"(?:\.rs)?(?![A-Za-z0-9_])"
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
        "IWorkPackage",
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
RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_METHODS = (
    "slide_transition",
    "set_slide_transition",
    "clear_slide_transition",
)
RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_METHODS
)
RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_SOURCES = (
    IWA_KEYNOTE_SOURCE_ROOT / "editor" / "transition_lifecycle.rs",
)
RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_MODULES = ("transition_lifecycle",)
RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_EXAMPLES = (
    Path("crates/litchi-iwa/examples/clear_keynote_transition.rs"),
    Path("crates/litchi-iwa/examples/edit_keynote_transition.rs"),
    Path("crates/litchi-iwa/examples/set_keynote_transition_effect.rs"),
)
IWA_KEYNOTE_SLIDE_TRANSITION_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(transition_lifecycle)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_KEYNOTE_README_SLIDE_TRANSITION_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:keynote|editor|reopened)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>slide_transition|set_slide_transition|clear_slide_transition)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?KeynoteEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>slide_transition|set_slide_transition|clear_slide_transition)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_slide_transition|clear_slide_transition)"
        r"\b[ \t\r\n]*\("
    ),
)
IWA_KEYNOTE_README_SLIDE_TRANSITION_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>"
    r"clear_keynote_transition|edit_keynote_transition|set_keynote_transition_effect"
    r")(?:\.rs)?(?![A-Za-z0-9_])"
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
        "IWorkPackage",
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
RETIRED_IWA_KEYNOTE_SLIDE_DELETE_METHODS = ("remove_slide",)
RETIRED_IWA_KEYNOTE_SLIDE_DELETE_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SLIDE_DELETE_METHODS
)
RETIRED_IWA_KEYNOTE_SLIDE_DELETE_SOURCE = (
    IWA_KEYNOTE_SOURCE_ROOT / "editor" / "slide_delete.rs"
)
RETIRED_IWA_KEYNOTE_SLIDE_DELETE_MODULES = ("slide_delete",)
RETIRED_IWA_KEYNOTE_SLIDE_DELETE_EXAMPLE = Path(
    "crates/litchi-iwa/examples/remove_keynote_slide.rs"
)
IWA_KEYNOTE_SLIDE_DELETE_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(slide_delete)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_KEYNOTE_README_SLIDE_DELETE_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:keynote|editor|reopened)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>remove_slide)\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?KeynoteEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>remove_slide)\b[ \t\r\n]*\("
    ),
)
IWA_KEYNOTE_README_SLIDE_DELETE_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>remove_keynote_slide)"
    r"(?:\.rs)?(?![A-Za-z0-9_])"
)
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
        "IWorkPackage",
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
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_METHODS = (
    "set_slide_text_placeholder_visible",
    "set_slide_title_visible",
    "set_slide_body_visible",
)
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_METHODS
)
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_SOURCE = (
    IWA_KEYNOTE_SOURCE_ROOT / "editor" / "placeholder_visibility.rs"
)
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_MODULES = (
    "placeholder_visibility",
)
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_EXAMPLE = Path(
    "crates/litchi-iwa/examples/set_keynote_placeholder_visibility.rs"
)
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_PUBLIC_TYPES = (
    "KeynoteSlideTextPlaceholder",
)
RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_PUBLIC_TYPE_SET = frozenset(
    RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_PUBLIC_TYPES
)
IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(placeholder_visibility)\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_KEYNOTE_README_PLACEHOLDER_VISIBILITY_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:keynote|editor)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_slide_text_placeholder_visible|"
        r"set_slide_title_visible|set_slide_body_visible)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?KeynoteEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>set_slide_text_placeholder_visible|"
        r"set_slide_title_visible|set_slide_body_visible)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_slide_text_placeholder_visible|"
        r"set_slide_title_visible|set_slide_body_visible)"
        r"\b[ \t\r\n]*\("
    ),
)
IWA_KEYNOTE_README_PLACEHOLDER_VISIBILITY_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>set_keynote_placeholder_visibility)"
    r"(?:\.rs)?(?![A-Za-z0-9_])"
)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_METHODS = (
    "set_slide_number_visible",
)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_METHOD_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_METHODS
)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_SOURCE = (
    IWA_KEYNOTE_SOURCE_ROOT / "editor" / "slide_number.rs"
)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_MODULES = ("slide_number",)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_EXAMPLE = Path(
    "crates/litchi-iwa/examples/set_keynote_slide_number.rs"
)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_TESTS = (
    "slide_number_visibility_matches_native_ownership_and_round_trips_exactly",
    "slide_number_visibility_rejects_inconsistent_native_state_transactionally",
)
RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_TEST_SET = frozenset(
    RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_TESTS
)
IWA_KEYNOTE_EDITOR_TEST_SOURCE = IWA_KEYNOTE_SOURCE_ROOT / "editor" / "tests.rs"
IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(slide_number)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_KEYNOTE_README_SLIDE_NUMBER_VISIBILITY_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_slide_number_visible)\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?KeynoteEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>set_slide_number_visible)\b[ \t\r\n]*\("
    ),
)
IWA_KEYNOTE_README_SLIDE_NUMBER_VISIBILITY_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>set_keynote_slide_number)"
    r"(?:\.rs)?(?![A-Za-z0-9_])"
)
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
        "IWorkPackage",
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
IWA_NUMBERS_SOURCE_ROOT = Path("crates/litchi-iwa/src/numbers")
IWA_NUMBERS_SEMANTIC_WORKBOOK_SOURCE = (
    IWA_NUMBERS_SOURCE_ROOT / "editor" / "semantic" / "workbook.rs"
)
RETIRED_IWA_NUMBERS_NAMES_METHODS = ("rename_sheet", "rename_table")
RETIRED_IWA_NUMBERS_NAMES_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_NAMES_METHODS
)
RETIRED_IWA_NUMBERS_NAMES_EXAMPLE = Path(
    "crates/litchi-iwa/examples/rename_numbers_items.rs"
)
IWA_NUMBERS_README = Path("crates/litchi-iwa/README.md")
NUMBERS_NAMES_IMPLEMENTATION_SOURCES = (
    Path("crates/litchi-numbers/src/names.rs"),
    Path("crates/litchi-numbers/src/package/names.rs"),
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
IWA_NUMBERS_README_NAMES_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>rename_sheet|rename_table)\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?NumbersEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>rename_sheet|rename_table)\b"
        r"[ \t\r\n]*\("
    ),
)
IWA_NUMBERS_README_NAMES_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>rename_numbers_items)(?:\.rs)?"
    r"(?![A-Za-z0-9_])"
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
RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_METHODS = (
    "table_header_settings",
    "set_table_header_settings",
)
RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_METHODS
)
RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_EXAMPLE = Path(
    "crates/litchi-iwa/examples/edit_numbers_table_headers.rs"
)
IWA_NUMBERS_README_TABLE_HEADER_SETTINGS_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>table_header_settings|set_table_header_settings)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?NumbersEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>table_header_settings|set_table_header_settings)"
        r"\b[ \t\r\n]*\("
    ),
)
IWA_NUMBERS_README_TABLE_HEADER_SETTINGS_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>edit_numbers_table_headers)(?:\.rs)?"
    r"(?![A-Za-z0-9_])"
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
        "IWorkPackage",
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
RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_METHODS = (
    "table_title_settings",
    "set_table_title_settings",
)
RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_METHODS
)
RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_EXAMPLE = Path(
    "crates/litchi-iwa/examples/edit_numbers_table_title.rs"
)
RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_TESTS = (
    "table_title_settings_are_lossless_transactional_and_wire_exact",
    "table_title_settings_restore_native_presence_exactly",
    "table_title_settings_reject_missing_render_styles_transactionally",
    "table_title_settings_reject_malformed_wire_transactionally",
)
RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_TEST_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_TESTS
)
IWA_NUMBERS_README_TABLE_TITLE_SETTINGS_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>table_title_settings|set_table_title_settings)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?NumbersEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>table_title_settings|set_table_title_settings)"
        r"\b[ \t\r\n]*\("
    ),
)
IWA_NUMBERS_README_TABLE_TITLE_SETTINGS_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>edit_numbers_table_title)(?:\.rs)?"
    r"(?![A-Za-z0-9_])"
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
        "IWorkPackage",
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
RETIRED_IWA_NUMBERS_TABLE_DIMENSION_METHODS = (
    "table_dimension_size",
    "set_table_dimension_size",
    "table_row_height",
    "set_table_row_height",
    "table_column_width",
    "set_table_column_width",
)
RETIRED_IWA_NUMBERS_TABLE_DIMENSION_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_DIMENSION_METHODS
)
RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPES = (
    "Dimension",
    "Points",
    "Size",
)
RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPE_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPES
)
IWA_NUMBERS_PRIVATE_TABLE_DIMENSION_ALIASES = frozenset(
    {
        "NumbersTableDimension",
        "NumbersTablePoints",
        "NumbersTableDimensionSize",
    }
)
RETIRED_IWA_NUMBERS_TABLE_DIMENSION_EXAMPLE = Path(
    "crates/litchi-iwa/examples/edit_numbers_table_dimension.rs"
)
RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TESTS = (
    "table_dimension_sizes_are_typed_transactional_and_wire_exact",
    "table_dimension_size_preserves_unknown_header_fields",
    "table_dimension_size_rejects_malformed_headers_transactionally",
)
RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TEST_SET = frozenset(
    RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TESTS
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
    "TransactionError",
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
        "IWorkPackage",
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
RETIRED_IWA_NUMBERS_TABLE_CELL_EDITOR_SOURCE = Path(
    "crates/litchi-iwa/src/numbers/editor/semantic/table.rs"
)
RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_SOURCE = Path(
    "crates/litchi-iwa/src/numbers/editor/model.rs"
)
RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_SOURCE = Path(
    "crates/litchi-iwa/src/numbers/editor/table_cells.rs"
)
RETIRED_IWA_NUMBERS_TABLE_CELL_TEST_SOURCE = Path(
    "crates/litchi-iwa/src/numbers/editor/tests.rs"
)
RETIRED_IWA_NUMBERS_TABLE_CELL_EXAMPLE = Path(
    "crates/litchi-iwa/examples/edit_numbers_cell.rs"
)
RETIRED_IWA_NUMBERS_TABLE_CELL_SOURCE_INVENTORY = (
    RETIRED_IWA_NUMBERS_TABLE_CELL_EDITOR_SOURCE,
    RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_SOURCE,
    RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_SOURCE,
    RETIRED_IWA_NUMBERS_TABLE_CELL_TEST_SOURCE,
    RETIRED_IWA_NUMBERS_TABLE_CELL_EXAMPLE,
)
RETIRED_IWA_NUMBERS_TABLE_CELL_METHODS = (
    "set_cell",
    "set_cells",
    "clear_cell",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_HELPERS = (
    "set_cell_in_package",
    "set_cells_in_package",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_HELPERS = (
    "apply_numbers",
)
RETIRED_IWA_NUMBERS_TABLE_CELL_TESTS = (
    "semantic_edits_round_trip_through_public_reader",
    "cell_batch_roundtrips_mixed_values_and_clear",
    "cell_batch_refreshes_formula_chain_from_final_state",
    "cell_batch_rejects_invalid_inputs_transactionally",
    "failed_edit_is_transactional",
    "cell_edits_keep_sparse_row_headers_in_lockstep",
    "source_created_large_table_allocates_sparse_tiles_for_batch_writes",
    "rich_text_cell_updates_preserve_the_payload_reference",
    "shared_rich_text_cell_update_uses_copy_on_write",
    "replacing_rich_text_releases_list_and_payload_objects",
    "segmented_string_entries_round_trip_and_remain_interned",
    "segmented_shared_rich_text_uses_copy_on_write_and_cleans_up",
    "formula_cells_can_be_cleared_with_refcount_cleanup",
    "cell_write_refreshes_transitive_formula_caches_in_dependency_order",
    "cell_write_rejects_unsupported_impacted_formula_transactionally",
)
IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_SOURCE = Path(
    "crates/litchi-iwa/src/numbers/editor.rs"
)
IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_HELPERS = (
    "test_set_cell",
    "test_set_cells",
    "test_clear_cell",
)
IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_HELPER_SET = frozenset(
    IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_HELPERS
)
IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_DECLARATION = re.compile(
    r"#[ \t\r\n]*\[[ \t\r\n]*cfg[ \t\r\n]*\([ \t\r\n]*test"
    r"[ \t\r\n]*\)[ \t\r\n]*\][ \t\r\n]*pub[ \t\r\n]*\("
    r"[ \t\r\n]*crate[ \t\r\n]*\)[ \t\r\n]+fn[ \t\r\n]+"
    r"(?:r#)?(?P<name>test_(?:set_cells?|clear_cell))\b"
)
IWA_NUMBERS_EXAMPLE_ROOT = Path("crates/litchi-iwa/examples")
IWA_TABLE_LOCK_SOURCE = Path("crates/litchi-iwa/src/table_lock.rs")
IWA_NUMBERS_TABLE_INFO_SOURCE = (
    IWA_NUMBERS_SOURCE_ROOT / "editor" / "semantic" / "model.rs"
)
NUMBERS_SOURCE_ROOT = Path("crates/litchi-numbers/src")
NUMBERS_PACKAGE_SOURCE = NUMBERS_SOURCE_ROOT / "package.rs"
NUMBERS_PACKAGE_TEST_MODULE = re.compile(
    r"^[ \t]*#[ \t]*\[[ \t]*cfg[ \t]*\([ \t]*test[ \t]*\)[ \t]*\]",
    re.MULTILINE,
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
RETIRED_IWA_NUMBERS_SHEET_ORDER_METHODS = ("move_sheet",)
RETIRED_IWA_NUMBERS_SHEET_ORDER_METHOD_SET = frozenset(
    RETIRED_IWA_NUMBERS_SHEET_ORDER_METHODS
)
RETIRED_IWA_NUMBERS_SHEET_ORDER_EXAMPLE = Path(
    "crates/litchi-iwa/examples/move_numbers_sheet.rs"
)
RETIRED_IWA_NUMBERS_SHEET_ORDER_TESTS = (
    "reorders_and_removes_sheets_transactionally",
    "sheet_list_crud_preserves_raw_references_and_restores_exact_component",
    "duplicate_sheet_references_fail_transactionally",
)
RETIRED_IWA_NUMBERS_SHEET_ORDER_TEST_SET = frozenset(
    RETIRED_IWA_NUMBERS_SHEET_ORDER_TESTS
)
IWA_NUMBERS_EDITOR_TEST_SOURCE = IWA_NUMBERS_SOURCE_ROOT / "editor" / "tests.rs"
IWA_NUMBERS_README_SHEET_ORDER_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:numbers|numbers_editor)"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>move_sheet)\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?NumbersEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>move_sheet)\b[ \t\r\n]*\("
    ),
)
IWA_NUMBERS_README_SHEET_ORDER_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>move_numbers_sheet)(?:\.rs)?"
    r"(?![A-Za-z0-9_])"
)
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
        "IWorkPackage",
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
RETIRED_IWA_PAGES_DOCUMENT_SOURCE = IWA_PAGES_SOURCE_ROOT / "document.rs"
RETIRED_IWA_PAGES_DOCUMENT_TYPES = (
    "PagesDocument",
    "PagesDocumentState",
    "PagesDocumentStats",
)
RETIRED_IWA_PAGES_DOCUMENT_TYPE_SET = frozenset(
    RETIRED_IWA_PAGES_DOCUMENT_TYPES
)
IWA_PAGES_MODULE_SOURCE = IWA_PAGES_SOURCE_ROOT / "mod.rs"
IWA_PAGES_DOCUMENT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(document)\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_PAGES_DOCUMENT_LOCAL_REEXPORT = re.compile(
    r"^[ \t]*pub(?:\([^()]*\))?[ \t\r\n]+use[ \t\r\n]+"
    r"(?:(?:r#)?self[ \t\r\n]*::[ \t\r\n]*|"
    r"(?:r#)?crate[ \t\r\n]*::[ \t\r\n]*(?:r#)?pages"
    r"[ \t\r\n]*::[ \t\r\n]*(?:\{[ \t\r\n]*)?)?"
    r"(?:r#)?(?P<module>document)\b",
    re.MULTILINE,
)
WORKSPACE_CRATES_ROOT = Path("crates")
IWA_HOST_SOURCE_ROOT = Path("crates/litchi-iwa/src")
IWA_PAGES_FOCUSED_READER_TYPES = frozenset({"Document", "Package"})
IWA_NUMBERS_SOURCE_ROOT = Path("crates/litchi-iwa/src/numbers")
RETIRED_IWA_NUMBERS_DOCUMENT_SOURCE = IWA_NUMBERS_SOURCE_ROOT / "document.rs"
RETIRED_IWA_NUMBERS_DOCUMENT_TYPES = (
    "NumbersDocument",
    "NumbersDocumentState",
    "NumbersDocumentStats",
)
RETIRED_IWA_NUMBERS_DOCUMENT_TYPE_SET = frozenset(
    RETIRED_IWA_NUMBERS_DOCUMENT_TYPES
)
RETIRED_IWA_NUMBERS_SHEET_SOURCE = IWA_NUMBERS_SOURCE_ROOT / "sheet.rs"
RETIRED_IWA_NUMBERS_SHEET_TYPES = ("NumbersSheet",)
RETIRED_IWA_NUMBERS_READER_TYPE_SET = frozenset(
    (*RETIRED_IWA_NUMBERS_DOCUMENT_TYPES, *RETIRED_IWA_NUMBERS_SHEET_TYPES)
)
IWA_NUMBERS_MODULE_SOURCE = IWA_NUMBERS_SOURCE_ROOT / "mod.rs"
IWA_NUMBERS_READER_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(?P<module>document|sheet)\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_NUMBERS_READER_LOCAL_REEXPORT = re.compile(
    r"^[ \t]*pub(?:\([^()]*\))?[ \t\r\n]+use[ \t\r\n]+"
    r"(?:\{[ \t\r\n]*)?"
    r"(?:(?:r#)?self[ \t\r\n]*::[ \t\r\n]*|"
    r"(?:r#)?crate[ \t\r\n]*::[ \t\r\n]*(?:r#)?numbers"
    r"[ \t\r\n]*::[ \t\r\n]*(?:\{[ \t\r\n]*)?)?"
    r"(?:r#)?(?P<module>document|sheet)\b",
    re.MULTILINE,
)
IWA_NUMBERS_FOCUSED_READER_TYPES = frozenset({"Document", "Package"})
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
RETIRED_IWA_PAGES_PAGE_LAYOUT_SOURCE = IWA_PAGES_SOURCE_ROOT / "editor" / "page_layout.rs"
RETIRED_IWA_PAGES_PAGE_LAYOUT_METHODS = ("page_layout", "set_page_layout")
RETIRED_IWA_PAGES_PAGE_LAYOUT_METHOD_SET = frozenset(
    RETIRED_IWA_PAGES_PAGE_LAYOUT_METHODS
)
PAGES_SOURCE_ROOT = Path("crates/litchi-pages/src")
PAGES_DOCUMENT_PUBLIC_API_SOURCES = (
    PAGES_SOURCE_ROOT / "document.rs",
    PAGES_SOURCE_ROOT / "lib.rs",
)
PAGES_DOCUMENT_PUBLIC_MARKERS = frozenset(
    {
        "Body",
        "DEFAULT_MAX_TEXT_BYTES",
        "Document",
        "DocumentReadOptions",
        "DocumentSourceLimitKind",
        "DocumentSourceLimits",
        "DocumentSourceLimitsError",
        "Error",
        "IoKind",
        "MAX_BODY_STORAGES",
        "MAX_SECTIONS",
        "ReadError",
        "ReadLimitKind",
        "Result",
        "Root",
        "SemanticLimitKind",
        "SemanticLimits",
        "SemanticLimitsError",
    }
)
PAGES_PACKAGE_SOURCE = PAGES_SOURCE_ROOT / "package.rs"
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
CARGO_SECTION_HEADER = re.compile(r"^[ \t]*\[([^\]]+)\][ \t]*(?:#.*)?$")
CARGO_PROST_DEPENDENCY = re.compile(
    r"^[ \t]*(?:prost|\"prost\")(?:[ \t]*\.[ \t]*workspace)?[ \t]*="
)
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
RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_METHODS = (
    "document_options",
    "set_document_options",
    "footnote_settings",
    "set_footnote_settings",
)
RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_METHOD_SET = frozenset(
    RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_METHODS
)
RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_SOURCES = (
    IWA_PAGES_SOURCE_ROOT / "editor" / "document_options.rs",
    IWA_PAGES_SOURCE_ROOT / "editor" / "document_options" / "wire.rs",
    IWA_PAGES_SOURCE_ROOT / "editor" / "footnote_settings.rs",
)
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
IWA_PAGES_PAGE_LAYOUT_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?page_layout\b[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_PAGES_DOCUMENT_SETTINGS_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(document_options|footnote_settings)\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
RETIRED_IWA_PAGES_SECTION_SETTINGS_METHODS = (
    "section_settings",
    "set_section_settings",
    "set_section_name",
)
RETIRED_IWA_PAGES_SECTION_SETTINGS_METHOD_SET = frozenset(
    RETIRED_IWA_PAGES_SECTION_SETTINGS_METHODS
)
RETIRED_IWA_PAGES_SECTION_SETTINGS_EXAMPLE = Path(
    "crates/litchi-iwa/examples/set_pages_section_settings.rs"
)
RETIRED_IWA_PAGES_SECTION_SETTINGS_TESTS = (
    "section_settings_crud_is_lossless_validated_and_transactional",
    "section_settings_reject_zero_starting_page_number_transactionally",
)
RETIRED_IWA_PAGES_SECTION_SETTINGS_TEST_SET = frozenset(
    RETIRED_IWA_PAGES_SECTION_SETTINGS_TESTS
)
IWA_PAGES_EDITOR_TEST_SOURCE = IWA_PAGES_SOURCE_ROOT / "editor" / "tests.rs"
IWA_PAGES_README = Path("crates/litchi-iwa/README.md")
IWA_PAGES_README_SECTION_SETTINGS_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:pages|editor)[ \t\r\n]*\."
        r"[ \t\r\n]*(?:r#)?(?P<method>section_settings|set_section_settings|set_section_name)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?PagesEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>section_settings|set_section_settings|set_section_name)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_section_settings|set_section_name)\b[ \t\r\n]*\("
    ),
)
IWA_PAGES_README_SECTION_SETTINGS_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>set_pages_section_settings)"
    r"(?:\.rs)?(?![A-Za-z0-9_])"
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
        "IWorkPackage",
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
RETIRED_IWA_PAGES_SECTION_BACKGROUND_METHODS = (
    "section_background",
    "set_section_background",
)
RETIRED_IWA_PAGES_SECTION_BACKGROUND_METHOD_SET = frozenset(
    RETIRED_IWA_PAGES_SECTION_BACKGROUND_METHODS
)
RETIRED_IWA_PAGES_SECTION_BACKGROUND_SOURCES = (
    IWA_PAGES_SOURCE_ROOT / "editor" / "section_background.rs",
    IWA_PAGES_SOURCE_ROOT / "editor" / "section_settings.rs",
)
RETIRED_IWA_PAGES_SECTION_BACKGROUND_MODULES = (
    "section_background",
    "section_settings",
)
RETIRED_IWA_PAGES_SECTION_BACKGROUND_EXAMPLE = Path(
    "crates/litchi-iwa/examples/set_pages_section_background.rs"
)
RETIRED_IWA_PAGES_SECTION_BACKGROUND_TESTS = (
    "solid_section_background_crud_preserves_nested_unknown_wire",
)
RETIRED_IWA_PAGES_SECTION_BACKGROUND_TEST_SET = frozenset(
    RETIRED_IWA_PAGES_SECTION_BACKGROUND_TESTS
)
IWA_PAGES_SECTION_BACKGROUND_MODULE = re.compile(
    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
    r"mod[ \t\r\n]+(?:r#)?(section_background|section_settings)\b"
    r"[ \t\r\n]*(?:;|\{)",
    re.MULTILINE,
)
IWA_PAGES_README_SECTION_BACKGROUND_CALLS = (
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?(?:pages|editor)[ \t\r\n]*\."
        r"[ \t\r\n]*(?:r#)?(?P<method>section_background|set_section_background)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])"
        r"(?:(?:r#)?[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]*::[ \t\r\n]*)*"
        r"(?:r#)?PagesEditor[ \t\r\n]*::[ \t\r\n]*"
        r"(?:r#)?(?P<method>section_background|set_section_background)"
        r"\b[ \t\r\n]*\("
    ),
    re.compile(
        r"(?<![A-Za-z0-9_])(?:r#)?[A-Za-z_][A-Za-z0-9_]*"
        r"[ \t\r\n]*\.[ \t\r\n]*(?:r#)?"
        r"(?P<method>set_section_background)\b[ \t\r\n]*\("
    ),
)
IWA_PAGES_README_SECTION_BACKGROUND_EXAMPLE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<example>set_pages_section_background)"
    r"(?:\.rs)?(?![A-Za-z0-9_])"
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
        "IWorkPackage",
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
CAMEL_CASE_WORD = re.compile(r"[A-Z]+(?=[A-Z][a-z]|$)|[A-Z]?[a-z]+|[0-9]+")
RUST_BYTE_SLICE = re.compile(
    r"&[ \t\r\n]*(?:'[A-Za-z_][A-Za-z0-9_]*[ \t\r\n]+)?"
    r"(?:mut[ \t\r\n]+)?\[[ \t\r\n]*u8[ \t\r\n]*\]"
)
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


def _rust_doc_identifier_occurrences(
    source: str, names: frozenset[str]
) -> list[tuple[str, int]]:
    """Return exact identifiers in Rust doc comments and `#[doc = ...]` attributes."""

    regions = [
        *re.finditer(r"^[ \t]*//[/!][^\r\n]*", source, re.MULTILINE),
        *re.finditer(r"/\*(?:\*|!)[\s\S]*?\*/", source),
        *re.finditer(r"#\s*\[\s*doc\s*=\s*[^\]]*\]", source),
    ]
    occurrences: set[tuple[str, int]] = set()
    for region in regions:
        for identifier in RUST_IDENTIFIER.finditer(region.group(0)):
            name = identifier.group(1)
            if name not in names:
                continue
            offset = region.start() + identifier.start(1)
            occurrences.add((name, source.count("\n", 0, offset) + 1))
    return sorted(occurrences, key=lambda item: (item[1], item[0]))


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

    if identifier.startswith("litchi_iwa"):
        return "archive/IWA type"
    if identifier in {"buffa", "prost", "prost_types"}:
        return "protobuf type"
    if identifier in {"IWorkPackage", "SourceCatalog"}:
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


def _numbers_names_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in the Numbers names API."""

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


def _numbers_table_header_settings_public_leak(identifier: str) -> str | None:
    """Classify implementation vocabulary forbidden in header transactions."""

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
    if identifier in {"litchi_iwa", "litchi_iwa_common"}:
        return "litchi-iwa formula facade"
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
    return bool(identifiers & KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES) or (
        _keynote_slide_transition_owner_declaration(declaration)
    )


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


def audit_iwa_keynote_source_topology(root: Path = ROOT) -> list[str]:
    """Prevent retired Keynote mutation surfaces from returning to the host."""

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    violations: list[str] = []
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_KEYNOTE_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Keynote method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    retired_example = root / RETIRED_IWA_KEYNOTE_SLIDE_NAME_EXAMPLE
    if retired_example.exists():
        violations.append(
            "retired litchi-iwa Keynote slide-name example returned: "
            + str(RETIRED_IWA_KEYNOTE_SLIDE_NAME_EXAMPLE)
        )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.exists():
        source = readme_path.read_text(encoding="utf-8")
        for match in IWA_KEYNOTE_README_SLIDE_NAME_CALL.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote slide-name README call "
                f"{match.group('method')}: {IWA_KEYNOTE_README}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_keynote_document_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the retired Keynote reader and its compatibility surface deleted."""

    violations: list[str] = []
    retired_source = root / RETIRED_IWA_KEYNOTE_DOCUMENT_SOURCE
    if retired_source.exists():
        violations.append(
            "retired litchi-iwa Keynote document reader source returned: "
            + str(RETIRED_IWA_KEYNOTE_DOCUMENT_SOURCE)
        )

    module_path = root / IWA_KEYNOTE_MODULE_SOURCE
    if module_path.is_file():
        module_source = _mask_rust_non_code(
            module_path.read_text(encoding="utf-8")
        )
        for match in IWA_KEYNOTE_DOCUMENT_MODULE.finditer(module_source):
            line_number = module_source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote document reader module "
                f"{match.group(1)}: {IWA_KEYNOTE_MODULE_SOURCE}:{line_number}"
            )
        for match in IWA_KEYNOTE_DOCUMENT_LOCAL_REEXPORT.finditer(module_source):
            line_number = module_source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote document reader local re-export "
                f"{match.group('module')}: {IWA_KEYNOTE_MODULE_SOURCE}:{line_number}"
            )

    caller_paths: set[Path] = set()
    for caller_root in IWA_KEYNOTE_DOCUMENT_CALLER_ROOTS:
        caller_path = root / caller_root
        if caller_path.is_dir():
            caller_paths.update(caller_path.rglob("*.rs"))
    for path in sorted(caller_paths):
        raw_source = path.read_text(encoding="utf-8")
        source = _mask_rust_non_code(raw_source)
        for match in RUST_IDENTIFIER.finditer(source):
            name = match.group(1)
            if name not in RETIRED_IWA_KEYNOTE_DOCUMENT_TYPE_SET:
                continue
            line_number = source.count("\n", 0, match.start(1)) + 1
            violations.append(
                "retired litchi-iwa Keynote document reader type usage "
                f"{name}: {path.relative_to(root)}:{line_number}"
            )
        for name, line_number in _rust_doc_identifier_occurrences(
            raw_source, RETIRED_IWA_KEYNOTE_DOCUMENT_TYPE_SET
        ):
            violations.append(
                "retired litchi-iwa Keynote document reader rustdoc reference "
                f"{name}: {path.relative_to(root)}:{line_number}"
            )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for match in RUST_IDENTIFIER.finditer(source):
            name = match.group(1)
            if name not in RETIRED_IWA_KEYNOTE_DOCUMENT_TYPE_SET:
                continue
            line_number = source.count("\n", 0, match.start(1)) + 1
            violations.append(
                "retired litchi-iwa Keynote document reader README reference "
                f"{name}: {IWA_KEYNOTE_README}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_keynote_show_settings_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Keynote show-settings APIs and files out of the host."""

    violations: list[str] = []
    for retired, label in (
        (RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_SOURCE, "source"),
        (RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_EXAMPLE, "example"),
    ):
        if (root / retired).exists():
            violations.append(
                f"retired litchi-iwa Keynote show-settings {label} returned: {retired}"
            )

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Keynote show-settings method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    editor_path = root / IWA_KEYNOTE_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_KEYNOTE_SHOW_SETTINGS_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote show-settings module "
                f"{match.group(1)}: {IWA_KEYNOTE_EDITOR_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_KEYNOTE_README_SHOW_SETTINGS_CALLS:
            for match in pattern.finditer(source):
                method_offset = match.start("method")
                line_number = source.count("\n", 0, method_offset) + 1
                violations.append(
                    "retired litchi-iwa Keynote show-settings README call "
                    f"{match.group('method')}: {IWA_KEYNOTE_README}:{line_number}"
                )

    return sorted(set(violations))


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


def audit_iwa_keynote_soundtrack_settings_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Keynote soundtrack-setting ownership out of the host."""

    violations: list[str] = []
    for retired, label in (
        (RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_SOURCE, "source"),
        (RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_EXAMPLE, "example"),
    ):
        if (root / retired).exists():
            violations.append(
                "retired litchi-iwa Keynote soundtrack settings "
                f"{label} returned: {retired}"
            )

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Keynote soundtrack settings method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    wire_path = root / IWA_KEYNOTE_SOUNDTRACK_WIRE_SOURCE
    if wire_path.is_file():
        source = wire_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_METHOD_SET:
                continue
            violations.append(
                "retired litchi-iwa Keynote soundtrack settings wire helper "
                f"{name}: {IWA_KEYNOTE_SOUNDTRACK_WIRE_SOURCE}:{line_number}"
            )

    editor_path = root / IWA_KEYNOTE_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_KEYNOTE_SOUNDTRACK_SETTINGS_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start(1)) + 1
            violations.append(
                "retired litchi-iwa Keynote soundtrack settings module "
                f"{match.group(1)}: {IWA_KEYNOTE_EDITOR_SOURCE}:{line_number}"
            )

    tests_path = root / IWA_KEYNOTE_EDITOR_TEST_SOURCE
    if tests_path.is_file():
        source = tests_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_TEST_SET:
                continue
            violations.append(
                "retired litchi-iwa Keynote soundtrack settings test "
                f"{name}: {IWA_KEYNOTE_EDITOR_TEST_SOURCE}:{line_number}"
            )

    for caller_source in IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLER_SOURCES:
        caller_path = root / caller_source
        if not caller_path.is_file():
            continue
        source = _mask_rust_non_code(caller_path.read_text(encoding="utf-8"))
        for pattern in IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Keynote soundtrack settings caller "
                    f"{match.group('method')}: {caller_source}:{line_number}"
                )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Keynote soundtrack settings README call "
                    f"{match.group('method')}: {IWA_KEYNOTE_README}:{line_number}"
                )
        for match in IWA_KEYNOTE_README_SOUNDTRACK_SETTINGS_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Keynote soundtrack settings README example "
                f"reference {match.group('example')}: "
                f"{IWA_KEYNOTE_README}:{line_number}"
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


def audit_iwa_keynote_slide_transition_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Keynote transition mutations out of their former host."""

    violations: list[str] = []
    for retired in RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_SOURCES:
        if (root / retired).exists():
            violations.append(
                "retired litchi-iwa Keynote slide-transition source returned: "
                + str(retired)
            )
    for retired in RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_EXAMPLES:
        if (root / retired).exists():
            violations.append(
                "retired litchi-iwa Keynote slide-transition example returned: "
                + str(retired)
            )

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Keynote slide-transition method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    editor_path = root / IWA_KEYNOTE_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_KEYNOTE_SLIDE_TRANSITION_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote slide-transition module "
                f"{match.group(1)}: {IWA_KEYNOTE_EDITOR_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_KEYNOTE_README_SLIDE_TRANSITION_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Keynote slide-transition README call "
                    f"{match.group('method')}: {IWA_KEYNOTE_README}:{line_number}"
                )
        for match in IWA_KEYNOTE_README_SLIDE_TRANSITION_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Keynote slide-transition README example reference "
                f"{match.group('example')}: {IWA_KEYNOTE_README}:{line_number}"
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


def audit_iwa_keynote_slide_delete_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Keynote slide deletion out of its former host."""

    violations: list[str] = []
    if (root / RETIRED_IWA_KEYNOTE_SLIDE_DELETE_SOURCE).exists():
        violations.append(
            "retired litchi-iwa Keynote slide-delete source returned: "
            + str(RETIRED_IWA_KEYNOTE_SLIDE_DELETE_SOURCE)
        )
    if (root / RETIRED_IWA_KEYNOTE_SLIDE_DELETE_EXAMPLE).exists():
        violations.append(
            "retired litchi-iwa Keynote slide-delete example returned: "
            + str(RETIRED_IWA_KEYNOTE_SLIDE_DELETE_EXAMPLE)
        )

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_KEYNOTE_SLIDE_DELETE_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Keynote slide-delete method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    editor_path = root / IWA_KEYNOTE_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_KEYNOTE_SLIDE_DELETE_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote slide-delete module "
                f"{match.group(1)}: {IWA_KEYNOTE_EDITOR_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_KEYNOTE_README_SLIDE_DELETE_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Keynote slide-delete README call "
                    f"{match.group('method')}: {IWA_KEYNOTE_README}:{line_number}"
                )
        for match in IWA_KEYNOTE_README_SLIDE_DELETE_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Keynote slide-delete README example reference "
                f"{match.group('example')}: {IWA_KEYNOTE_README}:{line_number}"
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


def audit_iwa_keynote_placeholder_visibility_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Keynote placeholder visibility ownership out of the host."""

    violations: list[str] = []
    for retired, label in (
        (RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_SOURCE, "source"),
        (RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_EXAMPLE, "example"),
    ):
        if (root / retired).exists():
            violations.append(
                "retired litchi-iwa Keynote placeholder visibility "
                f"{label} returned: {retired}"
            )

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Keynote placeholder visibility method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )
            for declaration, line_number in _rust_public_declarations(source):
                for match in RUST_IDENTIFIER.finditer(declaration):
                    name = match.group(1)
                    if (
                        name
                        not in RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_PUBLIC_TYPE_SET
                    ):
                        continue
                    identifier_line = line_number + declaration.count(
                        "\n", 0, match.start(1)
                    )
                    violations.append(
                        "retired litchi-iwa Keynote placeholder visibility public type "
                        f"{name}: {path.relative_to(root)}:{identifier_line}"
                    )

    editor_path = root / IWA_KEYNOTE_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote placeholder visibility module "
                f"{match.group(1)}: {IWA_KEYNOTE_EDITOR_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_KEYNOTE_README_PLACEHOLDER_VISIBILITY_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Keynote placeholder visibility README call "
                    f"{match.group('method')}: {IWA_KEYNOTE_README}:{line_number}"
                )
        for match in IWA_KEYNOTE_README_PLACEHOLDER_VISIBILITY_EXAMPLE.finditer(
            source
        ):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Keynote placeholder visibility README example "
                f"reference {match.group('example')}: "
                f"{IWA_KEYNOTE_README}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_keynote_slide_number_visibility_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired per-slide number visibility ownership out of the host."""

    violations: list[str] = []
    for retired, label in (
        (RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_SOURCE, "source"),
        (RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_EXAMPLE, "example"),
    ):
        if (root / retired).exists():
            violations.append(
                "retired litchi-iwa Keynote slide-number visibility "
                f"{label} returned: {retired}"
            )

    source_root = root / IWA_KEYNOTE_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Keynote slide-number visibility method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    editor_path = root / IWA_KEYNOTE_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Keynote slide-number visibility module "
                f"{match.group(1)}: {IWA_KEYNOTE_EDITOR_SOURCE}:{line_number}"
            )

    tests_path = root / IWA_KEYNOTE_EDITOR_TEST_SOURCE
    if tests_path.is_file():
        source = tests_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_TEST_SET:
                continue
            violations.append(
                "retired litchi-iwa Keynote slide-number visibility test "
                f"{name}: {IWA_KEYNOTE_EDITOR_TEST_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_KEYNOTE_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_KEYNOTE_README_SLIDE_NUMBER_VISIBILITY_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Keynote slide-number visibility README call "
                    f"{match.group('method')}: {IWA_KEYNOTE_README}:{line_number}"
                )
        for match in IWA_KEYNOTE_README_SLIDE_NUMBER_VISIBILITY_EXAMPLE.finditer(
            source
        ):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Keynote slide-number visibility README example "
                f"reference {match.group('example')}: "
                f"{IWA_KEYNOTE_README}:{line_number}"
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


def audit_iwa_numbers_names_source_topology(root: Path = ROOT) -> list[str]:
    """Keep retired Numbers naming mutations and examples out of the host."""

    violations: list[str] = []
    example_path = root / RETIRED_IWA_NUMBERS_NAMES_EXAMPLE
    if example_path.exists():
        violations.append(
            "retired litchi-iwa Numbers names example returned: "
            + str(RETIRED_IWA_NUMBERS_NAMES_EXAMPLE)
        )

    source_root = root / IWA_NUMBERS_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_NUMBERS_NAMES_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Numbers names method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    readme_path = root / IWA_NUMBERS_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_NUMBERS_README_NAMES_CALLS:
            for match in pattern.finditer(source):
                method_offset = match.start("method")
                line_number = source.count("\n", 0, method_offset) + 1
                violations.append(
                    "retired litchi-iwa Numbers names README call "
                    f"{match.group('method')}: {IWA_NUMBERS_README}:{line_number}"
                )
        for match in IWA_NUMBERS_README_NAMES_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Numbers names README example reference "
                f"{match.group('example')}: {IWA_NUMBERS_README}:{line_number}"
            )

    return sorted(set(violations))


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
            if not _is_numbers_names_public_declaration(
                declaration, dedicated_source=complete_source_scope
            ):
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


def audit_iwa_numbers_sheet_order_source_topology(root: Path = ROOT) -> list[str]:
    """Keep retired Numbers sheet-order ownership out of the host facade."""

    violations: list[str] = []
    example_path = root / RETIRED_IWA_NUMBERS_SHEET_ORDER_EXAMPLE
    if example_path.exists():
        violations.append(
            "retired litchi-iwa Numbers sheet-order example returned: "
            + str(RETIRED_IWA_NUMBERS_SHEET_ORDER_EXAMPLE)
        )

    source_root = root / IWA_NUMBERS_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_NUMBERS_SHEET_ORDER_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Numbers sheet-order method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    tests_path = root / IWA_NUMBERS_EDITOR_TEST_SOURCE
    if tests_path.is_file():
        source = tests_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_NUMBERS_SHEET_ORDER_TEST_SET:
                continue
            violations.append(
                "retired litchi-iwa Numbers sheet-order test "
                f"{name}: {IWA_NUMBERS_EDITOR_TEST_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_NUMBERS_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_NUMBERS_README_SHEET_ORDER_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Numbers sheet-order README call "
                    f"{match.group('method')}: {IWA_NUMBERS_README}:{line_number}"
                )
        for match in IWA_NUMBERS_README_SHEET_ORDER_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Numbers sheet-order README example reference "
                f"{match.group('example')}: {IWA_NUMBERS_README}:{line_number}"
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


def audit_iwa_numbers_table_header_settings_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Numbers table-header methods out of the host facade."""

    violations: list[str] = []
    example = root / RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_EXAMPLE
    if example.exists():
        violations.append(
            "retired litchi-iwa Numbers table-header settings example returned: "
            + str(RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_EXAMPLE)
        )

    source_root = root / IWA_NUMBERS_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Numbers table-header settings method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    readme_path = root / IWA_NUMBERS_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_NUMBERS_README_TABLE_HEADER_SETTINGS_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Numbers table-header settings README call "
                    f"{match.group('method')}: {IWA_NUMBERS_README}:{line_number}"
                )
        for match in IWA_NUMBERS_README_TABLE_HEADER_SETTINGS_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Numbers table-header settings README example "
                f"reference {match.group('example')}: "
                f"{IWA_NUMBERS_README}:{line_number}"
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


def audit_iwa_numbers_table_title_settings_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Numbers table-title ownership out of the host facade."""

    violations: list[str] = []
    example = root / RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_EXAMPLE
    if example.exists():
        violations.append(
            "retired litchi-iwa Numbers table-title settings example returned: "
            + str(RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_EXAMPLE)
        )

    source_root = root / IWA_NUMBERS_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Numbers table-title settings method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    tests_path = root / IWA_NUMBERS_EDITOR_TEST_SOURCE
    if tests_path.is_file():
        source = tests_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_TEST_SET:
                continue
            violations.append(
                "retired litchi-iwa Numbers table-title settings test "
                f"{name}: {IWA_NUMBERS_EDITOR_TEST_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_NUMBERS_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_NUMBERS_README_TABLE_TITLE_SETTINGS_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Numbers table-title settings README call "
                    f"{match.group('method')}: {IWA_NUMBERS_README}:{line_number}"
                )
        for match in IWA_NUMBERS_README_TABLE_TITLE_SETTINGS_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Numbers table-title settings README example "
                f"reference {match.group('example')}: "
                f"{IWA_NUMBERS_README}:{line_number}"
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


def audit_iwa_numbers_table_dimension_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep public Numbers sizing ownership out of the shared IWA host."""

    violations: list[str] = []
    example = root / RETIRED_IWA_NUMBERS_TABLE_DIMENSION_EXAMPLE
    if example.exists():
        violations.append(
            "retired litchi-iwa Numbers table-dimension example returned: "
            + str(RETIRED_IWA_NUMBERS_TABLE_DIMENSION_EXAMPLE)
        )

    tests_path = root / IWA_NUMBERS_EDITOR_TEST_SOURCE
    if tests_path.is_file():
        source = tests_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TEST_SET:
                continue
            violations.append(
                "retired litchi-iwa Numbers table-dimension test "
                f"{name}: {IWA_NUMBERS_EDITOR_TEST_SOURCE}:{line_number}"
            )

    source_root = root / IWA_NUMBERS_SOURCE_ROOT
    if not source_root.is_dir():
        return sorted(set(violations))

    for path in sorted(source_root.rglob("*.rs")):
        raw_source = path.read_text(encoding="utf-8")
        source = _mask_rust_non_code(raw_source)
        focused_aliases = set(RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPE_SET)
        focused_aliases.update(IWA_NUMBERS_PRIVATE_TABLE_DIMENSION_ALIASES)
        focused_modules: set[str] = set()

        focused_imports = tuple(
            re.finditer(
                r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
                r"use[ \t\r\n]+(?P<body>[^;]*\blitchi_numbers\b[^;]*);",
                source,
                re.MULTILINE,
            )
        )
        for imported in focused_imports:
            body = imported.group("body")
            for module_pattern in (
                r"\blitchi_numbers\b[ \t\r\n]+as[ \t\r\n]+"
                r"(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*)",
                r"\blitchi_numbers\b[ \t\r\n]*::[ \t\r\n]*"
                r"(?:(?:r#)?table[ \t\r\n]*::[ \t\r\n]*)?"
                r"(?:r#)?dimension\b(?:[ \t\r\n]+as[ \t\r\n]+"
                r"(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?",
            ):
                module_alias = re.search(module_pattern, body)
                if module_alias is not None:
                    focused_modules.add(module_alias.group("alias") or "dimension")
            if "*" in body:
                focused_aliases.update(RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPE_SET)
            for name in RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPE_SET:
                named = re.search(
                    rf"(?<![A-Za-z0-9_#])(?:r#)?{name}\b"
                    rf"(?:[ \t\r\n]+as[ \t\r\n]+(?:r#)?"
                    rf"(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?",
                    body,
                )
                if named is not None:
                    focused_aliases.add(named.group("alias") or name)

        type_aliases = tuple(
            re.finditer(
                r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?type"
                r"[ \t\r\n]+(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*)"
                r"[^=;]*=[ \t\r\n]*(?P<target>[^;]+);",
                source,
                re.MULTILINE,
            )
        )
        changed = True
        while changed:
            changed = False
            for alias in type_aliases:
                identifiers = {
                    match.group(1)
                    for match in RUST_IDENTIFIER.finditer(alias.group("target"))
                }
                if not (
                    identifiers & focused_aliases
                    or identifiers & focused_modules
                    and identifiers & RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPE_SET
                ):
                    continue
                alias_name = alias.group("alias")
                if alias_name not in focused_aliases:
                    focused_aliases.add(alias_name)
                    changed = True

        for declaration, line_number in _rust_public_declarations(raw_source):
            identifiers = [
                match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
            ]
            for method in sorted(
                set(identifiers) & RETIRED_IWA_NUMBERS_TABLE_DIMENSION_METHOD_SET
            ):
                identifier = next(
                    match
                    for match in RUST_IDENTIFIER.finditer(declaration)
                    if match.group(1) == method
                )
                identifier_line = line_number + declaration.count(
                    "\n", 0, identifier.start(1)
                )
                violations.append(
                    "retired litchi-iwa Numbers public table-dimension method "
                    f"{method}: {path.relative_to(root)}:{identifier_line}"
                )

            exposed = set(identifiers) & focused_aliases
            declaration_is_facade = identifiers[:2] in (
                ["pub", "type"],
                ["pub", "use"],
            )
            if declaration_is_facade:
                exposed.update(set(identifiers) & focused_modules)
                if re.search(
                    r"^[ \t]*pub[ \t\r\n]+use[ \t\r\n]+"
                    r"(?:r#)?litchi_numbers[ \t\r\n]*(?:;|::[ \t\r\n]*\*)",
                    declaration,
                ):
                    exposed.add("litchi_numbers")
            if set(identifiers) & focused_modules:
                exposed.update(
                    set(identifiers) & RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPE_SET
                )
            for name in sorted(exposed):
                identifier = next(
                    match
                    for match in RUST_IDENTIFIER.finditer(declaration)
                    if match.group(1) == name
                )
                identifier_line = line_number + declaration.count(
                    "\n", 0, identifier.start(1)
                )
                violations.append(
                    "retired litchi-iwa Numbers table-dimension public facade "
                    f"{name}: {path.relative_to(root)}:{identifier_line}"
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


def audit_iwa_numbers_table_cell_mutation_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired raw-ID Numbers cell writers out of their exact host scopes."""

    violations: list[str] = []
    example = root / RETIRED_IWA_NUMBERS_TABLE_CELL_EXAMPLE
    if example.exists():
        violations.append(
            "retired litchi-iwa Numbers table-cell mutation example returned: "
            + str(RETIRED_IWA_NUMBERS_TABLE_CELL_EXAMPLE)
        )

    retired_sources = (
        (
            RETIRED_IWA_NUMBERS_TABLE_CELL_EDITOR_SOURCE,
            frozenset(RETIRED_IWA_NUMBERS_TABLE_CELL_METHODS),
            "method",
        ),
        (
            RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_SOURCE,
            frozenset(RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_HELPERS),
            "model helper",
        ),
        (
            RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_SOURCE,
            frozenset(RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_HELPERS),
            "batch helper",
        ),
        (
            RETIRED_IWA_NUMBERS_TABLE_CELL_TEST_SOURCE,
            frozenset(RETIRED_IWA_NUMBERS_TABLE_CELL_TESTS),
            "test",
        ),
    )
    for relative_path, retired_names, description in retired_sources:
        path = root / relative_path
        if not path.is_file():
            continue
        source = path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in retired_names:
                continue
            violations.append(
                "retired litchi-iwa Numbers table-cell mutation "
                f"{description} {name}: {relative_path}:{line_number}"
            )

    fixture_path = root / IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_SOURCE
    if fixture_path.is_file():
        fixture_source = _mask_rust_non_code(
            fixture_path.read_text(encoding="utf-8")
        )
        gated_name_offsets = {
            match.start("name")
            for match in IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_DECLARATION.finditer(
                fixture_source
            )
        }
        for match in RUST_FUNCTION_DECLARATION.finditer(fixture_source):
            name = match.group(1)
            if (
                name not in IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_HELPER_SET
                or match.start(1) in gated_name_offsets
            ):
                continue
            line_number = fixture_source.count("\n", 0, match.start(1)) + 1
            violations.append(
                "litchi-iwa Numbers table-cell test fixture helper must be "
                f"private #[cfg(test)] {name}: "
                f"{IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_SOURCE}:{line_number}"
            )

    example_root = root / IWA_NUMBERS_EXAMPLE_ROOT
    if example_root.is_dir():
        for path in sorted(example_root.rglob("*.rs")):
            source = _mask_rust_non_code(path.read_text(encoding="utf-8"))
            for match in RUST_IDENTIFIER.finditer(source):
                name = match.group(1)
                if name not in IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_HELPER_SET:
                    continue
                line_number = source.count("\n", 0, match.start(1)) + 1
                violations.append(
                    "litchi-iwa Numbers example calls test-only table-cell fixture "
                    f"helper {name}: {path.relative_to(root)}:{line_number}"
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


def audit_iwa_pages_document_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the retired Pages reader and host compatibility surfaces deleted."""

    violations: list[str] = []
    retired_source = root / RETIRED_IWA_PAGES_DOCUMENT_SOURCE
    if retired_source.exists():
        violations.append(
            "retired litchi-iwa Pages document reader source returned: "
            + str(RETIRED_IWA_PAGES_DOCUMENT_SOURCE)
        )

    module_path = root / IWA_PAGES_MODULE_SOURCE
    if module_path.is_file():
        module_source = _mask_rust_non_code(
            module_path.read_text(encoding="utf-8")
        )
        for match in IWA_PAGES_DOCUMENT_MODULE.finditer(module_source):
            line_number = module_source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Pages document reader module "
                f"{match.group(1)}: {IWA_PAGES_MODULE_SOURCE}:{line_number}"
            )
        for match in IWA_PAGES_DOCUMENT_LOCAL_REEXPORT.finditer(module_source):
            line_number = module_source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Pages document reader local re-export "
                f"{match.group('module')}: {IWA_PAGES_MODULE_SOURCE}:{line_number}"
            )

    workspace_sources = root / WORKSPACE_CRATES_ROOT
    if workspace_sources.is_dir():
        for path in sorted(workspace_sources.glob("*/src/**/*.rs")):
            raw_source = path.read_text(encoding="utf-8")
            for declaration, line_number in _rust_public_declarations(raw_source):
                for match in RUST_IDENTIFIER.finditer(declaration):
                    name = match.group(1)
                    if name not in RETIRED_IWA_PAGES_DOCUMENT_TYPE_SET:
                        continue
                    identifier_line = line_number + declaration.count(
                        "\n", 0, match.start(1)
                    )
                    violations.append(
                        "retired Pages document reader workspace public name "
                        f"{name}: {path.relative_to(root)}:{identifier_line}"
                    )
            for name, line_number in _rust_doc_identifier_occurrences(
                raw_source, RETIRED_IWA_PAGES_DOCUMENT_TYPE_SET
            ):
                violations.append(
                    "retired Pages document reader workspace public rustdoc "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    host_source_root = root / IWA_HOST_SOURCE_ROOT
    if host_source_root.is_dir():
        for path in sorted(host_source_root.rglob("*.rs")):
            raw_source = path.read_text(encoding="utf-8")
            source = _mask_rust_non_code(raw_source)
            focused_aliases: set[str] = set()
            focused_modules = {"litchi_pages"}
            focused_imports = re.finditer(
                r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
                r"use[ \t\r\n]+(?P<body>[^;]*\blitchi_pages\b[^;]*);",
                source,
                re.MULTILINE,
            )
            for imported in focused_imports:
                body = imported.group("body")
                module_alias = re.search(
                    r"\blitchi_pages\b[ \t\r\n]+as[ \t\r\n]+"
                    r"(?:r#)?([A-Za-z_][A-Za-z0-9_]*)",
                    body,
                )
                if module_alias is not None:
                    focused_modules.add(module_alias.group(1))
                    if re.match(r"^[ \t]*pub\b", imported.group(0)):
                        line_number = source.count("\n", 0, imported.start()) + 1
                        violations.append(
                            "retired litchi-iwa Pages document reader focused host "
                            f"facade {module_alias.group(1)}: "
                            f"{path.relative_to(root)}:{line_number}"
                        )
                self_alias = re.search(
                    r"(?<![A-Za-z0-9_#])(?:r#)?self[ \t\r\n]+as"
                    r"[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)",
                    body,
                )
                if self_alias is not None:
                    focused_modules.add(self_alias.group(1))
                    if re.match(r"^[ \t]*pub\b", imported.group(0)):
                        line_number = source.count("\n", 0, imported.start()) + 1
                        violations.append(
                            "retired litchi-iwa Pages document reader focused host "
                            f"facade {self_alias.group(1)}: "
                            f"{path.relative_to(root)}:{line_number}"
                        )
                if "*" in body:
                    focused_aliases.update(IWA_PAGES_FOCUSED_READER_TYPES)
                    if re.match(r"^[ \t]*pub\b", imported.group(0)):
                        line_number = source.count("\n", 0, imported.start()) + 1
                        violations.append(
                            "retired litchi-iwa Pages document reader focused host "
                            f"facade glob: {path.relative_to(root)}:{line_number}"
                        )
                for name in IWA_PAGES_FOCUSED_READER_TYPES:
                    named = re.search(
                        rf"(?<![A-Za-z0-9_#])(?:r#)?{name}\b"
                        rf"(?:[ \t\r\n]+as[ \t\r\n]+(?:r#)?"
                        rf"(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?",
                        body,
                    )
                    if named is not None:
                        focused_aliases.add(named.group("alias") or name)

            type_aliases = tuple(
                re.finditer(
                    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?type"
                    r"[ \t\r\n]+(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*)"
                    r"[^=;]*=[ \t\r\n]*(?P<target>[^;]+);",
                    source,
                    re.MULTILINE,
                )
            )
            changed = True
            while changed:
                changed = False
                for alias in type_aliases:
                    identifiers = {
                        match.group(1)
                        for match in RUST_IDENTIFIER.finditer(alias.group("target"))
                    }
                    focused_target = bool(identifiers & focused_aliases) or (
                        bool(identifiers & focused_modules)
                        and bool(identifiers & IWA_PAGES_FOCUSED_READER_TYPES)
                    )
                    if focused_target and alias.group("alias") not in focused_aliases:
                        focused_aliases.add(alias.group("alias"))
                        changed = True

            for declaration, line_number in _rust_public_declarations(raw_source):
                identifiers = {
                    match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
                }
                exposed = sorted(identifiers & focused_aliases)
                if identifiers & focused_modules:
                    exposed.extend(
                        sorted(
                            IWA_PAGES_FOCUSED_READER_TYPES & identifiers
                            - set(exposed)
                        )
                    )
                for name in exposed:
                    identifier = next(
                        match
                        for match in RUST_IDENTIFIER.finditer(declaration)
                        if match.group(1) == name
                    )
                    identifier_line = line_number + declaration.count(
                        "\n", 0, identifier.start(1)
                    )
                    violations.append(
                        "retired litchi-iwa Pages document reader focused host "
                        f"facade {name}: {path.relative_to(root)}:{identifier_line}"
                    )

    caller_paths: set[Path] = set()
    for caller_root in (
        Path("crates/litchi-iwa/src"),
        Path("crates/litchi-iwa/tests"),
        Path("crates/litchi-iwa/examples"),
    ):
        caller_path = root / caller_root
        if caller_path.is_dir():
            caller_paths.update(caller_path.rglob("*.rs"))
    for path in sorted(caller_paths):
        raw_source = path.read_text(encoding="utf-8")
        source = _mask_rust_non_code(raw_source)
        for match in RUST_IDENTIFIER.finditer(source):
            name = match.group(1)
            if name not in RETIRED_IWA_PAGES_DOCUMENT_TYPE_SET:
                continue
            line_number = source.count("\n", 0, match.start(1)) + 1
            violations.append(
                "retired litchi-iwa Pages document reader type usage "
                f"{name}: {path.relative_to(root)}:{line_number}"
            )
        for name, line_number in _rust_doc_identifier_occurrences(
            raw_source, RETIRED_IWA_PAGES_DOCUMENT_TYPE_SET
        ):
            violations.append(
                "retired litchi-iwa Pages document reader rustdoc reference "
                f"{name}: {path.relative_to(root)}:{line_number}"
            )

    readme_path = root / IWA_PAGES_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for match in RUST_IDENTIFIER.finditer(source):
            name = match.group(1)
            if name not in RETIRED_IWA_PAGES_DOCUMENT_TYPE_SET:
                continue
            line_number = source.count("\n", 0, match.start(1)) + 1
            violations.append(
                "retired litchi-iwa Pages document reader README reference "
                f"{name}: {IWA_PAGES_README}:{line_number}"
            )

    return sorted(set(violations))


def audit_iwa_numbers_document_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the retired Numbers reader and host compatibility surfaces deleted."""

    violations: list[str] = []
    retired_source = root / RETIRED_IWA_NUMBERS_DOCUMENT_SOURCE
    if retired_source.exists():
        violations.append(
            "retired litchi-iwa Numbers document reader source returned: "
            + str(RETIRED_IWA_NUMBERS_DOCUMENT_SOURCE)
        )
    retired_sheet_source = root / RETIRED_IWA_NUMBERS_SHEET_SOURCE
    if retired_sheet_source.exists():
        violations.append(
            "retired litchi-iwa Numbers sheet reader source returned: "
            + str(RETIRED_IWA_NUMBERS_SHEET_SOURCE)
        )

    module_path = root / IWA_NUMBERS_MODULE_SOURCE
    if module_path.is_file():
        module_source = _mask_rust_non_code(
            module_path.read_text(encoding="utf-8")
        )
        for match in IWA_NUMBERS_READER_MODULE.finditer(module_source):
            line_number = module_source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Numbers reader module "
                f"{match.group('module')}: {IWA_NUMBERS_MODULE_SOURCE}:{line_number}"
            )
        for match in IWA_NUMBERS_READER_LOCAL_REEXPORT.finditer(module_source):
            line_number = module_source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Numbers reader local re-export "
                f"{match.group('module')}: {IWA_NUMBERS_MODULE_SOURCE}:{line_number}"
            )

    workspace_sources = root / WORKSPACE_CRATES_ROOT
    if workspace_sources.is_dir():
        for path in sorted(workspace_sources.rglob("*.rs")):
            raw_source = path.read_text(encoding="utf-8")
            if not any(
                name in raw_source for name in RETIRED_IWA_NUMBERS_READER_TYPE_SET
            ):
                continue
            source = _mask_rust_non_code(raw_source)
            for match in RUST_IDENTIFIER.finditer(source):
                name = match.group(1)
                if name not in RETIRED_IWA_NUMBERS_READER_TYPE_SET:
                    continue
                line_number = source.count("\n", 0, match.start(1)) + 1
                violations.append(
                    "retired Numbers document reader workspace type usage "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )
            for declaration, line_number in _rust_public_declarations(raw_source):
                for match in RUST_IDENTIFIER.finditer(declaration):
                    name = match.group(1)
                    if name not in RETIRED_IWA_NUMBERS_READER_TYPE_SET:
                        continue
                    identifier_line = line_number + declaration.count(
                        "\n", 0, match.start(1)
                    )
                    violations.append(
                        "retired Numbers document reader workspace public name "
                        f"{name}: {path.relative_to(root)}:{identifier_line}"
                    )
            for name, line_number in _rust_doc_identifier_occurrences(
                raw_source, RETIRED_IWA_NUMBERS_READER_TYPE_SET
            ):
                violations.append(
                    "retired Numbers document reader workspace public rustdoc "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    host_source_root = root / IWA_HOST_SOURCE_ROOT
    if host_source_root.is_dir():
        for path in sorted(host_source_root.rglob("*.rs")):
            raw_source = path.read_text(encoding="utf-8")
            source = _mask_rust_non_code(raw_source)
            focused_aliases: set[str] = set()
            focused_modules = {"litchi_numbers"}
            focused_imports = re.finditer(
                r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?"
                r"use[ \t\r\n]+(?P<body>[^;]*\blitchi_numbers\b[^;]*);",
                source,
                re.MULTILINE,
            )
            for imported in focused_imports:
                body = imported.group("body")
                module_alias = re.search(
                    r"\blitchi_numbers\b[ \t\r\n]+as[ \t\r\n]+"
                    r"(?:r#)?([A-Za-z_][A-Za-z0-9_]*)",
                    body,
                )
                if module_alias is not None:
                    focused_modules.add(module_alias.group(1))
                    if re.match(r"^[ \t]*pub\b", imported.group(0)):
                        line_number = source.count("\n", 0, imported.start()) + 1
                        violations.append(
                            "retired litchi-iwa Numbers document reader focused host "
                            f"facade {module_alias.group(1)}: "
                            f"{path.relative_to(root)}:{line_number}"
                        )
                self_alias = re.search(
                    r"(?<![A-Za-z0-9_#])(?:r#)?self[ \t\r\n]+as"
                    r"[ \t\r\n]+(?:r#)?([A-Za-z_][A-Za-z0-9_]*)",
                    body,
                )
                if self_alias is not None:
                    focused_modules.add(self_alias.group(1))
                    if re.match(r"^[ \t]*pub\b", imported.group(0)):
                        line_number = source.count("\n", 0, imported.start()) + 1
                        violations.append(
                            "retired litchi-iwa Numbers document reader focused host "
                            f"facade {self_alias.group(1)}: "
                            f"{path.relative_to(root)}:{line_number}"
                        )
                focused_module = re.search(
                    r"\blitchi_numbers\b[ \t\r\n]*::[ \t\r\n]*"
                    r"(?:\{[ \t\r\n]*)?(?:r#)?(?P<module>document|package)\b"
                    r"(?:[ \t\r\n]+as[ \t\r\n]+(?:r#)?"
                    r"(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?",
                    body,
                )
                if focused_module is not None:
                    alias = focused_module.group("alias") or focused_module.group(
                        "module"
                    )
                    focused_modules.add(alias)
                    if re.match(r"^[ \t]*pub\b", imported.group(0)):
                        line_number = source.count("\n", 0, imported.start()) + 1
                        violations.append(
                            "retired litchi-iwa Numbers document reader focused host "
                            f"facade {alias}: {path.relative_to(root)}:{line_number}"
                        )
                if "*" in body:
                    focused_aliases.update(IWA_NUMBERS_FOCUSED_READER_TYPES)
                    if re.match(r"^[ \t]*pub\b", imported.group(0)):
                        line_number = source.count("\n", 0, imported.start()) + 1
                        violations.append(
                            "retired litchi-iwa Numbers document reader focused host "
                            f"facade glob: {path.relative_to(root)}:{line_number}"
                        )
                for name in IWA_NUMBERS_FOCUSED_READER_TYPES:
                    named = re.search(
                        rf"(?<![A-Za-z0-9_#])(?:r#)?{name}\b"
                        rf"(?:[ \t\r\n]+as[ \t\r\n]+(?:r#)?"
                        rf"(?P<alias>[A-Za-z_][A-Za-z0-9_]*))?",
                        body,
                    )
                    if named is not None:
                        focused_aliases.add(named.group("alias") or name)

            type_aliases = tuple(
                re.finditer(
                    r"^[ \t]*(?:pub(?:\([^()]*\))?[ \t\r\n]+)?type"
                    r"[ \t\r\n]+(?:r#)?(?P<alias>[A-Za-z_][A-Za-z0-9_]*)"
                    r"[^=;]*=[ \t\r\n]*(?P<target>[^;]+);",
                    source,
                    re.MULTILINE,
                )
            )
            changed = True
            while changed:
                changed = False
                for alias in type_aliases:
                    identifiers = {
                        match.group(1)
                        for match in RUST_IDENTIFIER.finditer(alias.group("target"))
                    }
                    focused_target = bool(identifiers & focused_aliases) or (
                        bool(identifiers & focused_modules)
                        and bool(identifiers & IWA_NUMBERS_FOCUSED_READER_TYPES)
                    )
                    if focused_target and alias.group("alias") not in focused_aliases:
                        focused_aliases.add(alias.group("alias"))
                        changed = True

            for declaration, line_number in _rust_public_declarations(raw_source):
                identifiers = {
                    match.group(1) for match in RUST_IDENTIFIER.finditer(declaration)
                }
                exposed = sorted(identifiers & focused_aliases)
                if identifiers & focused_modules:
                    exposed.extend(
                        sorted(
                            IWA_NUMBERS_FOCUSED_READER_TYPES
                            & identifiers
                            - set(exposed)
                        )
                    )
                for name in exposed:
                    identifier = next(
                        match
                        for match in RUST_IDENTIFIER.finditer(declaration)
                        if match.group(1) == name
                    )
                    identifier_line = line_number + declaration.count(
                        "\n", 0, identifier.start(1)
                    )
                    violations.append(
                        "retired litchi-iwa Numbers document reader focused host "
                        f"facade {name}: {path.relative_to(root)}:{identifier_line}"
                    )

    readme_path = root / IWA_NUMBERS_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for match in RUST_IDENTIFIER.finditer(source):
            name = match.group(1)
            if name not in RETIRED_IWA_NUMBERS_READER_TYPE_SET:
                continue
            line_number = source.count("\n", 0, match.start(1)) + 1
            violations.append(
                "retired litchi-iwa Numbers document reader README reference "
                f"{name}: {IWA_NUMBERS_README}:{line_number}"
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


def audit_pages_document_public_api(root: Path = ROOT) -> list[str]:
    """Keep the focused Pages reader public API archive-free and semantic."""

    violations: list[str] = []
    for relative in PAGES_DOCUMENT_PUBLIC_API_SOURCES:
        path = root / relative
        if not path.is_file():
            continue
        source = path.read_text(encoding="utf-8")
        dedicated_source = relative == PAGES_SOURCE_ROOT / "document.rs"
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


def audit_numbers_package_no_eager_prost_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep the Numbers package ingress free of retired generated reads.

    This is intentionally scoped to ``package.rs``.  The remaining Numbers
    package helpers still have independent Prost-backed migrations, so a
    crate-wide dependency or source ban would be premature.  The first
    ``cfg(test)`` module is excluded so canonical Prost fixtures remain
    available to differential tests without weakening the production gate.
    """

    violations: list[str] = []
    source_path = root / NUMBERS_PACKAGE_SOURCE
    if not source_path.is_file():
        return violations

    raw_source = source_path.read_text(encoding="utf-8")
    masked_source = _mask_rust_non_code(raw_source)
    test_module = NUMBERS_PACKAGE_TEST_MODULE.search(masked_source)
    production_source = (
        raw_source[: test_module.start()] if test_module is not None else raw_source
    )
    production_code = _mask_rust_non_code(production_source)
    for label, pattern in NUMBERS_PACKAGE_NO_EAGER_PROST_SOURCE_PATTERNS:
        for match in pattern.finditer(production_code):
            line_number = production_code.count("\n", 0, match.start()) + 1
            violations.append(
                "focused litchi-numbers package production source uses "
                f"{label}: {NUMBERS_PACKAGE_SOURCE}:{line_number}"
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
        masked_source = _mask_rust_non_code(raw_source)
        test_module = PAGES_PACKAGE_TEST_MODULE.search(masked_source)
        production_source = (
            raw_source[: test_module.start()] if test_module is not None else raw_source
        )
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
            masked_source = _mask_rust_non_code(raw_source)
            test_module = KEYNOTE_PRODUCTION_TEST_MODULE.search(masked_source)
            production_source = (
                raw_source[: test_module.start()]
                if test_module is not None
                else raw_source
            )
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


def audit_iwa_pages_document_settings_source_topology(root: Path = ROOT) -> list[str]:
    """Keep retired Pages document-settings APIs and modules out of the host."""

    violations: list[str] = []
    for retired in RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_SOURCES:
        path = root / retired
        if path.exists():
            violations.append(
                "retired litchi-iwa Pages document-settings source returned: "
                + str(retired)
            )

    source_root = root / IWA_PAGES_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Pages document-settings method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    editor_path = root / IWA_PAGES_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_PAGES_DOCUMENT_SETTINGS_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Pages document-settings module "
                f"{match.group(1)}: {IWA_PAGES_EDITOR_SOURCE}:{line_number}"
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


def audit_iwa_pages_section_settings_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Pages section settings and name ownership out of the host."""

    violations: list[str] = []
    example = root / RETIRED_IWA_PAGES_SECTION_SETTINGS_EXAMPLE
    if example.exists():
        violations.append(
            "retired litchi-iwa Pages section-settings example returned: "
            + str(RETIRED_IWA_PAGES_SECTION_SETTINGS_EXAMPLE)
        )

    source_root = root / IWA_PAGES_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_PAGES_SECTION_SETTINGS_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Pages section-settings method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    tests_path = root / IWA_PAGES_EDITOR_TEST_SOURCE
    if tests_path.is_file():
        source = tests_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_PAGES_SECTION_SETTINGS_TEST_SET:
                continue
            violations.append(
                "retired litchi-iwa Pages section-settings test "
                f"{name}: {IWA_PAGES_EDITOR_TEST_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_PAGES_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_PAGES_README_SECTION_SETTINGS_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Pages section-settings README call "
                    f"{match.group('method')}: {IWA_PAGES_README}:{line_number}"
                )
        for match in IWA_PAGES_README_SECTION_SETTINGS_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Pages section-settings README example "
                f"reference {match.group('example')}: {IWA_PAGES_README}:{line_number}"
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


def audit_iwa_pages_section_background_source_topology(
    root: Path = ROOT,
) -> list[str]:
    """Keep retired Pages section-background ownership out of the host."""

    violations: list[str] = []
    for retired, label in (
        (RETIRED_IWA_PAGES_SECTION_BACKGROUND_EXAMPLE, "example"),
        *(
            (source, "source")
            for source in RETIRED_IWA_PAGES_SECTION_BACKGROUND_SOURCES
        ),
    ):
        if (root / retired).exists():
            violations.append(
                "retired litchi-iwa Pages section-background "
                f"{label} returned: {retired}"
            )

    source_root = root / IWA_PAGES_SOURCE_ROOT
    if source_root.is_dir():
        for path in sorted(source_root.rglob("*.rs")):
            source = path.read_text(encoding="utf-8")
            for name, line_number in _rust_function_declarations(source):
                if name not in RETIRED_IWA_PAGES_SECTION_BACKGROUND_METHOD_SET:
                    continue
                violations.append(
                    "retired litchi-iwa Pages section-background method "
                    f"{name}: {path.relative_to(root)}:{line_number}"
                )

    editor_path = root / IWA_PAGES_EDITOR_SOURCE
    if editor_path.is_file():
        source = _mask_rust_non_code(editor_path.read_text(encoding="utf-8"))
        for match in IWA_PAGES_SECTION_BACKGROUND_MODULE.finditer(source):
            line_number = source.count("\n", 0, match.start()) + 1
            violations.append(
                "retired litchi-iwa Pages section-background module "
                f"{match.group(1)}: {IWA_PAGES_EDITOR_SOURCE}:{line_number}"
            )

    tests_path = root / IWA_PAGES_EDITOR_TEST_SOURCE
    if tests_path.is_file():
        source = tests_path.read_text(encoding="utf-8")
        for name, line_number in _rust_function_declarations(source):
            if name not in RETIRED_IWA_PAGES_SECTION_BACKGROUND_TEST_SET:
                continue
            violations.append(
                "retired litchi-iwa Pages section-background test "
                f"{name}: {IWA_PAGES_EDITOR_TEST_SOURCE}:{line_number}"
            )

    readme_path = root / IWA_PAGES_README
    if readme_path.is_file():
        source = readme_path.read_text(encoding="utf-8")
        for pattern in IWA_PAGES_README_SECTION_BACKGROUND_CALLS:
            for match in pattern.finditer(source):
                line_number = source.count("\n", 0, match.start("method")) + 1
                violations.append(
                    "retired litchi-iwa Pages section-background README call "
                    f"{match.group('method')}: {IWA_PAGES_README}:{line_number}"
                )
        for match in IWA_PAGES_README_SECTION_BACKGROUND_EXAMPLE.finditer(source):
            line_number = source.count("\n", 0, match.start("example")) + 1
            violations.append(
                "retired litchi-iwa Pages section-background README example "
                f"reference {match.group('example')}: "
                f"{IWA_PAGES_README}:{line_number}"
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
        + audit_iwa_keynote_document_source_topology()
        + audit_iwa_keynote_show_settings_source_topology()
        + audit_keynote_show_settings_facade_source_topology()
        + audit_iwa_keynote_soundtrack_settings_source_topology()
        + audit_keynote_soundtrack_settings_facade_source_topology()
        + audit_iwa_keynote_slide_transition_source_topology()
        + audit_keynote_slide_transition_facade_source_topology()
        + audit_iwa_keynote_slide_delete_source_topology()
        + audit_keynote_slide_delete_facade_source_topology()
        + audit_iwa_keynote_placeholder_visibility_source_topology()
        + audit_iwa_keynote_slide_number_visibility_source_topology()
        + audit_keynote_placeholder_visibility_facade_source_topology()
        + audit_keynote_package_no_eager_prost_source_topology()
        + audit_numbers_package_no_eager_prost_source_topology()
        + audit_iwa_numbers_names_source_topology()
        + audit_numbers_names_facade_source_topology()
        + audit_iwa_numbers_sheet_order_source_topology()
        + audit_numbers_sheet_order_facade_source_topology()
        + audit_iwa_numbers_table_header_settings_source_topology()
        + audit_numbers_table_header_settings_facade_source_topology()
        + audit_iwa_numbers_table_title_settings_source_topology()
        + audit_numbers_table_title_settings_facade_source_topology()
        + audit_iwa_numbers_table_dimension_source_topology()
        + audit_numbers_table_dimension_facade_source_topology()
        + audit_numbers_formula_facade_source_topology()
        + audit_numbers_table_cells_facade_source_topology()
        + audit_numbers_table_cells_mutation_facade_source_topology()
        + audit_iwa_numbers_table_cell_mutation_source_topology()
        + audit_iwa_numbers_table_lock_source_topology()
        + audit_numbers_table_lock_facade_source_topology()
        + audit_iwa_numbers_document_source_topology()
        + audit_numbers_document_public_api()
        + audit_iwa_pages_document_source_topology()
        + audit_pages_document_public_api()
        + audit_pages_package_no_eager_prost_source_topology()
        + audit_iwa_pages_page_layout_source_topology()
        + audit_pages_page_layout_facade_source_topology()
        + audit_iwa_pages_document_settings_source_topology()
        + audit_pages_document_settings_facade_source_topology()
        + audit_iwa_pages_section_settings_source_topology()
        + audit_pages_section_settings_facade_source_topology()
        + audit_iwa_pages_section_background_source_topology()
        + audit_pages_section_background_facade_source_topology()
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
