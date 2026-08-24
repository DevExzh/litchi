#!/usr/bin/env python3
"""Build a deterministic, fail-closed summary from four ABBA reports.

The input reports are the JSON files emitted by ``tools/perf-baseline``.  The
four legs are ordered ``A1, B1, B2, A2``: A is the control implementation and
B is the candidate implementation.  This module deliberately has no
third-party dependencies so that it can be used from a clean checkout.

The summary is descriptive evidence.  A statistic is marked accepted only
when both candidate directions are lower and both same-implementation drift
values are within the configured ceilings.  No speedup claim is inferred from
an accepted statistic.  Allocator-instrumented reports are intentionally
rejected here because their elapsed samples include instrumentation overhead;
use the operation-metric comparator for allocation-only evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

if __package__:
    from . import perf_compare
else:  # pragma: no cover - exercised by the direct CLI entry point
    import perf_compare


SCHEMA_VERSION = 1
HARNESS_SCHEMA_VERSION = 1
HARNESS_TOOL_NAME = "litchi-perf-baseline"
TOOL_NAME = "litchi-perf-abba-summary"
TOOL_VERSION = "0.1.0"
LEG_ORDER = ("a1", "b1", "b2", "a2")
STATISTICS = ("p50", "mean", "p95", "p99")
REPORT_PROFILE_LEGACY = "legacy-v1"
REPORT_PROFILE_CURRENT = "current-v1"
REPORT_PROFILES = frozenset((REPORT_PROFILE_LEGACY, REPORT_PROFILE_CURRENT))
U64_MAX = (1 << 64) - 1
U32_MAX = (1 << 32) - 1
MIN_RETAINED_SAMPLES = 15
ENVIRONMENT_VARIANTS = frozenset(("git_revision",))
REQUIRED_TOOL_FIELDS = (
    "name",
    "version",
    "binary",
    "profile",
    "target_os",
    "target_arch",
    "instrumentation",
)
REQUIRED_BINARY_IDENTITY_FIELDS = (
    "path",
    "binary_sha256",
    "binary_bytes",
    "mode_bits",
    "executable",
    "profile",
)
OPTIONAL_BINARY_IDENTITY_FIELDS = frozenset(("label",))
UNIX_TARGET_OSES = frozenset(
    {
        "aix", "android", "dragonfly", "freebsd", "fuchsia", "haiku",
        "hurd", "illumos", "ios", "linux", "macos", "netbsd", "openbsd",
        "tvos", "visionos", "watchos",
    }
)
REQUIRED_CONFIGURATION_FIELDS = (
    "samples_per_case",
    "warmup_iterations_per_case",
    "cases",
    "corpus_shapes",
)
REQUIRED_ENVIRONMENT_FIELDS = (
    "rustc_version",
    "git_revision",
    "git_worktree_dirty",
    "logical_cpus_available",
    "allocator",
    "rustflags",
    "cargo_build_target",
    "perf_event_paranoid",
)
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
STUDENT_T_CRITICAL_95 = (
    12.706,
    4.303,
    3.182,
    2.776,
    2.571,
    2.447,
    2.365,
    2.306,
    2.262,
    2.228,
    2.201,
    2.179,
    2.160,
    2.145,
    2.131,
    2.120,
    2.110,
    2.101,
    2.093,
    2.086,
    2.080,
    2.074,
    2.069,
    2.064,
    2.060,
    2.056,
    2.052,
    2.048,
    2.045,
    2.042,
)
STUDENT_T_Z_975 = 1.959_963_984_540_054
DEFAULT_DRIFT_CEILINGS: dict[str, float] = {
    "p50": 5.0,
    "mean": 5.0,
    "p95": 10.0,
    "p99": 15.0,
}

# These selectors use fixed, non-CLI corpora.  Their shapes therefore cannot be
# listed in any of the configurable ``*_shapes`` fields emitted by the
# historical schema-1 harness.  Keep the complete manifest exact and
# case-local: an unknown generator, an accidentally substituted archive, or an
# added nested field must fail closed.
_XLSX_REPEAT_STORE_MEDIUM_CORPUS: dict[str, Any] = {
    "name": "xlsx-source-repeated-store-medium",
    "generator": "litchi-xlsx-source-repeated-store-corpus-v1",
    "package_format": "XLSX/OPC/ZIP",
    "shape": "medium",
    "payload_kind": "fixed-medium-grid-for-repeated-selected-store",
    "compression": "deflate",
    "entry_count": 9216,
    "archive_member_count": 17,
    "entry_bytes": 4,
    "uncompressed_payload_bytes": 4_231_168,
    "archive_bytes": 4_226_429,
    "archive_sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
    "target_entry": "Sheet1!A1",
    "target_payload_bytes": 1,
    "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
    "xlsx": {
        "sheet_count": 4,
        "rows_per_sheet": 48,
        "columns_per_sheet": 48,
        "one_percent_update_count": 93,
        "source_members": {
            "workbook": "xl/workbook.xml",
            "worksheets": [
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
                "xl/worksheets/sheet4.xml",
            ],
            "shared_strings": None,
            "styles": "xl/styles.xml",
        },
    },
}
_XLSX_REPEAT_STORE_OVERSIZED_CORPUS: dict[str, Any] = {
    "name": "xlsx-source-repeated-store-oversized",
    "generator": "litchi-xlsx-source-repeated-store-corpus-v1",
    "package_format": "XLSX/OPC/ZIP",
    "shape": "oversized",
    "payload_kind": "fixed-medium-grid-with-oversized-selected-worksheet",
    "compression": "deflate",
    "entry_count": 9216,
    "archive_member_count": 17,
    "entry_bytes": 4,
    "uncompressed_payload_bytes": 12_789_836,
    "archive_bytes": 4_236_114,
    "archive_sha256": "3cf797e44ef51189a4b62d040cf39ff2af670ebd909c6e806f387b51e72ecfec",
    "target_entry": "Sheet1!A1",
    "target_payload_bytes": 1,
    "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
    "xlsx": {
        "sheet_count": 4,
        "rows_per_sheet": 48,
        "columns_per_sheet": 48,
        "one_percent_update_count": 93,
        "source_members": {
            "workbook": "xl/workbook.xml",
            "worksheets": [
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
                "xl/worksheets/sheet4.xml",
            ],
            "shared_strings": None,
            "styles": "xl/styles.xml",
        },
    },
}
# These are deliberately distinct identities: the filesystem sample hashes the
# full workbook, while the repeated-store evidence hashes only its query
# projection.  They must never be compared for equality.
_XLSX_REPEAT_STORE_FULL_SEMANTIC_SHA256 = (
    "020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e"
)
_XLSX_REPEAT_STORE_PROJECTION_SHA256 = (
    "01c253bf3fc611835e0806414c6417a9cfbb012ff6e01f9bb55cec94236a6235"
)

# The DOCX section-layout selector uses one fixed, non-CLI corpus.  Keep its
# static manifest and section identity separate from the compiled smoke pins.
DOCX_SECTION_LAYOUT_CASE = "docx_source_backed_existing_section_layout_edit_save"
DOCX_SECTION_LAYOUT_FIXED_CORPUS_FIELDS: dict[str, Any] = {
    "name": "docx-source-backed-existing-section-layout",
    "generator": "litchi-docx-section-layout-source-v1",
    "package_format": "DOCX/OPC/ZIP",
    "shape": "existing-section-layout",
    "payload_kind": "deterministic-incompressible-media",
    "compression": "deflate",
    "entry_count": 7,
    "archive_member_count": 10,
    "entry_bytes": 512 * 1024,
    "target_entry": "word/document.xml",
    "xlsx": None,
}
DOCX_SECTION_LAYOUT_CORPUS_IDENTITY: dict[str, Any] = {
    **DOCX_SECTION_LAYOUT_FIXED_CORPUS_FIELDS,
    # Pinned from the compiled change-0273 smoke report.
    "uncompressed_payload_bytes": 2_116_441,
    "archive_bytes": 2_101_427,
    "archive_sha256": "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8",
    "target_payload_bytes": 18_855,
    "target_payload_sha256": "67cb9066f112bfee6905a23f3eabdc76b65c86cd0c6613dd60a5a18895ceac81",
}
DOCX_SECTION_LAYOUT_HASH_FIELDS = (
    "uncompressed_payload_bytes",
    "archive_bytes",
    "archive_sha256",
    "target_payload_bytes",
    "target_payload_sha256",
)
# Pinned from the deterministic source generator used by the compiled
# change-0273 smoke report. The report schema exposes the archive binding but
# not individual media members, so these values are independently derived
# from the four generator payloads.
DOCX_SECTION_LAYOUT_MEDIA_PAYLOAD_SHA256: tuple[str | None, ...] = (
    "e9f29948f8df63fcf0239ac1302d0da879ffc1025f62d87ab233ac31ce823e5e",
    "dd755641a26e08aea79988769fe1240d309e2467d7f57b1351647f78ed814e20",
    "68691a8c911bdd6660d475f31e80bce46ed06f3f38037e698c270d93ecb7b86c",
    "7f21c0cb8c0dee91242cf66e9c72feeb21c89422ce9e6a8efeff7c335c7e9417",
)
# Pinned from the deterministic changed-publication output in the compiled
# change-0273 smoke report.
DOCX_SECTION_LAYOUT_EXPECTED_OUTPUT_SHA256: str | None = (
    "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
)
DOCX_SECTION_LAYOUT_EXPECTED_OUTPUT_BYTES = 2_101_446
DOCX_SECTION_LAYOUT_HASHES_PENDING = False

# These selectors use a fixed, non-CLI corpus.  Its shape therefore cannot be
# listed in any of the configurable ``*_shapes`` fields emitted by the
# historical schema-1 harness.  Keep this exception exact and case-local so an
# unknown generator or an accidentally substituted corpus still fails closed.
FIXED_CASE_CORPUS_IDENTITIES: dict[str, dict[str, Any]] = {
    "ods_source_backed_one_edit_save": {
        "name": "ods-media-publication",
        "generator": "litchi-ods-media-publication-v1",
        "shape": "media-rich",
    },
    "ods_source_backed_one_percent_edit_save": {
        "name": "ods-media-publication",
        "generator": "litchi-ods-media-publication-v1",
        "shape": "media-rich",
    },
    "xlsx_eager_page_break_edit_save": {
        "name": "xlsx-page-break-media",
        "generator": "litchi-xlsx-page-break-source-edit-media-v1",
        "shape": "media-rich",
    },
    "xlsx_source_backed_page_break_edit_save": {
        "name": "xlsx-page-break-media",
        "generator": "litchi-xlsx-page-break-source-edit-media-v1",
        "shape": "media-rich",
    },
    "xlsx_source_repeated_store_medium": _XLSX_REPEAT_STORE_MEDIUM_CORPUS,
    "xlsx_source_repeated_store_medium_reacquisition_control": (
        _XLSX_REPEAT_STORE_MEDIUM_CORPUS
    ),
    "xlsx_source_repeated_store_oversized": _XLSX_REPEAT_STORE_OVERSIZED_CORPUS,
    "xlsx_source_repeated_store_oversized_reacquisition_control": (
        _XLSX_REPEAT_STORE_OVERSIZED_CORPUS
    ),
    "xls_numeric_eager_number_edit_save": {
        "name": "xls-comments-opaque-heavy",
        "generator": "litchi-xls-comments-opaque-heavy-v1",
        "shape": "256-comments-opaque-heavy",
    },
    "xls_numeric_source_backed_number_edit_save": {
        "name": "xls-comments-opaque-heavy",
        "generator": "litchi-xls-comments-opaque-heavy-v1",
        "shape": "256-comments-opaque-heavy",
    },
    "xls_numeric_plan_only_number_edit_save": {
        "name": "xls-comments-opaque-heavy",
        "generator": "litchi-xls-comments-opaque-heavy-v1",
        "shape": "256-comments-opaque-heavy",
    },
    "xls_numeric_eager_rk_mulrk_edit_save": {
        "name": "xls-rk-mulrk-deterministic",
        "generator": "litchi-xls-rk-mulrk-publication-v1",
        "shape": "one-rk-one-mulrk",
    },
    "xls_numeric_source_backed_rk_mulrk_edit_save": {
        "name": "xls-rk-mulrk-deterministic",
        "generator": "litchi-xls-rk-mulrk-publication-v1",
        "shape": "one-rk-one-mulrk",
    },
    "xls_numeric_plan_only_rk_mulrk_edit_save": {
        "name": "xls-rk-mulrk-deterministic",
        "generator": "litchi-xls-rk-mulrk-publication-v1",
        "shape": "one-rk-one-mulrk",
    },
    DOCX_SECTION_LAYOUT_CASE: DOCX_SECTION_LAYOUT_FIXED_CORPUS_FIELDS,
}

OPC_SOURCE_OVERLAY_MULTI_PART_CASES = frozenset(
    {
        "opc_source_overlay_multi_part_changed",
        "opc_source_overlay_multi_part_noop",
        "opc_source_overlay_multi_part_mixed",
    }
)
OPC_SOURCE_OVERLAY_MULTI_PART_COUNTS = (2, 8, 32)
# These physical corpus identity values are pinned from the compiled
# multi-Part changed smoke report at
# /dev/shm/opc-source-overlay-multi-part-changed-smoke/report.json.
OPC_SOURCE_OVERLAY_MULTI_PART_HASHES_PENDING = False
OPC_SOURCE_OVERLAY_MULTI_PART_CORPUS_IDENTITIES: dict[str, dict[str, Any]] = {
    "overlay-small": {
        "name_prefix": "overlay-small-compressible-count-",
        "generator": "litchi-opc-source-overlay-multi-part-v1",
        "package_format": "OPC/ZIP",
        "shape": "overlay-small",
        "payload_kind": "compressible",
        "compression": "deflate",
        "entry_count": 32,
        "archive_member_count": 34,
        "entry_bytes": 1024,
        "uncompressed_payload_bytes": 32 * 1024,
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 1024,
        "archive_bytes": 7451,
        "archive_sha256": "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509",
        "target_payload_sha256": "5b7b9793a43d08ca2c0670289d932541377407ed352ab9f6c145f63d19de9f98",
    },
    "overlay-large": {
        "name_prefix": "overlay-large-incompressible-count-",
        "generator": "litchi-opc-source-overlay-multi-part-v1",
        "package_format": "OPC/ZIP",
        "shape": "overlay-large",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 32,
        "archive_member_count": 34,
        "entry_bytes": 64 * 1024,
        "uncompressed_payload_bytes": 2 * 1024 * 1024,
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 64 * 1024,
        "archive_bytes": 2103195,
        "archive_sha256": "8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73",
        "target_payload_sha256": "e17b543eec6b4d3534978d7d59e7240dcbf0f2a2050fd80f32ea3daec266aa73",
    },
    "overlay-media-incompressible": {
        "name_prefix": "overlay-media-incompressible-incompressible-count-",
        "generator": "litchi-opc-source-overlay-multi-part-v1",
        "package_format": "OPC/ZIP",
        "shape": "overlay-media-incompressible",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 32,
        "archive_member_count": 34,
        "entry_bytes": 256 * 1024,
        "uncompressed_payload_bytes": 8 * 1024 * 1024,
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 256 * 1024,
        "archive_bytes": 8396580,
        "archive_sha256": "bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883",
        "target_payload_sha256": "3ad07c7e34d3dd6d9ff75b696ccbdd702777b6e4dea04b19bbe3d0aa6d21cdeb",
    },
}
OPC_SOURCE_OVERLAY_MULTI_PART_EXPECTED_EAGER_SHA256: dict[
    tuple[str, str, int], str
] = {
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-small",
        2,
    ): "980abb9efebb5f53105f304e1b70b7497882e366558de0e7739766c9b6e7e083",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-small",
        8,
    ): "fee10334d108aa7262bbe404e8f8848fa5bf9f11edbc63e1c1599e52b3fbe54a",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-small",
        32,
    ): "bdc0f47ca45be8e3654f4f21f93cec12dc505c2d45a0a308273cc24f4aa60cc2",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-small",
        2,
    ): "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-small",
        8,
    ): "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-small",
        32,
    ): "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-small",
        2,
    ): "387831c00e9d5f882c7f61209f21a2eda7a30ce3995925c252c80c0b7dedb169",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-small",
        8,
    ): "820b0ea99e9a5ee31b55acfe78340994e5b46f6e8337e6d7fd6d1a811a59491c",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-small",
        32,
    ): "f992e1df7ee8fbe235c2c46ccc8dab2c88dbaf328e60b11c563ed62b6c1a6adc",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-large",
        2,
    ): "adcf54d3e4f9a8c35190e0b9187fafff7e94aa1992afd0c410d45387fe925ec0",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-large",
        8,
    ): "ad37dc59d476bd09e0f5e4839c95b818ff8e3c9af315c8d3c4d246c3b1f5fea8",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-large",
        32,
    ): "4279caf9d8c0c0df131e392763c1c1fb9ec41b081999c4a50f381b03b55abb10",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-large",
        2,
    ): "8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-large",
        8,
    ): "8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-large",
        32,
    ): "8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-large",
        2,
    ): "90fdf6163d11b3e32ed9356a9fc82fc35448d40281c2f6b59013d92ff3c9fa91",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-large",
        8,
    ): "52e10126ef1972b1b346f510a8a46db81584a41f7505add7deeb84a4751e8ad0",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-large",
        32,
    ): "772f8360dcbdaa6d3146894e826fdb8f2d082b098da887b851f89a143289a5fc",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-media-incompressible",
        2,
    ): "1bed16c8ef6421f058535372802d31ecf1a91243be5f6048d2a1d17030e0c465",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-media-incompressible",
        8,
    ): "7458528265a57f89135eeec420053690b8a2dad666f0886bfb79be7b69859564",
    (
        "opc_source_overlay_multi_part_changed",
        "overlay-media-incompressible",
        32,
    ): "49984303765eed16b8cda0db34fe952a8beb97ee1ef38a2a29a3df0e93f65730",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-media-incompressible",
        2,
    ): "bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-media-incompressible",
        8,
    ): "bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883",
    (
        "opc_source_overlay_multi_part_noop",
        "overlay-media-incompressible",
        32,
    ): "bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-media-incompressible",
        2,
    ): "9852986f7addf22c22766587912b05951b8fd32c6f339f24a693dda88e7c4386",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-media-incompressible",
        8,
    ): "dee30e2f5ebaec5dc1e1fba871d789c5a6d6b5568a7883a7f85cbdef525d56a6",
    (
        "opc_source_overlay_multi_part_mixed",
        "overlay-media-incompressible",
        32,
    ): "cf178d5a31e2bda0aa3723c7085b6d913acb24229ce6e80041a0a0a8455c085f",
}

_MISSING = object()
_FILESYSTEM_EVIDENCE_KEYS = frozenset(
    {
        "case",
        "corpus",
        "warmup_iterations",
        "sample_count",
        "cache_states",
        "fresh_child_per_sample",
        "samples",
        "cold_verified_status",
        "cold_verified_samples",
        "cold_verified_claim_scope",
        "cold_verified_fincore_command",
        "cfb_owned",
        "tool",
        "configuration",
        "config",
    }
)
_FILESYSTEM_SAMPLE_KEYS = frozenset(
    {
        "sample_index",
        "cache_state",
        "child_process_id",
        "elapsed_ns",
        "parent_wall_ns",
        "cold_advice",
        "cold_verified",
        "logical_read_counter_scope",
        "logical_read_calls",
        "logical_read_requested_bytes",
        "logical_read_bytes",
        "logical_read_largest_requested_bytes",
        "logical_read_largest_returned_bytes",
        "logical_read_pattern",
        "max_concurrent_reads",
        "logical_read_request_sizes",
        "logical_read_request_size_buckets",
        "process_metrics",
        "allocation_metrics",
        "output_sha256",
        "output_bytes",
        "opc_materialized_parts",
        "cfb_changed_spans",
        "cfb_published_bytes",
        "cfb_phases",
        "pptx_source_replay",
        "docx_source_replay",
        "xlsx_source_sha256",
        "xlsx_semantic_sha256",
        "xlsx_repeat_store",
    }
)
_FILESYSTEM_SAMPLE_IDENTITY_KEYS = frozenset(
    {
        "sample_index",
        "cache_state",
        "cold_advice",
        "logical_read_counter_scope",
        "output_sha256",
        "xlsx_source_sha256",
        "xlsx_semantic_sha256",
    }
)
_FILESYSTEM_EVIDENCE_REQUIRED_KEYS = frozenset(
    {
        "case",
        "corpus",
        "warmup_iterations",
        "sample_count",
        "cache_states",
        "fresh_child_per_sample",
        "samples",
    }
)
_FILESYSTEM_SAMPLE_REQUIRED_KEYS = frozenset(
    {
        "sample_index",
        "cache_state",
        "elapsed_ns",
        "parent_wall_ns",
        "cold_advice",
        "logical_read_counter_scope",
        "logical_read_calls",
        "logical_read_requested_bytes",
        "logical_read_bytes",
        "max_concurrent_reads",
        "logical_read_request_sizes",
        "logical_read_request_size_buckets",
        "process_metrics",
        "output_sha256",
        "output_bytes",
        "opc_materialized_parts",
        "cfb_changed_spans",
        "cfb_published_bytes",
    }
)
# These counters were added to schema 1 after the first filesystem reports
# were emitted.  Schema 1 deliberately permits additive evidence, so an old
# sample may omit the pair.  A sample may never provide only one member: a
# partial pair would make the evidence ambiguous and is rejected below.
_FILESYSTEM_RANGE_SIZE_KEYS = frozenset(
    {
        "logical_read_largest_requested_bytes",
        "logical_read_largest_returned_bytes",
    }
)
# The operation-metrics envelope also grew additively while report
# ``schema_version`` remained 1.  These are the exact keys emitted before the
# aligned sample-index, latency-claim, range-shape, procfs, and sink-write
# additions.  They are kept separate from the current validator so accepting
# this historical envelope cannot weaken validation of a current envelope.
_LEGACY_OPERATION_METRICS_KEYS = frozenset(
    {
        "sample_count",
        "alignment",
        "source",
        "process",
        "sink",
        "publication",
        "materialization",
        "cfb_phases",
    }
)
_LEGACY_SOURCE_METRICS_KEYS = frozenset(
    {
        "status",
        "counter_scope",
        "logical_read_calls",
        "logical_read_requested_bytes",
        "logical_read_returned_bytes",
        "max_concurrent_reads",
    }
)
_LEGACY_PROCESS_METRICS_KEYS = frozenset(
    {
        "status",
        "user_cpu_ticks",
        "system_cpu_ticks",
        "clock_ticks_per_second",
        "minor_faults",
        "major_faults",
        "voluntary_context_switches",
        "nonvoluntary_context_switches",
        "rss_delta_bytes",
        "peak_rss_bytes",
    }
)
_LEGACY_SINK_METRICS_KEYS = frozenset({"status", "output_bytes"})
_COLD_ADVICE_VALUES = frozenset(
    ("not_requested", "requested", "unsupported", "failed")
)
_COLD_VERIFIED_STATUS_VALUES = frozenset(
    {
        "eligible",
        "ineligible_non_linux",
        "ineligible_linux_non64_bit",
        "ineligible_filesystem_unknown",
        "ineligible_filesystem_unsupported",
        "ineligible_fincore_unavailable",
        "ineligible_fincore_failed",
        "ineligible_fincore_invalid_json",
        "ineligible_fincore_multiple_records",
        "ineligible_fincore_path_mismatch",
        "ineligible_fincore_size_mismatch",
        "ineligible_fincore_metadata_unavailable",
        "ineligible_fincore_unrecognized_fallback",
        "ineligible_source_not_regular",
        "ineligible_source_empty",
        "ineligible_source_read_write_unavailable",
        "ineligible_source_hash_failed",
        "ineligible_source_page_size_unavailable",
        "ineligible_source_not_page_aligned",
        "ineligible_source_fsync_failed",
        "ineligible_source_advice_failed",
        "ineligible_source_resident",
        "ineligible_source_dirty",
        "ineligible_source_writeback",
        "ineligible_proc_io_unavailable",
        "ineligible_read_bytes_backwards",
        "ineligible_read_bytes_zero",
        "ineligible_prepared_query_control",
        "ineligible_source_alignment_unavailable",
        "ineligible_source_write_failed",
    }
)
_COLD_VERIFIED_SAMPLE_KEYS = frozenset(
    {
        "status",
        "filesystem_magic",
        "page_size_bytes",
        "source_bytes",
        "source_pages",
        "aligned_source_bytes",
        "aligned_source_sha256",
        "fsync_completed",
        "advice",
        "fincore_size_bytes",
        "resident_bytes",
        "dirty_bytes",
        "writeback_bytes",
        "fincore_tool",
        "fincore_sha256",
        "fincore_version",
        "fincore_stderr_sha256",
        "fincore_stderr_bytes",
        "fincore_version_stderr_sha256",
        "fincore_version_stderr_bytes",
        "fincore_method",
        "fincore_fallback",
        "read_bytes_before",
        "read_bytes_after",
        "read_bytes_delta",
    }
)

# XLSX repeated-store evidence is emitted by the filesystem harness after the
# timed operation.  Keep this schema local to the ABBA reader: these fields are
# not ordinary operation metrics and their role is selected by the explicit
# claim scope, never by a Git revision string.
_XLSX_REPEAT_STORE_KEYS = frozenset(
    {
        "implementation",
        "scenario",
        "selected_member",
        "selected_member_uncompressed_bytes",
        "cache_max_bytes",
        "cache_max_entries",
        "query_iterations",
        "query_names",
        "query_elapsed_ns",
        "timed_elapsed_total_ns",
        "control_reacquire_count",
        "timing_scope",
        "claim_scope",
        "budget_managed",
        "budget_memory_limit",
        "budget_input_bytes_limit",
        "budget_work_limit",
        "semantic_projection_sha256",
        "diagnostics_before",
        "diagnostics_after",
        "diagnostics_delta",
    }
)
_XLSX_REPEAT_STORE_COUNTER_KEYS = frozenset(
    {
        "cache_cold_loads",
        "cache_successful_loads",
        "cache_bypasses",
        "cache_oversized_bypasses",
        "cache_evictions",
        "source_read_calls",
        "source_read_bytes",
        "selected_member_read_calls",
        "selected_member_read_bytes",
        "budget_input_bytes_used",
        "budget_work_used",
    }
)
_XLSX_REPEAT_STORE_QUERY_NAMES = ("cell", "cells", "visit", "stored_extent")
_XLSX_REPEAT_STORE_CACHE_LIMITS: dict[str, tuple[int, int]] = {
    "medium": (8 * 1024 * 1024, 2),
    "oversized": (8 * 1024 * 1024, 128),
}
_XLSX_REPEAT_STORE_SELECTED_MEMBER = "xl/worksheets/sheet1.xml"
_XLSX_REPEAT_STORE_QUERY_ITERATIONS = 8
_XLSX_REPEAT_STORE_TIMING_SCOPE = (
    "semantic_query_only; explicit PartData reacquisition excluded"
)
_XLSX_REPEAT_STORE_PRIMARY_IMPLEMENTATION = "source_backed_cached_store"
_XLSX_REPEAT_STORE_STRUCTURAL_IMPLEMENTATION = (
    "explicit_part_data_reacquisition_structural_control"
)
_XLSX_REPEAT_STORE_PRIMARY_CLAIM_SCOPE = (
    "primary repeated-query selector; compare only the same selector across A/B revisions"
)
_XLSX_REPEAT_STORE_STRUCTURAL_CLAIM_SCOPE = (
    "structural cache/read control only; elapsed/query_ns must not be compared with candidate"
)
_XLSX_REPEAT_STORE_PRIMARY_SCOPE = "primary"
_XLSX_REPEAT_STORE_STRUCTURAL_SCOPE = "structural"
_XLSX_REPEAT_STORE_COUNTER_FIELDS = tuple(sorted(_XLSX_REPEAT_STORE_COUNTER_KEYS))
_XLSX_REPEAT_STORE_CASE_CONTRACT: dict[str, dict[str, Any]] = {
    "xlsx_source_repeated_store_medium": {
        "role": _XLSX_REPEAT_STORE_PRIMARY_SCOPE,
        "scenario": "medium",
        "selected_member_uncompressed_bytes": 63_294,
        "full_semantic_sha256": _XLSX_REPEAT_STORE_FULL_SEMANTIC_SHA256,
        "semantic_projection_sha256": _XLSX_REPEAT_STORE_PROJECTION_SHA256,
    },
    "xlsx_source_repeated_store_medium_reacquisition_control": {
        "role": _XLSX_REPEAT_STORE_STRUCTURAL_SCOPE,
        "scenario": "medium",
        "selected_member_uncompressed_bytes": 63_294,
        "full_semantic_sha256": _XLSX_REPEAT_STORE_FULL_SEMANTIC_SHA256,
        "semantic_projection_sha256": _XLSX_REPEAT_STORE_PROJECTION_SHA256,
    },
    "xlsx_source_repeated_store_oversized": {
        "role": _XLSX_REPEAT_STORE_PRIMARY_SCOPE,
        "scenario": "oversized",
        "selected_member_uncompressed_bytes": 8_389_041,
        "full_semantic_sha256": _XLSX_REPEAT_STORE_FULL_SEMANTIC_SHA256,
        "semantic_projection_sha256": _XLSX_REPEAT_STORE_PROJECTION_SHA256,
    },
    "xlsx_source_repeated_store_oversized_reacquisition_control": {
        "role": _XLSX_REPEAT_STORE_STRUCTURAL_SCOPE,
        "scenario": "oversized",
        "selected_member_uncompressed_bytes": 8_389_041,
        "full_semantic_sha256": _XLSX_REPEAT_STORE_FULL_SEMANTIC_SHA256,
        "semantic_projection_sha256": _XLSX_REPEAT_STORE_PROJECTION_SHA256,
    },
}

_XLSX_REPEAT_STORE_CORPUS_GENERATOR = "litchi-xlsx-source-repeated-store-corpus-v1"
_XLSX_REPEAT_STORE_CORPUS_NAMES = frozenset(
    {
        _XLSX_REPEAT_STORE_MEDIUM_CORPUS["name"],
        _XLSX_REPEAT_STORE_OVERSIZED_CORPUS["name"],
    }
)


class AbbaSummaryInputError(ValueError):
    """Raised when ABBA reports are not safely comparable."""


def _validate_json_tree(value: Any, location: str) -> None:
    """Reject values that cannot be represented by strict JSON."""

    if value is None or isinstance(value, (bool, str, int)):
        return
    if isinstance(value, float):
        if not math.isfinite(value):
            raise AbbaSummaryInputError(f"{location} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _validate_json_tree(item, f"{location}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                raise AbbaSummaryInputError(f"{location} has a non-string object key")
            _validate_json_tree(item, f"{location}.{key}")
        return
    raise AbbaSummaryInputError(
        f"{location} contains unsupported JSON value {type(value).__name__}"
    )


def _canonical_json(value: Any, location: str) -> str:
    """Return compact canonical JSON, rejecting values JSON cannot preserve."""

    _validate_json_tree(value, location)
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError, OverflowError) as error:
        raise AbbaSummaryInputError(f"{location} is not canonical JSON: {error}") from error


def _canonical_sha256(value: Any, location: str) -> str:
    """Hash canonical JSON without materializing a second report-sized string."""

    _validate_json_tree(value, location)
    digest = hashlib.sha256()

    def emit(item: Any) -> None:
        if isinstance(item, dict):
            digest.update(b"{")
            for index, key in enumerate(sorted(item)):
                if index:
                    digest.update(b",")
                digest.update(
                    json.dumps(key, separators=(",", ":"), allow_nan=False).encode(
                        "utf-8"
                    )
                )
                digest.update(b":")
                emit(item[key])
            digest.update(b"}")
            return
        if isinstance(item, list):
            digest.update(b"[")
            for index, child in enumerate(item):
                if index:
                    digest.update(b",")
                emit(child)
            digest.update(b"]")
            return
        digest.update(
            json.dumps(item, sort_keys=True, separators=(",", ":"), allow_nan=False).encode(
                "utf-8"
            )
        )

    emit(value)
    return digest.hexdigest()


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise AbbaSummaryInputError(f"{location} must be an object")
    return value


def _finite_number(value: Any, location: str, *, positive: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise AbbaSummaryInputError(f"{location} must be a number")
    try:
        number = float(value)
    except (OverflowError, ValueError) as error:
        raise AbbaSummaryInputError(f"{location} is outside the finite range") from error
    if not math.isfinite(number):
        raise AbbaSummaryInputError(f"{location} must be finite")
    if positive and number <= 0:
        raise AbbaSummaryInputError(f"{location} must be positive")
    return number


def _u64(value: Any, location: str, *, positive: bool = False) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise AbbaSummaryInputError(f"{location} must be an unsigned 64-bit integer")
    if value < (1 if positive else 0) or value > U64_MAX:
        requirement = "positive " if positive else ""
        raise AbbaSummaryInputError(
            f"{location} must be a {requirement}unsigned 64-bit integer"
        )
    return value


def _u32(value: Any, location: str, *, positive: bool = False) -> int:
    """Validate an optional process identifier without making it identity."""

    if isinstance(value, bool) or not isinstance(value, int):
        raise AbbaSummaryInputError(f"{location} must be an unsigned 32-bit integer")
    if value < (1 if positive else 0) or value > U32_MAX:
        requirement = "positive " if positive else ""
        raise AbbaSummaryInputError(
            f"{location} must be a {requirement}unsigned 32-bit integer"
        )
    return value


def _student_t_critical_95(degrees_of_freedom: int) -> float:
    if degrees_of_freedom == 0:
        return 0.0
    if degrees_of_freedom <= len(STUDENT_T_CRITICAL_95):
        return STUDENT_T_CRITICAL_95[degrees_of_freedom - 1]
    degrees = float(degrees_of_freedom)
    z = STUDENT_T_Z_975
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    return (
        z
        + (z3 + z) / (4.0 * degrees)
        + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * degrees * degrees)
        + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
        / (384.0 * degrees * degrees * degrees)
    )


def _float_close(reported: float, expected: float, location: str) -> None:
    tolerance = max(1e-12, abs(expected) * 1e-12)
    if not math.isfinite(reported) or abs(reported - expected) > tolerance:
        raise AbbaSummaryInputError(
            f"{location}={reported} disagrees with samples ({expected})"
        )


def recompute_statistics(elapsed: Any, location: str) -> dict[str, Any]:
    """Validate an ``elapsed_ns`` object using the Rust harness formula."""

    elapsed_object = _require_object(elapsed, location)
    if elapsed_object.get("unit") != "ns":
        raise AbbaSummaryInputError(f"{location}.unit must be 'ns'")
    samples_raw = elapsed_object.get("samples")
    if not isinstance(samples_raw, list) or len(samples_raw) < MIN_RETAINED_SAMPLES:
        raise AbbaSummaryInputError(
            f"{location}.samples must retain at least {MIN_RETAINED_SAMPLES} samples"
        )
    samples = [
        _u64(value, f"{location}.samples[{index}]")
        for index, value in enumerate(samples_raw)
    ]
    if samples != sorted(samples):
        raise AbbaSummaryInputError(f"{location}.samples must be sorted ascending")

    mean = 0.0
    squared_deviation_sum = 0.0
    for index, value in enumerate(samples):
        value_as_float = float(value)
        next_count = float(index + 1)
        delta = value_as_float - mean
        next_mean = mean + delta / next_count
        squared_deviation_sum += delta * (value_as_float - next_mean)
        mean = next_mean
    standard_deviation = (
        math.sqrt(squared_deviation_sum / float(len(samples) - 1))
        if len(samples) > 1
        else 0.0
    )
    margin = (
        _student_t_critical_95(len(samples) - 1)
        * standard_deviation
        / math.sqrt(float(len(samples)))
        if len(samples) > 1
        else 0.0
    )
    left = samples[(len(samples) - 1) // 2]
    right = samples[len(samples) // 2]
    p50 = left // 2 + right // 2 + (left % 2 + right % 2) // 2

    def nearest_rank(percentile: int) -> int:
        index = ((percentile * len(samples) + 99) // 100) - 1
        return samples[min(index, len(samples) - 1)]

    computed_u64 = {
        "min": samples[0],
        "p50": p50,
        "p95": nearest_rank(95),
        "p99": nearest_rank(99),
        "max": samples[-1],
    }
    computed_float = {
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95.lower": max(mean - margin, 0.0),
        "confidence_interval_95.upper": mean + margin,
    }
    for name, expected in computed_u64.items():
        if name not in elapsed_object:
            raise AbbaSummaryInputError(f"{location}.{name} is required")
        reported = _u64(elapsed_object[name], f"{location}.{name}")
        if reported != expected:
            raise AbbaSummaryInputError(
                f"{location}.{name}={reported} disagrees with samples ({expected})"
            )
    for name, expected in computed_float.items():
        field = name.split(".")[-1] if "." in name else name
        if "." in name:
            confidence = _require_object(
                elapsed_object.get("confidence_interval_95"),
                f"{location}.confidence_interval_95",
            )
            value = confidence.get(field)
        else:
            if name not in elapsed_object:
                raise AbbaSummaryInputError(f"{location}.{name} is required")
            value = elapsed_object[name]
        reported = _finite_number(value, f"{location}.{name}")
        _float_close(reported, expected, f"{location}.{name}")
    confidence = _require_object(
        elapsed_object.get("confidence_interval_95"),
        f"{location}.confidence_interval_95",
    )
    if confidence.get("method") != "two-sided Student's t interval for the mean":
        raise AbbaSummaryInputError(
            f"{location}.confidence_interval_95.method does not match the harness"
        )
    return {
        "sample_count": len(samples),
        **computed_u64,
        **{
            "mean": mean,
            "standard_deviation": standard_deviation,
            "confidence_interval_95": {
                "method": confidence["method"],
                "lower": computed_float["confidence_interval_95.lower"],
                "upper": computed_float["confidence_interval_95.upper"],
            },
        },
    }


def _result_key(result: Any, location: str) -> tuple[str, str, dict[str, Any]]:
    row = _require_object(result, location)
    case = row.get("case")
    if not isinstance(case, str) or not case:
        raise AbbaSummaryInputError(f"{location}.case must be a non-empty string")
    corpus = _require_object(row.get("corpus"), f"{location}.corpus")
    if not corpus:
        raise AbbaSummaryInputError(f"{location}.corpus must not be empty")
    corpus_identity = _canonical_json(corpus, f"{location}.corpus")
    return case, corpus_identity, row


def _index_results(report: dict[str, Any], label: str) -> dict[tuple[str, str], dict[str, Any]]:
    results = report.get("results")
    if not isinstance(results, list) or not results:
        raise AbbaSummaryInputError(f"{label}.results must be a non-empty list")
    indexed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, result in enumerate(results):
        case, corpus_identity, row = _result_key(result, f"{label}.results[{index}]")
        key = (case, corpus_identity)
        if key in indexed:
            raise AbbaSummaryInputError(
                f"{label}.results contains duplicate case/corpus identity for {case!r}"
            )
        indexed[key] = row
    return indexed


def _required_nonempty_string(
    value: Any, location: str, field: str, *, allow_none: bool = False
) -> str | None:
    if allow_none and value is None:
        return None
    if not isinstance(value, str) or not value:
        suffix = " or null" if allow_none else ""
        raise AbbaSummaryInputError(f"{location}.{field} must be a non-empty string{suffix}")
    return value


def _optional_string(value: Any, location: str, field: str) -> str | None:
    if value is not None and not isinstance(value, str):
        raise AbbaSummaryInputError(f"{location}.{field} must be a string or null")
    return value


def _required_bool(value: Any, location: str, field: str) -> bool:
    if not isinstance(value, bool):
        raise AbbaSummaryInputError(f"{location}.{field} must be a boolean")
    return value


def _required_positive_integer(value: Any, location: str, field: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise AbbaSummaryInputError(f"{location}.{field} must be a positive integer")
    return value


def _required_string_list(value: Any, location: str, field: str) -> list[str]:
    if (
        not isinstance(value, list)
        or not value
        or any(not isinstance(item, str) or not item for item in value)
        or len(value) != len(set(value))
    ):
        raise AbbaSummaryInputError(
            f"{location}.{field} must be a non-empty list of unique strings"
        )
    return value


def detect_report_profile(report: Any, label: str = "report") -> str:
    """Detect the report profile from raw report metadata only.

    The historical schema-1 harness reports predate executable identity.  A
    report carrying either additive current tool field (or the root binary
    identity object) is current; otherwise it must carry the complete legacy
    tool identity.  Callers comparing ABBA legs must run this for every raw
    report and reject a mixed set before consulting a summary.
    """

    root = _require_object(report, label)
    tool = _require_object(root.get("tool"), f"{label}.tool")
    if "binary" in tool or "instrumentation" in tool or "binary_identity" in root:
        return REPORT_PROFILE_CURRENT
    legacy_fields = {"name", "version", "profile", "target_os", "target_arch"}
    if legacy_fields <= set(tool):
        return REPORT_PROFILE_LEGACY
    raise AbbaSummaryInputError(
        f"{label} does not match a supported ABBA report profile"
    )


def detect_reports_profile(
    reports: Mapping[str, Any] | Sequence[Mapping[str, Any]],
) -> str:
    """Detect one common profile across all four raw reports."""

    report_values = _coerce_reports(reports, None, None, None)
    profiles = {
        detect_report_profile(report, label) for label, report in zip(LEG_ORDER, report_values)
    }
    if len(profiles) != 1:
        raise AbbaSummaryInputError("mixed legacy-v1/current-v1 ABBA report profiles")
    return next(iter(profiles))


def _validate_tool(
    tool: dict[str, Any], label: str, profile: str = REPORT_PROFILE_CURRENT
) -> None:
    if profile not in REPORT_PROFILES:
        raise AbbaSummaryInputError(f"unsupported ABBA report profile {profile!r}")
    required_fields = (
        REQUIRED_TOOL_FIELDS
        if profile == REPORT_PROFILE_CURRENT
        else ("name", "version", "profile", "target_os", "target_arch")
    )
    for field in required_fields:
        _required_nonempty_string(tool.get(field), f"{label}.tool", field)
    if tool["name"] != HARNESS_TOOL_NAME:
        raise AbbaSummaryInputError(
            f"{label}.tool.name must be {HARNESS_TOOL_NAME!r}"
        )
    if profile == REPORT_PROFILE_CURRENT:
        if tool["binary"] != HARNESS_TOOL_NAME:
            raise AbbaSummaryInputError(
                f"{label}.tool.binary must be {HARNESS_TOOL_NAME!r} for latency ABBA"
            )
        if tool["instrumentation"] != "none":
            raise AbbaSummaryInputError(
                f"{label}.tool.instrumentation must be 'none' for latency ABBA"
            )


def _validate_binary_identity(
    binary: Any, label: str, tool: Mapping[str, Any]
) -> dict[str, Any]:
    """Validate the untimed identity of the executable producing a report.

    The descriptor follows ``perf_resource_profile.binary_identity``.  Paths
    are provenance only and are not resolved against the reader's filesystem;
    the hash, byte count, mode bits, executable marker, and profile are the
    portable syntax/identity channels.  Unix mode bits are explicitly null on
    platforms without a portable permission-bit representation.
    """

    location = f"{label}.binary_identity"
    value = _require_object(binary, location)
    missing = set(REQUIRED_BINARY_IDENTITY_FIELDS) - set(value)
    if missing:
        raise AbbaSummaryInputError(
            f"{location} is missing required keys: {sorted(missing)}"
        )
    unknown = set(value) - set(REQUIRED_BINARY_IDENTITY_FIELDS) - OPTIONAL_BINARY_IDENTITY_FIELDS
    if unknown:
        raise AbbaSummaryInputError(f"{location} has unknown keys: {sorted(unknown)}")
    path = value["path"]
    if not isinstance(path, str) or not path:
        raise AbbaSummaryInputError(f"{location}.path must be a non-empty string")
    # Accept both POSIX and Windows absolute paths while keeping the path
    # itself opaque: reports may be compared on a different host.
    if not (path.startswith("/") or re.match(r"^[A-Za-z]:[\\/]", path) or path.startswith("\\\\")):
        raise AbbaSummaryInputError(f"{location}.path must be absolute")
    digest = _validate_output_sha256(value["binary_sha256"], f"{location}.binary_sha256")
    binary_bytes = value["binary_bytes"]
    if (
        isinstance(binary_bytes, bool)
        or not isinstance(binary_bytes, int)
        or binary_bytes <= 0
        or binary_bytes > (1 << 64) - 1
    ):
        raise AbbaSummaryInputError(
            f"{location}.binary_bytes must be a positive unsigned integer"
        )
    mode_bits = value["mode_bits"]
    if mode_bits is not None and (
        isinstance(mode_bits, bool)
        or not isinstance(mode_bits, int)
        or mode_bits < 0
        or mode_bits > 0o7777
        or mode_bits & 0o111 == 0
    ):
        raise AbbaSummaryInputError(
            f"{location}.mode_bits must be null or executable permission bits"
        )
    if tool["target_os"] in UNIX_TARGET_OSES and mode_bits is None:
        raise AbbaSummaryInputError(
            f"{location}.mode_bits must be present for Unix targets"
        )
    if value["executable"] is not True:
        raise AbbaSummaryInputError(f"{location}.executable must be true")
    profile = value["profile"]
    if not isinstance(profile, str) or not profile:
        raise AbbaSummaryInputError(f"{location}.profile must be a non-empty string")
    if profile != tool["profile"]:
        raise AbbaSummaryInputError(
            f"{location}.profile does not match {label}.tool.profile"
        )
    if "label" in value:
        _required_nonempty_string(value["label"], location, "label")
    return {
        **value,
        "binary_sha256": digest,
        "binary_bytes": binary_bytes,
        "mode_bits": mode_bits,
        "executable": True,
        "profile": profile,
    }


def _validate_environment(environment: dict[str, Any], label: str) -> None:
    location = f"{label}.environment"
    for field in REQUIRED_ENVIRONMENT_FIELDS:
        if field not in environment:
            raise AbbaSummaryInputError(f"{location}.{field} is required")
    _required_nonempty_string(environment["rustc_version"], location, "rustc_version")
    revision = _required_nonempty_string(environment["git_revision"], location, "git_revision")
    if revision is None:
        raise AbbaSummaryInputError(f"{location}.git_revision must not be null")
    if _required_bool(environment["git_worktree_dirty"], location, "git_worktree_dirty"):
        raise AbbaSummaryInputError(f"{location}.git_worktree_dirty must be false")
    _required_positive_integer(
        environment["logical_cpus_available"], location, "logical_cpus_available"
    )
    _required_nonempty_string(environment["allocator"], location, "allocator")
    for field in ("rustflags", "cargo_build_target", "perf_event_paranoid"):
        _optional_string(environment[field], location, field)
    for field in (
        "os",
        "kernel",
        "cpu_model",
        "filesystem_type",
        "cpu_affinity",
        "storage_identifier",
    ):
        if field in environment:
            _optional_string(environment[field], location, field)
    for field in ("total_memory_bytes", "page_size_bytes"):
        if field in environment and environment[field] is not None:
            _u64(environment[field], f"{location}.{field}")
    if "source_destination_same_device" in environment and environment[
        "source_destination_same_device"
    ] is not None:
        _required_bool(
            environment["source_destination_same_device"],
            location,
            "source_destination_same_device",
        )


def _validate_configuration(configuration: dict[str, Any], label: str) -> None:
    location = f"{label}.configuration"
    for field in REQUIRED_CONFIGURATION_FIELDS:
        if field not in configuration:
            raise AbbaSummaryInputError(f"{location}.{field} is required")
    samples = _required_positive_integer(
        configuration["samples_per_case"], location, "samples_per_case"
    )
    if samples < MIN_RETAINED_SAMPLES:
        raise AbbaSummaryInputError(
            f"{location}.samples_per_case must be at least {MIN_RETAINED_SAMPLES}"
        )
    _required_positive_integer(
        configuration["warmup_iterations_per_case"],
        location,
        "warmup_iterations_per_case",
    )
    _required_string_list(configuration["cases"], location, "cases")
    _required_string_list(configuration["corpus_shapes"], location, "corpus_shapes")
    for field, value in configuration.items():
        if field.endswith("_shapes") and field != "corpus_shapes":
            _required_string_list(value, location, field)
    if "filesystem_root_selected" in configuration:
        _required_bool(
            configuration["filesystem_root_selected"],
            location,
            "filesystem_root_selected",
        )


def _operation_metrics_identity_projection(value: Any) -> Any:
    """Remove candidate-sensitive vector values while retaining their shape."""

    if isinstance(value, dict):
        keys = set(value)
        if {"status", "scope"} <= keys <= {"status", "scope", "values"}:
            return {
                "status": value["status"],
                "scope": value["scope"],
                "values": (
                    "null"
                    if "values" in value and value["values"] is None
                    else "present"
                    if "values" in value
                    else "absent"
                ),
            }
        source = value.get("source")
        normalize_sample_indices = (
            isinstance(source, dict)
            and source.get("counter_scope") == "not_applicable_filesystem_xlsx"
        )
        return {
            key: (
                "<permutation>"
                if normalize_sample_indices and key == "sample_indices"
                else _operation_metrics_identity_projection(item)
            )
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [_operation_metrics_identity_projection(item) for item in value]
    return value


def _validate_legacy_operation_metrics(
    value: Any, path: str, elapsed_samples: list[Any], report_schema: int
) -> None:
    """Validate the pre-additive schema-1 operation-metrics envelope.

    This path is selected only for the exact historical key shape above.  It
    deliberately does not synthesize any of the newer vectors or alignment
    metadata from old observations.
    """

    if report_schema != HARNESS_SCHEMA_VERSION:
        raise AbbaSummaryInputError(
            f"{path} legacy validation requires report schema "
            f"{HARNESS_SCHEMA_VERSION}, got {report_schema}"
        )
    try:
        obj = perf_compare._require_exact_keys(
            value, path, set(_LEGACY_OPERATION_METRICS_KEYS)
        )
        sample_count = len(elapsed_samples)
        declared_sample_count = obj["sample_count"]
        if (
            isinstance(declared_sample_count, bool)
            or not isinstance(declared_sample_count, int)
            or declared_sample_count <= 0
        ):
            raise perf_compare.ComparisonInputError(
                f"{path}.sample_count must be a positive integer"
            )
        if declared_sample_count != sample_count:
            raise perf_compare.ComparisonInputError(
                f"{path}.sample_count={declared_sample_count} does not match "
                f"elapsed_ns.samples length {sample_count}"
            )
        if obj["alignment"] != "elapsed_ns.samples":
            raise perf_compare.ComparisonInputError(
                f"{path}.alignment must be 'elapsed_ns.samples'"
            )

        perf_compare._validate_status_group(
            obj["source"],
            f"{path}.source",
            set(_LEGACY_SOURCE_METRICS_KEYS),
            (
                "logical_read_calls",
                "logical_read_requested_bytes",
                "logical_read_returned_bytes",
                "max_concurrent_reads",
            ),
            sample_count,
        )
        source = obj["source"]
        if not isinstance(source["counter_scope"], str) or not source["counter_scope"]:
            raise perf_compare.ComparisonInputError(
                f"{path}.source.counter_scope must be a non-empty string"
            )

        perf_compare._validate_status_group(
            obj["process"],
            f"{path}.process",
            set(_LEGACY_PROCESS_METRICS_KEYS),
            (
                "user_cpu_ticks",
                "system_cpu_ticks",
                "clock_ticks_per_second",
                "minor_faults",
                "major_faults",
                "voluntary_context_switches",
                "nonvoluntary_context_switches",
                "rss_delta_bytes",
                "peak_rss_bytes",
            ),
            sample_count,
        )

        sink = perf_compare._require_exact_keys(
            obj["sink"], f"{path}.sink", set(_LEGACY_SINK_METRICS_KEYS)
        )
        sink_status = perf_compare._validate_metric_status(
            sink["status"], f"{path}.sink.status"
        )
        output_status = perf_compare._validate_metric_vector(
            sink["output_bytes"], f"{path}.sink.output_bytes", sample_count
        )
        if output_status != sink_status:
            raise perf_compare.ComparisonInputError(
                f"{path}.sink.status does not match {path}.sink.output_bytes.status"
            )

        perf_compare._validate_status_group(
            obj["publication"],
            f"{path}.publication",
            {"status", "changed_spans", "published_bytes"},
            ("changed_spans", "published_bytes"),
            sample_count,
        )
        perf_compare._validate_status_group(
            obj["materialization"],
            f"{path}.materialization",
            {"status", "opc_parts"},
            ("opc_parts",),
            sample_count,
        )
        phases = perf_compare._require_exact_keys(
            obj["cfb_phases"],
            f"{path}.cfb_phases",
            {"status", "open", "plan", "atomic_publication"},
        )
        phase_status = perf_compare._validate_metric_status(
            phases["status"], f"{path}.cfb_phases.status"
        )
        for phase in ("open", "plan", "atomic_publication"):
            perf_compare._validate_phase_set(
                phases[phase], f"{path}.cfb_phases.{phase}", phase_status, sample_count
            )
    except perf_compare.ComparisonInputError as error:
        raise AbbaSummaryInputError(str(error)) from error


def _operation_metrics_identity(
    row: dict[str, Any], location: str, report_schema: int, *, projected: bool = False
) -> str | None:
    if projected and "_operation_metrics_identity" in row:
        value = row["_operation_metrics_identity"]
        if value is not None and (
            not isinstance(value, str) or not value
        ):
            raise AbbaSummaryInputError(
                f"{location}._operation_metrics_identity must be a string or null"
            )
        return value
    if "operation_metrics" not in row:
        return None
    operation_metrics = row["operation_metrics"]
    if operation_metrics is None:
        raise AbbaSummaryInputError(
            f"{location}.operation_metrics must be omitted or an object"
        )
    elapsed = _require_object(row.get("elapsed_ns"), f"{location}.elapsed_ns")
    samples = elapsed.get("samples")
    if not isinstance(samples, list):
        raise AbbaSummaryInputError(f"{location}.elapsed_ns.samples must be a list")
    operation_location = f"{location}.operation_metrics"
    if set(operation_metrics) == set(_LEGACY_OPERATION_METRICS_KEYS):
        _validate_legacy_operation_metrics(
            operation_metrics, operation_location, samples, report_schema
        )
    else:
        try:
            # The comparator owns the exact operation-metrics schema, including
            # current nested vector and alignment rules.  Keep this call in
            # sync with perf_compare._collect_metrics: it takes the sample
            # list, not a precomputed count.
            elapsed_sample_order = elapsed.get(
                "sample_order", perf_compare._METRIC_SAMPLE_ORDER_MISSING
            )
            perf_compare._validate_operation_metrics(
                operation_metrics,
                operation_location,
                samples,
                report_schema,
                elapsed_sample_order=elapsed_sample_order,
            )
        except perf_compare.ComparisonInputError as error:
            raise AbbaSummaryInputError(str(error)) from error
    projected = _operation_metrics_identity_projection(operation_metrics)
    return _canonical_json(projected, f"{location}.operation_metrics.identity")


def _compare_operation_metrics_identity(
    rows: Mapping[str, dict[str, Any]],
    location: str,
    report_schema: int,
    *,
    projected: bool = False,
) -> tuple[str, str | None]:
    identities = {
        label: _operation_metrics_identity(
            rows[label], f"{label}.{location}", report_schema, projected=projected
        )
        for label in LEG_ORDER
    }
    present = {identity is not None for identity in identities.values()}
    if len(present) != 1:
        raise AbbaSummaryInputError(
            f"{location} operation_metrics presence differs between ABBA legs"
        )
    expected = identities["a1"]
    if expected is None:
        return "consistently_absent", None
    if any(identity != expected for identity in identities.values()):
        raise AbbaSummaryInputError(
            f"{location} operation_metrics identity differs between ABBA legs"
        )
    return "verified_equal", expected


def _filesystem_measurement_shape(value: Any) -> Any:
    """Retain non-numeric identity fields while eliding measurements.

    Numeric values and numeric sample vectors are candidate-sensitive.  Their
    values are therefore replaced with markers, while object keys, enum/string
    labels, and nested object shape remain bound across the four legs.
    """

    if value is None:
        return None
    if isinstance(value, bool):
        return "<boolean>"
    if isinstance(value, (int, float)):
        return "<number>"
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        # Most lists in a sample are candidate-sensitive numeric observations;
        # neither their values nor their cardinality is an implementation
        # identity.  The few structural evidence lists are projected one item
        # at a time by _filesystem_evidence_identity_projection below.
        return "<array>"
    if isinstance(value, dict):
        return {
            key: _filesystem_measurement_shape(item) for key, item in value.items()
        }
    raise AbbaSummaryInputError(
        "filesystem evidence contains unsupported value "
        f"{type(value).__name__}"
    )


def _validate_cold_verified_sample(value: Any, location: str) -> dict[str, Any]:
    sample = _require_object(value, location)
    unknown = set(sample) - _COLD_VERIFIED_SAMPLE_KEYS
    if unknown:
        raise AbbaSummaryInputError(
            f"{location} has unknown keys: {sorted(unknown)}"
        )
    status = sample.get("status")
    if not isinstance(status, str) or status not in _COLD_VERIFIED_STATUS_VALUES:
        raise AbbaSummaryInputError(
            f"{location}.status must be one of {sorted(_COLD_VERIFIED_STATUS_VALUES)}"
        )
    for field, item in sample.items():
        if field in {
            "aligned_source_sha256",
            "fincore_sha256",
            "fincore_stderr_sha256",
            "fincore_version_stderr_sha256",
        }:
            _validate_output_sha256(item, f"{location}.{field}")
        elif field in {"fsync_completed"}:
            _required_bool(item, location, field)
        elif field != "status":
            if isinstance(item, bool) or not isinstance(item, (int, str)):
                raise AbbaSummaryInputError(
                    f"{location}.{field} must be an integer or string"
                )
            if isinstance(item, int):
                _u64(item, f"{location}.{field}")
            elif not item:
                raise AbbaSummaryInputError(f"{location}.{field} must not be empty")
    return sample


def _validate_raw_allocation_metrics(value: Any, location: str) -> dict[str, Any]:
    """Validate the raw allocator sample shape emitted by the harness.

    ``perf_compare`` owns the allocator vocabulary for both operation and raw
    filesystem evidence.  Reuse its exact field list here, while accounting
    for the two compact ``Sample`` forms emitted by the Rust allocator target:
    a measured sample carries every counter and unavailable/overflow samples
    carry only status and scope.
    """

    allocation = _require_object(value, location)
    vector_fields = tuple(perf_compare._ALLOCATOR_VECTOR_FIELDS)
    expected_keys = {"status", "scope", *vector_fields}
    compact_keys = {"status", "scope"}
    keys = set(allocation)
    if keys not in (expected_keys, compact_keys):
        raise AbbaSummaryInputError(
            f"{location} has an invalid allocator schema"
        )
    status = allocation.get("status")
    if status not in {"measured", "unavailable", "overflow"}:
        raise AbbaSummaryInputError(
            f"{location}.status must be measured, unavailable, or overflow"
        )
    if allocation.get("scope") != "operation_global_system_allocator":
        raise AbbaSummaryInputError(
            f"{location}.scope must be 'operation_global_system_allocator'"
        )
    if status == "measured" and keys != expected_keys:
        raise AbbaSummaryInputError(
            f"{location}.measured samples must contain every allocator counter"
        )
    if status != "measured" and keys != compact_keys:
        raise AbbaSummaryInputError(
            f"{location}.{status} samples must omit allocator counters"
        )
    for field in vector_fields:
        if field in allocation:
            _u64(allocation[field], f"{location}.{field}")
    return allocation


def _validate_xlsx_repeat_store_counters(
    value: Any, location: str
) -> dict[str, Any]:
    counters = _require_object(value, location)
    if set(counters) != _XLSX_REPEAT_STORE_COUNTER_KEYS:
        raise AbbaSummaryInputError(
            f"{location} keys mismatch: expected {sorted(_XLSX_REPEAT_STORE_COUNTER_KEYS)}"
        )
    for field in _XLSX_REPEAT_STORE_COUNTER_FIELDS:
        _u64(counters[field], f"{location}.{field}")
    if counters["cache_cold_loads"] > counters["cache_successful_loads"]:
        raise AbbaSummaryInputError(
            f"{location}.cache_cold_loads exceeds cache_successful_loads"
        )
    if counters["cache_bypasses"] > counters["cache_successful_loads"]:
        raise AbbaSummaryInputError(
            f"{location}.cache_bypasses exceeds cache_successful_loads"
        )
    if counters["cache_oversized_bypasses"] > counters["cache_bypasses"]:
        raise AbbaSummaryInputError(
            f"{location}.cache_oversized_bypasses exceeds cache_bypasses"
        )
    if (
        counters["selected_member_read_calls"] > counters["source_read_calls"]
        or counters["selected_member_read_bytes"] > counters["source_read_bytes"]
    ):
        raise AbbaSummaryInputError(
            f"{location} selected-member reads exceed source reads"
        )
    if counters["budget_input_bytes_used"] > counters["source_read_bytes"]:
        raise AbbaSummaryInputError(
            f"{location}.budget_input_bytes_used exceeds source_read_bytes"
        )
    if counters["cache_cold_loads"] and counters["budget_work_used"] == 0:
        raise AbbaSummaryInputError(
            f"{location}.cache_cold_loads requires positive budget_work_used"
        )
    return counters


def _validate_xlsx_repeat_store(
    value: Any,
    location: str,
    *,
    case: str,
    report_label: str,
    corpus: dict[str, Any],
) -> tuple[str, dict[str, Any]]:
    """Validate one exact repeated-store evidence object.

    The explicit claim scope determines whether an object is a primary
    selector or a structural reacquisition control.  This is deliberately
    independent of ``environment.git_revision`` so the verifier remains
    useful for rebased or locally rebuilt ABBA legs.
    """

    contract = _XLSX_REPEAT_STORE_CASE_CONTRACT.get(case)
    if contract is None:
        raise AbbaSummaryInputError(
            f"{location} is not permitted on filesystem case {case!r}"
        )
    expected_corpus = FIXED_CASE_CORPUS_IDENTITIES[case]
    if _canonical_json(corpus, f"{location}.corpus") != _canonical_json(
        expected_corpus, f"{location}.expected_corpus"
    ):
        raise AbbaSummaryInputError(
            f"{location}.corpus does not match the pinned {case} corpus"
        )
    evidence = _require_object(value, location)
    if set(evidence) != _XLSX_REPEAT_STORE_KEYS:
        raise AbbaSummaryInputError(
            f"{location} keys mismatch: expected {sorted(_XLSX_REPEAT_STORE_KEYS)}"
        )
    implementation = _required_nonempty_string(
        evidence.get("implementation"), location, "implementation"
    )
    scenario = _required_nonempty_string(evidence.get("scenario"), location, "scenario")
    if scenario != contract["scenario"]:
        raise AbbaSummaryInputError(
            f"{location}.scenario does not match filesystem case {case!r}"
        )
    selected_member = _required_nonempty_string(
        evidence.get("selected_member"), location, "selected_member"
    )
    if selected_member != _XLSX_REPEAT_STORE_SELECTED_MEMBER:
        raise AbbaSummaryInputError(
            f"{location}.selected_member does not match the pinned worksheet"
        )
    selected_member_bytes = _u64(
        evidence.get("selected_member_uncompressed_bytes"),
        f"{location}.selected_member_uncompressed_bytes",
        positive=True,
    )
    if selected_member_bytes != contract["selected_member_uncompressed_bytes"]:
        raise AbbaSummaryInputError(
            f"{location}.selected_member_uncompressed_bytes does not match filesystem case {case!r}"
        )
    cache_max_bytes = _u64(
        evidence.get("cache_max_bytes"), f"{location}.cache_max_bytes", positive=True
    )
    cache_max_entries = _u64(
        evidence.get("cache_max_entries"),
        f"{location}.cache_max_entries",
        positive=True,
    )
    expected_cache_bytes, expected_cache_entries = _XLSX_REPEAT_STORE_CACHE_LIMITS[scenario]
    if (cache_max_bytes, cache_max_entries) != (
        expected_cache_bytes,
        expected_cache_entries,
    ):
        raise AbbaSummaryInputError(
            f"{location} cache limits do not match {scenario} scenario"
        )
    query_iterations = _u64(
        evidence.get("query_iterations"),
        f"{location}.query_iterations",
        positive=True,
    )
    if query_iterations != _XLSX_REPEAT_STORE_QUERY_ITERATIONS:
        raise AbbaSummaryInputError(
            f"{location}.query_iterations must be {_XLSX_REPEAT_STORE_QUERY_ITERATIONS}"
        )
    query_names = evidence.get("query_names")
    if query_names != list(_XLSX_REPEAT_STORE_QUERY_NAMES):
        raise AbbaSummaryInputError(
            f"{location}.query_names must be the four pinned semantic queries"
        )
    query_elapsed = evidence.get("query_elapsed_ns")
    if (
        not isinstance(query_elapsed, list)
        or len(query_elapsed) != len(_XLSX_REPEAT_STORE_QUERY_NAMES)
    ):
        raise AbbaSummaryInputError(
            f"{location}.query_elapsed_ns must contain four timings"
        )
    for index, elapsed in enumerate(query_elapsed):
        _u64(elapsed, f"{location}.query_elapsed_ns[{index}]", positive=True)
    timed_total = _u64(
        evidence.get("timed_elapsed_total_ns"),
        f"{location}.timed_elapsed_total_ns",
        positive=True,
    )
    if timed_total != sum(query_elapsed):
        raise AbbaSummaryInputError(
            f"{location}.timed_elapsed_total_ns does not equal query_elapsed_ns"
        )
    control_reacquire_count = _u64(
        evidence.get("control_reacquire_count"),
        f"{location}.control_reacquire_count",
    )
    timing_scope = _required_nonempty_string(
        evidence.get("timing_scope"), location, "timing_scope"
    )
    if timing_scope != _XLSX_REPEAT_STORE_TIMING_SCOPE:
        raise AbbaSummaryInputError(f"{location}.timing_scope is not recognized")
    claim_scope = _required_nonempty_string(
        evidence.get("claim_scope"), location, "claim_scope"
    )
    if claim_scope == _XLSX_REPEAT_STORE_PRIMARY_CLAIM_SCOPE:
        role = _XLSX_REPEAT_STORE_PRIMARY_SCOPE
        if implementation != _XLSX_REPEAT_STORE_PRIMARY_IMPLEMENTATION:
            raise AbbaSummaryInputError(
                f"{location}.implementation does not match primary claim_scope"
            )
        if control_reacquire_count != 0:
            raise AbbaSummaryInputError(
                f"{location}.control_reacquire_count must be zero for primary evidence"
            )
    elif claim_scope == _XLSX_REPEAT_STORE_STRUCTURAL_CLAIM_SCOPE:
        role = _XLSX_REPEAT_STORE_STRUCTURAL_SCOPE
        if implementation != _XLSX_REPEAT_STORE_STRUCTURAL_IMPLEMENTATION:
            raise AbbaSummaryInputError(
                f"{location}.implementation does not match structural claim_scope"
            )
        expected_reacquire_count = query_iterations * len(_XLSX_REPEAT_STORE_QUERY_NAMES)
        if control_reacquire_count != expected_reacquire_count:
            raise AbbaSummaryInputError(
                f"{location}.control_reacquire_count does not match structural queries"
            )
    else:
        raise AbbaSummaryInputError(f"{location}.claim_scope is not recognized")
    if role != contract["role"]:
        raise AbbaSummaryInputError(
            f"{location} scope does not match filesystem case {case!r}"
        )

    if evidence.get("budget_managed") is not True:
        raise AbbaSummaryInputError(f"{location}.budget_managed must be true")
    archive_bytes = corpus.get("archive_bytes")
    if archive_bytes is None:
        byte_summary = corpus.get("bytes")
        if isinstance(byte_summary, dict):
            archive_bytes = byte_summary.get("archive_bytes")
    archive_bytes = _u64(
        archive_bytes, f"{location}.corpus.archive_bytes", positive=True
    )
    expected_memory_limit = archive_bytes * 4 + 64 * 1024 * 1024
    if expected_memory_limit > U64_MAX:
        raise AbbaSummaryInputError(
            f"{location}.corpus.archive_bytes overflows the repeated-store memory limit"
        )
    if (
        _u64(evidence.get("budget_memory_limit"), f"{location}.budget_memory_limit")
        != expected_memory_limit
    ):
        raise AbbaSummaryInputError(
            f"{location}.budget_memory_limit does not match corpus bytes"
        )
    for field in ("budget_input_bytes_limit", "budget_work_limit"):
        if _u64(evidence.get(field), f"{location}.{field}") != U64_MAX:
            raise AbbaSummaryInputError(f"{location}.{field} must be u64::MAX")
    semantic_sha256 = _validate_output_sha256(
        evidence.get("semantic_projection_sha256"),
        f"{location}.semantic_projection_sha256",
    )
    if semantic_sha256 != contract["semantic_projection_sha256"]:
        raise AbbaSummaryInputError(
            f"{location}.semantic_projection_sha256 does not match the pinned "
            "repeated-store projection"
        )
    before = _validate_xlsx_repeat_store_counters(
        evidence.get("diagnostics_before"), f"{location}.diagnostics_before"
    )
    after = _validate_xlsx_repeat_store_counters(
        evidence.get("diagnostics_after"), f"{location}.diagnostics_after"
    )
    delta = _validate_xlsx_repeat_store_counters(
        evidence.get("diagnostics_delta"), f"{location}.diagnostics_delta"
    )
    for field in _XLSX_REPEAT_STORE_COUNTER_FIELDS:
        expected = before[field] + delta[field]
        if expected > U64_MAX or after[field] != expected:
            raise AbbaSummaryInputError(
                f"{location}.diagnostics_delta is inconsistent for {field}"
            )

    if scenario == "medium":
        if (
            before["cache_evictions"] == 0
            or before["cache_bypasses"] != 0
            or before["cache_oversized_bypasses"] != 0
        ):
            raise AbbaSummaryInputError(
                f"{location} medium evidence does not prove cache eviction"
            )
    else:
        if before["cache_oversized_bypasses"] == 0:
            raise AbbaSummaryInputError(
                f"{location} oversized evidence does not prove cache bypass"
            )

    if role == _XLSX_REPEAT_STORE_PRIMARY_SCOPE:
        if report_label in {"b1", "b2"}:
            if any(delta[field] != 0 for field in _XLSX_REPEAT_STORE_COUNTER_FIELDS):
                raise AbbaSummaryInputError(
                    f"{location}.diagnostics_delta must be zero for candidate primary evidence"
                )
        elif report_label in {"a1", "a2"} and not any(
            delta[field] > 0 for field in _XLSX_REPEAT_STORE_COUNTER_FIELDS
        ):
            raise AbbaSummaryInputError(
                f"{location}.diagnostics_delta must contain positive control evidence"
            )
    else:
        query_count = query_iterations * len(_XLSX_REPEAT_STORE_QUERY_NAMES)
        parts_per_query = 3 if scenario == "medium" else 1
        expected_cold_loads = query_count * parts_per_query
        if (
            delta["cache_cold_loads"] != expected_cold_loads
            or delta["cache_successful_loads"] != expected_cold_loads
            or delta["selected_member_read_calls"] != query_count
            or delta["selected_member_read_bytes"] == 0
            or delta["budget_input_bytes_used"] == 0
            or delta["budget_work_used"] == 0
        ):
            raise AbbaSummaryInputError(
                f"{location} structural control cache/read interval is not exact"
            )
        if scenario == "medium":
            if (
                delta["cache_bypasses"] != 0
                or delta["cache_oversized_bypasses"] != 0
                or delta["cache_evictions"] != expected_cold_loads
            ):
                raise AbbaSummaryInputError(
                    f"{location} medium structural cache interval is not exact"
                )
        elif (
            delta["cache_bypasses"] != query_count
            or delta["cache_oversized_bypasses"] != query_count
            or delta["cache_evictions"] != 0
        ):
            raise AbbaSummaryInputError(
                f"{location} oversized structural cache interval is not exact"
            )
    return role, evidence


def _looks_like_xlsx_repeat_store_corpus(corpus: Mapping[str, Any]) -> bool:
    """Recognize a repeated-store corpus marker before evidence dispatch.

    This is intentionally a one-way safety check.  A generic corpus is still
    allowed to use arbitrary names, but a report that retains the
    collision-free repeated-store name/generator pair cannot silently rename
    the selector and strip its evidence to enter the legacy path.  Archive
    hashes are deliberately not markers: the ordinary XLSX cell-values corpus
    shares the medium archive hash.
    """

    return (
        corpus.get("generator") == _XLSX_REPEAT_STORE_CORPUS_GENERATOR
        and corpus.get("name") in _XLSX_REPEAT_STORE_CORPUS_NAMES
    )


def _xlsx_repeat_store_identity_projection(
    value: Any, *, measurement: bool = False
) -> Any:
    """Retain structural constants and elide only timing/counter values.

    Unlike the generic filesystem projection, string arrays such as
    ``query_names`` remain explicit identities.  The selected member size,
    cache limits, query count, reacquisition count, and budget limits are
    structural constants and therefore remain exact.  Query timings and all
    diagnostics counter values remain measurements; the raw report canonical
    hash still binds their exact values.
    """

    if value is None or isinstance(value, str):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return "<number>" if measurement else value
    if isinstance(value, list):
        return [
            _xlsx_repeat_store_identity_projection(item, measurement=measurement)
            for item in value
        ]
    if isinstance(value, dict):
        projected: dict[str, Any] = {}
        for key, item in value.items():
            child_measurement = measurement or key in {
                "diagnostics_before",
                "diagnostics_after",
                "diagnostics_delta",
                "query_elapsed_ns",
                "timed_elapsed_total_ns",
            }
            projected[key] = _xlsx_repeat_store_identity_projection(
                item, measurement=child_measurement
            )
        return projected
    raise AbbaSummaryInputError(
        "xlsx_repeat_store contains unsupported value "
        f"{type(value).__name__}"
    )


def _validate_filesystem_sample(
    sample: Any, location: str, sample_count: int, cache_states: Sequence[str]
) -> dict[str, Any]:
    sample_object = _require_object(sample, location)
    missing = _FILESYSTEM_SAMPLE_REQUIRED_KEYS - set(sample_object)
    if missing:
        raise AbbaSummaryInputError(
            f"{location} is missing required keys: {sorted(missing)}"
        )
    unknown = set(sample_object) - _FILESYSTEM_SAMPLE_KEYS
    if unknown:
        raise AbbaSummaryInputError(
            f"{location} has unknown keys: {sorted(unknown)}"
        )
    range_size_keys = _FILESYSTEM_RANGE_SIZE_KEYS & set(sample_object)
    if range_size_keys and range_size_keys != _FILESYSTEM_RANGE_SIZE_KEYS:
        raise AbbaSummaryInputError(
            f"{location} must contain both logical read range-size counters or neither"
        )
    sample_index = _u64(sample_object.get("sample_index"), f"{location}.sample_index")
    if sample_index >= sample_count:
        raise AbbaSummaryInputError(
            f"{location}.sample_index={sample_index} is outside sample_count={sample_count}"
        )
    cache_state = _required_nonempty_string(
        sample_object.get("cache_state"), location, "cache_state"
    )
    if cache_state not in cache_states:
        raise AbbaSummaryInputError(
            f"{location}.cache_state={cache_state!r} is absent from cache_states"
        )
    if "child_process_id" in sample_object:
        _u32(sample_object["child_process_id"], f"{location}.child_process_id", positive=True)
    cold_advice = _required_nonempty_string(
        sample_object.get("cold_advice"), location, "cold_advice"
    )
    if cold_advice not in _COLD_ADVICE_VALUES:
        raise AbbaSummaryInputError(
            f"{location}.cold_advice has unknown value {cold_advice!r}"
        )
    if "cold_verified" in sample_object and sample_object["cold_verified"] is not None:
        _validate_cold_verified_sample(
            sample_object["cold_verified"], f"{location}.cold_verified"
        )
    if "logical_read_counter_scope" in sample_object:
        _required_nonempty_string(
            sample_object["logical_read_counter_scope"],
            location,
            "logical_read_counter_scope",
        )
    if (
        "logical_read_pattern" in sample_object
        and sample_object["logical_read_pattern"] is not None
    ):
        pattern = sample_object["logical_read_pattern"]
        if pattern not in {"sequential", "random", "unknown"}:
            raise AbbaSummaryInputError(
                f"{location}.logical_read_pattern has unknown value {pattern!r}"
            )
    if "output_sha256" in sample_object and sample_object["output_sha256"] is not None:
        _validate_output_sha256(sample_object["output_sha256"], f"{location}.output_sha256")
    for field in ("xlsx_source_sha256", "xlsx_semantic_sha256"):
        if field in sample_object:
            _validate_output_sha256(sample_object[field], f"{location}.{field}")
    if "allocation_metrics" in sample_object:
        _validate_raw_allocation_metrics(
            sample_object["allocation_metrics"], f"{location}.allocation_metrics"
        )
    if "xlsx_repeat_store" in sample_object:
        _require_object(sample_object["xlsx_repeat_store"], f"{location}.xlsx_repeat_store")
    for field in (
        "elapsed_ns",
        "parent_wall_ns",
        "logical_read_calls",
        "logical_read_requested_bytes",
        "logical_read_bytes",
        "logical_read_largest_requested_bytes",
        "logical_read_largest_returned_bytes",
        "max_concurrent_reads",
    ):
        if field in sample_object:
            _u64(sample_object[field], f"{location}.{field}")
    for field in (
        "output_bytes",
        "opc_materialized_parts",
        "cfb_changed_spans",
        "cfb_published_bytes",
    ):
        if field in sample_object and sample_object[field] is not None:
            _u64(sample_object[field], f"{location}.{field}")
    if "logical_read_request_sizes" in sample_object:
        request_sizes = sample_object["logical_read_request_sizes"]
        if not isinstance(request_sizes, list):
            raise AbbaSummaryInputError(
                f"{location}.logical_read_request_sizes must be a list"
            )
        for index, item in enumerate(request_sizes):
            _u64(item, f"{location}.logical_read_request_sizes[{index}]")
    if "logical_read_request_size_buckets" in sample_object:
        buckets = _require_object(
            sample_object["logical_read_request_size_buckets"],
            f"{location}.logical_read_request_size_buckets",
        )
        expected_buckets = {
            "bytes_0",
            "bytes_1_to_512",
            "bytes_513_to_4096",
            "bytes_4097_to_16384",
            "bytes_16385_to_65536",
            "bytes_over_65536",
        }
        if set(buckets) != expected_buckets:
            raise AbbaSummaryInputError(
                f"{location}.logical_read_request_size_buckets keys mismatch"
            )
        for field, item in buckets.items():
            _u64(item, f"{location}.logical_read_request_size_buckets.{field}")
    for field in (
        "process_metrics",
        "cfb_phases",
        "pptx_source_replay",
        "docx_source_replay",
    ):
        if field in sample_object and sample_object[field] is not None:
            _require_object(sample_object[field], f"{location}.{field}")
    return sample_object


def _filesystem_evidence_identity_projection(evidence: dict[str, Any]) -> str:
    projected: dict[str, Any] = {
        key: evidence[key]
        for key in (
            "case",
            "corpus",
            "warmup_iterations",
            "sample_count",
            "cache_states",
            "fresh_child_per_sample",
        )
    }
    for field in (
        "tool",
        "configuration",
        "config",
        "cold_verified_status",
        "cold_verified_claim_scope",
        "cold_verified_fincore_command",
    ):
        if field in evidence:
            projected[field] = evidence[field]
    if "cold_verified_samples" in evidence:
        projected["cold_verified_samples"] = [
            _filesystem_measurement_shape(item)
            for item in evidence["cold_verified_samples"]
        ]
    projected["samples"] = [
        {
            key: (
                sample[key]
                if key in _FILESYSTEM_SAMPLE_IDENTITY_KEYS
                else _xlsx_repeat_store_identity_projection(sample[key])
                if key == "xlsx_repeat_store"
                else _filesystem_measurement_shape(sample[key])
            )
            for key in sorted(sample)
        }
        for sample in evidence["samples"]
    ]
    if "cfb_owned" in evidence:
        projected["cfb_owned"] = [
            _filesystem_measurement_shape(item) for item in evidence["cfb_owned"]
        ]
    return _canonical_json(projected, "filesystem_evidence.identity")


def _filesystem_evidence_scope(identity: str, location: str) -> str | None:
    """Recover the validated repeated-store scope from an identity projection."""

    try:
        value = json.loads(identity)
    except (TypeError, ValueError) as error:
        raise AbbaSummaryInputError(
            f"{location} filesystem identity is not canonical JSON"
        ) from error
    samples = _require_object(value, location).get("samples")
    if not isinstance(samples, list):
        raise AbbaSummaryInputError(f"{location}.samples identity must be a list")
    scopes: set[str] = set()
    for index, sample in enumerate(samples):
        sample_object = _require_object(sample, f"{location}.samples[{index}]")
        repeated = sample_object.get("xlsx_repeat_store")
        if repeated is None:
            continue
        repeated_object = _require_object(
            repeated, f"{location}.samples[{index}].xlsx_repeat_store"
        )
        claim_scope = repeated_object.get("claim_scope")
        if claim_scope == _XLSX_REPEAT_STORE_PRIMARY_CLAIM_SCOPE:
            scopes.add(_XLSX_REPEAT_STORE_PRIMARY_SCOPE)
        elif claim_scope == _XLSX_REPEAT_STORE_STRUCTURAL_CLAIM_SCOPE:
            scopes.add(_XLSX_REPEAT_STORE_STRUCTURAL_SCOPE)
        else:
            raise AbbaSummaryInputError(
                f"{location}.samples[{index}].xlsx_repeat_store has an unknown claim scope"
            )
    if len(scopes) > 1:
        raise AbbaSummaryInputError(
            f"{location} filesystem identity mixes repeated-store scopes"
        )
    return next(iter(scopes), None)


def _validated_result_elapsed_by_sample(
    row: Mapping[str, Any], location: str, sample_count: int
) -> tuple[int, ...]:
    """Return authoritative elapsed values keyed by the original sample index.

    The harness sorts ``elapsed_ns.samples`` for statistics and records the
    inverse mapping in ``sample_order`` (sorted position -> original sample
    index).  Repeated-store evidence is additive, so it must bind to this
    exact permutation rather than guessing that the sorted position is the
    filesystem sample index.
    """

    elapsed = _require_object(row.get("elapsed_ns"), f"{location}.elapsed_ns")
    samples = elapsed.get("samples")
    if not isinstance(samples, list) or len(samples) != sample_count:
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.samples must contain exactly {sample_count} samples"
        )
    sample_values = tuple(
        _u64(value, f"{location}.elapsed_ns.samples[{index}]")
        for index, value in enumerate(samples)
    )
    sample_order = elapsed.get("sample_order")
    if not isinstance(sample_order, list) or len(sample_order) != sample_count:
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.sample_order must contain exactly {sample_count} entries"
        )
    normalized_order: list[int] = []
    for index, value in enumerate(sample_order):
        original_index = _u64(
            value, f"{location}.elapsed_ns.sample_order[{index}]"
        )
        if original_index >= sample_count:
            raise AbbaSummaryInputError(
                f"{location}.elapsed_ns.sample_order[{index}] is outside the sample range"
            )
        normalized_order.append(original_index)
    if set(normalized_order) != set(range(sample_count)):
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.sample_order must be an exact permutation"
        )
    by_original_index: list[int | None] = [None] * sample_count
    for sorted_position, original_index in enumerate(normalized_order):
        by_original_index[original_index] = sample_values[sorted_position]
    if any(value is None for value in by_original_index):  # pragma: no cover - permutation guard
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.sample_order must cover every sample"
        )
    return tuple(value for value in by_original_index if value is not None)


def _validate_filesystem_evidence(
    root: dict[str, Any],
    configuration: dict[str, Any],
    tool: dict[str, Any],
    indexed: Mapping[tuple[str, str], dict[str, Any]],
    label: str,
    *,
    report_role: str,
) -> tuple[bool, frozenset[str], dict[tuple[str, str], str], frozenset[int]]:
    raw = root.get("filesystem_evidence", _MISSING)
    for case, corpus_identity in indexed:
        corpus = json.loads(corpus_identity)
        if (
            _looks_like_xlsx_repeat_store_corpus(corpus)
            and case not in _XLSX_REPEAT_STORE_CASE_CONTRACT
        ):
            raise AbbaSummaryInputError(
                f"{label}.results case {case!r} uses a pinned repeated-store corpus "
                "but is not one of the exact repeated-store selectors"
            )
    if raw is _MISSING:
        if any(case in _XLSX_REPEAT_STORE_CASE_CONTRACT for case, _ in indexed):
            raise AbbaSummaryInputError(
                f"{label}.filesystem_evidence is required for repeated-store selectors"
            )
        return False, frozenset(), {}, frozenset()
    if not isinstance(raw, list):
        raise AbbaSummaryInputError(f"{label}.filesystem_evidence must be a list")
    evidence_index: dict[tuple[str, str], str] = {}
    filesystem_shapes: set[str] = set()
    report_child_process_ids: set[int] = set()
    report_pid_presence: bool | None = None
    configured_cache_states = configuration.get("filesystem_cache_states")
    if configured_cache_states is not None:
        configured_cache_states = _required_string_list(
            configured_cache_states,
            f"{label}.configuration",
            "filesystem_cache_states",
        )
    configured_fresh_child = configuration.get("filesystem_fresh_child_per_sample")
    if configured_fresh_child is not None:
        _required_bool(
            configured_fresh_child,
            f"{label}.configuration",
            "filesystem_fresh_child_per_sample",
        )
    for index, evidence in enumerate(raw):
        location = f"{label}.filesystem_evidence[{index}]"
        evidence_object = _require_object(evidence, location)
        missing = _FILESYSTEM_EVIDENCE_REQUIRED_KEYS - set(evidence_object)
        if missing:
            raise AbbaSummaryInputError(
                f"{location} is missing required keys: {sorted(missing)}"
            )
        unknown = set(evidence_object) - _FILESYSTEM_EVIDENCE_KEYS
        if unknown:
            raise AbbaSummaryInputError(
                f"{location} has unknown keys: {sorted(unknown)}"
            )
        case = _required_nonempty_string(evidence_object.get("case"), location, "case")
        corpus = _require_object(evidence_object.get("corpus"), f"{location}.corpus")
        if (
            _looks_like_xlsx_repeat_store_corpus(corpus)
            and case not in _XLSX_REPEAT_STORE_CASE_CONTRACT
        ):
            raise AbbaSummaryInputError(
                f"{location} uses a pinned repeated-store corpus on an unknown selector"
            )
        corpus_identity = _canonical_json(corpus, f"{location}.corpus")
        shape = corpus.get("shape")
        if not isinstance(shape, str) or not shape:
            raise AbbaSummaryInputError(
                f"{location}.corpus.shape must be a non-empty string"
            )
        key = (case, corpus_identity)
        if key in evidence_index:
            raise AbbaSummaryInputError(
                f"{label}.filesystem_evidence contains duplicate case/corpus identity"
            )
        if indexed.get(key) is None:
            raise AbbaSummaryInputError(
                f"{location}.case/corpus identity does not match a result row"
            )
        if "tool" in evidence_object:
            evidence_tool = _require_object(evidence_object["tool"], f"{location}.tool")
            if _canonical_json(evidence_tool, f"{location}.tool") != _canonical_json(
                tool, f"{label}.tool"
            ):
                raise AbbaSummaryInputError(
                    f"{location}.tool identity differs from report tool"
                )
        for config_field in ("configuration", "config"):
            if config_field in evidence_object:
                evidence_configuration = _require_object(
                    evidence_object[config_field], f"{location}.{config_field}"
                )
                if _canonical_json(
                    evidence_configuration, f"{location}.{config_field}"
                ) != _canonical_json(configuration, f"{label}.configuration"):
                    raise AbbaSummaryInputError(
                        f"{location}.{config_field} identity differs from report configuration"
                    )
        warmup_iterations = _required_positive_integer(
            evidence_object.get("warmup_iterations"), location, "warmup_iterations"
        )
        if warmup_iterations != configuration["warmup_iterations_per_case"]:
            raise AbbaSummaryInputError(
                f"{location}.warmup_iterations does not match configuration"
            )
        sample_count = _required_positive_integer(
            evidence_object.get("sample_count"), location, "sample_count"
        )
        if sample_count != configuration["samples_per_case"]:
            raise AbbaSummaryInputError(
                f"{location}.sample_count does not match configuration"
            )
        cache_states = _required_string_list(
            evidence_object.get("cache_states"), location, "cache_states"
        )
        if case in _XLSX_REPEAT_STORE_CASE_CONTRACT and cache_states != ["warm"]:
            raise AbbaSummaryInputError(
                f"{location}.cache_states must be exactly ['warm'] for repeated-store evidence"
            )
        if configured_cache_states is not None and cache_states != configured_cache_states:
            raise AbbaSummaryInputError(
                f"{location}.cache_states does not match configuration"
            )
        fresh_child = _required_bool(
            evidence_object.get("fresh_child_per_sample"),
            location,
            "fresh_child_per_sample",
        )
        if case in _XLSX_REPEAT_STORE_CASE_CONTRACT and fresh_child is not True:
            raise AbbaSummaryInputError(
                f"{location}.fresh_child_per_sample must be true for repeated-store evidence"
            )
        if configured_fresh_child is not None and fresh_child != configured_fresh_child:
            raise AbbaSummaryInputError(
                f"{location}.fresh_child_per_sample does not match configuration"
            )
        samples = evidence_object.get("samples")
        if not isinstance(samples, list):
            raise AbbaSummaryInputError(f"{location}.samples must be a list")
        expected_sample_total = sample_count * len(cache_states)
        if len(samples) != expected_sample_total:
            raise AbbaSummaryInputError(
                f"{location}.samples has {len(samples)} entries; expected "
                f"{expected_sample_total}"
            )
        result_row = indexed[key]
        authoritative_elapsed_by_sample = (
            _validated_result_elapsed_by_sample(
                result_row, f"{label}.results[{case}].elapsed_ns", sample_count
            )
            if case in _XLSX_REPEAT_STORE_CASE_CONTRACT
            and _XLSX_REPEAT_STORE_CASE_CONTRACT[case]["role"]
            == _XLSX_REPEAT_STORE_PRIMARY_SCOPE
            else None
        )
        pairs: list[tuple[int, str]] = []
        validated_samples = []
        range_size_presence: frozenset[str] | None = None
        repeated_store_roles: set[str] = set()
        repeated_store_semantics: set[str] = set()
        repeated_store_sources: set[str] = set()
        for sample_position, sample in enumerate(samples):
            validated_sample = _validate_filesystem_sample(
                sample,
                f"{location}.samples[{sample_position}]",
                sample_count,
                cache_states,
            )
            sample_range_size_presence = frozenset(
                _FILESYSTEM_RANGE_SIZE_KEYS & set(validated_sample)
            )
            if range_size_presence is None:
                range_size_presence = sample_range_size_presence
            elif sample_range_size_presence != range_size_presence:
                raise AbbaSummaryInputError(
                    f"{location}.samples must use one logical read range-size "
                    "counter schema consistently"
                )
            pairs.append(
                (validated_sample["sample_index"], validated_sample["cache_state"])
            )
            if "xlsx_repeat_store" in validated_sample:
                role, repeated_store = _validate_xlsx_repeat_store(
                    validated_sample["xlsx_repeat_store"],
                    f"{location}.samples[{sample_position}].xlsx_repeat_store",
                    case=case,
                    report_label=report_role,
                    corpus=corpus,
                )
                repeated_store_roles.add(role)
                repeated_store_semantics.add(
                    repeated_store["semantic_projection_sha256"]
                )
                for field in ("xlsx_source_sha256", "xlsx_semantic_sha256"):
                    if field not in validated_sample:
                        raise AbbaSummaryInputError(
                            f"{location}.samples[{sample_position}] is missing {field}"
                        )
                repeated_store_sources.add(validated_sample["xlsx_source_sha256"])
                if (
                    validated_sample["xlsx_semantic_sha256"]
                    != _XLSX_REPEAT_STORE_CASE_CONTRACT[case]["full_semantic_sha256"]
                ):
                    raise AbbaSummaryInputError(
                        f"{location}.samples[{sample_position}] XLSX semantic hash does not "
                        "match the pinned full-workbook semantic identity"
                    )
                sample_index = validated_sample["sample_index"]
                if repeated_store["timed_elapsed_total_ns"] != validated_sample["elapsed_ns"]:
                    raise AbbaSummaryInputError(
                        f"{location}.samples[{sample_position}] repeated-store timing "
                        "must equal filesystem elapsed"
                    )
                if authoritative_elapsed_by_sample is not None:
                    authoritative_elapsed = authoritative_elapsed_by_sample[sample_index]
                    if repeated_store["timed_elapsed_total_ns"] != authoritative_elapsed:
                        raise AbbaSummaryInputError(
                            f"{location}.samples[{sample_position}] repeated-store timing "
                            "must equal authoritative result elapsed"
                        )
            validated_samples.append(validated_sample)
        expected_pairs = [
            (sample_index, cache_state)
            for sample_index in range(sample_count)
            for cache_state in cache_states
        ]
        if pairs != expected_pairs:
            raise AbbaSummaryInputError(
                f"{location}.samples must contain each sample index/cache state exactly once"
            )
        repeated_store_presence = ["xlsx_repeat_store" in sample for sample in validated_samples]
        if case in _XLSX_REPEAT_STORE_CASE_CONTRACT and not any(repeated_store_presence):
            raise AbbaSummaryInputError(
                f"{location}.samples must contain xlsx_repeat_store for filesystem case {case!r}"
            )
        if any(repeated_store_presence) and not all(repeated_store_presence):
            raise AbbaSummaryInputError(
                f"{location}.samples must use one repeated-store evidence schema consistently"
            )
        if case in _XLSX_REPEAT_STORE_CASE_CONTRACT:
            role = next(iter(repeated_store_roles), None)
            if role == _XLSX_REPEAT_STORE_PRIMARY_SCOPE:
                for field in ("source", "sink", "output_sha256"):
                    if result_row.get(field) is not None:
                        raise AbbaSummaryInputError(
                            f"{label}.results[{case}].{field} must be absent or null for "
                            "repeated-store filesystem results"
                        )
        sample_pid_presence = [
            "child_process_id" in sample for sample in validated_samples
        ]
        evidence_has_pids = any(sample_pid_presence)
        if report_pid_presence is None:
            report_pid_presence = evidence_has_pids
        elif evidence_has_pids != report_pid_presence:
            raise AbbaSummaryInputError(
                f"{label}.filesystem_evidence must use one child_process_id "
                "presence schema across selectors"
            )
        if case in _XLSX_REPEAT_STORE_CASE_CONTRACT and not all(sample_pid_presence):
            raise AbbaSummaryInputError(
                f"{location}.samples must contain child_process_id for every repeated-store sample"
            )
        if any(sample_pid_presence):
            if not all(sample_pid_presence):
                raise AbbaSummaryInputError(
                    f"{location}.samples must use one child_process_id presence schema"
                )
            sample_pids = [sample["child_process_id"] for sample in validated_samples]
            if len(set(sample_pids)) != len(sample_pids):
                raise AbbaSummaryInputError(
                    f"{location}.samples child_process_id values must be unique"
                )
            duplicate_pids = report_child_process_ids.intersection(sample_pids)
            if duplicate_pids:
                raise AbbaSummaryInputError(
                    f"{location}.samples child_process_id values repeat another selector: "
                    f"{sorted(duplicate_pids)}"
                )
            report_child_process_ids.update(sample_pids)
        if repeated_store_roles and len(repeated_store_roles) != 1:
            raise AbbaSummaryInputError(
                f"{location}.samples mix primary and structural repeated-store scopes"
            )
        if repeated_store_semantics and len(repeated_store_semantics) != 1:
            raise AbbaSummaryInputError(
                f"{location}.samples have inconsistent XLSX semantic hashes"
            )
        if repeated_store_sources:
            corpus_source_sha256 = corpus.get("archive_sha256")
            _validate_output_sha256(
                corpus_source_sha256, f"{location}.corpus.archive_sha256"
            )
            if repeated_store_sources != {corpus_source_sha256}:
                raise AbbaSummaryInputError(
                    f"{location}.samples XLSX source hash differs from corpus archive"
                )
        if "cold_verified_status" in evidence_object:
            status = evidence_object["cold_verified_status"]
            if status is not None and status not in _COLD_VERIFIED_STATUS_VALUES:
                raise AbbaSummaryInputError(
                    f"{location}.cold_verified_status has unknown value {status!r}"
                )
        if "cold_verified_samples" in evidence_object:
            verified_samples = evidence_object["cold_verified_samples"]
            if not isinstance(verified_samples, list) or not verified_samples:
                raise AbbaSummaryInputError(
                    f"{location}.cold_verified_samples must be a non-empty list"
                )
            for verified_index, verified_sample in enumerate(verified_samples):
                _validate_cold_verified_sample(
                    verified_sample, f"{location}.cold_verified_samples[{verified_index}]"
                )
        for field in ("cold_verified_claim_scope", "cold_verified_fincore_command"):
            if field in evidence_object and evidence_object[field] is not None:
                _required_nonempty_string(evidence_object[field], location, field)
        if "cfb_owned" in evidence_object:
            owned = evidence_object["cfb_owned"]
            if not isinstance(owned, list):
                raise AbbaSummaryInputError(f"{location}.cfb_owned must be a list")
            for owned_index, item in enumerate(owned):
                owned_location = f"{location}.cfb_owned[{owned_index}]"
                owned_object = _require_object(item, owned_location)
                if set(owned_object) != {"sample_index", "cache_state", "evidence"}:
                    raise AbbaSummaryInputError(f"{owned_location} keys mismatch")
                owned_sample_index = _u64(
                    owned_object["sample_index"], f"{owned_location}.sample_index"
                )
                if owned_sample_index >= sample_count:
                    raise AbbaSummaryInputError(
                        f"{owned_location}.sample_index is outside sample_count"
                    )
                owned_cache_state = _required_nonempty_string(
                    owned_object["cache_state"], owned_location, "cache_state"
                )
                if owned_cache_state not in cache_states:
                    raise AbbaSummaryInputError(
                        f"{owned_location}.cache_state is absent from cache_states"
                    )
                _require_object(owned_object["evidence"], f"{owned_location}.evidence")
        filesystem_shapes.add(shape)
        evidence_index[key] = _filesystem_evidence_identity_projection(
            {**evidence_object, "samples": validated_samples}
        )
    expected_repeat_keys = {
        key for key in indexed if key[0] in _XLSX_REPEAT_STORE_CASE_CONTRACT
    }
    missing_repeat_keys = expected_repeat_keys - set(evidence_index)
    if missing_repeat_keys:
        raise AbbaSummaryInputError(
            f"{label}.filesystem_evidence is missing repeated-store selectors: "
            f"{sorted(missing_repeat_keys)}"
        )
    return True, frozenset(filesystem_shapes), evidence_index, frozenset(
        report_child_process_ids
    )


def validate_parallel_metrics(report: Any, label: str = "report") -> None:
    """Validate an emitted descriptive parallel-metrics envelope."""

    try:
        perf_compare.validate_parallel_metrics(report, label)
    except perf_compare.ComparisonInputError as error:
        raise AbbaSummaryInputError(str(error)) from error


def _validate_report(
    report: Any,
    label: str,
    *,
    profile: str | None = None,
    report_role: str | None = None,
    raw_canonical_sha256: str | None = None,
) -> tuple[
    int,
    dict[str, Any],
    dict[str, Any],
    dict[str, Any],
    dict[tuple[str, str], dict[str, Any]],
    str,
    frozenset[str],
    bool,
    dict[tuple[str, str], str],
    dict[str, Any] | None,
    frozenset[int],
]:
    if report_role is not None and report_role not in LEG_ORDER:
        raise AbbaSummaryInputError(
            f"{label} report_role must be one of {', '.join(LEG_ORDER)}"
        )
    semantic_role = report_role or label
    root = _require_object(report, label)
    detected_profile = detect_report_profile(root, label)
    if profile is not None and profile != detected_profile:
        raise AbbaSummaryInputError(
            f"{label} report profile {detected_profile!r} does not match {profile!r}"
        )
    selected_profile = profile or detected_profile
    if selected_profile not in REPORT_PROFILES:
        raise AbbaSummaryInputError(f"unsupported ABBA report profile {selected_profile!r}")
    if raw_canonical_sha256 is not None and SHA256_RE.fullmatch(raw_canonical_sha256) is None:
        raise AbbaSummaryInputError(f"{label}.report canonical SHA-256 is invalid")
    report_sha256 = raw_canonical_sha256 or _canonical_sha256(root, f"{label}.report")
    schema_version = root.get("schema_version")
    if isinstance(schema_version, bool) or not isinstance(schema_version, int):
        raise AbbaSummaryInputError(f"{label}.schema_version must be an integer")
    if schema_version != HARNESS_SCHEMA_VERSION:
        raise AbbaSummaryInputError(
            f"{label}.schema_version must be {HARNESS_SCHEMA_VERSION!r}"
        )
    validate_parallel_metrics(root, label)
    tool = _require_object(root.get("tool"), f"{label}.tool")
    if not tool:
        raise AbbaSummaryInputError(f"{label}.tool must not be empty")
    _validate_tool(tool, label, selected_profile)
    binary_identity = (
        _validate_binary_identity(root.get("binary_identity"), label, tool)
        if selected_profile == REPORT_PROFILE_CURRENT
        else None
    )
    environment = _require_object(root.get("environment"), f"{label}.environment")
    if not environment:
        raise AbbaSummaryInputError(f"{label}.environment must not be empty")
    _validate_environment(environment, label)
    configuration = _require_object(root.get("configuration"), f"{label}.configuration")
    if not configuration:
        raise AbbaSummaryInputError(f"{label}.configuration must not be empty")
    _validate_configuration(configuration, label)
    indexed = _index_results(root, label)
    _validate_opc_source_overlay_result_rows(indexed, configuration, label)
    _validate_docx_section_layout_result_rows(indexed, configuration, label)
    _validate_xls_numeric_source_summary(indexed, label)
    _validate_xls_numeric_operation_evidence(
        indexed,
        label,
        report_role=semantic_role,
    )
    filesystem_present, filesystem_shapes, filesystem_identity, filesystem_child_process_ids = (
        _validate_filesystem_evidence(
            root,
            configuration,
            tool,
            indexed,
            label,
            report_role=semantic_role,
        )
    )
    return (
        schema_version,
        tool,
        environment,
        configuration,
        indexed,
        report_sha256,
        filesystem_shapes,
        filesystem_present,
        filesystem_identity,
        binary_identity,
        filesystem_child_process_ids,
    )


def _stable_environment(environment: dict[str, Any]) -> dict[str, Any]:
    """Return environment facts expected to remain fixed across ABBA legs."""

    return {
        key: value for key, value in environment.items() if key not in ENVIRONMENT_VARIANTS
    }


def _opc_source_overlay_multi_part_identity(
    corpus: Mapping[str, Any],
    *,
    allow_pending_hashes: bool = False,
) -> tuple[str, int] | None:
    shape = corpus.get("shape")
    if not isinstance(shape, str):
        return None
    expected = OPC_SOURCE_OVERLAY_MULTI_PART_CORPUS_IDENTITIES.get(shape)
    if expected is None:
        return None
    pending_hash_fields = {
        field
        for field in ("archive_bytes", "archive_sha256", "target_payload_sha256")
        if expected.get(field) is None
    }
    if pending_hash_fields and not allow_pending_hashes:
        fields = ", ".join(sorted(pending_hash_fields))
        raise AbbaSummaryInputError(
            "multi-Part OPC fixed corpus identity constants are missing "
            f"for {shape}: fill {fields} from a compiled benchmark smoke run"
        )
    for field, value in expected.items():
        if field == "name_prefix" or field in pending_hash_fields:
            continue
        if corpus.get(field) != value:
            return None
    name = corpus.get("name")
    prefix = expected["name_prefix"]
    if not isinstance(name, str) or not name.startswith(prefix):
        return None
    suffix = name[len(prefix) :]
    if not suffix.isdigit():
        return None
    count = int(suffix)
    if count not in OPC_SOURCE_OVERLAY_MULTI_PART_COUNTS or name != f"{prefix}{count}":
        return None
    for field in ("archive_sha256", "target_payload_sha256"):
        value = corpus.get(field)
        if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
            return None
    archive_bytes = corpus.get("archive_bytes")
    if (
        isinstance(archive_bytes, bool)
        or not isinstance(archive_bytes, int)
        or archive_bytes <= 0
    ):
        return None
    if not pending_hash_fields:
        for field in ("archive_sha256", "target_payload_sha256"):
            if corpus.get(field) != expected[field]:
                return None
        if corpus.get("archive_bytes") != expected["archive_bytes"]:
            return None
    return shape, count


_OPC_SOURCE_OVERLAY_KEYS = frozenset(
    {
        "implementation",
        "timing_scope",
        "performance_claim",
        "overlay_mode",
        "replacement_semantics",
        "overlay_count",
        "source_shape",
        "payload_kind",
        "source_bytes",
        "source_sha256",
        "expected_eager_sha256",
        "source_cache_max_bytes",
        "source_cache_max_entries",
        "sink_max_bytes",
        "sink_max_write",
        "preparation_ns",
        "open_ns",
        "planning_ns",
        "publication_ns",
        "cache_before_publication_hits",
        "cache_before_publication_cold_loads",
        "cache_before_publication_retained_entries",
        "cache_before_publication_retained_bytes",
        "source_cache_after_publication_probe_hits",
        "source_cache_after_publication_probe_cold_loads",
        "source_cache_after_publication_probe_retained_entries",
        "source_cache_after_publication_probe_retained_bytes",
        "reopened_output_cache_hits",
        "reopened_output_cache_cold_loads",
        "reopened_output_cache_retained_entries",
        "reopened_output_cache_retained_bytes",
        "observed_after_publication_source_read_calls",
        "observed_after_publication_source_read_bytes",
        "observed_after_publication_ordinary_payload_read_calls",
        "observed_after_publication_ordinary_payload_read_bytes",
        "expected_eager_semantic_verified",
        "raw_members_and_order_preservation_verified",
        "equal_payload_noop_source_verified",
        "observed_output_sha256",
    }
)
_OPC_SOURCE_OVERLAY_PHASE_VECTOR_FIELDS = (
    "preparation_ns",
    "open_ns",
    "planning_ns",
    "publication_ns",
)
_OPC_SOURCE_OVERLAY_COUNTER_VECTOR_FIELDS = (
    "cache_before_publication_hits",
    "cache_before_publication_cold_loads",
    "cache_before_publication_retained_entries",
    "cache_before_publication_retained_bytes",
    "source_cache_after_publication_probe_hits",
    "source_cache_after_publication_probe_cold_loads",
    "source_cache_after_publication_probe_retained_entries",
    "source_cache_after_publication_probe_retained_bytes",
    "reopened_output_cache_hits",
    "reopened_output_cache_cold_loads",
    "reopened_output_cache_retained_entries",
    "reopened_output_cache_retained_bytes",
    "observed_after_publication_source_read_calls",
    "observed_after_publication_source_read_bytes",
    "observed_after_publication_ordinary_payload_read_calls",
    "observed_after_publication_ordinary_payload_read_bytes",
)
_OPC_SOURCE_OVERLAY_DYNAMIC_FIELDS = frozenset(
    {
        *_OPC_SOURCE_OVERLAY_PHASE_VECTOR_FIELDS,
        *_OPC_SOURCE_OVERLAY_COUNTER_VECTOR_FIELDS,
        "observed_output_sha256",
    }
)
_OPC_SOURCE_OVERLAY_SOURCE_VECTOR_FIELDS = (
    "read_calls",
    "read_bytes",
    "ordinary_payload_read_calls",
    "ordinary_payload_read_bytes",
    "max_in_flight_reads",
)
_OPC_SOURCE_OVERLAY_SINK_KEYS = frozenset(
    {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets"}
)
_OPC_SOURCE_OVERLAY_SINK_BUCKET_KEYS = frozenset(
    {
        "bytes_0",
        "bytes_1_to_512",
        "bytes_513_to_4096",
        "bytes_4097_to_16384",
        "bytes_16385_to_65536",
        "bytes_over_65536",
    }
)

_DOCX_SECTION_LAYOUT_TIMING_SCOPE = (
    "preparation, source-backed open, section query, columns edit, commit, "
    "and sequential publication; correctness, reopen, patch, refusal, cache, "
    "source, sink, and ZIP gates are excluded"
)
_DOCX_SECTION_LAYOUT_PHASE_VECTOR_FIELDS = (
    "preparation_ns",
    "open_ns",
    "query_ns",
    "edit_ns",
    "commit_ns",
    "publication_ns",
)
_DOCX_SECTION_LAYOUT_COUNTER_VECTOR_FIELDS = (
    "source_read_calls",
    "source_read_bytes",
    "ordinary_payload_materializations",
    "source_cache_before_publication_after_commit_hits",
    "source_cache_before_publication_after_commit_cold_loads",
    "source_cache_before_publication_after_commit_successful_loads",
    "source_cache_before_publication_after_commit_retained_entries",
    "source_cache_before_publication_after_commit_retained_bytes",
)
_DOCX_SECTION_LAYOUT_DYNAMIC_FIELDS = frozenset(
    {
        "sample_order",
        *_DOCX_SECTION_LAYOUT_PHASE_VECTOR_FIELDS,
        *_DOCX_SECTION_LAYOUT_COUNTER_VECTOR_FIELDS,
    }
)
_DOCX_SECTION_LAYOUT_ROOT_SOURCE_VECTOR_FIELDS = (
    "read_calls",
    "read_bytes",
    "ordinary_payload_read_calls",
    "ordinary_payload_read_bytes",
    "max_in_flight_reads",
    "ordinary_payload_materializations",
)
_DOCX_SECTION_LAYOUT_GATE_FIELDS = (
    "phase_sum_verified",
    "semantic_reopen_all_sections_verified",
    "table_cell_sectpr_excluded_verified",
    "raw_untouched_members_verified",
    "raw_member_order_and_comment_verified",
    "header_footer_relationships_verified",
    "source_immutability_verified",
    "patch_forward_inverse_verified",
    "publication_inverse_exact_source_verified",
    "stale_foreign_signed_noop_limits_partial_sink_verified",
    "cache_counters_verified",
    "sink_counters_verified",
)
_DOCX_SECTION_LAYOUT_KEYS = frozenset(
    {
        "implementation",
        "timing_scope",
        "performance_claim",
        "source_bytes",
        "source_sha256",
        "expected_output_sha256",
        "total_main_story_paragraph_count",
        "section_count",
        "paragraph_owned_section_positions",
        "selected_paragraph",
        "columns_before",
        "columns_after",
        "media_count",
        "media_bytes",
        "sample_order",
        *_DOCX_SECTION_LAYOUT_PHASE_VECTOR_FIELDS,
        *_DOCX_SECTION_LAYOUT_COUNTER_VECTOR_FIELDS,
        "output_sha256",
        *_DOCX_SECTION_LAYOUT_GATE_FIELDS,
    }
)


def _docx_section_layout_static_corpus_match(corpus: Any) -> bool:
    return isinstance(corpus, Mapping) and all(
        corpus.get(field) == value
        for field, value in DOCX_SECTION_LAYOUT_FIXED_CORPUS_FIELDS.items()
    )


def _validate_docx_section_layout_corpus_identity(
    corpus: Mapping[str, Any],
    location: str,
    *,
    allow_pending_hashes: bool = False,
) -> None:
    if not _docx_section_layout_static_corpus_match(corpus):
        raise AbbaSummaryInputError(
            f"{location} does not match the fixed DOCX section-layout corpus"
        )
    if set(corpus) != set(DOCX_SECTION_LAYOUT_CORPUS_IDENTITY):
        raise AbbaSummaryInputError(
            f"{location} schema does not match the fixed DOCX section-layout manifest"
        )
    for field in DOCX_SECTION_LAYOUT_HASH_FIELDS:
        value = corpus[field]
        if field.endswith("_sha256"):
            _validate_output_sha256(value, f"{location}.{field}")
        else:
            _u64(value, f"{location}.{field}", positive=True)
    pending = DOCX_SECTION_LAYOUT_HASHES_PENDING or any(
        value is None for value in DOCX_SECTION_LAYOUT_MEDIA_PAYLOAD_SHA256
    ) or DOCX_SECTION_LAYOUT_EXPECTED_OUTPUT_SHA256 is None
    if pending and not allow_pending_hashes:
        raise AbbaSummaryInputError(
            "DOCX section-layout fixed identity is pending "
            "change-0273-smoke archive/document/media/output hashes"
        )
    if not pending:
        for field in DOCX_SECTION_LAYOUT_HASH_FIELDS:
            if corpus[field] != DOCX_SECTION_LAYOUT_CORPUS_IDENTITY[field]:
                raise AbbaSummaryInputError(
                    f"{location}.{field} disagrees with the pinned DOCX section-layout identity"
                )


def _validate_opc_source_overlay(
    value: Any,
    location: str,
    *,
    case: str,
    corpus: Mapping[str, Any],
    samples_per_case: int,
    elapsed_samples: Sequence[Any],
    sample_order: Sequence[Any],
    source: Mapping[str, Any],
    sink: Mapping[str, Any],
    output_sha256: Any,
    allow_pending_hashes: bool = False,
) -> None:
    overlay = _require_object(value, location)
    if set(overlay) != _OPC_SOURCE_OVERLAY_KEYS:
        missing = sorted(_OPC_SOURCE_OVERLAY_KEYS - set(overlay))
        unknown = sorted(set(overlay) - _OPC_SOURCE_OVERLAY_KEYS)
        raise AbbaSummaryInputError(
            f"{location} schema mismatch (missing={missing!r}, unknown={unknown!r})"
        )
    if (
        isinstance(samples_per_case, bool)
        or not isinstance(samples_per_case, int)
        or samples_per_case <= 0
    ):
        raise AbbaSummaryInputError(f"{location} samples_per_case must be positive")
    if not isinstance(elapsed_samples, list):
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.samples must be a list"
        )
    elapsed = [
        _u64(item, f"{location}.elapsed_ns.samples[{index}]")
        for index, item in enumerate(elapsed_samples)
    ]
    if len(elapsed) != samples_per_case:
        raise AbbaSummaryInputError(
            f"{location} vectors must contain exactly {samples_per_case} samples"
        )
    if (
        not isinstance(sample_order, list)
        or len(sample_order) != samples_per_case
        or any(
            isinstance(item, bool)
            or not isinstance(item, int)
            or item < 0
            or item >= samples_per_case
            for item in sample_order
        )
        or sorted(sample_order) != list(range(samples_per_case))
    ):
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.sample_order must be an exact sample permutation"
        )
    sorted_elapsed = sorted(elapsed)
    if elapsed != sorted_elapsed:
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.samples must be sorted for overlay evidence"
        )
    if case not in OPC_SOURCE_OVERLAY_MULTI_PART_CASES:
        raise AbbaSummaryInputError(f"{location} is present for a non-multi-Part case")
    mode = case.removeprefix("opc_source_overlay_multi_part_")
    semantics = {
        "changed": "non-empty changed-payload replacement plan",
        "noop": "non-empty equal-payload replacement plan; semantic no-op",
        "mixed": "non-empty mixed changed/equal-payload replacement plan",
    }[mode]
    if overlay["implementation"] != "SourceBackedPackage::write_part_overlays_to_stream":
        raise AbbaSummaryInputError(f"{location}.implementation does not match the publisher")
    if not isinstance(overlay["timing_scope"], str) or not overlay["timing_scope"]:
        raise AbbaSummaryInputError(f"{location}.timing_scope must be non-empty")
    for required_phrase in (
        "preparation_ns",
        "open_ns",
        "planning_ns",
        "publication_ns",
        "excluded",
    ):
        if required_phrase not in overlay["timing_scope"]:
            raise AbbaSummaryInputError(
                f"{location}.timing_scope must name all reported phases and exclusions"
            )
    if overlay["performance_claim"] != "none":
        raise AbbaSummaryInputError(f"{location}.performance_claim must be 'none'")
    if overlay["overlay_mode"] != mode:
        raise AbbaSummaryInputError(f"{location}.overlay_mode disagrees with case")
    if overlay["replacement_semantics"] != semantics:
        raise AbbaSummaryInputError(
            f"{location}.replacement_semantics disagrees with overlay_mode"
        )
    identity = _opc_source_overlay_multi_part_identity(
        corpus, allow_pending_hashes=allow_pending_hashes
    )
    if identity is None:
        raise AbbaSummaryInputError(f"{location} corpus does not match fixed identity")
    shape, count = identity
    if overlay["overlay_count"] != count:
        raise AbbaSummaryInputError(f"{location}.overlay_count disagrees with corpus")
    expected_eager_identity = OPC_SOURCE_OVERLAY_MULTI_PART_EXPECTED_EAGER_SHA256.get(
        (case, shape, count)
    )
    if expected_eager_identity is None:
        raise AbbaSummaryInputError(
            f"{location} has no pinned expected eager digest for its case/shape/count cell"
        )
    if overlay["source_shape"] != shape or overlay["source_shape"] != corpus.get("shape"):
        raise AbbaSummaryInputError(f"{location}.source_shape disagrees with corpus")
    if overlay["payload_kind"] != corpus.get("payload_kind"):
        raise AbbaSummaryInputError(f"{location}.payload_kind disagrees with corpus")
    source_bytes = _u64(overlay["source_bytes"], f"{location}.source_bytes", positive=True)
    if source_bytes != corpus.get("archive_bytes"):
        raise AbbaSummaryInputError(f"{location}.source_bytes disagrees with corpus.archive_bytes")
    source_digest = _validate_output_sha256(
        overlay["source_sha256"], f"{location}.source_sha256"
    )
    if source_digest != corpus.get("archive_sha256"):
        raise AbbaSummaryInputError(f"{location}.source_sha256 disagrees with corpus")
    eager_digest = _validate_output_sha256(
        overlay["expected_eager_sha256"], f"{location}.expected_eager_sha256"
    )
    if eager_digest != expected_eager_identity:
        raise AbbaSummaryInputError(
            f"{location}.expected_eager_sha256 disagrees with the pinned case/shape/count identity"
        )
    output_digest = _validate_output_sha256(output_sha256, f"{location}.result.output_sha256")
    if mode == "noop" and output_digest != source_digest:
        raise AbbaSummaryInputError(f"{location} noop output differs from source digest")
    if mode != "noop" and output_digest == source_digest:
        raise AbbaSummaryInputError(
            f"{location} changed/mixed output must differ from source digest"
        )
    source_cache_max_bytes = _u64(
        overlay["source_cache_max_bytes"],
        f"{location}.source_cache_max_bytes",
        positive=True,
    )
    source_cache_max_entries = _u64(
        overlay["source_cache_max_entries"],
        f"{location}.source_cache_max_entries",
        positive=True,
    )
    if source_cache_max_bytes != corpus.get("uncompressed_payload_bytes"):
        raise AbbaSummaryInputError(
            f"{location}.source_cache_max_bytes disagrees with corpus cache policy"
        )
    if source_cache_max_entries != corpus.get("entry_count"):
        raise AbbaSummaryInputError(
            f"{location}.source_cache_max_entries disagrees with corpus cache policy"
        )
    sink_max_bytes = _u64(
        overlay["sink_max_bytes"], f"{location}.sink_max_bytes", positive=True
    )
    sink_max_write = _u64(
        overlay["sink_max_write"], f"{location}.sink_max_write", positive=True
    )
    if sink_max_write != 65_536:
        raise AbbaSummaryInputError(
            f"{location}.sink_max_write disagrees with the bounded sink policy"
        )
    sink_object = _require_object(sink, f"{location}.result.sink")
    if set(sink_object) != _OPC_SOURCE_OVERLAY_SINK_KEYS:
        raise AbbaSummaryInputError(
            f"{location}.result.sink schema does not match the bounded overlay sink"
        )
    sink_accepted_bytes = _u64(
        sink_object["accepted_bytes"], f"{location}.result.sink.accepted_bytes", positive=True
    )
    sink_write_calls = _u64(
        sink_object["write_calls"], f"{location}.result.sink.write_calls", positive=True
    )
    sink_largest_write = _u64(
        sink_object["largest_write"], f"{location}.result.sink.largest_write"
    )
    if sink_accepted_bytes > sink_max_bytes:
        raise AbbaSummaryInputError(
            f"{location}.result.sink.accepted_bytes exceeds sink_max_bytes"
        )
    if sink_largest_write > sink_max_write:
        raise AbbaSummaryInputError(
            f"{location}.result.sink.largest_write exceeds sink_max_write"
        )
    sink_formula_headroom = U64_MAX - 65_536
    if sink_accepted_bytes > sink_formula_headroom // 2:
        raise AbbaSummaryInputError(
            f"{location}.sink_max_bytes formula overflows u64"
        )
    expected_sink_max_bytes = sink_accepted_bytes * 2 + 65_536
    if sink_max_bytes != expected_sink_max_bytes:
        raise AbbaSummaryInputError(
            f"{location}.sink_max_bytes must equal 2*accepted_bytes+65536"
        )
    buckets = _require_object(
        sink_object["write_size_buckets"],
        f"{location}.result.sink.write_size_buckets",
    )
    if set(buckets) != _OPC_SOURCE_OVERLAY_SINK_BUCKET_KEYS:
        raise AbbaSummaryInputError(
            f"{location}.result.sink.write_size_buckets schema is incomplete"
        )
    bucket_counts = {
        field: _u64(value, f"{location}.result.sink.write_size_buckets.{field}")
        for field, value in buckets.items()
    }
    if sum(bucket_counts.values()) != sink_write_calls:
        raise AbbaSummaryInputError(
            f"{location}.result.sink write-size buckets disagree with write_calls"
        )
    if bucket_counts["bytes_over_65536"] != 0:
        raise AbbaSummaryInputError(
            f"{location}.result.sink observed a write above sink_max_write"
        )
    for field in (
        "expected_eager_semantic_verified",
        "raw_members_and_order_preservation_verified",
    ):
        if not _required_bool(overlay[field], location, field):
            raise AbbaSummaryInputError(f"{location}.{field} must be true")
    if _required_bool(overlay["equal_payload_noop_source_verified"], location, "equal_payload_noop_source_verified") != (mode == "noop"):
        raise AbbaSummaryInputError(
            f"{location}.equal_payload_noop_source_verified disagrees with overlay_mode"
        )

    def vector(field: str) -> list[int]:
        raw = overlay[field]
        if not isinstance(raw, list) or len(raw) != samples_per_case:
            raise AbbaSummaryInputError(
                f"{location}.{field} must contain exactly {samples_per_case} samples"
            )
        return [
            _u64(item, f"{location}.{field}[{index}]")
            for index, item in enumerate(raw)
        ]

    phases = {field: vector(field) for field in _OPC_SOURCE_OVERLAY_PHASE_VECTOR_FIELDS}
    counters = {field: vector(field) for field in _OPC_SOURCE_OVERLAY_COUNTER_VECTOR_FIELDS}
    digests = overlay["observed_output_sha256"]
    if not isinstance(digests, list) or len(digests) != samples_per_case:
        raise AbbaSummaryInputError(
            f"{location}.observed_output_sha256 must contain exactly {samples_per_case} samples"
        )
    if len(set(digests)) != 1:
        raise AbbaSummaryInputError(
            f"{location}.observed_output_sha256 is not deterministic across samples"
        )
    for index, digest in enumerate(digests):
        if _validate_output_sha256(digest, f"{location}.observed_output_sha256[{index}]") != output_digest:
            raise AbbaSummaryInputError(
                f"{location}.observed_output_sha256 disagrees with result.output_sha256"
            )
        phase_sum = sum(
            phases[field][index] for field in _OPC_SOURCE_OVERLAY_PHASE_VECTOR_FIELDS
        )
        if phase_sum > U64_MAX or phase_sum != sorted_elapsed[index]:
            raise AbbaSummaryInputError(
                f"{location} phase vectors must bind to sorted elapsed_ns.samples"
            )
    if any(counters[field][index] != 0 for field in (
        "cache_before_publication_retained_entries",
        "cache_before_publication_retained_bytes",
    ) for index in range(samples_per_case)):
        raise AbbaSummaryInputError(f"{location} cache retention oracle is non-zero")
    if any(
        counters[field][index] != 0
        for field in (
            "reopened_output_cache_hits",
            "reopened_output_cache_cold_loads",
            "reopened_output_cache_retained_entries",
            "reopened_output_cache_retained_bytes",
        )
        for index in range(samples_per_case)
    ):
        raise AbbaSummaryInputError(
            f"{location} reopened output is not a zero-cache lazy open"
        )
    expected_probe_bytes = count * corpus["entry_bytes"]
    if any(
        counters[field][index] != expected
        for field, expected in (
            ("source_cache_after_publication_probe_hits", 0),
            ("source_cache_after_publication_probe_cold_loads", count),
            ("source_cache_after_publication_probe_retained_entries", count),
            ("source_cache_after_publication_probe_retained_bytes", expected_probe_bytes),
        )
        for index in range(samples_per_case)
    ):
        raise AbbaSummaryInputError(
            f"{location} source cache replay does not match exact selected-Part evidence"
        )
    for nested_field, source_field in (
        (
            "observed_after_publication_source_read_calls",
            "read_calls",
        ),
        (
            "observed_after_publication_source_read_bytes",
            "read_bytes",
        ),
        (
            "observed_after_publication_ordinary_payload_read_calls",
            "ordinary_payload_read_calls",
        ),
        (
            "observed_after_publication_ordinary_payload_read_bytes",
            "ordinary_payload_read_bytes",
        ),
    ):
        if counters[nested_field] != source.get(source_field):
            raise AbbaSummaryInputError(
                f"{location}.{nested_field} disagrees with source.{source_field}"
            )
    if not all(
        counters[field][index] > 0
        for field in (
            "source_cache_after_publication_probe_cold_loads",
            "observed_after_publication_source_read_calls",
            "observed_after_publication_source_read_bytes",
            "observed_after_publication_ordinary_payload_read_bytes",
        )
        for index in range(samples_per_case)
    ):
        raise AbbaSummaryInputError(f"{location} source/cache evidence is empty")
    if any(
        counters["observed_after_publication_ordinary_payload_read_calls"][index] < count
        for index in range(samples_per_case)
    ):
        raise AbbaSummaryInputError(
            f"{location}.observed_after_publication_ordinary_payload_read_calls is below overlay_count"
        )


def _validate_docx_section_layout(
    value: Any,
    location: str,
    *,
    corpus: Mapping[str, Any],
    samples_per_case: int,
    elapsed_samples: Sequence[Any],
    sample_order: Sequence[Any],
    source: Mapping[str, Any],
    sink: Mapping[str, Any],
    output_sha256: Any,
    allow_pending_hashes: bool = False,
) -> None:
    overlay = _require_object(value, location)
    if set(overlay) != _DOCX_SECTION_LAYOUT_KEYS:
        missing = sorted(_DOCX_SECTION_LAYOUT_KEYS - set(overlay))
        unknown = sorted(set(overlay) - _DOCX_SECTION_LAYOUT_KEYS)
        raise AbbaSummaryInputError(
            f"{location} schema mismatch (missing={missing!r}, unknown={unknown!r})"
        )
    _validate_docx_section_layout_corpus_identity(
        corpus,
        f"{location}.corpus",
        allow_pending_hashes=allow_pending_hashes,
    )
    if (
        isinstance(samples_per_case, bool)
        or not isinstance(samples_per_case, int)
        or samples_per_case <= 0
    ):
        raise AbbaSummaryInputError(f"{location} samples_per_case must be positive")
    elapsed = [
        _u64(item, f"{location}.elapsed_ns.samples[{index}]")
        for index, item in enumerate(elapsed_samples)
    ]
    if len(elapsed) != samples_per_case:
        raise AbbaSummaryInputError(
            f"{location} vectors must contain exactly {samples_per_case} samples"
        )
    if elapsed != sorted(elapsed):
        raise AbbaSummaryInputError(
            f"{location}.elapsed_ns.samples must be sorted ascending"
        )

    def exact_permutation(raw: Any, field: str) -> list[int]:
        if (
            not isinstance(raw, list)
            or len(raw) != samples_per_case
            or any(
                isinstance(item, bool)
                or not isinstance(item, int)
                or item < 0
                or item >= samples_per_case
                for item in raw
            )
            or sorted(raw) != list(range(samples_per_case))
        ):
            raise AbbaSummaryInputError(
                f"{location}.{field} must be an exact sample permutation"
            )
        return raw

    elapsed_order = exact_permutation(sample_order, "elapsed_ns.sample_order")
    nested_order = exact_permutation(overlay["sample_order"], "sample_order")
    if nested_order != elapsed_order:
        raise AbbaSummaryInputError(
            f"{location}.sample_order disagrees with elapsed_ns.sample_order"
        )
    if overlay["implementation"] != "litchi-docx-source-backed-section-layout":
        raise AbbaSummaryInputError(f"{location}.implementation does not match the publisher")
    if overlay["timing_scope"] != _DOCX_SECTION_LAYOUT_TIMING_SCOPE:
        raise AbbaSummaryInputError(f"{location}.timing_scope does not match the named phases")
    if overlay["performance_claim"] != "none":
        raise AbbaSummaryInputError(f"{location}.performance_claim must be 'none'")

    source_bytes = _u64(overlay["source_bytes"], f"{location}.source_bytes", positive=True)
    if source_bytes != corpus["archive_bytes"]:
        raise AbbaSummaryInputError(f"{location}.source_bytes disagrees with corpus.archive_bytes")
    source_digest = _validate_output_sha256(
        overlay["source_sha256"], f"{location}.source_sha256"
    )
    if source_digest != corpus["archive_sha256"]:
        raise AbbaSummaryInputError(f"{location}.source_sha256 disagrees with corpus.archive_sha256")
    expected_output_digest = _validate_output_sha256(
        overlay["expected_output_sha256"], f"{location}.expected_output_sha256"
    )
    result_output_digest = _validate_output_sha256(
        output_sha256, f"{location}.result.output_sha256"
    )
    if expected_output_digest != result_output_digest:
        raise AbbaSummaryInputError(
            f"{location}.expected_output_sha256 disagrees with result.output_sha256"
        )
    if not DOCX_SECTION_LAYOUT_HASHES_PENDING:
        if DOCX_SECTION_LAYOUT_EXPECTED_OUTPUT_SHA256 is None:
            raise AbbaSummaryInputError(
                f"{location} has no pinned DOCX section-layout output digest"
            )
        if expected_output_digest != DOCX_SECTION_LAYOUT_EXPECTED_OUTPUT_SHA256:
            raise AbbaSummaryInputError(
                f"{location}.expected_output_sha256 disagrees with the pinned identity"
            )

    static_integers = {
        "total_main_story_paragraph_count": 257,
        "section_count": 3,
        "selected_paragraph": 64,
        "columns_before": 2,
        "columns_after": 3,
        "media_count": 4,
        "media_bytes": 512 * 1024,
    }
    for field, expected in static_integers.items():
        if _u64(overlay[field], f"{location}.{field}") != expected:
            raise AbbaSummaryInputError(
                f"{location}.{field} disagrees with the fixed section-layout identity"
            )
    positions = overlay["paragraph_owned_section_positions"]
    if (
        not isinstance(positions, list)
        or len(positions) != 2
        or any(isinstance(item, bool) or not isinstance(item, int) for item in positions)
        or positions != [64, 129]
    ):
        raise AbbaSummaryInputError(
            f"{location}.paragraph_owned_section_positions disagrees with the fixed identity"
        )

    def vector(field: str) -> list[int]:
        raw = overlay[field]
        if not isinstance(raw, list) or len(raw) != samples_per_case:
            raise AbbaSummaryInputError(
                f"{location}.{field} must contain exactly {samples_per_case} samples"
            )
        return [
            _u64(item, f"{location}.{field}[{index}]")
            for index, item in enumerate(raw)
        ]

    phases = {field: vector(field) for field in _DOCX_SECTION_LAYOUT_PHASE_VECTOR_FIELDS}
    counters = {field: vector(field) for field in _DOCX_SECTION_LAYOUT_COUNTER_VECTOR_FIELDS}
    for index in range(samples_per_case):
        phase_sum = sum(
            phases[field][index] for field in _DOCX_SECTION_LAYOUT_PHASE_VECTOR_FIELDS
        )
        if phase_sum > U64_MAX or phase_sum != elapsed[index]:
            raise AbbaSummaryInputError(
                f"{location} phase vectors must bind to sorted elapsed_ns.samples"
            )

    for field in _DOCX_SECTION_LAYOUT_ROOT_SOURCE_VECTOR_FIELDS:
        raw = source.get(field)
        if not isinstance(raw, list) or len(raw) != samples_per_case:
            raise AbbaSummaryInputError(
                f"{location}.source.{field} must contain exactly {samples_per_case} samples"
            )
        for index, item in enumerate(raw):
            _u64(item, f"{location}.source.{field}[{index}]")
    for nested_field, root_field in (
        ("source_read_calls", "read_calls"),
        ("source_read_bytes", "read_bytes"),
        ("ordinary_payload_materializations", "ordinary_payload_materializations"),
    ):
        if counters[nested_field] != source[root_field]:
            raise AbbaSummaryInputError(
                f"{location}.{nested_field} disagrees with source.{root_field}"
            )
    if any(value != 1 for value in counters["ordinary_payload_materializations"]):
        raise AbbaSummaryInputError(
            f"{location}.ordinary_payload_materializations must prove one retained main Part"
        )
    cache_fields = _DOCX_SECTION_LAYOUT_COUNTER_VECTOR_FIELDS[3:]
    if any(len(set(counters[field])) != 1 for field in cache_fields):
        raise AbbaSummaryInputError(f"{location} source cache diagnostics are not deterministic")
    for field, expected in (
        ("source_cache_before_publication_after_commit_cold_loads", 1),
        ("source_cache_before_publication_after_commit_successful_loads", 1),
        ("source_cache_before_publication_after_commit_retained_entries", 1),
    ):
        if any(value != expected for value in counters[field]):
            raise AbbaSummaryInputError(
                f"{location}.{field} disagrees with the one-main-Part cache contract"
            )
    if any(
        value == 0
        for value in counters[
            "source_cache_before_publication_after_commit_retained_bytes"
        ]
    ):
        raise AbbaSummaryInputError(f"{location} retained source cache bytes are empty")
    expected_retained_bytes = corpus["target_payload_bytes"]
    if any(
        value != expected_retained_bytes
        for value in counters[
            "source_cache_before_publication_after_commit_retained_bytes"
        ]
    ):
        raise AbbaSummaryInputError(
            f"{location} retained source cache bytes differ from the accessed document part"
        )
    if any(
        source[field][index] == 0
        for field in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes")
        for index in range(samples_per_case)
    ):
        raise AbbaSummaryInputError(f"{location} source read evidence is empty")

    output_digests = overlay["output_sha256"]
    if not isinstance(output_digests, list) or len(output_digests) != samples_per_case:
        raise AbbaSummaryInputError(
            f"{location}.output_sha256 must contain exactly {samples_per_case} samples"
        )
    for index, digest in enumerate(output_digests):
        if _validate_output_sha256(digest, f"{location}.output_sha256[{index}]") != result_output_digest:
            raise AbbaSummaryInputError(
                f"{location}.output_sha256 disagrees with result.output_sha256"
            )

    sink_object = _require_object(sink, f"{location}.result.sink")
    if set(sink_object) != _OPC_SOURCE_OVERLAY_SINK_KEYS:
        raise AbbaSummaryInputError(
            f"{location}.result.sink schema does not match the bounded sink"
        )
    sink_accepted_bytes = _u64(
        sink_object["accepted_bytes"], f"{location}.result.sink.accepted_bytes", positive=True
    )
    sink_write_calls = _u64(
        sink_object["write_calls"], f"{location}.result.sink.write_calls", positive=True
    )
    sink_largest_write = _u64(sink_object["largest_write"], f"{location}.result.sink.largest_write")
    if sink_largest_write > 65_536:
        raise AbbaSummaryInputError(
            f"{location}.result.sink.largest_write exceeds the configured sink ceiling"
        )
    buckets = _require_object(
        sink_object["write_size_buckets"], f"{location}.result.sink.write_size_buckets"
    )
    if set(buckets) != _OPC_SOURCE_OVERLAY_SINK_BUCKET_KEYS:
        raise AbbaSummaryInputError(
            f"{location}.result.sink.write_size_buckets schema is incomplete"
        )
    bucket_counts = {
        field: _u64(value, f"{location}.result.sink.write_size_buckets.{field}")
        for field, value in buckets.items()
    }
    if sum(bucket_counts.values()) != sink_write_calls:
        raise AbbaSummaryInputError(
            f"{location}.result.sink write-size buckets disagree with write_calls"
        )
    if bucket_counts["bytes_over_65536"] != 0:
        raise AbbaSummaryInputError(
            f"{location}.result.sink observed a write above the configured ceiling"
        )
    if sink_accepted_bytes == 0:
        raise AbbaSummaryInputError(f"{location}.result.sink accepted no output bytes")
    if sink_accepted_bytes != DOCX_SECTION_LAYOUT_EXPECTED_OUTPUT_BYTES:
        raise AbbaSummaryInputError(
            f"{location}.result.sink accepted bytes differ from the pinned output"
        )

    for field in _DOCX_SECTION_LAYOUT_GATE_FIELDS:
        if not _required_bool(overlay[field], location, field):
            raise AbbaSummaryInputError(f"{location}.{field} must be true")


def _validate_docx_section_layout_result_rows(
    indexed: Mapping[tuple[str, str], dict[str, Any]],
    configuration: Mapping[str, Any],
    label: str,
) -> None:
    samples_per_case = _required_positive_integer(
        configuration.get("samples_per_case"),
        f"{label}.configuration",
        "samples_per_case",
    )
    for (case, corpus_identity), row in indexed.items():
        source = row.get("source")
        if case != DOCX_SECTION_LAYOUT_CASE:
            if isinstance(source, dict) and source.get("docx_section_layout") is not None:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.docx_section_layout is only valid for its fixed selector"
                )
            continue
        source_object = _require_object(source, f"{label}.{case}.source")
        overlay = _require_object(
            source_object.get("docx_section_layout"),
            f"{label}.{case}.source.docx_section_layout",
        )
        sink_object = _require_object(row.get("sink"), f"{label}.{case}.sink")
        elapsed_object = _require_object(
            row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns"
        )
        _validate_docx_section_layout(
            overlay,
            f"{label}.{case}.source.docx_section_layout",
            corpus=json.loads(corpus_identity),
            samples_per_case=samples_per_case,
            elapsed_samples=elapsed_object.get("samples"),
            sample_order=elapsed_object.get("sample_order"),
            source=source_object,
            sink=sink_object,
            output_sha256=row.get("output_sha256"),
        )


def _validate_opc_source_overlay_result_rows(
    indexed: Mapping[tuple[str, str], dict[str, Any]],
    configuration: Mapping[str, Any],
    label: str,
) -> None:
    samples_per_case = _required_positive_integer(
        configuration.get("samples_per_case"), f"{label}.configuration", "samples_per_case"
    )
    for (case, corpus_identity), row in indexed.items():
        source = row.get("source")
        if case not in OPC_SOURCE_OVERLAY_MULTI_PART_CASES:
            if isinstance(source, dict) and source.get("opc_source_overlay") is not None:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.opc_source_overlay is only valid for multi-Part cases"
                )
            continue
        source_object = _require_object(source, f"{label}.{case}.source")
        overlay = _require_object(
            source_object.get("opc_source_overlay"),
            f"{label}.{case}.source.opc_source_overlay",
        )
        sink_object = _require_object(
            row.get("sink"), f"{label}.{case}.sink"
        )
        for field in _OPC_SOURCE_OVERLAY_SOURCE_VECTOR_FIELDS:
            raw = source_object.get(field)
            if not isinstance(raw, list) or len(raw) != samples_per_case:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.{field} must contain exactly {samples_per_case} samples"
                )
            for index, item in enumerate(raw):
                _u64(item, f"{label}.{case}.source.{field}[{index}]")
        elapsed_object = _require_object(
            row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns"
        )
        _validate_opc_source_overlay(
            overlay,
            f"{label}.{case}.source.opc_source_overlay",
            case=case,
            corpus=json.loads(corpus_identity),
            samples_per_case=samples_per_case,
            elapsed_samples=elapsed_object["samples"],
            sample_order=elapsed_object.get("sample_order"),
            source=source_object,
            sink=sink_object,
            output_sha256=row.get("output_sha256"),
        )


def _validate_configuration_rows(
    configuration: dict[str, Any],
    indexed: Mapping[tuple[str, str], dict[str, Any]],
    label: str,
    *,
    filesystem_shapes: Iterable[str] = (),
    projected: bool = False,
) -> None:
    """Check optional harness cardinality/selector declarations when present."""

    cases = configuration["cases"]
    actual_cases = {case for case, _ in indexed}
    if actual_cases != set(cases):
        raise AbbaSummaryInputError(
            f"{label}.configuration.cases does not match result cases"
        )
    multi_part_pairs: dict[str, list[tuple[str, int]]] = {}
    for case, corpus_identity in indexed:
        if case not in OPC_SOURCE_OVERLAY_MULTI_PART_CASES:
            continue
        corpus = json.loads(corpus_identity)
        identity = _opc_source_overlay_multi_part_identity(corpus)
        if identity is None:
            raise AbbaSummaryInputError(
                f"{label}.{case} corpus does not match the pinned multi-Part OPC identity"
            )
        multi_part_pairs.setdefault(case, []).append(identity)
    expected_multi_part_pairs = {
        (shape, count)
        for shape in OPC_SOURCE_OVERLAY_MULTI_PART_CORPUS_IDENTITIES
        for count in OPC_SOURCE_OVERLAY_MULTI_PART_COUNTS
    }
    for case in OPC_SOURCE_OVERLAY_MULTI_PART_CASES.intersection(cases):
        actual_pairs = multi_part_pairs.get(case, [])
        if len(actual_pairs) != len(expected_multi_part_pairs) or set(actual_pairs) != expected_multi_part_pairs:
            raise AbbaSummaryInputError(
                f"{label}.{case} must cover the complete multi-Part OPC 3-shape x 3-count matrix"
            )
    if DOCX_SECTION_LAYOUT_CASE in cases:
        docx_pairs = [
            json.loads(corpus_identity)
            for case, corpus_identity in indexed
            if case == DOCX_SECTION_LAYOUT_CASE
        ]
        if len(docx_pairs) != 1:
            raise AbbaSummaryInputError(
                f"{label}.{DOCX_SECTION_LAYOUT_CASE} must have exactly one fixed corpus identity"
            )
        if not _docx_section_layout_static_corpus_match(docx_pairs[0]):
            raise AbbaSummaryInputError(
                f"{label}.{DOCX_SECTION_LAYOUT_CASE} corpus does not match the fixed identity"
            )
    declared_shapes = {
        shape
        for field, values in configuration.items()
        if field.endswith("_shapes")
        for shape in values
    }
    actual_shapes: set[str] = set()
    undeclared_rows: list[tuple[str, dict[str, Any]]] = []
    for case, corpus_identity in indexed:
        corpus = json.loads(corpus_identity)
        shape = corpus.get("shape")
        if not isinstance(shape, str) or not shape:
            raise AbbaSummaryInputError(
                f"{label}.results case {case!r} corpus.shape must be a non-empty string"
            )
        actual_shapes.add(shape)
        if shape not in declared_shapes:
            undeclared_rows.append((case, corpus))
    filesystem_shape_set = set(filesystem_shapes)
    filesystem_exception = bool(filesystem_shape_set) and actual_shapes.issubset(
        filesystem_shape_set
    )
    fixed_case_exception = all(
        (
            (expected := FIXED_CASE_CORPUS_IDENTITIES.get(case)) is not None
            and all(corpus.get(field) == value for field, value in expected.items())
        )
        or _opc_source_overlay_multi_part_identity(corpus) is not None
        for case, corpus in undeclared_rows
    )
    if undeclared_rows and not filesystem_exception and not fixed_case_exception:
        raise AbbaSummaryInputError(
            f"{label}.configuration shape declarations do not cover result shapes"
        )
    samples_per_case = configuration.get("samples_per_case")
    for key, row in indexed.items():
        case = key[0]
        if projected and "_elapsed_statistics" in row:
            elapsed_statistics = _require_object(
                row["_elapsed_statistics"], f"{label}.{case}._elapsed_statistics"
            )
            if elapsed_statistics.get("sample_count") != samples_per_case:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.elapsed_ns sample count does not match samples_per_case"
                )
            continue
        elapsed = _require_object(row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
        samples = elapsed.get("samples")
        if not isinstance(samples, list) or len(samples) != samples_per_case:
            raise AbbaSummaryInputError(
                f"{label}.{case}.elapsed_ns.samples does not match samples_per_case"
            )


def _validate_drift_ceilings(value: Mapping[str, Any] | None) -> dict[str, float]:
    if value is None:
        return dict(DEFAULT_DRIFT_CEILINGS)
    if not isinstance(value, Mapping):
        raise AbbaSummaryInputError("drift ceilings must be an object")
    if set(value) != set(STATISTICS):
        raise AbbaSummaryInputError(
            f"drift ceilings must contain exactly {list(STATISTICS)!r}"
        )
    ceilings = {
        name: _finite_number(value[name], f"drift ceilings.{name}")
        for name in STATISTICS
    }
    if any(ceiling < 0 for ceiling in ceilings.values()):
        raise AbbaSummaryInputError("drift ceilings must be non-negative")
    return ceilings


def _identity_value(row: dict[str, Any], field: str, location: str) -> tuple[bool, str]:
    # Harness Option fields may be serialized as explicit null.  Null carries
    # no identity, so it is absence rather than a verified-equal payload.
    present = field in row and row[field] is not None
    value = row[field] if present else None
    if field == "source":
        value = _source_identity_projection(value)
    return present, _canonical_json(value, f"{location}.{field}")


CFB_OPEN_STREAM_SOURCE_MEASUREMENTS = (
    "expected_direct_physical_range",
    "logical_read_bytes",
    "logical_read_calls",
    "logical_read_range_sizes",
    "logical_read_ranges",
    "open_ns",
    "open_read_bytes",
    "open_read_calls",
    "open_read_range_sizes",
    "open_read_ranges",
    "operation_ns",
    "per_operation_ns",
    "per_operation_read_bytes",
    "per_operation_read_calls",
    "per_operation_read_range_sizes",
    "per_operation_read_ranges",
    "root_cache_read_bytes",
    "total_ns",
)

ODS_SOURCE_CELL_MEASUREMENTS = (
    "commit_ns",
    "content_source_read_bytes",
    "content_source_read_calls",
    "lifecycle_ns",
    "open_ns",
    "pictures_source_read_bytes",
    "pictures_source_read_calls",
    "publication_ns",
    "source_read_bytes",
    "source_read_calls",
    "source_read_range_overlap_bytes",
    "source_version_calls",
    "stage_ns",
    "untouched_source_read_bytes",
    "untouched_source_read_calls",
)

XLSX_CELL_VALUES_SOURCE_MEASUREMENTS = (
    "commit_ns",
    "open_ns",
    "plan_ns",
    "publication_ns",
    "reopen_ns",
)

XLS_NUMERIC_OPERATION_COUNTER_SCOPE = (
    "validated overlay logical pass shape; no runtime, allocator, or syscall counters"
)
XLS_NUMERIC_OPERATION_EVIDENCE_SCHEMA = "xls_numeric.operation_evidence.v1"
XLS_NUMERIC_LEGACY_SOURCE_COUNTER_SCOPE = "owned-source-ingress-only"
XLS_NUMERIC_CURRENT_SOURCE_COUNTER_SCOPE = (
    "owned-source-ingress-only; xls-overlay-operation-evidence-v1"
)
XLS_NUMERIC_OWNED_INPUT_SCOPE = (
    "complete in-memory CFB bytes; no positional/physical I/O"
)
XLS_NUMERIC_OPERATION_ATOMIC_SCOPE = (
    "public save durability contract; temporary-file, flush, fsync, rename, and parent-sync stage counters are not observable"
)
XLS_NUMERIC_CFB_MAX_BYTES = 2 * 1024 * 1024 * 1024
XLS_NUMERIC_CASE_CONTRACTS = {
    "xls_numeric_eager_number_edit_save": ("eager", "Number", 1),
    "xls_numeric_source_backed_number_edit_save": ("source_backed", "Number", 1),
    "xls_numeric_plan_only_number_edit_save": ("plan_only", "Number", 1),
    "xls_numeric_eager_rk_mulrk_edit_save": ("eager", "RK+MulRK", 3),
    "xls_numeric_source_backed_rk_mulrk_edit_save": (
        "source_backed",
        "RK+MulRK",
        3,
    ),
    "xls_numeric_plan_only_rk_mulrk_edit_save": ("plan_only", "RK+MulRK", 3),
}
# Native corpus manifests are part of the selector identity.  Keep the
# complete producer shape here so a report cannot replace a native corpus
# with a renamed or partially described synthetic one while preserving only
# its archive length.
XLS_NUMERIC_CORPUS_CONTRACTS = {
    "Number": {
        "name": "xls-comments-opaque-heavy",
        "generator": "litchi-xls-comments-opaque-heavy-v1",
        "package_format": "XLS/CFB",
        "shape": "256-comments-opaque-heavy",
        "payload_kind": "incompressible",
        "compression": "none",
        "entry_count": 257,
        "archive_member_count": 10,
        "entry_bytes": 2_097_152,
        "uncompressed_payload_bytes": 16_858_200,
        "archive_bytes": 16_995_840,
        "archive_sha256": "6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53",
        "target_entry": "Workbook",
        "target_payload_bytes": 80_946,
        "target_payload_sha256": "c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041",
        "xlsx": None,
    },
    "RK+MulRK": {
        "name": "xls-rk-mulrk-deterministic",
        "generator": "litchi-xls-rk-mulrk-publication-v1",
        "package_format": "XLS/CFB",
        "shape": "one-rk-one-mulrk",
        "payload_kind": "incompressible",
        "compression": "none",
        "entry_count": 4,
        "archive_member_count": 4,
        "entry_bytes": 98_304,
        "uncompressed_payload_bytes": 198_311,
        "archive_bytes": 202_752,
        "archive_sha256": "61a649b081c24821b02aa5e69b6ad1dc53b0232019d3668dd3776f402989c594",
        "target_entry": "Workbook",
        "target_payload_bytes": 1_665,
        "target_payload_sha256": "1b57f77d776cc8d0ed0f5154f7f7db2abca7f5fa7e23ecc69dd316c1c0b65967",
        "xlsx": None,
    },
}
XLS_NUMERIC_WORKBOOK_SIZE_CONTRACTS = {
    "Number": {"source_workbook_bytes": 80_946, "target_workbook_bytes": 80_946},
    "RK+MulRK": {"source_workbook_bytes": 1_665, "target_workbook_bytes": 1_665},
}
XLS_NUMERIC_OUTPUT_CONTRACTS = {
    "Number": {
        "output_sha256": "f8f37064dc842550445b674385c196640c07681465de558b74dd2480b040fc03",
        "source_fingerprint_sha256": "6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53",
        "target_fingerprint_sha256": "f8f37064dc842550445b674385c196640c07681465de558b74dd2480b040fc03",
        "replacement_bytes": 8,
        "changed_spans": 1,
        "sink_write_calls": 260,
        "source_read_calls": 260,
        "source_read_bytes": 16_995_840,
    },
    "RK+MulRK": {
        "output_sha256": "ddf5d5b81d677f9056e5b48815c134e36d6674647f4e5f8c9946f989f39cf260",
        "source_fingerprint_sha256": "61a649b081c24821b02aa5e69b6ad1dc53b0232019d3668dd3776f402989c594",
        "target_fingerprint_sha256": "ddf5d5b81d677f9056e5b48815c134e36d6674647f4e5f8c9946f989f39cf260",
        "replacement_bytes": 12,
        "changed_spans": 3,
        "sink_write_calls": 4,
        "source_read_calls": 4,
        "source_read_bytes": 202_752,
    },
}
XLS_NUMERIC_OPERATION_KEYS = frozenset(
    {
        "counter_scope",
        "source_mode",
        "source_bytes",
        "fingerprint_chunk_bytes",
        "publication_chunk_bytes",
        "fingerprint_buffer_bytes",
        "publication_buffer_bytes",
        "planning_fingerprint_scans",
        "planning_fingerprint_bytes",
        "planning_fingerprint_chunks",
        "composed_source_preflight_scans",
        "composed_source_preflight_bytes",
        "composed_source_preflight_chunks",
        "target_materialization_write_pre_scans",
        "target_materialization_write_pre_bytes",
        "target_materialization_write_pre_chunks",
        "target_materialization_emission_scans",
        "target_materialization_emission_bytes",
        "target_materialization_emission_chunks",
        "target_materialization_write_post_scans",
        "target_materialization_write_post_bytes",
        "target_materialization_write_post_chunks",
        "direct_write_pre_scans",
        "direct_write_pre_bytes",
        "direct_write_pre_chunks",
        "direct_emission_scans",
        "direct_emission_bytes",
        "direct_emission_chunks",
        "direct_write_post_scans",
        "direct_write_post_bytes",
        "direct_write_post_chunks",
        "atomic_save_pre_temp_scans",
        "atomic_save_pre_temp_bytes",
        "atomic_save_pre_temp_chunks",
        "atomic_save_emission_scans",
        "atomic_save_emission_bytes",
        "atomic_save_emission_chunks",
        "atomic_save_pre_rename_scans",
        "atomic_save_pre_rename_bytes",
        "atomic_save_pre_rename_chunks",
        "candidate_reopen_logical_artifact_bytes",
        "selected_stream_logical_bytes",
        "splice_count",
        "changed_span_count",
        "replacement_bytes",
        "target_materialization_vec_bytes",
        "target_materialization_clone_bytes",
        "publication_write_calls",
        "atomic_save_event_scope",
    }
)
XLS_NUMERIC_OPERATION_MEASUREMENT_FIELDS = frozenset(
    {
        "source_mode",
        "planning_fingerprint_scans",
        "planning_fingerprint_bytes",
        "planning_fingerprint_chunks",
        "composed_source_preflight_scans",
        "composed_source_preflight_bytes",
        "composed_source_preflight_chunks",
        "target_materialization_write_pre_scans",
        "target_materialization_write_pre_bytes",
        "target_materialization_write_pre_chunks",
        "target_materialization_emission_scans",
        "target_materialization_emission_bytes",
        "target_materialization_emission_chunks",
        "target_materialization_write_post_scans",
        "target_materialization_write_post_bytes",
        "target_materialization_write_post_chunks",
        "direct_write_pre_scans",
        "direct_write_pre_bytes",
        "direct_write_pre_chunks",
        "direct_emission_scans",
        "direct_emission_bytes",
        "direct_emission_chunks",
        "direct_write_post_scans",
        "direct_write_post_bytes",
        "direct_write_post_chunks",
        "atomic_save_pre_temp_scans",
        "atomic_save_pre_temp_bytes",
        "atomic_save_pre_temp_chunks",
        "atomic_save_emission_scans",
        "atomic_save_emission_bytes",
        "atomic_save_emission_chunks",
        "atomic_save_pre_rename_scans",
        "atomic_save_pre_rename_bytes",
        "atomic_save_pre_rename_chunks",
    }
)
XLS_NUMERIC_TIMING_VECTOR_FIELDS = frozenset(
    {"edit_ns", "set_ns", "commit_ns", "publication_ns", "total_ns"}
)
XLS_NUMERIC_SOURCE_COMMON_KEYS = frozenset(
    {
        "source_counter_scope",
        "implementation",
        "family",
        "source_backed",
        "target_artifact_retained_at_commit",
        "target_artifact_materialized_at_commit",
        "patch_or_inverse_supported",
        "update_count",
        "sample_count",
        "input_cfb_bytes",
        "output_cfb_bytes",
        "source_workbook_bytes",
        "target_workbook_bytes",
        "sink_capacity_bytes",
        "expected_output_sha256",
        "owned_input_scope",
        *XLS_NUMERIC_TIMING_VECTOR_FIELDS,
        "complete_target_materialized_bytes",
        "sink_bytes",
        "sink_write_calls",
        "sink_digests",
        "source_bytes",
        "source_workbook_bytes_per_sample",
        "target_workbook_bytes_per_sample",
    }
)
XLS_NUMERIC_SOURCE_OPTIONAL_KEYS = frozenset(
    {
        "splice_count",
        "replacement_bytes",
        "changed_spans",
        "source_fingerprints",
        "target_fingerprints",
        "operation_evidence_schema",
        "operation_evidence",
    }
)


def _validate_xls_numeric_vector(
    value: Any, location: str, sample_count: int, *, digest: bool = False
) -> list[Any]:
    if not isinstance(value, list) or len(value) != sample_count:
        raise AbbaSummaryInputError(
            f"{location} must be a list with {sample_count} samples"
        )
    for index, item in enumerate(value):
        item_location = f"{location}[{index}]"
        if digest:
            _validate_output_sha256(item, item_location)
        else:
            _u64(item, item_location)
    return value


def _validate_xls_numeric_sample_order(
    elapsed: dict[str, Any], location: str, sample_count: int
) -> tuple[list[int], list[int]]:
    samples = elapsed.get("samples")
    if not isinstance(samples, list) or len(samples) != sample_count:
        raise AbbaSummaryInputError(
            f"{location}.samples must be a list with {sample_count} samples"
        )
    sample_values = [
        _u64(value, f"{location}.samples[{index}]")
        for index, value in enumerate(samples)
    ]
    sample_order = elapsed.get("sample_order")
    if not isinstance(sample_order, list) or len(sample_order) != sample_count:
        raise AbbaSummaryInputError(
            f"{location}.sample_order must be a list with {sample_count} samples"
        )
    order_values = [
        _u64(value, f"{location}.sample_order[{index}]")
        for index, value in enumerate(sample_order)
    ]
    expected_order = set(range(sample_count))
    if set(order_values) != expected_order:
        raise AbbaSummaryInputError(
            f"{location}.sample_order must be a complete permutation of retained samples"
        )
    return sample_values, order_values


def _validate_xls_numeric_source_summary(
    indexed: Mapping[tuple[str, str], dict[str, Any]], label: str
) -> None:
    """Validate the native-XLS source-summary envelope and its cardinalities."""

    allowed_keys = XLS_NUMERIC_SOURCE_COMMON_KEYS | XLS_NUMERIC_SOURCE_OPTIONAL_KEYS
    for (case, corpus_identity), row in indexed.items():
        is_native_selector = case in XLS_NUMERIC_CASE_CONTRACTS
        corpus_value = row.get("corpus")
        native_corpus_family = next(
            (
                family
                for family, expected in XLS_NUMERIC_CORPUS_CONTRACTS.items()
                if corpus_value == expected
            ),
            None,
        )
        if native_corpus_family is not None and not is_native_selector:
            raise AbbaSummaryInputError(
                f"{label}.{case}.corpus is a native XLS numeric corpus but the selector is not native"
            )
        source = row.get("source", _MISSING)
        if is_native_selector and (source is _MISSING or source is None):
            raise AbbaSummaryInputError(
                f"{label}.{case}.source is required for native XLS numeric evidence"
            )
        if source is _MISSING or source is None:
            continue
        source_object = _require_object(source, f"{label}.{case}.source")
        numeric = source_object.get("xls_numeric", _MISSING)
        if is_native_selector and (numeric is _MISSING or numeric is None):
            raise AbbaSummaryInputError(
                f"{label}.{case}.source.xls_numeric is required for native XLS numeric evidence"
            )
        if numeric is _MISSING or numeric is None:
            continue
        location = f"{label}.{case}.source.xls_numeric"
        numeric_object = _require_object(numeric, location)
        contract = XLS_NUMERIC_CASE_CONTRACTS.get(case)
        if contract is None:
            raise AbbaSummaryInputError(
                f"{location}.case is not one of the six XLS numeric selectors"
            )
        expected_implementation, expected_family, expected_updates = contract
        corpus = _require_object(row.get("corpus"), f"{label}.{case}.corpus")
        try:
            parsed_corpus_identity = json.loads(corpus_identity)
        except (TypeError, ValueError) as error:
            raise AbbaSummaryInputError(
                f"{location} corpus identity is not valid canonical JSON"
            ) from error
        if parsed_corpus_identity != corpus:
            raise AbbaSummaryInputError(
                f"{location} corpus identity does not match result.corpus"
            )
        expected_corpus = XLS_NUMERIC_CORPUS_CONTRACTS[expected_family]
        if corpus != expected_corpus:
            raise AbbaSummaryInputError(
                f"{label}.{case}.corpus does not match the exact {expected_family} native corpus manifest"
            )
        corpus_bytes = _u64(
            corpus.get("archive_bytes"),
            f"{label}.{case}.corpus.archive_bytes",
            positive=True,
        )
        if corpus_bytes > XLS_NUMERIC_CFB_MAX_BYTES:
            raise AbbaSummaryInputError(
                f"{label}.{case}.corpus.archive_bytes exceeds the 2 GiB CFB limit"
            )
        _validate_output_sha256(
            corpus.get("archive_sha256"),
            f"{label}.{case}.corpus.archive_sha256",
        )
        authoritative_output_digest = _validate_output_sha256(
            row.get("output_sha256"), f"{label}.{case}.output_sha256"
        )
        output_contract = XLS_NUMERIC_OUTPUT_CONTRACTS[expected_family]
        if authoritative_output_digest != output_contract["output_sha256"]:
            raise AbbaSummaryInputError(
                f"{label}.{case}.output_sha256 differs from the exact {expected_family} native output"
            )
        sink = _require_object(row.get("sink"), f"{label}.{case}.sink")
        authoritative_sink_bytes = _u64(
            sink.get("accepted_bytes"),
            f"{label}.{case}.sink.accepted_bytes",
            positive=True,
        )
        authoritative_sink_writes = _u64(
            sink.get("write_calls"),
            f"{label}.{case}.sink.write_calls",
            positive=True,
        )
        if authoritative_sink_writes != output_contract["sink_write_calls"]:
            raise AbbaSummaryInputError(
                f"{label}.{case}.sink.write_calls differs from the exact {expected_family} native sink contract"
            )
        missing = XLS_NUMERIC_SOURCE_COMMON_KEYS - set(numeric_object)
        if missing:
            raise AbbaSummaryInputError(
                f"{location} is missing required keys: {sorted(missing)}"
            )
        unknown = set(numeric_object) - allowed_keys
        if unknown:
            raise AbbaSummaryInputError(
                f"{location} has unknown keys: {sorted(unknown)}"
            )
        implementation = _required_nonempty_string(
            numeric_object["implementation"], location, "implementation"
        )
        if implementation not in {"eager", "source_backed", "plan_only"}:
            raise AbbaSummaryInputError(f"{location}.implementation has an unknown value")
        if implementation != expected_implementation:
            raise AbbaSummaryInputError(
                f"{location}.implementation disagrees with selector {case!r}"
            )
        source_backed = _required_bool(
            numeric_object["source_backed"], location, "source_backed"
        )
        if source_backed != (implementation != "eager"):
            raise AbbaSummaryInputError(
                f"{location}.source_backed disagrees with implementation"
            )
        family = _required_nonempty_string(numeric_object["family"], location, "family")
        if family not in {"Number", "RK+MulRK"}:
            raise AbbaSummaryInputError(f"{location}.family has an unknown value")
        if family != expected_family:
            raise AbbaSummaryInputError(
                f"{location}.family disagrees with selector {case!r}"
            )
        _required_nonempty_string(numeric_object["source_counter_scope"], location, "source_counter_scope")
        owned_input_scope = _required_nonempty_string(
            numeric_object["owned_input_scope"], location, "owned_input_scope"
        )
        if owned_input_scope != XLS_NUMERIC_OWNED_INPUT_SCOPE:
            raise AbbaSummaryInputError(
                f"{location}.owned_input_scope is not the complete in-memory CFB contract"
            )
        update_count = _u64(numeric_object["update_count"], f"{location}.update_count", positive=True)
        if update_count != expected_updates:
            raise AbbaSummaryInputError(
                f"{location}.update_count disagrees with selector {case!r}"
            )
        sample_count = _u64(numeric_object["sample_count"], f"{location}.sample_count", positive=True)
        elapsed = _require_object(row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
        elapsed_samples, sample_order = _validate_xls_numeric_sample_order(
            elapsed, f"{label}.{case}.elapsed_ns", sample_count
        )
        source_read_calls = _validate_xls_numeric_vector(
            source_object.get("read_calls", _MISSING),
            f"{label}.{case}.source.read_calls",
            sample_count,
        )
        source_read_bytes = _validate_xls_numeric_vector(
            source_object.get("read_bytes", _MISSING),
            f"{label}.{case}.source.read_bytes",
            sample_count,
        )
        input_bytes = _u64(numeric_object["input_cfb_bytes"], f"{location}.input_cfb_bytes", positive=True)
        output_bytes = _u64(numeric_object["output_cfb_bytes"], f"{location}.output_cfb_bytes", positive=True)
        if input_bytes != corpus_bytes:
            raise AbbaSummaryInputError(
                f"{location}.input_cfb_bytes differs from corpus.archive_bytes"
            )
        if output_bytes != input_bytes:
            raise AbbaSummaryInputError(
                f"{location}.output_cfb_bytes differs from input_cfb_bytes"
            )
        expected_source_read_calls = output_contract["source_read_calls"]
        expected_source_read_bytes = output_contract["source_read_bytes"]
        if expected_source_read_bytes != input_bytes:
            raise AbbaSummaryInputError(
                f"{location} native source-read byte contract disagrees with input_cfb_bytes"
            )
        if expected_source_read_calls != authoritative_sink_writes:
            raise AbbaSummaryInputError(
                f"{location} native source-read call contract disagrees with result.sink.write_calls"
            )
        for index in range(sample_count):
            if source_read_calls[index] == 0:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.read_calls must be positive"
                )
            if source_read_calls[index] != expected_source_read_calls:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.read_calls differs from the exact {expected_family} native contract and source identity"
                )
            if source_read_bytes[index] != input_bytes:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.read_bytes disagrees with input_cfb_bytes"
                )
            if source_read_bytes[index] != expected_source_read_bytes:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.read_bytes differs from the exact {expected_family} native contract"
                )
        source_workbook_bytes = _u64(
            numeric_object["source_workbook_bytes"], f"{location}.source_workbook_bytes", positive=True
        )
        target_workbook_bytes = _u64(
            numeric_object["target_workbook_bytes"], f"{location}.target_workbook_bytes", positive=True
        )
        expected_workbook_sizes = XLS_NUMERIC_WORKBOOK_SIZE_CONTRACTS[expected_family]
        if source_workbook_bytes != expected_workbook_sizes["source_workbook_bytes"]:
            raise AbbaSummaryInputError(
                f"{location}.source_workbook_bytes disagrees with the native {expected_family} corpus"
            )
        if target_workbook_bytes != expected_workbook_sizes["target_workbook_bytes"]:
            raise AbbaSummaryInputError(
                f"{location}.target_workbook_bytes disagrees with the native {expected_family} corpus"
            )
        sink_capacity = _u64(
            numeric_object["sink_capacity_bytes"], f"{location}.sink_capacity_bytes", positive=True
        )
        if sink_capacity != output_bytes:
            raise AbbaSummaryInputError(f"{location}.sink_capacity_bytes differs from output_cfb_bytes")
        expected_digest = _validate_output_sha256(
            numeric_object["expected_output_sha256"], f"{location}.expected_output_sha256"
        )
        if expected_digest != authoritative_output_digest:
            raise AbbaSummaryInputError(
                f"{location}.expected_output_sha256 differs from result.output_sha256"
            )
        if authoritative_sink_bytes != output_bytes:
            raise AbbaSummaryInputError(
                f"{location}.result.sink.accepted_bytes differs from output_cfb_bytes"
            )
        expected_retained = implementation != "plan_only"
        for field, expected in (
            ("target_artifact_retained_at_commit", expected_retained),
            ("target_artifact_materialized_at_commit", expected_retained),
            ("patch_or_inverse_supported", expected_retained),
        ):
            if _required_bool(numeric_object[field], location, field) != expected:
                raise AbbaSummaryInputError(
                    f"{location}.{field} disagrees with implementation"
                )

        timing_vectors = {
            field: _validate_xls_numeric_vector(
                numeric_object[field], f"{location}.{field}", sample_count
            )
            for field in XLS_NUMERIC_TIMING_VECTOR_FIELDS
        }
        complete = _validate_xls_numeric_vector(
            numeric_object["complete_target_materialized_bytes"],
            f"{location}.complete_target_materialized_bytes",
            sample_count,
        )
        sink_bytes = _validate_xls_numeric_vector(
            numeric_object["sink_bytes"], f"{location}.sink_bytes", sample_count
        )
        sink_writes = _validate_xls_numeric_vector(
            numeric_object["sink_write_calls"], f"{location}.sink_write_calls", sample_count
        )
        sink_digests = _validate_xls_numeric_vector(
            numeric_object["sink_digests"], f"{location}.sink_digests", sample_count, digest=True
        )
        source_sizes = _validate_xls_numeric_vector(
            numeric_object["source_bytes"], f"{location}.source_bytes", sample_count
        )
        source_stream_sizes = _validate_xls_numeric_vector(
            numeric_object["source_workbook_bytes_per_sample"],
            f"{location}.source_workbook_bytes_per_sample",
            sample_count,
        )
        target_stream_sizes = _validate_xls_numeric_vector(
            numeric_object["target_workbook_bytes_per_sample"],
            f"{location}.target_workbook_bytes_per_sample",
            sample_count,
        )
        expected_materialized = output_bytes if expected_retained else 0
        for index in range(sample_count):
            phase_sum = sum(
                timing_vectors[field][index]
                for field in ("edit_ns", "set_ns", "commit_ns", "publication_ns")
            )
            if phase_sum > U64_MAX:
                raise AbbaSummaryInputError(
                    f"{location}.total_ns phase sum exceeds u64"
                )
            if timing_vectors["total_ns"][index] != phase_sum:
                raise AbbaSummaryInputError(
                    f"{location}.total_ns must equal the per-sample phase sum"
                )
            if complete[index] != expected_materialized:
                raise AbbaSummaryInputError(f"{location}.complete_target_materialized_bytes disagrees with implementation")
            if sink_bytes[index] != output_bytes:
                raise AbbaSummaryInputError(f"{location}.sink_bytes disagrees with output_cfb_bytes")
            if sink_writes[index] == 0:
                raise AbbaSummaryInputError(f"{location}.sink_write_calls must be positive")
            if sink_writes[index] != authoritative_sink_writes:
                raise AbbaSummaryInputError(
                    f"{location}.sink_write_calls differs from result.sink.write_calls"
                )
            if sink_writes[index] != output_contract["sink_write_calls"]:
                raise AbbaSummaryInputError(
                    f"{location}.sink_write_calls differs from the exact {expected_family} native sink contract"
                )
            if sink_digests[index] != expected_digest:
                raise AbbaSummaryInputError(f"{location}.sink_digests disagrees with expected_output_sha256")
            if sink_digests[index] != output_contract["target_fingerprint_sha256"]:
                raise AbbaSummaryInputError(
                    f"{location}.sink_digests differs from the exact {expected_family} native output"
                )
            if source_sizes[index] != input_bytes:
                raise AbbaSummaryInputError(f"{location}.source_bytes disagrees with input_cfb_bytes")
            if source_stream_sizes[index] != source_workbook_bytes:
                raise AbbaSummaryInputError(f"{location}.source_workbook_bytes_per_sample disagrees with source_workbook_bytes")
            if target_stream_sizes[index] != target_workbook_bytes:
                raise AbbaSummaryInputError(f"{location}.target_workbook_bytes_per_sample disagrees with target_workbook_bytes")
        for sorted_index, original_index in enumerate(sample_order):
            if timing_vectors["total_ns"][original_index] != elapsed_samples[sorted_index]:
                raise AbbaSummaryInputError(
                    f"{location}.total_ns does not bind to elapsed_ns.samples via sample_order"
                )

        optional_vector_fields = (
            "splice_count",
            "replacement_bytes",
            "changed_spans",
            "source_fingerprints",
            "target_fingerprints",
        )
        optional_present = {field: field in numeric_object for field in optional_vector_fields}
        if any(value != source_backed for value in optional_present.values()):
            raise AbbaSummaryInputError(
                f"{location} source-backed optional vector presence is inconsistent"
            )
        for field in optional_vector_fields:
            if field not in numeric_object:
                continue
            values = _validate_xls_numeric_vector(
                numeric_object[field],
                f"{location}.{field}",
                sample_count,
                digest=field.endswith("fingerprints"),
            )
            if field == "splice_count" and any(value != update_count for value in values):
                raise AbbaSummaryInputError(f"{location}.splice_count disagrees with update_count")
            if field == "replacement_bytes":
                if any(value == 0 for value in values):
                    raise AbbaSummaryInputError(f"{location}.replacement_bytes must be positive")
                if any(value != output_contract["replacement_bytes"] for value in values):
                    raise AbbaSummaryInputError(
                        f"{location}.replacement_bytes differs from the exact {expected_family} native contract"
                    )
            if field == "changed_spans":
                if any(value == 0 for value in values):
                    raise AbbaSummaryInputError(f"{location}.changed_spans must be positive")
                if any(value != output_contract["changed_spans"] for value in values):
                    raise AbbaSummaryInputError(
                        f"{location}.changed_spans differs from the exact {expected_family} native contract"
                    )
            if field == "source_fingerprints" and any(
                value != output_contract["source_fingerprint_sha256"] for value in values
            ):
                raise AbbaSummaryInputError(
                    f"{location}.source_fingerprints differs from the exact {expected_family} native corpus"
                )
            if field == "target_fingerprints" and any(
                value != output_contract["target_fingerprint_sha256"] for value in values
            ):
                raise AbbaSummaryInputError(
                    f"{location}.target_fingerprints differs from the exact {expected_family} native output"
                )

        schema = numeric_object.get("operation_evidence_schema", _MISSING)
        operation_present = "operation_evidence" in numeric_object
        if implementation == "eager":
            if numeric_object["source_counter_scope"] != XLS_NUMERIC_LEGACY_SOURCE_COUNTER_SCOPE:
                raise AbbaSummaryInputError(
                    f"{location}.source_counter_scope is invalid for eager implementation"
                )
            if schema is not _MISSING or operation_present:
                raise AbbaSummaryInputError(
                    f"{location} eager implementation must not carry operation evidence"
                )
        else:
            if numeric_object["source_counter_scope"] != XLS_NUMERIC_CURRENT_SOURCE_COUNTER_SCOPE:
                raise AbbaSummaryInputError(
                    f"{location}.source_counter_scope is not the current operation-evidence discriminator"
                )
            if schema is _MISSING:
                raise AbbaSummaryInputError(
                    f"{location} source-backed implementation requires operation_evidence_schema"
                )
            if schema != XLS_NUMERIC_OPERATION_EVIDENCE_SCHEMA:
                raise AbbaSummaryInputError(f"{location}.operation_evidence_schema is invalid")
            if not operation_present:
                raise AbbaSummaryInputError(
                    f"{location} current source-backed row requires operation_evidence"
                )
            if not isinstance(numeric_object["operation_evidence"], list):
                raise AbbaSummaryInputError(f"{location}.operation_evidence must be a list")


def _validate_xls_numeric_operation_evidence(
    indexed: Mapping[tuple[str, str], dict[str, Any]],
    label: str,
    *,
    report_role: str,
) -> None:
    """Validate additive native-XLS logical operation-shape evidence.

    The low-level CFB shape is deliberately redundant with its source length:
    recomputing every phase here catches report tampering and accidental
    generic/owned policy drift.  This validator does not infer timings or
    speedups, and it accepts historical rows that predate the additive field
    by validating only rows that actually carry ``operation_evidence``.
    """

    numeric_fields = XLS_NUMERIC_OPERATION_KEYS - {
        "counter_scope",
        "source_mode",
        "atomic_save_event_scope",
    }

    def phase_shape(
        evidence: dict[str, Any],
        prefix: str,
        scans: int,
        source_bytes: int,
        fingerprint_chunks: int,
        emission: bool = False,
    ) -> None:
        _u64(evidence[f"{prefix}_scans"], f"{prefix}.scans")
        _u64(evidence[f"{prefix}_bytes"], f"{prefix}.bytes")
        _u64(evidence[f"{prefix}_chunks"], f"{prefix}.chunks")
        expected_bytes = source_bytes * 2
        expected_chunks = (source_bytes + (65_536 if emission else 1_048_576) - 1) // (
            65_536 if emission else 1_048_576
        )
        if evidence[f"{prefix}_scans"] != scans:
            raise AbbaSummaryInputError(
                f"{prefix} scan count does not match source-mode contract"
            )
        if evidence[f"{prefix}_bytes"] != expected_bytes * scans:
            raise AbbaSummaryInputError(
                f"{prefix} logical bytes do not match source-mode contract"
            )
        if evidence[f"{prefix}_chunks"] != expected_chunks * scans:
            raise AbbaSummaryInputError(
                f"{prefix} logical chunks do not match source-mode contract"
            )

    for (case, _corpus_identity), row in indexed.items():
        source = row.get("source")
        if source is None:
            continue
        source_object = _require_object(source, f"{label}.{case}.source")
        numeric = source_object.get("xls_numeric")
        if numeric is None:
            continue
        numeric_object = _require_object(numeric, f"{label}.{case}.source.xls_numeric")
        schema = numeric_object.get("operation_evidence_schema", _MISSING)
        if schema is _MISSING:
            # Historical schema-1 operation evidence used a different field
            # shape. The explicit absence of the new schema marker, together
            # with the top-level legacy discriminator, keeps it opaque here.
            continue
        if schema != XLS_NUMERIC_OPERATION_EVIDENCE_SCHEMA:
            raise AbbaSummaryInputError(
                f"{label}.{case}.source.xls_numeric.operation_evidence_schema is invalid"
            )
        operation = numeric_object.get("operation_evidence", _MISSING)
        if operation is _MISSING:
            continue
        location = f"{label}.{case}.source.xls_numeric.operation_evidence"
        if operation is None or not isinstance(operation, list):
            raise AbbaSummaryInputError(f"{location} must be a list when present")
        if numeric_object.get("source_backed") is not True:
            raise AbbaSummaryInputError(
                f"{location} requires source.xls_numeric.source_backed=true"
            )
        implementation = numeric_object.get("implementation")
        if implementation not in {"source_backed", "plan_only"}:
            raise AbbaSummaryInputError(
                f"{location} requires source.xls_numeric implementation source_backed or plan_only"
            )
        elapsed = _require_object(row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
        elapsed_samples = elapsed.get("samples")
        if not isinstance(elapsed_samples, list) or len(operation) != len(elapsed_samples):
            raise AbbaSummaryInputError(
                f"{location} cardinality must equal elapsed_ns.samples"
            )
        sample_count = numeric_object.get("sample_count")
        if isinstance(sample_count, bool) or not isinstance(sample_count, int):
            raise AbbaSummaryInputError(
                f"{label}.{case}.source.xls_numeric.sample_count must be an integer"
            )
        if sample_count != len(operation):
            raise AbbaSummaryInputError(
                f"{location} cardinality must equal source.xls_numeric.sample_count"
            )

        source_bytes = _u64(
            numeric_object.get("input_cfb_bytes"),
            f"{label}.{case}.source.xls_numeric.input_cfb_bytes",
        )
        output_bytes = _u64(
            numeric_object.get("output_cfb_bytes"),
            f"{label}.{case}.source.xls_numeric.output_cfb_bytes",
        )
        target_workbook_bytes = _u64(
            numeric_object.get("target_workbook_bytes"),
            f"{label}.{case}.source.xls_numeric.target_workbook_bytes",
        )
        splice_values = numeric_object.get("splice_count")
        replacement_values = numeric_object.get("replacement_bytes")
        changed_values = numeric_object.get("changed_spans")
        sink_write_values = numeric_object.get("sink_write_calls")
        sink = _require_object(row.get("sink"), f"{label}.{case}.sink")
        authoritative_sink_writes = _u64(
            sink.get("write_calls"),
            f"{label}.{case}.sink.write_calls",
            positive=True,
        )
        for field, values in (
            ("splice_count", splice_values),
            ("replacement_bytes", replacement_values),
            ("changed_spans", changed_values),
            ("sink_write_calls", sink_write_values),
        ):
            if not isinstance(values, list) or len(values) != len(operation):
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.xls_numeric.{field} cardinality differs from operation_evidence"
                )
            for index, value in enumerate(values):
                _u64(value, f"{label}.{case}.source.xls_numeric.{field}[{index}]")
        if any(value != authoritative_sink_writes for value in sink_write_values):
            raise AbbaSummaryInputError(
                f"{label}.{case}.source.xls_numeric.sink_write_calls differs from result.sink.write_calls"
            )

        materialized = numeric_object.get("target_artifact_materialized_at_commit")
        if not isinstance(materialized, bool):
            raise AbbaSummaryInputError(
                f"{label}.{case}.source.xls_numeric.target_artifact_materialized_at_commit must be a boolean"
            )
        complete_materialized = numeric_object.get("complete_target_materialized_bytes")
        if not isinstance(complete_materialized, list) or len(complete_materialized) != len(operation):
            raise AbbaSummaryInputError(
                f"{label}.{case}.source.xls_numeric.complete_target_materialized_bytes cardinality is invalid"
            )

        for index, item in enumerate(operation):
            item_location = f"{location}[{index}]"
            evidence = _require_object(item, item_location)
            if set(evidence) != XLS_NUMERIC_OPERATION_KEYS:
                raise AbbaSummaryInputError(
                    f"{item_location} keys mismatch: expected {sorted(XLS_NUMERIC_OPERATION_KEYS)}"
                )
            for field in numeric_fields:
                _u64(evidence[field], f"{item_location}.{field}")
            counter_scope = _required_nonempty_string(
                evidence["counter_scope"], item_location, "counter_scope"
            )
            if counter_scope != XLS_NUMERIC_OPERATION_COUNTER_SCOPE:
                raise AbbaSummaryInputError(
                    f"{item_location}.counter_scope is not the validated CFB scope"
                )
            mode = _required_nonempty_string(
                evidence["source_mode"], item_location, "source_mode"
            )
            if mode not in {"generic_read_at", "owned_immutable_arc"}:
                raise AbbaSummaryInputError(
                    f"{item_location}.source_mode has an unknown value"
                )
            if implementation == "plan_only":
                expected_mode = "owned_immutable_arc"
            else:
                expected_mode = (
                    "generic_read_at"
                    if report_role in {"a1", "a2"}
                    else "owned_immutable_arc"
                )
            if mode != expected_mode:
                raise AbbaSummaryInputError(
                    f"{item_location}.source_mode violates the ABBA implementation-role contract"
                )
            atomic_scope = _required_nonempty_string(
                evidence["atomic_save_event_scope"], item_location, "atomic_save_event_scope"
            )
            if atomic_scope != XLS_NUMERIC_OPERATION_ATOMIC_SCOPE:
                raise AbbaSummaryInputError(
                    f"{item_location}.atomic_save_event_scope is not the public-save contract"
                )
            if evidence["source_bytes"] != source_bytes:
                raise AbbaSummaryInputError(
                    f"{item_location}.source_bytes differs from input_cfb_bytes"
                )
            if evidence["fingerprint_chunk_bytes"] != 1_048_576:
                raise AbbaSummaryInputError(
                    f"{item_location}.fingerprint_chunk_bytes has an invalid contract"
                )
            if evidence["publication_chunk_bytes"] != 65_536:
                raise AbbaSummaryInputError(
                    f"{item_location}.publication_chunk_bytes has an invalid contract"
                )
            if evidence["fingerprint_buffer_bytes"] != min(source_bytes, 1_048_576):
                raise AbbaSummaryInputError(
                    f"{item_location}.fingerprint_buffer_bytes has an invalid scope"
                )
            if evidence["publication_buffer_bytes"] != 65_536:
                raise AbbaSummaryInputError(
                    f"{item_location}.publication_buffer_bytes has an invalid scope"
                )
            fingerprint_chunks = (source_bytes + 1_048_576 - 1) // 1_048_576
            publication_chunks = (source_bytes + 65_536 - 1) // 65_536
            planning_scans = 2 if mode == "generic_read_at" else 1
            fenced_scans = 1 if mode == "generic_read_at" else 0
            if evidence["planning_fingerprint_scans"] != planning_scans:
                raise AbbaSummaryInputError(
                    f"{item_location}.planning_fingerprint_scans violates source-mode policy"
                )
            if evidence["planning_fingerprint_bytes"] != source_bytes * 2 * planning_scans:
                raise AbbaSummaryInputError(
                    f"{item_location}.planning_fingerprint_bytes violates source-mode policy"
                )
            if evidence["planning_fingerprint_chunks"] != fingerprint_chunks * planning_scans:
                raise AbbaSummaryInputError(
                    f"{item_location}.planning_fingerprint_chunks violates source-mode policy"
                )
            phase_shape(
                evidence,
                "composed_source_preflight",
                1,
                source_bytes,
                fingerprint_chunks,
            )
            for prefix in (
                "target_materialization_write_pre",
                "target_materialization_write_post",
                "direct_write_pre",
                "direct_write_post",
                "atomic_save_pre_temp",
                "atomic_save_pre_rename",
            ):
                phase_shape(
                    evidence,
                    prefix,
                    fenced_scans,
                    source_bytes,
                    fingerprint_chunks,
                )
            for prefix in (
                "target_materialization_emission",
                "direct_emission",
                "atomic_save_emission",
            ):
                phase_shape(
                    evidence,
                    prefix,
                    1,
                    source_bytes,
                    fingerprint_chunks,
                    emission=True,
                )
            if evidence["candidate_reopen_logical_artifact_bytes"] != (
                output_bytes if implementation == "source_backed" else source_bytes
            ):
                raise AbbaSummaryInputError(
                    f"{item_location}.candidate_reopen_logical_artifact_bytes violates publication contract"
                )
            if evidence["selected_stream_logical_bytes"] != target_workbook_bytes:
                raise AbbaSummaryInputError(
                    f"{item_location}.selected_stream_logical_bytes differs from target workbook length"
                )
            if evidence["splice_count"] != splice_values[index]:
                raise AbbaSummaryInputError(
                    f"{item_location}.splice_count differs from source summary"
                )
            if evidence["replacement_bytes"] != replacement_values[index]:
                raise AbbaSummaryInputError(
                    f"{item_location}.replacement_bytes differs from source summary"
                )
            if evidence["changed_span_count"] != changed_values[index]:
                raise AbbaSummaryInputError(
                    f"{item_location}.changed_span_count differs from source summary"
                )
            if evidence["publication_write_calls"] != sink_write_values[index]:
                raise AbbaSummaryInputError(
                    f"{item_location}.publication_write_calls differs from sink summary"
                )
            if evidence["publication_write_calls"] != authoritative_sink_writes:
                raise AbbaSummaryInputError(
                    f"{item_location}.publication_write_calls differs from result.sink.write_calls"
                )
            expected_materialized = output_bytes if materialized else 0
            if evidence["target_materialization_vec_bytes"] != expected_materialized:
                raise AbbaSummaryInputError(
                    f"{item_location}.target_materialization_vec_bytes violates materialization contract"
                )
            if evidence["target_materialization_clone_bytes"] != expected_materialized:
                raise AbbaSummaryInputError(
                    f"{item_location}.target_materialization_clone_bytes violates materialization contract"
                )
            _u64(
                complete_materialized[index],
                f"{label}.{case}.source.xls_numeric.complete_target_materialized_bytes[{index}]",
            )
            if complete_materialized[index] != expected_materialized:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.source.xls_numeric.complete_target_materialized_bytes disagrees with operation evidence"
                )


def _source_identity_projection(value: Any) -> Any:
    """Remove named measurements while retaining every source identity field."""
    if not isinstance(value, dict):
        return value
    projected = dict(value)
    cfb_open_stream = projected.get("cfb_open_stream")
    if isinstance(cfb_open_stream, dict):
        projected_cfb = dict(cfb_open_stream)
        for field in CFB_OPEN_STREAM_SOURCE_MEASUREMENTS:
            projected_cfb.pop(field, None)
        projected["cfb_open_stream"] = projected_cfb
    ods_source_cell = projected.get("ods_source_cell")
    if isinstance(ods_source_cell, dict):
        projected_ods = dict(ods_source_cell)
        for field in ODS_SOURCE_CELL_MEASUREMENTS:
            projected_ods.pop(field, None)
        projected["ods_source_cell"] = projected_ods
        # The harness duplicates the aggregate replay counters at the source
        # root.  They remain measurements for these exact ODS rows.
        projected.pop("read_calls", None)
        projected.pop("read_bytes", None)
    xlsx_cell_values = projected.get("xlsx_cell_values")
    if isinstance(xlsx_cell_values, dict):
        projected_xlsx = dict(xlsx_cell_values)
        for field in XLSX_CELL_VALUES_SOURCE_MEASUREMENTS:
            projected_xlsx.pop(field, None)
        projected["xlsx_cell_values"] = projected_xlsx
    xls_numeric = projected.get("xls_numeric")
    if isinstance(xls_numeric, dict):
        projected_xls = dict(xls_numeric)
        for field in XLS_NUMERIC_TIMING_VECTOR_FIELDS:
            projected_xls.pop(field, None)
        operation_evidence = projected_xls.get("operation_evidence")
        if isinstance(operation_evidence, list):
            projected_xls["operation_evidence"] = [
                {
                    key: item[key]
                    for key in item
                    if key not in XLS_NUMERIC_OPERATION_MEASUREMENT_FIELDS
                }
                if isinstance(item, dict)
                else item
                for item in operation_evidence
            ]
        projected["xls_numeric"] = projected_xls
    opc_source_overlay = projected.get("opc_source_overlay")
    if isinstance(opc_source_overlay, dict):
        projected_overlay = dict(opc_source_overlay)
        for field in _OPC_SOURCE_OVERLAY_DYNAMIC_FIELDS:
            projected_overlay.pop(field, None)
        projected["opc_source_overlay"] = projected_overlay
        for field in _OPC_SOURCE_OVERLAY_SOURCE_VECTOR_FIELDS:
            projected.pop(field, None)
    docx_section_layout = projected.get("docx_section_layout")
    if isinstance(docx_section_layout, dict):
        projected_docx = dict(docx_section_layout)
        for field in _DOCX_SECTION_LAYOUT_DYNAMIC_FIELDS:
            projected_docx.pop(field, None)
        projected["docx_section_layout"] = projected_docx
        for field in _DOCX_SECTION_LAYOUT_ROOT_SOURCE_VECTOR_FIELDS:
            projected.pop(field, None)
    return projected


def _compare_row_identity(
    rows: Mapping[str, dict[str, Any]],
    field: str,
    location: str,
) -> tuple[str, bool, str | None, Any]:
    identities = {
        label: _identity_value(rows[label], field, f"{label}.{location}")
        for label in LEG_ORDER
    }
    expected_presence = identities["a1"][0]
    if any(identity[0] != expected_presence for identity in identities.values()):
        raise AbbaSummaryInputError(
            f"{location} {field} presence differs between ABBA legs"
        )
    if not expected_presence:
        return "consistently_absent", False, None, None
    expected_identity = identities["a1"][1]
    if any(identity[1] != expected_identity for identity in identities.values()):
        raise AbbaSummaryInputError(
            f"{location} {field} identity differs between ABBA legs"
        )
    value = rows["a1"][field]
    if field == "source":
        value = _source_identity_projection(value)
    return "verified_equal", True, expected_identity, value


def _validate_output_sha256(value: Any, location: str) -> str:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        raise AbbaSummaryInputError(f"{location} must be a lowercase SHA-256 hex string")
    return value


def _coerce_reports(
    a1: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    b1: Mapping[str, Any] | None,
    b2: Mapping[str, Any] | None,
    a2: Mapping[str, Any] | None,
) -> tuple[Mapping[str, Any], Mapping[str, Any], Mapping[str, Any], Mapping[str, Any]]:
    if b1 is None and b2 is None and a2 is None:
        if isinstance(a1, Mapping):
            for labels in (LEG_ORDER, tuple(label.upper() for label in LEG_ORDER)):
                if all(label in a1 for label in labels):
                    return tuple(a1[label] for label in labels)  # type: ignore[return-value]
            raise AbbaSummaryInputError(
                f"reports mapping must contain {list(LEG_ORDER)!r}"
            )
        if isinstance(a1, Sequence) and not isinstance(a1, (str, bytes)) and len(a1) == 4:
            return tuple(a1)  # type: ignore[return-value]
        raise AbbaSummaryInputError("reports must be four reports in A1,B1,B2,A2 order")
    if b1 is None or b2 is None or a2 is None:
        raise AbbaSummaryInputError("all four ABBA reports are required")
    return a1, b1, b2, a2


def _projection_integrity(value: Any) -> str:
    """Return an integrity digest for an internal validated projection.

    Projection rows contain tuple keys, so they are not directly JSON
    serializable.  The tagged normalization below is deterministic for the
    bounded projection tree and deliberately ignores only the checker-added
    top-level ``_canonical_sha256`` transport key.  It detects accidental or
    adversarial mutation of the recomputed statistic markers before the
    private projected-summary path consumes them.
    """

    def normalize(item: Any, *, root: bool = False) -> Any:
        if isinstance(item, dict):
            pairs = []
            for key, child in item.items():
                if root and key == "_canonical_sha256":
                    continue
                pairs.append((normalize(key), normalize(child)))
            pairs.sort(
                key=lambda pair: json.dumps(
                    pair[0], sort_keys=True, separators=(",", ":"), allow_nan=False
                )
            )
            return {"__dict__": pairs}
        if isinstance(item, tuple):
            return {"__tuple__": [normalize(child) for child in item]}
        if isinstance(item, list):
            return [normalize(child) for child in item]
        if isinstance(item, frozenset):
            normalized = [normalize(child) for child in item]
            normalized.sort(
                key=lambda child: json.dumps(
                    child, sort_keys=True, separators=(",", ":"), allow_nan=False
                )
            )
            return {"__frozenset__": normalized}
        if item is None or isinstance(item, (bool, int, float, str)):
            return item
        raise AbbaSummaryInputError(
            f"projection contains unsupported provenance value {type(item).__name__}"
        )

    try:
        payload = json.dumps(
            normalize(value, root=True),
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise AbbaSummaryInputError(
            f"cannot fingerprint validated projection: {error}"
        ) from error
    return hashlib.sha256(payload).hexdigest()


class _ValidatedProjection(dict[str, Any]):
    """Private dict carrier for a projection validated in this module.

    The class is intentionally not part of the module's public API.  The
    projected summarizer accepts only instances created by ``_project_report``
    and checks the integrity digest before consuming their private markers.
    This keeps a caller from fabricating a normal mapping containing trusted
    ``_elapsed_statistics`` values.
    """

    __slots__ = ("_integrity",)

    def __init__(self, value: dict[str, Any]) -> None:
        super().__init__(value)
        self._integrity = _projection_integrity(self)

    def verify_integrity(self) -> None:
        if _projection_integrity(self) != self._integrity:
            raise AbbaSummaryInputError(
                "validated projection was mutated after raw validation"
            )


def _project_report(
    report: Mapping[str, Any],
    label: str = "report",
    *,
    profile: str | None = None,
    report_role: str | None = None,
    expected_revision: str | None = None,
    minimum_samples: int = MIN_RETAINED_SAMPLES,
    raw_canonical_sha256: str | None = None,
) -> dict[str, Any]:
    """Validate one raw report and retain only summary identity channels.

    The returned mapping intentionally contains no elapsed sample values and
    no filesystem measurement payloads.  It is suitable for retaining one
    leg while the next compressed report is decoded.  The private markers are
    consumed only by :func:`_summarize_projected_reports`; raw reports always
    take the full sample-validation path.
    """

    if isinstance(minimum_samples, bool) or not isinstance(minimum_samples, int):
        raise AbbaSummaryInputError("minimum_samples must be an integer")
    if minimum_samples < MIN_RETAINED_SAMPLES:
        raise AbbaSummaryInputError(
            f"minimum_samples must be at least {MIN_RETAINED_SAMPLES}"
        )
    validated = _validate_report(
        report,
        label,
        profile=profile,
        report_role=report_role,
        raw_canonical_sha256=raw_canonical_sha256,
    )
    report_schema, tool, environment, configuration, indexed, report_sha256 = validated[:6]
    filesystem_shapes = validated[6]
    filesystem_present = validated[7]
    filesystem_identity = validated[8]
    binary_identity = validated[9]
    filesystem_child_process_ids = validated[10]
    if expected_revision is not None:
        if not isinstance(expected_revision, str) or not expected_revision:
            raise AbbaSummaryInputError("expected_revision must be a non-empty string")
        if environment["git_revision"] != expected_revision:
            raise AbbaSummaryInputError(
                f"{label}.environment.git_revision does not match expected revision"
            )
    _validate_configuration_rows(
        configuration,
        indexed,
        label,
        filesystem_shapes=filesystem_shapes,
    )
    projected_index: dict[tuple[str, str], dict[str, Any]] = {}
    for key, row in indexed.items():
        case, corpus_identity = key
        elapsed = recompute_statistics(
            row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns"
        )
        if elapsed["sample_count"] < minimum_samples:
            raise AbbaSummaryInputError(
                f"{label}.{case}.elapsed_ns has fewer than {minimum_samples} samples"
            )
        for field in ("source", "sink", "output_sha256"):
            if field in row and row[field] is not None:
                if field == "output_sha256":
                    _validate_output_sha256(
                        row[field], f"{label}.{case}.{field}"
                    )
                else:
                    _canonical_json(row[field], f"{label}.{case}.{field}")
        operation_metrics_identity = _operation_metrics_identity(
            row, f"{label}.{case}", report_schema
        )
        projected_row: dict[str, Any] = {
            "case": case,
            "corpus": json.loads(corpus_identity),
            "_elapsed_statistics": elapsed,
            "_operation_metrics_identity": operation_metrics_identity,
        }
        for field in ("source", "sink", "output_sha256"):
            if field in row:
                value = row[field]
                projected_row[field] = (
                    _source_identity_projection(value) if field == "source" else value
                )
            present, identity = _identity_value(
                row, field, f"{label}.{case}[{corpus_identity}]"
            )
            projected_row[f"_{field}_present"] = present
            projected_row[f"_{field}_identity"] = identity
        projected_index[key] = projected_row
    report_profile = detect_report_profile(report, label)
    return _ValidatedProjection({
        "_profile": report_profile,
        "schema_version": report_schema,
        "tool": tool,
        "environment": environment,
        "configuration": configuration,
        "results": projected_index,
        "report_sha256": report_sha256,
        "filesystem_shapes": tuple(sorted(filesystem_shapes)),
        "filesystem_present": filesystem_present,
        "filesystem_identity": filesystem_identity,
        "binary_identity": binary_identity,
        "filesystem_child_process_ids": filesystem_child_process_ids,
    })


def _summarize_projected_reports(
    reports: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    *,
    drift_ceilings: Mapping[str, Any] | None = None,
    cases: Iterable[str] | None = None,
    shapes: Iterable[str] | None = None,
    ceilings: Mapping[str, Any] | None = None,
    profile: str | None = None,
) -> dict[str, Any]:
    """Summarize four sequentially validated report projections."""

    if not isinstance(reports, (Mapping, Sequence)) or isinstance(reports, (str, bytes)):
        raise AbbaSummaryInputError("projected reports must be four report mappings")
    projected_values = _coerce_reports(reports, None, None, None)
    profiles: set[str] = set()
    validated: dict[str, tuple[Any, ...]] = {}
    for label, projected in zip(LEG_ORDER, projected_values):
        if not isinstance(projected, _ValidatedProjection):
            raise AbbaSummaryInputError(
                f"{label} projected report lacks private validation provenance"
            )
        projected.verify_integrity()
        detected = projected.get("_profile")
        if detected not in REPORT_PROFILES:
            raise AbbaSummaryInputError(f"{label} projected report has an invalid profile")
        profiles.add(detected)
        indexed = projected.get("results")
        if not isinstance(indexed, Mapping) or not indexed:
            raise AbbaSummaryInputError(f"{label} projected report results must be non-empty")
        normalized_index: dict[tuple[str, str], dict[str, Any]] = {}
        for key, row in indexed.items():
            if (
                not isinstance(key, tuple)
                or len(key) != 2
                or not isinstance(key[0], str)
                or not isinstance(key[1], str)
                or not isinstance(row, dict)
            ):
                raise AbbaSummaryInputError(f"{label} projected report has an invalid result key")
            if "_elapsed_statistics" not in row:
                raise AbbaSummaryInputError(
                    f"{label} projected result is missing recomputed elapsed statistics"
                )
            normalized_index[key] = row
        schema_version = projected.get("schema_version")
        tool = projected.get("tool")
        environment = projected.get("environment")
        configuration = projected.get("configuration")
        report_sha256 = projected.get("report_sha256")
        if not isinstance(schema_version, int) or not isinstance(tool, dict):
            raise AbbaSummaryInputError(f"{label} projected report root identity is invalid")
        if not isinstance(environment, dict) or not isinstance(configuration, dict):
            raise AbbaSummaryInputError(f"{label} projected report metadata is invalid")
        if not isinstance(report_sha256, str) or SHA256_RE.fullmatch(report_sha256) is None:
            raise AbbaSummaryInputError(f"{label} projected report SHA-256 is invalid")
        filesystem_shapes = projected.get("filesystem_shapes")
        if not isinstance(filesystem_shapes, (tuple, list)):
            raise AbbaSummaryInputError(f"{label} projected filesystem shapes are invalid")
        filesystem_identity = projected.get("filesystem_identity")
        if not isinstance(filesystem_identity, dict):
            raise AbbaSummaryInputError(f"{label} projected filesystem identity is invalid")
        filesystem_present = projected.get("filesystem_present")
        if not isinstance(filesystem_present, bool):
            raise AbbaSummaryInputError(f"{label} projected filesystem presence is invalid")
        filesystem_child_process_ids = projected.get("filesystem_child_process_ids")
        if not isinstance(filesystem_child_process_ids, (frozenset, set, tuple, list)):
            raise AbbaSummaryInputError(
                f"{label} projected filesystem child-process identities are invalid"
            )
        normalized_child_process_ids: frozenset[int] = frozenset()
        for child_process_id in filesystem_child_process_ids:
            _u32(child_process_id, f"{label}.projected_child_process_id", positive=True)
            normalized_child_process_ids = normalized_child_process_ids | frozenset(
                (child_process_id,)
            )
        binary_identity = projected.get("binary_identity")
        if detected == REPORT_PROFILE_CURRENT and not isinstance(binary_identity, dict):
            raise AbbaSummaryInputError(f"{label} current projection lacks binary identity")
        if detected == REPORT_PROFILE_LEGACY and binary_identity is not None:
            raise AbbaSummaryInputError(f"{label} legacy projection carries binary identity")
        validated[label] = (
            schema_version,
            tool,
            environment,
            configuration,
            normalized_index,
            report_sha256,
            frozenset(filesystem_shapes),
            filesystem_present,
            filesystem_identity,
            binary_identity,
            normalized_child_process_ids,
        )
    if len(profiles) != 1:
        raise AbbaSummaryInputError("mixed legacy-v1/current-v1 ABBA report profiles")
    report_profile = next(iter(profiles))
    if profile is not None and profile != report_profile:
        raise AbbaSummaryInputError(
            f"projected report profile {report_profile!r} does not match {profile!r}"
        )
    return _summarize_reports_impl(
        drift_ceilings=drift_ceilings,
        cases=cases,
        shapes=shapes,
        ceilings=ceilings,
        profile=report_profile,
        _validated=(report_profile, validated),
    )


def _parse_selectors(values: Iterable[str] | None) -> set[str] | None:
    if values is None:
        return None
    if isinstance(values, str):
        values = (values,)
    selectors: set[str] = set()
    for value in values:
        if not isinstance(value, str):
            raise AbbaSummaryInputError("selectors must contain only strings")
        for selector in value.split(","):
            selector = selector.strip()
            if selector:
                selectors.add(selector)
    if not selectors:
        raise AbbaSummaryInputError("selectors must contain at least one non-empty value")
    return selectors


def _delta_percent(control: float, candidate: float) -> float:
    if control <= 0:
        raise AbbaSummaryInputError("elapsed statistics must have a positive control value")
    delta = (control - candidate) / control * 100.0
    if not math.isfinite(delta):
        raise AbbaSummaryInputError("candidate reduction is not finite")
    return delta


def _drift_percent(first: float, second: float) -> float:
    if first <= 0:
        raise AbbaSummaryInputError("elapsed statistics must have positive drift baselines")
    drift = (second - first) / first * 100.0
    if not math.isfinite(drift):
        raise AbbaSummaryInputError("same-implementation drift is not finite")
    return drift


def _result_summary(
    rows: Mapping[str, dict[str, Any]],
    *,
    case: str,
    corpus_identity: str,
    drift_ceilings: Mapping[str, float],
    report_schema: int,
    projected: bool = False,
) -> dict[str, Any]:
    source_status, source_present, source_identity, source_value = _compare_row_identity(
        rows, "source", f"{case}[{corpus_identity}]"
    )
    sink_status, sink_present, sink_identity, sink_value = _compare_row_identity(
        rows, "sink", f"{case}[{corpus_identity}]"
    )
    for label in LEG_ORDER:
        if "output_sha256" in rows[label] and rows[label]["output_sha256"] is not None:
            _validate_output_sha256(
                rows[label]["output_sha256"], f"{label}.{case}.output_sha256"
            )
    output_status, output_present, output_identity, output_value = _compare_row_identity(
        rows, "output_sha256", f"{case}[{corpus_identity}]"
    )
    operation_metrics_status, operation_metrics_identity = (
        _compare_operation_metrics_identity(
            rows, f"{case}[{corpus_identity}]", report_schema, projected=projected
        )
    )

    elapsed: dict[str, dict[str, Any]] = {}
    for label in LEG_ORDER:
        if projected and "_elapsed_statistics" in rows[label]:
            elapsed[label] = _require_object(
                rows[label]["_elapsed_statistics"],
                f"{label}.{case}._elapsed_statistics",
            )
        else:
            elapsed[label] = recompute_statistics(
                rows[label].get("elapsed_ns"), f"{label}.{case}.elapsed_ns"
            )
    sample_counts = {elapsed[label]["sample_count"] for label in LEG_ORDER}
    if len(sample_counts) != 1:
        raise AbbaSummaryInputError(
            f"{case}[{corpus_identity}] ABBA legs have different sample counts"
        )

    candidate_reduction = {
        "a1_to_b1": {
            name: _delta_percent(elapsed["a1"][name], elapsed["b1"][name])
            for name in STATISTICS
        },
        "a2_to_b2": {
            name: _delta_percent(elapsed["a2"][name], elapsed["b2"][name])
            for name in STATISTICS
        },
    }
    adverse_both: list[str] = []
    drift = {
        "control": {
            name: _drift_percent(elapsed["a1"][name], elapsed["a2"][name])
            for name in STATISTICS
        },
        "candidate": {
            name: _drift_percent(elapsed["b1"][name], elapsed["b2"][name])
            for name in STATISTICS
        },
    }
    drift_within_ceiling = {
        implementation: {
            name: abs(drift[implementation][name]) <= drift_ceilings[name]
            for name in STATISTICS
        }
        for implementation in ("control", "candidate")
    }
    accepted: list[str] = []
    rejected: dict[str, str] = {}
    for name in STATISTICS:
        reasons: list[str] = []
        first_reduction = candidate_reduction["a1_to_b1"][name]
        second_reduction = candidate_reduction["a2_to_b2"][name]
        if first_reduction < 0 and second_reduction < 0:
            adverse_both.append(name)
            reasons.append("candidate is not lower in both paired directions")
        elif first_reduction <= 0 or second_reduction <= 0:
            if (first_reduction < 0 < second_reduction) or (
                second_reduction < 0 < first_reduction
            ):
                reasons.append("paired directions disagree")
            else:
                reasons.append("candidate is not lower in both paired directions")
        for implementation in ("control", "candidate"):
            if not drift_within_ceiling[implementation][name]:
                reasons.append(
                    f"{implementation} drift {drift[implementation][name]:+.6f}% "
                    f"exceeds {drift_ceilings[name]:g}% ceiling"
                )
        if reasons:
            rejected[name] = "; ".join(reasons)
        else:
            accepted.append(name)

    corpus = json.loads(corpus_identity)
    return {
        "case": case,
        "shape": corpus.get("shape"),
        "corpus": corpus,
        "source": source_value,
        "sink": sink_value,
        "output_sha256": output_value,
        "identity": {
            "corpus": corpus_identity,
            "source_present": source_present,
            "source_canonical_json": source_identity,
            "source_status": source_status,
            "sink_present": sink_present,
            "sink_canonical_json": sink_identity,
            "sink_status": sink_status,
            "output_sha256_status": output_status,
            "output_sha256_canonical_json": output_identity,
            "operation_metrics_canonical_json": operation_metrics_identity,
            "operation_metrics_status": operation_metrics_status,
        },
        "elapsed_ns": {
            "sample_count": next(iter(sample_counts)),
            "legs_ns": elapsed,
            "candidate_reduction_percent": candidate_reduction,
            "same_implementation_drift_percent": drift,
            "drift_ceiling_percent": dict(drift_ceilings),
            "drift_within_ceiling": drift_within_ceiling,
            "adverse_both_statistics": adverse_both,
            "accepted_statistics": accepted,
            "rejected_statistics": rejected,
        },
    }


def summarize_reports(
    a1: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    b1: Mapping[str, Any] | None = None,
    b2: Mapping[str, Any] | None = None,
    a2: Mapping[str, Any] | None = None,
    *,
    drift_ceilings: Mapping[str, Any] | None = None,
    cases: Iterable[str] | None = None,
    shapes: Iterable[str] | None = None,
    reports: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    ceilings: Mapping[str, Any] | None = None,
    profile: str | None = None,
) -> dict[str, Any]:
    """Validate four raw reports and return a deterministic summary.

    Projection-only validated state is intentionally unavailable through this
    public entry point; it is consumed by the private streaming boundary.
    """

    return _summarize_reports_impl(
        a1,
        b1,
        b2,
        a2,
        drift_ceilings=drift_ceilings,
        cases=cases,
        shapes=shapes,
        reports=reports,
        ceilings=ceilings,
        profile=profile,
    )


def _summarize_reports_impl(
    a1: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    b1: Mapping[str, Any] | None = None,
    b2: Mapping[str, Any] | None = None,
    a2: Mapping[str, Any] | None = None,
    *,
    drift_ceilings: Mapping[str, Any] | None = None,
    cases: Iterable[str] | None = None,
    shapes: Iterable[str] | None = None,
    reports: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    ceilings: Mapping[str, Any] | None = None,
    profile: str | None = None,
    _validated: tuple[str, Mapping[str, tuple[Any, ...]]] | None = None,
) -> dict[str, Any]:
    """Validate four reports and return a deterministic machine-readable summary.

    ``a1`` may be a four-item sequence, a mapping keyed by ``a1``, ``b1``,
    ``b2`` and ``a2``, or the first report when the remaining three reports are
    passed as positional arguments.
    """

    if _validated is not None and any(
        value is not None for value in (a1, b1, b2, a2, reports)
    ):
        raise AbbaSummaryInputError("_validated cannot be combined with raw reports")
    if _validated is not None:
        report_profile, validated = _validated
        if profile is not None and profile != report_profile:
            raise AbbaSummaryInputError(
                f"validated report profile {report_profile!r} does not match {profile!r}"
            )
        profile = report_profile
    elif reports is not None:
        if a1 is not None or b1 is not None or b2 is not None or a2 is not None:
            raise AbbaSummaryInputError("reports cannot be combined with positional ABBA reports")
        a1 = reports
    if _validated is None and a1 is None:
        raise AbbaSummaryInputError("four ABBA reports are required")
    if drift_ceilings is not None and ceilings is not None:
        raise AbbaSummaryInputError("use drift_ceilings or ceilings, not both")
    ceiling_values = _validate_drift_ceilings(
        drift_ceilings if drift_ceilings is not None else ceilings
    )
    if _validated is None:
        report_values = _coerce_reports(a1, b1, b2, a2)
        detected_profiles = {
            detect_report_profile(report, label)
            for label, report in zip(LEG_ORDER, report_values)
        }
        if len(detected_profiles) != 1:
            raise AbbaSummaryInputError("mixed legacy-v1/current-v1 ABBA report profiles")
        detected_profile = next(iter(detected_profiles))
        if profile is not None and profile != detected_profile:
            raise AbbaSummaryInputError(
                f"raw report profile {detected_profile!r} does not match {profile!r}"
            )
        profile = detected_profile
        validated = {
            label: _validate_report(report, label, profile=profile)
            for label, report in zip(LEG_ORDER, report_values)
        }
    elif profile not in REPORT_PROFILES:
        raise AbbaSummaryInputError(f"unsupported ABBA report profile {profile!r}")
    schema_versions = {item[0] for item in validated.values()}
    if len(schema_versions) != 1:
        raise AbbaSummaryInputError("harness schema_version differs between ABBA legs")
    tool_identities = {
        _canonical_json(item[1], f"{label}.tool") for label, item in validated.items()
    }
    if len(tool_identities) != 1:
        raise AbbaSummaryInputError("harness tool identity differs between ABBA legs")
    environments_by_leg = {label: item[2] for label, item in validated.items()}
    control_revision = environments_by_leg["a1"]["git_revision"]
    if control_revision != environments_by_leg["a2"]["git_revision"]:
        raise AbbaSummaryInputError("control A1/A2 git_revision differs")
    candidate_revision = environments_by_leg["b1"]["git_revision"]
    if candidate_revision != environments_by_leg["b2"]["git_revision"]:
        raise AbbaSummaryInputError("candidate B1/B2 git_revision differs")
    if control_revision == candidate_revision:
        raise AbbaSummaryInputError(
            "control and candidate git_revision must be distinct"
        )
    environment_identities = {
        label: _canonical_json(item[2], f"{label}.environment")
        for label, item in validated.items()
    }
    stable_environment_identities = {
        label: _canonical_json(_stable_environment(item[2]), f"{label}.environment")
        for label, item in validated.items()
    }
    if len(set(stable_environment_identities.values())) != 1:
        raise AbbaSummaryInputError(
            "stable environment identity differs between ABBA legs"
        )
    configurations = {
        _canonical_json(item[3], f"{label}.configuration") for label, item in validated.items()
    }
    if len(configurations) != 1:
        raise AbbaSummaryInputError("harness configuration differs between ABBA legs")

    seen_child_process_ids: set[int] = set()
    for label in LEG_ORDER:
        child_process_ids = validated[label][10]
        duplicate_child_process_ids = seen_child_process_ids.intersection(
            child_process_ids
        )
        if duplicate_child_process_ids:
            raise AbbaSummaryInputError(
                "filesystem child_process_id values must be globally unique across "
                f"ABBA legs; {label} repeats {sorted(duplicate_child_process_ids)}"
            )
        seen_child_process_ids.update(child_process_ids)

    binary_identity_values: dict[str, dict[str, Any] | None] = {
        label: item[9] for label, item in validated.items()
    }
    if profile == REPORT_PROFILE_CURRENT:
        binary_identities = {
            label: _canonical_json(item[9], f"{label}.binary_identity")
            for label, item in validated.items()
        }
        if binary_identities["a1"] != binary_identities["a2"]:
            raise AbbaSummaryInputError("control binary identity differs between A1 and A2")
        if binary_identities["b1"] != binary_identities["b2"]:
            raise AbbaSummaryInputError("candidate binary identity differs between B1 and B2")
        if (
            binary_identity_values["a1"] is None
            or binary_identity_values["b1"] is None
            or binary_identity_values["a1"]["binary_sha256"]
            == binary_identity_values["b1"]["binary_sha256"]
        ):
            raise AbbaSummaryInputError(
                "control and candidate executable SHA-256 hashes must differ"
            )

    for label, (
        _,
        _,
        _,
        configuration,
        indexed,
        _,
        filesystem_shapes,
        _,
        _,
        _,
        _,
    ) in validated.items():
        _validate_configuration_rows(
            configuration,
            indexed,
            label,
            filesystem_shapes=filesystem_shapes,
            projected=_validated is not None,
        )

    filesystem_presence = {item[7] for item in validated.values()}
    if len(filesystem_presence) != 1:
        raise AbbaSummaryInputError(
            "filesystem_evidence presence differs between ABBA legs"
        )
    filesystem_identity_sets = {
        frozenset(item[8]) for item in validated.values()
    }
    if len(filesystem_identity_sets) != 1:
        raise AbbaSummaryInputError(
            "filesystem_evidence case/corpus identities differ between ABBA legs"
        )
    first_filesystem_identity = validated["a1"][8]
    for key in sorted(first_filesystem_identity):
        expected_identity = first_filesystem_identity[key]
        for label in LEG_ORDER:
            if validated[label][8][key] != expected_identity:
                raise AbbaSummaryInputError(
                    "filesystem_evidence identity differs between ABBA legs for "
                    f"{key[0]}[{key[1]}]"
                )
    filesystem_scopes = {
        key: _filesystem_evidence_scope(
            first_filesystem_identity[key],
            f"filesystem_evidence[{key[0]}[{key[1]}]]",
        )
        for key in sorted(first_filesystem_identity)
    }

    result_sets = {frozenset(item[4]) for item in validated.values()}
    if len(result_sets) != 1:
        raise AbbaSummaryInputError("case/corpus result identities differ between ABBA legs")
    selected_cases = _parse_selectors(cases)
    selected_shapes = _parse_selectors(shapes)
    first_index = validated["a1"][4]
    selected_keys = []
    for case, corpus_identity in sorted(first_index):
        if filesystem_scopes.get((case, corpus_identity)) == _XLSX_REPEAT_STORE_STRUCTURAL_SCOPE:
            continue
        corpus = json.loads(corpus_identity)
        if selected_cases is not None and case not in selected_cases:
            continue
        shape = corpus.get("shape")
        if selected_shapes is not None and not isinstance(shape, str):
            raise AbbaSummaryInputError(
                f"{case}[{corpus_identity}] cannot be selected by shape without a string shape"
            )
        if selected_shapes is not None and shape not in selected_shapes:
            continue
        selected_keys.append((case, corpus_identity))
    if not selected_keys:
        raise AbbaSummaryInputError("selectors did not match any case/corpus result")

    # Validate every row before applying selectors.  A selector controls what
    # is emitted; it must not hide an unsafe mismatch in another report row.
    all_summaries: dict[tuple[str, str], dict[str, Any]] = {}
    for case, corpus_identity in sorted(first_index):
        if filesystem_scopes.get((case, corpus_identity)) == _XLSX_REPEAT_STORE_STRUCTURAL_SCOPE:
            continue
        rows = {label: validated[label][4][(case, corpus_identity)] for label in LEG_ORDER}
        all_summaries[(case, corpus_identity)] = _result_summary(
            rows,
            case=case,
            corpus_identity=corpus_identity,
            drift_ceilings=ceiling_values,
            report_schema=next(iter(schema_versions)),
            projected=_validated is not None,
        )

    results = [all_summaries[key] for key in selected_keys]

    tool = json.loads(next(iter(tool_identities)))
    configuration = json.loads(next(iter(configurations)))
    environments = {
        label: json.loads(environment_identities[label]) for label in LEG_ORDER
    }
    stable_environment = json.loads(stable_environment_identities["a1"])
    report_identities = {
        label: {"canonical_sha256": validated[label][5]} for label in LEG_ORDER
    }
    status_counts = {
        field: {
            status: sum(
                result["identity"][f"{field}_status"] == status for result in results
            )
            for status in ("verified_equal", "consistently_absent")
        }
        for field in ("source", "sink")
    }
    output_status_counts = {
        status: sum(
            result["identity"]["output_sha256_status"] == status for result in results
        )
        for status in ("verified_equal", "consistently_absent")
    }
    operation_metrics_status_counts = {
        status: sum(
            result["identity"]["operation_metrics_status"] == status
            for result in results
        )
        for status in ("verified_equal", "consistently_absent")
    }
    native_xls_result_count = sum(
        result["case"] in XLS_NUMERIC_CASE_CONTRACTS for result in results
    )
    if native_xls_result_count and (
        status_counts["source"]["verified_equal"] != native_xls_result_count
    ):
        raise AbbaSummaryInputError(
            "native XLS numeric source_identity_verified must be exactly true"
        )
    implementation_identity: dict[str, Any] = {
        "control": {
            "git_revision": control_revision,
            "legs": ["a1", "a2"],
        },
        "candidate": {
            "git_revision": candidate_revision,
            "legs": ["b1", "b2"],
        },
        "distinct": True,
    }
    # The strict claim/evidence package uses these raw report hashes together
    # with the binary identities below as its external binding.  They are
    # fingerprints of untrusted JSON, not an authenticity mechanism.
    verification: dict[str, Any] = {
        "result_count": len(results),
        "tool_identity_verified": True,
        "configuration_identity_verified": True,
        "environment_stable_identity_verified": True,
        "environment_legs_recorded": True,
        "source_identity": status_counts["source"],
        "sink_identity": status_counts["sink"],
        "output_sha256_identity": output_status_counts,
        "operation_metrics_identity": operation_metrics_status_counts,
        "source_identity_verified": status_counts["source"]["verified_equal"] == len(results),
        "sink_identity_verified": status_counts["sink"]["verified_equal"] == len(results),
        "output_sha256_identity_verified": output_status_counts["verified_equal"]
        == len(results),
        "case_corpus_identity_verified": True,
        "filesystem_evidence_identity_verified": True,
        "statistics_recomputed_from_samples": True,
    }
    if profile == REPORT_PROFILE_CURRENT:
        assert binary_identity_values["a1"] is not None
        assert binary_identity_values["b1"] is not None
        implementation_identity["control"].update(
            {
                "binary_sha256": binary_identity_values["a1"]["binary_sha256"],
                "binary_identity": binary_identity_values["a1"],
            }
        )
        implementation_identity["candidate"].update(
            {
                "binary_sha256": binary_identity_values["b1"]["binary_sha256"],
                "binary_identity": binary_identity_values["b1"],
            }
        )
        verification.update(
            {
                "binary_identity_verified": True,
                "binary_hashes_distinct": True,
            }
        )
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
        "protocol": {
            "order": ["a1_control", "b1_candidate", "b2_candidate", "a2_control"],
            "statistics": list(STATISTICS),
            "drift_ceiling_percent": ceiling_values,
            "percentiles": (
                "p50 = Rust u64 floor midpoint; p95/p99 = integer nearest-rank"
            ),
            "dispersion": "sample standard deviation via Rust Welford update",
            "uncertainty": (
                "95% two-sided Student's t interval for the mean; embedded harness "
                "critical-value table with Cornish-Fisher tail"
            ),
        },
        "harness_identity": {
            "schema_version": next(iter(schema_versions)),
            "tool": tool,
            "configuration": configuration,
        },
        "environment": {
            "stable": stable_environment,
            "legs": environments,
        },
        "implementation_identity": implementation_identity,
        "report_identity": report_identities,
        "results": results,
        "verification": verification,
    }


build_summary = summarize_reports
summarize_abba = summarize_reports


def load_report(path: Path) -> dict[str, Any]:
    def reject_nonfinite(value: str) -> None:
        raise ValueError(f"non-finite JSON value {value!r}")

    def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        value: dict[str, Any] = {}
        for key, item in pairs:
            if key in value:
                raise ValueError(f"duplicate JSON object key {key!r}")
            value[key] = item
        return value

    try:
        with path.open(encoding="utf-8") as handle:
            value = json.load(
                handle,
                object_pairs_hook=reject_duplicate_keys,
                parse_constant=reject_nonfinite,
            )
    except (OSError, ValueError) as error:
        raise AbbaSummaryInputError(f"cannot read {path}: {error}") from error
    return _require_object(value, str(path))


def _parse_ceiling_argument(value: str) -> dict[str, float]:
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError:
        parsed = {}
        for item in value.split(","):
            name, separator, number = item.partition("=")
            if not separator:
                raise AbbaSummaryInputError(
                    "--drift-ceilings must be JSON or comma-separated name=value pairs"
                )
            parsed[name.strip()] = float(number)
    if not isinstance(parsed, dict):
        raise AbbaSummaryInputError("--drift-ceilings must describe an object")
    return _validate_drift_ceilings(parsed)


def _write_json(path: Path, value: dict[str, Any]) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
    except FileExistsError as error:
        raise AbbaSummaryInputError(f"output already exists: {path}") from error
    except OSError as error:
        raise AbbaSummaryInputError(f"cannot write {path}: {error}") from error


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "reports",
        nargs="*",
        type=Path,
        metavar="REPORT",
        help="four reports in A1,B1,B2,A2 order",
    )
    for label in LEG_ORDER:
        aliases = {
            "a1": ("--a1", "--control-a", "--before-a"),
            "b1": ("--b1", "--candidate-a", "--after-a"),
            "b2": ("--b2", "--candidate-b", "--after-b"),
            "a2": ("--a2", "--control-b", "--before-b"),
        }[label]
        parser.add_argument(*aliases, type=Path, help=f"{label.upper()} harness report")
    parser.add_argument("--json-out", "--output", type=Path)
    parser.add_argument(
        "--drift-ceilings",
        help="JSON object or comma-separated p50=5,mean=5,p95=10,p99=15",
    )
    parser.add_argument("--case", dest="cases", action="append")
    parser.add_argument("--shape", dest="shapes", action="append")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    try:
        named = [getattr(arguments, label) for label in LEG_ORDER]
        if arguments.reports and any(path is not None for path in named):
            raise AbbaSummaryInputError("use positional reports or --a1/--b1/--b2/--a2, not both")
        if arguments.reports:
            if len(arguments.reports) != 4:
                raise AbbaSummaryInputError("exactly four positional reports are required")
            paths = arguments.reports
        else:
            if any(path is None for path in named):
                raise AbbaSummaryInputError("all four --a1/--b1/--b2/--a2 reports are required")
            paths = named
        reports = [load_report(path) for path in paths]
        ceilings = (
            _parse_ceiling_argument(arguments.drift_ceilings)
            if arguments.drift_ceilings is not None
            else None
        )
        summary = summarize_reports(
            reports,
            drift_ceilings=ceilings,
            cases=arguments.cases,
            shapes=arguments.shapes,
        )
        if arguments.json_out is not None:
            _write_json(arguments.json_out, summary)
        json.dump(summary, sys.stdout, indent=2, sort_keys=True, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (OSError, AbbaSummaryInputError, ValueError) as error:
        print(f"{TOOL_NAME}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
