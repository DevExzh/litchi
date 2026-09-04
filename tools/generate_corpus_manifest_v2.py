#!/usr/bin/env python3
"""Generate the additive schema-2 corpus catalog from the checked V1 identity.

This tool intentionally performs a conservative migration.  Metadata which is
not represented by the V1 identity object is emitted as ``unknown``/``null``;
the migration never turns missing security, producer, or limit evidence into a
negative assertion.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from typing import Any


MANIFEST_VERSION = 2
MANIFEST_KIND = "corpus-catalog"
CANONICALIZATION = {
    "algorithm": "sorted-json-utf8-compact-v1",
    "hash": "sha256",
}

# This is the source-audited family map for the checked default corpus matrix.
# Keep it in lockstep with the table in CORPUS_MANIFEST_V2.md and the Rust
# migration.  The values are generator-contract metadata, not archive scans.
FAMILY_MAP: dict[str, dict[str, Any]] = {
    "litchi-cfb-synthetic-v1": {
        "family": "cfb",
        "kind": "synthetic",
        "source_kind": "generated",
        "source_path": "tools/perf-baseline/src/lib.rs",
        "producer": "Litchi deterministic generator",
        "license_spdx": "Apache-2.0",
        "license_evidence": "repository-license",
        "redistributable": True,
        "algorithm_id": "litchi-perf.cfb-payload-v1",
        "seed_spec": "indexed-formula-v1",
        "source_functions": [
            "build_cfb_corpus",
            "cfb_entry_name",
            "payload_bytes",
            "CorpusShape",
            "PayloadKind",
        ],
    },
    "litchi-opc-synthetic-v2": {
        "family": "opc",
        "kind": "synthetic",
        "source_kind": "generated",
        "source_path": "tools/perf-baseline/src/lib.rs",
        "producer": "Litchi deterministic generator",
        "license_spdx": "Apache-2.0",
        "license_evidence": "repository-license",
        "redistributable": True,
        "algorithm_id": "litchi-perf.opc-payload-v1",
        "seed_spec": "indexed-formula-v1",
        "source_functions": [
            "build_opc_corpus",
            "entry_name",
            "payload_bytes",
            "CorpusShape",
            "PayloadKind",
        ],
    },
    "litchi-legacy-writer-v1": {
        "family": "legacy-writer",
        "kind": "synthetic",
        "source_kind": "generated",
        "source_path": "tools/perf-baseline/src/lib.rs",
        "producer": "Litchi deterministic generator",
        "license_spdx": "Apache-2.0",
        "license_evidence": "repository-license",
        "redistributable": True,
        "algorithm_id": "litchi-perf.legacy-writer-v1",
        "seed_spec": "none",
        "source_functions": [
            "build_writer_corpus",
            "write_fresh_doc",
            "write_fresh_xls",
            "write_fresh_ppt",
            "WriterShape",
            "writer_text",
            "writer_payload_text",
        ],
    },
    "litchi-xlsx-synthetic-v1": {
        "family": "xlsx",
        "kind": "synthetic",
        "source_kind": "generated",
        "source_path": "tools/perf-baseline/src/lib.rs",
        "producer": "Litchi deterministic generator",
        "license_spdx": "Apache-2.0",
        "license_evidence": "repository-license",
        "redistributable": True,
        "algorithm_id": "litchi-perf.xlsx-integer-grid-v1",
        "seed_spec": "none",
        "source_functions": [
            "build_xlsx_corpus",
            "build_xlsx_workbook",
            "xlsx_one_percent_updates",
            "xlsx_value",
            "xlsx_sheet_name",
            "xlsx_address",
            "XlsxShape",
        ],
    },
}

_COMPRESSIBLE_PAYLOAD_FORMULA = (
    "BLOCK[(offset + index) % 45], "
    "BLOCK=litchi-perf-baseline-compressible-payload-v1\\n"
)
_INCOMPRESSIBLE_PAYLOAD_FORMULA = (
    "state=(index*0x9e3779b97f4a7c15+0xd1b54a32d192ed03) mod 2^64; "
    "xorshift64 shifts=(13,7,17); byte=state>>24"
)


def canonical_bytes(value: Any) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        allow_nan=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def format_slug(value: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-")
    return slug.lower()


def content_id(corpus: dict[str, Any]) -> str:
    return f"{format_slug(corpus['package_format'])}:sha256:{corpus['archive_sha256']}"


def generator_parameters(
    legacy: dict[str, Any], family: dict[str, Any] | None
) -> dict[str, Any]:
    """Return source-audited generator facts without inventing archive data."""

    parameters: dict[str, Any] = {
        "legacy_shape": legacy["shape"],
        "legacy_payload_kind": legacy["payload_kind"],
    }
    if family is None:
        return parameters

    family_id = family["family"]
    parameters.update(
        {
            "family": family_id,
            "package_format": legacy["package_format"],
            "compression": legacy["compression"],
            "entry_count": legacy["entry_count"],
            "entry_bytes": legacy["entry_bytes"],
            "archive_member_count": legacy["archive_member_count"],
            "source_functions": list(family["source_functions"]),
        }
    )

    if family_id in {"cfb", "opc"}:
        is_cfb = family_id == "cfb"
        target_index = (
            legacy["entry_count"] - 1
            if is_cfb
            else legacy["entry_count"] // 2
        )
        parameters.update(
            {
                "target_index": target_index,
                "target_name_pattern": (
                    "benchmark_stream_{index:05}.bin"
                    if is_cfb
                    else "benchmark/parts/{index:05}.bin"
                ),
                "payload_formula": (
                    _COMPRESSIBLE_PAYLOAD_FORMULA
                    if legacy["payload_kind"] == "compressible"
                    else _INCOMPRESSIBLE_PAYLOAD_FORMULA
                ),
            }
        )
        if legacy["payload_kind"] == "compressible":
            parameters["payload_block_bytes"] = 45
        if family_id == "opc":
            parameters["relationship"] = {
                "id": "rIdBenchmarkMain",
                "type": (
                    "http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships/officeDocument"
                ),
                "target": legacy["target_entry"],
                "mode": "Internal",
                "part_content_type": "application/octet-stream",
            }
        return parameters

    if family_id == "legacy-writer":
        writer_format = legacy["package_format"].split("/", 1)[0].lower()
        parameters.update(
            {
                "writer_format": writer_format,
                "writer_shape": legacy["shape"],
                "target_stream": legacy["target_entry"],
                "writer_text_template": (
                    "litchi-perf-baseline-{kind}-v1-"
                    "{first:03}-{second:05}-{third:03} deterministic payload"
                ),
                "payload_heavy_repeat_block": (
                    "litchi-perf-baseline-payload-heavy-v1 "
                ),
                "payload_heavy_bytes": {
                    "doc": 20_000,
                    "xls": 32_700,
                    "ppt": 40_000,
                }[writer_format],
            }
        )
        if writer_format == "doc":
            parameters["paragraph_count"] = legacy["entry_count"]
        elif writer_format == "xls":
            if legacy["shape"] == "tiny":
                dimensions = {"sheets": 1, "rows": 4, "columns": 4}
            elif legacy["shape"] == "large":
                dimensions = {"sheets": 4, "rows": 128, "columns": 16}
            else:
                dimensions = {"sheets": 128, "string_bytes": 32_700}
            parameters["dimensions"] = dimensions
            parameters["numeric_value_formula"] = (
                "(sheet * rows * columns + row * columns + column) as f64"
            )
        else:
            if legacy["shape"] == "tiny":
                dimensions = {"slides": 1, "text_boxes_per_slide": 2}
            elif legacy["shape"] == "large":
                dimensions = {"slides": 12, "text_boxes_per_slide": 12}
            else:
                dimensions = {"slides": 16, "text_boxes_per_slide": 8}
            parameters["dimensions"] = dimensions
            parameters["textbox_position_formula"] = (
                "x=36+(box_number % 3)*180; "
                "y=36+(box_number / 3)*90; width=144; height=54"
            )
        return parameters

    if family_id == "xlsx":
        xlsx = legacy["xlsx"]
        if xlsx is not None:
            parameters.update(
                {
                    "sheet_count": xlsx["sheet_count"],
                    "rows_per_sheet": xlsx["rows_per_sheet"],
                    "columns_per_sheet": xlsx["columns_per_sheet"],
                    "one_percent_update_count": xlsx[
                        "one_percent_update_count"
                    ],
                }
            )
        parameters.update(
            {
                "value_formula": "(sheet * 1_000_000 + row * 1_000 + column) as i32",
                "sheet_name_formula": "index == 0 ? Sheet1 : Bench{index:02}",
                "address_formula": "column_to_letters(column + 1) + (row + 1)",
                "one_percent_update_formula": "ceil(cell_count / 100)",
            }
        )
        return parameters

    return parameters


def migrate_corpus(legacy: dict[str, Any]) -> dict[str, Any]:
    generator = legacy["generator"]
    shape = legacy["shape"]
    payload_kind = legacy["payload_kind"]
    family = FAMILY_MAP.get(generator)
    generated = family is not None or "synthetic" in generator
    categories = ["legacy-migrated"]
    if generated:
        categories.append("synthetic")
    if payload_kind == "compressible":
        categories.append("highly-compressible")
    elif payload_kind == "incompressible":
        categories.append("incompressible")
    if shape == "many-small":
        categories.append("many-small-parts")
    elif shape == "few-large":
        categories.append("few-large-parts")
    categories = sorted(set(categories))
    revision_match = re.search(r"-v([^\-]+)$", generator)
    revision = f"v{revision_match.group(1)}" if revision_match else None
    shape_parameters: dict[str, Any] = {
        "entry_count": legacy["entry_count"],
        "entry_bytes": legacy["entry_bytes"],
    }
    xlsx = legacy.get("xlsx")
    if xlsx:
        for key in ("sheet_count", "rows_per_sheet", "columns_per_sheet"):
            shape_parameters[key] = xlsx[key]
    generator_metadata = {
        "id": generator,
        "kind": family["kind"] if family else ("synthetic" if generated else "unknown"),
        "revision": revision,
        "algorithm_id": family["algorithm_id"] if family else None,
        "seed_spec": family["seed_spec"] if family else None,
        "parameters": generator_parameters(legacy, family),
    }
    provenance_metadata = {
        "source_kind": family["source_kind"] if family else ("generated" if generated else "unknown"),
        "source_path": family["source_path"] if family else None,
        "producer": family["producer"] if family else ("Litchi deterministic generator" if generated else None),
        "producer_version": None,
        # This field identifies fixture/input bytes, never generator source.
        "source_sha256": None,
        "license_spdx": family["license_spdx"] if family else ("Apache-2.0" if generated else None),
        "license_evidence": family["license_evidence"] if family else ("repository-license" if generated else None),
        "redistributable": family["redistributable"] if family else (True if generated else None),
    }
    return {
        "id": content_id(legacy),
        "name": legacy["name"],
        "legacy_v1": legacy,
        "format": legacy["package_format"],
        "size_class": shape if shape in {"tiny", "medium", "large"} else "unknown",
        "categories": categories,
        "generator": generator_metadata,
        "provenance": provenance_metadata,
        "bytes": {
            "archive_bytes": legacy["archive_bytes"],
            "archive_sha256": legacy["archive_sha256"],
            "logical_payload_bytes": legacy["uncompressed_payload_bytes"],
            "text_bytes": None,
            "media_bytes": None,
            "metadata_bytes": None,
        },
        "shape_parameters": shape_parameters,
        "relationships": {
            "status": "unknown",
            "relationship_count": None,
            "dependency_closure_nodes": None,
            "dependency_closure_edges": None,
            "max_depth": None,
            "max_out_degree": None,
        },
        "security": {
            "encryption": {
                "state": "unknown",
                "kind": None,
                "evidence": "not-recorded",
            },
            "signature": {
                "state": "unknown",
                "kind": None,
                "evidence": "not-recorded",
            },
            "protection": {
                "state": "unknown",
                "scopes": [],
                "evidence": "not-recorded",
            },
            "macros": {
                "state": "unknown",
                "kind": None,
                "evidence": "not-recorded",
            },
            "external_links": {
                "state": "unknown",
                "count": None,
                "targets_sha256": None,
                "evidence": "not-recorded",
            },
        },
        "input": {
            "validity": "unknown",
            "malformation_kind": None,
            "expected_behavior": "unknown",
            "within_limits": None,
            "evidence": "not-recorded",
        },
        "limits": {
            "profile_id": None,
            "profile_sha256": None,
            "observed": {
                "input_bytes": legacy["archive_bytes"],
                "members": legacy["archive_member_count"],
                "relationships": None,
                "materialized_bytes": None,
            },
        },
        "members": {"status": "unavailable", "items": []},
        "targets": [
            {
                "entry": legacy["target_entry"],
                "logical_bytes": legacy["target_payload_bytes"],
                "sha256": legacy["target_payload_sha256"],
            }
        ],
        "coverage": {
            "timed_cases": [],
            "guard_cases": [],
            "inventory_only": False,
        },
    }


def generate(v1: dict[str, Any], revision: str | None) -> dict[str, Any]:
    corpora = {
        content_id(corpus): migrate_corpus(corpus)
        for corpus in v1["corpora"].values()
    }
    bindings = []
    for case, names in v1["case_corpora"].items():
        for name in names:
            legacy = v1["corpora"][name]
            corpus_id = content_id(legacy)
            corpora[corpus_id]["coverage"]["timed_cases"].append(case)
            bindings.append(
                {
                    "case": case,
                    "corpus_id": corpus_id,
                    "legacy_name": name,
                    "legacy_archive_sha256": legacy["archive_sha256"],
                    "role": "timed",
                }
            )
    for corpus in corpora.values():
        corpus["coverage"]["timed_cases"] = sorted(
            set(corpus["coverage"]["timed_cases"])
        )
    bindings.sort(key=lambda binding: (binding["case"], binding["corpus_id"]))
    catalog = {
        "manifest_version": MANIFEST_VERSION,
        "manifest_kind": MANIFEST_KIND,
        "catalog_id": "litchi-perf-corpus-v2",
        "canonicalization": CANONICALIZATION,
        "catalog_sha256": "",
        "content_set_sha256": "",
        "build": {
            "tool": "litchi-perf-baseline",
            "tool_version": "0.1.0",
            "git_revision": revision,
            "git_worktree_dirty": False if revision else None,
            "source_files": [],
        },
        "corpora": [corpora[key] for key in sorted(corpora)],
        "case_bindings": bindings,
    }
    content_set = {
        "corpora": [
            {
                "id": corpus["id"],
                "archive_sha256": corpus["bytes"]["archive_sha256"],
                "members": [
                    {
                        "ordinal": member["ordinal"],
                        "name": member["name"],
                        "sha256": member["sha256"],
                    }
                    for member in corpus["members"]["items"]
                ],
            }
            for corpus in catalog["corpora"]
        ],
        "case_bindings": [
            {
                "case": binding["case"],
                "corpus_id": binding["corpus_id"],
                "role": binding["role"],
            }
            for binding in bindings
        ],
    }
    catalog["content_set_sha256"] = sha256(canonical_bytes(content_set))
    without_hash = dict(catalog)
    without_hash.pop("catalog_sha256")
    catalog["catalog_sha256"] = sha256(canonical_bytes(without_hash))
    return catalog


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--v1", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--revision")
    args = parser.parse_args()
    v1 = json.loads(args.v1.read_text(encoding="utf-8"))
    catalog = generate(v1, args.revision)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(catalog, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
