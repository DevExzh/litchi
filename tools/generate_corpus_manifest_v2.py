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


def migrate_corpus(legacy: dict[str, Any]) -> dict[str, Any]:
    generator = legacy["generator"]
    shape = legacy["shape"]
    payload_kind = legacy["payload_kind"]
    generated = "synthetic" in generator
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
    return {
        "id": content_id(legacy),
        "name": legacy["name"],
        "legacy_v1": legacy,
        "format": legacy["package_format"],
        "size_class": shape if shape in {"tiny", "medium", "large"} else "unknown",
        "categories": categories,
        "generator": {
            "id": generator,
            "kind": "synthetic" if generated else "unknown",
            "revision": revision,
            "algorithm_id": None,
            "seed_spec": None,
            "parameters": {
                "legacy_shape": shape,
                "legacy_payload_kind": payload_kind,
            },
        },
        "provenance": {
            "source_kind": "generated" if generated else "unknown",
            "source_path": None,
            "producer": "Litchi deterministic generator" if generated else None,
            "producer_version": None,
            "source_sha256": None,
            "license_spdx": "Apache-2.0" if generated else None,
            "license_evidence": "repository-license" if generated else None,
            "redistributable": True if generated else None,
        },
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
