#!/usr/bin/env python3
"""Verify bounded smoke receipts without claiming sealed measurements."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import subprocess
from pathlib import Path


LANES = {
    "neutral_open_tiny",
    "host_open_tiny",
    "neutral_open_relationship",
    "host_stage_noop_tiny",
    "host_stage_rename_relationship",
    "host_commit_rename_relationship",
    "host_save_reopen_relationship",
    "host_inverse_relationship",
    "host_exact_cap_relationship",
    "host_refusal_opaque",
    "host_refusal_limit",
}
REFUSALS = {"host_refusal_opaque", "host_refusal_limit"}
NEUTRAL = {"neutral_open_tiny", "neutral_open_relationship"}
EXACT_PRESERVATION = {
    "host_open_tiny",
    "host_stage_noop_tiny",
    "host_inverse_relationship",
    "host_refusal_opaque",
    "host_refusal_limit",
}
OPAQUE_ERROR = (
    "Invalid format: XLDM outer identity proof failed: "
    "MS-XLDM storage must contain at least three complete 4096-byte pages"
)
LIMIT_ERROR = "Invalid format: Data Model rewritten workbook bytes limit exceeded"
SOURCE_MANIFEST_FORMAT = "xlsb-model-identity-cargo-source-closure-v3"


def fail(message: str) -> None:
    raise SystemExit(f"smoke verification failed: {message}")


def load(path: Path) -> dict:
    try:
        value = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        fail(f"{path}: {error}")
    if not isinstance(value, dict):
        fail(f"{path}: receipt is not an object")
    return value


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve(root: Path, shown: str) -> Path:
    path = Path(shown)
    return path if path.is_absolute() else root / path


def committed_blob_sha256(root: Path, commit: str, shown: str) -> str:
    try:
        committed = subprocess.check_output(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "cat-file",
                "blob",
                f"{commit}:{shown}",
            ]
        )
    except subprocess.CalledProcessError as error:
        raise AssertionError(f"cannot read committed source input blob: {shown}") from error
    return hashlib.sha256(committed).hexdigest()


def verify_git_snapshot(paths: list[Path], root: Path, commit: str) -> None:
    """Verify source bytes against the pinned commit, independent of Git status."""

    root = root.resolve()
    relative: list[str] = []
    for path in sorted({path.resolve() for path in paths}, key=str):
        if not path.is_file():
            fail(f"retained Git input missing: {path}")
        try:
            relative.append(path.relative_to(root).as_posix())
        except ValueError as error:
            raise AssertionError(
                f"non-Git local input has no retained source snapshot: {path}"
            ) from error
    if not relative:
        return
    current = subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), "rev-parse", "HEAD"],
        text=True,
    ).strip()
    if current != commit:
        raise AssertionError(f"Git snapshot commit changed: {commit} -> {current}")
    missing: list[str] = []
    for shown in relative:
        tracked = subprocess.run(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "ls-files",
                "--error-unmatch",
                "--",
                shown,
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        if tracked.returncode != 0:
            missing.append(shown)
    if missing:
        raise AssertionError(
            "local Git inputs are not tracked: " + ", ".join(sorted(missing))
        )
    changed = [
        shown
        for shown in relative
        if sha256(root / shown) != committed_blob_sha256(root, commit, shown)
    ]
    if changed:
        raise AssertionError(
            "local Git inputs differ from committed snapshot: "
            + ", ".join(sorted(set(changed)))
        )


def verify_source_manifest(
    path: Path, root: Path, *, require_transitive: bool = True
) -> tuple[int, str, str]:
    lines = path.read_text().splitlines()
    if not lines or lines[0] != f"format={SOURCE_MANIFEST_FORMAT}":
        fail("source manifest format changed")
    git_commit = None
    git_head = None
    metadata_hash = None
    packages = {}
    files = {}
    extras = []
    for line in lines[1:]:
        parts = line.split("\t")
        if line.startswith("git_commit="):
            if git_commit is not None:
                fail("source manifest Git commit is duplicated")
            git_commit = line.split("=", 1)[1]
            if len(git_commit) != 40 or any(
                character not in "0123456789abcdef" for character in git_commit
            ):
                fail("source manifest Git commit is malformed")
        elif line.startswith("git_head="):
            if git_head is not None:
                fail("source manifest Git head is duplicated")
            git_head = line.split("=", 1)[1]
            if len(git_head) != 40 or any(
                character not in "0123456789abcdef" for character in git_head
            ):
                fail("source manifest Git head is malformed")
        elif line.startswith("metadata_sha256="):
            if metadata_hash is not None:
                fail("source manifest metadata hash is duplicated")
            metadata_hash = line.split("=", 1)[1]
            if len(metadata_hash) != 64:
                fail("source manifest metadata hash is malformed")
        elif line.startswith("package="):
            if len(parts) != 7:
                fail(f"malformed source package line: {line}")
            key = (parts[0][len("package=") :], parts[1], parts[3])
            if key in packages:
                fail(f"duplicate source package line: {key}")
            try:
                count = int(parts[5])
            except ValueError:
                fail(f"source package file count is malformed: {line}")
            if count <= 0 or len(parts[4]) != 64 or len(parts[6]) != 64:
                fail(f"source package digest/count is malformed: {line}")
            packages[key] = (parts[2], count, parts[6])
        elif line.startswith("file="):
            if len(parts) != 5:
                fail(f"malformed source file line: {line}")
            key = (parts[0][len("file=") :], parts[1], parts[2])
            files.setdefault(key, []).append((parts[3], parts[4]))
        elif line.startswith("extra=\t"):
            if len(parts) != 3 or len(parts[2]) != 64:
                fail(f"malformed source extra line: {line}")
            extras.append((parts[1], parts[2]))
        else:
            fail(f"unknown source manifest line: {line}")
    if git_commit is None or git_head is None or metadata_hash is None:
        fail("source manifest provenance is incomplete")
    if git_head != git_commit:
        fail("source manifest Git head differs from pinned commit")
    if not packages or set(packages) != set(files):
        fail("source package/file manifest sets differ")
    if require_transitive and len(packages) < 100:
        fail(f"source manifest is not transitively complete: only {len(packages)} packages")
    checked = 0
    local_inputs = []
    for key, (source, expected_count, expected_tree) in packages.items():
        entries = files[key]
        if len(entries) != expected_count:
            fail(f"source package file count changed: {key}")
        tree_payload = "\n".join(f"{shown}\t{digest}" for shown, digest in entries)
        if hashlib.sha256(tree_payload.encode()).hexdigest() != expected_tree:
            fail(f"source package tree changed: {key}")
        for shown, expected in entries:
            current = resolve(root, shown)
            if not current.is_file() or sha256(current) != expected:
                fail(f"source manifest input changed or disappeared: {shown}")
            if source == "path":
                local_inputs.append(current)
            checked += 1
    for shown, expected in extras:
        current = resolve(root, shown)
        if not current.is_file() or sha256(current) != expected:
            fail(f"source manifest extra changed or disappeared: {shown}")
        local_inputs.append(current)
        checked += 1
    if require_transitive and checked < 1000:
        fail(f"source manifest checked too few files: {checked}")
    try:
        verify_git_snapshot(local_inputs, root, git_commit)
    except (AssertionError, OSError) as error:
        fail(str(error))
    return checked, git_commit, metadata_hash


def verify_allocator(sample: dict, lane: str) -> None:
    required = (
        "live_before",
        "live_after",
        "direct_allocated_bytes",
        "realloc_old_bytes",
        "realloc_new_bytes",
        "deallocated_bytes",
        "requested_alloc_bytes",
        "peak_live_delta",
        "alloc_balance_ok",
        "alloc_invalid",
        "allocation_failed",
    )
    for key in required:
        if key not in sample:
            fail(f"{lane}: missing allocator field {key}")
    expected_requested = sample["direct_allocated_bytes"] + sample["realloc_new_bytes"]
    if sample["requested_alloc_bytes"] != expected_requested:
        fail(f"{lane}: requested allocator bytes equation failed")
    expected_live = (
        sample["live_before"]
        + sample["direct_allocated_bytes"]
        + sample["realloc_new_bytes"]
        - sample["realloc_old_bytes"]
        - sample["deallocated_bytes"]
    )
    if sample["live_after"] != expected_live:
        fail(f"{lane}: live allocator equation failed")
    if not sample["alloc_balance_ok"] or sample["alloc_invalid"]:
        fail(f"{lane}: allocator balance/validity failed")
    if sample["allocation_failed"] != 0:
        fail(f"{lane}: allocator reported a failed allocation")


def verify_time_file(path: Path, lane: str) -> None:
    if not path.exists():
        fail(f"{lane}: missing time sidecar")
    text = path.read_text()
    if "Maximum resident set size (kbytes):" not in text:
        fail(f"{lane}: time sidecar has no RSS field")
    if "Exit status: 0" not in text:
        fail(f"{lane}: process exit status was not zero")


def verify_matrix_runner_capture(results: Path, manifest: Path) -> None:
    """Require the runner's exact matrix invocation artifacts before review."""

    argv_capture = results / "matrix-correctness.argv.json"
    stdout = results / "matrix-correctness.stdout.json"
    stderr = results / "matrix-correctness.stderr.log"
    exit_status = results / "matrix-correctness.exit.txt"
    binary_receipt = results / "binary.sha256"
    for path in (argv_capture, stdout, stderr, exit_status, binary_receipt):
        if not path.is_file():
            fail(f"correctness matrix capture is missing: {path.name}")
    binary_lines = binary_receipt.read_text().splitlines()
    if len(binary_lines) != 1 or len(binary_lines[0]) < 66:
        fail("binary hash capture is malformed")
    binary_line = binary_lines[0]
    binary_digest = binary_line[:64]
    if (
        len(binary_digest) != 64
        or any(character not in "0123456789abcdef" for character in binary_digest)
        or binary_line[64:66] != "  "
    ):
        fail("binary hash capture is malformed")
    hashed_binary = binary_line[66:]
    if not hashed_binary:
        fail("binary hash capture has no binary path")
    try:
        captured = json.loads(argv_capture.read_text())
    except (OSError, json.JSONDecodeError) as error:
        fail(f"matrix argv capture is malformed: {error}")
    if not isinstance(captured, dict) or set(captured) != {"argv"}:
        fail("matrix argv capture is malformed")
    argv = captured.get("argv")
    if argv != [hashed_binary, "--matrix-correctness"]:
        fail("matrix argv does not match the hashed binary invocation")
    if exit_status.read_text().strip() != "0":
        fail("correctness matrix command did not exit successfully")
    if stderr.read_bytes() != b"":
        fail("correctness matrix stderr was not empty")
    load(stdout)
    verify_correctness_receipt(stdout, manifest)


MATRIX_CHECK_FIELDS = (
    "table_ids_equal",
    "table_names_equal",
    "relationship_ids_equal",
    "relationship_endpoints_equal",
    "relationship_paths_equal",
    "time_grouping_ids_equal",
    "all_equal",
)
MATRIX_LIMIT_FIELDS = {
    "raw.payload": "raw_payload_bytes",
    "raw.string_units": "raw_string_units",
    "max_tables": "max_tables",
    "max_relationships": "max_relationships",
    "max_time_groupings": "max_time_groupings",
    "max_time_grouping_columns": "max_time_grouping_columns",
    "max_rewrite_bytes": "max_rewrite_bytes",
    "max_records": "max_records",
    "max_part_bytes": "max_part_bytes",
    "max_graph_parts": "max_graph_parts",
    "max_graph_relationships": "max_graph_relationships",
    "max_metadata_bytes": "max_metadata_bytes",
    "max_connection_bytes": "max_connection_bytes",
    "max_connections": "max_connections",
}

MATRIX_SOURCE_PROOF_COMPLETE = "complete_xldm140"
MATRIX_SOURCE_PROOF_UNAVAILABLE_GRAPH_LIMIT = "unavailable_graph_work_limit"
MATRIX_NAME_TARGETS = {
    "same_length_ascii": "TableX",
    "shorter": "T",
    "longer": "Table1-renamed-with-more-bytes",
    "escaped_xml": "A&B<case>",
    "unicode": "表一",
}
EXPECTED_OLAP_PROOF_LIMITS = {
    "max_items": 1_000_000,
    "max_string_bytes": 256 * 1024 * 1024,
    "max_source_bytes": 512 * 1024 * 1024,
    "max_work": 4_000_000,
}
MATRIX_MEMBER_MANIFEST_FIELDS = {
    "parts",
    "relationships",
    "content_types",
    "inner_members",
}


def verify_matrix_limits(receipt: dict, recipe: dict, manifest: dict) -> None:
    limits = receipt.get("caller_limits")
    recipe_limits = recipe.get("caller_limits")
    if not isinstance(limits, dict) or limits != recipe_limits:
        fail("correctness receipt caller limits are missing or disagree with recipe")
    expected = {
        MATRIX_LIMIT_FIELDS[key]
        for key in manifest.get("caller_limits", [])
        if key in MATRIX_LIMIT_FIELDS
    }
    if set(limits) != expected:
        fail("correctness receipt caller-limit field set is incomplete or changed")
    if any(not isinstance(value, int) or value <= 0 for value in limits.values()):
        fail("correctness receipt caller limits are not positive integers")
    if recipe.get("limits_mode") != "Limits::DEFAULT":
        fail("correctness receipt does not identify the applied caller-limit mode")


def verify_olap_proof_limits(receipt: dict, recipe: dict, manifest: dict) -> None:
    limits = receipt.get("olap_proof_limits")
    recipe_limits = recipe.get("olap_proof_limits")
    manifest_limits = manifest.get("olap_proof_limits")
    recipe_manifest_limits = manifest.get("recipe", {}).get("olap_proof_limits")
    if not isinstance(limits, dict) or limits != recipe_limits:
        fail("receipt OLAP proof limits are missing or disagree with recipe")
    if limits != manifest_limits or limits != recipe_manifest_limits:
        fail("receipt OLAP proof limits are not bound to the manifest")
    if limits != EXPECTED_OLAP_PROOF_LIMITS:
        fail("receipt OLAP proof defaults differ from reviewed defaults")


def verify_matrix_semantic_vector(
    vector: object, tables: int, relationships: int, label: str
) -> None:
    if not isinstance(vector, dict):
        fail(f"{label}: semantic vector is missing")
    if set(vector) != {"tables", "relationships", "time_groupings"}:
        fail(f"{label}: semantic vector fields are incomplete")
    vector_tables = vector.get("tables")
    vector_relationships = vector.get("relationships")
    groupings = vector.get("time_groupings")
    if not isinstance(vector_tables, list) or len(vector_tables) != tables:
        fail(f"{label}: table identity vector length is wrong")
    if not isinstance(vector_relationships, list) or len(vector_relationships) != relationships:
        fail(f"{label}: relationship identity vector length is wrong")
    if not isinstance(groupings, list):
        fail(f"{label}: time-grouping vector is missing")
    for index, table in enumerate(vector_tables):
        if not isinstance(table, dict) or set(table) != {
            "table_id",
            "xml_name",
            "metadata_path",
            "dimension_object_id",
        } or any(
            not isinstance(table.get(field), str) or not table[field]
            for field in ("table_id", "xml_name", "metadata_path", "dimension_object_id")
        ):
            fail(f"{label}: table identity {index} is incomplete")
    for index, relationship in enumerate(vector_relationships):
        if not isinstance(relationship, dict) or set(relationship) != {
            "relationship_id",
            "metadata_path",
            "containing_table",
            "primary_table",
            "primary_column",
            "foreign_column",
            "expected_index_key",
        } or any(
            not isinstance(relationship.get(field), str) or not relationship[field]
            for field in (
                "relationship_id",
                "metadata_path",
                "containing_table",
                "primary_table",
                "primary_column",
                "foreign_column",
                "expected_index_key",
            )
        ):
            fail(f"{label}: relationship identity {index} is incomplete")
    for index, grouping in enumerate(groupings):
        if not isinstance(grouping, dict) or set(grouping) != {
            "table_name",
            "column_id",
            "column_ids",
        } or any(
            not isinstance(grouping.get(field), str) or not grouping[field]
            for field in ("table_name", "column_id")
        ):
            fail(f"{label}: time-grouping identity {index} is incomplete")
        if not isinstance(grouping.get("column_ids"), list) or any(
            not isinstance(column, str) or not column for column in grouping["column_ids"]
        ):
            fail(f"{label}: time-grouping columns {index} are incomplete")


def expected_matrix_relationships(
    tables: int, relationships: int, endpoint_layout: str
) -> list[dict[str, str]]:
    if endpoint_layout not in {"selected_table", "distributed"}:
        fail(f"source: unknown endpoint layout {endpoint_layout!r}")
    pairs = [
        (containing, primary)
        for containing in range(tables)
        for primary in range(tables)
        if primary != containing
    ]
    if endpoint_layout == "selected_table":
        pairs.sort(key=lambda pair: (pair[0] != 0, pair[0], pair[1]))
    else:
        pairs.sort(
            key=lambda pair: (
                (pair[1] + tables - pair[0]) % tables,
                pair[0],
                pair[1],
            )
        )
    if relationships > len(pairs):
        fail("source: relationship count exceeds deterministic pair capacity")
    expected = []
    for index, (containing, primary) in enumerate(pairs[:relationships], start=1):
        containing_id = f"T{containing + 1}"
        relationship_id = f"Rel{index}"
        expected.append(
            {
                "relationship_id": relationship_id,
                "metadata_path": (
                    f"Model.1.db/{containing_id}.0.dim/"
                    f"R${containing_id}${relationship_id}.1.tbl.xml"
                ),
                "containing_table": containing_id,
                "primary_table": f"Table{primary + 1}",
                "primary_column": "Key",
                "foreign_column": "Key",
                "expected_index_key": f"R${containing_id}${relationship_id}",
            }
        )
    # The XLDM closure projection groups relationship members by containing
    # dimension before exposing them.  Keep the generated Rel identity from
    # the endpoint recipe, but verify the same deterministic projection order
    # that the public semantic reader returns.
    return sorted(
        expected,
        key=lambda relationship: (
            int(relationship["containing_table"][1:]),
            int(relationship["relationship_id"][3:]),
        ),
    )


def verify_matrix_source_vector(
    vector: dict, tables: int, relationships: int, endpoint_layout: str
) -> None:
    verify_matrix_semantic_vector(vector, tables, relationships, "source")
    table_rows = vector["tables"]
    table_ids = [table["table_id"] for table in table_rows]
    if table_ids != [f"T{index}" for index in range(1, tables + 1)]:
        fail("source: table identity order or IDs are not the deterministic fixture")
    if [table["xml_name"] for table in table_rows] != [
        f"Table{index}" for index in range(1, tables + 1)
    ]:
        fail("source: table XML names are not the deterministic fixture")
    relationship_rows = vector["relationships"]
    expected_relationships = expected_matrix_relationships(
        tables, relationships, endpoint_layout
    )
    if relationship_rows != expected_relationships:
        fail("source: relationship topology is not the deterministic endpoint recipe")
    table_names = {table["xml_name"] for table in table_rows}
    table_id_set = set(table_ids)
    for index, table in enumerate(table_rows, start=1):
        expected_dimension_id = (
            f"11111111-2222-3333-4444-{0x5555_5555_5500 + index:012X}"
        )
        expected_metadata_path = f"Model.1.db/T{index}.0.dim/T{index}.1.tbl.xml"
        if table["dimension_object_id"] != expected_dimension_id:
            fail("source: dimension object ID is not the deterministic recipe")
        if table["metadata_path"] != expected_metadata_path:
            fail("source: table metadata path is not the deterministic recipe")
    for relationship in relationship_rows:
        if relationship["containing_table"] not in table_names | table_id_set:
            fail("source: relationship containing endpoint is outside the table set")
        if relationship["primary_table"] not in table_names | table_id_set:
            fail("source: relationship primary endpoint is outside the table set")
        if relationship["primary_column"] != "Key" or relationship["foreign_column"] != "Key":
            fail("source: relationship columns are not the deterministic Key closure")
        index_match = re.fullmatch(
            r"R\$(T\d+)\$(Rel\d+)", relationship["expected_index_key"]
        )
        if index_match is None:
            fail("source: relationship index key is not deterministic")
        containing_id, relationship_id = index_match.groups()
        if relationship["relationship_id"] != relationship_id:
            fail("source: relationship index key does not match relationship identity")
        if relationship["containing_table"] != containing_id:
            fail("source: relationship containing endpoint does not match its index key")
        containing_index = int(containing_id[1:])
        primary_match = re.fullmatch(r"Table(\d+)", relationship["primary_table"])
        if primary_match is None or int(primary_match.group(1)) == containing_index:
            fail("source: relationship primary endpoint is not a valid distinct table")
        expected_path = (
            f"Model.1.db/{containing_id}.0.dim/"
            f"R${containing_id}${relationship_id}.1.tbl.xml"
        )
        if relationship["metadata_path"] != expected_path:
            fail("source: relationship metadata path is not the deterministic recipe")
    expected_groupings = tables if tables <= 16 else 0
    if len(vector["time_groupings"]) != expected_groupings:
        fail("source: time-grouping shape is outside the fixture contract")
    expected_grouping_rows = [
        {"table_name": f"Table{index}", "column_id": "Key", "column_ids": ["Year"]}
        for index in range(1, expected_groupings + 1)
    ]
    if vector["time_groupings"] != expected_grouping_rows:
        fail("source: time-grouping identity is not the deterministic recipe")


def expected_matrix_mutable_inner_paths(source: dict) -> list[str]:
    paths = {"BackupLog"}
    selected = [table for table in source["tables"] if table["table_id"] == "T1"]
    if len(selected) != 1:
        fail("source: cannot derive mutable paths without exactly one T1 table")
    paths.add(selected[0]["metadata_path"])
    for relationship in source["relationships"]:
        if relationship["containing_table"] in {"T1", "Table1"} or relationship[
            "primary_table"
        ] in {"T1", "Table1"}:
            paths.add(relationship["metadata_path"])
    return sorted(paths)


def expected_matrix_semantic(source: dict, name_profile: str) -> dict:
    target = MATRIX_NAME_TARGETS.get(name_profile)
    if target is None:
        fail(f"unknown matrix name profile: {name_profile}")
    expected = copy.deepcopy(source)
    selected = [table for table in expected["tables"] if table["table_id"] == "T1"]
    if len(selected) != 1 or selected[0]["xml_name"] != "Table1":
        fail("source: selected T1 table is not the expected Table1 identity")
    selected[0]["xml_name"] = target
    for relationship in expected["relationships"]:
        if relationship["containing_table"] == "Table1":
            relationship["containing_table"] = target
        if relationship["primary_table"] == "Table1":
            relationship["primary_table"] = target
    for grouping in expected["time_groupings"]:
        if grouping["table_name"] == "Table1":
            grouping["table_name"] = target
    return expected


def recompute_matrix_check(actual: dict, expected: dict) -> dict:
    table_ids_equal = [table["table_id"] for table in actual["tables"]] == [
        table["table_id"] for table in expected["tables"]
    ]
    table_names_equal = [
        (table["table_id"], table["xml_name"]) for table in actual["tables"]
    ] == [
        (table["table_id"], table["xml_name"]) for table in expected["tables"]
    ]
    relationship_ids_equal = [
        relationship["relationship_id"] for relationship in actual["relationships"]
    ] == [
        relationship["relationship_id"] for relationship in expected["relationships"]
    ]
    endpoint_fields = (
        "containing_table",
        "primary_table",
        "primary_column",
        "foreign_column",
        "expected_index_key",
    )
    relationship_endpoints_equal = [
        tuple(relationship[field] for field in endpoint_fields)
        for relationship in actual["relationships"]
    ] == [
        tuple(relationship[field] for field in endpoint_fields)
        for relationship in expected["relationships"]
    ]
    relationship_paths_equal = [
        relationship["metadata_path"] for relationship in actual["relationships"]
    ] == [
        relationship["metadata_path"] for relationship in expected["relationships"]
    ]
    time_grouping_ids_equal = actual["time_groupings"] == expected["time_groupings"]
    return {
        "table_ids_equal": table_ids_equal,
        "table_names_equal": table_names_equal,
        "relationship_ids_equal": relationship_ids_equal,
        "relationship_endpoints_equal": relationship_endpoints_equal,
        "relationship_paths_equal": relationship_paths_equal,
        "time_grouping_ids_equal": time_grouping_ids_equal,
        "all_equal": actual == expected,
    }


def verify_matrix_semantic_transition(result: dict, tables: int, relationships: int) -> None:
    source = result.get("source_semantic")
    candidate = result.get("semantic_observed")
    reopened = result.get("reopened_semantic_observed")
    if not isinstance(source, dict) or not isinstance(candidate, dict) or not isinstance(reopened, dict):
        fail("successful correctness result does not contain complete semantic vectors")
    verify_matrix_source_vector(
        source, tables, relationships, result["endpoint_layout"]
    )
    verify_matrix_semantic_vector(candidate, tables, relationships, "candidate")
    verify_matrix_semantic_vector(reopened, tables, relationships, "reopened")
    expected = expected_matrix_semantic(source, result["name_profile"])
    if candidate != expected:
        fail("candidate semantic vector differs from independently recomputed identity")
    if reopened != expected:
        fail("reopened semantic vector differs from independently recomputed identity")
    candidate_check = recompute_matrix_check(candidate, expected)
    reopened_check = recompute_matrix_check(reopened, expected)
    if result.get("semantic") != candidate_check:
        fail("candidate semantic equality booleans do not match recomputed vectors")
    if result.get("reopened_semantic") != reopened_check:
        fail("reopened semantic equality booleans do not match recomputed vectors")
    if not all(candidate_check.values()) or not all(reopened_check.values()):
        fail("recomputed semantic identity gate failed")


def verify_matrix_check(check: object, label: str) -> None:
    if not isinstance(check, dict) or set(check) != set(MATRIX_CHECK_FIELDS):
        fail(f"{label}: semantic equality check fields are incomplete")
    if any(check[field] is not True for field in MATRIX_CHECK_FIELDS):
        fail(f"{label}: semantic equality check failed")


def verify_sha256_map(value: object, label: str) -> dict[str, str]:
    if not isinstance(value, dict):
        fail(f"{label}: member hash map is missing")
    for key, digest in value.items():
        if not isinstance(key, str) or not key:
            fail(f"{label}: member path is malformed")
        if not isinstance(digest, str) or len(digest) != 64 or any(
            character not in "0123456789abcdef" for character in digest
        ):
            fail(f"{label}: member digest is malformed")
    return value


def verify_matrix_preservation(
    value: object, source: dict, label: str = "forward preservation"
) -> None:
    if not isinstance(value, dict) or set(value) != {
        "source",
        "candidate",
        "mutable_inner_paths",
    }:
        fail(f"{label}: complete source/candidate member manifests are missing")
    source_manifest = value["source"]
    candidate_manifest = value["candidate"]
    if not isinstance(source_manifest, dict) or set(source_manifest) != MATRIX_MEMBER_MANIFEST_FIELDS:
        fail(f"{label}: source member manifest is incomplete")
    if not isinstance(candidate_manifest, dict) or set(candidate_manifest) != MATRIX_MEMBER_MANIFEST_FIELDS:
        fail(f"{label}: candidate member manifest is incomplete")
    source_parts = verify_sha256_map(source_manifest["parts"], f"{label}.source.parts")
    candidate_parts = verify_sha256_map(
        candidate_manifest["parts"], f"{label}.candidate.parts"
    )
    source_relationships = verify_sha256_map(
        source_manifest["relationships"], f"{label}.source.relationships"
    )
    candidate_relationships = verify_sha256_map(
        candidate_manifest["relationships"], f"{label}.candidate.relationships"
    )
    source_inner = verify_sha256_map(
        source_manifest["inner_members"], f"{label}.source.inner_members"
    )
    candidate_inner = verify_sha256_map(
        candidate_manifest["inner_members"], f"{label}.candidate.inner_members"
    )
    if (
        not isinstance(source_manifest["content_types"], str)
        or len(source_manifest["content_types"]) != 64
        or any(
            character not in "0123456789abcdef"
            for character in source_manifest["content_types"]
        )
    ):
        fail(f"{label}: source content-type digest is malformed")
    if (
        not isinstance(candidate_manifest["content_types"], str)
        or len(candidate_manifest["content_types"]) != 64
        or any(
            character not in "0123456789abcdef"
            for character in candidate_manifest["content_types"]
        )
    ):
        fail(f"{label}: candidate content-type digest is malformed")
    mutable = value["mutable_inner_paths"]
    if not isinstance(mutable, list) or any(not isinstance(path, str) or not path for path in mutable):
        fail(f"{label}: mutable inner path list is malformed")
    if mutable != sorted(set(mutable)):
        fail(f"{label}: mutable inner path list is not sorted and unique")
    expected_mutable = expected_matrix_mutable_inner_paths(source)
    if mutable != expected_mutable:
        fail(f"{label}: mutable inner path scope is broader than the selected T1 closure")
    if not set(source_parts) == set(candidate_parts):
        fail(f"{label}: forward package member set changed")
    if not set(source_relationships) == set(candidate_relationships):
        fail(f"{label}: forward relationship owner set changed")
    if not set(source_inner) == set(candidate_inner):
        fail(f"{label}: forward XLDM member set changed")
    allowed_outer_changes = {"/xl/workbook.bin", "/xl/model/item.data"}
    for path in source_parts:
        if path not in allowed_outer_changes and source_parts[path] != candidate_parts[path]:
            fail(f"{label}: unrelated OPC member changed: {path}")
    if source_relationships != candidate_relationships:
        fail(f"{label}: relationship XML changed outside the admitted model rewrite")
    if source_manifest["content_types"] != candidate_manifest["content_types"]:
        fail(f"{label}: content-types XML changed outside the admitted model rewrite")
    if not set(mutable).issubset(source_inner):
        fail(f"{label}: mutable path is absent from the source XLDM member set")
    for path in source_inner:
        if path not in mutable and source_inner[path] != candidate_inner[path]:
            fail(f"{label}: unrelated XLDM member changed: {path}")


def verify_matrix_result(result: dict, expected: tuple[str, int, int], manifest: dict) -> None:
    family, tables, relationships = expected
    if (result.get("family"), result.get("tables"), result.get("relationships")) != expected:
        fail("correctness result scale identity does not match the catalog")
    if result.get("endpoint_layout") not in manifest["recipe"]["endpoint_layouts"]:
        fail("correctness result endpoint layout is outside the recipe")
    if result.get("name_profile") not in manifest["recipe"]["name_profiles"]:
        fail("correctness result name profile is outside the recipe")
    if not isinstance(result.get("source_bytes"), int) or result["source_bytes"] <= 0:
        fail("correctness result source byte count is missing")
    source_hash = result.get("source_sha256")
    if not isinstance(source_hash, str) or len(source_hash) != 64 or any(
        character not in "0123456789abcdef" for character in source_hash
    ):
        fail("correctness result source hash is malformed")
    if result.get("source_unchanged") is not True:
        fail("correctness result source was changed")
    status = result.get("status")
    if status == "passed":
        if result.get("expected_success") is not True:
            fail("successful correctness result has the wrong expected-success value")
        if result.get("source_proof_status") != MATRIX_SOURCE_PROOF_COMPLETE:
            fail("successful correctness result does not establish complete source proof")
        for field in ("staged_bytes", "candidate_bytes", "output_bytes"):
            if not isinstance(result.get(field), int) or result[field] <= 0:
                fail(f"successful correctness result has no {field} boundary")
        if result.get("no_op_exact") is not True or result.get("inverse_exact") is not True:
            fail("successful correctness result failed exact no-op/inverse gates")
        verify_matrix_semantic_transition(result, tables, relationships)
        verify_matrix_preservation(result.get("preservation"), result["source_semantic"])
        if result.get("error") is not None:
            fail("successful correctness result contains an error")
    elif status == "expected_refusal":
        if result.get("expected_success") is not False:
            fail("expected refusal has the wrong expected-success value")
        if result.get("source_proof_status") != MATRIX_SOURCE_PROOF_UNAVAILABLE_GRAPH_LIMIT:
            fail("expected refusal does not state that full source proof is unavailable")
        if any(result.get(field) is not None for field in ("staged_bytes", "candidate_bytes", "output_bytes")):
            fail("expected refusal reports candidate bytes")
        if result.get("no_op_exact") is not False or result.get("inverse_exact") is not False:
            fail("expected refusal reports positive mutation gates")
        if any(
            result.get(field) is not None
            for field in (
                "source_semantic",
                "semantic_observed",
                "semantic",
                "reopened_semantic_observed",
                "reopened_semantic",
            )
        ):
            fail("expected refusal reports semantic candidate vectors")
        if result.get("preservation") is not None:
            fail("expected refusal reports forward member preservation without a full proof")
        error = result.get("error")
        if not isinstance(error, dict) or error.get("class") != "limit_exceeded":
            fail("expected refusal does not record the exact limit_exceeded class")
        if error.get("typed_match") is not True or "graph work limit exceeded" not in error.get("message", ""):
            fail("expected refusal does not record the graph-work resource")
    else:
        fail(f"unknown correctness result status: {status!r}")


def verify_correctness_receipt(path: Path, manifest_path: Path) -> None:
    receipt = load(path)
    manifest = load(manifest_path)
    if receipt.get("schema") != "xlsb-model-identity-profile-v1-correctness":
        fail("correctness receipt has the wrong schema")
    if receipt.get("fixture_kind") != "synthetic_complete_xldm140":
        fail("correctness receipt has the wrong fixture kind")
    if receipt.get("correctness_only") is not True or receipt.get("timings_collected") is not False:
        fail("correctness receipt is not explicitly timing-free")
    if receipt.get("native_acceptance_claim") not in (None, False):
        fail("correctness receipt makes a native acceptance claim")
    recipe = receipt.get("recipe")
    if not isinstance(recipe, dict):
        fail("correctness receipt recipe is missing")
    manifest_recipe = manifest.get("recipe", {})
    for field in (
        "id",
        "version",
        "source",
        "generator",
        "name_profiles",
        "endpoint_layouts",
    ):
        manifest_value = manifest_recipe.get(
            field,
            manifest.get(
                "recipe_source" if field == "source" else "recipe_generator"
                if field == "generator"
                else None
            ),
        )
        if recipe.get(field) != manifest_value:
            fail(f"correctness recipe field {field} is not bound to the corpus manifest")
    if recipe.get("scale_matrix") != manifest.get("scale_matrix"):
        fail("correctness recipe scale matrix differs from the corpus catalog")
    verify_matrix_limits(receipt, recipe, manifest)
    verify_olap_proof_limits(receipt, recipe, manifest)
    if receipt.get("caller_limits") != recipe.get("caller_limits"):
        fail("correctness top-level caller limits differ from recipe limits")
    if receipt.get("olap_proof_limits") != recipe.get("olap_proof_limits"):
        fail("correctness top-level OLAP proof limits differ from recipe limits")
    coverage = receipt.get("coverage")
    results = receipt.get("results")
    if not isinstance(coverage, dict) or not isinstance(results, list):
        fail("correctness coverage/results are missing")
    expected_matrix = {
        (case["family"], case["tables"], case["relationships"])
        for case in manifest.get("scale_matrix", [])
    }
    expected_layouts = set(manifest_recipe.get("endpoint_layouts", []))
    expected_profiles = set(manifest_recipe.get("name_profiles", []))
    expected_runs = len(expected_matrix) * len(expected_layouts) * len(expected_profiles)
    if coverage != {
        "cases": len(expected_matrix),
        "endpoint_layouts": len(expected_layouts),
        "name_profiles": len(expected_profiles),
        "runs": expected_runs,
        "passed": coverage.get("passed"),
        "expected_refusals": coverage.get("expected_refusals"),
    }:
        fail("correctness coverage dimensions are not the complete matrix")
    if len(results) != expected_runs:
        fail("correctness result count does not cover the complete matrix")
    seen = set()
    for result in results:
        if not isinstance(result, dict):
            fail("correctness result is not an object")
        coordinate = (
            result.get("family"),
            result.get("tables"),
            result.get("relationships"),
        )
        key = coordinate + (result.get("endpoint_layout"), result.get("name_profile"))
        if coordinate not in expected_matrix or key in seen:
            fail("correctness result has a duplicate or unknown matrix coordinate")
        seen.add(key)
        verify_matrix_result(result, coordinate, manifest)
    if seen != {
        coordinate + (layout, profile)
        for coordinate in expected_matrix
        for layout in expected_layouts
        for profile in expected_profiles
    }:
        fail("correctness result coordinates are incomplete")
    passed = sum(result.get("status") == "passed" for result in results)
    refusals = sum(result.get("status") == "expected_refusal" for result in results)
    if coverage.get("passed") != passed or coverage.get("expected_refusals") != refusals:
        fail("correctness status counts do not match the result vectors")
    print(f"verified complete XLSB identity matrix: {passed} passed, {refusals} bounded refusals")


def verify_binary_receipts(results: Path) -> str:
    before_lines = (results / "binary.sha256").read_text().splitlines()
    after_lines = (results / "binary-after.sha256").read_text().splitlines()
    if len(before_lines) != 1 or before_lines != after_lines:
        fail("profile executable changed during smoke")
    fields = before_lines[0].split()
    if len(fields) != 2:
        fail("binary digest receipt is malformed")
    digest, shown = fields
    if len(digest) != 64 or any(char not in "0123456789abcdef" for char in digest):
        fail("binary digest is malformed")
    binary = Path(shown)
    if not binary.is_absolute() or binary.name != "xlsb-model-identity-profile":
        fail("binary path is malformed")
    build = (results / "build-provenance.txt").read_text()
    if f"binary={shown}" not in build or f"{digest}  {shown}" not in build:
        fail("build provenance does not bind the binary digest")
    if binary.is_file() and sha256(binary) != digest:
        fail("live smoke binary hash changed")
    return digest


def verify_provenance(
    results: Path,
    root: Path,
    manifest: dict,
    manifest_commit: str,
    metadata_hash: str,
    source_manifest_hash: str,
    binary_digest: str,
) -> None:
    provenance = (results / "provenance.txt").read_text().splitlines()
    values = {}
    for line in provenance:
        if "=" in line:
            key, value = line.split("=", 1)
            values[key] = value
    current = subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), "rev-parse", "HEAD"],
        text=True,
    ).strip()
    if values.get("git_head") != current or values.get("git_head") != manifest_commit:
        fail("provenance Git head does not match the committed source manifest")
    expected_host = manifest.get("host_feature_commit")
    if expected_host and values.get("git_baseline") != expected_host:
        fail("provenance host implementation pin does not match corpus manifest")
    if values.get("neutral_baseline") != manifest.get("neutral_baseline_commit"):
        fail("provenance neutral baseline does not match corpus manifest")
    if values.get("git_status_relevant") != "":
        fail("relevant source tree was dirty or provenance was incomplete")
    metadata_before = results / "metadata-before.json"
    metadata_after = results / "metadata-after.json"
    if metadata_before.read_bytes() != metadata_after.read_bytes():
        fail("Cargo metadata changed during smoke")
    if sha256(metadata_before) != metadata_hash:
        fail("source manifest metadata hash does not bind metadata-before.json")
    if values.get("metadata_before_sha256") != metadata_hash or values.get(
        "metadata_after_sha256"
    ) != metadata_hash:
        fail("provenance metadata hashes do not match source manifest")
    if values.get("source_manifest_before_sha256") != source_manifest_hash or values.get(
        "source_manifest_after_sha256"
    ) != source_manifest_hash:
        fail("provenance source-manifest hashes do not match receipts")
    if values.get("binary_sha256") != binary_digest or values.get(
        "binary_after_sha256"
    ) != binary_digest:
        fail("provenance binary hashes do not match receipts")


def verify_phases(sample: dict, lane: str) -> None:
    phases = sample.get("phases")
    if not isinstance(phases, dict):
        fail(f"{lane}: phase receipt is missing")
    expected = {
        "open_ns",
        "stage_ns",
        "commit_ns",
        "save_ns",
        "reopen_ns",
        "inverse_ns",
        "validation_ns",
    }
    if set(phases) != expected:
        fail(f"{lane}: phase receipt fields are not explicit: {sorted(phases)}")
    for name, value in phases.items():
        if value is not None and (type(value) is not int or value <= 0):
            fail(f"{lane}: phase {name} is not a positive integer or null")
    required = {
        "neutral_open_tiny": {"open_ns"},
        "host_open_tiny": {"open_ns", "validation_ns"},
        "neutral_open_relationship": {"open_ns"},
        "host_stage_noop_tiny": {"stage_ns", "commit_ns", "validation_ns"},
        "host_stage_rename_relationship": {"stage_ns", "commit_ns", "validation_ns"},
        "host_commit_rename_relationship": {"commit_ns", "validation_ns"},
        "host_save_reopen_relationship": {"save_ns", "reopen_ns", "validation_ns"},
        "host_inverse_relationship": {"inverse_ns", "save_ns", "validation_ns"},
        "host_exact_cap_relationship": {"commit_ns", "validation_ns"},
        "host_refusal_opaque": {"stage_ns", "validation_ns"},
        "host_refusal_limit": {"commit_ns", "validation_ns"},
    }[lane]
    for name in required:
        if phases[name] is None:
            fail(f"{lane}: required phase {name} was not measured")


def verify_preservation(sample: dict, lane: str) -> None:
    preservation = sample.get("preservation")
    if lane in NEUTRAL:
        if preservation is not None:
            fail(f"{lane}: neutral lane unexpectedly claims host preservation")
        return
    if not isinstance(preservation, dict):
        fail(f"{lane}: preservation manifest was not reported")
    booleans = (
        "all_parts_equal",
        "unchanged_parts_equal",
        "relationships_equal",
        "content_types_equal",
        "inner_all_equal",
        "inner_unchanged_equal",
    )
    for field in booleans:
        if type(preservation.get(field)) is not bool:
            fail(f"{lane}: preservation field {field} is not boolean")
    for field in ("relationship_owner_count", "content_types_bytes"):
        value = preservation.get(field)
        if type(value) is not int or value <= 0:
            fail(f"{lane}: preservation count {field} is not positive")
    inner_count = preservation.get("inner_member_count")
    if type(inner_count) is not int or (lane not in REFUSALS and inner_count <= 0):
        fail(f"{lane}: preservation count inner_member_count is invalid")
    if lane in EXACT_PRESERVATION:
        for field in ("all_parts_equal", "relationships_equal", "content_types_equal", "inner_all_equal"):
            if preservation[field] is not True:
                fail(f"{lane}: exact preservation field {field} failed")
    else:
        for field in (
            "unchanged_parts_equal",
            "relationships_equal",
            "content_types_equal",
            "inner_unchanged_equal",
        ):
            if preservation[field] is not True:
                fail(f"{lane}: changed preservation field {field} failed")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--source-before", type=Path)
    parser.add_argument("--source-after", type=Path)
    parser.add_argument("--correctness-receipt", type=Path)
    args = parser.parse_args()
    if args.correctness_receipt is not None:
        verify_correctness_receipt(args.correctness_receipt, args.manifest)
        return
    if args.results is None or args.source_before is None or args.source_after is None:
        parser.error("smoke verification requires --results, --source-before, and --source-after")

    manifest = load(args.manifest)
    verify_matrix_runner_capture(args.results, args.manifest)
    if manifest.get("fixture_kind") != "synthetic_complete_xldm140":
        fail("manifest fixture kind is not complete synthetic XLDM 140")
    if manifest.get("native_acceptance_claim") is not False:
        fail("manifest makes a native acceptance claim")
    if args.source_before.read_bytes() != args.source_after.read_bytes():
        fail("source manifest changed during the run")
    manifest_checked, manifest_commit, metadata_hash = verify_source_manifest(
        args.source_before, args.root.resolve()
    )
    source_manifest_hash = sha256(args.source_before)
    binary_digest = verify_binary_receipts(args.results)
    verify_provenance(
        args.results,
        args.root.resolve(),
        manifest,
        manifest_commit,
        metadata_hash,
        source_manifest_hash,
        binary_digest,
    )

    receipts = {}
    for lane in sorted(LANES):
        path = args.results / f"{lane}.json"
        receipt = load(path)
        receipts[lane] = receipt
        if receipt.get("schema") != "xlsb-model-identity-profile-v1-smoke":
            fail(f"{lane}: wrong schema")
        if receipt.get("fixture_kind") != "synthetic_complete_xldm140":
            fail(f"{lane}: wrong fixture kind")
        if receipt.get("source_backed_api") is not True:
            fail(f"{lane}: source-backed API flag missing")
        if receipt.get("native_acceptance_claim") is not False:
            fail(f"{lane}: native acceptance flag is not false")
        recipe = receipt.get("recipe")
        if not isinstance(recipe, dict):
            fail(f"{lane}: recipe is missing")
        verify_olap_proof_limits(receipt, recipe, manifest)
        if receipt.get("sample_count") != 1 or receipt.get("warmup") != 0:
            fail(f"{lane}: smoke process does not have one sample and zero warmups")
        if receipt.get("table_count") != 1:
            fail(f"{lane}: smoke fixture is not one table")
        if receipt.get("relationship_count") not in (0, 1):
            fail(f"{lane}: smoke relationship count is outside the contract")
        samples = receipt.get("samples")
        if not isinstance(samples, list) or len(samples) != 1:
            fail(f"{lane}: sample list shape is invalid")
        sample = samples[0]
        if not isinstance(sample, dict):
            fail(f"{lane}: sample is not an object")
        verify_allocator(sample, lane)
        verify_phases(sample, lane)
        verify_preservation(sample, lane)
        verify_time_file(args.results / f"{lane}.time.txt", lane)
        if (args.results / f"{lane}.stderr.log").read_bytes() != b"":
            fail(f"{lane}: process stderr was not empty")
        if sample.get("source_unchanged") is not True:
            fail(f"{lane}: source changed or source gate was unavailable")
        opaque_ok = sample.get("opaque_ok")
        if opaque_ok is False:
            fail(f"{lane}: opaque preservation gate failed")
        if lane not in {"neutral_open_tiny", "neutral_open_relationship"} and opaque_ok is not True:
            fail(f"{lane}: opaque preservation gate was unavailable")
        expected_success = lane not in REFUSALS
        if receipt.get("expected_success") is not expected_success:
            fail(f"{lane}: expected-success field is incorrect")
        if sample.get("actual_success") is not expected_success:
            fail(f"{lane}: actual-success field is incorrect")
        if sample.get("semantic_ok") is not True:
            fail(f"{lane}: semantic gate failed")
        error = sample.get("error")
        if expected_success:
            if error is not None:
                fail(f"{lane}: successful lane has an error")
        else:
            if not isinstance(error, dict) or error.get("typed_match") is not True:
                fail(f"{lane}: refusal has no typed error receipt")
            if sample.get("candidate_bytes") is not None:
                fail(f"{lane}: refusal has candidate bytes")
            if error.get("class") != "invalid_format":
                fail(f"{lane}: refusal class is not exact invalid_format")
            expected_message = OPAQUE_ERROR if lane == "host_refusal_opaque" else LIMIT_ERROR
            if error.get("message") != expected_message:
                fail(f"{lane}: refusal resource/message changed: {error.get('message')!r}")
        if lane == "host_exact_cap_relationship":
            if sample.get("exact_cap_ok") is not True:
                fail(f"{lane}: exact cap case did not succeed")
            if sample.get("one_under_cap_refused") is not True:
                fail(f"{lane}: one-byte-under cap was accepted")

    if receipts["neutral_open_tiny"]["input_sha256"] != receipts["host_open_tiny"]["input_sha256"]:
        fail("neutral and host tiny fixtures differ")
    if receipts["neutral_open_relationship"]["input_sha256"] != receipts[
        "host_stage_rename_relationship"
    ]["input_sha256"]:
        fail("neutral and host relationship fixtures differ")
    print(
        f"verified {len(receipts)} synthetic identity smoke lanes "
        f"({manifest_checked} source inputs, binary {binary_digest})"
    )


if __name__ == "__main__":
    main()
