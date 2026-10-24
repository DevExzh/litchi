#!/usr/bin/env python3
"""Validate a completed API-debt disposition without approving it.

``final_api_inventory.py`` intentionally emits a blank, human-review scaffold.
This tool is a later audit step: it checks that inventory and disposition rows
are an exact one-to-one relation, that every row has nonblank human evidence,
and that the decision is one of the contract's concrete categories. It never
changes the files and does not infer that a row is semantically resolved.

Accepted decisions are:

* ``FIXED`` — a source repair is evidenced;
* ``BUILDER_DRAFT`` — a proven authoring builder is intentionally mutable;
* ``SUBSTRATE_TRANSPORT`` — only a transport-only CFB/ZIP substrate row;
* ``DEFERRED_REDACTION_22`` — only an explicitly identified item-22
  redaction row.

The last category is the sole permitted deferral. Placeholder decisions from
the scaffold and broad ``deferred``/``out-of-scope``/``wontfix`` values are
rejected.
"""
from __future__ import annotations

import argparse
import csv
from pathlib import Path


INVENTORY_FIELDS = ("crate", "category", "path", "detail", "target")
DISPOSITION_FIELDS = INVENTORY_FIELDS + ("decision", "evidence", "owner")
ROW_KEY = INVENTORY_FIELDS
ACCEPTED_DECISIONS = frozenset({
    "FIXED",
    "BUILDER_DRAFT",
    "SUBSTRATE_TRANSPORT",
    "DEFERRED_REDACTION_22",
})
SUBSTRATE_CRATES = frozenset({"litchi-cfb", "soapberry-zip"})
SCAFFOLD_DECISIONS = frozenset({
    "DRAFT_CANDIDATE_REVIEW",
    "SUBSTRATE_SCOPE_REVIEW",
    "UNASSIGNED_REQUIRED_REVIEW",
})


def _read_rows(path: Path, fields: tuple[str, ...], label: str) -> tuple[list[dict[str, str]], list[str]]:
    errors: list[str] = []
    try:
        with path.open(newline="", encoding="utf-8") as stream:
            reader = csv.reader(stream, delimiter="\t", strict=True)
            try:
                header = next(reader)
            except StopIteration:
                return [], [f"{label}: missing header"]
            if header != list(fields):
                errors.append(
                    f"{label}: expected header {' '.join(fields)}, "
                    f"got {header!r}"
                )
            rows: list[dict[str, str]] = []
            for line, values in enumerate(reader, 2):
                if len(values) != len(fields):
                    errors.append(
                        f"{label} line {line}: expected {len(fields)} fields, "
                        f"got {len(values)}"
                    )
                # Keep malformed rows available for one-to-one diagnostics,
                # while padding short rows so later checks cannot traceback.
                normalized = (values + [""] * len(fields))[:len(fields)]
                row = dict(zip(fields, normalized))
                rows.append(row)
                for field in ("crate", "category", "path"):
                    if not row[field].strip():
                        errors.append(
                            f"{label} line {line}: {field} must be nonblank"
                        )
    except (OSError, csv.Error) as error:
        return [], [f"{label}: cannot read: {error}"]
    return rows, errors


def _key(row: dict[str, str]) -> tuple[str, ...]:
    return tuple(row.get(field) or "" for field in ROW_KEY)


def _duplicate_errors(rows: list[dict[str, str]], label: str) -> list[str]:
    seen: dict[tuple[str, ...], int] = {}
    errors: list[str] = []
    for number, row in enumerate(rows, 2):
        key = _key(row)
        if key in seen:
            errors.append(
                f"{label}: duplicate row {number} (already at line {seen[key]}): "
                + "\t".join(key)
            )
        else:
            seen[key] = number
    return errors


def _decision_errors(row: dict[str, str], line: int) -> list[str]:
    errors: list[str] = []
    decision = row.get("decision", "").strip()
    evidence = row.get("evidence", "").strip()
    owner = row.get("owner", "").strip()
    label = f"disposition line {line}"
    for field, value in (("decision", decision), ("evidence", evidence), ("owner", owner)):
        if not value:
            errors.append(f"{label}: {field} must be nonblank")
    if decision and decision not in ACCEPTED_DECISIONS:
        errors.append(f"{label}: unsupported decision {decision!r}")
        if decision in SCAFFOLD_DECISIONS:
            errors.append(f"{label}: scaffold decision is still pending human review")
        return errors
    if decision == "BUILDER_DRAFT":
        if row.get("category") != "mut_draft" or "builder" not in row.get("path", "").split("::"):
            errors.append(f"{label}: BUILDER_DRAFT requires a mut_draft builder path")
    elif decision == "SUBSTRATE_TRANSPORT":
        if row.get("crate") not in SUBSTRATE_CRATES:
            errors.append(f"{label}: SUBSTRATE_TRANSPORT is limited to CFB/ZIP substrate crates")
    elif decision == "DEFERRED_REDACTION_22":
        lowered = evidence.lower()
        if row.get("category") != "redaction" or "redaction" not in lowered or "22" not in lowered:
            errors.append(
                f"{label}: DEFERRED_REDACTION_22 requires a redaction row whose evidence names item 22"
            )
    return errors


def validate(inventory: Path, disposition: Path) -> list[str]:
    """Return audit failures; an empty list means only structural checks passed."""
    inventory_rows, errors = _read_rows(inventory, INVENTORY_FIELDS, "inventory")
    disposition_rows, disposition_errors = _read_rows(
        disposition, DISPOSITION_FIELDS, "disposition")
    errors.extend(disposition_errors)
    errors.extend(_duplicate_errors(inventory_rows, "inventory"))
    errors.extend(_duplicate_errors(disposition_rows, "disposition"))

    inventory_keys = {_key(row) for row in inventory_rows}
    disposition_keys = {_key(row) for row in disposition_rows}
    for key in sorted(inventory_keys - disposition_keys):
        errors.append("disposition: missing inventory row: " + "\t".join(key))
    for key in sorted(disposition_keys - inventory_keys):
        errors.append("disposition: extra row not present in inventory: " + "\t".join(key))
    for line, row in enumerate(disposition_rows, 2):
        errors.extend(_decision_errors(row, line))
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inventory", type=Path)
    parser.add_argument("disposition", type=Path)
    args = parser.parse_args(argv)
    errors = validate(args.inventory, args.disposition)
    if errors:
        print("\n".join(errors))
        return 1
    print(
        f"validated {sum(1 for _ in args.inventory.open(encoding='utf-8')) - 1} "
        "disposition rows; no approval inferred"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
