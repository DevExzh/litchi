#!/usr/bin/env python3
"""Validate the outside-clock PPT scoped-edit correctness receipt."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
from typing import NoReturn


EXPECTED_PATH = "clip-0000.avi"
EXPECTED_MEDIA_ID = 1
REQUIRED = {
    "mode",
    "fixture",
    "media_id",
    "requested_external_media_path",
    "external_media_path",
    "source_len",
    "target_len",
    "source_sha256",
    "target_sha256",
    "inverse_source_sha256",
    "inverse_matches_source",
}
OPTIONAL_REPLAY = {"replay_target_sha256", "replay_matches_target"}


def fail(message: str) -> NoReturn:
    raise SystemExit(message)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        fail("usage: validate_ppt_correctness_receipt.py RECEIPT FIXTURE")
    receipt_path, fixture_path = map(Path, argv[1:])
    row = json.loads(receipt_path.read_text(encoding="utf-8"))
    if not isinstance(row, dict):
        fail("correctness receipt is not an object")
    missing = sorted(REQUIRED - row.keys())
    if missing:
        fail(f"missing correctness fields: {', '.join(missing)}")
    fixture = fixture_path.read_bytes()
    expected_source = hashlib.sha256(fixture).hexdigest()
    if row["mode"] != "scoped_patch_correctness":
        fail("receipt mode is not the correctness companion")
    if not isinstance(row["fixture"], str):
        fail("receipt fixture is not a path")
    if Path(row["fixture"]).resolve() != fixture_path.resolve():
        fail("receipt fixture is not the selected source fixture")
    if type(row["media_id"]) is not int or row["media_id"] != EXPECTED_MEDIA_ID:
        fail("receipt media ID is not 1")
    if row["requested_external_media_path"] != EXPECTED_PATH:
        fail("receipt requested path is not clip-0000.avi")
    if row["external_media_path"] != EXPECTED_PATH:
        fail("receipt target path is not clip-0000.avi")
    if (
        type(row["source_len"]) is not int
        or type(row["target_len"]) is not int
        or row["source_len"] != len(fixture)
        or row["target_len"] <= 0
    ):
        fail("receipt extents are invalid")
    if row["source_sha256"] != expected_source:
        fail("receipt source hash does not match the fixture")
    if row["source_sha256"] == row["target_sha256"]:
        fail("correctness target bytes are unchanged")
    replay_fields = OPTIONAL_REPLAY & row.keys()
    if replay_fields and replay_fields != OPTIONAL_REPLAY:
        fail("forward replay fields must be supplied together")
    if replay_fields:
        if row["replay_target_sha256"] != row["target_sha256"]:
            fail("forward replay bytes do not match the committed target")
        if row["replay_matches_target"] is not True:
            fail("forward replay did not reproduce the committed target")
    if row["inverse_source_sha256"] != row["source_sha256"]:
        fail("inverse source hash does not match the source")
    if row["inverse_matches_source"] is not True:
        fail("inverse did not restore the source")
    for field in (
        "source_sha256",
        "target_sha256",
        "inverse_source_sha256",
    ):
        value = row[field]
        if not isinstance(value, str) or len(value) != 64:
            fail(f"{field} is not a SHA-256 hex digest")
        try:
            int(value, 16)
        except ValueError:
            fail(f"{field} is not a SHA-256 hex digest")
    summary = {
        "status": "passed",
        "source_sha256": row["source_sha256"],
        "target_sha256": row["target_sha256"],
        "external_media_path": row["external_media_path"],
        "inverse_matches_source": row["inverse_matches_source"],
    }
    if replay_fields:
        summary.update({
            "replay_target_sha256": row["replay_target_sha256"],
            "replay_matches_target": row["replay_matches_target"],
        })
    print(json.dumps(summary, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
