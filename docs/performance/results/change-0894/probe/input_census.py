#!/usr/bin/env python3
"""Create a complete, hash-bound XLSX input census for change 0894.

The Rust probe repeats the file hash before opening each input.  This small
helper is retained as a host-side manifest receipt: it makes discovery order,
duplicate elimination, metadata failures, and unreadable roots explicit before
either baseline or candidate binary is run.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from pathlib import Path
from typing import Any, Iterable


SCHEMA = "litchi.xlsx-input-census-0894.v1"
CHUNK = 128 * 1024


def absolute(path: Path) -> Path:
    try:
        return path.resolve(strict=False)
    except OSError:
        return Path(os.path.abspath(os.fspath(path)))


def walk_xlsx(root: Path) -> tuple[set[Path], list[str]]:
    files: set[Path] = set()
    errors: list[str] = []

    def visit(path: Path) -> None:
        try:
            stat = path.stat()
        except OSError as error:
            errors.append(f"{path}: {error}")
            return
        if path.is_file():
            if path.suffix.casefold() == ".xlsx":
                files.add(absolute(path))
            return
        if not path.is_dir():
            errors.append(f"{path}: not a regular file or directory")
            return
        try:
            children = sorted(path.iterdir(), key=lambda child: os.fspath(child))
        except OSError as error:
            errors.append(f"{path}: {error}")
            return
        for child in children:
            # Directory symlinks are deliberately not followed.  A symlinked
            # file remains an explicit input and is resolved in its receipt.
            try:
                if child.is_symlink() and child.is_dir():
                    errors.append(f"{child}: skipped directory symlink")
                    continue
            except OSError as error:
                errors.append(f"{child}: {error}")
                continue
            visit(child)

    visit(root)
    return files, errors


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while True:
            block = stream.read(CHUNK)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def file_record(path: Path, roots: Iterable[Path]) -> dict[str, Any]:
    record: dict[str, Any] = {
        "path": os.fspath(path),
        "bytes": None,
        "sha256": None,
        "metadata_error": None,
        "relative_names": [],
    }
    try:
        record["bytes"] = path.stat().st_size
    except OSError as error:
        record["metadata_error"] = str(error)
        return record
    try:
        record["sha256"] = sha256(path)
    except OSError as error:
        record["metadata_error"] = str(error)
    for root in roots:
        try:
            record["relative_names"].append(os.fspath(path.relative_to(root)))
        except ValueError:
            continue
    return record


def write_line(stream: Any, value: dict[str, Any]) -> None:
    stream.write(json.dumps(value, sort_keys=True, separators=(",", ":")))
    stream.write("\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("roots", nargs="+", type=Path)
    parser.add_argument("--output", type=Path, default=None)
    arguments = parser.parse_args()

    roots = [absolute(root) for root in arguments.roots]
    files: set[Path] = set()
    errors: list[str] = []
    for root in roots:
        discovered, root_errors = walk_xlsx(root)
        files.update(discovered)
        errors.extend(root_errors)

    output = sys.stdout if arguments.output is None else arguments.output.open("w", encoding="utf-8")
    try:
        for path in sorted(files, key=lambda item: os.fspath(item)):
            write_line(
                output,
                {
                    "schema": SCHEMA,
                    "kind": "input",
                    "status": "metadata_error"
                    if (record := file_record(path, roots))["metadata_error"]
                    else "discovered",
                    "input": record,
                },
            )
        for error in errors:
            write_line(
                output,
                {
                    "schema": SCHEMA,
                    "kind": "discovery_error",
                    "status": "error",
                    "error": error,
                },
            )
        write_line(
            output,
            {
                "schema": SCHEMA,
                "kind": "summary",
                "roots": [os.fspath(root) for root in roots],
                "inputs": len(files),
                "discovery_errors": len(errors),
                "historical_change_0642_probe_corpus": {
                    "status": "not_included",
                    "reason": "the retained change-0642 packet has probe sources but no XLSX corpus file",
                },
            },
        )
    finally:
        if output is not sys.stdout:
            output.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
