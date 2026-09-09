#!/usr/bin/env python3
"""Probe the 0484 CLI fixtures with an independent DOCX consumer.

The replayable-tail CLI creates the source and candidate DOCX fixtures and a
report containing their Rust-side oracle hashes.  This probe copies those
inputs into an immutable attempt directory, checks the fixture manifests and
report against the bytes, then asks LibreOffice to export both documents as
UTF-8 text.  The expected text is extracted independently with Python's
``zipfile`` and ``xml.etree.ElementTree`` modules; no Litchi semantic helper
is used here.

The result is a synthetic LibreOffice readback only.  It makes no Microsoft
Office compatibility claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import signal
import subprocess
import sys
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

from common import ENV, REPO, ROOT, TEMP, meta, now, read, sha, write


SCHEMA = "docx-replayable-tail-append-consumer-probe-v1"
FIXTURE_SCHEMA = "docx-replayable-tail-append-fixture-v1"
REPORT_SCHEMA = "docx-replayable-tail-append-v1"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WORD = f"{{{WORD_NS}}}"
ATTEMPT = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
TEXT_MODES = ("short", "near_limit")
SOURCE_COUNT = 64
AUTHORED_COUNT = 64
CHUNK_MODE = "fixed64"
SAMPLES = 1
WARMUPS = 1
NEAR_TEXT_BYTES = 60 * 1024
COMMAND_TIMEOUT_SECONDS = 300


class ProbeError(RuntimeError):
    """An input, custody, consumer, or independent-readback failure."""


def fail(message: str) -> None:
    raise ProbeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def attempt_token(value: str) -> str:
    if ATTEMPT.fullmatch(value) is None:
        fail("--attempt must be a non-empty path-safe token")
    return value


def file_hash(path: Path) -> str:
    value = sha(path)
    require(SHA256.fullmatch(value) is not None, f"{path}: invalid SHA-256")
    return value


def text_hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def normalized_text(raw: bytes) -> str:
    """Decode UTF-8 while preserving all characters and paragraph newlines."""

    try:
        value = raw.decode("utf-8-sig")
    except UnicodeDecodeError as error:
        fail(f"LibreOffice text output is not UTF-8: {error}")
    # LibreOffice normally emits LF here.  Normalize platform line endings
    # without stripping trailing paragraph separators or any authored spaces.
    return value.replace("\r\n", "\n").replace("\r", "\n")


def _text_node(node: ET.Element) -> str:
    if node.tag == WORD + "t":
        return node.text or ""
    if node.tag == WORD + "tab":
        return "\t"
    if node.tag in {WORD + "br", WORD + "cr"}:
        return "\n"
    return ""


def document_text(path: Path) -> tuple[str, dict[str, Any]]:
    """Read Word paragraphs from document.xml independently of the library."""

    try:
        with zipfile.ZipFile(path, "r") as archive:
            xml = archive.read("word/document.xml")
    except (OSError, KeyError, zipfile.BadZipFile) as error:
        fail(f"{path}: DOCX/XML read failed: {error}")
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as error:
        fail(f"{path}: document.xml is not XML: {error}")
    paragraphs: list[str] = []
    body = root.find(f"{WORD}body")
    require(body is not None, f"{path}: document.xml has no w:body")
    for paragraph in body.iter(WORD + "p"):
        chunks: list[str] = []
        for node in paragraph.iter():
            chunks.append(_text_node(node))
        paragraphs.append("".join(chunks))
    value = "\n".join(paragraphs) + ("\n" if paragraphs else "")
    return value, {
        "xml_bytes": len(xml),
        "xml_sha256": hashlib.sha256(xml).hexdigest(),
        "paragraph_count": len(paragraphs),
        "text_bytes": len(value.encode("utf-8")),
        "text_sha256": text_hash(value),
        "contains_entity_text": "<&>" in value,
        "non_ascii_codepoints": sum(ord(char) > 127 for char in value),
    }


def safe_fixture_path(directory: Path, relative: str, field: str) -> Path:
    require(isinstance(relative, str) and relative, f"{field}: fixture path is missing")
    candidate = (directory / relative).resolve()
    base = directory.resolve()
    require(candidate != base and base in candidate.parents, f"{field}: fixture path escapes its directory")
    require(candidate.is_file() and not candidate.is_symlink(), f"{field}: fixture file is missing: {relative}")
    return candidate


def verify_manifest(path: Path, fixture_dir: Path) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict), f"{path}: manifest must be an object")
    require(value.get("schema") == FIXTURE_SCHEMA and value.get("version") == 1, f"{path}: fixture schema differs")
    require(value.get("source_count") == SOURCE_COUNT, f"{path}: source count differs")
    require(value.get("authored_count") == AUTHORED_COUNT, f"{path}: authored count differs")
    require(value.get("chunk_mode") == CHUNK_MODE, f"{path}: chunk mode differs")
    require(value.get("text_mode") in TEXT_MODES, f"{path}: text mode differs")
    for role in ("source", "candidate"):
        artifact = value.get(role)
        require(isinstance(artifact, dict), f"{path}: {role} artifact metadata is missing")
        artifact_path = safe_fixture_path(fixture_dir, artifact.get("file"), f"{path}:{role}")
        actual = meta(artifact_path)
        require(actual.get("bytes") == artifact.get("bytes"), f"{path}: {role} byte count differs")
        require(actual.get("sha256") == artifact.get("sha256"), f"{path}: {role} archive hash differs")
        require(SHA256.fullmatch(str(artifact.get("sha256", ""))) is not None, f"{path}: {role} hash is invalid")
        artifact["path"] = artifact_path
        artifact["actual"] = actual
    source_xml = document_xml_hash(safe_fixture_path(fixture_dir, value["source"]["file"], f"{path}:source"))
    candidate_xml = document_xml_hash(safe_fixture_path(fixture_dir, value["candidate"]["file"], f"{path}:candidate"))
    require(source_xml == value.get("source_main_xml_sha256"), f"{path}: source main XML hash differs")
    require(candidate_xml == value.get("candidate_main_xml_sha256"), f"{path}: candidate main XML hash differs")
    require(SHA256.fullmatch(source_xml) is not None and SHA256.fullmatch(candidate_xml) is not None, f"{path}: XML hash is invalid")
    value["manifest_path"] = path
    value["manifest_sha256"] = file_hash(path)
    return value


def document_xml_hash(path: Path) -> str:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            content = archive.read("word/document.xml")
    except (OSError, KeyError, zipfile.BadZipFile) as error:
        fail(f"{path}: cannot read document.xml for binding: {error}")
    return hashlib.sha256(content).hexdigest()


def verify_report(
    path: Path,
    manifests: list[dict[str, Any]],
    fixture_dir: Path,
) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    value = read(path)
    require(isinstance(value, dict), f"{path}: report must be an object")
    require(value.get("schema") == REPORT_SCHEMA and value.get("version") == 1, f"{path}: report schema differs")
    config = value.get("config")
    require(isinstance(config, dict), f"{path}: report config is missing")
    require(config.get("source_counts") == [SOURCE_COUNT], f"{path}: report source count differs")
    require(config.get("authored_counts") == [AUTHORED_COUNT], f"{path}: report authored count differs")
    require(config.get("chunk_modes") == [CHUNK_MODE], f"{path}: report chunk mode differs")
    require(config.get("text_modes") == list(TEXT_MODES), f"{path}: report text mode set differs")
    require(config.get("samples") == SAMPLES and config.get("warmups") == WARMUPS, f"{path}: report sample settings differ")
    require(config.get("fixture_dir") == str(fixture_dir), f"{path}: fixture directory binding differs")
    cases = value.get("cases")
    require(isinstance(cases, list) and len(cases) == len(manifests), f"{path}: report case count differs")
    by_mode: dict[str, dict[str, Any]] = {}
    for case in cases:
        require(isinstance(case, dict), f"{path}: report case is invalid")
        require(case.get("source_count") == SOURCE_COUNT, f"{path}: case source count differs")
        require(case.get("authored_count") == AUTHORED_COUNT, f"{path}: case authored count differs")
        require(case.get("chunk_mode") == CHUNK_MODE, f"{path}: case chunk mode differs")
        mode = case.get("text_mode")
        require(mode in TEXT_MODES and mode not in by_mode, f"{path}: report text case set differs")
        source = case.get("source")
        authored = case.get("authored")
        oracle = case.get("oracle")
        require(
            isinstance(source, dict) and isinstance(authored, dict) and isinstance(oracle, dict),
            f"{path}: case archive/authored metadata is missing",
        )
        require(authored.get("authored_count") == AUTHORED_COUNT, f"{path}: case authored proof count differs")
        require(authored.get("chunk_mode") == CHUNK_MODE, f"{path}: case authored proof chunk mode differs")
        require(authored.get("text_mode") == mode, f"{path}: case authored proof text mode differs")
        require(isinstance(authored.get("text_bytes"), int), f"{path}: case authored text byte proof is missing")
        if mode == "near_limit":
            require(
                authored["text_bytes"] == NEAR_TEXT_BYTES * AUTHORED_COUNT,
                f"{path}: near text proof is not 60 KiB per authored paragraph",
            )
        for field in ("candidate_xml_exact", "candidate_semantic_exact", "physical_order_exact", "inverse_exact"):
            require(oracle.get(field) is True, f"{path}: case oracle {field} failed")
        by_mode[mode] = case
    require(set(by_mode) == set(TEXT_MODES), f"{path}: report does not contain both text modes")
    manifests_by_mode = {manifest["text_mode"]: manifest for manifest in manifests}
    require(set(manifests_by_mode) == set(TEXT_MODES), "fixture manifests do not contain both text modes")
    for mode, case in by_mode.items():
        manifest = manifests_by_mode[mode]
        source = case["source"]
        oracle = case["oracle"]
        require(source.get("archive_sha256") == manifest["source"]["sha256"], f"{mode}: report source archive binding differs")
        require(oracle.get("candidate_archive_sha256") == manifest["candidate"]["sha256"], f"{mode}: report candidate archive binding differs")
        require(source.get("main_xml_sha256") == manifest["source_main_xml_sha256"], f"{mode}: report source XML binding differs")
        require(oracle.get("candidate_main_xml_sha256") == manifest["candidate_main_xml_sha256"], f"{mode}: report candidate XML binding differs")
    return value, by_mode


def run_command(argv: list[str], label: str, cwd: Path, directory: Path) -> dict[str, Any]:
    stdout = directory / f"{label}.stdout"
    stderr = directory / f"{label}.stderr"
    receipt = directory / f"{label}.json"
    require(not any(path.exists() for path in (stdout, stderr, receipt)), f"{label}: command artifacts already exist")
    started = now()
    timed_out = False
    launch_error: str | None = None
    process: subprocess.Popen[bytes] | None = None
    try:
        with stdout.open("xb") as out, stderr.open("xb") as err:
            process = subprocess.Popen(
                argv,
                cwd=cwd,
                env=ENV,
                stdout=out,
                stderr=err,
                start_new_session=True,
            )
            try:
                exit_code = process.wait(timeout=COMMAND_TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                timed_out = True
                os.killpg(process.pid, signal.SIGKILL)
                exit_code = process.wait()
    except OSError as error:
        exit_code = None
        launch_error = f"{type(error).__name__}: {error}"
    result: dict[str, Any] = {
        "argv": argv,
        "cwd": str(cwd),
        "started_utc": started,
        "finished_utc": now(),
        "exit_code": exit_code,
        "timed_out": timed_out,
        "stdout": meta(stdout) if stdout.is_file() else None,
        "stderr": meta(stderr) if stderr.is_file() else None,
    }
    if launch_error is not None:
        result["launch_error"] = launch_error
    write(receipt, result)
    require(exit_code == 0 and not timed_out, f"{label} failed; receipt retained: {receipt}")
    return result


def copy_create(source: Path, destination: Path) -> dict[str, Any]:
    require(not destination.exists(), f"refusing to replace retained artifact: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    copied = meta(destination)
    require(copied == meta(source), f"copied artifact differs: {source}")
    return copied


def readback_row(
    mode: str,
    source: Path,
    candidate: Path,
    text_dir: Path,
    authored_text_bytes: int,
) -> dict[str, Any]:
    source_expected, source_details = document_text(source)
    candidate_expected, candidate_details = document_text(candidate)
    source_output_path = text_dir / "source.txt"
    candidate_output_path = text_dir / "candidate.txt"
    require(source_output_path.is_file() and candidate_output_path.is_file(), f"{mode}: LibreOffice text artifacts are missing")
    source_output = normalized_text(source_output_path.read_bytes())
    candidate_output = normalized_text(candidate_output_path.read_bytes())
    row = {
        "text_mode": mode,
        "source_docx": meta(source),
        "candidate_docx": meta(candidate),
        "source_txt": meta(source_output_path),
        "candidate_txt": meta(candidate_output_path),
        "source_expected": source_details,
        "candidate_expected": candidate_details,
        "source_output": {
            "text_bytes": len(source_output.encode("utf-8")),
            "text_sha256": text_hash(source_output),
            "preview": source_output[:128],
            "suffix": source_output[-128:],
        },
        "candidate_output": {
            "text_bytes": len(candidate_output.encode("utf-8")),
            "text_sha256": text_hash(candidate_output),
            "preview": candidate_output[:128],
            "suffix": candidate_output[-128:],
        },
        "authored_text_bytes": authored_text_bytes,
        "contains_entity_text": "<&>" in candidate_expected,
        "verified": source_output == source_expected and candidate_output == candidate_expected,
    }
    require(row["contains_entity_text"], f"{mode}: candidate text lost <&> entity content")
    require(
        candidate_details["text_bytes"] - source_details["text_bytes"]
        == authored_text_bytes + AUTHORED_COUNT,
        f"{mode}: independent paragraph text size differs from authored proof",
    )
    return row


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fixture-dir", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--attempt", required=True)
    args = parser.parse_args(argv)
    try:
        attempt = attempt_token(args.attempt)
        fixture_dir = args.fixture_dir.resolve(strict=True)
        report = args.report.resolve(strict=True)
        require(fixture_dir.is_dir(), f"fixture directory is not a directory: {fixture_dir}")
        manifests_paths = sorted(fixture_dir.glob("*-hashes.json"))
        require(len(manifests_paths) == len(TEXT_MODES), "fixture directory must contain exactly short and near manifests")
        manifests = [verify_manifest(path, fixture_dir) for path in manifests_paths]
        _, report_cases = verify_report(report, manifests, fixture_dir)
        retained = ROOT / "consumer" / attempt
        require(not retained.exists(), f"refusing to replace consumer attempt: {retained}")
        retained.mkdir(parents=True)
        work = TEMP / f"consumer-0484-{attempt}"
        require(not work.exists(), f"refusing to replace consumer work directory: {work}")
        work.mkdir(parents=True)
        report_copy = retained / "cli-report.json"
        report_meta = copy_create(report, report_copy)
        rows: list[dict[str, Any]] = []
        retained_inputs: list[dict[str, Any]] = []
        for manifest in sorted(manifests, key=lambda value: value["text_mode"]):
            mode = manifest["text_mode"]
            folder = retained / mode
            folder.mkdir()
            source = folder / "source.docx"
            candidate = folder / "candidate.docx"
            manifest_copy = folder / "fixture-hashes.json"
            source_meta = copy_create(manifest["source"]["path"], source)
            candidate_meta = copy_create(manifest["candidate"]["path"], candidate)
            manifest_meta = copy_create(manifest["manifest_path"], manifest_copy)
            require(source_meta["sha256"] == manifest["source"]["sha256"], f"{mode}: retained source hash differs")
            require(candidate_meta["sha256"] == manifest["candidate"]["sha256"], f"{mode}: retained candidate hash differs")
            text_dir = folder / "text"
            text_dir.mkdir()
            profile = work / f"profile-{mode}"
            profile.mkdir()
            libreoffice = run_command(
                [
                    "/usr/bin/libreoffice",
                    "--headless",
                    "-env:UserInstallation=" + profile.as_uri(),
                    "--convert-to",
                    "txt:Text (encoded):UTF8",
                    "--outdir",
                    str(text_dir),
                    str(source),
                    str(candidate),
                ],
                "libreoffice",
                REPO,
                folder,
            )
            row = readback_row(
                mode,
                source,
                candidate,
                text_dir,
                report_cases[mode]["authored"]["text_bytes"],
            )
            write(folder / "readback.json", row)
            require(row["verified"], f"{mode}: independent TXT readback differs; artifacts retained")
            rows.append(row)
            retained_inputs.append(
                {
                    "text_mode": mode,
                    "fixture_manifest": {
                        "path": str(manifest["manifest_path"]),
                        "sha256": manifest["manifest_sha256"],
                    },
                    "copied_manifest": manifest_meta,
                    "source": source_meta,
                    "candidate": candidate_meta,
                    "libreoffice": libreoffice,
                    "report_case": report_cases[mode],
                }
            )
        result = {
            "schema": SCHEMA,
            "version": 1,
            "status": "pass",
            "attempt": attempt,
            "fixture_dir": str(fixture_dir),
            "fixture_manifests": [
                {"path": str(value["manifest_path"]), "sha256": value["manifest_sha256"]}
                for value in manifests
            ],
            "report": {"path": str(report), "sha256": file_hash(report), "copied": report_meta},
            "inputs": retained_inputs,
            "cases": rows,
            "finished_utc": now(),
            "scope": "independent Python DOCX ZIP/XML extraction and LibreOffice UTF-8 TXT export; synthetic consumer probe only; no Microsoft Office compatibility claim",
        }
        write(retained / "result.json", result)
        print(f"LibreOffice independently read back {len(rows)} 0484 DOCX fixture pairs.")
        return 0
    except (ProbeError, OSError, ValueError, KeyError, json.JSONDecodeError) as error:
        print(f"consumer-probe.py: FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
