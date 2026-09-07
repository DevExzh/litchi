#!/usr/bin/env python3
"""Bind independently verified source-tail candidate ODP fixtures.

The candidate binary reports hashes and member identities, but it does not
carry the archive bytes needed to prove that a style-free generated page was
inserted at the source XML boundary.  This helper retains one source/output
fixture per shape, runs the native ZIP/XML oracle over those retained files,
and writes an immutable receipt consumed by the candidate report oracle.
It performs no build and does not execute a benchmark binary.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from typing import Any


SCHEMA = "litchi-0457-odp-source-tail-candidate-output-binding-v1"
NATIVE_SCHEMA = "litchi-0457-native-odp-append-oracle-v1"
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
CONTENT_NAME = "content.xml"
MANIFEST_NAME = "META-INF/manifest.xml"
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
MANIFEST_FULL_PATH = f"{{{MANIFEST_NS}}}full-path"
MANIFEST_MEDIA_TYPE = f"{{{MANIFEST_NS}}}media-type"
OPAQUE_PATH = "Opaque/litchi-perf-odp-existing-append-opaque.bin"
REQUIRED_MEMBERS = {
    "mimetype",
    CONTENT_NAME,
    "styles.xml",
    "meta.xml",
    OPAQUE_PATH,
    MANIFEST_NAME,
}


class BindError(ValueError):
    pass


def fail(message: str) -> None:
    raise BindError(message)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def canonical(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def copy_new(source: Path, destination: Path) -> None:
    if not source.is_file() or source.is_symlink():
        fail(f"{source}: expected a regular file")
    if destination.exists() or destination.is_symlink():
        fail(f"refusing to overwrite retained fixture {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    if source.stat().st_size != destination.stat().st_size or sha(source) != sha(destination):
        fail(f"fixture copy verification failed for {destination}")


def import_native(path: Path):
    if not path.is_file() or path.is_symlink():
        fail(f"native output oracle is missing: {path}")
    spec = importlib.util.spec_from_file_location("change0457_native_output_oracle", path)
    if spec is None or spec.loader is None:
        fail(f"cannot load native output oracle {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    if not callable(getattr(module, "parse_zip", None)):
        fail(f"native output oracle {path} does not expose parse_zip")
    return module


def variant_text(index: int) -> str:
    variants = (
        "plain slide",
        "Unicode café Δ 中",
        'entities <&> "quoted"',
        "mixed façade Ω <&> value",
    )
    return variants[index % 4]


def title(index: int) -> str:
    return f"litchi-perf-odp-buffered-title-{index:05d} {variant_text(index)}"


def body(index: int) -> str:
    return f"litchi-perf-odp-buffered-body-{index:05d} {variant_text(index)}"


def manifest_media_types(archive: Any) -> dict[str, str]:
    entry = archive.by_name.get(MANIFEST_NAME.encode("utf-8"))
    if entry is None:
        fail("fixture has no META-INF/manifest.xml")
    try:
        root = ET.fromstring(entry.decoded)
    except ET.ParseError as error:
        fail(f"manifest.xml is not parseable: {error}")
    result: dict[str, str] = {}
    for item in root.iter(f"{{{MANIFEST_NS}}}file-entry"):
        full_path = item.get(MANIFEST_FULL_PATH)
        media_type = item.get(MANIFEST_MEDIA_TYPE)
        if full_path and media_type is not None and full_path != "/":
            result[full_path] = media_type
    return result


def member_identities(native: Any, path: Path) -> list[dict[str, Any]]:
    archive = native.parse_zip(path)
    media_types = manifest_media_types(archive)
    members: list[dict[str, Any]] = []
    method_names = {0: "Store", 8: "Deflate"}
    for entry in archive.entries:
        if entry.is_dir:
            fail(f"fixture contains a directory member: {entry.name!r}")
        method = entry.central["method"]
        if method not in method_names:
            fail(f"fixture member {entry.name!r} uses unsupported method {method}")
        members.append(
            {
                "compressed_bytes": len(entry.compressed_payload),
                "compressed_sha256": sha_bytes(entry.compressed_payload),
                "compression_method": method_names[method],
                "crc32": entry.central["crc32"],
                "data_descriptor": bool(entry.local_fixed["flags"] & 0x08),
                "decoded_bytes": len(entry.decoded),
                "decoded_sha256": sha_bytes(entry.decoded),
                "media_type": media_types.get(entry.name, "application/octet-stream"),
                "path": entry.name,
            }
        )
    members.sort(key=lambda item: item["path"])
    if {item["path"] for item in members} != REQUIRED_MEMBERS:
        fail(f"fixture member set differs from the six-member corpus: {path}")
    return members


def member_map(members: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    return {item["path"]: item for item in members}


def validate_against_control(shape: str, output_root: Path, source_path: Path, source_members: list[dict[str, Any]], source: dict[str, Any]) -> None:
    if source.get("archive_sha256") != sha(source_path) or source.get("archive_bytes") != source_path.stat().st_size:
        fail(f"{shape} source fixture differs from corpus-bindings.json")
    source_map = member_map(source_members)
    if source_map[CONTENT_NAME]["decoded_sha256"] != source.get("content_xml_sha256") or source_map[CONTENT_NAME]["decoded_bytes"] != source.get("content_xml_bytes"):
        fail(f"{shape} source content.xml differs from corpus-bindings.json")
    if source_map["styles.xml"]["decoded_sha256"] != "d9881e91085516246a19c30d9e5cde39a8b10d7e42120b135f48f5ca8afef8d2":
        fail(f"{shape} source styles.xml is not the pinned control member")
    if source_map["meta.xml"]["decoded_sha256"] != "c7e55a3560c73aa42da85eec4751c3e78b5cc53ff964f50acba6c5cd105e6719":
        fail(f"{shape} source meta.xml is not the pinned control member")


def run_native(native_path: Path, source: Path, output: Path, request: dict[str, str]) -> dict[str, Any]:
    completed = subprocess.run(
        [
            sys.executable,
            "-B",
            str(native_path),
            "--source",
            str(source),
            "--output",
            str(output),
            "--title",
            request["title"],
            "--body",
            request["body"],
            "--name",
            request["name"],
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        fail(f"native output oracle rejected {output.name}: {completed.stderr.strip() or completed.stdout.strip()}")
    try:
        receipt = json.loads(completed.stdout)
    except json.JSONDecodeError as error:
        fail(f"native output oracle emitted invalid JSON: {error}")
    receipt = obj(receipt, "native output receipt")
    if receipt.get("schema") != NATIVE_SCHEMA or receipt.get("status") != "validated":
        fail("native output oracle did not return a validated receipt")
    return receipt


def fixture_record(path: Path, relative_to: Path) -> dict[str, Any]:
    return {
        "path": path.relative_to(relative_to).as_posix(),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def bind_shape(
    shape: str,
    source_input: Path,
    output_input: Path,
    output_root: Path,
    native_path: Path,
    native: Any,
    control_shape: dict[str, Any],
) -> dict[str, Any]:
    fixture_dir = output_root / "output-fixtures"
    source_path = fixture_dir / f"{shape}-source.odp"
    output_path = fixture_dir / f"{shape}-output.odp"
    copy_new(source_input, source_path)
    copy_new(output_input, output_path)
    source_members = member_identities(native, source_path)
    output_members = member_identities(native, output_path)
    control_source = obj(control_shape["source"], f"corpus-bindings.shapes.{shape}.source")
    validate_against_control(shape, output_root, source_path, source_members, control_source)
    source_map = member_map(source_members)
    output_map = member_map(output_members)
    for path in REQUIRED_MEMBERS - {CONTENT_NAME}:
        if source_map[path] != output_map[path]:
            fail(f"{shape} untouched member changed before native receipt binding: {path}")
    request = {"title": title(SHAPES[shape]), "body": body(SHAPES[shape]), "name": f"page{SHAPES[shape] + 1}"}
    native_result = run_native(native_path, source_path, output_path, request)
    native_source = obj(native_result["source"], f"native {shape} source")
    native_output = obj(native_result["output"], f"native {shape} output")
    content = obj(native_result["content"], f"native {shape} content")
    if native_source.get("sha256") != sha(source_path) or native_source.get("bytes") != source_path.stat().st_size:
        fail(f"{shape} native source receipt does not bind retained source")
    if native_output.get("sha256") != sha(output_path) or native_output.get("bytes") != output_path.stat().st_size:
        fail(f"{shape} native output receipt does not bind retained output")
    if content.get("source_content_xml_sha256") != source_map[CONTENT_NAME]["decoded_sha256"] or content.get("source_content_xml_bytes") != source_map[CONTENT_NAME]["decoded_bytes"]:
        fail(f"{shape} native source content identity differs from parsed member")
    if content.get("output_content_xml_sha256") != output_map[CONTENT_NAME]["decoded_sha256"] or content.get("output_content_xml_bytes") != output_map[CONTENT_NAME]["decoded_bytes"]:
        fail(f"{shape} native output content identity differs from parsed member")
    native_result["source"]["path"] = source_path.relative_to(output_root).as_posix()
    native_result["output"]["path"] = output_path.relative_to(output_root).as_posix()
    return {
        "request": request,
        "fixtures": {
            "source": fixture_record(source_path, output_root),
            "output": fixture_record(output_path, output_root),
        },
        "source": {
            "archive_sha256": sha(source_path),
            "archive_bytes": source_path.stat().st_size,
            "content_xml_sha256": source_map[CONTENT_NAME]["decoded_sha256"],
            "content_xml_bytes": source_map[CONTENT_NAME]["decoded_bytes"],
            "members": source_members,
        },
        "output": {
            "archive_sha256": sha(output_path),
            "archive_bytes": output_path.stat().st_size,
            "content_xml_sha256": output_map[CONTENT_NAME]["decoded_sha256"],
            "content_xml_bytes": output_map[CONTENT_NAME]["decoded_bytes"],
            "members": output_members,
        },
        "native_result": native_result,
        "native_result_sha256": sha_bytes(canonical(native_result)),
    }


def update_protocol(protocol_path: Path, output_binding_path: Path, binding_sha: str, native_path: Path, output_root: Path) -> None:
    if not protocol_path.is_file() or protocol_path.is_symlink():
        fail(f"protocol is missing: {protocol_path}")
    runs = output_root / "runs"
    if runs.is_dir() and any(runs.iterdir()):
        fail("capture has started; refusing to rewrite protocol output binding")
    protocol = obj(load(protocol_path), "candidate protocol")
    output = obj(protocol.get("output_binding"), "candidate protocol.output_binding")
    if output.get("path") != output_binding_path.name:
        fail("candidate protocol output binding path differs")
    if output.get("sha256") not in ("pending-output-preflight", "bound-at-capture"):
        fail("candidate protocol already has a final output binding; refusing overwrite")
    output["sha256"] = binding_sha
    binder_path = output_root / "bind-output.py"
    output["binder_path"] = "bind-output.py"
    output["binder_sha256"] = sha(binder_path)
    native_record = obj(output.get("native_oracle"), "candidate protocol.output_binding.native_oracle")
    native_relative = Path(os.path.relpath(native_path, output_root)).as_posix()
    if native_relative != "../native/verify-output.py":
        fail("native output oracle must be docs/performance/results/change-0457/native/verify-output.py")
    native_record["path"] = native_relative
    native_record["sha256"] = sha(native_path)
    protocol_path.write_text(json.dumps(protocol, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--native-oracle", type=Path, required=True)
    parser.add_argument("--output-binding", type=Path, default=Path(__file__).resolve().parent / "output-binding.json")
    parser.add_argument("--protocol", type=Path, default=None)
    parser.add_argument("--fixture", nargs=3, action="append", metavar=("SHAPE", "SOURCE", "OUTPUT"), required=True)
    args = parser.parse_args()
    try:
        output_root = Path(__file__).resolve().parent
        output_binding_path = args.output_binding.resolve()
        if output_binding_path.parent != output_root:
            fail("output binding must live in the candidate bundle")
        if output_binding_path.exists() or output_binding_path.is_symlink():
            fail(f"refusing to overwrite {output_binding_path}")
        native_path = args.native_oracle.resolve()
        if Path(os.path.relpath(native_path, output_root)).as_posix() != "../native/verify-output.py":
            fail("native output oracle must be docs/performance/results/change-0457/native/verify-output.py")
        native = import_native(native_path)
        corpus = obj(load(output_root / "corpus-bindings.json"), "candidate corpus bindings")
        if corpus.get("schema") != "litchi-0457-odp-source-tail-candidate-corpus-binding-v1":
            fail("candidate corpus binding schema differs")
        control_shapes = obj(corpus.get("shapes"), "candidate corpus bindings.shapes")
        repo_root = args.repo_root.resolve()
        fixture_args = {}
        for shape, source_text, output_text in args.fixture:
            if shape not in SHAPES or shape in fixture_args:
                fail(f"fixtures must contain each shape exactly once: {shape}")
            source_path = Path(source_text)
            output_path = Path(output_text)
            if not source_path.is_absolute():
                source_path = repo_root / source_path
            if not output_path.is_absolute():
                output_path = repo_root / output_path
            fixture_args[shape] = (source_path.resolve(), output_path.resolve())
        if set(fixture_args) != set(SHAPES):
            fail("fixtures must contain tiny, medium, and large")
        shapes = {}
        for shape in ("tiny", "medium", "large"):
            shapes[shape] = bind_shape(shape, *fixture_args[shape], output_root, native_path, native, obj(control_shapes[shape], f"candidate corpus bindings.shapes.{shape}"))
        binding = {
            "schema": SCHEMA,
            "purpose": "independently verified source/output ZIP/XML fixtures for the style-free source-tail append; control output is used only for semantic/order/text/count references",
            "binder": {
                "path": "bind-output.py",
                "sha256": sha(output_root / "bind-output.py"),
            },
            "native_oracle": {
                "path": "../native/verify-output.py",
                "sha256": sha(native_path),
            },
            "shapes": shapes,
        }
        output_binding_path.write_text(json.dumps(binding, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        binding_sha = sha(output_binding_path)
        if args.protocol is not None:
            update_protocol(args.protocol.resolve(), output_binding_path, binding_sha, native_path, output_root)
        print(json.dumps({"status": "pass", "output_binding": str(output_binding_path), "sha256": binding_sha}, sort_keys=True))
        return 0
    except (OSError, TypeError, KeyError, ValueError, BindError) as error:
        print(f"BIND OUTPUT INVALID: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
