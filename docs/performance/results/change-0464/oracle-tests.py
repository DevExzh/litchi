#!/usr/bin/env python3
"""Small negative tests for the independent 0464 PPTX pair oracle.

Every mutation is written below a temporary directory.  The retained source,
destination, and published artifacts are never modified.
"""
from __future__ import annotations

import argparse
import importlib.util
import posixpath
import sys
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

HERE = Path(__file__).resolve().parent
ORACLE_PATH = HERE / "pair-oracle.py"
DEFAULT_SOURCE = HERE / "inputs/source.pptx"
DEFAULT_OUTPUT = HERE / "inputs/destination.pptx"


def load_oracle():
    spec = importlib.util.spec_from_file_location("litchi_0464_pair_oracle", ORACLE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load pair-oracle.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def rewrite_zip(source: Path, target: Path, changes: dict[str, bytes], extra: tuple[str, bytes] | None = None) -> None:
    with ZipFile(source) as archive, ZipFile(target, "w", compression=ZIP_DEFLATED) as output:
        for info in archive.infolist():
            payload = changes.get(info.filename, archive.read(info.filename))
            output.writestr(info.filename, payload)
        if extra is not None:
            output.writestr(*extra)


def expect_failure(oracle, label: str, diagnostic: str, **kwargs) -> str:
    try:
        oracle.verify_pair(**kwargs)
    except oracle.VerificationError as error:
        if diagnostic not in str(error):
            raise AssertionError(
                f"{label}: expected diagnostic containing {diagnostic!r}, got {error!r}"
            )
        return f"{label}: rejected ({error})"
    raise AssertionError(f"{label}: tampered pair was accepted")


def run(source: Path, destination: Path, output: Path, source_index: int, insertion_index: int) -> list[str]:
    oracle = load_oracle()
    base = {
        "source_path": source,
        "destination_path": destination,
        "output_path": output,
        "source_index": source_index,
        "insertion_index": insertion_index,
        "same_input_destination": source.resolve() == destination.resolve(),
    }
    positive = oracle.verify_pair(**base)
    if positive.get("status") != "pass":
        raise AssertionError("base pair did not pass")
    results = ["base positive: pass"]
    wrong_insertion = insertion_index - 1 if insertion_index else insertion_index + 1
    results.append(expect_failure(
        oracle, "wrong insertion index", "slide order differs",
        **{**base, "insertion_index": wrong_insertion},
    ))
    with tempfile.TemporaryDirectory(prefix="litchi-0464-oracle-tests-") as directory:
        temp = Path(directory)
        if source.read_bytes() == destination.read_bytes():
            distinct_destination = temp / "destination-copy.pptx"
            distinct_destination.write_bytes(source.read_bytes())
            distinct_result = oracle.verify_pair(
                source, distinct_destination, output, source_index, insertion_index,
                same_input_destination=False,
            )
            if distinct_result["input_identity_fact"] != "same_bytes_distinct_paths":
                raise AssertionError("distinct archive identity was not reported explicitly")
            results.append("same-bytes distinct paths: accepted without an independent-producer claim")
        else:
            results.append("formal distinct inputs: same-byte alias case not applicable")
        destination_package = oracle.Package(
            oracle._snapshot(destination.read_bytes(), "test destination"), "test destination"
        )
        output_package = oracle.Package(
            oracle._snapshot(output.read_bytes(), "test output"), "test output"
        )
        destination_names = set(destination_package.raw.order)
        added = set(output_package.raw.order) - destination_names
        copied_slide_candidates = [name for name in output_package.slides if name not in destination_names]
        if len(copied_slide_candidates) != 1:
            raise AssertionError(f"expected one copied slide, got {copied_slide_candidates}")
        copied_slide = copied_slide_candidates[0]
        copied_rels = oracle._rels_path(copied_slide)
        if copied_rels not in added:
            raise AssertionError("copied slide relationship sidecar is missing")
        payload_candidates = [
            name for name in added
            if name != copied_slide and not name.endswith(".rels")
            and not output_package.content_type(name).endswith("+xml")
        ]
        copied_payload = payload_candidates[0] if payload_candidates else copied_slide
        original_payload = output_package.raw.payloads[copied_payload]
        original_rels = output_package.raw.payloads[copied_rels]
        changed_payload = original_payload[:-1] + bytes([original_payload[-1] ^ 1])
        payload_path = temp / "payload-tampered.pptx"
        rewrite_zip(output, payload_path, {copied_payload: changed_payload})
        payload_diagnostic = "copied payload differs" if copied_payload != copied_slide else "expected one copied selected-slide addition"
        results.append(expect_failure(
            oracle, "copied payload mutation", payload_diagnostic,
            **{**base, "output_path": payload_path},
        ))
        wrong_target = None
        for relationship in output_package.relationships(copied_slide):
            if relationship.target_mode is not None:
                continue
            target = output_package.target(copied_slide, relationship)
            if target not in added:
                continue
            destination_target = next(
                name for name in destination_package.raw.order
                if name not in {destination_package.presentation, destination_package.presentation_rels, "[Content_Types].xml"}
            )
            replacement = posixpath.relpath(destination_target, posixpath.dirname(copied_slide))
            needle = f'Target="{relationship.target}"'.encode()
            replacement_bytes = f'Target="{replacement}"'.encode()
            if needle in original_rels:
                wrong_target = original_rels.replace(needle, replacement_bytes, 1)
                break
        if wrong_target is None:
            raise AssertionError("copied relationship has no internal copied target")
        rels_path = temp / "relationship-tampered.pptx"
        rewrite_zip(output, rels_path, {copied_rels: wrong_target})
        results.append(expect_failure(
            oracle, "copied relationship retargeting", "copied relationship target is not a new member",
            **{**base, "output_path": rels_path},
        ))
        extra_path = temp / "extra-member.pptx"
        rewrite_zip(output, extra_path, {}, extra=("ppt/unexpected.bin", b"unexpected"))
        results.append(expect_failure(
            oracle, "unexpected output member", "no content type for",
            **{**base, "output_path": extra_path},
        ))
    return results


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--destination", type=Path)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--source-index", type=int, default=0)
    parser.add_argument("--insertion-index", type=int, default=1)
    args = parser.parse_args(argv)
    destination = args.destination or args.source
    try:
        results = run(args.source, destination, args.output, args.source_index, args.insertion_index)
    except (AssertionError, OSError, ValueError) as error:
        print(f"oracle negative tests failed: {error}", file=sys.stderr)
        return 1
    print("\n".join(results))
    print(f"negative_cases={sum(1 for result in results if ': rejected (' in result)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
