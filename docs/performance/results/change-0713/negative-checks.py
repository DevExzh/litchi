#!/usr/bin/env python3
"""Exercise the 0713 analyzer's rejection paths without mutating evidence files.

Each negative case replaces one ``read`` result with a deep-copied, corrupted
value in memory.  The analyzer therefore sees the corruption, while the
retained report and receipt files remain untouched.  The unmodified analyzer
is also rerun against a temporary output and its bytes are compared with the
sealed analysis report.
"""

from __future__ import annotations

import copy
import contextlib
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import sys
import tempfile
from typing import Any, Callable


HERE = Path(__file__).resolve().parent
ANALYZER_PATH = HERE / "analyze.py"
ANALYSIS_PATH = HERE / "analysis.json"

spec = importlib.util.spec_from_file_location("analyzer0713_negative_checks", ANALYZER_PATH)
if spec is None or spec.loader is None:
    raise RuntimeError(f"cannot load analyzer: {ANALYZER_PATH}")
ANALYZER = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ANALYZER)


def digest_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def digest_path(path: Path) -> str:
    return digest_bytes(path.read_bytes())


def run_analyzer(
    *,
    output: Path,
    target: Path | None = None,
    mutate: Callable[[Any], None] | None = None,
) -> tuple[int, str, bool, bytes | None]:
    """Run ``ANALYZER.main`` with an optional in-memory read mutation."""

    original_read = ANALYZER.read
    original_argv = sys.argv

    def patched_read(path: Path) -> Any:
        value = original_read(path)
        if target is not None and mutate is not None and Path(path).resolve() == target.resolve():
            value = copy.deepcopy(value)
            mutate(value)
        return value

    if target is not None:
        ANALYZER.read = patched_read
    stdout = io.StringIO()
    stderr = io.StringIO()
    try:
        sys.argv = [str(ANALYZER_PATH), "--output", str(output)]
        with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
            return_code = ANALYZER.main()
    finally:
        sys.argv = original_argv
        ANALYZER.read = original_read
    transcript = stdout.getvalue() + stderr.getvalue()
    output_bytes = output.read_bytes() if output.is_file() else None
    return return_code, transcript, output_bytes is not None, output_bytes


def expect_rejection(
    name: str,
    target: Path,
    mutate: Callable[[Any], None],
    reference_bytes: bytes,
) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="litchi-0713-negative-") as directory:
        output = Path(directory) / "unexpected-analysis.json"
        return_code, transcript, output_present, _ = run_analyzer(
            output=output,
            target=target,
            mutate=mutate,
        )
    if return_code == 0:
        raise AssertionError(f"{name}: analyzer accepted the in-memory corruption")
    if output_present:
        raise AssertionError(f"{name}: analyzer retained output after rejection")
    return {
        "name": name,
        "rejected": True,
        "return_code": return_code,
        "message": transcript.strip(),
        "target": str(target.relative_to(HERE)),
        "retained_target_sha256": digest_path(target),
        "reference_analysis_unchanged": digest_bytes(ANALYSIS_PATH.read_bytes())
        == digest_bytes(reference_bytes),
    }


def main() -> int:
    reference_bytes = ANALYSIS_PATH.read_bytes()
    protected_paths = [
        ANALYSIS_PATH,
        HERE / "baseline-A1-native-generated-edit.json",
        HERE / "baseline-A1-native-numbered-list-edit.json",
        HERE / "candidate-B1-native-numbered-list-edit.json",
        HERE / "baseline-A1-native-numbered-list-edit.receipt.json",
    ]
    before_digests = {str(path): digest_path(path) for path in protected_paths}

    with tempfile.TemporaryDirectory(prefix="litchi-0713-positive-") as directory:
        output = Path(directory) / "analysis.json"
        return_code, transcript, output_present, rerun_bytes = run_analyzer(output=output)
    if return_code != 0:
        raise AssertionError(f"unmodified analyzer rerun failed: {transcript.strip()}")
    if not output_present or rerun_bytes != reference_bytes:
        raise AssertionError("unmodified analyzer rerun was not byte-identical")

    checks = [
        expect_rejection(
            "wrong-raw-elapsed-statistics",
            HERE / "baseline-A1-native-generated-edit.json",
            lambda value: value["results"][0]["elapsed_ns"].__setitem__(
                "mean", value["results"][0]["elapsed_ns"]["mean"] + 1.0
            ),
            reference_bytes,
        ),
        expect_rejection(
            "wrong-receipt-command-argv",
            HERE / "baseline-A1-native-numbered-list-edit.receipt.json",
            lambda value: value["command"].__setitem__(0, "tampered-taskset"),
            reference_bytes,
        ),
        expect_rejection(
            "wrong-decoded-manifest-quantity",
            HERE / "baseline-A1-native-numbered-list-edit.json",
            lambda value: value["results"][0]["corpus"].__setitem__(
                "target_payload_bytes", 1242
            ),
            reference_bytes,
        ),
        expect_rejection(
            "wrong-deterministic-parity-output",
            HERE / "candidate-B1-native-numbered-list-edit.json",
            lambda value: value["results"][0]["source"]["ordinary_save"].__setitem__(
                "atomic_publication_steps", "in-memory parity corruption"
            ),
            reference_bytes,
        ),
    ]

    after_digests = {str(path): digest_path(path) for path in protected_paths}
    if after_digests != before_digests:
        raise AssertionError("a retained evidence file changed during negative checks")

    output = {
        "schema_version": 1,
        "status": "pass",
        "analyzer_sha256": digest_path(ANALYZER_PATH),
        "analysis_sha256": digest_bytes(reference_bytes),
        "original_rerun": {
            "return_code": 0,
            "byte_equal_to_analysis": True,
            "temporary_output_cleaned": True,
        },
        "checks": checks,
        "retained_inputs_unchanged": True,
    }
    (HERE / "negative-checks.json").write_text(json.dumps(output, indent=2) + "\n")
    print(f"PASS {len(checks)} in-memory analyzer rejection checks; original rerun byte-identical")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
