#!/usr/bin/env python3
"""Run twelve fresh-process candidate MCE mechanism traces."""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
from pathlib import Path

P = Path(__file__).resolve().parent
PACKET = P.parent
ROOT = P.parents[4]
REAL = ROOT / "test-data" / "libreoffice-core" / "sd" / "qa" / "unit" / "data" / "pptx" / "slide-section-test.pptx"


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    build = json.loads((P / "build.json").read_text())
    if build.get("status") != "completed":
        raise AssertionError("mechanism build is not complete")
    binary = Path(build["binary"])
    if sha(binary) != build["binary_sha256"]:
        raise AssertionError("mechanism binary hash changed after build")
    if sha(ROOT / "crates" / "litchi-ooxml-common" / "src" / "mce" / "codec.rs") != build["original_codec_sha256"]:
        raise AssertionError("production codec is still instrumented or otherwise changed")
    for name, digest in build["probe_sha256"].items():
        if sha(P / name) != digest:
            raise AssertionError(f"standalone probe changed after build: {name}")
    if not REAL.is_file():
        raise AssertionError(f"missing real source deck: {REAL}")

    outdir = P / "trace-runs"
    outdir.mkdir(exist_ok=True)
    rows: list[dict[str, object]] = []
    for repeat in range(2):
        for case, source in (("real", REAL), ("generated", "generated:12x8")):
            for workflow in ("noop", "one", "two"):
                name = f"{case}-{workflow}-r{repeat}"
                phase = outdir / f"{name}.phase"
                stdout = outdir / f"{name}.stdout"
                stderr = outdir / f"{name}.stderr"
                if stdout.exists() or stderr.exists() or phase.exists():
                    raise AssertionError(f"refuse to overwrite mechanism trace: {name}")
                phase.write_text("unset")
                source_argument = str(source) if isinstance(source, Path) else source
                command = [str(binary), source_argument, workflow, "1", str(phase)]
                try:
                    with stdout.open("w") as out, stderr.open("w") as err:
                        result = subprocess.run(
                            command,
                            cwd=ROOT,
                            env=os.environ | {"LITCHI_0704_MECHANISM_PHASE_FILE": str(phase)},
                            stdout=out,
                            stderr=err,
                        )
                finally:
                    phase.unlink(missing_ok=True)
                row: dict[str, object] = {
                    "schema": "litchi-0704-mce-mechanism-run-v1",
                    "name": name,
                    "case": case,
                    "source": source_argument,
                    "workflow": workflow,
                    "repeat": repeat,
                    "command": command,
                    "exit_code": result.returncode,
                    "binary": str(binary),
                    "binary_sha256": sha(binary),
                    "candidate_source_map_sha256": build["candidate_source_map_sha256"],
                    "source_archive_sha256": sha(source) if isinstance(source, Path) else None,
                    "stdout": str(stdout.relative_to(P)),
                    "stderr": str(stderr.relative_to(P)),
                    "stdout_sha256": sha(stdout),
                    "stderr_sha256": sha(stderr),
                }
                rows.append(row)
                (P / "runs.json").write_text(json.dumps(rows, indent=2) + "\n")
                print(name, result.returncode, flush=True)
                if result.returncode:
                    raise RuntimeError(stderr.read_text())

    command = ["python3", str(P / "analyze_trace_0704.py")]
    command.extend(str(P / row["stderr"]) for row in rows)
    command.extend(("--json-out", str(P / "trace-summary.json")))
    with (P / "analyze.log").open("w") as output:
        subprocess.run(command, cwd=ROOT, stdout=output, stderr=subprocess.STDOUT, check=True)
    print("PASS: 12 fresh candidate mechanism traces analyzed", flush=True)


if __name__ == "__main__":
    main()
