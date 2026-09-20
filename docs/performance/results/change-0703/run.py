#!/usr/bin/env python3
"""Run two fresh-process repeats of six diagnostic workflows; no timers."""
import hashlib
import json
import os
import subprocess
from pathlib import Path
P = Path(__file__).resolve().parent
ROOT = P.parents[3]
REAL = "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    build = json.loads((P / "build.json").read_text())
    assert build["status"] == "completed"
    binary = Path(build["binary"])
    assert sha(binary) == build["binary_sha256"]
    assert sha(ROOT / "crates/litchi-ooxml-common/src/mce/codec.rs") == build["original_codec_sha256"]
    for name, digest in build["probe_sha256"].items():
        assert sha(P / name) == digest, name
    outdir = P / "trace-runs"
    outdir.mkdir(exist_ok=True)
    rows = []
    for repeat in range(2):
        for case, source in [("real", REAL), ("generated", "generated:12x8")]:
            for workflow in ["noop", "one", "two"]:
                name = f"{case}-{workflow}-r{repeat}"
                phase = outdir / (name + ".phase")
                stdout = outdir / (name + ".stdout")
                stderr = outdir / (name + ".stderr")
                assert not stdout.exists() and not stderr.exists(), "refuse to overwrite trace"
                phase.write_text("unset")
                command = [str(binary), source, workflow, "1", str(phase)]
                try:
                    with stdout.open("w") as out, stderr.open("w") as err:
                        result = subprocess.run(command, cwd=ROOT,
                            env=os.environ | {"LITCHI_0703_PHASE_FILE": str(phase)},
                            stdout=out, stderr=err)
                finally:
                    phase.unlink(missing_ok=True)
                rows.append(dict(name=name, case=case, source=source, workflow=workflow,
                    repeat=repeat, command=command, exit_code=result.returncode,
                    binary_sha256=sha(binary), source_archive_sha256=sha(ROOT / source) if case=="real" else None,
                    stdout=str(stdout.relative_to(P)), stderr=str(stderr.relative_to(P)),
                    stdout_sha256=sha(stdout), stderr_sha256=sha(stderr)))
                (P / "runs.json").write_text(json.dumps(rows, indent=2) + "\n")
                print(name, result.returncode, flush=True)
                if result.returncode:
                    raise RuntimeError(stderr.read_text())
    command = ["python3", str(P / "analyze_trace_0703.py")]
    command += [str(P / row["stderr"]) for row in rows]
    command += ["--json-out", str(P / "trace-summary.json")]
    with (P / "analyze.log").open("w") as out:
        subprocess.run(command, cwd=ROOT, stdout=out, stderr=subprocess.STDOUT, check=True)

if __name__ == "__main__":
    main()
