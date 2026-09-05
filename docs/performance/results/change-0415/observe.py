#!/usr/bin/env python3
"""Capture separate writer and full-readback processes at three logical sizes."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--probe", type=Path, required=True)
    parser.add_argument("--scratch-prefix", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    probe = args.probe.resolve()
    interop = Path(__file__).resolve().with_name("interop.py")
    with probe.open("rb") as source:
        binary_sha256 = hashlib.file_digest(source, "sha256").hexdigest()
    runs = []
    for mode in ("borrowed", "owned"):
        for size in (67108864, 268435456, 4294967297):
            stem = f"{mode}-{size}"
            archive = Path(f"{args.scratch_prefix}-{stem}.zip").resolve()
            operations = (
                ("write", [str(probe), "write", mode, str(size), str(archive)]),
                ("readback", [str(probe), "verify", str(archive)]),
                ("python", [sys.executable, str(interop), "verify", str(archive)]),
            )
            for kind, command in operations:
                name = f"{stem}-{kind}"
                report = args.output / f"{name}.json"
                timing = (args.output / f"{name}.time.txt").resolve()
                argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(timing)] + command
                started = datetime.datetime.now(datetime.timezone.utc).isoformat()
                with report.open("w") as output:
                    result = subprocess.run(argv, stdout=output, stderr=subprocess.PIPE, text=True, check=False)
                runs.append({"name": name, "argv": argv, "started": started,
                             "finished": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                             "exit_code": result.returncode, "stderr": result.stderr})
                (args.output / "capture.json").write_text(json.dumps({"binary_sha256": binary_sha256, "runs": runs}, indent=2) + "\n")
                if result.returncode:
                    raise SystemExit(f"{name}: {result.stderr}")
                print(name, flush=True)


if __name__ == "__main__":
    main()
