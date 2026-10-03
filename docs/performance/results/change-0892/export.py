#!/usr/bin/env python3
"""Export auditable native sampled-IP views from retained perf data.

The exports are text views of already captured ``perf.data`` files.  No
workload or build is started.  Each command and output is create-only so a
failed export remains visible for review.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import subprocess

import verify


P = Path(__file__).resolve().parent
COMMANDS = {
    "leaf-samples.txt": [
        "script", "-G", "-F", "pid,tid,time,period,event,ip,sym,dso",
    ],
    "call-stacks.txt": [
        "script", "-F", "pid,tid,time,period,event,ip,sym,dso", "--inline",
    ],
    "self-symbols.txt": [
        "report", "--stdio", "--no-children", "--call-graph", "none",
        "--sort", "dso,symbol", "--show-nr-samples", "--show-total-period",
        "--percent-limit", "0", "--stdio-color", "never", "--field-separator", "|",
    ],
    "header.txt": ["report", "--header-only", "--stdio"],
    "buildids.txt": ["buildid-list"],
}


def export(capture_root: Path) -> dict[str, object]:
    folders = sorted(path for path in capture_root.glob("child-*") if path.is_dir())
    if not folders:
        raise ValueError(f"no native child folders under {capture_root}")
    rows: list[dict[str, object]] = []
    for folder in folders:
        data = folder / "perf.data"
        if not data.is_file():
            raise FileNotFoundError(data)
        if (folder / "exports.json").exists():
            raise FileExistsError(f"export receipt already exists; preserve prior attempt: {folder / 'exports.json'}")
        child_rows: list[dict[str, object]] = []
        for name, arguments in COMMANDS.items():
            output = folder / name
            stderr = folder / f"{name}.stderr"
            if output.exists() or stderr.exists():
                raise FileExistsError(f"export output already exists; preserve prior attempt: {output}")
            argv = ["/usr/bin/perf", *arguments, "-i", str(data)]
            with output.open("x", encoding="utf-8") as out, stderr.open("x", encoding="utf-8") as err:
                code = subprocess.run(argv, cwd=P, env=verify.CHILD_ENV, stdout=out, stderr=err, check=False).returncode
            row = {
                "argv": argv,
                "exit_code": code,
                "output": name,
                "stderr": stderr.name,
                "output_identity": verify.descriptor(output),
                "stderr_identity": verify.descriptor(stderr),
            }
            child_rows.append(row)
            if code != 0:
                raise RuntimeError(f"perf export failed for {folder.name}/{name}; attempt retained")
        receipt = {
            "schema": "litchi.xlsx-scan-cycle-profile-export-0892.v1",
            "status": "pass",
            "child": folder.name,
            "records": child_rows,
            "perf_data": verify.descriptor(data),
            "helper": verify.descriptor(Path(__file__)),
            "performance_claim": False,
        }
        verify.write_once(folder / "exports.json", receipt)
        rows.append({"child": folder.name, "receipt": str((folder / "exports.json").resolve())})
        print(folder.name + " exported", flush=True)
    result = {
        "schema": "litchi.xlsx-scan-cycle-profile-export-receipt-0892.v1",
        "status": "pass",
        "capture_root": str(capture_root.resolve()),
        "children": rows,
        "performance_claim": False,
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--capture-root", type=Path, default=P / "capture")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    result = export(args.capture_root.resolve())
    if args.output is not None:
        verify.write_once(args.output.resolve(), result)
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
