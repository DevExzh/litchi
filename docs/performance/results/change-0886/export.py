#!/usr/bin/env python3
"""Export auditable sampled-IP and call-stack views from retained perf data."""
import hashlib
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
for folder in sorted(P.glob("child-*")):
    prefix = ["/usr/bin/perf"]
    source = ["-i", str(folder / "perf.data")]
    commands = {
        "leaf-samples.txt": prefix + ["script", *source, "-G", "-F", "pid,tid,time,period,event,ip,sym,dso"],
        "call-stacks.txt": prefix + ["script", *source, "-F", "pid,tid,time,period,event,ip,sym,dso", "--inline"],
        "self-symbols.txt": prefix + ["report", *source, "--stdio", "--no-children", "--call-graph", "none",
            "--sort", "dso,symbol", "--show-nr-samples", "--show-total-period", "--percent-limit", "0",
            "--stdio-color", "never", "--field-separator", "|"],
        "header.txt": prefix + ["report", *source, "--header-only", "--stdio"],
    }
    rows = []
    for name, argv in commands.items():
        with (folder / name).open("xb") as out, (folder / (name + ".stderr")).open("xb") as err:
            code = subprocess.run(argv, stdout=out, stderr=err).returncode
        rows.append({"argv": argv, "exit_code": code, "output": name,
                     "sha256": hashlib.sha256((folder / name).read_bytes()).hexdigest()})
        if code:
            raise SystemExit(code)
    with (folder / "exports.json").open("x") as out:
        json.dump(rows, out, indent=2)
        out.write("\n")
    print(folder.name + " exported", flush=True)
