#!/usr/bin/env python3
"""Separate runner-scoped guest-instruction diagnostic, not native cycles."""
import datetime
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

P = Path(__file__).resolve().parent
sys.path.insert(0, str(P.parent / "change-0885"))
import verify

plan = json.loads((P / "callgrind-plan.json").read_text())
binary = Path(plan["binary"]["path"])
if hashlib.sha256(binary.read_bytes()).hexdigest() != plan["binary"]["sha256"]:
    raise ValueError("binary mismatch")
for i in range(plan["children"]):
    folder = P / f"callgrind-{i:02}"
    folder.mkdir()
    argv = ["/usr/bin/taskset", "-c", "12", "/usr/bin/valgrind", "--tool=callgrind",
            "--collect-atstart=no", "--toggle-collect=" + plan["toggle_function"],
            "--callgrind-out-file=" + str(folder / "callgrind.out"), str(binary),
            "--case", "xlsx_producer_medium_source_full_cell_scan", "--warmup", "0",
            "--samples", str(plan["samples"]), "--filesystem-cache", "warm",
            "--json", str(folder / "report.json"), "--producer-evidence", str(folder / "producer.json")]
    env = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}
    receipt = {"argv": argv, "environment": env, "started_utc": datetime.datetime.now(datetime.timezone.utc).isoformat()}
    (folder / "command-start.json").write_text(json.dumps(receipt, indent=2) + "\n")
    with (folder / "stdout.txt").open("xb") as out, (folder / "stderr.txt").open("xb") as err:
        code = subprocess.run(argv, env=env, stdout=out, stderr=err).returncode
    receipt.update(exit_code=code, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    receipt["artifacts"] = {f.name: {"bytes": f.stat().st_size, "sha256": hashlib.sha256(f.read_bytes()).hexdigest()}
                            for f in folder.iterdir() if f.is_file()}
    (folder / "command.json").write_text(json.dumps(receipt, indent=2) + "\n")
    if code:
        raise SystemExit(code)
    verify._validate_report(json.loads((folder / "report.json").read_text()), json.loads((folder / "producer.json").read_text()),
                            "native", {"samples": plan["samples"], "warmup": 0})
    raw = (folder / "callgrind.out").read_text()
    totals = re.findall(r"^totals:\s+(\d+)", raw, re.M)
    if not totals or int(totals[-1]) == 0:
        raise ValueError("Callgrind collected no instructions; preserve failed attempt")
    print(folder.name + " complete; semantic replay passed", flush=True)
