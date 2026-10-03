#!/usr/bin/env python3
"""Run the frozen whole-process cycle profile, preserving every attempt."""
import datetime
import hashlib
import json
from pathlib import Path
import subprocess
import sys

P = Path(__file__).resolve().parent
sys.path.insert(0, str(P.parent / "change-0885"))
import verify

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def write(path, value):
    with path.open("x") as f:
        json.dump(value, f, indent=2)
        f.write("\n")

plan = json.loads((P / "plan.json").read_text())
binary = Path(plan["binary"]["path"])
if sha(binary) != plan["binary"]["sha256"]:
    raise ValueError("retained executable changed")
for index in range(plan["children"]):
    folder = P / f"child-{index:02}"
    folder.mkdir()
    argv = ["/usr/bin/perf", "record", "--no-buildid-cache", "-e", plan["event"],
            "-F", str(plan["frequency_hz"]), "--call-graph", plan["call_graph"],
            "-o", str(folder / "perf.data"), "--", "/usr/bin/taskset", "-c", "12",
            str(binary), "--case", plan["case"], "--warmup", str(plan["warmup"]),
            "--samples", str(plan["samples"]), "--filesystem-cache", "warm",
            "--json", str(folder / "report.json"), "--producer-evidence", str(folder / "producer.json")]
    env = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}
    receipt = {"argv": argv, "environment": env, "cwd": str(P),
               "plan_sha256": sha(P / "plan.json"), "capture_helper_sha256": sha(Path(__file__)),
               "started_utc": datetime.datetime.now(datetime.timezone.utc).isoformat()}
    write(folder / "command-start.json", receipt)
    with (folder / "stdout.txt").open("xb") as out, (folder / "stderr.txt").open("xb") as err:
        code = subprocess.run(argv, env=env, cwd=P, stdout=out, stderr=err).returncode
    receipt.update(exit_code=code, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    receipt["artifacts"] = {f.name: {"bytes": f.stat().st_size, "sha256": sha(f)}
                            for f in folder.iterdir() if f.is_file()}
    write(folder / "command.json", receipt)
    if code:
        raise SystemExit(code)
    verify._validate_report(json.loads((folder / "report.json").read_text()),
                            json.loads((folder / "producer.json").read_text()),
                            "native", {"samples": plan["samples"], "warmup": plan["warmup"]})
    print(f"profile {index} complete; semantic replay passed", flush=True)
if sha(binary) != plan["binary"]["sha256"]:
    raise ValueError("retained executable changed during capture")
