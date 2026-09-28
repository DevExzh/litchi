"""Root-owned Python regression gates; retain each attempt and source snapshot."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def descriptor(path):
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def main():
    attempt = 0
    while (P / f"quality-{attempt}").exists():
        attempt += 1
    out = P / f"quality-{attempt}"
    out.mkdir()
    sources = []
    for source in (ROOT / "tools/perf_allocation_schema.py", ROOT / "tools/test_perf_allocation_schema.py",
                   P / "preflight.py", P / "fixtures.json", Path(__file__)):
        dest = out / source.name
        shutil.copy2(source, dest)
        sources.append({"original": descriptor(source), "snapshot": descriptor(dest)})
    receipt = {"schema": "litchi.performance.0826.quality.v1", "status": "running", "sources": sources, "rows": []}
    path = out / "receipt.json"
    def write():
        path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    write()
    for index, flags in enumerate(([], ["-O"])):
        command = ["python3", "-B", *flags, "-m", "unittest", "tools.test_perf_allocation_schema", "-v"]
        log = out / f"{index}.log"
        started = time.time()
        with log.open("x") as stream:
            result = subprocess.run(command, cwd=ROOT, env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1"},
                                    stdout=stream, stderr=subprocess.STDOUT)
        receipt["rows"].append({"command": command, "exit_code": result.returncode,
                                "started": started, "ended": time.time(), "log": descriptor(log)})
        receipt["status"] = "running" if result.returncode == 0 else "failed"
        write()
        if result.returncode:
            raise RuntimeError(f"regression gate failed; retained {path}")
        for source in sources:
            if descriptor(ROOT / source["original"]["path"]) != source["original"]:
                raise RuntimeError("source changed during regression gate")
        print(f"0826 Python gate {index + 1} PASS", flush=True)
    receipt["status"] = "pass"
    write()
    with (P / "quality.json").open("x") as stream:
        stream.write(json.dumps(receipt, indent=2, sort_keys=True) + "\n")


if __name__ == "__main__":
    main()
