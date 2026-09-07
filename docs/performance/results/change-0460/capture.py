#!/usr/bin/env python3
"""Run the frozen ordinary ODP phase matrix serially from retained binaries."""
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def write(path, value):
    with path.open("x") as stream:
        json.dump(value, stream, indent=2)
        stream.write("\n")


def artifact(path):
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    variant, repeat = sys.argv[1:]
    assert variant in ["baseline", "candidate"] and repeat in ["R1", "R2"]
    protocol_path = ROOT / "protocol.json"
    protocol = json.loads(protocol_path.read_text())
    assert protocol["capture_sha256"] == sha(Path(__file__))
    assert protocol["oracle_sha256"] == sha(ROOT / "oracle.py")
    binding_path = ROOT / (variant + "-binding.json")
    if variant == "baseline":
        assert protocol["baseline_binding_sha256"] == sha(binding_path)
    binding = json.loads(binding_path.read_text())
    for value in binding["binaries"].values():
        path = Path(value["path"])
        assert path.is_file() and not path.is_symlink()
        assert sha(path) == value["sha256"] and path.stat().st_size == value["bytes"]
    spec = importlib.util.spec_from_file_location("oracle460", ROOT / "oracle.py")
    oracle = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(oracle)
    runs = ROOT / "runs" / variant
    runs.mkdir(parents=True, exist_ok=True)
    for lane in protocol["order"]:
        if lane["repeat"] != repeat:
            continue
        directory = runs / lane["id"]
        directory.mkdir()
        report = directory / "report.json"
        substitutions = dict(lane, binary=binding["binaries"][lane["instrumentation"]]["path"], report=str(report))
        argv = [part.format_map(substitutions) for part in protocol["argv"]]
        command = ["/usr/bin/time", "-v", "-o", str(directory / "resource.log"), "taskset", "-c", str(protocol["cpu"]), *argv]
        environment = os.environ | protocol["environment"]
        observed_allocator = {key: value for key, value in environment.items() if key.startswith("MALLOC_") or key == "GLIBC_TUNABLES"}
        record = {"schema": "litchi-0460-capture-v1", "change": 460, "variant": variant, "lane": lane, "argv": command,
                  "cwd": str(REPO), "protocol_sha256": sha(protocol_path), "binary_binding_sha256": sha(binding_path),
                  "environment_fixed": protocol["environment"], "allocator_environment": observed_allocator,
                  "started_utc": now()}
        with (directory / "stdout.log").open("xb") as out, (directory / "stderr.log").open("xb") as err:
            result = subprocess.run(command, cwd=REPO, env=environment, stdout=out, stderr=err)
        record["exit_code"] = result.returncode
        try:
            assert result.returncode == 0, "workload exited unsuccessfully"
            record["oracle"] = oracle.validate(report, lane, protocol)
            binary = Path(substitutions["binary"])
            assert sha(binary) == binding["binaries"][lane["instrumentation"]]["sha256"]
            record["status"] = "pass"
        except Exception as error:
            record["status"] = "failed"
            record["error"] = str(error)
        record["finished_utc"] = now()
        record["artifacts"] = [artifact(path) for path in sorted(directory.iterdir()) if path.is_file()]
        write(directory / "receipt.json", record)
        print(lane["id"], record["status"], flush=True)
        if record["status"] != "pass":
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
