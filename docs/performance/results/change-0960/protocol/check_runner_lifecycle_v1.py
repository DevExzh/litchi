#!/usr/bin/env python3
"""Offline lifecycle control for the 0960 runner v7 cleanup identity."""
from __future__ import annotations
import importlib.util
import json
import os
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0960")
RUNNER = ROOT / "tools" / "screen_0960_v7.py"
OUTPUT = ROOT / "review" / "runner-lifecycle-controls-v1.json"

spec = importlib.util.spec_from_file_location("screen_0960_v7_lifecycle", RUNNER)
assert spec is not None and spec.loader is not None
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

class AlreadyReapedParent:
    pid = 44000
    returncode = 0
    def wait(self, timeout=None):
        return self.returncode

old_getpgid = module.os.getpgid
old_killpg = module.os.killpg
live = {"value": True}
calls: list[tuple[int, int]] = []

def missing_parent(pid):
    raise ProcessLookupError(pid)

def group_call(pgid, signum):
    calls.append((pgid, int(signum)))
    if signum == 0:
        if live["value"]:
            return None
        raise ProcessLookupError(pgid)
    live["value"] = False

module.os.getpgid = missing_parent
module.os.killpg = group_call
try:
    result = module.cleanup(AlreadyReapedParent(), None)
finally:
    module.os.getpgid = old_getpgid
    module.os.killpg = old_killpg
assert result["reaped"] is True
assert result["process_group_clear"] is True
assert result["actions"] == ["SIGTERM"]
assert calls[0] == (AlreadyReapedParent.pid, 0)
assert (AlreadyReapedParent.pid, int(module.signal.SIGTERM)) in calls
encoded = (json.dumps({
    "schema": "litchi.performance.0960.runner-lifecycle-controls.v1",
    "status": "passed",
    "scope": "offline mocked cleanup only; parent is already reaped while a descendant remains in its start_new_session group",
    "runner": {"path": str(RUNNER.resolve()), "sha256": __import__("hashlib").sha256(RUNNER.read_bytes()).hexdigest()},
    "checks": [{"name": "reaped_parent_live_group_uses_pid_as_pgid", "status": "pass", "actions": result["actions"], "calls": calls, "process_group_clear": result["process_group_clear"]}],
}, indent=2, sort_keys=True) + "\n").encode()
fd = os.open(OUTPUT, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o664)
with os.fdopen(fd, "wb") as stream:
    stream.write(encoded)
    stream.flush()
    os.fsync(stream.fileno())
print(json.dumps({"status": "passed", "output": str(OUTPUT), "bytes": len(encoded)}, sort_keys=True))
