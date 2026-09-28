"""Root-only owner-scoped Callgrind publications for large capture.

The resulting Ir counts are attribution evidence only. This driver never
interprets them as latency, RSS, phase fractions, or speedup.
"""

import subprocess
import time

import custody as c


P = c.P
plan = c.read(P / "plan.json")
assert plan["profile"]["owner"] == "namespace_uri_probe::capture_region_0793"
assert plan["profile"]["shape"] == "large"
assert plan["profile"]["samples"] == 1
assert plan["profile"]["warmup"] == 0
assert not (P / "profiles").exists()
out = P / "profiles"
out.mkdir()

builds = {
    leg: c.read(P / f"build-{leg}/build.json")
    for leg in ("before", "after")
}
source = c.source()
expected_sources = {
    leg: c.read(builds[leg]["source"]["path"])
    for leg in ("before", "after")
}
assert c.changed_files(
    expected_sources["before"], expected_sources["after"]
) == set(plan["source_allowlist"])
assert source == expected_sources["after"]
for leg, build in builds.items():
    assert c.read(build["source"]["path"]) == expected_sources[leg]
    binary = build["binaries"]["profile"]
    assert c.artifact(binary["path"]) == binary

owner = plan["profile"]["owner"]
rows = []
for repeat, order in enumerate(plan["profile"]["orders"]):
    for leg in order:
        stem = f"{repeat}-{leg}"
        report = out / f"{stem}.json"
        raw = out / f"{stem}.callgrind"
        log = out / f"{stem}.log"
        binary = builds[leg]["binaries"]["profile"]
        command = [
            "taskset",
            "-c",
            str(plan["cpu"]),
            "valgrind",
            "--tool=callgrind",
            "--collect-atstart=no",
            "--toggle-collect=" + owner,
            "--zero-before=" + owner,
            "--dump-after=" + owner,
            "--callgrind-out-file=" + str(raw),
            binary["path"],
            "--mode",
            "capture",
            "--shape",
            "large",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--output",
            str(report),
        ]
        started = time.time()
        with log.open("w") as stream:
            result = subprocess.run(
                command,
                cwd=c.ROOT,
                stdout=stream,
                stderr=subprocess.STDOUT,
            )
        row = {
            "schema": "litchi.performance.0808.callgrind-receipt.v1",
            "repeat": repeat,
            "leg": leg,
            "command": command,
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "binary": binary,
            "driver_sha256": c.sha(P / "profile.py"),
            "artifacts": {
                path.name: c.artifact(path)
                for path in out.glob(stem + ".*")
                if path.is_file()
            },
        }
        rows.append(row)
        c.write(out / "receipts.json", rows)
        assert result.returncode == 0, log
        assert c.source() == source
        numbered = out / f"{stem}.callgrind.1"
        assert numbered.is_file(), "owner dump did not produce one numbered publication"
        assert not (out / f"{stem}.callgrind.2").exists(), "unexpected second publication"
        lines = numbered.read_text().splitlines()
        summaries = [
            int(line.split(":", 1)[1].strip())
            for line in lines
            if line.startswith("summary:")
        ]
        assert len(summaries) == 1 and summaries[0] > 1000
        assert any(owner in line for line in lines), "exact owner missing from publication"
        assert any("capture_internal" in line for line in lines)
        print(stem, "Callgrind PASS", flush=True)

c.write(
    out / "complete.json",
    {
        "schema": "litchi.performance.0808.callgrind.complete.v1",
        "processes": len(rows),
        "reports": len(rows),
        "samples": len(rows),
        "plan_sha256": c.sha(P / "plan.json"),
        "build_before_sha256": c.sha(P / "build-before/build.json"),
        "build_after_sha256": c.sha(P / "build-after/build.json"),
        "receipts": c.artifact(out / "receipts.json"),
        "scope": "namespace_uri_probe::capture_region_0793 only; Ir attribution without latency or RSS claim",
    },
)
