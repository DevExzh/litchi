"""Serial 0788 capture driver.

This module is intentionally root-only.  It owns the child schedule and the
parent side of the opt-in RSS phase protocol; it never retries a child and it
never overwrites an existing lane.  The build and capture receipts retain
absolute paths because the packet is later relocated and replayed by the
offline validators.
"""

from __future__ import annotations

import gzip
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Iterable

import custody as c


PACKET = c.P
ROOT = c.ROOT
PLAN_PATH = PACKET / "plan.json"
AFFINITY = ",".join(str(value) for value in c.read(PLAN_PATH)["affinity"])
SENTINEL = (1 << 64) - 1
PHASE_ACK = b"+\n"
PROC_FILES = ("smaps", "smaps_rollup", "maps", "status", "stat")
TIME_FORMAT = "%M %R %F"
TIME_PROGRAM = "/usr/bin/time"
TASKSET = "taskset"
HEAPTRACK = "/usr/bin/heaptrack"
HEAPTRACK_PRINT = "/usr/bin/heaptrack_print"
RECEIPT_SCHEMA = "litchi.cached-part-memory-receipts.0788.v1"


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def artifact_if(path: Path) -> dict[str, Any] | None:
    return c.artifact(path) if path.is_file() else None


def assert_new(path: Path) -> None:
    require(not path.exists(), f"capture would overwrite existing path: {path}")


def case_args(case: dict[str, Any], samples: int, warmup: int,
              report: Path) -> list[str]:
    args: list[str] = []
    for key in ("route", "shape", "state", "task_floor", "workers"):
        args.extend(["--" + key.replace("_", "-"), str(case[key])])
    args.extend(["--samples", str(samples), "--warmup", str(warmup),
                 "--output", str(report)])
    return args


def time_command(program: Iterable[str], time_file: Path) -> list[str]:
    """Run ``program`` under the requested affinity and outer GNU time."""
    return [TASKSET, "-c", AFFINITY, TIME_PROGRAM, "-f", TIME_FORMAT,
            "-o", str(time_file), *program]


def write_json(path: Path, value: Any) -> None:
    assert_new(path)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def parse_time_file(time_file: Path, rss_file: Path) -> None:
    """Retain the raw GNU-time line and expose a replay-friendly JSON receipt."""
    require(time_file.is_file(), f"GNU time did not create {time_file}")
    fields = time_file.read_text().strip().split()
    require(len(fields) == 3 and fields[0].isdigit() and fields[2].isdigit(),
            f"unexpected GNU time output in {time_file}: {fields!r}")
    write_json(rss_file, {
        "ru_maxrss_kib": int(fields[0]),
        "elapsed_seconds": float(fields[1]),
        "major_faults": int(fields[2]),
    })


def load_builds(legs: Iterable[str]) -> dict[str, dict[str, Any]]:
    builds: dict[str, dict[str, Any]] = {}
    for leg in legs:
        path = PACKET / "builds" / f"{leg}.json"
        require(path.is_file(), f"missing build receipt: {path}")
        build = c.read(path)
        require(isinstance(build, dict), f"malformed build receipt: {path}")
        binaries = build.get("binaries")
        require(isinstance(binaries, dict), f"build binaries missing: {path}")
        for binary in binaries.values():
            require(c.artifact(binary["path"]) == binary,
                    f"build binary changed: {path}")
        source = build.get("source")
        require(isinstance(source, dict), f"build source missing: {path}")
        source_path = Path(source["path"])
        require(c.artifact(source_path) == source,
                f"build source artifact changed: {path}")
        builds[leg] = build
    return builds


def frozen_source(builds: dict[str, dict[str, Any]], anchor: str) -> dict[str, Any]:
    """Bind the live worktree to this lane's historical build source.

    Paired lanes retain before and after binaries together, so their
    production source snapshots intentionally differ.  The selected anchor
    is the source that must still be live while the lane runs; tool and probe
    sources are common inputs and must agree across both builds.
    """
    require(anchor in builds, f"source anchor {anchor!r} is not built")
    expected = c.read(builds[anchor]["source"]["path"])
    require(isinstance(expected, dict), "build source snapshot is malformed")
    for leg, build in builds.items():
        source = c.read(build["source"]["path"])
        require(source.get("tool") == expected.get("tool"),
                f"build tool source differs for {leg}")
        require(source.get("probe") == expected.get("probe"),
                f"probe source differs for {leg}")
    live: dict[str, Any] = {
        "production": c.source(),
        "tool": c.tool_source(),
    }
    if "probe" in expected:
        live["probe"] = {
            path.name: c.sha(path)
            for path in sorted((PACKET / "probe-src").iterdir())
            if path.is_file()
        }
    require(live == expected,
            f"live source differs from {anchor} build source")
    c.unchanged(expected)
    return expected


def check_source(frozen: dict[str, Any]) -> None:
    c.unchanged(frozen)


def common_row(lane: str, case: dict[str, Any], *, repeat: int | None,
               leg: str | None, mode: str, samples: int, warmup: int,
               block: int | None = None, protocol: int | None = None,
               order: int | None = None) -> dict[str, Any]:
    row: dict[str, Any] = {
        "schema": RECEIPT_SCHEMA,
        "lane": lane,
        "case": dict(case),
        **case,
        "mode": mode,
        "samples": samples,
        "warmup": warmup,
    }
    if repeat is not None:
        row["repeat"] = repeat
    if leg is not None:
        row["leg"] = leg
    if block is not None:
        row["block"] = block
    if protocol is not None:
        row["protocol"] = protocol
    if order is not None:
        row["order"] = order
    return row


def add_artifact(row: dict[str, Any], key: str, path: Path) -> None:
    value = artifact_if(path)
    if value is not None:
        row[key] = value


def record_row(out: Path, rows: list[dict[str, Any]], stem: str,
               row: dict[str, Any]) -> None:
    receipt = out / f"{stem}.receipt.json"
    assert_new(receipt)
    write_json(receipt, row)
    rows.append(row)
    write_json_or_replace(out / "receipts.json", {
        "schema": RECEIPT_SCHEMA,
        "lane": row.get("lane"),
        "receipts": rows,
    })


def write_json_or_replace(path: Path, value: Any) -> None:
    # The aggregate receipt is deliberately checkpointed after every child.
    # The lane itself is new, so this is not evidence overwrite.
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def finish_lane(out: Path, rows: list[dict[str, Any]], frozen: dict[str, Any],
                lane: str) -> None:
    check_source(frozen)
    complete = {
        "schema": RECEIPT_SCHEMA,
        "lane": lane,
        "children": len(rows),
        "receipts": c.artifact(out / "receipts.json"),
        "source": c.artifact(out / "source.json"),
    }
    write_json(out / "complete.json", complete)


def standard_child(out: Path, lane: str, case: dict[str, Any],
                   binary: dict[str, Any], *, mode: str, samples: int,
                   warmup: int, repeat: int | None, leg: str | None,
                   block: int | None, protocol: int | None,
                   order: int | None, frozen: dict[str, Any],
                   rows: list[dict[str, Any]], stem: str) -> None:
    report = out / f"{stem}.json"
    time_file = out / f"{stem}.time"
    rss = out / f"{stem}.rss"
    log = out / f"{stem}.log"
    for path in (report, time_file, rss, log):
        assert_new(path)
    command = time_command(
        [binary["path"], *case_args(case, samples, warmup, report)], time_file)
    row = common_row(lane, case, repeat=repeat, leg=leg, mode=mode,
                     samples=samples, warmup=warmup, block=block,
                     protocol=protocol, order=order)
    row.update({"command": command, "binary": binary,
                "started": time.time()})
    result: subprocess.CompletedProcess[Any] | None = None
    error: str | None = None
    try:
        with log.open("wb") as stream:
            result = subprocess.run(command, cwd=ROOT, stdout=stream,
                                    stderr=subprocess.STDOUT)
    except OSError as exc:
        error = f"capture process failed to start: {exc}"
    row["ended"] = time.time()
    row["exit_code"] = None if result is None else result.returncode
    if error is not None:
        row["error"] = error
    add_artifact(row, "log", log)
    add_artifact(row, "time", time_file)
    if time_file.is_file():
        try:
            parse_time_file(time_file, rss)
        except Exception as exc:  # retain the raw failed artifact
            row["time_error"] = str(exc)
    add_artifact(row, "rss", rss)
    add_artifact(row, "report", report)
    record_row(out, rows, stem, row)
    require(result is not None and result.returncode == 0,
            f"{lane} child failed: {stem}: {error or result.returncode}")
    require(report.is_file(), f"{lane} child omitted report: {report}")
    report_value = c.read(report)
    require(isinstance(report_value.get("samples"), list)
            and len(report_value["samples"]) == samples,
            f"{lane} report sample count changed: {report}")
    require(rss.is_file(), f"{lane} child omitted RSS/time receipt: {rss}")
    check_source(frozen)
    print(stem, "passed", flush=True)


def read_stat_identity(text: str, expected_pid: int) -> tuple[int, int]:
    close = text.rfind(")")
    require(close > 0, "malformed /proc stat")
    try:
        pid = int(text[:text.index(" ")])
        tail = text[close + 2:].split()
        ppid = int(tail[1])
        starttime = int(tail[19])
    except (ValueError, IndexError) as exc:
        raise RuntimeError(f"malformed /proc stat identity: {exc}") from exc
    require(pid == expected_pid, f"/proc stat PID changed: {pid} != {expected_pid}")
    return ppid, starttime


def proc_snapshot(pid: int, wrapper_pid: int, binary_path: str,
                  sequence: int, phase: str, sample: int) -> dict[str, Any]:
    proc = Path("/proc") / str(pid)
    require(proc.is_dir(), f"benchmark PID disappeared: {pid}")
    exe = os.readlink(proc / "exe")
    require(os.path.realpath(exe) == os.path.realpath(binary_path),
            f"benchmark executable changed: {exe} != {binary_path}")
    raw = {name: (proc / name).read_text(encoding="utf-8", errors="replace")
           for name in PROC_FILES}
    ppid, starttime = read_stat_identity(raw["stat"], pid)
    require(ppid == wrapper_pid,
            f"benchmark parent is not /usr/bin/time: {ppid} != {wrapper_pid}")
    wrapper_exe = os.readlink(Path("/proc") / str(wrapper_pid) / "exe")
    require(os.path.realpath(wrapper_exe) == os.path.realpath(TIME_PROGRAM),
            f"benchmark parent executable is not time: {wrapper_exe}")
    task_ids = sorted(int(entry.name) for entry in (proc / "task").iterdir()
                      if entry.name.isdigit() and entry.is_dir())
    require(task_ids, f"benchmark task list is empty: {pid}")
    sample_value: int | str = (str(SENTINEL)
                                if sample == SENTINEL else sample)
    return {
        "sequence": sequence,
        "phase": phase,
        "sample": sample_value,
        "pid": pid,
        "ppid": ppid,
        "starttime": starttime,
        "exe": exe,
        "task_ids": task_ids,
        "raw": raw,
    }


def expected_protocol(samples: int) -> list[tuple[str, int]]:
    result = [("startup", SENTINEL), ("corpus_ready", SENTINEL),
              ("warmup_done", SENTINEL)]
    indexes = [0] if samples == 1 else [0, samples - 1]
    for sample in indexes:
        result.extend((phase, sample) for phase in (
            "package_ready", "after_preload", "after_operation",
            "after_batch_drop", "after_package_drop"))
    result.extend((phase, SENTINEL) for phase in (
        "samples_done", "report_written", "report_dropped"))
    return result


def parse_marker(line: bytes) -> tuple[str, int, int]:
    require(line.endswith(b"\n"), f"phase record was not newline terminated: {line!r}")
    try:
        fields = line[:-1].decode("ascii").split("\t")
        require(len(fields) == 4 and fields[0] == "RSS0788",
                f"invalid phase record: {line!r}")
        return fields[1], int(fields[2]), int(fields[3])
    except (UnicodeDecodeError, ValueError) as exc:
        raise RuntimeError(f"invalid phase record: {line!r}: {exc}") from exc


def write_snapshots(path: Path, snapshots: list[dict[str, Any]]) -> None:
    assert_new(path)
    encoded = json.dumps(snapshots, indent=2, sort_keys=True).encode() + b"\n"
    with path.open("wb") as stream:
        with gzip.GzipFile(fileobj=stream, mode="wb", mtime=0) as compressed:
            compressed.write(encoded)


def memory_child(out: Path, case: dict[str, Any], binary: dict[str, Any],
                 *, repeat: int, protocol: int, samples: int, warmup: int,
                 mode: str, leg: str, order: int, frozen: dict[str, Any],
                 rows: list[dict[str, Any]], stem: str) -> None:
    report = out / f"{stem}.json"
    time_file = out / f"{stem}.time"
    rss = out / f"{stem}.rss"
    log = out / f"{stem}.log"
    snapshots_file = out / f"{stem}.snapshots.json.gz"
    for path in (report, time_file, rss, log, snapshots_file):
        assert_new(path)
    command = time_command(
        [binary["path"], *case_args(case, samples, warmup, report)], time_file)
    row = common_row("memory", case, repeat=repeat, leg=leg, mode=mode,
                     samples=samples, warmup=warmup, protocol=protocol,
                     order=order)
    row.update({"command": command, "binary": binary,
                "started": time.time(), "phase_env": mode == "on"})
    environment = os.environ.copy()
    environment.pop("LITCHI_RSS_PHASES", None)
    if mode == "on":
        environment["LITCHI_RSS_PHASES"] = "1"
    snapshots: list[dict[str, Any]] = []
    transcript: list[bytes] = []
    result: subprocess.Popen[bytes] | None = None
    error: str | None = None
    log_stream = log.open("wb")
    try:
        result = subprocess.Popen(command, cwd=ROOT, env=environment,
                                  stdin=subprocess.PIPE if mode == "on" else subprocess.DEVNULL,
                                  stdout=subprocess.PIPE, stderr=log_stream)
        if mode == "off":
            stdout, _ = result.communicate()
            if stdout:
                transcript.append(stdout)
                error = "phase-disabled child wrote to stdout"
        else:
            require(result.stdout is not None and result.stdin is not None,
                    "phase-enabled child pipes were not created")
            wrapper_pid = result.pid
            expected = expected_protocol(samples)
            identity: tuple[int, int, int] | None = None
            for index, (wanted_phase, wanted_sample) in enumerate(expected):
                line = result.stdout.readline()
                if not line:
                    error = (f"phase protocol ended at {index}/{len(expected)}; "
                             f"wanted {wanted_phase}/{wanted_sample}")
                    break
                transcript.append(line)
                try:
                    phase, sample, pid = parse_marker(line)
                    require((phase, sample) == (wanted_phase, wanted_sample),
                            f"phase {index} changed: {(phase, sample)!r} != "
                            f"{(wanted_phase, wanted_sample)!r}")
                    snapshot = proc_snapshot(
                        pid, wrapper_pid, binary["path"], index, phase, sample)
                    current_identity = (snapshot["pid"], snapshot["starttime"],
                                        snapshot["ppid"])
                    if identity is None:
                        identity = current_identity
                    else:
                        require(current_identity == identity,
                                "benchmark PID/starttime/PPID changed during protocol")
                    snapshots.append(snapshot)
                    result.stdin.write(PHASE_ACK)
                    result.stdin.flush()
                except Exception as exc:
                    error = str(exc)
                    break
            if error is None:
                # The final acknowledgement lets the benchmark write and drop
                # its report.  Waiting before reading the remainder closes the
                # pipe without imposing a timeout or killing the child.
                result.stdin.close()
                result.wait()
                extra = result.stdout.read()
                if extra:
                    transcript.append(extra)
                    error = "phase-enabled child wrote an unexpected extra record"
            else:
                result.stdin.close()
                result.wait()
                remainder = result.stdout.read()
                if remainder:
                    transcript.append(remainder)
    except OSError as exc:
        error = f"memory capture process failed: {exc}"
        if result is not None and result.stdin is not None:
            try:
                result.stdin.close()
            except OSError:
                pass
        if result is not None:
            result.wait()
    finally:
        log_stream.close()
    if transcript:
        with log.open("ab") as stream:
            stream.write(b"\nPARENT_PHASE_TRANSCRIPT\n")
            for item in transcript:
                stream.write(item)
    row["ended"] = time.time()
    row["exit_code"] = None if result is None else result.returncode
    if error is not None:
        row["error"] = error
    add_artifact(row, "log", log)
    add_artifact(row, "time", time_file)
    if time_file.is_file():
        try:
            parse_time_file(time_file, rss)
        except Exception as exc:
            row["time_error"] = str(exc)
    add_artifact(row, "rss", rss)
    add_artifact(row, "report", report)
    if mode == "on":
        write_snapshots(snapshots_file, snapshots)
        add_artifact(row, "snapshots", snapshots_file)
    else:
        row["snapshots"] = None
    record_row(out, rows, stem, row)
    require(result is not None and result.returncode == 0,
            f"memory child failed: {stem}: {error or result.returncode}")
    require(error is None, f"memory protocol failed: {stem}: {error}")
    require(report.is_file(), f"memory child omitted report: {report}")
    report_value = c.read(report)
    require(isinstance(report_value.get("samples"), list)
            and len(report_value["samples"]) == samples,
            f"memory report sample count changed: {report}")
    if mode == "on":
        require(len(snapshots) == len(expected_protocol(samples)),
                f"memory snapshot count changed: {stem}")
    require(rss.is_file(), f"memory child omitted RSS/time receipt: {rss}")
    check_source(frozen)
    print(stem, "passed", flush=True)


def heaptrack_child(out: Path, case: dict[str, Any], binary: dict[str, Any],
                    *, repeat: int, leg: str, frozen: dict[str, Any],
                    rows: list[dict[str, Any]], stem: str, samples: int,
                    warmup: int) -> None:
    report = out / f"{stem}.json"
    time_file = out / f"{stem}.time"
    rss = out / f"{stem}.rss"
    log = out / f"{stem}.log"
    trace_pattern = out / f"{stem}.heaptrack.%p"
    summary = out / f"{stem}.heaptrack-print.log"
    flamegraph = out / f"{stem}.heaptrack-peak.stacks"
    for path in (report, time_file, rss, log, summary, flamegraph):
        assert_new(path)
    command = time_command([
        HEAPTRACK, "--record-only", "--output", str(trace_pattern),
        binary["path"], *case_args(case, samples, warmup, report),
    ], time_file)
    row = common_row("heaptrack", case, repeat=repeat, leg=leg,
                     mode="heaptrack", samples=samples, warmup=warmup)
    row.update({"command": command, "binary": binary,
                "started": time.time()})
    result: subprocess.CompletedProcess[Any] | None = None
    error: str | None = None
    try:
        with log.open("wb") as stream:
            result = subprocess.run(command, cwd=ROOT, stdout=stream,
                                    stderr=subprocess.STDOUT)
    except OSError as exc:
        error = f"heaptrack process failed to start: {exc}"
    row["ended"] = time.time()
    row["exit_code"] = None if result is None else result.returncode
    if error is not None:
        row["error"] = error
    trace_candidates = sorted(
        path for path in out.glob(f"{stem}.heaptrack.*")
        if path.is_file() and path.suffix in (".gz", ".zst"))
    if len(trace_candidates) == 1:
        row["trace"] = c.artifact(trace_candidates[0])
        row["raw_trace"] = row["trace"]
    elif trace_candidates:
        row["trace_error"] = f"expected one heaptrack trace, found {len(trace_candidates)}"
    add_artifact(row, "log", log)
    add_artifact(row, "time", time_file)
    if time_file.is_file():
        try:
            parse_time_file(time_file, rss)
        except Exception as exc:
            row["time_error"] = str(exc)
    add_artifact(row, "rss", rss)
    add_artifact(row, "report", report)
    print_result: subprocess.CompletedProcess[Any] | None = None
    if len(trace_candidates) == 1:
        print_command = [
            HEAPTRACK_PRINT, "-f", str(trace_candidates[0]), "-m", "0",
            "-p", "1", "-a", "1", "-T", "0", "-n", "12",
            "--flamegraph-cost-type", "peak", "-F", str(flamegraph),
        ]
        row["print_command"] = print_command
        try:
            with summary.open("wb") as stream:
                print_result = subprocess.run(
                    print_command, cwd=ROOT, stdout=stream,
                    stderr=subprocess.STDOUT)
        except OSError as exc:
            row["print_error"] = str(exc)
        row["print_exit_code"] = (None if print_result is None
                                   else print_result.returncode)
    else:
        row["print_exit_code"] = None
    add_artifact(row, "summary", summary)
    add_artifact(row, "flamegraph", flamegraph)
    record_row(out, rows, stem, row)
    require(result is not None and result.returncode == 0,
            f"heaptrack child failed: {stem}: {error or result.returncode}")
    require(len(trace_candidates) == 1,
            f"heaptrack trace missing or ambiguous: {stem}")
    require(print_result is not None and print_result.returncode == 0,
            f"heaptrack_print failed: {stem}")
    require(report.is_file(), f"heaptrack child omitted report: {report}")
    report_value = c.read(report)
    require(isinstance(report_value.get("samples"), list)
            and len(report_value["samples"]) == samples,
            f"heaptrack report sample count changed: {report}")
    check_source(frozen)
    print(stem, "passed", flush=True)


def qualification_or_native(lane: str, plan: dict[str, Any],
                            builds: dict[str, dict[str, Any]],
                            frozen: dict[str, Any]) -> None:
    qualification = lane.startswith("qualification-")
    spec_name = "qualification" if qualification else "native"
    spec = plan[spec_name]
    out = PACKET / lane
    assert_new(out)
    out.mkdir()
    write_json(out / "source.json", frozen)
    rows: list[dict[str, Any]] = []
    cases = spec["cases"]
    blocks = 1 if qualification else spec["blocks"]
    for block in range(blocks):
        order = [lane.split("-", 1)[1]] if qualification else spec["orders"][block]
        for case in cases:
            for leg in order:
                stem = (f"{block}-{case['shape']}-{case['state']}-"
                        f"{case['task_floor']}-{case['workers']}-{leg}")
                standard_child(
                    out, lane, case, builds[leg]["binaries"]["native"],
                    mode="qualification" if qualification else "native",
                    samples=spec["samples"], warmup=spec["warmup"],
                    repeat=None, leg=leg, block=block, protocol=None,
                    order=None, frozen=frozen, rows=rows, stem=stem)
    finish_lane(out, rows, frozen, lane)


def memory_lane(plan: dict[str, Any], builds: dict[str, dict[str, Any]],
                frozen: dict[str, Any]) -> None:
    spec = plan["memory"]
    out = PACKET / "memory"
    assert_new(out)
    out.mkdir()
    write_json(out / "source.json", frozen)
    rows: list[dict[str, Any]] = []
    for repeat, order_list in enumerate(spec["orders"]):
        for protocol, protocol_spec in enumerate(spec["protocols"]):
            samples, warmup = protocol_spec["samples"], protocol_spec["warmup"]
            for case_index, case in enumerate(spec["cases"]):
                for order, label in enumerate(order_list):
                    mode, leg = label.split("-", 1)
                    stem = (f"r{repeat}-p{protocol}-c{case_index}-o{order}-"
                            f"{mode}-{leg}-{case['shape']}-{case['state']}-"
                            f"{case['task_floor']}-{case['workers']}")
                    memory_child(
                        out, case, builds[leg]["binaries"]["memory"],
                        repeat=repeat, protocol=protocol, samples=samples,
                        warmup=warmup, mode=mode, leg=leg, order=order,
                        frozen=frozen, rows=rows, stem=stem)
    finish_lane(out, rows, frozen, "memory")


def heaptrack_lane(plan: dict[str, Any], builds: dict[str, dict[str, Any]],
                   frozen: dict[str, Any]) -> None:
    spec = plan["heaptrack"]
    out = PACKET / "heaptrack"
    assert_new(out)
    out.mkdir()
    write_json(out / "source.json", frozen)
    rows: list[dict[str, Any]] = []
    for repeat, order_list in enumerate(spec["orders"]):
        for case_index, case in enumerate(spec["cases"]):
            for order, leg in enumerate(order_list):
                stem = (f"r{repeat}-c{case_index}-o{order}-{leg}-"
                        f"{case['shape']}-{case['state']}-{case['task_floor']}-"
                        f"{case['workers']}")
                heaptrack_child(
                    out, case, builds[leg]["binaries"]["native"],
                    repeat=repeat, leg=leg, frozen=frozen, rows=rows,
                    stem=stem, samples=spec["samples"], warmup=spec["warmup"])
    finish_lane(out, rows, frozen, "heaptrack")


def main() -> None:
    require(len(sys.argv) == 2,
            "usage: capture.py qualification-before|qualification-after|native|memory|heaptrack")
    lane = sys.argv[1]
    require(lane in ("qualification-before", "qualification-after", "native",
                     "memory", "heaptrack"), f"invalid capture lane: {lane}")
    plan = c.read(PLAN_PATH)
    if lane in ("qualification-before", "qualification-after"):
        leg = lane.split("-", 1)[1]
        builds = load_builds([leg])
    elif lane == "native" or lane == "heaptrack":
        builds = load_builds(["before", "after"])
    else:
        builds = load_builds(["before", "after"])
    anchor = (lane.split("-", 1)[1]
              if lane.startswith("qualification-") else "after")
    frozen = frozen_source(builds, anchor)
    if lane in ("qualification-before", "qualification-after", "native"):
        qualification_or_native(lane, plan, builds, frozen)
    elif lane == "memory":
        memory_lane(plan, builds, frozen)
    else:
        heaptrack_lane(plan, builds, frozen)


if __name__ == "__main__":
    main()
