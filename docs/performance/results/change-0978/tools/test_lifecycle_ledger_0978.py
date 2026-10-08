#!/usr/bin/env python3
"""Offline lifecycle/ledger controls for the 0978 adapter boundary."""
from __future__ import annotations

import sys
import struct
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import lifecycle_ledger_0978 as ledger  # noqa: E402


PID = 77
CPU = 31
PERIOD = 1_000_000
TARGET = "/tmp/doc_retention_probe"
BUILD_ID = "0123456789abcdef0123456789abcdef01234567"


def _raw(kind: int, size: int, *, pid: int = PID, tid: int = PID, comm: str | None = None, time_ns: int = 0, cpu: int = 31, misc: int = 0, build_id: str | None = None) -> bytes:
    raw = bytearray(size)
    struct.pack_into("<IHH", raw, 0, kind, misc, size)
    if kind in (3, 10):
        struct.pack_into("<II", raw, 8, pid, tid)
    if kind == 3 and comm is not None:
        raw[16 : 16 + len(comm) + 1] = comm.encode() + b"\0"
    if kind == 10:
        struct.pack_into("<QQQ", raw, 16, 0x400000, 0x2000, 0)
        if build_id is not None:
            value = bytes.fromhex(build_id)
            struct.pack_into("<I", raw, 40, len(value))
            raw[44 : 44 + len(value)] = value
        struct.pack_into("<II", raw, 64, 5, 2)
        raw[72 : 72 + len(TARGET) + 1] = TARGET.encode() + b"\0"
    if kind == 4:
        struct.pack_into("<IIII", raw, 8, pid, tid, pid, tid)
    if kind == 3:
        start = (raw.find(b"\0", 16) + 8) & ~7
    elif kind == 10:
        start = (raw.find(b"\0", 72) + 8) & ~7
    elif kind == 4:
        start = 32
    else:
        start = None
    if start is not None and start + 24 <= len(raw):
        struct.pack_into("<IIQII", raw, start, tid, pid, time_ns, cpu if cpu >= 0 else 0xFFFFFFFF, 0)
    return bytes(raw)


def _dump() -> Path:
    records = [
        (_raw(3, 56, comm="perf-exec", time_ns=0, cpu=0), "0 0 0x0 [0x38]: PERF_RECORD_COMM: perf-exec:77/77"),
        (_raw(9, 56), "31 105 0x38 [0x38]: PERF_RECORD_SAMPLE(IP, 0x2): 77/77: 0x400000 period: 1000000"),
        (_raw(3, 56, comm="python3.14", time_ns=110, cpu=31, misc=0x2002), "31 110 0x40 [0x38]: PERF_RECORD_COMM exec: python3.14:77/77"),
        (_raw(3, 56, comm="target", time_ns=115, cpu=31, misc=0x2002), "31 115 0x78 [0x38]: PERF_RECORD_COMM exec: target:77/77"),
        (_raw(10, 120, time_ns=120, cpu=31, misc=0x6002, build_id=BUILD_ID), f"31 120 0xb0 [0x78]: PERF_RECORD_MMAP2 77/77: [0x400000(0x2000) @ 0 <{BUILD_ID}>]: r-xp {TARGET}"),
        (_raw(9, 56), "31 125 0x128 [0x38]: PERF_RECORD_SAMPLE(IP, 0x2): 77/77: 0x400100 period: 1000000"),
        (_raw(9, 56), "31 135 0x160 [0x38]: PERF_RECORD_SAMPLE(IP, 0x2): 77/77: 0x400110 period: 1000000"),
        (_raw(4, 56, time_ns=160, cpu=31), "31 160 0x198 [0x38]: PERF_RECORD_EXIT(77:77):(77:77)"),
    ]
    lines: list[str] = []
    for raw, descriptor in records:
        lines.append(f". ... raw event: size {len(raw)} bytes")
        for offset in range(0, len(raw), 16):
            chunk = raw[offset : offset + 16]
            lines.append(f".  {offset:04x}:  " + " ".join(f"{byte:02x}" for byte in chunk))
        lines.append(descriptor)
    handle = tempfile.NamedTemporaryFile("w", suffix=".perf-mmap-dump", delete=False)
    path = Path(handle.name)
    with handle:
        handle.write("\n".join(lines) + "\n")
    return path


def _samples() -> list[dict]:
    return [
        # Valid 72-byte zero-callchain/zero-stack sample: unknown, retained.
        {"index": 1, "time_ns": 105, "ip": 0x400000, "raw_bytes": 56, "period": PERIOD, "pid": PID, "tid": PID, "cpu": CPU, "mode": "user", "callchain_nr": 0, "stack_size": 0},
        # A target-window libc/loader/other-DSO physical leaf is known physical
        # work but is not a target-source claim.
        {"index": 2, "time_ns": 125, "ip": 0x400100, "raw_bytes": 56, "period": PERIOD, "pid": PID, "tid": PID, "cpu": CPU, "mode": "user", "callchain_nr": 1, "stack_size": 8},
        {"index": 3, "time_ns": 135, "ip": 0x400110, "raw_bytes": 56, "period": PERIOD, "pid": PID, "tid": PID, "cpu": CPU, "mode": "user", "callchain_nr": 1, "stack_size": 8},
    ]


def _bind_sample_identity(rows: list[dict], stream: dict) -> list[dict]:
    """Attach the scanner's descriptor/raw-event source identities."""
    for row, descriptor in zip(rows, stream["sample_records"]):
        row["dump_line"] = descriptor["line"]
        row["raw_event_line"] = descriptor["raw_event_line"]
        row["physical_offset"] = descriptor["record_offset"]
    return rows


def test_empty_and_non_target_owner() -> None:
    path = _dump()
    try:
        stream = ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        rows = _bind_sample_identity(_samples(), stream)
        packet = ledger.build_ledger(
            stream,
            rows,
            expected_pid=PID,
            expected_cpu=CPU,
            expected_period=PERIOD,
            target_path=TARGET,
            target_build_id=BUILD_ID,
            target_comm="target",
            terminal_lost_samples=0,
            marker="profiled_measured_case",
            marker_dso=TARGET,
            attributions={
                2: {"physical_leaf": {"ip": 0x400100, "dso_path": TARGET, "build_id": BUILD_ID, "symbol": "libc_like", "map_record_sequence": 5}, "target_source_eligible": True},
                3: {"physical_leaf": {"ip": 0x400110, "dso_path": TARGET, "build_id": BUILD_ID, "symbol": "profiled_measured_case", "map_record_sequence": 5}, "target_source_eligible": True, "marker_frame": {"present": True, "resolved": True, "dso_path": TARGET, "build_id": BUILD_ID, "symbol": "profiled_measured_case", "ip": 0x400110, "map_record_sequence": 5}},
            },
        )
        assert packet["gates"]["valid_empty_is_unknown"] is True
        assert packet["gates"]["all_sample_empty_gate"] is False
        assert packet["ledger"]["D"] == 3 * PERIOD
        assert packet["ledger"]["M"] == PERIOD
        assert packet["ledger"]["R"] == 2 * PERIOD
        assert packet["ledger"]["U_target"] == 0
        assert packet["ledger"]["phase_periods"]["startup"] == PERIOD
        assert packet["ledger"]["target_known_owner_period"][f"{TARGET}::libc_like"] == PERIOD
        assert packet["ledger"]["conservation"] == {
            "D_equals_phase_sum": True,
            "R_equals_D_minus_M": True,
            "T_equals_known_plus_U_target": True,
            "M_equals_known_plus_U_marker": True,
        }
    finally:
        path.unlink()


def test_duplicate_sample_index_rejected() -> None:
    path = _dump()
    try:
        stream = ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        rows = _bind_sample_identity(_samples(), stream)
        rows[1]["index"] = 1
        try:
            ledger.build_ledger(
                stream,
                rows,
                expected_pid=PID,
                expected_cpu=CPU,
                expected_period=PERIOD,
                target_path=TARGET,
                target_build_id=BUILD_ID,
                target_comm="target",
                terminal_lost_samples=0,
            )
        except ledger.LedgerFailure as error:
            assert error.code == "sample_identity"
        else:
            raise AssertionError("duplicate sample index was admitted")
    finally:
        path.unlink()


def test_physical_offset_join_rejected() -> None:
    path = _dump()
    try:
        stream = ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        rows = _bind_sample_identity(_samples(), stream)
        rows[0]["physical_offset"] += 1
        try:
            ledger.build_ledger(
                stream,
                rows,
                expected_pid=PID,
                expected_cpu=CPU,
                expected_period=PERIOD,
                target_path=TARGET,
                target_build_id=BUILD_ID,
                target_comm="target",
                terminal_lost_samples=0,
            )
        except ledger.LedgerFailure as error:
            assert error.code == "sample_identity"
        else:
            raise AssertionError("scanner physical offset mutation was admitted")
    finally:
        path.unlink()


def test_marker_ancestor_with_unknown_leaf_is_conserved() -> None:
    path = _dump()
    try:
        stream = ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        rows = _bind_sample_identity(_samples(), stream)
        packet = ledger.build_ledger(
            stream,
            rows,
            expected_pid=PID,
            expected_cpu=CPU,
            expected_period=PERIOD,
            target_path=TARGET,
            target_build_id=BUILD_ID,
            target_comm="target",
            terminal_lost_samples=0,
            marker="profiled_measured_case",
            marker_dso=TARGET,
            attributions={
                3: {
                    "physical_unknown": True,
                    "target_source_eligible": False,
                    "marker_frame": {
                        "present": True,
                        "resolved": True,
                        "dso_path": TARGET,
                        "build_id": BUILD_ID,
                        "symbol": "profiled_measured_case",
                        "ip": 0x400110,
                        "map_record_sequence": 5,
                    },
                }
            },
        )
        assert packet["ledger"]["M"] == PERIOD
        assert packet["ledger"]["U_marker"] == PERIOD
        assert packet["ledger"]["marker_known_owner_period"] == {}
        assert packet["ledger"]["conservation"]["M_equals_known_plus_U_marker"] is True
    finally:
        path.unlink()


def test_startup_owner_uses_pre_exec_map() -> None:
    maps = [{
        "kind": "PERF_RECORD_MMAP2",
        "record_sequence": 1,
        "start": 0x500000,
        "length": 0x1000,
        "path": "/tmp/preexec.so",
        "build_id": BUILD_ID,
    }]
    owner, eligible, resolved, active = ledger._owner_from_attribution(
        {"physical_leaf": {"ip": 0x500010, "dso_path": "/tmp/preexec.so", "build_id": BUILD_ID, "symbol": "preexec", "map_record_sequence": 1}, "target_source_eligible": False},
        sample={"index": 1},
        active_maps=maps,
        record_sequence=2,
        comm_sequence=5,
        target_path=TARGET,
        target_build_id=BUILD_ID,
    )
    assert owner == "/tmp/preexec.so::preexec"
    assert eligible is False
    assert resolved is True
    assert active is not None and active["record_sequence"] == 1


def test_target_source_requires_executable_mapping() -> None:
    maps = [{
        "kind": "PERF_RECORD_MMAP2",
        "record_sequence": 1,
        "start": 0x500000,
        "length": 0x1000,
        "path": TARGET,
        "build_id": BUILD_ID,
        "permissions": "r--p",
        "pf_x": False,
    }]
    try:
        ledger._owner_from_attribution(
            {"physical_leaf": {"ip": 0x500010, "dso_path": TARGET, "build_id": BUILD_ID, "symbol": "nonexec", "map_record_sequence": 1}, "target_source_eligible": True},
            sample={"index": 1},
            active_maps=maps,
            record_sequence=2,
            comm_sequence=1,
            target_path=TARGET,
            target_build_id=BUILD_ID,
        )
    except ledger.LedgerFailure as error:
        assert error.code == "owner_binding"
    else:
        raise AssertionError("non-executable target mapping was admitted as source")


def test_initial_zero_sample_id_is_narrowly_bound() -> None:
    path = _dump()
    try:
        text = path.read_text()
        path.write_text(text.replace("0 0 0x0 [0x38]: PERF_RECORD_COMM: perf-exec:77/77", "1 0 0x0 [0x38]: PERF_RECORD_COMM: perf-exec:77/77", 1))
        try:
            ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        except ledger.LedgerFailure as error:
            assert error.code == "raw_metadata"
        else:
            raise AssertionError("edited all-zero bootstrap descriptor was admitted")
    finally:
        path.unlink()


def test_lost_samples_count_is_bound() -> None:
    raw = bytearray(32)
    struct.pack_into("<IHHQ", raw, 0, 13, 0, 32, 3)
    descriptor = "31 100 0x0 [0x20]: PERF_RECORD_LOST_SAMPLES: 3"
    path = Path(tempfile.mktemp(suffix=".perf-mmap-dump"))
    lines = [f". ... raw event: size {len(raw)} bytes"]
    for offset in range(0, len(raw), 16):
        chunk = raw[offset : offset + 16]
        lines.append(f".  {offset:04x}:  " + " ".join(f"{byte:02x}" for byte in chunk))
    lines.append(descriptor)
    path.write_text("\n".join(lines) + "\n")
    try:
        parsed = ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        assert parsed["terminal_lost_values"] == [3]
        assert parsed["records"][0]["raw_metadata"]["fields"]["lost"] == 3
    finally:
        path.unlink()


def test_text_raw_metadata_mismatch_rejected() -> None:
    path = _dump()
    try:
        text = path.read_text()
        path.write_text(text.replace(f"r-xp {TARGET}", "r-xp /tmp/mutated-target", 1))
        try:
            ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        except ledger.LedgerFailure as error:
            assert error.code == "raw_metadata"
        else:
            raise AssertionError("text-only MMAP2 mutation was admitted")
    finally:
        path.unlink()


def test_thread_map_u64_layout() -> None:
    raw = bytearray(40)
    struct.pack_into("<IHHQ", raw, 0, 73, 0, 40, 1)
    struct.pack_into("<Q", raw, 16, PID)
    raw[24 : 24 + 4] = b"main"
    descriptor = "31 100 0x0 [0x28]: PERF_RECORD_THREAD_MAP nr: 1 thread: 77"
    path = Path(tempfile.mktemp(suffix=".perf-mmap-dump"))
    lines = [f". ... raw event: size {len(raw)} bytes"]
    for offset in range(0, len(raw), 16):
        chunk = raw[offset : offset + 16]
        lines.append(f".  {offset:04x}:  " + " ".join(f"{byte:02x}" for byte in chunk))
    lines.append(descriptor)
    path.write_text("\n".join(lines) + "\n")
    try:
        parsed = ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        assert parsed["record_kind_counts"] == {"PERF_RECORD_THREAD_MAP": 1}
        assert parsed["records"][0]["raw_metadata"]["fields"]["threads"] == [PID]
        raw[16] ^= 1
        mutated = lines[:1]
        for offset in range(0, len(raw), 16):
            chunk = raw[offset : offset + 16]
            mutated.append(f".  {offset:04x}:  " + " ".join(f"{byte:02x}" for byte in chunk))
        mutated.append(descriptor)
        path.write_text("\n".join(mutated) + "\n")
        try:
            ledger.parse_lifecycle_dump(path, expected_pid=PID, expected_cpu=CPU)
        except ledger.LedgerFailure as error:
            assert error.code == "raw_metadata"
        else:
            raise AssertionError("THREAD_MAP raw mutation was admitted")
    finally:
        path.unlink()


if __name__ == "__main__":
    test_empty_and_non_target_owner()
    test_duplicate_sample_index_rejected()
    test_physical_offset_join_rejected()
    test_marker_ancestor_with_unknown_leaf_is_conserved()
    test_startup_owner_uses_pre_exec_map()
    test_target_source_requires_executable_mapping()
    test_initial_zero_sample_id_is_narrowly_bound()
    test_lost_samples_count_is_bound()
    test_text_raw_metadata_mismatch_rejected()
    test_thread_map_u64_layout()
    print("0978 lifecycle ledger controls: PASS")
