#!/usr/bin/env python3
"""Focused synthetic controls for the 0978 variable-layout scanner.

The fixtures exercise the byte layout at the scanner boundary.  They do not
stand in for perf capture qualification: retained-capture replay is separate.
"""
from __future__ import annotations

import struct
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import raw_stack_0978 as scanner  # noqa: E402


PID = 4242
CPU = 31
PERIOD = 1_000_000
STACK_CAP = 32_768
NAMES = ["IP", "TID", "TIME", "CALLCHAIN", "CPU", "PERIOD", "REGS_USER", "STACK_USER"]


def contract() -> dict:
    return scanner.make_event_contract(
        sample_type_names=NAMES, regs_user_mask=0b10101,
        period=PERIOD, stack_cap=STACK_CAP, source="synthetic-0978",
    )


def sample_bytes(*, callchain: list[int], abi: int, stack: bytes, dyn: int = 0, misc: int = 2) -> bytes:
    """Build the kernel order: IP,TID,TIME,CPU,PERIOD,CALLCHAIN,REGS,STACK."""
    body = bytearray(struct.pack("<QIIQIIQ", 0x500000, PID, PID, 123456789, CPU, 0, PERIOD))
    body.extend(struct.pack("<Q", len(callchain)))
    body.extend(struct.pack("<" + "Q" * len(callchain), *callchain) if callchain else b"")
    body.extend(struct.pack("<Q", abi))
    if abi:
        body.extend(struct.pack("<QQQ", 0x1111, 0x2222, 0x3333)[: 8 * 3])
    body.extend(struct.pack("<Q", len(stack)))
    body.extend(stack)
    if stack:
        body.extend(struct.pack("<Q", dyn))
    header_size = 8 + len(body)
    return struct.pack("<IHH", 9, misc, header_size) + body


def raw_bytes(record_type: int, payload: bytes) -> bytes:
    if record_type == 9:
        return payload
    return struct.pack("<IHH", record_type, 0, len(payload)) + payload[8:]


def dump_for(records: list[tuple[bytes, str | None]]) -> str:
    lines: list[str] = []
    for data, sample_line in records:
        lines.append("")
        lines.append(f". ... raw event: size {len(data)} bytes")
        for offset in range(0, len(data), 16):
            chunk = data[offset : offset + 16]
            lines.append(f".  {offset:04x}:  " + " ".join(f"{byte:02x}" for byte in chunk))
        if sample_line is not None:
            lines.append(sample_line)
    return "\n".join(lines) + "\n"


def sample_line(data: bytes, *, misc: int = 2) -> str:
    ip, pid, tid = struct.unpack_from("<QII", data, 8)
    return (
        f"{CPU} 123456789 0x500000 [0x{len(data):x}]: "
        f"PERF_RECORD_SAMPLE(IP, 0x{misc:x}): {pid}/{tid}: 0x{ip:x} period: {PERIOD}"
    )


def run_dump(text: str, *, expected_pid: int = PID) -> dict:
    with tempfile.NamedTemporaryFile("w", suffix=".dump", delete=False) as handle:
        handle.write(text)
        path = Path(handle.name)
    try:
        return scanner.scan_dump(
            path, expected_pid=expected_pid, expected_cpu=CPU,
            expected_period=PERIOD, stack_cap=STACK_CAP,
            event_contract=contract(), retain_samples=True,
        )
    finally:
        path.unlink()


def expect_failure(text: str, codes: set[str]) -> None:
    try:
        run_dump(text)
    except scanner.ScanFailure as error:
        assert error.code in codes, (error.code, codes)
    else:
        raise AssertionError(f"expected ScanFailure in {codes}")


def test_variable_layouts() -> None:
    rich = sample_bytes(callchain=[0x501000, 0x502000], abi=2, stack=b"abc", dyn=2)
    none = sample_bytes(callchain=[], abi=0, stack=b"")
    result = run_dump(dump_for([(rich, sample_line(rich)), (none, sample_line(none))]))
    assert result["sample_count"] == 2
    assert result["sample_conservation"] is True
    assert result["raw_record_type_counts"] == {"9": 2}
    assert result["raw_record_size_counts"][str(len(none))] == 1
    assert result["regs_abi_counts"] == {"64": 1, "NONE": 1}
    assert result["stack_shape_counts"] == {"nonzero": 1, "zero": 1}
    assert result["sample_ledger"][0]["callchain_nr"] == 2
    assert result["sample_ledger"][0]["stack_dyn_size"] == 2
    assert result["sample_ledger"][1]["missing_unwindability"] == ["callchain", "user_stack", "user_regs"]


def test_all_records_and_conservation() -> None:
    rich = sample_bytes(callchain=[0x501000], abi=1, stack=b"12345678", dyn=8)
    other = struct.pack("<IHHQ", 1, 0, 16, 0x1234)
    result = run_dump(dump_for([(other, None), (rich, sample_line(rich))]))
    assert result["raw_event_count"] == 2
    assert result["raw_record_type_counts"] == {"1": 1, "9": 1}
    assert result["raw_unconsumed_count"] == 1
    assert result["sample_conservation"] is True
    assert result["sample_count"] == 1


def test_physical_identity_and_summary_controls() -> None:
    rich = sample_bytes(callchain=[0x501000], abi=1, stack=b"12345678", dyn=8)
    location = f"0x1000@/tmp/synthetic.perf [0x{len(rich):x}]: event: 9\n"
    result = run_dump(location + dump_for([(rich, sample_line(rich))]))
    assert result["physical_identity_valid"] is True
    assert result["physical_location_count"] == 1
    bad_cpu = sample_line(rich).replace(f"{CPU} 123456789", "30 123456789", 1)
    bad = run_dump(location + dump_for([(rich, bad_cpu)]))
    assert bad["pid_tid_cpu_valid"] is False
    assert any(row["reason"] == "summary_cpu_mismatch" for row in bad["violations"])
    duplicate = location + dump_for([(rich, sample_line(rich))]) + location + dump_for([(rich, sample_line(rich))])
    expect_failure(duplicate, {"raw_record_identity"})


def test_failure_controls() -> None:
    rich = sample_bytes(callchain=[0x501000], abi=2, stack=b"12345678", dyn=8)
    trailing = rich + b"\0\0\0\0"
    trailing = struct.pack("<IHH", 9, 2, len(trailing)) + trailing[8:]
    expect_failure(dump_for([(trailing, sample_line(trailing))]), {"sample_trailing_bytes"})
    truncated = rich[:-8]
    truncated = struct.pack("<IHH", 9, 2, len(truncated)) + truncated[8:]
    expect_failure(dump_for([(truncated, sample_line(truncated))]), {"sample_truncated", "stack_truncated"})
    expect_failure("\n. ... raw event: size 65536 bytes\n", {"raw_record_header"})
    ascii_hex = ".  0000:  " + " ".join(["00"] * 16) + "  ab ................"
    assert len(scanner._parse_hex_line(ascii_hex)[1]) == 16
    try:
        scanner._parse_hex_line(".  0000:  " + " ".join(["00"] * 17) + "  malformed")
    except scanner.ScanFailure as error:
        assert error.code == "raw_hex_parse"
    else:
        raise AssertionError("oversized hex column was admitted")
    bad = contract()
    bad["sample_type_mask"] |= 1 << 30
    with tempfile.NamedTemporaryFile("w", suffix=".dump", delete=False) as handle:
        path = Path(handle.name)
        handle.write("")
    try:
        try:
            scanner.scan_dump(path, expected_pid=PID, expected_cpu=CPU,
                              expected_period=PERIOD, stack_cap=STACK_CAP,
                              event_contract=bad)
        except scanner.ScanFailure as error:
            assert error.code == "evlist_contract"
        else:
            raise AssertionError("unsupported sample bit was admitted")
    finally:
        path.unlink()


def test_evlist_contract_controls() -> None:
    line = ("instructions:u: type: 4 (cpu), size: 144, config: 0xc0 (instructions), "
            "{ sample_period, sample_freq }: 1000000, sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER, "
            "read_format: ID|LOST, disabled: 1, exclude_kernel: 1, exclude_hv: 1, enable_on_exec: 1, task: 1, "
            "sample_id_all: 1, sample_regs_user: 0xff0fff, sample_stack_user: 32768")
    parsed = scanner.parse_evlist_contract(line, expected_period=PERIOD, expected_stack_cap=STACK_CAP)
    assert parsed["sample_type_names"][3] == "CALLCHAIN"
    for mutation in (line.replace("sample_type:", "sample_type: IP|"), line + ", inherit: 1", line.replace("sample_type:", "sample_type: IP|TID, sample_type:")):
        try:
            scanner.parse_evlist_contract(mutation, expected_period=PERIOD, expected_stack_cap=STACK_CAP)
        except scanner.ScanFailure:
            pass
        else:
            raise AssertionError("malformed evlist contract was admitted")


if __name__ == "__main__":
    test_variable_layouts()
    test_all_records_and_conservation()
    test_physical_identity_and_summary_controls()
    test_failure_controls()
    test_evlist_contract_controls()
    print("0978 synthetic scanner controls: PASS")
