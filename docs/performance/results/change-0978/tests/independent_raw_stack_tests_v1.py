#!/usr/bin/env python3
"""Independent mutation controls for the 0978 variable-layout scanner.

These tests construct perf-script text and raw record bytes from the Linux
UAPI framing rules.  They assert conservation and externally visible
invariants; they do not call scanner-private parsers or copy scanner logic.
Run with ``python3 -B`` so this external test does not create bytecode files.
"""
from __future__ import annotations

import importlib.util
import struct
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Iterable


SCANNER_PATH = Path("/home/zhuhe/litchi-performance-0978/tools/raw_stack_0978.py")
spec = importlib.util.spec_from_file_location("raw_stack_0978_under_test", SCANNER_PATH)
if spec is None or spec.loader is None:
    raise RuntimeError(f"cannot load scanner at {SCANNER_PATH}")
scanner = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = scanner
spec.loader.exec_module(scanner)


PID = 4242
CPU = 31
PERIOD = 1_000_000
STACK_CAP = 32_768
FULL_NAMES = [
    "IP",
    "TID",
    "TIME",
    "CALLCHAIN",
    "CPU",
    "PERIOD",
    "REGS_USER",
    "STACK_USER",
]


def make_contract(*, names: Iterable[str] = FULL_NAMES, regs_mask: int = 0b10101) -> dict:
    return scanner.make_event_contract(
        sample_type_names=list(names),
        regs_user_mask=regs_mask,
        period=PERIOD,
        stack_cap=STACK_CAP,
        source="independent-0978-mutation-tests-v1",
    )


def make_sample_record(
    *,
    names: Iterable[str] = FULL_NAMES,
    regs_mask: int = 0b10101,
    abi: int = 2,
    callchain: tuple[int, ...] = (0x501000, 0x502000),
    stack: bytes = b"stack",
    dyn_size: int = 5,
    pid: int = PID,
    tid: int = PID,
    cpu: int = CPU,
    time_ns: int = 123456789,
    period: int = PERIOD,
    misc: int = 2,
) -> bytes:
    """Serialize one sample in kernel PERF_SAMPLE enum order.

    This is a fixture writer, rather than a second decoder: values are supplied
    explicitly and the assertions below inspect only scanner output and gates.
    """
    selected = set(names)
    body = bytearray()
    if "IP" in selected:
        body.extend(struct.pack("<Q", 0x500000))
    if "TID" in selected:
        body.extend(struct.pack("<II", pid, tid))
    if "TIME" in selected:
        body.extend(struct.pack("<Q", time_ns))
    if "CPU" in selected:
        body.extend(struct.pack("<II", cpu, 0))
    if "PERIOD" in selected:
        body.extend(struct.pack("<Q", period))
    if "CALLCHAIN" in selected:
        body.extend(struct.pack("<Q", len(callchain)))
        for value in callchain:
            body.extend(struct.pack("<Q", value))
    if "REGS_USER" in selected:
        body.extend(struct.pack("<Q", abi))
        if abi:
            for index in range(regs_mask.bit_count()):
                body.extend(struct.pack("<Q", 0x1111 + index))
    if "STACK_USER" in selected:
        body.extend(struct.pack("<Q", len(stack)))
        body.extend(stack)
        if stack:
            body.extend(struct.pack("<Q", dyn_size))
    size = 8 + len(body)
    if size > 0xFFFF:
        raise ValueError("fixture record exceeds perf header size")
    return struct.pack("<IHH", 9, misc, size) + body


def sample_line(
    *,
    record_size: int = 0x48,
    pid: int = PID,
    tid: int = PID,
    cpu: int = CPU,
    time_ns: int = 123456789,
    period: int = PERIOD,
    misc: int = 2,
    ip: int = 0x500000,
) -> str:
    return (
        f"{cpu} {time_ns} 0x500000 [0x{record_size:x}]: "
        f"PERF_RECORD_SAMPLE(IP, 0x{misc:x}): "
        f"{pid}/{tid}: 0x{ip:x} period: {period}"
    )


def sample_for(data: bytes, **kwargs: object) -> str:
    return sample_line(record_size=len(data), **kwargs)


def raw_event_header(data: bytes, *, declared_size: int | None = None) -> str:
    size = len(data) if declared_size is None else declared_size
    return f". ... raw event: size {size} bytes"


def hex_line(offset: int, payload: bytes) -> str:
    return f".  {offset:04x}:  " + " ".join(f"{value:02x}" for value in payload)


def hex_line_with_ascii_gutter(offset: int, payload: bytes) -> str:
    # perf-script reserves a 48-character (16 * ``"xx "``) byte column before
    # its printable gutter.  Keep that layout even on the short final row.
    # The gutter deliberately begins with hex-looking tokens so a greedy byte
    # regex cannot silently treat ASCII as record bytes.
    byte_column = "".join(f"{value:02x} " for value in payload).ljust(48)
    return f".  {offset:04x}:  {byte_column}  ab de ad  |ASCII gutter|"


def location_line(data: bytes, *, offset: int, event: int = 9, path: str = "/synthetic/perf.data") -> str:
    return f"0x{offset:x}@{path} [0x{len(data):x}]: event: {event}"


def dump_for(events: Iterable[tuple[bytes, list[str], str | None] | tuple[bytes, list[str], str | None, str]]) -> str:
    lines: list[str] = []
    for index, event in enumerate(events):
        data, hex_rows, decoded_line = event[:3]
        location = event[3] if len(event) == 4 else location_line(data, offset=0x1000 + index * 0x100)
        lines.append("")
        lines.append(location)
        lines.append(raw_event_header(data))
        lines.extend(hex_rows)
        if decoded_line is not None:
            lines.append(decoded_line)
    return "\n".join(lines) + "\n"


def full_hex_rows(data: bytes, *, split_at: int | None = None) -> list[str]:
    if split_at is None:
        return [hex_line(offset, data[offset : offset + 16]) for offset in range(0, len(data), 16)]
    if not 0 < split_at < len(data):
        return [hex_line(0, data)]
    return [hex_line(0, data[:split_at]), hex_line(split_at, data[split_at:])]


def run_dump(
    text: str,
    *,
    contract: dict | None = None,
    expected_pid: int = PID,
    expected_cpu: int | None = CPU,
    expected_period: int | None = PERIOD,
    stack_cap: int = STACK_CAP,
    callback: list[dict] | None = None,
) -> dict:
    with tempfile.NamedTemporaryFile("w", suffix=".dump", delete=False) as handle:
        path = Path(handle.name)
        handle.write(text)
    try:
        return scanner.scan_dump(
            path,
            expected_pid=expected_pid,
            expected_cpu=expected_cpu,
            expected_period=expected_period,
            stack_cap=stack_cap,
            event_contract=contract or make_contract(),
            sample_callback=(callback.append if callback is not None else None),
            retain_samples=True,
        )
    finally:
        path.unlink(missing_ok=True)


def assert_scan_failure(test: unittest.TestCase, text: str, codes: set[str], **kwargs: object) -> None:
    with test.assertRaises(scanner.ScanFailure) as context:
        run_dump(text, **kwargs)
    test.assertIn(context.exception.code, codes, str(context.exception))


class IndependentRawStackMutationTests(unittest.TestCase):
    def test_variable_abi_register_masks_and_zero_stack_are_conserved(self) -> None:
        contract = make_contract(regs_mask=0b10101)
        rich64 = make_sample_record(abi=2, stack=b"abc", dyn_size=2)
        rich32 = make_sample_record(
            abi=1, callchain=(0x503000,), stack=b"12345678", dyn_size=8,
        )
        empty = make_sample_record(abi=0, callchain=(), stack=b"", dyn_size=0)
        text = dump_for(
            (
                data,
                full_hex_rows(data),
                sample_for(data),
                location_line(data, offset=offset),
            )
            for offset, data in zip((0x1000, 0x2000, 0x3000), (rich64, rich32, empty))
        )
        callback: list[dict] = []
        result = run_dump(text, contract=contract, callback=callback)
        self.assertEqual(result["raw_event_count"], 3)
        self.assertEqual(result["raw_sample_record_count"], 3)
        self.assertEqual(result["sample_count"], 3)
        self.assertTrue(result["sample_conservation"])
        self.assertEqual(len(callback), 3)
        self.assertEqual(result["regs_abi_counts"], {"32": 1, "64": 1, "NONE": 1})
        self.assertEqual(result["stack_shape_counts"], {"nonzero": 2, "zero": 1})
        self.assertEqual(result["sample_ledger"][0]["regs_words"], 3)
        self.assertEqual(result["sample_ledger"][1]["regs_words"], 3)
        self.assertEqual(result["sample_ledger"][2]["regs_words"], 0)
        self.assertIsNone(result["sample_ledger"][2]["stack_dyn_size"])
        self.assertFalse(result["sample_ledger"][2]["stack_available"])

    def test_register_mask_and_abi_truncation_fail_closed(self) -> None:
        # A changed mask is valid only when the corresponding register payload
        # has the new width.  The same bytes under the original mask must fail
        # closed rather than shifting the stack fields into register slots.
        narrow_contract = make_contract(regs_mask=0b11)
        valid_narrow = make_sample_record(regs_mask=0b11, abi=2, stack=b"abc", dyn_size=3)
        valid = run_dump(
            dump_for([(valid_narrow, full_hex_rows(valid_narrow), sample_for(valid_narrow))]),
            contract=narrow_contract,
        )
        self.assertEqual(valid["sample_ledger"][0]["regs_words"], 2)

        text = dump_for([(valid_narrow, full_hex_rows(valid_narrow), sample_for(valid_narrow))])
        assert_scan_failure(self, text, {"regs_truncated", "sample_trailing_bytes", "stack_truncated"})

        unknown_abi = make_sample_record(abi=3, stack=b"abc", dyn_size=3)
        text = dump_for([(unknown_abi, full_hex_rows(unknown_abi), sample_for(unknown_abi))])
        assert_scan_failure(self, text, {"regs_abi"})

    def test_zero_stack_has_no_dynamic_trailer(self) -> None:
        valid = make_sample_record(abi=0, callchain=(), stack=b"")
        invalid = struct.pack("<IHH", 9, 2, len(valid) + 8) + valid[8:] + struct.pack("<Q", 0)
        text = dump_for([(invalid, full_hex_rows(invalid), sample_for(invalid))])
        assert_scan_failure(self, text, {"sample_trailing_bytes", "stack_truncated"})

    def test_missing_duplicate_and_overlapping_hex_rows_fail(self) -> None:
        data = make_sample_record()
        line = sample_for(data)
        missing = dump_for([(data, full_hex_rows(data[:-1]), line)])
        assert_scan_failure(self, missing, {"raw_record_truncated"})

        rows = full_hex_rows(data)
        duplicated = dump_for([(data, rows + rows, line)])
        assert_scan_failure(self, duplicated, {"raw_hex_overlap"})

        split = 16
        overlap = dump_for(
            [(data, [hex_line(0, data[:split]), hex_line(split - 1, data[split - 1 : split + 15])], line)]
        )
        assert_scan_failure(self, overlap, {"raw_hex_overlap"})

    def test_fixed_width_hex_column_ignores_hex_like_ascii_gutter(self) -> None:
        data = make_sample_record()
        rows = [
            hex_line_with_ascii_gutter(offset, data[offset : offset + 16])
            for offset in range(0, len(data), 16)
        ]
        result = run_dump(dump_for([(data, rows, sample_for(data))]))
        self.assertEqual(result["raw_event_count"], 1)
        self.assertEqual(result["raw_sample_record_count"], 1)
        self.assertEqual(result["sample_count"], 1)
        self.assertTrue(result["sample_conservation"])

    def test_sample_pair_and_raw_record_conservation(self) -> None:
        data = make_sample_record()
        line = sample_for(data)
        unpaired = dump_for([(data, full_hex_rows(data), None)])
        assert_scan_failure(self, unpaired, {"sample_conservation"})

        duplicate_description = dump_for(
            [(data, full_hex_rows(data), line + "\n" + line)]
        )
        assert_scan_failure(self, duplicate_description, {"sample_double_pair"})

        # Two byte-identical raw records are two physical records at distinct
        # offsets; a conservation checker must retain both rather than dedup
        # them by decoded values.
        repeated = dump_for(
            [
                (data, full_hex_rows(data), sample_for(data), location_line(data, offset=0x1000)),
                (data, full_hex_rows(data), sample_for(data), location_line(data, offset=0x2000)),
            ]
        )
        result = run_dump(repeated)
        self.assertEqual(result["raw_event_count"], 2)
        self.assertEqual(result["raw_sample_record_count"], 2)
        self.assertEqual(result["sample_count"], 2)
        self.assertTrue(result["sample_conservation"])

        same_offset = dump_for(
            [
                (data, full_hex_rows(data), sample_for(data), location_line(data, offset=0x1000)),
                (data, full_hex_rows(data), sample_for(data), location_line(data, offset=0x1000)),
            ]
        )
        assert_scan_failure(
            self,
            same_offset,
            {"raw_record_identity", "duplicate_sample_identity", "sample_duplicate", "sample_conservation"},
        )

    def test_unrelated_raw_records_are_counted_but_not_paired_as_samples(self) -> None:
        sample = make_sample_record()
        unrelated = struct.pack("<IHHQ", 1, 0, 16, 0x1234)
        text = dump_for(
            [
                (unrelated, full_hex_rows(unrelated), None, location_line(unrelated, offset=0x4000, event=1)),
                (sample, full_hex_rows(sample), sample_for(sample), location_line(sample, offset=0x5000)),
            ]
        )
        result = run_dump(text)
        self.assertEqual(result["raw_event_count"], 2)
        self.assertEqual(result["raw_record_type_counts"], {"1": 1, "9": 1})
        self.assertEqual(result["raw_unconsumed_count"], 1)
        self.assertEqual(result["raw_sample_record_count"], 1)
        self.assertEqual(result["sample_count"], 1)
        self.assertTrue(result["sample_conservation"])

    def test_summary_and_raw_header_mismatches_are_visible(self) -> None:
        data = make_sample_record()

        cpu_text_mismatch = dump_for(
            [(data, full_hex_rows(data), sample_for(data, cpu=CPU + 1))]
        )
        cpu_result = run_dump(cpu_text_mismatch)
        self.assertFalse(cpu_result["pid_tid_cpu_valid"])
        self.assertTrue(any("cpu" in item["reason"] for item in cpu_result["violations"]))

        time_text_mismatch = dump_for(
            [(data, full_hex_rows(data), sample_for(data, time_ns=123456790))]
        )
        time_result = run_dump(time_text_mismatch)
        self.assertTrue(any("time" in item["reason"] for item in time_result["violations"]))

        period_text_mismatch = dump_for(
            [(data, full_hex_rows(data), sample_for(data, period=PERIOD + 1))]
        )
        period_result = run_dump(period_text_mismatch)
        self.assertFalse(period_result["period_valid"])
        self.assertTrue(any("period" in item["reason"] for item in period_result["violations"]))

        invalid_period = make_sample_record(period=0)
        invalid_period_result = run_dump(
            dump_for([(invalid_period, full_hex_rows(invalid_period), sample_for(invalid_period, period=0))])
        )
        self.assertFalse(invalid_period_result["period_valid"])
        self.assertTrue(any(item["reason"] == "sample_period" for item in invalid_period_result["violations"]))

        invalid_pid = make_sample_record(pid=PID + 1, tid=PID + 1)
        invalid_pid_result = run_dump(
            dump_for([(invalid_pid, full_hex_rows(invalid_pid), sample_for(invalid_pid, pid=PID + 1, tid=PID + 1))])
        )
        self.assertFalse(invalid_pid_result["pid_tid_cpu_valid"])
        self.assertTrue(any(item["reason"] == "sample_pid_tid" for item in invalid_pid_result["violations"]))

        invalid_cpu = make_sample_record(cpu=CPU + 1)
        invalid_cpu_result = run_dump(
            dump_for([(invalid_cpu, full_hex_rows(invalid_cpu), sample_for(invalid_cpu, cpu=CPU + 1))])
        )
        self.assertFalse(invalid_cpu_result["pid_tid_cpu_valid"])
        self.assertTrue(any(item["reason"] == "sample_cpu" for item in invalid_cpu_result["violations"]))

        raw_misc_mismatch = make_sample_record(misc=1)
        misc_result = run_dump(
            dump_for([(raw_misc_mismatch, full_hex_rows(raw_misc_mismatch), sample_for(raw_misc_mismatch, misc=2))])
        )
        self.assertTrue(any(item["reason"] == "summary_misc_mismatch" for item in misc_result["violations"]))

        bad_header_size = struct.pack("<IHH", 9, 2, len(data) - 1) + data[8:]
        assert_scan_failure(
            self,
            dump_for([(bad_header_size, full_hex_rows(bad_header_size), sample_for(bad_header_size))]),
            {"record_size_mismatch"},
        )

        bad_type = struct.pack("<IHH", 1, 2, len(data)) + data[8:]
        assert_scan_failure(
            self,
            dump_for([(bad_type, full_hex_rows(bad_type), sample_for(bad_type))]),
            {"raw_record_identity", "sample_record_type"},
        )

    def test_evlist_duplicate_attributes_and_empty_types_fail(self) -> None:
        line = (
            "instructions:u: type: 4 (cpu), size: 144, config: 0xc0 (instructions), "
            "{ sample_period, sample_freq }: 1000000, "
            "sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER, "
            "read_format: ID|LOST, disabled: 1, exclude_kernel: 1, exclude_hv: 1, "
            "enable_on_exec: 1, task: 1, sample_id_all: 1, "
            "sample_regs_user: 0xff0fff, sample_stack_user: 32768"
        )
        mutations = (
            line.replace("sample_type:", "sample_type: IP|TID, sample_type:", 1),
            line.replace(
                "sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
                "sample_type: IP||TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
            ),
            line.replace(
                "sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
                "sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER|IP",
            ),
            line + ", sample_regs_user: 0xff0fff",
        )
        for malformed in mutations:
            with self.subTest(malformed=malformed):
                with self.assertRaises(scanner.ScanFailure) as context:
                    scanner.parse_evlist_contract(malformed, expected_period=PERIOD, expected_stack_cap=STACK_CAP)
                self.assertEqual(context.exception.code, "evlist_contract")

    def test_oversized_raw_event_is_rejected_before_unbounded_allocation(self) -> None:
        text = raw_event_header(b"", declared_size=0x10000) + "\n"
        assert_scan_failure(
            self,
            text,
            {"raw_record_header", "raw_record_size", "raw_record_too_large"},
        )


if __name__ == "__main__":
    unittest.main()
