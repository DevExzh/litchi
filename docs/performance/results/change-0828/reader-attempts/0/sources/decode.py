"""Root-owned perf decoding and live-binary owner/phase symbol evidence."""

from __future__ import annotations

import gzip
import hashlib
import io
import re
import subprocess
import sys
import time
from pathlib import Path

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
BUILD = c.read(P / "build.json")
FROZEN = c.read(BUILD["frozen_inputs"]["path"])
PERF = P / "perf"
SYMBOLS = P / "symbols"
c.check_no_overrides()
c.stable(FROZEN)

FP = BUILD["binaries"][PLAN["perf"]["binary"]]["artifact"]
assert c.artifact(FP["path"]) == FP
EXPECTED_OWNER = PLAN["perf"]["owner"]
BINARIES = {
    name: BUILD["binaries"][name]["artifact"]
    for name in ("ordinary", "fp")
}
for _descriptor in BINARIES.values():
    assert c.artifact(_descriptor["path"]) == _descriptor


def command_output(command: list[str], stdout_path: Path, stderr_path: Path) -> int:
    assert not stdout_path.exists() and not stderr_path.exists()
    with stdout_path.open("w") as stdout, stderr_path.open("w") as stderr:
        result = subprocess.run(command, cwd=c.ROOT, stdout=stdout, stderr=stderr)
    return result.returncode


def parse_nm(path: Path) -> list[dict[str, str]]:
    rows = []
    pattern = re.compile(r"^\s*([0-9A-Fa-f]+)\s+([0-9A-Fa-f]+)\s+(\S)\s+(\S+)\s*$")
    for line in path.read_text(errors="replace").splitlines():
        match = pattern.match(line)
        if match:
            rows.append({
                "address": match.group(1),
                "size": match.group(2),
                "type": match.group(3),
                "symbol": match.group(4),
            })
    return rows


def parse_nm_text(text: str) -> list[dict[str, str]]:
    rows = []
    pattern = re.compile(r"^\s*([0-9A-Fa-f]+)\s+([0-9A-Fa-f]+)\s+(\S)\s+(.+?)\s*$")
    for line in text.splitlines():
        match = pattern.match(line)
        if match:
            rows.append({
                "address": match.group(1),
                "size": match.group(2),
                "type": match.group(3),
                "symbol": match.group(4),
                "line": line,
            })
    return rows


def static_owner_evidence() -> dict[str, object]:
    started = time.time()
    out = SYMBOLS
    assert not out.exists(), "refusing to overwrite symbol evidence"
    out.mkdir()
    nm = out / "nm-match.txt"
    nm_err = out / "nm-defined.log"
    nm_command = ["nm", "-S", "--defined-only", FP["path"]]
    result = subprocess.run(nm_command, cwd=c.ROOT, capture_output=True, check=False)
    nm_err.write_bytes(result.stderr)
    assert result.returncode == 0, nm_err
    rows = parse_nm_text(result.stdout.decode(errors="replace"))
    demangled = out / "nm-demangled-match.txt"
    demangled_err = out / "nm-demangled.log"
    demangled_command = ["nm", "-C", "-S", "--defined-only", FP["path"]]
    result = subprocess.run(demangled_command, cwd=c.ROOT, capture_output=True, check=False)
    demangled_err.write_bytes(result.stderr)
    assert result.returncode == 0
    demangled_rows = parse_nm_text(result.stdout.decode(errors="replace"))

    phase_owners = PLAN["probe"].get("phase_owners")
    assert isinstance(phase_owners, dict)
    assert set(phase_owners) == {"capture", "set_text", "publish"}
    expected_owners = {"edit_region": EXPECTED_OWNER, **phase_owners}
    owner_lines: dict[str, list[str]] = {}
    matched_rows: dict[str, dict[str, str]] = {}
    assembly_paths: dict[str, dict[str, object]] = {}
    command_rows: dict[str, list[str]] = {}
    phase_nm_lines: list[str] = []
    phase_demangled_lines: list[str] = []
    for label, expected in expected_owners.items():
        owner_rows = [row for row in demangled_rows if row["symbol"] == expected]
        assert len(owner_rows) == 1, (
            f"expected one exact demangled {label} owner {expected!r}, found {len(owner_rows)}"
        )
        owner = owner_rows[0]
        raw_rows = [row for row in rows if (
            row["address"], row["size"], row["type"]
        ) == (owner["address"], owner["size"], owner["type"])]
        assert len(raw_rows) == 1, f"expected one raw symbol for {label} owner {expected!r}"
        matched = raw_rows[0]
        owner_lines[label] = [owner["line"]]
        matched_rows[label] = matched
        if label != "edit_region":
            phase_nm_lines.append(matched["line"])
            phase_demangled_lines.append(owner["line"])

        assembly = out / f"{label}-assembly.txt"
        assembly_err = out / f"{label}-assembly.log"
        assembly_command = [
            "objdump", "--demangle=rust", f"--disassemble={matched['symbol']}", "--wide",
            "--line-numbers", FP["path"],
        ]
        assert command_output(assembly_command, assembly, assembly_err) == 0
        assembly_text = assembly.read_text(errors="replace")
        assert matched["symbol"] in assembly_text
        assert re.search(r"push\s+%rbp", assembly_text)
        assert re.search(r"mov\s+%rsp,%rbp", assembly_text)
        assert re.search(r"\bcall\b", assembly_text), f"{label} owner must retain a call boundary"
        assembly_paths[label] = {
            "assembly": c.artifact(assembly),
            "log": c.artifact(assembly_err),
            "command": assembly_command,
        }
        command_rows[label] = assembly_command
    nm.write_text(matched_rows["edit_region"]["line"] + "\n")
    demangled.write_text(
        owner_lines["edit_region"][0]
        + "\n"
        + "\n".join(
            f"# phase {label}: {owner}"
            for label, owner in phase_owners.items()
        )
        + "\n"
    )
    phase_nm = out / "phase-nm-match.txt"
    phase_demangled = out / "phase-nm-demangled-match.txt"
    phase_nm.write_text("\n".join(phase_nm_lines) + "\n")
    phase_demangled.write_text("\n".join(phase_demangled_lines) + "\n")

    matched = matched_rows["edit_region"]
    c.stable(FROZEN)
    result = {
        "schema": "litchi.performance.0828.symbol-evidence.v1",
        "started": started,
        "ended": time.time(),
        "frame_pointer_and_call_verified": True,
        "owner": EXPECTED_OWNER,
        "owner_function": PLAN["probe"]["owner_function"],
        "phase_owners": phase_owners,
        "owners": expected_owners,
        "owner_lines": owner_lines,
        "matched_rows": matched_rows,
        "matched_mangled_symbol": matched["symbol"],
        "matched_address": matched["address"],
        "matched_size": matched["size"],
        "matched_type": matched["type"],
        "qualified_demangled_lines": owner_lines["edit_region"],
        "binary": FP,
        "binaries": BINARIES,
        "binary_sha256": FP["sha256"],
        "source": c.artifact(P / "build-0/source.json"),
        "probe": c.artifact(P / "build-0/probe.json"),
        "build": c.artifact(P / "build.json"),
        "frozen_inputs": c.artifact(BUILD["frozen_inputs"]["path"]),
        "source_sha256": c.sha(P / "build-0/source.json"),
        "probe_sha256": c.sha(P / "build-0/probe.json"),
        "build_sha256": c.sha(P / "build.json"),
        "commands": {
            "nm": nm_command,
            "nm_demangled": demangled_command,
            "objdump": command_rows,
        },
        "nm": c.artifact(nm),
        "nm_log": c.artifact(nm_err),
        "nm_demangled": c.artifact(demangled),
        "nm_demangled_log": c.artifact(demangled_err),
        "phase_nm": c.artifact(phase_nm),
        "phase_nm_demangled": c.artifact(phase_demangled),
        "assembly": assembly_paths["edit_region"]["assembly"],
        "assembly_log": assembly_paths["edit_region"]["log"],
        "phase_assemblies": {
            label: {
                "assembly": value["assembly"],
                "log": value["log"],
            }
            for label, value in assembly_paths.items()
            if label != "edit_region"
        },
    }
    c.write(out / "symbol.json", result)
    return result


def deterministic_gzip(raw: bytes) -> bytes:
    output = io.BytesIO()
    with gzip.GzipFile(filename="", mode="wb", fileobj=output, compresslevel=9, mtime=0) as stream:
        stream.write(raw)
    stored = output.getvalue()
    assert gzip.decompress(stored) == raw
    return stored


def compress_member(
    path: Path, original: dict[str, object], *, repeat: int, kind: str,
) -> dict[str, object]:
    stored_path = path.with_name(path.name + ".gz")
    assert not stored_path.exists(), f"refusing to overwrite {stored_path}"
    raw = path.read_bytes()
    assert hashlib.sha256(raw).hexdigest() == original["sha256"]
    stored_path.write_bytes(deterministic_gzip(raw))
    return {
        "repeat": repeat,
        "kind": kind,
        "label": f"{repeat}.{kind}",
        "original": original,
        "compressed": c.artifact(stored_path),
        "compression": "gzip",
        "gzip_mtime": 0,
        "decompressed_sha256": hashlib.sha256(gzip.decompress(stored_path.read_bytes())).hexdigest(),
    }


def main() -> None:
    symbols_only = sys.argv[1:] == ["--symbols"]
    assert not sys.argv[1:] or symbols_only, "usage: decode.py [--symbols]"
    if symbols_only:
        static = static_owner_evidence()
        c.write(SYMBOLS / "complete.json", {
            "schema": "litchi.performance.0828.symbols.complete.v1",
            "status": "pass",
            "symbol": c.artifact(SYMBOLS / "symbol.json"),
            "binary": FP,
            "owners": static["owners"],
            "phase_owners": static["phase_owners"],
            "source": static["source"],
            "build": static["build"],
            "plan_sha256": c.sha(P / "plan.json"),
        })
        c.stable(FROZEN)
        print("0828 symbols PASS: exact edit and phase owners bound before timing", flush=True)
        return
    assert PERF.is_dir(), "run capture.py perf before decode.py"
    assert not (PERF / "decode-complete.json").exists(), "refusing to overwrite decode result"
    symbol_path = SYMBOLS / "symbol.json"
    assert symbol_path.is_file(), "run decode.py --symbols before decode"
    static = c.read(symbol_path)
    assert static["owner"] == EXPECTED_OWNER and static["binary"] == FP
    assert static["phase_owners"] == PLAN["probe"]["phase_owners"]
    complete = c.read(PERF / "complete.json")
    assert complete["schema"] == "litchi.performance.0828.perf.complete.v1"
    if complete["status"] == "unavailable":
        c.write(PERF / "decode-complete.json", {
            "schema": "litchi.performance.0828.decode.complete.v1",
            "status": "unavailable",
            "typed_unavailable": True,
            "profile_fabricated": False,
            "owner": EXPECTED_OWNER,
            "phase_owners": PLAN["probe"]["phase_owners"],
            "reason": complete.get("reason", "perf capture unavailable"),
            "symbol": c.artifact(SYMBOLS / "symbol.json"),
            "preflight": complete.get("preflight"),
            "plan_sha256": c.sha(P / "plan.json"),
            "build_sha256": c.sha(P / "build.json"),
        })
        c.stable(FROZEN)
        print("0828 decode unavailable: retained typed perf denial and static owner evidence", flush=True)
        return

    receipts_path = PERF / "receipts.json"
    receipts = c.read(receipts_path)
    assert isinstance(receipts, list) and len(receipts) == PLAN["perf"]["reports"]
    rows: list[dict[str, object]] = []
    compression: list[dict[str, object]] = []
    owner_counts: list[dict[str, object]] = []
    for receipt in receipts:
        repeat = receipt["repeat"]
        raw = Path(receipt["raw"]["path"])
        assert raw.is_file() and c.artifact(raw) == receipt["raw"]
        frames = PERF / f"{repeat}.frames"
        log = PERF / f"{repeat}.decode.log"
        command = ["perf", "script", "--no-inline", "--ns", "-i", str(raw)]
        started = time.time()
        exit_code = command_output(command, frames, log)
        row: dict[str, object] = {
            "schema": "litchi.performance.0828.decode-receipt.v1",
            "repeat": repeat,
            "command": command,
            "started": started,
            "ended": time.time(),
            "exit_code": exit_code,
            "raw": receipt["raw"],
            "frames": c.artifact(frames) if frames.exists() else None,
            "log": c.artifact(log),
            "binary": FP,
            "binaries": BINARIES,
            "symbol": c.artifact(SYMBOLS / "symbol.json"),
        }
        rows.append(row)
        c.write(PERF / "decode-receipts.json", rows)
        assert exit_code == 0 and frames.is_file() and frames.stat().st_size > 0
        frame_data = frames.read_bytes()
        text = frame_data.decode(errors="replace")
        owner_count = text.count(EXPECTED_OWNER)
        phase_owner_counts = {
            label: text.count(owner)
            for label, owner in static["phase_owners"].items()
        }
        owner_counts.append({
            "repeat": repeat,
            "frames_bytes": frames.stat().st_size,
            "owner_text_hits": owner_count,
            "owner_present": owner_count > 0,
            "phase_owner_text_hits": phase_owner_counts,
            "phase_owners_present": all(count > 0 for count in phase_owner_counts.values()),
        })
        stable = c.stable(FROZEN)
        del stable
    assert all(row["owner_present"] for row in owner_counts), "owner absent from decoded frames"

    for row in rows:
        repeat = row["repeat"]
        raw = Path(row["raw"]["path"])
        frames = Path(row["frames"]["path"])
        compression.append(compress_member(raw, row["raw"], repeat=repeat, kind="raw"))
        compression.append(compress_member(frames, row["frames"], repeat=repeat, kind="frames"))
    c.write(PERF / "compression.json", compression)
    for member in compression:
        original = Path(member["original"]["path"])
        stored = Path(member["compressed"]["path"])
        restored = gzip.decompress(stored.read_bytes())
        assert len(restored) == member["original"]["bytes"]
        assert hashlib.sha256(restored).hexdigest() == member["original"]["sha256"]
        assert c.artifact(original) == member["original"]
        original.unlink()
    c.write(PERF / "frame-owner-counts.json", owner_counts)
    c.write(PERF / "decode-complete.json", {
        "schema": "litchi.performance.0828.decode.complete.v1",
        "status": "available",
        "typed_unavailable": False,
        "profile_fabricated": False,
        "reports": len(rows),
        "frames": len(rows),
        "owner": EXPECTED_OWNER,
        "phase_owners": PLAN["probe"]["phase_owners"],
        "owner_counts": c.artifact(PERF / "frame-owner-counts.json"),
        "symbol": c.artifact(SYMBOLS / "symbol.json"),
        "receipts": c.artifact(PERF / "decode-receipts.json"),
        "compression": c.artifact(PERF / "compression.json"),
        "command_contract": ["perf", "script", "--no-inline", "--ns"],
        "raw_and_frames_retained": True,
        "retention": "lossless deterministic gzip; original descriptors bind decompressed bytes",
        "uncompressed_copies_removed": True,
        "plan_sha256": c.sha(P / "plan.json"),
        "build_sha256": c.sha(P / "build.json"),
    })
    c.stable(FROZEN)
    print("0828 decode complete: owner/phase symbol evidence and retained raw/gzip artifacts", flush=True)


if __name__ == "__main__":
    main()
