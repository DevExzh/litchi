"""Root-only perf decode and deterministic gzip custody.

The raw perf files are decoded while all three exact executables remain in the
owned target. Both the binary perf data and the resulting non-inline frame
text are retained as deterministic gzip members with their uncompressed and
compressed identities recorded.
"""

import gzip
import hashlib
import io
import subprocess
import time
from pathlib import Path

import custody as c


P = c.P
OUT = P / "perf"
assert OUT.is_dir(), "run capture.py perf before decode.py"
assert not (OUT / "decode-receipts.json").exists(), "refusing to overwrite decode receipts"
assert not (OUT / "compression.json").exists(), "refusing to overwrite compression receipt"
PLAN = c.read(P / "plan.json")
BUILD = c.read(P / "build/build.json")
FROZEN = c.read(BUILD["frozen_inputs"]["path"])
PERF = c.read(OUT / "receipts.json")
assert c.frozen_driver_hashes() == FROZEN["drivers"]
assert len(PERF) == PLAN["perf"]["reports"] == 2
assert PLAN["perf"]["event"] == "cycles:u"
assert PLAN["perf"]["frequency_hz"] == 499
assert PLAN["perf"]["call_graph"] == "fp"
assert PLAN["perf"]["owner"] == "namespace_uri_probe::capture_region_0793"

source = c.source()
assert source["revision"] == PLAN["source_revision"]
assert len(source["files"]) == PLAN["source_file_count"]
probe = c.assert_probe(BUILD["probe"])
root_inputs = c.assert_root_inputs()
architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()

for variant, binary in BUILD["binaries"].items():
    assert c.artifact(binary["path"]) == binary


def stable() -> None:
    assert c.source() == source
    assert c.assert_probe(probe) == probe
    assert c.assert_root_inputs() == root_inputs
    assert c.architecture_hashes() == architecture
    assert c.assert_unrelated() == unrelated
    for binary in BUILD["binaries"].values():
        assert c.artifact(binary["path"]) == binary


def deterministic_gzip(raw: bytes) -> bytes:
    output = io.BytesIO()
    with gzip.GzipFile(
        filename="", mode="wb", fileobj=output, compresslevel=9, mtime=0
    ) as stream:
        stream.write(raw)
    stored = output.getvalue()
    assert gzip.decompress(stored) == raw
    return stored


rows = []
logical = []
frame_receipts = []
for receipt in PERF:
    repeat = receipt["repeat"]
    binary = receipt["binary"]
    assert binary == BUILD["binaries"][PLAN["perf"]["binary"]]
    assert c.artifact(binary["path"]) == binary
    # Keep this spelling explicit so every path comes from the immutable
    # capture receipt; no glob can silently select a different perf file.
    raw = Path(receipt["raw"]["path"])
    assert raw.is_file() and c.artifact(raw) == receipt["raw"]
    frames = OUT / f"{repeat}.frames"
    log = OUT / f"{repeat}.decode.log"
    assert not frames.exists() and not log.exists()
    command = ["perf", "script", "--no-inline", "--ns", "-i", str(raw)]
    started = time.time()
    with frames.open("w") as stream, log.open("w") as errors:
        result = subprocess.run(
            command,
            cwd=c.ROOT,
            stdout=stream,
            stderr=errors,
        )
    row = {
        "schema": "litchi.performance.0811.decode-receipt.v1",
        "repeat": repeat,
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "binary": binary,
        "raw": receipt["raw"],
        "frames": c.artifact(frames) if frames.exists() else None,
        "log": c.artifact(log),
        "source": c.artifact(P / "build/source.json"),
        "probe": probe,
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
    }
    rows.append(row)
    c.write(OUT / "decode-receipts.json", rows)
    assert result.returncode == 0, log
    assert frames.is_file() and frames.stat().st_size > 0
    row["frames"] = c.artifact(frames)
    c.write(OUT / "decode-receipts.json", rows)
    stable()
    print("decode", repeat, "PASS", flush=True)

for receipt in rows:
    repeat = receipt["repeat"]
    raw = Path(receipt["raw"]["path"])
    frames = Path(receipt["frames"]["path"])
    for kind, path, original in (
        ("raw", raw, receipt["raw"]),
        ("frames", frames, receipt["frames"]),
    ):
        assert path.is_file()
        uncompressed = path.read_bytes()
        assert hashlib.sha256(uncompressed).hexdigest() == original["sha256"]
        stored_path = path.with_name(path.name + ".gz")
        assert not stored_path.exists()
        stored = deterministic_gzip(uncompressed)
        stored_path.write_bytes(stored)
        compressed = c.artifact(stored_path)
        assert gzip.decompress(stored_path.read_bytes()) == uncompressed
        member = {
            "repeat": repeat,
            "kind": kind,
            "original": original,
            "compressed": compressed,
            "compression": "gzip",
            "gzip_mtime": 0,
            "decompressed_sha256": hashlib.sha256(
                gzip.decompress(stored_path.read_bytes())
            ).hexdigest(),
        }
        logical.append(member)
        if kind == "frames":
            frame_receipts.append(member)
        path.unlink()

c.write(OUT / "compression.json", logical)
c.write(OUT / "frame-receipts.json", frame_receipts)
c.write(
    OUT / "decode-complete.json",
    {
        "schema": "litchi.performance.0811.decode.complete.v1",
        "reports": len(rows),
        "logical_artifacts": len(logical),
        "receipts": c.artifact(OUT / "decode-receipts.json"),
        "compression": c.artifact(OUT / "compression.json"),
        "frames": c.artifact(OUT / "frame-receipts.json"),
        "plan_sha256": c.sha(P / "plan.json"),
        "build_sha256": c.sha(P / "build/build.json"),
        "command_contract": ["perf", "script", "--no-inline", "--ns"],
        "raw_and_frames_retained_as": "deterministic gzip members",
    },
)
stable()
print("0811 perf decode complete", flush=True)
