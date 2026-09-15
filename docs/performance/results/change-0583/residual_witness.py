#!/usr/bin/env python3
"""Build the witnesses that price change 0583's two residuals.

Change 0583 bounds a neighbour's payload with `max(local, central)` when the
probe reads that neighbour's fixed local header.  Two shapes keep a
local-versus-central payload differential open, and each gets one archive here
so the record's Limitations section is measured rather than argued.

  1. `residual-prune-distant-csize`  -- the predecessor is pruned by the
     zero-I/O central-directory bracket (`local_header_offset + 30 +
     central compressed_size + 131094`), so its local header is never read and
     the maximum never runs.  Closing this needs one 30-byte read per
     predecessor, which is the archive-wide cost change 0580 removed.

  2. `residual-zip64-sentinel-csize` -- the local `compressed_size` is the
     ZIP64 sentinel `0xFFFFFFFF` and the real length is in a local ZIP64 extra
     field.  The probe does not read the variable region, so it falls back to
     the central length.  Closing this needs a second, variable-length read.

  3. `residual-descriptor-flag-disagreement` -- the CENTRAL record declares a
     data descriptor and the LOCAL header does not.  The bound is a
     `Descriptor`, whose payload end is the offset the descriptor is read at,
     so it must stay where the central length puts it and the maximum is not
     applied; but a reader following local headers sees bit 3 clear and uses
     the inflated local size.  Closing this needs a bound that can refuse
     without resolving, which the current `LocalSpanBound` cannot express.

Usage: residual_witness.py <change-0582 results dir> <out dir>
"""
import sys, pathlib, importlib.util

spec = importlib.util.spec_from_file_location(
    "build_corpus", pathlib.Path(sys.argv[1]) / "build_corpus.py")
bc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bc)

out = pathlib.Path(sys.argv[2]); out.mkdir(parents=True, exist_ok=True)

# (1) The bracket is 30 + central csize + 131094 past the predecessor's offset.
#     Placing the target beyond it prunes the predecessor at zero I/O.
BRACKET_END = 30 + 16 + 2 * 65535 + 24            # 131140, predecessor at 0
TARGET = 200_000                                   # comfortably past it
assert TARGET > BRACKET_END
prune = bc.build_zip([
    bc.Member("pred.bin", b"P" * 16, local_csize=300_000),
    bc.Member("target.bin", b"T" * 16, local_offset=TARGET),
])
(out / "residual-prune-distant-csize.zip").write_bytes(prune)

# (2) The ZIP64 sentinel in the local compressed-size field, with a local ZIP64
#     extra field carrying the real 64-bit length.  Header id 0x0001, then
#     uncompressed then compressed size, both u64.
zip64_extra = (b"\x01\x00" + (16).to_bytes(2, "little")
               + (300_000).to_bytes(8, "little") + (300_000).to_bytes(8, "little"))
sentinel = bc.build_zip([
    bc.Member("pred.bin", b"P" * 16, local_csize=0xFFFFFFFF,
              local_usize=0xFFFFFFFF, local_extra=zip64_extra),
    bc.Member("target.bin", b"T" * 16, local_offset=200),
])
(out / "residual-zip64-sentinel-csize.zip").write_bytes(sentinel)

# (3) Central bit 3 set, local bit 3 clear, local compressed size inflated, and
#     a valid descriptor exactly where the central length puts it -- so the
#     resolve succeeds and nothing refuses.
payload = b"P" * 16
disagree = bc.build_zip([
    bc.Member("pred.bin", payload, descriptor="unsigned", local_flags=0x0000,
              local_csize=100_000, local_usize=100_000,
              local_crc=bc.zlib.crc32(payload) & 0xFFFFFFFF),
    bc.Member("target.bin", b"T" * 16, local_offset=200),
])
(out / "residual-descriptor-flag-disagreement.zip").write_bytes(disagree)

for path in sorted(out.iterdir()):
    print(f"{path.name} {len(path.read_bytes())} bytes")
