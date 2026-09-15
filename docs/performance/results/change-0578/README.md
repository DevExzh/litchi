# change-0578 evidence

Evidence for [0578](../../0578-zip-passthrough-is-already-bounded.md), a refuted
hypothesis. `performance_claim: none`. Peak-retained-byte and allocation counts
only; no timing was measured and none is claimed.

## What was measured

Peak retained heap bytes and allocation counts for a ZIP save that passes
unchanged members from a positional `FileReader` source to a sequential,
non-seek sink. Three scenarios over the same synthetic package (21 small
deflated XML members plus one large `word/media/big.bin`):

- `A_copy_through` — `PreservationAction::Copy` for every member with one small
  member `Regenerate`d. The production save shape.
- `A_copy_all` — the same with no edited member, isolating the index term.
- `B_precompressed_token` — capture the large member as a
  `VerifiedPrecompressedEntry` and republish it via
  `RegeneratedEntry::new_precompressed_shared`. The cross-document copy shape.
- `C_write_precompressed_slice` — `ZipArchiveWriter::write_precompressed_file`,
  fed by a caller that materializes the member's compressed range itself,
  because `VerifiedPrecompressedEntry::compressed_payload` is `pub(crate)` and
  nothing else can reach the entry point.

## Files

| file | contents |
| --- | --- |
| `environment.txt` | revision, toolchain, CPU, memory, allocator identity |
| `stage1-peak-by-size.csv` | axis 1 — peak against largest-member size, all three scenarios, Store and Deflate |
| `stage1-peak-by-size-copyall.csv` | axis 1b — the same series with a pure copy-all plan |
| `stage1-peak-by-count.csv` | axis 2 — peak against member count at a fixed 4 KiB member size |
| `stage1-peak-real-fixtures.csv` | axis 3 — five real `test-data` fixtures copied through verbatim |
| `stage1-peak-zip64.csv` | axis 4 — one 4,362,076,160-byte stored member, past the ZIP64 boundary |
| `zip64-fixture.sha256` | SHA-256 of the axis-4 archive, which was deleted after measurement |
| `peak-bytes-probe.rs` | the probe source, including its counting global allocator |
| `peak-bytes-probe.Cargo.toml` | the probe manifest; point the path dependency at a worktree of the revision below |

CSV columns are the scenario inputs followed by `region_peak_bytes`,
`allocations`, `allocated_bytes`, `output_bytes`. `region_peak_bytes` is the
high-water mark of live heap bytes inside the measured region, rebased to the
live total at region entry.

## Reproducing

Measured against a **detached git worktree** of
`32d25e08806d93f792ffd4954d83acc9db9c5301` with an isolated `CARGO_TARGET_DIR`,
so no concurrently mutated tree could affect the result.

```sh
git worktree add --detach /tmp/w 32d25e08806d93f792ffd4954d83acc9db9c5301
mkdir -p /tmp/probe/src
cp peak-bytes-probe.rs        /tmp/probe/src/main.rs
cp peak-bytes-probe.Cargo.toml /tmp/probe/Cargo.toml
# point the soapberry-zip path dependency at /tmp/w/crates/soapberry-zip
CARGO_TARGET_DIR=/tmp/probe-target cargo build --release --manifest-path /tmp/probe/Cargo.toml -j 8

P=/tmp/probe-target/release/zip-passthrough-probe
"$P" sizes         /tmp/work
"$P" sizes-copyall /tmp/work
"$P" counts        /tmp/work
"$P" fixture       /tmp/work <absolute fixture paths…>
"$P" zip64         /tmp/big      # writes and reads a 4.06 GiB archive
```

The axis-4 run needs about 4.1 GiB of free space for the generated archive. The
probe removes every synthetic archive it creates except the ZIP64 one, which was
hashed and deleted by hand.

## Caveats

The counting allocator is callback-ordered heap accounting from a
single-threaded probe: it excludes allocator-internal fragmentation, RSS, and
the page cache holding the source file. The probe is scratchpad tooling and is
not part of the workspace build; it contains the only `unsafe` involved, in the
`GlobalAlloc` wrapper, which mirrors
`tools/perf-baseline/src/bin/support/counting_allocator.rs`. `soapberry-zip`'s
own `deny(unsafe_code)` is untouched.

## End-to-end files (axis 5)

| file | contents |
| --- | --- |
| `stage1-peak-e2e-opc.csv` | axis 5 — both production `litchi-opc` save doors against a growing unchanged media member, with an FNV-1a digest of each output stream |
| `opc-save-peak-probe.rs` | the end-to-end probe source |
| `opc-save-peak-probe.Cargo.toml` | its manifest |

```sh
mkdir -p /tmp/e2e/src
cp opc-save-peak-probe.rs         /tmp/e2e/src/main.rs
cp opc-save-peak-probe.Cargo.toml /tmp/e2e/Cargo.toml
# point the litchi-opc and litchi-core path dependencies at /tmp/w/crates/…
CARGO_TARGET_DIR=/tmp/e2e-target cargo build --release --manifest-path /tmp/e2e/Cargo.toml -j 8
/tmp/e2e-target/release/opc-save-peak-probe \
    <repo>/test-data/ooxml/docx/drawing.docx /tmp/work-e2e
```

`output_fnv1a64` is a digest of the complete output byte stream, folded in the
sink so it allocates nothing. Matching digests between the `A_package_writer` and
`B_source_backed` rows for one `media_bytes` value mean the two production save
paths emitted identical archives.
