# Change 0416 evidence

This bundle covers small entries with ZIP64 local framing and ordinary central
sizes, including Python's `force_zip64` output. It is a compatibility change
with performance guards, not a speedup claim. Control is `4794d2e5e`; the
candidate and exact source identities are recorded with the completed capture.

## Corpus

`interop.py` uses Python's standard-library ZIP writer with deterministic
timestamps, names and payloads. It generates sequential and seekable output,
then makes explicitly identified unsigned-descriptor and central-ZIP64 variants.
The unsigned fixture includes an actual payload whose CRC equals `0x08074b50`;
the marker cannot safely be interpreted from its first four bytes alone.
`manifest.json` records archive hashes, member sizes, CRCs, payload hashes and
local/central/descriptor framing. Verification independently drains every
payload using Python and zlib.

```sh
python3 docs/performance/results/change-0416/interop.py verify
```

To regenerate without modifying retained evidence, select a fresh output
directory using `generate --output-dir PATH`, then compare its archives and
manifest. The Python and zlib versions are recorded because compression output
may change across versions. `native-fixtures.json` binds three existing tracked
LibreOffice-resaved artifacts; these are archive guards and do not represent a
new native-application verification session.

## Timed operations

The standalone `probe.rs` uses the same source and dependency lock for both
revisions. Input file reads and the central-metadata oracle are prepared before
the timer. Each sample constructs and drops a fresh archive/index. Per-member
metadata comparisons and the observation digest are inside the timer.

- `borrowed` constructs `ArchiveReader` and attempts `read_stored_borrowed` for
  every file. At least one Store entry is required. Store access validates local
  framing and overlap and scans the stored payload for CRC verification; it
  does not decompress Deflate members.
- `indexed` constructs `IndexedArchive` over the retained byte slice through
  `ReaderAt`, allocates its operation scratch and builds a `PreservationIndex`.
  This validates local headers and descriptors without decoding member bodies.
  Tail and metadata ranges can still include payload bytes incidentally.
- `count` is a separate indexed observation with explicit call/accepted-byte
  counters. Its diagnostic elapsed time is excluded from normal latency rows.
- `capability` records admission/refusal separately. Its elapsed field includes
  setup and is never compared as successful-work latency.

`capture.py` serializes fresh CPU-2 processes in A1/B1/B2/A2 order. Each ordinary
row retains 300 samples after 30 warmups and a process `/usr/bin/time -v` report.
Indexed-only fixtures cover the existing native artifacts. Local-only inputs
are capability cases: a control refusal is not a zero-time baseline.
`summarize.py` recomputes exact medians and nearest-rank p95/p99 values from raw
vectors, checks observations, and retains candidate regressions separately from
within-revision drift. Positive changes above 5% are review triggers.

The machine is a shared KVM guest. Affinity and task-workload serialization do
not control unrelated background activity. This is warm in-memory evidence,
not cold-cache, physical-storage, remote-source, concurrency or scaling evidence.
Process RSS includes retained source and oracle allocations; it is not an
operation-local allocation measurement.

## Fresh build

Use clean detached worktrees at the recorded revisions. Create an external
Cargo package with an empty `[workspace]`, edition 2021, package name
`zip64-stream-probe`, version `0.0.0`, and these dependencies:

```toml
[dependencies]
soapberry-zip = { path = "/absolute/selected-tree/crates/soapberry-zip" }
crc32fast = "1"
serde_json = "1"
```

Declare a `zip64-read-probe` binary using the retained `probe.rs`; use the
retained probe lockfile. Build each revision in sequence:

```sh
env CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=1 \
  RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes' \
  cargo +1.98.1 build --release --locked --manifest-path /absolute/probe/Cargo.toml
```

Copy the executable to a distinct path before changing the selected dependency
and rebuilding. `capture.py --help` documents fixture, executable and worktree
arguments. Exact capture argv and artifact identities are retained in the
capture manifest. Serialize measurements against task-owned builds, tests,
profiling and analysis workloads.

## Guard results and review

Control `4794d2e5e` and candidate `d18cd04a2` use identical probe source and
locks. Both captures contain 11 rows and 44 processes; the follow-up uses
1,000 samples after 100 warmups. All per-row metadata observations agree.
All seven local-only capability fixtures refuse on control and succeed on
candidate in both modes, in each capture. No refusal latency enters a ratio.

| Fixture / mode | A1 → B1 p50 (µs) | A2 → B2 p50 (µs) | Paired p50 change |
| --- | ---: | ---: | ---: |
| central64-many256 / borrowed | 1277.736 → 1246.600 | 1278.815 → 1246.465 | -2.44% / -2.53% |
| central64-many256 / indexed | 84.680 → 86.070 | 85.760 → 85.710 | +1.64% / -0.06% |
| central64-tiny / borrowed | 0.650 → 0.640 | 0.630 → 0.670 | -1.54% / +6.35% |
| central64-tiny / indexed | 1.710 → 1.710 | 1.710 → 1.700 | +0.00% / -0.58% |
| zip32-many256 / borrowed | 1140.300 → 1179.650 | 1136.380 → 1183.864 | +3.45% / +4.18% |
| zip32-many256 / indexed | 81.710 → 83.940 | 84.490 → 81.950 | +2.73% / -3.01% |
| zip32-tiny / borrowed | 0.610 → 0.620 | 0.610 → 0.620 | +1.64% / +1.64% |
| zip32-tiny / indexed | 1.630 → 1.640 | 1.630 → 1.640 | +0.61% / +0.61% |
| native-docx / indexed | 4.520 → 4.570 | 4.510 → 4.600 | +1.11% / +2.00% |
| native-pptx / indexed | 17.480 → 17.820 | 17.540 → 17.800 | +1.95% / +1.48% |
| native-xlsx / indexed | 5.380 → 5.260 | 5.210 → 5.300 | -2.23% / +1.73% |

The original capture flags one candidate p99 increase, +5.88% for tiny
central-ZIP64 borrowed access. The follow-up flags the same row in the second
pair: p50 630 → 670 ns (+6.35%) and p95 660 → 700 ns (+6.06%).
No other follow-up row crosses the 5% review threshold. Both summaries retain
every p50/p95/p99/RSS comparison and within-revision drift. There is no clean
latency-regression pass and no speedup claim. Follow-up process RSS spans
2,176–2,760 KiB; no RSS comparison crosses 5%. Retain this necessary input and
preservation capability with the small central-ZIP64 tail/median trigger open;
the next performance investigation should isolate that sub-microsecond path
before changing its implementation. No aggregate hides this row.

## Resource diagnostics

All seven matched indexed fixtures have identical logical ReaderAt call and
accepted-byte counts between revisions. ZIP32 tiny uses 14 calls / 650 bytes;
ZIP32 many256 uses 1,030 / 76,890. Candidate local-only many256 uses 1,030 /
84,058. These are metadata-range observations over memory; reads can include
payload bytes incidentally, and they are not physical device or remote calls.

Separate whole-process heaptrack runs on ZIP32 many256 indexed access include
100 samples, 10 warmups, source/oracle setup and JSON output. Both revisions
record 143,646 allocation calls and 530.62 KB peak heap. `heaptrack_print`
reports 28,167 temporary allocations; raw heaptrack stderr reports 28,166.
Both sources are retained, and this one-count provenance difference is not
used as an optimization result. Heaptrack reports 544 retained bytes at process exit for each; that
number is not an object-lifetime leak diagnosis. Its RSS includes instrumentation
and is excluded from normal RSS comparisons. These totals do not establish
operation-local allocation attribution or local-only input memory behavior.

Separate CPU/PMU processes each run 20,000 samples and 100 warmups on the same
ZIP32 many256 indexed fixture. Control/candidate sampling records 337/339
samples with zero lost samples. The retained reports and flame graphs identify
name normalization, preservation indexing and copying as active work. These
short paired diagnostics do not attribute the tiny borrowed guard trigger or
establish a statistically supported CPU improvement. Exact cycles, instructions,
branches, branch misses and page faults remain in the two stat CSVs. Control
and candidate instructions total 33,968,036,179 and 34,955,721,524 (+2.91%);
cycles total 7,603,116,888 and 7,505,690,538. These single whole-process
captures are descriptive, not a hardware-counter improvement claim.

## Validation and limitations

The final ZIP/OPC/ODF all-feature/all-target suite passes 1,246 tests. The 12
ODT embedded-resource GC integration tests and its new local-framing unit test
pass. Both explicitly selected large ZIP64 tests pass; 43 doctests pass and two
are ignored. Scoped warning-denied Clippy and four-crate warning-denied rustdoc
pass. Clippy retains the five existing command-level allowances documented in
0415. Changed-file formatting passes; full workspace formatting reports only
unchanged `litchi-docx/tests/glossary_authoring.rs`. Boundaries, CRUD coverage,
eight strict registered claims and report classification pass. Initial compile,
fixture and offset failures are retained alongside successful final checks.

ZIP and OPC each complete 1,000 AddressSanitizer/sanitizer-coverage fuzz runs
from the 13 fixtures. This is bounded smoke coverage. The generator reproduces
all fixture bytes and its manifest exactly. `verify.py` checks source/binary
capture bindings, corpus hashes, run order and argv, raw-vector summaries,
diagnostic hashes/observations and required check results.

Standalone export replay passes without the temporary binaries or worktrees.
Four deliberate source, observation, capture-argv and binary-binding corruptions
are rejected; restoring the files passes again. See `checks/replay.json`.

Task worktrees, executables, generated large input, fuzz corpora/targets and
created fuzz lockfiles were removed. Shared targets and the user goal file
were preserved. Post-cleanup verification passes; `checks/cleanup.json` records
the exact removals. `SHA256SUMS` inventories every retained bundle file except
itself and can be checked with `sha256sum -c SHA256SUMS` inside this directory.
