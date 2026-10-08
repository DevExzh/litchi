# Batch 0971 process-instruction attribution proposal

Status: read-only design review. Current checkout HEAD was revalidated as
`29f8c61d51b1c81f4f2afe6b9f61984e6347a6a6`. No production source, harness
source, build, benchmark, or native workload was changed or launched. The
current HEAD is a custody observation only; the proposed experiment must use
the retained historical binary below.

## Finding from retained 0970 evidence

The retained 0970 capture is 96 serial children across four cells, each using
five warmups and 80 measured samples. Its manifest is
`/home/zhuhe/litchi-performance-0970/capture/diagnostic-v3/capture-manifest-v2.json`
(SHA-256
`b0943c4572304dccc3a70dde17cce3894a35535e4c52fb7efb2526eb760b8685`).
The historical executable is

```text
/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile
SHA-256 0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb
BuildID e5c457384eeebd057932c444efda77618f137399
```

For the pinned `docfloat/e1/default` children, the timed p50 range is
`978289`–`1270130.5 ns`; whole-process `instructions` range is
`6987088875`–`7494203611`, with instruction-count CV `2.737%`. Their
p50-to-instruction Pearson/Spearman associations are `0.9553/0.8811`. The
pinned `docnohf/e1/default` control has instruction-count CV `0.052%` and
p50 range `115170.5`–`117785.5 ns`; its instruction association is weak and
negative (`-0.1581/-0.2587`). These values are descriptive evidence from
`/home/zhuhe/litchi-performance-0970/review/raw-independent-scalars-0970-v1.json`
(SHA-256
`c9bca1f01c575678f20632e7ecdd7ded0075e7baf7094523b22914f11abfcd70`).

The 0970 perf sidecars measured the launched `/usr/bin/time`/Python launcher
and inherited baseline process tree, while the harness `elapsed_ns` measured
only the operation-local region. Therefore the existing association does not
locate the instruction difference in the timed operation. A direct `perf
record` around the retained executable can make that attribution testable.

## Why the retained binary is usable without a rebuild

The retained ELF is PIE, `with debug_info, not stripped`; it has `.symtab`,
`.debug_info`, `.debug_line`, `.eh_frame`, and valid unwind entries. The exact
symbol and source checks are:

- `doc_retention_probe_0942_cfb_profile::profiled_measured_case` is a physical
  function at `0xec4f0..0xec56d`, with DWARF declaration at historical
  `tools/perf-cfb-profile/src/lib.rs:338`.
- `doc_retention_probe_0942_cfb_profile::run_case` is at
  `0xf15f0..0xf26f4`, with DWARF declaration at `lib.rs:290`.
- `addr2line -f -C` resolves the marker, `run_case`,
  `litchi_doc::body_text::Snapshot::open_bounded`, and
  `RevisionEditor::finish_unchecked` to the retained historical source tree.
- The historical harness source is
  `/home/zhuhe/litchi-performance-0969/source-baseline/tools/perf-cfb-profile/src/lib.rs`
  (SHA-256
  `b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc`), and
  its Cargo manifest is unchanged at SHA-256
  `41dee12e4eb7d1e3853a33a7743afdf6b5d6519d5e7456e61ce4978aa9177945`.

`profiled_measured_case` is the existing unique `#[inline(never)]` marker. It
calls `run_case`, whose body covers the source clone, bounded snapshot open,
edit/replacement loop, retention releases, commit, output `Vec` copy, and
local destruction. `timed_case` starts the clock immediately before the
marker and stops immediately after it (`lib.rs:353-358`). Preparation,
correctness routes, per-sample validation, hashing, and JSON serialization are
outside that clock. The ordinary 0942 manifest has no enabled
`performance-diagnostics` feature, so there are no phase callbacks in this
binary; the physical marker is the only exact boundary selector and is
sufficient for operation-level attribution.

Both warmups and measured samples call the marker. The retained CLI therefore
has **85 marker calls** per child (`5` warmups plus `80` measured), not 80.
The proposal keeps the historical `--warmup 5 --samples 80` contract and
labels marker samples as the combined 85-call marker scope. Separating warmup
from measured calls would require a harness change and is outside this
experiment.

## Harmless profiler capability check

The fixed `instructions:u` and `cycles:u` events were recorded with DWARF
callgraphs against `/usr/bin/true`; both returned zero, produced five samples,
and had zero lost samples. DWARF 8192 and 16384 byte smoke captures also
returned zero and decoded five samples. The receipt is
`/home/zhuhe/litchi-performance-0971/receipts/perf-record-capability-probe-v1.json`
(SHA-256
`50bc1b5c3253ed44cc3b1516e258a6028db40a3f049336814264365b539141f2`). It
records `/usr/bin/perf` SHA-256
`174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b`,
`perf_event_paranoid=1`, `perf_event_max_sample_rate=25000`, and
`perf_event_max_stack=127`.

This smoke check does not validate marker recovery on the retained binary. A
first exact-binary pilot must measure marker recovery and stop if the result
is not sufficient. The prior 8192-byte captures are known to truncate a
roughly 10.9 KiB timed stack; use `dwarf,16384` for the bounded pilot and
report any remaining unresolved/truncated portion. Do not silently replace
unresolved samples with qualified attribution.

## Smallest useful root-GO experiment

Use 12 fresh, serial, CPU-31-pinned children of the retained binary for the
noisy `docfloat/e1/default` cell:

```text
/usr/bin/taskset --cpu-list 31 /usr/bin/perf record \
  -e instructions:u -c 5000000 --call-graph dwarf,16384 \
  --period --timestamp --sample-cpu --buildid-all --no-buildid-cache \
  --user-callchains -o <child>.perf.data -- \
  /home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile \
  --case docfloat \
  --input /home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc \
  --edits 1 --retention default --warmup 5 --samples 80 \
  > <child>.report.json 2> <child>.perf.stderr
```

The fixture is 335,360 bytes with SHA-256
`41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`. The
period gives roughly 1,400 instruction samples per seven-billion-instruction
child before filtering, with an expected few hundred marker samples; it is
bounded and avoids a frequency claim. If reviewers require a low-variation
control, add exactly 12 CPU-31-pinned `docnohf/e1/default` children using the
same command and the 26,112-byte fixture SHA-256
`45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48`. Do not
expand beyond 12 target children or 24 total children.

The first target child is a pilot and remains part of the 12-child set. Before
continuing, require a predeclared sufficient-marker gate, such as at least
100 samples with the exact physical marker in that child, no lost perf events,
and a reported weighted marker-period share. If the gate fails, preserve the
pilot and stop; do not choose a different child or selectively rerun a fast or
slow process.

Keep one raw `perf.data`, stderr, stdout JSON, perf header/build-ID export,
and exact command/affinity/binary/fixture hashes per child. Decode both forms:

```text
/usr/bin/perf script --show-lost-events --no-inline \
  -F cpu,pid,tid,time,event,period,ip,sym,symoff,dso -i <child>.perf.data
/usr/bin/perf script --show-lost-events --inline -i <child>.perf.data
```

The physical parser must select the exact fully qualified
`doc_retention_probe_0942_cfb_profile::profiled_measured_case` frame. It must
report total samples/period, marker samples/period, deepest resolved
instruction, unresolved frames, stack truncation, and lost events. Inline
source annotations may refine the owner line but cannot substitute for the
physical marker. A child with no marker, ambiguous marker, or material
unresolved/truncated remainder is descriptive only.

For every child, parse the harness report's timed p50 and associate it with
all marker samples from that child. Rank exclusive deepest-resolved
instruction owners separately from inclusive marker-ancestor tags; do not add
overlapping inclusive percentages. Compute descriptive Pearson/Spearman
associations across all 12 children for p50 versus marker-period totals and
per-owner marker-period shares. Analyze every child; do not select fast/slow
tails for a second run.

## Interpretation and limits

A repeatable shift in marker-owner instruction shares that tracks the timed
p50 across pinned fresh processes would support process-specific instruction
path variation as a contributor to the 0970 spread. A stable marker profile
with varying p50 would leave scheduling, cache, allocator, or frequency
variation unresolved. The capture cannot identify a hash seed directly; it can
only show whether fresh processes execute different source-resolved paths.

The direct profiler changes overhead and is diagnostic evidence only. It is
not latency-comparable to 0969/0970 and cannot support a speedup or slowdown
claim. The retained historical binary remains byte-identical, no forced-frame-
pointer or instrumented rebuild is needed, and no current-HEAD runtime is
admitted. If exact-binary marker recovery is inadequate, the result is
`attribution unsupported` and a separately reviewed diagnostic build would be
required; do not infer a source path from whole-process counters alone.
