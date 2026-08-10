# XLSX commit reuses its bounded validated worksheet store

Date: 2026-08-11

Production base: `2c93ab84277ce073c67feed8c0622fa4d6f2d4ca`

Scope: owned XLSX worksheet commits followed by a public read of the changed
worksheet. OLE2, RTF and ODF production code are unchanged, and iWork/IWA
crates were explicitly excluded.

## Hypothesis

A changed worksheet commit compacts its final XML, parses that exact byte
sequence into an owned semantic `Store`, and validates every staged semantic
change and shared-style reference against it. Publication then constructs a
new `Workbook` whose worksheet `OnceLock<Store>` values are empty. The first
public `cell`, `cells`, `row`, `column`, or extent query reparses the same
published bytes into the same semantic form.

Moving the already validated store into the new snapshot should remove that
duplicate parse. It is safe only while the final part bytes and the style and
shared-string identity domains remain exactly the ones used by commit-time
validation.

## Change and bounded retention policy

Commit now retains a private candidate containing the changed worksheet URI,
the exact `Arc<Vec<u8>>` staged for publication, and the already validated
`Store`. After package construction and every existing calculation/web
integrity check succeeds, the new workbook adopts the store only when:

- source and target style lineages are pointer-identical;
- source and target shared-string lineages are pointer-identical;
- both source and target entries are worksheets at the same part URI; and
- the final published part blob is pointer-identical to the staged bytes.

A later active-tab, reference, or other composition rewrite replaces the part
`Arc` and therefore refuses adoption. Created and removed sheets are outside
this first tranche. No validation, error boundary, patch state, exact-source
authorization, save behavior, public API, dependency edge, runtime, lock,
durability rule, or unsafe-code boundary changes.

The initial unrestricted prototype won the dense-wide latency cell but failed
the memory gate: one-sample commit-only Heaptrack peak heap rose from 91.79 to
100.04 MiB (**+8.99%**) and profiler RSS rose 8.56%. That version was rejected.
The accepted policy hands off only stores with at most 4,096 stored cells and
at most 1 MiB of final worksheet XML. The generated medium target has 1,024
cells; the 65,536-cell dense-wide target deliberately takes the cold-cache
fallback.

## Dedicated public benchmark

The new opt-in `xlsx_one_cell_commit_first_read` case prepares a fixed
`Sheet1!A1` update outside timing, then times public commit plus the first
public read of that cell. Patch/value checks stay in the timed operation's
success path and a complete workbook semantic verification runs once outside
timing. Existing commit-only cases now retain only their final result for the
same post-timing verification.

The harness has 108 selectable cases while the 36-case / 198-record default
matrix is unchanged. CI runs a tiny smoke and a scheduled tiny/medium/dense-wide
attribution slice.

The common-harness baseline executable SHA-256 is
`eee2ed9b0173d3a409e209e20a16f21de54f7cb1e41e6589882d68cd9c25dc25`.
The final bounded executable SHA-256 is
`b4d65486b3176eb3c14bbc7dae84c3f16bf53da576d86af04739833b44896b31`.
Its measured `.text` section SHA-256 is
`a0740f7c1e52b6cdc2afb97ca767e2120a202f8cc55dd36aa158d01ed0339b33`.
The rejected unrestricted prototype executable was
`56bb7f3f8f61e3bb93472a32095b784617298fa941fcee0ba96f94808c834a7b`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic medium corpus has four 32-by-32
worksheets, 4,096 cells total, nine archive members, a 15,254-byte archive,
and archive SHA-256
`9574867b4f1ab4d30ce150de32d2a0b01267d15399ec9edd2c0d57b4bc60fab6`.

## Matched latency measurement

The primary ABBA run used 50 warmups and 500 samples per leg. Pooling 1,000
samples per state gives:

| Medium XLSX one-cell commit + first read | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 4.431 ms | 3.402 ms | **-23.23%** |
| p95 | 4.665 ms | 3.587 ms | **-23.11%** |
| p99 | 4.899 ms | 3.949 ms | **-19.39%** |
| mean | 4.460 ms | 3.427 ms | **-23.15%** |

The approximate independent-sample 95% interval for the mean delta is
`[-23.38%, -22.93%]` of the before mean. Both after legs have lower p50 and
mean than both before legs.

## Guardrails and rejected prototype

| Workload | p50 delta | p95 delta | p99 delta | Mean delta |
|---|---:|---:|---:|---:|
| Medium commit only | -1.90% | -2.26% | -0.62% | -1.80% |
| Medium commit + save | -1.95% | -5.16% | -13.09% | -2.85% |
| Dense-wide commit + first read, bounded fallback | -0.87% | -1.76% | -0.61% | -1.05% |
| Dense-wide unrestricted prototype | -26.62% | -25.79% | -27.26% | -26.51% |

The medium guards use 50 warmups and 500 samples per ABBA leg. The deliberately
excluded dense-wide guard uses five warmups and 20 samples per leg. The
unrestricted prototype used ten warmups and 50 samples per leg; its latency
result is retained only as rejected evidence because of the memory trigger.

## Allocations, memory, and CPU attribution

Matched Heaptrack processes used two warmups and 20 primary samples plus one
complete post-timing verifier:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 4,284,012 | 3,384,082 | **-21.01%** |
| Temporary allocations | 639,446 | 541,436 | **-15.33%** |
| Peak heap | 3.03 MiB | 3.16 MiB | +4.29% |
| Heaptrack RSS | 14.05 MiB | 14.25 MiB | +1.42% |
| Leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

The bounded retained cache stays below the 5% review threshold on the selected
workload. An independent one-sample commit-only profile reports the same
3.03/3.16 MiB peak direction. Uninstrumented GNU Time ABBA processes used 100
warmups and 1,000 samples per leg; maximum RSS was 30,848/30,976 KiB before
and 30,848/30,848 KiB after, flat at the measurement granularity.

Matched `perf stat` ABBA processes at the same sample count give:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 10,562.920 ms | 8,000.890 ms | -24.25% |
| cycles | 52,086,313,173 | 39,436,354,287 | -24.29% |
| instructions | 210,575,521,590 | 162,029,436,468 | -23.05% |
| branches | 42,549,015,960 | 32,727,398,511 | -23.08% |
| branch misses | 55,545,234 | 50,261,929 | -9.51% |
| cache references | 1,049,719,647 | 823,375,108 | -21.56% |
| cache misses | 84,894,963 | 74,728,818 | -11.97% |
| page faults | 18,800 | 18,786 | -0.07% |
| context switches | 149 | 98 | -34.23% |
| CPU migrations | 1 | 2 | +1 event |

Sampled profiles move the expected worksheet parser work: `NsReader`'s top
exclusive share falls from 13.31% to 9.36%, and the former public first-read
parse no longer executes for an eligible changed sheet. Process-wide profiles
still include source-sheet edit preparation and the one complete verifier, so
their relative shares are supporting attribution rather than removed-time
claims.

## Correctness verification

- focused tests prove that the changed sheet is preinitialized, an untouched
  sheet stays cold, oversized stores are not retained, and equal-but-distinct
  bytes or different style/shared-string lineages refuse adoption;
- all-feature/all-target `litchi-xlsx` tests pass, including 730 library tests,
  integration suites, security/encryption cases, examples, patch/inverse and
  real-producer fixtures;
- warning-denied library Clippy passes, and all-target Clippy passes with only
  the two named pre-existing test-lint allowances (`needless_question_mark`
  and `module_inception`);
- the complete harness test suite and warning-denied harness Clippy pass;
- formatting, workflow YAML, all retained JSON, evidence hashes, final binary
  identity, `git diff --check`, and staged-scope checks are commit gates.

Warning-denied rustdoc remains blocked by existing private/broken intra-doc
links in unchanged XLSX modules. Warning-denied harness dependencies likewise
surface the existing `litchi-odf-common` GenericArray deprecations; the harness
itself passes `-D warnings`. These are not reported as passing gates.

Raw primary, guard, rejected, Heaptrack, GNU Time, `perf stat`, and sampled
profile files are under `docs/performance/results/`; their digests are in
`xlsx-store-handoff-sha256.txt`.

## Next non-iWork audits

1. OLE2: benchmark PPT text-edit source-editor reuse; public text editing
   currently opens the same CFB editor for preflight and resolution.
2. RTF: add byte-1252, LZFu, LibreOffice watermark and relative-font-size
   coverage before another parser specialization.
3. ODF: do not fuse the ODP transition-style prepass into page parsing unless
   malformed-input error precedence can be preserved; a repeated-query cache
   needs a separate retained-memory policy and benchmark.
4. OOXML: attribute XLSX bulk action-plan flattening and source-backed editable
   publication independently of this bounded changed-sheet cache.

iWork remains deferred while the `iwa-*` crates are modified independently.
