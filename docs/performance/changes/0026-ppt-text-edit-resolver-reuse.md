# PPT text editing reuses its validated persisted-record editor

Date: 2026-08-11

Production base: `f9cd26da793a8342393e6c6d2862c5c56bc3734d`

Scope: direct native PPT/OLE2 shape-text transactions. OOXML, RTF and ODF
production code are unchanged, and iWork/IWA crates were explicitly excluded.

## Hypothesis and change

`text_edit::Snapshot::edit_text` opened a complete persisted-record `Editor`
to reject signed, encrypted, protected or otherwise unsupported sources, then
discarded it. Shape resolution parsed the public presentation selector and
opened the same CFB editor again solely to read the selected slide's live
persisted record.

The transaction now resolves the semantic slide/shape selector into a private
persist-ID/shape-ID result, but does not surface that result yet. It then runs
the same complete editor preflight. Only after that gate succeeds does it read
the persisted record from that editor and construct the transaction. This
keeps the allocation-heavy presentation and editor owners sequential while
removing the second editor open, CFB validation, protection traversal, stream
capture, Current User parse and live persist-map construction.

Public error precedence is unchanged: an editor failure is still mapped to
`UnsupportedSource` before a stored selector error is returned. A focused
signed-storage test proves that a protected source plus an out-of-range target
still returns `UnsupportedSource`. No candidate is staged before the editor
gate. Standalone resolution, anchor editing, commit-time fresh-editor source
comparison, persisted-record replacement, final editor finish, complete
snapshot reopen and independent text readback are unchanged.

There is no public API, dependency edge, cache, retained snapshot state,
runtime, lock, unsafe code, durability rule, package limit or source-ownership
change.

## Dedicated public benchmark

The new opt-in `ppt_text_edit_one_edit_save` case constructs the public
`text_edit::Snapshot` outside timing, then times direct `edit_text`, one
same-length middle-shape replacement and `commit`. One final result undergoes
exact patch replay, inverse restoration, direct text-edit readback and a
complete generic public presentation reopen outside timing. Every iteration
also checks the exact published bytes against the deterministic changed
artifact produced before measurement.

The harness now has 109 selectable cases; the 36-case / 198-record default
matrix is unchanged. The native OLE2 smoke/release matrix has 20 cases and 40
tiny/large records.

The common-harness baseline executable SHA-256 is
`bcf596b6219ab726d5f9face8fb2a3d51af7078660cb0c7e83330f4e5d11301c`.
The final executable SHA-256 is
`582e9721df55b5a2e51135eba22ba0832796bc835174f97ad0d8493e502c4f5e`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic large corpus has 144 text boxes,
four CFB streams, 9,072 bytes of logical text, a 40,960-byte archive, archive
SHA-256 `229052cd918c0e5b7ef44070bafe20833531eee119b5943b18499503e225ff52`,
and PowerPoint Document SHA-256
`bef446ada643821b87531c06be7564b7ff8ca5539bb6a39766fbd28c11f65523`.

## Matched latency measurement

The primary ABBA run used 30 warmups and 1,000 samples per leg. Pooling 2,000
samples per state gives:

| Large direct PPT text edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 206.209 us | 177.089 us | **-14.12%** |
| p95 | 250.666 us | 208.066 us | **-16.99%** |
| p99 | 289.850 us | 236.485 us | **-18.41%** |
| mean | 213.799 us | 180.899 us | **-15.39%** |

The deterministic-bootstrap 95% interval for the mean delta is
`[-15.92%, -14.86%]` of the before mean. Both after legs have lower p50 and
mean than both before legs.

## End-to-end and unaffected guardrails

| Workload | p50 delta | p95 delta | p99 delta | Mean delta |
|---|---:|---:|---:|---:|
| Root slide-order one-shape edit/save | -3.59% | -4.95% | -0.75% | -3.23% |
| Ordinary PPT open | +0.48% | +4.61% | +3.93% | +1.05% |
| One selected shape | -1.78% | -4.06% | -9.92% | -2.13% |
| Exact no-op edit/save | -0.68% | +1.16% | +2.01% | -0.75% |
| Root slide-order snapshot open | -0.40% | -6.64% | -2.79% | -1.18% |

The ordinary root edit uses 50 warmups and 500 samples per leg. Open,
selected-shape and root-open guards use 1,000 samples per leg; exact no-op uses
2,000. All unaffected cells remain inside the 5% p50/mean review threshold.

## Allocations, memory and counters

Matched Heaptrack processes used 1,000 primary samples and one complete
post-timing verifier:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 3,433,239 | 3,311,999 | **-3.53%** |
| Temporary allocations | 380,942 | 357,896 | **-6.05%** |
| Peak heap | 859.30 KiB | 859.31 KiB | unchanged |
| Heaptrack RSS | 11.95 MiB | 12.03 MiB | +0.67% |
| Leaked bytes | 544 B | 544 B | unchanged |

Uninstrumented GNU Time ABBA processes used 100 warmups and 10,000 samples per
leg. Maximum RSS was 30,976/30,976 KiB before and 30,848/30,976 KiB after,
flat at the 128 KiB measurement granularity.

Matched `perf stat` ABBA processes at the same sample count give:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 4,513.440 ms | 4,350.920 ms | -3.60% |
| cycles | 21,995,522,784 | 21,414,861,864 | -2.64% |
| instructions | 68,829,211,564 | 67,209,937,365 | -2.35% |
| branches | 13,328,340,060 | 12,947,748,400 | -2.86% |
| branch misses | 95,730,095 | 87,949,020 | -8.13% |
| cache references | 3,101,023,949 | 2,883,354,129 | -7.02% |
| cache misses | 229,198,347 | 212,974,703 | -7.08% |
| page faults | 75,613 | 314,117 | **+315.43%** |
| context switches | 919 | 766 | -16.65% |
| CPU migrations | 2 | 2 | unchanged |

The page-fault trigger is retained rather than averaged away. A separate
minor/major split records 37,825 versus 157,057 minor faults and zero major
faults in both states. The system allocator faults more demand-zero pages over
10,000 repeated lifetimes after the allocation sequence changes, despite
fewer allocation calls. Direct latency, process task time, allocation count,
cache counters, peak heap and uninstrumented RSS all improve or remain within
the review threshold; no page-locality improvement is claimed.

## Correctness verification

- a focused differential test proves the split selector plus supplied editor
  resolves the same semantic target, persist ID, native shape ID, raw slide
  record, text encoding, payload, text and resize capability as standalone
  resolution;
- the protected-source/out-of-range test preserves `UnsupportedSource` error
  precedence;
- existing exact patch/inverse, durable patch, same-length, length-changing
  byte/UTF-16 and semantic-position tests pass;
- the public benchmark verifies deterministic changed bytes, exact replay,
  inverse restoration, direct readback and every slide/shape through the
  generic facade;
- complete all-feature/all-target PPT and harness suites, warning-denied
  Clippy, formatting, workflow counts, evidence hashes, `git diff --check` and
  staged-scope checks are commit gates.

Raw primary, guard, Heaptrack, GNU Time and `perf stat` evidence is under
`docs/performance/results/`; its digests are in
`ppt-text-edit-sha256.txt`.

## Next non-iWork audits

1. ODF: profile a bounded adaptive ODS facade cell-locator index on medium and
   large full-cell scans before accepting retained lookup state.
2. OOXML: attribute XLSX writer action regrouping on medium and dense-wide 1%
   commits before flattening any validated maps.
3. RTF: add byte-1252, LZFu, LibreOffice watermark and relative-font-size
   coverage before another parser specialization.
4. OLE2: attribute the remaining final owner/public-reader validation layers;
   keep PPT anchor and standalone resolver paths independent of this tranche.

iWork remains deferred while the `iwa-*` crates are modified independently.
