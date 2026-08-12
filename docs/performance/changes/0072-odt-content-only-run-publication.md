# Change 0072: ODT content-only inline-run publication

Date: 2026-08-12

Production base: `900f222347d6b8fe3134a1f924e1469f327c6609`

Status: accepted

## Hypothesis and implementation

`AppendRun` changes only modeled paragraph content in `content.xml`, but its
scalar transaction dispatch still called the complete ODF package rebuild. On
a media-rich document that regenerated and recompressed every unchanged
resource. The existing ODT content-only publisher already proves exact source
lineage, compact source/target XML, same package topology, eligible ZIP
framing, unchanged manifest/security state, and raw identity of all members
other than `content.xml`.

The operation now calls `MutableDocument::to_bytes_content_only`, the same
accepted publication boundary used by paragraph replacement and line-break
insertion. Both styled and unstyled appended runs use the path. No public API,
patch vocabulary, dependency, cache, runtime, lock, unsafe code, archive type,
or global state changes. Structural and mixed operations, oversized XML,
resource edits, signatures, encryption, manifest size metadata, and
unsupported ZIP layouts retain the established full rebuild or refusal policy.

Initial guard measurements also exposed that exact no-op `Edit::commit` paid
the changed path's 3,432-byte stack-frame reservation before its early return.
The public method now returns the same source-sharing no-op `Commit` before
entering a private non-inlined changed-operation helper. Every changed commit
still enters the exact previous validation/dispatch body. This isolates the
frame without weakening source, envelope, operation, candidate, or final
validation.

The packaged-transaction regression exercises both unstyled and styled
append-run commits over an ODT with an opaque resource. It proves that only
`content.xml` changes at the raw ZIP local/central-record level, reopens the
exact text and style, replays the patch, applies the inverse back to the exact
source archive, and refuses stale replay.

## Matched corpus and protocol

The opt-in `odt_media_append_run_edit_save` case reuses the deterministic
`litchi-odt-media-paragraph-publication-v1` corpus: 200 paragraphs, eight
incompressible 2 MiB resources, 13 ZIP members, and 16,787,016 logical bytes.
The archive is 16,786,287 bytes with SHA-256
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
The target is paragraph 100; both paths append the exact unstyled text
` appended run`.

The control binary, SHA-256
`7eda79d72dfb9ea00ad1572f965d8f0c0ef3df579c73dbf6be2b9ab928adc537`,
contains the new harness case with the old full-rebuild dispatch. The candidate
binary, SHA-256
`7db856fff49eea43e7d08e0d889e536b21648bc723360e50426dcb59cdda8856`,
adds the content-only dispatch and no-op helper boundary. Both publish the same
deterministic artifact, SHA-256
`8b63becd2100f11255baff345e71cf66ae85aa7812313deaf54b782ce3d8f5db`.

The timed interval contains snapshot open, public inline-run staging, commit,
and final byte observation. Complete paragraph/media/manifest reopen, patch
replay, exact inverse restoration, stale-source refusal, raw-member identity,
and output hashing stay outside timing. On CPU 2, the retained order was
control A, candidate A, candidate B, control B. Each leg used 50 warmups and
500 samples, yielding 1,000 samples per state. Exact inputs and results are in
the [`measurement summary`](../results/odt-append-run-publication-summary.json).

## Results

| Metric | Full rebuild | Content-only | Delta |
|---|---:|---:|---:|
| pooled samples | 1,000 | 1,000 | — |
| p50 | 225.431 ms | 3.635 ms | **-98.39% (62.01x)** |
| mean | 225.944 ms | 3.670 ms | **-98.38% (61.56x)** |
| mean 95% interval | 225.709-226.179 ms | 3.656-3.684 ms | disjoint |
| p95 | 232.932 ms | 4.030 ms | **-98.27% (57.80x)** |
| p99 | 238.738 ms | 4.310 ms | **-98.19% (55.40x)** |

Every ordered leg improves independently. Control p50/mean between-leg drift
is -1.50%/-1.35%; candidate drift is -4.96%/-4.29%, within the 5% policy.

## Regression, allocation, and counter evidence

The final frozen binaries were rerun in isolated CPU-pinned A/B/B/A sequences.
Large exact no-op edit/save changes -1.10% p50, -3.73% mean, and -4.17% p95,
confirming that the helper boundary removes the frame without regression.
Large one-edit/save changes -0.73% p50, +0.06% mean, and +3.33% p95; large
one-percent edit/save changes -2.02%, -2.55%, and -2.51%. The existing
media-rich paragraph and line-break transactions improve by 2.24-2.71% and
0.32-1.16% across p50/mean/p95. All central guards remain within the 5%
policy; isolated p99 noise is disclosed in the summary but is not a central
acceptance metric.

One-sample Heaptrack attribution reports allocation calls 19,767 -> 18,383
(-7.00%) and temporary allocations 3,644 -> 3,127 (-14.19%). Peak heap is flat
at 107.17 -> 107.14 MiB. Heaptrack-inclusive RSS falls 119.26 -> 117.49 MiB,
and two uninstrumented GNU Time processes per state report mean maximum RSS
111,032 -> 111,096 KiB (+0.06%, flat). Both states retain the same 1.78 KiB of
profiler/runtime leakage.

One matched process-wide `perf stat` run per state used two warmups and ten
samples:

| Counter | Full rebuild | Content-only | Delta |
|---|---:|---:|---:|
| task clock | 3,666.12 ms | 932.59 ms | -74.56% |
| cycles | 17.934 billion | 4.527 billion | -74.76% |
| instructions | 44.794 billion | 9.639 billion | -78.48% |
| branches | 7.632 billion | 1.279 billion | -83.24% |
| branch misses | 164.007 million | 13.952 million | -91.49% |
| cache references | 1.196 billion | 308.112 million | -74.23% |
| cache misses | 22.263 million | 15.302 million | -31.27% |
| page faults | 256,034 | 156,046 | -39.05% |

CPU migrations were zero. Ten-sample `perf record` profiles lost no samples.
The control spends 57.79% of sampled cycles in `deflate_medium` and 17.42% in
`longest_match`; the candidate spends 19.36% and 6.21%. Multiplying those
shares by sampled cycles attributes approximate absolute reductions of 92.45%
and 91.97%, respectively. Corpus construction, the one changed `content.xml`
compression, and complete untimed verification remain in the process profile.

## Validation and limitations

Passed on the final source:

- complete all-feature ODT tests, including styled/unstyled raw-member
  identity, text/style, patch/inverse and stale-source coverage;
- complete performance-harness tests plus deterministic debug/release cases;
- warning-denied ODT and performance-harness all-target Clippy and ODT rustdoc;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the GenericArray deprecation cleanup from `1194fbc7f`; and
- formatter, workflow/JSON/YAML parsing, whitespace, and final-diff checks.

This change authorizes only the existing `AppendRun` transaction over the
already accepted content-only preservation boundary and isolates exact no-op
commit dispatch. It does not add new formatting semantics, structural edits,
resource mutation, positional ZIP reads, real-producer coverage, or a new
security policy. OLE2, OOXML, RTF, other ODF family production crates, and
every iWork/IWA crate are unchanged by this batch.
