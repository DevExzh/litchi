# Change 0074: ODT content-only hyperlink publication

Date: 2026-08-12

Production base: `eba9ff256a640181d09b53e1ced4993b299563af`

Status: accepted

## Hypothesis and implementation

`AppendHyperlink` changes only modeled paragraph content in `content.xml`, but
its scalar transaction dispatch still called the complete ODF package rebuild.
On a media-rich document that regenerated and recompressed every unchanged
resource. The existing ODT content-only publisher already proves exact source
lineage, compact source/target XML, unchanged package topology, eligible ZIP
framing, unchanged manifest/security state, and raw identity of all members
other than `content.xml`.

The operation now calls `MutableDocument::to_bytes_content_only`, the same
accepted publication boundary used by paragraph replacement, line-break
insertion, and inline-run append. No public API, patch vocabulary, dependency,
cache, runtime, lock, unsafe code, archive type, or global state changes.
Structural and mixed operations, oversized XML, resource edits, signatures,
encryption, manifest size metadata, and unsupported ZIP layouts retain the
established full-rebuild or refusal policy.

The packaged-transaction regression appends an inert hyperlink over an ODT
with an opaque resource. It proves that only `content.xml` changes at the raw
ZIP local/central-record level, reopens the exact text and URL, replays the
patch, applies the inverse back to the exact source archive, and refuses stale
replay.

## Matched corpus and protocol

The opt-in `odt_media_append_hyperlink_edit_save` case reuses the deterministic
`litchi-odt-media-paragraph-publication-v1` corpus: 200 paragraphs, eight
incompressible 2 MiB resources, 13 ZIP members, and 16,787,016 logical bytes.
The archive is 16,786,287 bytes with SHA-256
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
The target is paragraph 100; both paths append the exact text
` performance link` with inert URL
`https://example.invalid/performance`.

The control binary, SHA-256
`4ac963b520f2e8e5cf025121e28344ccbb11d42803d3e6c6a2a124d7d592079a`,
contains the new harness case with the old full-rebuild dispatch. The candidate
binary, SHA-256
`b87ec233282fb0cddd49602039fd1c23014d941e1c4a9005472bf28c4a47b43a`,
changes only the dispatch and focused proof. Both publish the same deterministic
artifact, SHA-256
`5c55c1e220a434868fd7d57d967b9bc0b84ad31711aebb16da05074bd3608f8d`.

The timed interval contains snapshot open, public hyperlink staging, commit,
and final byte observation. Complete paragraph/hyperlink/media/manifest reopen,
patch replay, exact inverse restoration, stale-source refusal, raw-member
identity, and output hashing stay outside timing. On CPU 2, the retained order
was control A, candidate A, candidate B, control B. Each leg used 50 warmups
and 500 samples, yielding 1,000 samples per state. A preliminary sequence was
discarded before analysis after overlapping yielded processes were detected;
the retained sequence was rerun serially without overlap. Exact retained inputs
and results are in the
[`measurement summary`](../results/odt-append-hyperlink-publication-summary.json).

## Results

| Metric | Full rebuild | Content-only | Delta |
|---|---:|---:|---:|
| pooled samples | 1,000 | 1,000 | — |
| p50 | 221.443 ms | 3.988 ms | **-98.20% (55.52x)** |
| mean | 221.714 ms | 4.027 ms | **-98.18% (55.06x)** |
| mean 95% interval | 221.493-221.936 ms | 4.013-4.040 ms | disjoint |
| p95 | 227.046 ms | 4.376 ms | **-98.07% (51.88x)** |
| p99 | 234.302 ms | 4.833 ms | **-97.94% (48.48x)** |

Every ordered leg improves independently. Control p50/mean between-leg drift
is -0.25%/-0.10%; candidate drift is +2.29%/+2.87%, within the 5% policy.

## Regression, allocation, and counter evidence

The ordinary large ODT open, list-paragraphs, one-paragraph, full-text, no-op,
one-edit, and one-percent-edit guards were measured in a serial A/B/B/A
sequence with 400 samples per state. Every pooled p50 and mean regression is
below 5%; every p95 regression is below 10%. The largest mean change is +4.12%
for full text, and the central edit/save guards change -2.22% to +1.15% mean.
The existing media-rich paragraph, line-break, and append-run publication cases
improve across p50, mean, and p95.

One-sample Heaptrack attribution reports allocation calls 19,805 -> 18,421
(-6.99%) and temporary allocations 3,651 -> 3,134 (-14.16%). Peak heap is flat
at 107.19 -> 107.16 MiB. Heaptrack-inclusive RSS falls 119.57 -> 117.68 MiB,
and two uninstrumented GNU Time processes per state report mean maximum RSS
112,634 -> 111,536 KiB (-0.98%). Both states retain the same 1.78 KiB of
profiler/runtime leakage.

Three matched `perf stat` repeats per state used two warmups and ten samples.
Mean task clock falls 3,600.22 -> 925.17 ms (-74.30%), cycles fall
17.656 -> 4.537 billion (-74.31%), instructions fall 44.640 -> 9.671 billion
(-78.34%), branches fall 7.606 -> 1.285 billion (-83.11%), and branch misses
fall 163.979 -> 13.992 million (-91.47%). CPU migrations were zero.

Ten-sample `perf record` profiles lost no samples. The control spends 52.55%
of sampled cycles in `deflate_medium` and 16.97% in `longest_match`; the
candidate spends 15.89% and 5.02%. Multiplying those shares by sampled cycles
attributes approximate absolute reductions of 92.27% and 92.44%, respectively.
Corpus construction, compression of the one changed XML member, and complete
untimed verification remain in the process profile.

## Validation and limitations

Passed on the final source:

- complete all-feature ODT tests, including raw-member identity, exact
  hyperlink text/URL, patch/inverse, stale-source, security-envelope, and
  oversized-content fallback coverage;
- complete performance-harness tests plus deterministic debug/release cases;
- warning-denied ODT and performance-harness all-target Clippy and ODT rustdoc;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the GenericArray deprecation cleanup from `1194fbc7f`; and
- formatter, workflow/JSON/YAML parsing, whitespace, ADR-hash, and final-diff
  checks.

This change authorizes only the existing `AppendHyperlink` transaction over the
already accepted content-only preservation boundary. It does not activate the
URL, fetch external content, add a relationship, change formatting semantics,
perform a structural or resource edit, add positional ZIP reads, establish
real-producer/native-Office coverage, or change security policy. OLE2, OOXML,
RTF, other ODF family production crates, and every iWork/IWA crate are unchanged
by this batch.
