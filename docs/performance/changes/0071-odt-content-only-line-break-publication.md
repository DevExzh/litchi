# Change 0071: ODT content-only line-break publication

Date: 2026-08-12

Production base: `6f5774d99ae32ae816aa0b6f1335cd23f5f84af7`

Status: accepted

## Hypothesis and implementation

`AppendLineBreak` changes only the modeled paragraph content in `content.xml`,
but the scalar transaction dispatch still called the complete ODF package
rebuild. On a media-rich document that needlessly regenerated and recompressed
every unchanged resource. The existing ODT content-only publisher already
proves exact source lineage, compact source/target XML, same package topology,
eligible ZIP framing, unchanged manifest/security state, and raw identity of
all members other than `content.xml`.

The operation now calls `MutableDocument::to_bytes_content_only`, the same
accepted publication boundary used by paragraph replacement. No public API,
patch vocabulary, dependency, cache, runtime, lock, unsafe code, archive type,
or global state changes. Structural and mixed operations, oversized XML,
resource edits, signatures, encryption, manifest size metadata, and unsupported
ZIP layouts retain the established full rebuild or refusal policy.

The packaged-transaction regression independently exercises paragraph
replacement and line-break commits over an ODT with an opaque resource. For
the line-break result it proves that only `content.xml` changes at the raw ZIP
local/central-record level, reopens the exact `Before\n` text, replays the
patch, and applies the inverse back to the exact source archive.

## Matched corpus and protocol

The opt-in `odt_media_line_break_edit_save` case reuses the deterministic
`litchi-odt-media-paragraph-publication-v1` corpus: 200 paragraphs, eight
incompressible 2 MiB resources, 13 ZIP members, and 16,787,016 logical bytes.
The archive is 16,786,287 bytes with SHA-256
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
The target is paragraph 100; both paths append exactly one line break.

The control binary, SHA-256
`b8fb6ddaab2f6928461b225b5a910d7f502df004e54d1e66fbafc73f8b294827`,
contains the new harness case with the old full-rebuild dispatch. The candidate
binary, SHA-256
`dc414be1a3555ec5b69f2e158b64970b9ed687f66799b9c013937e9fd6dd9cdc`,
changes only that dispatch. Both publish the same deterministic artifact,
SHA-256
`c38d45741cdad624da8ac0d22e268f4074e8073c16a7586af13620a1e20e2c20`.

The timed interval contains snapshot open, public line-break staging, commit,
and final byte observation. Complete paragraph/media/manifest reopen, patch
replay, exact inverse restoration, stale-source refusal, raw-member identity,
and output hashing stay outside timing. On CPU 2, the retained order was
control A, candidate A, candidate B, control B. Each leg used 50 warmups and
500 samples, yielding 1,000 samples per state. Exact inputs and results are in
the [`measurement summary`](../results/odt-line-break-publication-summary.json).

## Results

| Metric | Full rebuild | Content-only | Delta |
|---|---:|---:|---:|
| pooled samples | 1,000 | 1,000 | — |
| p50 | 217.532 ms | 3.985 ms | **-98.17% (54.59x)** |
| mean | 218.050 ms | 4.005 ms | **-98.16% (54.44x)** |
| mean 95% interval | 217.819-218.281 ms | 3.992-4.018 ms | disjoint |
| p95 | 224.790 ms | 4.316 ms | **-98.08% (52.09x)** |
| p99 | 229.972 ms | 4.654 ms | **-97.98% (49.41x)** |

Every ordered leg improves independently. Control p50/mean between-leg drift
is 0.22%/0.19%; candidate drift is 2.28%/2.54%, within the 5% policy.

## Regression, allocation, and counter evidence

The three threshold-sensitive controls were rerun in isolated CPU-pinned
A/B/B/A sequences. Large one-paragraph edit/save changes +0.39% p50, +0.50%
mean, and +0.50% p95. Large exact no-op changes +3.68% p50, +2.61% mean, and
+9.96% p95, within the 5% central / 10% tail gates. Medium open changes -0.99%
p50, -0.13% mean, and -2.28% p95. All other mixed-matrix central guards remain
within the same policy; the largest other tail is list-medium at +9.58% p95.
The existing media-rich paragraph transaction remains neutral-to-improved at
-1.43% p50, -0.39% mean, and -2.10% p95.

One-sample Heaptrack attribution reports allocation calls 20,034 -> 18,652
(-6.90%) and temporary allocations 3,479 -> 2,963 (-14.83%). Peak heap is flat
at 107.17 -> 107.13 MiB. Heaptrack-inclusive RSS is flat at 119.19 -> 119.38
MiB, and two uninstrumented GNU Time processes per state report mean maximum
RSS 110,964 -> 111,032 KiB (+0.06%). Both states retain the same 1.78 KiB of
profiler/runtime leakage.

Two matched process-wide `perf stat` repeats per state used two warmups and ten
samples:

| Counter | Full rebuild | Content-only | Delta |
|---|---:|---:|---:|
| task clock | 3,514.41 ms | 917.94 ms | -73.88% |
| cycles | 17.228 billion | 4.546 billion | -73.61% |
| instructions | 44.609 billion | 9.662 billion | -78.34% |
| branches | 7.600 billion | 1.283 billion | -83.12% |
| branch misses | 163.298 million | 13.995 million | -91.43% |
| cache references | 1.178 billion | 306.351 million | -74.00% |
| cache misses | 21.016 million | 14.676 million | -30.17% |
| page faults | 222,547 | 156,047 | -29.88% |

CPU migrations were zero. Ten-sample `perf record` profiles lost no samples.
The control spends 51.22% of sampled cycles in `deflate_medium` and 17.88% in
`longest_match`; the candidate spends 18.00% and 6.42%. Multiplying those
shares by sampled cycles attributes approximate absolute reductions of 90.68%
and 90.48%, respectively. Corpus construction, the one changed `content.xml`
compression, and complete untimed verification remain in the process profile.

## Validation and limitations

Passed on the final source:

- complete all-feature ODT tests, including raw-member identity, text,
  patch/inverse and stale-source coverage;
- complete performance-harness tests plus deterministic debug/release cases;
- warning-denied ODT and performance-harness all-target Clippy and ODT rustdoc;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the GenericArray deprecation cleanup from `1194fbc7f`; and
- formatter, JSON/YAML parsing, whitespace, and final-diff checks.

This change authorizes only the existing `AppendLineBreak` transaction over the
already accepted content-only preservation boundary. It does not add arbitrary
inline formatting, structural edits, resource mutation, positional ZIP reads,
real-producer coverage, or a new security policy. OLE2, OOXML, RTF, other ODF
family production crates, and every iWork/IWA crate are unchanged by this
batch.
