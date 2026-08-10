# RTF ASCII transport batching

Date: 2026-08-11

Production base: `d71bede640954e8f202df459559cdf4dd7da4a04`

Scope: native RTF parser transport-byte accumulation only. OLE2, OOXML and
ODF production code are unchanged, and iWork/IWA crates were explicitly
excluded.

## Hypothesis

After parser-state specialization, matched large-corpus profiles attributed
15.37% of open samples and 14.46% of one-edit/save samples exclusively to
`SmallVec::extend`. `append_transport_bytes` invoked that generic extension
path once for every source character even though the generated corpus, and the
ordinary RTF source syntax it represents, is ASCII.

## Change

`append_transport_bytes` now recognizes an all-ASCII token and extends its
destination once from `str::bytes`. Non-ASCII input retains the existing
character-by-character checked conversion, including acceptance of byte-valued
characters, rejection of code points above `u8::MAX`, and preservation of the
valid prefix before an error.

Focused tests prove one extension call for a complete ASCII span, the existing
Latin-1-valued fallback, and the existing partial-prefix `InvalidUnicode`
behavior. This is private work elimination. It changes no public type,
dependency, snapshot, transaction, patch, output, limit, security policy,
runtime, lock, cache or unsafe-code boundary.

## Matched latency measurement

The harness and corpus are unchanged. The before release executable SHA-256 is
`178c8f728fb91f9ed3e43c2465c8dd710850393d16f4a945de6cebfda866c59a`;
the after SHA-256 is
`953a0a45a34e9b527de222418468bf5dac81695fd70aee72b8512619cc117c77`.

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator and CPU 2 pinned with `taskset`. The
deterministic large RTF contains 10,000 paragraphs and 540,051 source bytes.
Its source SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`.
Every timed open is semantically inspected, and every changed output is
reopened with its complete paragraph and text semantics verified outside
timing.

The primary run used 50 warmups and 500 samples per leg in before-A, after-A,
after-B, before-B order. Pooling the two legs gives 1,000 raw samples per
state; the statistics below are recomputed from those samples.

| Large RTF workload | Before p50 | After p50 | p50 delta | p95 delta | Mean delta |
|---|---:|---:|---:|---:|---:|
| Open | 3.159 ms | 2.316 ms | **-26.67%** | -24.79% | -26.56% |
| One-paragraph edit/save | 7.795 ms | 7.307 ms | **-6.26%** | -5.22% | -5.73% |

The approximate independent-sample 95% interval for the open mean difference
is `[-0.866, -0.833] ms`, or `[-27.08%, -26.05%]` of the before mean. The
one-edit/save interval is `[-0.478, -0.419] ms`, or
`[-6.11%, -5.36%]`. Both are wholly beneficial.

Raw primary reports and their SHA-256 digests:

- `abba-rtf-ascii-transport-primary-before-a.json`:
  `1ffaef5e60e81bfe5315bc911d6186353fbd5ffb9a9a07992712c22c054b8744`
- `abba-rtf-ascii-transport-primary-after-a.json`:
  `e205b53e07935241939bf278500dbf19ea1547a93363d2b92d6943dd10d37094`
- `abba-rtf-ascii-transport-primary-after-b.json`:
  `043586c67a5c3e428d2b3a8e0246308ac92ebc6e6732bf57c6bb5926993bd983`
- `abba-rtf-ascii-transport-primary-before-b.json`:
  `75fed8a022082e3872c7ad44c9f861513f3f9fe0f1619376e67c87aa053f07ee`

## Guardrails

An independent four-leg run used 30 warmups and 500 samples per leg for both
medium and large corpora. The table pools 1,000 samples per state.

| Guardrail | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|
| Open, medium | -30.09% | -29.54% | -26.27% |
| Open, large | -27.38% | -26.46% | -22.52% |
| One edit/save, medium | -10.07% | -8.71% | -4.10% |
| One edit/save, large | -8.49% | -8.54% | -9.08% |
| Full text, medium | 0.00% | +0.18% | -5.54% |
| Full text, large | -4.07% | -4.44% | -6.23% |
| Exact no-op edit/save, large | -0.21% | -0.29% | +4.38% |
| Exact stream save, large | -7.53% | -7.74% | -13.96% |

The sub-microsecond medium stream-save timer initially moved by more than 5%
in the broad matrix despite being outside the changed parser path. A dedicated
four-leg run with 20,000 samples per leg resolves its pooled before/after p50
to 150/150 ns, p95 to 230/210 ns, and mean to 153.64/155.20 ns (+1.01%). The
raw reports are the `abba-rtf-ascii-transport-guardrails-*.json` and
`abba-rtf-ascii-transport-stream-medium-*.json` files under `results/`.

## Profile, counters and memory

Matched `perf record` runs used ten warmups and 120 samples. On large open,
the exclusive `SmallVec::extend` frame falls from 15.37% to 2.56% and
`append_transport_bytes` from 2.46% to 0.60%. On large one-edit/save, those
frames fall from 14.46% to 1.60% and from 2.31% to 0.46%, respectively.

Matched `perf stat` ABBA processes used 20 warmups and 200 large one-edit/save
samples per leg:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 8,732 ms | 7,747 ms | -11.28% |
| cycles | 42,634,364,563 | 37,852,281,402 | -11.22% |
| instructions | 158,927,695,513 | 129,677,716,632 | -18.40% |
| branches | 40,789,809,543 | 35,064,282,934 | -14.04% |
| branch misses | 37,333,187 | 19,112,248 | -48.81% |
| cache references | 2,854,161,106 | 2,795,587,543 | -2.05% |
| cache misses | 133,133,160 | 130,345,247 | -2.09% |

Heaptrack over two warmups and 20 samples reports identical whole-process
allocation calls (1,416,591), temporary allocations (240,090), peak heap
(56.98 MiB), and leaked bytes (544 B). Instrumented peak RSS is 65.77/66.15
MiB (+0.58%). Uninstrumented GNU Time ABBA maximum RSS is 54,592/54,588 KiB
before and 54,588/54,720 KiB after; the maximum-to-maximum delta is +0.23%.
User time falls from 3.13/3.18 s to 2.59/2.62 s.

Raw evidence is in `rtf-ascii-transport-*-perf-report.txt`,
`rtf-ascii-transport-perf-stat-*.csv`,
`rtf-ascii-transport-*-heaptrack.txt`, and
`rtf-ascii-transport-time-*.txt`.

## Correctness verification

- the complete `litchi-rtf --all-features` suite passed, including 295 library
  unit tests, every integration suite and nine doctests;
- warning-denied all-target/all-feature Clippy and warning-denied crate rustdoc
  passed;
- the `parse_rtf` fuzz target and its production dependency graph compile
  offline;
- the unchanged benchmark harness's 23 tests and warning-denied Clippy passed;
- focused ASCII batching, byte-valued fallback and invalid-Unicode regressions
  passed; and
- formatting and `git diff --check` are final commit gates.

The parser's code-page decoding, limits, immutable source, exact no-op bytes,
checked edit, durable patch/inverse, stale-source, opaque-syntax, candidate
parse/readback, compressed-input and forward-only sink contracts remain
covered.

## Rejected ODF candidate and next audits

This tranche also measured, rejected and fully reverted an ODT changed-commit
candidate that adopted its already parsed final `Document` instead of copying
and reparsing the final archive. Large one-edit/save p50 improved 5.70%, p95
3.60%, and mean 5.20%; medium one-edit/save improved 6.39% p50 and 7.50% mean.
However, a dedicated medium one-paragraph read guard regressed 4.66% p50,
17.64% p95 and 6.33% mean. That exceeds the common-workload review threshold,
so none of the production or test code remains.

The next non-iWork candidates remain independently gated:

1. Remove the measured regenerated OPC logical-payload copy at the private ZIP
   publication boundary while retaining topology fallback and framing.
2. Measure exact validated CFB render reuse at the native object-editor finish
   boundary with retained-byte peak-memory evidence.
3. Attribute another ODF package-parse or unchanged-member publication path;
   do not revive the rejected ODT final-document handoff without explaining
   and eliminating its read regression.
4. Extend RTF evidence to formatted/media-heavy, compressed/code-page,
   malformed/security, real-producer and broad-edit corpora before another
   parser specialization.

iWork remains deferred while the `iwa-*` crates are modified independently.
