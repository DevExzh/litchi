# RTF parser-state specialization

Date: 2026-08-11

Production base: `dde814dcc656551b138c03d3e18389188d5b9171`

Scope: native RTF document-body text parsing only. OLE2, OOXML and ODF are
unchanged, and iWork/IWA crates were explicitly excluded.

## Hypothesis

`flush_text_buffer` cloned the complete parser `State` for every ordinary body
text run even though the common path needs only the effective encoding and the
copyable character and paragraph properties. A 100-sample large one-edit/save
profile attributed 8.53% of process samples exclusively to `State::clone`.
Revision runs need the complete state for author/date metadata, but ordinary
text in the deterministic corpus does not.

## Change

The parser now borrows the current state long enough to:

- reject non-body destinations;
- resolve the effective text encoding; and
- copy the `Formatting` and `Paragraph` values used by the emitted block.

It clones the complete state only when `revision_type` is present. Deletion
text still contributes revision metadata without entering the visible body;
inserted text still records its exact body range. A focused regression combines
an explicit Japanese font code page, bold centered ordinary text, insertion,
deletion and revision-author metadata so both the common and exceptional paths
are exercised together.

This is private work elimination. It does not change a public type, dependency
edge, snapshot identity, transaction, patch, output, resource limit, security
policy, runtime, lock, cache or unsafe-code boundary.

## Matched latency measurement

The harness and corpus are unchanged. The before release executable SHA-256 is
`3146676f551e652b436e2738a1a9832c5a2205eb58fe96ad53f956d31a0a7728`;
the after SHA-256 is
`178c8f728fb91f9ed3e43c2465c8dd710850393d16f4a945de6cebfda866c59a`.

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator and CPU 2 pinned with `taskset`. The
deterministic large RTF contains 10,000 paragraphs and 540,051 source bytes.
Its source SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`;
the changed output is 540,052 bytes. Every timed output is reopened and its
complete paragraph and text semantics are verified outside timing.

The primary run used 50 warmups and 500 samples per leg in before-A, after-A,
after-B, before-B order. Pooling the two legs gives 1,000 raw samples per
state; statistics below are recomputed from those samples.

| Large RTF one-paragraph edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 8.630 ms | 7.634 ms | **-11.54%** |
| p95 | 9.201 ms | 8.091 ms | **-12.06%** |
| mean | 8.701 ms | 7.682 ms | **-11.71%** |

The approximate independent-sample 95% interval for the mean difference is
`[-1.050, -0.988] ms`, or `[-12.07%, -11.35%]` of the before mean. Matched A
and B p50 comparisons improve 12.19% and 11.17%. Within-state p50 drift is
1.47% before and 0.33% after.

Raw primary reports and their SHA-256 digests:

- `abba-rtf-state-clone-one-edit-before-a.json`:
  `6a54ae4f0b5c05c234068eb0d8dedfd4e800d01439042f20a58f85241f6904db`
- `abba-rtf-state-clone-one-edit-after-a.json`:
  `e8890e2ea4bc02cf0376a634e729590e0d9c6353f2f35ccb98f13bac984ab6f6`
- `abba-rtf-state-clone-one-edit-after-b.json`:
  `d512490a5e1586b22a69b13c092c39fc66b3566dec14dc0fa6bcf0fc3a10e403`
- `abba-rtf-state-clone-one-edit-before-b.json`:
  `0404dbec3839a578e707f485318a2d56622f6230a6c42869e16240f890a5d9eb`

## Guardrails

An independent four-leg run used 50 warmups and 500 samples per leg for both
medium and large corpora. The table pools 1,000 samples per state.

| Guardrail | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| Open, medium | 69.807 us | 55.256 us | -20.84% | -21.64% | -22.40% |
| Open, large | 4.010 ms | 3.205 ms | -20.09% | -19.94% | -20.43% |
| One edit/save, medium | 116.010 us | 99.578 us | -14.16% | -14.34% | -14.26% |
| Full text, large | 24.584 us | 24.254 us | -1.34% | -2.26% | -5.80% |
| Exact no-op edit/save, large | 86.159 us | 85.337 us | -0.95% | -0.66% | -2.99% |
| Exact stream save, large | 89.123 us | 90.675 us | +1.74% | +2.89% | +8.32% |

The save-only path cannot execute the changed parser branch. Because its
sub-100-us p95 moved more than 5% in the broad matrix, a dedicated follow-up
used 5,000 samples per completed leg. Its before-A p50/p95 was 88.392/110.782
us; the pooled after-A/after-B p50/p95 was 85.809/103.273 us (-2.92%/-6.78%),
with mean -4.12%. The intended fourth leg was terminated when unrelated host
CPU contention starved the pinned process; it is not represented as a passing
ABBA result. The complete four-leg 500-sample guard remains the decision
record, and the higher-sample completed legs reject a save-only regression.

The raw reports are the `abba-rtf-state-clone-guardrails-*.json` and completed
`abba-rtf-state-clone-stream-save-*.json` files under `results/`.

## Profile, counters and memory

Matched `perf record` runs used ten warmups and 100 large one-edit/save samples.
The exclusive `State::clone` frame accounts for 8.53% of baseline samples and
is absent from the after report; `flush_text_buffer` itself accounts for 0.80%
after specialization.

Matched `perf stat` ABBA processes used 20 warmups and 200 samples per leg:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 9,871 ms | 8,785 ms | -11.00% |
| cycles | 48,373,268,975 | 43,292,479,928 | -10.50% |
| instructions | 175,176,116,163 | 158,919,974,150 | -9.28% |
| branches | 42,456,389,647 | 40,788,356,916 | -3.93% |
| branch misses | 34,901,996 | 34,100,599 | -2.30% |
| cache references | 3,131,296,423 | 2,861,702,330 | -8.61% |
| cache misses | 138,367,793 | 133,592,965 | -3.45% |

Heaptrack over two warmups and 20 samples reports identical whole-process
allocation calls (1,416,592), temporary allocations (240,091), peak heap
(56.98 MiB), and leaked bytes (544 B). This matches the mechanism: cloning the
large state copied inline values but did not allocate. Instrumented peak RSS is
66.05/66.07 MiB, also flat.

Uninstrumented GNU Time ABBA runs used 20 warmups and 200 samples. Maximum RSS
was 54,592/54,720 KiB before and 54,696/54,592 KiB after; the maximum-to-maximum
delta is -0.04%. User time fell from 3.71/3.69 s to 3.12/3.19 s.

Raw evidence is in `rtf-state-clone-*-perf-report.txt`,
`rtf-state-clone-perf-stat-*.csv`, `rtf-state-clone-*-heaptrack.txt`, and
`rtf-state-clone-time-*.txt`.

## Correctness verification

- the complete `litchi-rtf --all-features` suite passed, including 292 library
  unit tests, every integration suite and nine doctests;
- warning-denied all-target/all-feature Clippy and warning-denied crate rustdoc
  passed;
- the `parse_rtf` fuzz target and its production dependency graph compile
  offline;
- the unchanged benchmark harness's 23 tests and warning-denied Clippy passed;
- the focused formatting/code-page/revision regression passed; and
- formatting and `git diff --check` are final commit gates.

The final text-block construction, revision metadata, limits, immutable source,
exact no-op bytes, candidate parse, complete semantic readback, durable patch,
inverse, stale-source, opaque-syntax, compressed-input and forward-only sink
contracts remain covered.

## Rejected ODF candidate and next audits

This tranche also measured, rejected and fully reverted an ODS worksheet
snapshot change that adopted the already parsed target package instead of
copying the source archive and parsing the target package again. On the large
one-cell edit/save workload, pooled p50 moved from 327.900 to 326.458 ms
(-0.44%), p95 regressed from 339.622 to 340.655 ms (+0.30%), and mean improved
0.51%. That is below the materiality threshold, so none of its production or
test code remains.

The next non-iWork candidates remain independently gated:

1. Remove regenerated OPC payload copies at the private publication boundary,
   retaining topology fallback, ZIP framing and complete OOXML readback.
2. Measure reuse of an exact validated CFB render across the shared DOC/XLS
   object-editor finish boundary, with retained-byte peak-memory evidence.
3. Extend RTF coverage to formatted/media-heavy, compressed/code-page,
   malformed/security, real-producer and broad-edit corpora before specializing
   another parser branch.
4. Continue ODF package-parse and unchanged-member publication attribution;
   the rejected ODS snapshot adoption is not a reusable speedup claim.

iWork remains deferred while the `iwa-*` crates are modified independently.
