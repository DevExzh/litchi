# Change 0048: retained RTF body source span

Date: 2026-08-11
Production base: `f0469cffe9a8dcb27270bc88265e25dfc3aadbd8`
Status: accepted

## Hypothesis and measured owner

Every changed ordinary RTF body-text commit parsed the source twice. The
initial `Document::from_bytes` lexer already produced exact token ranges and
the parser structurally validated the complete document. Later,
`Edit::commit` cloned the 540,051-byte ASCII source into a `String`, allocated
a second token and span vector, and scanned root depth only to rediscover the
body range. Candidate publication still performed its required complete parse
and semantic readback after that locator.

Heaptrack attributed 588 allocation calls across 20 measured edits to this
second `ordinary_body_source_span` lexer subtree. The hypothesis was that
retaining a conservatively proven range from the first parse would remove that
work without weakening candidate validation or expanding the set of editable
documents.

## Implementation and fallback boundary

The parser's existing complete structural-preflight loop now also observes the
first literal root-level text offset, the root closing offset, later nested
groups, and root-level binary tokens. This adds no second token-vector pass.
For direct, uncompressed ASCII input only, a non-empty, contiguous and
binary-free range is packed into the private immutable document model. A
changed commit checks ASCII transport and range bounds before using it.

Empty bodies, non-ASCII/code-page transport, LZFu transport, root-level binary
data, a nested group after body text, a range larger than 32-bit offsets, and
any absent or invalid private range take the established full locator path.
That path and all of its typed refusals remain present. The optimization does
not make CP-1252 or compressed bodies editable.

The first prototype calculated the range with a second token-vector scan at
model construction and regressed the large open guard. It was removed. The
accepted version derives the range inside work the parser already performs and
uses one packed `u64` in the immutable model. The changed commit still runs:

1. the existing plain-body editability and opaque-preservation checks;
2. bounded replacement encoding and source splicing;
3. complete `Snapshot::from_bytes_with_limits` candidate parse; and
4. semantic text/property readback before publication.

No public API, dependency, feature, cache, lock, executor, runtime, unsafe-code,
patch, signature, encryption, source-version or sink contract changes.

## Matched latency evidence

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86-64 AMD EPYC
9575F VM, Rust system allocator and CPU 2 pinned with `taskset`. The fixed large
plain corpus contains 10,000 paragraphs, 499,999 logical text bytes and 540,051
source bytes. Its SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`.
Every iteration performs the public paragraph replacement, complete changed
commit, forward-only save, exact output comparison, ordinary reopen and full
semantic verification.

The control executable SHA-256 is
`3cf6c077ff116355558724e51038382b01894876ece10ab8c62c1874e71271b8`;
the measured accepted candidate is
`4008ca067234a83e08ed5e7f3a3069da13d9425db769b90f174217cb6d0f2b69`.
Their `.text` hashes are respectively
`d052acef3b92331c64a4b89fb4304fb03d7ab53e845f52335a69bfa10f68994a`
and
`a035e0b6a38a036e03f51495ea4512511f763f40fa477081d9b29e2569f517e1`.

The headline ABBA run used 30 warmups and 250 measured samples per leg in
control-A/candidate-A/candidate-B/control-B order. Pooling 500 samples per state
gives:

| Large plain RTF one-edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 6.052671 ms | 5.403987 ms | **-10.72%** |
| mean | 6.090771 ms | 5.474709 ms | **-10.11%** |
| p95 | 6.472859 ms | 5.905858 ms | **-8.76%** |
| p99 | 7.179094 ms | 6.742781 ms | **-6.08%** |

Both candidate legs improve against their adjacent control: means improve
9.31% and 10.93%, while p50 improves 10.32% and 10.83%. Candidate-leg p50s
are 5.459 and 5.363 ms; control-leg p50s are 6.087 and 6.015 ms.

A separate 50-warmup/500-sample-per-leg ABBA scaling run shows the same result
on smaller inputs:

| Shape | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| tiny | 11.546 us | 9.863 us | **-14.58%** | -13.59% | +4.15% |
| medium | 81.693 us | 70.847 us | **-13.28%** | -14.08% | -17.80% |

Tiny p95 remains within the 5% guardrail and its p99 improves 10.23%.

## Allocation, memory and hardware counters

Matched Heaptrack processes used 20 large samples. The complete process
contains corpus creation, expected-output construction, all ordinary source
and candidate parses, saves and verification, so peak memory remains dominated
by retained snapshots rather than the short-lived locator.

| Heaptrack process metric | Before | After | Delta |
|---|---:|---:|---:|
| allocation calls | 1,306,120 | 1,305,511 | -609 / -0.05% |
| temporary allocations | 220,087 | 220,066 | -21 / -0.01% |
| peak heap | 56.98 MiB | 56.98 MiB | flat |
| profiler peak RSS | 57.45 MiB | 58.05 MiB | +1.04% |

The before-only `ordinary_body_source_span` subtree accounts for 588 calls:
260 token-vector growth calls for each of the token and span vectors, 20 final
allocations in each vector, 13 other growth calls, and 21 source-string calls.
The candidate has no allocation stack under that function. This directly
confirms removal of the second source copy and lexer vectors. The unchanged
544-byte profiler/runtime leak remains.

Uninstrumented GNU Time ABBA processes with ten warmups and 100 samples report
55,016/55,256 KiB before and 55,380/55,380 KiB after. The 0.44-0.66% movement
is flat within the guardrail.

Matched process-wide `perf stat` ABBA runs at the same 100-sample setting give:

| Counter, A+B | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 4,128.82 ms | 3,578.37 ms | -13.33% |
| cycles | 19,928,965,039 | 17,397,503,496 | -12.70% |
| instructions | 53,750,403,813 | 48,032,051,139 | **-10.64%** |
| branches | 12,911,246,010 | 11,549,678,831 | -10.55% |
| branch misses | 11,497,341 | 9,789,311 | -14.86% |
| cache references | 1,651,263,411 | 1,611,179,410 | -2.43% |
| cache misses | 91,714,809 | 93,504,515 | +1.95% |
| page faults | 1,518,848 | 1,094,309 | -27.95% |

CPU migrations are zero in every leg. The small cache-miss increase is
disclosed; it does not produce a latency, heap or RSS regression.

## Guardrails and verification

Separate final-candidate, per-case large guards retain open and list timing
outside the changed commit:

| Unchanged case | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|
| open | -6.00% | -6.42% | -8.07% | -8.10% |
| list paragraphs | -9.15% | -8.29% | -5.88% | +5.12% |

The list p99 movement is 3.296 us and its p50, mean and p95 all improve. A
2,000-sample-per-state exact no-op guard has substantial control-leg frequency
drift (19.817 versus 14.720 us p50), so it is not used as performance evidence;
its pooled p50/mean/p95 move -4.24%/-7.12%/-2.68%, and the implementation and
exact-source write path are unchanged.

Focused tests prove the cached range for ordinary ASCII, and absence for empty,
post-body nested, binary, non-ASCII and compressed sources. Existing transaction
tests retain exact output, patch/inverse, stale-source, opaque preservation,
modeled-header byte identity, structural/property refusal and complete
candidate readback. Transport contracts retain CP-1252 exact no-op and changed
refusal, LZFu exact no-op and changed refusal, and producer-watermark behavior.

The complete all-feature RTF suite passes: 297 library unit tests, every
integration suite and nine doctests. Warning-denied all-target/all-feature
Clippy and warning-denied rustdoc pass. The `parse_rtf` fuzz target and its
production graph compile offline; `cargo-fuzz` is not installed on this host.
The unchanged benchmark harness passes 32 tests and warning-denied Clippy. The
final 63-record capability-bounded RTF smoke spans tiny/medium/large plain,
CP-1252, LZFu and watermark cases. Formatting and `git diff --check` pass.

## Evidence and remaining work

Committed evidence includes:

- `abba-rtf-body-span-{before,after}-*.json` for the headline result;
- `abba-rtf-body-span-shapes-*.json` for tiny/medium scaling;
- `abba-rtf-body-span-{open,list,noop}-*.json` for guards;
- `rtf-body-span-{before,after}-heaptrack.txt` for allocation attribution;
- `rtf-body-span-perf-*.csv` and `rtf-body-span-time-*.txt` for counters/RSS;
- `rtf-body-span-variant-smoke.json` for capability coverage; and
- `rtf-body-span-sha256.txt` for executable, `.text` and artifact digests.

The next bounded RTF work needs a newly attributed owner frame or broader
formatted/media/security corpus; this change makes no broad RTF editing claim.
The strongest separate ODF candidate is selective ODP slide parsing with full
EOF validation. OOXML multi-Part source publication and DOC opaque-heavy owner
stage attribution remain larger follow-ups. iWork/IWA remains deferred while
other agents modify `iwa-*` crates.
