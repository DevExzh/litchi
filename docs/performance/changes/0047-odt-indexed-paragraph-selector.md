# Change 0047: ODT indexed paragraph selector

Date: 2026-08-11
Production base: `ae5d750174c8acc3831c2194427bd59768a73a3c`
Status: accepted

## Hypothesis

The public ODT one-paragraph workflow had to call `Document::paragraphs()` and
retain every structured paragraph before selecting one value. On the large
semantic corpus that meant constructing 10,000 paragraph elements and strings
for a single middle-paragraph read. The XML still has to be scanned through
EOF: returning as soon as the target closes would skip malformed trailing XML,
resource limits, attributes and suppressed-content structure.

The hypothesis was that a full validating scan which retains only the selected
paragraph would materially reduce time and peak memory without weakening the
ODF decoder contract or changing the established all-block path.

## Implementation

`Document::paragraph(index)` and the lower-level
`TextElements::parse_paragraph_at(xml, index)` are additive selectors. Paragraph
indices are zero-based, exclude headings, and keep the existing start order for
paragraphs nested in frames, annotations, sections and tables.

The selector scans the complete XML and applies the same namespace resolution,
tracked-change/note/ruby suppression, nesting depth, paragraph-plus-heading
count, per-block text, aggregate text, whitespace-control, entity, attribute
value and duplicate-expanded-attribute validation as the collection parser.
Non-target blocks retain only their text-byte count while open; the selected
paragraph alone retains an `Element` and decoded `String`. An absent index
returns `Ok(None)` only after successful EOF validation.

The benchmark's `odt_semantic_one_paragraph` case now calls the public selector.
Corpus creation, package open, full semantic verification and the expected-text
comparison stay outside the timed interval.

An initial shared runtime-mode parser improved selection, but made the unchanged
10,000-paragraph list guard about 6% slower. It was fully removed. The accepted
design leaves the established all-block parser and its structured/full-text
callers isolated from the selective retention state.

## Matched latency measurement

The before release executable SHA-256 is
`2f4ad1f6d30950e3445d2f468927b8505602720fff5945c036852996705a68c4`;
the measured accepted-after executable SHA-256 is
`bd2c5e3f5d8d9775416a7ee74d61070e5217c424882858d345a1e93b8e2a2047`.
The final executable rebuilt after test-only and documentation additions is
`3cf6c077ff116355558724e51038382b01894876ece10ab8c62c1874e71271b8`;
its `.text` section is byte-identical to the measured after executable. The
before and after `.text` section hashes are respectively
`825a650486d81550d4053a05be7602355dfece752cad955c7dd45e99c62a4d93`
and
`d052acef3b92331c64a4b89fb4304fb03d7ab53e845f52335a69bfa10f68994a`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86-64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`. The deterministic
large ODT has 10,000 paragraphs, 490,000 logical text bytes, a 28,420-byte
archive, and archive SHA-256
`9d724c649cb5e4b4adce30c4ede2059ff9efc26109c1b84ac8460df00ecf89a9`.
The selected paragraph is index 5,000, and every iteration verifies it against
the generated expected text before recording the sample.

Four balanced control-A/candidate-A/candidate-B/control-B cycles used 50
warmups and 250 measured samples per leg. Pooling 2,000 samples per state gives:

| Large ODT middle paragraph | Collection then select | Indexed selector | Delta |
|---|---:|---:|---:|
| p50 | 3.201880 ms | 1.647189 ms | **-48.56% (1.94x)** |
| mean | 3.232256 ms | 1.669948 ms | **-48.33%** |
| p95 | 3.490128 ms | 1.818835 ms | **-47.89%** |
| p99 | 3.873926 ms | 2.085393 ms | **-46.17%** |

The approximate independent-sample 95% interval for the mean delta is
`[-48.63%, -48.04%]` of the before mean. Every within-cycle A/B p50 spread is
below 3.1%; all candidate-leg p50s span 1.620-1.678 ms. The control span across
all four cycles is 3.091-3.270 ms, and all samples are retained rather than
selecting favorable cycles.

The selector also improves smaller public queries:

| Shape | Paragraphs | Before p50 | After p50 | Delta |
|---|---:|---:|---:|---:|
| tiny | 24 | 10.033 us | 6.028 us | -39.92% |
| medium | 200 | 64.079 us | 33.326 us | -47.99% |

Each smaller cell pools 1,000 samples per state from a separate balanced ABBA
run with 50 warmups and 500 samples per leg.

## Allocation, memory and counters

Matched Heaptrack processes used five warmups and 20 large samples. The whole
process includes deterministic corpus construction and two unchanged complete
semantic verification passes per iteration, so the selective-query savings are
diluted:

| Heaptrack process metric | Before | After | Delta |
|---|---:|---:|---:|
| allocation calls | 5,547,918 | 4,047,369 | **-27.05%** |
| temporary allocations | 552,185 | 551,886 | -0.05% |
| peak heap | 22.92 MiB | 17.25 MiB | **-24.74%** |
| profiler peak RSS | 39.07 MiB | 33.09 MiB | -15.31% |

The identical 1.78 KiB profiler/runtime leak remains. Uninstrumented GNU Time
ABBA processes with 20 warmups and 500 samples report 34,632 KiB before and
30,848 KiB after in both legs, a **10.93%** maximum-RSS reduction.

Matched `perf stat` ABBA processes at the same 500-sample setting give:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 14,952.0 ms | 12,919.9 ms | -13.59% |
| cycles | 73,445,143,754 | 63,162,356,427 | -14.00% |
| instructions | 261,075,854,367 | 227,520,168,388 | -12.85% |
| branches | 55,120,671,680 | 48,285,546,712 | -12.40% |
| branch misses | 40,628,650 | 38,617,276 | -4.95% |
| cache references | 4,884,823,413 | 3,317,704,345 | -32.08% |
| cache misses | 199,638,522 | 159,299,215 | -20.21% |
| page faults | 1,421,600 | 1,418,295 | -0.23% |

## Guardrails and correctness

One balanced large-corpus ABBA guard used 20 warmups and 100 samples per leg:

| Unchanged case | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|
| open | +0.61% | +1.24% | +2.99% |
| list paragraphs | +0.38% | -0.25% | +0.71% |
| full text | -3.04% | -3.47% | -4.56% |
| exact no-op edit/save | +101 ns / +3.82% | -63 ns / -2.23% | +0.43 us |
| one edit/save | +0.20% | -0.55% | -1.95% |
| 1% edit/save | -0.78% | -0.24% | +2.20% |

The nanosecond no-op distribution is disclosed in absolute units; its p95
percentage is not treated as a material regression. All other p50, mean and p95
guards remain within 5%.

Focused tests prove public parity and out-of-range behavior, heading exclusion,
nested start order, note/ruby/tracked-change suppression, malformed trailing
XML, a late over-limit whitespace run, and a late duplicate expanded
attribute. The selector always scans through EOF. The existing all-block,
real-producer nested-frame, transaction, media, package and security suites
remain unchanged.

The all-feature ODT library suite, integration suites and doctests pass, as do
warning-denied all-target/all-feature Clippy, warning-denied rustdoc, the ODF
fuzz build, the benchmark harness tests and Clippy, formatting, JSON parsing and
`git diff --check`. The existing CI ODF smoke and release matrix already invoke
`odt_semantic_one_paragraph`; the same case now exercises the indexed public
API. No dependency, feature, unsafe-code, cache, lock, source-version,
publication, patch, signature or encryption boundary changes. OLE2, OOXML and
RTF production code are unchanged, and iWork/IWA is excluded.

## Evidence

- `abba-odt-indexed-paragraph-repeat-{1,2,3,4}-*.json`: headline ABBA cycles;
- `abba-odt-indexed-paragraph-size-*.json`: tiny/medium scaling;
- `abba-odt-indexed-paragraph-guards-*.json`: unchanged-path guardrails;
- `odt-indexed-paragraph-{before,after}-heaptrack.txt`: allocation and peak
  memory attribution;
- `odt-indexed-paragraph-time-*.txt`: uninstrumented maximum RSS; and
- `odt-indexed-paragraph-perf-stat-*.csv`: process hardware counters.

Raw reports retain the harness environment, corpus hash, complete latency
samples and confidence intervals. The broad remaining ODF work is positional
source-backed reads, repeated ODP selectors, non-text/structural edits and
real-producer media matrices; this change makes no broad lazy-read claim.
