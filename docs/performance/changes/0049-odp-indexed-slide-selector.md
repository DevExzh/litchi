# Change 0049: ODP indexed slide selector

Date: 2026-08-11
Production base: `84168afab9c6110e3737a677cce85c40326987e5`
Status: accepted

## Hypothesis and benchmark correction

`Presentation::slide(index)` called `Presentation::slides()`, retained every
semantic `Slide`, and then selected one value. The nominal
`odp_semantic_one_slide` benchmark repeated the same collection call directly,
so it did not exercise the public selector.

The benchmark now times `Presentation::slide(index)` and keeps package open,
complete semantic verification, and the expected-text comparison outside the
timed interval. The control executable was frozen after that benchmark-only
correction and before the parser implementation changed. The hypothesis was
that a full validating XML scan could retain semantic content only for the
requested page without changing error, ordering, namespace, style, limit, or
EOF behavior.

## Implementation and compatibility boundary

The private ODP page parser is compile-time specialized for collection and
indexed retention. The existing `slides()` and ODG/ODS parser callers use the
`SELECT_ONE = false` instantiation. `Presentation::slide(index)` uses the
`SELECT_ONE = true` instantiation and retains at most one completed slide.

For non-selected pages, the indexed instantiation still:

- resolves drawing-page transition styles from all of `styles.xml` and
  `content.xml`, including unused inheritance errors;
- scans content through `Event::Eof` and performs the same namespace-aware XML,
  attribute, hyperlink, plugin, enhanced-geometry, shape-tree, 3D, modern and
  legacy animation validation;
- maintains the document-global shape and animation counters and nesting
  limits; and
- decodes text, CDATA, references, and bounded ODF whitespace controls.

It does not normalize/store non-target paragraph text, attach completed child
shapes, or construct completed non-target `Slide` values. The target retains
the same title, body, notes, transition, animations, nested shapes, links,
media metadata, drawing attributes, and zero-based index as `slides()[index]`.
Empty and out-of-range selectors, including `usize::MAX`, return `Ok(None)`
only after the complete validating passes succeed.

No public type, dependency, feature, runtime, cache, lock, unsafe-code,
publication, patch, signature, encryption, source-version, or sink contract
changed. This is an in-memory semantic selector, not positional ZIP/XML I/O or
early termination.

## Matched latency evidence

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86-64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`. The deterministic
large ODP has 100 slides, 8,700 logical payload bytes, five archive members, a
3,424-byte archive, and archive SHA-256
`afb69ac66dffbc9f3ef19db360161af636abb818bfce689c6f4964fc521778c6`.
Every timed query selects slide 50 and compares its full visible text with the
generated expectation before recording the sample.

The control executable SHA-256 is
`f032c946580051baa8bbc36b99e60e6dfa9564040adcfe126be0ad6e1f879f1f`;
the candidate is
`c8f9597076a8f1a00b4a3001546f1b031a9f5643b4dcbd763d51cacb4dd37f7e`.
Their `.text` section hashes are respectively
`05dd7c1ddd5166c77021e58bd498964ae3862250f5d1b903613173b31dc45456`
and
`8baabbcb830459a0468ec3a44b35a354238006598bb2d116201db04feb3698ca`.

Ten order-balanced pairs used 50 warmups and 1,000 measured samples per leg.
Pooling 10,000 samples per state gives:

| Large ODP middle slide | Collection then select | Indexed selector | Delta |
|---|---:|---:|---:|
| p50 | 1.019059 ms | 0.977380 ms | **-4.09%** |
| mean | 1.040605 ms | 0.996934 ms | **-4.20%** |
| p95 | 1.154716 ms | 1.094953 ms | **-5.18%** |
| p99 | 1.477091 ms | 1.543663 ms | +4.51% |

The approximate independent-sample 95% interval for the mean delta is
`[-4.43%, -3.96%]` of the control mean. The p99 movement remains below the 5%
review threshold and is disclosed rather than hidden in an aggregate.

Ten balanced-order tiny/medium pairs used 50 warmups and 2,000 samples per leg:

| Shape | Slides | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|---:|
| tiny | 3 | 36.580 us | 36.701 us | +0.33% | +0.12% | -0.79% |
| medium | 12 | 125.162 us | 123.220 us | -1.55% | -0.94% | +1.75% |

The medium p99 is scheduler-sensitive across individual legs (149.7-218.9 us
for the candidate and 149.7-211.8 us for the control); its pooled +28.3 us is
not treated as a stable regression because p50, mean, and p95 are within 1.8%
and the leg ordering does not reproduce a consistent direction.

## Allocation, memory, and counters

Matched Heaptrack processes used 20 large samples. The whole process also
contains deterministic corpus creation and an unchanged complete semantic
verification after every query, so the selector savings are diluted:

| Heaptrack process metric | Before | After | Delta |
|---|---:|---:|---:|
| allocation calls | 567,244 | 545,364 | **-3.86%** |
| temporary allocations | 257,458 | 249,538 | **-3.08%** |
| peak heap | 696.21 KiB | 696.21 KiB | flat |
| profiler peak RSS | 12.41 MiB | 12.23 MiB | -1.45% |

The identical 1.78 KiB profiler/runtime leak remains. Uninstrumented GNU Time
ABBA processes with 20 warmups and 500 samples report 30,848 KiB in both
control legs and 30,848/30,976 KiB in candidate legs; the maximum movement is
+0.41% and is flat within the guardrail.

Matched process-wide `perf stat` ABBA runs at the same 500-sample setting give:

| Counter, A+B | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 3,749.15 ms | 3,484.83 ms | -7.05% |
| cycles | 18,356,507,345 | 17,244,414,941 | -6.06% |
| instructions | 78,890,631,707 | 77,947,762,668 | **-1.20%** |
| branches | 19,062,805,518 | 18,827,837,024 | -1.23% |
| cache references | 314,121,698 | 275,188,120 | -12.39% |
| cache misses | 11,312,720 | 12,299,342 | +8.72% |
| page faults | 17,261 | 17,276 | +0.09% |

The cache-miss increase is disclosed; it does not produce a latency, heap, or
RSS regression. Branch-miss counts are omitted from the table because the
first control leg was a 2.6x outlier while the other three legs were tightly
clustered, so pooling it would imply a code effect unsupported by repetition.

## Guardrails and verification

A balanced large-corpus guard uses 20 warmups and 200 pooled samples per state;
the slower save cases use a separate 500-sample-per-state ABBA run:

| Unchanged case | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|
| open | -9.04% | -16.13% | -39.85% |
| list slides | -0.69% | -1.07% | -6.04% |
| full text | +1.97% | +0.97% | -4.08% |
| exact no-op edit/save | +0.42% | +0.71% | +0.87% |
| one edit/save | +0.19% | -0.12% | -1.17% |
| media text-box edit/save | -1.17% | -0.86% | -1.72% |

Focused tests prove parser and public-facade parity for every present index,
empty/out-of-range/`usize::MAX`, a semantic failure after a valid selected
page, and cyclic transition styles. Existing suites retain prefix aliases,
titles/body/notes, transitions, shapes/groups/3D/geometry, hyperlinks, inert
events/media, modern and legacy animation, malformed/truncated input, global
resource limits, exact no-op, patch/inverse, signatures, encryption, media,
real Impress packages, and complete transaction readback.

The complete all-feature ODP suite passes (123 library tests, all integration
suites, and 21 doctests), as do warning-denied all-target/all-feature ODP
Clippy, warning-denied rustdoc, the 32-test warning-denied benchmark harness,
formatting, JSON parsing, and `git diff --check`. The current ODF fuzz package
(detector and ODT targets) compiles offline; there is no dedicated ODP fuzz
target in the tree. The workspace-wide iWork/IWA gate was not rerun because
those crates are explicitly excluded while other agents modify them.

## Evidence and remaining work

Committed evidence includes:

- `odp-indexed-slide-large-{forward,reverse}-*.json` for the headline result;
- `odp-indexed-slide-small-{forward,reverse}-*.json` for tiny/medium scaling;
- `odp-indexed-slide-{guards,save-guards}-*.json` for unchanged paths;
- `odp-indexed-slide-{before,after}-heaptrack.txt` for allocation and peak
  memory attribution;
- `odp-indexed-slide-perf-*.csv` and `odp-indexed-slide-time-*.txt` for
  hardware counters and maximum RSS; and
- the executable and `.text` hashes above for control identity.

Broader ODF work remains positional package/XML reads, repeated slide queries,
non-text and structural edits, resource-adding publication, real-producer
media/security matrices, and final-result copy attribution. OLE2, OOXML, RTF,
and iWork/IWA production code are unchanged by this batch.
