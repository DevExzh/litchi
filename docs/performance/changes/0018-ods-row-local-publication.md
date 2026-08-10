# ODS row-local worksheet publication

Date: 2026-08-11

Production base: `397abec7412cc7f5b2223865611f4cb06e65f77b`

Scope: packaged ODS same-topology worksheet cell publication only. ODT and ODP
are unchanged, and iWork/IWA crates were explicitly excluded.

## Hypothesis

A one-cell ODS worksheet commit parsed the source sheet into a typed draft and
then regenerated every worksheet row through `replace_tables`. On the large
32,768-cell corpus, the complete row writer and its validation accounted for a
material part of the 350+ ms changed-publication path even though only one row
was different.

The flat-ODS transaction already had a bounded row-local splicer. Reusing that
implementation for an eligible packaged edit should preserve untouched XML
byte-for-byte and serialize only the changed logical row, while retaining the
package and semantic readback boundary.

## Change

Packaged worksheet commit now attempts a private row-local publication when
worksheet count, names, styles, and row anchors retain the source topology. It
compares the parsed source and candidate sheets, validates only rows that will
be regenerated, emits those rows through the existing bounded row writer, and
copies every untouched source span exactly. The shared helper borrows staged
sheets rather than cloning them.

Eligibility and safety are deliberately conservative:

- a structural worksheet change, rename, style change, or row insertion with
  no physical source anchor falls back to the established full-table writer;
- an unmodeled direct table child retains the established refusal;
- an untouched row may contain opaque markup and is copied exactly, while a
  changed row with opaque elements, attributes, comments, processing
  instructions, or doctype markup is refused before publication;
- row validation reconstructs the original ancestor start tags so namespace
  bindings declared on the table remain available;
- source, modeled-sheet, XML-depth, generated-row, and total-output limits
  remain enforced.

The resulting `content.xml` still passes compactness validation. The package
is still rebuilt and reopened, the final worksheet snapshot is parsed, and
every typed sheet must equal the staged draft before the commit is returned.
Exact no-op bytes, source-checked patches, exact inverse restoration, resource
checks, and signed/encrypted mutation policy are unchanged. No public API,
dependency edge, executor, cache, durability contract, or format capability
changed.

## Matched latency measurement

Both binaries use the identical unchanged harness at the production base. The
before binary SHA-256 is
`45fc835674fcea43ccc9ee10956dd244cc2ca2e7c4ff7756dc7a7ca909e8bcb9`;
the after binary SHA-256 is
`3146676f551e652b436e2738a1a9832c5a2205eb58fe96ad53f956d31a0a7728`.

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic large ODS contains two 128 by 128
worksheets and 32,768 cells. Its 98,892-byte archive has SHA-256
`7f0c43561602aedac7c5e91915f55b3515371d327ae69ac7fc0fe42b655db3f2`;
the selected 54-byte cell payload has SHA-256
`df7ce48fc58e88d5db56d2b3e286d5a61e50fde7ec89f563cd3b4c129c04bd02`.

The primary ABBA run used three warmups and 30 samples per leg. Pooling the two
legs gives 60 raw samples per state; pooled statistics are recomputed from the
samples rather than from leg medians.

| Large ODS one-cell edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 359.011 ms | 324.774 ms | **-9.54%** |
| p95 | 375.180 ms | 333.435 ms | **-11.13%** |
| mean | 359.088 ms | 325.616 ms | **-9.32%** |

The approximate independent-sample 95% interval for the mean delta is
`[-10.02%, -8.62%]`. Matched A and B p50 comparisons improved by 10.30% and
7.69%, respectively. Within-state p50 drift was 2.88% before and 0.04% after.

Raw primary reports and their SHA-256 digests:

- `abba-ods-row-splice-one-edit-before-a.json`:
  `258c30c7a0a69136db4db9e6832bfa6adc76aaac5a09bb71d82b88b9f1ee62c7`
- `abba-ods-row-splice-one-edit-before-b.json`:
  `3033ca73bea1f6d41abd99c3e87228368412f4d0d6cfe70641f1010d9e9f8f51`
- `abba-ods-row-splice-one-edit-after-a.json`:
  `570d54d890f4ca1cee6ee4fb8d559f2b8ae5cbc4f4a91c17e37016914151248b`
- `abba-ods-row-splice-one-edit-after-b.json`:
  `5197462165423afac16d35c6581042a4fac5251ed42b2e941a651ca28850cc56`

The independent medium-corpus ABBA run used the same three warmups and 30
samples per leg. Pooled p50 improved from 21.959 ms to 20.374 ms (-7.22%),
p95 improved 9.50%, and mean improved 7.60%. Raw reports are the
`abba-ods-row-splice-medium-*.json` files beside the primary reports.

## Guardrails

An independent large-input ABBA run used three warmups and 30 samples per leg
for five public operations. The list and one-cell selectors are nanosecond
operations over an already-opened model, so their relative tail movements are
timer noise; their means stay within 3% and the changed branch cannot execute
in any guardrail.

| Guardrail | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| Open | 47.466 ms | 48.663 ms | +2.52% | +2.10% | +0.06% |
| List sheets | 140 ns | 160 ns | +14.29% | -2.61% | +7.28% |
| One cell | 1.562 us | 1.663 us | +6.47% | +2.86% | +17.85% |
| Full cell text | 3.186 ms | 3.182 ms | -0.13% | -0.93% | -4.14% |
| Exact no-op edit/save | 66.084 ms | 66.592 ms | +0.77% | +0.72% | -0.36% |

Both states reported identical input and selected-payload hashes. The harness
also verified the complete logical grid, exact no-op bytes, changed bytes,
forward patch, inverse restoration, diagnostics, and full snapshot reopen.
Raw reports are the `abba-ods-row-splice-guards-*.json` files.

## Allocations, RSS, and hardware counters

Matched one-sample Heaptrack processes used the same large one-edit workload.
These are whole-process totals and include corpus construction and exhaustive
post-timing verification in both states:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 6,179,548 | 5,818,058 | -5.85% |
| Temporary allocations | 1,464,538 | 1,366,130 | -6.72% |
| Peak heap | 65.23 MiB | 47.50 MiB | -27.18% |
| Heaptrack RSS | 73.99 MiB | 69.32 MiB | -6.31% |
| Leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

Uninstrumented GNU Time ABBA runs used three warmups and 30 samples per leg.
Maximum RSS was 69,172/74,456 KiB before and 65,200/65,200 KiB after, so there
is no measured RSS regression.

Matched `perf stat` ABBA runs over the same 33 iterations per leg reported:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 27,948 ms | 26,500 ms | -5.18% |
| cycles | 137,159,283,239 | 129,654,401,225 | -5.47% |
| instructions | 557,168,068,879 | 518,620,095,329 | -6.92% |
| branches | 118,430,561,901 | 108,373,506,227 | -8.49% |
| branch misses | 78,046,422 | 76,538,744 | -1.93% |
| cache references | 3,747,087,671 | 3,486,590,780 | -6.95% |
| cache misses | 200,897,529 | 187,682,084 | -6.58% |
| page faults | 434,129 | 403,677 | -7.01% |
| CPU migrations | 0 | 0 | unchanged |

The counter direction agrees with removing complete table serialization from
the eligible changed path. The remaining dominant samples are XML namespace
resolution, event reading, typed worksheet parsing, and package/readback
validation; those mandatory layers were not weakened.

Raw evidence is in `perf-ods-row-splice-*.csv`,
`time-ods-row-splice-*.txt`, and
`heaptrack-ods-row-splice-{before,after}.txt`.

## Correctness verification

- the new regression proves byte-exact preservation of an untouched opaque
  row, durable typed readback of a neighboring edited row, exact inverse
  restoration, and refusal when the opaque row itself is touched;
- existing worksheet transaction coverage proves structural add/move fallback,
  exact-source patching, stale-source refusal, and mutation atomicity;
- the complete `litchi-ods --all-features` test and doctest suite passed: 236
  tests, including security, compactness, malformed-input, resource, real-corpus,
  and document-transaction coverage;
- warning-denied production-library clippy and the changed integration-test
  clippy passed;
- the unchanged benchmark harness's 23 tests and warning-denied clippy passed;
- all benchmark JSON parsed successfully, and `git diff --check` and formatting
  checks passed.

The broader warning-denied ODS all-target clippy command remains blocked by
six pre-existing lints in unrelated test/module code. Warning-denied ODS
rustdoc remains blocked by the pre-existing broken `super::Cell::text` link in
`model/hyperlink.rs`. Neither blocked file changed in this batch.

The final compactness audit, package reopen, snapshot parse, and complete typed
sheet readback remain the publication boundary. A later ODS optimization
should target package parsing or shared package bytes rather than remove any of
those retained validation layers.
