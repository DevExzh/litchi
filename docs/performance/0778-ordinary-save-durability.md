# 0778 — preserve DOCX note relationships and measure save durability

Status: final captures complete; independent artifact and statistical replay pass.

This batch fixes lost DOCX footnote/endnote relationships after an unrelated
paragraph edit, discovered by an independent preflight preservation check.
It also measures current-source ordinary DOCX, XLSX and PPTX saves with the
default policy, explicit `Full`, `FileOnly` and `NoSync`. Preflight exposed and
prompted a DOCX note-relationship preservation fix; the default durability
policy is unchanged. These policies intentionally offer different
crash and power-loss persistence guarantees; a lower observed latency for a
weaker policy is not an optimization of the default contract.

The experiment follows [0714](0714-docx-atomic-publication-attribution.md), which
located the synchronization boundary in ordinary DOCX publication, and
[0773](0773-save-durability-integration.md), which integrated the explicit
policy and verified its syscall sequences without claiming a latency ratio.
Historical timings are context only and are not used as a baseline here.

## Scope and method

The source is based on `1dff77b6df`. The harness adds an untimed artifact exporter
that reuses the existing corpus, edit, filesystem-save and sequential-output
paths. The timed runner is unchanged. The packet records the exact source,
dependency locks, binaries, fixtures, host, commands and scripts.

The frozen matrix covers the three fixed generated-medium corpora plus
`NumberedList.docx`, `alt-chunk-header.docx`,
`ConditionalFormattingSamples.xlsx` and `slide-section-test.pptx`. The altChunk
DOCX intentionally refuses the edit and must preserve its source archive
exactly. It is a refusal control, not a successful mutation measurement.

Four process blocks use balanced policy orders for lifecycle and atomic
publication. Each native child collects 100 samples after ten warmups. Edit
and counting-sink controls are collected separately. Allocation observations,
syscall traces and whole-process user counters use separate processes. RSS is
whole-process peak memory, not operation-region live allocation. Phase
quantiles and allocation peaks must not be added or subtracted to reconstruct
lifecycle costs.

Every timed filesystem save starts with an absent destination. Corpus setup
performs default-policy reference publications before timing and warms source
data. This is not existing-destination, physically cold, remote-source,
concurrent, large-corpus or device-floor evidence. The generic harness shape
flag does not enlarge these generated ordinary-save corpora.

## Independent output checks

The untimed exporter retains each source, all four filesystem-policy outputs
and a sequential-output control. The Python oracle uses ZIP and XML readers
independent of Litchi to check the named edit and preservation outside its
explicit mutation closure. Qualification binds those artifacts to the exact
corpora and output identities used by the timed executable. The exporter also
reopens every output through the public reader.

These checks are scoped to paragraph text, decoded cell values, shape text,
named relationship semantics and unchanged decoded members. They do not prove
full lexical, style, ZIP-metadata or external-application compatibility inside
rewritten parts. Refused DOCX source equality is a separate whole-archive check.
Relationship comparison includes type, target and normalized `TargetMode`.
The real XLSX content-type rewrite is checked narrowly: its printer-settings
`.bin` default becomes explicit overrides for the same 15 retained parts,
with every effective retained-part content type unchanged. The calculation
chain's part, relationship and override are the only permitted removals.

## Preservation correction found during preflight

The first exported `NumberedList.docx` lost the main document's `rId7`
footnote and `rId8` endnote relationships while retaining both note payloads.
The independent graph oracle rejected this output. The existing writer
unconditionally skipped note relationships while copying the main document's
graph, but a paragraph-only edit generated no replacement note XML. Keeping
the payloads alone did not preserve their package reachability.

The writer now computes replacement footnote/endnote XML once before copying
relationships. It preserves each existing note edge if that note collection
has no generated replacement, including its ID, target and target mode.
When authored notes generate a replacement, either Transitional or Strict
source relationship types are recognized and replaced by the existing
generated-note route. This avoids retaining an old Strict edge alongside a
new generated edge. The surrounding rollback guard still covers generation
and publication failures.

Regressions serialize and reopen a source with custom note targets before
appending, verify complete relationship tuples and byte-exact note payloads,
and exercise authored footnotes and endnotes replacing existing Strict edges.
These are relationship-type controls, not a claim of full Strict-document
conversion support. The rejected preflight is retained under `preflight/`;
its outputs and qualification records are not used for the final matrix.

## Results

The final capture contains 280 native processes (28,000 timed samples), 140 separate allocation processes, 28 syscall traces and 32 user-counter processes plus counter qualification. No capture was rerun or discarded for variation. The following values are medians of four process p50s, in milliseconds; each process has 100 samples. Default omits the policy flag.

### Ordinary lifecycle

| Corpus | Default | Full | FileOnly | NoSync |
|---|---:|---:|---:|---:|
| Generated DOCX | 5.758 | 5.763 | 3.894 | 0.865 |
| Generated XLSX | 14.257 | 14.265 | 11.232 | 4.812 |
| Generated PPTX | 6.860 | 6.872 | 5.011 | 1.756 |
| NumberedList DOCX | 5.785 | 5.809 | 3.937 | 0.525 |
| altChunk DOCX (refused) | 5.927 | 5.939 | 4.070 | 0.495 |
| ConditionalFormatting XLSX | 10.488 | 10.471 | 8.387 | 3.256 |
| Slide-section PPTX | 14.291 | 14.338 | 12.444 | 8.577 |

### Atomic publication

| Corpus | Default | Full | FileOnly | NoSync |
|---|---:|---:|---:|---:|
| Generated DOCX | 5.304 | 5.305 | 3.446 | 0.422 |
| Generated XLSX | 11.347 | 11.338 | 8.389 | 1.963 |
| Generated PPTX | 5.225 | 5.240 | 3.370 | 0.127 |
| NumberedList DOCX | 5.433 | 5.421 | 3.573 | 0.187 |
| altChunk DOCX (refused) | 5.447 | 5.445 | 3.596 | 0.042 |
| ConditionalFormatting XLSX | 7.722 | 7.722 | 5.618 | 0.482 |
| Slide-section PPTX | 5.982 | 5.982 | 4.138 | 0.260 |

Full remains the default. Lower latencies under FileOnly and NoSync accompany intentionally weaker persistence guarantees. The 28 traces confirm one atomic replacement per measured save: default/Full each perform one file and one parent-directory sync, FileOnly only the file sync, and NoSync neither. Trace timing is diagnostic and is not substituted for native timing. These observations establish neither a default-contract speedup nor a device latency floor.

All 50 process-spread flags above 5% are retained: 40 p99, eight p95 and two mean flags; no p50 or whole-process RSS spread is flagged. The largest is 58.64% for the refused DOCX counting-sink p95. Default versus explicit Full has 14 flagged metric series, all in p95/p99, despite identical durability semantics. Those tails demonstrate variation and are not optimization claims. See [all spread flags](results/change-0778/spread-flags.csv), [all equivalent-policy controls](results/change-0778/default-full-controls.csv), and [70 groups including edit/counting controls](results/change-0778/native-summary.csv). [All 280 process summaries](results/change-0778/native-processes.csv) retain means and tail quantiles. Spread is `(max − min) / min`; no pooling of process quantiles is used.

### Allocation and user work

The allocation lane uses two processes per group, three samples each, without warmups. For each lifecycle/publication case, all four policies have identical allocation calls, cumulative allocated bytes, peak bytes above region entry and net live-byte change. All repeated metric p50s have zero spread. Small absolute region-peak offsets between policies already exist at region entry and do not indicate an operation memory change. Whole-process RSS and region allocation are separate observations.

The table gives default-policy lifecycle allocation. Bytes are cumulative requested allocation, not copied bytes or RSS; peak above entry is not a phase-additive quantity. Owner destruction occurs outside the measured region.

| Corpus | Allocation calls | Allocated bytes | Peak above entry (bytes) |
|---|---:|---:|---:|
| Generated DOCX | 15,571 | 3,166,487 | 1,066,433 |
| Generated XLSX | 29,845 | 1,100,160,722 | 6,466,398 |
| Generated PPTX | 13,488 | 7,046,682 | 907,099 |
| NumberedList DOCX | 4,693 | 2,559,705 | 728,659 |
| altChunk DOCX (refused) | 5,466 | 2,189,479 | 263,282 |
| ConditionalFormatting XLSX | 27,515 | 39,278,130 | 2,390,762 |
| Slide-section PPTX | 56,813 | 16,706,874 | 2,050,945 |

[All allocation metrics](results/change-0778/allocation-summary.csv) include atomic publication and edit/counting controls. The generated XLSX's 1.10 GB cumulative allocation is a follow-up investigation opportunity; it does not imply that much live memory or a proven removable copy cost.

User counters compare whole processes with 3 versus 23 samples, using `(counter23 − counter3) / 20` for each of two repeats. This includes incremental report work and is not an exact operation-region count. All four events report 100% running time. Default-to-NoSync mean marginal user-instruction changes are −0.007% (generated DOCX), +0.131% (generated XLSX), +0.034% (generated PPTX), and −0.008% (NumberedList). User instructions are essentially unchanged in these observations; kernel work is outside these counters. Cycles vary substantially and are retained without a causal precision claim. [Both repeats and all four events](results/change-0778/counter-summary.csv) are retained; counters are never divided by separately measured elapsed time.

The host is shared and the measurements are descriptive. Broader cold/range/concurrency/scaling coverage and complete CRUD validation remain open. No CRUD registry row is promoted, and the non-iWork goal remains active.


## Quality and admission history

Final source `0b4799d649` passes all 14 gates: five all-feature DOCX owner gates
and nine harness/build gates. The final test commands report 2,492 passed and
33 ignored across 67 suites. All seven exported corpora pass the independent
oracle, all 28 qualification processes pass, and policy plus stream outputs
are byte-identical for each corpus. The refusal control remains source-exact.

The rejected preflight passed nine quality/build gates at `3cd8e197b0`: formatting, all-target
checking, binary tests, warning-denied Clippy, warning-denied rustdoc, source
policy, coverage-index validation, and the native/export and allocation builds.
Its binary suites passed 61 tests across 21 suites, and its focused exporter
test passed. These are preflight results, not final-fix test counts.

The full library suite passed 556 tests with one ignored on `2c43a3dd98`.
The following binary suite in that attempt failed a global-allocation live-byte
assertion under two test threads, followed by two poisoned-lock failures. The
binary suites passed with one test thread. A second attempt then found two
Clippy warnings in the new exporter. That preflight source differed from its full-library-tested predecessor only
by replacing two lazy optional copies with `then_some`; a source-bound diff
check and focused retest retain that distinction. The later production
preservation fixes receive fresh owner and harness gates. Failed attempts remain in the packet.

The first independent admission attempt rejected generated XLSX's workbook
recalculation metadata because the draft oracle's mutation closure omitted
the existing calculation-invalidation contract. The actual changes are an
exact `calcPr` replacement/addition and, for the real fixture, removal of the
calculation-chain part, relationship and content-type override. Admission must
check that precise closure and preserve all unrelated workbook and graph
semantics; passing qualification alone is insufficient.

The first DOCX owner run passed after preserving untouched notes. Review then
added Strict relationship replacement coverage. That new regression initially
used plain-text note payloads and failed XML publication before reaching its
assertions; the fixture was corrected to valid XML. Final DOCX owner gates at
`0b4799d649` pass, including 1,936 tests and 32 ignored across 65 suites.
Both the earlier passing owner run and the invalid-fixture attempt are retained.

## Integration and replay

The [packet README](results/change-0778/README.md) gives offline replay commands.
Source, fixture and executable identities are checked before removing the owned
build target; `cleanup.json` retains exact executable witnesses. The sealed
packet retains raw captures, rejected preflight evidence, independent reviews
and derived tables. Worktree removal is recorded after integration below.

The historical [next-step note](results/change-0778/next-step.md) is superseded
by the [0779 current-source correction](0779-opc-bounded-input-growth.md): ZIP
first-read coalescing already landed in 0611, with later structural prefetch.
The sealed note is retained as history and is not an outstanding implementation
recommendation.

Integrated by fast-forward at `8f6ca3af97`. The owned target and filesystem
root, copied root lock, three reference symlinks, worktree and temporary branch
are removed. Offline validation and all six derived-table checks pass from
main after the original worktree is gone, including the complete packet seal.
All three unrelated main-file hashes and all pre-existing worktree records
remain unchanged. The main dependency lock and reference checkouts are retained.
