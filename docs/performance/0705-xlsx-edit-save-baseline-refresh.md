# 0705 — refresh XLSX source-backed edit/save attribution

`performance_claim: none`

This diagnostic returns to XLSX after the bounded PPTX retention work in
[0704](0704-pptx-bounded-slide-mce-retention.md). It measures current source
without a production change. The historical [0520](changes/0520-xlsx-source-edit-phase-attribution.md)
phase breakdown predates unchanged-cell readback in
[0525](changes/0525-xlsx-unchanged-cell-readback.md), shared planning traversal
in [0546](changes/0546-xlsx-shared-traversal-retained.md), layout facts in
[0622](0622-xlsx-compact-source-facts.md), and fact-builder/style work in
[0635](0635-xlsx-facts-builder-and-chains.md). Those older timings are not a
current bottleneck ranking or a matched baseline for a speedup claim.

The existing source-backed one-percent scalar-cell selector separates open,
planning, staged sets plus commit, and sequential publication. It uses fresh
editor/cache instances over an instrumented in-memory positional provider.
The sum excludes sink setup, remaining handle destruction, output reopen and
oracles. Publication includes dropping its returned snapshot. The synthetic
shapes are medium and dense-sparse; all four worksheets are selected, so this
does not establish selective-subset scaling or physical cold-cache behavior.

A separate producer control edits a numeric cell on the planning sheet in
the producer Edit variant, which contains a shared-string table. It does not
edit the shared-string sheet, and its combined timer is not a phase breakdown.

The [packet](results/change-0705/README.md) retains source and binary identities,
commands, raw samples, independent analysis and validation. Native timing and
allocator instrumentation run in separate children. Operation-region allocator
peaks are neither whole-workflow peaks nor RSS, and phase peaks are not added.

The previous emitted-output parser fusion remains rejected under
[0516](changes/0516-xlsx-output-fusion-rejection.md). No new parser, cache or
validation shortcut is justified solely by the historical cost breakdown.
OLE2/OOXML work and the broader GOAL remain active; iWork is excluded.

## Current native results

Four CPU-12-pinned children retain 100 samples after 20 warmups each.
The second repetition reverses shape order. Phase shares divide summed
phase times by summed matched totals; they are not sums of percentiles.

| Shape / repeat | Total p50 ms | p95 ms | p99 ms | Planning share | Commit share | Publication share |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| medium / 1 | 16.656 | 16.809 | 16.831 | 37.62% | 21.02% | 40.87% |
| dense-sparse / 1 | 32.763 | 33.250 | 33.814 | 36.48% | 21.09% | 42.15% |
| medium / 2 | 16.659 | 16.832 | 16.938 | 37.65% | 21.04% | 40.81% |
| dense-sparse / 2 | 33.263 | 34.365 | 34.609 | 35.73% | 20.77% | 43.17% |

Publication is the largest measured phase, with planning close behind. Commit
is approximately one fifth of the interval. Two-repeat variation exceeds 5%
in five open-phase metrics, including dense-sparse p50 (+7.51%) and p99
(+39.34%). All flags remain in `analysis.json`; they are same-build variation,
not optimization regressions. The sample size does not establish cross-host
or universal tail behavior. Unrelated build activity was observed during the
build and one Cargo process immediately before capture; no continuous host
isolation is claimed and no causal explanation is assigned to variation.

Medium contains 9,216 cells and updates 93; dense-sparse contains 17,792 and
updates 178. Each measured workflow materializes six payloads, preserves twelve
untouched members, and emits at most 64 KiB per accepted sink write. Logical
source reads total 4,233,620 / 4,259,054 bytes respectively, including
publication transfer. All worksheets are selected; zero unselected reads is
not a selective-subset result. Corpus, output and resource observations repeat.

The producer control's two total medians are 1.924 / 1.939 ms, with identical
corpus, output and sink identities. Separate short native children observe
whole-child RSS of 64,420 / 64,048 KiB (medium) and 83,520 / 84,136 KiB
(dense-sparse). Those peaks include corpus construction and untimed oracles.
They are not operation-region retention measurements or memory improvements.

## Allocation attribution

Four separate instrumented children retain three samples after two warmups.
The table reports one operation-region sample per shape; relative counts and
bytes repeat across the children. Absolute process live gauges include retained
harness state and are kept separately in the raw evidence.

| Shape / region | Allocation calls | Requested bytes | Peak above region start | Net live change |
| --- | ---: | ---: | ---: | ---: |
| medium / plan | 67,845 | 12,139,796 | 3,191,768 | +2,200,064 |
| medium / staging | 389 | 1,899,381 | 28,521 | +12,227 |
| medium / commit_core | 42,383 | 3,205,114 | 2,058,609 | +2,046,254 |
| medium / publication | 19,197 | 4,126,154 | 1,095,243 | -15,633 |
| dense-sparse / plan | 129,411 | 15,850,404 | 7,602,062 | +4,200,215 |
| dense-sparse / staging | 725 | 2,916,044 | 35,828 | +25,944 |
| dense-sparse / commit_core | 80,110 | 5,968,832 | 3,931,745 | +3,902,769 |
| dense-sparse / publication | 36,573 | 5,747,529 | 1,604,554 | -15,633 |

Planning requests the most bytes and allocates 67,845 / 129,411 times on
these shapes. Publication requests fewer bytes despite taking the largest
native time share. This is evidence to investigate distinct owners, not a
reason to equate allocation count with latency. Staging and commit-core counts
reconcile to the combined commit region; their peaks are not additive. No
instrumented elapsed time is used for the native timing conclusions.

## Publication instruction attribution and next action

Four further fresh children collect Callgrind guest instructions only inside
`SourceBackedEditor::publish_multi_commit_to_stream`. Raw incoming edges
uniquely identify the measured call; five untimed lifecycle calls and the
termination dump are retained separately. The plan initially predicted three
lifecycle calls, omitting two foreign-source refusal entries. Its correction
is documented; no capture was discarded or substituted.

| Shape / repeat | Publication Ir | Source XML audits Ir | Audit share | Preservation writer Ir | Writer share |
| --- | ---: | ---: | ---: | ---: | ---: |
| medium / 1 | 95,565,632 | 52,136,751 | 54.56% | 42,725,977 | 44.71% |
| dense-sparse / 1 | 173,074,628 | 100,615,603 | 58.13% | 71,528,637 | 41.33% |
| medium / 2 | 95,566,310 | 52,137,480 | 54.56% | 42,726,238 | 44.71% |
| dense-sparse / 2 | 173,074,283 | 100,615,724 | 58.13% | 71,527,941 | 41.33% |

These are disjoint immediate children of the nested OPC topology writer.
The analyzer verifies raw edge counts and self-plus-direct-child accounting.
Nested XML attribute/parser costs overlap the audit branch and must not be
added to it. Guest instruction fractions are not native time fractions or
predicted speedups. Each profile reports Valgrind's `brk segment overflow`
notice but exits successfully with matching output and resource identities;
no allocator or RSS conclusion uses these profile processes. The scoped
method excludes the returned snapshot drop present in native publication.

The next investigation is the source-compatible XML auditor. There are ten
calls per measured publication: five original/replacement pairs. Current
`SourceBackedPackage` validates both sides before changed XML publication;
`validate_source_part_xml` delegates to `xml_minifier::audit::verify_source`.
The source-compatible grammar and budgets differ from XLSX's semantic parser.
Neither audit may be omitted merely because semantic readback succeeded.
Before an implementation, attribute the auditor's parsing and attribute work
and check existing typed source-proof routes for compatible work elimination.
Preservation writing remains the second substantial publication branch;
planning's allocations are a separate follow-up.

## Verification and limits

Both release builds succeed. The first full native build and the subsequent
workspace-input recheck produce the same binary digest. The independent
packet audit verifies all source/constraint hashes, native statistics and
phase alignment, allocator reconciliation, producer/RSS identities, and raw
publication call-edge accounting. The initial retained native analysis
predated its final `checks` metadata; the old derived file and failed replay
log are archived, and the corrected replay changes no raw measurement or
prior derived value.

All six evidence gates pass: crate boundaries, strict and structural claims,
report classification, CRUD coverage and the non-iWork gate manifest. The
existing executable semantic/preservation oracles pass throughout capture.
No Rust production or harness source changes, and no new full Rust test,
fuzz, native Office, cross-platform, physical cold-device or scaling campaign
is claimed. Claims and coverage registries are unchanged. Owned build and
analysis scratch is removed after capture, with binary identities retained
for post-cleanup audit. The broader GOAL remains active.
