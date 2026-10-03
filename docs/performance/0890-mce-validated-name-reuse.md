# 0890 MCE validated-name reuse

This batch evaluates repeated attribute-name validation in the legacy MCE
codec. The retained 0886 native profiles attribute 2.76–3.35% of sampled self
periods to `expand_parts`; the profile motivates a candidate but does not
predict its end-to-end benefit. The candidate is rejected: its paired native p50 ratio is 0.99381,
with a 95% bootstrap interval of 0.98861–1.01112. It fails both parts
of the frozen adoption gate. Production is restored exactly to the base.

## Candidate and preservation contract

The existing directive scan validates each ordinary attribute in source
order. The candidate records when that scan has successfully validated an
unprefixed name. The writer can then reuse its known empty namespace and
avoid a second QName validation. Recognized MCE directives also reuse the
existing keep flag so the writer can count their removal without expanding
their names again. These facts live only for the current start tag.

The candidate leaves the first validation point, duplicate checks, directive
token accounting, namespace declarations, and ProcessContent checks in place.
Inherited opaque elements bypass the directive scan and retain their original
writer validation. The additional boolean is checked against the former
attribute layout on the tested compiler. There is no new allocation, dependency,
public API, or persistent cache.

| Constraint | Evidence |
| --- | --- |
| ADR 0001: correctness before performance | Complete codec/stream outcomes and focused error-order regressions |
| ADR 0002/0024: owner boundaries | Two files in the existing shared OOXML MCE owner; unchanged dependencies |
| ADR 0005: measured work elimination | Matched release protocol and explicit adoption gate |
| ADR 0006: preservation and bounded refusals | Raw/semantic transcripts, complete reports, typed errors, and exact/one-over limit controls |

The [v2 source review](results/change-0890/source-review-v2.json) approves
measurement. The initial freeze and review remain retained; v2 adds an
ordinary malformed-unprefixed-name ordering case, with production code
unchanged between freezes.

## Correctness evidence

The baseline is commit `f790c2228cbd9e89c2892ba3ffaeb49802fac43b`.
Two baseline and two candidate processes each evaluate 20,000 generated seeds,
20,000 aliasing seeds, and 67 members from three real packages. Seed counts
include repeated inputs: the generated populations have 13,811 and 19,418
distinct input hashes respectively.

All four reports are byte-identical: 226,180,145 bytes, SHA-256
`53f4e5b747b9c019aae3922e7df1ef7f98bbbb4a047f74b467443e8d49eefe57`.
Each contains 200,201 structured MCE records, 200,405 legacy outcomes, and
10 required runtime controls. [Independent cross-version verification](results/change-0890/root-oracle-comparison.json)
checks exact populations, configurations, complete fields, and control
expectations. This combines independently specified small controls with broad
determinism/equality checks; it is not an independent full MCE implementation.

The 24 diagnostic tests and 151 focused MCE tests pass, including the new
layout and error-order regressions. Shared format suites pass 1,381 XLSX,
1,230 DOCX, and 851 PPTX tests. One existing PPTX test that prints a golden
transcript remains ignored. Warning-denied owner Clippy and rustdoc pass.
The continuation passes all 399 common-owner tests, isolated workspace
formatting, and formatting of the owner in the original workspace. The original workspace formatting command reports differences only
in the unrelated untracked `crates/litchi/examples/bench_word_parse.rs`; that
file is preserved, and tracked-source formatting is checked in isolation.
The [post-capture source checks](results/change-0890/attempts/candidate-diagnostic-f790c2228c-v2-retry1/post-candidate-source-verification.json)
bind the frozen library, diagnostic harness, fixtures, and binary identities.

Preparation failures are retained. An initial full-tree extraction exceeded
the temporary-directory quota before any build or workload; its owned partial
export was removed and replaced by a bounded source export. This event has a
[retrospective note](results/change-0890/source-preparation-retrospective.json),
not a contemporaneous command log. The first isolated baseline build omitted
two compile-time fixture inputs and failed before running the diagnostic.
The corrected baseline source includes those exact inputs. The first candidate
diagnostic test run passed 23 tests and failed the package-discovery fixture
lookup because its manifest path pointed into the bounded source archive.
The successful continuation builds the diagnostic tests using the original
repository manifest and fixtures, while the differential executable remains
built from the isolated candidate source. Source and binary identities are
checked after those operations. A continuation preflight was stopped after
review found that its source inventory was recomputed once per file; no test
had started. The corrected continuation computes that inventory once and
passes. Both observers of that single interrupted process retain their receipts.

## Release protocol

The before binaries are the retained 0885 release executables. Their selected
source tree differs from the current base only in the separate 0889 diagnostic
binary files, which are explicitly excluded from release workload reuse.
The candidate is reconstructed from that same retained source with the frozen
two-file overlay, using the same compiler and release settings.

The protocol retains the four 0887 scenarios: producer XLSX full-cell scan,
ordinary XLSX full-cell scan, selected DOCX paragraph, and selected PPTX slide.
Eight native paired blocks use three warmups and 50 samples per child; four
allocator paired blocks use three warmups and 15 samples for the primary
producer scan. The 72 children use alternating before/after order and CPU 12.
Whole-child ratios are the resampling unit; samples are not pooled.

The primary adoption gate is a paired native p50 median ratio at most 0.98
with the 95% bootstrap upper endpoint below 1.0. Per-scenario tail latency,
RSS, and spread flags remain visible for review. Each role/mode executes one
freshly created, byte-verified binary copy, reused across its children. This
controls copy provenance but does not establish that executable-mapping RSS
variation has disappeared. Warm-cache results have no exclusive-host or
physical-cold-cache guarantee.

## Matched results and decision

All 72 fresh processes complete and pass their workload oracles: 64 native
children contribute 3,200 measured iterations, and eight allocator children
contribute 120. There are no retries or replacement captures. The
[packet verification](results/change-0890/verification.json) reconstructs the
candidate source and replays the command, corpus, binary, and reduction evidence.

The before/after columns below are medians of child p50 values from this
matched capture. The ratio and interval summarize within-pair ratios, so the
ratio need not equal the quotient of the two descriptive medians.

| Scenario | Before p50, ms | Candidate p50, ms | Paired ratio | 95% bootstrap interval |
| --- | ---: | ---: | ---: | ---: |
| Producer XLSX full-cell scan (primary) | 3.5555065 | 3.5332740 | 0.99381 | 0.98861–1.01112 |
| Ordinary XLSX full-cell scan | 0.1328735 | 0.1319480 | 0.99422 | 0.99111–0.99907 |
| Selected DOCX paragraph | 0.1048630 | 0.1077825 | 1.02837 | 1.00434–1.04080 |
| Selected PPTX slide | 0.1784230 | 0.1711705 | 0.96175 | 0.94914–0.97305 |

| Scenario | Paired p95 ratio | Paired p99 ratio | Paired process-RSS ratio |
| --- | ---: | ---: | ---: |
| Producer XLSX | 0.99075 | 0.98999 | 0.93482 |
| Ordinary XLSX | 0.97113 | 0.94255 | 0.92363 |
| DOCX | 1.01325 | 1.01646 | 0.96993 |
| PPTX | 0.96292 | 0.96973 | 0.95898 |

The [reduction](results/change-0890/reduction.json) retains three control
metric flags covering five individual pairs: ordinary XLSX p99 ratios
1.08447 and 1.05557; DOCX RSS ratio 1.06323; and PPTX p99 ratios 1.05285
and 1.07553. No primary metric or native p50 spread flag fires. DOCX's
2.84% paired p50 increase remains visible even though no pair crosses the
5% review threshold. A PPTX block also has large p99 values on both sides
(3.257103 ms before, 3.222012 ms after); it remains included. No sample is
removed or attributed to a specific host event without evidence.

The allocator population has exactly 25,933 allocation calls and 37,618,900
allocated bytes per operation on both sides, with zero failed calls. Median
region peak live bytes change from 9,252,257 to 9,252,252, a five-byte
observation, not a meaningful general memory improvement. Whole-process RSS
is separate from these allocator counters; the managed-copy protocol does
not establish a causal explanation for RSS differences.

The [decision](results/change-0890/decision.json) rejects the candidate on the
failed primary gate. Improvements in other controls and lower median RSS
do not replace that gate. Both changed owner files, including the new tests,
are restored byte-for-byte to `f790c2228c`; the complete tested version
remains in [candidate-v2.patch](results/change-0890/candidate-v2.patch).
0885 remains the current-source release reference. This result rules out
adopting this bounded reuse implementation under the stated protocol; it
does not prove that all repeated-name work is unoptimizable.

## Reproduction and cleanup

[The packet README](results/change-0890/README.md) documents capture and
retained-only replay. The native and allocator binaries, source archive,
source patch, raw reports, logs, failures, and reviews remain retained.
Large binaries and the 0885 source archive are local ignored artifacts;
replay requires them in addition to the committed evidence.
Correctness scratch has been removed with a
[verified inventory receipt](results/change-0890/cleanup-correctness.json).
The release source, target, and managed executable copies have also been
removed. The initial [release cleanup receipt](results/change-0890/cleanup-release.json)
retains a post-removal verification failure: its verifier resolved a packet-relative
artifact path against the repository root. The deletion scope was correct;
the separate [corrected custody check](results/change-0890/cleanup-release-custody-verification.json)
verifies retained artifacts and absence of the removed roots without deleting
anything. The original helper and failure remain unchanged.

[Retained-only plan verification](results/change-0890/post-cleanup-freeze.json)
and [complete packet replay](results/change-0890/post-cleanup-verification.json)
pass after cleanup. The [independent capture review](results/change-0890/capture-review.json)
reconciles raw child statistics, pairing, intervals, and flags. No worktree was
created for 0890; unrelated worktrees and the untracked benchmark file are
outside its scope. The broader non-iWork performance goal remains open.
