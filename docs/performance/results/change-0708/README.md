# 0708 XLSX validator name-storage experiment packet

This packet records a private candidate experiment for the source-backed XLSX
one-percent scalar-cell edit/save route. It is bound to revision
`553227504ed384cf6724c81913930823a85ffa09`, CPU 12, and the frozen source and
measurement constraints. The native/allocator pilot rejects the candidate,
production is restored, and the seven focused guard tests are retained. There
is no admitted performance claim. iWork is excluded.

The candidate tests a private `ElementName` stack representation in
`crates/litchi-xlsx/src/cell_values/validation.rs`: modeled names use static
bytes, empty-event names are borrowed for validation, and unfamiliar names
retain owned bytes. The candidate production patch and seven-test patch are
retained as evidence; no public API or benchmark harness source is part of the
candidate boundary.

## Current evidence

The focused baseline and candidate lanes each pass 41 tests: 30 validation,
five facts-oracle, and six planning-error tests. The existing owned reference
is unchanged. Baseline and candidate source manifests each contain 7,282
entries, with candidate deltas confined to `crates/litchi-xlsx/`. Focused test
results establish semantic coverage; their process durations are not a timing
comparison. The baseline and candidate builds and the native/allocator
acquisition are recorded in the packet; conditional profile/RSS acquisition
was deferred after the native gate failed. The restored focused lane also
passes all 41 tests, and restored formatting, all-features/all-targets XLSX
check, and `-D warnings` Clippy all pass.

The first restored Clippy attempt reported a pre-existing `useless_format` in
one retained shared-traversal guard. Root replaced it with the identical
literal and reran the full restored focused and quality lanes successfully.
`quality-correction.json`, `retained-test-correction.patch`,
`retained-test-corrections.json`, the interrupted preflight artifacts, and the
passing restored receipts preserve that correction history. It changed only a
retained test and did not alter production or measurement source retroactively.

The standalone layout observation is 24 bytes per candidate enum slot versus
16 bytes per baseline `Box<[u8]>` slot. At the unchanged maximum XML depth of
256, the logical slot difference is 2,048 bytes. This observation excludes
allocation capacity, live allocation, peak memory, and RSS, so it is retained
as a design cost rather than a memory result.

## Pilot result

The independent analysis passes its custody, arithmetic, parity, and flag
checks and records `decision: reject` for the native/allocator pilot. All
eight primary rows fail the joint gate: total p50 and mean each require 2%,
and planning p50 requires 5% in every shape/repeat.

| Pair | Shape / repeat | Total p50 | Total mean | Planning p50 | Row |
| --- | --- | ---: | ---: | ---: | --- |
| A1 → B1 | dense-sparse / 1 | 1.130% | 1.167% | 1.406% | fail |
| A1 → B1 | dense-sparse / 2 | 2.147% | 2.316% | 0.210% | fail |
| A1 → B1 | medium / 1 | 0.757% | 0.742% | 0.339% | fail |
| A1 → B1 | medium / 2 | 0.671% | 0.625% | −0.262% | fail |
| A2 → B2 | dense-sparse / 1 | 0.452% | 0.468% | 0.423% | fail |
| A2 → B2 | dense-sparse / 2 | 0.799% | 0.862% | 0.415% | fail |
| A2 → B2 | medium / 1 | 1.090% | 1.090% | 0.006% | fail |
| A2 → B2 | medium / 2 | 1.853% | 1.684% | −0.052% | fail |

The allocator gate passes all four shape/repeat rows. Planning allocation
calls fall 27.477% on medium (67,845 → 49,203) and 27.746% on dense-sparse
(129,411 → 93,505). Publication allocation calls remain unchanged at 19,197
and 36,573. Planning peak above region start grows by 64 bytes per shape, and
the logical stack-slot observation remains +2,048 bytes at depth 256. These
allocator reductions and memory observations cannot override the failed
native gate.

The packet retains 142 paired over-5% flags (84 adverse, 58 favorable), 38
same-build A/A flags, 21 repeat-drift flags (14 higher, 7 lower, −9.17% to
+10.87%), and 24 deduplicated allocator flags. All 39 native identity rows
match across output, corpus, source, and counters. Conditional Callgrind and
RSS acquisition is deferred because the native pilot failed. Production is
restored and the seven focused guard tests remain retained evidence.

## Frozen plan and gates

The primary native matrix uses medium and dense-sparse deterministic
four-worksheet workbooks, two repeats, and 200 samples after 20 warmups. The
ABBA order is baseline noise 1, baseline noise 2, baseline A1, candidate B1,
candidate B2, baseline A2, with reversed shape order on repeat two. Guard
cases cover one-edit, managed, vendor-extension, and noncompact source-backed
routes as recorded in `plan.json`.

Admission requires every primary shape/repeat to reach at least 2% total p50
reduction, 2% total mean reduction, and 5% planning p50 reduction. The
separate allocator lane uses two repeats, five samples, and no warmup per
shape, reports operation regions independently, and requires at least 20%
planning allocation-call reduction for every shape/repeat. Changes over 5% and
repeat drift are retained for review. If the native pilot passes, the
conditional profile requires at least 3% planning instruction reduction for
every shape/repeat, with RSS, memory, focused quality, and broader consumer
gates. A failed gate permits no benefit claim.

The native interval includes open, planning, staging plus commit, and
sequential publication. It excludes sink setup, remaining handle destruction,
reopen, and semantic or preservation oracles. Allocator values are
operation-region diagnostics and do not sum phase peaks. Conditional
Callgrind values are guest-instruction attribution, and conditional RSS is a
whole-child process signal; neither is substituted for native latency or
operation-level memory evidence.

## Packet map

| Evidence | Artifact |
| --- | --- |
| Frozen main plan and constraints | [`plan.json`](plan.json), [`constraints.json`](constraints.json), [`investigation.json`](investigation.json) |
| Candidate design and independent review | [`candidate-design.json`](candidate-design.json), [`review.json`](review.json), [`candidate-production.patch`](candidate-production.patch), [`candidate-tests.patch`](candidate-tests.patch) |
| Build and source custody | [`build.py`](build.py), [`build-baseline.json`](build-baseline.json), [`build-candidate.json`](build-candidate.json), [`source-baseline.json`](source-baseline.json), [`source-candidate.json`](source-candidate.json), and the build logs |
| Focused and restored correctness receipts | [`quality-baseline-focused.json`](quality-baseline-focused.json), [`quality-candidate-focused.json`](quality-candidate-focused.json), [`quality-restored-focused.json`](quality-restored-focused.json), their source manifests, and the focused logs |
| Restored quality and correction custody | [`quality-restored-quality.json`](quality-restored-quality.json), [`quality-correction.json`](quality-correction.json), [`retained-test-correction.patch`](retained-test-correction.patch), [`retained-test-corrections.json`](retained-test-corrections.json), and the restored quality logs |
| Layout and memory design observation | [`layout.json`](layout.json), [`layout.py`](layout.py), [`layout-probe.rs`](layout-probe.rs), and layout build logs |
| Native/allocator capture and analysis | [`capture.py`](capture.py), [`analyze.py`](analyze.py), [`analysis.json`](analysis.json), with native and allocator child receipts and raw outputs |
| Conditional profile/RSS mechanism | [`mechanism-plan.json`](mechanism-plan.json), [`mechanism.py`](mechanism.py), [`analyze_mechanism.py`](analyze_mechanism.py) |

The raw child artifacts are source- and binary-bound. The capture scripts
record source before/after custody, binary identity, plan/script/constraint
hashes, corpus and output identities, and per-child receipt data. A fresh
acquisition must honor the frozen plan and preserve every raw result; no
derived summary can replace those receipts.

## Correctness and memory boundary

The candidate must preserve exact close-name matching, strict/transitional
dialect behavior, namespace and unbound-prefix checks, first-error order,
copied-subtree opacity, source limits, depth and reservation behavior,
authoritative fallback, and typed atomic refusals. The review notes that a
copied close mismatch is rejected by `NsReader` before the validator's `End`
callback and that the new differential reference does not model production
depth/`try_reserve`; the existing 300-level exact refusal remains part of the
required full XLSX quality gate. The production maximum XML depth remains 256.

The layout result is only a logical stack-slot comparison: 24 bytes for the
candidate enum and 16 bytes for the baseline box, or 2,048 additional logical
bytes at depth 256. The allocator lane records a 64-byte increase in planning
peak above region start for each primary shape. Process RSS was not measured
because the native pilot failed.

## Reproduction limits

The packet contains the native and allocator result, numerical pilot
comparison, regression flags, custody proof, and rejection disposition. It has
no Callgrind or RSS result because the native gate failed. The focused tests
and build/source receipts do not prove end-to-end performance beyond this
scoped pilot. The workload is synthetic and in-memory, and the packet makes
no claim about native Office producers, cold storage, cross-platform behavior,
broad CRUD coverage, parallel scaling, or universal documents. The broader
non-iWork goal remains active.
