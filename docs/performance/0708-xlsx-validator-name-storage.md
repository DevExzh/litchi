# 0708 — XLSX validator name-storage candidate

`performance_claim: none`; the native/allocator pilot rejects the candidate
and production is restored. The reductions below are scoped pilot results,
not an admitted performance claim.

The 0707 source-backed XLSX planning profile identified the validator as a
substantial nested owner inside worksheet validation and parsing. This batch
tests a bounded private representation change in the validator's open-element
stack. The experiment is bound to revision `553227504ed384cf6724c81913930823a85ffa09`,
CPU 12, and the source-backed XLSX one-percent scalar-cell edit/save route.
The native and allocator capture is complete; the conditional profile and RSS
lanes are deferred after the failed native gate.
iWork is outside this batch.

## Pilot result and disposition

The independent packet analysis passes its arithmetic, custody, parity, and
flag checks and records `decision: reject` for the native and allocator pilot.
All eight primary shape/repeat rows fail the joint admission requirement. The
candidate's best total p50 reduction is 2.147% (dense-sparse, repeat 2), but
that row's planning p50 reduction is only 0.210%, below the required 5%.

| Pair | Shape / repeat | Total p50 | Total mean | Planning p50 | Joint row |
| --- | --- | ---: | ---: | ---: | --- |
| A1 → B1 | dense-sparse / 1 | 1.130% | 1.167% | 1.406% | fail |
| A1 → B1 | dense-sparse / 2 | 2.147% | 2.316% | 0.210% | fail |
| A1 → B1 | medium / 1 | 0.757% | 0.742% | 0.339% | fail |
| A1 → B1 | medium / 2 | 0.671% | 0.625% | −0.262% | fail |
| A2 → B2 | dense-sparse / 1 | 0.452% | 0.468% | 0.423% | fail |
| A2 → B2 | dense-sparse / 2 | 0.799% | 0.862% | 0.415% | fail |
| A2 → B2 | medium / 1 | 1.090% | 1.090% | 0.006% | fail |
| A2 → B2 | medium / 2 | 1.853% | 1.684% | −0.052% | fail |

The separate allocator gate passes every shape and repeat. Planning allocation
calls fall from 67,845 to 49,203 on medium (27.477%) and from 129,411 to
93,505 on dense-sparse (27.746%). Commit-core calls also fall from 42,383 to
23,712 (44.053%) and from 80,110 to 44,175 (44.857%), while publication calls
are unchanged at 19,197 and 36,573. These allocation reductions do not
override the failed native rows. Planning peak above region start increases
by 64 bytes for each shape; the candidate's logical stack-slot cost remains
8 bytes per active element and 2,048 bytes at depth 256.

The complete native matrix has 142 paired over-5% flags: 84 adverse and 58
favorable. The packet also records 38 same-build A/A flags, with a maximum
absolute drift of 34.21%, and 21 repeat-drift flags: 14 higher and 7 lower,
from −9.17% to +10.87%. The allocator lane contributes 24 deduplicated
over-5% flags in its planning and commit-region count/deallocation metrics.
These are review diagnostics, not additional performance claims.

Because the native pilot fails, the conditional Callgrind and RSS lanes are
deferred. The candidate production source is restored exactly, and the seven
focused guard tests are retained with the evidence packet. No production
optimization is admitted.

## Problem and candidate

`Validator` must retain each open element's local name because the XML reader
advances its event storage before the matching close event. The baseline keeps
`Vec<Box<[u8]>>`, copying every start-event local name into an owned boxed
slice. An empty element is validated immediately but still copies its local
name into a temporary vector. Ordinary worksheet and workbook paths use a
small fixed vocabulary, so the candidate tests whether those copies can be
avoided without changing validation or preservation behavior.

The private candidate changes only
[`validation.rs`](../../crates/litchi-xlsx/src/cell_values/validation.rs):

* `elements` becomes `Vec<ElementName>`.
* `ElementName::Static(&'static [u8])` represents the 13 modeled local names:
  `worksheet`, `dimension`, `sheetData`, `row`, `c`, `f`, `v`, `is`, `t`, `r`,
  `workbook`, `sheets`, and `sheet`.
* `ElementName::Owned(Box<[u8]>)` preserves an unfamiliar local name in the
  same owned form used by the baseline.
* Start events borrow the reader's local-name bytes while validation runs and
  retain the static or owned representation only after the existing admission,
  depth, and `try_reserve` checks succeed.
* Empty events pass their borrowed local name to validation and retain no name.
  Close validation reads the retained value through `AsRef<[u8]>`.

The candidate does not add a cache, global state, unsafe code, parser fusion,
public API, or proof reuse. Its hypothesis is that static modeled names and
borrowed empty-event names reduce avoidable name-copy work in the planning
validator while unknown and copied content remains owned and exact.

## Correctness boundary

The experiment must preserve the existing validation order and all established
typed outcomes. In particular, the candidate is required to retain:

* exact local-name and close-name byte comparison, including unknown names;
* strict and transitional dialect handling, namespace resolution, and
  unbound-prefix refusals;
* first-error precedence and the existing error text and owner;
* copied or opaque subtree behavior, including modeled-looking names under a
  foreign namespace and namespace rebinding;
* the existing XML depth limit, stack reservation and allocation-failure path;
* source admission, authoritative fallback, budgets, cancellation, and atomic
  refusal behavior.

The independent review records no remaining concrete defect after the
`expected.as_ref()` correction. It also records two limits of the differential
coverage: a copied close mismatch is rejected by `NsReader` before the
validator's `End` callback, and the new reference helper does not reproduce
the production depth or `try_reserve` path. The existing 300-level exact
refusal in the XLSX admission tests remains required. The candidate keeps the
production XML maximum depth at 256.

The candidate test patch adds seven tests while leaving the frozen owned
reference unchanged: five validator borrow/differential tests and two shared
traversal fallback tests. They cover modeled names in start and empty forms,
both dialects and namespace forms, composed-context refusals, opaque
modeled-looking names, namespace rebinding, close-mismatch transparency, and
reader-versus-validator first-error order.

The baseline and candidate focused lanes each pass 41 tests: 30 validation
tests, five facts-oracle tests, and six planning-error tests. Their logs and
source manifests are retained in the packet. These are correctness receipts;
their process durations are affected by compilation and are not a performance
comparison.

## Frozen measurement protocol

The primary native case is
`xlsx_source_backed_cell_values_one_percent_edit_save` over deterministic
medium and dense-sparse four-worksheet inputs. Each of two repeats retains 200
samples after 20 warmups. The order is baseline noise 1, baseline noise 2,
baseline A1, candidate B1, candidate B2, baseline A2; the second repeat
reverses the shape order. Native timing covers the harness's open, planning,
staging plus commit, and sequential publication regions. It excludes sink
setup, remaining handle destruction, reopen, and semantic or preservation
oracles.

The primary comparison requires every matched shape and repeat to meet all of
the following: total p50 reduction of at least 2%, total mean reduction of at
least 2%, and planning p50 reduction of at least 5%. The separate allocator
lane retains two repeats, five samples, and no warmup for medium and
dense-sparse shapes. It reports planning, staging, commit-core, aggregate
commit, and publication regions separately; region peaks are never summed. It
requires planning allocation-call reduction of at least 20% for every shape
and repeat. Any native or allocator metric with an absolute change over 5% is
flagged for review, as are repeat-drift diagnostics.

Guard coverage expands the frozen plan with one-edit medium and dense-sparse
cases, managed one-percent medium and dense-sparse cases, vendor-extension and
noncompact one-percent cases, and managed noncompact cases. A candidate that
passes only a primary subset cannot support a benefit claim.

Callgrind and RSS are conditional follow-up evidence. They run only if the
native pilot satisfies its primary gates. The profile lane uses two repeats,
one sample, no warmup, and attributes the outer
`litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets` owner;
the required planning-instruction gate is at least 3% reduction for every
shape and repeat. The profile is a guest-instruction diagnostic, not a latency
or allocator claim. The RSS lane uses two repeats, three samples, and two
warmups with `/usr/bin/time -v`; it is a whole-child process signal including
setup and oracle work, with a 5% review threshold rather than an
operation-level peak claim.

Every child is bound to the frozen plan, source census, build record, binary
hash, corpus, source identity, sink identity, and output identity. The packet
retains baseline and candidate source maps with 7,282 entries. Candidate
source changes are restricted to `crates/litchi-xlsx/`. Baseline, candidate,
allocator, profile, and RSS lanes remain separate evidence types.

## Layout and memory limits

The standalone layout probe reports a 24-byte logical `ElementName` slot for
the static/owned enum versus 16 bytes for the baseline `Box<[u8]>` slot. At the
unchanged maximum XML depth of 256, that is a calculated 2,048-byte increase
in stack-slot storage in the maximally deep case. This is a logical slot-size
calculation only: it does not measure allocation capacity, live bytes, peak
memory, or RSS. Static modeled names avoid their per-name heap allocation;
unknown names still use owned storage. The allocator lane observes a 64-byte
increase in planning peak above region start for each primary shape. This is
separate from the logical layout calculation and from process-level RSS,
which was deferred after the native gate failed.

## Status and limits

The baseline, candidate, and restored focused lanes each pass 41 tests: 30
validation, five facts-oracle, and six planning-error tests. The restored
checkout also passes `cargo fmt --all --check`, the locked all-features,
all-targets XLSX check, and the corresponding Clippy command with
`-D warnings`. The first restored Clippy attempt exposed a pre-existing
`useless_format` in a retained shared-traversal guard; root replaced that
format call with the identical string literal, then reran the complete
restored focused and quality lanes successfully. The exact correction patch,
before/after hashes, interrupted preflight evidence, and passing receipts are
retained in the packet. The correction changed only a retained test and did
not modify production or measurement source retroactively.

The production checkout is restored with the seven new guard tests retained.
The independent analysis verifies all 39 native output/corpus/source/counter
identity rows, stable constraints and source custody, recomputed statistics,
allocator phase reconciliation, and non-summed region peaks.

The pilot's native and allocator results are complete, but no Callgrind or RSS
result is claimed because the native gate failed. The reductions in the pilot
table are scoped to the synthetic in-memory route and its named guards; they
do not establish a retained speedup.

This experiment covers one synthetic in-memory XLSX source-backed edit/save
route and its named guards. It does not establish native Office-producer,
cold-storage, cross-platform, broad CRUD, parallel-scaling, or universal
workload behavior. The profile and layout observations cannot substitute for
the native and allocator gates. The broader non-iWork performance goal remains
active.

The complete custody packet is in
[`results/change-0708/`](results/change-0708/README.md).
