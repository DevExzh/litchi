# 0495: managed DOCX document edits and owner reservations

0495 is the necessary measured enabler for the ordinary opened-document DOCX
lifecycle. The current working tree adds an owner-carrying managed
`Snapshot`/`Edit`/`Commit`/`Patch` path, source-authorized forward publication,
and a complete-artifact inverse capability. It also freezes a before/after
harness so the managed path can be measured against the existing unmanaged
control.

The formal and pilot captures are now verified and accepted as descriptive
provider evidence. `claim_authorized` remains false: the paired comparison is
limited to unmanaged before/after rows, while managed-after rows are a
capability and finite-budget baseline with no managed-before comparator. The
full performance goal remains open.

## Current implementation

For an eligible ordinary source-backed DOCX, `Package::document_snapshot()`
and `Package::edit_document()` now admit the main document through the finite
`ExecutionContext`. The managed snapshot keeps the source lineage, source
version, and main Part URI with its source XML or managed `PartData`; a managed
payload cannot escape through the old shared `Arc<Vec<u8>>` route. Signed or
otherwise unsupported changed inputs remain typed refusals. Exact reads and
no-op handling may use the managed `PartData` fallback where source policy
requires it, but that fallback does not authorize a changed signed document.

The ordinary managed editing seam now also has a base-relative reconstruction
plan for repeated direct-body paragraph replacements in one isolated edit. Each
replacement validates the authored text, scans the exact source text slots,
proves the ranges against the immutable base `SourceXmlPart`, rebuilds the
candidate from that base, and performs semantic readback before retaining the
operation ledger. This avoids asking a derived one-shot `SourceXmlPart` for a
second proof and supports disjoint paragraphs in caller-selected order. The
dedicated `replace_body_paragraph_texts` batch helper and specialized hyperlink,
field, revision, content-control, table, transfer, and other structural
mutators still use their typed managed refusal until they have the corresponding
source-preserving ownership and publication proof. The base-relative plan is
implemented and covered by the focused gates; the frozen 0495 benchmark still
measures one paragraph only, so no multi-paragraph timing claim follows.

The managed reservation ledger covers the retained objects that the current
path creates:

* `ManagedAdmission` charges snapshot/index metadata and object reservations
  before layout growth. Its transient scan token charges parser memory,
  objects, depth, and work before scanning begins.
* An empty managed commit uses an allocation-free `OperationList::Empty`.
  Non-empty operation ledgers admit the temporary peak while a `Vec<Operation>`
  is converted to `Arc<[Operation]>`, including the analogous Vec-to-Arc
  overlap while the inverse ledger is built. The admission is made before
  either retained representation grows.
* Immutable snapshots, source-backed paragraph views, projected candidates,
  commits, and patches retain shared admission owners. Dropping one view does
  not invalidate a source range; dropping the last owner releases the RAII
  reservation.
* The staged edit admits operation metadata and caller-authored text storage.
  Commit construction admits the inverse operation ledger and its text
  storage, so `Patch::inverse()` can retain the same source owners and managed
  reservations.
* `Snapshot::shared_xml()` remains unavailable for managed/source-backed
  storage. This prevents a caller from retaining an uncharged detached XML
  allocation. Owner-carrying views or typed refusal remain the rule for
  semantic graphs that have not yet received this treatment.
* The managed read-only `Document` path retains a paragraph index with
  `DocumentIndexAdmission` memory/object reservations, including its
  construction overlap. Each parser-backed text query admits its temporary
  memory, objects, depth, and work through a query guard. Collection-returning
  managed views still refuse when they would expose an unbudgeted Arc-backed
  graph; caller-owned text queries remain available.

The source-backed unmanaged main-document path now stores an
`OwnedWithIdentity` snapshot. That scoped identity is retained through rewritten
candidates, commit construction, `Patch::inverse()`, and durable replay/restore,
so a separately opened package with equal main-document XML but different media
or opaque members is refused before output. This keeps the compatibility path
unmanaged while preserving the lineage, source version, and Part URI needed for
the exact-source check.

The current parser work envelope is deliberately conservative. A query guard
may pre-consume the square of (`xml_len` + 1) in `Resource::Work`, so a large
paragraph or document can be refused under a finite work limit even when its
actual parse would fit. This is a safety ceiling, not a measured operation
count or a CPU instruction claim. Future work should integrate observed
per-event accounting without weakening the finite-work boundary.

Managed read guards check execution state around parser-backed reads. Read-only
`collapsed` and run-symbol access uses those guards and parser admission, while
managed `set_collapsed` and symbol mutations refuse before detaching
source-owned XML. Direct paragraph inlines, hyperlinks, nested runs, and
unknown run content now retain owner-carrying `XmlRef` ranges; the remaining
unwired collection graphs stay typed-refused or bounded separately.

Inherited namespace resolution now scans the retained owner through the selected
start event, lets the reader pop a preceding sibling before that event, and
copies the effective bindings into a resolver lease. The admission reserves
`32 * owner_span_len` plus fixed state, reader, and admission metadata before
the scan; its Work reservation includes the
`(fragment_len + 1) * (owner_span_len + 1)` fragment-by-owner cross-term, with
per-event work charged against active bindings. A namespace query currently
rescans from the owner prefix to reconstruct inherited bindings. That is a
correctness path, and no selective namespace-query performance claim is made.

These reservations establish the ownership seam needed for a bounded managed
opened-edit measurement. They do not establish that every allocation in the
broader DOCX object model, durable history, or composition graph is admitted;
those gaps remain listed below.

## Source-authorized publication and inverse

`publish_document_commit_to_stream()` applies the patch to the current exact
source snapshot. A changed managed candidate is published through
`SourceTopologyPlan` and its `SourceXmlPart` proof, preserving untouched ZIP
members and rejecting stale source versions, unsupported topology, signatures,
dependency closures, and sink or resource failures with typed errors. An exact
no-op uses the complete source artifact path.

Changed ordinary document publication also runs a bounded exact-source policy
scan over an existing `settings.xml` Part before output. The policy permits
only a foreign-bound `mc:Ignorable` attribute: every listed prefix must resolve
to a bound namespace outside the Word main and MCE namespaces. The settings
Part is then copied byte-for-byte; the main-document edit does not rewrite its
opaque or policy bytes. `mc:ProcessContent`, `mc:AlternateContent`, unresolved
or non-foreign ignorable prefixes, namespace declarations with escaped URI
text, and direct enabled `w:documentProtection`, `w:writeProtection`, or
`w:trackRevisions` policy are refused before sink output. Explicitly disabled
protection or revision markers remain admissible under this source policy.

`publish_document_commit_with_inverse_to_stream()` adds the stronger artifact
capability. `DocumentPublication` retains the original `SourceArtifact`, the
original artifact fingerprint and length, and the accepted forward artifact
fingerprint. `publish_document_inverse_to_stream()` authenticates the complete
current artifact before restoring the exact original package, including opaque
members and ZIP framing. A separately opened foreign package with equal main
document XML but different opaque members is refused. The focused managed
tests cover this capability, source mutation, cancellation, output limits, and
partial sink output; this API is an implementation enabler and is not a 0495
timed row.

`Patch::inverse().apply()` has a narrower meaning. It reverses the in-memory
document XML snapshot and checks the patch's exact document source identity; it
does not restore a complete DOCX artifact. The frozen harness field named
`inverse_restores_source_verified` is this XML-only preflight oracle, and its
scope is `untimed_preflight_commit_patch_oracles`. It must not be reported as
complete-artifact inverse evidence.

The managed `AuthoredXmlFragment::markup_with_execution_context()` and
`text_with_execution_context()` constructors charge the temporary audit
wrapper and parser workspace for memory, objects, depth, and work before the
audit runs. Their owned input `Vec` becomes the retained fragment payload; the
surrounding caller operation admits that retained payload separately. The
scoped audit reservation is released after the audit, including on cancellation
or failure.

## Harness and measurement boundary

The new `docx-managed-edit` harness uses the same public lifecycle for its two
API modes:

* `unmanaged-api` is the paired before/after regression control using the
  compatibility constructor without an `ExecutionContext`.
* `managed-api` is the candidate-after capability path with a finite explicit
  `ExecutionContext` and managed budget gauges.

The old managed typed refusal is retained as a negative capability probe and
historical baseline. A refusal from that probe is not a successful timing row.
The harness keeps the unmanaged before/after comparison separate from the
managed-after capability measurement; it does not compare different source
revisions or API modes as an optimization claim.

The frozen provider matrix has six explicit arms over the deterministic 0494
media-rich corpus: `owned`, `instrumented`, `file-warm`, 4 KiB `short`, delayed
64 KiB `delayed`, and the zero-fixed-delay `range-zero` control. Normal and
allocator roles use the same operation and source configuration. Setup,
provider/context construction, sink reservation, and fixture preparation stay
outside the operation clock. The clock covers fresh package open, one paragraph
edit and commit, sequential publication, the specified drops, and timed commit
diagnostics. Output, reopen, media, source-version, range, patch, stale/foreign
source, budget-release, and cache checks are outside the clock. The harness does
not exercise the complete-artifact inverse API; that capability is covered only
by the focused production tests described above.

The retained `change-0495` directory contains the frozen `protocol.json`,
formal and pilot captures/analyses, verification receipts, source and harness
manifests, design reviews, and final gate receipts. The formal run retains 72
processes and 2,160 samples; pilot2 retains 36 processes and 108 samples. The
six-provider matrix covers `owned`, `instrumented`, `file-warm`, 4 KiB `short`,
delayed 64 KiB `delayed`, and zero-fixed-delay `range-zero`, in normal and
allocator roles. Every formal row produced the expected 16,793,048-byte
output and passed source, sink, semantic, untouched-media, and resource
checks.

The result remains descriptive evidence. The paired review is unmanaged
before/after only; managed-after rows establish successful finite-context
capability, budget transitions, and release behavior without a managed-before
speed or RSS baseline. Eight whole-child RSS observations and three latency
comparison cells cross the five-percent review threshold. They remain
follow-up flags with unresolved causality, and repeat instability remains
visible in the formal review. Phase-local CPU and whole-child RSS attribution,
plus repeats of the unstable file-warm and short/allocator arms, remain the
explicit measured follow-up. A post-publication cache observation is also
needed for any cache-retention claim.

## Current verification state

The final production-gate receipts pass. Final 4 gate 3 reports 478 harness
tests passed and 1 ignored; gate 9 reports 937 DOCX unit tests; gate 10 reports
119 integration tests, grouped by the receipt as `16 + 5 + 36 + 12 + 6 + 25 +
8 + 11`; and gate 15 reports 388 OPC unit tests. The final DOCX doctest receipt
reports 74 passed and 31 ignored, the final OPC doctest receipt reports 5
passed, and final 6 helper tests report 58 Python checks. Final 4 gates 1, 2,
4, 5, 11, 12, 13, 14, 16, and 17 also pass.

A supplemental AddressSanitizer/libFuzzer smoke ran the existing DOCX
`parse_docx`, `tail_append`, and `source_backed_tail_append_stream` targets and
OPC `parse_opc`, with 1,000 executions per target and exit code zero for all
four. The [supplemental fuzz evidence](../results/change-0495-fuzz/) is separate
from the sealed performance bundle. This bounded smoke exercises the existing
fuzz entry points; the managed-edit API itself is covered by the focused
production tests above, not by a new managed-edit fuzz target.

The sealed bundle covers 1,486 files with seal SHA-256
`bed878c517ba2aeac8dca25bcb78b19d82a259811abebfe6fa1c8cfd3b024a13`. Its
cleanup receipt removed 5.85 GiB of temporary custody data and retained only
the four replay binaries needed for local reproduction.

Candidate 23 is the final 27-Rust-path source scope; its only test-only addition
is the precise OPC namespace refusal. These are correctness, build, and source
scope receipts only. The 63-fixture corpus parity gate remains custody evidence.
The formal/profile result is descriptive: no provider ranking, managed-before
comparison, broad speedup, or claim classification is promoted.

The normal managed p50 observations for the six providers were:

| Provider | Repeat 1 (ms) | Repeat 2 (ms) |
| --- | ---: | ---: |
| owned | 5.592 | 5.556 |
| instrumented | 3.277 | 3.298 |
| file-warm | 5.682 | 5.853 |
| short | 4.066 | 4.060 |
| delayed | 572.748 | 576.147 |
| range-zero | 182.327 | 181.990 |

These are managed capability/budget observations, not a managed-before speed
comparison or provider ranking. In the allocator role, unmanaged before-to-after
allocation calls changed from 22,859 to 9,696 and allocated bytes from
5,721,334 to 1,623,696, while the peak live increment changed by 0.485%.
Managed-after versus unmanaged-after calls changed by 111.788%, allocated bytes
by 162.992%, and peak live increment by 2.058%; this is a descriptive API delta
with no managed-before comparator.

The managed resource snapshots have distinct phases: `cache_before` is
post-open, `cache_live` is post-edit/pre-publication, and resource `live` is
post-publication. Resource `after_drop` follows package consumption and
returned-snapshot/commit release. There is no post-publication cache snapshot,
so the run does not establish cache-after-drop behavior.

**Correction (change [0629](../0629-facade-docx-budget-test-bisect.md)).**
The coordinator should append the following to
`docs/performance/changes/0495-docx-managed-document-edits.md`, after
*Current verification state*:

## Correction (change 0629)

`44a4710699ef17041d5969240c30984dffbc3319` left one test outside this change's
gate list red, and it stayed red for 94 commits. The facade test
`document::doc::tests::managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal`
in `crates/litchi/src/document/doc.rs` sizes a memory budget to the package it
opens (1,154 bytes) and calls `Document::paragraph_text` on a managed
source-backed DOCX. The parser workspace this change introduced —
`xml.len() * 32 + 131_072` bytes of `Resource::Memory`, reserved by
`ensure_source_document_xml` before quick-xml's namespace reader allocates,
and by `admit_document_query_parser` for every parser-backed text query —
needs 137,280 bytes for that 194-byte document, so the query refuses with
`Memory budget exceeded in facade-managed-docx-paragraph-text: observed
137474, limit 1154`. The same commit sized `litchi-docx`'s own managed tests
from this formula (`source_document_scan_workspace` in
`crates/litchi-docx/tests/source_backed_managed.rs`); the facade test was not
migrated with them, because `litchi`'s default feature set is empty and
`cargo test -p litchi` never compiles it.

**No behaviour of this change is in question.** The contract the test
asserts — caller-owned text queries remain available on a managed document
while collection-returning views are refused — is the contract this record
states, and it holds under any budget that admits the workspace. What was
stale is the budget the test handed the facade. Change
[0629](../0629-facade-docx-budget-test-bisect.md) bisected the failure to this
commit (last good `de8ee88b0`), corrected the test to assert the contract
rather than the pre-fence number, and left the fence itself untouched. Change
0621 recorded the failure as pre-existing at `1e4198321`, which was correct;
changes 0588, 0591, 0592, 0593 and 0594 are not implicated.

## Scenario classification

The scope is kept distinct so a later report cannot combine unlike workflows:

| Scenario | 0495 status |
| --- | --- |
| Synthetic managed read-only access | Existing capability/read paths only; no 0495 edit timing claim. |
| Opened managed document edit | Verified six-provider capability and finite-budget baseline across 24 managed formal processes and 720 samples within the 72-process/2,160-sample full matrix; no managed-before speed, RSS, allocation, or optimization comparison. |
| Unmanaged opened edit | Paired before/after review for the same harness and corpus; eight whole-child RSS and three latency cells cross the five-percent flag threshold, with unresolved causality and repeat instability retained. |
| Native producers | Excluded; no producer result is implied by the managed API. |
| Cold filesystem source | Excluded from this six-arm harness capture; 0494 cold evidence remains historical. |
| Concurrent scaling | Excluded; no worker or multi-core scaling result is implied. |

The implementation leaves the 0494 provider/read-ahead and cold baseline
records as historical evidence. 0495 does not add a cold-filesystem,
independent-producer, borrowed-source, atomic-save, concurrent-scaling, or
broader CRUD result. The full performance goal remains open.

## Remaining correctness and goal boundaries

The following remain genuine open work rather than hidden success criteria:

* Managed table, row/cell, notes/comments, smart-tag, and other collection
  graphs must either retain the source owner with budgeted collection storage
  or return a typed refusal. Direct paragraph inlines, hyperlinks, nested runs,
  and opaque run content now use retained `XmlRef` ranges, while their broader
  collection and mutation surfaces still require separate review.
* Durable patch serialization, bounded history/redo, prepared-edit composition,
  conflict graphs, and eviction/drop must retain owner reservations based on
  retained bytes. The current managed history path still refuses where it has
  no reservation ledger. `Patch::inverse()`'s XML reversal does not close
  this durable-history requirement.
* The dedicated managed paragraph-batch helper, non-paragraph mutators,
  cross-package transfers, relationship/resource dependency publication, and
  changed signed, encrypted, protected, MCE, or ambiguous inputs remain
  fail-closed until their exact source and budget contracts are implemented.
* The source-backed unmanaged main-document edit/commit path now retains scoped
  source identity through rewrite, commit, patch inverse, and durable restore,
  and refuses the equal-main-XML/different-media foreign package before output.
  That result is scoped to this main-document path; it does not establish
  equivalent identity coverage across every DOCX facade or complete-artifact
  scenario.
* Genuinely borrowed source workflows, atomic filesystem replacement, native
  producers, cold storage behavior, concurrent scaling, and the remaining
  Office CRUD scenarios are outside this slice and remain open in
  `docs/GOAL.md`.

iWork and `/home/zhuhe/code/litchi-spec-gaps` are excluded. The protected
worktree was not inspected or modified.

## Evidence and references

The implementation and measurement boundary is documented in the retained
[0495 README](../results/change-0495/README.md), [harness plan](../results/change-0495/harness-plan.md),
[results review](../results/change-0495/results-review.md),
[profile review](../results/change-0495/profile-review.md),
[ownership review](../results/change-0495/ownership-review.md), and
[publication review](../results/change-0495/publication-review.md). The latter
reviews distinguish managed forward publication, document-XML patch reversal,
and complete-artifact inverse authority; the results/profile reviews preserve
the measurement boundary and unresolved flags.

The current ADR refresh receipt covers all 30 recorded ADR files; their SHA-256
values still match `docs/performance/results/change-0495/adr-refresh.json`.
The final ADR compliance matrix maps ADR 0001/0002/0003/0005/0006/0008/0011/
0024 to owner-retained snapshots, finite admission, borrowed parser lifetimes,
source-authority publication and inverse, typed refusals, and the final custody
gates. The matrix records the remaining durable-history, composition, broad
mutator, and scaling boundaries. This change record was checked with
`git diff --check`. No Cargo command or workload capture was run while writing
this document.
