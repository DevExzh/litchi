# Batch 0506 ADR review: ODG source-span batching

Review date: 2026-09-11.  Review revision: `56e396d05`.

The proposed change is a private optimization in
`crates/litchi-odg/src/package/snapshot.rs`: shape source-span discovery may
collect the 16 already-requested attributes in one checked traversal, using
fixed stack slots.  The slots correspond to `control`, `name`, `layer`,
`x`, `y`, `width`, `height`, `x1`, `y1`, `x2`, `y2`, `viewBox`, `points`,
`transform`, `d`, and `style-name`, in the same order as the existing span
fields.  Page discovery currently requests only `name` and `style-name` and
does not need to move to this path.

This review finds the design compatible with the accepted ADRs, subject to
the implementation and verification conditions below.  It is a design gate,
not evidence that the candidate has been implemented or tested.  No
The reviewer changed no production code, manifest, lockfile, or build output;
the source audit below records the candidate that the coder has since
materialized in the shared worktree.

The fast path must retain the current `attribute_source_span` semantics:

1. Iterate every raw attribute and map iterator errors before applying the
   local-name prefilter already introduced by batch 0505.
2. Resolve matching qualified names against the current namespace resolver;
   compare the resolved namespace and local name with the requested pair.
3. Preserve the exact raw qualified spelling (or directly record the exact
   value span) so aliases, namespace rebinding, and lexical prefix choices do
   not change the byte range selected for a later edit.
4. Detect duplicate matching expanded names, including two aliases bound to
   the same namespace, and retain the existing error.
5. On *any* discovery or raw value-span error, discard partial slots and call
   the existing helper in its original field order.  This fallback must cover
   malformed attributes, duplicate matches, malformed quoting or `=`, and
   missing raw spans.  The old helper therefore remains authoritative for
   error ordering and exact error behavior on exceptional input.
6. Convert the successful value offsets to `tag_start`-relative source spans
   without copying source XML or canonicalizing names.  The fixed slots must
   not become a user-controlled vector or alter source, budget, or publication
   state.

## Source audit of the materialized candidate

The current uncommitted source matches that shape: `ShapeAttributeSpans` is a
private 16-slot array, the shape path consumes all slots in the original field
order, and the batch loop parses every `element.attributes()` item before
storing only matching raw `QName` values.  Its wrapper drops the partial batch
and invokes the old ordered helpers when the batch returns an error.  The
source diff changes no manifest, dependency, public export, budget constant,
snapshot field, or edit/publication path.

The slots currently borrow `QName` values from the event for the duration of
the helper.  This is lifetime-safe and source-preserving, but it changes the
stack/register shape; the prior batch-0505 borrowed-key experiment showed an
RSS regression.  Batch 0506 therefore still needs its own RSS and latency
evidence before any performance result is accepted.  The locked quick-xml
implementation derives the resolver's local name from the same qualified name
bytes used by `local_name()`; retaining an explicit resolved-local comparison
would make that equivalence easier to audit but is not a new API or ownership
requirement.

## Post-capture disposition

The source review above preceded timing acceptance. The final comparison and
[root review decision](review-decision.json) retain the approximately 10%
plain-large RSS regression as an accepted tradeoff for the measured work and
latency reduction. The cause is unproven; the earlier borrowing experiment
does not establish a causal explanation. The final suite passes 108 tests,
including direct fallback parity and exact public edit/inverse coverage.

## Applicability matrix

| ADR | Applicability and review disposition |
|---:|---|
| 0001 | Applies directly to the private implementation boundary. No ordinary API, raw identifier, archive type, unsafe code, panic path, or unsupported-content behavior may be introduced. |
| 0002 | Applies to dependency direction. `litchi-odg` remains the owner of its ODG grammar and may not add a peer-family or umbrella dependency; the candidate has no dependency change. |
| 0003 | Applies to immutable snapshots and source-checked edits. Discovery is read-only parser state; it must leave snapshot, edit, commit, patch, and exact no-op semantics unchanged. |
| 0004 | Applies only as a non-regression boundary. No semantic type, selector, facade name, or public raw escape hatch changes; existing shape semantics remain authoritative. |
| 0005 | Primary applicability. The optimization removes repeated bounded attribute work with fixed storage, retains explicit limits and synchronous execution, and adds no provider, cache, thread, or budget bypass. Latency or allocation claims require representative measurements under the measurement contract. |
| 0006 | Primary correctness boundary. Preserve raw namespace spelling, lexical values, ordering, unknown attributes, and validation behavior. Every raw attribute must still be checked, and any fast-path error must use the ordered fallback described above. |
| 0007 | Applies to the ODF drawing model. The helper only indexes existing source attributes; it must not reinterpret shapes, activate content, or change unknown/inert drawing data. |
| 0008 | Applies to verification. The implementation requires focused alias, namespace-rebinding, duplicate, malformed-trailing-attribute, source-offset, formatting, diff, and boundary checks before acceptance; this review supplies none of those execution results. |
| 0009 | No direct change. ODF detection ownership and its fuzz boundary remain with the existing owner; the ODG parser optimization must not move detection logic. |
| 0010 | No direct change. Archive grammar remains below the facade; no ZIP implementation type or archive traversal is introduced. |
| 0011 | No direct change. This is outside OOXML physical-package ownership and cannot add an archive edge to an ODF family crate. |
| 0012 | No direct change. BIFF8 reference domains and panic-free encoding remain untouched. |
| 0013 | No direct change. PPTX notes ownership and atomic deletion remain untouched. |
| 0014 | Historical reader decision is amended by ADR 0015; no core-properties reader or ownership code is involved. |
| 0015 | No direct change. OOXML core-properties CRUD, preservation, and typed publication remain untouched. |
| 0016 | No direct change. BIFF8 writer-location validation remains untouched. |
| 0017 | No direct change. OOXML producer templates remain format-owned and untouched. |
| 0018 | No direct change. XLSX calculation-chain ownership and invalidation remain untouched. |
| 0019 | No direct change. DOCX web-settings ownership remains untouched. |
| 0020 | No direct change. PPTX table-style ownership and graph validation remain untouched. |
| 0021 | No direct change. DOCX glossary ownership and inert resource handling remain untouched. |
| 0022 | No direct change. PPTX embedded-font ownership, payload sharing, and activation boundary remain untouched. |
| 0023 | Primary ownership boundary. `litchi-odg` is the dedicated family owner; the private helper must stay there, use the existing common substrate only, and not recreate or export a shared family model. |
| 0024 | Primary current-topology check. The workspace package and dependency inventory do not change; ODG remains a dedicated family package and no topology claim follows from this optimization. |
| 0025 | No direct change. OGraph chart-area transactions are a separate owner and grammar. |
| 0026 | No direct change. Shared OLE directory metadata binding is unrelated. |
| 0027 | No direct change. XLS sheet-anchor ownership is unrelated. |
| 0028 | No direct change. The IWA migration-host exit, debt ledger, and iWork ownership are outside this ODG batch. |
| 0029 | No direct change. The archive-free IWA index and graph adapter remain outside this batch. |

The operative decision hierarchy is therefore: correctness and preservation
first, then the private bounded work reduction, with performance claims gated
by the measured evidence required by ADR 0005 and the verification gates in
ADR 0008. A separate ADR or exception is unnecessary if the implementation
keeps the boundaries above.

## ADR custody hashes

The mechanically generated [ADR hash manifest](adr-manifest.json) records all
30 input files, including the README. Root verified every SHA-256 against the
current file bytes before final evidence custody. The manifest is the single
hash inventory; this review does not duplicate its values.
