# OPC prepared topology and effective candidate readback

This is the design and evidence contract for source-preserving OPC topology
publication. The prepared-topology API and its borrowed candidate callback are
implemented in the current worktree and remain under OPC/DOCX review. This
note does not grant production approval or make a performance claim; the
required evidence below remains the acceptance checklist.

The immediate consumer is the source-backed DOCX SVG attach/detach lifecycle.
The same boundary is intended for later source-preserving owners that need to
validate a changed relationship, content-type, and Part closure before writing
to an external sequential sink.

## Problem and boundary

[`SourceTopologyPlan`](../../../crates/litchi-opc/src/source_backed.rs) is
currently opaque and consumed by
`SourceBackedPackage::write_topology_to_stream`. That writer owns the source
ZIP preservation plan, relationship dialect, content-type lexical preservation,
source freshness, signature/encryption policy, and final sink accounting. A
format owner needs one additional stage between planning and publication:

1. prepare the effective logical topology from the source and the plan;
2. read the prepared topology through typed OPC views and perform complete
   owner-specific semantic readback; and
3. publish the already prepared source-preserving topology only after that
   readback succeeds.

The candidate stage must not write to the caller's sink. A failed plan,
readback, source check, cancellation check, or resource reservation therefore
leaves the external sink untouched. An exact semantic no-op keeps the existing
fast path: it copies the source artifact exactly and does not prepare or reopen
a candidate.

This design treats an effective typed candidate as a semantic view over the
source and prepared overlays. It is equivalent to a complete candidate reopen
for semantic purposes only when the view exposes every effective Part,
relationship, content-type binding, and payload that the owner can observe.
It does not by itself validate ZIP framing or prove that a later physical
writer emitted the intended bytes; those are separate publication checks below.

## OPC seam

The smallest ownership-preserving API is a one-shot prepared topology:

```text
SourceBackedPackage::prepare_topology(
    self,
    SourceTopologyPlan,
) -> Result<PreparedTopology>

PreparedTopology::with_candidate(|candidate: &EffectiveTopology| {
    owner_readback(candidate)
}) -> Result<T>

PreparedTopology::publish_to_stream(writer) -> Result<()>

SourceBackedPackage::with_prepared_topology(
    &self,
    SourceTopologyPlan,
    |candidate: &EffectiveTopology| owner_readback(candidate),
) -> Result<T>
```

The owning and borrowed forms share one preparation implementation. The
borrowed callback lets an edit validate its complete candidate before returning
a commit while leaving the source package usable. It retains overlays and
reservations through the callback, checks destination and transferred source
freshness before and after it, and writes no external bytes. Candidate views
cannot escape the callback. Publication prepares again against its consuming
source boundary. Both forms refuse candidate callbacks for an exact empty plan;
the owning publication path still copies that source exactly.

The important properties are:

* `prepare_topology` consumes the source package and plan once. It performs the
  existing topology preflight, including operation limits, canonical versus
  noncanonical relationship restrictions, duplicate names and relationship IDs,
  internal target checks, content-type checks, signature/encryption policy,
  source-boundary checks, and XML/source-splice validation.
* `PreparedTopology` retains the source snapshot, lineage, source version,
  execution context, read limits, generated relationship and content-type
  bytes, changed Part overlays, omitted members, appended members, source
  authorization tokens, and every managed reservation needed until readback
  and publication finish. It must not rebuild these values after the callback.
* `with_candidate` gives the owner a read-only effective view. The callback
  cannot publish or mutate the source. Source freshness and cancellation are
  checked before the callback, during bounded reads, and after it returns.
* `publish_to_stream` consumes the prepared value and uses the existing
  preservation writer. It checks source freshness and execution policy before
  the first external byte and while the sink accepts bytes. It is the only
  stage that charges the caller's `Resource::OutputBytes` budget.

The implementation should extract the existing preparation work from
`write_topology_to_stream` rather than duplicate it in DOCX. In particular,
the generated `ChangedOverlay` values, omitted physical members, appended
entries, relationship publications, content-type replacement, source guards,
and memory reservations currently assembled in the source-backed writer are
the natural contents of `PreparedTopology`.

## Effective candidate contract

`EffectiveTopology` is an OPC-owned typed overlay, not a serialized ZIP byte
buffer. It must present the same logical surface needed by a reopened
source-backed package:

* package relationships and every retained Part relationship set;
* the effective Part catalog, including added, replaced, and removed Parts;
* effective content-type defaults and overrides, with duplicate detection;
* physical member presence for collision and removal checks;
* lazy reads of unchanged source payloads through the original `ReadAt` source;
* reads of changed and added payloads through the prepared, source-authorized
  allocations; and
* source lineage/version and the same read, XML, graph, cancellation, and
  execution limits used by preparation.

The view must not materialize unchanged ZIP members merely to expose them.
Managed Part payloads remain budgeted handles for the lifetime of the view and
must not escape as an unaccounted `Arc<Vec<u8>>`. Generated relationship and
content-type XML may be shared with the prepared writer, but its reservations
must remain live through the callback and final publication.

The effective graph validator must prove at least the following for every
changed owner:

* the resulting relationship IDs are unique and each changed relationship has
  the planned type, target, and target mode;
* internal targets introduced by the plan resolve to retained or added Parts;
* removed Parts have no retained inbound relationship, whether inherited from
  the source or introduced by the plan;
* the effective content-type map has no duplicate default or override and each
  retained Part has the expected binding; and
* a Part removed by the plan has no remaining selected closure owner. Existing
  source dangling edges may remain only under an explicit preservation policy;
  the candidate must prove that the edit did not introduce a new dangling edge.

OPC's ordinary graph reader intentionally permits a relationship target with no
ZIP member. Therefore opening a physical candidate alone is not sufficient for
the no-new-dangling-edge requirement. The prepared graph must compare the
effective changed edge set against the source and the plan.

## DOCX readback contract

The DOCX owner should adapt its lifecycle readback to `EffectiveTopology`.
For the SVG attach/detach closure, the callback must verify:

* the selected inline or anchor and direct picture still resolve;
* the raster relationship, type, target, content type, and payload bytes are
  unchanged;
* the requested SVG owner is present or absent with the expected relationship
  ID, type, internal target, content type, and payload hash;
* the selected story XML resolves the same owner state as the prepared graph;
* the SVG Part has no outbound relationship; and
* shared-target incoming-edge and final media/content-type removal semantics
  agree with the prepared effective graph.

The callback must compare package-wide relationship and content-type fingerprints
or an equivalent prepared graph report. Comparing only selected story bytes and
selected SVG payload state is insufficient: it cannot detect a stale detached
relationship, an incorrectly retained media Part, an unexpected content-type
override, or a changed unrelated edge.

The semantic callback is the answer to the lifecycle design requirement that
the candidate be reopened after staged XML, relationship, media, and
content-type changes, provided the effective view is complete and typed. The
physical ZIP output remains a separate proof obligation; the callback must not
be described as validating ZIP local headers, central-directory offsets,
compression, or final member ordering.

## Publication and physical validation

`PreparedTopology::publish_to_stream` retains the existing source-preserving
physical authority. It must continue to:

* copy untouched source members lazily and exactly;
* emit only the prepared changed and appended members;
* preserve the admitted source lexical form for noncanonical relationship and
  content-type members while retaining current add/remove-only restrictions;
* reject unsupported opaque members, encrypted entries, signatures, and
  source namespace forms according to the existing policy;
* enforce prospective member, XML, relationship, archive, and output limits
  before allocating or accepting output; and
* report the exact accepted external byte count, including typed incomplete
  output after a sink failure.

Physical evidence must reopen the bytes emitted by `publish_to_stream` with
the ordinary bounded OPC reader and compare its typed graph to the effective
candidate report. It must also inspect raw ZIP records for unchanged members,
source preamble/trailing bytes policy, and changed-member placement where the
source-preservation contract promises those properties. This final reopen is a
test/evidence gate and does not justify retaining a complete output archive in
memory during production publication.

If a product requirement later insists on reopening the literal physical ZIP
*before* the caller's sink, the API must accept an explicit OPC-owned scratch
store (for example, a caller-provided bounded file-backed `ReadAt`/write
object). Scratch bytes need a distinct limit/accounting dimension or an
explicit caller reservation. The default API must not silently allocate a
complete ZIP `Vec`, create an unbounded temporary file, or use a network-backed
staging service.

## Resource and freshness rules

The prepared path must keep resource dimensions distinct:

| Resource | Preparation/readback/publication charge |
| --- | --- |
| `Memory` | Prepared changed XML/relationship/content-type allocations, effective graph metadata, and live managed view state. |
| `InputBytes` | Source reads and effective candidate payload reads, charged cumulatively under the caller context. |
| `Work` | XML/relationship parsing, graph traversal, hashing, and candidate semantic reread. |
| `OutputBytes` | Bytes accepted by the external publication sink only. |
| Scratch budget, if physical staging is opted in | Bytes retained by the caller-selected scratch store, with an explicit separate limit. |

Every prospective generated replacement and appended member must be checked
against its limit before constructing the temporary buffer. The prepared value
must retain all reservations until the callback and publication complete; a
reservation must not be released merely because the generated bytes are
temporarily unreachable from the plan.

The source version and lineage are checked when preparation starts, before
readback, after readback, before external publication, and after publication.
No candidate readback may authorize a plan for a different source package or
silently adopt a source revision observed during the callback.

## Required evidence before approval

The implementation is reviewable only after the following evidence exists.
The evidence should record exact source hashes and the commands used; prior
historical logs must not be relabeled as evidence for this API.

### OPC preparation tests

* exact empty-plan copy remains byte-for-byte and performs no candidate
  callback;
* one replacement, one addition, one removal, relationship append/remove, and
  content-type add/remove expose the expected effective candidate graph;
* canonical and admitted noncanonical relationship sources preserve their
  documented lexical bytes and reject unsupported replacement/mixed operations;
* duplicate Part names, duplicate relationship IDs, missing internal targets,
  duplicate content-type mappings, signature/encryption inputs, stale sources,
  and failed read/output/memory/work admission checks fail before external
  output; cancellation or budget exhaustion during publication retains the
  existing accepted-byte and incomplete-output accounting;
* a callback failure leaves the external sink untouched and drops all prepared
  reservations; and
* a sink short write/error reports accepted bytes through the existing typed
  incomplete-output contract.

### DOCX lifecycle tests

* native inline and floating SVG detach, synthetic inline and floating attach,
  shared SVG targets, final media removal, and exact inverse;
* strict core relationship dialect, unknown extension/comment/attribute
  preservation, duplicate/ambiguous owners, linked owners, malformed owners,
  and MCE ancestry refusal;
* candidate readback rejects an intentionally altered changed relationship,
  target mode/type, content-type binding, retained media Part, or story owner;
* no-op skips preparation and candidate readback; and
* rejected candidate/readback and low-budget paths leave the caller sink empty.

### Physical output and resource evidence

* every focused output is reopened by the bounded OPC reader and its typed
  graph equals the effective candidate report;
* unchanged source members retain their raw ZIP bytes and declared source
  metadata where promised;
* a large unchanged source does not become a full in-memory candidate during
  normal publication;
* managed probes show peak memory is bounded by source-backed cache plus
  prepared overlays and candidate metadata, not total ZIP output;
* low `Memory`, `InputBytes`, `Work`, and `OutputBytes` limits fail before the
  first external byte; and
* if physical scratch support is added, low scratch limits and scratch-store
  failures are covered without affecting the caller's output accounting.

Until these API, semantic, physical, and resource checks are materialized and
reviewed, the implementation remains pending acceptance. Passing focused tests
does not establish the complete managed-memory or performance contract.
