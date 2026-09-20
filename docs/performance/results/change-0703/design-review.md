# 0703 source review: capture-local PPTX projection reuse

Status: read-only design review at `f69d660800af19391a01c69331c919ce8c9d114c`.
No production source, Cargo metadata, or lockfile was changed for this review.
There is no performance claim in this document. The historical call counts below
are provenance for choosing a probe; they are not measurements of the current
binary.

## Finding

Opened capture repeats the default MCE projection of the presentation part
because it constructs the same borrowed `PresentationPart` through three
independent reads. `capture_internal` first calls
`PresentationPart::from_package`, then calls `presentation.slide_references`,
then creates `Presentation` and calls `capture_slides`
([opened/model.rs:618-628](../../../../crates/litchi-pptx/src/opened/model.rs#L618)).
The third call enters `Presentation::catalog`; that view already has a
successful `OnceLock<Vec<SlideReference>>` memo, but it is empty because the
first reference vector was parsed outside the view
([presentation/model.rs:15-59](../../../../crates/litchi-pptx/src/presentation/model.rs#L15)).

The historical 0693/0697 capture sequence for a 13-slide transitional package
was:

```text
presentation, presentation, presentation,
slide1 ... slide13,
presentation, presentation
```

That sequence is retained in
[`change-0697/sequences/all.txt`](../change-0697/sequences/all.txt); it must be
re-counted with fresh temporary instrumentation at this HEAD before any
implementation decision. The source mapping is clear: the first presentation
call validates the root, the second parses the direct catalog, and the third
parses the view catalog. Each slide's first pass uses one
`SlidePart::from_part_with_name` projection; later slides after a name error
use `SlidePart::from_part`, which also preprocesses once
([presentation/package.rs:49-107](../../../../crates/litchi-pptx/src/presentation/package.rs#L49)).

The historical 0697 isolated attribution put the five presentation calls at
3.32% of the sequence and slide processing at 96.86%; slide 11 alone was
42.99% ([0697 attribution](../../0697-mce-context-ownership-attribution.md#fresh-native-results)).
That makes this a bounded seam with potentially limited absolute return. The
catalog candidate should be selected only if the fresh 0703 trace confirms
enough duplicate bytes or calls, rather than because it is the smallest safe
change.

The notes load after the capture identity loop can process the presentation
again. `load_index_with_slide_root_proofs` calls `root_conformance` and then
`scan_xml` on the same presentation payload
([notes/package.rs:217-232](../../../../crates/litchi-pptx/src/notes/package.rs#L217)).
`root_conformance` tries Transitional and Strict in order and can therefore
make one or two MCE calls; `scan_xml` makes another call when it is reached
([notes/codec.rs:311-353](../../../../crates/litchi-pptx/src/notes/codec.rs#L311)).
The existing slide-root proof removes the repeated slide root scan, but it does
not retain processed XML and it does not cover this presentation scan
([parts/slide.rs:902-929](../../../../crates/litchi-pptx/src/parts/slide.rs#L902)).

## Capture and commit boundaries

The capture path has a useful small seam. `Presentation::catalog` already
memoizes a successful catalog for the lifetime of one immutable borrowed view;
only the capture caller bypasses it by parsing a separate vector first. A
private capture-only change can obtain the catalog from that view once, check
`limits.max_parts`, and then let `capture_slides` reuse the same memo. This
changes the three presentation projections (root validation plus two catalog
passes) to two (root validation plus one catalog pass), without retaining any
processed XML or changing the public `PresentationPart` methods. The root
validation and catalog parse remain in their current order, and slide
relationship, target, identity, name, and notes checks remain after the catalog
limit check.

The larger changed-transaction path is different. `Transaction::commit`
compacts changed slides, fingerprints the staged package using the source's
payload digest memo, captures the patch, and only then captures a new snapshot
for a non-empty patch
([opened/transaction.rs:1190-1241](../../../../crates/litchi-pptx/src/opened/transaction.rs#L1190)).
The call to `capture_with_revision_and_digests` skips repeated payload hashing,
but its documentation explicitly retains every capture validation
([opened/model.rs:580-598](../../../../crates/litchi-pptx/src/opened/model.rs#L580)).
`Patch::capture` compares resource bytes and relationships; it does not invoke
MCE ([opened/patch.rs:210-260](../../../../crates/litchi-pptx/src/opened/patch.rs#L210)).

An exact no-op patch returns `self.source` before recapture
([opened/transaction.rs:1222-1233](../../../../crates/litchi-pptx/src/opened/transaction.rs#L1222)).
Consequently, retaining an MCE projection in `Transaction` cannot improve the
no-op commit path. A changed transaction also has several invalidation seams:
slide XML and relationships can change, `move_slide` changes the presentation
part, notes edits rewrite a notes part, dependency transfers add or copy parts,
slide removal captures and publishes an intermediate snapshot, and signing
handling may remove a signature graph before the final capture. At publication,
`capture_candidate` can rebind only when the complete packages compare equal;
otherwise it falls back to capture with a payload-digest parent
([opened/patch.rs:814-850](../../../../crates/litchi-pptx/src/opened/patch.rs#L814)).
These facts make a transaction-retained processed-XML cache a larger ownership
and invalidation change than the capture-local catalog seam.

## MCE profile, ownership, and source identity

The low-level helper performs the 64 MiB PresentationML raw-size preflight,
then takes a second `Part::blob()` observation and passes that exact slice to
`process_ooxml`, returning a `Cow` beside the source witness
([parts/mod.rs:26-57](../../../../crates/litchi-pptx/src/parts/mod.rs#L26)).
This two-observation shape is deliberate for foreign `Part` implementations.
It must not be replaced by a cache keyed only by a pointer obtained from one
unverified `blob_arc()` call.

`process_ooxml` always uses the default OOXML capability and MCE limit profiles
([mce/codec.rs:1451-1454](../../../../crates/litchi-ooxml-common/src/mce/codec.rs#L1451));
the lower-level `process_markup_compatibility` API accepts explicit profiles,
and semantic slide text uses its own semantic limits. A cache shared across
those consumers would need the complete capability and limit identity. The
capture experiment must remain scoped to the default `process_ooxml` route.
When no MCE namespace is present, the result is `Cow::Borrowed`; a marked
document can return `Cow::Owned` and its output is bounded by the MCE profile.
Reuse must still run each consumer's raw and processed-size checks.

The existing fingerprint memo is not an MCE cache. `PartDigests` retains an
`Arc<Vec<u8>>` only to make its (address, length) digest key safe against ABA,
and `memo_key` rejects a foreign part when `blob_arc()` does not alias the
visible `blob()` slice
([opened/model.rs:313-429](../../../../crates/litchi-pptx/src/opened/model.rs#L313)).
It stores no transformed XML and re-feeds part metadata and relationships on
every fingerprint pass. A future transformed-XML cache would need the same
alias proof, a retained source owner for any address key, and an explicit
bounded transformed-byte budget. The current
`Limits::max_retained_candidate_bytes` is specifically the cross-slide-copy
serialized archive retention ceiling, so it is not authority for a general MCE
cache.

## Limits and semantic constraints

The capture-local catalog reuse has no new retained byte budget: the existing
`Presentation` view already owns its bounded reference vector, and the
processed `Cow` remains temporary. Any design that keeps a processed
presentation through notes must account for the independent limits: the part
reader allows up to `MAX_PART_XML_BYTES` (64 MiB), the default MCE output limit
is 512 MiB, and notes caps the presentation raw and processed scan at 32 MiB
([parts/mod.rs:23-24](../../../../crates/litchi-pptx/src/parts/mod.rs#L23),
[notes/mod.rs:38-44](../../../../crates/litchi-pptx/src/notes/mod.rs#L38)).
Holding a large `Cow::Owned` until notes validation would change peak retained
bytes. A proof-only design must retain only a bounded source witness and the
minimum validated inventory, with fallible admission and an uncached fallback.

The following are hard constraints for any later candidate:

* Every raw-size, MCE input/output, notes XML, slide-count, relationship,
  namespace, and allocation limit remains in force. Reuse may skip repeated
  transformation only after the same source and profile are proved.
* Root/content-type validation, catalog parse errors, the `max_parts` limit,
  slide relationship and identity checks, deferred name-error behavior, notes
  conformance, and typed refusal precedence remain observable in the existing
  order. A source identity mismatch must use the old uncached path.
* MCE output bytes, `Cow` ownership, and all notes and snapshot semantic values
  must remain exact. Unknown markup, namespace choices, relationships, and
  opaque package members remain source-authoritative.
* No public API, global cache, ambient process state, executor, or archive type
  is introduced. The work belongs to the PPTX owner and its existing
  `litchi-opc` boundary ([ADR 0024](../../../adr/0024-current-topology.md#L15)).
* Any retention is local to one capture or an explicitly owned immutable
  snapshot. It must not rely on `Arc<RwLock<_>>`, change immutable snapshot
  semantics, or make dirty edit state evictable.

These constraints follow ADR 0001's strict typed layers and preservation rules,
ADR 0003's immutable snapshots and atomic commits, ADR 0005's measured and
bounded resource policy, ADR 0006's validation/preservation/security contract,
ADR 0011's OPC ownership boundary, ADR 0024's current OOXML topology, and the
accepted lazy-decode and execution-budget boundaries in ADR 0030 and ADR 0031.
The full accepted-ADR manifest is `docs/adr/README.md`; no proposed ADR is
needed for the capture-local experiment.

## Alternatives

| Option | Scope and benefit to test | Main risk or cost | Disposition |
| --- | --- | --- | --- |
| Reuse `Presentation::catalog` in `capture_internal` | One existing successful catalog memo serves the length check and slide pass; no transformed XML is retained. | It changes the private capture caller from two catalog passes to one, so the probe must cover foreign or adversarial parts and refusal order. | Best first source candidate after the diagnostic trace, subject to measured duplicate work and ROI. |
| Fuse main root and catalog over one `ProcessedXml` | Could remove another default-profile presentation MCE pass and keep one temporary projection. | Must preserve the raw-size preflight, source witness, root-before-catalog errors, `Cow` lifetime, and fallback behavior for source identity changes. | Separate experiment if the fresh trace shows material main-part work. |
| Carry a processed presentation proof into notes | Could remove the post-slide presentation MCE calls while retaining no XML. | Requires a bounded scanner result, raw pointer/length witness, profile binding, and exact fallback/error masking; notes inventory and allocation behavior become part of the seam. | Larger follow-up, not the first change. |
| Retain transformed XML in `Transaction` or `Snapshot` | Could span capture and a changed commit in theory. | No-op commits already return the source; edits invalidate multiple parts; publication may drift; transformed bytes add memory and have no current budget. | Defer. |
| Global or cross-format MCE cache | Broad reuse across parts and profiles. | Violates source ownership and profile boundaries, risks stale/foreign bytes, and conflicts with ADR 0005's explicit state model. | Reject. |

## Best next experiment

First run a temporary diagnostic at this exact HEAD over the existing real and
generated packages, with no-op, one-edit, and two-edit transaction prefixes;
include the notes prefix if it is already available in the same harness. A
strict package and refusal fixtures are future production gates, not
requirements for completing this bounded diagnostic. Record for every
`process_ooxml` call:

1. phase and caller label;
2. part name when known;
3. raw pointer and length, raw SHA-256, and output pointer and length;
4. `Cow::Borrowed` versus `Cow::Owned`;
5. success or typed error and the profile label
   (`process_ooxml-default`).

The trace must be restored before any source census or gate. Compare only call
counts, duplicate raw/output identities, and byte totals; do not infer latency,
allocation, RSS, or throughput from the trace. It must distinguish the two
catalog passes and root call from the notes calls. Later production gates must
add strict/retry and refusal paths, foreign `Part` identity cases, exact
semantic and error-order parity, bounded memory checks, and representative
timing/resource measurements.

If the trace confirms the expected duplicate catalog work, the smallest
candidate is to let `capture_internal` obtain its `references` from the
already-created `Presentation` view's `catalog` and reuse that memo in
`capture_slides`. Keep the public `PresentationPart` API unchanged. Required
checks are the existing focused opened tests plus exact candidate/fresh
snapshot parity, MCE output/ownership oracle coverage, foreign `Part` identity
fallbacks, late name/relationship/notes refusals, and a before/after resident
byte check. Only a fresh representative timing and resource packet can decide
whether this small source change is worth retaining.

If the catalog seam is retained and the trace still shows meaningful repeated
main-part work, design a second candidate around one capture-local
`ProcessedXml`/bounded proof. Do not carry that candidate into a transaction or
global cache until its memory admission, source identity, profile key, and
publication/rebind behavior have an explicit proof and measurements.
