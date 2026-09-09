# DOCX bounded tail append source review

Status: the opening design review below is historical. It was written before
the DOCX tail append implementation appeared in the coordinator worktree and
records its acceptance conditions. The later source-review sections and the
current checkpoint at the end are the status of the implementation.

This review is based on the explicit-window contract in
[`change-0481/window-contract.md`](../change-0481/window-contract.md), the
accepted ADRs (especially 0001, 0003, 0005, 0006, 0010, 0011, and 0024), the
existing source-backed paragraph grammar in
[`paragraph_copy.rs`](../../../../crates/litchi-docx/src/source_backed/paragraph_copy.rs),
the bounded authoring rules in
[`streaming.rs`](../../../../crates/litchi-docx/src/streaming.rs), and the
decoded-member owner in
[`litchi-opc/src/source_backed/splice.rs`](../../../../crates/litchi-opc/src/source_backed/splice.rs).

## Review verdict

The operation can be accepted as the M1 one-paragraph decoded-splice
milestone only if the DOCX layer owns a semantic source and candidate scan.
`xml_minifier::audit::verify_authored_reader` is useful as a bounded XML and
compactness check, but it does not recognize WordprocessingML roots, direct
body children, plain paragraph grammar, or section placement. Passing that
generic audit alone is therefore a correctness blocker.

The implementation must keep the physical publication in OPC and keep the
DOCX public surface free of ZIP entry IDs, archive readers, locks, or complete
XML buffers. The current OPC plan already checks raw source and candidate
length/hash proofs and performs a fresh replay, but those checks do not replace
the format-owned semantic proof.

## Source admission and scanner shape

The source pass must retain the existing paragraph-copy admission policy:

* the package relationship and main Part must agree on Transitional versus
  Strict WordprocessingML;
* the main Part must be `/word/document.xml` with the ordinary non-macro main
  content type;
* external relationships, signatures, macros, unsupported story dependencies,
  and enforced document/write protection or tracked revisions must refuse the
  changed operation before output;
* the root must be `document`, containing one direct `body`, and the body may
  contain only direct plain `p` children plus at most one final direct
  `sectPr`; and
* a plain paragraph remains the narrow `p` -> `r` -> `t` grammar. Attributes
  remain limited to namespace declarations and the existing `xml:space`
  values. Unknown body wrappers, tables, MCE branches, complex runs, and
  unsupported markup must remain typed refusals.

The scanner should consume a verified decoded `BufRead` and retain only scalar
proofs, a bounded scope/namespace state, bounded token scratch, and hashes. It
must not copy the main XML, retain paragraph ranges, collect an opaque
`sectPr`, or build a complete candidate. The source pass needs a decoded byte
offset for the exact start of the insertion event. `quick_xml` event positions
are acceptable only when the surrounding reader proves that they are decoded
raw offsets; deriving an offset from reserialized or normalized XML is not.

The candidate pass must scan a reader made from source prefix, the bounded
authored fragment, and source suffix. It must prove the same grammar and
well-formedness, exactly one additional direct paragraph, the expected source
order, and the expected section placement. A generic XML parse or a paragraph
count by itself is insufficient.

## Namespace binding at the insertion point

The generated paragraph must be namespace-valid in the body context for both
accepted Word dialects and for all accepted root prefix forms. In particular,
the source may use a default Word namespace:

```xml
<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <body>…</body>
</document>
```

An inserted `<w:p>` without a local `xmlns:w` declaration is unbound in that
document. Conversely, blindly inheriting or reusing a prefix is unsafe if a
source scope shadows it. The authored fragment should either use a fixed
self-contained `w` declaration on its root, or the scanner must capture and
authenticate the in-scope Word binding used by the fragment. The candidate
scanner must resolve the generated names and reject an unbound or foreign
binding; it must not merely compare lexical prefixes. A redundant declaration
is acceptable only if the admission rules and exact candidate proof explicitly
allow it.

The namespace declaration belongs to the authored fragment and is therefore
part of its length/hash and append limit. It must not be injected by a later
OPC replay step. Namespace declarations in the opaque section span remain
source bytes and must never be normalized.

## Final `sectPr` and opaque source span

For a final direct `w:sectPr`, record the start offset of its opening lexical
event, not the end of the body or a reconstructed element. Insert the generated
paragraph immediately before that offset. For a body without `sectPr`, record
the start offset of the direct `</w:body>` event. Whitespace before either
anchor remains in the source prefix and must be preserved byte-for-byte.

The scanner must reject a `sectPr` that is nested in a paragraph, wrapped in
another body child, non-final, or duplicated. It may treat the admitted final
section element as opaque for semantic purposes, but it still has to consume
its complete well-formed span under bounded depth/events/token/attribute
limits. The candidate proof should carry a compact section flag plus source
and candidate anchor/span values (and a section digest when computed); the
candidate anchor must equal the source anchor shifted by the fragment length,
and its final section span must be the unchanged source suffix. Do not parse,
serialize, or regenerate `sectPr` as part of this operation.

Opaque does not mean “assume the bytes are safe” and it does not mean “force
every descendant into the root Word namespace.” Unknown section children,
foreign attributes, and in-scope declarations that the source policy admits
must remain untouched and must be included in the source-span proof. The
scanner still has to establish balanced XML and bounded namespace/token state;
it must reject malformed names and whichever comments, processing
instructions, CDATA, DTD, or active MCE forms the stated admission policy
excludes.

Comments, processing instructions, CDATA, DTDs, MCE wrappers, and other
unsupported markup remain refusals according to the explicit-window contract,
including when they are encountered outside the admitted opaque section policy.
If the implementation chooses a narrower section policy, it must refuse those
inputs rather than silently discard or normalize them.

## Authored fragment and proof matching

The smallest accepted authored input is one caller-owned UTF-8 plain-text
paragraph. Empty text is still an authored empty `w:t` paragraph; it must not
be classified as an exact no-op. Invalid XML characters, tabs, line breaks,
and structural markup must fail before publication, using the existing
streaming writer's escaping and typed input conventions.

The DOCX plan needs a compact format-owned proof that binds at least:

* source lineage/version and main Part identity;
* source decoded length and SHA-256;
* decoded insertion offset and whether a final `sectPr` exists;
* source paragraph/event/depth counts;
* authored fragment length, capacity/limit, and SHA-256;
* candidate decoded length and SHA-256;
* candidate paragraph/event/depth counts and shifted section anchor/span; and
* all relevant source, append, output, parser, and workspace limits.

The low-level `SourcePartSpliceProof` carries the raw source/fragment/
candidate length/hash scalars, but it does not carry the DOCX grammar or section
facts. Those DOCX facts must be checked before calling
`prepare_source_part_splice`, retained in the DOCX plan/publication, and
compared again where the public inverse or rehydrated operation needs them.
The fragment digest must be checked both before and during replay. The source
and candidate digest checks authenticate the raw prefix/suffix construction;
the semantic scan authenticates that the inserted bytes are exactly one plain
paragraph at the intended body position.

## No-output preflight and publication

`prepare` must not write to a caller sink. The sequence should be:

1. Check source freshness, cancellation, topology, package security policy,
   limits, and all required workspace reservations.
2. Scan and hash the complete source member.
3. Encode and bound the authored paragraph, including its namespace binding.
4. Scan and hash the complete candidate through a prefix/fragment/suffix
   reader, checking the DOCX semantic proof.
5. Call the OPC decoded-splice preparation, which rechecks physical source
   identity, source/candidate raw proofs, fragment capacity, XML audit, output
   limits, and preservation/replay workspace.
6. Recheck freshness/cancellation and return the plan. Only `write_to` may
   begin ZIP output after this point.

Any source, candidate, topology, limit, allocation, cancellation, or proof
failure in these stages must leave the sink untouched. `write_to` must use a
fresh source reader and the same fragment/proof, then map failures after the
first accepted byte to the existing incomplete-output/partial-progress type.
The sink is sequential and caller-owned; no API may imply rollback.

The changed path must preserve the existing signature/encryption/protection
policy. The exact no-op path in the generic OPC owner may copy a signed or
malformed source when its proof explicitly describes zero insertion, but the
DOCX tail operation must not turn empty authored text into that branch.

## Inverse and provenance

An immediate inverse must authorize the exact complete candidate archive
fingerprint before it emits anything, then copy the retained original
`SourceArtifact` byte-for-byte. It must retain source lineage and use the
current candidate execution context when the reopened candidate is managed.
Foreign, stale, or semantically different candidates must fail with no output.
An authored-fragment digest is not enough to make a durable inverse: exact
physical restoration requires an explicit retained or durable original source
provider, as stated by the window contract.

The public DOCX publication must expose only format-owned proof and typed
errors. It must not expose `PreservationEntryId`, ZIP compression details,
`PartView`, `Arc<RwLock<_>>`, or an OPC archive handle through ordinary
signatures. A public method may delegate to the OPC publication internally,
but it must preserve the DOCX source/candidate proof and error provenance.

## Resource accounting and truthful claims

The 0482 OPC owner charges the retained fragment capacity, XML audit envelope,
preservation index, replay/compressor state, ZIP output, and decoded reader
work. The DOCX semantic scanner adds its own bounded state: event/token
windows, namespace resolver state, scope stack, counters, and any section
hashing scratch. These allocations must either be charged through the execution
context before source bytes are consumed or be included explicitly in the
documented operation bound. `NsReader` must not be allowed to grow an
unbounded token buffer from an attacker-controlled tag.

The resulting M1 claim may say that complete main/candidate XML is not
materialized and that the decoded payload scan is bounded by explicit parser,
fragment, and package-metadata limits. It must not claim complete-window RSS,
constant memory, or memory proportional only to the paragraph until caller
scratch, package metadata, ZIP preservation index, and compressor state are
all part of one enforced contract. Any benchmark must report source storage,
catalog/index, ZIP codec, caller text/fragment, and allocator overhead
separately.

## Required focused cases before acceptance

The implementation should have tests for:

* Transitional and Strict roots with `w:` prefixes and default Word
  namespaces, including insertion before a final `sectPr`;
* no `sectPr`, one final opaque `sectPr` with unusual lexical bytes, and
  exact preservation of its complete source span;
* non-final, duplicate, nested, and wrapped `sectPr` refusals;
* unknown direct body children, wrappers/tables/MCE, complex paragraphs,
  malformed XML or namespace declarations (including duplicate raw attributes
  and namespace declarations), comments/PI/CDATA/DTD policy, invalid text
  characters, empty text, and fragment-limit boundaries;
* source mutation or cancellation during source scan, candidate scan, and
  replay, with an empty sink for preflight failures; cancellation or execution
  failure after the first accepted byte must retain typed OPC
  `IncompleteOutput` progress, including forward and inverse publication, and
  a simultaneous source change must retain source-freshness precedence;
* source/candidate length/hash/offset/section-proof mismatches;
* signed, encrypted, macro-enabled, protected, external-link, and unsupported
  dependency refusals;
* Store and Deflate main members, raw untouched members and ZIP metadata,
  sequential short writes/failures, and exact immediate inverse; and
* managed and unmanaged positional sources, including a reopened candidate
  used for inverse authorization, with explicit options checked for both
  semantic and physical ownership.

The benchmark contract in `harness-contract.md` deliberately measures only
the no-`sectPr` route initially. That is suitable for a scoped performance
comparison after the correctness gate, but it cannot stand in for the section
placement and namespace tests above.

## Handoff

Coder handoff: make the DOCX scanner the authority for Word grammar,
namespace binding, body/section placement, and semantic source/candidate
readback. Pass only authenticated compact raw proofs and a bounded fragment to
OPC. Keep all preflight checks before the first sink byte and keep empty text a
real paragraph.

Coordinator handoff: do not accept a green generic XML-audit or OPC replay
test as DOCX completion. Require the semantic refusal matrix, exact opaque
`sectPr` preservation, namespace-default positive case, proof mismatch and
no-output checks, and a source review of the final implementation before
calling this batch complete.

## Concrete implementation review — scanner pass

This pass reviews the implementation currently in
[`tail_append.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append.rs).
The initial design findings above remain the acceptance history; the findings
below are source-level blockers in the current revision.

### Resolved in this source pass

The latest scanner revision now blocks direct body content after a closed
`sectPr` (for both `Start` and `Empty` section events), rejects a second
document root, and compares the package relationship dialect with the scanned
root before fragment generation. These are no longer current blockers; keep
the non-final-section, root-count, and relationship/root-mismatch fixtures as
regression cases. The verified-reader mapping also now preserves an OPC
transport/freshness error over a simultaneous callback error, as required by
the 0482 error-precedence contract.

The following earlier scanner findings are also resolved in the current source
pass:

* The opaque branch now enables `quick_xml` end-name checks, resolves every
  opaque element, validates ordinary attributes and expanded-name duplicates,
  and rejects unknown qualified names. It therefore establishes the bounded
  lexical/name proof while still admitting bound foreign vendor content.

* Comments, processing instructions, CDATA, doctypes, general references, and
  non-whitespace text are explicit refusals inside the opaque section. Keep
  those events in the refusal matrix so a future broadening is deliberate.

* `validate_candidate` now requires the candidate insertion offset to equal
  the source offset shifted by the fragment length, along with the generated
  paragraph marker, event-count increase, depth evidence, and unchanged
  section span. Retain proof-mismatch fixtures for these relationships.

### Proof, policy, and resource gaps

8. **The execution authority split is resolved (regression guard).** The public
   execution-context setter and unmanaged fragment lease are gone. `effective_options`
   derives its private execution value solely from the source-backed package,
   which owns semantic scans, fragment allocation, OPC audit/replay, and output
   accounting. Forward `Publication` no longer retains an execution or
   cancellation option; inverse publication delegates to OPC as a new operation,
   so a reopened managed candidate selects its current package context. The
   extra public cancellation token is documented as applying to semantic parser
   events and forward sink callbacks, while the package context supplies the
   full OPC cancellation fence. Keep managed budget and reopened-candidate
   context tests to prevent a second authority from returning.

9. **The semantic scanner reservation now has an explicit parser envelope,
   but the envelope still needs an evidence check.** `scan_reader` reserves
   before constructing `NsReader` and includes token windows, retained open
   names, namespace resolver/declaration vectors, attribute scratch, and the
   `depth + 1` admission state. The settings-root probe uses the same helper.
   Keep a focused resource test or documented derivation that tracks
   quick-xml's vector growth and any future parser-state changes; without that
   regression proof the formula can silently become an undercharge.

10. **The DOCX-to-OPC audit profile is resolved (regression guard).**
    `make_splice_limits` now passes the candidate, depth, event, token, text,
    and per-event attribute ceilings explicitly; the aggregate attribute budget
    is no longer accidentally tied to the event count. Independent audit hard
    ceilings remain an intentional refusal boundary. Keep an over-ceiling
    policy fixture so the refusal stays pre-output and visible rather than
    silently widening the OPC profile.

11. **The public DOCX proof is diagnostic rather than standalone.**
    `SourceVersion` already carries the source identity and revision, and the
    plan borrows the canonical package and main Part, so this is not by itself
    a stale-source safety hole for the current non-rehydratable API. The
    format proof still omits an explicit part-name/artifact/fragment wrapper,
    however, and the low-level scalar proof is what callers see through
    `Plan::splice_proof`. Keep the non-rehydratable ownership invariant clear;
    if proof values are later persisted or rehydrated, add explicit
    DOCX-owned bindings before doing so.

12. **Opaque namespace admission is now guarded (regression coverage remains
    required).** `NsReader` rejects reserved-prefix rebinding, while the
    opaque attribute pass rejects reserved default values, empty qualified
    declarations, unknown qualified attributes, and duplicate expanded
    ordinary attributes. The pinned `quick-xml 0.41.0`
    `BytesStart::attributes()` iterator also performs its default raw-key
    duplicate check, including repeated namespace declarations; keep that
    dependency behavior covered by malformed fixtures or make the check
    explicit if the dependency is upgraded. `require_opaque_namespace` still
    permits genuinely unqualified vendor names, as the opaque policy requires.
    Keep malformed declaration, unknown-prefix, reserved-binding, duplicate
    raw-key, and duplicate-expanded-name fixtures so these checks remain tied
    to the preserved raw span.

13. **The compact lexical fixture is aligned with the accepted audit closure
    (regression guard).** The focused strict and transitional section fixtures
    are compact, so the DOCX semantic and OPC audit passes admit the same
    bytes. Preserve the explicit 0482 compactness policy and keep a pretty
    structural-whitespace fixture as an intentional refusal; do not treat that
    refusal as source-preserving acceptance.

14. **The public DOCX surface exposes generic scalar OPC proof diagnostics.**
    `Plan::splice_proof` returns `SourcePartSpliceProof`, and
    `Publication::candidate_artifact_fingerprint` returns
    `SourceArtifactFingerprint`. These values do not expose ZIP entry IDs or
    archive handles and are therefore not an immediate ADR ownership breach.
    DOCX-owned wrappers would make the boundary clearer if this API becomes a
    durable or format-specific proof surface; preserve the current scalar
    values only while the plan remains borrowed and non-rehydratable.

15. **The fragment-owner transfer is now accounted and package-bound
    (regression coverage remains useful).** The allocator reserves the
    requested length, rejects any larger reported `Vec` capacity before the
    owner escapes, and transfers the single reservation into the OPC plan.
    Pointer identity rejects fragments allocated by another package. Keep
    focused coverage for foreign-package refusal, mutation before proof
    hashing, source/context cancellation during zero-fill, and reservation
    release after a rejected handoff; these are regression tests rather than
    current ownership blockers.

16. **Cancellation scope is now explicit (regression guard).** The separate
    public token is documented for semantic parser events and forward sink
    callbacks. The source package's execution context remains the cancellation
    authority for OPC preflight, replay, exact no-op copying, and inverse
    fingerprint/copy work. Tests should keep those scopes distinct rather than
    expecting an options-only token to interrupt OPC work before a sink callback.

17. **Changed-entry encryption preflight is resolved in the current source
    (regression guard).** The changed path checks `has_encrypted_entries()`
    before topology and decoded main-part scans, while the explicit no-op route
    retains its separate exact-copy policy. Keep the encrypted metadata test
    asserting no payload read and no output.

18. **Publication error precedence is resolved (regression guard).** OPC now
    recognizes a bare `litchi_core::ExecutionError` carried through a format
    sink, so an options-token cancellation remains typed and accepted bytes
    remain inside `IncompleteOutput`. The existing result-first precedence keeps
    authoritative source freshness and transport errors ahead of the callback
    side channel. Keep managed and unmanaged mid-prefix cancellation tests.

19. **Inverse cancellation authority is resolved (regression guard).** The
    publication does not carry the forward options token into a later inverse.
    OPC inverse fingerprinting, workspace reservation, replay/copy, and output
    checks all use the current candidate package context, with its full typed
    cancellation and freshness precedence. This is a new operation and should
    retain current-package managed and unmanaged inverse coverage.

20. **The checked-writer retry loop is resolved in the current source
    (regression guard).** `PublicationCheckedWriter::fail` now uses terminal
    `io::Error::other` after recording the typed execution failure, matching
    OPC's `ContextCheckedSink` and avoiding `Write::write_all`'s retry behavior
    for `Interrupted`. Keep a cancellation publication test to prevent a
    future change from restoring a retryable error kind.

21. **Nested Word `sectPr` rejection is resolved (regression guard).** The
    opaque branch now rejects any nested section-properties element whose
    resolved namespace is either admitted WordprocessingML dialect, while a
    same-local-name foreign vendor element remains eligible for the opaque
    policy. Keep both cross-dialect fixtures and assert an empty sink.

22. **Owned text capacity is now bounded and charged (regression guard).** The
    `Into<Cow<'text, str>>` admission checks both `String::len()` and an owned
    `String`'s retained capacity against `max_text_bytes`, reserves that full
    capacity under the managed package context, and drops the input and lease
    after fragment encoding. On an unmanaged package the finite text-capacity
    limit is the explicit bound; there is no hidden execution owner to charge.
    Keep a short, over-capacity `String` fixture and a managed reservation test
    so a future length-only optimization cannot reopen the gap.

23. **QName and XML 1.0 admission is resolved (regression guard).** Source,
    candidate, and opaque section events now validate qualified names with the
    shared XML 1.0 grammar, including namespace declaration keys and ordinary
    attributes. XML declarations are restricted to the admitted UTF-8/XML 1.0
    form. Keep malformed prefix/local-name, non-UTF-8 name, and declaration
    fixtures.

24. **Inverse context authority is resolved (regression guard).** `Publication`
    no longer stores preparation options. The OPC inverse selects the current
    candidate package context for fingerprinting, reservations, replay, and
    output, with the retained source context used only where OPC's unmanaged
    fallback requires it. Keep a reopened candidate with a distinct managed
    context in the inverse matrix.

25. **Settings admission is now fenced at its owning phase boundaries
    (regression guard).** `validate_topology` runs the borrowed source guard,
    reserves separate MCE output and scratch owners, and checks the effective
    package context and extra token before and after the synchronous MCE pass.
    An owned MCE result retains only its bounded output lease; the scratch
    lease is released before the post-MCE guard. That guard recounts the actual
    processed bytes, namespace state, events, and semantic payload under the
    workspace remaining after retained output. The complete borrowed model path
    then reserves `settings_model_workspace_requirement` and prepays three
    event units per observed model pass before parsing settings, extensions,
    and mail-merge data, with a cancellation/work check at the phase boundary.
    This is the documented cooperative boundary for codecs that do not accept
    an observer; it is no longer an unfenced owner or work gap. Keep MCE output
    expansion, post-MCE namespace growth, model-limit, and boundary-cancellation
    fixtures.

26. **XML declaration cardinality and placement are resolved (regression
    guard).** Source, candidate, and root probes reject duplicate declarations,
    validate the XML 1.0/UTF-8 form, and admit a declaration only in the
    initial prolog position. Keep duplicate, late, and malformed declaration
    fixtures.

27. **The public options description now matches the authority cleanup
    (regression guard).** `Edit::with_options` describes additional cooperative
    cancellation options, while the source-backed package constructor owns the
    execution budget used by semantic and OPC work. Keep the managed default
    path and extra-token scope tests aligned with that documentation.

The settings model/MCE fence and options documentation are current and are no
longer open blockers. Execution authority, inverse context, cancellation scope,
cross-dialect nested-section rejection, QName validation, XML declaration
cardinality, owned-input capacity, and raw/expanded attribute duplicate
rejection are covered; retain their regression fixtures. The audit profile and
compact lexical fixture are deliberate bounded-closure policy choices. The
future bounded authored-stream objective remains open as documented by change
0481; this one-paragraph API does not close that work.

## Current sourcefreeze checkpoint

The latest read-only pass found no concrete remaining API, semantic, typed-error,
publication, or resource-ownership safety blocker in the current implementation.
The source scanner remains the authority for the admitted Word grammar,
namespace binding, opaque final-section validation, insertion-anchor shift, and
candidate readback; OPC remains the authority for physical source freshness,
replay, output progress, and inverse authorization. The MCE/model workspace
leases and phase-boundary work/cancellation checks are now present, and the
public options text matches the package-owned execution authority. The initial
“implementation not present” text above, and the earlier MCE/options findings,
are retained as history only.

Acceptance still depends on retaining the focused malformed-input,
namespace/section-preservation, proof-mismatch, cancellation/no-output, and
resource-envelope regression coverage. This review did not run tests, builds,
formatting, profiles, or Git operations. The bounded authored-stream objective
from change 0481 remains future work.
