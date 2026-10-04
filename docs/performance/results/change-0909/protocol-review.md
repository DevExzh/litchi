# 0909 PPTX owning-input control protocol review

This is a read-only protocol review for the bounded owning-input PPTX control.
It does not build, run a workload, collect a profile, or change production
code. The control is a correctness and physical-preservation qualification;
it is not a performance result. A later timing or instruction experiment must
bind this protocol and retain its own source, binary, timer, and sink receipts.

## Review status and material risk

The 0909 source base is the 0908 child `57de5495f899a43574110f8c572823a7d9907e5f`,
with the source and lockfile bindings in `source-base.json`. The control must
use the same 0907 multi-slide corpus and edit specification:

* 200 slides with eight text boxes per slide;
* eight deterministic 2 MiB media Parts;
* selected slides `[0, 28, 57, 85, 114, 142, 171, 199]`;
* 64 deterministic text replacements;
* the 0907 source archive (229 OPC Parts, 445 ZIP members, 17,017,139 source
  bytes, SHA-256
  `61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`).

The main risk is an overclaim about preservation. A semantic reopen, an OPC
Part digest, or equality with the 0907 eager output does not prove that an
untouched ZIP member was copied. Conversely, equality of an untouched raw
member proves an output property but cannot prove that the writer avoided
decoding, recompressing, or copying it internally. The current `PackageWriter`
has private `Copy` and `Regenerate` planning decisions, but no public trace or
counter exposes those decisions. The control therefore needs an independent
raw ZIP census for the published bytes and must report the residual path-proof
gap explicitly. Any later instruction or wall-clock claim requires a separate
profile or timing scope; this packet cannot infer avoided Deflate work from raw
equality.

## Current control-source blockers

The first control source reviewed in this batch has several qualification
gaps that the result packet must close before it can report `qualified`:

* The current snapshot now records owned-source state and exact source length
  and digest at ingress, after the changed commit, and after a no-op commit.
  Keep those observations tied to the actual package used by each operation,
  and make the published gate derive from them rather than from a standalone
  literal. The changed state should be reported as loss of exact-source
  authorization while the retained archive remains available for targeted
  preservation; the no-op state should retain authorization and exact bytes.
* The Rust `raw_census` is generated through the harness's own
  `raw_zip_members`/`raw_zip_layout` helpers. It is useful inventory, but it is
  not independent evidence. The separate `zip-census.py` pass must run over
  the retained source, no-op, forward, inverse, and replay artifacts, bind the
  script/method identity, and classify every changed member against the
  committed patch closure. Until that receipt exists, a
  `raw_census_artifacts_exported` flag and a hard-coded `status: "qualified"`
  overstate the result.
* The current Rust inventory records member hashes, order, and comment hashes,
  but it does not itself establish the allowed changed-member set or the
  central-record field comparison. The independent census must reject any
  changed member outside the proven closure and must report differences rather
  than treating the inventory as a preservation gate.
  The Rust `RawZipCensus` also omits the retained forward-replay artifact, so
  the external census must include it or state that replay was not physically
  audited.
* The current source exercises both partial and zero-byte sinks, but the
  serialized `Gates` value reports only the partial sink. The qualification
  packet must retain a zero-sink receipt as well: zero accepted bytes, the
  documented refusal, and a successful retry on a fresh sink. The partial sink
  must retain and report its exact accepted prefix. Preserve the typed
  `IncompleteOutput`/zero-byte I/O classification where the writer exposes it;
  an error string alone cannot verify the sink contract.
* The exact and under-limit input probes now run, but the finite-limit gate is
  also currently a literal and discards the over-limit error identity. Record
  the accepted exact-bound case and the typed over-limit refusal so this gate
  cannot pass after a probe was accidentally removed or weakened.
* The report currently compresses inverse evidence into
  `inverse_exact_source_verified`. Preserve separate resource-byte, semantic,
  and full-archive inverse results. The generated fixture may pass all three,
  but `Patch` itself promises the first two logical forms and does not make a
  full ZIP-byte witness part of the general patch contract.
* The final coverage declaration is primary-corpus-only. It does not exercise
  borrowed `from_bytes` ingress, a physical-only same-logical-source variant,
  an unrelated-resource stale mutation, or a second fresh full transaction.
  Keep those as explicit residual coverage and do not generalize this one
  owning route into a format-wide provenance guarantee.
* The runner now writes the source, no-op, forward, inverse, and replay
  archives as each stage succeeds, so a failure retains the reached archive
  evidence. `write_failure` still records only the final error and corpus
  manifest: it does not identify the failed stage or retain the accepted sink
  prefix, zero-sink refusal, patch bytes, or partial census state. Add a
  bounded pre-failure receipt (or explicitly state that a stage was never
  reached) before returning, while keeping failed output out of the
  successful-output oracle.

These are evidence and reporting requirements; they do not require a new
writer architecture. Once the control records actual ownership state, binds
the independent census, exercises both sink boundaries, and preserves
pre-failure evidence, the existing transaction, commit, durable patch, stale,
foreign, and replay route can be qualified under the gates below.

## The two ingress routes

The 0908 eager leg follows this route:

```text
SourceBackedPackage::from_read_at(source)
    -> SourceBackedPackage::into_opc_package()
    -> Package::from_opc_package(opc)
```

`into_opc_package` creates a fresh `OpcPackage`, decodes the source Parts, and
adopts their logical payloads. It does not transfer the source archive, the
physical member index, or a verified compressed-member token. `from_opc_package`
sets `physical_source_provenance` to `false`. The subsequent writer therefore
has no `preservation_source` to pass to its targeted preservation planner and
falls back to a complete materialization. This is the route whose `to_bytes`
owner dominated the 0908 profile.

The owning control must use one of the source-preserving PPTX constructors:

* `Package::from_vec` or `from_vec_with_limits` is the primary control;
* `Package::from_reader` is a valid owning route, but it reads the input into a
  newly owned buffer and must expose that copy as a separate scope;
* `Package::open` is also source-preserving, with filesystem input and its own
  I/O scope.

`Package::from_bytes` is borrowed ingress. It deliberately calls the
non-provenance OPC path and is not an owning control. A successful semantic
edit after `from_bytes` must not be described as physical source preservation.
An `Arc` or a caller-retained `Vec` also does not change this distinction
unless the package constructor explicitly admits that allocation as owned
source storage.

The primary `from_vec` control should clone the corpus into a fresh `Vec`
before its operation boundary, then move that `Vec` into `Package::from_vec`.
That clone is test setup and must be recorded as such. It must not be silently
included in a claim about the owning package's parse, edit, commit, or publish
cost. If a `from_reader(Cursor::new(&corpus))` lane is added, its input copy is
inside the constructor and must either be timed separately or reported as an
explicit part of that lane. The two lanes are not interchangeable evidence.

The control should record the capability state before and after mutation:

| point | expected state |
| --- | --- |
| immediately after `from_vec` | an owned exact source is retained; the source digest and length are available; physical provenance is true |
| after a no-op transaction | exact-source publication remains authorized and output is the source byte-for-byte |
| after a changed commit is applied | exact-source shortcut is revoked, but the retained source archive and preservation provenance remain available for targeted publication |
| after any refusal | the package and the retained source remain unchanged; no output may have been accepted when validation failed before emission |

The public OPC `is_unmodified_owned_source`, `exact_source_sha256`, and
`exact_source_len` observations can support the first two rows. The changed
row must not be inferred from a digest alone: a changed package must be shown
to publish the requested logical bytes and to retain the original archive as
the preservation source. A changed output equal to the source is a failed
edit-oracle result, not evidence that the edit was free.

## Operation and transaction scope

The owning control should use the same opened transaction family as the 0907
eager leg:

```text
Package::from_vec(owned_source)
    -> opened_presentation_transaction()
    -> set_shape_texts() for the same eight slides
    -> Transaction::commit()
    -> Package::apply_opened_presentation_commit()
    -> sequential OPC publication
```

The transaction and commit are material correctness work. They capture and
validate the complete opened presentation, dependency closure, relationship
state, notes policy, and complete logical revision. The owning control must not
skip `opened_presentation`, complete revision fingerprinting, or candidate
reopen merely because only eight slides are edited. The 0908 profile showed a
large snapshot owner, but that cost is part of the current stale-source and
publication contract.

The control's untimed correctness path should prepare an expected semantic
oracle outside the publication sample and then exercise the same operation on
fresh owning packages. If a later benchmark is added, its boundary must be
declared before capture. To compare with the 0907 eager helper, the measured
owner would need to include owning open, the 64 replacements, commit,
qualification replays, publication, and the bounded sequential sink. Input
buffer creation and the expected-output oracle should remain outside that
boundary, exactly as their roles are declared. A from-reader input copy needs
its own labelled phase.

The control must retain the production operation and the qualification replays
as separate labels even when a future timer combines them. The existing eager
helper applies the inverse and forward patch after publication; those checks
are useful gates but are not independent production publication phases. Do not
subtract their cost from a future route total without adding explicit phase
boundaries.

The expected output must be derived from the same logical edit, not by requiring
the owning output to equal the 0907 eager output byte-for-byte. The eager
output was produced by a provenance-free full writer and can have different
compression records, timestamps, headers, central offsets, and archive
framing. The owning output is expected to retain source physical records for
untouched members. Equality with the old eager SHA-256 is therefore neither a
required semantic oracle nor a preservation proof.

## What inverse identity means here

The opened `Patch` is exact over its logical OPC resource states. Each delta
contains the before and after content type, decoded Part bytes, and
relationships; the optional root delta contains the package relationships.
Its complete revision is a logical fingerprint of that graph and opaque
non-Part payloads. It intentionally excludes ZIP member offsets, compression
streams, timestamps, local-header framing, central-directory order, and the
archive comment.

The control must report these properties separately:

1. **Resource-byte inverse identity.** After applying the durable inverse to
   the published candidate, every Part payload, content type, relationship,
   root relationship, opaque logical member, and complete logical revision
   equals the original source graph. This is the minimum reversible-patch
   gate. Comparing only slide names and text is too weak.
2. **Semantic inverse identity.** Reopen the inverse result and verify the
   complete slide order, selected and unselected slide semantics, media and
   relationship closure, and the same source revision. This is useful evidence
   but cannot replace the resource-byte check.
3. **Full archive byte identity.** Compare the entire inverse artifact with the
   original ZIP bytes, including all local records, compressed data, central
   records, EOCD comment, and any accepted trailing framing. This is a stronger
   physical result and must be recorded independently.

The first two are expected from the current `Patch::inverse` contract. The
third is not implied by that contract. A targeted writer can copy every
untouched source member while regenerating a changed slide member; applying
the inverse can restore the slide's decoded bytes and logical state while
emitting a different Deflate stream or header for that member. Such a result
must be labelled “resource-byte and semantic inverse passed; full archive
identity not established.” It must not be called an exact source restoration.

For a production eager provenance bridge, full archive identity after inverse
is a valuable admission gate for this corpus. If the existing route cannot
provide it, the candidate must either document the narrower resource-byte
contract and preserve all untouched physical members, or refuse a claim of
byte-for-byte physical reversibility. A semantic pass alone cannot authorize a
general provenance bridge.

The no-op case is different: an unmodified owned package has an exact-source
shortcut and must publish the complete original archive byte-for-byte. This
is an unconditional no-op gate, including the source ZIP comment and framing.

## Independent raw ZIP census

The control must build a raw census outside the semantic package reader and
outside any prospective writer decision. The census should be bounded and
create-only, and should retain enough data to compare source and candidate
without treating an `ArchiveReader` Part view as physical evidence. For every
member, record at least:

* normalized and raw name bytes, duplicate/case-fold collisions, and member
  count;
* local-header bytes and the complete local span through the compressed data;
* compression method and flags, CRC, compressed and uncompressed sizes,
  timestamps, version fields, extra fields, and local-header name bytes;
* central-directory record bytes, with the relative local-header offset
  zeroed only for comparisons where an earlier regenerated member necessarily
  moves the offset;
* physical local order, central-directory order, archive comment, EOCD and
  archive-end offset, and any trailing bytes allowed by the source policy.

For the unchanged-member comparison, the expected changed set must come from
the committed patch and the complete dependency closure. Do not hard-code
“the eight slide XML files” until the control records the patch resources and
proves that no presentation relationships, content types, notes, or other
closure member changed. Then require:

* source and candidate member sets agree, unless the patch explicitly adds or
  removes a member;
* every source member outside the proven changed closure has identical local
  span/compressed bytes and identical central record fields apart from the
  offset field forced by layout movement;
* every opaque/non-Part member remains present and raw-identical;
* for a same-topology edit, local and central member order remains unchanged;
* archive comment and accepted trailing framing remain unchanged;
* every changed member is in the patch's allowed closure, and no changed
  member is silently excused as “serializer noise.”

The retained `zip-census.py` helper is suitable for this independent pass: it
has no Litchi import, parses local and central records and descriptors from
the archive bytes, and uses the standard-library reader only for a separate
decoded-payload hash. Its central-record equality excludes only the four-byte
local-header offset, which is expected to move when an earlier changed member
has a different length. The helper's generated-fixture mode allows exactly the
eight selected slide members and rejects all other member, order, or comment
changes. A later control result should bind the helper and its method receipt
by hash rather than silently replacing it with a semantic `ArchiveReader`
comparison.

The raw census must also compare repeated forward outputs. A stable output
digest and stable per-member census make the result deterministic evidence;
they do not prove that the internal planner selected `Copy`. The report must
include this limitation and leave any avoided-decompression or avoided-Deflate
claim unmade until a separate source-level or instruction-level experiment
qualifies it.

The inverse census should first compare logical resource bytes, then compare
the complete raw archive. If exact archive identity fails, include the member
and field differences rather than collapsing them into the semantic pass.

## Correctness and refusal gates

The control is admissible only if all required gates have explicit receipts.
The existing PPTX boundary and opened-transaction tests provide useful
precedent for most of these checks, but a result packet must state which gates
were actually exercised by this corpus and route.

### Source and operation identity

* Bind the source snapshot, lockfiles, corpus generator, source archive length,
  member count, uncompressed payload inventory, and SHA-256.
* Bind the exact slide positions, shape counts, replacement strings, and
  expected changed-resource set.
* Verify the original source buffer is unchanged after every successful and
  refused operation.
* Repeat the changed operation on fresh owning packages and require identical
  logical results, output digest, member census, changed set, and sink summary.

### Publication and reopen

* The changed package must report a non-empty commit and the expected number
  of shape replacements.
* `apply_opened_presentation_commit` must publish atomically and leave the
  source snapshot unchanged.
* Publish through a sequential non-seek sink, reopen the accepted output, and
  verify the full slide graph, text, relationships, media, unknown members,
  and output logical revision.
* Verify the package can publish the same candidate repeatedly without
  changing the retained source or output census.
* Verify the exact no-op route emits the original source bytes and does not
  manufacture a normalized archive.

### Durable forward and inverse

* Serialize and parse the committed patch under the retained finite limits;
  the decoded patch must equal the original.
* Apply the forward patch to a fresh `Package::from_vec` source, then apply its
  inverse to a fresh owning package opened from the published bytes.
* Check resource-byte identity, semantic identity, complete logical revision,
  and full archive identity as separate results using the rules above.
* Apply the forward patch twice or apply it to the restored source and require
  a typed stale/conflict refusal where the patch contract demands it; do not
  interpret a repeated success as harmless last-writer-wins behavior.
* Check the inverse and forward paths before any accepted sink bytes are
  emitted when their preconditions fail.

### Stale, foreign, and physical variants

* Mutate a patched resource in the destination after planning; application
  must refuse and the destination must remain unchanged.
* Mutate an unrelated resource or relationship outside the write set; the
  complete-candidate publication must either follow the documented
  complete-graph behavior or refuse, with no partial assignment.
* Apply the patch to a foreign semantic source and require refusal.
* Include a physical-only source variant with the same logical graph when one
  can be generated without changing the semantic corpus. Record whether the
  current patch contract accepts it because its logical revision matches.
  Do not call that acceptance an exact physical-source authorization: the
  current opened revision deliberately excludes ZIP layout and compression.
* Verify borrowed `from_bytes` ingress is rejected for any control that
  requires physical provenance, using the existing typed provenance refusal
  where that API applies.

### Limits, unsupported content, and sink failures

* Reuse the existing finite `ReadLimits` and opened `Limits` policy. Exercise
  at least one exact-limit and one over-limit case for input, snapshot or
  patch resources, and require the typed limit identity.
* Preserve existing typed refusals for unsupported dependency closure,
  markup-compatibility content, signatures, unknown physical members where the
  operation requires complete physical authorization, and malformed payloads.
  For an ordinary same-package text edit, an unknown member that the writer
  can safely copy should be checked for raw preservation rather than forced
  into a refusal; refusal expectations are operation-specific.
* Exercise the existing partial and zero-byte sink probes. A partial sink must
  report typed incomplete output with the exact accepted prefix; a zero-byte
  sink must fail before accepted output when planning succeeds but emission is
  refused by the sink. Retry publication to a fresh sink and verify the
  package remains usable and the raw census is unchanged.
* Keep sink failure, allocation failure, cancellation, and publication
  refusal outside the successful output oracle. A failed candidate must not
  be presented as a partially valid archive.

The `CountingSink` and `PrefixFailSink` helpers, the PPTX slide-boundary
partial/zero sink checks, and the existing opened-transaction stale/foreign,
unknown, MCE, signed, dependency, and limits tests are suitable starting
points. They do not by themselves prove this media-rich owning control's raw
closure, source retention, or inverse identity and must not be cited as if
they did.

## Admission and reporting rule

The control may be called a qualified owning-input preservation control only
when the source identity, ownership state, same-corpus edit, semantic reopen,
resource-byte inverse, stale/foreign refusal, source immutability, finite-limit,
and sink gates all pass. The raw census must pass for every member outside the
proven changed closure. Full archive inverse identity must be reported as its
own boolean; it is required for a byte-for-byte physical reversibility claim,
but a resource-byte/semantic-only result must carry the narrower wording.

A passing control still proves a capability and an output property, not a
production speedup. It justifies investigating a provenance bridge that keeps
the existing `PackageWriter` preservation machinery at the eager boundary. It
does not justify routing a format facade directly through `SourceBackedPackage`,
skipping complete snapshot validation, changing compression policy, exposing
physical archive types, or claiming that a raw-copy member avoided all decode
or recompression work.

The final 0909 packet should retain the runner, exact source/build/capture
receipts, stdout/stderr, raw census records, semantic and patch receipts, and
cleanup evidence. The result must state whether the control was `qualified`,
`qualified_with_physical_inverse_gap`, or `refused/unqualified`, with the
failed gate and evidence path. No later benchmark or profile may silently
promote the control to a performance claim without its own timing scope and
native validation.
