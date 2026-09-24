# DOCX `stylesWithEffects` performance and memory measurement contract

Status: **measurement contract; current attribution timing remains gated**.
The production baseline `8702fd4db` passed the fresh 52-lane correctness smoke
retained in `fa927a8a9`, with independent review. A current profile may start
only after its production source, transitive local inputs, fixture copies,
and harness are frozen in a clean committed checkout. Its prerequisite must
match the approved current smoke's exact committed file set and bytes.
Historical captures remain separate evidence; they do not establish timings
for the current attribution harness.

The profile is an absolute, scenario-scoped observation of the public
`litchi_docx::styles::effects` API. It must not be presented as a native Word
rendering result, a general DOCX throughput claim, or a before/after
improvement until a separately matched control and candidate run exists.

This plan is grounded in [`docs/GOAL.md`](../../../GOAL.md), accepted
[`ADR 0021`](../../../adr/0021-docx-glossary-ownership.md), the approved
[`stylesWithEffects` design](../docx-styles-effects-design.md), the current
public owner facade at
`crates/litchi-docx/src/package/package/styles_with_effects.rs`, the owner
implementation at `crates/litchi-docx/src/styles/effects.rs`, and
`crates/litchi-docx/tests/styles_with_effects.rs`. Their source hashes must be
recorded for each frozen capture; this contract does not approve a moving
source tree.

## Contract under test

The owner is a package resource with two semantic owners:

* `Owner::MainDocument` resolves only from the main document part.
* `Owner::Glossary` resolves only from the validated glossary-document owner.

The public package surface currently under test is:

* `Package::styles_with_effects(owner) -> Result<Snapshot>`;
* `Package::put_styles_with_effects(owner, Resource) -> Result<bool>`;
* `Package::remove_styles_with_effects(owner) -> Result<bool>`;
* `Package::apply_styles_with_effects_patch(owner, &Patch) -> Result<Snapshot>`;
* `Snapshot::{resource, is_empty, owner, conformance, edit}`;
* `Resource::{from_xml, xml_bytes, styles, definitions, projection, conformance}`;
* `Transaction::{replace_resource, clear_resource, commit}`; and
* `Patch::{inverse, apply_to_snapshot, before, after, owner}`.

The first profile must remain resource-level. The typed style projection is
read-only and the current API has no scalar style setter. A future scalar
mutation lane may be added only when a committed public setter exists; the
plan therefore measures scalar **projection reads** separately from whole
resource replacement and does not invent an unsupported mutation API.

The profile must preserve these design decisions from
[`docx-styles-effects-design.md`](../docx-styles-effects-design.md):

* absent owner state is a valid `Snapshot` with `resource() == None`;
* fresh glossary creation retains its existing four-resource seed and does not
  auto-create a fifth effects resource;
* main and glossary effects resources are independent and are never
  deduplicated merely because they contain style definitions;
* source-backed snapshots retain provenance-bearing payload ownership;
* exact no-op publication preserves all package bytes and signatures;
* changed publication validates the owner, content type, relationship closure,
  package-wide two-part limit, and source preconditions atomically; and
* resource removal deletes the target only when the owner relationship is its
  exclusive inbound edge.

No lane may claim a complete visual-effects grammar, style cascade, layout,
rendering, or Office acceptance. The native fixtures currently contain copies
of style definitions and provide topology/source-preservation evidence; the
design scan found no `glow`, `shadow`, `reflection`, `textOutline`,
`textFill`, `scene3d`, `props3d`, or `w14:` token in the effects members.

## Frozen native inputs

The required native files live in the repository's existing `test-data`
submodules. The future runner must copy only these four files into its owned
fixture directory, hash them before staging, and retain the source path, byte
count, and SHA-256 in a replayable manifest. It must not copy a whole submodule
or build from a dirty active workspace.

| Fixture source path | Package bytes | Package SHA-256 | Effects members (uncompressed XML bytes) |
| --- | ---: | --- | --- |
| `test-data/poi/test-data/document/Bug54849.docx` | 27,566 | `f54182713ea5ce5d77b9593d3d9d24e645460043cec0b40ef59c932385f084d3` | `word/stylesWithEffects.xml`: 19,883 / `799de1f7a4ce43f0ca101dc750e8a8bd6e75bcb721f7744d4787dad576cda3b1`; `word/glossary/stylesWithEffects.xml`: 16,138 / `d27f6ced340ffa173b3b861b08a4006e76673a411687b4145f5e46dc6dae13eb` |
| `test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx` | 16,142 | `bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f` | `word/stylesWithEffects.xml`: 15,710 / `00c5cda7671bf545a8c97312f14b2b8bc0ee7fa469b36c25ee158c8a5c1c1568` |
| `test-data/ooxml/docx/ComplexNumberedLists.docx` | 14,458 | `297a085a7d433af2eeee7661e8db21539452cb585096484774a1e9f5f258b0b6` | `word/stylesWithEffects.xml`: 15,955 / `b4bf5d355a45daf0a1085e73fe27041b5db22bfa23820f18ffaa7f9c8cb70f18` |
| `test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx` | 25,741 | `8ccd581d8f0ae102b220228ad26b3974821a7ce8e3ff7df4b78f7da8a0d06ed9` | `word/stylesWithEffects.xml`: 20,117 / `e72df38e71a351ebaaf7b102e8ab7862b7e5eb04cc86e187f8510746b70d3f54`; `word/glossary/stylesWithEffects.xml`: 16,244 / `78112f02ff4e0b94a6c99f8688d86c6fa91d54d6c57c504001453b9c73e3dad1` |

The package-byte counts above are the current committed fixture observations.
The effects-member lengths and hashes above are part of the approved design
and must be checked again by the smoke harness before any timing run. A fixture
receipt must fail closed when a package or member hash, length, relationship
count, content type, or owner-presence expectation differs.

The fixture matrix must assert these owner states:

* `Bug54849.docx`: main and glossary present, with different effects bytes;
* `ms-office-2010-signed.docx`: main present, glossary absent, signature
  present;
* `ComplexNumberedLists.docx`: main present, glossary absent; and
* `testGlossary.docx`: main and glossary present, with independent bytes.

Each native capture must also compare the effects resource against ordinary
`word/styles.xml` and, where present, `word/glossary/styles.xml` to prove that
the two owner resources were not silently merged or deduplicated.

## Correctness smoke before profiling

The smoke harness must call the public API and retain raw receipts. It must
never synthesize a refusal, timing, allocation count, graph count, or source
result. Every expected failure must be observed at the API call that fails.

For every operation that fails, the harness must:

1. catch the typed error before looking at its display text;
2. read the package back after failure and compare complete physical package
   bytes, metadata, relationship XML, content types, owner graph, opaque
   members, and signature state as applicable;
3. assert that no partial output or partial relationship member was emitted;
4. reopen the unchanged bytes through the same public reader and revalidate the
   owner state; and
5. record the typed variant and structured fields in the receipt. A string
   class label is diagnostic only.

The exact-cap refusal lane must match
`DocxError::Opc(OpcError::ReadLimit { resource, actual, maximum })` and must
verify the expected `ReadResource` for the selected cap, `actual > maximum`,
unchanged physical and metadata readback, and no output. A generic
`is_err()` or message substring is not an acceptance gate. The cap matrix must
cover both an existing owner replacement and a source-less owner addition:

* `Parts`;
* `TotalPartBytes`;
* `TotalRelationships`;
* `TotalRelationshipXmlEvents`;
* `TotalRelationshipXmlBytes`;
* `RelationshipParts`; and
* `RelationshipGraphNodes`.

For every cap, run an exact-fit control where the projected value equals the
limit and succeeds, and a one-unit-under-projected refusal where the cap is
`projected value - 1`. Record the source and projected package metrics used to
derive the limit; do not hard-code a refusal number copied from an earlier
run. Select a fixture whose source still opens under the refusal limit. A
replacement that leaves a count unchanged cannot test that count's mutation
boundary; use an addition for that count and report the replacement case as
not applicable, rather than treating an ingress refusal as an edit refusal.
For source-bound edits, also assert that a projected aggregate-cap violation
is rejected by `Transaction::commit` before replacement metadata is
serialized; publication-only rejection does not prove this staging boundary.
If a selected operation also exercises the per-relationship-member
`RelationshipXmlBytes` cap, keep that lane separate and identify the member
whose typed limit failure was observed.

The smoke matrix must include the malformed and topology cases already covered
by `tests/styles_with_effects.rs`: duplicate owner relationships, third/orphan
effects part, external target, wrong content type, outbound effects-part
relationship, shared inbound target, malformed root/namespace/XML, caller XML
event/depth limits, absent owner, missing glossary owner, and signed changed
publication without explicit `unsign`. These are correctness gates and are not
performance samples unless the scenario is also in the named timing matrix.

## Named operation matrix

The future profile uses stable lane names and runs each lane against the
small, scaled, and near-limit resource classes where the operation applies.
Fixture setup, ZIP fixture copying, synthetic XML generation, and caller-owned
replacement construction occur before the timed allocation window.

### Native capture and cheap source paths

* `native_capture_main_<fixture>` and
  `native_capture_glossary_<fixture>` for each present owner. The timed scope
  is package open plus `styles_with_effects(owner)` and the selected typed
  projection read; native fixture decompression and file I/O are measured in a
  separate setup receipt.
* `native_absent_owner_<fixture>` for each absent main/glossary owner. The
  result is an empty snapshot and must not synthesize a part.
* `source_snapshot_noop_main` and `source_snapshot_noop_glossary` on
  `Bug54849.docx`: load the owner snapshot, clone it, call `edit().commit()`
  without a replacement, then pass the resulting empty patch through the
  public exact no-op path (and call `put_styles_with_effects` with the same
  retained resource). Record the shared payload identity in smoke, but use
  bytes and source tokens as the correctness gates because pointer values are
  process-local.
* `projection_scalar_read_<owner>_<scale>`: read `len`, one ID/name lookup,
  one default-style lookup, and the bounded numbering projection from the
  retained typed projection. This is a read-only scalar observation; it must
  not be reported as a mutation or as a full style cascade.

### Resource replacement and reversibility

* `replace_resource_main_<scale>` and `replace_resource_glossary_<scale>`:
  load a present owner, construct a valid same-conformance `Resource` with one
  deterministic inert extension marker changed, stage
  `replace_resource(Some(resource))`, commit, and publish through the public
  package patch path. The marker must be verified semantically and all
  unrelated package members must remain byte-identical.
* `remove_present_<owner>_<scale>`: clear one owner resource, publish, reopen,
  and verify that the effects relationship/target is gone while the other
  owner and unrelated graph remain unchanged.
* `add_absent_<owner>_<scale>`: start from a package with the selected owner
  absent, publish a validated `Resource`, reopen it, and verify the exact
  owner relationship/content-type closure. A missing glossary document must
  refuse before mutation; an existing glossary must not gain a fifth seeded
  resource merely because the effects resource is added.
* The current profile's `inverse_replace_main` and `inverse_remove_main` lanes
  publish a changed patch, construct/apply `patch.inverse()` and serialize its
  output under `inverse_ns`, then start `inverse_reopen_ns` immediately before
  `Package::from_reader` and owner loading. Require an exact package-byte hash
  and owner-resource hash equal to the original. The restored reopen belongs
  to the observed inverse cost, not to setup; that cost is the sum of these
  two disjoint clocks.
* `stale_patch_<owner>` is a smoke-only refusal. Apply the same patch twice or
  apply it to a different owner/source and require a typed source/owner/graph
  precondition error plus complete source readback.
* `signed_noop` and `signed_changed` are smoke lanes. The exact replacement
  must return `false`, preserve the signature, and preserve bytes. A changed
  resource must refuse while signed and succeed only after the explicit
  `unsign` operation; no profile row may imply that changed signed publication
  is implicit.

### Main/glossary independence

Run paired operations on `Bug54849.docx` and `testGlossary.docx`:

* read and replace main while hashing glossary effects XML and graph members;
* read and replace glossary while hashing main effects XML and graph members;
* remove one owner and confirm the other owner's target, bytes, and typed
  snapshot are unchanged; and
* replace both sequentially in fresh processes to prove that the two resource
  payloads are not combined into one size or one allocation observation.

The report must show owner and fixture in every row. An aggregate “effects
resource” row without owner identity is not sufficient evidence.

## Resource scales

The scale generator must be deterministic, source-controlled, and retained by
hash. It must produce a valid `w:styles` root in the package's established
Strict or Transitional conformance, preserve an inert unknown extension child,
and keep the typed projection bounded. It must never claim to exercise
unmodeled visual effects.

Use these classes unless the frozen smoke records a fixture-specific reason to
adjust them:

* **small**: the native effects XML member, approximately 15–20 KiB;
* **scaled**: deterministic valid resources at 64 KiB, 1 MiB, and 8 MiB;
* **near-limit**: a valid resource just below the 32 MiB resource ceiling
  (`MAX_XML_BYTES`), with XML event/depth counts recorded and below their
  corresponding limits.

The generator must record exact XML bytes, event count, maximum depth, typed
style count, opaque-marker byte count, and SHA-256. For every scale, keep the
whole package input hash separate from the effects-member hash. The near-limit
lane must run only after a smoke proves the exact part/XML limits and must not
be treated as a hard OOM or scratch-budget claim.

## Measurement boundaries and phases

Each receipt must distinguish wall time from process memory and Rust allocator
traffic. Use disjoint phase clocks whose sum is no greater than
`elapsed_ns`:

* `capture_ns`: package open and owner snapshot load for capture lanes;
* `snapshot_ns`: owner snapshot/projection capture for mutation lanes when it
  is part of the named operation;
* `stage_ns`: transaction creation, resource replacement/clear, and commit
  preparation;
* `commit_ns`: source/graph validation and commit formation;
* `publish_ns`: package candidate publication and assignment;
* `reopen_ns`: ordinary post-publication package reopen and owner load;
* `inverse_reopen_ns`: package reopen of the restored inverse output, timed from
  before `Package::from_reader`;
* `inverse_ns`: inverse patch validation and publication;
* `projection_ns`: typed style projection or scalar lookup when explicitly
  selected;
* `opaque_ns`: complete unchanged/opaque member checks, including ordinary
  `Package` validation cost;
* `graph_ns`: package/relationship metrics and owner-closure checks;
* `readback_ns`: source readback after a refusal or failed publication; and
* `validation_ns`: semantic result checks not already named above.

If a phase intentionally overlaps another, name that overlap in the receipt
and exclude it from the disjoint sum. Do not hide `opaque_ns` or `graph_ns`
inside an unnamed validation phase. Fixture generation, copying replacement
bytes, and report serialization are outside the operation boundary and have
their own setup/driver receipts.

For exact inverse, the restored package reopen must be timed in
`inverse_reopen_ns`; starting that timer after `from_reader` would make the
inverse lane invalid. For refusals, `readback_ns` and typed-error validation
must be separate phases.

## Memory, source, opaque, and graph evidence

Every measured sample reports these quantities independently:

* `elapsed_ns` for the named operation;
* `direct_allocated_bytes`, `realloc_old_bytes`, `realloc_new_bytes`,
  `requested_alloc_bytes`, allocation/reallocation/deallocation call counts;
* `live_before`, `live_after`, `peak_live_delta`, allocator failure and
  underflow flags, and the checked live-byte equation;
* process RSS from an external `/usr/bin/time -v` sidecar; and
* output bytes and phase timings.

`peak_live_delta` is allocator live memory, not RSS. RSS includes runtime,
allocator arenas, mappings, and page effects. The report must not substitute
one for the other or claim a hard memory cap from either metric.

Per sample and per fixture retain:

* whole input package SHA-256 and byte length;
* each effects XML SHA-256, byte length, conformance, owner, and typed style
  count;
* source/opaque bytes for the selected effects member and every unrelated
  member asserted unchanged;
* package part count, aggregate part bytes, total relationship count,
  relationship-part count, relationship graph nodes, relationship XML bytes,
  and relationship XML event count; and
* allocator receipt, RSS sidecar, process exit status, stderr, and the exact
  command/environment.

For a no-op, changed operation, inverse, or refusal, compare complete package
bytes after reopening. A captured story/resource snapshot is not a substitute
for reading back the published package. For main/glossary rows, compare the
unselected owner separately. For signed rows, include signature members and
signature state in the opaque/physical comparison.

## Fail-closed source and build receipts

The future evidence directory should retain:

* copied native fixtures and a fixture manifest with source paths, lengths, and
  hashes;
* deterministic scale-generator inputs and generated XML hashes;
* a source manifest for the committed production files, current public API and
  tests, `Cargo.toml`/`Cargo.lock`, and every harness source file;
* clean checkout HEAD, tree/status, Rust toolchain, target, flags, Cargo
  metadata before/after, build log, and binary hashes;
* raw JSON receipts and `/usr/bin/time -v` sidecars for every lane/process;
* reports recomputed only from raw receipts; and
* a verifier receipt whose `passed` field is true only when all lanes, samples,
  hashes, phase bounds, allocator equations, typed errors, physical/metadata
  readback, and RSS/process status pass.

The profile settings are three fresh processes, at least two warmups, and at
least twenty measured samples per process, with `CARGO_INCREMENTAL=0`,
`LC_ALL=C`, release `--locked --offline`, and `RUSTFLAGS`,
`CARGO_ENCODED_RUSTFLAGS`, and `RUSTC_BOOTSTRAP` unset. Run control and
candidate sequentially and never during another agent's timed workload.

The verifier must reject negative, boolean, string, fractional, overflowing,
or otherwise impossible numeric receipts; phase values must be unsigned,
disjoint, and bounded by elapsed time; allocator equations must balance; RSS
must be one positive decimal marker; and every expected refusal must match its
typed `ReadResource`. A lane that unexpectedly succeeds, emits partial output,
or returns a generic error fails the run rather than becoming a passing row.

Until these source/build inputs are frozen and the smoke is independently
reviewed, no full timing, allocation, RSS, scaling, or speedup result should be
recorded for this owner.
