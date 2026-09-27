# 0780 candidate review — static MCE baseline capabilities

Review status: **the applied/live production source has no reachable semantic or
bounded-resource correctness blocker in this read-only audit.** Retention is
still conditional on the root-owned quality gates and paired public-workflow
evidence required by ADR 0005. This review did not run Cargo, native commands,
or profilers.

The frozen candidate source is
`candidate/model.rs` at SHA
`c5a213ce4ef9f507b24cf4bc0319a6d4192c501b6b3ee3ed6b145863b966c4ba`.
The authoritative applied snapshot and the live worktree source are both SHA
`44169877f7c689b9266583610f6dfb84138b2a1039c7edc22b5d82ce75eed070` and have
the same bytes. The applied snapshot includes the quality-0 test-only fix to
the independent-oracle loop; the frozen `model.rs` and `model.patch` retain
the earlier `for &namespace in OLD_BASELINE_NAMESPACES` spelling, which does
not compile for an array of `&str`. This archive discrepancy is a packet
bookkeeping issue, not a production-logic blocker, but it must be reconciled
before the packet is sealed.

## Semantic verdict

The representation in `candidate/applied-model.rs:266-367` is sound for the
existing API contract:

* `Capabilities::ooxml_baseline()` stores `NamespaceSet::Baseline`, and the
  static table contains the same 17 exact OOXML Transitional/Strict and XML
  namespace spellings as the old constructor (`:293-299`, `:340-367`). The
  membership check is exact and case-sensitive, so empty strings, trailing
  slashes, and near-match URIs remain unsupported.
* `Capabilities::new()` remains empty (`:283-290`). A baseline profile that
  receives an arbitrary namespace materializes all 17 old strings plus that
  namespace into an owned set (`:302-321`). A repeated baseline registration
  remains semantically a no-op, while explicit profiles retain ordinary
  `HashSet<String>` insertion and clone isolation.
* `understands` is the only capability query used by the MCE scope
  (`crates/litchi-ooxml-common/src/mce/scope.rs:149-170`). The extension set,
  `Default`, public methods, MCE limits, MustUnderstand decisions, selected
  branches, and emitted bytes are otherwise unchanged. The codec's borrowed
  no-MCE path remains unchanged (`crates/litchi-ooxml-common/src/mce/codec.rs:743-758`).
* The private field type change has no production direct-field consumers in the
  workspace. `Capabilities` still derives `Clone`; its extension set is
  untouched. Derived `Debug` is observably different for a baseline value
  (`NamespaceSet::Baseline` replaces the old heap-backed set), so diagnostics
  that snapshot `Debug` text should treat this as an intentional representation
  change rather than a stable formatting contract.

The added tests at `candidate/applied-model.rs:498-739` provide a useful
independent differential check. They compare the static profile with a
separately hard-coded old 17-name profile for positive and near-miss URI
membership, `new`/`default`, clone isolation, custom registration, extension
preservation, legacy and streaming MCE output/report/event traces,
AlternateContent selection, MustUnderstand behavior, malformed XML, an
unbound ignorable prefix, and an input-limit refusal. The layout test is at
`:609-623`; the corrected independent oracle is at `:525-530`.

## Layout and resource verdict

The candidate avoids the bool representation's inline owner growth. Its
`NamespaceSet` enum relies on the current `HashSet<String>` pointer niche, and
the focused test asserts both
`size_of::<NamespaceSet>() == size_of::<HashSet<String>>()` and
`size_of::<Capabilities>() == size_of::<HashSet<String>>() +
size_of::<HashSet<Name>>()` (`candidate/applied-model.rs:609-623`). This is a
target/compiler layout guard, not a portable language guarantee. If a
supported target fails either assertion, the owner envelope must be updated or
the representation changed before adoption.

The DOCX settings admission envelope remains conservative without a source
change. `mce_workspace.rs:31-54` lists the 17 baseline names plus the two
settings extensions, and `:388-402` charges their payloads, a 19-entry table,
the extension table, and `size_of::<Capabilities>()`. The settings path adds
exactly those two custom namespaces
(`crates/litchi-docx/src/settings/extensions/package.rs:21-29`). The first
custom registration in the candidate reserves at least 18 entries and
materializes the old baseline; the second settings registration produces the
same 19 logical names that the existing envelope charges. Baseline-only calls
now retain the old fixed charge as an overbound while avoiding the 17 owned
strings and baseline table. No mutable global cache, registry, unsafe code, or
new allocation policy was introduced.

This conclusion is about the existing logical owner envelope. It does not
claim exact allocator capacity, RSS, or typed recovery from global allocator
OOM; ordinary `String` conversion and set insertion behavior remains subject to
the existing allocation policy. Custom profiles with more registrations than
the DOCX settings profile still require their caller-specific limits and
measurements.

## Required bounded action

1. Use `candidate/applied-model.rs` and the matching live source as the
   correctness reference. Before sealing the evidence packet, either refresh
   `candidate/model.rs` and `candidate/model.patch` with the one-line
   test-only oracle fix or record the applied snapshot as the final candidate;
   do not leave two claimed candidate sources with different compile status.
2. Keep the layout assertion tied to every target on which the existing
   `capabilities_owner` formula is used. A failed assertion is a real resource
   accounting blocker, even though the semantic representation remains valid.
3. Gate the optimization on paired public evidence: exercise the default-only
   PPTX capture/commit/publication path (the packet's `commit/large` lane), a
   custom-registration path such as PPTX extension capabilities and the DOCX
   settings two-extension path, and the `capabilities/tiny` constructor
   allocation diagnostic. Require output/readback parity plus allocation and
   process/RSS evidence. The historical 0760 profile's approximate 7.5%
   constructor attribution motivates this probe but is not a 0780 result.

ADR 0005 requires finite owner accounting and measured performance evidence;
ADR 0006 requires compatibility and fail-closed behavior. The applied source
keeps the existing MCE limits and owner envelope under the layout guard, and
the appended differential tests cover the principal semantic routes. No cold
cache, concurrency, complete-CRUD, native Office, or net-throughput claim can
be made from this source review alone.

## Follow-through

The packet bookkeeping point above is resolved: the final live source,
`candidate/applied-model.rs`, and `candidate/applied-model.patch` now describe
the same SHA `44169877f7c689b9266583610f6dfb84138b2a1039c7edc22b5d82ce75eed070`.
`candidate/model.rs` and `candidate/model.patch` remain retained as the
failed-draft history, and `implementation-notes.md` directs reproduction to
the applied patch. The original review text is preserved; this note closes
that artifact discrepancy.


## Coordinator reconciliation

The original draft and failed quality-0 attempt remain historical evidence.
`candidate/applied-model.patch` and `candidate/applied-model.rs` are the final
formatted source with the test-only iterator fix. The implementation notes
explicitly direct reproduction to that applied patch. This resolves the
artifact ambiguity without overwriting the failed draft. All five new tests,
including the layout and independent membership oracle, pass in quality-1.
