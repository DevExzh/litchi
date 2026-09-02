# Change 0371: shared ODF content validation

## Scope

Change 0371 moves the proven plain-`Reader` namespace tracker and ODF content
document validator into `litchi-odf-common`. The doc-hidden `core::private`
surface exposes `BindingTracker`, `BindingTrackerError`, and
`ContentDocumentValidator` only to owning family crates. `litchi-odt` now
uses this shared substrate and removes its duplicate tracker and validation
handler. No ordinary CRUD API, package handle, raw type, or runtime handle is
added.

The common tracker uses checked `u32` depth rather than relying on
quick-xml's `NsReader` namespace depth. It preserves deferred scope removal
for empty and end elements, namespace rebinding and unbinding, and the
reserved `xml`/`xmlns` rules. At most 256 namespace declarations are accepted
per element. `litchi-odf-common` now declares its direct `memchr` dependency
for the extracted tracker.

## Validation contract

`ContentDocumentValidator` retains the 256 MiB `content.xml` input bound and
adds a shared maximum nesting depth of 4,096. Opening the 4,097th element
returns a typed invalid-format error instead of entering quick-xml's
`u16`-depth namespace path. Namespace resolution is performed only while the
validator still needs the `office` namespace, avoiding unnecessary deep-path
allocation. Existing ODT content-root, body, version, XML-reference, and
tokenizer checks remain in the shared event stream.

The ODT catalog depth regression now expects the shared error:
`Invalid format: ODT content.xml nesting exceeds maximum depth of 4096`.
Tracker tests cover the 256/257 declaration boundary, reserved bindings,
scope restoration, prefix unbinding, empty-element deferred pop, and depths
above `u16::MAX`. The validator test constructs a balanced 4,097-level input
and verifies rejection without a panic.

## Verification

Exact Rust formatting passed. The locked, offline release tests for
`litchi-odf-common` and `litchi-odt` ran with one Cargo job and one test thread.
All remaining focused tests passed when two pre-existing failures in
unmodified writer code were skipped by exact name:
`encryption_authoring_uses_no_unsafe_code` rejects the word `unsafe` inside an
error string, and `metadata_is_validated_and_bounded_before_member_output`
expects a MIME case to return an error but receives `Ok(())`.

Strict scoped Clippy exposed one pre-existing `large_enum_variant` diagnostic
in unmodified `litchi-odf-common/src/package/model.rs`. The same scoped Clippy
run passed with only that exact lint allowed. The crate-boundary gate passed
for 64 workspace packages and 240 internal dependency declarations.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. This correctness and
security-hardening batch has no clean A/B timing evidence. It makes no
latency, throughput, allocation, RSS, physical-I/O, cold-cache, fixed-memory,
or OOM-prevention claim.
