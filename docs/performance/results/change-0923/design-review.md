# 0923 — DOCX source fingerprint reuse design review

## Disposition

The opportunity is admissible only through an explicit owned-byte ingress. It
is not admissible as a generic `ReadAt` optimization. This review records the
boundary and approval conditions for a later candidate; it does not approve a
production implementation or a performance claim.

## Source and freshness findings

`litchi-opc::SourceArtifact::fingerprint` currently checks the captured source
version, reserves its managed 64 KiB scratch buffer, reads the complete source,
checks freshness after each read and at the end, charges managed `Work`, and
honors cancellation through the existing read/context path. `SourceSnapshot`
stores `Arc<dyn ReadAt>`, so `SourceBackedPackage::from_vec` and
`from_vec_with_limits` currently erase their `OwnedSource` before the snapshot
is retained.

`SourceVersion` and `SourceLineage` cannot prove immutable bytes. The existing
DOCX test `section_layout_commit_rechecks_complete_artifact_before_target_reuse`
(`crates/litchi-docx/tests/source_backed_sections.rs:824`) mutates an unused
payload while a custom adapter continues to report `(8085, 0)`; publication
must return `SectionLayoutFingerprintConflict`. A generic memo keyed by a
version, lineage, length, or main-part bytes would bypass that proof and is
therefore rejected. `from_read_at(Arc::new(OwnedSource))` remains generic after
type erasure and must retain the full scan behavior.

## Bounded route

Refactor only the owned constructors through a private route that retains the
exact immutable `Arc<Vec<u8>>` alongside the trait object. Carry a private
owned-source state through cloned `SourceSnapshot`/`SourceArtifact` handles;
the state must share one small digest memo for that exact allocation. Generic
constructors set the capability to absent. `from_reader` inherits the owned
route because it already moves its bounded input into `from_vec`.

The first fingerprint request must use the existing complete scan and ordering,
then store only a successful digest. A later request may call
`ensure_current` and return that digest for the same owned state. It must not
eagerly hash during open, memoize failures, or change `write_to_stream`, exact
byte authorization, publication hashes, or inverse checks. A new source or
publication receives new state; no memo is inherited by a different package.

The memo is a fixed-size derived value over bytes the owner already holds. Its
implementation must state and reserve its bounded cost through the accepted
fallible reservation rule; an opaque non-fallible `Arc` allocation is not
evidence of that requirement. The owner-held strong reference is load-bearing:
it prevents an allocation identity from being recycled and proves that the
memo still names the bytes retained by this package.

## Managed and generic exclusions

All existing `from_read_at_*_execution_context` paths remain memo-disabled.
Their fingerprint reads reserve `Resource::Memory`, charge `Resource::Work`
and input observations, and expose cancellation and source-change checks. A
memo hit that silently removed those observations could change a budget
refusal or cancellation result. A future managed owned constructor must either
retain this full path or receive a separate accepted accounting contract; it
must not select the unmanaged memo implicitly.

ADR 0005's 2026-09-16 amendment and ADR 0032 permit a small `f(bytes)` memo
only when the implementation proves owner-held allocation identity, strong
retention, own-allocation projection, bounded/fallible state, successful-value
storage, debug re-derivation, and semantic invisibility of hits and misses.
Their opened-presentation wording is a permitted shape, not blanket authority
for arbitrary DOCX `ReadAt` adapters or a generic fingerprint cache.

## Required controls before implementation review

The paired checks must cover:

- first and repeated fingerprints, fresh `SourceArtifact` handles, and clones
  sharing one owned memo;
- equal bytes in a different allocation and a custom generic adapter remaining
  uncached;
- the existing stable-version changed-payload refusal, revision changes,
  short reads, malformed source errors, and the established check/error order;
- managed work, memory, input, and cancellation behavior with no memo route;
- exact no-op and changed publication bytes, inverse authorization, and output
  failure atomicity; and
- an owned-ingress baseline/candidate pair plus the existing
  `InstrumentedSource` generic control. Results must be labeled as an owned
  route comparison and must not claim generic fingerprint improvement.

Until those controls and the bounded reservation design are present, the
generic route has no admissible optimization and should continue recomputing
the complete fingerprint.
