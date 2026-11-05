# 0362ul61 — PPTX raw shape and motion limits

This leaf carries the PPTX shape, transition, and laser raw-boundary
migration from integration base `a942320908`.

## Public boundary

The borrowed and owned shape readers now expose only `read(..., Limits)`;
the shape-tag reader also requires the selected shape policy, and the old
`read_with` and default twins are gone.  Transition `read`, `write`, and
`write_to` likewise require the transition policy.  Transition output has its
own `output_bytes` budget, separate from opaque-reader retention.

Laser raw operations require a `Limits` policy.  `Usage` is a separate mutable
aggregate ledger passed through slide walks, preserving total input, trace,
and point accounting across slides without mutating the policy.  Failed reads
stage their charges and leave the caller's ledger unchanged.  Laser and
transition fragment writers first count the exact data-dependent output,
preflight the selected bound, make one fallible exact reservation, and emit
through a bounded sink that rolls an existing destination buffer back on
failure.  This covers variable-width coordinates, times, escaped preset
names, and retained raw transition XML; no payload-dependent estimate is used.

`OpenOptions` retains an optional semantic profile for shape, transition, and
laser codecs.  Owning packages, borrowed presentations, opened snapshots and
edits, source-backed snapshots and edits, slide/master/layout projections,
cross-copy/replay candidates, and package-level shape-owned consumers reuse
the selected profile on every read, write, admission, and read-back path.
Default owners carry no profile override.  The package editor intersects the
retained `OpcPackage::read_limits()` policy with the laser safety ceilings
once, then threads that profile through input admission, dialect scanning,
fragment writing, complete-slide size preflight, and read-back.  The
insertion path computes the complete projected slide size before allocating
its output buffer; it does not fall back to a codec default when an owner was
opened under a tighter profile.

## 0362ul61e replay correction

The follow-up closes the recipient-policy gaps found after the first review.
Part patches capture their before/after payload totals during the existing
commit walks and retain the changed-write-set delta, so destination admission
uses the recipient owner's retained aggregate without a proportional scan of
unchanged parts. Exact no-op patches also retain the complete visible-content
identity of their recorded base. A same-owner replay is O(1); an independent
recipient uses an already cached identity when one exists, or performs one
deferred content proof and caches the result. Thus a cold no-op commit itself
does not fingerprint the package, while equal independent opens still pass and
an unrelated recipient returns `PatchConflict` before duplicate publication.
Source-backed catalogs retain ordinary-part, physical-archive,
and relationship XML byte/event aggregates from the existing central-directory
and relationship admission passes; repeated relationship-owner visits are
collapsed to one physical `.rels` metric. Eager materialization transfers the
ordinary-part proof to the owner, and native builder publication refreshes the
same aggregate plus its lazy identity/memo cells after in-place graph writes.
Shape-owned classification,
change-tracking,
master-layout, and placeholder replays validate changed PresentationML
payloads under the recipient's selected shape profile before publication.

Source-backed slide patches now match the recorded content and closure rather
than process-local source lineage or version tokens. A changed replay always
adopts the recipient execution context, output bound, semantic profile, and
reservation ownership; batch and SVG/media replays retain unchanged recipient
slides without re-admitting their opaque XML. Generic source snapshots retain
the selected laser policy but do not parse laser XML until a laser admission or
replay boundary needs usage. That usage is cached per exact snapshot and
policy, so repeated same-policy replay does not rescan the slide. Every
selected slide in a batch shares one laser `Usage` ledger. Changed slides are
re-admitted under the recipient policy; unchanged slides contribute their
captured source usage without a second XML scan. Thus the selected batch scope
preserves aggregate byte, trace, and point ceilings rather than silently
charging only changed slides. A package-wide caller obtains package-wide
accounting by passing every slide through the same ledger.

All PPTX source, package, edit, unit-test, integration-test, and example
consumers were migrated to the canonical signatures.  The minimal and full
PPTX glossary profiles register the required shape, transition, laser, and
package-loader signatures.

## Source checks

The source-only checks for this leaf are JSON parsing and static API/call-site
audits.  Cargo, rustdoc, native tools, and performance capture were not run in
this wave; the parent coordinator must perform independent compilation and
validation after composing the neighboring raw-content and raw-properties
leaves.

Focused source tests cover exact and one-byte-under output boundaries, failure
atomicity, aggregate laser usage across multiple slides, and retained-owner
shape, transition, and laser profiles that refuse work under a custom bound.
The main performance consideration is the extra linear counting pass in the
laser and transition writers; it replaces unsafe estimates and payload-sized
reallocations, while the default owner still stores only an optional profile
bundle and direct appends remain bounded.
