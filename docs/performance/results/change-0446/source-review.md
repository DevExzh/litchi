# Production review

The sole production delta is in `ContentTypeMap`'s Override branch:
`PackURI::new(&partname)` becomes `PackURI::new(partname)`.
`required_attributes` returns two owned Strings. The Part-name String has no
subsequent use. `PackURI::new<S: Into<String>>` first executes `uri.into()`;
borrowing clones the string, whereas passing ownership retains its allocation.
The constructor validates the same bytes and returns the same typed result.
Content-type validation still follows URI validation. The map still performs
case-insensitive duplicate detection and owns its keys and values.

Raw and normalized attribute bounds, XML decoding/normalization, malformed URI,
unknown attribute, duplicate/equivalent override and content-type checks are
unchanged. Error formatting reads the same string. The moved allocation can
have a different capacity; allocated-byte and peak observations therefore need
measurement rather than assuming every memory metric improves. This is safe
Rust ownership transfer, with no unsafe code or new public API/dependency.

Three required manifest parses remain in the lifecycle: source-backed opening,
source manifest readback after freshness checks during publication, and generated
content-type validation. Existing managed-memory reservations and their lifetime
remain. Caching source manifest bytes or skipping a parse is outside this change.
ADR ownership, preservation, bounds, validation and publication requirements
(0002/0003/0005/0006/0008/0023/0024) are unaffected. The accepted ADR tree remains
c950b6c8be822561b498d7bbe87c460873dcbf49.

The unchanged harness uses the plain-source selector for both builds. Timed
source catalog open, plan construction, sequential publication and catalog drop
are covered; prepared input/payload, source wrapper, sink setup, hash finalization,
report oracles and endpoint probes are outside the timer. Independent archive and
report oracles retain exact fixture/output identity checks. No semantic Office
owner or native compatibility result follows from synthetic binary OPC Parts.
