# Non-simple OLE Property Set validation

The source-backed non-simple owner exposes lazy CONTENTS inspection, typed
indirect stream/storage references, bounded physical-element edits, and
source-checked commits with exact inverse patches. Unknown streams and storage
metadata are retained. Changed publication validates the property identifier,
`propN` name, kind and physical-reference closure; known indirect VARIANT tags
also undergo this validation when supplied as opaque values. Generic Property
Set parsing continues to leave non-simple indirect variants opaque.

Admission validates source/catalog/path limits before retention or mutation.
The shared binary serializer checks exact encoded lengths and wire `u32`
representability before materializing output. CONTENTS validation and bounded
serialization precede candidate cloning. The CFB output planner admits exact
layout size before copying payloads; a capped output sink remains active.
The shared opaque `DirectoryNameKey` keeps indexed name lookup consistent with
CFB UTF-16 length and simple-uppercase equality, hashing and ordering.

Root Rust 1.95.0 gates passed 568 unit/integration tests across litchi-cfb and
litchi-ole-common (324 and 244 respectively), plus fourteen doctests. One native
corpus integration test and one doctest remain ignored. Strict all-target,
all-feature Clippy, formatting and whitespace checks passed. The receipt binds
nineteen source/documentation files and compressed/raw logs. Independent review
checked the local specification, reference closure, serializer bounds and CFB
layout; focused tests additionally cover 0/4095/4096-byte physical streams,
reopened sizes/payloads, exact output caps and cap-minus-one atomic refusal.

Evidence is synthetic. This owner does not perform filesystem alternate-stream
lookup, external activation, rendering or execution. Unsupported directory kinds
and root-name rewrites refuse rather than dropping producer data. No native
PropertyBag producer acceptance or measured performance claim is made.
