# 0362ul94sb — Ribbon publication follow-up

This follow-up closes two review findings in the authored XML publication
seam from 0362ul94sa.

The private `try_add_source_part` route remains the cheap source-materialized
insertion path. It validates the incoming name and reserves its destination
maps before revoking exact-source authorization, but does not rescan every
existing payload to recompute an aggregate byte total. Source-backed callers
already admit their payloads, and the batch materialization path retains its
indexed preflight. The public authored add route keeps its full part, part
count, aggregate-byte, XML-event, depth, attribute and allocation checks
before mutation.

Authored and strict XML replacement now resolve the stored canonical
`PackURI` before validation, mutation and provenance-map insertion. A
case-insensitive relationship or caller alias therefore cannot leave the
replacement allocation under a key different from the actual package part.

The OPC regression covers both replacement routes with a stored uppercase
part and a lowercase alias, and confirms that canonical source provenance can
be captured after replacement. No public alias, deprecation or limit escape
was added.

No Cargo, compiler, build, rustdoc, native or performance command was run.
