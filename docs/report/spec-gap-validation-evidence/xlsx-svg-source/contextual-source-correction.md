# Contextual SVG source correction

Status: shared contextual codec and XLSX source-host integration implemented
and independently reviewed; allocation verification remains pending. Scoped
evidence is in `contextual-codec-verification.json` and
`contextual-host-verification.json`. This note records the
correction criteria for the retained-source amplification reproduced in
`retained-scope-probe.rs`; it does not claim lifecycle approval.

The recorded 149,433-byte input produces 32 owned SVG values whose standalone
sources total 4,205,334 bytes. That count excludes retained namespace models.
The host already shares its namespace context, but each shared-codec read first
materializes a complete fragment containing that context and retains another
owned copy. Sharing the host context alone therefore does not solve the issue.

The correction must retain raw host ranges as the exact byte authority and
share immutable inherited bindings across values. Contextual parsing must avoid
both retained and transient per-owner copies of the entire inherited scope.
Namespace completion belongs at standalone export, with the output length
checked before construction or bounded streaming to the caller's sink.

Opaque content may refer to prefixes in attribute values or text. Removing
bindings merely because they do not occur in element or attribute names is not
a valid optimization. Local declarations, shadowing, default undeclarations,
and opaque children must preserve their namespace meaning after scalar edits.

Existing standalone reads must retain their exact-source behavior. A contextual
value must not present an incomplete host fragment as standalone-ready source;
it needs an explicit contextual representation or no standalone source fast
path. Export and readback must remain valid independently of the host document.

Verification must cover shared-context ownership and measured allocation growth
as picture count and inherited scope grow independently. Summing `source()`
lengths alone is insufficient if the representation changes. Include scalar
edits, opaque QName-valued content, namespace shadowing, exact host preservation,
standalone export/readback, and refusal before output allocation at a small
caller cap. Cross-host shared-codec checks are required before integration is
approved. No speedup or peak-memory result is established by this note.
