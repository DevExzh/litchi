# 0362nca — Complete the ODF namespace limit error boundary

Source review of `1710be959a` confirms that 0362nc repairs the direct limit
checks in `ResolvedReader`, but the namespace URI dimension still crosses the
core error boundary identified in 0362nb. The full post-repair test and Clippy
receipts remain valid for their tested source; they do not close this remaining
contract issue. No runtime regression or performance result is claimed here.

`ResolvedReader::push_scope` calls the public
`namespace::normalize_namespace_uri`. Its private `decode_namespace_uri`
constructs the URI limit through `namespace_limit`, which returns a core
`ResourceLimit` carrying the scope string `ODF raw XML namespace URI bytes`.
The reader's `?` then invokes the crate error conversion, and
`RawXmlLimitKind::from_scope` reconstructs `NamespaceUriBytes` from that string.
The public normalization helper itself also returns the core Result.

The repair must carry `RawXmlLimitKind::NamespaceUriBytes` directly through
normalization and its public raw boundary, with explicit conversion only in
legacy consumers that require a core error. Migrate all callers in the same
commit. Verify exact-limit and over-limit lexical/normalized URI behavior,
the returned typed error before conversion, observation/maximum values and
the existing namespace and owner admission semantics. Keep the default
historical envelope and caller-supplied policy intact. The worker repair is
tracked as 0362nd; native qualification remains required.

Follow-up source review of `3f802aec9d` closes this specific error-origin
return. The decoder now constructs a private URI-specific typed error before
either boundary conversion. The raw normalizer maps it directly to the ODF
kind; the legacy resolver explicitly maps it to core. The worker's focused
namespace/reader tests, full common/ODS/ODT suites and Clippy pass as recorded
in 0362nd. Coordinator integration and performance qualification remain open.
