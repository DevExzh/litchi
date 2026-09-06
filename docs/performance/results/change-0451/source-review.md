# Source review and ownership

The production change is confined to OPC's source-backed package owner. The
new public `PartView::data_and_authorize_precompressed` returns `PartData` and
an existing opaque OPC-authorized transfer token. ZIP types remain private.
No format facade, dependency, unsafe policy, global pool or ambient provider changes.
Accepted ADR tree: `c950b6c8be822561b498d7bbe87c460873dcbf49`, unchanged from
the earlier full read (0002/0010/0011/0024 ownership).

Authorization keeps exact source boundaries, encryption/signature refusals,
leaf relationships, URI/content type, source artifact and revision. Both paths
share limits and final source/context fences. Cold loading calls the ZIP combined
primitive; CRC, sizes and complete Deflate consumption stay ZIP-owned. Format
callers still owe semantic validation before publication (0003/0006).

Before capture, OPC reserves `2*C + 4096` for compressed capture, writer staging
and scratch. The ordinary cache admission reserves decoded `U` and an object.
The combined token privately pins `CachedPayload`, including both reservations,
because its expected-byte `Arc` outlives returned data or the package. The old
expected-byte API keeps caller-owned decoded accounting. No public managed Arc
escape is introduced. Store boundary tests account for `3*U + 4096`, then show
memory/object charges reach zero after publication (0005).

Cold authorization charges `C` work and cache admission charges `U`. A cache hit
or waiter instead charges `U` before fresh capture/verification. Elected loaders
and bypass loads share the existing publication rollback. Ordinary reads still
use their accounting/session dispatch. Deterministic races cover both combined
and ordinary loader directions, allocation sharing and cancellation/source-change
rollback after provisional publication. Freshness takes precedence over
cancellation and ZIP failures at the retained fences.

Do not derive `Clone` for the authorization token to integrate it into reusable
PPTX plans: the existing shared reservation covers one writer staging allocation.
Reusable publication needs explicit writer reservations for each concurrent
publication, alongside source and semantic identity. No one-shot consumption or
mutable shared token is introduced into immutable plans here.

The native test reads an existing POI PPTX image and publishes an OPC transfer
after dropping package/data handles. It proves compressed identity and budget
lifetime for native input; it is not an Office-application roundtrip. Eight
counter cases are deterministic assertions, not latency/allocator samples.
