# Change 0383: PPTX source-backed cross-slide chart-leaf copy

## Scope

Change 0383 extends the bounded source-backed cross-slide copy operation with
zero or more direct ordinary chart graphic frames alongside the existing
picture set. A chart is admitted only when it is a direct leaf under the one
selected slide shape tree, has exactly one namespace-resolved `r:id` binding,
and that binding resolves to one internal canonical `/ppt/charts/` part with
the destination slide's strict/transitional dialect, chart content type, and
valid chart root. The chart part must be a relationship-free leaf and otherwise
self-contained. ChartEx, embedded workbooks, `externalData`, style/color
parts, chart drawing/user-shape content, and every chart outbound relationship
remain outside this operation.

Distinct source chart parts are copied once even when separate source frames
share a chart. Separate slide bindings are retained. Destination chart part
URIs and relationship IDs are allocated deterministically, and only exact
namespace-resolved `r:id` byte values are rewritten; unrelated XML and
attributes remain byte-preserved. Image handling retains the 0382 restrictions
and deduplicates distinct physical media in the same way.

The operation refuses external, wrong-type, missing, outbound, or unreferenced
bindings; ChartEx or other broader graph closures; malformed, ambiguous,
nested, or misplaced hosts; stray chart-namespace content; MCE, DTD, or PI
content; unresolved or rebound namespaces; and unsupported chart dependencies.
Stale source/destination snapshots, foreign source/destination pairings,
signed packages, read-limit violations, cancellation, and unsupported
relationship-ID collisions fail before publication except for allocator-owned
selected chart/image relationship-ID collisions. Destination anchors and all
other package members, raw members, source snapshots, and signatures retain
their existing preservation and invalidation semantics. Partial sink behavior
is preserved, and no durable inverse is provided.

## Evidence

The focused `source_backed_cross_copy` suite passed `52/52`, including chart
binding, shared-chart physical deduplication, namespace, malformed-input,
freshness, signature, cancellation, partial-sink, and collision coverage. The
isolated typed-cancellation regression passed `1/1`. The default-feature
library gate passed `531` tests with one named filtered test. The all-features
primary library gate passed `533` tests with one named filtered test, and its
integration gates were green with the three existing exact exclusions. The
documentation/doctest evidence was `6` passed and `2` ignored.

Clippy was green with the inherited allowances `clippy::nonminimal_bool`,
`clippy::clone_on_copy`, and `clippy::needless_lifetimes`. The crate-boundary
gate reported `64` packages, `240` internal dependencies, and `14` accepted
debt edges.

Validation used one Cargo process, `CARGO_BUILD_JOBS=1`, a 6 GiB virtual-memory
cap, and a 10 GiB `MemAvailable` admission gate. These controls are
resource-capped/OOM-mitigating execution policy, not proof of OOM prevention.
No performance measurements were taken.

`performance_claim: none`

`claim_authorized: false`

