# Next memory and lifecycle work

0422 supplies a separate callback-order region peak with serialized boundaries.
It does not establish retained ownership or bound aggregate document memory.

The next measurement should retain explicit snapshots after planning, commit,
publication, and each object drop. Keep the existing lifecycle timer/region
boundary stable; use a distinct diagnostic selector if destruction is moved
inside the measured interval. The first plan's patch and the application
candidate were both alive at the old finish boundary, so live-after was not a
post-drop footprint. Snapshot probes should retain fixed-size observations
without allocating inside the measured phase. Do not nest allocator regions;
that remains an unavailable observation by design.

Use these boundaries in admitted near-output-limit media cases with exact and
one-short resource limits, preserving typed preflight/refusal behavior. Label
exit-minus-entry live bytes as signed net process change. Tie any retained-byte
accounting to concrete owners and release points; neither process RSS nor the
new region peak proves an ownership-specific budget by itself.

Then establish a matched source-backed media-rich PPTX lifecycle using the owned
corpus's exact archives, slide positions and collision pattern. Include source
and destination opening in both lifecycle selectors. The source-backed API
already admits eight 2 MiB relationship-free image leaves. Verify nine added
OPC parts and ten ZIP members, remapped names/relationships/content types,
image bytes, copied slide semantics, deterministic per-writer output and
untouched destination raw records. Existing plain source-backed phase timing
is not a matched lifecycle comparison.

Retain refusal gates for unknown non-Part members, encryption, signatures,
macros/protection, external or non-leaf media relationships, stale source,
layout/dialect mismatch, malformed sizes, cancellation and resource limits.
Broader source-backed OPC/XLSX workflows, cold/range access, native-producer
breadth, managed-cache contention and bounded scaling still require evidence.

A V3 diagnostic regression policy also needs matched corrected captures for its
chosen control/candidate roles and scenario contract. Do not silently upgrade
historical exact-tool policies or use allocator elapsed time as a regression
claim while adopting the new metric.
