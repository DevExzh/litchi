# Change 0382: PPTX source-backed cross-slide image batch

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Scope

Change 0382 extends the source-backed PPTX cross-slide copy closure from one
direct embedded picture to a nonempty caller-bounded set of direct `p:pic`
leaves under exactly one direct `p:spTree`. Each selected picture contains
exactly one direct `p:blipFill/a:blip r:embed` reference to an internal,
relationship-free `/ppt/media/` `image/*` leaf. Distinct media targets are
copied once. Each selected image receives a deterministic destination media
URI and selected relationship-ID allocation/rewrite without XML
normalization.

Semantic picture parsing preserves bounded foreign, non-MCE,
non-relationship `a:blip` attributes opaquely and accepts one valid
unqualified `cstate` token. Namespace-safe copy rewrites only the full-slide
resolved relationship-namespace `r:embed`. A full-slide unbound lexical
`r:embed` is rejected as `UnsupportedRelationship`; `r:link`, unknown
relationship attributes, MCE, and duplicate or ambiguous resolved embeds
remain refused.

The destination anchor permits and preserves other valid existing relationships
while anchoring exactly one dialect-correct internal `slideLayout`. The
full-slide namespace-aware `SourceSlide::images` inventory is an inventory
fence. Source catalog relationship reconstruction is fallible, physical ZIP
media deduplication is asserted, and strict XML end-name and unresolved-prefix
fences remain active.

Unselected XML and package members remain preserved. The existing source and
destination freshness, signature, cancellation, partial-sink, and resource
fences remain in force. No image decoding, conversion, or rendering is
performed, and no durable inverse or patch wire is added.

## Refusal boundary

The selected image relationship IDs may be allocated and rewritten to avoid a
destination collision. Layout, non-selected, and unsupported relationship-ID
collisions still refuse. Planning also refuses broader dependency graphs,
MCE, malformed, ambiguous, or misplaced blips, external or linked targets,
missing or mistyped media, media with outbound relationships, unreferenced
image relationships, and unsupported topology. Duplicate shape trees and
shared or otherwise unsupported relationships fail closed before output.

Wrong-type non-selected and non-anchor slide bindings are rejected at open.
Planning still validates every binding defense-in-depth. The evidence does not
claim a planner test for malformed objects.

The accepted topology remains exactly one direct `p:spTree` and a nonempty
caller-bounded set of direct picture leaves. Selected media bytes and content
types remain inert opaque payloads; the operation does not decode, convert,
render, or semantically normalize image or slide XML.

## Resource-accounting boundary

The existing resource fences remain applied to the complete caller-bounded
selection and each distinct media payload. Declared sizes are checked before
payload reads, checked candidate accounting remains in force, and actual
payload sizes are verified after reads. These are bounded refusal and
accounting conditions, not fixed-memory or system-level OOM evidence.

## Validation

- Focused `source_backed_cross_copy`: `41/41` passed.
- Default-feature library: `531` passed with the exact pre-existing
  `stale_and_unsupported_raw_xml_fail_before_publication` test excluded.
- All-features library: `533/533` passed; all integration binaries passed
  with the exact exclusions
  `stale_and_unsupported_raw_xml_fail_before_publication`,
  `malformed_presentation_children_are_reported_by_their_owner`, and
  `noncanonical_style_target_survives_transactional_raw_save`.
- Doctests: `6` passed and `2` ignored.
- Strict Clippy passed with warnings denied and only the existing allowances
  `clippy::nonminimal_bool`, `clippy::clone_on_copy`, and
  `clippy::needless_lifetimes`.
- The crate-boundary gate passed for 64 workspace packages and 240 internal
  dependency declarations with 14 existing debt entries.

Validation used one Cargo invocation at a time with `CARGO_BUILD_JOBS=1`,
incremental and debug build state disabled, one dedicated target, serial test
threads, a 6 GiB per-process virtual-memory cap, and a `>=10 GiB`
available-memory launch threshold. These are OOM-mitigating,
resource-capped controls, not proof of OOM prevention.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. This evidence is limited
to the nonempty caller-bounded direct-picture source-backed copy topology,
distinct-media deduplication, selected relationship-ID rewrite,
preservation, freshness, signature, cancellation, partial-sink, resource,
and typed-refusal invariants exercised above. No measured hotspot, latency,
allocation-volume, RSS, physical-I/O, cold-cache, throughput, scaling,
fixed-memory, image-processing, broad media-rich, real-producer, durable
inverse, or system-level OOM-prevention claim follows.
