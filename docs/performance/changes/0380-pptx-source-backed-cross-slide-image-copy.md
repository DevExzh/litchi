# Change 0380: PPTX source-backed cross-slide image copy

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Scope

Change 0380 extends the source-backed cross-presentation slide-copy closure
from a dependency-free slide to a slide with exactly one direct embedded image
leaf. The accepted slide contains exactly one direct `p:spTree`, exactly one
direct `p:pic`, and exactly one direct `p:blipFill/a:blip r:embed`. The embed
relationship must target one internal, relationship-free `/ppt/media/`
`image/*` part. The source slide retains its single supported layout
relationship and the previously bounded layout/master/theme equivalence
contract.

`plan_cross_slide_copy` captures the image bytes, content type, URI,
relationship identity, source/destination lineage, topology, and freshness.
`publish_cross_slide_copy_to_stream` copies the opaque payload and allocates a
deterministic collision-free destination media URI, including when the source
URI is already occupied. Unrelated destination members remain preserved. No
image decoding, conversion, rendering, or semantic rewriting is performed.

This tranche does not add a durable or inverse patch representation.

## Refusal boundary

Planning refuses an image/layout relationship-ID collision rather than
remapping the relationship ID. It also refuses missing or external image
targets, non-image content types, image parts with outbound relationships,
multiple or unreferenced images, extra/shared/external relationships,
duplicate or misplaced shape trees, misplaced or ambiguous blip branches,
MCE, charts, diagrams, tables, notes, comments, and arbitrary or media-rich
dependency graphs.

The existing dialect, layout/master/theme equivalence, physical-member,
signature, stale source/destination, foreign-editor, same-source,
cancellation, and partial-sink classifications remain in force. Unsupported
topology fails before output. A sink failure during publication may report
classified partial output under the existing streaming contract; candidate
construction and validation still occur before publication begins.

## Resource-accounting boundary

The planner reads the media part's declared uncompressed size before any
image-payload read and includes the complete staged candidate in a checked
managed-memory reservation. It verifies the actual payload size after the
read. The focused range test proves that an insufficient reservation is
refused before a source read overlaps the image payload. This is bounded
resource-accounting evidence, not a successful-path physical-read bound or a
system-level OOM-prevention result.

## Validation

- Focused `source_backed_cross_copy`: `22/22` passed.
- Default-feature library: `531` passed with the exact pre-existing
  `stale_and_unsupported_raw_xml_fail_before_publication` test excluded.
- All-features library: `533/533` passed; all integration binaries passed with
  the exact audited exclusions
  `stale_and_unsupported_raw_xml_fail_before_publication`,
  `malformed_presentation_children_are_reported_by_their_owner`, and
  `noncanonical_style_target_survives_transactional_raw_save`.
- Doctests: `6` passed and `2` ignored.
- Strict Clippy passed with only the unrelated pre-existing allowances
  `clippy::nonminimal_bool`, `clippy::clone_on_copy`, and
  `clippy::needless_lifetimes`; none occurs in the Change 0380 files.
- The crate-boundary gate passed for 64 workspace packages and 240 internal
  dependency declarations with 14 existing debt entries.
- Independent topology/API, resource/freshness, and test/compile reviewers
  accepted the bounded closure.

Validation ran one Cargo process and one test invocation at a time with
`CARGO_BUILD_JOBS=1`, incremental and dev/test debug build state disabled,
serial test threads, one dedicated target, a 6 GiB per-process virtual-memory
cap, and a `>=10 GiB` available-memory launch threshold. These are
OOM-mitigating, resource-capped controls, not evidence that OOM is prevented.

## Remaining gaps

No latency, allocation, RSS, physical-I/O, cold-cache, throughput,
fixed-memory, image-processing, broad media-rich, real-producer, or
system-level OOM-prevention claim is accepted. Multiple images, linked or
external images, relationship-ID remapping, shared media, media with outbound
relationships, and general slide dependency closure remain outside this
surface.
