# Change 0341: PPTX source-backed slide-order move

Status: implemented/accepted

`performance_claim: none`

This change provides a deliberately narrow source-backed PPTX slide-order
mutation seam. It is a structural operation on an existing presentation
source, not a general slide-editing or package-copying facility.

## Locked scope

- Move exactly one existing slide to a different position in the presentation
  slide list.
- Operate on exactly one canonical, non-MCE direct root slide list. MCE-wrapped
  or alternate list representations are outside this operation.
- Identify the slide through an opaque public source snapshot and preserve the
  slide's stable source identity, slide ID, and relationship ID (`rId`). Raw
  XML IDs are not exposed by the public API.
- Keep the operation source-backed and lossless for members it does not edit.
- Permit only a root-only overlay for the presentation root relationship/list
  structure. No unrelated part is materialized or rewritten.
- Preserve unchanged members by copying their original raw bytes.
- Perform no slide, media, layout, theme, notes, chart, or other semantic
  payload reads as part of locating or applying the move.

## Invariants

- The move changes only the direct slide-list order and the minimum root
  relationship/list overlay required to represent that order.
- Stable slide IDs and relationship IDs remain associated with their original
  slides; the move does not renumber, regenerate, or retarget them.
- Custom shows are stable `r:id` slide selections. They are preserved and
  remain valid because the move does not rewrite those IDs.
- Mere presence of `viewProps` or `presProps` is allowed. The strict existing
  codecs inspect those parts; a nonempty outline-slide collection or numeric
  `sldRg` range is refused because it can own positional semantics.
- The operation remains panic-free and returns typed errors for invalid input,
  stale sources, cancellation, limits, encryption, signatures, and unsupported
  topology.
- Canonical-name, package-boundary, relationship-target, and encryption guards
  are enforced before an overlay can be committed.
- A patch and its inverse are process-local and are valid only for the exact
  source lineage, source version, and raw presentation-root bytes from which
  they were created. They are not durable or portable patch artifacts.
- If the source is signed, an exact no-op is allowed. An effective signed move
  is refused with the typed signed-mutation refusal; no signature is silently
  invalidated or regenerated.
- Source freshness and cancellation are checked at the operation boundaries
  and before committing the overlay.

## Explicit refusals

The operation returns a typed refusal, without a partial overlay, when
positional semantics may be owned or affected by any of the following:

- presentation sections;
- `modifyVerifier` or other encryption-related mutation surfaces;
- MCE/alternate-content slide-list topology;
- an unknown positional root surface;
- a nonempty outline-slide collection in `viewProps` or `presProps`;
- a numeric `sldRg` range in `viewProps` or `presProps`;
- zero, multiple, noncanonical, or non-direct root slide lists;
- missing, duplicate, ambiguous, or non-source-backed slide identities;
- a target outside the permitted package relationship boundary;
- an effective move requested against an encrypted or signed source.

The refusal set is conservative: if the implementation cannot prove that the
single direct slide list is the sole positional owner, it refuses the move.

## Deferred from this change

- slide add, remove, duplicate, or cross-presentation copy;
- durable or serialized patch storage;
- MCE-aware mutation;
- synchronization or mutation of sections, custom shows, view/presentation
  properties, or any other positional metadata;
- slide, media, layout, theme, notes, chart, or other semantic model edits;
- benchmark or throughput claims.

## OOM-safe validation

Validation was strictly serialized with one isolated external Cargo target and
`CARGO_BUILD_JOBS=1`; no concurrent rebuilds or test matrices were used.

`cargo fmt -p litchi-pptx` passed.

```text
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0341-target CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 cargo test -p litchi-pptx --test source_backed_edit source_slide_order -- --test-threads=1
```

Passed: 4/4 tests, 17 filtered.

```text
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0341-target CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 cargo test -p litchi-pptx --lib opened_slide_name_index_preserves_identity_after_reorder -- --test-threads=1
```

Passed: 1/1 test, 528 filtered.

```text
CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0341-target CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 cargo test -p litchi-pptx --lib composes_order_shape_and_notes_in_one_durable_inverse_commit -- --test-threads=1
```

Passed: 1/1 test, 528 filtered.

The exact external target was deleted after validation. A scoped diff check
passed. Final source reviews found no P0 or P1 findings.

Post-cleanup observations were 136 GiB available on the root filesystem,
53 MiB used in `/dev/shm`, 9.7 GiB available in unrelated `/tmp`, and 18 GiB
available memory with swap nearly full. The worktree audit found only the
protected main worktree, and `.cargo-targets` was empty after cleanup.

## Locality evidence

No benchmark, memory, latency, or throughput claim is made in this change.
The evidence is correctness-only: the inventory and mutation inspect only
presentation-root metadata, perform no slide/media semantic reads, leave
unchanged members raw-copied, and allocate only bounded process-local patch /
inverse state.
