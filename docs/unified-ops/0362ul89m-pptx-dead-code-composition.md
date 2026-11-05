# 0362ul89m — combined PPTX dead-code return

This private source composition follows the actual isolated compiler failure
on `237d6801a641b85b4477c432a26ac53e531987ed`: 113 dead-code errors,
with the prior type/borrow errors absent. Record 0362ul95 and the individual
89i/89j/89k/89l records retain the exact failed capture and artifact hashes.

All four scoped returns passed independent source review:

- Opened/package/source: `25fa1830445a34c41c5ca71b76e9542702501729`.
  Review returned two helpers that were still unused under cfg(test);
  `stage_shape_property_chunks` and `compaction_scene` were then removed.
- Shape/tag/master/slide: `a4b2b30c7cc3b84686cf5c3bebfdbf0e747b0dd9`.
- Content/comments/media/notes: `c551c4e266a90b670c17774b594428b5d643d4b1`,
  with documentation-only follow-up `2fb0defad219c13a1a0e8f066801c8715c958db2`.
- Presentation metadata: `aaa93fdca174b2125903832943a0bca7de53d42e`,
  with documentation-only follow-up `f3805b5e0f88c8be0d34c99c1572d02836e6fb19`.

Deleted private helpers have no current consumers. Remaining cfg(test)
helpers have actual test callers; their imports, reexports and private
collection fields follow the same condition. Production selected-limit and
semantic-limit implementations remain intact, with no public capability
removed and no warning suppression, compatibility alias or default fallback
introduced. Production custom-show and section routes retain their separate
bounded writers and opaque-content handling.

Coordinator composition applies the reviewed patches without overlap and
formats all 39 changed Rust paths. Direct rustfmt and `git diff --check`
pass. A fresh isolated all-features/all-targets compiler run is still
required; tests, Clippy, glossary, golden transcripts and performance remain
open. Source review does not reclassify the retained failed captures.
