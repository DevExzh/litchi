# 0362r — DOCX source-backed focused edit surface

This slice is based on `1c7eb1e84c`, which already contains the primary replay
owner return-shape migration. It closes the two remaining public tail-append
edit aliases without changing the source-backed publication algorithm.

## API boundary

- `source_backed::paragraph_stream::Edit` is now the actual public focused
  edit type. The stream engine's `ParagraphStreamEdit` is crate-private and
  is only constructed by the package/source-backed implementation. The public
  `TailAppendStreamEdit` re-export and the intermediate public alias were
  removed.
- The one-paragraph family already owns its focused `tail_append::Edit`; the
  `TailAppendEdit` re-export was removed. The root source-backed module no
  longer exposes either prefixed edit type.
- Source-backed limits are selected at package edit creation, or through the
  focused `tail_append::EditOptions::new(Limits)` passed to
  `Edit::with_options`. Its optional policy accessor is
  `EditOptions::limits_value()`, following the API naming rule. The public
  `Edit::with_limits` setters are gone. The
  stream family takes its required `ParagraphStreamLimits` at package entry;
  its public focused edit has no post-staging limit replacement. This keeps a
  caller policy from being silently replaced after source admission.
- The options type is named `EditOptions`; the previous `Options` spelling is
  not retained as an API alias. Existing cancellation options and the package
  execution context remain unchanged.

All DOCX unit, integration, source-read-policy, and fuzz consumers were
migrated in this commit. The public low-level stream traits, replay stores,
proofs, plans, patches, and bounded publication behavior remain available at
their canonical raw modules. The raw `variables`, `paragraph_copy`, and
`story_hyperlinks` record models are deliberately not reclassified as owner
packages here: their remaining raw XML/patch ingress and snapshot publication
rows require separate policy and owner-boundary work.

## Validation

The focused command, run with the capped W8 environment and two build/test
workers, was:

```text
source (deleted-coord-artifact) && \
  CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2 \
  cargo test -p litchi-docx \
    --test source_backed_tail_append \
    --test source_backed_tail_append_stream \
    --test source_read_policy
```

It passed 40 tail-append tests, 34 stream tests (including the external
canonical-type assertion), and one source-read-policy test. The raw log is
outside the repository at
`(deleted-coord-artifact)/focused-edit-tests.log` (6,903 bytes,
SHA-256 `32a0598615ce6a472b1d228c4a8b75ee7aaccbe93319000e9fe8640706b55f87`).
The lower stream engine is not part of the public rustdoc surface after this
change; the focused public path is `source_backed::paragraph_stream::Edit`.

`cargo doc -p litchi-docx --no-deps` also passed. Its log is outside the
repository at `(deleted-coord-artifact)/rustdoc.log` (1,338 bytes, SHA-256
`63affd95cc4236ee8e0689acddb3a82e93b93cb57d8af0eb06e50907b4b2abb7`). The
current `source_backed/paragraph_stream/index.html`,
`source_backed/index.html`, and `source_backed/struct.Package.html` pages show
the focused `paragraph_stream::Edit` path and no public
`TailAppendEdit`, `TailAppendStreamEdit`, or `ParagraphStreamEdit` item.
The all-target crate Clippy gate also passed with `-D warnings`; its external
log is `(deleted-coord-artifact)/clippy.log` (3,282 bytes, SHA-256
`d8bb1477efc57516af3c79e7ad47f020fa127c13d336b8adc430d0eefdd669db`).

No benchmark binary or large evidence file is part of this record. This is a
surface and policy-shape cleanup; the existing bounded source scan, deferred
write, inverse, cancellation, and source-freshness paths are unchanged.
