# Native Keynote drawable-comment regression fixtures

These ten files were generated through the focused `litchi-keynote` API and
verified in Keynote 14.4 on 2026-09-08. Each candidate opened without a repair
warning, displayed the expected comment state, was saved, had its document
window actually closed, and was reopened at its exact path. Strict focused
readback then passed. Every saved file differs from its generated candidate;
`native-receipt.json` records both SHA-256 hashes.

The nine cross-component cases cover root update/removal, reply add/update/
removal, root creation on a foreign drawable, shared-root isolation,
shared-reply preservation, and root-plus-reply creation on a foreign drawable.
The foreign drawable is the **Author and Date** footer text box. The tenth case
covers creation through the unique unrooted annotation-author registry.
Baseline movie, caption, square-comment, and slide semantics remain covered
by the focused readback assertions.

Ordinary tests read these checked-in native files:

```sh
cargo test -p litchi-keynote --test slide_drawable_comments_cross_component_native native_resaved
cargo test -p litchi-keynote --test slide_drawable_comment_registry native_resaved
```

For another native verification run, export to an owned scratch directory;
do not overwrite the checked-in fixtures while Keynote has them open:

```sh
LITCHI_KEYNOTE_DRAWABLE_COMMENTS_CROSS_COMPONENT_NATIVE_OUTPUT_DIR=/private/tmp/keynote-comment-candidates \
  cargo test -p litchi-keynote --test slide_drawable_comments_cross_component_native export_cross_component_native_candidates
LITCHI_KEYNOTE_DRAWABLE_COMMENT_REGISTRY_OUTPUT_DIR=/private/tmp/keynote-registry-candidate \
  cargo test -p litchi-keynote --test slide_drawable_comment_registry unrooted_unique_registry_creates_and_cleans_generated_author
```

Open, inspect, save, actually close, and reopen each candidate in Keynote.
Then run the readback tests with
`LITCHI_KEYNOTE_DRAWABLE_COMMENTS_CROSS_COMPONENT_NATIVE_DIR` and
`LITCHI_KEYNOTE_DRAWABLE_COMMENT_REGISTRY_NATIVE_DIR` pointing to those saved
directories. Export and readback use separate environment variables so a
readback run cannot replace native output with a freshly generated candidate.
Promote new goldens and update their receipt only after that native lifecycle
and semantic verification succeeds.
