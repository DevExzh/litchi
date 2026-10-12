# 0215 — Unified legacy PowerPoint reads

The PPT facade previously exposed family text and slide reads while unified
`get`, `find`, `query` and semantic views returned a whole-format refusal. This
wave replaces its uninhabited node placeholder with the native-backed adapter
and connects the static dispatch used by path resolution, compiled query plans,
node properties, children, views and capabilities.

Native identities remain behind explicit raw paths and queries. Ordinary paths
use positions and supported semantic selectors. The source owner and read-only
presentation remain authoritative; the adapter does not translate the document
into a second semantic model. Traversal selects native slides as needed and
query termination stops before unselected slide traversal. `Slide::shared_shapes()` shares
the native parsed shape cache without cloning text or group payloads; repeated
borrowed and shared access on one native slide sees the same immutable array.
Placeholder queries use semantic kinds while preserving physical `/shape[N]`
paths. Native master-unit anchors are converted to typed EMUs using 3175/2;
plain lines omit unavailable geometry/rotation/hidden fields. Non-picture name
selectors are not advertised because the native projection lacks those names.

Text, annotated, outline and statistics views use the retained native projection.
Group traversal has explicit depth and shape-count limits. Unavailable chart
counts and picture alternate-text statistics are omitted rather than guessed;
this is a deliberate delta from the PPTX view schema. Element/property gaps
remain discoverable through capabilities. Mutation dispatch is still pending.

The native projection now exposes its existing freshness check as
`Presentation::check_source`. The facade uses it even when edit ownership was
released with `retain_source(false)`, so cached shape and catalog reads cannot
bypass the positional source fence. Signed files continue to permit reads.

Native `ShapeContainer` children now use concrete `ShapeEnum` values instead
of boxed shape trait objects. Callers pass the corresponding enum variant to
`add_child`; unused `Shape::clone_box` is removed. This keeps native shared
shape arrays and unified nodes `Send + Sync` with static dispatch.

## Verification

The native PPT all-feature suite passes 1,403 tests/doctests (11 existing
ignored cases); the facade all-feature suite passes 1,570. Final focused facade tests pass 30 cases after traversal-size and
feature-gating fixes. The PPT-only suite also covers native fixture placeholders, typed geometry, source-change errors,
iterator fusion, signed reads and released edit ownership. Node cloning makes
no allocations. The selected-get and first-match tests measure constant facade
allocation overhead relative to equivalent native reads at 1 versus 100 slides.

This does **not** establish constant total allocation: native selected-slide
reads allocate more for larger decks (observed total get: 68 allocations / 15,402
bytes at 1 slide, 73 / 34,374 at 100). The native slide factory's prerequisite
work remains a performance follow-up. Each `slide_at` returns a new slide; the
shared shape cache is per slide handle, not a presentation-wide retained cache.

PPT-only and no-feature builds, warning-denied all-target Clippy, warning-denied
rustdoc for both changed crates, the PPT glossary profile, crate boundaries
(66 packages / 268 dependency declarations), and the OfficeCLI delta checker
pass. Repository lockfiles are unchanged; an external offline lock/cache was
used for the locally available dependency versions. The base checkout remains
`32ad64e3207655b4ba558fc2793d9a7396655cdf`, already an ancestor of this branch.

## Exploratory paired read timing

Source implementation: `cc93039cee7e8327920dcd836b296f9759b50965`.

[Machine-readable samples](0215-ppt-read-results.json) retain seven alternating
AB/BA samples of 100 reads, with fixture creation, opening, path/query compilation
and text equality checks outside timing. The example uses release optimization
with LTO disabled; binary and benchmark-source hashes are recorded, not binaries.

| Slides | Source | Native µs/read (median) | Get/native paired median | First/native paired median |
|--------|--------|------------------------|--------------------------|----------------------------|
| 1 | memory | 2.83 | 1.221× | 1.229× |
| 1 | positional | 3.03 | 2.111× | 2.539× |
| 100 | memory | 6.50 | 1.105× | 1.105× |
| 100 | positional | 6.70 | 1.504× | 1.706× |
| 1000 | memory | 38.21 | 1.025× | 1.019× |
| 1000 | positional | 38.54 | 1.092× | 1.122× |

The source-backed facade repeats freshness checks at public read boundaries;
this is a material cost for small decks. Larger decks expose native prerequisite
work even for a selected first slide. These observations do not satisfy the
full goal's read-scaling gate. Follow-ups: profile and remove unrelated native
slide-factory prerequisite work, consolidate redundant facade freshness checks
without weakening mutation detection, and rerun the frozen harness across the
required scenarios. The probe establishes correctness and measures overhead;
it does not qualify all-format or cold/open/write performance.

## Remaining scope

Unified PPT mutation dispatch, complete read property/story coverage, frozen
harness performance qualification, and the remaining all-format/Python/CLI
requirements remain open. This record does not claim the full goal is complete.
