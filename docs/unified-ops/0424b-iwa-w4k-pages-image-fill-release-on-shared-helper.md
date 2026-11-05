# 0424b — W4-K: Pages image-fill release on the shared helper

- Status: merge-ready on branch `uo2/iwa-w4k`, after
  [0424a](0424a-iwa-w4k-cell-edits-on-pivot-and-category-tables.md)
  (merged at `baa2644511`).
- Date: 2026-11-03.
- `performance_claim`: the Pages release route costs 1.4% fewer leaf
  instructions than its previous private copy, and setting an image fill
  costs 0.6% more. Both are inside the 5% budget (§3).

## 1. What moved

Pages `drawable::style` (`package/body_drawable_style.rs`) had its own
copy of the image-fill release from record 0415: `Release`,
`released_image` and `release_images`, about 150 lines over its own
editable-archive list. All of it is deleted. The commit now does the
following:

- **Capture.** Before a fill set, a fill reset or a new image fill, the
  commit captures the old fill with `litchi_iwa_style::shape::stored_image`.
  This is the image the shape's own style stores; an inherited fill is not
  the shape's to release.
- **Release.** Before `workspace.finish()`, the commit releases on the
  shared workspace with `shape::release_image_fills`, record 0426's helper
  on `data_owner::release`:
  - the style drops its message-info data reference;
  - the registered owner is dropped;
  - data nothing references any more loses its `DataInfo`.
- **Deletion.** A `Data/` member the package does not hold is not deleted.
  This is the old copy's guard, kept as a filter on the returned members.

There is no public API change, and the Numbers route of record 0426 is
unchanged.

## 2. The keep rule: the shared rule is the host's

The two rules for keeping an unreferenced asset differ:

- **Record 0415's copy:** any `DataMetadataMap` in the registry kept every
  released asset.
- **`data_owner::release`:** an asset is kept only by a map entry naming
  its data, or by a message-info data reference. This is the host's
  `EmbeddedMediaAsset::is_referenced`, which the host's
  `remove_orphaned_image_asset` uses for every format.

The coordinator allowed either a parameter for the difference or a proof
that the rules are equivalent. I took neither: the evidence shows the
shared rule is the host's, and the 0415 rule was a deviation from it.

- **Equivalence where the old parity ran.** The 0413b/0415 Pages fixtures
  are source-built, and their registry has no map (`PackageMetadata` field
  10 is absent). A temporary digest test, removed before commit, ran five
  leaf routes on cached fixtures, with the previous owner and with the new
  one:
  - the image fill with scale-to-fill and a tint;
  - the image fill tiled;
  - releases by a solid fill, by a reset and by a new image.

  The output and inverse-restored packages are byte-identical. All Pages
  image-fill parity cases still pass unchanged
  (`pages_body_shape_image_fill_parity`,
  `pages_body_shape_image_fill_release_parity`,
  `pages_shape_image_fill_read_parity`, the chart image-fill cases).
- **Where the rules differ.** Every native Pages, Keynote and Numbers
  fixture in `test-data/` has a map. On `basic.pages` with a host rectangle
  and a host image fill, a solid fill gives these `Data/` members:

  | Route | `Data/` members after the solid fill |
  |---|---|
  | host | `[]` (the orphan is deleted) |
  | previous leaf copy (0415 rule) | `["Data/poster-24.png"]` (an orphan left behind) |
  | leaf on the shared helper | `[]` |

  The new tier-B case `pages_body_shape_image_fill_release_native_map_parity`
  (solid, reset and new image on that fixture) pins the host's behaviour.
  So the move removes a Pages-only orphaned-media leak rather than adding
  a hook to preserve it.

## 3. Performance

Callgrind measured two release parity binaries with debug info, neither
committed:

- the previous owner, SHA-256
  `67ecd74115eb02006377077c13761b3ea58f593197c47be05d32e0338f90a584`;
- the shared helper, SHA-256
  `2a837c27b9acdafb3456d3f42619dd86f35832f7d34b3e1d6fc98a5b4e30c707`.

The runs pinned CPU 5 and used ABBA order, with `--toggle-collect` on
`*probe_host*` / `*probe_leaf*`.

| Route (probe) | Previous leaf Ir | Shared leaf Ir | New/previous | Host Ir (new binary) |
|---|---:|---:|---:|---:|
| Solid fill over an image fill: the release (`shape_image_fill_release_cost_abba`, new) | 5,447,596 | 5,371,752 | 0.986 | 29,708,245 |
| Image fill on a plain shape (`shape_image_fill_cost_abba`, 0413b) | 11,191,756 | 11,256,358 | 1.006 | 35,215,027 |

## 4. Verification

The gate list in `(deleted-coord-artifact)`, with `litchi-pages` added to
clippy, tests and docs, runs on the merge of trunk. Its results are
reported in the hand-off.
