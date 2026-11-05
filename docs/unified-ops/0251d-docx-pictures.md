# 0251d — DOCX inline pictures: read and add

Plan step 7a item 5 of [0248](0248-coordinator-restart-plan.md) §6 Phase B.
Builds on the package-graph changes of [0251b](0251b-docx-hyperlink-writes.md).

## Leaf

**`Edit::insert_picture_at(paragraph_path, at, &NewPicture) -> Position`** on
the eager package-bound `document::Edit` (`document/owned/picture.rs`). It
stages two changes, published by one commit:

- the image as a new part `media/imageN.<ext>` beside the main part, with
  an `image` relationship (the Strict type in Strict documents);
- a new run holding an inline drawing that references it.

The run goes before the paragraph's direct run `at`, or after its last child.
The return value is the picture's position among the paragraph's inline
pictures, in `Paragraph::images` order. It is counted from the paragraph's
markup before the insertion point, so callers need not read the committed
document back.

**`NewPicture::{new(bytes), with_size(width, height), with_name, with_description}`.**

- The format comes from the signature: PNG, JPEG, GIF, BMP, TIFF, EMF or WMF.
  Unrecognized bytes are refused (`InvalidPropertyValue`).
- Without a size, the pixel size at 96 DPI is used: PNG `IHDR`, JPEG
  start-of-frame, GIF and BMP headers. With one dimension, the other
  follows the aspect ratio. Otherwise the size falls back to one inch.
- An extent outside 1 EMU to 22 inches is refused.
- The drawing identifier (`wp:docPr/@id`) is one above the main document's
  largest. The default name is `Picture N`.
- Every drawing namespace, and `r` for `r:embed`, is declared on the
  inserted markup.

Replay and exact inverse follow from 0251b (the inverse removes the media
part and its relationship). Source-backed owners refuse the graph change.

**Reader fix.** `InlineImage::{name, description}` now unescape their
attribute values. Before, `descr="A &lt;logo&gt;"` read back as the escaped
text.

## Unified surface

**Path.** `/body/p[N]/picture[M]`, and the same step under any paragraph:
cell, content control or nested table. `picture` was a gap (`R4x`) and is
now read.

- Properties are `width` and `height` (EMU lengths), `alt`, `name` and
  `relId`.
- Paragraphs list their pictures after their runs and hyperlinks. Available
  path suggestions name `picture`.

**add.** `add <paragraph> picture` (also `image` or `img`).

- `src` is required. It is a base64 `data:` URI or a local file path, read
  under the default staged-byte limit (16 MiB), as in PPTX 0252.
- `width`, `height`, `alt` and `name` are optional.
- Placement resolves against the paragraph's runs, like run add, and must
  land beside a direct run.
- The returned path is `…/picture[M]`.

Both owners go through the eager editor (0251b).

**Capabilities.**

- `picture` gains get, query and add.
- It has no set or remove yet.
- Both DOCX capability goldens are re-blessed and the OfficeCLI delta is
  regenerated.

The source reader duplicates PPTX's `read_source`. The duplicate stays local
because the PPTX module belongs to a parallel branch; a later change can
share it. The `docx` feature now enables the optional `base64` dependency.

## OfficeCLI parity

| OfficeCLI `picture` | Litchi |
|---|---|
| add `src` (path, URL, data URI or bytes) | path or base64 data URI; no URL fetch, no SVG fallback |
| add, get `width`, `height`, `alt` | same (no set) |
| get `relId` | same |
| `crop*`, `anchor`, `wrap`, `behindText`, positions, `link`, `decorative`, `effectExtent` | missing (floating pictures and effects) |
| get `id`, `contentType`, `fileSize` | missing |
| remove | missing |

## Measurements

The harness, machine and method are the same as 0251a. The case is
`picture_add_png`:

- leaf: `insert_picture_at` of a 4×2 PNG, then commit;
- unified: `add /body/p[151] picture`, whose source is the same PNG as a
  data URI;
- eager owner; 200-paragraph fixture.

| Trial | Leaf µs | Unified µs | Ratio |
|---|---:|---:|---:|
| 1 | 191.9 | 197.3 | 1.028 |
| 2 | 191.0 | 197.0 | 1.031 |
| 3 | 192.9 | 196.6 | 1.019 |

The first version located the new picture by reading the committed
paragraph back. That cost 1.94×, so it was replaced by the ordinal the leaf
returns.

Evidence: `results/0251d-docx-pictures/trial-{1,2,3}.json`. Regenerate with:

```
DOCX_BENCH_CASE=picture_add_png taskset -c 5 target/release/examples/docx_element_set_bench
```

## Verification

| Command | Exit |
|---|---|
| `cargo test -p litchi-docx --all-features` | 0 (`element_graph_edits`: media part and drawing published atomically, sizes, names, ids, replay, exact inverse, refusals) |
| `cargo test -p litchi-ops --all-features` | 0 (`docx_unified_nested`: data-URI and file sources, body and cell paragraphs, ordinals, reopen, refusals on both owners) |

## Not claimed

- Picture set and remove.
- Floating (anchored) pictures, cropping and click links.
- URL sources.
- Pictures in headers, footers or notes.
