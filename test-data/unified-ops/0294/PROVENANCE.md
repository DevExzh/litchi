# Generated CRUD checklist fixtures (record 0294)

Record 0289's CRUD checklist left claimed cells unreached because no file in
`test-data/`, and no small LibreOffice, Apache POI or Tika file, holds the
element. The files here are written by Litchi's own writers in
`crates/litchi-ops/tests/acceptance_fixtures.rs`. That test regenerates them,
checks that two generations are byte-identical and that the checked-in bytes
equal them, and rewrites them with
`LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test acceptance_fixtures`.

They are synthetic, not producer files: they show that each element's
adapter code works, not that every third-party variant of the element does.
They are part of this repository and carry its license (see the workspace
`Cargo.toml`). Only `video.ppt` derives from an upstream file (Apache POI,
Apache-2.0, notice in `test-data/poi/NOTICE`); every other file copies nothing
upstream.

| File | Bytes | SHA-256 | Generator | Reaches |
|---|---|---|---|---|
| `modern-comments.pptx` | 10,725 | `2b1af0c58ef084017c981338ede7cc7931e198380df3485f952b02ad061f9bcb` | the synthetic deck of `crates/litchi-ops/tests/support/pptx_presentation.rs` (`bytes()`), packaged with the `zip` crate | PPTX `modernComment`, `reply` |
| `connectors.ppt` | 7,168 | `70d729c72c1c0ae4993763bc02bb5c0ae91bec94a2d545055f8624360802ce41` | `litchi_ppt::writer::Builder`: one slide, two `add_connector`, one `add_line` | PPT `connector` (and a sibling pair) |
| `elements.ppt` | 11,776 | `245ead4cce6d1cafc1cb619e16f12b6d8e2856daab90c0478e94cec987e4a262` | `litchi_ppt::writer::Builder` (record 0323): a text box with a URL shape hyperlink and a text hyperlink over characters 6..11, a `SlideComment`, a 2×2 table, a Microsoft Graph bar chart, and a second slide | PPT `comment`, `hyperlink`, `run`, `tr`, `tc`, `col`, `chart` |
| `table.xlsb` | 5,806 | `1722f8502afc2e15ae6c302fc188d551404cfcde5edf22256ccc02007dfe67b9` | `litchi_xlsb::Builder` with a `MutableWorksheet` holding a 4×3 range and one `Table` (`Items`) | XLSB `table` |
| `comments.odp` | 2,157 | `c80b5c8ee9f3c518246d461d27b2f212000ee3f5763511577c14254b0f960c1d` | `litchi_odp::Builder` (one slide), then two page annotations through the leaf edit's `add_annotation` | ODP `comment` |
| `layers.odg` | 908 | `118de79bef150787a5474aff373b4a0a8ab3980a3837192e25eaf458e8761d04` | `litchi_odf_common::core::PackageWriter` with content XML written in the test: one page with a page-level `draw:layer-set` of two layers | ODG `layer` |
| `frames.oth` | 1,639 | `3ba59ef90e5fab2ef93596e443302eaff52b6cf5e668c08fc9ad700fb3b0b665` | `PackageWriter`, `text-web` mimetype: a paragraph with two picture frames and a text-box frame, an embedded PNG, and `styles.xml` with two paragraph styles | OTH `frame`, `picture`, `textbox`, `style` |
| `marks.doc` | 15,360 | `b2697f643b8aade5b675fbc0e0f7b380df3882aee92bb6cbe192cb907ce38ca0` | `litchi_doc::Builder` (record 0324): a body paragraph with two tab stops (`tab_stops_to_add`), a floating text box (`insert_floating_text_box`) and a closing paragraph | DOC `tab`, `textbox` |
| `frames.odm` | 1,218 | `eb3c61eafbec73e700a464706f151ac939b5eb1389a77c724dd98f5fdb430504` | as `frames.oth`, `text-master` mimetype, no `styles.xml` | ODM `frame`, `picture`, `textbox` |
| `parts.odt` | 3,212 | `bef3e97a356fb8b1f4b1956d9bb42b0cee290baed5e27a1bf577d0cb208b3220` | `PackageWriter` (record 0328), content, styles and formula XML written in the test: a common and an automatic list style and two lists, a paragraph style with two tab stops, five text fields, a table of contents, a table with three declared columns, a custom shape and a rectangle, a formula subdocument (`Object 1/`) and an inline formula, and two master pages with five header and footer regions | ODT `header`, `footer`, `numbering`, `abstractNum`, `num`, `level`, `field`, `toc`, `tab`, `shape`, `equation`, `column` |
| `parts.oth` | 3,219 | `a264abeefbdbda1d164e8261568e00c58d73c9a8e5df814e8a3816cec1b38725` | as `parts.odt`, `text-web` mimetype | the same OTH kinds |
| `parts.odm` | 3,223 | `ae577d28afe0b74b0befb67f0076b14b9662b5974120555e51ca0dd08f2d7cdd` | as `parts.odt`, `text-master` mimetype | the same ODM kinds |
| `parts.odp` | 3,196 | `7cb7864322129dec6fb58ec7ed712544264c01260e7e54ce8bcb0b435714278c` | `PackageWriter` (record 0328), content, styles, formula and chart XML written in the test: a text link with a title, two line breaks, click listeners (a URL, the next page), a `draw:a` link, a table with three declared columns, a video and an audio plugin with inert payloads, a formula subdocument and an inline formula, a chart subdocument with two axes and two series, and a second slide | ODP `hyperlink`, `linebreak`, `col`, `video`, `audio`, `equation`, `series`, `axis` |
| `parts.rtf` | 1,028 | `ce02b25cade06b8b67d67008ed848cb61e92bcdaed45a247294a7d300b851738` | hand-written RTF bytes (record 0328) in the test: a document section and a continuous landscape two-column section, the bookmarks `Intro` and `Second`, `PAGE`, `NUMPAGES` and `TOC` fields, a paragraph with a left and a right dot-leader tab stop, a rectangle shape and a text-box shape | RTF `section`, `bookmark`, `field`, `toc`, `tab`, `shape`, `textbox` |
| `video.ppt` | 43,520 | `f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a` | record 0341: Apache POI's `test-data/slideshow/sound.ppt` (Apache-2.0; copied as `test-data/poi/test-data/slideshow/sound.ppt`, SHA-256 `d6425bec…c0b1`) rewritten in place: its one 40-byte `ExWAVAudioEmbeddedContainer` becomes an `ExAviMovieContainer` of the same size holding the same `ExMediaAtom` (id 1) and an empty path atom, so no offset moves | PPT `video` |
| `shapes.key` | 456,779 | `72f0ff952c4f9838db8eac3ffd75d93cb8a7e66dfa778bb9b69bb44dddd016cb` | record 0341: `test-data/iwork/keynote/basic.key` plus two rectangles and an image from `litchi_keynote`'s drawable edit, then a wire-level rewrite of the slide archive: the first rectangle's Bézier path keeps its move-to and first line-to (a straight line), and the image becomes a `TSD.GroupArchive` holding only its drawable (no children) | Keynote `connector`, `group` |

The embedded PNG is the 4×2 image the checklist's picture recipes use.
