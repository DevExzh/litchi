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
`Cargo.toml`); nothing upstream is copied.

| File | Bytes | SHA-256 | Generator | Reaches |
|---|---|---|---|---|
| `modern-comments.pptx` | 10,725 | `2b1af0c58ef084017c981338ede7cc7931e198380df3485f952b02ad061f9bcb` | the synthetic deck of `crates/litchi-ops/tests/support/pptx_presentation.rs` (`bytes()`), packaged with the `zip` crate | PPTX `modernComment`, `reply` |
| `connectors.ppt` | 7,168 | `70d729c72c1c0ae4993763bc02bb5c0ae91bec94a2d545055f8624360802ce41` | `litchi_ppt::writer::Builder`: one slide, two `add_connector`, one `add_line` | PPT `connector` (and a sibling pair) |
| `elements.ppt` | 11,776 | `245ead4cce6d1cafc1cb619e16f12b6d8e2856daab90c0478e94cec987e4a262` | `litchi_ppt::writer::Builder` (record 0323): a text box with a URL shape hyperlink and a text hyperlink over characters 6..11, a `SlideComment`, a 2×2 table, a Microsoft Graph bar chart, and a second slide | PPT `comment`, `hyperlink`, `run`, `tr`, `tc`, `col`, `chart` |
| `table.xlsb` | 5,806 | `1722f8502afc2e15ae6c302fc188d551404cfcde5edf22256ccc02007dfe67b9` | `litchi_xlsb::Builder` with a `MutableWorksheet` holding a 4×3 range and one `Table` (`Items`) | XLSB `table` |
| `comments.odp` | 2,157 | `c80b5c8ee9f3c518246d461d27b2f212000ee3f5763511577c14254b0f960c1d` | `litchi_odp::Builder` (one slide), then two page annotations through the leaf edit's `add_annotation` | ODP `comment` |
| `layers.odg` | 908 | `118de79bef150787a5474aff373b4a0a8ab3980a3837192e25eaf458e8761d04` | `litchi_odf_common::core::PackageWriter` with content XML written in the test: one page with a page-level `draw:layer-set` of two layers | ODG `layer` |
| `frames.oth` | 1,639 | `3ba59ef90e5fab2ef93596e443302eaff52b6cf5e668c08fc9ad700fb3b0b665` | `PackageWriter`, `text-web` mimetype: a paragraph with two picture frames and a text-box frame, an embedded PNG, and `styles.xml` with two paragraph styles | OTH `frame`, `picture`, `textbox`, `style` |
| `frames.odm` | 1,218 | `eb3c61eafbec73e700a464706f151ac939b5eb1389a77c724dd98f5fdb430504` | as `frames.oth`, `text-master` mimetype, no `styles.xml` | ODM `frame`, `picture`, `textbox` |

The embedded PNG is the 4×2 image the checklist's picture recipes use.
