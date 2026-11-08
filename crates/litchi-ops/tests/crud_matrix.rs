//! The CRUD checklist of record 0289 (plan step 25, design §14 layer 2).
//!
//! For every claimed format (the iWork formats since record 0296), every
//! element kind its capability report
//! lists is crossed with the verbs `get`, `find`, `query`, `view`, `set`,
//! `add`, `remove`, `move`, `swap`, `edit`, `batch`, `create` and `save`.
//! Each cell is classified from [`Capabilities`] (supported, typed gap with
//! its follow-up, out of scope), and every cell the report claims is
//! supported is exercised at least once against pinned fixtures from
//! `test-data/` or documents `create()` and `add` make, then republished
//! and reopened.
//!
//! Values are generated from the registry's value types, so a `set` cell
//! passes when one declared property is accepted, and an `add` cell when one
//! property combination (or a [`RECIPES`] entry) is accepted; the TSV names
//! what passed. `move` and `swap` are format verbs, not declared per
//! element: their cells record what the first sibling pair observed.
//!
//! Set `LITCHI_REPORT_DIR` to a scratch directory outside the repository to
//! emit the optional checklist. Capability coverage is asserted directly. A
//! claimed-supported cell that fails is a bug: [`KNOWN_FAILURES`] lists the
//! ones still open, each with its cause (records 0290, 0298 and 0303 closed
//! every one record 0289 found), and the test fails on any other.
//!
//! Record 0341 adds the ADR 0008 rows, one of each per format, built from
//! the format's successful changing `set` cells: **atomic failure** (a
//! batch whose item 1 fails is refused naming item 1 and leaves the
//! snapshot, its file and refused save destinations byte-identical),
//! **patch inverse** (the commit's patch replays onto its source, its
//! inverse restores the exact source bytes, and the patch is refused on
//! another artifact) and **composition** (two sets made on two threads from
//! one snapshot compose in either order to one result holding both values,
//! and each edit's patch is refused on the other's result). They are
//! asserted directly; [`ADR_NOT_APPLICABLE`] lists the
//! rows a format cannot support, with the typed reason (none today).

#![cfg(all(
    feature = "doc",
    feature = "docx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ppt",
    feature = "pptx",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "odg",
    feature = "odc",
    feature = "odi",
    feature = "odb",
    feature = "oth",
    feature = "odm",
    feature = "odf-formula",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Constraint, Element, Support, Verb};
use litchi_ops::value::ValueType;
use litchi_ops::view::Mode;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, NewNode, Placement, Props, create, open, open_bytes,
};

/// One format of the checklist: the corpus extension and the pinned
/// fixtures (relative to `test-data/`), chosen with [`discover_fixtures`]
/// as the smallest files that hold each element kind.
struct Spec {
    format: FileFormat,
    name: &'static str,
    extension: &'static str,
    fixtures: &'static [&'static str],
}

const SPECS: &[Spec] = &[
    Spec {
        format: FileFormat::Docx,
        name: "docx",
        extension: "docx",
        fixtures: &[
            "libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-html.docx",
            "libreoffice-core/sw/qa/extras/ooxmlexport/data/listWithLgl.docx",
            "ooxml/docx/Hyperlink.docx",
            "officecli/examples/word/revisions.docx",
            "ooxml/docx/table-indent.docx",
            "libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf157011_ins_del_empty_cols.docx",
            "libreoffice-core/sw/qa/extras/ooxmlexport/data/custom-styles-TOC-comma.docx",
            "libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx",
            "ooxml/docx/comment.docx",
            "libreoffice-core/oox/qa/unit/data/chart-data-label-char-color.docx",
            // Record 0319: notes, headers and footers, and numbering.
            "ooxml/docx/footnotes.docx",
            "ooxml/docx/endnotes.docx",
            "ooxml/docx/headerFooter.docx",
            "ooxml/docx/ComplexNumberedLists.docx",
            // Record 0319: a bookmark, a text box and a watermark.
            "libreoffice-core/sw/qa/extras/ooxmlexport/data/lvlPicBulletId.docx",
            "libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx",
            "libreoffice-core/sw/qa/extras/ooxmlexport/data/textWatermark.docx",
        ],
    },
    Spec {
        format: FileFormat::Xlsx,
        name: "xlsx",
        extension: "xlsx",
        fixtures: &[
            "ooxml/xlsx/DataValidationListTooLong.xlsx",
            "poi/test-data/spreadsheet/47889.xlsx",
            "ooxml/xlsx/column_style.xlsx",
            "ooxml/xlsx/ExcelTables.xlsx",
            "ooxml/xlsx/named-ranges-global.xlsx",
            "ooxml/xlsx/SimpleScatterChart.xlsx",
            "libreoffice-core/chart2/qa/extras/data/xlsx/boxWhisker.xlsx",
        ],
    },
    Spec {
        format: FileFormat::Pptx,
        name: "pptx",
        extension: "pptx",
        fixtures: &[
            "libreoffice-core/sd/qa/unit/data/pptx/tdf163852.pptx",
            "ooxml/pptx/shape-glow-effect.pptx",
            "ooxml/pptx/shapes.pptx",
            "ooxml/pptx/radar-chart.pptx",
            "libreoffice-core/oox/qa/unit/data/tdf131082.pptx",
            "ooxml/pptx/connectorConnection.pptx",
            "ooxml/pptx/smartart-composite-infer-right.pptx",
            "libreoffice-core/sd/qa/unit/data/pptx/ole.pptx",
            "libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx",
            "libreoffice-core/sd/qa/unit/data/pptx/tdf89064.pptx",
            "libreoffice-core/sd/qa/unit/data/pptx/tdf114821.pptx",
            "libreoffice-core/sd/qa/unit/data/pptx/tdf169496_hidden_graphic.pptx",
            "ooxml/pptx/sample.pptx",
            // Record 0294: a generated fixture (test-data/unified-ops/0294).
            "unified-ops/0294/modern-comments.pptx",
            // Record 0321: Office Math in a markup-compatibility choice, and
            // embedded audio and video (Apache POI, Apache-2.0).
            "libreoffice-core/oox/qa/unit/data/import-mce.pptx",
            "poi/test-data/slideshow/EmbeddedAudio.pptx",
            "poi/test-data/slideshow/EmbeddedVideo.pptx",
        ],
    },
    Spec {
        format: FileFormat::Xlsb,
        name: "xlsb",
        extension: "xlsb",
        fixtures: &[
            "ooxml/xlsb/date.xlsb",
            "libreoffice-core/sc/qa/unit/data/xlsb/shared_formula.xlsb",
            "libreoffice-core/sc/qa/unit/data/xlsb/tdf94627.xlsb",
            "poi/test-data/spreadsheet/testVarious.xlsb",
            // Record 0294: a generated fixture (test-data/unified-ops/0294).
            "unified-ops/0294/table.xlsb",
        ],
    },
    Spec {
        format: FileFormat::Doc,
        name: "doc",
        extension: "doc",
        fixtures: &[
            "ole/doc/documentProperties.doc",
            "ole/doc/hyperlink.doc",
            "ole/doc/table-merged-cells.doc",
            "poi/test-data/hpsf/TestNon4ByteBoundary.doc",
            "libreoffice-core/sw/qa/extras/ww8import/data/changes-in-footnote.doc",
            "ole/doc/endingnote.doc",
            "ole/doc/commented-table.doc",
            "ole/doc/image-comment-at-char.doc",
            // Record 0324: styles and lists, bookmarks, fields and a table of
            // contents, sections, floating shapes, and (generated) body tab
            // stops and a text box.
            "ole/doc/Lists.doc",
            "ole/doc/watermark.doc",
            "poi/test-data/document/Bug53453Section.doc",
            "ole/doc/FloatingPictures.doc",
            "unified-ops/0294/marks.doc",
        ],
    },
    Spec {
        format: FileFormat::Xls,
        name: "xls",
        extension: "xls",
        fixtures: &[
            "libreoffice-core/sc/qa/unit/data/xls/external-ref.xls",
            "ole/xls/SimpleWithColours.xls",
            "ole/xls/WithExtendedStyles.xls",
            "libreoffice-core/sc/qa/unit/data/xls/formats.xls",
            "ole/xls/SimpleWithAutofilter.xls",
            // Record 0327: notes, pictures, conditional formats, hyperlinks
            // and page breaks of producer workbooks (Apache POI, Apache-2.0).
            "ole/xls/SimpleWithComments.xls",
            "ole/xls/SimpleWithImages.xls",
            "ole/xls/WithConditionalFormatting.xls",
            "ole/xls/WithHyperlink.xls",
            "ole/xls/SimpleWithPageBreaks.xls",
        ],
    },
    Spec {
        format: FileFormat::Ppt,
        name: "ppt",
        extension: "ppt",
        fixtures: &[
            "poi/test-data/slideshow/incorrect_slide_order.ppt",
            "poi/test-data/slideshow/WithComments.ppt",
            "ole/ppt/text_shapes.ppt",
            "ole/ppt/ppt_with_png.ppt",
            "poi/test-data/slideshow/headers_footers.ppt",
            "poi/test-data/slideshow/basic_test_ppt_file.ppt",
            "poi/test-data/slideshow/42485.ppt",
            // Record 0294: a POI table deck, copied with provenance.
            "poi/test-data/slideshow/bug61881.ppt",
            // Record 0294: a generated fixture (test-data/unified-ops/0294).
            "unified-ops/0294/connectors.ppt",
            // Record 0323: comments, shape and text hyperlinks, a table and a
            // Graph chart (generated); the round-trip theme, an OLE workbook
            // and a sound frame (Apache POI).
            "unified-ops/0294/elements.ppt",
            "poi/test-data/slideshow/headers_footers_2007.ppt",
            "poi/test-data/slideshow/testPPT_oleWorkbook.ppt",
            "poi/test-data/slideshow/sound.ppt",
            // Record 0341: POI's sound frame rewritten to play a movie.
            "unified-ops/0294/video.ppt",
        ],
    },
    Spec {
        format: FileFormat::Rtf,
        name: "rtf",
        extension: "rtf",
        fixtures: &[
            "rtf/testDefaultEncodingParse.rtf",
            "rtf/column-break.rtf",
            "libreoffice-core/sw/qa/extras/rtfexport/data/table-rtl.rtf",
            "rtf/test437Encoding.rtf",
            "libreoffice-core/sw/qa/extras/rtfexport/data/tdf87034.rtf",
            "libreoffice-core/sw/qa/writerfilter/rtftok/data/endnote-at-section-end.rtf",
            "libreoffice-core/sw/qa/extras/rtfexport/data/tdf94377.rtf",
            "rtf/hyperlink.rtf",
            "libreoffice-core/sw/qa/extras/rtfexport/data/n751020.rtf",
            "libreoffice-core/sw/qa/extras/rtfexport/data/tdf167569.rtf",
            "libreoffice-core/sw/qa/extras/rtfexport/data/n192129.rtf",
            // Record 0328: sections, bookmarks, fields, a table of contents,
            // tab stops, a shape and a text box (hand-written).
            "unified-ops/0294/parts.rtf",
        ],
    },
    Spec {
        format: FileFormat::Odt,
        name: "odt",
        extension: "odt",
        fixtures: &[
            "odf/corpus/writer-header-footer.odt",
            "odf/corpus/writer-table.odt",
            "odf/corpus/writer-definition-lists.odt",
            "odf/corpus/writer-images-frames.odt",
            "odf/corpus/writer-table-of-contents.odt",
            "libreoffice-core/sw/qa/extras/layout/data/tdf124423.odt",
            // Record 0328: headers and footers, numbering, fields, a table of
            // contents, tab stops, shapes, equations and table columns
            // (generated, test-data/unified-ops/0294).
            "unified-ops/0294/parts.odt",
        ],
    },
    Spec {
        format: FileFormat::Ods,
        name: "ods",
        extension: "ods",
        fixtures: &[
            "odf/corpus/calc-formulas.ods",
            "odf/corpus/calc-two-sheets.ods",
            "odfdo/tests/samples/simple_table_named_range.ods",
            "odf/ods/grid-padding-package.ods",
        ],
    },
    Spec {
        format: FileFormat::Odp,
        name: "odp",
        extension: "odp",
        fixtures: &[
            "odf/odp/tdf105502.odp",
            "odf/odp/tdf169979.odp",
            "odf/corpus/impress-basic.odp",
            "odf/odp/text-in-image.odp",
            "odf/odp/cellspan.odp",
            "libreoffice-core/chart2/qa/extras/data/odp/tdf121189.odp",
            "libreoffice-core/sd/qa/unit/data/odp/trailing-paragraphs.odp",
            "libreoffice-core/sd/qa/unit/data/odp/group_rotation.odp",
            "libreoffice-core/sd/qa/unit/data/odp/tdf53970_linked.odp",
            // Record 0294: a generated fixture (test-data/unified-ops/0294).
            "unified-ops/0294/comments.odp",
            // Record 0328: hyperlinks, line breaks, columns, video and audio,
            // equations and chart series and axes (generated).
            "unified-ops/0294/parts.odp",
        ],
    },
    Spec {
        format: FileFormat::Odg,
        name: "odg",
        extension: "odg",
        fixtures: &[
            "libreoffice-core/desktop/qa/data/BlankDrawDocument.odg",
            "libreoffice-core/svx/qa/unit/data/FontWork.odg",
            "libreoffice-core/desktop/qa/data/3page.odg",
            // Record 0294: a generated fixture (test-data/unified-ops/0294).
            "unified-ops/0294/layers.odg",
        ],
    },
    Spec {
        format: FileFormat::Odc,
        name: "odc",
        extension: "odc",
        fixtures: &[
            "odf/odc-producer-evidence/odfdom-created.odc",
            "odf/odc-producer-evidence/odfdom-resaved.odc",
        ],
    },
    Spec {
        format: FileFormat::Odi,
        name: "odi",
        extension: "odi",
        fixtures: &[],
    },
    Spec {
        format: FileFormat::Odb,
        name: "odb",
        extension: "odb",
        fixtures: &[
            "libreoffice-core/extras/source/database/biblio.odb",
            "libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb",
        ],
    },
    Spec {
        format: FileFormat::Oth,
        name: "oth",
        extension: "oth",
        fixtures: &[
            // Record 0294: a generated fixture (test-data/unified-ops/0294).
            "unified-ops/0294/frames.oth",
            // Record 0328: the ODT parts fixture as a web template.
            "unified-ops/0294/parts.oth",
        ],
    },
    Spec {
        format: FileFormat::Odm,
        name: "odm",
        extension: "odm",
        fixtures: &[
            "libreoffice-core/sw/qa/extras/uiwriter/data/DUMMY.odm",
            "libreoffice-core/vcl/qa/cppunit/pdfexport/data/master.odm",
            // Record 0294: a generated fixture (test-data/unified-ops/0294).
            "unified-ops/0294/frames.odm",
            // Record 0328: the ODT parts fixture as a master document.
            "unified-ops/0294/parts.odm",
        ],
    },
    Spec {
        format: FileFormat::Odf,
        name: "odf",
        extension: "odf",
        fixtures: &[
            "odf/native-resave/source/font-styles.odf",
            "libreoffice-core/starmath/qa/extras/data/tdf151842.odf",
        ],
    },
    // Record 0296: the iWork formats. Record 0298 added the multi-slide
    // Keynote deck (a slide's `remove` no longer meets the leaf's
    // last-slide refusal) and the decks holding pictures, movies, audio,
    // tables and charts.
    Spec {
        format: FileFormat::Pages,
        name: "pages",
        extension: "pages",
        fixtures: &[
            "iwork/pages/body-sections-unicode.pages",
            "iwork/pages/basic.pages",
            // Record 0299: body tables, cells and anchored paragraphs.
            "iwork/pages/body-table-read-native.pages",
            "iwork/pages/body-table-visible.pages",
            // Record 0309: a body table without formulas takes cell values.
            "iwork/pages/body-table-name-focused.pages",
            // Record 0404: a native chart-grid fixture exercises the chart
            // get/find/query/view/set checklist cells on a real chart node.
            "iwork/pages/chart-grid-authored-native.pages",
        ],
    },
    Spec {
        format: FileFormat::Keynote,
        name: "keynote",
        extension: "key",
        fixtures: &[
            "iwork/keynote/multi-slide.key",
            "iwork/keynote/basic.key",
            "iwork/keynote/slide-table-read-native.key",
            "iwork/keynote/chart-grid-authored-native.key",
            "iwork/keynote/image-adjustments-native.key",
            "iwork/keynote/media-lifecycle-audio-remove-native.key",
            "iwork/keynote/media-lifecycle-remove-native.key",
            // Record 0341: a line and a group (generated).
            "unified-ops/0294/shapes.key",
        ],
    },
    Spec {
        format: FileFormat::Numbers,
        name: "numbers",
        extension: "numbers",
        // Record 0297: a merged range and a two-sheet workbook join the
        // checklist, so `mergeCell` reads and sheet/table moves are pinned.
        fixtures: &[
            "iwork/numbers/basic.numbers",
            "iwork/numbers/cell-value-native.numbers",
            "iwork/numbers/table-merges-native.numbers",
            "iwork/numbers/table-relocation-native.numbers",
            // Record 0404: a native chart-grid fixture exercises the chart
            // get/find/query/view/set checklist cells on a real chart node.
            "iwork/numbers/chart-grid-authored-native.numbers",
        ],
    },
];

/// The verbs of the checklist, in column order.
const VERBS: [&str; 13] = [
    "get", "find", "query", "view", "set", "add", "remove", "move", "swap", "edit", "batch",
    "create", "save",
];

/// An `add` the generic property attempts cannot express: the parent (on
/// the fixture labelled `fixture`), the properties and whether the node
/// goes first.
struct Recipe {
    format: &'static str,
    element: &'static str,
    fixture: &'static str,
    parent: &'static str,
    props: &'static [(&'static str, &'static str)],
    start: bool,
}

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4nGP4z8DwHxkzoAsAAA8hD/EEN8afAAAAAElFTkSuQmCC";

/// Inert media payloads (record 0321): the leaf stores them undecoded.
const MP4: &str = "data:video/mp4;base64,AAAAGGZ0eXBtcDQyAAAAAG1wNDJpc29t";
const MP3: &str = "data:audio/mpeg;base64,SUQzAwAAAAAAAA==";

const RECIPES: &[Recipe] = &[
    // Record 0338: an added XLSX cell names its address and a value.
    Recipe {
        format: "xlsx",
        element: "cell",
        fixture: "create()",
        parent: "/sheet[1]",
        props: &[("ref", "H8"), ("value", "added")],
        start: false,
    },
    // Record 0324: a DOC hyperlink names exactly one of `url` and `anchor`.
    Recipe {
        format: "doc",
        element: "hyperlink",
        fixture: "ole/doc/Lists.doc",
        parent: "/body/p[2]",
        props: &[("url", "http://example.com/"), ("text", "Litchi")],
        start: false,
    },
    // Record 0323: a PPT comment needs its author; a hyperlink its target.
    Recipe {
        format: "ppt",
        element: "comment",
        fixture: "unified-ops/0294/elements.ppt",
        parent: "/slide[2]",
        props: &[("author", "Ada"), ("text", "Check")],
        start: false,
    },
    Recipe {
        format: "ppt",
        element: "hyperlink",
        fixture: "unified-ops/0294/elements.ppt",
        parent: "/slide[1]/shape[1]",
        props: &[("link", "https://example.test/added")],
        start: false,
    },
    // Record 0297: a Numbers merged range needs an A1 `ref` of two cells or
    // more inside the table's grid.
    Recipe {
        format: "numbers",
        element: "mergeCell",
        fixture: "create()",
        parent: "/sheet[1]/table[1]",
        props: &[("ref", "B2:C3")],
        start: false,
    },
    // Record 0288: ODC axis/series add work on a chart Litchi made; the
    // odfdom fixtures hold content outside the leaf's lossless surface.
    Recipe {
        format: "odc",
        element: "axis",
        fixture: "create()",
        parent: "/",
        props: &[("dimension", "x")],
        start: false,
    },
    Recipe {
        format: "odc",
        element: "series",
        fixture: "create()",
        parent: "/",
        props: &[("values", "Sheet1.B1:B3")],
        start: false,
    },
    // Record 0292: the legend is added through the typed definition of a
    // chart Litchi made; a MathML token needs its tag and text together.
    Recipe {
        format: "odc",
        element: "legend",
        fixture: "create()",
        parent: "/",
        props: &[("position", "bottom")],
        start: false,
    },
    Recipe {
        format: "odf",
        element: "element",
        fixture: "create()",
        parent: "/",
        props: &[("tag", "mi"), ("text", "x")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "hyperlink",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("url", "https://example.test/"), ("text", "link")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "table",
        fixture: "seed",
        parent: "/body",
        props: &[("rows", "2"), ("cols", "2")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "comment",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("author", "Ada"), ("text", "Check")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "picture",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("src", PNG), ("alt", "Logo")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "chart",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("chartType", "column"), ("data", "A:1,2;B:3,4")],
        start: false,
    },
    // Record 0319: notes are anchored on a paragraph and need their text; a
    // level is added to a definition by its ilvl.
    Recipe {
        format: "docx",
        element: "footnote",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("text", "Note")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "endnote",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("text", "End")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "level",
        fixture: "ooxml/docx/ComplexNumberedLists.docx",
        parent: "/numbering/abstractNum[1]",
        props: &[("ilvl", "8"), ("format", "decimal"), ("lvlText", "%9.")],
        start: false,
    },
    // Record 0325: bookmarks and fields go into a paragraph, a table of
    // contents under the body.
    Recipe {
        format: "docx",
        element: "bookmark",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("name", "Mark"), ("text", "here")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "field",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("fieldType", "page")],
        start: false,
    },
    Recipe {
        format: "docx",
        element: "toc",
        fixture: "seed",
        parent: "/body",
        props: &[("levels", "1-3")],
        start: false,
    },
    // Record 0295: DOCX SmartArt is added under a paragraph from an outline.
    Recipe {
        format: "docx",
        element: "smartart",
        fixture: "seed",
        parent: "/body/p[1]",
        props: &[("layout", "process"), ("text", "Plan\nBuild")],
        start: false,
    },
    // Record 0335: a node joins the graphic the SmartArt cell added.
    Recipe {
        format: "docx",
        element: "node",
        fixture: "seed + add smartart",
        parent: "/smartart[1]",
        props: &[("text", "Ship")],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "chart",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("chartType", "column"), ("data", "A:1,2;B:3,4")],
        start: false,
    },
    // Record 0320: worksheet features need their anchor (a cell, a range, a
    // row or column, an image).
    Recipe {
        format: "xlsx",
        element: "comment",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("ref", "B2"), ("text", "note"), ("author", "Ann")],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "hyperlink",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("ref", "A1"), ("url", "https://example.test/")],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "picture",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[(
            "src",
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==",
        )],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "shape",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("geometry", "rect"), ("text", "box")],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "conditionalFormatting",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[
            ("type", "cellIs"),
            ("ref", "A1:A9"),
            ("value", "5"),
            ("fill", "FFC7CE"),
        ],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "dataValidation",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("type", "list"), ("ref", "A1:A9"), ("formula1", "a,b")],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "autofilter",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("range", "A1:C9")],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "rowbreak",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("row", "10")],
        start: false,
    },
    Recipe {
        format: "xlsx",
        element: "colbreak",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("col", "4")],
        start: false,
    },
    // Record 0322: the XLSB worksheet features take the XLSX recipes.
    Recipe {
        format: "xlsb",
        element: "comment",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("ref", "B2"), ("text", "note"), ("author", "Ann")],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "hyperlink",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("ref", "A1"), ("url", "https://example.test/")],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "picture",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[(
            "src",
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==",
        )],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "shape",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("geometry", "rect"), ("text", "box")],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "conditionalFormatting",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[
            ("type", "cellIs"),
            ("ref", "A1:A9"),
            ("value", "5"),
            ("fill", "FFC7CE"),
        ],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "dataValidation",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("type", "list"), ("ref", "A1:A9"), ("formula1", "a,b")],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "autofilter",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("range", "A1:C9")],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "rowbreak",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("row", "10")],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "colbreak",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("col", "4")],
        start: false,
    },
    // Record 0327: the XLS worksheet features take the XLSX recipes (a CF
    // fill is a palette color in BIFF8).
    Recipe {
        format: "xls",
        element: "comment",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("ref", "B2"), ("text", "note"), ("author", "Ann")],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "hyperlink",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("ref", "A1"), ("url", "https://example.test/")],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "picture",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[(
            "src",
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==",
        )],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "shape",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("geometry", "rect"), ("text", "box")],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "conditionalFormatting",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[
            ("type", "cellIs"),
            ("ref", "A1:A9"),
            ("value", "5"),
            ("fill", "FF0000"),
        ],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "dataValidation",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("type", "list"), ("ref", "A1:A9"), ("formula1", "a,b")],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "autofilter",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("range", "A1:C9")],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "rowbreak",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("row", "10")],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "colbreak",
        fixture: "create()",
        parent: "/Sheet1",
        props: &[("col", "4")],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "slide",
        fixture: "create()",
        parent: "/",
        props: &[("layout", "Blank")],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "shape",
        fixture: "create() + add slide",
        parent: "/slide[1]",
        props: &[("text", "Box"), ("width", "1in"), ("height", "1in")],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "chart",
        fixture: "create() + add slide",
        parent: "/slide[1]",
        props: &[("chartType", "column"), ("data", "A:1,2;B:3,4")],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "smartart",
        fixture: "create() + add slide",
        parent: "/slide[1]",
        props: &[("layout", "process"), ("text", "Plan\nBuild\nShip")],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "node",
        fixture: "create() + add slide + add smartart",
        parent: "/slide[1]/smartart[1]",
        props: &[("text", "Learn")],
        start: false,
    },
    // Record 0321: media is embedded with a poster image.
    Recipe {
        format: "pptx",
        element: "video",
        fixture: "create() + add slide",
        parent: "/slide[1]",
        props: &[("src", MP4), ("poster", PNG)],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "audio",
        fixture: "create() + add slide",
        parent: "/slide[1]",
        props: &[("src", MP3), ("poster", PNG)],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "media",
        fixture: "create() + add slide",
        parent: "/slide[1]",
        props: &[("src", MP4), ("poster", PNG)],
        start: false,
    },
    // Record 0330: claims that agree with the registry. A picture needs
    // its image and a legacy comment its author, both on a producer's
    // slide; an XLSB defined name needs its name and a sheet-qualified
    // reference.
    Recipe {
        format: "pptx",
        element: "picture",
        fixture: "ooxml/pptx/shapes.pptx",
        parent: "/slide[1]",
        props: &[("src", PNG), ("name", "Logo")],
        start: false,
    },
    Recipe {
        format: "pptx",
        element: "comment",
        fixture: "ooxml/pptx/shapes.pptx",
        parent: "/slide[1]",
        props: &[("author", "Ada"), ("text", "Check")],
        start: false,
    },
    // Record 0346: a modern comment needs its author and, with no clock,
    // its creation stamp.
    Recipe {
        format: "pptx",
        element: "modernComment",
        fixture: "unified-ops/0294/modern-comments.pptx",
        parent: "/slide[1]",
        props: &[
            ("author", "Ada"),
            ("text", "Check"),
            ("created", "2026-10-19T09:30:00Z"),
        ],
        start: false,
    },
    Recipe {
        format: "xlsb",
        element: "namedrange",
        fixture: "create()",
        parent: "/",
        props: &[("name", "Totals"), ("ref", "Sheet1!$A$1:$B$2")],
        start: false,
    },
    Recipe {
        format: "xls",
        element: "row",
        fixture: "create()",
        parent: "/sheet[1]",
        props: &[],
        start: true,
    },
    Recipe {
        format: "xls",
        element: "column",
        fixture: "create()",
        parent: "/sheet[1]",
        props: &[],
        start: true,
    },
    Recipe {
        format: "ods",
        element: "namedrange",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/",
        props: &[("name", "Totals"), ("ref", "sheet1!$A$1:$C$2")],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "comment",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1/B2",
        props: &[("text", "check"), ("author", "Reviewer")],
        start: false,
    },
    // Record 0327: the ODS worksheet features take the XLSX recipes.
    Recipe {
        format: "ods",
        element: "hyperlink",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[("ref", "A1"), ("url", "https://example.test/")],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "picture",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[(
            "src",
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==",
        )],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "shape",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[("geometry", "rect"), ("text", "box")],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "conditionalFormatting",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[
            ("type", "cellIs"),
            ("ref", "A1:A9"),
            ("value", "5"),
            ("fill", "FF0000"),
        ],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "dataValidation",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[("type", "list"), ("ref", "A1:A2"), ("formula1", "a,b")],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "autofilter",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[("range", "A1:C3")],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "rowbreak",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[("row", "2")],
        start: false,
    },
    Recipe {
        format: "ods",
        element: "colbreak",
        fixture: "odf/corpus/calc-two-sheets.ods",
        parent: "/sheet1",
        props: &[("col", "2")],
        start: false,
    },
    Recipe {
        format: "odp",
        element: "shape",
        fixture: "create() + add slide",
        parent: "/slide[1]",
        props: &[("text", "Box"), ("width", "8cm"), ("height", "2cm")],
        start: false,
    },
    Recipe {
        format: "keynote",
        element: "picture",
        fixture: "iwork/keynote/multi-slide.key",
        parent: "/slide[1]",
        props: &[("src", PNG)],
        start: false,
    },
    Recipe {
        format: "keynote",
        element: "slide",
        fixture: "iwork/keynote/multi-slide.key",
        parent: "/",
        props: &[("from", "/slide[2]")],
        start: false,
    },
    // Record 0329: a new table copies the deck's (document's) first table.
    Recipe {
        format: "keynote",
        element: "table",
        fixture: "iwork/keynote/slide-table-read-native.key",
        parent: "/slide[1]",
        props: &[("rows", "2"), ("cols", "2"), ("data", "a,b;c,d")],
        start: false,
    },
    Recipe {
        format: "pages",
        element: "table",
        fixture: "iwork/pages/body-table-visible.pages",
        parent: "/body",
        props: &[("rows", "2"), ("cols", "2"), ("data", "a,b;c,d")],
        start: false,
    },
    Recipe {
        format: "odg",
        element: "shape",
        fixture: "create() + add page",
        parent: "/page[1]",
        props: &[
            ("x", "1cm"),
            ("y", "1cm"),
            ("width", "4cm"),
            ("height", "2cm"),
            ("text", "Added"),
        ],
        start: false,
    },
];

/// One step of a seed document.
enum Step {
    Add(
        &'static str,
        &'static str,
        &'static [(&'static str, &'static str)],
    ),
    Set(&'static str, &'static [(&'static str, &'static str)]),
}

/// Documents built from `create()` with the unified verbs, labelled
/// `seed`, for the cells no pinned fixture can exercise (a second slide to
/// remove, a merged range, paragraphs on a format whose `create()` has none).
const SEEDS: &[(&str, &[Step])] = &[
    (
        "docx",
        &[
            Step::Add("/body", "paragraph", &[("text", "One")]),
            Step::Add("/body", "paragraph", &[("text", "Two")]),
        ],
    ),
    (
        "pptx",
        &[
            Step::Add("/", "slide", &[("layout", "Blank")]),
            Step::Add("/", "slide", &[("layout", "Blank")]),
            Step::Add(
                "/slide[1]",
                "shape",
                &[("text", "Box"), ("width", "1in"), ("height", "1in")],
            ),
        ],
    ),
    (
        "rtf",
        &[
            Step::Add("/body", "paragraph", &[("text", "One")]),
            Step::Add("/body", "paragraph", &[("text", "Two")]),
        ],
    ),
    ("ods", &[Step::Set("/Sheet1/A1:B2", &[("merge", "true")])]),
    // Record 0404: seed a named, non-empty second section so the claimed
    // section removal cell exercises a removable sibling, not the required
    // first section or an empty section that Pages correctly refuses.
    (
        "pages",
        &[
            Step::Add("/body", "paragraph", &[("text", "Intro")]),
            Step::Add("/", "section", &[("name", "Appendix")]),
            Step::Set("/section[2]", &[("text", "Tail")]),
        ],
    ),
    // Record 0309: character properties are written on runs of text boxes.
    (
        "keynote",
        &[Step::Add("/slide[1]", "shape", &[("text", "Box\nSecond")])],
    ),
];

/// Claimed-supported cells that fail, with the cause record 0289 found:
/// `(format, element, verb, cause)`. Each entry is a recorded bug. Record
/// 0290 fixed F1, F2, F4, F6 and F7; record 0298 fixed F8 (Keynote slide
/// removal, now exercised on the multi-slide deck); record 0303 fixed F3
/// (PPTX OLE removal under markup compatibility).
const KNOWN_FAILURES: &[(&str, &str, &str, &str)] = &[];

fn test_data() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
}

/// A document under test: where it came from and its handle.
struct Fixture {
    label: String,
    file: File,
}

/// One element instance: the fixture it is in and its canonical path.
#[derive(Clone, PartialEq, Eq)]
struct Instance {
    fixture: usize,
    path: String,
}

/// How the capability report classifies a cell.
#[derive(Clone, PartialEq, Eq)]
enum Claim {
    Supported,
    /// `move`/`swap`: supported by the format, not declared per element.
    FormatVerb,
    /// A typed gap with its follow-up.
    Gap(String),
    /// Not offered for the element and nothing declares a follow-up.
    OutOfScope(String),
}

/// What exercising a cell observed.
#[derive(Clone, PartialEq, Eq)]
enum Observed {
    /// Exercised successfully; the note says how.
    Ok(String),
    /// Refused with a typed code (expected for a gap).
    Refused(String),
    /// Failed although claimed (a bug unless listed).
    Failed(String),
    /// No instance of the element was reachable.
    Unreached,
    /// Not exercised (a cell with no claim, or no successful set to replay).
    Skipped,
}

struct Cell {
    claim: Claim,
    observed: Observed,
}

fn guarded<T>(f: impl FnOnce() -> litchi_ops::Result<T>) -> Result<T, String> {
    match catch_unwind(AssertUnwindSafe(f)) {
        Ok(Ok(value)) => Ok(value),
        Ok(Err(error)) => Err(format!("{}: {}", error.code().as_str(), error.message())),
        Err(panic) => {
            let message = panic
                .downcast_ref::<&str>()
                .map(|text| (*text).to_owned())
                .or_else(|| panic.downcast_ref::<String>().cloned())
                .unwrap_or_default();
            Err(format!("panic: {message}"))
        },
    }
}

fn code_of(message: &str) -> &str {
    message.split(':').next().unwrap_or("")
}

fn is_bug(error: &str) -> bool {
    code_of(error) == Code::InternalError.as_str() || error.starts_with("panic")
}

/// Distinct errors, at most three, in the order seen.
#[derive(Default)]
struct Errors(Vec<String>);

impl Errors {
    fn push(&mut self, error: String) {
        if self.0.len() < 3 && !self.0.contains(&error) {
            self.0.push(error);
        }
    }

    fn joined(&self, empty: &str) -> String {
        if self.0.is_empty() {
            empty.to_owned()
        } else {
            self.0.join(" | ")
        }
    }
}

/// Publishes and reopens `file`, the round trip every write cell needs.
fn republish(file: &File) -> Result<File, String> {
    let bytes = guarded(|| file.to_bytes())?;
    guarded(|| open_bytes(bytes))
}

fn load_fixtures(spec: &Spec) -> Vec<Fixture> {
    let mut fixtures = Vec::new();
    match guarded(|| create(spec.format)) {
        Ok(file) => fixtures.push(Fixture {
            label: "create()".to_owned(),
            file,
        }),
        // The iWork leaves have no from-scratch builder yet (record 0296,
        // `IWORK_API_REQUESTS.md` request 1): their checklist starts from
        // the pinned fixtures only.
        Err(error)
            if matches!(
                spec.format,
                FileFormat::Pages | FileFormat::Keynote | FileFormat::Numbers
            ) && code_of(&error) == Code::UnsupportedCapability.as_str() => {},
        Err(error) => panic!("{}: create failed: {error}", spec.name),
    }
    for (format, steps) in SEEDS {
        if *format != spec.name {
            continue;
        }
        let mut file = fixtures[0].file.clone();
        for step in *steps {
            file = match step {
                Step::Add(parent, kind, props) => guarded(|| {
                    let mut node = NewNode::new(*kind);
                    for (key, value) in *props {
                        node = node.prop(*key, *value);
                    }
                    file.add(*parent, node).map(|added| added.file)
                }),
                Step::Set(path, props) => guarded(|| {
                    let mut set = Props::new();
                    for (key, value) in *props {
                        set.push(*key, *value);
                    }
                    file.set(*path, &set)
                }),
            }
            .unwrap_or_else(|error| panic!("{format}: seed step failed: {error}"));
        }
        fixtures.push(Fixture {
            label: "seed".to_owned(),
            file,
        });
    }
    for relative in spec.fixtures {
        let bytes = std::fs::read(test_data().join(relative))
            .unwrap_or_else(|error| panic!("{}: fixture {relative}: {error}", spec.name));
        let file = guarded(|| open_bytes(bytes))
            .unwrap_or_else(|error| panic!("{}: fixture {relative}: {error}", spec.name));
        assert_eq!(file.format(), Some(spec.format), "{relative}");
        fixtures.push(Fixture {
            label: (*relative).to_owned(),
            file,
        });
    }
    fixtures
}

/// Walks one fixture breadth first, at most `budget` nodes, and returns
/// each node's kind and canonical path.
fn walk_one(file: &File, budget: usize) -> Vec<(&'static str, String)> {
    let result = catch_unwind(AssertUnwindSafe(|| {
        let mut local = Vec::new();
        let Ok(root) = file.root() else {
            return local;
        };
        local.push((root.kind(), root.path().to_string()));
        let mut queue = std::collections::VecDeque::from([root]);
        while let Some(node) = queue.pop_front() {
            for child in node.children() {
                let Ok(child) = child else { break };
                local.push((child.kind(), child.path().to_string()));
                if local.len() >= budget {
                    return local;
                }
                queue.push_back(child);
            }
        }
        local
    }));
    result.unwrap_or_default()
}

fn names(element: &Element) -> Vec<&'static str> {
    std::iter::once(element.name())
        .chain(element.aliases())
        .collect()
}

/// Every instance of `element`: walked nodes, query matches and path
/// templates instantiated with each fixture's first sheet, interleaved
/// across fixtures so consecutive attempts use different documents.
fn instances(
    element: &Element,
    fixtures: &[Fixture],
    walked: &[Vec<(&'static str, String)>],
    writes: bool,
) -> Vec<Instance> {
    let kinds = names(element);
    let mut per_fixture: Vec<Vec<String>> = Vec::new();
    for (index, fixture) in fixtures.iter().enumerate() {
        let mut list: Vec<String> = walked[index]
            .iter()
            .filter(|(kind, _)| kinds.contains(kind))
            .map(|(_, path)| path.clone())
            .take(6)
            .collect();
        if element.operations().query
            && let Ok(found) = guarded(|| {
                let mut paths = Vec::new();
                for node in fixture.file.query(element.name())?.take(6) {
                    paths.push(node?.path().to_string());
                }
                Ok(paths)
            })
        {
            list.extend(found);
        }
        let sheet = walked[index]
            .iter()
            .find(|(kind, _)| *kind == "sheet")
            .map(|(_, path)| path.trim_start_matches('/').to_owned());
        for (template, _) in element.paths() {
            if let Some(path) = instantiate(template, sheet.as_deref())
                && guarded(|| fixture.file.get(path.as_str()).map(|node| node.kind()))
                    .is_ok_and(|kind| kinds.contains(&kind))
            {
                list.push(path);
            }
        }
        let mut seen = BTreeSet::new();
        list.retain(|path| seen.insert(path.clone()));
        // Slide content before master and layout content; writes try
        // drawables before placeholders, whose runs keep their layout's
        // styles in Keynote (record 0309).
        list.sort_by_key(|path| {
            (
                path.contains("master") || path.contains("layout"),
                writes && path.contains("/placeholder["),
            )
        });
        per_fixture.push(list);
    }
    // Writes try the documents Litchi made first (added, seeded, created),
    // whose content every writer admits; reads try the pinned corpus
    // first, to exercise producers' content.
    let mut order: Vec<usize> = (0..fixtures.len()).collect();
    order.sort_by_key(|&index| {
        let label = &fixtures[index].label;
        let made = if label.contains(" + add ") {
            0
        } else if label == "seed" {
            1
        } else if label == "create()" {
            2
        } else {
            3
        };
        if writes { made } else { 3 - made }
    });
    let mut result = Vec::new();
    for round in 0..6 {
        for &fixture in &order {
            if let Some(path) = per_fixture[fixture].get(round) {
                result.push(Instance {
                    fixture,
                    path: path.clone(),
                });
            }
        }
    }
    result
}

/// A path template with its placeholders filled, when all of them can be.
fn instantiate(template: &str, sheet: Option<&str>) -> Option<String> {
    if !template.starts_with('/') {
        return None;
    }
    let mut path = template.to_owned();
    if path.contains("<sheet>") {
        path = path.replace("<sheet>", sheet?);
    }
    path = path
        .replace("<A1>:<A1>", "A1:B2")
        .replace("<A1>", "A1")
        .replace("<letters>", "A")
        .replace("[N]", "[1]")
        .replace("[K]", "[1]")
        .replace("[M]", "[1]");
    if path.contains('<') || path.contains("NAME") || path.contains("TYPE") || path.contains("ROLE")
    {
        return None;
    }
    Some(path)
}

/// The current value of `key` on the node at `path`, as the view prints it.
fn current(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).ok()?;
    let format = node.view().format().ok()?;
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

/// Candidate values for a property: changed values first, then the current
/// value (an exact no-op still exercises admission and dispatch).
fn candidates(ty: ValueType, constraint: Constraint, current: Option<&str>) -> Vec<String> {
    let mut values: Vec<String> = match ty {
        ValueType::Bool => vec!["true".into(), "false".into()],
        ValueType::Int => match constraint {
            Constraint::Int { min, max } => {
                let pick = current.and_then(|text| text.parse::<i64>().ok()).map_or(
                    min.max(1).min(max),
                    |value| {
                        if value < max { value + 1 } else { min }
                    },
                );
                vec![pick.to_string(), min.to_string()]
            },
            _ => vec!["2".into(), "1".into()],
        },
        ValueType::Number => match constraint {
            Constraint::Number { min, max } => {
                vec![((min + max) / 2.0).to_string(), min.to_string()]
            },
            _ => vec!["1.5".into(), "1".into()],
        },
        ValueType::Text | ValueType::Scalar => vec!["Litchi".into(), "7".into()],
        ValueType::Token(tokens) => tokens
            .values()
            .iter()
            .take(4)
            .map(|value| (*value).to_owned())
            .collect(),
        ValueType::Length(_) => vec!["12pt".into(), "2cm".into()],
        ValueType::Color(_) => vec!["FF0000".into(), "red".into()],
        _ => vec!["1".into()],
    };
    values.retain(|value| Some(value.as_str()) != current);
    if let Some(current) = current {
        values.push(current.to_owned());
    }
    values.truncate(5);
    values
}

/// Tries every settable property with its candidate values until one set
/// succeeds and republishes; returns the property, the value and whether
/// it equals the current value. `changed` selects the pass: changed values
/// only, or the current values (an exact no-op still runs admission and
/// dispatch).
fn try_set(
    element: &Element,
    file: &File,
    path: &str,
    changed: bool,
    errors: &mut Errors,
) -> Option<(String, String, bool)> {
    for property in element.properties().filter(|property| property.set()) {
        let now = current(file, path, property.name());
        let values = if changed {
            candidates(property.value_type(), property.constraint(), now.as_deref())
                .into_iter()
                .filter(|value| Some(value.as_str()) != now.as_deref())
                .collect()
        } else {
            now.clone().into_iter().collect::<Vec<_>>()
        };
        // The chart-grid fixtures intentionally start at 12pt. Use a known
        // distinct point size so the new chart CRUD cell proves mutation,
        // reopen and transaction paths rather than only a no-op admission.
        let values = if changed && element.name() == "chart" && property.name() == "font.size" {
            let target = if now.as_deref() == Some("17pt") {
                "19pt"
            } else {
                "17pt"
            };
            vec![target.to_owned()]
        } else {
            values
        };
        for value in values {
            let props = Props::new().with(property.name(), value.as_str());
            match guarded(|| file.set(path, &props)).and_then(|file| republish(&file)) {
                Ok(_) => {
                    let unchanged = now.as_deref() == Some(value.as_str());
                    return Some((property.name().to_owned(), value, unchanged));
                },
                Err(error) => {
                    let bug = is_bug(&error);
                    errors.push(format!("{}={value}: {error}", property.name()));
                    if bug {
                        return None;
                    }
                },
            }
        }
    }
    None
}

/// The property sets `add` is tried with, smallest first.
fn add_attempts(element: &Element) -> Vec<Vec<(String, String)>> {
    let addable: Vec<(String, String)> = element
        .properties()
        .filter(|property| property.add())
        .map(|property| {
            let value = candidates(property.value_type(), property.constraint(), None)
                .into_iter()
                .next()
                .unwrap_or_default();
            (property.name().to_owned(), value)
        })
        .collect();
    let mut attempts = vec![Vec::new()];
    for key in ["text", "name"] {
        if let Some(pair) = addable.iter().find(|(name, _)| name == key) {
            attempts.push(vec![pair.clone()]);
        }
    }
    if !addable.is_empty() {
        attempts.push(addable);
    }
    attempts
}

/// The properties of an `add`, as a JSON object in insertion order, so the
/// Python mirror (`crates/pyo3-litchi/tests/test_crud_mirror.py`) can
/// replay it. Data URIs are abbreviated to `<PNG>`, which stands for
/// [`PNG`].
fn props_json(props: &[(String, String)]) -> String {
    let fields: Vec<String> = props
        .iter()
        .map(|(key, value)| {
            let value = if value == PNG {
                "<PNG>"
            } else {
                value.as_str()
            };
            format!(
                "{}:{}",
                serde_json::Value::from(key.as_str()),
                serde_json::Value::from(value)
            )
        })
        .collect();
    format!("{{{}}}", fields.join(","))
}

/// One added element: the fixture it went into, the new file and path.
struct AddedNode {
    from: usize,
    file: File,
    path: String,
    how: String,
    parent: String,
    node: NewNode<'static>,
}

fn try_add(
    spec: &Spec,
    element: &Element,
    fixtures: &[Fixture],
    walked: &[Vec<(&'static str, String)>],
    errors: &mut Errors,
) -> Option<AddedNode> {
    let mut tries: Vec<(usize, String, Vec<(String, String)>, bool)> = Vec::new();
    for recipe in RECIPES
        .iter()
        .filter(|recipe| recipe.format == spec.name && recipe.element == element.name())
    {
        for (index, fixture) in fixtures.iter().enumerate() {
            if fixture.label == recipe.fixture {
                let props = recipe
                    .props
                    .iter()
                    .map(|(key, value)| ((*key).to_owned(), (*value).to_owned()))
                    .collect();
                tries.push((index, recipe.parent.to_owned(), props, recipe.start));
            }
        }
    }
    if tries.is_empty() {
        let mut parents = Vec::new();
        for parent in element.parents() {
            parents.extend(
                instances(&parent, fixtures, walked, true)
                    .into_iter()
                    .take(6),
            );
        }
        for parent in parents {
            for attempt in add_attempts(element) {
                tries.push((parent.fixture, parent.path.clone(), attempt, false));
            }
        }
    }
    for (index, parent, props, start) in tries {
        let mut node = NewNode::new(element.name());
        for (key, value) in &props {
            node = node.prop(key.clone(), value.clone());
        }
        if start {
            node = node.at(Placement::Start);
        }
        let file = &fixtures[index].file;
        match guarded(|| file.add(parent.as_str(), node.clone())) {
            Ok(added) => match republish(&added.file) {
                Ok(_) => {
                    return Some(AddedNode {
                        from: index,
                        path: added.path.to_string(),
                        file: added.file,
                        how: format!(
                            "under {parent} {}{}",
                            props_json(&props),
                            if start { " at start" } else { "" }
                        ),
                        parent,
                        node,
                    });
                },
                Err(error) => errors.push(format!("republish: {error}")),
            },
            Err(error) => {
                let bug = is_bug(&error);
                errors.push(error);
                if bug {
                    return None;
                }
            },
        }
    }
    None
}

fn format_verb(verb: &str) -> Verb {
    match verb {
        "get" => Verb::Get,
        "find" => Verb::Find,
        "query" => Verb::Query,
        "view" => Verb::View,
        "set" => Verb::Set,
        "add" => Verb::Add,
        "remove" => Verb::Remove,
        "move" => Verb::Move,
        "swap" => Verb::Swap,
        "edit" => Verb::Edit,
        "batch" => Verb::Batch,
        "create" => Verb::Create,
        _ => Verb::Save,
    }
}

fn element_claim(caps: &Capabilities, element: &Element, verb: &str) -> Claim {
    if let Support::Unsupported {
        follow_up, reason, ..
    } = caps.verb(format_verb(verb))
    {
        return if follow_up.starts_with("none") || reason.contains("out of scope") {
            Claim::OutOfScope(follow_up.to_owned())
        } else {
            Claim::Gap(follow_up.to_owned())
        };
    }
    if let Some(gap) = element.gap() {
        return Claim::Gap(gap.follow_up().to_owned());
    }
    let ops = element.operations();
    let declared = match verb {
        "get" | "find" | "view" => ops.get,
        "query" => ops.query,
        "set" | "edit" | "batch" => ops.set,
        "add" => ops.add,
        "remove" => ops.remove,
        // Moves and swaps are format verbs: the report does not list them
        // per element. Creation and saving apply to the whole format.
        "move" | "swap" => return Claim::FormatVerb,
        _ => true,
    };
    if declared {
        return Claim::Supported;
    }
    let write_gap = element
        .properties()
        .find_map(|property| property.write_gap())
        .filter(|_| matches!(verb, "set" | "edit" | "batch"));
    // An operation gap the report lists for this verb and element (record
    // 0292: the object formats refuse at element level, not verb level).
    let object_format = matches!(
        caps.format(),
        Some(
            FileFormat::Odg | FileFormat::Odc | FileFormat::Odi | FileFormat::Odb | FileFormat::Odf
        )
    );
    let operation_gap = || {
        caps.gaps().filter(|_| object_format).find(|gap| {
            matches!(
                gap.subject(),
                litchi_ops::capabilities::Subject::Operation { verb: gap_verb, element: gap_element, .. }
                    if gap_verb == verb && element.name().eq_ignore_ascii_case(gap_element)
            )
        })
    };
    match write_gap.or_else(operation_gap) {
        Some(gap) if gap.follow_up().starts_with("none") => {
            Claim::OutOfScope(gap.follow_up().to_owned())
        },
        Some(gap) => Claim::Gap(gap.follow_up().to_owned()),
        None => Claim::OutOfScope(format!("{verb} not offered for this element")),
    }
}

fn observed_error(claim: &Claim, error: String) -> Observed {
    if *claim == Claim::Supported || is_bug(&error) {
        Observed::Failed(error)
    } else {
        Observed::Refused(code_of(&error).to_owned())
    }
}

/// Two instances with the same parent path in one fixture.
fn sibling_pair(instances: &[Instance]) -> Option<(usize, String, String)> {
    for (i, a) in instances.iter().enumerate() {
        for b in &instances[i + 1..] {
            if a.fixture == b.fixture && parent(&a.path) == parent(&b.path) && a.path != b.path {
                return Some((a.fixture, a.path.clone(), b.path.clone()));
            }
        }
    }
    None
}

fn parent(path: &str) -> &str {
    path.rsplit_once('/').map_or("", |(head, _)| head)
}

/// The cells of one format, keyed by `(element, verb)`, in element order,
/// with format-level notes.
struct Report {
    order: Vec<String>,
    cells: BTreeMap<(String, &'static str), Cell>,
    notes: Vec<String>,
    /// The ADR 0008 rows, in [`ADR_ROWS`] order (record 0341).
    adr: Vec<AdrObserved>,
}

#[allow(
    clippy::too_many_lines,
    reason = "one pass over the verbs of the checklist"
)]
fn run_format(spec: &Spec) -> Report {
    let caps = Capabilities::of(spec.format);
    let mut fixtures = load_fixtures(spec);
    let mut walked: Vec<Vec<(&'static str, String)>> = fixtures
        .iter()
        .map(|fixture| walk_one(&fixture.file, 3000))
        .collect();
    let elements: Vec<Element> = caps.elements().collect();

    // Adds run first, in element order, so a later element can be added
    // under (and every later cell can reach) one an earlier add created.
    let mut adds: BTreeMap<&'static str, Result<(String, String), String>> = BTreeMap::new();
    for element in &elements {
        if element_claim(&caps, element, "add") != Claim::Supported {
            continue;
        }
        let mut errors = Errors::default();
        match try_add(spec, element, &fixtures, &walked, &mut errors) {
            Some(added) => {
                let label = format!("{} + add {}", fixtures[added.from].label, element.name());
                adds.insert(
                    element.name(),
                    Ok((
                        added.path.clone(),
                        format!("on `{}` {}", fixtures[added.from].label, added.how),
                    )),
                );
                // A second, identical add gives `move` and `swap` a pair of
                // siblings Litchi made (kept only when it succeeds).
                let twice = guarded(|| added.file.add(added.parent.as_str(), added.node.clone()))
                    .ok()
                    .filter(|again| republish(&again.file).is_ok());
                let file = match twice {
                    Some(again) => again.file,
                    None => added.file,
                };
                walked.push(walk_one(&file, 3000));
                fixtures.push(Fixture { label, file });
            },
            None => {
                adds.insert(element.name(), Err(errors.joined("no parent instance")));
            },
        }
    }

    let mut report = Report {
        order: Vec::new(),
        cells: BTreeMap::new(),
        notes: Vec::new(),
        adr: Vec::new(),
    };
    let mut candidates: Vec<Candidate> = Vec::new();
    let mut found_by_element: Vec<Vec<Instance>> = Vec::new();
    for (element_index, element) in elements.iter().enumerate() {
        let name = element.name().to_owned();
        report.order.push(name.clone());
        let read = instances(element, &fixtures, &walked, false);
        let found = instances(element, &fixtures, &walked, true);
        found_by_element.push(found.clone());
        let claim = |verb: &str| element_claim(&caps, element, verb);
        let mut cells: Vec<(&'static str, Observed)> = Vec::new();

        // Reads, on every instance found (at most eight): a producer file
        // that one verb cannot read is a failure of that verb.
        for verb in ["get", "find", "view"] {
            let mut observed = if read.is_empty() {
                Observed::Unreached
            } else {
                Observed::Ok(String::new())
            };
            let mut count = 0usize;
            for instance in read.iter().take(8) {
                let fixture = &fixtures[instance.fixture];
                let path = instance.path.as_str();
                let result = guarded(|| match verb {
                    "get" => fixture.file.get(path).map(|node| node.kind()),
                    "find" => fixture
                        .file
                        .find(path)
                        .map(|node| node.map_or("<none>", |node| node.kind())),
                    _ => fixture.file.get(path).and_then(|node| {
                        let view = node.view();
                        view.text()?;
                        view.format()?;
                        Ok(node.kind())
                    }),
                });
                match result {
                    Ok("<none>") => {
                        observed = Observed::Failed(format!("find({path}) returned None"));
                        break;
                    },
                    Ok(_) => count += 1,
                    Err(error) => {
                        observed = observed_error(
                            &claim(verb),
                            format!("{error} (`{}` {path})", fixture.label),
                        );
                        break;
                    },
                }
            }
            if let Observed::Ok(note) = &mut observed {
                *note = format!(
                    "{count} instances, first `{}` {}",
                    fixtures[read[0].fixture].label, read[0].path
                );
            }
            cells.push((verb, observed));
        }
        let observed = if found.is_empty() {
            Observed::Unreached
        } else if claim("query") == Claim::Supported {
            let mut errors = Errors::default();
            let mut hit = None;
            for fixture in &fixtures {
                match guarded(|| {
                    let mut count = 0usize;
                    for node in fixture.file.query(element.name())?.take(3) {
                        node?;
                        count += 1;
                    }
                    Ok(count)
                }) {
                    Ok(0) => {},
                    Ok(_) => {
                        hit = Some(fixture.label.clone());
                        break;
                    },
                    Err(error) => errors.push(error),
                }
            }
            match hit {
                Some(label) => Observed::Ok(format!("matches in {label}")),
                None => Observed::Failed(errors.joined(&format!(
                    "no match in any fixture although {} exists",
                    found[0].path
                ))),
            }
        } else {
            Observed::Skipped
        };
        cells.push(("query", observed));

        // Set, then the same set through an edit closure and a batch.
        let mut replay: Option<(usize, String, String, String)> = None;
        let mut unchanged = false;
        let observed = if found.is_empty() {
            Observed::Unreached
        } else if claim("set") != Claim::Supported {
            Observed::Skipped
        } else {
            let mut errors = Errors::default();
            'passes: for changed in [true, false] {
                for instance in found.iter().take(10) {
                    let file = &fixtures[instance.fixture].file;
                    if let Some((key, value, same)) =
                        try_set(element, file, &instance.path, changed, &mut errors)
                    {
                        unchanged = same;
                        replay = Some((instance.fixture, instance.path.clone(), key, value));
                        break 'passes;
                    }
                    if errors.0.iter().any(|error| is_bug(error)) {
                        break 'passes;
                    }
                }
            }
            match &replay {
                Some((fixture, path, key, value)) => Observed::Ok(format!(
                    "`{}` {path} {}{}",
                    fixtures[*fixture].label,
                    props_json(&[(key.clone(), value.clone())]),
                    if unchanged {
                        " (value unchanged: no-op only)"
                    } else {
                        ""
                    }
                )),
                None => Observed::Failed(errors.joined("no settable property")),
            }
        };
        cells.push(("set", observed));
        if let Some((fixture, path, key, value)) = &replay
            && !unchanged
        {
            candidates.push(Candidate {
                element: element_index,
                fixture: *fixture,
                path: path.clone(),
                key: key.clone(),
                value: value.clone(),
            });
        }
        let (edit, batch) = match &replay {
            None => (Observed::Skipped, Observed::Skipped),
            Some((fixture, path, key, value)) => {
                let file = &fixtures[*fixture].file;
                let props = Props::new().with(key.as_str(), value.as_str());
                let edit = guarded(|| file.edit(|tx| tx.set(path.as_str(), &props)))
                    .and_then(|commit| republish(commit.file()));
                let mut item = BatchItem::new("set");
                item.path = Some(path.as_str().into());
                item.props = Some(props.clone());
                let batch = guarded(|| file.batch([item], BatchMode::Atomic))
                    .and_then(|commit| republish(commit.file()));
                (
                    edit.map_or_else(
                        |error| observed_error(&claim("edit"), error),
                        |_| Observed::Ok("tx.set".into()),
                    ),
                    batch.map_or_else(
                        |error| observed_error(&claim("batch"), error),
                        |_| Observed::Ok("set item".into()),
                    ),
                )
            },
        };
        cells.push(("edit", edit));
        cells.push(("batch", batch));

        let observed = match adds.get(element.name()) {
            None => Observed::Skipped,
            Some(Ok((path, how))) => Observed::Ok(format!("{path} {how}")),
            Some(Err(error)) => observed_error(&claim("add"), error.clone()),
        };
        cells.push(("add", observed));

        let observed = if found.is_empty() {
            Observed::Unreached
        } else if claim("remove") != Claim::Supported {
            Observed::Skipped
        } else {
            let mut errors = Errors::default();
            let mut removed = None;
            // Keep the scan bounded, but reach a later sibling when the
            // first candidate is structurally non-removable (for example,
            // Pages' required first section).
            for instance in found.iter().take(24) {
                let file = &fixtures[instance.fixture].file;
                match guarded(|| file.remove(instance.path.as_str()))
                    .and_then(|file| republish(&file))
                {
                    Ok(_) => {
                        removed = Some(format!(
                            "`{}` {}",
                            fixtures[instance.fixture].label, instance.path
                        ));
                        break;
                    },
                    Err(error) => {
                        let bug = is_bug(&error);
                        errors.push(error);
                        if bug {
                            break;
                        }
                    },
                }
            }
            match removed {
                Some(path) => Observed::Ok(path),
                None => Observed::Failed(errors.joined("no instance")),
            }
        };
        cells.push(("remove", observed));

        // Moves and swaps need two siblings in one fixture.
        let pair = sibling_pair(&found);
        for verb in ["move", "swap"] {
            let observed = match &pair {
                None => Observed::Unreached,
                Some((fixture, a, b)) => {
                    let file = &fixtures[*fixture].file;
                    let result = guarded(|| {
                        if verb == "move" {
                            file.move_to(a.as_str(), Placement::after(b.as_str())?)
                        } else {
                            file.swap(a.as_str(), b.as_str())
                        }
                    })
                    .and_then(|file| republish(&file));
                    match result {
                        Ok(_) => Observed::Ok(format!("`{}` {a} {b}", fixtures[*fixture].label)),
                        Err(error) if is_bug(&error) => Observed::Failed(error),
                        Err(error) => Observed::Refused(code_of(&error).to_owned()),
                    }
                },
            };
            cells.push((verb, observed));
        }
        cells.push(("create", Observed::Skipped));
        cells.push(("save", Observed::Skipped));
        for (verb, observed) in cells {
            report.cells.insert(
                (name.clone(), verb),
                Cell {
                    claim: claim(verb),
                    observed,
                },
            );
        }
    }

    // Creation and saving are format verbs; every element row carries the
    // format result.
    let created = guarded(|| create(spec.format)).and_then(|file| {
        let reopened = republish(&file)?;
        if reopened.format() == Some(spec.format) {
            Ok(file)
        } else {
            Err(format!("reopened as {:?}", reopened.format()))
        }
    });
    // A format without a creator (the iWork formats, record 0296) saves
    // its first pinned fixture.
    let base = created.clone().or_else(|error| {
        fixtures
            .first()
            .map(|fixture| fixture.file.clone())
            .ok_or(error)
    });
    let saved = match &base {
        Ok(file) => (|| {
            let dir = tempfile::tempdir().map_err(|error| error.to_string())?;
            let path = dir.path().join(format!("saved.{}", spec.extension));
            guarded(|| file.save_as(&path))?;
            let reopened = guarded(|| open(&path))?;
            if reopened.format() == Some(spec.format) {
                Ok(())
            } else {
                Err(format!("reopened as {:?}", reopened.format()))
            }
        })(),
        Err(error) => Err(error.clone()),
    };
    for ((_, verb), cell) in &mut report.cells {
        let result = match *verb {
            "create" => created.as_ref().map(|_| ()).map_err(Clone::clone),
            "save" => saved.clone(),
            _ => continue,
        };
        cell.observed = match result {
            Ok(()) => Observed::Ok("format".into()),
            Err(error) => observed_error(&cell.claim, error),
        };
    }

    // Format-level views, on the first pinned fixture.
    let view_fixture = &fixtures[usize::from(fixtures.len() > 1)];
    let mut views = Vec::new();
    for mode in Mode::ALL {
        let claimed = caps.view(*mode).is_supported();
        let result = guarded(|| view_fixture.file.view(*mode));
        views.push(match (claimed, result) {
            (true, Ok(_)) => format!("{} ok", mode.name()),
            (true, Err(error)) => format!("{} FAILED ({})", mode.name(), clean(&error)),
            (false, Ok(_)) => format!("{} gap but succeeded", mode.name()),
            (false, Err(error)) => format!("{} gap ({})", mode.name(), code_of(&error)),
        });
    }
    report.adr = adr_rows(spec, &fixtures, &candidates, &mut |a: &Candidate| {
        let mut errors = Errors::default();
        let mut more = Vec::new();
        for instance in found_by_element[a.element].iter().take(10) {
            if instance.fixture != a.fixture || instance.path == a.path {
                continue;
            }
            let file = &fixtures[instance.fixture].file;
            if let Some((key, value, false)) = try_set(
                &elements[a.element],
                file,
                &instance.path,
                true,
                &mut errors,
            ) {
                more.push(Candidate {
                    element: a.element,
                    fixture: instance.fixture,
                    path: instance.path.clone(),
                    key,
                    value,
                });
            }
        }
        more
    });
    report.notes.push(format!(
        "fixtures: {}",
        fixtures
            .iter()
            .map(|fixture| format!("`{}`", fixture.label))
            .collect::<Vec<_>>()
            .join(", ")
    ));
    report.notes.push(format!(
        "document views on `{}`: {}",
        view_fixture.label,
        views.join(", ")
    ));
    report
}

/// The ADR 0008 rows of the checklist (record 0341), in column order.
const ADR_ROWS: [&str; 3] = ["atomic failure", "patch inverse", "composition"];

/// ADR 0008 rows a format cannot support, with the reason:
/// `(format, row, reason)`. A listed row that fails with a typed error is
/// reported `n/a`; one that passes fails the test (remove it).
const ADR_NOT_APPLICABLE: &[(&str, &str, &str)] = &[];

/// One successful, changing `set` of the checklist: the fixture, the path
/// and the property it wrote.
#[derive(Clone)]
struct Candidate {
    element: usize,
    fixture: usize,
    path: String,
    key: String,
    value: String,
}

impl Candidate {
    fn props(&self) -> Props<'_> {
        Props::new().with(self.key.as_str(), self.value.as_str())
    }

    fn json(&self) -> String {
        props_json(&[(self.key.clone(), self.value.clone())])
    }
}

/// What one ADR 0008 row observed: its evidence fields for the TSV (the
/// Python mirror replays them: fixture, path, props, second path, second
/// props, typed code) and its note, or why it failed.
enum AdrObserved {
    Ok { fields: Vec<String>, note: String },
    Failed(String),
}

/// Runs `f`, keeping a typed error whole; a panic becomes `Err(Err(_))`.
fn attempt<T>(
    f: impl FnOnce() -> litchi_ops::Result<T>,
) -> Result<T, Result<litchi_ops::Error, String>> {
    match catch_unwind(AssertUnwindSafe(f)) {
        Ok(Ok(value)) => Ok(value),
        Ok(Err(error)) => Err(Ok(error)),
        Err(_) => Err(Err("panic".to_owned())),
    }
}

fn describe(error: &Result<litchi_ops::Error, String>) -> String {
    match error {
        Ok(error) => format!("{}: {}", error.code().as_str(), error.message()),
        Err(panic) => panic.clone(),
    }
}

/// A typed refusal (not `internal_error`, not a panic), as its code.
fn typed(error: Result<litchi_ops::Error, String>) -> Result<&'static str, String> {
    match error {
        Ok(error) if error.code() != Code::InternalError => Ok(error.code().as_str()),
        error => Err(describe(&error)),
    }
}

/// A path that parses like `path` but names nothing: its last index is
/// 9999, or a missing child when it has no index.
fn missing_path(path: &str) -> String {
    match path.rfind('[') {
        Some(open)
            if path.ends_with(']') && path[open + 1..path.len() - 1].parse::<u32>().is_ok() =>
        {
            format!("{}[9999]", &path[..open])
        },
        _ => format!("{}/missing[9999]", path.trim_end_matches('/')),
    }
}

/// The input bytes a fixture was opened from, when it is a pinned file.
fn input_bytes(label: &str) -> Option<Vec<u8>> {
    if label.contains(" + add ") || label == "seed" || label == "create()" {
        return None;
    }
    std::fs::read(test_data().join(label)).ok()
}

/// Atomic failure: a batch whose item 1 names nothing, after a valid set
/// at item 0, is refused naming item 1, and leaves the snapshot, its file
/// on disk and two refused save destinations byte-identical.
fn adr_atomic(spec: &Spec, fixtures: &[Fixture], candidate: &Candidate) -> AdrObserved {
    let fixture = &fixtures[candidate.fixture];
    let run = || -> Result<(String, &'static str), String> {
        let before = guarded(|| fixture.file.to_bytes())?;
        let dir = tempfile::tempdir().map_err(|error| error.to_string())?;
        let source = dir.path().join(format!("source.{}", spec.extension));
        std::fs::write(&source, &before).map_err(|error| error.to_string())?;
        let origin = guarded(|| open(&source))?;
        let props = candidate.props();
        let missing = missing_path(&candidate.path);
        let mut first = BatchItem::new("set");
        first.path = Some(candidate.path.as_str().into());
        first.props = Some(props.clone());
        let mut second = BatchItem::new("set");
        second.path = Some(missing.as_str().into());
        second.props = Some(props.clone());
        // The first item alone changes the document, so the refusal below
        // rolls a real change back.
        let alone = guarded(|| origin.batch([first.clone()], BatchMode::Atomic))?;
        if guarded(|| alone.file().to_bytes())? == before {
            return Err("item 0 alone does not change the document".to_owned());
        }
        let code = match attempt(|| origin.batch([first, second], BatchMode::Atomic)) {
            Ok(_) => {
                return Err(format!(
                    "the batch with a missing {missing} at item 1 succeeded"
                ));
            },
            Err(Ok(error)) if error.code() != Code::InternalError => {
                if error.item() != Some(1) {
                    return Err(format!(
                        "the refusal names item {:?}, not 1: {}",
                        error.item(),
                        error.message()
                    ));
                }
                error.code().as_str()
            },
            Err(error) => return Err(format!("untyped refusal: {}", describe(&error))),
        };
        if guarded(|| origin.to_bytes())? != before {
            return Err("the snapshot changed after the refused batch".to_owned());
        }
        if std::fs::read(&source).map_err(|error| error.to_string())? != before {
            return Err("the snapshot's file changed after the refused batch".to_owned());
        }
        // A destination of another format and one under a regular file:
        // both saves are refused and neither existing file changes.
        let other = if spec.extension == "docx" {
            "xlsx"
        } else {
            "docx"
        };
        let sentinel = b"litchi sentinel".to_vec();
        let wrong = dir.path().join(format!("destination.{other}"));
        let blocker = dir.path().join("blocker");
        for path in [&wrong, &blocker] {
            std::fs::write(path, &sentinel).map_err(|error| error.to_string())?;
        }
        let under = blocker.join(format!("saved.{}", spec.extension));
        for (what, path) in [("another format", &wrong), ("a path under a file", &under)] {
            if let Err(error) = typed(match attempt(|| origin.save_as(path)) {
                Ok(_) => return Err(format!("save_as to {what} succeeded")),
                Err(error) => error,
            }) {
                return Err(format!("save_as to {what}: {error}"));
            }
        }
        for path in [&wrong, &blocker] {
            if std::fs::read(path).map_err(|error| error.to_string())? != sentinel {
                return Err(format!("{} changed", path.display()));
            }
        }
        let mut names: Vec<String> = std::fs::read_dir(dir.path())
            .map_err(|error| error.to_string())?
            .flatten()
            .map(|entry| entry.file_name().to_string_lossy().into_owned())
            .collect();
        names.sort();
        let expected = {
            let mut expected = vec![
                "blocker".to_owned(),
                format!("destination.{other}"),
                format!("source.{}", spec.extension),
            ];
            expected.sort();
            expected
        };
        if names != expected {
            return Err(format!("the refused saves left files behind: {names:?}"));
        }
        Ok((missing, code))
    };
    match run() {
        Ok((missing, code)) => AdrObserved::Ok {
            note: format!(
                "`{}` batch [set {} {}, set {missing}] refused `{code}` naming item 1; snapshot, its file and two refused save destinations unchanged",
                fixture.label,
                candidate.path,
                candidate.json()
            ),
            fields: vec![
                fixture.label.clone(),
                candidate.path.clone(),
                candidate.json(),
                missing,
                String::new(),
                code.to_owned(),
            ],
        },
        Err(error) => {
            AdrObserved::Failed(format!("`{}` {}: {error}", fixture.label, candidate.path))
        },
    }
}

/// Patch inverse: an edit's patch replays onto its source, its inverse
/// restores the exact source bytes, the patch refuses another artifact
/// (typed), and a separately opened, byte-identical copy of the source takes
/// it (content identity, record 0362).
fn adr_inverse(fixtures: &[Fixture], candidate: &Candidate) -> Result<AdrObserved, String> {
    let fixture = &fixtures[candidate.fixture];
    let before = guarded(|| fixture.file.to_bytes())?;
    let props = candidate.props();
    let commit = guarded(|| {
        fixture
            .file
            .edit(|tx| tx.set(candidate.path.as_str(), &props))
    })?;
    if !commit.diagnostics().changed() {
        // Not a change: the caller tries the next candidate.
        return Err("the edit changed nothing".to_owned());
    }
    let run = || -> Result<(usize, &'static str, bool, &'static str), String> {
        let after = guarded(|| commit.file().to_bytes())?;
        if after == before {
            return Err("the edit's bytes equal the source".to_owned());
        }
        let forward = guarded(|| commit.patch().apply(&fixture.file))?;
        if guarded(|| forward.to_bytes())? != after {
            return Err("the patch replayed onto its source differs from the commit".to_owned());
        }
        let inverse = commit.patch().inverse();
        let restored = guarded(|| inverse.apply(commit.file()))?;
        if guarded(|| restored.to_bytes())? != before {
            return Err("the inverse patch does not restore the exact source bytes".to_owned());
        }
        // Another artifact: the edited result is not the patch's source.
        let code = match attempt(|| commit.patch().apply(commit.file())) {
            Ok(_) => return Err("the patch applied to its own result".to_owned()),
            Err(error) => typed(error).map_err(|error| format!("stale apply: {error}"))?,
        };
        // A separately opened, byte-identical copy takes the patch: a patch
        // applies by content identity in every format (record 0362).
        let foreign = guarded(|| open_bytes(before.clone()))?;
        let copy = match attempt(|| commit.patch().apply(&foreign)) {
            Ok(applied) if guarded(|| applied.to_bytes())? == after => "accepts (content-checked)",
            Ok(_) => return Err("a byte-identical copy took the patch to other bytes".to_owned()),
            Err(error) => {
                let code = typed(error).map_err(|error| format!("copy apply: {error}"))?;
                return Err(format!("a byte-identical copy refused the patch: `{code}`"));
            },
        };
        let input = input_bytes(&fixture.label).is_some_and(|input| input == before);
        Ok((before.len(), code, input, copy))
    };
    Ok(match run() {
        Ok((bytes, code, input, copy)) => AdrObserved::Ok {
            note: format!(
                "`{}` set {} {}: the patch replays onto its source, the inverse restores the exact {bytes} source bytes{}, the patch on its own result refuses `{code}`; a byte-identical reopened copy {copy}",
                fixture.label,
                candidate.path,
                candidate.json(),
                if input {
                    " (the input file's bytes)"
                } else {
                    ""
                }
            ),
            fields: vec![
                fixture.label.clone(),
                candidate.path.clone(),
                candidate.json(),
                String::new(),
                String::new(),
                code.to_owned(),
            ],
        },
        Err(error) => {
            AdrObserved::Failed(format!("`{}` {}: {error}", fixture.label, candidate.path))
        },
    })
}

/// Composition: two edits prepared on two threads from one snapshot each
/// compose with the other in either order to the same bytes holding both
/// values, and either edit's patch refuses (typed) the other's result.
fn adr_compose(fixtures: &[Fixture], a: &Candidate, b: &Candidate) -> Result<AdrObserved, String> {
    let fixture = &fixtures[a.fixture];
    let file = &fixture.file;
    let (props_a, props_b) = (a.props(), b.props());
    let (first, second) = std::thread::scope(|scope| {
        let first = scope.spawn(|| guarded(|| file.edit(|tx| tx.set(a.path.as_str(), &props_a))));
        let second = scope.spawn(|| guarded(|| file.edit(|tx| tx.set(b.path.as_str(), &props_b))));
        (
            first.join().unwrap_or_else(|_| Err("panic".to_owned())),
            second.join().unwrap_or_else(|_| Err("panic".to_owned())),
        )
    });
    let (first, second) = (first?, second?);
    if !first.diagnostics().changed() || !second.diagnostics().changed() {
        return Err("an edit changed nothing".to_owned());
    }
    let run = || -> Result<(&'static str, &'static str, &'static str), String> {
        let ab = guarded(|| first.file().set(b.path.as_str(), &props_b))?;
        let ba = guarded(|| second.file().set(a.path.as_str(), &props_a))?;
        let value = |file: &File, candidate: &Candidate| {
            current(file, &candidate.path, &candidate.key).or_else(|| {
                let node = file.get(candidate.path.as_str()).ok()?;
                let text = node.view().text().ok()??.into_owned();
                Some(format!("text {text}"))
            })
        };
        let (want_a, want_b) = (value(first.file(), a), value(second.file(), b));
        for (order, composed) in [("a then b", &ab), ("b then a", &ba)] {
            let reopened = republish(composed)?;
            if value(&reopened, a) != want_a || value(&reopened, b) != want_b {
                return Err(format!(
                    "{order} dropped an edit: {:?} {:?}, expected {want_a:?} {want_b:?}",
                    value(&reopened, a),
                    value(&reopened, b)
                ));
            }
        }
        // Each order is repeatable; the two orders give the same bytes or,
        // where the format numbers new strings, links or records in
        // insertion order, the same document (text, node tree, both
        // values).
        let ab_bytes = guarded(|| ab.to_bytes())?;
        let again = guarded(|| first.file().set(b.path.as_str(), &props_b))?;
        if guarded(|| again.to_bytes())? != ab_bytes {
            return Err("a then b is not repeatable".to_owned());
        }
        let same = if ab_bytes == guarded(|| ba.to_bytes())? {
            "the same bytes"
        } else {
            let logical = |file: &File| -> Result<(String, Vec<(&'static str, String)>), String> {
                let reopened = republish(file)?;
                Ok((guarded(|| reopened.text())?, walk_one(&reopened, 3000)))
            };
            if logical(&ab)? != logical(&ba)? {
                return Err("the two orders give different documents".to_owned());
            }
            "the same document (text, node tree, both values; bytes differ in the order new strings, links or records are written)"
        };
        let stale_a = match attempt(|| second.patch().apply(first.file())) {
            Ok(_) => return Err("b's patch applied to a's result".to_owned()),
            Err(error) => typed(error).map_err(|error| format!("b's patch on a: {error}"))?,
        };
        let stale_b = match attempt(|| first.patch().apply(second.file())) {
            Ok(_) => return Err("a's patch applied to b's result".to_owned()),
            Err(error) => typed(error).map_err(|error| format!("a's patch on b: {error}"))?,
        };
        Ok((same, stale_a, stale_b))
    };
    Ok(match run() {
        Ok((same, stale_a, stale_b)) => AdrObserved::Ok {
            note: format!(
                "`{}` set {} {} and set {} {} on two threads: both orders hold both values and give {same}; each patch on the other's result refuses `{stale_a}`/`{stale_b}`",
                fixture.label,
                a.path,
                a.json(),
                b.path,
                b.json()
            ),
            fields: vec![
                fixture.label.clone(),
                a.path.clone(),
                a.json(),
                b.path.clone(),
                b.json(),
                format!("{stale_a}/{stale_b}"),
            ],
        },
        Err(error) => AdrObserved::Failed(format!(
            "`{}` {} and {}: {error}",
            fixture.label, a.path, b.path
        )),
    })
}

/// `b` with a value other than `a`'s generated text, so two names never
/// collide (two sheets or pages renamed alike conflict by design).
fn distinct(a: &Candidate, mut b: Candidate) -> Candidate {
    if b.value == a.value && b.value == "Litchi" {
        b.value = "Lychee".to_owned();
    }
    b
}

/// Whether two paths address independent nodes (neither holds the other)
/// or, on one node, different properties.
fn independent(a: &Candidate, b: &Candidate) -> bool {
    if a.fixture != b.fixture {
        return false;
    }
    if a.path == b.path {
        return a.key != b.key;
    }
    let nested = |outer: &str, inner: &str| {
        inner.starts_with(outer) && (outer.ends_with('/') || inner[outer.len()..].starts_with('/'))
    };
    !nested(&a.path, &b.path) && !nested(&b.path, &a.path)
}

/// The three ADR 0008 rows of one format, from its changing sets (and, for
/// composition, a second set found on another instance when the checklist
/// holds no independent pair).
fn adr_rows(
    spec: &Spec,
    fixtures: &[Fixture],
    candidates: &[Candidate],
    extra: &mut dyn FnMut(&Candidate) -> Vec<Candidate>,
) -> Vec<AdrObserved> {
    // Node sets before document-root sets, so item 1 of the atomic batch
    // names a missing sibling of a real node.
    let mut ordered: Vec<Candidate> = candidates.to_vec();
    ordered.sort_by_key(|candidate| candidate.path == "/");
    let candidates = ordered.as_slice();
    let none = |what: &str| AdrObserved::Failed(format!("no changing set {what}"));
    let atomic = candidates.first().map_or_else(
        || none("to batch"),
        |candidate| adr_atomic(spec, fixtures, candidate),
    );
    let mut inverse = None;
    let mut skipped = Errors::default();
    for candidate in candidates.iter().take(12) {
        match adr_inverse(fixtures, candidate) {
            Ok(observed) => {
                inverse = Some(observed);
                break;
            },
            Err(error) => skipped.push(error),
        }
    }
    let inverse = inverse.unwrap_or_else(|| {
        AdrObserved::Failed(format!(
            "no candidate edit: {}",
            skipped.joined("no changing set")
        ))
    });
    let mut compose = None;
    let mut skipped = Errors::default();
    // Two instances of one element kind first (two paragraphs, two cells,
    // two slides), then two element kinds, then two properties of one node.
    let mut pairs: Vec<(Candidate, Candidate)> = Vec::new();
    for a in candidates.iter().take(8) {
        for b in extra(a).into_iter().take(2) {
            if independent(a, &b) {
                pairs.push((a.clone(), distinct(a, b)));
            }
        }
    }
    let mut across = Vec::new();
    for (index, a) in candidates.iter().enumerate() {
        for b in &candidates[index + 1..] {
            if independent(a, b) {
                across.push((a.clone(), distinct(a, b.clone())));
            }
        }
    }
    across.sort_by_key(|(a, b)| a.path == b.path);
    pairs.extend(across);
    for (a, b) in pairs.iter().take(12) {
        match adr_compose(fixtures, a, b) {
            Ok(observed) => {
                compose = Some(observed);
                break;
            },
            Err(error) => skipped.push(error),
        }
    }
    let compose = compose.unwrap_or_else(|| {
        AdrObserved::Failed(format!(
            "no independent pair: {}",
            skipped.joined("no two independent changing sets")
        ))
    });
    vec![atomic, inverse, compose]
}

fn short(claim: &Claim, observed: &Observed) -> &'static str {
    match (claim, observed) {
        (Claim::Supported, Observed::Ok(_)) => "ok",
        (Claim::Supported, Observed::Failed(_)) => "FAIL",
        (Claim::Supported, Observed::Unreached) => "unreached",
        (Claim::Supported, Observed::Refused(_)) => "refused",
        (Claim::Supported, Observed::Skipped) => "-",
        (Claim::FormatVerb, Observed::Ok(_)) => "ok",
        (Claim::FormatVerb, Observed::Refused(_)) => "refused",
        (Claim::FormatVerb, Observed::Failed(_)) => "FAIL",
        (Claim::FormatVerb, _) => "no pair",
        (Claim::Gap(_), Observed::Ok(_)) => "gap (works)",
        (Claim::Gap(_), _) => "gap",
        (Claim::OutOfScope(_), Observed::Ok(_)) => "n/a (works)",
        (Claim::OutOfScope(_), _) => "n/a",
    }
}

fn detail(observed: &Observed) -> String {
    match observed {
        Observed::Ok(how) => format!("ok {how}"),
        Observed::Refused(code) => format!("refused {code}"),
        Observed::Failed(error) => format!("FAILED {error}"),
        Observed::Unreached => "unreached".into(),
        Observed::Skipped => "-".into(),
    }
}

/// Keeps an evidence line short and free of table separators.
fn clean(text: &str) -> String {
    let text = text.replace(['\t', '\n', '|'], " ");
    if text.chars().count() > 160 {
        let mut short: String = text.chars().take(157).collect();
        short.push_str("...");
        short
    } else {
        text
    }
}

#[test]
#[allow(clippy::too_many_lines, reason = "renders the checklist and checks it")]
fn crud_checklist_per_format() {
    let mut markdown = String::from(
        "# CRUD checklist (record 0289, generated)\n\n\
         Generated by `crates/litchi-ops/tests/crud_matrix.rs`; rewrite with\n\
         `LITCHI_REPORT_DIR=/path/to/scratch cargo test -p litchi-ops --all-features --test crud_matrix`.\n\n\
         Cells: `ok` claimed and exercised; `FAIL` claimed but failed (a bug);\n\
         `unreached` claimed but no fixture instance; `-` claimed, not exercised\n\
         (no successful `set` to replay); `gap` typed gap with a follow-up (TSV);\n\
         `n/a` not offered and no follow-up declared; `(works)` the report\n\
         understates. `move`/`swap` are format verbs, not declared per element:\n\
         `ok` the first sibling pair moved and swapped, `refused` a typed\n\
         refusal for that element kind (not a failure), `no pair` no two siblings.\n\
         Claimed cells count the per-element `get`/`find`/`query`/`view`/`set`/\n\
         `add`/`remove`/`edit`/`batch` claims plus `create`/`save` per row.\n\n\
         ADR 0008 rows (record 0341), one per format: `atomic failure` a\n\
         batch whose item 1 names nothing, after a valid set, is refused naming\n\
         item 1 and leaves the snapshot, its file on disk and two refused save\n\
         destinations byte-identical; `patch inverse` an edit's patch replays\n\
         onto its source, its inverse restores the exact source bytes and the\n\
         patch on its own result is refused (typed), and whether a\n\
         byte-identical reopened copy takes it is reported; `composition` two\n\
         sets made on two threads from one snapshot compose in either order\n\
         (each repeatable) to the same bytes, or to the same document where a\n\
         format writes new strings, links or records in edit order, holding\n\
         both values, and each edit's patch on the other's result is refused\n\
         (typed). `n/a` a typed reason the format cannot support the row.\n\
         Evidence: `crud-adr0008.tsv`.\n\n",
    );
    let mut summary = String::from(
        "| format | elements | claimed cells | ok | FAIL | unreached | gap | n/a | move/swap ok | move/swap refused | atomic failure | patch inverse | composition |\n\
         |---|---|---|---|---|---|---|---|---|---|---|---|---|\n",
    );
    let mut adr_tsv =
        String::from("format\trow\tstatus\tfixture\tpath\tprops\tpath2\tprops2\tcode\tnote\n");
    let mut adr_failures: Vec<String> = Vec::new();
    let mut adr_listed: Vec<(&str, &str)> = Vec::new();
    let mut tsv = String::from("format\telement\tverb\tclaim\tfollow_up\tobserved\n");
    let mut failures: Vec<(&str, String, &str, String)> = Vec::new();
    let mut unreached = Vec::new();
    let mut body = String::new();
    for spec in SPECS {
        let report = run_format(spec);
        let mut counts = [0usize; 8];
        let _ = writeln!(body, "## {}\n", spec.name);
        for note in &report.notes {
            let _ = writeln!(body, "- {note}");
        }
        let mut adr_short = Vec::new();
        for (row, observed) in ADR_ROWS.iter().zip(&report.adr) {
            let not_applicable = ADR_NOT_APPLICABLE
                .iter()
                .find(|(format, listed, _)| *format == spec.name && listed == row);
            let (status, fields, note) = match (observed, not_applicable) {
                (AdrObserved::Ok { fields, note }, None) => ("ok", fields.clone(), note.clone()),
                (AdrObserved::Ok { .. }, Some(_)) => {
                    adr_failures.push(format!("{} {row}: listed n/a but passes", spec.name));
                    ("ok", Vec::new(), "listed n/a but passes".to_owned())
                },
                (AdrObserved::Failed(error), Some((_, _, reason))) => {
                    adr_listed.push((spec.name, row));
                    ("n/a", Vec::new(), format!("{reason} ({})", clean(error)))
                },
                (AdrObserved::Failed(error), None) => {
                    adr_failures.push(format!("{} {row}: {error}", spec.name));
                    ("FAIL", Vec::new(), clean(error))
                },
            };
            adr_short.push(status);
            let _ = writeln!(body, "- ADR 0008 {row}: {status}, {note}");
            let mut fields = fields;
            fields.resize(6, String::new());
            let _ = writeln!(
                adr_tsv,
                "{}\t{row}\t{status}\t{}\t{}",
                spec.name,
                fields
                    .iter()
                    .map(|field| field.replace(['\t', '\n'], " "))
                    .collect::<Vec<_>>()
                    .join("\t"),
                note.replace(['\t', '\n'], " ")
            );
        }
        let _ = writeln!(body, "\n| element | {} |", VERBS.join(" | "));
        let _ = writeln!(body, "|---|{}", "---|".repeat(VERBS.len()));
        for element in &report.order {
            let row: Vec<&str> = VERBS
                .iter()
                .map(|verb| {
                    let cell = &report.cells[&(element.clone(), *verb)];
                    short(&cell.claim, &cell.observed)
                })
                .collect();
            let _ = writeln!(body, "| {element} | {} |", row.join(" | "));
            for verb in VERBS {
                let cell = &report.cells[&(element.clone(), verb)];
                let (claim, follow_up) = match &cell.claim {
                    Claim::Supported => ("supported", String::new()),
                    Claim::FormatVerb => ("format-verb", String::new()),
                    Claim::Gap(follow_up) => ("gap", follow_up.clone()),
                    Claim::OutOfScope(why) => ("n/a", why.clone()),
                };
                let _ = writeln!(
                    tsv,
                    "{}\t{element}\t{verb}\t{claim}\t{}\t{}",
                    spec.name,
                    clean(&follow_up),
                    match &cell.observed {
                        // Successful evidence is replayed by the Python
                        // mirror, so it is never shortened.
                        Observed::Ok(_) => detail(&cell.observed).replace(['\t', '\n', '|'], " "),
                        observed => clean(&detail(observed)),
                    }
                );
                match (&cell.claim, &cell.observed) {
                    (Claim::Supported, observed) => {
                        counts[0] += 1;
                        match observed {
                            Observed::Ok(_) => counts[1] += 1,
                            Observed::Failed(error) => {
                                counts[2] += 1;
                                failures.push((spec.name, element.clone(), verb, clean(error)));
                            },
                            Observed::Unreached => {
                                counts[3] += 1;
                                if !matches!(verb, "move" | "swap") {
                                    unreached.push(format!("{} {element} {verb}", spec.name));
                                }
                            },
                            _ => {},
                        }
                    },
                    (Claim::FormatVerb, observed) => match observed {
                        Observed::Ok(_) => counts[6] += 1,
                        Observed::Refused(_) => counts[7] += 1,
                        Observed::Failed(error) => {
                            failures.push((spec.name, element.clone(), verb, clean(error)));
                        },
                        _ => {},
                    },
                    (Claim::Gap(_), _) => counts[4] += 1,
                    (Claim::OutOfScope(_), _) => counts[5] += 1,
                }
            }
        }
        body.push('\n');
        let _ = writeln!(
            summary,
            "| {} | {} | {} | {} | {} | {} | {} | {} | {} | {} | {} |",
            spec.name,
            report.order.len(),
            counts[0],
            counts[1],
            counts[2],
            counts[3],
            counts[4],
            counts[5],
            counts[6],
            counts[7],
            adr_short.join(" | ")
        );
    }
    markdown.push_str(&summary);
    markdown.push('\n');
    markdown.push_str(&body);
    if let Some(output) = std::env::var_os("LITCHI_REPORT_DIR").map(PathBuf::from) {
        std::fs::create_dir_all(&output).unwrap();
        std::fs::write(output.join("crud-checklist.md"), &markdown).unwrap();
        std::fs::write(output.join("crud-checklist.tsv"), &tsv).unwrap();
        std::fs::write(output.join("crud-adr0008.tsv"), &adr_tsv).unwrap();
    }
    let unexpected: Vec<String> = failures
        .iter()
        .filter(|(format, element, verb, _)| {
            !KNOWN_FAILURES
                .iter()
                .any(|known| known.0 == *format && known.1 == element && known.2 == *verb)
        })
        .map(|(format, element, verb, error)| format!("{format} {element} {verb}: {error}"))
        .collect();
    let fixed: Vec<String> = KNOWN_FAILURES
        .iter()
        .filter(|known| {
            !failures.iter().any(|(format, element, verb, _)| {
                known.0 == *format && known.1 == element && known.2 == *verb
            })
        })
        .map(|known| format!("{} {} {}", known.0, known.1, known.2))
        .collect();
    eprintln!(
        "unreached claimed cells ({}):\n{}",
        unreached.len(),
        unreached.join("\n")
    );
    assert!(
        unexpected.is_empty(),
        "claimed-supported cells failed ({}):\n{}",
        unexpected.len(),
        unexpected.join("\n")
    );
    assert!(
        fixed.is_empty(),
        "known failures now pass; remove them: {fixed:?}"
    );
    assert!(
        adr_failures.is_empty(),
        "ADR 0008 rows failed ({}):\n{}",
        adr_failures.len(),
        adr_failures.join("\n")
    );
    let stale: Vec<_> = ADR_NOT_APPLICABLE
        .iter()
        .filter(|(format, row, _)| !adr_listed.contains(&(*format, *row)))
        .collect();
    assert!(
        stale.is_empty(),
        "ADR 0008 n/a rows not observed: {stale:?}"
    );
    assert!(
        unreached.is_empty(),
        "claimed-supported cells were not exercised: {unreached:?}"
    );
}

/// Prints, per element kind, the smallest corpus files that contain it, to
/// choose the pinned fixtures:
/// `cargo test -p litchi-ops --all-features --test crud_matrix -- --ignored --nocapture`.
#[test]
#[ignore = "fixture discovery for maintainers; prints, asserts nothing"]
fn discover_fixtures() {
    for spec in SPECS {
        let caps = Capabilities::of(spec.format);
        let elements: Vec<Element> = caps.elements().collect();
        let mut files = corpus(spec.extension);
        files.sort();
        let mut best: BTreeMap<&'static str, Vec<String>> = BTreeMap::new();
        for (size, relative) in files {
            let Ok(bytes) = std::fs::read(test_data().join(&relative)) else {
                continue;
            };
            let Ok(file) = guarded(|| open_bytes(bytes)) else {
                continue;
            };
            if file.format() != Some(spec.format) || file.is_signed() {
                continue;
            }
            let fixture = [Fixture {
                label: relative.clone(),
                file,
            }];
            let walked = vec![walk_one(&fixture[0].file, 3000)];
            for element in &elements {
                if !instances(element, &fixture, &walked, false).is_empty() {
                    let list = best.entry(element.name()).or_default();
                    if list.len() < 3 {
                        list.push(format!("{relative} ({size})"));
                    }
                }
            }
        }
        for element in &elements {
            eprintln!(
                "{} {}: {}",
                spec.name,
                element.name(),
                best.get(element.name())
                    .map_or("NONE".to_owned(), |list| list.join(" | "))
            );
        }
    }
}

/// Every corpus file with `extension` under 1 MiB, with its size.
fn corpus(extension: &str) -> Vec<(u64, String)> {
    let root = test_data();
    let mut found = Vec::new();
    let mut stack = vec![root.clone()];
    while let Some(dir) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else if path
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case(extension))
            {
                let size = entry.metadata().map_or(u64::MAX, |meta| meta.len());
                if size < 1024 * 1024 {
                    let relative = path
                        .strip_prefix(&root)
                        .unwrap()
                        .to_string_lossy()
                        .into_owned();
                    found.push((size, relative));
                }
            }
        }
    }
    found
}
