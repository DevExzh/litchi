//! Generated fixtures for the CRUD checklist cells no producer file in the
//! corpus reaches (record 0294).
//!
//! Record 0289 left 70 claimed cells unreached because no file under
//! `test-data/` (and no small LibreOffice, Apache POI or Tika file) holds the
//! element. Each fixture here is built with Litchi's own writers: the leaf
//! builders and edit transactions, and the ODF package writer of
//! `litchi-odf-common` for content XML written in this file. They are
//! checked in under `test-data/unified-ops/0294/` so the Rust checklist
//! (`tests/crud_matrix.rs`), the corpus reports and the Python mirror all
//! open the same bytes, and this test proves they are what the generators
//! write: rewrite them with `LITCHI_BLESS=1` after an intended change.
//! `test-data/unified-ops/0294/PROVENANCE.md` lists each file's generator.
//!
//! The fixtures are synthetic: they show that the adapter code of each
//! element works, not that a third-party producer's variant of it does.

#![cfg(all(
    feature = "doc",
    feature = "pptx",
    feature = "ppt",
    feature = "xlsb",
    feature = "odt",
    feature = "odp",
    feature = "odg",
    feature = "oth",
    feature = "odm"
))]

use std::io::Cursor;
use std::path::{Path, PathBuf};

use litchi_odf_common::core::PackageWriter;

#[path = "support/pptx_presentation.rs"]
#[allow(dead_code, reason = "only the full synthetic deck is used here")]
mod presentation;

const XMLNS: &str = concat!(
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"xmlns:dc="http://purl.org/dc/elements/1.1/" "#,
    r#"xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" "#,
    r#"office:version="1.3""#
);

/// A 4x2 PNG (the same image as `crud_matrix.rs`'s data URI).
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x02, 0x08, 0x06, 0x00, 0x00, 0x00, 0x7F, 0xA8, 0x7D,
    0x63, 0x00, 0x00, 0x00, 0x12, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0xF0,
    0x1F, 0x19, 0x33, 0xA0, 0x0B, 0x00, 0x00, 0x0F, 0x21, 0x0F, 0xF1, 0x04, 0x37, 0xC6, 0x9F, 0x00,
    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/unified-ops/0294")
}

/// A text body with two sibling frames in one paragraph, one holding a
/// picture and one a text box, and a second picture frame (the frame,
/// picture and text-box cells of OTH and ODM, with a sibling pair for move
/// and swap).
fn frames_body() -> &'static str {
    concat!(
        r#"<text:p text:style-name="Litchi">Frames"#,
        r#"<draw:frame draw:name="Picture1" text:anchor-type="as-char" svg:width="2cm" svg:height="1cm">"#,
        r#"<draw:image xlink:href="Pictures/litchi.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>"#,
        r#"<draw:frame draw:name="Box1" text:anchor-type="as-char" svg:width="4cm" svg:height="2cm">"#,
        r#"<draw:text-box><text:p>Inside the box</text:p></draw:text-box></draw:frame>"#,
        r#"<draw:frame draw:name="Picture2" text:anchor-type="as-char" svg:width="2cm" svg:height="1cm">"#,
        r#"<draw:image xlink:href="Pictures/litchi.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>"#,
        r#"</text:p><text:p>After</text:p>"#
    )
}

fn text_package(mimetype: &str, styles: bool) -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {XMLNS}><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
        frames_body()
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if styles {
        let styles = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles {XMLNS}><office:styles><style:style style:name="Standard" style:family="paragraph"/><style:style style:name="Litchi" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-weight="bold"/></style:style></office:styles></office:document-styles>"#
        );
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    }
    writer
        .add_file_with_media_type("Pictures/litchi.png", PNG, "image/png")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

/// A page with a page-level layer set and a shape on a layer.
fn odg_layers() -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {XMLNS}><office:body><office:drawing><draw:page draw:name="One"><draw:layer-set><draw:layer draw:name="layout"/><draw:layer draw:name="notes"/></draw:layer-set><draw:rect draw:name="Box" draw:layer="layout" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><text:p>Box</text:p></draw:rect></draw:page></office:drawing></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

/// A one-slide presentation from the ODP builder with two page comments
/// added through the leaf's annotation transaction.
fn odp_comments() -> Vec<u8> {
    use litchi_odp::annotation::{Anchor, Annotation};
    let mut shape = litchi_odp::Shape::new();
    shape.name = Some("Title".to_string());
    shape.text = "Comments".to_string();
    let slide = litchi_odp::Slide {
        title: None,
        text: "body".to_string(),
        index: 0,
        notes: None,
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: vec![shape],
    };
    let mut builder = litchi_odp::Builder::new();
    builder.add_slide_element(slide).unwrap();
    let mut package =
        litchi_odp::Package::from_bytes(&builder.build().unwrap().to_bytes().unwrap()).unwrap();
    for text in ["First comment", "Second comment"] {
        let mut edit = package.edit().unwrap();
        edit.add_annotation(&Anchor::page(0), &Annotation::new(text))
            .unwrap();
        package = edit.commit().unwrap().into_package();
    }
    package.to_bytes().unwrap()
}

/// One slide with two straight connectors and a plain line.
fn ppt_connectors() -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer.add_connector(slide, 20, 20, 200, 60).unwrap();
    writer.add_connector(slide, 20, 100, 200, 140).unwrap();
    writer.add_line(slide, 20, 180, 200, 180).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut bytes).unwrap();
    bytes.into_inner()
}

/// Record 0323's element deck: a text box with a shape and a text
/// hyperlink, a comment, a 2x2 table and a Microsoft Graph chart on the
/// first slide, and a second slide the links can jump to.
fn ppt_elements() -> Vec<u8> {
    use litchi_ppt::writer::{Chart, ChartKind, Hyperlink, SlideComment, Table};
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 20, 20, 400, 60, "Hello brave new world")
        .unwrap();
    let link = writer.add_hyperlink(Hyperlink::url("https://example.test/shape"));
    writer.set_last_shape_hyperlink(slide, link).unwrap();
    let text = writer.add_hyperlink(Hyperlink::url("https://example.test/text"));
    writer
        .set_last_shape_text_hyperlink(slide, litchi_ppt::TextRange::new(6, 11).unwrap(), text)
        .unwrap();
    writer
        .add_comment(
            slide,
            SlideComment::new("Ada Lovelace", "Check the figures", 10, 20),
        )
        .unwrap();
    let mut table = Table::new(2, 2).unwrap();
    for (row, texts) in [["A1", "B1"], ["A2", "B2"]].iter().enumerate() {
        for (column, text) in texts.iter().enumerate() {
            table.set_cell_text(row, column, *text).unwrap();
        }
    }
    writer.add_table(slide, 20, 120, table).unwrap();
    let mut chart = Chart::new(ChartKind::Bar);
    chart.set_title("Sales");
    chart.set_categories(["Q1", "Q2"]);
    chart.add_series(Some("2026"), vec![1.0, 2.0]).unwrap();
    writer.add_chart(slide, 300, 120, 200, 150, chart).unwrap();
    writer.add_slide().unwrap();
    let mut bytes = Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut bytes).unwrap();
    bytes.into_inner()
}

/// A worksheet with a header row and a three-column table over it.
fn xlsb_table() -> Vec<u8> {
    use litchi_xlsb::package::table::{Column, Range, Table, Type};
    let mut sheet = litchi_xlsb::writer::MutableWorksheet::new("Sheet1");
    for (column, name) in ["Item", "Count", "Price"].into_iter().enumerate() {
        sheet.set_cell(0, column as u32, name);
    }
    for row in 1..=3u32 {
        sheet.set_cell(row, 0, format!("Item {row}"));
        sheet.set_cell(row, 1, f64::from(row));
        sheet.set_cell(row, 2, f64::from(row) * 1.5);
    }
    sheet
        .add_table(Table {
            id: 1,
            name: Some("Items".to_string()),
            display_name: Some("Items".to_string()),
            range: Range {
                first_row: 0,
                last_row: 3,
                first_column: 0,
                last_column: 2,
            },
            table_type: Type::Range,
            header_row_count: 1,
            columns: ["Item", "Count", "Price"]
                .into_iter()
                .zip(1u32..)
                .map(|(name, id)| Column {
                    id,
                    name: Some(name.to_string()),
                    ..Column::default()
                })
                .collect(),
            ..Table::default()
        })
        .unwrap();
    let mut workbook = litchi_xlsb::Builder::new();
    workbook.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    workbook.build().unwrap().write_to(&mut bytes).unwrap();
    bytes.into_inner()
}

/// Record 0324: a DOC body paragraph with two tab stops, a floating text
/// box and a closing paragraph (the DOC `tab` and `textbox` cells, with a
/// sibling pair of tab stops).
fn doc_marks() -> Vec<u8> {
    use litchi_doc::parts::pap::{TabAlignment, TabLeader, TabStop};
    use litchi_doc::writer::{
        FloatingPosition, Kind as DrawingKind, ParagraphFormatting, Shape as DrawingShape,
    };
    let mut builder = litchi_doc::Builder::new();
    builder
        .add_formatted_paragraph(
            "Name\tValue",
            ParagraphFormatting {
                tab_stops_to_add: vec![
                    TabStop {
                        position: 2880,
                        alignment: TabAlignment::Left,
                        leader: TabLeader::None,
                    },
                    TabStop {
                        position: 5760,
                        alignment: TabAlignment::Right,
                        leader: TabLeader::Dots,
                    },
                ],
                ..ParagraphFormatting::default()
            },
        )
        .unwrap();
    builder
        .insert_floating_text_box(
            DrawingShape::new(DrawingKind::Rectangle, 3600, 1800).unwrap(),
            FloatingPosition::new(2500, 1200),
            "Boxed text",
        )
        .unwrap();
    builder.add_paragraph("after").unwrap();
    let mut bytes = Cursor::new(Vec::new());
    builder.build().unwrap().write_to(&mut bytes).unwrap();
    bytes.into_inner()
}

/// The namespaces of the ODT parts fixture, beyond [`XMLNS`].
const ODT_XMLNS: &str = concat!(
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#
);

/// A text document (of the ODT, OTH or ODM `mimetype`) with every part
/// record 0328 reads: a common and an
/// automatic list style, a list, tab stops, text fields, a table of
/// contents, a table with declared columns, a custom shape and a
/// rectangle, a formula subdocument and an inline formula, and two master
/// pages with headers and footers.
fn odt_parts(mimetype: &str) -> Vec<u8> {
    let content = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {xmlns} {odt}>"#,
            r#"<office:automatic-styles>"#,
            r#"<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties><style:tab-stops><style:tab-stop style:position="2.5cm"/><style:tab-stop style:position="10cm" style:type="right" style:leader-style="dotted" style:leader-text="."/></style:tab-stops></style:paragraph-properties></style:style>"#,
            r#"<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>"#,
            r#"<style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="4.5cm"/></style:style>"#,
            r##"<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="solid" svg:stroke-color="#0000ff" svg:stroke-width="0.05cm" style:wrap="none"/></style:style>"##,
            r#"<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:margin-left="1.27cm" fo:text-indent="-0.635cm"/></style:list-level-properties><style:text-properties style:font-name="OpenSymbol"/></text:list-level-style-bullet></text:list-style>"#,
            r#"</office:automatic-styles><office:body><office:text>"#,
            r#"<text:h text:outline-level="1">Parts</text:h>"#,
            r#"<text:table-of-content text:name="Contents1"><text:table-of-content-source text:outline-level="3"><text:index-title-template>Contents</text:index-title-template><text:table-of-content-entry-template text:outline-level="1"><text:index-entry-link-start/><text:index-entry-text/><text:index-entry-tab-stop style:type="right"/><text:index-entry-page-number/><text:index-entry-link-end/></text:table-of-content-entry-template></text:table-of-content-source><text:index-body><text:p>Parts<text:tab/>1</text:p></text:index-body></text:table-of-content>"#,
            r#"<text:p text:style-name="P1">Name<text:tab/>Value</text:p>"#,
            r#"<text:list text:style-name="Numbering_20_123"><text:list-item><text:p>First</text:p></text:list-item><text:list-item><text:p>Second</text:p></text:list-item></text:list>"#,
            r#"<text:list text:style-name="L1"><text:list-item><text:p>Bullet</text:p></text:list-item></text:list>"#,
            r#"<text:p>Page <text:page-number text:select-page="current">1</text:page-number> of <text:page-count>2</text:page-count>, <text:date text:fixed="true" text:date-value="2026-10-17">17/10/2026</text:date>, <text:title>Parts</text:title>, <text:chapter text:display="name" text:outline-level="1">Parts</text:chapter></text:p>"#,
            r#"<table:table table:name="Grid"><table:table-column table:style-name="co1" table:number-columns-repeated="2"/><table:table-column table:style-name="co2"/><table:table-row><table:table-cell><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row></table:table>"#,
            r#"<text:p>Shapes<draw:custom-shape draw:name="Star" draw:style-name="gr1" text:anchor-type="char" svg:x="1cm" svg:y="0.5cm" svg:width="3cm" svg:height="2cm"><text:p>Starred</text:p><draw:enhanced-geometry draw:type="star5"/></draw:custom-shape><draw:rect draw:name="Box" text:anchor-type="paragraph" svg:x="5cm" svg:y="1cm" svg:width="2cm" svg:height="1cm"/></text:p>"#,
            r#"<text:p>Formula <draw:frame draw:name="Formula1" text:anchor-type="as-char" svg:width="3cm" svg:height="0.8cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame> and <draw:frame draw:name="Formula2" text:anchor-type="as-char" svg:width="2cm" svg:height="0.6cm"><draw:object><math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><semantics><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><annotation encoding="StarMath 5.0">x + 1</annotation></semantics></math></draw:object></draw:frame></text:p>"#,
            r#"<text:p>End</text:p>"#,
            r#"</office:text></office:body></office:document-content>"#
        ),
        xmlns = XMLNS,
        odt = ODT_XMLNS,
    );
    let styles = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles {xmlns} {odt}>"#,
            r#"<office:styles><style:style style:name="Standard" style:family="paragraph"/>"#,
            r#"<style:style style:name="Header" style:family="paragraph" style:parent-style-name="Standard"/>"#,
            r##"<text:list-style style:name="Numbering_20_123" style:display-name="Numbering 123"><text:list-level-style-number text:level="1" style:num-prefix="(" style:num-suffix=")" style:num-format="1" text:start-value="3"><style:list-level-properties fo:text-align="end" text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="space" fo:margin-left="0.635cm" fo:text-indent="-0.635cm"/></style:list-level-properties><style:text-properties fo:font-weight="bold" fo:color="#00ff00" fo:font-size="12pt"/></text:list-level-style-number><text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix="." text:display-levels="2"><style:list-level-properties text:space-before="1.27cm" text:min-label-width="0.635cm"/></text:list-level-style-number></text:list-style>"##,
            r#"</office:styles><office:master-styles>"#,
            r#"<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p text:style-name="Header">Top <text:page-number text:select-page="current">1</text:page-number></text:p></style:header><style:header-left><text:p>Even top</text:p></style:header-left><style:footer><text:p>Bottom</text:p></style:footer></style:master-page>"#,
            r#"<style:master-page style:name="Landscape" style:page-layout-name="pm2"><style:header-first><text:p>First top</text:p></style:header-first><style:footer style:display="false"><text:p>Hidden bottom</text:p></style:footer></style:master-page>"#,
            r#"</office:master-styles></office:document-styles>"#
        ),
        xmlns = XMLNS,
        odt = ODT_XMLNS,
    );
    let formula = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><math xmlns="http://www.w3.org/1998/Math/MathML" display="block">"#,
        r#"<semantics><mrow><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></mrow>"#,
        r#"<annotation encoding="StarMath 5.0">a^2 + b^2 = c^2</annotation></semantics></math>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    writer
        .add_manifest_directory("Object 1/", "application/vnd.oasis.opendocument.formula")
        .unwrap();
    writer
        .add_file_with_media_type("Object 1/content.xml", formula.as_bytes(), "text/xml")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

/// The namespaces of the ODP elements fixture, beyond [`XMLNS`].
const ODP_XMLNS: &str = concat!(
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0""#
);

/// A two-slide presentation with every element record 0328 reads: a text
/// hyperlink with a title, line breaks, shape click actions (a URL, the
/// next page, a `draw:a` link), a table with declared columns, a video and
/// an audio plugin, a formula subdocument and an inline formula, and a chart
/// subdocument with two axes and two series.
fn odp_elements() -> Vec<u8> {
    let frame = |name: &str, y: u32| {
        format!(r#"draw:name="{name}" svg:x="1cm" svg:y="{y}cm" svg:width="6cm" svg:height="1cm""#)
    };
    let content = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {xmlns} {odp}>"#,
            r#"<office:automatic-styles>"#,
            r#"<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="4cm"/></style:style>"#,
            r#"<style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="2.5cm"/></style:style>"#,
            r#"</office:automatic-styles><office:body><office:presentation>"#,
            r#"<draw:page draw:name="One" draw:master-page-name="Default">"#,
            r#"<draw:frame {text}><draw:text-box>"#,
            r#"<text:p>Visit <text:a xlink:type="simple" xlink:href="https://example.com/" office:title="Example">example</text:a> now</text:p>"#,
            r#"<text:p>Line one<text:line-break/>Line two<text:line-break/>Line three</text:p>"#,
            r#"</draw:text-box></draw:frame>"#,
            r#"<draw:custom-shape {button}><office:event-listeners><presentation:event-listener script:event-name="dom:click" presentation:action="show" xlink:href="https://example.org/" xlink:type="simple" xlink:show="embed" xlink:actuate="onRequest"/></office:event-listeners><text:p>Go</text:p><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>"#,
            r#"<draw:custom-shape {next}><office:event-listeners><presentation:event-listener script:event-name="dom:click" presentation:action="next-page"/></office:event-listeners><text:p>Next</text:p><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>"#,
            r#"<draw:a xlink:type="simple" xlink:href="https://example.net/"><draw:frame {linked}><draw:text-box><text:p>Linked box</text:p></draw:text-box></draw:frame></draw:a>"#,
            r#"<draw:custom-shape {plain}><text:p>Plain</text:p><draw:enhanced-geometry draw:type="ellipse"/></draw:custom-shape>"#,
            r#"<draw:frame {grid}><table:table><table:table-column table:style-name="co1" table:number-columns-repeated="2"/><table:table-column table:style-name="co2"/><table:table-row><table:table-cell><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row></table:table></draw:frame>"#,
            r#"<draw:frame {movie}><draw:plugin xlink:href="Media/clip.mp4" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="application/vnd.sun.star.media"><draw:param draw:name="Loop" draw:value="true"/><draw:param draw:name="Mute" draw:value="false"/></draw:plugin></draw:frame>"#,
            r#"<draw:frame {sound}><draw:plugin xlink:href="Media/tone.wav" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="audio/x-wav"/></draw:frame>"#,
            r#"<draw:frame {formula}><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>"#,
            r#"<draw:frame {inline}><draw:object><math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mrow><mi>y</mi><mo>=</mo><mn>2</mn></mrow><annotation encoding="StarMath 5.0">y = 2</annotation></semantics></math></draw:object></draw:frame>"#,
            r#"<draw:frame presentation:class="chart" {chart}><draw:object xlink:href="./Object 2" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>"#,
            r#"</draw:page>"#,
            r#"<draw:page draw:name="Two" draw:master-page-name="Default"><draw:frame {title}><draw:text-box><text:p>Second</text:p></draw:text-box></draw:frame></draw:page>"#,
            r#"</office:presentation></office:body></office:document-content>"#
        ),
        xmlns = XMLNS,
        odp = ODP_XMLNS,
        text = frame("Text", 1),
        button = frame("Button", 3),
        next = frame("Next", 5),
        linked = frame("Linked", 7),
        plain = frame("Plain", 9),
        grid = frame("Grid", 11),
        movie = frame("Movie", 13),
        sound = frame("Sound", 15),
        formula = frame("Formula", 17),
        inline = frame("InlineFormula", 18),
        chart = frame("Chart", 19),
        title = frame("Title2", 1),
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles {XMLNS}><office:styles/><office:master-styles><style:master-page style:name="Default"/></office:master-styles></office:document-styles>"#
    );
    let formula = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><math xmlns="http://www.w3.org/1998/Math/MathML" display="block">"#,
        r#"<semantics><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow>"#,
        r#"<annotation encoding="StarMath 5.0">E = m c^2</annotation></semantics></math>"#
    );
    let chart = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" "#,
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3">"#,
        r#"<office:body><office:chart><chart:chart chart:class="chart:bar" svg:width="8cm" svg:height="6cm">"#,
        r#"<chart:title><text:p>Sales</text:p></chart:title>"#,
        r#"<chart:plot-area table:cell-range-address="local-table.$A$1:.$C$3">"#,
        r#"<chart:axis chart:dimension="x" chart:name="primary-x"><chart:title><text:p>Quarter</text:p></chart:title><chart:categories table:cell-range-address="local-table.$A$2:.$A$3"/></chart:axis>"#,
        r#"<chart:axis chart:dimension="y" chart:name="primary-y"/>"#,
        r#"<chart:series chart:values-cell-range-address="local-table.$B$2:.$B$3" chart:label-cell-address="local-table.$B$1" chart:class="chart:bar"><chart:data-point chart:repeated="2"/></chart:series>"#,
        r#"<chart:series chart:values-cell-range-address="local-table.$C$2:.$C$3" chart:label-cell-address="local-table.$C$1"/>"#,
        r#"</chart:plot-area></chart:chart></office:chart></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    writer
        .add_manifest_directory("Object 1/", "application/vnd.oasis.opendocument.formula")
        .unwrap();
    writer
        .add_file_with_media_type("Object 1/content.xml", formula.as_bytes(), "text/xml")
        .unwrap();
    writer
        .add_manifest_directory("Object 2/", "application/vnd.oasis.opendocument.chart")
        .unwrap();
    writer
        .add_file_with_media_type("Object 2/content.xml", chart.as_bytes(), "text/xml")
        .unwrap();
    writer
        .add_file_with_media_type("Media/clip.mp4", b"litchi video", "video/mp4")
        .unwrap();
    writer
        .add_file_with_media_type("Media/tone.wav", b"litchi audio", "audio/x-wav")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

/// Every generated fixture: file name and bytes.
fn fixtures() -> Vec<(&'static str, Vec<u8>)> {
    vec![
        ("modern-comments.pptx", presentation::bytes()),
        ("connectors.ppt", ppt_connectors()),
        ("marks.doc", doc_marks()),
        ("elements.ppt", ppt_elements()),
        ("table.xlsb", xlsb_table()),
        ("comments.odp", odp_comments()),
        ("layers.odg", odg_layers()),
        ("parts.odp", odp_elements()),
        (
            "parts.odt",
            odt_parts("application/vnd.oasis.opendocument.text"),
        ),
        (
            "parts.oth",
            odt_parts("application/vnd.oasis.opendocument.text-web"),
        ),
        (
            "parts.odm",
            odt_parts("application/vnd.oasis.opendocument.text-master"),
        ),
        (
            "frames.oth",
            text_package("application/vnd.oasis.opendocument.text-web", true),
        ),
        (
            "frames.odm",
            text_package("application/vnd.oasis.opendocument.text-master", false),
        ),
    ]
}

#[test]
fn generated_fixtures_are_current() {
    let dir = dir();
    let bless = std::env::var_os("LITCHI_BLESS").is_some();
    let mut stale = Vec::new();
    for (name, bytes) in fixtures() {
        // Two generations are byte-identical (ADR 0006), so the check below
        // is meaningful.
        let again = fixtures()
            .into_iter()
            .find(|(other, _)| *other == name)
            .unwrap()
            .1;
        assert!(again == bytes, "{name}: generation is not deterministic");
        litchi_ops::open_bytes(bytes.clone()).unwrap_or_else(|error| panic!("{name}: {error}"));
        let path = dir.join(name);
        if bless {
            std::fs::create_dir_all(&dir).unwrap();
            std::fs::write(&path, &bytes).unwrap();
        } else if std::fs::read(&path).ok().as_deref() != Some(bytes.as_slice()) {
            stale.push(name);
        }
    }
    assert!(
        stale.is_empty(),
        "generated fixtures differ from test-data/unified-ops/0294 ({stale:?}); \
         review and bless with LITCHI_BLESS=1"
    );
}
