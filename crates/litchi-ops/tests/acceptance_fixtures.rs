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
    feature = "pptx",
    feature = "ppt",
    feature = "xlsb",
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

/// Every generated fixture: file name and bytes.
fn fixtures() -> Vec<(&'static str, Vec<u8>)> {
    vec![
        ("modern-comments.pptx", presentation::bytes()),
        ("connectors.ppt", ppt_connectors()),
        ("table.xlsb", xlsb_table()),
        ("comments.odp", odp_comments()),
        ("layers.odg", odg_layers()),
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
