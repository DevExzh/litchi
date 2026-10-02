//! Phase-3 (P2) leniency acceptance tests for WordprocessingML reads.
//!
//! Each finding patches a real test-data package to carry the deviation and
//! asserts that it opens under the default posture through the owning and
//! source-backed facades with the repair recorded; the unmodified fixture is
//! the conformant control and records nothing. Hard errors stay hard.

use std::io;
use std::path::Path;

use litchi_docx::{Defect, NumberingDefect, Package, StyleDefect, source_backed};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

fn fixture(name: &str) -> Vec<u8> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/docx")
        .join(name);
    std::fs::read(&path).unwrap_or_else(|error| panic!("read {}: {error}", path.display()))
}

fn member(bytes: &[u8], name: &str) -> String {
    let archive = ArchiveReader::new(bytes).expect("fixture archive");
    String::from_utf8(archive.read(name).expect("archive member")).expect("UTF-8 member")
}

/// Rewrite `name` in the archive by replacing the first `from` with `to`.
fn patch(bytes: &[u8], name: &str, from: &str, to: &str) -> Vec<u8> {
    let archive = ArchiveReader::new(bytes).expect("fixture archive");
    let mut writer = StreamingArchiveWriter::new();
    let mut patched = false;
    for member in archive.file_names() {
        let mut content = archive.read(member).expect("archive member");
        if member == name {
            let text = String::from_utf8(content).expect("UTF-8 member");
            assert!(text.contains(from), "{name} contains {from}");
            content = text.replacen(from, to, 1).into_bytes();
            patched = true;
        }
        writer.write_stored(member, &content).expect("write member");
    }
    assert!(patched, "fixture has {name}");
    writer.finish_to_bytes().expect("finish archive")
}

/// The markup-compatibility processed XML the typed codecs read.
fn visible(xml: &str) -> String {
    litchi_ooxml_common::mce::process_str(xml)
        .expect("MCE processing")
        .into_owned()
}

fn open(bytes: &[u8]) -> Package {
    Package::from_reader(io::Cursor::new(bytes.to_vec())).expect("package opens")
}

fn open_source_backed(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_reader(io::Cursor::new(bytes.to_vec()))
        .expect("source-backed package opens")
}

fn defects(report: &litchi_docx::ToleranceReport) -> Vec<(Defect, u64)> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect()
}

// ---------------------------------------------------------------------------
// F2.9 — list level (`w:ilvl`) above 8; style outline level above 9
// ---------------------------------------------------------------------------

const LISTS: &str = "ComplexNumberedLists.docx";
const FIRST_LEVEL: &str = r#"<w:ilvl w:val="1"/><w:numId w:val="1"/>"#;
const LIST_PARAGRAPH_PPR: &str = r#"<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>"#;

#[test]
fn deep_paragraph_level_is_clamped_and_recorded() {
    let bytes = patch(
        &fixture(LISTS),
        "word/document.xml",
        FIRST_LEVEL,
        r#"<w:ilvl w:val="12"/><w:numId w:val="1"/>"#,
    );
    let package = open(&bytes);
    let document = package.document().expect("document");
    let paragraphs = document.paragraphs().expect("paragraphs");
    let index = paragraphs
        .iter()
        .position(|paragraph| paragraph.text().expect("text") == "2-a")
        .expect("2-a paragraph");
    let numbering = paragraphs[index]
        .numbering()
        .expect("ilvl 12 no longer fails numbering")
        .expect("direct numbering");
    assert_eq!((numbering.num_id, numbering.level), (1, 8));
    assert_eq!(
        paragraphs[index].style_id().expect("style id").as_deref(),
        Some("ListParagraph")
    );

    let items = document.list_items().expect("list items resolve");
    assert!(items.len() > 3);
    assert_eq!(
        defects(&document.tolerance_report()),
        [(
            Defect::Numbering(NumberingDefect::LevelOutOfRange),
            index as u64
        )]
    );

    // The source-backed facade reads the same clamped level, and its owned
    // materialization resolves and records the list.
    let source = open_source_backed(&bytes);
    let pinned = source.document().expect("pinned document");
    let numbering = pinned.paragraphs().expect("paragraphs")[index]
        .numbering()
        .expect("numbering")
        .expect("direct numbering");
    assert_eq!(numbering.level, 8);
    let owned = source.to_owned_package().expect("owned");
    let owned_document = owned.document().expect("document");
    owned_document.list_items().expect("list items");
    assert_eq!(
        defects(&owned_document.tolerance_report()),
        [(
            Defect::Numbering(NumberingDefect::LevelOutOfRange),
            index as u64
        )]
    );
}

#[test]
fn deep_style_levels_are_repaired_and_recorded() {
    let original = fixture(LISTS);
    let replacement = r#"<w:pPr><w:numPr><w:ilvl w:val="9"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="12"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>"#;
    let bytes = patch(
        &original,
        "word/styles.xml",
        LIST_PARAGRAPH_PPR,
        replacement,
    );
    // The "2-a" paragraph drops its direct numbering, so list resolution
    // inherits the style's (clamped) numbering and loads the style table.
    let bytes = patch(
        &bytes,
        "word/document.xml",
        r#"<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>"#,
        "",
    );
    // Offsets are within the visible (markup-compatibility processed) XML,
    // which expands empty elements.
    let styles_xml = visible(&member(&bytes, "word/styles.xml"));
    let ilvl = styles_xml.find(r#"<w:ilvl w:val="9">"#).expect("ilvl") as u64;
    let outline = styles_xml
        .find(r#"<w:outlineLvl w:val="12">"#)
        .expect("outline") as u64;

    let package = open(&bytes);
    let document = package.document().expect("document");
    let mut styles = document.styles().expect("styles");
    let style = styles
        .get_by_id("ListParagraph")
        .expect("style table loads")
        .expect("style");
    assert_eq!(style.numbering().map(|value| value.level), Some(8));
    assert_eq!(style.outline(), None);
    let expected = [
        (Defect::Style(StyleDefect::LevelOutOfRange), ilvl),
        (Defect::Style(StyleDefect::OutlineLevelOutOfRange), outline),
    ];
    assert_eq!(
        defects(styles.tolerance_report().expect("report")),
        expected
    );

    let items = document.list_items().expect("list items resolve");
    let inherited = items
        .iter()
        .find(|item| item.text == "2-a")
        .expect("2-a stays numbered through its style");
    assert_eq!(inherited.numbering.level, 8);
    assert_eq!(defects(&document.tolerance_report()), expected);

    // The source-backed package's owned materialization reports the same.
    let owned = open_source_backed(&bytes)
        .to_owned_package()
        .expect("owned");
    let owned_document = owned.document().expect("document");
    owned_document.list_items().expect("list items");
    assert_eq!(defects(&owned_document.tolerance_report()), expected);
}

#[test]
fn body_text_outline_level_is_conformant() {
    let bytes = patch(
        &fixture(LISTS),
        "word/styles.xml",
        LIST_PARAGRAPH_PPR,
        r#"<w:pPr><w:outlineLvl w:val="9"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>"#,
    );
    let package = open(&bytes);
    let document = package.document().expect("document");
    let mut styles = document.styles().expect("styles");
    assert_eq!(
        styles
            .get_by_id("ListParagraph")
            .expect("styles load")
            .expect("style")
            .outline(),
        None
    );
    assert!(styles.tolerance_report().expect("report").is_clean());
}

#[test]
fn conformant_list_fixture_records_nothing() {
    let bytes = fixture(LISTS);
    let package = open(&bytes);
    let document = package.document().expect("document");
    document.list_items().expect("list items");
    assert!(document.tolerance_report().is_clean());
    let source = open_source_backed(&bytes);
    assert!(
        source
            .document()
            .expect("pinned")
            .tolerance_report()
            .is_clean()
    );
}

#[test]
fn non_numeric_levels_stay_fatal() {
    let bytes = patch(
        &fixture(LISTS),
        "word/document.xml",
        FIRST_LEVEL,
        r#"<w:ilvl w:val="deep"/><w:numId w:val="1"/>"#,
    );
    let package = open(&bytes);
    let document = package.document().expect("document");
    assert!(document.list_items().is_err());

    let bytes = patch(
        &fixture(LISTS),
        "word/styles.xml",
        LIST_PARAGRAPH_PPR,
        r#"<w:pPr><w:outlineLvl w:val="-1"/></w:pPr>"#,
    );
    let package = open(&bytes);
    let mut styles = package
        .document()
        .expect("document")
        .styles()
        .expect("styles");
    assert!(styles.len().is_err());
}
