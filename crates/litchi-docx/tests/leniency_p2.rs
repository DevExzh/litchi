//! Phase-3 (P2) leniency acceptance tests for WordprocessingML reads.
//!
//! Each finding patches a real test-data package to carry the deviation and
//! asserts that it opens under the default posture through the owning and
//! source-backed facades with the repair recorded; the unmodified fixture is
//! the conformant control and records nothing. Hard errors stay hard.

use std::io;
use std::path::Path;

use litchi_docx::{
    Defect, DrawingDefect, FieldDefect, NumberingDefect, Package, SectionDefect, StyleDefect,
    source_backed,
};
use litchi_drawingml::geom::Preset;
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
    patch_with(bytes, name, |text| {
        assert!(text.contains(from), "{name} contains {from}");
        text.replacen(from, to, 1)
    })
}

/// Rewrite the text of archive member `name` with `rewrite`.
fn patch_with(bytes: &[u8], name: &str, rewrite: impl FnOnce(String) -> String) -> Vec<u8> {
    let archive = ArchiveReader::new(bytes).expect("fixture archive");
    let mut writer = StreamingArchiveWriter::new();
    let mut rewrite = Some(rewrite);
    for member in archive.file_names() {
        let mut content = archive.read(member).expect("archive member");
        if member == name {
            let text = String::from_utf8(content).expect("UTF-8 member");
            content = rewrite.take().expect("member once")(text).into_bytes();
        }
        writer.write_stored(member, &content).expect("write member");
    }
    assert!(rewrite.is_none(), "fixture has {name}");
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

// ---------------------------------------------------------------------------
// F2.10 — unknown `a:prstGeom@prst`
// ---------------------------------------------------------------------------

const SHAPES: &str = "../../libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx";

/// The shape fixture with its `wps` choice unwrapped from
/// `mc:AlternateContent` (dropping the VML fallback), the way third-party
/// producers write a Word shape; the markup-compatibility reader would
/// otherwise select the fallback.
fn shape_fixture() -> Vec<u8> {
    patch_with(&fixture(SHAPES), "word/document.xml", |xml| {
        let open = xml
            .find("<mc:AlternateContent>")
            .expect("alternate content");
        let choice = "<mc:Choice Requires=\"wps\">";
        let body = xml[open..].find(choice).expect("wps choice") + open + choice.len();
        let close = xml.find("</mc:Choice>").expect("choice end");
        let end_tag = "</mc:AlternateContent>";
        let end = xml.find(end_tag).expect("alternate content end") + end_tag.len();
        format!("{}{}{}", &xml[..open], &xml[body..close], &xml[end..])
    })
}

fn shape_presets(
    objects: &[litchi_docx::drawing::Object],
) -> Vec<(Option<Preset>, Option<String>)> {
    objects
        .iter()
        .filter(|object| object.kind() != litchi_docx::drawing::Kind::Other)
        .map(|object| (object.preset(), object.unknown_preset().map(str::to_owned)))
        .collect()
}

#[test]
fn unknown_shape_preset_is_kept_as_raw_geometry_and_recorded() {
    let original = shape_fixture();
    let control = open(&original);
    let control_document = control.document().expect("document");
    let control_objects = control_document.drawing_objects().expect("drawings");
    assert_eq!(
        shape_presets(&control_objects),
        [(Some(Preset::Rect), None)]
    );
    let index = control_document
        .paragraphs()
        .expect("paragraphs")
        .iter()
        .position(|paragraph| !paragraph.drawing_objects().expect("drawings").is_empty())
        .expect("shape paragraph") as u64;
    assert!(control_document.tolerance_report().is_clean());

    let bytes = patch(
        &original,
        "word/document.xml",
        r#"<a:prstGeom prst="rect""#,
        r#"<a:prstGeom prst="customShape""#,
    );
    let expected = [(Defect::Drawing(DrawingDefect::UnknownPreset), index)];

    let package = open(&bytes);
    let document = package.document().expect("document");
    let objects = document
        .drawing_objects()
        .expect("drawing inventory survives");
    assert_eq!(objects.len(), control_objects.len());
    assert_eq!(
        shape_presets(&objects),
        [(None, Some("customShape".to_owned()))]
    );
    assert_eq!(objects[0].name(), control_objects[0].name());
    assert_eq!(defects(&document.tolerance_report()), expected);

    let source = open_source_backed(&bytes);
    let pinned = source.document().expect("pinned document");
    let objects = pinned
        .drawing_objects()
        .expect("drawing inventory survives");
    assert_eq!(
        shape_presets(&objects),
        [(None, Some("customShape".to_owned()))]
    );
    assert_eq!(defects(&pinned.tolerance_report()), expected);

    let source_control = open_source_backed(&original);
    let pinned = source_control.document().expect("pinned document");
    pinned.drawing_objects().expect("drawings");
    assert!(pinned.tolerance_report().is_clean());
}

#[test]
fn missing_shape_preset_is_kept_and_recorded() {
    let bytes = patch(
        &shape_fixture(),
        "word/document.xml",
        r#"<a:prstGeom prst="rect""#,
        "<a:prstGeom",
    );
    let package = open(&bytes);
    let document = package.document().expect("document");
    let objects = document
        .drawing_objects()
        .expect("drawing inventory survives");
    assert_eq!(shape_presets(&objects), [(None, Some(String::new()))]);
    assert_eq!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .map(|tolerated| tolerated.defect)
            .collect::<Vec<_>>(),
        [Defect::Drawing(DrawingDefect::MissingPreset)]
    );
}

#[test]
fn malformed_shape_preset_attribute_stays_fatal() {
    let bytes = patch(
        &shape_fixture(),
        "word/document.xml",
        r#"<a:prstGeom prst="rect""#,
        r#"<a:prstGeom prst="rect" prst="rect""#,
    );
    // A duplicated attribute is ill-formed XML: the document does not open.
    assert!(open(&bytes).document().is_err());
    assert!(open_source_backed(&bytes).document().is_err());
}

// ---------------------------------------------------------------------------
// F2.11 — `w:fldSimple` without `w:instr`
// ---------------------------------------------------------------------------

const MAIL_MERGE: &str = "../../libreoffice-core/sw/qa/extras/ooxmlexport/data/mailmerge.docx";
const MERGE_FIELD: &str = r#"<w:fldSimple w:instr=" MERGEFIELD Name ">"#;

fn field_summary(fields: &[litchi_docx::Field]) -> Vec<(String, Option<String>)> {
    fields
        .iter()
        .map(|field| {
            (
                field.instruction().to_owned(),
                field.result().map(str::to_owned),
            )
        })
        .collect()
}

#[test]
fn simple_field_without_instruction_is_kept_and_recorded() {
    let original = fixture(MAIL_MERGE);
    // A well-formed PAGE field, then the merge field stripped of w:instr.
    let bytes = patch(
        &original,
        "word/document.xml",
        MERGE_FIELD,
        r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:fldSimple>"#,
    );
    let expected_summary = vec![
        ("PAGE".to_owned(), Some("1".to_owned())),
        (String::new(), Some("«Name»".to_owned())),
    ];

    let package = open(&bytes);
    let document = package.document().expect("document");
    let fields = document.fields().expect("field listing survives");
    assert_eq!(field_summary(&fields), expected_summary);
    let recorded = document.tolerance_report();
    assert_eq!(
        recorded
            .defects()
            .iter()
            .map(|tolerated| tolerated.defect)
            .collect::<Vec<_>>(),
        [Defect::Field(FieldDefect::MissingInstruction)]
    );
    // Offsets are within the visible (markup-compatibility processed) XML.
    let xml = visible(&member(&bytes, "word/document.xml"));
    let bad = xml.find("<w:fldSimple>").expect("bad field") as u64;
    assert_eq!(recorded.defects()[0].context, bad);
    // Typed projections over the listing read past the empty instruction.
    document.hyperlink_fields().expect("hyperlink fields");
    assert_eq!(document.field_count().expect("count"), 2);

    let source = open_source_backed(&bytes);
    let pinned = source.document().expect("pinned document");
    let fields = pinned.fields().expect("field listing survives");
    assert_eq!(field_summary(&fields), expected_summary);
    assert_eq!(
        pinned
            .tolerance_report()
            .defects()
            .iter()
            .map(|tolerated| tolerated.defect)
            .collect::<Vec<_>>(),
        [Defect::Field(FieldDefect::MissingInstruction)]
    );
    assert_eq!(pinned.tolerance_report().defects()[0].context, bad);
}

#[test]
fn conformant_simple_fields_record_nothing() {
    let original = fixture(MAIL_MERGE);
    let package = open(&original);
    let document = package.document().expect("document");
    assert_eq!(
        field_summary(&document.fields().expect("fields")),
        [("MERGEFIELD Name".to_owned(), Some("«Name»".to_owned()))]
    );
    assert!(document.tolerance_report().is_clean());
    let source = open_source_backed(&original);
    let pinned = source.document().expect("pinned");
    pinned.fields().expect("fields");
    assert!(pinned.tolerance_report().is_clean());
}

#[test]
fn unbalanced_simple_field_stays_fatal() {
    let bytes = patch(
        &fixture(MAIL_MERGE),
        "word/document.xml",
        MERGE_FIELD,
        r#"</w:fldSimple><w:fldSimple w:instr=" MERGEFIELD Name ">"#,
    );
    // Ill-formed XML: an end tag with no start.
    assert!(
        open(&bytes)
            .document()
            .and_then(|document| document.fields())
            .is_err()
    );
}

// ---------------------------------------------------------------------------
// F2.12 — header/footer relationship ID that is not an XML ID
// ---------------------------------------------------------------------------

const HEADERS: &str = "headerFooter.docx";
const DEFAULT_HEADER: &str = r#"<w:headerReference w:type="default" r:id="rId5"/>"#;
const UNUSED_MCE_BINDING: &str =
    r#" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006""#;

/// `headerFooter.docx` without its unused markup-compatibility namespace
/// binding, which the source-backed section inventory refuses outright.
fn header_fixture() -> Vec<u8> {
    patch(
        &fixture(HEADERS),
        "word/document.xml",
        UNUSED_MCE_BINDING,
        "",
    )
}

/// The header fixture with the default header reference written as the
/// numeric ID `5`. (The OPC reader itself requires XML IDs in `.rels`, so
/// the reference cannot name an existing relationship.)
fn numeric_header_id() -> Vec<u8> {
    patch(
        &header_fixture(),
        "word/document.xml",
        DEFAULT_HEADER,
        r#"<w:headerReference w:type="default" r:id="5"/>"#,
    )
}

fn section_defects(report: &litchi_docx::ToleranceReport) -> Vec<Defect> {
    report
        .defects()
        .iter()
        .map(|tolerated| tolerated.defect)
        .collect()
}

#[test]
fn numeric_header_relationship_id_is_kept_and_recorded() {
    let bytes = numeric_header_id();
    let xml = visible(&member(&bytes, "word/document.xml"));
    let reference = xml
        .find(r#"<w:headerReference w:type="default" r:id="5">"#)
        .or_else(|| xml.find(r#"<w:headerReference w:type="default" r:id="5"/>"#))
        .expect("reference") as u64;
    let expected = [(
        Defect::Section(SectionDefect::NonConformingRelationshipId),
        reference,
    )];

    let package = open(&bytes);
    let document = package.document().expect("document");
    let inventory = document
        .section_inventory()
        .expect("a numeric r:id no longer fails the section inventory");
    let ids = inventory.sections()[0]
        .headers()
        .iter()
        .map(|header| header.relationship_id.clone())
        .collect::<Vec<_>>();
    assert!(ids.contains(&"5".to_owned()), "{ids:?}");
    assert_eq!(defects(inventory.tolerance_report()), expected);
    assert_eq!(defects(&document.tolerance_report()), expected);

    let source = open_source_backed(&bytes);
    let pinned = source.document().expect("pinned document");
    let inventory = pinned.section_inventory().expect("section inventory");
    assert_eq!(
        section_defects(inventory.tolerance_report()),
        [Defect::Section(SectionDefect::NonConformingRelationshipId)]
    );
    assert_eq!(
        section_defects(&pinned.tolerance_report()),
        [Defect::Section(SectionDefect::NonConformingRelationshipId)]
    );
}

#[test]
fn conformant_header_relationship_ids_record_nothing() {
    let bytes = header_fixture();
    let package = open(&bytes);
    let document = package.document().expect("document");
    document.section_inventory().expect("sections");
    assert!(document.tolerance_report().is_clean());
    let source = open_source_backed(&bytes);
    let pinned = source.document().expect("pinned");
    pinned.section_inventory().expect("sections");
    assert!(pinned.tolerance_report().is_clean());
}

#[test]
fn empty_header_relationship_id_stays_fatal() {
    let bytes = patch(
        &header_fixture(),
        "word/document.xml",
        DEFAULT_HEADER,
        r#"<w:headerReference w:type="default" r:id=""/>"#,
    );
    let package = open(&bytes);
    assert!(
        package
            .document()
            .expect("document")
            .section_inventory()
            .is_err()
    );
}

// ---------------------------------------------------------------------------
// F2.6 follow-up — style-level `w:numPr` without `w:numId`
// ---------------------------------------------------------------------------

/// `ComplexNumberedLists.docx` with numbering on the `ListParagraph` style,
/// a `ListDeep` style based on it whose `w:numPr` carries only `w:ilvl`, and
/// the "2-a" paragraph styled `ListDeep` without direct numbering.
fn ilvl_only_style_fixture(deep_num_pr: &str) -> Vec<u8> {
    let bytes = patch(
        &fixture(LISTS),
        "word/styles.xml",
        LIST_PARAGRAPH_PPR,
        r#"<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>"#,
    );
    let bytes = patch(
        &bytes,
        "word/styles.xml",
        "</w:styles>",
        &format!(
            r#"<w:style w:type="paragraph" w:styleId="ListDeep"><w:name w:val="List Deep"/><w:basedOn w:val="ListParagraph"/><w:pPr>{deep_num_pr}</w:pPr></w:style></w:styles>"#
        ),
    );
    patch(
        &bytes,
        "word/document.xml",
        r#"<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>2-a"#,
        r#"<w:pPr><w:pStyle w:val="ListDeep"/></w:pPr><w:r><w:t>2-a"#,
    )
}

#[test]
fn ilvl_only_style_numbering_is_inherited_and_recorded() {
    let bytes = ilvl_only_style_fixture(r#"<w:numPr><w:ilvl w:val="2"/></w:numPr>"#);
    let styles_xml = visible(&member(&bytes, "word/styles.xml"));
    let deep = styles_xml.find(r#"w:styleId="ListDeep""#).expect("style");
    let num_pr = (styles_xml[deep..].find("<w:numPr>").expect("numPr") + deep) as u64;
    let expected = [(Defect::Style(StyleDefect::NumPrMissingNumId), num_pr)];

    let package = open(&bytes);
    let document = package.document().expect("document");
    let mut styles = document.styles().expect("styles");
    let style = styles
        .get_by_id("ListDeep")
        .expect("style table no longer fails")
        .expect("style");
    assert_eq!(style.numbering(), None);
    let resolved = styles
        .resolved_numbering("ListDeep")
        .expect("resolves")
        .expect("inherited numbering");
    assert_eq!((resolved.num_id, resolved.level), (1, 2));
    assert_eq!(
        defects(styles.tolerance_report().expect("report")),
        expected
    );

    let items = document.list_items().expect("list items resolve");
    let item = items
        .iter()
        .find(|item| item.text == "2-a")
        .expect("2-a numbered through its style");
    assert_eq!((item.numbering.num_id, item.numbering.level), (1, 2));
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
fn empty_style_num_pr_inherits_and_is_recorded() {
    let bytes = ilvl_only_style_fixture("<w:numPr/>");
    let package = open(&bytes);
    let document = package.document().expect("document");
    let items = document.list_items().expect("list items resolve");
    let item = items
        .iter()
        .find(|item| item.text == "2-a")
        .expect("2-a numbered through its style");
    assert_eq!((item.numbering.num_id, item.numbering.level), (1, 0));
    assert_eq!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .map(|tolerated| tolerated.defect)
            .collect::<Vec<_>>(),
        [Defect::Style(StyleDefect::NumPrMissingNumId)]
    );
}

#[test]
fn complete_style_num_pr_records_nothing() {
    let bytes =
        ilvl_only_style_fixture(r#"<w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr>"#);
    let package = open(&bytes);
    let document = package.document().expect("document");
    let items = document.list_items().expect("list items resolve");
    let item = items
        .iter()
        .find(|item| item.text == "2-a")
        .expect("2-a numbered through its style");
    assert_eq!(item.numbering.level, 2);
    assert!(document.tolerance_report().is_clean());
}

#[test]
fn nested_style_num_pr_stays_fatal() {
    let bytes = ilvl_only_style_fixture("<w:numPr><w:numPr/></w:numPr>");
    let package = open(&bytes);
    let mut styles = package
        .document()
        .expect("document")
        .styles()
        .expect("styles");
    assert!(styles.len().is_err());
}
