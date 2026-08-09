//! Regression tests for the mutable document owner.

use super::MutableDocument;
use crate::Builder;
use crate::Document;
use crate::core::{OwnedPackage, PackageWriter};
use crate::elements::parser::OrderElement;
use crate::elements::table::{Table, TableCell, TableRow};
use crate::elements::text::{Hyperlink, List, ListItem, Paragraph};
use crate::page_layout::PageUsage;

const MINIMAL_CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Original</text:p></office:text></office:body></office:document-content>"#;

fn source_document() -> Document {
    let mut builder = Builder::new();
    builder.add_paragraph("Before table").unwrap();

    let mut table = Table::new();
    table.set_name("Data");
    let mut row = TableRow::new();
    let mut cell = TableCell::new();
    cell.set_text("Cell content");
    row.add_cell(cell);
    table.add_row(row);
    builder.add_table(table).unwrap();

    builder.add_heading("After table", 2).unwrap();
    builder
        .add_bulleted_list(vec!["First item", "Second item"])
        .unwrap();
    builder.add_paragraph("Document end").unwrap();

    Document::from_bytes(builder.build().unwrap()).unwrap()
}

#[test]
fn mutable_metadata_is_deterministic_and_preserves_source_timestamps() {
    let mut fresh = MutableDocument::new();
    fresh.metadata_mut().created = Some("2026-08-10T01:02:03Z".parse().unwrap());
    fresh.metadata_mut().modified = Some("2026-08-10T04:05:06Z".parse().unwrap());
    let fresh_package = OwnedPackage::from_bytes(fresh.to_bytes().unwrap()).unwrap();
    let fresh_meta = String::from_utf8(fresh_package.get_file("meta.xml").unwrap()).unwrap();
    assert!(fresh_meta.contains("<meta:creation-date>2026-08-10T01:02:03Z</meta:creation-date>"));
    assert!(fresh_meta.contains("<dc:date>2026-08-10T04:05:06Z</dc:date>"));
    litchi_odf_common::compact_xml::validate(fresh_meta.as_bytes()).unwrap();

    const SOURCE_META: &str = r#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><meta:generator>Producer/1</meta:generator><dc:date>2024-01-02T03:04:05Z</dc:date></office:meta></office:document-meta>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer
        .add_file("content.xml", MINIMAL_CONTENT.as_bytes())
        .unwrap();
    writer.add_file("meta.xml", SOURCE_META.as_bytes()).unwrap();
    let source = Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    let mut edited = MutableDocument::from_document(source).unwrap();
    edited.add_paragraph("Changed").unwrap();
    let output = OwnedPackage::from_bytes(edited.to_bytes().unwrap()).unwrap();
    assert_eq!(output.get_file("meta.xml").unwrap(), SOURCE_META.as_bytes());
}

#[test]
fn mutable_default_metadata_has_no_ambient_timestamp() {
    let document = MutableDocument::new();
    let package = OwnedPackage::from_bytes(document.to_bytes().unwrap()).unwrap();
    let meta = String::from_utf8(package.get_file("meta.xml").unwrap()).unwrap();
    assert!(!meta.contains("<meta:creation-date>"));
    assert!(!meta.contains("<dc:date>"));
    litchi_odf_common::compact_xml::validate(meta.as_bytes()).unwrap();
}

fn rich_note_body() -> crate::NoteBodyContent {
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    crate::NoteBodyContent::new(vec![crate::MetaFieldNode::Element(
        crate::MetaFieldElement {
            namespace_uri: TEXT.to_string(),
            local_name: "p".to_string(),
            attributes: Vec::new(),
            children: vec![
                crate::MetaFieldNode::Text("Styled ".to_string()),
                crate::MetaFieldNode::Element(crate::MetaFieldElement {
                    namespace_uri: TEXT.to_string(),
                    local_name: "span".to_string(),
                    attributes: vec![crate::MetaFieldAttribute {
                        namespace_uri: TEXT.to_string(),
                        local_name: "style-name".to_string(),
                        value: "Emphasis".to_string(),
                    }],
                    children: vec![crate::MetaFieldNode::Text("body".to_string())],
                }),
            ],
        },
    )])
    .unwrap()
}

fn element_kinds(document: &Document) -> Vec<&'static str> {
    document
        .elements()
        .unwrap()
        .iter()
        .map(|element| match element {
            OrderElement::Paragraph(_) => "paragraph",
            OrderElement::NumberedParagraph(_) => "numbered-paragraph",
            OrderElement::Heading(_) => "heading",
            OrderElement::Table(_) => "table",
            OrderElement::List(_) => "list",
        })
        .collect()
}

#[test]
fn conversion_preserves_top_level_order_without_nested_paragraph_duplicates() {
    let mutable = MutableDocument::from_document(source_document()).unwrap();

    assert_eq!(mutable.paragraphs().len(), 2);
    assert_eq!(mutable.tables().len(), 1);
    assert_eq!(mutable.headings().len(), 1);
    assert_eq!(mutable.lists().len(), 1);
    assert_eq!(mutable.headings()[0].text().unwrap(), "After table");
    assert_eq!(mutable.headings()[0].level(), Some(2));
    assert_eq!(mutable.lists()[0].items().unwrap().len(), 2);
}

#[test]
fn mutable_hyperlink_authoring_round_trips_through_an_odt_package() {
    let mut mutable = MutableDocument::new();
    mutable
        .add_hyperlink("https://example.test/", "External")
        .unwrap();
    let mut internal = Hyperlink::with_href("#bookmark", "Internal").unwrap();
    internal.set_name("bookmark-link");
    mutable.add_hyperlink_element(internal).unwrap();

    let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.hyperlinks().unwrap(),
        vec![
            ("External".to_string(), "https://example.test/".to_string()),
            ("Internal".to_string(), "#bookmark".to_string()),
        ]
    );
    assert_eq!(document.text().unwrap(), "External\nInternal");
}

#[test]
fn mutable_ruby_annotation_and_style_crud_round_trip_through_an_odt_package() {
    let first_style = crate::ruby_family::Style::new(
        "RubyAbove",
        Some(crate::ruby_family::Properties {
            position: Some(crate::ruby_family::Position::Above),
            alignment: Some(crate::ruby_family::Alignment::Center),
        }),
    )
    .unwrap();
    let second_style = crate::ruby_family::Style::new(
        "RubyAbove",
        Some(crate::ruby_family::Properties {
            position: Some(crate::ruby_family::Position::Below),
            alignment: Some(crate::ruby_family::Alignment::DistributeLetter),
        }),
    )
    .unwrap();
    let first = crate::ruby_family::Annotation::new(
        Some(first_style.name.clone()),
        crate::ruby_family::Base::from_text("語").unwrap(),
        "ご",
        None,
    )
    .unwrap();
    let replacement = crate::ruby_family::Annotation::new(
        Some(first_style.name.clone()),
        crate::ruby_family::Base::from_text("文").unwrap(),
        "ぶん",
        None,
    )
    .unwrap();

    let mut mutable = MutableDocument::new();
    mutable.add_paragraph("Read ").unwrap();
    assert_eq!(mutable.set_ruby_style(&first_style).unwrap(), None);
    assert_eq!(
        mutable.set_ruby_style(&second_style).unwrap(),
        Some(first_style)
    );
    mutable.insert_ruby_annotation(0, &first).unwrap();
    assert_eq!(
        mutable.replace_ruby_annotation(0, &replacement).unwrap(),
        first
    );

    let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(document.ruby_styles().unwrap().styles, vec![second_style]);
    assert_eq!(
        document.ruby_annotations().unwrap().annotations,
        vec![replacement.clone()]
    );

    assert_eq!(mutable.remove_ruby_annotation(0).unwrap(), replacement);
    assert!(mutable.ruby_annotations().unwrap().annotations.is_empty());
    assert!(mutable.remove_ruby_style("RubyAbove").unwrap().is_some());
    assert!(mutable.ruby_styles().unwrap().styles.is_empty());
}

#[test]
fn mutable_ruby_range_wrapping_round_trips_through_an_odt_package() {
    let annotation = crate::ruby_family::Annotation::new(
        None,
        crate::ruby_family::Base::from_text("字").unwrap(),
        "じ",
        None,
    )
    .unwrap();
    let mut mutable = MutableDocument::new();
    mutable.add_paragraph("Read 漢字").unwrap();
    let start = "Read 漢".len();
    mutable
        .wrap_ruby_annotation(0, start..start + "字".len(), &annotation)
        .unwrap();

    let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.ruby_annotations().unwrap().annotations,
        vec![annotation]
    );
}

#[test]
fn mutable_line_numbering_configuration_round_trips_without_generation() {
    let first = crate::line_numbering::Configuration {
        number_lines: Some(true),
        number_format: Some(crate::line_numbering::Format::LowerAlpha),
        letter_sync: Some(true),
        style_name: Some("LineNumbers".to_string()),
        increment: Some(2),
        number_position: Some(crate::line_numbering::Position::Inner),
        offset: Some(crate::line_numbering::NonNegativeLength::new("0.2in").unwrap()),
        count_empty_lines: Some(false),
        count_in_text_boxes: Some(true),
        restart_on_page: Some(false),
        separator: Some(crate::line_numbering::Separator {
            increment: Some(4),
            text: " · ".to_string(),
        }),
    };
    let replacement = crate::line_numbering::Configuration {
        number_lines: Some(false),
        number_format: Some(crate::line_numbering::Format::UpperRoman),
        increment: Some(1),
        ..crate::line_numbering::Configuration::default()
    };

    let mut mutable = MutableDocument::new();
    assert_eq!(mutable.line_numbering_configuration().unwrap(), None);
    assert_eq!(
        mutable.set_line_numbering_configuration(&first).unwrap(),
        None
    );
    assert_eq!(
        mutable.line_numbering_configuration().unwrap(),
        Some(first.clone())
    );
    assert_eq!(
        mutable
            .set_line_numbering_configuration(&replacement)
            .unwrap(),
        Some(first)
    );

    let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.line_numbering_configuration().unwrap(),
        Some(replacement.clone())
    );
    assert_eq!(
        mutable.clear_line_numbering_configuration().unwrap(),
        Some(replacement)
    );
    assert_eq!(mutable.line_numbering_configuration().unwrap(), None);
    let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(document.line_numbering_configuration().unwrap(), None);
}

#[test]
fn read_modify_write_keeps_paragraph_table_heading_and_list_order() {
    let mutable = MutableDocument::from_document(source_document()).unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();

    assert_eq!(
        element_kinds(&round_trip),
        ["paragraph", "table", "heading", "list", "paragraph"]
    );

    let elements = round_trip.elements().unwrap();
    let OrderElement::Table(table) = &elements[1] else {
        panic!("second element should remain a table");
    };
    assert_eq!(table.name(), Some("Data"));
    assert_eq!(
        table
            .row(0)
            .unwrap()
            .unwrap()
            .cell(0)
            .unwrap()
            .unwrap()
            .text()
            .unwrap(),
        "Cell content"
    );

    let OrderElement::List(list) = &elements[3] else {
        panic!("fourth element should remain a list");
    };
    let items = list.items().unwrap();
    assert_eq!(items[0].text().unwrap(), "First item");
    assert_eq!(items[1].text().unwrap(), "Second item");
}

#[test]
fn mutable_document_can_add_headings_and_lists() {
    let mut document = MutableDocument::new();
    assert!(document.add_heading("Invalid", 0).is_err());
    document.add_heading("Title", 1).unwrap();

    let mut list = List::new();
    let mut item = ListItem::new();
    let mut paragraph = Paragraph::new();
    paragraph.set_text("Item");
    item.add_paragraph(paragraph);
    list.add_item(item);
    document.add_list(list).unwrap();

    let round_trip = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert_eq!(element_kinds(&round_trip), ["heading", "list"]);
}

#[test]
fn read_modify_write_preserves_auxiliary_package_parts_and_media_types() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer
        .add_file("content.xml", MINIMAL_CONTENT.as_bytes())
        .unwrap();
    writer
        .add_file("settings.xml", b"document settings")
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/photo.bin", b"image", "image/x-test")
        .unwrap();
    writer
        .add_manifest_entry("Object 1/", "application/vnd.oasis.opendocument.text")
        .unwrap();
    writer
        .add_file("Object 1/content.xml", b"embedded object")
        .unwrap();
    writer
        .add_file_with_media_type(
            "custom/data.bin",
            b"custom payload",
            "application/x-litchi-test",
        )
        .unwrap();
    writer
        .add_file("META-INF/documentsignatures.xml", b"stale signature")
        .unwrap();

    let source_bytes = writer.finish_to_bytes().unwrap();
    let source = Document::from_bytes(source_bytes.clone()).unwrap();
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    let mut mutable = MutableDocument::from_document(source).unwrap();
    mutable.add_paragraph("Modified").unwrap();
    let output = OwnedPackage::from_bytes(mutable.to_bytes().unwrap()).unwrap();

    assert_eq!(
        output.get_file("settings.xml").unwrap(),
        b"document settings"
    );
    assert_eq!(output.get_file("Pictures/photo.bin").unwrap(), b"image");
    assert_eq!(
        output.get_file("Object 1/content.xml").unwrap(),
        b"embedded object"
    );
    assert_eq!(
        output.get_file("custom/data.bin").unwrap(),
        b"custom payload"
    );
    assert!(!output.has_file("META-INF/documentsignatures.xml").unwrap());

    let package = output.package().unwrap();
    assert_eq!(
        package.manifest().get_media_type("Pictures/photo.bin"),
        Some("image/x-test")
    );
    assert_eq!(
        package.manifest().get_media_type("Object 1/"),
        Some("application/vnd.oasis.opendocument.text")
    );
    assert_eq!(
        package.manifest().get_media_type("custom/data.bin"),
        Some("application/x-litchi-test")
    );
}

#[test]
fn mutable_note_crud_round_trips_through_an_odt_package() {
    let mut mutable = MutableDocument::new();
    mutable.add_paragraph("Before").unwrap();
    let first = crate::Note::new(crate::NoteClass::Footnote, "1", "Initial").unwrap();
    mutable.insert_note(0, &first).unwrap();
    assert_eq!(mutable.footnotes().unwrap(), vec![first.clone()]);
    assert!(mutable.endnotes().unwrap().is_empty());

    let replacement = crate::Note::new(crate::NoteClass::Endnote, "i", "Replacement").unwrap();
    assert_eq!(mutable.replace_note(0, &replacement).unwrap(), first);
    assert_eq!(mutable.endnotes().unwrap(), vec![replacement.clone()]);
    assert_eq!(mutable.remove_note(0).unwrap(), replacement);
    assert!(mutable.notes().unwrap().is_empty());

    let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert!(document.notes().unwrap().is_empty());
}

#[test]
fn mutable_document_round_trips_structured_note_authoring() {
    let mut mutable = MutableDocument::new();
    mutable.add_paragraph("Before").unwrap();
    let note =
        crate::Note::with_rich_body(crate::NoteClass::Footnote, "1", rich_note_body()).unwrap();
    mutable.insert_note(0, &note).unwrap();

    let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let notes = document.notes().unwrap();
    assert_eq!(notes, vec![note]);
}

#[test]
fn edits_master_page_regions_through_the_public_mutable_document_api() {
    const STYLES: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:styles><style:style style:name="preserved"/></office:styles><office:automatic-styles><style:page-layout style:name="pm1" style:page-usage="left"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Old header</text:p></style:header><style:footer><text:p>Old footer</text:p></style:footer></style:master-page></office:master-styles></office:document-styles>"#;

    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer
        .add_file("content.xml", MINIMAL_CONTENT.as_bytes())
        .unwrap();
    writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
    let source = Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
    let mut mutable = MutableDocument::from_document(source).unwrap();
    let layouts = mutable.page_layouts().unwrap();
    assert_eq!(layouts[0].page_usage, PageUsage::Left);
    assert_eq!(
        layouts[0].properties.as_ref().unwrap().attribute(
            Some("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),
            "page-width",
        ),
        Some("21cm")
    );

    mutable
        .set_header_footer_text(
            "Standard",
            crate::header_footer::Kind::Header,
            "New & <header>",
        )
        .unwrap();
    mutable
        .clear_header_footer("Standard", crate::header_footer::Kind::Footer)
        .unwrap();
    let pages = mutable.master_pages().unwrap();
    assert_eq!(
        pages[0]
            .region(crate::header_footer::Kind::Header)
            .unwrap()
            .text,
        "New & <header>"
    );
    assert!(
        pages[0]
            .region(crate::header_footer::Kind::Footer)
            .is_none()
    );

    let output = OwnedPackage::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let styles = String::from_utf8(output.get_file("styles.xml").unwrap()).unwrap();
    assert!(styles.contains("<style:style style:name=\"preserved\"/>"));
    assert!(!styles.contains("Old footer"));
    let round_trip = Document::from_bytes(output.as_bytes().to_vec()).unwrap();
    assert_eq!(round_trip.page_layouts().unwrap(), layouts);
    assert_eq!(
        round_trip.master_pages().unwrap()[0]
            .region(crate::header_footer::Kind::Header)
            .unwrap()
            .text,
        "New & <header>"
    );
}

#[test]
fn creates_a_master_page_and_header_in_a_new_document() {
    let mut mutable = MutableDocument::new();
    mutable.add_master_page("Standard", "pm1").unwrap();
    let layout = r#"<s:page-layout xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" s:name="pm1" s:page-usage="mirrored"><s:page-layout-properties f:page-width="21cm" f:page-height="29.7cm"/></s:page-layout>"#;
    mutable.set_page_layout_xml("pm1", layout).unwrap();
    mutable
        .set_header_footer_text(
            "Standard",
            crate::header_footer::Kind::Header,
            "Created header",
        )
        .unwrap();
    let rich = r#"<s:header xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p>Page <t:page-number/></t:p></s:header>"#;
    mutable
        .set_header_footer_xml("Standard", crate::header_footer::Kind::Header, rich)
        .unwrap();

    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let pages = round_trip.master_pages().unwrap();
    let layouts = round_trip.page_layouts().unwrap();
    assert_eq!(layouts.len(), 1);
    assert_eq!(layouts[0].xml, layout);
    assert_eq!(layouts[0].page_usage, PageUsage::Mirrored);
    assert_eq!(pages.len(), 1);
    assert_eq!(pages[0].name, "Standard");
    assert_eq!(pages[0].page_layout_name.as_deref(), Some("pm1"));
    assert_eq!(
        pages[0]
            .region(crate::header_footer::Kind::Header)
            .unwrap()
            .text,
        "Page "
    );
    assert_eq!(
        pages[0]
            .region(crate::header_footer::Kind::Header)
            .unwrap()
            .xml,
        rich
    );
    let styles = String::from_utf8(round_trip.get_file("styles.xml").unwrap()).unwrap();
    assert!(styles.contains(layout));
}

#[test]
fn insert_image_round_trips_through_package_and_read_api() {
    use crate::frame::{Anchor, Length};

    let mut doc = MutableDocument::new();
    doc.add_paragraph("Before image").unwrap();
    let png = minimal_png();
    let path = doc
        .insert_image(
            1,
            &png,
            &Length::centimeters(10.0),
            &Length::centimeters(4.0),
            Anchor::AsChar,
        )
        .unwrap();
    assert_eq!(path, "Pictures/image1.png");
    doc.add_paragraph("After image").unwrap();
    assert!(
        doc.insert_image(
            99,
            &png,
            &Length::points(1.0),
            &Length::points(1.0),
            Anchor::Page
        )
        .is_err()
    );
    assert!(
        doc.insert_image(
            0,
            b"not-an-image",
            &Length::points(1.0),
            &Length::points(1.0),
            Anchor::Page
        )
        .is_err()
    );

    let round_trip = Document::from_bytes(doc.to_bytes().unwrap()).unwrap();
    // Text content survives around the frame.
    let text = round_trip.text().unwrap();
    assert!(text.contains("Before image"));
    assert!(text.contains("After image"));
    // The frame is discoverable with identity, geometry, and anchor.
    let images = round_trip.images().unwrap();
    assert_eq!(images.len(), 1);
    let frame = images[0].frame.as_ref().unwrap();
    assert_eq!(frame.width.as_deref(), Some("10cm"));
    assert_eq!(frame.height.as_deref(), Some("4cm"));
    assert_eq!(frame.anchor_type.as_deref(), Some("as-char"));
    assert_eq!(images[0].package_path(), Some("Pictures/image1.png"));
    // Payload is stored verbatim.
    assert_eq!(round_trip.image_bytes(&images[0]).unwrap(), Some(png));
}

#[test]
fn insert_image_coexists_with_existing_media() {
    use crate::frame::{Anchor, Length};

    let mut doc = MutableDocument::new();
    let first = doc
        .insert_image(
            0,
            &minimal_png(),
            &Length::points(8.0),
            &Length::points(8.0),
            Anchor::Page,
        )
        .unwrap();
    let second = doc
        .insert_image(
            1,
            &minimal_jpeg(),
            &Length::points(8.0),
            &Length::points(8.0),
            Anchor::Page,
        )
        .unwrap();
    assert_eq!(first, "Pictures/image1.png");
    assert_eq!(second, "Pictures/image2.jpg");

    let round_trip = Document::from_bytes(doc.to_bytes().unwrap()).unwrap();
    let images = round_trip.images().unwrap();
    assert_eq!(images.len(), 2);
    let mut paths: Vec<_> = images
        .iter()
        .filter_map(|image| image.package_path())
        .collect();
    paths.sort_unstable();
    assert_eq!(paths, ["Pictures/image1.png", "Pictures/image2.jpg"]);
}

#[test]
fn insert_text_box_round_trips_story_text() {
    use crate::frame::{Anchor, Length};

    let mut doc = MutableDocument::new();
    doc.add_paragraph("Intro").unwrap();
    let name = doc
        .insert_text_box(
            1,
            "boxed <text> & more\nsecond line",
            &Length::inches(2.0),
            &Length::inches(1.0),
            Anchor::Paragraph,
        )
        .unwrap();
    assert_eq!(name, "Text Box 1");

    let round_trip = Document::from_bytes(doc.to_bytes().unwrap()).unwrap();
    let content = String::from_utf8(round_trip.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("draw:text-box"));
    assert!(content.contains("boxed &lt;text&gt; &amp; more"));
    assert!(content.contains("second line"));
    assert!(content.contains("text:anchor-type=\"paragraph\""));
    assert!(round_trip.text().unwrap().contains("Intro"));
}

fn minimal_png() -> Vec<u8> {
    // 1x1 transparent PNG.
    let mut bytes = b"\x89PNG\r\n\x1a\n".to_vec();
    bytes.extend_from_slice(&[0, 0, 0, 13]);
    bytes.extend_from_slice(b"IHDR");
    bytes.extend_from_slice(&[0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0]);
    bytes.extend_from_slice(&[0x1f, 0x15, 0xc4, 0x89]);
    bytes.extend_from_slice(&[0, 0, 0, 11]);
    bytes.extend_from_slice(b"IDAT");
    bytes.extend_from_slice(&[
        0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d,
    ]);
    bytes.extend_from_slice(&[0x0a, 0x2d, 0xb4]);
    bytes.extend_from_slice(&[0, 0, 0, 0]);
    bytes.extend_from_slice(b"IEND");
    bytes.extend_from_slice(&[0xae, 0x42, 0x60, 0x82]);
    bytes
}

fn minimal_jpeg() -> Vec<u8> {
    let mut bytes = b"\xff\xd8\xff\xe0".to_vec();
    bytes.extend_from_slice(&[0, 16]);
    bytes.extend_from_slice(b"JFIF\0");
    bytes.extend_from_slice(&[1, 1, 0, 0, 1, 0, 1, 0, 0, 0xff, 0xd9]);
    bytes
}
