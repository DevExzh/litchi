//! Focused semantic and codec tests for the document facade.

use super::codec::{parse_hyperlinks, parse_image_references};
use super::model::Document;
use crate::constants;
use crate::core::PackageWriter;
use crate::elements::parser::OrderElement;

fn document(content: &str) -> Document {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_TEXT).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, content.as_bytes())
        .unwrap();
    Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
}

fn raw_package(content: &[u8]) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("mimetype", constants::ODF_TEXT.as_bytes())
        .unwrap();
    writer
        .write_stored(constants::ODF_CONTENT, content)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn prepared_detection_transfers_the_same_archive_index_into_semantic_open() {
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>prepared</t:p></o:text></o:body></o:document-content>"#;
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_TEXT).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, content.as_bytes())
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let prepared = litchi_odf_common::detect::prepared(bytes.clone()).unwrap();
    let index_identity = prepared.prepared_index_identity();
    let document = Document::from_prepared_package(prepared).unwrap();

    assert_eq!(document.prepared_index_identity(), index_identity);
    assert_eq!(document.text().unwrap(), "prepared");
    assert_eq!(
        Document::from_bytes(bytes).unwrap().text().unwrap(),
        "prepared"
    );
}

#[test]
fn prepared_detection_rejects_a_wrong_family_before_semantic_open() {
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:spreadsheet/></o:body></o:document-content>"#;
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, content.as_bytes())
        .unwrap();
    let prepared = litchi_odf_common::detect::prepared(writer.finish_to_bytes().unwrap())
        .expect("valid ODS package should be detected");

    assert!(Document::from_prepared_package(prepared).is_err());
}

#[test]
fn prepared_detection_rejects_junk_and_malformed_odt_content_xml() {
    for content in [
        br#"<junk/>"#.as_slice(),
        br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:spreadsheet/></o:body></o:document-content>"#.as_slice(),
        &[0xff, 0xfe][..],
    ] {
        let prepared = litchi_odf_common::detect::prepared(raw_package(content))
            .expect("archive structure should prepare before XML validation");
        assert!(Document::from_prepared_package(prepared).is_err());
    }
}

#[test]
fn prepared_semantic_open_rejects_invalid_xml_references_and_late_declarations() {
    for content in [
        br#"<!--bad--comment--><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>bad comment</o:text></o:body></o:document-content>"#.as_slice(),
        br#"<?xml version="1.0"?><?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>duplicate declaration</o:text></o:body></o:document-content>"#.as_slice(),
        br#"
<?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>whitespace-delayed declaration</o:text></o:body></o:document-content>"#.as_slice(),
        br#"<!--comment--><?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>late declaration</o:text></o:body></o:document-content>"#.as_slice(),
        br#"<?odf-prologue?><?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>late declaration</o:text></o:body></o:document-content>"#.as_slice(),
        br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>&#0;</o:text></o:body></o:document-content>"#.as_slice(),
        br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>&#xD800;</o:text></o:body></o:document-content>"#.as_slice(),
        br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>&#x110000;</o:text></o:body></o:document-content>"#.as_slice(),
        br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text>ok</o:text></o:body></o:document-content><?xml version="1.0"?>"#.as_slice(),
    ] {
        let prepared = litchi_odf_common::detect::prepared(raw_package(content))
            .expect("archive structure should prepare before XML validation");
        assert!(Document::from_prepared_package(prepared).is_err());
    }
}

#[test]
fn text_model_accepts_arbitrary_prefixes_and_decodes_mixed_text() {
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:h t:outline-level="2">Title &amp; More</t:h><t:p t:style-name="Body">A<t:span>B</t:span>C<t:s t:c="2"/>D<![CDATA[!]]></t:p></o:text></o:body></o:document-content>"#;
    let document = document(content);

    assert_eq!(document.text().unwrap(), "Title & More\nABC  D!");
    assert_eq!(document.paragraph_count().unwrap(), 1);
    let paragraph = document.paragraphs().unwrap().remove(0);
    assert_eq!(paragraph.style_name(), Some("Body"));
    assert_eq!(paragraph.text().unwrap(), "ABC  D!");
    let paragraph = document.paragraph(0).unwrap().unwrap();
    assert_eq!(paragraph.style_name(), Some("Body"));
    assert_eq!(paragraph.text().unwrap(), "ABC  D!");
    assert!(document.paragraph(1).unwrap().is_none());

    let elements = document.elements().unwrap();
    assert_eq!(elements.len(), 2);
    let OrderElement::Heading(heading) = &elements[0] else {
        panic!("first document element is not a heading");
    };
    assert_eq!(heading.level(), Some(2));
    assert_eq!(heading.text().unwrap(), "Title & More");
    let OrderElement::Paragraph(paragraph) = &elements[1] else {
        panic!("second document element is not a paragraph");
    };
    assert_eq!(paragraph.text().unwrap(), "ABC  D!");
}

#[test]
fn indexed_paragraph_lookup_validates_content_after_the_match() {
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>Selected</t:p><t:p><t:s t:c="1000001"/></t:p></o:text></o:body></o:document-content>"#;
    let document = document(content);

    assert!(document.paragraph(0).is_err());
}

#[test]
fn references_fields_and_images_are_namespace_aware_and_decoded() {
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:body><o:text><t:p><t:bookmark t:name="point &amp; one"/><t:bookmark-start t:name="range"/>ab<t:s t:c="2"/>c<t:bookmark-end t:name="range"/></t:p><t:p><t:a x:type="simple" x:href="https://example.invalid/?a=1&amp;b=2">A<t:span>B &amp; C</t:span><t:s t:c="2"/>D<![CDATA[!]]></t:a><t:date s:data-style-name="N1" t:fixed="true" t:date-value="2026-07-16">July &amp; 16</t:date><t:word-count>42</t:word-count><d:frame><d:image x:href="Pictures/a&amp;b.png"/><d:image x:href="https://example.invalid/image.png"/><d:image><o:binary-data>AA==</o:binary-data></d:image></d:frame></t:p></o:text></o:body></o:document-content>"#;
    let document = document(content);

    assert_eq!(
        document.hyperlinks().unwrap(),
        vec![(
            "AB & C  D!".to_string(),
            "https://example.invalid/?a=1&b=2".to_string()
        )]
    );
    assert_eq!(
        document.bookmark_names().unwrap(),
        vec!["point & one".to_string(), "range".to_string()]
    );
    let bookmarks = document.bookmarks().unwrap();
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(bookmarks[0].name(), Some("point & one"));
    let ranges = document.bookmark_ranges().unwrap();
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].name, "range");
    assert_eq!(ranges[0].start, Some((0, 0)));
    assert_eq!(ranges[0].end, Some((0, 5)));
    assert!(ranges[0].is_complete());

    let fields = document.fields().unwrap();
    assert_eq!(fields.len(), 2);
    assert_eq!(fields[0].field_type(), "text:date");
    assert_eq!(fields[0].value(), "July & 16");
    assert_eq!(fields[0].format(), Some("N1"));
    assert_eq!(fields[1].field_type(), "text:word-count");
    assert_eq!(fields[1].value(), "42");
    assert_eq!(
        document.image_paths().unwrap(),
        vec![
            "Pictures/a&b.png".to_string(),
            "https://example.invalid/image.png".to_string()
        ]
    );
}

#[test]
fn reference_readers_reject_malformed_xml_and_duplicate_expanded_attributes() {
    let duplicate = r#"<t:p xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:y="http://www.w3.org/1999/xlink"><t:a x:href="a" y:href="b">bad</t:a></t:p>"#;
    assert!(parse_hyperlinks(duplicate).is_err());
    let missing_name =
        r#"<t:p xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:bookmark/></t:p>"#;
    assert!(crate::elements::bookmark::BookmarkParser::parse_bookmarks(missing_name).is_err());
    let nonempty = r#"<t:p xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:bookmark t:name="bad">content</t:bookmark></t:p>"#;
    assert!(crate::elements::bookmark::BookmarkParser::parse_bookmarks(nonempty).is_err());
    assert!(parse_image_references("<d:image").is_err());
    assert!(crate::elements::field::FieldParser::parse_fields("<t:date>").is_err());
}
