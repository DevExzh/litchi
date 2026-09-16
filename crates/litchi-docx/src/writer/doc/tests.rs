use super::super::toc::TableOfContents;
use crate::alt::{Chunk, Conformance, Rel};
use crate::settings::ProtectionType;

use super::*;

/// A pretty-printed main document whose body children are indented, so a
/// three-byte offset shift truncates every preserved range it produces.
const INDENTED_BODY: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
    "  <w:body>\n",
    "    <w:p w:rsidR=\"00A1\"><w:r><w:t>start</w:t></w:r></w:p>\n",
    "    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>\n",
    "    <w:altChunk r:id=\"rId1\" />\n",
    "    <w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>\n",
    "  </w:body>\n",
    "</w:document>\n",
);

/// The same body with the final section properties before, not after, the
/// alternative-format anchor, as `alt-chunk-header.docx` writes them.
const SECTION_BEFORE_ALT: &str = concat!(
    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
    "  <w:body>\n",
    "    <w:p><w:r><w:t>start</w:t></w:r></w:p>\n",
    "    <w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>\n",
    "    <w:altChunk r:id=\"rId1\" />\n",
    "  </w:body>\n",
    "</w:document>\n",
);

/// The UTF-8 byte order mark, spelled here so the tests read the same way the
/// parser does.
const MARK: &str = "\u{feff}";

#[test]
fn test_create_empty_document() {
    let doc = MutableDocument::new();
    assert_eq!(doc.paragraph_count(), 0);
    assert_eq!(doc.table_count(), 0);
}

#[test]
fn test_add_paragraph() {
    let mut doc = MutableDocument::new();
    doc.add_paragraph_with_text("Hello, World!");
    assert_eq!(doc.paragraph_count(), 1);
}

#[test]
fn removing_alt_chunk_keeps_pending_toc_insertion_in_bounds() {
    let mut doc = MutableDocument::new();
    let chunk = Chunk::new(Rel::new("rIdAlt1").unwrap(), None);
    doc.insert_alt(0, chunk, Conformance::Transitional).unwrap();
    doc.add_toc(TableOfContents::new()).unwrap();

    doc.remove_alt(0).unwrap();
    doc.generate_toc_if_needed().unwrap();

    assert!(doc.alts().is_empty());
    assert_eq!(doc.paragraph_count(), 1);
    assert!(doc.to_xml().unwrap().contains("TOC"));
}

#[test]
fn headings_use_style_ids_instead_of_display_names() {
    let mut doc = MutableDocument::new();
    for level in 0..=9 {
        doc.add_heading(&format!("Level {level}"), level).unwrap();
    }

    let xml = doc.to_xml().unwrap();
    assert!(xml.contains(r#"<w:pStyle w:val="Title"/>"#));
    for level in 1..=9 {
        assert!(xml.contains(&format!(r#"<w:pStyle w:val="Heading{level}"/>"#)));
        assert!(!xml.contains(&format!(r#"<w:pStyle w:val="Heading {level}"/>"#)));
    }
    assert!(doc.add_heading("invalid", 10).is_err());
}

#[test]
fn test_add_table() {
    let mut doc = MutableDocument::new();
    let table = doc.add_table(2, 3);
    assert_eq!(table.row_count(), 2);
    table.cell(0, 0).unwrap().set_text("Cell 1");
    assert_eq!(doc.table_count(), 1);
}

#[test]
fn test_xml_generation() {
    let mut doc = MutableDocument::new();
    doc.add_paragraph_with_text("Test paragraph");

    let xml = doc.to_xml().unwrap();
    assert!(xml.contains("<w:document"));
    assert!(xml.contains("<w:body>"));
    assert!(xml.contains("<w:p>"));
    assert!(xml.contains("Test paragraph"));
}

#[test]
fn test_run_formatting() {
    let mut doc = MutableDocument::new();
    let para = doc.add_paragraph();
    para.add_run_with_text("Bold text").bold(true);
    para.add_run_with_text("Italic text").italic(true);

    let xml = doc.to_xml().unwrap();
    assert!(xml.contains("<w:b/>"));
    assert!(xml.contains("<w:i/>"));
}

#[test]
fn appending_preserves_existing_body_xml_exactly() {
    let input = r#"<?xml version="1.0"?><q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:extension"><q:body>
  <!--keep--><q:p q:rsidR="00AB"><q:r><q:t><![CDATA[A < B]]></q:t></q:r><x:payload data="1 &amp; 2"/></q:p>
  <q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl>
  <x:custom><![CDATA[opaque <xml>]]></x:custom>
  <q:sectPr><q:pgSz q:w="20000" q:h="10000"/></q:sectPr>
</q:body></q:document>"#;
    let mut document = MutableDocument::from_xml(input).unwrap();
    assert_eq!(document.paragraph_count(), 1);
    assert_eq!(document.table_count(), 1);

    document.add_paragraph_with_text("appended");
    let output = document.to_xml().unwrap();
    assert!(output.starts_with(
            r#"<?xml version="1.0"?><q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:extension"><q:body"#
        ));
    assert!(
        output
            .contains(r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#)
    );
    assert!(output.contains(
            r#"<q:p q:rsidR="00AB"><q:r><q:t><![CDATA[A < B]]></q:t></q:r><x:payload data="1 &amp; 2"/></q:p>"#
        ));
    assert!(output.contains(
        r#"<q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl>"#
    ));
    assert!(output.contains(r#"<x:custom><![CDATA[opaque <xml>]]></x:custom>"#));
    assert!(output.contains(r#"<q:sectPr><q:pgSz q:w="20000" q:h="10000"/></q:sectPr>"#));
    assert!(output.contains("appended"));
    assert_eq!(output.matches("sectPr").count(), 2);
    assert!(output.ends_with("</q:body></q:document>"));
}

#[test]
fn existing_document_parser_rejects_missing_or_truncated_body() {
    assert!(MutableDocument::from_xml("<w:document/>").is_err());
    assert!(MutableDocument::from_xml(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>"#
        )
        .is_err());
}

#[test]
fn protection_patching_preserves_unrelated_settings_exactly() {
    let input = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:extension"><!--before--><q:smartTagType q:namespaceuri="urn:test" q:name="person" q:url="https://example.test"/><q:documentProtection q:edit="readOnly" q:enforcement="1" x:keep="yes"/><x:opaque><![CDATA[a < b]]></x:opaque><q:doNotEmbedSmartTags/></q:settings>"#;
    let mut document = MutableDocument::new();
    document.set_protection_with_password(
        ProtectionType::Comments,
        "hash&\"value".into(),
        "salt<value".into(),
    );

    let output = document.generate_settings_xml(Some(input)).unwrap();
    let output = String::from_utf8(output).unwrap();
    assert!(output.contains(r#"<q:smartTagType q:namespaceuri="urn:test" q:name="person" q:url="https://example.test"/>"#));
    assert!(output.contains(r#"<x:opaque><![CDATA[a < b]]></x:opaque>"#));
    assert!(output.contains("<q:doNotEmbedSmartTags/>"));
    assert!(output.contains(r#"<q:documentProtection q:edit="comments" q:enforcement="1" q:hash="hash&amp;&quot;value" q:salt="salt&lt;value"/>"#));
    assert_eq!(output.matches("documentProtection").count(), 1);
}

#[test]
fn protection_patching_removes_only_protection_and_handles_empty_roots() {
    let input = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="125"/><w:documentProtection w:edit="forms"/><w:savePreviewPicture/></w:settings>"#;
    let mut document = MutableDocument::new();
    document.remove_protection();
    let output = document.generate_settings_xml(Some(input)).unwrap();
    assert_eq!(
        String::from_utf8(output).unwrap(),
        r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="125"/><w:savePreviewPicture/></w:settings>"#
    );

    document.set_protection(ProtectionType::ReadOnly);
    let empty = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
    let output = String::from_utf8(document.generate_settings_xml(Some(empty)).unwrap()).unwrap();
    assert_eq!(
        output,
        r#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:documentProtection s:edit="readOnly" s:enforcement="1"/></s:settings>"#
    );

    let default_namespace =
        br#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    let output = String::from_utf8(
        document
            .generate_settings_xml(Some(default_namespace))
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        output,
        r#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:edit="readOnly" w:enforcement="1"/></settings>"#
    );
}

#[test]
fn a_byte_order_mark_shifts_no_preserved_body_range() {
    let marked = format!("{MARK}{INDENTED_BODY}");
    let plain = MutableDocument::from_xml(INDENTED_BODY).unwrap();
    let with_mark = MutableDocument::from_xml(&marked).unwrap();

    assert_eq!(with_mark.paragraph_count(), plain.paragraph_count());
    assert_eq!(with_mark.table_count(), plain.table_count());
    assert_eq!(with_mark.alts().len(), plain.alts().len());

    let plain_xml = plain.to_xml().unwrap();
    assert!(plain_xml.contains(r#"<w:p w:rsidR="00A1"><w:r><w:t>start</w:t></w:r></w:p>"#));
    assert!(plain_xml.contains(r#"<w:altChunk r:id="rId1"/>"#));
    assert_eq!(with_mark.to_xml().unwrap(), format!("{MARK}{plain_xml}"));
}

#[test]
fn a_byte_order_mark_survives_one_appended_paragraph() {
    let mut document = MutableDocument::from_xml(&format!("{MARK}{INDENTED_BODY}")).unwrap();
    document.add_paragraph_with_text("appended");

    let output = document.to_xml().unwrap();
    assert!(output.starts_with(MARK));
    assert_eq!(output.matches(MARK).count(), 1);
    assert!(output.contains(r#"<w:p w:rsidR="00A1"><w:r><w:t>start</w:t></w:r></w:p>"#));
    assert!(output.contains("appended"));
    assert!(output.ends_with("</w:body></w:document>"), "{output}");
}

#[test]
fn body_final_section_properties_must_remain_the_last_body_child() {
    let refusal = match MutableDocument::from_xml(SECTION_BEFORE_ALT) {
        Ok(_) => panic!("out-of-order section properties were admitted"),
        Err(error) => error.to_string(),
    };
    assert!(
        refusal.contains("body-final section properties are not the final body child"),
        "{refusal}"
    );
    let marked = format!("{MARK}{SECTION_BEFORE_ALT}");
    let marked_refusal = match MutableDocument::from_xml(&marked) {
        Ok(_) => panic!("a byte order mark admitted out-of-order section properties"),
        Err(error) => error.to_string(),
    };
    assert_eq!(marked_refusal, refusal);
}
