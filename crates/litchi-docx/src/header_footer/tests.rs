//! Regression tests for the layered story model and codec.

use super::{Kind, MAX_XML_BYTES, MAX_XML_DEPTH, Role, Story};

fn story(xml: &[u8], kind: Kind) -> Story {
    Story::from_xml_bytes(xml.to_vec(), kind).unwrap()
}

#[test]
fn extracts_header_and_footer_text() {
    let header = story(
        b"<w:hdr><w:p><w:r><w:t>Header Text</w:t></w:r></w:p></w:hdr>",
        Kind::Primary,
    );
    let footer = story(
        b"<w:ftr><w:p><w:r><w:t>Footer Text</w:t></w:r></w:p></w:ftr>",
        Kind::Primary,
    );
    assert_eq!(header.role(), Role::Header);
    assert_eq!(footer.role(), Role::Footer);
    assert_eq!(header.text().unwrap(), "Header Text");
    assert_eq!(footer.text().unwrap(), "Footer Text");
}

#[test]
fn kind_xml_and_display_round_trip() {
    let values = [
        (Kind::Primary, "default", "Primary", 1),
        (Kind::FirstPage, "first", "First Page", 2),
        (Kind::EvenPage, "even", "Even Page", 3),
    ];
    for (value, xml, display, repr) in values {
        assert_eq!(value.to_xml(), xml);
        assert_eq!(Kind::from_xml(xml), Some(value));
        assert_eq!(value.to_string(), display);
        assert_eq!(value as u8, repr);
    }
    assert_eq!(Kind::from_xml("invalid"), None);
    assert_eq!(Kind::default(), Kind::Primary);
}

#[test]
fn kind_is_carried_by_the_story_without_a_tuple_wrapper() {
    let xml = b"<w:hdr><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:hdr>";
    assert_eq!(story(xml, Kind::Primary).kind(), Kind::Primary);
    assert_eq!(story(xml, Kind::FirstPage).kind(), Kind::FirstPage);
    assert_eq!(story(xml, Kind::EvenPage).kind(), Kind::EvenPage);
}

#[test]
fn raw_xml_is_lossless_and_ordered() {
    let xml = br#"<h:hdr xmlns:h="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:example"><x:before/><h:p><h:r><h:t>one</h:t></h:r></h:p><x:after/></h:hdr>"#;
    let header = story(xml, Kind::Primary);
    assert_eq!(header.xml_bytes(), xml);
    let paragraphs = header.paragraphs().unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].text().unwrap(), "one");
}

#[test]
fn empty_story_has_no_blocks() {
    let header = story(b"<w:hdr></w:hdr>", Kind::Primary);
    assert_eq!(header.text().unwrap(), "");
    assert_eq!(header.paragraph_count().unwrap(), 0);
    assert_eq!(header.table_count().unwrap(), 0);
}

#[test]
fn extracts_multiple_paragraphs_and_runs() {
    let header = story(
        br#"<w:hdr><w:p><w:r><w:t>First Paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Second Paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Run One</w:t></w:r><w:r><w:t> Run Two</w:t></w:r></w:p></w:hdr>"#,
        Kind::Primary,
    );
    let text = header.text().unwrap();
    assert!(text.contains("First Paragraph"));
    assert!(text.contains("Second Paragraph"));
    assert!(text.contains("Run One"));
    assert!(text.contains("Run Two"));
    assert_eq!(header.paragraph_count().unwrap(), 3);
    assert_eq!(header.paragraphs().unwrap().len(), 3);
}

#[test]
fn extracts_tables_in_source_order() {
    let header = story(
        br#"<w:hdr><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell 2</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>"#,
        Kind::Primary,
    );
    assert_eq!(header.table_count().unwrap(), 2);
    assert_eq!(header.tables().unwrap().len(), 2);
    assert!(header.text().unwrap().contains("Cell 1"));
}

#[test]
fn namespace_aliases_cdata_and_foreign_elements_are_handled() {
    let xml = br#"<h:hdr xmlns:h="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml"><false:p><false:r><false:t>ignored</false:t></false:r></false:p><h:p><h:r><h:t><![CDATA[A < B]]></h:t></h:r></h:p><h:tbl><h:tr><h:tc><h:p><h:r><h:t>cell</h:t></h:r></h:p></h:tc></h:tr></h:tbl><h:p/><false:tbl/></h:hdr>"#;
    let header = story(xml, Kind::Primary);
    assert_eq!(header.text().unwrap(), "A < Bcell");
    assert_eq!(header.paragraph_count().unwrap(), 3);
    assert_eq!(header.table_count().unwrap(), 1);
    let paragraphs = header.paragraphs().unwrap();
    assert_eq!(paragraphs[0].text().unwrap(), "A < B");
    assert_eq!(paragraphs[0].runs().unwrap()[0].text().unwrap(), "A < B");
    assert_eq!(header.tables().unwrap()[0].row_count().unwrap(), 1);
}

#[test]
fn malformed_story_is_rejected_at_the_codec_boundary() {
    assert!(Story::from_xml_bytes(
        br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r/>"#.to_vec(),
        Kind::Primary,
    )
    .is_err());
    assert!(Story::from_xml_bytes(b"<w:body/>".to_vec(), Kind::Primary).is_err());
}

#[test]
fn story_limits_are_bounded() {
    assert!(Story::from_xml_bytes(vec![b' '; MAX_XML_BYTES + 1], Kind::Primary).is_err());

    let mut deep = String::from("<w:hdr>");
    for _ in 0..MAX_XML_DEPTH {
        deep.push_str("<w:p>");
    }
    for _ in 0..MAX_XML_DEPTH {
        deep.push_str("</w:p>");
    }
    deep.push_str("</w:hdr>");
    assert!(Story::from_xml_bytes(deep.into_bytes(), Kind::Primary).is_err());
}

#[test]
fn mixed_content_and_unicode_are_preserved() {
    let header = story(
        "<w:hdr><w:p><w:r><w:t>Unicode: 你好世界 🎉</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table content</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p></w:hdr>".as_bytes(),
        Kind::Primary,
    );
    assert_eq!(header.paragraph_count().unwrap(), 3);
    assert_eq!(header.table_count().unwrap(), 1);
    let text = header.text().unwrap();
    assert!(text.contains("你好世界"));
    assert!(text.contains("🎉"));
    assert!(text.contains("Table content"));
    assert!(text.contains("After"));
}

#[test]
fn clone_shares_the_same_semantic_story() {
    let header = story(
        b"<w:hdr><w:p><w:r><w:t>Clonable Content</w:t></w:r></w:p></w:hdr>",
        Kind::Primary,
    );
    let cloned = header.clone();
    assert_eq!(header.text().unwrap(), cloned.text().unwrap());
    assert_eq!(header.kind(), cloned.kind());
    assert_eq!(header.role(), cloned.role());
}

#[test]
fn nested_elements_tabs_and_breaks_are_supported() {
    let header = story(
        br#"<w:hdr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Bold Centered Text</w:t></w:r><w:r><w:tab/><w:t>After</w:t></w:r></w:p></w:hdr>"#,
        Kind::Primary,
    );
    let text = header.text().unwrap();
    assert!(text.contains("Bold Centered Text"));
    assert!(text.contains("After"));
    assert_eq!(header.paragraph_count().unwrap(), 1);
}

#[test]
fn vml_watermarks_remain_available() {
    let xml = br##"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1" fillcolor="#808080"><v:textpath style="font-family:'Cambria';font-size:1pt" string="CONFIDENTIAL"/></v:shape></w:pict></w:r></w:p></w:hdr>"##;
    let header = story(xml, Kind::Primary);
    let watermarks = header.watermarks().unwrap();
    assert_eq!(watermarks.len(), 1);
    assert_eq!(watermarks[0].get_text(), "CONFIDENTIAL");
    assert_eq!(watermarks[0].color(), "#808080");
}
