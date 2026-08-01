use litchi_rtf::{CustomXmlAttribute, CustomXmlTag, RtfDocument, RtfWriter};
use std::borrow::Cow;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_namespaced_tag_with_attributes_and_round_trips() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\*\xmlnstbl {\xmlns1 urn:example:test}}"#,
        r#"{\xmlopen \xmlns1 employee}{\*\xmlattrname id}{\*\xmlattrvalue 5}Body{\xmlclose employee}}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "Body");
    let tags = document.custom_xml_tags();
    assert_eq!(tags.len(), 1);
    let tag = &tags[0];
    assert_eq!(tag.name, "employee");
    assert_eq!(tag.namespace, Some(1));
    assert_eq!(tag.attributes.len(), 1);
    assert_eq!(tag.attributes[0].name, "id");
    assert_eq!(tag.attributes[0].value, "5");
    assert_eq!(tag.position, 0);
    assert_eq!(tag.content, "Body");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), "Body");
    assert_eq!(reparsed.custom_xml_tags(), tags);
    assert_eq!(reparsed.xml_namespaces(), document.xml_namespaces());
}

#[test]
fn parses_nested_attributes_inside_xmlopen_and_nested_tags() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1 pre{\xmlopen {\*\xmlattrname a}{\*\xmlattrvalue b}outer}"#,
        r#"x{\xmlopen inner}y{\xmlclose inner}z{\xmlclose outer}post}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "prexyzpost");
    let tags = document.custom_xml_tags();
    assert_eq!(tags.len(), 2);
    assert_eq!(tags[0].name, "outer");
    assert_eq!(tags[0].namespace, None);
    assert_eq!(tags[0].attributes.len(), 1);
    assert_eq!(tags[0].attributes[0].name, "a");
    assert_eq!(tags[0].attributes[0].value, "b");
    assert_eq!(tags[0].position, 3);
    assert_eq!(tags[0].content, "xyz");
    assert_eq!(tags[1].name, "inner");
    assert_eq!(tags[1].position, 4);
    assert_eq!(tags[1].content, "y");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.custom_xml_tags(), tags);
}

#[test]
fn parses_empty_and_unicode_tag_content() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\xmlopen empty}{\xmlclose empty}"#,
        r#"{\xmlopen empl\u20320?}X{\xmlclose empl\u20320?}}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "X");
    let tags = document.custom_xml_tags();
    assert_eq!(tags.len(), 2);
    assert_eq!(tags[0].name, "empty");
    assert_eq!(tags[0].content, "");
    assert_eq!(tags[0].position, 0);
    assert_eq!(tags[1].name, "empl你");
    assert_eq!(tags[1].content, "X");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.custom_xml_tags(), tags);
}

#[test]
fn typed_constructors_validate() {
    assert!(CustomXmlTag::new(Cow::Borrowed(""), None, Vec::new(), 0, Cow::Borrowed("")).is_err());
    assert!(
        CustomXmlTag::new(
            Cow::Borrowed("a"),
            Some(0),
            Vec::new(),
            0,
            Cow::Borrowed("")
        )
        .is_err()
    );
    assert!(CustomXmlAttribute::new(Cow::Borrowed(" "), Cow::Borrowed("v")).is_err());
    let duplicate = vec![
        CustomXmlAttribute::new(Cow::Borrowed("n"), Cow::Borrowed("1")).unwrap(),
        CustomXmlAttribute::new(Cow::Borrowed("n"), Cow::Borrowed("2")).unwrap(),
    ];
    assert!(CustomXmlTag::new(Cow::Borrowed("a"), None, duplicate, 0, Cow::Borrowed("")).is_err());
}

#[test]
fn rejects_malformed_custom_xml_markup() {
    let cases = [
        // Unclosed tag.
        r#"{\rtf1{\xmlopen a}Body}"#,
        // Mismatched close name.
        r#"{\rtf1{\xmlopen a}Body{\xmlclose b}}"#,
        // Close without open.
        r#"{\rtf1 Body{\xmlclose a}}"#,
        // Empty tag name.
        r#"{\rtf1{\xmlopen }Body}"#,
        // Namespace reference without a namespace table.
        r#"{\rtf1{\xmlopen \xmlns1 a}B{\xmlclose a}}"#,
        // Namespace reference outside the valid range.
        r#"{\rtf1{\*\xmlnstbl {\xmlns1 urn:x}}{\xmlopen \xmlns0 a}B{\xmlclose a}}"#,
        // Unknown namespace reference.
        r#"{\rtf1{\*\xmlnstbl {\xmlns1 urn:x}}{\xmlopen \xmlns2 a}B{\xmlclose a}}"#,
        // Duplicate namespace selector.
        r#"{\rtf1{\*\xmlnstbl {\xmlns1 urn:x}}{\xmlopen \xmlns1\xmlns1 a}B{\xmlclose a}}"#,
        // Orphan attribute destinations.
        r#"{\rtf1{\*\xmlattrname id}Body}"#,
        r#"{\rtf1{\*\xmlattrvalue 5}Body}"#,
        // Attribute value without a name.
        r#"{\rtf1{\xmlopen a}{\*\xmlattrvalue 5}B{\xmlclose a}}"#,
        // Attribute name without a value.
        r#"{\rtf1{\xmlopen a}{\*\xmlattrname id}B{\xmlclose a}}"#,
        r#"{\rtf1{\xmlopen a}{\*\xmlattrname id}{\*\xmlattrname id2}B{\xmlclose a}}"#,
        r#"{\rtf1{\xmlopen a}{\*\xmlattrname id}B{\xmlclose a}{\*\xmlattrvalue 5}}"#,
        // Duplicate attribute names within one tag.
        r#"{\rtf1{\xmlopen a}{\*\xmlattrname n}{\*\xmlattrvalue 1}{\*\xmlattrname n}{\*\xmlattrvalue 2}B{\xmlclose a}}"#,
        // Starred tag destinations.
        r#"{\rtf1{\*\xmlopen a}Body{\*\xmlclose a}}"#,
        // Unstarred attribute destinations.
        r#"{\rtf1{\xmlopen a}{\xmlattrname id}B{\xmlclose a}}"#,
        // Binary data inside a tag destination.
        r#"{\rtf1{\xmlopen a\bin2 xx}B{\xmlclose a}}"#,
        // Nested group inside a close destination.
        r#"{\rtf1{\xmlopen a}B{\xmlclose a{x}}}"#,
        // Improper nesting.
        r#"{\rtf1{\xmlopen a}{\xmlopen b}B{\xmlclose a}{\xmlclose b}}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn rejects_excessive_nesting_depth() {
    let mut rtf = String::from(r#"{\rtf1"#);
    for _ in 0..65 {
        rtf.push_str(r"{\xmlopen t}");
    }
    rtf.push('B');
    for _ in 0..65 {
        rtf.push_str(r"{\xmlclose t}");
    }
    rtf.push('}');
    assert!(RtfDocument::parse(&rtf).is_err());

    let mut valid = String::from(r#"{\rtf1"#);
    for _ in 0..64 {
        valid.push_str(r"{\xmlopen t}");
    }
    valid.push('B');
    for _ in 0..64 {
        valid.push_str(r"{\xmlclose t}");
    }
    valid.push('}');
    let document = RtfDocument::parse(&valid).unwrap();
    assert_eq!(document.custom_xml_tags().len(), 64);
}

#[test]
fn rejects_custom_xml_markup_in_non_body_stories() {
    let cases = [
        // Footnote story.
        r"{\rtf1 body{\footnote{\xmlopen t}note{\xmlclose t}}}",
        r"{\rtf1 body{\footnote{\*\xmlattrname a}note}}",
        // Endnote story.
        r"{\rtf1 body{\endnote{\xmlopen t}note{\xmlclose t}}}",
        // Header and footer stories.
        r"{\rtf1{\header{\xmlopen t}h{\xmlclose t}}body}",
        r"{\rtf1{\footer{\*\xmlattrvalue v}f}body}",
        // Shape text story.
        r#"{\rtf1 body{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt{\xmlopen t}s{\xmlclose t}}}}}"#,
        r#"{\rtf1 body{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt{\*\xmlattrname a}s}}}}"#,
        // Field instruction and result stories.
        r"{\rtf1 body{\field{\*\fldinst X}{\fldrslt{\xmlopen t}r{\xmlclose t}}}}",
        r"{\rtf1 body{\field{\*\fldinst{\xmlopen t}X{\xmlclose t}}{\fldrslt r}}}",
        r"{\rtf1 body{\field{\*\fldinst X}{\*\xmlattrname a}{\fldrslt r}}}",
    ];
    for rtf in cases {
        assert!(
            RtfDocument::parse(rtf).is_err(),
            "accepted non-body custom XML markup {rtf}"
        );
    }
}

#[test]
fn coexists_with_bookmarks_and_body_markup() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\*\bkmkstart bm}pre{\*\bkmkend bm}"#,
        r#"{\xmlopen wrap}\b bold\b0{\xmlclose wrap}}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "prebold");
    let bookmark = &document.bookmarks().bookmarks()[0];
    assert_eq!(bookmark.content, "pre");
    let tags = document.custom_xml_tags();
    assert_eq!(tags.len(), 1);
    assert_eq!(tags[0].position, 3);
    assert_eq!(tags[0].content, "bold");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.custom_xml_tags(), tags);
    assert_eq!(
        reparsed.bookmarks().bookmarks()[0].content,
        bookmark.content
    );
}
