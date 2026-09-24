//! Focused source-preserving coverage for the exact ODRAWXML §2.44 owner.

use litchi_drawingml::chart::extension::formatcode2::{
    Element, MAX_VALUE_BYTES, NAMESPACE, Value, read, read_attribute, read_attribute_with_bindings,
    write, write_attribute, write_attribute_value,
};

const PREFIXED: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<!-- keep this comment -->
<c16r2:formatcode2 xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">[$-en-US]#,##0.00</c16r2:formatcode2>
"#;

#[test]
fn read_write_noop_preserves_exact_source_and_scalar_edit_keeps_comments() {
    let mut value = read(PREFIXED).expect("valid formatcode2 fragment");
    assert_eq!(value.value(), "[$-en-US]#,##0.00");
    assert_eq!(value.source(), Some(PREFIXED));
    assert_eq!(write(&value).unwrap(), PREFIXED);

    value.set_value("[$-fr-FR]#,##0.00").unwrap();
    let output = write(&value).unwrap();
    assert!(output.starts_with(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"));
    assert!(
        output
            .windows(b"keep this comment".len())
            .any(|window| window == b"keep this comment")
    );
    assert!(
        output
            .windows(NAMESPACE.len())
            .any(|window| window == NAMESPACE.as_bytes())
    );
    assert!(
        std::str::from_utf8(&output)
            .unwrap()
            .contains("[$-fr-FR]#,##0.00")
    );
    assert_eq!(read(&output).unwrap().value(), "[$-fr-FR]#,##0.00");
}

#[test]
fn cdata_and_comments_are_preserved_during_scalar_edit() {
    let source = br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart"><![CDATA[General]]><!--opaque--></f:formatcode2>"#;
    let mut value = read(source).unwrap();
    assert_eq!(value.value(), "General");
    value.set_value("0.00").unwrap();
    let output = write(&value).unwrap();
    assert!(
        std::str::from_utf8(&output)
            .unwrap()
            .contains("<![CDATA[0.00]]>")
    );
    assert!(
        std::str::from_utf8(&output)
            .unwrap()
            .contains("<!--opaque-->")
    );
    assert_eq!(read(&output).unwrap().value(), "0.00");
}

#[test]
fn entity_and_empty_root_edits_keep_the_scalar_semantics() {
    let entity_source = br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">&amp;<!--opaque--></f:formatcode2>"#;
    let mut entity = read(entity_source).unwrap();
    assert_eq!(entity.value(), "&");
    entity.set_value("0").unwrap();
    let entity_output = write(&entity).unwrap();
    assert!(
        std::str::from_utf8(&entity_output)
            .unwrap()
            .contains("<!--opaque-->")
    );
    assert_eq!(read(&entity_output).unwrap().value(), "0");

    let empty_source =
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart"/>"#;
    let mut empty = read(empty_source).unwrap();
    assert_eq!(empty.value(), "");
    empty.set_value("General").unwrap();
    let empty_output = write(&empty).unwrap();
    assert_eq!(read(&empty_output).unwrap().value(), "General");
}

#[test]
fn st_xstring_escapes_decode_once_and_reencode_controls() {
    let source = br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">_x0041__x0001_</f:formatcode2>"#;
    let parsed = read(source).unwrap();
    assert_eq!(parsed.value(), "A\u{1}");

    let value = Element::new("literal _x0041_\u{1}\u{fffe}").unwrap();
    let output = write(&value).unwrap();
    let lexical = std::str::from_utf8(&output).unwrap();
    assert!(lexical.contains("_x005F_x0041__x0001__xFFFE_"));
    assert_eq!(read(&output).unwrap().value(), value.value());
}

#[test]
fn qualified_attribute_owner_preserves_host_start_tag_and_edits_only_value() {
    let source = br#"<c:numFmt xmlns:c='http://schemas.microsoft.com/office/drawing/2015/06/chart' formatCode='0' c:formatcode2='_x0041_' data='keep'>"#;
    let mut attribute = read_attribute(source).unwrap();
    assert_eq!(attribute.value(), "A");
    assert_eq!(attribute.source(), Some(source.as_slice()));
    assert_eq!(write_attribute(&attribute).unwrap(), source);

    attribute.set_value("0\u{1}").unwrap();
    let output = write_attribute(&attribute).unwrap();
    let output_text = std::str::from_utf8(&output).unwrap();
    assert!(output_text.contains("formatCode='0'"));
    assert!(output_text.contains("data='keep'"));
    assert!(output_text.contains("c:formatcode2='0_x0001_'"));
    assert_eq!(read_attribute(&output).unwrap().value(), "0\u{1}");

    for invalid in [
        br#"<c:numFmt xmlns:c="http://schemas.microsoft.com/office/drawing/2015/06/chart" formatcode2="x">"#.as_slice(),
        br#"<c:numFmt xmlns:c="urn:other" c:formatcode2="x">"#,
        br#"<c:numFmt c:formatcode2="x">"#,
        br#"<c:numFmt xmlns:c="http://schemas.microsoft.com/office/drawing/2015/06/chart" c:formatcode2="x"></c:numFmt>"#,
    ] {
        assert!(read_attribute(invalid).is_err(), "accepted invalid attribute tag");
    }
}

#[test]
fn qualified_attribute_can_use_explicit_inherited_bindings_without_source_decoration() {
    let source = br#"<c:numFmt c:formatcode2='General' data='keep'/>"#;
    let mut attribute = read_attribute_with_bindings(source, &[("c", NAMESPACE)]).unwrap();
    assert_eq!(attribute.value(), "General");
    assert_eq!(attribute.source(), Some(source.as_slice()));
    assert_eq!(write_attribute(&attribute).unwrap(), source);

    attribute.set_value("0.00").unwrap();
    let edited = write_attribute(&attribute).unwrap();
    assert_eq!(edited, br#"<c:numFmt c:formatcode2='0.00' data='keep'/>"#);
    assert_eq!(
        read_attribute_with_bindings(&edited, &[("c", NAMESPACE)])
            .unwrap()
            .value(),
        "0.00"
    );

    let local = br#"<c:numFmt xmlns:c="urn:local" c:formatcode2="x"/>"#;
    assert!(read_attribute_with_bindings(local, &[("c", NAMESPACE)]).is_err());
    let inherited = br#"<c:numFmt c:formatcode2="x"/>"#;
    assert!(read_attribute_with_bindings(inherited, &[("c", "urn:other")]).is_err());
}

#[test]
fn detached_write_is_bounded_and_reopens() {
    let value = Element::new("[$-ja-JP]#,##0").unwrap();
    let output = write(&value).unwrap();
    assert!(output.len() < MAX_VALUE_BYTES);
    assert_eq!(read(&output).unwrap().value(), value.value());
}

#[test]
fn namespace_and_simple_content_grammar_fail_closed() {
    let cases = [
        br#"<formatcode2 xmlns="http://schemas.microsoft.com/office/drawing/2015/06/chart" bad="1">x</formatcode2>"#.as_slice(),
        br#"<f:formatcode2 xmlns:f="urn:other">x</f:formatcode2>"#,
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart"><f:x/></f:formatcode2>"#,
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart" xmlns:x="" x:y="z">x</f:formatcode2>"#,
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart" xmlns="http://www.w3.org/XML/1998/namespace">x</f:formatcode2>"#,
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">&unknown;</f:formatcode2>"#,
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart"><</f:formatcode2>"#,
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">]]></f:formatcode2>"#,
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart" xmlns:f="urn:duplicate">x</f:formatcode2>"#,
        b"<f:formatcode2 xmlns:f=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">\x01</f:formatcode2>",
    ];
    for source in cases {
        assert!(
            read(source).is_err(),
            "accepted invalid fragment: {:?}",
            source
        );
    }
}

#[test]
fn declaration_rules_are_bounded() {
    let fragment = br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">x</f:formatcode2>"#;
    let valid = [
        br#"<?xml version="1.0" encoding="UTF-8"?>"#.as_slice(),
        fragment,
    ]
    .concat();
    assert!(read(&valid).is_ok());
    for declaration in [
        b"<?xml version=\"1.1\"?>".as_slice(),
        b"<?xml version=\"2.0\"?>".as_slice(),
        b"<?xml encoding=\"wat\"?>".as_slice(),
    ] {
        let mut source = declaration.to_vec();
        source.extend_from_slice(fragment);
        assert!(read(&source).is_err(), "accepted invalid declaration");
    }
}

#[test]
fn namespace_entities_are_decoded_before_owner_matching() {
    let escaped_namespace = NAMESPACE.replace('/', "&#x2F;");
    let element =
        format!(r#"<f:formatcode2 xmlns:f="{escaped_namespace}">General</f:formatcode2>"#);
    assert_eq!(read(element.as_bytes()).unwrap().value(), "General");

    let attribute = format!(r#"<series xmlns:f="{escaped_namespace}" f:formatcode2="General"/>"#);
    assert_eq!(
        read_attribute(attribute.as_bytes()).unwrap().value(),
        "General"
    );

    let duplicate_expanded =
        br#"<series xmlns:a="urn:x&#x2F;y" xmlns:b="urn:x/y" a:x="1" b:x="2"/>"#;
    assert!(read_attribute(duplicate_expanded).is_err());
}

#[test]
fn attribute_whitespace_uses_xml_character_references() {
    let value = Value::new("a\t\n\r").unwrap();
    assert_eq!(write_attribute_value(&value).unwrap(), b"a&#x9;&#xA;&#xD;");
}

#[test]
fn comments_and_processing_instructions_are_xml_checked_but_preserved() {
    let source = br#"<!-- <? legal-comment --><f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">General<?xml-stylesheet href="x"?></f:formatcode2>"#;
    assert_eq!(write(&read(source).unwrap()).unwrap(), source);

    let root = br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">General</f:formatcode2>"#;
    let invalid_comment = [b"<!--\x01-->".as_slice(), root].concat();
    assert!(read(&invalid_comment).is_err());
    let invalid_pi = [b"<?1bad?>".as_slice(), root].concat();
    assert!(read(&invalid_pi).is_err());
    let reserved_pi = [b"<?XML?>".as_slice(), root].concat();
    assert!(read(&reserved_pi).is_err());
}
