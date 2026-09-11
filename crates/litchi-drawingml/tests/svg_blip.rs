use litchi_drawingml::svg_blip::{
    Attribute, MAX_RELATIONSHIP_ID_BYTES, NAMESPACE, Namespace, Reference, RelationshipId, SvgBlip,
    XML_NAMESPACE, read, write,
};

const SVG_BLIP: &[u8] = br#"<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:future="urn:example:future" r:embed="rIdSvg" future:mode="preserve"><future:payload data="x"/></asvg:svgBlip>"#;

#[test]
fn reads_typed_relationship_and_preserves_unknown_markup() {
    let value = read(SVG_BLIP).expect("SVG blip fixture must parse");
    assert_eq!(
        value.embedded().expect("embedded relationship").as_str(),
        "rIdSvg"
    );
    assert!(value.linked().is_none());
    assert_eq!(value.source(), Some(SVG_BLIP));
    assert!(
        value
            .namespaces()
            .iter()
            .any(|item| item.uri() == NAMESPACE)
    );
    assert_eq!(value.attributes()[0].name(), "future:mode");
    assert_eq!(value.attributes()[0].value(), "preserve");
    assert_eq!(value.children().len(), 1);
    assert_eq!(
        value.children()[0].as_bytes(),
        br#"<future:payload data="x"/>"#
    );
    assert_eq!(write(&value).expect("source-backed write"), SVG_BLIP);
}

#[test]
fn changing_relationship_keeps_future_content() {
    let mut value = read(SVG_BLIP).expect("SVG blip fixture must parse");
    value
        .set_reference(Reference::linked("rIdExternal").expect("relationship ID"))
        .expect("valid relationship reference");
    let output = write(&value).expect("modified SVG blip write");
    let output = String::from_utf8(output).expect("XML output is UTF-8");
    assert!(output.contains("r:link=\"rIdExternal\""));
    assert!(output.contains("future:mode=\"preserve\""));
    assert!(output.contains("<future:payload data=\"x\"/>"));
    let reparsed = read(output.as_bytes()).expect("modified SVG blip must parse");
    assert_eq!(
        reparsed.linked().expect("linked relationship").as_str(),
        "rIdExternal"
    );
    assert!(reparsed.embedded().is_none());
}

#[test]
fn rejects_invalid_root_and_relationship_values() {
    let wrong_root =
        br#"<asvg:other xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"/>"#;
    assert!(read(wrong_root).is_err());
    let duplicate_reference = br#"<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="same" r:link="same"/>"#;
    assert!(read(duplicate_reference).is_err());
    let too_long = "r".repeat(MAX_RELATIONSHIP_ID_BYTES + 1);
    assert!(RelationshipId::new(too_long).is_err());
    assert!(Namespace::new(Some("xml"), "urn:invalid").is_err());
    assert!(Namespace::new(Some("xmlns"), "urn:invalid").is_err());
    assert!(Namespace::new(None, XML_NAMESPACE).is_err());
    assert!(Attribute::new("xmlns:future", "urn:invalid").is_err());
}

#[test]
fn constructs_empty_svg_blip_with_typed_reference() {
    let reference = Reference::embedded("rId1").expect("relationship ID");
    let value = SvgBlip::new(reference).expect("valid SVG blip");
    let output = write(&value).expect("new SVG blip write");
    assert!(output.starts_with(b"<asvg:svgBlip"));
    assert!(
        output
            .windows(b"r:embed=\"rId1\"".len())
            .any(|window| window == b"r:embed=\"rId1\"")
    );
}

#[test]
fn edits_default_namespace_and_alternate_relationship_prefix_without_unbound_names() {
    let source = concat!(
        r#"<svgBlip xmlns="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:r="urn:foreign" xmlns:future="urn:example:future" rel:embed="rIdSvg"><future:payload/>"#,
        " \n",
        r#"<!--preserve--><![CDATA[future]]></svgBlip>"#
    ).as_bytes();
    let mut value = read(source).expect("default SVG blip fixture must parse");
    value
        .set_reference(Reference::linked("rIdExternal").expect("relationship ID"))
        .expect("valid relationship reference");
    let output = write(&value).expect("modified SVG blip write");
    let output = String::from_utf8(output).expect("XML output is UTF-8");
    assert!(output.starts_with("<svgBlip xmlns=\""));
    assert!(output.contains("rel:link=\"rIdExternal\""));
    assert!(!output.contains(" r:link="));
    assert!(output.contains("<!--preserve-->"));
    assert!(output.contains("<![CDATA[future]]>"));
    assert_eq!(
        read(output.as_bytes())
            .expect("edited default SVG blip")
            .linked()
            .unwrap()
            .as_str(),
        "rIdExternal"
    );
}

#[test]
fn accepts_ancestor_bound_root_prefix_for_unknown_attributes() {
    let source = br#"<asvg:svgBlip asvg:future="x" r:embed="rIdSvg"/>"#;
    let mut value = read(source).expect("ancestor-bound SVG fragment");
    value
        .set_reference(Reference::linked("rIdExternal").expect("relationship ID"))
        .expect("valid relationship reference");
    let output = write(&value).expect("edited ancestor-bound SVG fragment");
    assert!(output.starts_with(b"<asvg:svgBlip"));
    assert!(
        output
            .windows(b"xmlns:asvg=\"".len())
            .any(|window| window == b"xmlns:asvg=\"")
    );
    assert!(
        output
            .windows(b"asvg:future=\"x\"".len())
            .any(|window| window == b"asvg:future=\"x\"")
    );
    assert!(read(&output).is_ok());
}

#[test]
fn rejects_non_xml_content_after_the_root() {
    let trailing_text = [SVG_BLIP, b"trailing"].concat();
    assert!(read(&trailing_text).is_err());
    let trailing_cdata = [SVG_BLIP, b"<![CDATA[trailing]]>"].concat();
    assert!(read(&trailing_cdata).is_err());
    let trailing_reference = [SVG_BLIP, b"&future;"].concat();
    assert!(read(&trailing_reference).is_err());
}

#[test]
fn retains_default_namespace_undeclaration_through_relationship_edits() {
    let source = br#"<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns="" r:embed="rIdSvg"><future/></asvg:svgBlip>"#;
    let mut value = read(source).expect("a default namespace can be cleared");
    assert_eq!(write(&value).unwrap(), source);
    assert!(
        value
            .namespaces()
            .iter()
            .any(|namespace| { namespace.prefix().is_none() && namespace.uri().is_empty() })
    );
    value
        .set_reference(Reference::embedded("rIdUpdated").unwrap())
        .unwrap();
    let output = write(&value).unwrap();
    let reopened = read(&output).unwrap();
    assert_eq!(reopened.embedded().unwrap().as_str(), "rIdUpdated");
    assert!(
        reopened
            .namespaces()
            .iter()
            .any(|namespace| { namespace.prefix().is_none() && namespace.uri().is_empty() })
    );
    assert_eq!(reopened.children()[0].as_bytes(), b"<future/>");
    assert!(Namespace::new(None, "").is_ok());
    assert!(Namespace::new(Some(""), "").is_ok());
    assert!(Namespace::new(Some("future"), "").is_err());
    assert!(read(br#"<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:future="" future:attr="x"/>"#).is_err());
}
