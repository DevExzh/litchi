use litchi_drawingml::svg_blip::{
    Attribute, MAX_CONTEXT_DEPTH, MAX_NAMESPACE_DECLARATIONS, MAX_RELATIONSHIP_ID_BYTES, NAMESPACE,
    Namespace, NamespaceContext, RELATIONSHIP_NAMESPACE, Reference, RelationshipId, SvgBlip,
    XML_NAMESPACE, read, read_contextual, write, write_contextual, write_contextual_to,
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
fn standalone_reads_keep_source_equality_and_exact_write() {
    let first = read(SVG_BLIP).expect("first standalone SVG blip");
    let second = read(SVG_BLIP).expect("second standalone SVG blip");
    assert_eq!(first, second);
    assert_eq!(first.source(), Some(SVG_BLIP));
    assert_eq!(write(&first).unwrap(), SVG_BLIP);
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
fn contextual_read_shares_scope_and_keeps_raw_source_distinct() {
    let context = NamespaceContext::empty()
        .child([
            Namespace::new(Some("asvg"), NAMESPACE).unwrap(),
            Namespace::new(Some("r"), RELATIONSHIP_NAMESPACE).unwrap(),
            Namespace::new(Some("future"), "urn:example:future").unwrap(),
            Namespace::new(Some("opaque"), "urn:example:opaque").unwrap(),
        ])
        .unwrap();
    let source = br#"<asvg:svgBlip r:embed="rIdSvg" future:marker="opaque:value"><future:payload data="future:value"/></asvg:svgBlip>"#;
    let first = read_contextual(source, &context).expect("contextual SVG blip");
    let second = read_contextual(source, &context).expect("second contextual SVG blip");

    assert_eq!(first.source(), None);
    assert_eq!(first.raw_source(), Some(source.as_slice()));
    assert_eq!(first.namespace_context(), Some(&context));
    assert_eq!(first.namespace_context().unwrap().binding_count(), 4);
    assert!(
        first
            .namespace_context()
            .unwrap()
            .shares_storage(second.namespace_context().unwrap())
    );

    let exported = write_contextual(&first, 4096).expect("standalone contextual export");
    assert!(
        exported
            .windows(b"xmlns:opaque=\"urn:example:opaque\"".len())
            .any(|window| window == b"xmlns:opaque=\"urn:example:opaque\"")
    );
    assert!(
        exported
            .windows(b"future:marker=\"opaque:value\"".len())
            .any(|window| window == b"future:marker=\"opaque:value\"")
    );
    assert_eq!(
        read(&exported).unwrap().embedded().unwrap().as_str(),
        "rIdSvg"
    );

    let mut sink = Vec::new();
    assert!(write_contextual_to(&mut sink, &first, source.len()).is_err());
    assert!(sink.is_empty(), "cap refusal must precede sink writes");
}

#[test]
fn contextual_edits_preserve_opaque_qnames_shadowing_and_undeclarations() {
    let context = NamespaceContext::empty()
        .child([
            Namespace::new(Some("asvg"), NAMESPACE).unwrap(),
            Namespace::new(Some("r"), RELATIONSHIP_NAMESPACE).unwrap(),
            Namespace::new(Some("future"), "urn:outer").unwrap(),
            Namespace::new(None, "urn:outer-default").unwrap(),
        ])
        .unwrap()
        .child([
            Namespace::new(Some("future"), "urn:inner").unwrap(),
            Namespace::new(None, "").unwrap(),
        ])
        .unwrap();
    let source = br#"<asvg:svgBlip r:embed="rIdSvg" future:marker="future:value"><future:payload data="future:value"/></asvg:svgBlip>"#;
    let mut value = read_contextual(source, &context).expect("contextual SVG blip");
    value
        .set_reference(Reference::linked("rIdExternal").unwrap())
        .unwrap();
    assert_eq!(value.source(), None);
    assert_eq!(value.raw_source(), Some(source.as_slice()));

    let exported = write(&value).expect("edited contextual export");
    let text = String::from_utf8(exported.clone()).unwrap();
    assert!(text.contains("xmlns:future=\"urn:inner\""));
    assert!(text.contains("xmlns=\"\""));
    assert!(text.contains("future:marker=\"future:value\""));
    assert!(text.contains("<future:payload data=\"future:value\"/>"));
    assert!(text.contains("r:link=\"rIdExternal\""));
    assert_eq!(
        read(&exported).unwrap().linked().unwrap().as_str(),
        "rIdExternal"
    );
}

#[test]
fn contextual_rejects_unbound_opaque_descendant_prefixes() {
    let context = NamespaceContext::empty()
        .child([Namespace::new(Some("asvg"), NAMESPACE).unwrap()])
        .unwrap();
    let source = br#"<asvg:svgBlip><future:payload/></asvg:svgBlip>"#;
    assert!(read_contextual(source, &context).is_err());
}

#[test]
fn contextual_does_not_apply_default_namespace_to_unprefixed_attributes() {
    let context = NamespaceContext::empty()
        .child([
            Namespace::new(Some("asvg"), NAMESPACE).unwrap(),
            Namespace::new(None, RELATIONSHIP_NAMESPACE).unwrap(),
        ])
        .unwrap();
    let source = br#"<asvg:svgBlip embed="not-a-relationship"/>"#;
    let value = read_contextual(source, &context).expect("unprefixed attribute is opaque");
    assert!(value.embedded().is_none());
    assert_eq!(value.attributes()[0].name(), "embed");
}

#[test]
fn contextual_shadowed_relationship_prefix_is_not_redeclared() {
    let context = NamespaceContext::empty()
        .child([
            Namespace::new(Some("asvg"), NAMESPACE).unwrap(),
            Namespace::new(Some("r"), RELATIONSHIP_NAMESPACE).unwrap(),
            Namespace::new(Some("rel"), RELATIONSHIP_NAMESPACE).unwrap(),
            Namespace::new(Some("future"), "urn:future").unwrap(),
        ])
        .unwrap();
    let source = br#"<asvg:svgBlip xmlns:r="urn:foreign" rel:embed="rIdSvg"><future:payload/></asvg:svgBlip>"#;
    let mut value = read_contextual(source, &context).expect("shadowed contextual SVG blip");
    value
        .set_reference(Reference::linked("rIdExternal").unwrap())
        .unwrap();
    let output = String::from_utf8(write(&value).unwrap()).unwrap();
    assert_eq!(output.matches("xmlns:r=").count(), 1);
    assert!(output.contains("xmlns:rel="));
    assert!(output.contains("rel:link=\"rIdExternal\""));
    assert!(!output.contains("r:link="));
    assert!(read(output.as_bytes()).is_ok());
}

#[test]
fn contextual_export_refuses_too_many_root_declarations() {
    let mut context = NamespaceContext::empty();
    context = context
        .child([Namespace::new(Some("asvg"), NAMESPACE).unwrap()])
        .unwrap();
    for index in 0..MAX_NAMESPACE_DECLARATIONS {
        let prefix = format!("p{index}");
        context = context
            .child([Namespace::new(Some(&prefix), format!("urn:{index}")).unwrap()])
            .unwrap();
    }
    let source = br#"<asvg:svgBlip/>"#;
    let value = read_contextual(source, &context).expect("contextual SVG blip");
    assert!(write_contextual(&value, usize::MAX).is_err());
}

#[test]
fn contextual_context_depth_is_bounded() {
    let mut context = NamespaceContext::empty();
    for index in 0..MAX_CONTEXT_DEPTH {
        let prefix = format!("p{index}");
        context = context
            .child([Namespace::new(Some(&prefix), format!("urn:{index}")).unwrap()])
            .unwrap();
    }
    assert!(context.depth() <= MAX_CONTEXT_DEPTH);
    let prefix = format!("p{MAX_CONTEXT_DEPTH}");
    assert!(
        context
            .child([Namespace::new(Some(&prefix), "urn:overflow").unwrap()])
            .is_err()
    );
}

#[test]
fn contextual_sink_fallback_checks_cap_before_source_write() {
    let value = read(SVG_BLIP).expect("standalone SVG blip");
    let mut sink = Vec::new();
    assert!(write_contextual_to(&mut sink, &value, SVG_BLIP.len() - 1).is_err());
    assert!(sink.is_empty());
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
