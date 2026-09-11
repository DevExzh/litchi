//! Adversarial grammar, XString, source-retention, and budget coverage for
//! [MS-ODRAWXML] §2.44 formatcode2.

use formatcode2::{
    Element, MAX_ATTRIBUTES, MAX_DEPTH, MAX_NAMESPACE_BYTES, MAX_NODES, MAX_VALUE_BYTES,
    MAX_XML_BYTES, NAMESPACE, read, read_attribute, read_attribute_shared,
    read_attribute_shared_with_bindings, read_attribute_with_bindings, read_shared, write,
    write_attribute, write_attribute_to, write_to,
};
use litchi_drawingml::{Error, chart::extension::formatcode2};

const ROOT_PREFIX: &str = "f";
const MAX_LEXICAL_VALUE_BYTES: usize = MAX_VALUE_BYTES * 7;

fn encoded_namespace() -> String {
    NAMESPACE.replace('/', "&#x2F;")
}

fn fragment(value: &str) -> Vec<u8> {
    format!(
        r#"<{ROOT_PREFIX}:formatcode2 xmlns:{ROOT_PREFIX}="{NAMESPACE}">{value}</{ROOT_PREFIX}:formatcode2>"#
    )
    .into_bytes()
}

fn assert_invalid(source: &[u8], expected: &str) {
    match read(source) {
        Err(Error::Invalid(message)) => assert_eq!(message, expected),
        Err(other) => panic!("expected Invalid({expected:?}), got {other:?}"),
        Ok(value) => panic!("accepted invalid formatcode2 value {:?}", value.value()),
    }
}

fn assert_xml_error(source: &[u8]) {
    match read(source) {
        Err(Error::Xml(_)) => {},
        Err(other) => panic!("expected XML parser refusal, got {other:?}"),
        Ok(value) => panic!("accepted malformed formatcode2 value {:?}", value.value()),
    }
}

fn assert_rejected(source: &[u8]) {
    match read(source) {
        Err(Error::Invalid(_) | Error::Xml(_)) => {},
        Err(other) => panic!("expected malformed XML refusal, got {other:?}"),
        Ok(value) => panic!("accepted malformed formatcode2 value {:?}", value.value()),
    }
}

fn assert_limit(source: &[u8], resource: &'static str, limit: usize) {
    match read(source) {
        Err(Error::Limit {
            resource: actual_resource,
            limit: actual_limit,
        }) => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_limit, limit);
        },
        Err(other) => panic!("expected Limit({resource:?}, {limit}), got {other:?}"),
        Ok(value) => panic!("accepted over-budget formatcode2 value {:?}", value.value()),
    }
}

fn assert_attribute_limit(source: &[u8], resource: &'static str, limit: usize) {
    match read_attribute(source) {
        Err(Error::Limit {
            resource: actual_resource,
            limit: actual_limit,
        }) => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_limit, limit);
        },
        Err(other) => panic!("expected Limit({resource:?}, {limit}), got {other:?}"),
        Ok(value) => panic!(
            "accepted over-budget formatcode2 attribute value {:?}",
            value.value()
        ),
    }
}

fn assert_attribute_binding_limit(
    source: &[u8],
    bindings: &[(&str, &str)],
    resource: &'static str,
    limit: usize,
) {
    match read_attribute_with_bindings(source, bindings) {
        Err(Error::Limit {
            resource: actual_resource,
            limit: actual_limit,
        }) => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_limit, limit);
        },
        Err(other) => panic!("expected Limit({resource:?}, {limit}), got {other:?}"),
        Ok(value) => panic!(
            "accepted over-budget inherited formatcode2 attribute value {:?}",
            value.value()
        ),
    }
}

fn assert_write_limit(value: &Element, resource: &'static str, limit: usize) {
    match write(value) {
        Err(Error::Limit {
            resource: actual_resource,
            limit: actual_limit,
        }) => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_limit, limit);
        },
        Err(other) => panic!("expected Limit({resource:?}, {limit}), got {other:?}"),
        Ok(output) => panic!(
            "wrote over-budget formatcode2 output of {} bytes",
            output.len()
        ),
    }
}

#[test]
fn section_244_namespace_and_xsd_string_whitespace_are_semantic() {
    let source = b"<f:formatcode2 xmlns:f=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\"> \t\n\r  [$-en-US]#,##0.00  </f:formatcode2>";
    let value = read(source).expect("valid section 2.44 element");
    assert_eq!(value.value(), " \t\n\r  [$-en-US]#,##0.00  ");
    assert_eq!(value.source(), Some(source.as_slice()));
    assert_eq!(write(&value).unwrap(), source);

    let detached = Element::new(" \t\n\r  [$-en-US]#,##0.00  ").unwrap();
    assert_eq!(detached.value(), value.value());
    assert_eq!(detached, value);
}

#[test]
fn encoded_namespace_uri_references_resolve_for_elements_attributes_and_duplicates() {
    let encoded = encoded_namespace();
    let element_source = format!(r#"<f:formatcode2 xmlns:f="{encoded}">General</f:formatcode2>"#);
    let element = read(element_source.as_bytes()).expect("encoded element namespace URI");
    assert_eq!(element.value(), "General");
    assert_eq!(write(&element).unwrap(), element_source.as_bytes());

    let attribute_source = format!(r#"<c:series xmlns:c="{encoded}" c:formatcode2="General"/>"#);
    let mut attribute =
        read_attribute(attribute_source.as_bytes()).expect("encoded attribute namespace URI");
    assert_eq!(attribute.value(), "General");
    assert_eq!(
        write_attribute(&attribute).unwrap(),
        attribute_source.as_bytes()
    );
    attribute.set_value("0.00").unwrap();
    let edited = write_attribute(&attribute).unwrap();
    assert_eq!(read_attribute(&edited).unwrap().value(), "0.00");
    assert!(
        std::str::from_utf8(&edited)
            .unwrap()
            .contains(r#"c:formatcode2="0.00""#)
    );

    let duplicate_element = format!(
        r#"<f:formatcode2 xmlns:f="{NAMESPACE}" xmlns:a="{NAMESPACE}" xmlns:b="{encoded}" a:opaque="1" b:opaque="2">x</f:formatcode2>"#
    );
    assert_invalid(
        duplicate_element.as_bytes(),
        "formatcode2 element has duplicate expanded attributes",
    );

    let duplicate_attribute = format!(
        r#"<c:series xmlns:c="urn:host" xmlns:a="{NAMESPACE}" xmlns:b="{encoded}" a:opaque="1" b:opaque="2" a:formatcode2="General"/>"#
    );
    match read_attribute(duplicate_attribute.as_bytes()) {
        Err(Error::Invalid(message)) => {
            assert_eq!(
                message,
                "formatcode2 element has duplicate expanded attributes"
            )
        },
        other => panic!("expected encoded duplicate expanded attributes refusal, got {other:?}"),
    }
}

#[test]
fn bom_declarations_comments_and_processing_instructions_follow_xml_order() {
    let root = fragment("General");
    let mut bom = vec![0xEF, 0xBB, 0xBF];
    bom.extend_from_slice(&root);
    let value = read(&bom).expect("UTF-8 BOM is allowed");
    assert_eq!(value.value(), "General");
    assert_eq!(write(&value).unwrap(), bom);

    let with_declaration = [
        b"\xEF\xBB\xBF<?xml version=\"1.0\" encoding=\"UTF-8\"?>".as_slice(),
        b"<!--before--><?before?>".as_slice(),
        root.as_slice(),
        b"<?after?><!--after-->".as_slice(),
    ]
    .concat();
    let value = read(&with_declaration).expect("well-formed declaration and surrounding markup");
    assert_eq!(value.value(), "General");
    assert_eq!(write(&value).unwrap(), with_declaration);

    let mut declaration_after_comment = b"<!--before--><?xml version=\"1.0\"?>".to_vec();
    declaration_after_comment.extend_from_slice(&root);
    assert_invalid(
        &declaration_after_comment,
        "formatcode2 XML declaration is misplaced",
    );

    for declaration in [
        b"<?xml version=\"1.1\"?>".as_slice(),
        b"<?xml version=\"2.0\"?>".as_slice(),
    ] {
        let source = [declaration, root.as_slice()].concat();
        assert_invalid(&source, "formatcode2 XML declaration must use version 1.0");
    }
    let source = [
        b"<?xml version=\"1.0\" encoding=\"UTF-16\"?>".as_slice(),
        root.as_slice(),
    ]
    .concat();
    assert_invalid(
        &source,
        "formatcode2 XML declaration must use UTF-8 encoding",
    );

    let duplicate = [
        b"<?xml version=\"1.0\"?><?xml version=\"1.0\"?>".as_slice(),
        root.as_slice(),
    ]
    .concat();
    assert_invalid(&duplicate, "formatcode2 XML declaration is misplaced");

    for declaration in [
        r#"<?xml version="1.0"?>"#,
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        r#"<?xml version="1.0" standalone="yes"?>"#,
        r#"<?xml version="1.0" standalone="no"?>"#,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="no"?>"#,
    ] {
        let source = [declaration.as_bytes(), root.as_slice()].concat();
        let value = read(&source).expect("allowed XML declaration controls");
        assert_eq!(write(&value).unwrap(), source);
    }

    for (declaration, expected) in [
        (
            r#"<?xml version="1.0" standalone="maybe"?>"#,
            "formatcode2 XML declaration standalone must be yes or no",
        ),
        (
            r#"<?xml version="1.0" unknownfoo="x"?>"#,
            "formatcode2 XML declaration has an unknown attribute",
        ),
    ] {
        let source = [declaration.as_bytes(), root.as_slice()].concat();
        assert_invalid(&source, expected);
    }
    let duplicate_version = [
        br#"<?xml version="1.0" version="1.0"?>"#.as_slice(),
        root.as_slice(),
    ]
    .concat();
    assert_xml_error(&duplicate_version);
}

#[test]
fn xstring_references_controls_and_literal_escape_syntax_round_trip() {
    let source = fragment("_x0041__x005F_x0042__xD83D__xDE00__x0001__xFFFE_ &amp; &#x3C; &#62;");
    let value = read(&source).expect("valid SpreadsheetML XString escapes");
    assert_eq!(value.value(), "A_x0042_😀\u{1}\u{fffe} & < >");
    assert_eq!(write(&value).unwrap(), source);

    let detached_value = "literal _x0041_ _x005F_ & < > \" ' \t\n\r\0\u{1}\u{fffe}\u{ffff}😀";
    let detached = Element::new(detached_value).expect("ST_Xstring supports escaped controls");
    let output = write(&detached).unwrap();
    let output_text = std::str::from_utf8(&output).unwrap();
    assert!(output_text.contains("_x005F_x0041_"));
    assert!(output_text.contains("_x0000_"));
    assert!(output_text.contains("_x0001_"));
    assert!(output_text.contains("_xFFFE_"));
    assert!(output_text.contains("_xFFFF_"));
    assert!(!output.iter().any(|byte| matches!(byte, 0 | 1 | 13)));
    assert!(output_text.contains("\t\n"));
    assert!(output_text.contains("_x000D_"));
    assert!(!output_text.contains("_x0009_"));
    assert!(!output_text.contains("_x000A_"));
    assert_eq!(read(&output).unwrap().value(), detached_value);

    for encoded in ["_xD800_", "_xDC00_", "_xD800__x0041_", "_xDE00__xD83D_"] {
        let source = fragment(encoded);
        assert!(
            read(&source).is_err(),
            "accepted malformed XString {encoded}"
        );
    }
}

#[test]
fn xml_references_and_controls_fail_at_the_relevant_boundary() {
    assert_invalid(
        &fragment("&unknown;"),
        "formatcode2 contains an unknown general entity",
    );
    assert_xml_error(&fragment("&#0;"));
    assert_invalid(
        &fragment("&#xFFFE;"),
        "formatcode2 reference contains an invalid XML character",
    );
    assert_rejected(&fragment("<"));
    assert_invalid(
        &fragment("]]>"),
        "formatcode2 text contains the forbidden ]]> delimiter",
    );

    let detached_control = Element::new("a\u{1}b").unwrap();
    let output = write(&detached_control).unwrap();
    assert!(std::str::from_utf8(&output).unwrap().contains("a_x0001_b"));
    assert_eq!(read(&output).unwrap().value(), "a\u{1}b");
}

#[test]
fn reserved_namespace_bindings_and_unbound_qnames_are_rejected_exactly() {
    let cases = [
        (
            format!(
                r#"<f:formatcode2 xmlns:f="{NAMESPACE}" xmlns:p="http://www.w3.org/2000/xmlns/">x</f:formatcode2>"#
            ),
            "formatcode2 declaration binds the reserved XMLNS namespace",
        ),
        (
            format!(
                r#"<f:formatcode2 xmlns:f="{NAMESPACE}" xmlns:p="http://www.w3.org/XML/1998/namespace">x</f:formatcode2>"#
            ),
            "formatcode2 declaration binds the XML namespace to a non-xml prefix",
        ),
        (
            format!(
                r#"<f:formatcode2 xmlns:f="{NAMESPACE}" xmlns:xml="urn:wrong">x</f:formatcode2>"#
            ),
            "formatcode2 xml prefix has the wrong namespace",
        ),
        (
            format!(
                r#"<f:formatcode2 xmlns:f="{NAMESPACE}" xmlns:xmlns="urn:wrong">x</f:formatcode2>"#
            ),
            "formatcode2 namespace declaration prefix is invalid",
        ),
        (
            format!(r#"<f:formatcode2 xmlns:f="{NAMESPACE}" xmlns:p="">x</f:formatcode2>"#),
            "formatcode2 prefixed namespace binding is empty",
        ),
    ];
    assert_invalid(cases[0].0.as_bytes(), cases[0].1);
    assert_xml_error(cases[1].0.as_bytes());
    assert_xml_error(cases[2].0.as_bytes());
    for (source, expected) in cases.into_iter().skip(3) {
        assert_invalid(source.as_bytes(), expected);
    }

    assert_invalid(
        br#"<f:formatcode2>x</f:formatcode2>"#,
        "formatcode2 element has an undeclared prefix 'f',",
    );
    let source = format!(r#"<f:formatcode2 xmlns:f="{NAMESPACE}" p:x="1">x</f:formatcode2>"#);
    assert_invalid(
        source.as_bytes(),
        "formatcode2 attribute has an undeclared prefix",
    );
}

#[test]
fn root_grammar_rejects_attributes_children_extra_roots_and_outer_data() {
    let source = format!(r#"<f:formatcode2 xmlns:f="{NAMESPACE}" data="1">x</f:formatcode2>"#);
    assert_invalid(
        source.as_bytes(),
        "formatcode2 root has unexpected attributes",
    );
    let source = format!(r#"<f:formatcode2 xmlns:f="{NAMESPACE}"><f:other/></f:formatcode2>"#);
    assert_invalid(
        source.as_bytes(),
        "formatcode2 value contains a child element",
    );

    let root = fragment("x");
    let multiple = [root.as_slice(), root.as_slice()].concat();
    assert_invalid(
        &multiple,
        "formatcode2 fragment has more than one root element",
    );
    assert_invalid(
        b"prefix <f:formatcode2 xmlns:f=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">x</f:formatcode2>",
        "formatcode2 has non-whitespace text outside its root",
    );
    assert_invalid(
        b"<f:formatcode2 xmlns:f=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">x</f:formatcode2>suffix",
        "formatcode2 has non-whitespace text outside its root",
    );
    assert_invalid(
        b"<![CDATA[x]]><f:formatcode2 xmlns:f=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">x</f:formatcode2>",
        "formatcode2 contains data outside its root",
    );
    assert_invalid(
        b"<!DOCTYPE formatcode2><f:formatcode2 xmlns:f=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">x</f:formatcode2>",
        "formatcode2 cannot contain a document type",
    );
}

#[test]
fn split_text_comments_cdata_and_refs_are_exactly_retained_on_noop() {
    let source = br#"<?xml version="1.0" encoding="UTF-8"?><!--before--><f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">left<!--between-->right<![CDATA[]]]]><![CDATA[>]]>&amp;<?inside?>tail</f:formatcode2><!--after-->"#;
    let value = read(source).expect("split simple-content events are valid");
    assert_eq!(value.value(), "leftright]]>&tail");
    assert_eq!(write(&value).unwrap(), source);

    let mut cdata = read(
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart"><![CDATA[General]]><!--keep--></f:formatcode2>"#,
    )
    .unwrap();
    cdata.set_value("0.00").unwrap();
    let edited = write(&cdata).unwrap();
    let edited_text = std::str::from_utf8(&edited).unwrap();
    assert!(edited_text.contains("<![CDATA[0.00]]>"));
    assert!(edited_text.contains("<!--keep-->"));
    assert_eq!(read(&edited).unwrap().value(), "0.00");
}

#[test]
fn scalar_edit_keeps_opaque_markup_and_escapes_a_forbidden_cdata_delimiter() {
    let source = br#"<!--before--><f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">old<!--opaque--><![CDATA[tail]]><?opaque?></f:formatcode2><!--after-->"#;
    let mut value = read(source).unwrap();
    assert_eq!(value.value(), "oldtail");
    value.set_value("new]]>value").unwrap();
    let edited = write(&value).unwrap();
    let text = std::str::from_utf8(&edited).unwrap();
    assert!(text.contains("<!--before-->"));
    assert!(text.contains("<!--opaque-->"));
    assert!(text.contains("<?opaque?>"));
    assert!(text.contains("<!--after-->"));
    assert!(!text.contains("new]]>value"));
    assert_eq!(read(&edited).unwrap().value(), "new]]>value");

    let split = br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart"><![CDATA[]]]]><![CDATA[>]]></f:formatcode2>"#;
    let split_value = read(split).unwrap();
    assert_eq!(split_value.value(), "]]>");
    assert_eq!(write(&split_value).unwrap(), split);
}

#[test]
fn detached_and_shared_writes_are_deterministic_and_sink_safe() {
    let detached = Element::new("General & _x0041_\u{1}").unwrap();
    let first = write(&detached).unwrap();
    let second = write(&detached).unwrap();
    assert_eq!(first, second);
    let reopened = read(&first).unwrap();
    assert_eq!(reopened.value(), detached.value());
    assert!(reopened.source().is_some());

    let source = std::sync::Arc::<[u8]>::from(
        br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart">General</f:formatcode2>"#.as_slice(),
    );
    let parsed = read_shared(source.clone()).unwrap();
    assert_eq!(parsed.source(), Some(source.as_ref()));
    let mut sink = Vec::new();
    write_to(&mut sink, &parsed).unwrap();
    assert_eq!(sink, source.as_ref());

    let mut edited = parsed.clone();
    edited.set_value("0").unwrap();
    sink.clear();
    write_to(&mut sink, &edited).unwrap();
    assert_eq!(read(&sink).unwrap().value(), "0");
}

#[test]
fn qualified_attribute_owner_preserves_host_start_tag_and_detaches_cleanly() {
    let source = format!(
        r#"<c:series xmlns:c="{NAMESPACE}" id="7" c:formatcode2="_x0041__x0001_" data-extra="opaque"/>"#
    );
    let mut attribute = read_attribute(source.as_bytes()).expect("qualified formatcode2 attribute");
    assert_eq!(attribute.value(), "A\u{1}");
    assert_eq!(attribute.source(), Some(source.as_bytes()));
    assert_eq!(write_attribute(&attribute).unwrap(), source.as_bytes());

    attribute.set_value("0.00").unwrap();
    let edited = write_attribute(&attribute).unwrap();
    let edited_text = std::str::from_utf8(&edited).unwrap();
    assert!(edited_text.contains(r#"id="7""#));
    assert!(edited_text.contains(r#"data-extra="opaque""#));
    assert!(edited_text.contains(r#"c:formatcode2="0.00""#));
    assert_eq!(read_attribute(&edited).unwrap().value(), "0.00",);

    let detached = formatcode2::Value::new("literal _x0041_\u{1}").unwrap();
    let lexical = formatcode2::write_attribute_value(&detached).unwrap();
    assert_eq!(
        std::str::from_utf8(&lexical).unwrap(),
        "literal _x005F_x0041__x0001_",
    );
    let whitespace = formatcode2::Value::new("a\t\n\r").unwrap();
    assert_eq!(
        std::str::from_utf8(&formatcode2::write_attribute_value(&whitespace).unwrap()).unwrap(),
        "a&#x9;&#xA;&#xD;",
    );
    let whitespace_source =
        format!(r#"<c:series xmlns:c="{NAMESPACE}" c:formatcode2="a&#x9;&#xA;_x000D_"/>"#);
    let parsed_whitespace = read_attribute(whitespace_source.as_bytes()).unwrap();
    assert_eq!(parsed_whitespace.value(), "a\t\n\r");
    assert_eq!(
        write_attribute(&parsed_whitespace).unwrap(),
        whitespace_source.as_bytes()
    );
    assert_eq!(
        std::str::from_utf8(&formatcode2::write_attribute_value(&detached).unwrap()).unwrap(),
        "literal _x005F_x0041__x0001_",
    );
}

fn assert_inherited_attribute_roundtrip(source: &[u8], bindings: &[(&str, &str)]) {
    let mut attribute =
        read_attribute_with_bindings(source, bindings).expect("inherited namespace bindings");
    assert_eq!(attribute.value(), "General");
    assert_eq!(attribute.source(), Some(source));
    assert_eq!(write_attribute(&attribute).unwrap(), source);

    attribute.set_value("0.00").unwrap();
    let edited = write_attribute(&attribute).unwrap();
    let reopened =
        read_attribute_with_bindings(&edited, bindings).expect("edited inherited attribute");
    assert_eq!(reopened.value(), "0.00");
    assert_eq!(reopened.source(), Some(edited.as_slice()));
}

#[test]
fn qualified_attribute_noop_shared_sink_and_grammar_failures_are_exact() {
    let source = std::sync::Arc::<[u8]>::from(
        format!(r#"<c:series xmlns:c="{NAMESPACE}" c:formatcode2="General" unknown="1"/>"#)
            .into_bytes(),
    );
    let attribute = read_attribute_shared(source.clone()).unwrap();
    assert_eq!(attribute.source(), Some(source.as_ref()));
    let mut sink = Vec::new();
    write_attribute_to(&mut sink, &attribute).unwrap();
    assert_eq!(sink, source.as_ref());

    let missing = format!(r#"<c:series xmlns:c="{NAMESPACE}" formatcode2="General"/>"#);
    match read_attribute(missing.as_bytes()) {
        Err(Error::Invalid(message)) => assert_eq!(message, "formatcode2 attribute is absent"),
        other => panic!("expected absent attribute refusal, got {other:?}"),
    }
    let foreign = br#"<c:series xmlns:c="urn:other" c:formatcode2="General"/>"#;
    match read_attribute(foreign) {
        Err(Error::Invalid(message)) => assert_eq!(message, "formatcode2 attribute is absent"),
        other => panic!("expected foreign attribute refusal, got {other:?}"),
    }
    let undeclared = br#"<series c:formatcode2="General"/>"#;
    match read_attribute(undeclared) {
        Err(Error::Invalid(message)) => {
            assert_eq!(message, "formatcode2 attribute has an undeclared prefix")
        },
        other => panic!("expected undeclared attribute refusal, got {other:?}"),
    }
}

#[test]
fn inherited_attribute_bindings_cover_prefix_default_and_local_shadowing() {
    let inherited_prefix = br#"<c:series f:formatcode2="General"/>"#;
    assert_inherited_attribute_roundtrip(inherited_prefix, &[("c", "urn:host"), ("f", NAMESPACE)]);

    let inherited_default = br#"<series f:formatcode2="General"/>"#;
    assert_inherited_attribute_roundtrip(inherited_default, &[("", "urn:host"), ("f", NAMESPACE)]);

    let opaque = br#"<c:series f:formatcode2="General" c:opaque="host" q:future="keep"/>"#;
    assert_inherited_attribute_roundtrip(
        opaque,
        &[("c", "urn:host"), ("f", NAMESPACE), ("q", "urn:opaque")],
    );

    let local_shadow = format!(
        r#"<series xmlns="urn:local-host" xmlns:f="{}" f:formatcode2="General"/>"#,
        NAMESPACE
    );
    assert_inherited_attribute_roundtrip(
        local_shadow.as_bytes(),
        &[("", "urn:inherited-host"), ("f", "urn:wrong")],
    );

    let shadowed_target = br#"<series xmlns:f="urn:wrong" f:formatcode2="General"/>"#;
    match read_attribute_with_bindings(shadowed_target, &[("f", NAMESPACE)]) {
        Err(Error::Invalid(message)) => assert_eq!(message, "formatcode2 attribute is absent"),
        other => panic!("expected local namespace shadow refusal, got {other:?}"),
    }

    let shared_source = std::sync::Arc::<[u8]>::from(inherited_prefix.to_vec().into_boxed_slice());
    let shared = read_attribute_shared_with_bindings(
        shared_source.clone(),
        &[("c", "urn:host"), ("f", NAMESPACE)],
    )
    .expect("shared inherited namespace bindings");
    assert_eq!(shared.source(), Some(shared_source.as_ref()));
    assert_eq!(write_attribute(&shared).unwrap(), shared_source.as_ref());

    let duplicate = [("f", NAMESPACE), ("f", NAMESPACE)];
    match read_attribute_with_bindings(inherited_prefix, &duplicate) {
        Err(Error::Invalid(_)) => {},
        other => panic!("expected duplicate inherited binding refusal, got {other:?}"),
    }
    let conflicting = [("f", NAMESPACE), ("f", "urn:wrong")];
    match read_attribute_with_bindings(inherited_prefix, &conflicting) {
        Err(Error::Invalid(message)) => assert_eq!(
            message,
            "formatcode2 inherited namespace prefix has conflicting bindings"
        ),
        other => panic!("expected conflicting inherited binding refusal, got {other:?}"),
    }
}

#[test]
fn qualified_attribute_namespace_and_output_limits_are_real() {
    let long_prefix = "p".repeat(MAX_NAMESPACE_BYTES + 1);
    let source = format!(
        r#"<c:series xmlns:c="{NAMESPACE}" xmlns:{long_prefix}="urn:other" {long_prefix}:formatcode2="x"/>"#
    );
    assert_attribute_limit(
        &source.into_bytes(),
        "formatcode2 namespace bytes",
        MAX_NAMESPACE_BYTES,
    );

    let padding = " ".repeat(700_000);
    let source = format!(r#"<c:series xmlns:c="{NAMESPACE}"{padding} c:formatcode2="a"/>"#);
    let mut attribute = read_attribute(source.as_bytes()).unwrap();
    attribute
        .set_value("\u{1}".repeat(MAX_VALUE_BYTES))
        .unwrap();
    match write_attribute(&attribute) {
        Err(Error::Limit { resource, limit }) => {
            assert_eq!(resource, "formatcode2 attribute XML bytes");
            assert_eq!(limit, MAX_XML_BYTES);
        },
        other => panic!("expected qualified attribute output limit, got {other:?}"),
    }
}

#[test]
fn namespace_prefix_uri_and_xml_input_limits_are_real_and_exact() {
    let long_prefix = "p".repeat(MAX_NAMESPACE_BYTES + 1);
    let source = format!(
        r#"<{long_prefix}:formatcode2 xmlns:{long_prefix}="{NAMESPACE}">x</{long_prefix}:formatcode2>"#
    );
    assert_limit(
        source.as_bytes(),
        "formatcode2 namespace bytes",
        MAX_NAMESPACE_BYTES,
    );

    let long_uri = "u".repeat(MAX_NAMESPACE_BYTES + 1);
    let source = format!(r#"<f:formatcode2 xmlns:f="{long_uri}">x</f:formatcode2>"#);
    assert_limit(
        source.as_bytes(),
        "formatcode2 namespace bytes",
        MAX_NAMESPACE_BYTES,
    );

    let over_input = format!(
        r#"<f:formatcode2 xmlns:f="{NAMESPACE}"><!--{}-->x</f:formatcode2>"#,
        "x".repeat(MAX_XML_BYTES),
    );
    assert_limit(
        over_input.as_bytes(),
        "formatcode2 XML bytes",
        MAX_XML_BYTES,
    );

    let lexical = format!(
        r#"<f:formatcode2 xmlns:f="{NAMESPACE}">{}</f:formatcode2>"#,
        "a".repeat(MAX_LEXICAL_VALUE_BYTES + 1),
    );
    assert_limit(
        lexical.as_bytes(),
        "formatcode2 lexical value bytes",
        MAX_LEXICAL_VALUE_BYTES,
    );

    let decoded_value = format!(
        r#"<f:formatcode2 xmlns:f="{NAMESPACE}">{}</f:formatcode2>"#,
        "a".repeat(MAX_VALUE_BYTES + 1),
    );
    assert_limit(
        decoded_value.as_bytes(),
        "formatcode2 value bytes",
        MAX_VALUE_BYTES,
    );

    let mut attributes = format!(r#"<f:formatcode2 xmlns:f="{NAMESPACE}""#);
    for index in 0..MAX_ATTRIBUTES {
        attributes.push_str(&format!(r#" xmlns:p{index}="urn:p{index}""#));
    }
    attributes.push_str(">x</f:formatcode2>");
    assert_limit(
        &attributes.into_bytes(),
        "formatcode2 attributes",
        MAX_ATTRIBUTES,
    );

    let mut local_attribute_declarations = format!(r#"<series xmlns:f="{NAMESPACE}""#);
    for index in 0..MAX_ATTRIBUTES {
        local_attribute_declarations.push_str(&format!(r#" xmlns:p{index}="urn:p{index}""#));
    }
    local_attribute_declarations.push_str(r#" f:formatcode2="x"/>"#);
    assert_attribute_limit(
        local_attribute_declarations.as_bytes(),
        "formatcode2 attributes",
        MAX_ATTRIBUTES,
    );

    let mut owned_bindings = vec![("f".to_owned(), NAMESPACE.to_owned())];
    for index in 0..MAX_ATTRIBUTES {
        owned_bindings.push((format!("p{index}"), format!("urn:p{index}")));
    }
    let bindings = owned_bindings
        .iter()
        .map(|(prefix, namespace)| (prefix.as_str(), namespace.as_str()))
        .collect::<Vec<_>>();
    assert_attribute_binding_limit(
        br#"<series f:formatcode2="x"/>"#,
        &bindings,
        "formatcode2 inherited namespace bindings",
        MAX_ATTRIBUTES,
    );
}

#[test]
fn output_limit_charges_actual_xstring_expansion_before_allocating() {
    let filler = "x".repeat(700_000);
    let source =
        format!(r#"<f:formatcode2 xmlns:f="{NAMESPACE}">a<!--{filler}--></f:formatcode2>"#);
    assert!(source.len() < MAX_XML_BYTES);
    let mut value = read(source.as_bytes()).expect("source is under the XML bound");
    value
        .set_value("\u{1}".repeat(MAX_VALUE_BYTES))
        .expect("semantic value is under the value bound");
    assert_write_limit(&value, "formatcode2 XML bytes", MAX_XML_BYTES);
}

#[test]
fn public_limit_constants_match_the_adversarial_contract() {
    assert_eq!(MAX_ATTRIBUTES, 64);
    assert_eq!(MAX_NODES, 16);
    assert_eq!(MAX_DEPTH, 8);
    assert_eq!(MAX_NAMESPACE_BYTES, 4 * 1024);
    assert_eq!(MAX_VALUE_BYTES, 64 * 1024);
    assert_eq!(MAX_XML_BYTES, 1 << 20);
}
