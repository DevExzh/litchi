//! Adversarial tests for the namespace-bound PowerPoint 2014 ink-action
//! profile.  The legacy `actions::read` codec intentionally has a broader,
//! opaque compatibility contract; these tests exercise only `read_profile`.

use litchi_drawingml::{
    Error,
    ink::{
        ACTION_NAMESPACE, INKML_NAMESPACE, MAX_ATTRIBUTE_VALUE_BYTES, MAX_DEPTH, MAX_NODES,
        MAX_SOURCE_BYTES, actions,
    },
};

fn root(body: &str, length_unit: &str, time_unit: &str) -> String {
    format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="{length_unit}" timeUnit="{time_unit}">{body}</iact:actions>"#
    )
}

fn assert_rejected(xml: &[u8]) {
    match actions::read_profile(xml) {
        Err(Error::Invalid(_) | Error::Xml(_)) => {},
        Err(other) => panic!("expected profile grammar rejection, got {other:?}"),
        Ok(profile) => panic!("accepted invalid profile {:?}", profile.source()),
    }
}

fn assert_invalid(xml: &[u8]) {
    match actions::read_profile(xml) {
        Err(Error::Invalid(_)) => {},
        Err(other) => panic!("expected Invalid profile error, got {other:?}"),
        Ok(profile) => panic!("accepted invalid profile {:?}", profile.source()),
    }
}

fn assert_limit(xml: &[u8], resource: &'static str, limit: usize) {
    match actions::read_profile(xml) {
        Err(Error::Limit {
            resource: actual_resource,
            limit: actual_limit,
        }) => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_limit, limit);
        },
        Err(other) => panic!("expected {resource:?} limit {limit}, got {other:?}"),
        Ok(profile) => panic!("accepted over-budget profile {:?}", profile.source()),
    }
}

#[test]
fn profile_preserves_character_reference_whitespace_in_element_only_containers() {
    let body = r#"&#32;<iact:actionGroup type="add" startTime="0">&#x9;<iact:action type="add" startTime="0">&#10;<iact:property name="style"/>&#xD;<iact:actionDataGroup>&#00032;<iact:actionData>&#x000020;</iact:actionData></iact:actionDataGroup></iact:action></iact:actionGroup>"#;
    let source = root(body, "m", "s");
    let profile = actions::read_profile(source.as_bytes()).unwrap();
    assert_eq!(actions::write_profile(&profile).unwrap(), source.as_bytes());
    for reference in ["&#65;", "&#xA0;", "&amp;", "&#0;", "&unknown;"] {
        assert_rejected(root(reference, "m", "s").as_bytes());
    }
    for source in [format!("&#32;{source}"), format!("{source}&#32;")] {
        assert_rejected(source.as_bytes());
    }
}

#[test]
fn profile_empty_property_content_rejects_whitespace_but_preserves_comments() {
    for body in [" ", "&#32;", "<![CDATA[ ]]>", "text"] {
        let action = format!(
            r#"<iact:action type="add" startTime="0"><iact:property name="style">{body}</iact:property></iact:action>"#
        );
        assert_rejected(root(&action, "m", "s").as_bytes());
    }
    let source = root(
        r#"<iact:action type="add" startTime="0"><iact:property name="style"><!--keep--></iact:property></iact:action>"#,
        "m",
        "s",
    );
    let profile = actions::read_profile(source.as_bytes()).unwrap();
    assert_eq!(actions::write_profile(&profile).unwrap(), source.as_bytes());
}

#[test]
fn profile_rejects_raw_xml_delimiters_and_all_cdata_outside_root() {
    let property = r#"<iact:action type="add" startTime="0"><iact:property name="style" value="<"/></iact:action>"#;
    assert_rejected(root(property, "m", "s").as_bytes());
    let trace = format!(
        r#"<iact:action type="add" startTime="0"><iact:actionData><i:trace xmlns:i="{INKML_NAMESPACE}">bad]]&gt;</i:trace></iact:actionData></iact:action>"#
    );
    let valid = root(&trace, "m", "s");
    let profile = actions::read_profile(valid.as_bytes()).unwrap();
    assert_eq!(actions::write_profile(&profile).unwrap(), valid.as_bytes());
    assert_rejected(valid.replace("]]&gt;", "]]>").as_bytes());
    let valid = root(&trace.replace("bad]]&gt;", "<![CDATA[<?x]]>"), "m", "s");
    let profile = actions::read_profile(valid.as_bytes()).unwrap();
    assert_eq!(actions::write_profile(&profile).unwrap(), valid.as_bytes());
    for cdata in ["<![CDATA[]]>", "<![CDATA[ ]]>"] {
        assert_rejected(format!("{cdata}{valid}").as_bytes());
        assert_rejected(format!("{valid}{cdata}").as_bytes());
    }
    let valid = root(&property.replace("value=\"<\"", "value=\"&lt;\""), "m", "s");
    assert!(actions::read_profile(valid.as_bytes()).is_ok());
}

#[test]
fn profile_rejects_reserved_namespace_aliases_inside_opaque_payloads() {
    for namespace in [
        "http://www.w3.org/XML/1998/namespace",
        "http://www.w3.org/2000/xmlns/",
    ] {
        for namespace in [namespace.to_owned(), namespace.replace('/', "&#x2F;")] {
            for declaration in [
                format!(r#"xmlns="{namespace}""#),
                format!(r#"xmlns:p="{namespace}" p:future="z""#),
            ] {
                let body = format!(
                    r#"<i:definitions xmlns:i="{INKML_NAMESPACE}"><i:trace {declaration}/></i:definitions>"#
                );
                assert_rejected(root(&body, "m", "s").as_bytes());
            }
        }
    }
    let body = format!(
        r#"<i:definitions xmlns:i="{INKML_NAMESPACE}"><future xmlns="" xmlns:p="urn:ordinary" p:value="z"/></i:definitions>"#
    );
    let source = root(&body, "m", "s");
    let profile = actions::read_profile(source.as_bytes()).unwrap();
    assert_eq!(actions::write_profile(&profile).unwrap(), source.as_bytes());
}

#[test]
fn profile_replays_exact_source_and_exposes_typed_children_in_schema_order() {
    let source = br###"<?xml version="1.0" encoding="UTF-8"?><!--before--><iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:inkml="http://www.w3.org/2003/InkML" xml:id="root" lengthUnit="cm" timeUnit="ms"><inkml:definitions><inkml:future><inkml:payload>opaque</inkml:payload></inkml:future></inkml:definitions><iact:action xml:id="a0" type="add" startTime="+0.25"><iact:property name="dataType"/><iact:property name="style" value="instant"/><iact:property name="vendor" value="opaque-value"/><iact:actionDataGroup xml:id="g0" name="path"><iact:actionData xml:id="d0" name="target" ref="#stroke"><iact:transform/><inkml:trace><inkml:future><inkml:payload>opaque</inkml:payload></inkml:future></inkml:trace><inkml:traceView/></iact:actionData></iact:actionDataGroup></iact:action><iact:actionGroup xml:id="ag0" type="transform" startTime="1.0"><iact:action type="remove" startTime="2"><iact:actionData/></iact:action></iact:actionGroup></iact:actions><!--after-->"###;

    let profile = actions::read_profile(source).expect("valid strict action profile");
    assert_eq!(profile.source(), source);
    assert_eq!(actions::write_profile(&profile).unwrap(), source);
    assert_eq!(profile.xml(profile.definitions_span().unwrap()), br#"<inkml:definitions><inkml:future><inkml:payload>opaque</inkml:payload></inkml:future></inkml:definitions>"#);
    assert_eq!(profile.xml_id(), Some("root"));
    assert_eq!(profile.length_unit(), actions::LengthUnit::Centimeter);
    assert_eq!(profile.time_unit(), actions::TimeUnit::Millisecond);
    assert_eq!(profile.children().len(), 2);

    let actions::RootChild::Action(action) = &profile.children()[0] else {
        panic!("first root child must be an action");
    };
    assert_eq!(action.xml_id(), Some("a0"));
    assert_eq!(action.start_time(), "+0.25");
    assert_eq!(action.action_type().as_str(), "add");
    assert_eq!(action.properties().len(), 3);
    assert_eq!(action.properties()[0].name(), "dataType");
    assert_eq!(action.properties()[0].value(), "ink");
    assert_eq!(action.properties()[1].name(), "style");
    assert_eq!(action.properties()[1].value(), "instant");
    assert_eq!(action.properties()[2].name(), "vendor");
    assert_eq!(action.properties()[2].value(), "opaque-value");
    assert_eq!(action.children().len(), 4);
    let actions::ActionChild::DataGroup(group) = &action.children()[3] else {
        panic!("fourth action child must be a data group");
    };
    assert_eq!(group.xml_id(), Some("g0"));
    assert_eq!(group.name(), "path");
    assert_eq!(group.data().len(), 1);
    let data = &group.data()[0];
    assert_eq!(data.xml_id(), Some("d0"));
    assert_eq!(data.name(), "target");
    assert_eq!(data.reference(), Some("#stroke"));
    assert_eq!(data.children().len(), 3);
    assert!(matches!(
        data.children()[0],
        actions::DataChild::Transform(_)
    ));
    assert!(matches!(data.children()[1], actions::DataChild::Trace(_)));
    assert!(matches!(
        data.children()[2],
        actions::DataChild::TraceView(_)
    ));
    assert!(
        std::str::from_utf8(data.xml(&profile))
            .unwrap()
            .contains("<inkml:future>")
    );

    let actions::RootChild::ActionGroup(group) = &profile.children()[1] else {
        panic!("second root child must be an action group");
    };
    assert_eq!(group.xml_id(), Some("ag0"));
    assert_eq!(group.action_type().as_str(), "transform");
    assert_eq!(group.start_time(), "1.0");
    assert_eq!(group.actions().len(), 1);
}

#[test]
fn all_schema_length_and_time_units_round_trip_through_public_enums() {
    for (lexical, expected) in [
        ("m", actions::LengthUnit::Meter),
        ("cm", actions::LengthUnit::Centimeter),
        ("mm", actions::LengthUnit::Millimeter),
        ("in", actions::LengthUnit::Inch),
        ("pt", actions::LengthUnit::Point),
        ("pc", actions::LengthUnit::Pica),
        ("em", actions::LengthUnit::Em),
        ("ex", actions::LengthUnit::Ex),
    ] {
        let profile = actions::read_profile(root("", lexical, "s").as_bytes())
            .expect("standard InkML length unit");
        assert_eq!(profile.length_unit(), expected);
        assert_eq!(profile.length_unit().as_str(), lexical);
    }
    for (lexical, expected) in [
        ("s", actions::TimeUnit::Second),
        ("ms", actions::TimeUnit::Millisecond),
    ] {
        let profile = actions::read_profile(root("", "m", lexical).as_bytes())
            .expect("standard InkML time unit");
        assert_eq!(profile.time_unit(), expected);
        assert_eq!(profile.time_unit().as_str(), lexical);
    }
}

#[test]
fn nonstandard_units_and_missing_required_units_are_rejected() {
    for unit in ["px", "inch", "second", "MS", "m ", ""] {
        let source = root("", unit, "s");
        assert_invalid(source.as_bytes());
    }
    for unit in ["min", "ms ", "", "seconds"] {
        let source = root("", "m", unit);
        assert_invalid(source.as_bytes());
    }
    assert_invalid(
        br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" timeUnit="s"/>"#,
    );
    assert_invalid(
        br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="m"/>"#,
    );
}

#[test]
fn xsd_decimal_whitespace_is_collapsed_before_decimal_validation() {
    // xsd:decimal has the fixed `collapse` whitespace facet.  XML attribute
    // references are retained by the XML layer, while the schema value space
    // treats the surrounding whitespace as insignificant.
    let source = root(
        r#"<iact:action type="add" startTime=" &#x2B;0.25&#x20; "/><iact:actionGroup type="remove" startTime="&#x20;-1.0&#x9;"><iact:action type="add" startTime="0"/></iact:actionGroup>"#,
        "m",
        "s",
    );
    let profile = actions::read_profile(source.as_bytes()).expect("xsd:decimal whitespace");
    let actions::RootChild::Action(action) = &profile.children()[0] else {
        panic!("root action expected");
    };
    assert_eq!(action.start_time().trim(), "+0.25");
    let actions::RootChild::ActionGroup(group) = &profile.children()[1] else {
        panic!("root action group expected");
    };
    assert_eq!(group.start_time().trim(), "-1.0");
    assert_eq!(actions::write_profile(&profile).unwrap(), source.as_bytes());

    for value in ["", ".", "+", "-", "1e2", "+ 1", "1 2"] {
        let source = root(
            &format!(r#"<iact:action type="add" startTime="{value}"/>"#),
            "m",
            "s",
        );
        assert_invalid(source.as_bytes());
    }
}

#[test]
fn empty_action_type_is_allowed_by_the_unrestricted_user_union_member() {
    // ST_ActionType is a union of the reserved enumeration and
    // ST_ActionTypeUser, whose xsd:string base admits the empty lexical value.
    // The profile remains inert and does not infer execution semantics from it.
    let source = root(
        r#"<iact:action type="" startTime="0"/><iact:action type="vendor-action" startTime="1"/>"#,
        "m",
        "s",
    );
    let profile = actions::read_profile(source.as_bytes()).expect("empty custom action type");
    let actions::RootChild::Action(empty) = &profile.children()[0] else {
        panic!("first root action expected");
    };
    assert_eq!(empty.action_type().as_str(), "");
    let actions::RootChild::Action(custom) = &profile.children()[1] else {
        panic!("second root action expected");
    };
    assert_eq!(custom.action_type().as_str(), "vendor-action");
    assert_eq!(actions::write_profile(&profile).unwrap(), source.as_bytes());
}

#[test]
fn property_names_values_and_action_data_names_are_lexical_and_inert() {
    let source = root(
        r#"<iact:action type="transform" startTime="0"><iact:property name="dataType" value="pointEraser"/><iact:property name="style" value="instant"/><iact:property name="future-property" value="future-value"/><iact:actionData name="not-a-reserved-name"/><iact:actionDataGroup name="future-group"><iact:actionData name="future-data"/></iact:actionDataGroup></iact:action>"#,
        "m",
        "s",
    );
    let profile = actions::read_profile(source.as_bytes()).expect("lexical custom names");
    let actions::RootChild::Action(action) = &profile.children()[0] else {
        panic!("root action expected");
    };
    assert_eq!(action.action_type().as_str(), "transform");
    assert_eq!(action.properties()[0].value(), "pointEraser");
    assert_eq!(action.properties()[1].value(), "instant");
    assert_eq!(action.properties()[2].name(), "future-property");
    assert_eq!(action.properties()[2].value(), "future-value");
    let actions::ActionChild::Data(data) = &action.children()[3] else {
        panic!("first data child expected");
    };
    assert_eq!(data.name(), "not-a-reserved-name");
    let actions::ActionChild::DataGroup(group) = &action.children()[4] else {
        panic!("second data child expected");
    };
    assert_eq!(group.name(), "future-group");
    assert_eq!(group.data()[0].name(), "future-data");
}

#[test]
fn recognized_child_order_and_transform_cardinality_are_enforced() {
    let prefix = format!(r#"xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s""#);
    let invalid = [
        format!(
            r#"<iact:actions {prefix}><iact:action type="add" startTime="0"><iact:actionData/><iact:property name="late"/></iact:action></iact:actions>"#
        ),
        format!(
            r#"<iact:actions {prefix}><iact:action type="add" startTime="0"><iact:actionData><iact:transform/><iact:transform/></iact:actionData></iact:action></iact:actions>"#
        ),
        format!(
            r#"<iact:actions xmlns:inkml="{INKML_NAMESPACE}" {prefix}><iact:action type="add" startTime="0"><iact:actionData><inkml:trace/><iact:transform/></iact:actionData></iact:action></iact:actions>"#
        ),
        format!(r#"<iact:actions {prefix}><iact:actionData/></iact:actions>"#),
        format!(
            r#"<iact:actions {prefix}><iact:actionGroup type="add" startTime="0"/></iact:actions>"#
        ),
        format!(
            r#"<iact:actions {prefix}><iact:action type="add" startTime="0"><iact:actionDataGroup/></iact:action></iact:actions>"#
        ),
    ];
    for source in invalid {
        assert_rejected(source.as_bytes());
    }
}

#[test]
fn reserved_action_types_do_not_execute_semantic_name_requirements() {
    let source = root(
        r#"<iact:action type="add" startTime="0"><iact:actionData name="target"/></iact:action><iact:action type="remove" startTime="1"><iact:actionData name="path"/></iact:action><iact:action type="transform" startTime="2"/>"#,
        "m",
        "s",
    );
    let profile = actions::read_profile(source.as_bytes())
        .expect("profile must remain inert about reserved action semantics");
    assert_eq!(profile.children().len(), 3);
    assert_eq!(profile.actions().count(), 3);
}

#[test]
fn namespace_aliases_are_resolved_but_wrong_namespace_is_rejected() {
    let source = format!(
        r#"<alias:actions xmlns:alias="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" lengthUnit="m" timeUnit="s"><alias:action type="add" startTime="0"><alias:actionData><i:trace/></alias:actionData></alias:action></alias:actions>"#
    );
    let profile = actions::read_profile(source.as_bytes()).expect("prefix aliases are lexical");
    assert_eq!(profile.children().len(), 1);
    assert_eq!(actions::write_profile(&profile).unwrap(), source.as_bytes());

    let wrong_root = r#"<alias:actions xmlns:alias="urn:wrong" lengthUnit="m" timeUnit="s"/>"#;
    assert_invalid(wrong_root.as_bytes());
    let wrong_child = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:other="urn:wrong" lengthUnit="m" timeUnit="s"><other:action type="add" startTime="0"/></iact:actions>"#
    );
    assert_invalid(wrong_child.as_bytes());
}

#[test]
fn strict_profile_rejects_unknown_root_attributes_and_foreign_or_unqualified_children() {
    let cases = [
        format!(
            r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s" future="1"/>"#
        ),
        format!(
            r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:f="urn:foreign" lengthUnit="m" timeUnit="s" f:future="1"/>"#
        ),
        format!(
            r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:f="urn:foreign" lengthUnit="m" timeUnit="s"><f:action type="add" startTime="0"/></iact:actions>"#
        ),
        format!(
            r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s"><action type="add" startTime="0"/></iact:actions>"#
        ),
        format!(
            r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s"><future/></iact:actions>"#
        ),
    ];
    for source in cases {
        assert_invalid(source.as_bytes());
    }
}

#[test]
fn encoded_namespace_aliases_resolve_to_the_same_expanded_names() {
    let encoded_action = ACTION_NAMESPACE.replace('/', "&#x2F;");
    let encoded_inkml = INKML_NAMESPACE.replace('/', "&#x2F;");
    let source = format!(
        r#"<a:actions xmlns:a="{encoded_action}" xmlns:i="{encoded_inkml}" lengthUnit="m" timeUnit="s"><a:action type="add" startTime="0"><a:actionData><i:trace/></a:actionData></a:action></a:actions>"#
    );
    let profile = actions::read_profile(source.as_bytes()).expect("numeric URI references resolve");
    assert_eq!(profile.children().len(), 1);
    assert_eq!(actions::write_profile(&profile).unwrap(), source.as_bytes());
}

#[test]
fn duplicate_expanded_attributes_and_duplicate_typed_attributes_are_rejected() {
    let duplicate_expanded = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:a="urn:opaque" xmlns:b="urn:opaque" lengthUnit="m" timeUnit="s" a:future="one" b:future="two"/>"#
    );
    assert_invalid(duplicate_expanded.as_bytes());

    let encoded_action = ACTION_NAMESPACE.replace('/', "&#x2F;");
    let encoded_duplicate = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:a="{ACTION_NAMESPACE}" xmlns:b="{encoded_action}" lengthUnit="m" timeUnit="s" a:future="one" b:future="two"/>"#
    );
    assert_invalid(encoded_duplicate.as_bytes());

    let duplicate_typed = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="m" lengthUnit="cm" timeUnit="s"/>"#
    );
    assert_rejected(duplicate_typed.as_bytes());

    let duplicate_xml_id = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xml:id="one" xml:id="two" lengthUnit="m" timeUnit="s"/>"#
    );
    assert_rejected(duplicate_xml_id.as_bytes());
}

#[test]
fn opaque_inkml_descendants_are_retained_while_known_containers_fail_closed() {
    let source = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" lengthUnit="m" timeUnit="s"><i:definitions><i:future><i:trace>opaque</i:trace></i:future></i:definitions><iact:action type="add" startTime="0"><iact:actionData><i:trace><i:future><i:payload>opaque</i:payload></i:future></i:trace></iact:actionData></iact:action></iact:actions>"#
    );
    let profile = actions::read_profile(source.as_bytes()).expect("opaque InkML descendants");
    assert_eq!(profile.source(), source.as_bytes());
    assert!(
        std::str::from_utf8(profile.xml(profile.definitions_span().unwrap()))
            .unwrap()
            .contains("<i:future>")
    );

    let prefix = format!(r#"xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s""#);
    for source in [
        format!(r#"<iact:actions {prefix}><iact:property name="x"/></iact:actions>"#),
        format!(r#"<iact:actions {prefix}><iact:transform/></iact:actions>"#),
        format!(
            r#"<iact:actions {prefix}><iact:action type="add" startTime="0"><iact:actionData><iact:action/></iact:actionData></iact:action></iact:actions>"#
        ),
        format!(
            r#"<iact:actions xmlns:i="{INKML_NAMESPACE}" {prefix}><iact:action type="add" startTime="0"><iact:actionData><i:future/></iact:actionData></iact:action></iact:actions>"#
        ),
        format!(
            r#"<iact:actions {prefix}><iact:action type="add" startTime="0"><iact:actionDataGroup><iact:actionData><iact:actionData/></iact:actionData></iact:actionDataGroup></iact:action></iact:actions>"#
        ),
        format!(
            r#"<iact:actions xmlns:i="{INKML_NAMESPACE}" {prefix}><iact:action type="add" startTime="0"/><i:definitions/></iact:actions>"#
        ),
    ] {
        assert_rejected(source.as_bytes());
    }
}

#[test]
fn xml_declaration_bom_comments_and_document_boundaries_follow_profile_grammar() {
    let body = root("", "m", "s");
    let source = [
        b"\xEF\xBB\xBF<?xml version=\"1.0\" encoding=\"UTF-8\"?>".as_slice(),
        b"<!--before-->".as_slice(),
        body.as_bytes(),
        b"<!--after-->".as_slice(),
    ]
    .concat();
    let profile = actions::read_profile(&source).expect("well-formed XML profile");
    assert_eq!(actions::write_profile(&profile).unwrap(), source);

    let late_declaration = [
        b"<!--before-->".as_slice(),
        body.as_bytes(),
        b"<?xml version=\"1.0\"?>".as_slice(),
    ]
    .concat();
    assert_rejected(&late_declaration);
    let duplicate_root = [body.as_bytes(), body.as_bytes()].concat();
    assert_rejected(&duplicate_root);
    for source in [
        [b"text".as_slice(), body.as_bytes()].concat(),
        [body.as_bytes(), b"text".as_slice()].concat(),
        [b"<!DOCTYPE iact:actions>".as_slice(), body.as_bytes()].concat(),
        [body.as_bytes(), b"<?after?>".as_slice()].concat(),
        [b"<![CDATA[data]]>".as_slice(), body.as_bytes()].concat(),
    ] {
        assert_rejected(&source);
    }
    assert_rejected(
        br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="m" timeUnit="s"><iact:action type="add" startTime="0">"#,
    );
}

#[test]
fn profile_depth_node_attribute_and_source_limits_are_bounded() {
    let mut exact_depth = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" xmlns:x="urn:opaque" lengthUnit="m" timeUnit="s"><i:definitions>"#
    );
    for _ in 0..MAX_DEPTH - 3 {
        exact_depth.push_str("<x:node>");
    }
    exact_depth.push_str("<x:leaf/>");
    for _ in 0..MAX_DEPTH - 3 {
        exact_depth.push_str("</x:node>");
    }
    exact_depth.push_str("</i:definitions></iact:actions>");
    assert!(actions::read_profile(exact_depth.as_bytes()).is_ok());

    let mut overflow_depth = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" xmlns:x="urn:opaque" lengthUnit="m" timeUnit="s"><i:definitions>"#
    );
    for _ in 0..MAX_DEPTH - 2 {
        overflow_depth.push_str("<x:node>");
    }
    overflow_depth.push_str("<x:leaf/>");
    for _ in 0..MAX_DEPTH - 2 {
        overflow_depth.push_str("</x:node>");
    }
    overflow_depth.push_str("</i:definitions></iact:actions>");
    assert_limit(
        overflow_depth.as_bytes(),
        "ink actions XML depth",
        MAX_DEPTH,
    );

    let mut exact_nodes = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" xmlns:x="urn:opaque" lengthUnit="m" timeUnit="s"><i:definitions>"#
    );
    for _ in 0..MAX_NODES - 2 {
        exact_nodes.push_str("<x:n/>");
    }
    exact_nodes.push_str("</i:definitions></iact:actions>");
    assert!(exact_nodes.len() < MAX_SOURCE_BYTES);
    assert!(actions::read_profile(exact_nodes.as_bytes()).is_ok());

    let mut overflow_nodes = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" xmlns:x="urn:opaque" lengthUnit="m" timeUnit="s"><i:definitions>"#
    );
    for _ in 0..MAX_NODES - 1 {
        overflow_nodes.push_str("<x:n/>");
    }
    overflow_nodes.push_str("</i:definitions></iact:actions>");
    assert_limit(
        overflow_nodes.as_bytes(),
        "ink actions XML nodes",
        MAX_NODES,
    );

    let mut attributes = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" xmlns:x="urn:opaque" lengthUnit="m" timeUnit="s"><i:definitions><x:container"#
    );
    for index in 0..257 {
        attributes.push_str(&format!(r#" x:a{index}="v""#));
    }
    attributes.push_str("/></i:definitions></iact:actions>");
    assert_limit(
        attributes.as_bytes(),
        "ink actions attributes per element",
        256,
    );

    let over_source = vec![b' '; MAX_SOURCE_BYTES + 1];
    assert_limit(&over_source, "ink actions source bytes", MAX_SOURCE_BYTES);
}

#[test]
fn attribute_value_hard_cap_is_reported_before_profile_retention() {
    let value = "x".repeat(MAX_ATTRIBUTE_VALUE_BYTES + 1);
    let source = format!(
        r#"<iact:actions xmlns:iact="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s" xml:id="{value}"/>"#
    );
    assert_limit(
        source.as_bytes(),
        "ink actions attribute value bytes",
        MAX_ATTRIBUTE_VALUE_BYTES,
    );
}
