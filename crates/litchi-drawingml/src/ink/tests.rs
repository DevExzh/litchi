use super::{
    BrushPropertyName, ContextKind, InkEffect, MAX_ATTRIBUTE_VALUE_BYTES, MAX_DEPTH, MAX_NODES,
    MAX_SOURCE_BYTES, SemanticType, SourceSpan, actions, read, read_metadata,
    read_metadata_with_source_spans, read_shared, write,
};
use std::sync::Arc;

const INK: &[u8] = br##"<?xml version="1.0" encoding="UTF-8"?><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:msink="http://schemas.microsoft.com/ink/2010/main" xmlns:a16="http://schemas.microsoft.com/office/drawing/2016/ink"><inkml:definitions><inkml:brush xml:id="br0"><inkml:brushProperty name="width" value="0.06667" units="cm"/><inkml:brushProperty name="inkEffects" value="pencil"/></inkml:brush></inkml:definitions><inkml:traceGroup><inkml:annotationXML><msink:context id="{8646EB18-6E67-4FFA-8739-E20C3C1A0F80}" type="writingRegion" semanticType="comment" rotatedBoundingBox="0,0 10,10"/></inkml:annotationXML><inkml:trace contextRef="#ctx0" brushRef="#br0">1 2 3</inkml:trace></inkml:traceGroup></inkml:ink>"##;

#[test]
fn reads_typed_context_trace_and_extended_brush_property() {
    let document = read(INK).expect("InkML fixture must parse");
    assert_eq!(document.context_count(), 1);
    assert_eq!(document.trace_count(), 1);
    assert_eq!(document.brush_property_count(), 2);
    assert!(matches!(
        document.contexts()[0].kind(),
        ContextKind::WritingRegion
    ));
    assert!(matches!(
        document.contexts()[0].semantic_type(),
        Some(SemanticType::Comment)
    ));
    assert!(matches!(
        document.brush_properties()[1].name(),
        BrushPropertyName::InkEffects
    ));
    assert!(matches!(
        document.brush_properties()[1].ink_effect(),
        Some(InkEffect::Pencil)
    ));
    assert_eq!(document.traces()[0].context_ref(), Some("#ctx0"));
    assert_eq!(document.traces()[0].data(&document), b"1 2 3");
    assert_eq!(
        document.contexts()[0].xml(&document),
        br#"<msink:context id="{8646EB18-6E67-4FFA-8739-E20C3C1A0F80}" type="writingRegion" semanticType="comment" rotatedBoundingBox="0,0 10,10"/>"#
    );
}

#[test]
fn effective_brush_properties_apply_profile_defaults_without_changing_source() {
    let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:brush><i:brushProperty name="width" value="bad" units="bogus"/><i:brushProperty name="height"/><i:brushProperty name="color" value="red" units="cm"/><i:brushProperty name="transparency" value="256"/><i:brushProperty name="tip" value="bad"/><i:brushProperty name="rasterOp" value="bad"/><i:brushProperty name="antiAliased" value="maybe"/><i:brushProperty name="fitToCurve" value="maybe"/><i:brushProperty name="ignorePressure"/><i:brushProperty name="inkEffects" value="future"/><i:brushProperty name="anchorX" value="bad"/><i:brushProperty name="anchorY" value="1" units="cm"/><i:brushProperty name="scaleFactor" value="2"/><i:brushProperty name="future" value="opaque"/></i:brush></i:definitions></i:ink>"##;
    let document = read(source).expect("generic InkML reader retains source lexicals");
    assert_eq!(document.brush_property_count(), 14);
    let effective: Vec<_> = document
        .brush_properties()
        .iter()
        .map(|property| property.effective())
        .collect();
    assert_eq!(effective[0].as_ref().unwrap().value(), ".053");
    assert_eq!(effective[0].as_ref().unwrap().units(), Some("cm"));
    assert!(effective[0].as_ref().unwrap().defaulted());
    assert_eq!(effective[1].as_ref().unwrap().value(), ".001");
    assert_eq!(effective[2].as_ref().unwrap().value(), "#000000");
    assert_eq!(effective[3].as_ref().unwrap().value(), "0");
    assert_eq!(effective[4].as_ref().unwrap().value(), "ellipse");
    assert_eq!(effective[5].as_ref().unwrap().value(), "copyPen");
    assert_eq!(effective[6].as_ref().unwrap().value(), "true");
    assert_eq!(effective[7].as_ref().unwrap().value(), "false");
    assert_eq!(effective[8].as_ref().unwrap().value(), "false");
    assert_eq!(effective[9].as_ref().unwrap().value(), "none");
    assert_eq!(effective[10].as_ref().unwrap().value(), "0");
    assert_eq!(effective[11].as_ref().unwrap().value(), "0");
    assert_eq!(effective[12].as_ref().unwrap().value(), "2");
    assert!(!effective[12].as_ref().unwrap().defaulted());
    assert!(effective[13].is_none());
    assert_eq!(document.brush_properties()[0].value(), "bad");
    assert_eq!(document.brush_properties()[0].units(), Some("bogus"));
    assert_eq!(document.source(), source);
}

#[test]
fn effective_brush_properties_use_normative_units_and_xsd_whitespace() {
    let source = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:brush><i:brushProperty name="width" value="&#x20;1.25&#x9;" units="m"/><i:brushProperty name="height" value="1" units="px"/><i:brushProperty name="transparency" value="&#xA;42&#xD;"/><i:brushProperty name="antiAliased" value="&#x20;&#x31;&#x9;"/></i:brush></i:definitions></i:ink>"##;
    let document = read(source).expect("schema whitespace fixture must parse");
    let effective: Vec<_> = document
        .brush_properties()
        .iter()
        .map(|property| property.effective())
        .collect();

    assert_eq!(effective[0].as_ref().unwrap().value(), "1.25");
    assert_eq!(effective[0].as_ref().unwrap().units(), Some("m"));
    assert!(!effective[0].as_ref().unwrap().defaulted());
    assert_eq!(effective[1].as_ref().unwrap().value(), ".001");
    assert_eq!(effective[1].as_ref().unwrap().units(), Some("cm"));
    assert!(effective[1].as_ref().unwrap().defaulted());
    assert_eq!(effective[2].as_ref().unwrap().value(), "42");
    assert!(!effective[2].as_ref().unwrap().defaulted());
    assert_eq!(effective[3].as_ref().unwrap().value(), "true");
    assert!(!effective[3].as_ref().unwrap().defaulted());
    assert_eq!(document.source(), source);
}

#[test]
fn source_backed_write_is_byte_exact() {
    let document = read(INK).expect("InkML fixture must parse");
    assert_eq!(write(&document).expect("source-backed write"), INK);
    let clone = document.clone();
    assert_eq!(clone, document);
}

#[test]
fn read_shared_retains_the_caller_payload_and_spans() {
    let payload = Arc::new(INK.to_vec());
    let document = read_shared(Arc::clone(&payload)).expect("shared InkML fixture must parse");

    assert!(Arc::ptr_eq(&document.source, &payload));
    assert_eq!(Arc::strong_count(&payload), 2);
    assert_eq!(document.source(), INK);
    assert_eq!(write(&document).expect("shared source-backed write"), INK);
    assert_eq!(document.traces()[0].data(&document), b"1 2 3");

    drop(payload);
    assert_eq!(document.source(), INK);
    assert_eq!(document.contexts()[0].xml(&document), br#"<msink:context id="{8646EB18-6E67-4FFA-8739-E20C3C1A0F80}" type="writingRegion" semanticType="comment" rotatedBoundingBox="0,0 10,10"/>"#);
}

#[test]
fn metadata_matches_shared_document_and_clones_share_storage() {
    let metadata = read_metadata(INK).expect("InkML metadata fixture must parse");
    let payload = Arc::new(INK.to_vec());
    let document = read_shared(Arc::clone(&payload)).expect("shared InkML fixture must parse");

    assert_eq!(&metadata, document.metadata());
    assert_eq!(metadata.context_count(), 1);
    assert_eq!(metadata.trace_count(), 1);
    assert_eq!(metadata.brush_property_count(), 2);

    let clone = metadata.clone();
    assert_eq!(metadata.contexts().as_ptr(), clone.contexts().as_ptr());
    assert_eq!(metadata.traces().as_ptr(), clone.traces().as_ptr());
    assert_eq!(
        metadata.brush_properties().as_ptr(),
        clone.brush_properties().as_ptr()
    );
}

#[test]
fn filtered_metadata_requires_exact_source_spans() {
    let document = read(INK).expect("InkML fixture must parse");
    let context = document.contexts()[0].source_span();
    let trace = document.traces()[0].source_span();
    let property = document.brush_properties()[0].source_span();
    let metadata = read_metadata_with_source_spans(INK, &[context], &[trace], &[property], &[])
        .expect("exact semantic spans must be accepted");
    assert_eq!(metadata.context_count(), 1);
    assert_eq!(metadata.trace_count(), 1);
    assert_eq!(metadata.brush_property_count(), 1);

    assert!(
        read_metadata_with_source_spans(
            INK,
            &[SourceSpan::new(context.start(), context.end() + 1)],
            &[],
            &[],
            &[],
        )
        .is_err()
    );
    assert!(read_metadata_with_source_spans(INK, &[context, context], &[], &[], &[]).is_err());
    assert!(
        read_metadata_with_source_spans(
            INK,
            &[SourceSpan::new(INK.len() + 1, INK.len() + 1)],
            &[],
            &[],
            &[],
        )
        .is_err()
    );
}

#[test]
fn metadata_owns_typed_values_without_retaining_input() {
    let mut input = INK.to_vec();
    let metadata = read_metadata(&input).expect("InkML metadata fixture must parse");
    input.fill(b'x');
    drop(input);

    assert_eq!(metadata.contexts()[0].kind(), &ContextKind::WritingRegion);
    assert_eq!(
        metadata.contexts()[0].semantic_type(),
        Some(&SemanticType::Comment)
    );
    assert_eq!(metadata.traces()[0].context_ref(), Some("#ctx0"));
    assert_eq!(metadata.brush_properties()[1].value(), "pencil");
}

#[test]
fn read_metadata_rejects_malformed_and_oversized_payloads() {
    let malformed =
        br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML">&future;</inkml:ink>"#;
    assert!(read_metadata(malformed).is_err());

    let oversized = vec![b' '; MAX_SOURCE_BYTES + 1];
    assert!(read_metadata(&oversized).is_err());
}

#[test]
fn read_shared_rejects_malformed_and_oversized_payloads_before_retention() {
    let malformed = Arc::new(
        br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML">&future;</inkml:ink>"#.to_vec(),
    );
    let malformed_count = Arc::strong_count(&malformed);
    assert!(read_shared(Arc::clone(&malformed)).is_err());
    assert_eq!(Arc::strong_count(&malformed), malformed_count);

    let oversized = Arc::new(vec![b' '; MAX_SOURCE_BYTES + 1]);
    let oversized_count = Arc::strong_count(&oversized);
    assert!(read_shared(Arc::clone(&oversized)).is_err());
    assert_eq!(Arc::strong_count(&oversized), oversized_count);
}

#[test]
fn rejects_invalid_context_guid_and_points() {
    let invalid_guid = br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:msink="http://schemas.microsoft.com/ink/2010/main"><msink:context id="{bad}" type="inkDrawing"/></inkml:ink>"#;
    assert!(read(invalid_guid).is_err());
    let invalid_points = br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:msink="http://schemas.microsoft.com/ink/2010/main"><msink:context type="inkDrawing" rotatedBoundingBox="0,0 nope"/></inkml:ink>"#;
    assert!(read(invalid_points).is_err());
}

#[test]
fn ink_actions_are_typed_and_preserved() {
    let source = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0.25"><iact:actionData name="stroke"/></iact:action><iact:actionGroup type="transform" startTime="1"><iact:action type="remove" startTime="2"/></iact:actionGroup></iact:actions>"#;
    let actions = actions::read(source).expect("ink actions fixture must parse");
    assert_eq!(actions.length_unit(), "cm");
    assert_eq!(actions.time_unit(), "s");
    assert_eq!(actions.action_group_count(), 1);
    assert_eq!(actions.actions().len(), 2);
    assert_eq!(actions.actions()[0].action_type().as_str(), "add");
    assert_eq!(actions.actions()[1].start_time(), "2");
    assert_eq!(
        actions::write(&actions).expect("source-backed action write"),
        source
    );
}

#[test]
fn ink_actions_reject_document_level_references_and_invalid_units() {
    let trailing_reference = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"/>&future;"#;
    assert!(actions::read(trailing_reference).is_err());
    let empty_unit = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="" timeUnit="s"/>"#;
    assert!(actions::read(empty_unit).is_err());
}

#[test]
fn ink_actions_support_ancestor_bound_prefix_fragments() {
    let source = br#"<iact:actions lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0"/></iact:actions>"#;
    let parsed = actions::read(source).expect("ancestor-bound action fragment");
    assert_eq!(parsed.actions().len(), 1);
    assert_eq!(
        actions::write(&parsed).expect("exact action fragment"),
        source
    );
}

#[test]
fn rejects_forbidden_ink_markup() {
    let dtd = br#"<!DOCTYPE inkml:ink><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>"#;
    assert!(read(dtd).is_err());
    let late_declaration =
        br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/><?xml version="1.0"?>"#;
    assert!(read(late_declaration).is_err());
}

#[test]
fn rejects_invalid_text_cdata_references_and_utf8_before_retaining_source() {
    let mut invalid_utf8 = br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML">"#.to_vec();
    invalid_utf8.push(0xff);
    invalid_utf8.extend_from_slice(br#"</inkml:ink>"#);
    assert!(read(&invalid_utf8).is_err());

    let top_level_cdata =
        br#"<![CDATA[ ]]><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>"#;
    assert!(read(top_level_cdata).is_err());

    let unknown_reference =
        br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML">&future;</inkml:ink>"#;
    assert!(read(unknown_reference).is_err());
    let invalid_character_reference =
        br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML">&#0;</inkml:ink>"#;
    assert!(read(invalid_character_reference).is_err());
}

#[test]
fn accepts_and_preserves_a_utf8_bom_before_the_xml_declaration() {
    let mut source = vec![0xef, 0xbb, 0xbf];
    source.extend_from_slice(
        br#"<?xml version="1.0" encoding="UTF-8"?><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>"#,
    );
    let document = read(&source).expect("UTF-8 BOM should be accepted");
    assert_eq!(document.source(), source.as_slice());
}

#[test]
fn rejects_duplicate_declarations_and_bound_foreign_root_prefixes() {
    let duplicate = br#"<?xml version="1.0"?><?xml version="1.0"?><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>"#;
    assert!(read(duplicate).is_err());
    let foreign = br#"<inkml:ink xmlns:inkml="urn:foreign"/>"#;
    assert!(read(foreign).is_err());

    let action_foreign =
        br#"<iact:actions xmlns:iact="urn:foreign" lengthUnit="cm" timeUnit="s"/>"#;
    assert!(actions::read(action_foreign).is_err());

    let malformed_declaration = br#"<?xml version="1.0" foo="bar"?><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>"#;
    assert!(read(malformed_declaration).is_err());
    let invalid_numeric_reference =
        br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s">&#+65;</iact:actions>"#;
    assert!(actions::read(invalid_numeric_reference).is_err());
}

#[test]
fn rejects_duplicate_expanded_attributes_in_opaque_elements() {
    let duplicate = br#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:a="urn:opaque" xmlns:b="urn:opaque" a:value="one" b:value="two"/>"#;
    assert!(read(duplicate).is_err());
    let action_duplicate = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:a="urn:opaque" xmlns:b="urn:opaque" lengthUnit="cm" timeUnit="s" a:value="one" b:value="two"/>"#;
    assert!(actions::read(action_duplicate).is_err());
}

#[test]
fn strict_ink_actions_profile_exposes_ordered_typed_children() {
    let source = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:inkml="http://www.w3.org/2003/InkML" xml:id="root" lengthUnit="cm" timeUnit="ms"><inkml:definitions><inkml:future/></inkml:definitions><iact:action xml:id="a0" type="add" startTime="0"><iact:property name="dataType"/><iact:actionDataGroup xml:id="g0" name="path"><iact:actionData xml:id="d0" name="target"><iact:transform/><inkml:trace/><inkml:traceView/></iact:actionData></iact:actionDataGroup></iact:action><iact:actionGroup xml:id="ag0" type="transform" startTime="1"><iact:action type="remove" startTime="2"/></iact:actionGroup></iact:actions>"#;
    let profile = actions::read_profile(source).expect("strict action profile");
    assert_eq!(profile.length_unit(), actions::LengthUnit::Centimeter);
    assert_eq!(profile.time_unit(), actions::TimeUnit::Millisecond);
    assert_eq!(profile.xml_id(), Some("root"));
    assert!(profile.definitions_span().is_some());
    assert_eq!(profile.children().len(), 2);
    assert_eq!(actions::write_profile(&profile).unwrap(), source);

    let actions::RootChild::Action(action) = &profile.children()[0] else {
        panic!("first root child must be an action");
    };
    assert_eq!(action.properties().len(), 1);
    assert_eq!(action.xml_id(), Some("a0"));
    assert_eq!(action.properties()[0].name(), "dataType");
    assert_eq!(action.properties()[0].value(), "ink");
    assert_eq!(action.children().len(), 2);
    assert!(matches!(
        action.children()[0],
        actions::ActionChild::Property(_)
    ));
    let actions::ActionChild::DataGroup(group) = &action.children()[1] else {
        panic!("second action child must be a data group");
    };
    assert_eq!(group.name(), "path");
    assert_eq!(group.xml_id(), Some("g0"));
    assert_eq!(group.data().len(), 1);
    assert_eq!(group.data()[0].name(), "target");
    assert_eq!(group.data()[0].xml_id(), Some("d0"));
    assert_eq!(group.data()[0].children().len(), 3);
    assert!(matches!(
        group.data()[0].children()[0],
        actions::DataChild::Transform(_)
    ));
    assert!(matches!(
        group.data()[0].children()[1],
        actions::DataChild::Trace(_)
    ));
    assert!(matches!(
        group.data()[0].children()[2],
        actions::DataChild::TraceView(_)
    ));
    assert_eq!(
        action.xml_from(profile.source()),
        br#"<iact:action xml:id="a0" type="add" startTime="0"><iact:property name="dataType"/><iact:actionDataGroup xml:id="g0" name="path"><iact:actionData xml:id="d0" name="target"><iact:transform/><inkml:trace/><inkml:traceView/></iact:actionData></iact:actionDataGroup></iact:action>"#
    );
    let actions::RootChild::ActionGroup(group) = &profile.children()[1] else {
        panic!("second root child must be an action group");
    };
    assert_eq!(group.xml_id(), Some("ag0"));
}

#[test]
fn strict_ink_actions_profile_rejects_invalid_structure_but_keeps_legacy_read() {
    let prefix = r#"xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s""#;
    let legacy = br#"<iact:actions lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0"/></iact:actions>"#;
    assert!(actions::read(legacy).is_ok());
    assert!(actions::read_profile(legacy).is_err());
    let legacy_opaque = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:x="urn:future" lengthUnit="cm" timeUnit="s"><x:future/></iact:actions>"#;
    assert!(actions::read(legacy_opaque).is_ok());
    assert!(actions::read_profile(legacy_opaque).is_err());

    let cases = [
        format!("<iact:actions {prefix}><iact:property name=\"x\"/></iact:actions>"),
        format!(
            "<iact:actions {prefix}><iact:action type=\"add\" startTime=\"0\"><iact:actionData name=\"stroke\"/><iact:property name=\"late\"/></iact:action></iact:actions>"
        ),
        format!(
            "<iact:actions xmlns:inkml=\"http://www.w3.org/2003/InkML\" {prefix}><iact:action type=\"add\" startTime=\"0\"><iact:actionData><inkml:trace/><iact:transform/></iact:actionData></iact:action></iact:actions>"
        ),
        format!(
            "<iact:actions {prefix}><iact:action type=\"add\" startTime=\"0\"><iact:actionData><iact:transform/><iact:transform/></iact:actionData></iact:action></iact:actions>"
        ),
        format!(
            "<iact:actions {prefix}><iact:action type=\"add\" startTime=\"0\"><iact:action><iact:actionData/></iact:action></iact:action></iact:actions>"
        ),
        format!(
            "<iact:actions xmlns:inkml=\"http://www.w3.org/2003/InkML\" {prefix}><iact:action type=\"add\" startTime=\"0\"/><inkml:definitions/></iact:actions>"
        ),
        format!(
            "<iact:actions {prefix}><iact:actionGroup type=\"add\" startTime=\"0\"/></iact:actions>"
        ),
        format!(
            "<iact:actions {prefix}><iact:action type=\"add\" startTime=\"0\"><iact:actionDataGroup/></iact:action></iact:actions>"
        ),
        format!("<iact:actions {prefix} lengthUnit=\"px\"/>"),
    ];
    for case in cases {
        assert!(actions::read_profile(case.as_bytes()).is_err(), "{case}");
    }
    assert!(actions::read_profile(
        br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="px" timeUnit="s"/>"#
    )
    .is_err());
}

#[test]
fn strict_ink_actions_profile_preserves_opaque_descendants_and_scopes_reserved_semantics() {
    let source = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:x="urn:future" xmlns:inkml="http://www.w3.org/2003/InkML" lengthUnit="m" timeUnit="s"><iact:action type="add" startTime="0"><iact:actionData><iact:transform><x:future><x:payload>opaque</x:payload></x:future></iact:transform></iact:actionData></iact:action></iact:actions>"#;
    let profile = actions::read_profile(source).expect("opaque action descendant");
    let actions::RootChild::Action(action) = &profile.children()[0] else {
        panic!("root action expected");
    };
    assert_eq!(action.action_type().as_str(), "add");
    assert!(matches!(
        action.children().first(),
        Some(actions::ActionChild::Data(_))
    ));
    assert_eq!(profile.source(), source);
}

#[test]
fn strict_ink_actions_profile_rejects_schema_owned_unknowns_and_canonicalizes_namespaces() {
    let encoded = br#"<iact:actions xmlns:iact="http:&#x2F;&#x2F;schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"/>"#;
    assert!(actions::read_profile(encoded).is_ok());

    let unknown_root_attribute = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s" future="x"/>"#;
    assert!(actions::read_profile(unknown_root_attribute).is_err());

    let unknown_action_attribute = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0" future="x"/></iact:actions>"#;
    assert!(actions::read_profile(unknown_action_attribute).is_err());

    let foreign_root_child = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:x="urn:future" lengthUnit="cm" timeUnit="s"><x:future/></iact:actions>"#;
    assert!(actions::read_profile(foreign_root_child).is_err());

    let foreign_action_child = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:x="urn:future" lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0"><x:future/></iact:action></iact:actions>"#;
    assert!(actions::read_profile(foreign_action_child).is_err());

    let wrong_namespace_action = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:x="urn:future" lengthUnit="cm" timeUnit="s"><x:action type="add" startTime="0"/></iact:actions>"#;
    assert!(actions::read_profile(wrong_namespace_action).is_err());

    let encoded_duplicate = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" xmlns:x="urn:future" xmlns:a="urn:foo&#x2F;" xmlns:b="urn:foo/" lengthUnit="cm" timeUnit="s"><iact:action type="add" startTime="0"><iact:actionData><iact:transform a:value="one" b:value="two"/></iact:actionData></iact:action></iact:actions>"#;
    assert!(actions::read_profile(encoded_duplicate).is_err());
}

#[test]
fn strict_ink_actions_profile_collapses_decimal_whitespace_and_allows_empty_user_type() {
    let source = br#"<iact:actions xmlns:iact="http://schemas.microsoft.com/office/powerpoint/2014/inkAction" lengthUnit="cm" timeUnit="s"><iact:action type="" startTime="&#x9; 1&#xA;"/></iact:actions>"#;
    let profile =
        actions::read_profile(source).expect("xsd decimal whitespace and empty user type");
    let actions::RootChild::Action(action) = &profile.children()[0] else {
        panic!("root action expected");
    };
    assert_eq!(action.action_type().as_str(), "");
    assert_eq!(action.start_time(), "1");
    assert_eq!(profile.source(), source);
}

#[test]
fn accepts_exact_depth_and_rejects_empty_element_overflow() {
    let mut exact = String::from(
        r#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:x="urn:opaque">"#,
    );
    for _ in 1..MAX_DEPTH - 1 {
        exact.push_str("<x:node>");
    }
    exact.push_str("<x:leaf/>");
    for _ in 1..MAX_DEPTH - 1 {
        exact.push_str("</x:node>");
    }
    exact.push_str("</inkml:ink>");
    assert!(read(exact.as_bytes()).is_ok());

    let mut overflow = String::from(
        r#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:x="urn:opaque">"#,
    );
    for _ in 1..MAX_DEPTH {
        overflow.push_str("<x:node>");
    }
    overflow.push_str("<x:leaf/>");
    for _ in 1..MAX_DEPTH {
        overflow.push_str("</x:node>");
    }
    overflow.push_str("</inkml:ink>");
    assert!(read(overflow.as_bytes()).is_err());
}

#[test]
fn accepts_exact_node_and_attribute_value_limits() {
    let mut exact_nodes = String::from(
        r#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" xmlns:x="urn:opaque">"#,
    );
    for _ in 1..MAX_NODES {
        exact_nodes.push_str("<x:node/>");
    }
    exact_nodes.push_str("</inkml:ink>");
    assert!(read(exact_nodes.as_bytes()).is_ok());

    let mut over_nodes = exact_nodes[..exact_nodes.len() - 11].to_owned();
    over_nodes.push_str("<x:node/></inkml:ink>");
    assert!(read(over_nodes.as_bytes()).is_err());

    let exact_value = "a".repeat(MAX_ATTRIBUTE_VALUE_BYTES);
    let exact_attribute =
        format!(r#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" data="{exact_value}"/>"#);
    assert!(read(exact_attribute.as_bytes()).is_ok());
    let over_value = "a".repeat(MAX_ATTRIBUTE_VALUE_BYTES + 1);
    let over_attribute =
        format!(r#"<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML" data="{over_value}"/>"#);
    assert!(read(over_attribute.as_bytes()).is_err());
}
