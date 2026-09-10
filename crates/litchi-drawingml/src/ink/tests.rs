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
