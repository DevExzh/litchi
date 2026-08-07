use super::*;

fn fixture() -> XmlMapInfo {
    XmlMapInfo {
        selection_namespaces: "xmlns:x='urn:test'".into(),
        schemas: vec![XmlSchema {
            id: "schema-1".into(),
            schema_reference: Some("urn:test".into()),
            namespace: Some("urn:test".into()),
            payload_xml: Some(
                format!(r#"<x:schema xmlns:x="urn:test" xmlns="{NS_TEXT}"/>"#).into_bytes(),
            ),
        }],
        maps: vec![XmlMap {
            id: 1,
            name: "map".into(),
            root_element: "root".into(),
            schema_id: "schema-1".into(),
            show_import_export_validation_errors: true,
            auto_fit: true,
            append: false,
            preserve_sort_auto_filter_layout: true,
            preserve_format: true,
            data_binding: None,
        }],
    }
}

#[test]
fn parses_validates_and_serializes_both_conformance_families() {
    let value = fixture();
    validate_xml_map_info(&value).unwrap();
    for conformance in [XmlMapConformance::Transitional, XmlMapConformance::Strict] {
        let xml = serialize_xml_map_info(&value, conformance).unwrap();
        assert_eq!(parse_xml_map_info(&xml).unwrap(), value);
    }
}

#[test]
fn source_patch_preserves_unmodeled_markup() {
    let before = fixture();
    let source = serialize_xml_map_info(&before, XmlMapConformance::Transitional).unwrap();
    let mut after = before.clone();
    after.maps[0].name = "edited".into();
    let patched = patch_xml_map_info_source(
        &source,
        &before,
        &after,
        XmlMapConformance::Transitional,
        XmlMapConformance::Transitional,
    )
    .unwrap();
    assert_eq!(parse_xml_map_info(&patched).unwrap(), after);
}

#[test]
fn exposes_the_codec_limits() {
    assert_eq!(XmlMapLimits::default(), XmlMapLimits::DEFAULT);
    assert_eq!(XmlMapLimits::DEFAULT.max_maps, 65_536);
}

#[test]
fn caller_limits_cover_every_bounded_resource() {
    let value = fixture();
    let xml = serialize_xml_map_info(&value, XmlMapConformance::Transitional).unwrap();

    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_part_bytes = xml.len() - 1;
    assert!(parse_xml_map_info_with_limits(&xml, &limits).is_err());
    assert!(
        serialize_xml_map_info_with_limits(&value, XmlMapConformance::Transitional, &limits,)
            .is_err()
    );

    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_schemas = 0;
    assert!(validate_xml_map_info_with_limits(&value, &limits).is_err());
    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_maps = 0;
    assert!(validate_xml_map_info_with_limits(&value, &limits).is_err());
    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_string_bytes = value.selection_namespaces.len() - 1;
    assert!(validate_xml_map_info_with_limits(&value, &limits).is_err());
    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_opaque_bytes = value.schemas[0].payload_xml.as_ref().unwrap().len() - 1;
    assert!(validate_xml_map_info_with_limits(&value, &limits).is_err());
    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_depth = 0;
    assert!(validate_xml_map_info_with_limits(&value, &limits).is_err());
    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_events = 1;
    assert!(validate_xml_map_info_with_limits(&value, &limits).is_err());
}

#[test]
fn enforces_office_map_and_binding_boundaries() {
    let mut value = fixture();
    value.maps[0].id = i32::MAX as u32;
    value.maps[0].name = "n".repeat(65_535);
    validate_xml_map_info(&value).unwrap();

    value.maps[0].id = 0;
    assert!(validate_xml_map_info(&value).is_err());
    value.maps[0].id = i32::MAX as u32 + 1;
    assert!(validate_xml_map_info(&value).is_err());
    value.maps[0].id = 1;
    value.maps[0].name.push('n');
    assert!(validate_xml_map_info(&value).is_err());

    value.maps[0].name = "map".into();
    value.maps[0].data_binding = Some(DataBinding {
        data_binding_name: Some("binding".into()),
        file_binding: Some(true),
        connection_id: Some(i32::MAX as u32),
        file_binding_name: Some("file".into()),
        load_mode: 1,
        payload_xml: None,
    });
    validate_xml_map_info(&value).unwrap();
    value.maps[0].data_binding.as_mut().unwrap().connection_id = None;
    assert!(validate_xml_map_info(&value).is_err());
    {
        let binding = value.maps[0].data_binding.as_mut().unwrap();
        binding.file_binding = Some(false);
        binding.file_binding_name = None;
    }
    assert!(validate_xml_map_info(&value).is_ok());
    value.maps[0].data_binding.as_mut().unwrap().connection_id = Some(1);
    assert!(validate_xml_map_info(&value).is_err());
}

#[test]
fn detects_conformance_and_canonicalizes_boolean_words() {
    let value = fixture();
    for conformance in [XmlMapConformance::Transitional, XmlMapConformance::Strict] {
        let xml = serialize_xml_map_info(&value, conformance).unwrap();
        let text = std::str::from_utf8(&xml).unwrap();
        assert!(text.contains("AutoFit=\"true\""));
        assert!(text.contains("Append=\"false\""));
        let parsed = parse_xml_map_info_with_conformance(&xml).unwrap();
        assert_eq!(parsed.info, value);
        assert_eq!(parsed.conformance, conformance);
    }

    let xml = serialize_xml_map_info(&value, XmlMapConformance::Transitional).unwrap();
    let numeric = String::from_utf8(xml)
        .unwrap()
        .replace("=\"true\"", "=\"1\"")
        .replace("=\"false\"", "=\"0\"");
    assert!(parse_xml_map_info(numeric.as_bytes()).is_err());
}

#[test]
fn rejects_numeric_map_and_file_binding_booleans() {
    let namespace = NS_TEXT;
    let valid = format!(
        r#"<MapInfo xmlns="{namespace}" SelectionNamespaces=""><Schema ID="s"/><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="false" AutoFit="true" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"><DataBinding FileBinding="true" ConnectionID="1" DataBindingLoadMode="1"/></Map></MapInfo>"#,
    );
    assert!(parse_xml_map_info(valid.as_bytes()).is_ok());

    for invalid in [
        valid.replacen("AutoFit=\"true\"", "AutoFit=\"1\"", 1),
        valid.replacen("Append=\"false\"", "Append=\"0\"", 1),
        valid.replacen("FileBinding=\"true\"", "FileBinding=\"1\"", 1),
    ] {
        assert!(parse_xml_map_info(invalid.as_bytes()).is_err());
    }
}

#[test]
fn schema_language_is_ignored_but_preserved_by_source_patch() {
    let namespace = NS_TEXT;
    let source = format!(
        r#"<?xml version="1.0"?><MapInfo xmlns="{namespace}" SelectionNamespaces=""><Schema ID="s" SchemaLanguage="en-US"/><Map ID="1" Name="old" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="false" AutoFit="true" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"/></MapInfo>"#,
    );
    let before = parse_xml_map_info(source.as_bytes()).unwrap();
    let mut after = before.clone();
    after.maps[0].name = "renamed".into();
    let patched = patch_xml_map_info_source(
        source.as_bytes(),
        &before,
        &after,
        XmlMapConformance::Transitional,
        XmlMapConformance::Transitional,
    )
    .unwrap();
    let patched = std::str::from_utf8(&patched).unwrap();
    assert!(patched.contains("SchemaLanguage=\"en-US\""));
    assert!(patched.contains("Name=\"renamed\""));
}

#[test]
fn source_patch_preserves_comments_extensions_and_canonicalizes_edits() {
    let namespace = NS_TEXT;
    let source = format!(
        r#"<?xml version="1.0"?><MapInfo xmlns="{namespace}" xmlns:f="urn:future" SelectionNamespaces=""><!--keep--><Schema ID="s"><f:payload/></Schema><Map ID="1" Name="old" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="false" AutoFit="true" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"/></MapInfo>"#,
    );
    let before = parse_xml_map_info(source.as_bytes()).unwrap();
    let mut after = before.clone();
    after.maps[0].name = "new".into();
    after.maps[0].append = true;
    let patched = patch_xml_map_info_source(
        source.as_bytes(),
        &before,
        &after,
        XmlMapConformance::Transitional,
        XmlMapConformance::Transitional,
    )
    .unwrap();
    let patched = std::str::from_utf8(&patched).unwrap();
    assert!(patched.contains("<!--keep-->"));
    assert!(patched.contains("<f:payload/>"));
    assert!(patched.contains("Name=\"new\""));
    assert!(patched.contains("Append=\"true\""));
}

#[test]
fn serializes_a_large_catalog_with_amortized_growth() {
    let mut value = fixture();
    value.schemas[0].payload_xml = None;
    value.maps = (1..=10_000)
        .map(|id| XmlMap {
            id,
            name: format!("catalog-map-{id:05}-{}", "n".repeat(48)),
            root_element: format!("root{id}"),
            schema_id: "schema-1".into(),
            show_import_export_validation_errors: id % 2 == 0,
            auto_fit: true,
            append: false,
            preserve_sort_auto_filter_layout: true,
            preserve_format: true,
            data_binding: None,
        })
        .collect();

    let xml = serialize_xml_map_info(&value, XmlMapConformance::Transitional).unwrap();
    assert!(xml.len() < XmlMapLimits::DEFAULT.max_part_bytes);
    assert_eq!(parse_xml_map_info(&xml).unwrap(), value);
}

#[test]
fn source_patch_honors_relaxed_limits_above_default() {
    let before = fixture();
    let source = serialize_xml_map_info(&before, XmlMapConformance::Transitional).unwrap();
    let mut after = before.clone();
    after.selection_namespaces = "n".repeat(XmlMapLimits::DEFAULT.max_part_bytes + 1024);
    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_part_bytes = after.selection_namespaces.len() + source.len() + 1024;
    limits.max_string_bytes = after.selection_namespaces.len();

    let patched = patch_xml_map_info_source_with_limits(
        &source,
        &before,
        &after,
        XmlMapConformance::Transitional,
        XmlMapConformance::Transitional,
        &limits,
    )
    .unwrap();
    assert!(patched.len() > XmlMapLimits::DEFAULT.max_part_bytes);
    assert_eq!(
        parse_xml_map_info_with_limits(&patched, &limits).unwrap(),
        after
    );
}

#[test]
fn source_patch_enforces_the_exact_output_boundary() {
    let before = fixture();
    let source = serialize_xml_map_info(&before, XmlMapConformance::Transitional).unwrap();
    let mut after = before.clone();
    after.selection_namespaces.push_str("-longer");
    let expected = patch_xml_map_info_source(
        &source,
        &before,
        &after,
        XmlMapConformance::Transitional,
        XmlMapConformance::Transitional,
    )
    .unwrap();

    let mut limits = XmlMapLimits::DEFAULT;
    limits.max_part_bytes = expected.len();
    assert_eq!(
        patch_xml_map_info_source_with_limits(
            &source,
            &before,
            &after,
            XmlMapConformance::Transitional,
            XmlMapConformance::Transitional,
            &limits,
        )
        .unwrap(),
        expected,
    );
    limits.max_part_bytes -= 1;
    assert!(
        patch_xml_map_info_source_with_limits(
            &source,
            &before,
            &after,
            XmlMapConformance::Transitional,
            XmlMapConformance::Transitional,
            &limits,
        )
        .is_err()
    );
}

#[test]
fn borrowed_projection_serializes_and_patches_without_payload_clones() {
    let before = fixture();
    let before_ref = XmlMapInfoRef::try_from(&before).unwrap();
    assert!(std::ptr::eq(
        before_ref.schemas[0].payload_xml.unwrap().as_ptr(),
        before.schemas[0].payload_xml.as_ref().unwrap().as_ptr(),
    ));
    assert!(std::ptr::eq(
        before_ref.maps[0].name.as_ptr(),
        before.maps[0].name.as_ptr(),
    ));

    let owned_xml = serialize_xml_map_info(&before, XmlMapConformance::Transitional).unwrap();
    let borrowed_xml =
        serialize_xml_map_info_ref(&before_ref, XmlMapConformance::Transitional).unwrap();
    assert_eq!(borrowed_xml, owned_xml);
    validate_xml_map_info_ref(&before_ref).unwrap();

    let mut after = before.clone();
    after.maps[0].name = "borrowed-edit".into();
    let after_ref = XmlMapInfoRef::try_from(&after).unwrap();
    let owned_patch = patch_xml_map_info_source(
        &owned_xml,
        &before,
        &after,
        XmlMapConformance::Transitional,
        XmlMapConformance::Transitional,
    )
    .unwrap();
    let borrowed_patch = patch_xml_map_info_source_ref(
        &owned_xml,
        &before_ref,
        &after_ref,
        XmlMapConformance::Transitional,
        XmlMapConformance::Transitional,
    )
    .unwrap();
    assert_eq!(borrowed_patch, owned_patch);
}
