use litchi_core::sheet::Result;
use litchi_xlsx::xml_maps::{
    DataBinding, XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapSchema, XmlSchema,
    parse_map_info, parse_xml_map_info, serialize_map_info, serialize_xml_map_info,
};

fn legacy_value() -> XmlMapInfo {
    let schema: XmlMapSchema = XmlSchema {
        id: "schema-1".to_owned(),
        schema_reference: None,
        namespace: Some("urn:example".to_owned()),
        payload_xml: None,
    };
    let binding: XmlMapDataBinding = DataBinding {
        data_binding_name: None,
        file_binding: None,
        connection_id: None,
        file_binding_name: None,
        load_mode: 0,
        payload_xml: None,
    };
    XmlMapInfo {
        selection_namespaces: "xmlns:x='urn:example'".to_owned(),
        schemas: vec![schema],
        maps: vec![XmlMap {
            id: 1,
            name: "Map1".to_owned(),
            root_element: "root".to_owned(),
            schema_id: "schema-1".to_owned(),
            show_import_export_validation_errors: false,
            auto_fit: false,
            append: false,
            preserve_sort_auto_filter_layout: false,
            preserve_format: false,
            data_binding: Some(binding),
        }],
    }
}

fn legacy_inherent_round_trip(value: &XmlMapInfo) -> Result<XmlMapInfo> {
    let xml = value.to_xml(false)?;
    XmlMapInfo::parse(&xml)
}

#[test]
fn legacy_types_struct_literals_and_inherent_methods_remain_usable() -> Result<()> {
    let value = legacy_value();
    assert_eq!(legacy_inherent_round_trip(&value)?, value);
    Ok(())
}

#[test]
fn xlsx_adapters_retain_the_boxed_error_result_surface() -> Result<()> {
    let _: fn(&[u8]) -> Result<XmlMapInfo> = XmlMapInfo::parse;
    let _: fn(&XmlMapInfo, bool) -> Result<Vec<u8>> = XmlMapInfo::to_xml;
    let _: fn(&[u8]) -> Result<XmlMapInfo> = parse_xml_map_info;
    let _: fn(&[u8]) -> Result<XmlMapInfo> = parse_map_info;
    let _: fn(&XmlMapInfo, XmlMapConformance) -> Result<Vec<u8>> = serialize_xml_map_info;
    let _: fn(&XmlMapInfo, XmlMapConformance) -> Result<Vec<u8>> = serialize_map_info;

    let value = legacy_value();
    let xml = serialize_xml_map_info(&value, XmlMapConformance::Transitional)?;
    assert_eq!(parse_xml_map_info(&xml)?, value);
    Ok(())
}
