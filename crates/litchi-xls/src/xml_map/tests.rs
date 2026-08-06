//! Focused XML-map round-trip and rejection coverage.

use super::{
    DataBinding, LoadMode, Map, MapId, MapInfo, OpaqueXml, Schema, SchemaId, XPath, parse, write,
};

fn valid_xml() -> &'static [u8] {
    br#"<?xml version="1.0" encoding="utf-8"?>
<MapInfo SelectionNamespaces="xmlns:x='urn:example'" xmlns="urn:excel">
  <Schema ID="main"><x:schema xmlns:x="http://www.w3.org/2001/XMLSchema"><x:element name="root"/></x:schema></Schema>
  <Map ID="7" Name="Example" RootElement="root" SchemaID="main" ShowImportExportValidationErrors="true" AutoFit="false" Append="true" PreserveSortAFLayout="false" PreserveFormat="true">
    <DataBinding DataBindingName="binding" FileBinding="source.xml" FileBindingName="source" DataBindingLoadMode="1"><x:source xmlns:x="urn:source"/></DataBinding>
  </Map>
</MapInfo>"#
}

#[test]
fn parses_and_serializes_the_complete_vertical_slice() {
    let parsed = parse(valid_xml());
    assert!(parsed.is_ok());
    let value = parsed.unwrap_or_else(|error| panic!("parse failed: {error}"));
    assert_eq!(value.schemas().len(), 1);
    assert_eq!(value.maps().len(), 1);
    assert_eq!(value.maps()[0].id().get(), 7);
    assert_eq!(
        value.maps()[0].data_binding().map(|v| v.load_mode()),
        Some(LoadMode::Normal)
    );
    assert_eq!(value.schemas()[0].payload().as_bytes(), br#"<x:schema xmlns:x="http://www.w3.org/2001/XMLSchema"><x:element name="root"/></x:schema>"#);

    let encoded = write(&value);
    assert!(encoded.is_ok());
    let reparsed = encoded
        .and_then(|xml| parse(&xml))
        .unwrap_or_else(|error| panic!("round-trip parse failed: {error}"));
    assert_eq!(reparsed, value);
}

#[test]
fn typed_constructors_enforce_ids_and_xpath_bounds() {
    assert!(MapId::new(0).is_err());
    assert!(MapId::new(2_147_483_648).is_err());
    assert!(MapId::new(2_147_483_647).is_ok());
    assert!(XPath::new("x".repeat(32_000)).is_err());
    assert!(XPath::new("x".repeat(31_999)).is_ok());
}

#[test]
fn rejects_invalid_booleans_dependencies_bindings_and_unknown_attributes() {
    let cases = [
        br#"<MapInfo SelectionNamespaces=""><Schema ID="s"><x/></Schema><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="1" AutoFit="false" Append="false" PreserveSortAFLayout="false" PreserveFormat="false"/></MapInfo>"#.as_slice(),
        br#"<MapInfo SelectionNamespaces=""><Schema ID="s"><x/></Schema><Map ID="1" Name="m" RootElement="r" SchemaID="missing" ShowImportExportValidationErrors="false" AutoFit="false" Append="false" PreserveSortAFLayout="false" PreserveFormat="false"/></MapInfo>"#.as_slice(),
        br#"<MapInfo SelectionNamespaces=""><Schema ID="s"><x/></Schema><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="false" AutoFit="false" Append="false" PreserveSortAFLayout="false" PreserveFormat="false"><DataBinding FileBinding="true" DataBindingLoadMode="0"/></Map></MapInfo>"#.as_slice(),
        br#"<MapInfo SelectionNamespaces=""><Schema ID="s" unexpected="x"><x/></Schema></MapInfo>"#.as_slice(),
        br#"<!DOCTYPE MapInfo [<!ENTITY x SYSTEM "file:///etc/passwd">]><MapInfo SelectionNamespaces=""><Schema ID="s"><x>&x;</x></Schema></MapInfo>"#.as_slice(),
    ];
    for xml in cases {
        assert!(parse(xml).is_err(), "invalid XML map was accepted");
    }
}

#[test]
fn rejects_duplicate_map_names_and_schema_cycles() {
    let duplicate = br#"<MapInfo SelectionNamespaces=""><Schema ID="s"><x/></Schema><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="false" AutoFit="false" Append="false" PreserveSortAFLayout="false" PreserveFormat="false"/><Map ID="2" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="false" AutoFit="false" Append="false" PreserveSortAFLayout="false" PreserveFormat="false"/></MapInfo>"#;
    assert!(parse(duplicate).is_err());

    let cycle = br#"<MapInfo SelectionNamespaces=""><Schema ID="a" SchemaRef="b"><x/></Schema><Schema ID="b" SchemaRef="a"><x/></Schema></MapInfo>"#;
    assert!(parse(cycle).is_err());
}

#[test]
fn public_model_can_build_and_write_inert_metadata() {
    let payload = OpaqueXml::try_new(br#"<x:data xmlns:x="urn:x"/>"#.to_vec());
    assert!(payload.is_ok());
    let payload = payload.unwrap_or_else(|error| panic!("payload failed: {error}"));
    let schema_id = SchemaId::new("schema");
    assert!(schema_id.is_ok());
    let schema_id = schema_id.unwrap_or_else(|error| panic!("schema failed: {error}"));
    let schema = Schema::try_new(schema_id.clone(), payload);
    assert!(schema.is_ok());
    let schema = schema.unwrap_or_else(|error| panic!("schema failed: {error}"));
    let map_id = MapId::new(1);
    assert!(map_id.is_ok());
    let map_id = map_id.unwrap_or_else(|error| panic!("map ID failed: {error}"));
    let map = Map::try_new(
        map_id, "map", "root", schema_id, false, true, false, true, true,
    );
    assert!(map.is_ok());
    let map = map.unwrap_or_else(|error| panic!("map failed: {error}"));
    let info = MapInfo::try_new("", vec![schema], vec![map]);
    assert!(info.is_ok());
    let binding = DataBinding::try_new("data.xml", LoadMode::DelayLoad);
    assert!(binding.is_ok());
    assert_eq!(
        binding
            .unwrap_or_else(|error| panic!("binding failed: {error}"))
            .load_mode(),
        LoadMode::DelayLoad
    );
    assert!(write(&info.unwrap_or_else(|error| panic!("info failed: {error}"))).is_ok());
}
