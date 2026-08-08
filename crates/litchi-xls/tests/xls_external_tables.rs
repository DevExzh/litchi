use std::io::Cursor;

use litchi_xls::writer::Writer;
use litchi_xls::{
    ExternalTableField, ExternalTableMetadata, ExternalTableVersion, ListColumnId, ListObject,
    ListObjectColumn, ListObjectFeatureVersion, ListObjectId, ListObjectRange,
    ListObjectSourceMetadata, ListObjectStyleOptions, Map, MapId, MapInfo, OpaqueXml, Schema,
    SchemaId, WebColumnType, WebDefaultValue, WebFieldInfo, WebInvalidCell, WebTableField,
    WebTableMetadata, Workbook, XmlColumnMapping, XmlDataType, XmlTableField, XmlTableMetadata,
};

fn map_info(id: u32, name: &str, root: &str) -> MapInfo {
    let schema_id = SchemaId::new(format!("schema-{id}")).unwrap();
    let schema_xml = format!(
        "<x:schema xmlns:x=\"http://www.w3.org/2001/XMLSchema\"><x:element name=\"{root}\"/></x:schema>"
    );
    let schema = Schema::try_new(
        schema_id.clone(),
        OpaqueXml::try_new(schema_xml.into_bytes()).unwrap(),
    )
    .unwrap();
    let map = Map::try_new(
        MapId::new(id).unwrap(),
        name,
        root,
        schema_id,
        false,
        true,
        false,
        true,
        true,
    )
    .unwrap();
    MapInfo::try_new("", vec![schema], vec![map]).unwrap()
}

fn external_table() -> ListObject {
    let columns = vec![
        ListObjectColumn::try_new(ListColumnId::try_new(1).unwrap(), "City").unwrap(),
        ListObjectColumn::try_new(ListColumnId::try_new(2).unwrap(), "Population").unwrap(),
    ];
    let metadata = ExternalTableMetadata::try_new(vec![
        ExternalTableField::try_new(columns[0].id(), "CITY_NAME", 41).unwrap(),
        ExternalTableField::try_new(columns[1].id(), "POPULATION", 42)
            .unwrap()
            .with_aggregate_format_bytes(vec![0x11, 0x22, 0x33])
            .unwrap(),
    ])
    .unwrap()
    .with_version(ExternalTableVersion::Excel2007)
    .with_build_number(1234);
    ListObject::try_new(
        ListObjectId::try_new(17).unwrap(),
        "ExternalCities",
        ListObjectRange::try_new(0, 4, 0, 1).unwrap(),
        columns,
        ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_external_data(metadata)
    .unwrap()
}

#[test]
fn feature12_external_metadata_round_trips_inertly() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Query results").unwrap();
    writer.add_list_object(sheet, external_table()).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let first = output.into_inner();
    let workbook = Workbook::new(Cursor::new(first.clone())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(table.feature_version(), ListObjectFeatureVersion::Feature12);
    let metadata = table.external_metadata().unwrap();
    assert_eq!(metadata.version(), ExternalTableVersion::Excel2007);
    assert_eq!(metadata.build_number(), 1234);
    assert_eq!(metadata.fields()[0].source_name(), "CITY_NAME");
    assert_eq!(metadata.fields()[1].query_field_id(), 42);
    assert_eq!(
        metadata.fields()[1].aggregate_format_bytes(),
        &[0x11, 0x22, 0x33]
    );
    assert!(
        table.opaque_feature().is_some(),
        "producer fragments remain available for lossless output"
    );
}

#[test]
fn external_metadata_rejects_duplicate_and_mismatched_ownership() {
    let duplicate = ExternalTableMetadata::try_new(vec![
        ExternalTableField::try_new(ListColumnId::try_new(1).unwrap(), "A", 9).unwrap(),
        ExternalTableField::try_new(ListColumnId::try_new(2).unwrap(), "a", 9).unwrap(),
    ]);
    assert!(duplicate.is_err());
    let mismatched = ExternalTableMetadata::try_new(vec![
        ExternalTableField::try_new(ListColumnId::try_new(2).unwrap(), "A", 9).unwrap(),
        ExternalTableField::try_new(ListColumnId::try_new(1).unwrap(), "B", 10).unwrap(),
    ])
    .unwrap();
    let columns = vec![
        ListObjectColumn::try_new(ListColumnId::try_new(1).unwrap(), "A").unwrap(),
        ListObjectColumn::try_new(ListColumnId::try_new(2).unwrap(), "B").unwrap(),
    ];
    assert!(
        ListObject::try_new(
            ListObjectId::try_new(2).unwrap(),
            "Ownership",
            ListObjectRange::try_new(0, 2, 0, 1).unwrap(),
            columns,
            ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
        )
        .unwrap()
        .with_external_data(mismatched)
        .is_err()
    );
}

#[test]
fn external_substructure_bounds_are_enforced() {
    let field = ExternalTableField::try_new(ListColumnId::try_new(1).unwrap(), "A", 1).unwrap();
    assert!(field.clone().with_auto_filter_bytes(vec![0; 5]).is_err());
    let mut oversized = vec![0; 6 + 2081];
    oversized[..4].copy_from_slice(&2081u32.to_le_bytes());
    assert!(field.clone().with_auto_filter_bytes(oversized).is_err());
    assert!(field.with_header_cache_bytes(vec![8, 0, 0, 0]).is_err());
}

#[test]
fn feature11_web_source_round_trips_typed_inert_sync_metadata() {
    let columns = vec![
        ListObjectColumn::try_new(ListColumnId::try_new(1).unwrap(), "Title").unwrap(),
        ListObjectColumn::try_new(ListColumnId::try_new(2).unwrap(), "Active").unwrap(),
    ];
    let fields = vec![
        WebTableField::try_new(
            columns[0].id(),
            "Title",
            WebColumnType::Text,
            WebFieldInfo::new(0x409)
                .with_default_value(WebDefaultValue::String("draft".into()))
                .with_validation_formula("LEN([Title])>0")
                .unwrap(),
        )
        .unwrap()
        .with_calculated_formula_tokens(vec![0x1e, 1, 0])
        .unwrap(),
        WebTableField::try_new(
            columns[1].id(),
            "Active",
            WebColumnType::Boolean,
            WebFieldInfo::new(0x409)
                .with_default_value(WebDefaultValue::Boolean(true))
                .with_required(true),
        )
        .unwrap(),
    ];
    let metadata = WebTableMetadata::try_new(fields)
        .unwrap()
        .with_provider_name("Provider")
        .unwrap()
        .with_entry_id("web-17")
        .unwrap()
        .with_deleted_row_ids(vec![100, 101])
        .unwrap()
        .with_changed_row_ids(vec![102])
        .unwrap()
        .with_invalid_cells(vec![WebInvalidCell::new(102, columns[1].id())])
        .unwrap();
    let table = ListObject::try_new(
        ListObjectId::try_new(17).unwrap(),
        "WebTasks",
        ListObjectRange::try_new(0, 4, 0, 1).unwrap(),
        columns,
        ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_web_source(metadata)
    .unwrap();
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Web").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(table.feature_version(), ListObjectFeatureVersion::Feature11);
    let ListObjectSourceMetadata::Web(metadata) = table.source_metadata().unwrap() else {
        panic!("expected Web source")
    };
    assert_eq!(metadata.fields()[0].data_type(), WebColumnType::Text);
    assert_eq!(
        metadata.fields()[0].calculated_formula_tokens(),
        Some(&[0x1e, 1, 0][..])
    );
    assert_eq!(
        metadata.fields()[1].info().default_value(),
        Some(&WebDefaultValue::Boolean(true))
    );
    assert_eq!(metadata.deleted_row_ids(), [100, 101]);
    assert_eq!(metadata.changed_row_ids(), [102]);
    assert_eq!(
        metadata.invalid_cells(),
        [WebInvalidCell::new(102, ListColumnId::try_new(2).unwrap())]
    );
    assert!(table.opaque_feature().is_none());
}

#[test]
fn feature11_xml_source_round_trips_typed_mapping() {
    let columns =
        vec![ListObjectColumn::try_new(ListColumnId::try_new(1).unwrap(), "Name").unwrap()];
    let field = XmlTableField::try_new(
        columns[0].id(),
        "name",
        XmlDataType::try_new(0x2125).unwrap(),
    )
    .unwrap()
    .with_mapping(XmlColumnMapping::try_new(9, "/root/item/name", true).unwrap());
    let metadata = XmlTableMetadata::try_new(vec![field])
        .unwrap()
        .with_entry_id("23")
        .unwrap();
    let table = ListObject::try_new(
        ListObjectId::try_new(23).unwrap(),
        "XmlNames",
        ListObjectRange::try_new(0, 3, 0, 0).unwrap(),
        columns,
        ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_xml_source(metadata)
    .unwrap();
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("XML").unwrap();
    assert!(writer.add_list_object(sheet, table.clone()).is_err());
    assert!(writer.xml_map().is_none());
    assert!(
        writer
            .put_xml_map(map_info(9, "Names", "root"))
            .unwrap()
            .is_none()
    );
    writer.add_list_object(sheet, table).unwrap();
    assert!(writer.put_xml_map(map_info(10, "Other", "other")).is_err());
    assert_eq!(writer.xml_map().unwrap().maps()[0].id().get(), 9);
    assert!(writer.remove_xml_map().is_err());
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let mut package = litchi_cfb::OleFile::open(Cursor::new(bytes.clone())).unwrap();
    let xml_streams = package
        .list_streams()
        .into_iter()
        .filter(|path| path.as_slice() == ["XML"])
        .collect::<Vec<_>>();
    assert_eq!(xml_streams.len(), 1);
    let xml = package.open_stream(&["XML"]).unwrap();
    assert!(!xml.iter().any(|byte| matches!(byte, b'\n' | b'\r' | b'\t')));
    assert!(!xml.windows(2).any(|bytes| bytes == b"> "));
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.xml_map().unwrap().maps()[0].id().get(), 9);
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    let ListObjectSourceMetadata::Xml(metadata) = table.source_metadata().unwrap() else {
        panic!("expected XML source")
    };
    let mapping = metadata.fields()[0].mapping().unwrap();
    assert_eq!(metadata.entry_id(), Some("23"));
    assert_eq!(metadata.fields()[0].data_type().value(), 0x2125);
    assert_eq!(mapping.map_id(), 9);
    assert_eq!(mapping.xpath(), "/root/item/name");
    assert!(mapping.can_be_single());
    assert!(table.opaque_feature().is_none());
}

#[test]
fn feature12_web_source_round_trips_typed_with_loaded_total_formula() {
    let columns = vec![
        ListObjectColumn::try_new(ListColumnId::try_new(1).unwrap(), "Title")
            .unwrap()
            .with_total_formula_tokens(vec![0x1e, 1, 0])
            .unwrap(),
    ];
    let fields = vec![
        WebTableField::try_new(
            columns[0].id(),
            "Title",
            WebColumnType::Text,
            WebFieldInfo::new(0x409),
        )
        .unwrap(),
    ];
    let metadata = WebTableMetadata::try_new(fields)
        .unwrap()
        .with_entry_id("web-feature12")
        .unwrap();
    let table = ListObject::try_new(
        ListObjectId::try_new(41).unwrap(),
        "WebFeature12",
        ListObjectRange::try_new(0, 3, 0, 0).unwrap(),
        columns,
        ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_totals_row(true)
    .unwrap()
    .with_web_source(metadata)
    .unwrap();
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Web12").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(table.feature_version(), ListObjectFeatureVersion::Feature12);
    let ListObjectSourceMetadata::Web(metadata) = table.source_metadata().unwrap() else {
        panic!("expected Web source");
    };
    assert_eq!(metadata.fields()[0].data_type(), WebColumnType::Text);
    assert_eq!(
        table.columns()[0].total_formula_tokens(),
        Some(&[0x1e, 1, 0][..])
    );
    assert!(table.opaque_feature().is_none());
}

#[test]
fn feature12_xml_source_round_trips_typed_with_loaded_total_formula() {
    let columns = vec![
        ListObjectColumn::try_new(ListColumnId::try_new(1).unwrap(), "Name")
            .unwrap()
            .with_total_formula_tokens(vec![0x1e, 1, 0])
            .unwrap(),
    ];
    let field = XmlTableField::try_new(
        columns[0].id(),
        "name",
        XmlDataType::try_new(0x2125).unwrap(),
    )
    .unwrap()
    .with_mapping(XmlColumnMapping::try_new(12, "/root/name", true).unwrap());
    let metadata = XmlTableMetadata::try_new(vec![field])
        .unwrap()
        .with_entry_id("42")
        .unwrap();
    let table = ListObject::try_new(
        ListObjectId::try_new(42).unwrap(),
        "XmlFeature12",
        ListObjectRange::try_new(0, 3, 0, 0).unwrap(),
        columns,
        ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_totals_row(true)
    .unwrap()
    .with_xml_source(metadata)
    .unwrap();
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Xml12").unwrap();
    writer
        .put_xml_map(map_info(12, "Feature12", "root"))
        .unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(table.feature_version(), ListObjectFeatureVersion::Feature12);
    let ListObjectSourceMetadata::Xml(metadata) = table.source_metadata().unwrap() else {
        panic!("expected XML source");
    };
    assert_eq!(metadata.entry_id(), Some("42"));
    assert_eq!(metadata.fields()[0].mapping().unwrap().map_id(), 12);
    assert_eq!(
        table.columns()[0].total_formula_tokens(),
        Some(&[0x1e, 1, 0][..])
    );
    assert!(table.opaque_feature().is_none());
}

#[test]
fn feature11_source_types_reject_mismatched_or_unbounded_metadata() {
    assert!(XmlDataType::try_new(0xdead_beef).is_err());
    assert!(XmlColumnMapping::try_new(1, "x".repeat(32000), false).is_err());
    let info = WebFieldInfo::new(0x409).with_default_value(WebDefaultValue::Boolean(true));
    assert!(
        WebTableField::try_new(
            ListColumnId::try_new(1).unwrap(),
            "Text",
            WebColumnType::Text,
            info
        )
        .is_err()
    );
    let columns =
        vec![ListObjectColumn::try_new(ListColumnId::try_new(1).unwrap(), "Name").unwrap()];
    let metadata = XmlTableMetadata::try_new(vec![
        XmlTableField::try_new(
            columns[0].id(),
            "name",
            XmlDataType::try_new(0x2125).unwrap(),
        )
        .unwrap(),
    ])
    .unwrap()
    .with_entry_id("99")
    .unwrap();
    assert!(
        ListObject::try_new(
            ListObjectId::try_new(23).unwrap(),
            "XmlNames",
            ListObjectRange::try_new(0, 2, 0, 0).unwrap(),
            columns,
            ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap()
        )
        .unwrap()
        .with_xml_source(metadata)
        .is_err()
    );
}
