use std::io::Cursor;

use litchi_ole::xls::writer::XlsWriter;
use litchi_ole::xls::{
    XlsExternalTableField, XlsExternalTableMetadata, XlsExternalTableVersion, XlsListColumnId,
    XlsListObject, XlsListObjectColumn, XlsListObjectFeatureVersion, XlsListObjectId,
    XlsListObjectRange, XlsListObjectSourceMetadata, XlsListObjectStyleOptions, XlsWebColumnType,
    XlsWebDefaultValue, XlsWebFieldInfo, XlsWebInvalidCell, XlsWebTableField, XlsWebTableMetadata,
    XlsWorkbook, XlsXmlColumnMapping, XlsXmlDataType, XlsXmlTableField, XlsXmlTableMetadata,
};

fn external_table() -> XlsListObject {
    let columns = vec![
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "City").unwrap(),
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(2).unwrap(), "Population").unwrap(),
    ];
    let metadata = XlsExternalTableMetadata::try_new(vec![
        XlsExternalTableField::try_new(columns[0].id(), "CITY_NAME", 41).unwrap(),
        XlsExternalTableField::try_new(columns[1].id(), "POPULATION", 42)
            .unwrap()
            .with_aggregate_format_bytes(vec![0x11, 0x22, 0x33])
            .unwrap(),
    ])
    .unwrap()
    .with_version(XlsExternalTableVersion::Excel2007)
    .with_build_number(1234);
    XlsListObject::try_new(
        XlsListObjectId::try_new(17).unwrap(),
        "ExternalCities",
        XlsListObjectRange::try_new(0, 4, 0, 1).unwrap(),
        columns,
        XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_external_data(metadata)
    .unwrap()
}

#[test]
fn feature12_external_metadata_round_trips_inertly() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Query results").unwrap();
    writer.add_list_object(sheet, external_table()).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let first = output.into_inner();
    let workbook = XlsWorkbook::new(Cursor::new(first.clone())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(
        table.feature_version(),
        XlsListObjectFeatureVersion::Feature12
    );
    let metadata = table.external_metadata().unwrap();
    assert_eq!(metadata.version(), XlsExternalTableVersion::Excel2007);
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
    let duplicate = XlsExternalTableMetadata::try_new(vec![
        XlsExternalTableField::try_new(XlsListColumnId::try_new(1).unwrap(), "A", 9).unwrap(),
        XlsExternalTableField::try_new(XlsListColumnId::try_new(2).unwrap(), "a", 9).unwrap(),
    ]);
    assert!(duplicate.is_err());
    let mismatched = XlsExternalTableMetadata::try_new(vec![
        XlsExternalTableField::try_new(XlsListColumnId::try_new(2).unwrap(), "A", 9).unwrap(),
        XlsExternalTableField::try_new(XlsListColumnId::try_new(1).unwrap(), "B", 10).unwrap(),
    ])
    .unwrap();
    let columns = vec![
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "A").unwrap(),
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(2).unwrap(), "B").unwrap(),
    ];
    assert!(
        XlsListObject::try_new(
            XlsListObjectId::try_new(2).unwrap(),
            "Ownership",
            XlsListObjectRange::try_new(0, 2, 0, 1).unwrap(),
            columns,
            XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
        )
        .unwrap()
        .with_external_data(mismatched)
        .is_err()
    );
}

#[test]
fn external_substructure_bounds_are_enforced() {
    let field =
        XlsExternalTableField::try_new(XlsListColumnId::try_new(1).unwrap(), "A", 1).unwrap();
    assert!(field.clone().with_auto_filter_bytes(vec![0; 5]).is_err());
    let mut oversized = vec![0; 6 + 2081];
    oversized[..4].copy_from_slice(&2081u32.to_le_bytes());
    assert!(field.clone().with_auto_filter_bytes(oversized).is_err());
    assert!(field.with_header_cache_bytes(vec![8, 0, 0, 0]).is_err());
}

#[test]
fn feature11_web_source_round_trips_typed_inert_sync_metadata() {
    let columns = vec![
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "Title").unwrap(),
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(2).unwrap(), "Active").unwrap(),
    ];
    let fields = vec![
        XlsWebTableField::try_new(
            columns[0].id(),
            "Title",
            XlsWebColumnType::Text,
            XlsWebFieldInfo::new(0x409)
                .with_default_value(XlsWebDefaultValue::String("draft".into()))
                .with_validation_formula("LEN([Title])>0")
                .unwrap(),
        )
        .unwrap()
        .with_calculated_formula_tokens(vec![0x1e, 1, 0])
        .unwrap(),
        XlsWebTableField::try_new(
            columns[1].id(),
            "Active",
            XlsWebColumnType::Boolean,
            XlsWebFieldInfo::new(0x409)
                .with_default_value(XlsWebDefaultValue::Boolean(true))
                .with_required(true),
        )
        .unwrap(),
    ];
    let metadata = XlsWebTableMetadata::try_new(fields)
        .unwrap()
        .with_provider_name("Provider")
        .unwrap()
        .with_entry_id("web-17")
        .unwrap()
        .with_deleted_row_ids(vec![100, 101])
        .unwrap()
        .with_changed_row_ids(vec![102])
        .unwrap()
        .with_invalid_cells(vec![XlsWebInvalidCell::new(102, columns[1].id())])
        .unwrap();
    let table = XlsListObject::try_new(
        XlsListObjectId::try_new(17).unwrap(),
        "WebTasks",
        XlsListObjectRange::try_new(0, 4, 0, 1).unwrap(),
        columns,
        XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_web_source(metadata)
    .unwrap();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Web").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(
        table.feature_version(),
        XlsListObjectFeatureVersion::Feature11
    );
    let XlsListObjectSourceMetadata::Web(metadata) = table.source_metadata().unwrap() else {
        panic!("expected Web source")
    };
    assert_eq!(metadata.fields()[0].data_type(), XlsWebColumnType::Text);
    assert_eq!(
        metadata.fields()[0].calculated_formula_tokens(),
        Some(&[0x1e, 1, 0][..])
    );
    assert_eq!(
        metadata.fields()[1].info().default_value(),
        Some(&XlsWebDefaultValue::Boolean(true))
    );
    assert_eq!(metadata.deleted_row_ids(), [100, 101]);
    assert_eq!(metadata.changed_row_ids(), [102]);
    assert_eq!(
        metadata.invalid_cells(),
        [XlsWebInvalidCell::new(
            102,
            XlsListColumnId::try_new(2).unwrap()
        )]
    );
    assert!(table.opaque_feature().is_none());
}

#[test]
fn feature11_xml_source_round_trips_typed_mapping() {
    let columns =
        vec![XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "Name").unwrap()];
    let field = XlsXmlTableField::try_new(
        columns[0].id(),
        "name",
        XlsXmlDataType::try_new(0x2125).unwrap(),
    )
    .unwrap()
    .with_mapping(XlsXmlColumnMapping::try_new(9, "/root/item/name", true).unwrap());
    let metadata = XlsXmlTableMetadata::try_new(vec![field])
        .unwrap()
        .with_entry_id("23")
        .unwrap();
    let table = XlsListObject::try_new(
        XlsListObjectId::try_new(23).unwrap(),
        "XmlNames",
        XlsListObjectRange::try_new(0, 3, 0, 0).unwrap(),
        columns,
        XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_xml_source(metadata)
    .unwrap();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("XML").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    let XlsListObjectSourceMetadata::Xml(metadata) = table.source_metadata().unwrap() else {
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
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "Title")
            .unwrap()
            .with_total_formula_tokens(vec![0x1e, 1, 0])
            .unwrap(),
    ];
    let fields = vec![
        XlsWebTableField::try_new(
            columns[0].id(),
            "Title",
            XlsWebColumnType::Text,
            XlsWebFieldInfo::new(0x409),
        )
        .unwrap(),
    ];
    let metadata = XlsWebTableMetadata::try_new(fields)
        .unwrap()
        .with_entry_id("web-feature12")
        .unwrap();
    let table = XlsListObject::try_new(
        XlsListObjectId::try_new(41).unwrap(),
        "WebFeature12",
        XlsListObjectRange::try_new(0, 3, 0, 0).unwrap(),
        columns,
        XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_totals_row(true)
    .unwrap()
    .with_web_source(metadata)
    .unwrap();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Web12").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(
        table.feature_version(),
        XlsListObjectFeatureVersion::Feature12
    );
    let XlsListObjectSourceMetadata::Web(metadata) = table.source_metadata().unwrap() else {
        panic!("expected Web source");
    };
    assert_eq!(metadata.fields()[0].data_type(), XlsWebColumnType::Text);
    assert_eq!(
        table.columns()[0].total_formula_tokens(),
        Some(&[0x1e, 1, 0][..])
    );
    assert!(table.opaque_feature().is_none());
}

#[test]
fn feature12_xml_source_round_trips_typed_with_loaded_total_formula() {
    let columns = vec![
        XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "Name")
            .unwrap()
            .with_total_formula_tokens(vec![0x1e, 1, 0])
            .unwrap(),
    ];
    let field = XlsXmlTableField::try_new(
        columns[0].id(),
        "name",
        XlsXmlDataType::try_new(0x2125).unwrap(),
    )
    .unwrap()
    .with_mapping(XlsXmlColumnMapping::try_new(12, "/root/name", true).unwrap());
    let metadata = XlsXmlTableMetadata::try_new(vec![field])
        .unwrap()
        .with_entry_id("42")
        .unwrap();
    let table = XlsListObject::try_new(
        XlsListObjectId::try_new(42).unwrap(),
        "XmlFeature12",
        XlsListObjectRange::try_new(0, 3, 0, 0).unwrap(),
        columns,
        XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_totals_row(true)
    .unwrap()
    .with_xml_source(metadata)
    .unwrap();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Xml12").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(
        table.feature_version(),
        XlsListObjectFeatureVersion::Feature12
    );
    let XlsListObjectSourceMetadata::Xml(metadata) = table.source_metadata().unwrap() else {
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
    assert!(XlsXmlDataType::try_new(0xdead_beef).is_err());
    assert!(XlsXmlColumnMapping::try_new(1, "x".repeat(32000), false).is_err());
    let info = XlsWebFieldInfo::new(0x409).with_default_value(XlsWebDefaultValue::Boolean(true));
    assert!(
        XlsWebTableField::try_new(
            XlsListColumnId::try_new(1).unwrap(),
            "Text",
            XlsWebColumnType::Text,
            info
        )
        .is_err()
    );
    let columns =
        vec![XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "Name").unwrap()];
    let metadata = XlsXmlTableMetadata::try_new(vec![
        XlsXmlTableField::try_new(
            columns[0].id(),
            "name",
            XlsXmlDataType::try_new(0x2125).unwrap(),
        )
        .unwrap(),
    ])
    .unwrap()
    .with_entry_id("99")
    .unwrap();
    assert!(
        XlsListObject::try_new(
            XlsListObjectId::try_new(23).unwrap(),
            "XmlNames",
            XlsListObjectRange::try_new(0, 2, 0, 0).unwrap(),
            columns,
            XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap()
        )
        .unwrap()
        .with_xml_source(metadata)
        .is_err()
    );
}
