use super::codec::{append_frt, parse_string};
use super::*;
use crate::Result;

#[test]
fn table_flags_preserve_known_and_future_bits() {
    let raw = 0xA1FF_FFFE;
    let flags = TableFlags::from_raw(raw);
    assert_eq!(flags.raw(), raw);
    assert!(flags.auto_filter());
    assert!(flags.persists_auto_filter());
    assert!(flags.shows_insert_row());
    assert!(flags.insert_row_inserts_cells());
    assert!(flags.loads_deleted_row_ids());
    assert!(flags.shows_total_row());
    assert!(flags.needs_commit());
    assert!(flags.is_single_cell());
    assert!(flags.applies_auto_filter());
    assert!(flags.forces_insert_row_visible());
    assert!(flags.uses_compressed_xml());
    assert!(flags.loads_provider_name());
    assert!(flags.loads_changed_row_ids());
    assert_eq!(flags.version_nibble(), 0xF);
    assert!(flags.loads_entry_id());
    assert!(flags.loads_invalid_cells());
    assert!(flags.has_good_build());
    assert!(flags.is_published());
    assert_ne!(flags.unknown_bits(), 0);
}

#[test]
fn table_feature_flags_are_typed_and_round_trip_through_feature11() {
    let value = table(2, 3)
        .with_table_flags(
            TableFlags::default_table()
                .with_show_insert_row(true)
                .with_unknown_bits(0x8000_0000),
        )
        .unwrap();
    let records = value.to_feature_record_bytes().unwrap();
    let parsed = parse_feature_records(&value, &records).unwrap();
    assert!(parsed.table_flags().shows_insert_row());
    assert_eq!(parsed.table_flags().unknown_bits(), 0x8000_0000);
    assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);
}

fn table(column_count: usize, name_len: usize) -> ListObject {
    let columns = (0..column_count)
        .map(|index| {
            ListObjectColumn::try_new(
                ListColumnId::try_new(index as u32 + 1).unwrap(),
                format!("C{index}_{}", "x".repeat(name_len)),
            )
            .unwrap()
        })
        .collect();
    ListObject::try_new(
        ListObjectId::try_new(1).unwrap(),
        "TableOne",
        ListObjectRange::try_new(0, 2, 0, column_count as u16 - 1).unwrap(),
        columns,
        ListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
}

fn payload(record: &[u8]) -> &[u8] {
    &record[4..]
}

fn parse_feature_records(table: &ListObject, records: &[Vec<u8>]) -> Result<ListObject> {
    let mut collector = ListObjectCollector::new();
    let header = feature_header_record(std::slice::from_ref(table))?;
    collector.feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))?;
    for record in records {
        let record_type = u16::from_le_bytes(record[..2].try_into().unwrap());
        collector.feed_record(record_type, payload(record))?;
    }
    for record in table.to_list12_record_bytes()? {
        collector.feed_record(LIST12_RECORD_TYPE, payload(&record))?;
    }
    Ok(collector.finish()?.remove(0))
}

#[test]
fn continue_frt11_rejects_orphans_bad_echoes_and_short_predecessors() {
    let continuation = {
        let mut value = Vec::new();
        append_frt(&mut value, CONTINUE_FRT11_RECORD_TYPE, None);
        value
    };
    assert!(
        ListObjectCollector::new()
            .feed_record(CONTINUE_FRT11_RECORD_TYPE, &continuation)
            .is_err()
    );

    let short = table(2, 3);
    let mut collector = ListObjectCollector::new();
    let header = feature_header_record(std::slice::from_ref(&short)).unwrap();
    collector
        .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
        .unwrap();
    let feature = short.to_feature_record_bytes().unwrap();
    collector
        .feed_record(FEATURE11_RECORD_TYPE, payload(&feature[0]))
        .unwrap();
    assert!(
        collector
            .feed_record(CONTINUE_FRT11_RECORD_TYPE, &continuation)
            .is_err()
    );

    let long = table(256, 220);
    let mut records = long.to_feature_record_bytes().unwrap();
    assert!(records.len() > 2);
    records[1][4] = 0;
    let mut collector = ListObjectCollector::new();
    let header = feature_header_record(std::slice::from_ref(&long)).unwrap();
    collector
        .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
        .unwrap();
    collector
        .feed_record(FEATURE11_RECORD_TYPE, payload(&records[0]))
        .unwrap();
    assert!(
        collector
            .feed_record(CONTINUE_FRT11_RECORD_TYPE, payload(&records[1]))
            .is_err()
    );
}

#[test]
fn unsupported_feature12_bytes_are_retained_and_autofilter12_chain_is_strict() {
    let value = table(2, 3).with_header_row(false).unwrap();
    let header = feature_header_record(std::slice::from_ref(&value)).unwrap();
    let mut feature = value.to_feature_record_bytes().unwrap().remove(0);
    feature[4 + 35..4 + 39].copy_from_slice(&4u32.to_le_bytes());
    let list12 = value.to_list12_record_bytes().unwrap();
    let mut collector = ListObjectCollector::new();
    collector
        .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
        .unwrap();
    collector
        .feed_record(FEATURE12_RECORD_TYPE, payload(&feature))
        .unwrap();
    collector
        .feed_record(LIST12_RECORD_TYPE, payload(&list12[0]))
        .unwrap();
    let mut autofilter = Vec::new();
    append_frt(&mut autofilter, AUTO_FILTER12_RECORD_TYPE, None);
    collector
        .feed_record(AUTO_FILTER12_RECORD_TYPE, &autofilter)
        .unwrap();
    let mut continuation = Vec::new();
    append_frt(
        &mut continuation,
        crate::sort_data::CONTINUE_FRT12_RECORD_TYPE,
        None,
    );
    assert!(
        collector
            .feed_record(crate::sort_data::CONTINUE_FRT12_RECORD_TYPE, &continuation)
            .is_err()
    );

    let mut collector = ListObjectCollector::new();
    collector
        .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))
        .unwrap();
    collector
        .feed_record(FEATURE12_RECORD_TYPE, payload(&feature))
        .unwrap();
    for record in list12 {
        collector
            .feed_record(LIST12_RECORD_TYPE, payload(&record))
            .unwrap();
    }
    let parsed = collector.finish().unwrap().remove(0);
    assert!(parsed.opaque_feature().is_some());
    assert_eq!(parsed.to_feature_record_bytes().unwrap()[0], feature);
}

#[test]
fn external_feature12_is_lossless_and_hostile_versions_cardinality_and_feature11_are_rejected() {
    let base_table = table(2, 3);
    let metadata = ExternalTableMetadata::try_new(vec![
        ExternalTableField::try_new(base_table.columns[0].id, "SOURCE_A", 41).unwrap(),
        ExternalTableField::try_new(base_table.columns[1].id, "SOURCE_B", 42).unwrap(),
    ])
    .unwrap();
    let value = base_table.with_external_data(metadata).unwrap();
    let header = feature_header_record(std::slice::from_ref(&value)).unwrap();
    let feature = value.to_feature_record_bytes().unwrap();
    let list12 = value.to_list12_record_bytes().unwrap();
    let parse = |records: &[Vec<u8>]| -> Result<Vec<ListObject>> {
        let mut collector = ListObjectCollector::new();
        collector.feed_record(FEAT_HDR11_RECORD_TYPE, payload(&header))?;
        collector.feed_record(FEATURE12_RECORD_TYPE, payload(&records[0]))?;
        for continuation in &records[1..] {
            collector.feed_record(CONTINUE_FRT11_RECORD_TYPE, payload(continuation))?;
        }
        for record in &list12 {
            collector.feed_record(LIST12_RECORD_TYPE, payload(record))?;
        }
        collector.finish()
    };
    let parsed = parse(&feature).unwrap().remove(0);
    assert_eq!(parsed.to_feature_record_bytes().unwrap(), feature);
    assert_eq!(
        parsed.external_metadata().unwrap().fields()[1].query_field_id(),
        42
    );

    let mut bad_version = feature.clone();
    let flags = u32::from_le_bytes(bad_version[0][4 + 63..4 + 67].try_into().unwrap());
    bad_version[0][4 + 63..4 + 67]
        .copy_from_slice(&((flags & !0x000F_0000) | 0x000A_0000).to_le_bytes());
    assert!(parse(&bad_version).is_err());

    let mut bad_count = feature.clone();
    let (_, count_offset) =
        parse_string(payload(&bad_count[0]), 99, FEATURE12_RECORD_TYPE, "rgbName").unwrap();
    bad_count[0][4 + count_offset..4 + count_offset + 2].copy_from_slice(&3u16.to_le_bytes());
    assert!(parse(&bad_count).is_err());

    let ordinary = table(2, 3);
    let ordinary_header = feature_header_record(std::slice::from_ref(&ordinary)).unwrap();
    let mut feature11 = ordinary.to_feature_record_bytes().unwrap().remove(0);
    feature11[4 + 35..4 + 39].copy_from_slice(&3u32.to_le_bytes());
    let mut collector = ListObjectCollector::new();
    collector
        .feed_record(FEAT_HDR11_RECORD_TYPE, payload(&ordinary_header))
        .unwrap();
    collector
        .feed_record(FEATURE11_RECORD_TYPE, payload(&feature11))
        .unwrap();
    assert!(collector.finish().is_err());
}

#[test]
fn feature11_web_lfdt_values_and_defaults_round_trip_strictly() {
    let base = table(WebColumnType::ALL.len(), 1);
    let fields =
        WebColumnType::ALL
            .iter()
            .copied()
            .enumerate()
            .map(|(index, kind)| {
                let info =
                    match kind {
                        WebColumnType::Text
                        | WebColumnType::Choice
                        | WebColumnType::MultipleChoices => WebFieldInfo::new(1033)
                            .with_default_value(WebDefaultValue::String(format!("v{index}"))),
                        WebColumnType::Boolean => WebFieldInfo::new(1033)
                            .with_default_value(WebDefaultValue::Boolean(true)),
                        WebColumnType::Number | WebColumnType::Currency => WebFieldInfo::new(1033)
                            .with_default_value(WebDefaultValue::Number(12.5)),
                        WebColumnType::DateTime => WebFieldInfo::new(1033)
                            .with_default_value(WebDefaultValue::DateTime(45_000.25)),
                        _ => WebFieldInfo::new(1033),
                    };
                WebTableField::try_new(
                    base.columns[index].id,
                    format!("SOURCE_{index}"),
                    kind,
                    info,
                )
                .unwrap()
            })
            .collect();
    let value = base
        .with_web_source(WebTableMetadata::try_new(fields).unwrap())
        .unwrap();
    assert_eq!(value.feature_version(), ListObjectFeatureVersion::Feature11);
    let records = value.to_feature_record_bytes().unwrap();
    let parsed = parse_feature_records(&value, &records).unwrap();
    let ListObjectSourceMetadata::Web(metadata) = parsed.source_metadata().unwrap() else {
        panic!("expected Web metadata")
    };
    assert_eq!(
        metadata
            .fields()
            .iter()
            .map(WebTableField::data_type)
            .collect::<Vec<_>>(),
        WebColumnType::ALL
    );
    assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);

    assert!(
        WebTableField::try_new(
            value.columns[0].id,
            "INVALID_DEFAULT",
            WebColumnType::Note,
            WebFieldInfo::new(0).with_default_value(WebDefaultValue::String("x".to_string())),
        )
        .is_err()
    );
}

#[test]
fn feature11_xml_lfxidt_is_exhaustive_and_preserves_ignored_storage() {
    let base = table(XmlDataType::ALL.len(), 1);
    let fields = XmlDataType::ALL
        .iter()
        .copied()
        .enumerate()
        .map(|(index, kind)| {
            XmlTableField::try_new(base.columns[index].id, format!("XML_{index}"), kind).unwrap()
        })
        .collect();
    let value = base
        .with_xml_source(XmlTableMetadata::try_new(fields).unwrap())
        .unwrap();
    let mut records = value.to_feature_record_bytes().unwrap();
    assert_eq!(records.len(), 1);
    let (_, count_offset) =
        parse_string(payload(&records[0]), 99, FEATURE11_RECORD_TYPE, "rgbName").unwrap();
    let first_field = count_offset + 2;
    records[0][4 + 35 + 26..4 + 35 + 28].copy_from_slice(&0xa55au16.to_le_bytes());
    let flags = u32::from_le_bytes(records[0][4 + 63..4 + 67].try_into().unwrap());
    records[0][4 + 63..4 + 67].copy_from_slice(&(flags | 0x8280_0001).to_le_bytes());
    records[0][4 + 35 + 32] = 0x5a;
    records[0][4 + 35 + 48] = 0xa5;
    let field_flags = u32::from_le_bytes(
        records[0][4 + first_field + 24..4 + first_field + 28]
            .try_into()
            .unwrap(),
    );
    records[0][4 + first_field + 24..4 + first_field + 28]
        .copy_from_slice(&(field_flags | 0x8000_0030).to_le_bytes());

    let parsed = parse_feature_records(&value, &records).unwrap();
    let ListObjectSourceMetadata::Xml(metadata) = parsed.source_metadata().unwrap() else {
        panic!("expected XML metadata")
    };
    assert_eq!(metadata.ignored_fixed_word(), 0xa55a);
    assert_eq!(metadata.ignored_flags(), 0x8280_0001);
    assert_eq!(metadata.ignored_fixed_tail()[0], 0x5a);
    assert_eq!(metadata.ignored_fixed_tail()[16], 0xa5);
    assert_eq!(metadata.fields()[0].ignored_flags(), 0x8000_0030);
    assert_eq!(
        metadata
            .fields()
            .iter()
            .map(XmlTableField::data_type)
            .collect::<Vec<_>>(),
        XmlDataType::ALL
    );
    assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);

    let mut invalid = records.clone();
    invalid[0][4 + first_field + 8..4 + first_field + 12].copy_from_slice(&0x212eu32.to_le_bytes());
    assert!(parse_feature_records(&value, &invalid).is_err());
    let mut wrong_source = records.clone();
    wrong_source[0][4 + first_field + 4..4 + first_field + 8].copy_from_slice(&1u32.to_le_bytes());
    assert!(parse_feature_records(&value, &wrong_source).is_err());
}

#[test]
fn feature11_web_ignored_flags_round_trip_but_reserved_and_source_bits_fail() {
    let base = table(1, 1);
    let field = WebTableField::try_new(
        base.columns[0].id,
        "SOURCE",
        WebColumnType::Text,
        WebFieldInfo::new(1033).with_default_value(WebDefaultValue::String("default".to_string())),
    )
    .unwrap();
    let value = base
        .with_web_source(WebTableMetadata::try_new(vec![field]).unwrap())
        .unwrap();
    let canonical = value.to_feature_record_bytes().unwrap();
    let (_, count_offset) =
        parse_string(payload(&canonical[0]), 99, FEATURE11_RECORD_TYPE, "rgbName").unwrap();
    let field_offset = count_offset + 2;
    let (_, after_source) = parse_string(
        payload(&canonical[0]),
        field_offset + 36,
        FEATURE11_RECORD_TYPE,
        "source",
    )
    .unwrap();
    let (_, after_caption) = parse_string(
        payload(&canonical[0]),
        after_source,
        FEATURE11_RECORD_TYPE,
        "caption",
    )
    .unwrap();
    let web_info_offset = after_caption + 6;

    let mut ignored = canonical.clone();
    let flags = u32::from_le_bytes(ignored[0][4 + 63..4 + 67].try_into().unwrap());
    ignored[0][4 + 63..4 + 67].copy_from_slice(&(flags | 0x8280_0001).to_le_bytes());
    let field_flags = u32::from_le_bytes(
        ignored[0][4 + field_offset + 24..4 + field_offset + 28]
            .try_into()
            .unwrap(),
    );
    ignored[0][4 + field_offset + 24..4 + field_offset + 28]
        .copy_from_slice(&(field_flags | 0x8000_0030).to_le_bytes());
    let display = u32::from_le_bytes(
        ignored[0][4 + web_info_offset + 8..4 + web_info_offset + 12]
            .try_into()
            .unwrap(),
    );
    ignored[0][4 + web_info_offset + 8..4 + web_info_offset + 12]
        .copy_from_slice(&(display | 0x8000_0000).to_le_bytes());
    let validation = u32::from_le_bytes(
        ignored[0][4 + web_info_offset + 12..4 + web_info_offset + 16]
            .try_into()
            .unwrap(),
    );
    ignored[0][4 + web_info_offset + 12..4 + web_info_offset + 16]
        .copy_from_slice(&(validation | 0x4000_0000).to_le_bytes());
    let parsed = parse_feature_records(&value, &ignored).unwrap();
    assert_eq!(parsed.to_feature_record_bytes().unwrap(), ignored);

    let mut bad_lfdt = canonical.clone();
    bad_lfdt[0][4 + field_offset + 4..4 + field_offset + 8].copy_from_slice(&12u32.to_le_bytes());
    assert!(parse_feature_records(&value, &bad_lfdt).is_err());
    let mut bad_xml_type = canonical.clone();
    bad_xml_type[0][4 + field_offset + 8..4 + field_offset + 12]
        .copy_from_slice(&XmlDataType::DataTypeString.value().to_le_bytes());
    assert!(parse_feature_records(&value, &bad_xml_type).is_err());
    let mut reserved = canonical.clone();
    let field_flags = u32::from_le_bytes(
        reserved[0][4 + field_offset + 24..4 + field_offset + 28]
            .try_into()
            .unwrap(),
    );
    reserved[0][4 + field_offset + 24..4 + field_offset + 28]
        .copy_from_slice(&(field_flags | 0x40).to_le_bytes());
    assert!(parse_feature_records(&value, &reserved).is_err());

    let mut unsupported_feature12 = canonical.clone();
    unsupported_feature12[0][..2].copy_from_slice(&FEATURE12_RECORD_TYPE.to_le_bytes());
    unsupported_feature12[0][4..6].copy_from_slice(&FEATURE12_RECORD_TYPE.to_le_bytes());
    assert!(parse_feature_records(&value, &unsupported_feature12).is_err());
}

#[test]
fn feature12_single_cell_xml_source_round_trips() {
    let mut base = table(1, 1);
    base.range = ListObjectRange::try_new(0, 0, 0, 0).unwrap();
    let base = base.with_header_row(false).unwrap();
    let field =
        XmlTableField::try_new(base.columns[0].id, "single", XmlDataType::DataTypeString).unwrap();
    let metadata = XmlTableMetadata::try_new(vec![field])
        .unwrap()
        .with_single_cell(true)
        .unwrap();
    let value = base.with_xml_source(metadata).unwrap();
    let records = value.to_feature_record_bytes().unwrap();
    assert_eq!(
        u16::from_le_bytes(records[0][..2].try_into().unwrap()),
        FEATURE12_RECORD_TYPE
    );
    let parsed = parse_feature_records(&value, &records).unwrap();
    assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);
}

#[test]
fn cached_disk_header_typed_builder_and_noncanonical_string_round_trip() {
    let mut raw = Vec::new();
    raw.extend_from_slice(&3u32.to_le_bytes());
    raw.extend_from_slice(&[0xaa, 0xbb, 0xcc]);
    raw.extend_from_slice(&6u16.to_le_bytes());
    raw.push(1);
    raw.extend("Header".encode_utf16().flat_map(u16::to_le_bytes));
    let base = table(1, 1).with_header_row(false).unwrap();
    let field = ExternalTableField::try_new(base.columns[0].id, "SOURCE", 7)
        .unwrap()
        .with_header_cache_bytes(raw.clone())
        .unwrap();
    assert_eq!(
        field.cached_disk_header().formatting_bytes(),
        &[0xaa, 0xbb, 0xcc]
    );
    assert_eq!(field.cached_disk_header().style_name(), Some("Header"));
    assert_eq!(field.header_cache_bytes(), raw);
    let value = base
        .clone()
        .with_external_data(ExternalTableMetadata::try_new(vec![field]).unwrap())
        .unwrap();
    let records = value.to_feature_record_bytes().unwrap();
    let parsed = parse_feature_records(&value, &records).unwrap();
    let field = &parsed.external_metadata().unwrap().fields()[0];
    assert_eq!(field.header_cache_bytes(), raw);
    assert_eq!(field.cached_disk_header().style_name(), Some("Header"));
    assert_eq!(parsed.to_feature_record_bytes().unwrap(), records);

    let built = CachedDiskHeader::try_new(vec![1, 2])
        .unwrap()
        .with_style_name("BuiltInHeader")
        .unwrap();
    assert_eq!(built.formatting_bytes(), &[1, 2]);
    assert_eq!(built.style_name(), Some("BuiltInHeader"));
    assert_eq!(built.clone().without_style_name().style_name(), None);
}

#[test]
fn cached_disk_header_presence_lengths_and_flags_are_strict() {
    assert!(CachedDiskHeader::try_new(vec![0; MAX_FEATURE_BYTES]).is_err());
    assert!(
        ExternalTableField::try_new(ListColumnId::try_new(1).unwrap(), "SOURCE", 1,)
            .unwrap()
            .with_header_cache_bytes(vec![2, 0, 0, 0, 1])
            .is_err()
    );

    let header = CachedDiskHeader::try_new(vec![1])
        .unwrap()
        .with_style_name("HeaderStyle")
        .unwrap();
    let headered = table(1, 1);
    let field = ExternalTableField::try_new(headered.columns[0].id, "SOURCE", 1)
        .unwrap()
        .with_cached_disk_header(header)
        .unwrap();
    assert!(
        headered
            .with_external_data(ExternalTableMetadata::try_new(vec![field]).unwrap())
            .is_err()
    );

    let base = table(1, 1).with_header_row(false).unwrap();
    let field = ExternalTableField::try_new(base.columns[0].id, "SOURCE", 1)
        .unwrap()
        .with_cached_disk_header(
            CachedDiskHeader::try_new(vec![0x10])
                .unwrap()
                .with_style_name("HeaderStyle")
                .unwrap(),
        )
        .unwrap();
    let value = base
        .clone()
        .with_external_data(ExternalTableMetadata::try_new(vec![field]).unwrap())
        .unwrap();
    let records = value.to_feature_record_bytes().unwrap();
    let (_, count_offset) =
        parse_string(payload(&records[0]), 99, FEATURE12_RECORD_TYPE, "rgbName").unwrap();
    let mut field_offset = count_offset + 2;
    let table_flags = u32::from_le_bytes(records[0][4 + 63..4 + 67].try_into().unwrap());
    if table_flags & 0x0010_0000 != 0 {
        field_offset = parse_string(
            payload(&records[0]),
            field_offset,
            FEATURE12_RECORD_TYPE,
            "entryId",
        )
        .unwrap()
        .1;
    }

    let mut missing_flag = records.clone();
    let flags = u32::from_le_bytes(
        missing_flag[0][4 + field_offset + 24..4 + field_offset + 28]
            .try_into()
            .unwrap(),
    );
    missing_flag[0][4 + field_offset + 24..4 + field_offset + 28]
        .copy_from_slice(&(flags & !0x200).to_le_bytes());
    assert!(parse_feature_records(&value, &missing_flag).is_err());

    let empty_field = ExternalTableField::try_new(base.columns[0].id, "SOURCE", 1).unwrap();
    let empty_value = base
        .with_external_data(ExternalTableMetadata::try_new(vec![empty_field]).unwrap())
        .unwrap();
    let mut spurious_flag = empty_value.to_feature_record_bytes().unwrap();
    let (_, count_offset) = parse_string(
        payload(&spurious_flag[0]),
        99,
        FEATURE12_RECORD_TYPE,
        "rgbName",
    )
    .unwrap();
    let mut field_offset = count_offset + 2;
    let table_flags = u32::from_le_bytes(spurious_flag[0][4 + 63..4 + 67].try_into().unwrap());
    if table_flags & 0x0010_0000 != 0 {
        field_offset = parse_string(
            payload(&spurious_flag[0]),
            field_offset,
            FEATURE12_RECORD_TYPE,
            "entryId",
        )
        .unwrap()
        .1;
    }
    let flags = u32::from_le_bytes(
        spurious_flag[0][4 + field_offset + 24..4 + field_offset + 28]
            .try_into()
            .unwrap(),
    );
    spurious_flag[0][4 + field_offset + 24..4 + field_offset + 28]
        .copy_from_slice(&(flags | 0x200).to_le_bytes());
    assert!(parse_feature_records(&empty_value, &spurious_flag).is_err());
}
