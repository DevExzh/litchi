//! Synthetic Brt-record stream tests for the PivotCache definition parser.

use super::model::*;
use super::parse::parse_pivot_cache_definition;
use crate::xlsb::error::{Error, Result};
use litchi_xlsb::raw::{Error as WireError, Kind, Stage, Writer, kind as rt};

fn wide_string(value: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

fn nullable_wide_string(value: Option<&str>) -> Vec<u8> {
    match value {
        Some(value) => wide_string(value),
        None => u32::MAX.to_le_bytes().to_vec(),
    }
}

fn stream(records: &[(Kind, Vec<u8>)]) -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    for (record_type, payload) in records {
        writer.write_record(*record_type, payload).unwrap();
    }
    data
}

fn parse(records: &[(Kind, Vec<u8>)]) -> Result<PivotCacheDefinition> {
    parse_pivot_cache_definition(&stream(records))
}

/// `BrtBeginPivotCacheDef` payload with refreshed-by and records rel id.
fn definition_payload() -> Vec<u8> {
    let mut data = vec![
        3,           // bVerCacheLastRefresh
        0,           // bVerCacheRefreshableMin
        2,           // bVerCacheCreated
        0b0001_0001, // fSaveData | fEnableRefresh
    ];
    data.extend_from_slice(&(-1i32).to_le_bytes()); // citmGhostMax
    data.extend_from_slice(&44_000.5f64.to_le_bytes()); // xnumRefreshedDate
    data.push(0b0000_0011); // fLoadRefreshedWho | fLoadRelIDRecords
    data.extend_from_slice(&5u32.to_le_bytes()); // cRecords
    data.extend_from_slice(&wide_string(" analyst "));
    data.extend_from_slice(&wide_string("rIdRecords"));
    data
}

fn minimal_field_payload(name: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&0u16.to_le_bytes()); // flags
    data.extend_from_slice(&u32::MAX.to_le_bytes()); // ifmt = default
    data.extend_from_slice(&0u16.to_le_bytes()); // wTypeSql
    data.extend_from_slice(&0u32.to_le_bytes()); // ihdb
    data.extend_from_slice(&0x7FFFu32.to_le_bytes()); // isxtl
    data.extend_from_slice(&0u32.to_le_bytes()); // cIsxtmps
    data.extend_from_slice(&wide_string(name));
    data
}

#[test]
fn parses_worksheet_range_source() {
    let mut source = Vec::new();
    source.extend_from_slice(&0u32.to_le_bytes()); // iSrcType = sheet
    source.extend_from_slice(&0u32.to_le_bytes()); // dwConnID

    let mut range = vec![0x00, 0x00, 0b0000_0010]; // fLoadSheet
    range.extend_from_slice(&wide_string("Data Sheet"));
    for value in [0i32, 99, 1, 7] {
        range.extend_from_slice(&value.to_le_bytes());
    }

    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_SOURCE, source),
        (rt::BEGIN_PCDS_RANGE, range),
        (rt::END_PCDS_RANGE, Vec::new()),
        (rt::END_PCD_SOURCE, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();

    assert_eq!(definition.version_last_refresh, 3);
    assert_eq!(definition.version_created, 2);
    assert!(definition.save_data);
    assert!(definition.enable_refresh);
    assert!(!definition.refresh_on_load);
    assert_eq!(definition.ghost_items_max, -1);
    assert_eq!(definition.refreshed_date_serial, 44_000.5);
    assert_eq!(definition.record_count, 5);
    assert_eq!(definition.refreshed_by.as_deref(), Some(" analyst "));
    assert_eq!(definition.records_rel_id.as_deref(), Some("rIdRecords"));

    let source = definition.source.unwrap();
    assert_eq!(source.source_type, PivotCacheSourceType::Worksheet);
    assert_eq!(source.connection_id, None);
    let worksheet = source.worksheet.unwrap();
    assert_eq!(worksheet.sheet_name.as_deref(), Some("Data Sheet"));
    assert_eq!(worksheet.named_range, None);
    assert_eq!(
        worksheet.range,
        Some(PivotCacheRange {
            first_row: 0,
            last_row: 99,
            first_column: 1,
            last_column: 7,
        })
    );
}

#[test]
fn parses_named_range_and_consolidation_sources() {
    // Named-range worksheet source.
    let mut named = vec![0x01, 0x00, 0x00]; // fName
    named.extend_from_slice(&wide_string("MyRange"));
    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_SOURCE, 0u32.to_le_bytes().repeat(2)),
        (rt::BEGIN_PCDS_RANGE, named),
        (rt::END_PCDS_RANGE, Vec::new()),
        (rt::END_PCD_SOURCE, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();
    let worksheet = definition.source.unwrap().worksheet.unwrap();
    assert_eq!(worksheet.named_range.as_deref(), Some("MyRange"));
    assert_eq!(worksheet.range, None);
    assert_eq!(worksheet.sheet_name, None);

    // Consolidation source with one set and one page.
    let mut set = Vec::new();
    for value in [1u32, u32::MAX, u32::MAX, u32::MAX] {
        set.extend_from_slice(&value.to_le_bytes());
    }
    set.push(0x00); // fName = 0 -> range
    set.push(0x00); // fBuiltIn
    set.push(0b0000_0010); // fLoadSheet
    set.extend_from_slice(&wide_string("Q1"));
    for value in [4i32, 20, 0, 3] {
        set.extend_from_slice(&value.to_le_bytes());
    }

    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_SOURCE, {
            let mut data = Vec::new();
            data.extend_from_slice(&2u32.to_le_bytes()); // iSrcType = consolidation
            data.extend_from_slice(&0u32.to_le_bytes());
            data
        }),
        (rt::BEGIN_PCDS_CONSOL, 1u16.to_le_bytes().to_vec()), // fAutoPage
        (rt::BEGIN_PCDSC_SETS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDSC_SET, set),
        (rt::END_PCDSC_SET, Vec::new()),
        (rt::END_PCDSC_SETS, Vec::new()),
        (rt::BEGIN_PCDSC_PAGES, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDSC_PAGE, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDSCP_ITEM, wide_string("Region1")),
        (rt::END_PCDSCP_ITEM, Vec::new()),
        (rt::END_PCDSC_PAGE, Vec::new()),
        (rt::END_PCDSC_PAGES, Vec::new()),
        (rt::END_PCDS_CONSOL, Vec::new()),
        (rt::END_PCD_SOURCE, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();
    let source = definition.source.unwrap();
    assert_eq!(source.source_type, PivotCacheSourceType::Consolidation);
    let consolidation = source.consolidation.unwrap();
    assert!(consolidation.auto_page);
    assert_eq!(consolidation.sets.len(), 1);
    assert_eq!(consolidation.sets[0].item_indexes[0], 1);
    assert_eq!(consolidation.sets[0].sheet_name.as_deref(), Some("Q1"));
    assert_eq!(
        consolidation.sets[0].range,
        Some(PivotCacheRange {
            first_row: 4,
            last_row: 20,
            first_column: 0,
            last_column: 3,
        })
    );
    assert_eq!(consolidation.pages.len(), 1);
    assert_eq!(consolidation.pages[0].item_names, ["Region1".to_string()]);
}

/// `PCDIAddlInfo` with the given flags, optional caption, and member props.
fn item_info(flags: u16, caption: Option<&str>, member_props: &[i32]) -> Vec<u8> {
    let mut data = flags.to_le_bytes().to_vec();
    if flags & 0b100 != 0 {
        data.extend_from_slice(&nullable_wide_string(caption));
    }
    data.extend_from_slice(&(member_props.len() as u32).to_le_bytes());
    for prop in member_props {
        data.extend_from_slice(&prop.to_le_bytes());
    }
    data
}

#[test]
fn parses_fields_with_all_shared_item_types() {
    let mut atbl = Vec::new();
    // fHasBlankItem | fMixedTypesIgnoringBlanks | fNumField | fNumMinMaxValid
    atbl.extend_from_slice(&0b0000_0001_0111_0000u16.to_le_bytes());
    atbl.extend_from_slice(&8u32.to_le_bytes()); // citems
    atbl.extend_from_slice(&1.5f64.to_le_bytes()); // xnumMin
    atbl.extend_from_slice(&99.5f64.to_le_bytes()); // xnumMax

    let mut number = 42.5f64.to_le_bytes().to_vec();
    number.extend_from_slice(&item_info(0b001, None, &[])); // fGhost
    let mut boolean = vec![0x01];
    boolean.extend_from_slice(&item_info(0, None, &[]));
    let mut error = vec![0x2A]; // #N/A
    error.extend_from_slice(&item_info(0, None, &[]));
    let mut string = wide_string("hello");
    string.extend_from_slice(&item_info(0b100, Some("Greeting"), &[3]));
    let mut datetime = Vec::new();
    datetime.extend_from_slice(&2020u16.to_le_bytes());
    datetime.extend_from_slice(&1u16.to_le_bytes());
    datetime.extend_from_slice(&[2, 3, 4, 5]);
    datetime.extend_from_slice(&item_info(0, None, &[]));

    let mut run = Vec::new();
    run.extend_from_slice(&0x0001u16.to_le_bytes()); // mdSxoper = numbers
    run.extend_from_slice(&2u32.to_le_bytes());
    run.extend_from_slice(&7.0f64.to_le_bytes());
    run.extend_from_slice(&8.0f64.to_le_bytes());

    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_FIELDS, 2u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCD_FIELD, minimal_field_payload("F1")),
        (rt::BEGIN_PCDF_ATBL, atbl),
        (rt::PCDIA_MISSING, item_info(0b011, None, &[])), // ghost + calculated
        (rt::PCDIA_NUMBER, number),
        (rt::PCDIA_BOOLEAN, boolean),
        (rt::PCDIA_ERROR, error),
        (rt::PCDIA_STRING, string),
        (rt::PCDIA_DATETIME, datetime),
        (rt::PCDI_STRING, wide_string("plain")),
        (rt::BEGIN_PCDI_RUN, run),
        (rt::END_PCDI_RUN, Vec::new()),
        (rt::END_PCDF_ATBL, Vec::new()),
        (rt::END_PCD_FIELD, Vec::new()),
        (rt::BEGIN_PCD_FIELD, {
            let mut data = Vec::new();
            data.extend_from_slice(&0b1000u16.to_le_bytes()); // fCaption
            data.extend_from_slice(&14u32.to_le_bytes()); // ifmt = built-in date
            data.extend_from_slice(&0u16.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            data.extend_from_slice(&0x7FFFu32.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            data.extend_from_slice(&wide_string("F2"));
            data.extend_from_slice(&wide_string("Field Two"));
            data
        }),
        (rt::END_PCD_FIELD, Vec::new()),
        (rt::END_PCD_FIELDS, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();

    assert_eq!(definition.fields.len(), 2);
    let field = &definition.fields[0];
    assert_eq!(field.name, "F1");
    assert_eq!(field.number_format, None);
    let stats = field.shared_items.stats.unwrap();
    assert!(stats.has_blank_item);
    assert!(stats.mixed_types_ignoring_blanks);
    assert!(stats.numeric_field);
    assert_eq!(stats.item_count, 8);
    assert_eq!(stats.minimum, Some(1.5));
    assert_eq!(stats.maximum, Some(99.5));

    let items = &field.shared_items.items;
    assert_eq!(items.len(), 9);
    assert_eq!(items[0].value, PivotCacheItemValue::Missing);
    let missing_info = items[0].additional.as_ref().unwrap();
    assert!(missing_info.ghost);
    assert!(missing_info.calculated);
    assert_eq!(items[1].value, PivotCacheItemValue::Number(42.5));
    assert!(items[1].additional.as_ref().unwrap().ghost);
    assert_eq!(items[2].value, PivotCacheItemValue::Boolean(true));
    assert_eq!(
        items[3].value,
        PivotCacheItemValue::Error(PivotCacheErrorCode::NA)
    );
    assert_eq!(items[4].value, PivotCacheItemValue::String("hello".into()));
    let string_info = items[4].additional.as_ref().unwrap();
    assert_eq!(string_info.caption.as_deref(), Some("Greeting"));
    assert_eq!(string_info.member_property_items, [3]);
    assert_eq!(
        items[5].value,
        PivotCacheItemValue::DateTime(PivotCacheDateTime {
            year: 2020,
            month: 1,
            day: 2,
            hour: 3,
            minute: 4,
            second: 5,
        })
    );
    assert_eq!(items[6].value, PivotCacheItemValue::String("plain".into()));
    assert_eq!(items[6].additional, None);
    assert_eq!(items[7].value, PivotCacheItemValue::Number(7.0));
    assert_eq!(items[8].value, PivotCacheItemValue::Number(8.0));

    let second = &definition.fields[1];
    assert_eq!(second.caption.as_deref(), Some("Field Two"));
    assert_eq!(second.number_format, Some(14));
    assert_eq!(second.shared_items.items.len(), 0);
    assert_eq!(second.shared_items.stats, None);
}

#[test]
fn parses_range_and_discrete_grouping() {
    let mut group_range = Vec::new();
    group_range.push(0x04); // iByType = days
    group_range.push(0b0000_0101); // fAutoStart | fDates
    group_range.extend_from_slice(&40_000.0f64.to_le_bytes());
    group_range.extend_from_slice(&44_000.0f64.to_le_bytes());
    group_range.extend_from_slice(&7.0f64.to_le_bytes());

    let mut discrete_item = 3u32.to_le_bytes().to_vec();
    discrete_item.extend_from_slice(&item_info(0, None, &[]));

    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_FIELDS, 2u32.to_le_bytes().to_vec()),
        // Base date field with range grouping items.
        (rt::BEGIN_PCD_FIELD, minimal_field_payload("Date")),
        (rt::BEGIN_PCDF_GROUP, {
            let mut data = Vec::new();
            data.extend_from_slice(&(-1i32).to_le_bytes()); // ifdbParent
            data.extend_from_slice(&(-1i32).to_le_bytes()); // ifdbBase
            data
        }),
        (rt::BEGIN_PCDFG_RANGE, group_range),
        (rt::END_PCDFG_RANGE, Vec::new()),
        (rt::BEGIN_PCDFG_ITEMS, 2u32.to_le_bytes().to_vec()),
        (rt::PCDI_STRING, wide_string("<1/1/2020")),
        (rt::PCDI_STRING, wide_string(">12/31/2020")),
        (rt::END_PCDFG_ITEMS, Vec::new()),
        (rt::END_PCDF_GROUP, Vec::new()),
        (rt::END_PCD_FIELD, Vec::new()),
        // Grouped field with a discrete grouping over the base field.
        (rt::BEGIN_PCD_FIELD, minimal_field_payload("Group1")),
        (rt::BEGIN_PCDF_GROUP, {
            let mut data = Vec::new();
            data.extend_from_slice(&(-1i32).to_le_bytes());
            data.extend_from_slice(&0i32.to_le_bytes()); // ifdbBase = field 0
            data
        }),
        (rt::BEGIN_PCDFG_DISCRETE, 2u32.to_le_bytes().to_vec()),
        (rt::PCDI_INDEX, 0u32.to_le_bytes().to_vec()),
        (rt::PCDI_INDEX, 2u32.to_le_bytes().to_vec()),
        (rt::END_PCDFG_DISCRETE, Vec::new()),
        (rt::BEGIN_PCDFG_ITEMS, 1u32.to_le_bytes().to_vec()),
        (rt::PCDIA_STRING, {
            let mut data = wide_string("Group A");
            data.extend_from_slice(&item_info(0, None, &[]));
            data
        }),
        (rt::END_PCDFG_ITEMS, Vec::new()),
        (rt::END_PCDF_GROUP, Vec::new()),
        (rt::END_PCD_FIELD, Vec::new()),
        (rt::END_PCD_FIELDS, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();

    let range_grouping = definition.fields[0].grouping.as_ref().unwrap();
    assert_eq!(range_grouping.parent_field, None);
    assert_eq!(range_grouping.base_field, None);
    let range = range_grouping.range.unwrap();
    assert_eq!(range.group_by, PivotCacheGroupBy::Days);
    assert!(range.auto_start);
    assert!(!range.auto_end);
    assert!(range.dates);
    assert_eq!(range.start, 40_000.0);
    assert_eq!(range.end, 44_000.0);
    assert_eq!(range.interval, 7.0);
    assert_eq!(range_grouping.items.len(), 2);
    assert_eq!(range_grouping.discrete, None);

    let discrete_grouping = definition.fields[1].grouping.as_ref().unwrap();
    assert_eq!(discrete_grouping.base_field, Some(0));
    assert_eq!(
        discrete_grouping.discrete.as_ref().unwrap().item_indexes,
        [0, 2]
    );
    assert_eq!(discrete_grouping.items.len(), 1);
    assert_eq!(
        discrete_grouping.items[0].value,
        PivotCacheItemValue::String("Group A".into())
    );
}

#[test]
fn parses_olap_hierarchies() {
    let mut hierarchy = Vec::new();
    hierarchy.extend_from_slice(&0b0001_0100u16.to_le_bytes()); // fAttributeHierarchy | fOnlyOneField
    hierarchy.extend_from_slice(&2u32.to_le_bytes()); // cLevels
    hierarchy.extend_from_slice(&(-1i32).to_le_bytes()); // isetParent
    hierarchy.extend_from_slice(&0i32.to_le_bytes()); // iconSet
    hierarchy.push(0b0000_0001); // fLoadDimUnq
    hierarchy.extend_from_slice(&0u16.to_le_bytes()); // wAttributeMemberValueType
    hierarchy.extend_from_slice(&wide_string("[Dim].[Hier]"));
    hierarchy.extend_from_slice(&wide_string("Hier Caption"));
    hierarchy.extend_from_slice(&wide_string("[Dim]"));

    let mut level = vec![0x01]; // fGroupLevel
    level.extend_from_slice(&wide_string("[Dim].[Hier].[GroupLvl]"));
    level.extend_from_slice(&wide_string("Group Level"));

    let mut group = Vec::new();
    group.extend_from_slice(&1i32.to_le_bytes()); // iGrpNum
    group.push(0x01); // fLoadParent
    group.extend_from_slice(&wide_string("Group1"));
    group.extend_from_slice(&wide_string("[Dim].[Hier].[Group1]"));
    group.extend_from_slice(&wide_string("Group One"));
    group.extend_from_slice(&wide_string("[Dim].[Hier].[All]"));

    let mut member = 0u32.to_le_bytes().to_vec(); // fGroup = 0
    member.extend_from_slice(&wide_string("[Dim].[Hier].[Member1]"));

    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_SOURCE, {
            let mut data = Vec::new();
            data.extend_from_slice(&1u32.to_le_bytes()); // iSrcType = external
            data.extend_from_slice(&7u32.to_le_bytes()); // dwConnID
            data
        }),
        (rt::END_PCD_SOURCE, Vec::new()),
        (rt::BEGIN_PCD_HIERARCHIES, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCD_HIERARCHY, hierarchy),
        (rt::BEGIN_PCDH_FIELDS_USAGE, {
            let mut data = Vec::new();
            data.extend_from_slice(&2u32.to_le_bytes());
            data.extend_from_slice(&0i32.to_le_bytes());
            data.extend_from_slice(&(-1i32).to_le_bytes());
            data
        }),
        (rt::END_PCDH_FIELDS_USAGE, Vec::new()),
        (rt::BEGIN_PCDHG_LEVELS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDHG_LEVEL, level),
        (rt::END_PCDHG_LEVEL, Vec::new()),
        (rt::END_PCDHG_LEVELS, Vec::new()),
        (rt::BEGIN_PCDHGL_GROUPS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDHGL_GROUP, group),
        (rt::BEGIN_PCDHGLG_MEMBERS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDHGLG_MEMBER, member),
        (rt::END_PCDHGLG_MEMBER, Vec::new()),
        (rt::END_PCDHGLG_MEMBERS, Vec::new()),
        (rt::END_PCDHGL_GROUP, Vec::new()),
        (rt::END_PCDHGL_GROUPS, Vec::new()),
        (
            rt::PCD_H14, // BrtPCDH14
            {
                let mut data = vec![0; 4]; // FRTBlank
                data.push(0b0000_0101); // fFlattenHierarchies | fHierarchizeDistinct
                data.extend_from_slice(&1u32.to_le_bytes());
                data.extend_from_slice(&(-2i32).to_le_bytes());
                data
            },
        ),
        (rt::END_PCD_HIERARCHY, Vec::new()),
        (rt::END_PCD_HIERARCHIES, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();

    let source = definition.source.unwrap();
    assert_eq!(source.source_type, PivotCacheSourceType::External);
    assert_eq!(source.connection_id, Some(7));

    assert_eq!(definition.hierarchies.len(), 1);
    let hierarchy = &definition.hierarchies[0];
    assert_eq!(hierarchy.unique_name, "[Dim].[Hier]");
    assert_eq!(hierarchy.caption, "Hier Caption");
    assert_eq!(hierarchy.dimension_unique_name.as_deref(), Some("[Dim]"));
    assert!(hierarchy.attribute_hierarchy);
    assert!(hierarchy.only_one_field);
    assert!(!hierarchy.measure);
    assert!(!hierarchy.hidden);
    assert_eq!(hierarchy.level_count, 2);
    assert_eq!(hierarchy.set_parent_index, None);
    assert_eq!(hierarchy.field_usage, [0, -1]);
    assert_eq!(hierarchy.grouping_levels.len(), 1);
    assert!(hierarchy.grouping_levels[0].group_level);
    assert_eq!(hierarchy.grouping_levels[0].caption, "Group Level");
    assert_eq!(hierarchy.grouping_groups.len(), 1);
    let group = &hierarchy.grouping_groups[0];
    assert_eq!(group.group_number, 1);
    assert_eq!(group.name, "Group1");
    assert_eq!(
        group.parent_unique_name.as_deref(),
        Some("[Dim].[Hier].[All]")
    );
    assert_eq!(group.members.len(), 1);
    assert!(!group.members[0].is_group);
    assert_eq!(group.members[0].unique_name, "[Dim].[Hier].[Member1]");
    let ext14 = hierarchy.ext14.as_ref().unwrap();
    assert!(ext14.flatten_hierarchies);
    assert!(ext14.hierarchize_distinct);
    assert_eq!(ext14.hierarchy_indexes, [-2]);
}

#[test]
fn parses_calculated_items_and_members() {
    // Calculated item with a formula, one name/pair, and one rule filter.
    let mut calc_item = 0xFFFF_FFFFu32.to_le_bytes().to_vec(); // reserved = -1
    calc_item.extend_from_slice(&3u32.to_le_bytes()); // cce
    calc_item.extend_from_slice(&[0x1E, 0x05, 0x00]); // PtgInt 5
    calc_item.extend_from_slice(&0u32.to_le_bytes()); // cb

    let mut name = Vec::new();
    name.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // ifdb = -1
    name.push(0x00); // ifn = SUM
    name.push(0x00); // flags
    name.extend_from_slice(&[0, 0]); // padding

    let mut pair = Vec::new();
    pair.push(0x00); // flags
    pair.extend_from_slice(&0u32.to_le_bytes()); // ifield
    pair.extend_from_slice(&3i32.to_le_bytes()); // iitem
    pair.extend_from_slice(&[0, 0, 0]); // padding

    let mut filter = Vec::new();
    filter.extend_from_slice(&0i32.to_le_bytes()); // isxvd
    filter.extend_from_slice(&1u32.to_le_bytes()); // cItems
    filter.extend_from_slice(&[0b0000_0101, 0x00, 0x01]); // itmtypeData|itmtypeSUM + fSelected

    let mut calc_mem = Vec::new();
    calc_mem.extend_from_slice(&0b011u32.to_le_bytes()); // fLoadMemberName | fLoadSourceHier
    calc_mem.extend_from_slice(&5i32.to_le_bytes()); // wSolveOrder
    calc_mem.extend_from_slice(&0u32.to_le_bytes()); // fSet = calculated member
    calc_mem.extend_from_slice(&wide_string("[Measures].[Calc]"));
    calc_mem.extend_from_slice(&wide_string("1+1"));
    calc_mem.extend_from_slice(&wide_string("Calc"));
    calc_mem.extend_from_slice(&wide_string("[Measures]"));

    let mut calc_mem14 = vec![0; 4]; // FRTBlank
    calc_mem14.push(0b0000_0011); // fFlattenHierarchies | fDynamicSet
    calc_mem14.extend_from_slice(&wide_string("Folder"));

    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_CALC_ITEMS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCD_CALC_ITEM, calc_item),
        (rt::BEGIN_P_NAMES, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_P_NAME, name),
        (rt::BEGIN_PN_PAIRS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PN_PAIR, pair),
        (rt::END_PN_PAIR, Vec::new()),
        (rt::END_PN_PAIRS, Vec::new()),
        (rt::END_P_NAME, Vec::new()),
        (rt::END_P_NAMES, Vec::new()),
        (rt::BEGIN_PR_FILTERS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PR_FILTER, filter),
        (rt::BEGIN_PRF_ITEM, 7u32.to_le_bytes().to_vec()),
        (rt::END_PRF_ITEM, Vec::new()),
        (rt::END_PR_FILTER, Vec::new()),
        (rt::END_PR_FILTERS, Vec::new()),
        (rt::END_PCD_CALC_ITEM, Vec::new()),
        (rt::END_PCD_CALC_ITEMS, Vec::new()),
        (rt::BEGIN_PCD_CALC_MEMS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCD_CALC_MEM, calc_mem),
        (rt::BEGIN_PCD_CALC_MEM14, calc_mem14),
        (rt::END_PCD_CALC_MEM14, Vec::new()),
        (rt::END_PCD_CALC_MEM, Vec::new()),
        (rt::END_PCD_CALC_MEMS, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();

    assert_eq!(definition.calculated_items.len(), 1);
    let item = &definition.calculated_items[0];
    assert_eq!(item.formula.tokens, [0x1E, 0x05, 0x00]);
    assert_eq!(item.formula.extra, Vec::<u8>::new());
    assert_eq!(item.names.len(), 1);
    assert_eq!(item.names[0].field_index, u32::MAX);
    assert_eq!(item.names[0].function, PivotNameFunction::Sum);
    assert!(!item.names[0].err_name);
    assert_eq!(item.names[0].pairs.len(), 1);
    assert_eq!(item.names[0].pairs[0].field_index, 0);
    assert_eq!(item.names[0].pairs[0].item_index, 3);
    assert!(!item.names[0].pairs[0].physical);
    assert_eq!(item.filters.len(), 1);
    assert_eq!(item.filters[0].field, 0);
    assert_eq!(item.filters[0].item_types, 0b101);
    assert!(item.filters[0].selected);
    assert_eq!(item.filters[0].items, [7]);

    assert_eq!(definition.calculated_members.len(), 1);
    let member = &definition.calculated_members[0];
    assert_eq!(member.name, "[Measures].[Calc]");
    assert_eq!(member.mdx, "1+1");
    assert_eq!(member.solve_order, 5);
    assert!(!member.is_set);
    assert_eq!(member.member_name.as_deref(), Some("Calc"));
    assert_eq!(member.source_hierarchy.as_deref(), Some("[Measures]"));
    assert_eq!(member.parent_unique, None);
    let ext14 = member.ext14.as_ref().unwrap();
    assert!(ext14.flatten_hierarchies);
    assert!(ext14.dynamic_set);
    assert_eq!(ext14.display_folder, "Folder");
    assert_eq!(ext14.long_mdx, None);
}

#[test]
fn parses_tuple_cache_and_pcd14() {
    let mut set = Vec::new();
    set.extend_from_slice(&u32::MAX.to_le_bytes()); // cTuples unknown
    set.extend_from_slice(&3u32.to_le_bytes()); // iRankMax
    set.extend_from_slice(&1u32.to_le_bytes()); // ssoType = SSOASC
    set.push(0x01); // fQueryFailed
    set.extend_from_slice(&wide_string("{[Dim].[Hier].Members}"));

    let mut pcd14 = vec![0; 4]; // FRTBlank
    pcd14.push(0b0000_0101); // fSlicerData | fSrvSupportSubQueryNonVisual
    pcd14.extend_from_slice(&9i32.to_le_bytes()); // icacheId

    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCDSD_TUPLE_CACHE, Vec::new()),
        (rt::BEGIN_PCDSDTC_ENTRIES, 2u32.to_le_bytes().to_vec()),
        (rt::PCDI_NUMBER, 1.5f64.to_le_bytes().to_vec()),
        (rt::PCDI_STRING, wide_string("entry")),
        (rt::END_PCDSDTC_ENTRIES, Vec::new()),
        (rt::BEGIN_PCDSDTC_QUERIES, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDSDTC_QUERY, wide_string("SELECT ...")),
        (rt::END_PCDSDTC_QUERY, Vec::new()),
        (rt::END_PCDSDTC_QUERIES, Vec::new()),
        (rt::BEGIN_PCDSDTC_SETS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCDSDTC_SET, set),
        (rt::END_PCDSDTC_SET, Vec::new()),
        (rt::END_PCDSDTC_SETS, Vec::new()),
        (rt::END_PCDSD_TUPLE_CACHE, Vec::new()),
        (rt::BEGIN_PCD14, pcd14),
        (rt::END_PCD14, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();

    let cache = definition.tuple_cache.unwrap();
    assert_eq!(
        cache.entries,
        [
            PivotCacheItemValue::Number(1.5),
            PivotCacheItemValue::String("entry".into()),
        ]
    );
    assert_eq!(cache.queries, ["SELECT ...".to_string()]);
    assert_eq!(cache.sets.len(), 1);
    assert_eq!(cache.sets[0].tuple_count, None);
    assert_eq!(cache.sets[0].max_rank, 3);
    assert_eq!(cache.sets[0].sort_order, 1);
    assert!(cache.sets[0].query_failed);
    assert_eq!(cache.sets[0].definition, "{[Dim].[Hier].Members}");

    let ext14 = definition.ext14.unwrap();
    assert!(ext14.slicer_data);
    assert!(ext14.server_support_subquery_non_visual);
    assert_eq!(ext14.cache_id, 9);
}

#[test]
fn skips_unknown_records_and_collections() {
    let definition = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        // Unknown standalone record with garbage payload.
        (Kind::new(0x0F00).unwrap(), vec![0xDE, 0xAD, 0xBE, 0xEF]),
        // FRT wrapper containing an unknown record.
        (rt::FRT_BEGIN, vec![0; 12]),
        (Kind::new(0x0F01).unwrap(), vec![1, 2, 3]),
        (rt::FRT_END, Vec::new()),
        (rt::BEGIN_PCD_FIELDS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_PCD_FIELD, minimal_field_payload("F")),
        // Unmodelled KPI collection inside the field.
        (rt::BEGIN_PCD_KPIS, 1u32.to_le_bytes().to_vec()), // BrtBeginPCDKPIs
        (rt::BEGIN_PCD_KPI, vec![9; 8]),                   // BrtBeginPCDKPI
        (rt::END_PCD_KPI, Vec::new()),                     // BrtEndPCDKPI
        (rt::END_PCD_KPIS, Vec::new()),                    // BrtEndPCDKPIs
        (Kind::new(0x0F02).unwrap(), vec![0xFF]), // unknown record inside the field collection
        (rt::END_PCD_FIELD, Vec::new()),
        (rt::END_PCD_FIELDS, Vec::new()),
        // Unmodelled top-level collection with nested content.
        (rt::BEGIN_PCD_SFCI_ENTRIES, 1u32.to_le_bytes().to_vec()),
        (rt::PCD_SFCI_ENTRY, wide_string("#,##0.00")), // BrtPCDSFCIEntry
        (rt::END_PCD_SFCI_ENTRIES, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap();

    assert_eq!(definition.fields.len(), 1);
    assert_eq!(definition.fields[0].name, "F");
    assert_eq!(definition.tuple_cache, None);
}

#[test]
fn rejects_malformed_streams() {
    // Does not start with BrtBeginPivotCacheDef.
    let error = parse(&[(rt::BEGIN_PCD_FIELDS, 0u32.to_le_bytes().to_vec())]).unwrap_err();
    assert!(matches!(error, Error::UnexpectedRecord { .. }));

    // Truncated definition payload.
    let error = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, vec![1, 2, 3]),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap_err();
    assert!(matches!(
        error,
        Error::Wire(WireError::Truncated {
            stage: Stage::Value,
            ..
        })
    ));

    // Missing BrtEndPivotCacheDef.
    let error = parse(&[(rt::BEGIN_PIVOT_CACHE_DEF, definition_payload())]).unwrap_err();
    assert!(matches!(error, Error::UnexpectedEndOfStream(_)));

    // Invalid enumerated value (iSrcType out of range).
    let mut source = Vec::new();
    source.extend_from_slice(&9u32.to_le_bytes());
    source.extend_from_slice(&0u32.to_le_bytes());
    let error = parse(&[
        (rt::BEGIN_PIVOT_CACHE_DEF, definition_payload()),
        (rt::BEGIN_PCD_SOURCE, source),
        (rt::END_PCD_SOURCE, Vec::new()),
        (rt::END_PIVOT_CACHE_DEF, Vec::new()),
    ])
    .unwrap_err();
    assert!(matches!(error, Error::Unrecognized { .. }));
}
