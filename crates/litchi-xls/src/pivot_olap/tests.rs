//! Regression coverage for the layered PivotTable OLAP owner.

use super::*;

fn hierarchy() -> PivotHierarchy {
    PivotHierarchy {
        is_measure: false,
        outline_mode: true,
        multiple_page_items: false,
        subtotal_at_top: true,
        is_named_set: false,
        hidden_from_field_list: false,
        is_attribute_hierarchy: true,
        is_time_hierarchy: false,
        filter_inclusive: false,
        is_key_attribute_hierarchy: false,
        is_kpi: false,
        axis: PivotHierarchyAxis {
            row: true,
            column: false,
            page: false,
            data: false,
        },
        pivot_field_index: 2,
        axis_field_count: 1,
        drag_to_row: true,
        drag_to_column: true,
        drag_to_page: false,
        drag_to_data: true,
        drag_to_hide: true,
        unique_name: "[Product].[Category]".to_string(),
        display_name: "Category é".to_string(),
        default_member: "[Product].[Category].&[1]".to_string(),
        all_member: "[Product].[Category].[All]".to_string(),
        dimension: "[Product]".to_string(),
        level_fields: vec![2, -1],
        hidden_member_sets: vec![HiddenMemberSet {
            member_names: vec!["&[4]".to_string()],
        }],
    }
}

fn page_extension() -> PivotPageItemOlapExt {
    PivotPageItemOlapExt {
        hierarchy_index: 1,
        unique_name: "[Product].[Category].&[4]".to_string(),
        display_name: "Bikes".to_string(),
    }
}

fn field_extension() -> PivotFieldOlapExt {
    PivotFieldOlapExt {
        tensor_sort: true,
        drilled_level: false,
        items_drilled_by_default: true,
        member_property_in_report: true,
        member_property_in_tip: false,
        member_property_in_caption: true,
        hierarchy_index: 0,
        olap_level_index: 1,
        item_flags: vec![
            PivotItemOlapFlags {
                drilled_member: true,
                has_children: true,
                collapsed_member: false,
                has_children_estimated: true,
                olap_filter_selected: false,
            },
            PivotItemOlapFlags::default(),
        ],
    }
}

fn borrowed(records: &[(u16, Vec<u8>)]) -> Vec<(u16, &[u8])> {
    records
        .iter()
        .map(|(record_type, payload)| (*record_type, payload.as_slice()))
        .collect()
}

#[test]
fn view_header_round_trips_future_bytes() {
    let header = PivotViewOlapHeader {
        hierarchy_count: 2,
        page_extension_count: 1,
        field_extension_count: 3,
        future_bytes: vec![0xDE, 0xAD],
    };
    let payload = header.to_payload().expect("serialize");
    assert_eq!(PivotViewOlapHeader::parse(&payload).unwrap(), header);
    assert!(header.has_future_extensions());

    let mut invalid = header.clone();
    invalid.hierarchy_count = 0;
    assert!(invalid.to_payload().is_err());
}

#[test]
fn hierarchy_round_trips_unicode_and_hidden_levels() {
    let value = hierarchy();
    let payload = value.to_payload().expect("serialize");
    assert_eq!(PivotHierarchy::parse(&payload).unwrap(), value);
    assert_eq!(value.level_count(), 2);
    assert_eq!(value.hidden_level_count(), 1);
    assert_eq!(value.axis.axis_count(), 1);

    let mut invalid = value.clone();
    invalid.hidden_member_sets.push(HiddenMemberSet {
        member_names: vec!["&[9]".to_string()],
    });
    invalid.hidden_member_sets.push(HiddenMemberSet {
        member_names: vec!["&[10]".to_string()],
    });
    assert!(invalid.to_payload().is_err());
}

#[test]
fn page_and_field_extensions_round_trip() {
    let page = page_extension();
    assert_eq!(
        PivotPageItemOlapExt::parse(&page.to_payload().unwrap()).unwrap(),
        page
    );

    let field = field_extension();
    assert_eq!(
        PivotFieldOlapExt::parse(&field.to_payload().unwrap()).unwrap(),
        field
    );
    assert_eq!(field.item_count(), 2);
}

#[test]
fn sequence_package_enforces_order_counts_and_round_trip() {
    let sequence = OlapSequence::from_parts(
        vec![hierarchy()],
        vec![page_extension()],
        vec![field_extension()],
        vec![0xAA, 0xBB],
    )
    .expect("build sequence");
    let records = sequence.to_records().expect("encode sequence");
    let parsed = OlapSequence::parse(&borrowed(&records)).expect("parse sequence");
    assert_eq!(parsed, sequence);
    assert_eq!(parsed.to_records().expect("re-encode sequence"), records);

    let mut wrong_order = records.clone();
    wrong_order.swap(1, 2);
    assert!(OlapSequence::parse(&borrowed(&wrong_order)).is_err());

    let mut wrong_count = records[0].1.clone();
    // SXViewEx payload: FrtHeaderOld (4) then csxth (4).
    wrong_count[4..8].copy_from_slice(&2i32.to_le_bytes());
    let mut malformed = records.clone();
    malformed[0].1 = wrong_count;
    assert!(OlapSequence::parse(&borrowed(&malformed)).is_err());
}

#[test]
fn malformed_record_headers_and_reserved_bits_are_rejected() {
    let mut payload = page_extension().to_payload().unwrap();
    payload[0..2].copy_from_slice(&SXPI_EX_RECORD_TYPE.to_le_bytes());
    assert!(PivotPageItemOlapExt::parse(&payload).is_err());

    let mut field = field_extension().to_payload().unwrap();
    field[4..6].copy_from_slice(&0x0040u16.to_le_bytes());
    assert!(PivotFieldOlapExt::parse(&field).is_err());
}
