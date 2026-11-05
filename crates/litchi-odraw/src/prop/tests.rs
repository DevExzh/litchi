#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use crate::{Error, Record, RecordKind};

use super::{Array, ColorRef, IS_BLIP, IS_COMPLEX, Id, Props, Value};

fn opt_record(data: &[u8], instance: u16) -> Record<'_> {
    Record::from_parts(RecordKind::Opt, 3, instance, data, crate::Limits::default())
        .expect("valid fixture")
}

fn opt_record_with_version(data: &[u8], instance: u16, version: u8) -> Record<'_> {
    Record::from_parts(
        RecordKind::Opt,
        version,
        instance,
        data,
        crate::Limits::default(),
    )
    .expect("valid fixture")
}

fn push_property(data: &mut Vec<u8>, id: u16, value: i32) {
    data.extend_from_slice(&id.to_le_bytes());
    data.extend_from_slice(&value.to_le_bytes());
}

#[test]
fn decodes_packed_boolean_property_groups() {
    let mut data = Vec::new();
    push_property(&mut data, 0x01BF, 0x0014_0010);
    push_property(&mut data, 0x01FF, 0x0008_0008);
    push_property(&mut data, 0x023F, 0x0002_0002);
    push_property(&mut data, 0x00FF, 0x0020_0020);
    push_property(&mut data, 0x033F, 0x0001_0001);
    push_property(&mut data, 0x00BF, 0x001A_0012);

    let properties =
        Props::parse(&opt_record(&data, 6), crate::Limits::default()).expect("valid properties");

    assert_eq!(properties.get_bool(Id::Filled), Some(true));
    assert_eq!(properties.get_bool(Id::FillShape), Some(false));
    assert_eq!(properties.get_bool(Id::HitTestFill), None);
    assert!(properties.is_filled());
    assert_eq!(properties.get_bool(Id::AnyLine), Some(true));
    assert!(properties.has_line());
    assert_eq!(properties.get_bool(Id::Shadow), Some(true));
    assert_eq!(properties.get_bool(Id::ShadowObscured), None);
    assert!(properties.has_shadow());
    assert_eq!(properties.get_bool(Id::GeoTextBoldFont), Some(true));
    assert_eq!(properties.get_bool(Id::GeoTextUnderlineFont), None);
    assert_eq!(properties.get_bool(Id::ShapeBackgroundShape), Some(true));
    assert_eq!(properties.get_bool(Id::SelectText), Some(true));
    assert_eq!(properties.get_bool(Id::AutoTextMargin), Some(false));
    assert_eq!(properties.get_bool(Id::FitShapeToText), Some(true));
}

#[test]
fn decodes_explicit_false_boolean_group_bits() {
    let mut data = Vec::new();
    push_property(&mut data, 0x01FF, 0x0008_0000);
    push_property(&mut data, 0x023F, 0x0002_0000);

    let properties =
        Props::parse(&opt_record(&data, 2), crate::Limits::default()).expect("valid properties");

    assert_eq!(properties.get_bool(Id::AnyLine), Some(false));
    assert!(!properties.has_line());
    assert_eq!(properties.get_bool(Id::Shadow), Some(false));
    assert!(!properties.has_shadow());
}

#[test]
fn text_margins_use_ms_odraw_defaults() {
    let properties = Props::new();
    assert_eq!(
        properties.get_text_margins(),
        Some((0x0001_6530, 0x0000_B298, 0x0001_6530, 0x0000_B298))
    );
}

#[test]
fn fill_and_line_resolvers_apply_spec_defaults_without_hiding_absence() {
    let properties = Props::new();

    assert_eq!(properties.get_bool(Id::Filled), None);
    assert_eq!(properties.get_bool(Id::AnyLine), None);
    assert!(properties.is_filled());
    assert!(properties.has_line());
    assert!(!properties.has_shadow());
}

#[test]
fn decodes_writer_fill_enabled_and_disabled_masks() {
    let mut enabled_data = Vec::new();
    push_property(&mut enabled_data, 0x01BF, 0x0015_0011);
    let enabled = Props::parse(&opt_record(&enabled_data, 1), crate::Limits::default())
        .expect("valid properties");

    assert_eq!(enabled.get_bool(Id::Filled), Some(true));
    assert_eq!(enabled.get_bool(Id::FillShape), Some(false));
    assert_eq!(enabled.get_bool(Id::NoFillHitTest), Some(true));
    assert_eq!(enabled.get_bool(Id::HitTestFill), None);

    let mut disabled_data = Vec::new();
    push_property(&mut disabled_data, 0x01BF, 0x0010_0000);
    let disabled = Props::parse(&opt_record(&disabled_data, 1), crate::Limits::default())
        .expect("valid properties");
    assert_eq!(disabled.get_bool(Id::Filled), Some(false));
}

#[test]
fn accepts_direct_boolean_properties_from_lenient_producers() {
    let mut data = Vec::new();
    push_property(&mut data, Id::Filled.raw(), 1);

    let properties =
        Props::parse(&opt_record(&data, 1), crate::Limits::default()).expect("valid properties");

    assert_eq!(properties.get_bool(Id::Filled), Some(true));
}

#[test]
fn rejects_negative_complex_property_lengths_without_panicking() {
    let mut data = Vec::new();
    push_property(&mut data, IS_COMPLEX | Id::Vertices.raw(), -1);

    assert!(matches!(
        Props::parse(&opt_record(&data, 1), crate::Limits::default()),
        Err(Error::MalformedProperties { .. })
    ));
}

#[test]
fn preserves_distinct_unknown_property_ids() {
    let mut data = Vec::new();
    push_property(&mut data, 0x0600, 11);
    push_property(&mut data, 0x0601, 12);

    let properties =
        Props::parse(&opt_record(&data, 2), crate::Limits::default()).expect("valid properties");
    let first = Id::unknown(0x0600).expect("unassigned identifier");
    let second = Id::unknown(0x0601).expect("unassigned identifier");
    assert_eq!(first.raw(), 0x0600);
    assert_eq!(properties.get_int(first), Some(11));
    assert_eq!(properties.get_int(second), Some(12));
    assert!(Id::unknown(IS_BLIP | 0x0600).is_none());
    assert!(Id::unknown(Id::FillColor.raw()).is_none());
}

#[test]
fn preserves_order_flags_raw_ids_and_raw_values() {
    let mut data = Vec::new();
    push_property(&mut data, IS_BLIP | 0x0601, -7);
    push_property(&mut data, IS_BLIP | IS_COMPLEX | Id::GroupName.raw(), 3);
    data.extend_from_slice(&[0x41, 0x00, 0x00]);

    let properties =
        Props::parse(&opt_record(&data, 2), crate::Limits::default()).expect("valid properties");
    let entries = properties.iter().collect::<Vec<_>>();
    assert_eq!(entries[0].raw_id(), 0x0601);
    assert_eq!(entries[0].raw_opid(), IS_BLIP | 0x0601);
    assert!(entries[0].is_blip());
    assert!(!entries[0].is_complex());
    assert_eq!(entries[0].raw_value(), -7);
    assert_eq!(entries[1].id(), Id::GroupName);
    assert_eq!(
        entries[1].raw_opid(),
        IS_BLIP | IS_COMPLEX | Id::GroupName.raw()
    );
    assert!(entries[1].is_blip());
    assert!(entries[1].is_complex());
    assert_eq!(entries[1].raw_value(), 3);
    assert_eq!(
        properties.get_binary(Id::GroupName),
        Some(&[0x41, 0, 0][..])
    );
}

#[test]
fn rejects_duplicate_semantic_identifiers_even_when_flags_differ() {
    let mut data = Vec::new();
    push_property(&mut data, Id::FillColor.raw(), 1);
    push_property(&mut data, IS_BLIP | Id::FillColor.raw(), 2);

    assert!(matches!(
        Props::parse(&opt_record(&data, 2), crate::Limits::default()),
        Err(Error::MalformedProperties {
            reason: "duplicate property identifier"
        })
    ));
}

#[test]
fn requires_opt_family_version_three() {
    assert!(matches!(
        Props::parse(
            &opt_record_with_version(&[], 0, 2),
            crate::Limits::default()
        ),
        Err(Error::MalformedProperties {
            reason: "Opt-family property table must have recVer 3"
        })
    ));
}

#[test]
fn classifies_arrays_by_property_id_instead_of_payload_shape() {
    let array_bytes = [0, 0, 0, 0, 4, 0];

    let mut scalar_complex = Vec::new();
    push_property(&mut scalar_complex, IS_COMPLEX | Id::GroupName.raw(), 6);
    scalar_complex.extend_from_slice(&array_bytes);
    let scalar = Props::parse(&opt_record(&scalar_complex, 1), crate::Limits::default())
        .expect("valid complex value");
    assert!(matches!(scalar.get(Id::GroupName), Some(Value::Complex(_))));

    let mut typed_array = Vec::new();
    push_property(&mut typed_array, IS_COMPLEX | Id::Vertices.raw(), 6);
    typed_array.extend_from_slice(&array_bytes);
    let array = Props::parse(&opt_record(&typed_array, 1), crate::Limits::default())
        .expect("valid array value");
    assert!(matches!(array.get(Id::Vertices), Some(Value::Array(_))));
}

#[test]
fn validates_imsoarray_header_and_exact_extent() {
    let mut special = Vec::new();
    special.extend_from_slice(&2u16.to_le_bytes());
    special.extend_from_slice(&3u16.to_le_bytes());
    special.extend_from_slice(&0xFFF0u16.to_le_bytes());
    special.extend_from_slice(&[0; 8]);
    let special_array = Array::new(&special).expect("valid truncated-element array");
    assert_eq!(special_array.raw_element_size(), 0xFFF0);
    assert_eq!(special_array.element_size(), 4);
    assert_eq!(special_array.elements().count(), 2);

    let mut underallocated = Vec::new();
    underallocated.extend_from_slice(&2u16.to_le_bytes());
    underallocated.extend_from_slice(&1u16.to_le_bytes());
    underallocated.extend_from_slice(&4u16.to_le_bytes());
    underallocated.extend_from_slice(&[0; 8]);
    assert!(Array::new(&underallocated).is_err());

    let mut other_high_size = Vec::new();
    other_high_size.extend_from_slice(&1u16.to_le_bytes());
    other_high_size.extend_from_slice(&1u16.to_le_bytes());
    other_high_size.extend_from_slice(&0xFFF1u16.to_le_bytes());
    other_high_size.extend_from_slice(&[0; 4]);
    assert!(Array::new(&other_high_size).is_err());

    let mut trailing = Vec::new();
    trailing.extend_from_slice(&1u16.to_le_bytes());
    trailing.extend_from_slice(&1u16.to_le_bytes());
    trailing.extend_from_slice(&4u16.to_le_bytes());
    trailing.extend_from_slice(&[0; 5]);
    assert!(Array::new(&trailing).is_err());
}

#[test]
fn color_ref_is_lossless_and_decodes_only_direct_rgb() {
    let direct = ColorRef::from_raw(0x0033_2211);
    assert_eq!(direct.raw(), 0x0033_2211);
    assert_eq!(direct.flags(), 0);
    assert_eq!(direct.rgb(), Some((0x11, 0x22, 0x33)));

    let scheme = ColorRef::from_raw(0x0800_0004);
    assert_eq!(scheme.raw(), 0x0800_0004);
    assert_eq!(scheme.flags(), 0x08);
    assert_eq!(scheme.rgb(), None);
}

#[test]
fn compact_and_indexed_tables_preserve_lookup_order_and_duplicate_rejection() {
    for count in [0_u16, 1, 16, 17, 64] {
        let mut data = Vec::new();
        for index in 0..count {
            push_property(&mut data, 0x600 + index, i32::from(index));
        }
        let record = opt_record(&data, count);
        let props = Props::parse(&record, crate::Limits::default()).expect("table");
        assert_eq!(props.len(), usize::from(count));
        for index in 0..count {
            let id = Id::from(0x600 + index);
            assert!(props.has(id));
            assert_eq!(props.get_int(id), Some(i32::from(index)));
            assert_eq!(
                props.iter().nth(usize::from(index)).expect("ordered").id(),
                id
            );
        }
        assert!(!props.has(Id::from(0x700)));
        if count > 0 {
            push_property(&mut data, 0x4600, -1);
            assert!(matches!(
                Props::parse(&opt_record(&data, count + 1), crate::Limits::default()),
                Err(Error::MalformedProperties { .. })
            ));
        }
    }
}
