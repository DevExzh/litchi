use super::codec::{
    encode_xl_string_no_cch, write_sx_stream_id, write_sxdi, write_sxex, write_sxpi, write_sxvd,
    write_sxvi, write_sxvs,
};
use super::model::{PivotCacheFieldInfo, SxDiConfig, SxExConfig, SxVdConfig, SxViConfig};
use super::validation::{validate_sxdbb_index, validate_sxdbb_inputs};

#[test]
fn test_write_sxvs() {
    let mut buf = Vec::new();
    write_sxvs(&mut buf, 0x0001).unwrap();
    assert_eq!(&buf[0..2], &[0xE3, 0x00]);
    assert_eq!(&buf[2..4], &[0x02, 0x00]);
    assert_eq!(&buf[4..6], &[0x01, 0x00]);
}

#[test]
fn test_write_sxvd_no_name() {
    let mut buf = Vec::new();
    write_sxvd(
        &mut buf,
        &SxVdConfig {
            axis: 0x0001,
            subtotal_count: 0,
            subtotal_flags: 0,
            item_count: 5,
            name: None,
        },
    )
    .unwrap();
    assert_eq!(&buf[0..2], &[0xB1, 0x00]);
    assert_eq!(&buf[12..14], &[0xFF, 0xFF]);
}

#[test]
fn test_write_sxvi_data() {
    let mut buf = Vec::new();
    write_sxvi(
        &mut buf,
        &SxViConfig {
            item_type: 0x00FE,
            flags: 0,
            cache_index: 3,
            name: None,
        },
    )
    .unwrap();
    assert_eq!(&buf[0..2], &[0xB2, 0x00]);
}

#[test]
fn test_write_sxpi() {
    let mut buf = Vec::new();
    write_sxpi(&mut buf, &[(1, 0, 0), (2, 1, 0)]).unwrap();
    assert_eq!(&buf[0..2], &[0xB6, 0x00]);
    assert_eq!(&buf[2..4], &[12, 0]);
}

#[test]
fn test_write_sx_stream_id() {
    let mut buf = Vec::new();
    write_sx_stream_id(&mut buf, 0).unwrap();
    assert_eq!(&buf[0..2], &[0xD5, 0x00]);
    assert_eq!(&buf[2..4], &[0x02, 0x00]);
    assert_eq!(&buf[4..6], &[0x00, 0x00]);
}

#[test]
fn test_write_sxex_default() {
    let mut buf = Vec::new();
    write_sxex(&mut buf, &SxExConfig::default()).unwrap();
    assert_eq!(&buf[0..2], &[0xF1, 0x00]);
    assert_eq!(&buf[2..4], &[24, 0]);
    assert_eq!(&buf[6..8], &[0xFF, 0xFF]);
}

#[test]
fn test_encode_xl_string_no_cch_empty() {
    assert!(encode_xl_string_no_cch("").is_empty());
}

#[test]
fn test_write_sxdi_empty_name() {
    let mut buf = Vec::new();
    write_sxdi(
        &mut buf,
        &SxDiConfig {
            source_field_index: 0,
            function: 0,
            display_format: 0,
            base_field_index: 0,
            base_item_index: 0,
            num_format_index: 0,
            name: "",
        },
    )
    .unwrap();
    assert_eq!(&buf[16..18], &[0xFF, 0xFF]);
    assert_eq!(&buf[2..4], &[14, 0]);
}

#[test]
fn validation_rejects_shared_index_cardinality_mismatch() {
    let fields = [PivotCacheFieldInfo {
        name: "field",
        items: &[],
        is_numeric: false,
        unique_numeric_count: 0,
        grouping: None,
        group_child: None,
        is_source_field: true,
    }];

    assert!(validate_sxdbb_inputs(&fields, &[0]).is_err());
}

#[test]
fn validation_preserves_biff_index_width_rules() {
    assert!(validate_sxdbb_index(0xFF, false).is_ok());
    assert!(validate_sxdbb_index(0x100, false).is_err());
    assert!(validate_sxdbb_index(0xFFFF, true).is_ok());
}
