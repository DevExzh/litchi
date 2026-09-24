use super::codec::{
    encode_xl_string_no_cch, write_sx_stream_id, write_sxdi, write_sxex, write_sxpi, write_sxvd,
    write_sxvi, write_sxview, write_sxvs,
};
use super::model::{
    PivotCacheFieldInfo, SxDiConfig, SxExConfig, SxVdConfig, SxViConfig, SxViewConfig,
};
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

/// An empty `XLUnicodeStringNoCch` keeps its option byte, which the reader
/// consumes when `cch` is zero. Until change 0766 the encoder wrote nothing,
/// so a PivotTable with an empty name shifted its data field name by a byte.
#[test]
fn an_empty_no_cch_string_keeps_its_option_byte() {
    assert_eq!(encode_xl_string_no_cch(""), [0x00]);
    let mut buf = Vec::new();
    write_sxview(&mut buf, &view_named("", "Values")).unwrap();
    let view = crate::pivot_table::parse_sxview(&buf[4..]).unwrap();
    assert_eq!(view.name, "");
    assert_eq!(view.data_field_name, "Values");
}

fn view_named<'a>(name: &'a str, data_field_name: &'a str) -> SxViewConfig<'a> {
    SxViewConfig {
        first_row: 8,
        last_row: 10,
        first_col: 0,
        last_col: 1,
        first_header_row: 8,
        first_data_row: 9,
        first_data_col: 1,
        cache_index: 0,
        data_axis: 0,
        data_position: 0,
        field_count: 2,
        row_field_count: 1,
        col_field_count: 0,
        page_field_count: 0,
        data_field_count: 1,
        data_row_count: 2,
        data_col_count: 1,
        flags: 0x020B,
        auto_format_index: 1,
        name,
        data_field_name,
    }
}

/// PivotTable names count UTF-16 code units, so a supplementary-plane
/// character counts two; until change 0766 the count was one short and the
/// reader refused the workbook ("lone surrogate found").
#[test]
fn pivot_names_count_utf16_code_units_up_to_their_limits() {
    for (name, data_name) in [
        ("S😀".to_string(), "V😀".to_string()),
        (
            format!("{}😀", "t".repeat(253)),
            format!("{}😀", "d".repeat(252)),
        ),
        ("t".repeat(255), "d".repeat(254)),
    ] {
        let mut buf = Vec::new();
        write_sxview(&mut buf, &view_named(&name, &data_name)).unwrap();
        let view = crate::pivot_table::parse_sxview(&buf[4..]).unwrap();
        assert_eq!(view.name, name);
        assert_eq!(view.data_field_name, data_name);
    }
    for (name, data_name) in [
        ("t".repeat(256), "d".to_string()),
        (format!("{}😀", "t".repeat(254)), "d".to_string()),
        ("t".to_string(), "d".repeat(255)),
        ("t".to_string(), format!("{}😀", "d".repeat(253))),
    ] {
        let mut buf = Vec::new();
        assert!(matches!(
            write_sxview(&mut buf, &view_named(&name, &data_name)),
            Err(crate::Error::StringTooLong { .. })
        ));
        assert!(buf.is_empty());
    }
    let mut buf = Vec::new();
    assert!(write_sxview(&mut buf, &view_named("t", "")).is_err());

    for (name, fits) in [
        ("F😀".to_string(), true),
        (format!("{}😀", "f".repeat(253)), true),
        (format!("{}😀", "f".repeat(254)), false),
    ] {
        let mut buf = Vec::new();
        let written = write_sxvd(
            &mut buf,
            &SxVdConfig {
                axis: 0x0001,
                subtotal_count: 0,
                subtotal_flags: 0,
                item_count: 0,
                name: Some(&name),
            },
        );
        assert_eq!(written.is_ok(), fits, "{name:?}");
        if fits {
            let field = crate::pivot_table::parse_sxvd(&buf[4..]).unwrap();
            assert_eq!(field.name.as_deref(), Some(name.as_str()));
        }
    }
    for (name, fits) in [
        ("I😀".to_string(), true),
        (format!("{}😀", "i".repeat(252)), true),
        (format!("{}😀", "i".repeat(253)), false),
    ] {
        let mut buf = Vec::new();
        let written = write_sxvi(
            &mut buf,
            &SxViConfig {
                item_type: 0,
                flags: 0,
                cache_index: 0,
                name: Some(&name),
            },
        );
        assert_eq!(written.is_ok(), fits, "{name:?}");
        if fits {
            let item = crate::pivot_table::parse_sxvi(&buf[4..]).unwrap();
            assert_eq!(item.name.as_deref(), Some(name.as_str()));
        }
    }
    for (name, fits) in [
        ("D😀".to_string(), true),
        (format!("{}😀", "d".repeat(253)), true),
        (format!("{}😀", "d".repeat(254)), false),
    ] {
        let mut buf = Vec::new();
        let written = write_sxdi(
            &mut buf,
            &SxDiConfig {
                source_field_index: 0,
                function: 0,
                display_format: 0,
                base_field_index: 0,
                base_item_index: 0,
                num_format_index: 0,
                name: &name,
            },
        );
        assert_eq!(written.is_ok(), fits, "{name:?}");
        if fits {
            let item = crate::pivot_table::parse_sxdi(&buf[4..]).unwrap();
            assert_eq!(item.name, name);
        }
    }
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
