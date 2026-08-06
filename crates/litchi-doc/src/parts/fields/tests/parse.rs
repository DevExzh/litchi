use super::super::model::*;
use super::super::package::*;
use super::plcf;

#[test]
fn malformed_plcf_and_fieldlist_matrix_is_rejected() {
    let valid = plcf(&[1, 3, 5, 7], &[[0x13, 0x21], [0x14, 0], [0x15, 0x80]]);
    assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 7, &valid).is_ok());
    for invalid in [
        Vec::new(),
        vec![0; 5],
        plcf(&[1, 1], &[[0x13, 0x21]]),
        plcf(&[2, 1], &[[0x13, 0x21]]),
        plcf(&[1, 9], &[[0x13, 0x21]]),
        plcf(&[1, 3], &[[0x12, 0x21]]),
        plcf(&[1, 3], &[[0x14, 0]]),
        plcf(&[1, 3], &[[0x15, 0]]),
        plcf(&[1, 3], &[[0x13, 0x21]]),
        plcf(&[1, 3, 5], &[[0x13, 0x21], [0x15, 0x80]]),
        plcf(
            &[1, 2, 3, 4, 5],
            &[[0x13, 0x21], [0x14, 0], [0x14, 0], [0x15, 0x80]],
        ),
        plcf(
            &[1, 2, 3, 4, 5],
            &[[0x13, 0x07], [0x13, 0x21], [0x15, 0], [0x15, 0]],
        ),
    ] {
        assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 7, &invalid).is_err());
    }
}

#[test]
fn all_end_flags_and_reserved_descriptor_bits_are_preserved() {
    let descriptor = FieldDescriptor::from_bytes(&[0xF5, 0xFF]).unwrap();
    assert_eq!(descriptor.reserved_bits, 7);
    let FieldMarkerValue::End(flags) = descriptor.value else {
        panic!("end descriptor");
    };
    assert!(flags.differ && flags.zombie_embed && flags.results_dirty);
    assert!(flags.results_edited && flags.locked && flags.private_result);
    assert!(flags.nested && flags.has_separator);
    assert_eq!(descriptor.to_bytes(), [0xF5, 0xFF]);
}

#[test]
fn field_type_mapping_covers_specified_and_unknown_values() {
    assert_eq!(FieldType::from(0x0E), FieldType::Info);
    assert_eq!(FieldType::Info.as_u8(), 0x0E);
    assert_eq!(FieldType::from(0x3A), FieldType::EmbeddedObject);
    assert_eq!(FieldType::EmbeddedObject.as_u8(), 0x3A);
    assert_eq!(FieldType::from(0x3F), FieldType::BarCode);
    assert_eq!(FieldType::BarCode.as_u8(), 0x3F);
    assert_eq!(FieldType::from(0x5C), FieldType::BidiOutline);
    assert_eq!(FieldType::BidiOutline.as_u8(), 0x5C);
    assert_eq!(FieldType::from(0x5F), FieldType::Shape);
    assert_eq!(FieldType::Shape.as_u8(), 0x5F);
    assert_eq!(FieldType::from(0x46), FieldType::FormText);
    assert_eq!(FieldType::FormText.as_u8(), 0x46);
    assert_eq!(FieldType::from(0x47), FieldType::FormCheckbox);
    assert_eq!(FieldType::FormCheckbox.as_u8(), 0x47);
    assert_eq!(FieldType::from(0x53), FieldType::FormDropdown);
    assert_eq!(FieldType::FormDropdown.as_u8(), 0x53);
    assert_eq!(FieldType::from(0x58), FieldType::Hyperlink);
    assert_eq!(FieldType::from(0x34), FieldType::AutoNumOutline);
    assert_eq!(FieldType::from(0x35), FieldType::AutoNumLegal);
    assert_eq!(FieldType::from(0x36), FieldType::AutoNum);
    assert_eq!(FieldType::from(0x39), FieldType::Symbol);
    assert_eq!(FieldType::from(0x30), FieldType::Print);
    assert_eq!(FieldType::Print.as_u8(), 0x30);
    assert_eq!(FieldType::from(0x5A), FieldType::ListNumber);
    assert_eq!(FieldType::from(0x41), FieldType::Section);
    assert_eq!(FieldType::from(0x42), FieldType::SectionPages);
    assert_eq!(FieldType::from(0x45), FieldType::FileSize);
    assert_eq!(FieldType::from(0x04), FieldType::Unknown(0x04));
    assert_eq!(
        FieldType::from_keyword("hyperlink"),
        Some(FieldType::Hyperlink)
    );
    assert_eq!(FieldType::from_keyword("="), Some(FieldType::Formula));
    assert_eq!(
        FieldType::from_keyword("FTNREF"),
        Some(FieldType::FootnoteRef)
    );
    assert_eq!(FieldType::from_keyword("TC"), None);
    assert_eq!(FieldType::from_keyword("future-field"), None);
}
