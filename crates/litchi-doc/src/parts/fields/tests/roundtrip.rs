use super::super::model::*;
use super::super::package::*;
use super::plcf;

#[test]
fn nested_fields_flags_and_roundtrip_are_exact() {
    let data = plcf(
        &[1, 3, 5, 7, 9, 11, 13],
        &[
            [0x13, 0x07],
            [0x13, 0x21],
            [0x14, 0xFF],
            [0x15, 0xC1],
            [0x14, 0xA5],
            [0x15, 0xB1],
        ],
    );
    let table = FieldStoryTable::parse_plcf(FieldStory::Main, 20, &data).unwrap();
    assert_eq!(table.to_plcf_bytes().unwrap(), data);
    assert_eq!(table.fields().len(), 2);
    assert_eq!(table.fields()[0].field_type, FieldType::If);
    assert_eq!(table.fields()[0].separator_cp, Some(9));
    assert!(table.fields()[0].end_flags.locked);
    assert_eq!(table.fields()[1].field_type, FieldType::Page);
    assert_eq!(table.fields()[1].nesting_depth, 1);
    assert!(table.fields()[1].end_flags.nested);
}

#[test]
fn unrecognized_field_type_is_preserved_with_valid_boundaries() {
    let data = plcf(&[1, 3, 5, 7], &[[0x13, 0x4E], [0x14, 0], [0x15, 0x80]]);
    let table = FieldStoryTable::parse_plcf(FieldStory::Main, 7, &data).unwrap();

    assert_eq!(table.fields().len(), 1);
    assert_eq!(table.fields()[0].field_type, FieldType::Unknown(0x4E));
    assert_eq!(table.fields()[0].separator_cp, Some(3));
    assert_eq!(table.to_plcf_bytes().unwrap(), data);
}

#[cfg(test)]
mod terminal_cp_regression_tests {
    use super::{FieldStory, FieldStoryTable};

    fn field_plcf(cps: &[u32]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&[0x13, 0x21]);
        data.extend_from_slice(&[0x15, 0x00]);
        data
    }

    #[test]
    fn terminal_cp_is_not_story_bounded_but_marker_cps_are() {
        let data = field_plcf(&[1, 3, u32::MAX]);
        let parsed = FieldStoryTable::parse_plcf(FieldStory::Main, 3, &data)
            .expect("undefined terminal CP may exceed the story length");
        assert_eq!(parsed.terminal_cp, u32::MAX);
        assert_eq!(parsed.markers[0].position, 1);
        assert_eq!(parsed.markers[1].position, 3);

        let marker_outside_story = field_plcf(&[1, 4, 5]);
        assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 3, &marker_outside_story).is_err());
    }
}
