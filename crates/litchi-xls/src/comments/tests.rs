//! Regression coverage for the BIFF8 comment owner.

use super::codec::{CommentCollector, parse_note_record, parse_txo_runs};
use super::model::*;

fn obj(object_id: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&0x0015u16.to_le_bytes());
    data.extend_from_slice(&0x0012u16.to_le_bytes());
    data.extend_from_slice(&COMMENT_OBJECT_TYPE.to_le_bytes());
    data.extend_from_slice(&object_id.to_le_bytes());
    data.extend_from_slice(&0x4011u16.to_le_bytes());
    data.extend_from_slice(&[0; 12]);
    data.extend_from_slice(&0x000Du16.to_le_bytes());
    data.extend_from_slice(&0x0016u16.to_le_bytes());
    let mut guid = [0u8; 16];
    guid[0..2].copy_from_slice(&object_id.to_le_bytes());
    data.extend_from_slice(&guid);
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&[0; 4]);
    data.extend_from_slice(&[0; 4]);
    data
}

fn txo(character_count: u16, run_bytes: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&0x0212u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&[0; 6]);
    data.extend_from_slice(&character_count.to_le_bytes());
    data.extend_from_slice(&run_bytes.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data
}

fn client_textbox() -> [u8; 8] {
    [0, 0, 0x0D, 0xF0, 0, 0, 0, 0]
}

fn note(object_id: u16, flags: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&5u16.to_le_bytes());
    data.extend_from_slice(&3u16.to_le_bytes());
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(&object_id.to_le_bytes());
    data.extend_from_slice(&4u16.to_le_bytes());
    data.push(0);
    data.extend_from_slice(b"User");
    data.push(0xD0);
    data
}

fn runs(character_count: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&6u16.to_le_bytes());
    data.extend_from_slice(&[0; 4]);
    data.extend_from_slice(&character_count.to_le_bytes());
    data.extend_from_slice(&[0; 6]);
    data
}

#[test]
fn links_obj_txo_continues_and_note() {
    let mut collector = CommentCollector::new();
    collector.feed_record(OBJ_TYPE, &obj(7)).unwrap();
    collector
        .feed_record(MSODRAWING_TYPE, &client_textbox())
        .unwrap();
    collector.feed_record(TXO_TYPE, &txo(5, 16)).unwrap();
    collector.feed_record(CONTINUE_TYPE, b"\0Hello").unwrap();
    collector.feed_record(CONTINUE_TYPE, &runs(5)).unwrap();
    collector
        .feed_record(RECORD_TYPE, &note(7, 0x0182))
        .unwrap();
    let comments = collector.finish().unwrap();
    let comment = &comments[0];
    assert_eq!(comment.row(), 5);
    assert_eq!(comment.column(), 3);
    assert_eq!(comment.visibility(), Visibility::Visible);
    assert!(comment.row_hidden() && comment.column_hidden());
    assert_eq!(comment.identity().object_id(), 7);
    assert_eq!(comment.author(), "User");
    assert_eq!(comment.text(), "Hello");
    assert_eq!(comment.text_runs()[0].font_index(), 6);
}

#[test]
fn assembles_mixed_segmented_unicode_without_splitting_surrogates() {
    let mut collector = CommentCollector::new();
    collector.feed_record(OBJ_TYPE, &obj(8)).unwrap();
    collector
        .feed_record(MSODRAWING_TYPE, &client_textbox())
        .unwrap();
    collector.feed_record(TXO_TYPE, &txo(3, 16)).unwrap();
    collector.feed_record(CONTINUE_TYPE, b"\0A").unwrap();
    let mut wide = vec![1];
    wide.extend_from_slice(&0xD83Du16.to_le_bytes());
    wide.extend_from_slice(&0xDE00u16.to_le_bytes());
    collector.feed_record(CONTINUE_TYPE, &wide).unwrap();
    collector.feed_record(CONTINUE_TYPE, &runs(3)).unwrap();
    collector.feed_record(RECORD_TYPE, &note(8, 0)).unwrap();
    assert_eq!(collector.finish().unwrap()[0].text(), "A😀");
}

#[test]
fn retains_reserved_bits_and_rejects_broken_order_or_bad_last_run() {
    let mut malformed_note = note(1, 1);
    let parsed = parse_note_record(&malformed_note).unwrap();
    assert_eq!(parsed.metadata.reserved_flags(), 1);
    malformed_note[4..6].copy_from_slice(&0u16.to_le_bytes());
    malformed_note.pop();
    assert!(parse_note_record(&malformed_note).is_err());

    let mut collector = CommentCollector::new();
    collector.feed_record(OBJ_TYPE, &obj(1)).unwrap();
    assert!(collector.feed_record(TXO_TYPE, &txo(1, 16)).is_err());

    let mut collector = CommentCollector::new();
    collector.feed_record(OBJ_TYPE, &obj(2)).unwrap();
    assert!(collector.feed_record(MSODRAWING_TYPE, &[0; 8]).is_err());

    let mut bad_runs = runs(2);
    bad_runs[8..10].copy_from_slice(&1u16.to_le_bytes());
    assert!(parse_txo_runs(&bad_runs, 2).is_err());
}

#[test]
fn preserves_opaque_fields_unknown_obj_subrecords_padding_and_record_order() {
    let mut object = obj(9);
    object[0..4].copy_from_slice(&[0xAA, 0xBB, 0xCC, 0xDD]);
    object[8..10].copy_from_slice(&0x4013u16.to_le_bytes());
    object[10..22].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
    object[42..44].copy_from_slice(&2u16.to_le_bytes());
    object[44..48].copy_from_slice(&[0x11, 0x22, 0x33, 0x44]);
    let end = object.len() - 4;
    object.truncate(end);
    object.extend_from_slice(&[0x7F, 0x7F, 3, 0, 0xE1, 0xE2, 0xE3]);
    object.extend_from_slice(&[0, 0, 0, 0]);
    object.extend_from_slice(&[0xFA, 0xFB, 0xFC, 0xFD]);

    let mut text = txo(1, 16);
    text[0] |= 1;
    text[4..10].copy_from_slice(&[0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6]);
    let mut note_data = note(9, 0x0003 | 0x0400);
    note_data[10] = 0x80;

    let mut collector = CommentCollector::new();
    collector.feed_record(OBJ_TYPE, &object).unwrap();
    collector
        .feed_record(MSODRAWING_TYPE, &client_textbox())
        .unwrap();
    collector.feed_record(TXO_TYPE, &text).unwrap();
    collector.feed_record(CONTINUE_TYPE, b"\0A").unwrap();
    collector.feed_record(CONTINUE_TYPE, &runs(1)).unwrap();
    collector.feed_record(RECORD_TYPE, &note_data).unwrap();

    let comment = collector.finish().unwrap().pop().unwrap();
    assert_eq!(
        comment.object_properties().reserved_header(),
        &[0xAA, 0xBB, 0xCC, 0xDD]
    );
    assert_eq!(comment.object_properties().reserved_flags(), 2);
    assert_eq!(
        comment.object_properties().unused_bytes(),
        &[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    );
    assert_eq!(comment.identity().shared_value(), 2);
    assert_eq!(comment.identity().unused_bytes(), &[0x11, 0x22, 0x33, 0x44]);
    assert_eq!(comment.object_subrecords()[1].record_type(), 0x7F7F);
    assert!(!comment.object_subrecords()[1].is_known());
    assert_eq!(
        comment.object_subrecords()[1].payload(),
        &[0xE1, 0xE2, 0xE3]
    );
    assert_eq!(comment.object_padding(), &[0xFA, 0xFB, 0xFC, 0xFD]);
    assert_eq!(comment.text_properties().reserved_options(), 1);
    assert_eq!(
        comment.text_properties().reserved_fields(),
        &[0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6]
    );
    assert_eq!(comment.note_metadata().reserved_flags(), 0x0401);
    assert_eq!(comment.note_metadata().reserved_string_flags(), 0x80);
    assert_eq!(
        comment
            .records()
            .iter()
            .map(CommentRecord::kind)
            .collect::<Vec<_>>(),
        vec![
            RecordKind::Object,
            RecordKind::Drawing,
            RecordKind::TextObject,
            RecordKind::Continue,
            RecordKind::Continue,
            RecordKind::Note,
        ]
    );
}

#[test]
fn enforces_record_and_opaque_payload_bounds() {
    assert!(CommentRecord::new(RecordKind::Note, &vec![0; MAX_RECORD_BYTES + 1]).is_err());
    let mut collector = CommentCollector::new();
    assert!(
        collector
            .feed_record(OBJ_TYPE, &vec![0; MAX_RECORD_BYTES + 1])
            .is_err()
    );
}

#[test]
fn reads_poi_comment_fixtures() {
    use crate::Workbook;
    use std::fs::File;
    use std::path::Path;

    let fixture = |name: &str| {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    };

    let simple = Workbook::new(File::open(fixture("SimpleWithComments.xls")).unwrap()).unwrap();
    let comments = simple.xls_worksheet(0).unwrap().comments();
    assert_eq!(comments.len(), 3);
    assert_eq!(comments[0].author(), "Yegor Kozlov");
    assert_eq!(comments[0].text(), "Yegor Kozlov:\nfirst cell");
    assert_eq!(comments[1].text(), "Yegor Kozlov:\nsecond cell");
    assert_eq!(comments[2].visibility(), Visibility::Visible);
    assert_eq!(comments[0].identity().object_id(), 1);
    assert_ne!(comments[0].identity().guid(), &[0; 16]);
    assert_eq!(comments[0].text_runs().len(), 2);

    let drawing = Workbook::new(File::open(fixture("DrawingAndComments.xls")).unwrap()).unwrap();
    let comments = drawing.xls_worksheet(0).unwrap().comments();
    assert_eq!(comments.len(), 3);
    assert!(comments.iter().all(|comment| !comment.text().is_empty()));

    let libreoffice = Workbook::new(File::open(fixture("comments.xls")).unwrap()).unwrap();
    let comments = libreoffice.xls_worksheet(0).unwrap().comments();
    assert_eq!(comments.len(), 3);
    assert!(
        comments
            .iter()
            .all(|comment| comment.author() == "Sven Nissel")
    );
    assert_eq!(comments[0].text(), "comment top row1 (index0)\n");
}
