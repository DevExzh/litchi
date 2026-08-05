use std::io::Cursor;

use litchi_xls::writer::shape::{Anchor, Behavior, Point};
use litchi_xls::writer::{
    XlsCommentTextRunWrite, XlsCommentWriteOptions, XlsPivotTableConfig, XlsWriter,
};
use litchi_xls::{Visibility, XlsWorkbook};

fn workbook_records(bytes: Vec<u8>) -> Vec<(u16, Vec<u8>)> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset + 4 <= stream.len() {
        let record_type = u16::from_le_bytes(stream[offset..offset + 2].try_into().unwrap());
        let length = usize::from(u16::from_le_bytes(
            stream[offset + 2..offset + 4].try_into().unwrap(),
        ));
        let end = offset + 4 + length;
        records.push((record_type, stream[offset + 4..end].to_vec()));
        offset = end;
    }
    records
}

#[test]
fn comments_round_trip_unicode_runs_visibility_and_guid() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Comments").unwrap();
    let guid = [7u8; 16];
    writer
        .add_comment_with_options(
            sheet,
            3,
            2,
            "Author",
            "Hello 😀",
            XlsCommentWriteOptions {
                visible: true,
                shared: true,
                anchor: Some(
                    Anchor::new(
                        Point::new(3, 4).unwrap().offset(20, 10).unwrap(),
                        Point::new(8, 7).unwrap().offset(200, 900).unwrap(),
                        Behavior::MoveAndSize,
                    )
                    .unwrap(),
                ),
                text_runs: vec![XlsCommentTextRunWrite {
                    character_index: 0,
                    font_index: 0,
                }],
                font_when_empty: 0,
                guid: Some(guid),
            },
        )
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let comments = workbook.xls_worksheet(0).unwrap().comments();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].text(), "Hello 😀");
    assert_eq!(comments[0].visibility(), Visibility::Visible);
    assert!(comments[0].identity().shared());
    assert_eq!(comments[0].identity().guid(), &guid);
    assert_eq!(comments[0].text_runs().len(), 1);
}

#[test]
fn long_comment_splits_without_breaking_surrogates() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Long").unwrap();
    let text = "😀".repeat(5000);
    writer.add_comment(sheet, 0, 0, "A", &text).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().comments()[0].text(),
        text
    );
}

#[test]
fn comment_api_rejects_bounds_duplicates_and_bad_runs() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Bad").unwrap();
    assert!(writer.add_comment(sheet, 65_536, 0, "A", "x").is_err());
    assert!(writer.add_comment(sheet, 0, 256, "A", "x").is_err());
    assert!(writer.add_comment(sheet, 0, 0, "", "x").is_err());
    writer.add_comment(sheet, 0, 0, "A", "x").unwrap();
    assert!(writer.add_comment(sheet, 0, 0, "B", "y").is_err());
    assert!(
        writer
            .add_comment_with_options(
                sheet,
                1,
                0,
                "A",
                "x",
                XlsCommentWriteOptions {
                    text_runs: vec![XlsCommentTextRunWrite {
                        character_index: 1,
                        font_index: 0
                    }],
                    ..Default::default()
                }
            )
            .is_err()
    );
}

#[test]
fn comment_records_follow_client_boundaries_and_note_order() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Structure").unwrap();
    writer.add_comment(sheet, 0, 0, "A", "text").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let records = workbook_records(output.into_inner());
    let relevant = records
        .iter()
        .filter_map(|(record_type, _)| {
            matches!(
                *record_type,
                0x00EB | 0x00EC | 0x005D | 0x01B6 | 0x003C | 0x001C
            )
            .then_some(*record_type)
        })
        .collect::<Vec<_>>();
    assert_eq!(
        relevant,
        vec![
            0x00EB, 0x00EC, 0x005D, 0x00EC, 0x01B6, 0x003C, 0x003C, 0x001C
        ]
    );
    assert!(records.iter().all(|(_, data)| data.len() <= 8224));
}

#[test]
fn comments_and_pivot_objects_share_one_drawing_and_use_separate_object_ids() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Mixed").unwrap();
    writer
        .add_pivot_table(
            sheet,
            XlsPivotTableConfig {
                name: "Pivot".into(),
                source_type: 1,
                source_sheet_name: "Mixed".into(),
                source_first_row: 0,
                source_last_row: 0,
                source_first_col: 0,
                source_last_col: 0,
                first_row: 4,
                last_row: 4,
                first_col: 0,
                last_col: 0,
                first_header_row: 4,
                first_data_row: 4,
                first_data_col: 0,
                data_field_name: "Values".into(),
                data_axis: 0,
                data_position: 0,
                fields: Vec::new(),
                data_items: Vec::new(),
                page_entries: vec![(0, 0, 1)],
                source_data: Vec::new(),
            },
        )
        .unwrap();
    writer.add_comment(sheet, 1, 1, "A", "mixed").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let records = workbook_records(output.into_inner());
    let drawing_group = records
        .iter()
        .find(|(record_type, _)| *record_type == 0x00EB)
        .unwrap();
    assert_eq!(drawing_group.1.len(), 90);
    let comment_object = records
        .iter()
        .find(|(record_type, data)| {
            *record_type == 0x005D && data.get(4..6) == Some(&0x0019u16.to_le_bytes())
        })
        .unwrap();
    assert_eq!(
        u16::from_le_bytes(comment_object.1[6..8].try_into().unwrap()),
        2
    );
}
