use std::io::Cursor;

use litchi_xls::shapes::extract_shapes_from_workbook;
use litchi_xls::writer::{
    XlsPivotTableConfig, XlsShapeAnchor, XlsShapeColor, XlsShapeFill, XlsShapeKind, XlsShapeLine,
    XlsShapeText, XlsShapeTextRun, XlsShapeWrite, XlsWriter,
};

fn anchor() -> XlsShapeAnchor {
    XlsShapeAnchor {
        move_with_cells: true,
        size_with_cells: true,
        first_column: 1,
        first_column_offset: 10,
        first_row: 1,
        first_row_offset: 20,
        last_column: 4,
        last_column_offset: 900,
        last_row: 6,
        last_row_offset: 200,
    }
}

fn workbook_stream(bytes: Vec<u8>) -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    ole.open_stream(&["Workbook"]).unwrap()
}

fn records(stream: &[u8]) -> Vec<(u16, Vec<u8>)> {
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

fn write(writer: &mut XlsWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    workbook_stream(output.into_inner())
}

#[test]
fn primitives_emit_exact_client_record_order_and_parse_after_write() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Shapes").unwrap();

    let mut rectangle = XlsShapeWrite::new(XlsShapeKind::Rectangle, anchor());
    rectangle.object_id = Some(7);
    rectangle.fill = XlsShapeFill::Solid(XlsShapeColor::rgb(0x12, 0x34, 0x56));
    rectangle.line = XlsShapeLine::None;
    assert_eq!(writer.add_shape(sheet, rectangle).unwrap(), 7);

    let mut textbox = XlsShapeWrite::new(XlsShapeKind::TextBox, anchor());
    textbox.text = Some(XlsShapeText {
        value: "Hello 世界".to_string(),
        runs: vec![XlsShapeTextRun {
            character_index: 0,
            font_index: 0,
        }],
        font_when_empty: 0,
    });
    assert_eq!(writer.add_shape(sheet, textbox).unwrap(), 1);

    let stream = write(&mut writer);
    let records = records(&stream);
    let relevant = records
        .iter()
        .filter_map(|(kind, _)| {
            matches!(*kind, 0x00EB | 0x00EC | 0x005D | 0x01B6 | 0x003C).then_some(*kind)
        })
        .collect::<Vec<_>>();
    assert_eq!(
        relevant,
        vec![
            0x00EB, 0x00EC, 0x005D, 0x00EC, 0x005D, 0x00EC, 0x01B6, 0x003C, 0x003C
        ]
    );
    assert!(records.iter().all(|(_, data)| data.len() <= 8224));

    let drawings = records
        .iter()
        .filter(|(kind, _)| *kind == 0x00EC)
        .map(|(_, data)| data)
        .collect::<Vec<_>>();
    assert_eq!(drawings.len(), 3);
    assert!(drawings[0].windows(2).any(|pair| pair == [0x10, 0xF0]));
    assert!(drawings[0].windows(2).any(|pair| pair == [0x11, 0xF0]));
    assert!(drawings[1].windows(2).any(|pair| pair == [0x11, 0xF0]));
    assert_eq!(&drawings[2][2..4], &[0x0D, 0xF0]);

    let officeart = drawings
        .iter()
        .flat_map(|data| data.iter().copied())
        .collect::<Vec<_>>();
    assert_eq!(
        u32::from_le_bytes(officeart[4..8].try_into().unwrap()) as usize,
        officeart.len() - 8
    );

    let parsed = extract_shapes_from_workbook(&stream).unwrap();
    assert_eq!(parsed.len(), 2);
    assert_eq!(parsed[0].shape_id, 1025);
    assert_eq!(parsed[1].shape_id, 1026);
    assert_eq!(parsed[1].text.as_deref(), Some("Hello 世界"));
}

#[test]
fn all_safe_primitive_types_use_unique_shape_and_object_ids() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Kinds").unwrap();
    for kind in [
        XlsShapeKind::Rectangle,
        XlsShapeKind::RoundedRectangle,
        XlsShapeKind::Ellipse,
        XlsShapeKind::Line,
        XlsShapeKind::TextBox,
    ] {
        let mut shape = XlsShapeWrite::new(kind, anchor());
        if kind == XlsShapeKind::Line {
            shape.fill = XlsShapeFill::None;
        }
        writer.add_shape(sheet, shape).unwrap();
    }
    let stream = write(&mut writer);
    let records = records(&stream);
    let object_ids = records
        .iter()
        .filter(|(kind, data)| *kind == 0x005D && data.len() >= 8)
        .map(|(_, data)| u16::from_le_bytes(data[6..8].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(object_ids, vec![1, 2, 3, 4, 5]);
    let parsed = extract_shapes_from_workbook(&stream).unwrap();
    assert_eq!(
        parsed
            .iter()
            .map(|shape| shape.shape_id)
            .collect::<Vec<_>>(),
        vec![1025, 1026, 1027, 1028, 1029]
    );
}

#[test]
fn mixed_pivot_shape_and_comment_ids_are_collision_free_in_one_cluster() {
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
    let mut shape = XlsShapeWrite::new(XlsShapeKind::Ellipse, anchor());
    shape.object_id = Some(2);
    writer.add_shape(sheet, shape).unwrap();
    writer.add_comment(sheet, 1, 1, "A", "note").unwrap();

    let records = records(&write(&mut writer));
    assert_eq!(
        records.iter().filter(|(kind, _)| *kind == 0x00EB).count(),
        1
    );
    let drawing_group = records.iter().find(|(kind, _)| *kind == 0x00EB).unwrap();
    assert_eq!(drawing_group.1.len(), 90);
    let object_ids = records
        .iter()
        .filter(|(kind, data)| *kind == 0x005D && data.len() >= 8)
        .map(|(_, data)| u16::from_le_bytes(data[6..8].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(object_ids, vec![1, 2, 3]);
}

#[test]
fn shape_mutations_reject_malformed_input_and_are_atomic() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Atomic").unwrap();

    let mut invalid = XlsShapeWrite::new(XlsShapeKind::Rectangle, anchor());
    invalid.anchor.last_column = 0;
    assert!(writer.add_shape(sheet, invalid).is_err());

    let mut first = XlsShapeWrite::new(XlsShapeKind::Rectangle, anchor());
    first.object_id = Some(9);
    writer.add_shape(sheet, first).unwrap();
    let mut collision = XlsShapeWrite::new(XlsShapeKind::Ellipse, anchor());
    collision.object_id = Some(9);
    assert!(writer.add_shape(sheet, collision).is_err());
    assert!(writer.remove_shape(sheet, 0).is_err());
    assert_eq!(writer.clear_shapes(sheet).unwrap(), 1);
    assert_eq!(writer.clear_shapes(sheet).unwrap(), 0);

    let mut line = XlsShapeWrite::new(XlsShapeKind::Line, anchor());
    line.text = Some(XlsShapeText::new("unsupported"));
    assert!(writer.add_shape(sheet, line).is_err());
    assert!(
        records(&write(&mut writer))
            .iter()
            .all(|(kind, _)| *kind != 0x00EC && *kind != 0x005D)
    );
}

#[test]
fn long_unicode_shape_text_respects_continue_record_limits() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Long").unwrap();
    let mut shape = XlsShapeWrite::new(XlsShapeKind::TextBox, anchor());
    shape.text = Some(XlsShapeText::new("😀".repeat(5000)));
    writer.add_shape(sheet, shape).unwrap();
    let records = records(&write(&mut writer));
    assert!(records.iter().all(|(_, data)| data.len() <= 8224));
    assert!(records.iter().filter(|(kind, _)| *kind == 0x003C).count() >= 4);
}
