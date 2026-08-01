use std::io::Cursor;

use litchi_odraw::shape::Kind;
use litchi_ole::xls::shapes::extract_shapes_from_workbook;
use litchi_ole::xls::writer::{
    XlsGroupRect, XlsShapeAnchor, XlsShapeColor, XlsShapeFill, XlsShapeGroupChild,
    XlsShapeGroupWrite, XlsShapeKind, XlsShapeLine, XlsShapeText, XlsShapeWrite, XlsWriter,
};

fn anchor() -> XlsShapeAnchor {
    XlsShapeAnchor {
        move_with_cells: true,
        size_with_cells: true,
        first_column: 1,
        first_column_offset: 10,
        first_row: 1,
        first_row_offset: 20,
        last_column: 6,
        last_column_offset: 900,
        last_row: 9,
        last_row_offset: 200,
    }
}

fn group_with_children() -> XlsShapeGroupWrite {
    let mut group = XlsShapeGroupWrite::new(anchor());
    group.coordinates = XlsGroupRect::new(0, 0, 2000, 1000);
    let mut rectangle =
        XlsShapeGroupChild::new(XlsShapeKind::Rectangle, XlsGroupRect::new(0, 0, 900, 500));
    rectangle.fill = XlsShapeFill::Solid(XlsShapeColor::rgb(0x20, 0x40, 0x60));
    group.children.push(rectangle);
    let mut textbox = XlsShapeGroupChild::new(
        XlsShapeKind::TextBox,
        XlsGroupRect::new(900, 400, 2000, 1000),
    );
    textbox.line = XlsShapeLine::None;
    textbox.text = Some(XlsShapeText::new("Grouped 世界"));
    group.children.push(textbox);
    group
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
fn grouped_shapes_emit_exact_record_order_and_round_trip() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Groups").unwrap();

    let standalone = XlsShapeWrite::new(XlsShapeKind::Rectangle, anchor());
    assert_eq!(writer.add_shape(sheet, standalone).unwrap(), 1);
    assert_eq!(
        writer
            .add_shape_group(sheet, group_with_children())
            .unwrap(),
        2
    );

    let stream = write(&mut writer);
    let records = records(&stream);
    let relevant = records
        .iter()
        .filter_map(|(kind, _)| {
            matches!(*kind, 0x00EB | 0x00EC | 0x005D | 0x01B6 | 0x003C).then_some(*kind)
        })
        .collect::<Vec<_>>();
    // Prefix + primitive, group header, plain child, then textbox child with TXO.
    assert_eq!(
        relevant,
        vec![
            0x00EB, 0x00EC, 0x005D, 0x00EC, 0x005D, 0x00EC, 0x005D, 0x00EC, 0x005D, 0x00EC, 0x01B6,
            0x003C, 0x003C
        ]
    );
    assert!(records.iter().all(|(_, data)| data.len() <= 8224));

    // The group OBJ carries ftCmo (type 0x0000) plus the mandatory ftGmo marker.
    let objs = records
        .iter()
        .filter(|(kind, _)| *kind == 0x005D)
        .map(|(_, data)| data)
        .collect::<Vec<_>>();
    assert_eq!(objs.len(), 4);
    let group_obj = objs[1];
    assert_eq!(group_obj.len(), 32);
    assert_eq!(u16::from_le_bytes(group_obj[4..6].try_into().unwrap()), 0);
    assert_eq!(u16::from_le_bytes(group_obj[6..8].try_into().unwrap()), 2);
    assert_eq!(
        u16::from_le_bytes(group_obj[22..24].try_into().unwrap()),
        0x0006
    );
    assert_eq!(
        u16::from_le_bytes(group_obj[24..26].try_into().unwrap()),
        0x0002
    );
    assert!(
        objs.iter()
            .enumerate()
            .all(|(index, data)| index == 1 || data.len() == 26)
    );
    let object_ids = objs
        .iter()
        .map(|data| u16::from_le_bytes(data[6..8].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(object_ids, vec![1, 2, 3, 4]);

    // The DgContainer length spans every MsoDrawing fragment of the worksheet.
    let drawings = records
        .iter()
        .filter(|(kind, _)| *kind == 0x00EC)
        .map(|(_, data)| data)
        .collect::<Vec<_>>();
    assert_eq!(drawings.len(), 5);
    let escher = drawings
        .iter()
        .flat_map(|data| data.iter().copied())
        .collect::<Vec<u8>>();
    assert_eq!(u16::from_le_bytes(escher[2..4].try_into().unwrap()), 0xF002);
    assert_eq!(
        u32::from_le_bytes(escher[4..8].try_into().unwrap()) as usize,
        escher.len() - 8
    );

    // The group SpgrContainer declares a length spanning its header and children.
    let group_fragment = drawings[1];
    assert_eq!(
        u16::from_le_bytes(group_fragment[2..4].try_into().unwrap()),
        0xF003
    );
    assert_eq!(
        u32::from_le_bytes(group_fragment[4..8].try_into().unwrap()) as usize,
        (group_fragment.len() - 8) + drawings[2].len() + drawings[3].len() + drawings[4].len()
    );

    // The workbook drawing-group cluster accounts for group and child shape IDs.
    let drawing_group = records
        .iter()
        .find(|(kind, _)| *kind == 0x00EB)
        .map(|(_, data)| data)
        .unwrap();
    assert_eq!(
        u32::from_le_bytes(drawing_group[24..28].try_into().unwrap()),
        5
    );
    assert_eq!(
        u32::from_le_bytes(drawing_group[32..36].try_into().unwrap()),
        1
    );
    assert_eq!(
        u32::from_le_bytes(drawing_group[36..40].try_into().unwrap()),
        5
    );

    // Round-trip through the group-aware reader.
    let parsed = extract_shapes_from_workbook(&stream).unwrap();
    assert_eq!(parsed.len(), 2);
    assert_eq!(parsed[0].shape_id, 1025);
    assert_eq!(parsed[0].shape_type, Kind::Rectangle);
    assert!(!parsed[0].is_group);

    let group = &parsed[1];
    assert!(group.is_group);
    assert_eq!(group.shape_type, Kind::Group);
    assert_eq!(group.shape_id, 1026);
    assert_eq!(group.children.len(), 2);
    assert_eq!(group.children[0].shape_id, 1027);
    assert_eq!(group.children[0].shape_type, Kind::Rectangle);
    assert_eq!(group.children[1].shape_id, 1028);
    assert_eq!(group.children[1].shape_type, Kind::TextBox);
    assert_eq!(group.children[1].text.as_deref(), Some("Grouped 世界"));
}

#[test]
fn group_object_ids_are_collision_free_across_primitives_and_groups() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Ids").unwrap();

    let mut group = group_with_children();
    group.object_id = Some(2);
    group.children[0].object_id = Some(5);
    assert_eq!(writer.add_shape_group(sheet, group).unwrap(), 2);

    // The second group's children skip every identifier the first one claimed.
    assert_eq!(
        writer
            .add_shape_group(sheet, group_with_children())
            .unwrap(),
        3
    );

    let mut colliding = XlsShapeWrite::new(XlsShapeKind::Ellipse, anchor());
    colliding.object_id = Some(5);
    assert!(writer.add_shape(sheet, colliding).is_err());
    let free = XlsShapeWrite::new(XlsShapeKind::Ellipse, anchor());
    assert_eq!(writer.add_shape(sheet, free).unwrap(), 7);

    let stream = write(&mut writer);
    let records = records(&stream);
    let object_ids = records
        .iter()
        .filter(|(kind, data)| *kind == 0x005D && data.len() >= 8)
        .map(|(_, data)| u16::from_le_bytes(data[6..8].try_into().unwrap()))
        .collect::<Vec<_>>();
    // Primitives serialize before groups; every ID is unique.
    assert_eq!(object_ids, vec![7, 2, 5, 1, 3, 4, 6]);

    let parsed = extract_shapes_from_workbook(&stream).unwrap();
    assert_eq!(parsed.len(), 3);
    assert_eq!(parsed.iter().filter(|shape| shape.is_group).count(), 2);
}

#[test]
fn group_mutations_reject_malformed_input_and_are_atomic() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Atomic").unwrap();

    // Structurally invalid groups never partially register.
    assert!(
        writer
            .add_shape_group(sheet, XlsShapeGroupWrite::new(anchor()))
            .is_err()
    );
    let mut degenerate = group_with_children();
    degenerate.coordinates = XlsGroupRect::new(0, 0, 0, 1000);
    assert!(writer.add_shape_group(sheet, degenerate).is_err());
    let mut duplicated = group_with_children();
    duplicated.object_id = Some(9);
    duplicated.children[1].object_id = Some(9);
    assert!(writer.add_shape_group(sheet, duplicated).is_err());

    let mut primitive = XlsShapeWrite::new(XlsShapeKind::Rectangle, anchor());
    primitive.object_id = Some(3);
    writer.add_shape(sheet, primitive).unwrap();
    let mut colliding = group_with_children();
    colliding.children[0].object_id = Some(3);
    assert!(writer.add_shape_group(sheet, colliding).is_err());

    let group_id = writer
        .add_shape_group(sheet, group_with_children())
        .unwrap();
    assert_eq!(group_id, 1);
    assert!(writer.remove_shape_group(sheet, 99).is_err());
    let removed = writer.remove_shape_group(sheet, group_id).unwrap();
    assert_eq!(removed.object_id, Some(group_id));
    assert_eq!(removed.children.len(), 2);
    assert!(writer.remove_shape_group(sheet, group_id).is_err());

    // Only the standalone primitive remains in the serialized stream.
    let records = records(&write(&mut writer));
    let object_ids = records
        .iter()
        .filter(|(kind, data)| *kind == 0x005D && data.len() >= 8)
        .map(|(_, data)| u16::from_le_bytes(data[6..8].try_into().unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(object_ids, vec![3]);
}
