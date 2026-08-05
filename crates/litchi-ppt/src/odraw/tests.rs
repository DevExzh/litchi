use super::codec::text_from_ppt_records;
use super::{anchor, parse, text_from_drawing};
use litchi_odraw::RecordKind;

fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

#[test]
fn ppt_text_atoms_use_their_specified_encodings() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&4000u16.to_le_bytes());
    bytes.extend_from_slice(&4u32.to_le_bytes());
    bytes.extend_from_slice(&[0x3d, 0xd8, 0x00, 0xde]);
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&4008u16.to_le_bytes());
    bytes.extend_from_slice(&2u32.to_le_bytes());
    bytes.extend_from_slice(&[0x80, 0xe9]);

    let mut text = String::new();
    text_from_ppt_records(&bytes, &mut text).unwrap();

    assert_eq!(text, "😀\n\u{80}é");
}

#[test]
fn malformed_embedded_record_stops_without_panicking() {
    let mut text = String::new();
    assert!(text_from_ppt_records(&[0; 7], &mut text).is_err());
    assert!(text.is_empty());
}

#[test]
fn concatenated_drawing_roots_are_complete_and_strict() {
    let dg = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
    let root = record(0x0f, 0, RecordKind::DgContainer.raw(), &dg);
    let stream = [root.as_slice(), root.as_slice()].concat();

    assert!(parse(&stream).unwrap().is_empty());
    assert_eq!(text_from_drawing(&stream).unwrap(), "");

    let malformed = [stream.as_slice(), &[0]].concat();
    assert!(parse(&malformed).is_err());
    assert!(text_from_drawing(&malformed).is_err());
}

#[test]
fn ppt_client_anchor_projects_small_rect_order() {
    let mut shape_atom = Vec::new();
    shape_atom.extend_from_slice(&42u32.to_le_bytes());
    shape_atom.extend_from_slice(&0x0A00u32.to_le_bytes());
    let mut shape = record(2, 1, RecordKind::Sp.raw(), &shape_atom);
    let anchor_data = [
        20i16.to_le_bytes(),
        10i16.to_le_bytes(),
        110i16.to_le_bytes(),
        70i16.to_le_bytes(),
    ]
    .concat();
    shape.extend(record(0, 0, RecordKind::ClientAnchor.raw(), &anchor_data));
    let shape = record(0x0f, 0, RecordKind::SpContainer.raw(), &shape);
    let mut drawing_children = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
    drawing_children.extend(shape);
    let drawing = record(0x0f, 0, RecordKind::DgContainer.raw(), &drawing_children);

    let shapes = parse(&drawing).unwrap();
    let anchor = anchor(&shapes[0]).unwrap().unwrap();
    assert_eq!((anchor.left(), anchor.top()), (10, 20));
    assert_eq!((anchor.right(), anchor.bottom()), (110, 70));
    assert_eq!((anchor.width(), anchor.height()), (100, 50));
}
