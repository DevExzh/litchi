#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;

use litchi_odraw::write::Header;
use zerocopy::IntoBytes;

#[test]
fn wire_view_retains_unknown_record_header_payload_and_instance() {
    let mut known = EscherBuilder::new(header_version::SIMPLE, 0, record_type::DG);
    known.add_data(&[1; 8]);
    let unknown_header = Header::new(0x00, 0x345, 0x7FFE, 3);
    let mut unknown = unknown_header.as_bytes().to_vec();
    unknown.extend_from_slice(&[0xAA, 0xBB, 0xCC]);
    let mut root = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);
    root.add_data(&known.build().expect("known record"));
    root.add_data(&unknown);

    let bytes = root.build().expect("root record");
    let view = EscherDrawing::parse(&bytes).expect("wire view");
    let unknown_records = view.unknown_records().expect("unknown records");
    assert_eq!(unknown_records.len(), 1);
    let record = unknown_records[0];
    assert_eq!(record.raw_kind(), 0x7FFE);
    assert_eq!(record.version(), 0);
    assert_eq!(record.instance(), 0x345);
    assert_eq!(record.bytes(), &unknown);
}

#[test]
fn wire_shape_validation_rejects_truncated_shape_atoms() {
    let mut shape = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);
    let mut atom = EscherBuilder::new(header_version::SP, shape_type::RECTANGLE, record_type::SP);
    atom.add_data(&[1, 2, 3]);
    shape.add_data(&atom.build().expect("shape atom"));
    let bytes = shape.build().expect("shape container");

    let view = EscherDrawing::parse(&bytes).expect("wire view accepts record framing");
    let error = view
        .validate_shapes()
        .expect_err("shape atom must be rejected");
    assert_eq!(error.kind(), std::io::ErrorKind::InvalidData);
}
