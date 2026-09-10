use litchi_doc::body_text::Snapshot;
use litchi_doc::{
    DofrArray, DofrFrameKind, DofrPayload, PgpArray, PgpType, PrintDriver, SavedSelection,
    SelectionGeometry, SelectionStyle,
};

#[test]
fn public_auxiliary_table_apis_expose_bounded_typed_data() {
    let mut selection_bytes = vec![0u8; 36];
    selection_bytes[2] = 1;
    selection_bytes[4..8].copy_from_slice(&4i32.to_le_bytes());
    selection_bytes[8..12].copy_from_slice(&8i32.to_le_bytes());
    selection_bytes[20..24].copy_from_slice(&4i32.to_le_bytes());
    selection_bytes[24..26].copy_from_slice(&(SelectionStyle::Character as u16).to_le_bytes());
    selection_bytes[32..34].copy_from_slice(&(-100i16).to_le_bytes());
    selection_bytes[34..36].copy_from_slice(&100i16.to_le_bytes());
    let selection = SavedSelection::parse_bytes(&selection_bytes).unwrap();
    let mut transaction = selection.transaction();
    transaction
        .set_range(2, 6)
        .unwrap()
        .set_geometry(SelectionGeometry::Block { first: 1, limit: 2 })
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit.patch().apply(&selection).unwrap(),
        *commit.snapshot()
    );

    let mut pgp = Vec::new();
    pgp.extend_from_slice(&1u16.to_le_bytes());
    pgp.extend_from_slice(&9u32.to_le_bytes());
    pgp.extend_from_slice(&0u32.to_le_bytes());
    pgp.extend_from_slice(&0u32.to_le_bytes());
    pgp.extend_from_slice(&0x0101u16.to_le_bytes());
    pgp.extend_from_slice(&6u16.to_le_bytes());
    pgp.extend_from_slice(&120i32.to_le_bytes());
    pgp.extend_from_slice(&1u16.to_le_bytes());
    let groups = PgpArray::parse_bytes(&pgp).unwrap();
    assert_eq!(
        groups.entries()[0].options().kind(),
        Some(PgpType::BlockQuote)
    );
    assert_eq!(groups.entries()[0].options().dxa_left(), Some(120));

    let driver = PrintDriver::parse_bytes(b"printer\0port\0driver\0product\0").unwrap();
    assert_eq!(driver.driver(), b"driver");

    let mut dofr = Vec::new();
    dofr.extend_from_slice(&8u32.to_le_bytes());
    dofr.extend_from_slice(&0u32.to_le_bytes());
    let mut frame = vec![0u8; 44];
    frame[0..4].copy_from_slice(&44u32.to_le_bytes());
    frame[4..8].copy_from_slice(&1u32.to_le_bytes());
    frame[20..24].copy_from_slice(&2u32.to_le_bytes());
    dofr.extend_from_slice(&frame);
    let records = DofrArray::parse_bytes(&dofr).unwrap();
    let DofrPayload::Frame(frame) = records.get(1).unwrap().payload().unwrap() else {
        panic!("expected typed Dofr frame");
    };
    assert_eq!(frame.frame_kind(), DofrFrameKind::Frame);
}

#[test]
fn public_body_snapshot_reads_auxiliary_tables_without_raw_stream_access() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/NoHeadFoot.doc");
    let snapshot = Snapshot::parse(&std::fs::read(path).expect("real DOC fixture"))
        .expect("safe body snapshot");
    assert!(snapshot.saved_selection().is_ok());
    assert!(snapshot.dofr_records().is_ok());
}
