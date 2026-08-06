use litchi_doc::ole_controls::{
    Flags, Format, Metadata, OcxInfo, Persist1, Persist2, RgxOcxInfo, Story, parse_metadata,
    to_metadata_bytes,
};

#[test]
fn public_owner_round_trips_typed_object_metadata() {
    let persist1 = Persist1::try_new(
        false, false, false, false, false, false, true, false, true, 0x402D,
    )
    .unwrap();
    let persist2 = Persist2::try_new(false, true, false, 0xFFF0).unwrap();
    let metadata = Metadata::try_new(persist1, Format::Metafile, Some(persist2)).unwrap();
    let bytes = to_metadata_bytes(&metadata).unwrap();
    assert_eq!(parse_metadata(&bytes).unwrap(), metadata);
}

#[test]
fn public_owner_editor_keeps_the_source_snapshot_immutable() {
    let source = RgxOcxInfo::try_new(vec![OcxInfo::new(
        7,
        0,
        0,
        0,
        Flags::new(false, false, false, false, false, false, false, 0),
        Story::Main,
        0,
    )])
    .unwrap();
    let replacement = OcxInfo::new(
        8,
        1,
        0,
        0,
        Flags::new(true, false, false, false, false, false, false, 0),
        Story::Header,
        0,
    );
    let mut edit = source.edit();
    edit.replace(0, replacement).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.infos()[0].cookie(), 7);
    assert_eq!(commit.snapshot().infos()[0].cookie(), 8);
}
