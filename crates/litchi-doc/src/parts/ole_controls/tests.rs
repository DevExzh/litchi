use super::{
    Entry, FIB_INDEX_PLC_OCX, FieldCounts, Flags, Format, Metadata, ObjectPool, OcxInfo, Persist1,
    Persist2, RgxOcxInfo, StorageName, Story, parse, parse_bytes, parse_metadata, to_bytes,
    to_metadata_bytes,
};
use crate::parts::fib::FileInformationBlock;

const FIB_POINTER_COUNT: usize = FIB_INDEX_PLC_OCX + 1;

fn set_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
    let start = 154 + index * 8;
    fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
    fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
}

fn fib_with_pointer(offset: u32, length: u32) -> FileInformationBlock {
    let mut data = vec![0; 154 + FIB_POINTER_COUNT * 8];
    data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
    data[152..154].copy_from_slice(&(FIB_POINTER_COUNT as u16).to_le_bytes());
    set_pointer(&mut data, FIB_INDEX_PLC_OCX, offset, length);
    FileInformationBlock::parse(&data).expect("valid fixture FIB")
}

fn sample() -> OcxInfo {
    OcxInfo::new(
        0x1020_3040,
        0x5566_7788,
        0x90AB_CDEF,
        7,
        Flags::new(true, false, true, false, true, true, false, 0xA5),
        Story::HeaderTextbox,
        0x55AA,
    )
}

#[test]
fn round_trips_fixed_records_and_ignored_values() {
    let table = RgxOcxInfo::try_new(vec![
        sample(),
        OcxInfo::new(
            9,
            10,
            11,
            12,
            Flags::new(false, true, false, true, false, false, true, 0x3C),
            Story::Main,
            0xCAFE,
        ),
    ])
    .unwrap();
    let bytes = to_bytes(&table).unwrap();
    assert_eq!(bytes.len(), 4 + 2 * 20);
    let parsed = parse_bytes(&bytes).unwrap();
    assert_eq!(parsed, table);
    assert_eq!(parsed.infos()[0].accelerator_handle(), 0x90AB_CDEF);
    assert_eq!(parsed.infos()[0].flags().reserved_bits(), 0xA5);
    assert_eq!(parsed.infos()[0].reserved(), 0x55AA);
}

#[test]
fn enforces_cookie_uniqueness_and_record_shape() {
    let duplicate = RgxOcxInfo::try_new(vec![sample(), sample()]);
    assert!(duplicate.is_err());

    let mut bytes = vec![1, 0, 0, 0];
    bytes.extend_from_slice(&[0; 19]);
    assert!(parse_bytes(&bytes).is_err());
    bytes.push(0);
    assert!(parse_bytes(&bytes).is_err());

    assert!(parse_bytes(&[0, 0, 0, 0, 0]).is_err());
    assert!(parse_bytes(&u32::MAX.to_le_bytes()).is_err());
    assert!(parse_bytes(&[0; 3]).is_err());
}

#[test]
fn rejects_invalid_story_and_fifld() {
    let table = to_bytes(&RgxOcxInfo::try_new(vec![sample()]).unwrap()).unwrap();

    let mut bad_story = table.clone();
    bad_story[4 + 16..4 + 18].copy_from_slice(&5u16.to_le_bytes());
    assert!(parse_bytes(&bad_story).is_err());

    let mut bad_fifld = table;
    bad_fifld[4 + 14..4 + 16].copy_from_slice(&0xA4u16.to_le_bytes());
    assert!(parse_bytes(&bad_fifld).is_err());
}

#[test]
fn reads_the_fc_plcocx_table_pointer() {
    let payload = to_bytes(&RgxOcxInfo::try_new(vec![sample()]).unwrap()).unwrap();
    let mut table_stream = vec![0xCC; 5];
    let offset = table_stream.len();
    table_stream.extend_from_slice(&payload);
    let fib = fib_with_pointer(offset as u32, payload.len() as u32);
    assert_eq!(
        parse(&fib, &table_stream).unwrap().unwrap().infos(),
        [sample()]
    );

    let absent = fib_with_pointer(0, 0);
    assert!(parse(&absent, &[]).unwrap().is_none());

    let undefined_empty = fib_with_pointer(1, 0);
    assert!(parse(&undefined_empty, &[0]).unwrap().is_none());

    let truncated = fib_with_pointer(4, payload.len() as u32);
    assert!(parse(&truncated, &table_stream).is_err());
}

fn metadata_sample() -> Metadata {
    let persist1 = Persist1::try_new(
        true, true, true, false, true, true, true, true, true, 0x402D,
    )
    .unwrap();
    let persist2 = Persist2::try_new(true, true, true, 0xFFF0).unwrap();
    Metadata::try_new(persist1, Format::UnicodeText, Some(persist2)).unwrap()
}

#[test]
fn object_metadata_round_trips_exact_persist_words() {
    let metadata = metadata_sample();
    let bytes = to_metadata_bytes(&metadata).unwrap();
    assert_eq!(bytes, [0x7F, 0xF3, 0x14, 0x00, 0xFD, 0xFF]);
    assert_eq!(parse_metadata(&bytes).unwrap(), metadata);
    assert_eq!(Metadata::parse_bytes(&bytes).unwrap(), metadata);
    assert_eq!(metadata.to_bytes().unwrap(), bytes);
    assert_eq!(
        to_metadata_bytes(&parse_metadata(&bytes).unwrap()).unwrap(),
        bytes
    );
    assert!(metadata.is_activex());
    assert!(metadata.stores_control_data_in_stream());
}

#[test]
fn object_metadata_preserves_absent_and_zero_persist2() {
    let persist1 = Persist1::try_new(
        false, false, false, false, false, false, false, false, false, 0,
    )
    .unwrap();
    let without = Metadata::try_new(persist1, Format::Text, None).unwrap();
    assert_eq!(to_metadata_bytes(&without).unwrap(), [0, 0, 2, 0]);

    let with_zero = Metadata::try_new(
        persist1,
        Format::Text,
        Some(Persist2::try_new(false, false, false, 0).unwrap()),
    )
    .unwrap();
    assert_eq!(to_metadata_bytes(&with_zero).unwrap(), [0, 0, 2, 0, 0, 0]);
    assert_ne!(without, with_zero);
}

#[test]
fn object_metadata_rejects_invalid_domains_and_relationships() {
    assert!(parse_metadata(&[0, 0, 0x06, 0]).is_err());
    assert!(parse_metadata(&[0x00, 0x04, 0x02, 0x00]).is_err());
    assert!(parse_metadata(&[0x00, 0x20, 0x02, 0x00]).is_err());
    assert!(parse_metadata(&[0x00, 0x00, 0x02, 0x00, 0x02, 0x00]).is_err());
    assert!(parse_metadata(&[0, 0, 2, 0, 0]).is_err());
}

#[test]
fn validates_story_specific_field_indices() {
    let table = RgxOcxInfo::try_new(vec![OcxInfo::new(
        1,
        2,
        0,
        0,
        Flags::new(false, false, false, false, false, false, false, 0),
        Story::Main,
        0,
    )])
    .unwrap();
    assert!(
        table
            .validate_fields(FieldCounts::new(2, 0, 0, 0, 0, 0, 0))
            .is_err()
    );
    assert!(
        table
            .validate_fields(FieldCounts::new(3, 0, 0, 0, 0, 0, 0))
            .is_ok()
    );
}

#[test]
fn object_pool_validates_inert_activex_stream_metadata() {
    let metadata = metadata_sample();
    let name = StorageName::try_new("_42").unwrap();
    let entry = Entry::try_with_streams(
        name,
        Some("{00000000-0000-0000-0000-000000000000}".to_owned()),
        Some(metadata),
        true,
        true,
        true,
    )
    .unwrap();
    let pool = ObjectPool::try_new(vec![entry]).unwrap();
    let entry = pool.get("_42").unwrap();
    let active_x = entry.active_x().unwrap();
    assert!(active_x.stream_data());
    assert!(active_x.data_present());
    assert!(entry.print_present());
    assert!(entry.enhanced_print_present());
    assert_eq!(
        entry.class_id(),
        Some("{00000000-0000-0000-0000-000000000000}")
    );
    assert!(
        Entry::try_new(
            StorageName::try_new("_43").unwrap(),
            None,
            Some(metadata),
            false,
        )
        .is_err()
    );
    assert!(StorageName::try_new("bad").is_err());
}

#[test]
fn editor_commit_round_trips_without_mutating_the_source_snapshot() {
    let source = RgxOcxInfo::try_new(vec![sample()]).unwrap();
    let replacement = OcxInfo::new(
        0xABCD,
        3,
        4,
        5,
        Flags::new(false, true, false, true, false, false, true, 0x11),
        Story::Comment,
        0xCAFE,
    );
    let mut editor = source.edit();
    editor.replace(0, replacement).unwrap();
    editor
        .push(OcxInfo::new(
            99,
            1,
            2,
            3,
            Flags::new(false, false, false, false, false, false, false, 0),
            Story::Main,
            0,
        ))
        .unwrap();
    let commit = editor.commit().unwrap();
    assert_eq!(commit.patch().changes().len(), 2);
    assert_eq!(source.infos(), &[sample()]);
    let bytes = to_bytes(commit.snapshot()).unwrap();
    assert_eq!(&parse_bytes(&bytes).unwrap(), commit.snapshot());
}
