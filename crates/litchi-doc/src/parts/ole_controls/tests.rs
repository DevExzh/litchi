use super::{FIB_INDEX_PLC_OCX, Flags, OcxInfo, RgxOcxInfo, Story, parse, parse_bytes, to_bytes};
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

    let invalid_empty = fib_with_pointer(1, 0);
    assert!(parse(&invalid_empty, &[0]).is_err());

    let truncated = fib_with_pointer(4, payload.len() as u32);
    assert!(parse(&truncated, &table_stream).is_err());
}
