use super::codec::PLCOCX;
use super::{Controls, Document};
use crate::parts::fib::FileInformationBlock;

const FIB_POINTERS: usize = 93;

fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
    let declared = u16::from_le_bytes([fib[152], fib[153]]);
    let count = declared.max(u16::try_from(index + 1).unwrap());
    fib[152..154].copy_from_slice(&count.to_le_bytes());
    let start = 154 + index * 8;
    fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
    fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
}

fn rgx_ocx_info(cookies: &[u32]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(cookies.len() as u32).to_le_bytes());
    for cookie in cookies {
        data.extend_from_slice(&cookie.to_le_bytes());
    }
    data
}

fn fixture(cookies: &[u32]) -> (FileInformationBlock, Vec<u8>) {
    let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
    fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
    fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());

    let table = rgx_ocx_info(cookies);
    set_fib_pointer(&mut fib_data, PLCOCX, 0, table.len() as u32);
    (FileInformationBlock::parse(&fib_data).unwrap(), table)
}

#[test]
fn parses_ole_controls() {
    let (fib, table) = fixture(&[3, 0, 17]);
    let parsed = Controls::parse(&fib, &table)
        .unwrap()
        .expect("controls present");
    assert_eq!(parsed.len(), 3);
    assert!(!parsed.is_empty());
    assert_eq!(
        parsed
            .controls()
            .iter()
            .map(|control| control.cookie)
            .collect::<Vec<_>>(),
        [3, 0, 17]
    );
}

#[test]
fn absent_or_empty_table_yields_none() {
    // No `fcPlcocx` pointer at all.
    let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
    fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
    fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());
    let fib = FileInformationBlock::parse(&fib_data).unwrap();
    assert!(Controls::parse(&fib, &[]).unwrap().is_none());

    // A zero-length pointer is ignored per MS-DOC 2.5 (fcPlcocx).
    let (fib, table) = fixture(&[]);
    let mut fib_data = fib.raw_data().to_vec();
    set_fib_pointer(&mut fib_data, PLCOCX, 0, 0);
    let fib = FileInformationBlock::parse(&fib_data).unwrap();
    assert!(Controls::parse(&fib, &table).unwrap().is_none());
}

#[test]
fn parses_empty_control_list() {
    let (fib, table) = fixture(&[]);
    let parsed = Controls::parse(&fib, &table)
        .unwrap()
        .expect("empty table present");
    assert!(parsed.is_empty());
}

#[test]
fn accepts_word_padded_entries() {
    // Word 2003+ pads each `OcxInfo` beyond `dwCookie`; here 20 bytes per
    // entry, as emitted for the travel-form fixture.
    let mut data = Vec::new();
    data.extend_from_slice(&1u32.to_le_bytes());
    data.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    data.extend_from_slice(&[0xFF, 0xFF, 0xFF, 0xFF]);
    data.extend_from_slice(&[0; 4]);
    data.extend_from_slice(&[0; 2]);
    data.extend_from_slice(&[1, 0]);
    data.extend_from_slice(&[0; 4]);
    let parsed = Controls::parse_bytes(&data).unwrap();
    assert_eq!(parsed.controls()[0].cookie, 0xFFFF_FFFF);
    assert!(parsed.controls()[0].metadata.is_some());
    assert_eq!(parsed.entry_stride(), 20);
    assert_eq!(
        parsed.entry_padding(0),
        Some(&[0xFF, 0xFF, 0xFF, 0xFF, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0][..])
    );
    assert_eq!(parsed.to_bytes(), data);
}

#[test]
fn decodes_the_specified_ocx_info_body() {
    let mut data = Vec::new();
    data.extend_from_slice(&1u32.to_le_bytes());
    data.extend_from_slice(&17u32.to_le_bytes());
    data.extend_from_slice(&29u32.to_le_bytes());
    data.extend_from_slice(&0xAABB_CCDDu32.to_le_bytes());
    data.extend_from_slice(&3u16.to_le_bytes());
    data.extend_from_slice(&0x22CBu16.to_le_bytes());
    data.extend_from_slice(&8u16.to_le_bytes());
    data.extend_from_slice(&0x3344u16.to_le_bytes());

    let parsed = Controls::parse_bytes(&data).unwrap();
    let metadata = parsed.controls()[0].metadata.unwrap();
    assert_eq!(metadata.field_index, 29);
    assert_eq!(metadata.accelerator_handle, 0xAABB_CCDD);
    assert_eq!(metadata.accelerator_count, 3);
    assert!(metadata.flags.eats_return);
    assert!(metadata.flags.default_button);
    assert!(metadata.flags.right_to_left);
    assert!(metadata.flags.corrupt);
    assert_eq!(metadata.document, Document::HeaderTextbox);
    assert_eq!(parsed.to_bytes(), data);
}

#[test]
fn rejects_malformed_tables() {
    // Duplicate cookies.
    assert!(Controls::parse_bytes(&rgx_ocx_info(&[3, 3])).is_err());

    // Declared count disagrees with the byte length.
    assert!(Controls::parse_bytes(&rgx_ocx_info(&[3])[..6]).is_err());
    let mut misaligned = rgx_ocx_info(&[3, 4]);
    misaligned.pop();
    assert!(Controls::parse_bytes(&misaligned).is_err());

    // A nonzero count with no entries, including a hostile count that must
    // be rejected before any capacity reservation.
    assert!(Controls::parse_bytes(&1u32.to_le_bytes()).is_err());
    assert!(Controls::parse_bytes(&u32::MAX.to_le_bytes()).is_err());

    // A zero count with trailing bytes.
    assert!(Controls::parse_bytes(&[0; 8]).is_err());

    // Truncated header.
    assert!(Controls::parse_bytes(&[0, 0]).is_err());
}
