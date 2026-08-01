//! Round-trip tests for the BIFF8 PhoneticInfo record.

use litchi_xls::writer::XlsWriter;
use litchi_xls::{
    XlsPhoneticAlignment, XlsPhoneticFormat, XlsPhoneticInfo, XlsPhoneticRange, XlsPhoneticType,
    XlsWorkbook,
};
use std::io::Cursor;

fn written_workbook(phonetic_info: Option<XlsPhoneticInfo>) -> Vec<u8> {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Phonetic").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_phonetic_info(sheet, phonetic_info).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn phonetic_info_round_trips() {
    let info = XlsPhoneticInfo::try_new(
        XlsPhoneticFormat::new(4, XlsPhoneticType::Hiragana, XlsPhoneticAlignment::Center),
        vec![
            XlsPhoneticRange::new(1, 3, 0, 5).unwrap(),
            XlsPhoneticRange::new(7, 7, 2, 2).unwrap(),
        ],
    )
    .unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(written_workbook(Some(info.clone())))).unwrap();
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().phonetic_info(),
        Some(&info)
    );
}

#[test]
fn long_range_list_round_trips_across_continues() {
    // 4,000 ranges exceed one BIFF8 record and force Continue chunking.
    let ranges = (0..4_000u16)
        .map(|row| XlsPhoneticRange::new(row, row, 0, 4).unwrap())
        .collect::<Vec<_>>();
    let info = XlsPhoneticInfo::try_new(
        XlsPhoneticFormat::new(0, XlsPhoneticType::Any, XlsPhoneticAlignment::General),
        ranges,
    )
    .unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(written_workbook(Some(info.clone())))).unwrap();
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().phonetic_info(),
        Some(&info)
    );
}

#[test]
fn worksheet_without_phonetic_info_has_none() {
    let workbook = XlsWorkbook::new(Cursor::new(written_workbook(None))).unwrap();
    assert!(workbook.xls_worksheet(0).unwrap().phonetic_info().is_none());
}
