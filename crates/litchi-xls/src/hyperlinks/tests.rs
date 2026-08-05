//! Regression coverage for the BIFF8/MS-XLS hyperlink owner.

use super::codec::{FILE_MONIKER_CLSID, URL_MONIKER_CLSID, URL_SERIAL_GUID, parse_hlink_record};
use super::model::{RECORD_TYPE, TOOLTIP_RECORD_TYPE, XlsHyperlinkMoniker, XlsHyperlinkTargetKind};
use super::package::HyperlinkCollector;

const TEST_HLINK_CLSID: [u8; 16] = [
    0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];

fn range(data: &mut Vec<u8>, row: u16, column: u16) {
    for value in [row, row, column, column] {
        data.extend_from_slice(&value.to_le_bytes());
    }
}
fn string(data: &mut Vec<u8>, value: &str) {
    data.extend_from_slice(&(value.encode_utf16().count() as u32 + 1).to_le_bytes());
    for unit in value.encode_utf16().chain(std::iter::once(0)) {
        data.extend_from_slice(&unit.to_le_bytes());
    }
}
fn base(flags: u32) -> Vec<u8> {
    let mut data = Vec::new();
    range(&mut data, 4, 0);
    data.extend_from_slice(&TEST_HLINK_CLSID);
    data.extend_from_slice(&2u32.to_le_bytes());
    data.extend_from_slice(&flags.to_le_bytes());
    data
}
fn url_link() -> Vec<u8> {
    let mut data = base(0x17);
    string(&mut data, "Example");
    data.extend_from_slice(&URL_MONIKER_CLSID);
    let mut url = Vec::new();
    for u in "https://example.com"
        .encode_utf16()
        .chain(std::iter::once(0))
    {
        url.extend_from_slice(&u.to_le_bytes());
    }
    data.extend_from_slice(&(url.len() as u32 + 24).to_le_bytes());
    data.extend_from_slice(&url);
    data.extend_from_slice(&URL_SERIAL_GUID);
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&0xABA5u32.to_le_bytes());
    data
}

#[test]
fn parses_url_email_document_file_and_string_monikers() {
    let url = parse_hlink_record(&url_link()).unwrap();
    assert_eq!(url.target_kind(), XlsHyperlinkTargetKind::Url);
    assert_eq!(url.address(), Some("https://example.com"));
    let mut document = base(0x1C);
    string(&mut document, "place");
    string(&mut document, "Sheet1!A1");
    let document = parse_hlink_record(&document).unwrap();
    assert_eq!(document.target_kind(), XlsHyperlinkTargetKind::Document);
    assert_eq!(document.location(), Some("Sheet1!A1"));
    let mut file = base(1);
    file.extend_from_slice(&FILE_MONIKER_CLSID);
    file.extend_from_slice(&0u16.to_le_bytes());
    file.extend_from_slice(&9u32.to_le_bytes());
    file.extend_from_slice(b"file.xls\0");
    file.extend_from_slice(&0xFFFFu16.to_le_bytes());
    file.extend_from_slice(&0xDEADu16.to_le_bytes());
    file.extend_from_slice(&[0; 16]);
    file.extend_from_slice(&0u32.to_le_bytes());
    file.extend_from_slice(&0u32.to_le_bytes());
    let file = parse_hlink_record(&file).unwrap();
    assert_eq!(file.target_kind(), XlsHyperlinkTargetKind::File);
    assert_eq!(file.address(), Some("file.xls"));
    let mut unc = base(0x101);
    string(&mut unc, "\\\\server\\share");
    let unc = parse_hlink_record(&unc).unwrap();
    assert_eq!(unc.target_kind(), XlsHyperlinkTargetKind::Unc);
}

#[test]
fn links_only_an_immediately_following_matching_tooltip() {
    let mut collector = HyperlinkCollector::new();
    collector.feed_record(RECORD_TYPE, &url_link()).unwrap();
    let mut tooltip = Vec::new();
    tooltip.extend_from_slice(&TOOLTIP_RECORD_TYPE.to_le_bytes());
    range(&mut tooltip, 4, 0);
    for u in "Open site".encode_utf16().chain(std::iter::once(0)) {
        tooltip.extend_from_slice(&u.to_le_bytes());
    }
    collector
        .feed_record(TOOLTIP_RECORD_TYPE, &tooltip)
        .unwrap();
    assert_eq!(collector.finish()[0].tooltip(), Some("Open site"));
    let mut collector = HyperlinkCollector::new();
    assert!(
        collector
            .feed_record(TOOLTIP_RECORD_TYPE, &tooltip)
            .is_err()
    );
}

#[test]
fn rejects_bad_flags_range_and_url_tail() {
    let mut data = url_link();
    data[31] = 0x80;
    assert!(parse_hlink_record(&data).is_err());
    let mut data = url_link();
    data[0..2].copy_from_slice(&5u16.to_le_bytes());
    data[2..4].copy_from_slice(&4u16.to_le_bytes());
    assert!(parse_hlink_record(&data).is_err());
    let mut data = url_link();
    let last = data.len() - 24;
    data[last] ^= 1;
    assert!(parse_hlink_record(&data).is_err());
}

#[test]
fn accepts_and_retains_nonstandard_hlink_producer_clsid() {
    let mut data = url_link();
    let producer_clsid = [0xA5; 16];
    data[8..24].copy_from_slice(&producer_clsid);

    let link = parse_hlink_record(&data).unwrap();

    assert_eq!(link.class_id(), &producer_clsid);
    assert_eq!(link.address(), Some("https://example.com"));
}

#[test]
fn reads_poi_hyperlink_fixtures() {
    use crate::XlsWorkbook;
    use std::fs::File;
    use std::path::Path;
    let fixture = |name: &str| {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    };
    let workbook = XlsWorkbook::new(File::open(fixture("WithTwoHyperLinks.xls")).unwrap()).unwrap();
    let links = workbook.xls_worksheet(0).unwrap().hyperlinks();
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].range().first_row(), 4);
    assert_eq!(links[0].display_name(), Some("Foo"));
    assert_eq!(links[0].address(), Some("http://poi.apache.org/"));
    assert_eq!(links[1].range().first_column(), 1);
    let workbook =
        XlsWorkbook::new(File::open(fixture("HyperlinksOnManySheets.xls")).unwrap()).unwrap();
    assert_eq!(workbook.xls_worksheet(0).unwrap().hyperlinks().len(), 2);
    let email = &workbook.xls_worksheet(1).unwrap().hyperlinks()[0];
    assert_eq!(email.target_kind(), XlsHyperlinkTargetKind::Email);
    assert_eq!(email.address(), Some("mailto:dev@poi.apache.org"));
    let document = &workbook.xls_worksheet(2).unwrap().hyperlinks()[0];
    assert_eq!(document.target_kind(), XlsHyperlinkTargetKind::Document);
    assert_eq!(document.location(), Some("WebLinks!A1"));
}

#[test]
fn test_parse_hlink_record_too_short() {
    assert!(parse_hlink_record(&[0; 10]).is_err());
}

#[test]
fn test_parse_hlink_record_invalid_version() {
    let mut data = base(0x08);
    data[24..28].copy_from_slice(&99u32.to_le_bytes());
    string(&mut data, "Sheet1!A1");
    assert!(parse_hlink_record(&data).is_err());
}

#[test]
fn test_hyperlink_target_url() {
    let link = parse_hlink_record(&url_link()).unwrap();
    assert_eq!(link.target_kind(), XlsHyperlinkTargetKind::Url);
    assert_eq!(link.address(), Some("https://example.com"));
    assert!(link.absolute());
}

#[test]
fn test_hyperlink_target_document() {
    let mut data = base(0x08);
    string(&mut data, "Sheet1!A1");
    let link = parse_hlink_record(&data).unwrap();
    assert_eq!(link.target_kind(), XlsHyperlinkTargetKind::Document);
    assert_eq!(link.address(), Some("Sheet1!A1"));
}

#[test]
fn test_hyperlink_target_unc() {
    let mut data = base(0x101);
    string(&mut data, "\\\\server\\share\\file.txt");
    let link = parse_hlink_record(&data).unwrap();
    assert_eq!(link.target_kind(), XlsHyperlinkTargetKind::Unc);
    assert_eq!(link.address(), Some("\\\\server\\share\\file.txt"));
}

#[test]
fn test_hyperlink_target_file_with_long_name() {
    let mut data = base(0x01);
    data.extend_from_slice(&FILE_MONIKER_CLSID);
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&13u32.to_le_bytes());
    data.extend_from_slice(b"LONGFI~1.TXT\0");
    data.extend_from_slice(&0xFFFFu16.to_le_bytes());
    data.extend_from_slice(&0xDEADu16.to_le_bytes());
    data.extend_from_slice(&[0; 16]);
    data.extend_from_slice(&0u32.to_le_bytes());
    let unicode: Vec<u8> = "long_filename.txt"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    data.extend_from_slice(&(unicode.len() as u32 + 6).to_le_bytes());
    data.extend_from_slice(&(unicode.len() as u32).to_le_bytes());
    data.extend_from_slice(&3u16.to_le_bytes());
    data.extend_from_slice(&unicode);
    let link = parse_hlink_record(&data).unwrap();
    assert_eq!(link.address(), Some("long_filename.txt"));
    let XlsHyperlinkMoniker::File(file) = link.moniker().unwrap() else {
        panic!()
    };
    assert_eq!(file.ansi_path(), "LONGFI~1.TXT");
    assert_eq!(file.unicode_path(), Some("long_filename.txt"));
}

#[test]
fn test_hyperlink_target_file_without_long_name() {
    let mut data = base(0x01);
    data.extend_from_slice(&FILE_MONIKER_CLSID);
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&9u32.to_le_bytes());
    data.extend_from_slice(b"FILE.TXT\0");
    data.extend_from_slice(&0xFFFFu16.to_le_bytes());
    data.extend_from_slice(&0xDEADu16.to_le_bytes());
    data.extend_from_slice(&[0; 16]);
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    let link = parse_hlink_record(&data).unwrap();
    assert_eq!(link.address(), Some("FILE.TXT"));
}

#[test]
fn test_xls_hyperlink_clone() {
    let link = parse_hlink_record(&url_link()).unwrap();
    assert_eq!(link.clone(), link);
    assert_eq!(
        link.moniker().unwrap().clone(),
        link.moniker().unwrap().clone()
    );
}

#[test]
fn test_xls_hyperlink_debug() {
    let link = parse_hlink_record(&url_link()).unwrap();
    let debug = format!("{link:?}");
    assert!(debug.contains("XlsHyperlink"));
    assert!(debug.contains("https://example.com"));
}

#[test]
fn test_record_type_constant() {
    assert_eq!(RECORD_TYPE, 0x01B8);
    assert_eq!(TOOLTIP_RECORD_TYPE, 0x0800);
}
