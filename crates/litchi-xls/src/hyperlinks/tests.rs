//! Regression coverage for the BIFF8/MS-XLS hyperlink owner.

use super::codec::{
    FILE_MONIKER_CLSID, URL_MONIKER_CLSID, URL_SERIAL_GUID, parse_hlink_record, parse_tooltip,
};
use super::model::{HyperlinkMoniker, HyperlinkTargetKind, RECORD_TYPE, TOOLTIP_RECORD_TYPE};
use super::package::HyperlinkCollector;
use crate::SourceLimits;

fn limits() -> SourceLimits {
    SourceLimits::default()
}

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
fn tooltip(row: u16, column: u16, value: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&TOOLTIP_RECORD_TYPE.to_le_bytes());
    range(&mut data, row, column);
    for unit in value.encode_utf16().chain(std::iter::once(0)) {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
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
    let url = parse_hlink_record(&url_link(), &limits()).unwrap();
    assert_eq!(url.target_kind(), HyperlinkTargetKind::Url);
    assert_eq!(url.address(), Some("https://example.com"));
    let mut document = base(0x1C);
    string(&mut document, "place");
    string(&mut document, "Sheet1!A1");
    let document = parse_hlink_record(&document, &limits()).unwrap();
    assert_eq!(document.target_kind(), HyperlinkTargetKind::Document);
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
    let file = parse_hlink_record(&file, &limits()).unwrap();
    assert_eq!(file.target_kind(), HyperlinkTargetKind::File);
    assert_eq!(file.address(), Some("file.xls"));
    let mut unc = base(0x101);
    string(&mut unc, "\\\\server\\share");
    let unc = parse_hlink_record(&unc, &limits()).unwrap();
    assert_eq!(unc.target_kind(), HyperlinkTargetKind::Unc);
}

#[test]
fn links_only_an_immediately_following_matching_tooltip() {
    let mut collector = HyperlinkCollector::new(limits());
    collector.feed_record(RECORD_TYPE, &url_link()).unwrap();
    let tooltip = tooltip(4, 0, "Open site");
    collector
        .feed_record(TOOLTIP_RECORD_TYPE, &tooltip)
        .unwrap();
    assert_eq!(collector.finish()[0].tooltip(), Some("Open site"));
    let mut collector = HyperlinkCollector::new(limits());
    assert!(
        collector
            .feed_record(TOOLTIP_RECORD_TYPE, &tooltip)
            .is_err()
    );
}

#[test]
fn tooltip_parser_rejects_payload_above_selected_entry_limit_before_utf16_decode() {
    let data = tooltip(4, 0, "012345678901234567890123456789");
    let limits = SourceLimits::builder()
        .max_entry_bytes(data.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        parse_tooltip(&data, &limits),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::EntryBytes,
            observed,
            maximum,
        }) if observed == data.len() as u64 && maximum == (data.len() - 1) as u64
    ));
}

#[test]
fn collector_applies_selected_entry_limit_to_tooltip_before_decode() {
    let mut link = base(0x08);
    string(&mut link, "A");
    let data = tooltip(4, 0, "012345678901234567890123456789");
    let limits = SourceLimits::builder()
        .max_entry_bytes(link.len())
        .unwrap()
        .build()
        .unwrap();
    let mut collector = HyperlinkCollector::new(limits);
    collector.feed_record(RECORD_TYPE, &link).unwrap();
    assert!(matches!(
        collector.feed_record(TOOLTIP_RECORD_TYPE, &data),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::EntryBytes,
            observed,
            maximum,
        }) if observed == data.len() as u64 && maximum == link.len() as u64
    ));
}

#[test]
fn rejects_bad_flags_range_and_url_tail() {
    let mut data = url_link();
    data[31] = 0x80;
    assert!(parse_hlink_record(&data, &limits()).is_err());
    let mut data = url_link();
    data[0..2].copy_from_slice(&5u16.to_le_bytes());
    data[2..4].copy_from_slice(&4u16.to_le_bytes());
    assert!(parse_hlink_record(&data, &limits()).is_err());
    let mut data = url_link();
    let last = data.len() - 24;
    data[last] ^= 1;
    assert!(parse_hlink_record(&data, &limits()).is_err());
}

#[test]
fn accepts_and_retains_nonstandard_hlink_producer_clsid() {
    let mut data = url_link();
    let producer_clsid = [0xA5; 16];
    data[8..24].copy_from_slice(&producer_clsid);

    let link = parse_hlink_record(&data, &limits()).unwrap();

    assert_eq!(link.class_id(), &producer_clsid);
    assert_eq!(link.address(), Some("https://example.com"));
}

#[test]
fn reads_poi_hyperlink_fixtures() {
    use crate::raw::Reader;
    use std::fs::File;
    use std::path::Path;
    let fixture = |name: &str| {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    };
    let workbook = Reader::new(File::open(fixture("WithTwoHyperLinks.xls")).unwrap()).unwrap();
    let links = workbook.xls_worksheet(0).unwrap().hyperlinks();
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].range().first_row(), 4);
    assert_eq!(links[0].display_name(), Some("Foo"));
    assert_eq!(links[0].address(), Some("http://poi.apache.org/"));
    assert_eq!(links[1].range().first_column(), 1);
    let workbook = Reader::new(File::open(fixture("HyperlinksOnManySheets.xls")).unwrap()).unwrap();
    assert_eq!(workbook.xls_worksheet(0).unwrap().hyperlinks().len(), 2);
    let email = &workbook.xls_worksheet(1).unwrap().hyperlinks()[0];
    assert_eq!(email.target_kind(), HyperlinkTargetKind::Email);
    assert_eq!(email.address(), Some("mailto:dev@poi.apache.org"));
    let document = &workbook.xls_worksheet(2).unwrap().hyperlinks()[0];
    assert_eq!(document.target_kind(), HyperlinkTargetKind::Document);
    assert_eq!(document.location(), Some("WebLinks!A1"));
}

#[test]
fn test_parse_hlink_record_too_short() {
    assert!(parse_hlink_record(&[0; 10], &limits()).is_err());
}

#[test]
fn test_parse_hlink_record_invalid_version() {
    let mut data = base(0x08);
    data[24..28].copy_from_slice(&99u32.to_le_bytes());
    string(&mut data, "Sheet1!A1");
    assert!(parse_hlink_record(&data, &limits()).is_err());
}

#[test]
fn test_hyperlink_target_url() {
    let link = parse_hlink_record(&url_link(), &limits()).unwrap();
    assert_eq!(link.target_kind(), HyperlinkTargetKind::Url);
    assert_eq!(link.address(), Some("https://example.com"));
    assert!(link.absolute());
}

#[test]
fn test_hyperlink_target_document() {
    let mut data = base(0x08);
    string(&mut data, "Sheet1!A1");
    let link = parse_hlink_record(&data, &limits()).unwrap();
    assert_eq!(link.target_kind(), HyperlinkTargetKind::Document);
    assert_eq!(link.address(), Some("Sheet1!A1"));
}

#[test]
fn test_hyperlink_target_unc() {
    let mut data = base(0x101);
    string(&mut data, "\\\\server\\share\\file.txt");
    let link = parse_hlink_record(&data, &limits()).unwrap();
    assert_eq!(link.target_kind(), HyperlinkTargetKind::Unc);
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
    let link = parse_hlink_record(&data, &limits()).unwrap();
    assert_eq!(link.address(), Some("long_filename.txt"));
    let HyperlinkMoniker::File(file) = link.moniker().unwrap() else {
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
    let link = parse_hlink_record(&data, &limits()).unwrap();
    assert_eq!(link.address(), Some("FILE.TXT"));
}

#[test]
fn test_xls_hyperlink_clone() {
    let link = parse_hlink_record(&url_link(), &limits()).unwrap();
    assert_eq!(link.clone(), link);
    assert_eq!(
        link.moniker().unwrap().clone(),
        link.moniker().unwrap().clone()
    );
}

#[test]
fn test_xls_hyperlink_debug() {
    let link = parse_hlink_record(&url_link(), &limits()).unwrap();
    let debug = format!("{link:?}");
    assert!(debug.contains("Hyperlink"));
    assert!(debug.contains("https://example.com"));
}

#[test]
fn test_record_type_constant() {
    assert_eq!(RECORD_TYPE, 0x01B8);
    assert_eq!(TOOLTIP_RECORD_TYPE, 0x0800);
}

/// A URL without the optional serialization tail whose own last characters
/// spell the tail's GUID is still a URL: the tail follows the URL's NUL.
#[test]
fn a_url_ending_in_the_serialization_guid_is_not_a_tail() {
    let guid_units = URL_SERIAL_GUID
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes(*pair))
        .collect::<Vec<_>>();
    let mut units = "https://x/".encode_utf16().collect::<Vec<_>>();
    units.extend_from_slice(&guid_units);
    units.extend_from_slice(&[u16::from(b'a'), u16::from(b'b'), u16::from(b'c')]);
    let expected = String::from_utf16(&units).unwrap();

    let mut data = base(0x03);
    data.extend_from_slice(&URL_MONIKER_CLSID);
    let mut url = Vec::new();
    for unit in units.iter().copied().chain(std::iter::once(0)) {
        url.extend_from_slice(&unit.to_le_bytes());
    }
    data.extend_from_slice(&(url.len() as u32).to_le_bytes());
    data.extend_from_slice(&url);

    let link = parse_hlink_record(&data, &limits()).unwrap();
    assert_eq!(link.address(), Some(expected.as_str()));
    // The same URL with a real tail still parses its URI flags.
    assert_eq!(
        parse_hlink_record(&url_link(), &limits())
            .unwrap()
            .moniker()
            .map(|moniker| matches!(moniker, HyperlinkMoniker::Url(url) if url.serialization_uri_flags() == Some(0xABA5))),
        Some(true)
    );
}

#[test]
fn parser_rejects_payload_above_selected_entry_limit_before_string_decoding() {
    let data = url_link();
    let limits = SourceLimits::builder()
        .max_entry_bytes(data.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        parse_hlink_record(&data, &limits),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::EntryBytes,
            observed,
            maximum,
        }) if observed == data.len() as u64 && maximum == (data.len() - 1) as u64
    ));
}
