use litchi_rtf::{DocumentInfo, RtfDocument, RtfTimestamp, RtfWriter};
use std::borrow::Cow;
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_complete_typed_info_and_round_trips_without_body_leakage() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\info{\title Title}{\subject Subject}{\author Ada}{\manager Grace}"#,
        r#"{\company Example}{\operator Linus}{\category Test}{\keywords one;two}"#,
        r#"{\comment Summary}{\doccomm Document note}{\hlinkbase https://example.test/base/}"#,
        r#"{\creatim\yr2024\mo2\dy29\hr23\min59\sec58}{\revtim\yr2025\mo7}"#,
        r#"{\printim\yr2026\mo1\dy2}{\buptim\yr2023}"#,
        r#"\version4\vern9\edmins120\nofpages8\nofwords900\nofchars4200\nofcharsws5000\id77}Body}"#,
    )).unwrap();
    let info = document.info();
    assert_eq!(info.comment.as_deref(), Some("Summary"));
    assert_eq!(info.document_comment.as_deref(), Some("Document note"));
    assert_eq!(info.hyperlink_base.as_deref(), Some("https://example.test/base/"));
    assert_eq!(info.creation_timestamp.unwrap().day, Some(29));
    assert_eq!(info.revision_timestamp.unwrap().day, None);
    assert_eq!(info.backup_timestamp.unwrap().month, None);
    assert_eq!(info.pages, Some(8));
    assert_eq!(document.text(), "Body");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.info(), info);
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn writer_accepts_typed_and_legacy_timestamps_and_rejects_unsafe_values() {
    let mut output = Vec::new();
    let mut writer = RtfWriter::new(&mut output);
    writer.write_document_header().unwrap();
    let mut info = DocumentInfo::new();
    info.creation_timestamp = Some(RtfTimestamp {
        year: Some(2024), month: Some(2), day: Some(29),
        hour: None, minute: None, second: None,
    });
    info.revision_time = Some(Cow::Borrowed("2025-07-16T12:34:56"));
    writer.write_document_info(&info).unwrap();
    writer.write_str("}").unwrap();
    let parsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(parsed.info().creation_timestamp.unwrap().hour, None);
    assert_eq!(parsed.info().revision_timestamp.unwrap().minute, Some(34));

    let mut invalid = DocumentInfo::new();
    invalid.pages = Some(i32::MAX as u32 + 1);
    assert!(RtfWriter::new(Vec::new()).write_document_info(&invalid).is_err());
    invalid.pages = None;
    invalid.creation_timestamp = Some(RtfTimestamp {
        year: Some(2023), month: Some(2), day: Some(29),
        hour: None, minute: None, second: None,
    });
    assert!(RtfWriter::new(Vec::new()).write_document_info(&invalid).is_err());
}

#[test]
fn rejects_ambiguous_or_malformed_known_info_metadata() {
    let cases = [
        r#"{\rtf1{\info{\title A}{\title B}}}"#,
        r#"{\rtf1{\info\nofpages-1}}"#,
        r#"{\rtf1{\info\nofwords1\nofwords2}}"#,
        r#"{\rtf1{\info active text}}"#,
        r#"{\rtf1{\info}{\info}}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }

    let invalid_calendar = RtfDocument::parse(
        r#"{\rtf1{\info{\creatim\yr2023\mo2\dy29}{\revtim\hr24}}}"#,
    ).unwrap();
    assert!(!invalid_calendar.info().creation_timestamp.unwrap().is_valid());
    assert!(!invalid_calendar.info().revision_timestamp.unwrap().is_valid());
}

#[test]
fn parses_real_libreoffice_info_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/fdo80924.rtf",
        "sw/qa/extras/ooxmlexport/data/ooo39250-1-min.rtf",
        "sw/qa/extras/ooxmlexport/data/tdf154703_framePr2.rtf",
    ];
    let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../../3rdparty/libreoffice-core/");
    for (index, fixture) in FIXTURES.iter().enumerate() {
        let bytes = fs::read(format!("{root}{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(document.info().document_comment.is_some());
        if index > 0 {
            assert!(document.info().creation_timestamp.is_some());
            assert!(document.info().revision_timestamp.is_some());
        }
    }
}

#[test]
fn preserves_libreoffice_zero_timestamp_sentinels() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/text-change-tracking.rtf"
    );
    let bytes = fs::read(path).unwrap();
    let document = RtfDocument::parse_bytes(&bytes).unwrap();
    let revision = document.info().revision_timestamp.unwrap();
    let printed = document.info().print_timestamp.unwrap();

    assert_eq!((revision.year, revision.month, revision.day), (Some(0), Some(0), Some(0)));
    assert_eq!((printed.year, printed.month, printed.day), (Some(0), Some(0), Some(0)));
    assert!(!revision.is_valid());
    assert!(!printed.is_valid());

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.info().revision_timestamp, Some(revision));
    assert_eq!(reparsed.info().print_timestamp, Some(printed));
}
