#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

//! Round-trip tests for extended section page numbering: the full RTF 1.9.1
//! `\pgn*` format family, `\pgnrestart`/`\pgncont`, and `\pgnxN`/`\pgnyN`.

use litchi_rtf::{PageNumberFormat, PageNumberRestart, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(document).unwrap();
    String::from_utf8(bytes).unwrap()
}

fn round_trip(format: PageNumberFormat) {
    let source = format!(r"{{\rtf1\ansi\sectd\{} Body\par}}", format.control_word());
    let document = RtfDocument::parse(&source).unwrap();
    assert_eq!(document.sections()[0].properties.page_number_format, format);

    let first = write(&document);
    assert!(
        first.contains(format!("\\{}", format.control_word()).as_str()),
        "missing {} in {first}",
        format.control_word()
    );
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(reparsed.sections()[0].properties.page_number_format, format);
    assert_eq!(first, write(&reparsed));
}

#[test]
fn latin_page_number_formats_round_trip() {
    for format in [
        PageNumberFormat::Decimal,
        PageNumberFormat::UpperRoman,
        PageNumberFormat::LowerRoman,
        PageNumberFormat::UpperLetter,
        PageNumberFormat::LowerLetter,
    ] {
        round_trip(format);
    }
}

#[test]
fn cjk_and_indic_page_number_formats_round_trip() {
    for format in [
        PageNumberFormat::BidiAlphabetic,
        PageNumberFormat::BidiAbjad,
        PageNumberFormat::KoreanChosung,
        PageNumberFormat::Circle,
        PageNumberFormat::KanjiDigitless,
        PageNumberFormat::KanjiWithDigit,
        PageNumberFormat::KanjiThree,
        PageNumberFormat::KanjiFour,
        PageNumberFormat::DoubleDecimal,
        PageNumberFormat::KoreanGanada,
        PageNumberFormat::ChineseOne,
        PageNumberFormat::ChineseTwo,
        PageNumberFormat::ChineseThree,
        PageNumberFormat::ChineseFour,
        PageNumberFormat::HindiVowels,
        PageNumberFormat::HindiConsonants,
        PageNumberFormat::HindiNumbers,
        PageNumberFormat::HindiDescriptive,
        PageNumberFormat::ThaiLetters,
        PageNumberFormat::ThaiNumbers,
        PageNumberFormat::ThaiDescriptive,
        PageNumberFormat::VietnameseCardinal,
        PageNumberFormat::ZodiacOne,
        PageNumberFormat::ZodiacTwo,
        PageNumberFormat::ZodiacThree,
    ] {
        round_trip(format);
    }
}

#[test]
fn page_number_restart_and_offsets_round_trip() {
    let source = r"{\rtf1\ansi\sectd\pgnstarts7\pgnrestart\pgnx120\pgny-240 Body\par}";
    let document = RtfDocument::parse(source).unwrap();
    let properties = &document.sections()[0].properties;
    assert_eq!(properties.page_number_start, 7);
    assert_eq!(
        properties.page_number_restart,
        Some(PageNumberRestart::Restart)
    );
    assert_eq!(properties.page_number_offset_x, Some(120));
    assert_eq!(properties.page_number_offset_y, Some(-240));

    let first = write(&document);
    for expected in ["\\pgnstarts7", "\\pgnrestart", "\\pgnx120", "\\pgny-240"] {
        assert!(first.contains(expected), "missing {expected} in {first}");
    }
    let reparsed = RtfDocument::parse(&first).unwrap();
    let properties = &reparsed.sections()[0].properties;
    assert_eq!(
        properties.page_number_restart,
        Some(PageNumberRestart::Restart)
    );
    assert_eq!(properties.page_number_offset_x, Some(120));
    assert_eq!(properties.page_number_offset_y, Some(-240));
    assert_eq!(first, write(&reparsed));
}

#[test]
fn page_number_continue_round_trips() {
    let source = r"{\rtf1\ansi\sectd\pgncont Body\par}";
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(
        document.sections()[0].properties.page_number_restart,
        Some(PageNumberRestart::Continuous)
    );

    let first = write(&document);
    assert!(first.contains("\\pgncont"), "missing pgncont in {first}");
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(
        reparsed.sections()[0].properties.page_number_restart,
        Some(PageNumberRestart::Continuous)
    );
}

#[test]
fn omitted_page_number_controls_stay_omitted() {
    let source = r"{\rtf1\ansi\sectd Body\par}";
    let document = RtfDocument::parse(source).unwrap();
    let properties = &document.sections()[0].properties;
    assert_eq!(properties.page_number_restart, None);
    assert_eq!(properties.page_number_offset_x, None);
    assert_eq!(properties.page_number_offset_y, None);

    let first = write(&document);
    assert!(!first.contains("\\pgnrestart"), "unexpected in {first}");
    assert!(!first.contains("\\pgncont"), "unexpected in {first}");
    assert!(!first.contains("\\pgnx"), "unexpected in {first}");
    assert!(!first.contains("\\pgny"), "unexpected in {first}");
}
