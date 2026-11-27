use super::{Clock, Instruction, apply_case, format_number};

#[test]
fn instructions_split_into_arguments_and_switches() {
    let toc = Instruction::parse(r#" TOC \o "1-3" \h \z \u \n \p "-" "#);
    assert_eq!(toc.kind, "TOC");
    assert_eq!(toc.switch('o'), Some("1-3"));
    assert!(toc.has('h') && toc.has('z') && toc.has('u') && toc.has('n'));
    assert_eq!(toc.switch('n'), None);
    assert_eq!(toc.switch('p'), Some("-"));
    let seq = Instruction::parse(r"seq Figure \* ROMAN \r 3 \* MERGEFORMAT");
    assert_eq!(seq.kind, "SEQ");
    assert_eq!(seq.arguments, ["Figure"]);
    assert_eq!(seq.switch('r'), Some("3"));
    assert_eq!(seq.formats().collect::<Vec<_>>(), ["ROMAN"]);
    let property = Instruction::parse(r#"DOCPROPERTY "Last Saved By""#);
    assert_eq!(property.arguments, ["Last Saved By"]);
    let quoted = Instruction::parse(r#"DATE \@ "dddd, MMMM d""#);
    assert_eq!(quoted.switch('@'), Some("dddd, MMMM d"));
}

#[test]
fn clocks_are_explicit_and_validated() {
    let clock = Clock::parse("2024-02-29T13:07:09").unwrap();
    assert_eq!(clock.format("M/d/yyyy"), "2/29/2024");
    assert_eq!(clock.format("yy-MM-dd HH:mm:ss"), "24-02-29 13:07:09");
    assert_eq!(clock.format("h:mm AM/PM"), "1:07 PM");
    assert_eq!(clock.format("ddd d MMM 'at' h"), "Thu 29 Feb at 13");
    assert_eq!(Clock::parse("2024-01-01").unwrap().format("dddd"), "Monday");
    for bad in [
        "2023-02-29",
        "2024-13-01",
        "24-01-01",
        "2024-01-01T25:00",
        "2024-01-01Z",
    ] {
        assert!(Clock::parse(bad).is_err(), "{bad}");
    }
}

#[test]
fn numbers_and_cases_follow_word_formats() {
    assert_eq!(format_number(14, "ROMAN").as_deref(), Some("XIV"));
    assert_eq!(format_number(14, "roman").as_deref(), Some("xiv"));
    assert_eq!(format_number(28, "ALPHABETIC").as_deref(), Some("BB"));
    assert_eq!(format_number(3, "alphabetic").as_deref(), Some("c"));
    assert_eq!(format_number(3, "Ordinal"), None);
    assert_eq!(
        apply_case("hello world", "Caps").as_deref(),
        Some("Hello World")
    );
    assert_eq!(apply_case("hello", "FirstCap").as_deref(), Some("Hello"));
    assert_eq!(apply_case("x", "Hex"), None);
}
