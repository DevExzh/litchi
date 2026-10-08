//! Booleans (`ParseHelpers.IsTruthy`) and opacities.

use super::super::{ErrorKind, What, parse_bool, parse_opacity};
use super::kind;

#[test]
fn every_truthy_and_falsy_spelling() {
    for text in [
        "true", "1", "yes", "on", "TRUE", "Yes", "oN", " true ", "\ttrue\n",
    ] {
        assert_eq!(parse_bool(text), Ok(true), "{text:?}");
    }
    for text in ["false", "0", "no", "off", "", "   ", "FALSE", "No", "OFF"] {
        assert_eq!(parse_bool(text), Ok(false), "{text:?}");
    }
}

#[test]
fn invisible_characters_are_trimmed_after_white_space() {
    assert_eq!(parse_bool("\u{feff}true"), Ok(true));
    assert_eq!(parse_bool("true\u{200b}"), Ok(true));
    assert_eq!(parse_bool("\u{200c}\u{200d}\u{2060}on\u{a0}"), Ok(true));
    // OfficeCLI trims white space first, then the invisible set, so a space
    // inside the invisible characters stays.
    assert_eq!(kind(parse_bool("\u{200b} true")), ErrorKind::Syntax);
}

#[test]
fn anything_else_is_invalid_with_a_suggestion() {
    for text in ["2", "y", "n", "enabled", "truee", "t r u e", "１"] {
        let error = parse_bool(text).unwrap_err();
        assert_eq!(error.kind(), ErrorKind::Syntax, "{text:?}");
        assert_eq!(error.what(), What::Boolean);
    }
    let error = parse_bool("ture").unwrap_err();
    assert_eq!(error.suggestion().as_deref(), Some("did you mean 'true'?"));
    assert_eq!(
        error.to_string(),
        "invalid boolean: expected true/false, yes/no, 1/0 or on/off"
    );
}

#[test]
fn opacity_reads_fractions_and_percentages() {
    let cases: &[(&str, f64)] = &[
        ("0", 0.0),
        ("0.25", 0.25),
        ("1", 1.0),
        ("1.0", 1.0),
        ("2", 0.02),
        ("30", 0.3),
        ("100", 1.0),
        (" 0.5 ", 0.5),
    ];
    for &(text, expected) in cases {
        assert_eq!(parse_opacity(text), Ok(expected), "{text:?}");
    }
    assert_eq!(kind(parse_opacity("1.5")), ErrorKind::Ambiguous);
    assert_eq!(kind(parse_opacity("1.0001")), ErrorKind::Ambiguous);
    assert_eq!(kind(parse_opacity("101")), ErrorKind::OutOfRange);
    assert_eq!(kind(parse_opacity("-0.1")), ErrorKind::OutOfRange);
    assert_eq!(kind(parse_opacity("50%")), ErrorKind::Syntax);
    assert_eq!(kind(parse_opacity("")), ErrorKind::Empty);
    assert_eq!(kind(parse_opacity("1e999")), ErrorKind::NotFinite);
}
