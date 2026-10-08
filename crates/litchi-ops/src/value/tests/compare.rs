//! Query comparison: OfficeCLI `AttributeFilter.MatchOne` (`StringEquals`,
//! `DimensionEquals`, `NumericEquals`, `CompareNumeric`, `ExtractNumber`,
//! `ExtractUnit`, `MatchesTextFilter`) and typed equality.

use litchi_core::selector::{CompareOp, Condition, Key, Value};

use super::super::{
    Color, ColorContext, Length, LengthContext, LengthSpec, LengthStyle, Scalar, Test, Tokens,
    ValueType, Warning,
};
use crate::Code;

const OPS: [CompareOp; 7] = [
    CompareOp::Eq,
    CompareOp::Ne,
    CompareOp::Contains,
    CompareOp::Gt,
    CompareOp::Lt,
    CompareOp::Ge,
    CompareOp::Le,
];

/// `actual op literal` with both sides as text and no key type: pure
/// OfficeCLI semantics.
fn text(actual: &str, op: CompareOp, literal: &str) -> bool {
    Test::with_text(op, literal, None).eval(Some(Scalar::Text(actual)))
}

#[test]
fn string_equality_ignores_case_and_leading_hashes() {
    assert!(text("Heading1", CompareOp::Eq, "heading1"));
    assert!(text("#FF0000", CompareOp::Eq, "ff0000"));
    assert!(text("FF0000", CompareOp::Eq, "#ff0000"));
    assert!(text("##FF0000", CompareOp::Eq, "#FF0000"));
    assert!(text("Überschrift", CompareOp::Eq, "üBERSCHRIFT"));
    assert!(!text("Heading1", CompareOp::Eq, "Heading 1"));
    assert!(!text("abc", CompareOp::Eq, "abcd"));
    assert!(!text("Straße", CompareOp::Eq, "STRASSE"));
}

#[test]
fn dimension_equality_has_a_500_emu_tolerance_with_a_unit() {
    assert!(text("2cm", CompareOp::Eq, "20mm"));
    assert!(text("2.54cm", CompareOp::Eq, "1in"));
    assert!(text("2cm", CompareOp::Eq, "720400emu"));
    assert!(text("2cm", CompareOp::Eq, "720500"));
    assert!(!text("2cm", CompareOp::Eq, "720501"));
    // No unit on either side: no tolerance ("50" is not "150").
    assert!(!text("50", CompareOp::Eq, "150"));
    assert!(!text("914400", CompareOp::Eq, "914700"));
}

#[test]
fn numeric_equality() {
    assert!(text("25", CompareOp::Eq, "25.0"));
    assert!(text("11pt", CompareOp::Eq, "11"));
    assert!(text("1,000", CompareOp::Eq, "1000"));
    assert!(text("50%", CompareOp::Eq, "50"));
    assert!(text("(5)", CompareOp::Eq, "-5"));
    assert!(text("1E+20", CompareOp::Eq, "100000000000000000000"));
    assert!(!text("1,5", CompareOp::Eq, "15"));
    assert!(text("25.0", CompareOp::Ne, "26"));
    assert!(!text("25.0", CompareOp::Ne, "25"));
}

#[test]
fn absent_keys_match_only_not_equal() {
    for op in OPS {
        let test = Test::with_text(op, "x", None);
        assert_eq!(test.eval(None), op == CompareOp::Ne, "{op:?}");
    }
    assert!(!Test::exists().eval(None));
}

#[test]
fn exists_needs_non_empty_text() {
    assert!(Test::exists().eval(Some(Scalar::Text("x"))));
    assert!(!Test::exists().eval(Some(Scalar::Text(""))));
    assert!(Test::exists().eval(Some(Scalar::Bool(false))));
    assert!(Test::exists().eval(Some(Scalar::Int(0))));
}

#[test]
fn contains_is_a_case_insensitive_substring() {
    assert!(text("Hello World", CompareOp::Contains, "WORLD"));
    assert!(text("Hello World", CompareOp::Contains, ""));
    assert!(text("", CompareOp::Contains, ""));
    assert!(!text("", CompareOp::Contains, "a"));
    assert!(text("Ölfeld", CompareOp::Contains, "öL"));
    assert!(text("报告 2024", CompareOp::Contains, "报告"));
    assert!(text("abcabd", CompareOp::Contains, "abd"));
    assert!(!text("Straße", CompareOp::Contains, "SS"));
    assert!(!text("Hello", CompareOp::Contains, "hello world"));
}

#[test]
fn relational_operators_follow_compare_numeric() {
    let cases: &[(&str, CompareOp, &str, bool)] = &[
        ("5000", CompareOp::Gt, "4000", true),
        ("5000", CompareOp::Lt, "4000", false),
        ("5000", CompareOp::Ge, "5000", true),
        ("5000", CompareOp::Le, "5000", true),
        ("1,500", CompareOp::Gt, "1000", true),
        // Units differ: compared in EMU.
        ("2.54cm", CompareOp::Ge, "1in", true),
        ("2.54cm", CompareOp::Gt, "1in", false),
        ("12pt", CompareOp::Lt, "1in", true),
        ("20mm", CompareOp::Lt, "2.1cm", true),
        // One side without a unit: the numbers compare as written.
        ("2.54cm", CompareOp::Gt, "2", true),
        ("960pt", CompareOp::Gt, "100", true),
        ("2.54cm", CompareOp::Gt, "100", false),
        // `<n>emu` has no number for OfficeCLI, only an EMU value.
        ("1800emu", CompareOp::Gt, "1000", true),
        ("1800emu", CompareOp::Lt, "1cm", true),
        ("50%", CompareOp::Gt, "40%", true),
        // Mixed units that are not both lengths fall back to the numbers.
        ("5em", CompareOp::Gt, "3pt", true),
        // Non-numeric text never matches.
        ("abc", CompareOp::Gt, "1", false),
        ("abc", CompareOp::Lt, "1", false),
        ("张三", CompareOp::Gt, "5000", false),
        ("5", CompareOp::Gt, "abc", false),
    ];
    for &(actual, op, literal, expected) in cases {
        assert_eq!(
            text(actual, op, literal),
            expected,
            "{actual} {op} {literal}"
        );
    }
}

#[test]
fn relational_literals_that_are_not_numbers_warn() {
    assert_eq!(
        Test::with_text(CompareOp::Gt, "abc", None).warning(),
        Some(Warning::NonNumeric)
    );
    assert_eq!(
        Test::with_text(CompareOp::Le, "high", None).warning(),
        Some(Warning::NonNumeric)
    );
    assert_eq!(Test::with_text(CompareOp::Gt, "5", None).warning(), None);
    assert_eq!(Test::with_text(CompareOp::Gt, "2cm", None).warning(), None);
    assert_eq!(
        Test::with_text(CompareOp::Gt, "100emu", None).warning(),
        None
    );
    assert_eq!(Test::with_text(CompareOp::Gt, "50%", None).warning(), None);
    assert_eq!(Test::with_text(CompareOp::Eq, "abc", None).warning(), None);
    assert_eq!(Test::exists().warning(), None);
}

#[test]
fn header_rows_read_numeric_literals_without_units() {
    assert_eq!(
        Test::with_text(CompareOp::Eq, "5,000", None).literal_number(),
        Some(5000.0)
    );
    assert_eq!(
        Test::with_text(CompareOp::Eq, "5k", None).literal_number(),
        None
    );
    assert_eq!(
        Test::with_text(CompareOp::Eq, "12pt", None).literal_number(),
        None
    );
    assert_eq!(
        Test::with_text(CompareOp::Eq, "12pt", None).literal(),
        "12pt"
    );
}

#[test]
fn regex_values_are_unsupported() {
    let pattern = Value::regex("^Chapter \\d+").expect("representable");
    let error = Test::compare(CompareOp::Contains, &pattern, None).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(error.suggestion().is_some());
    let literal = Value::text("Chapter");
    assert!(Test::compare(CompareOp::Contains, &literal, None).is_ok());
}

#[test]
fn conditions_compile_from_the_grammar() {
    let exists = Condition::Exists(Key::new("bold"));
    assert!(
        Test::from_condition(&exists, None)
            .expect("compiles")
            .eval(Some(Scalar::Bool(true)))
    );
    let compare = Condition::Compare {
        key: Key::new("size"),
        op: CompareOp::Ge,
        value: Value::text("12pt"),
    };
    let test = Test::from_condition(&compare, Some(ValueType::Length(LengthSpec::FONT_SIZE)))
        .expect("compiles");
    assert!(test.eval(Some(Scalar::Length(Length::from_points(14)))));
    assert!(!test.eval(Some(Scalar::Length(Length::from_points(10)))));
}

#[test]
fn typed_equality_matches_what_the_key_would_accept() {
    let fill = Some(ValueType::Color(ColorContext::Drawing));
    let red = Some(Scalar::Color(Color::rgb(255, 0, 0)));
    for literal in [
        "red",
        "F00",
        "#ff0000",
        "rgb(255,0,0)",
        "hsl(0,100%,50%)",
        "FFFF0000",
    ] {
        assert!(
            Test::with_text(CompareOp::Eq, literal, fill).eval(red),
            "{literal}"
        );
        assert!(
            !Test::with_text(CompareOp::Ne, literal, fill).eval(red),
            "{literal}"
        );
    }
    assert!(!Test::with_text(CompareOp::Eq, "blue", fill).eval(red));
    // Without the key's type, `red` is only text, as in OfficeCLI.
    assert!(!Test::with_text(CompareOp::Eq, "red", None).eval(red));

    let accent = Some(Scalar::Color(
        Color::parse("dk1+lumMod75", ColorContext::Drawing).expect("parses"),
    ));
    assert!(Test::with_text(CompareOp::Eq, "dark1+lumMod75", fill).eval(accent));
    assert!(Test::with_text(CompareOp::Eq, "dk1+lumMod=75000", fill).eval(accent));
    assert!(!Test::with_text(CompareOp::Eq, "dk1", fill).eval(accent));

    let cell = Some(ValueType::Color(ColorContext::Spreadsheet));
    let tinted = Some(Scalar::Color(
        Color::parse("accent1+tint40", ColorContext::Spreadsheet).expect("parses"),
    ));
    assert!(Test::with_text(CompareOp::Eq, "accent1+tint39.8", cell).eval(tinted));
    assert!(!Test::with_text(CompareOp::Eq, "accent1+tint30", cell).eval(tinted));

    let bold = Some(ValueType::Bool);
    for (literal, expected) in [
        ("yes", true),
        ("1", true),
        ("on", true),
        ("TRUE", true),
        ("no", false),
        ("0", false),
    ] {
        assert_eq!(
            Test::with_text(CompareOp::Eq, literal, bold).eval(Some(Scalar::Bool(true))),
            expected,
            "{literal}"
        );
    }
    // An empty literal is not coerced: `[bold=]` does not mean false.
    assert!(!Test::with_text(CompareOp::Eq, "", bold).eval(Some(Scalar::Bool(false))));

    static UNDERLINE: Tokens = Tokens::new("underline", &["single", "double"], &[("sgl", 0)]);
    let underline = Some(ValueType::Token(&UNDERLINE));
    assert!(Test::with_text(CompareOp::Eq, "sgl", underline).eval(Some(Scalar::Token("single"))));
    assert!(Test::with_text(CompareOp::Eq, "Single", underline).eval(Some(Scalar::Text("single"))));
    assert!(!Test::with_text(CompareOp::Eq, "sgl", underline).eval(Some(Scalar::Token("double"))));

    let width = Some(ValueType::Length(LengthSpec::EXTENT));
    let inch = Some(Scalar::Length(Length::from_emu(914_400)));
    assert!(Test::with_text(CompareOp::Eq, "914400", width).eval(inch));
    assert!(Test::with_text(CompareOp::Eq, "72pt", width).eval(inch));
    let count = Some(ValueType::Int);
    assert!(Test::with_text(CompareOp::Eq, " +7 ", count).eval(Some(Scalar::Int(7))));
}

#[test]
fn typed_lengths_print_in_the_key_style_for_comparison() {
    let line = Some(ValueType::Length(LengthSpec::LINE_WIDTH));
    let width = Some(Scalar::Length(Length::from_emu(180_000)));
    // The line width prints `14.17pt`; OfficeCLI compares that text.
    assert!(Test::with_text(CompareOp::Contains, "14.17", line).eval(width));
    assert!(Test::with_text(CompareOp::Gt, "14", line).eval(width));
    assert!(Test::with_text(CompareOp::Eq, "14.17pt", line).eval(width));
    // As a position the same value prints `0.5cm`.
    let position = Some(ValueType::Length(LengthSpec::POSITION));
    assert!(Test::with_text(CompareOp::Contains, "0.5cm", position).eval(width));
}

fn pseudo_random(state: &mut u64) -> u64 {
    *state = state.wrapping_add(0x9E37_79B9_7F4A_7C15);
    let mut z = *state;
    z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
    z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
    z ^ (z >> 31)
}

const LITERALS: &[&str] = &[
    "0", "1", "-1", "100", "914400", "2.54", "2.54cm", "1in", "72pt", "1cm", "960pt", "11pt",
    "10.5", "50%", "1800emu", "3600", "abc", "true", "#FF0000", "accent1", "1e20", "1E+20", "0.5",
    "12700", "-2.54cm", "pt", "cm", "", "2em", "1,000", "(5)",
];

/// A typed value evaluates exactly as its canonical text does, for every
/// operator OfficeCLI's string semantics decide (typed equality, which only
/// adds matches for `=` and removes them for `!=`, is off here because the
/// tests carry no key type, or carry a type whose coercion adds nothing).
#[test]
fn typed_values_compare_as_their_canonical_text() {
    let mut state = 7_u64;
    let mut values: Vec<Scalar<'static>> = vec![
        Scalar::Bool(true),
        Scalar::Bool(false),
        Scalar::Int(0),
        Scalar::Int(-5),
        Scalar::Int(914_400),
        Scalar::Number(1e20),
        Scalar::Number(2.54),
        Scalar::Number(-0.5),
        Scalar::Number(1.5e-7),
        Scalar::Token("single"),
        Scalar::Color(Color::rgb(255, 0, 0)),
        Scalar::Color(Color::parse("accent1+lumMod75", ColorContext::Drawing).expect("parses")),
    ];
    for emu in [
        0, 1, 63, 127, 1_800, 3_600, 12_700, 180_000, 914_400, 12_192_000, -914_400,
    ] {
        values.push(Scalar::Length(Length::from_emu(emu)));
    }
    for _ in 0..200 {
        let emu = (pseudo_random(&mut state) % 40_000_000) as i64 - 20_000_000;
        values.push(Scalar::Length(Length::from_emu(emu)));
    }
    for style in [
        LengthStyle::Auto,
        LengthStyle::PointsFirst,
        LengthStyle::Lossy,
        LengthStyle::LineWidth,
        LengthStyle::Points,
    ] {
        let ty = Some(ValueType::Length(LengthSpec::new(
            LengthContext::Emu,
            style,
        )));
        for value in &values {
            let canonical = value.display(ty).to_string();
            for op in OPS {
                for literal in LITERALS {
                    let typed = Test::with_text(op, literal, None);
                    let styled = Test::with_text(op, literal, ty);
                    let reference =
                        Test::with_text(op, literal, None).eval(Some(Scalar::Text(&canonical)));
                    if style == LengthStyle::Auto {
                        assert_eq!(
                            typed.eval(Some(*value)),
                            reference,
                            "{value:?} {op} {literal:?}"
                        );
                    }
                    if !matches!(op, CompareOp::Eq | CompareOp::Ne)
                        || !matches!(value, Scalar::Length(_))
                    {
                        assert_eq!(
                            styled.eval(Some(*value)),
                            reference,
                            "{value:?} {style:?} {op} {literal:?}"
                        );
                    } else if reference {
                        // Typed equality only adds matches for `=`.
                        assert!(
                            styled.eval(Some(*value)) || op == CompareOp::Ne,
                            "{value:?} {op} {literal:?}"
                        );
                    }
                }
            }
        }
    }
}
