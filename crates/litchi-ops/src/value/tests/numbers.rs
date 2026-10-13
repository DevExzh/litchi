//! Numbers: .NET `NumberStyles` parsing, exact decimals, double rendering.

use super::super::number::{Styles, parse_f64, parse_i32, parse_i64};
use super::super::{
    Decimal, DotNetDouble, ErrorKind, parse_integer, parse_number, parse_numeric_text,
};
use super::kind;

#[test]
fn float_style_accepts_dotnet_float_forms() {
    // `NumberStyles.Float`: [ws][sign]digits[.digits][e[sign]digits][ws],
    // white space being U+0009..U+000D and U+0020 only.
    let cases: &[(&str, f64)] = &[
        ("0", 0.0),
        ("+1", 1.0),
        ("-1", -1.0),
        ("1.5", 1.5),
        (".5", 0.5),
        ("5.", 5.0),
        ("1e3", 1000.0),
        ("1E-3", 0.001),
        ("1.5e+2", 150.0),
        ("5.e1", 50.0),
        (" \t1.5\n\x0b\x0c\r ", 1.5),
        ("007", 7.0),
    ];
    for &(text, expected) in cases {
        assert_eq!(parse_number(text), Ok(expected), "{text:?}");
    }
    assert!(parse_number("-0").expect("parses").is_sign_negative());
}

#[test]
fn float_style_rejects_everything_else() {
    for text in [
        "1,000", "1.2.3", "e5", "1e", "1e+", "--1", "+-1", "- 1", "1 2", "0x10", "NaN", "Infinity",
        "inf", "\u{a0}1", "1\u{a0}", "(1)", "1-", "\u{a4}1", "１", "1_000",
    ] {
        assert_eq!(kind(parse_number(text)), ErrorKind::Syntax, "{text:?}");
    }
    assert_eq!(kind(parse_number("")), ErrorKind::Empty);
    assert_eq!(kind(parse_number("  ")), ErrorKind::Empty);
    assert_eq!(kind(parse_number("1e400")), ErrorKind::NotFinite);
    assert_eq!(kind(parse_number("-1e400")), ErrorKind::NotFinite);
}

#[test]
fn integer_style_is_sign_digits_and_white_space() {
    assert_eq!(parse_integer("42"), Ok(42));
    assert_eq!(parse_integer(" -42 "), Ok(-42));
    assert_eq!(parse_integer("+0"), Ok(0));
    assert_eq!(parse_integer("9223372036854775807"), Ok(i64::MAX));
    assert_eq!(parse_integer("-9223372036854775808"), Ok(i64::MIN));
    for text in [
        "9223372036854775808",
        "-9223372036854775809",
        "1.0",
        "1e3",
        "1_000",
        "1,000",
        "- 1",
    ] {
        assert_eq!(kind(parse_integer(text)), ErrorKind::Syntax, "{text:?}");
    }
    assert_eq!(kind(parse_integer("")), ErrorKind::Empty);
    assert_eq!(parse_i32("2147483647"), Some(i32::MAX));
    assert_eq!(parse_i32("2147483648"), None);
    assert_eq!(parse_i32(" 75 "), Some(75));
}

#[test]
fn any_style_follows_dotnet_and_numeric_text() {
    // OfficeCLI `NumericText.TryParse`: `NumberStyles.Any` plus well-formed
    // thousands groups.
    let accepted: &[(&str, f64)] = &[
        ("1,234", 1234.0),
        ("1,234,567.89", 1_234_567.89),
        ("-1,234", -1234.0),
        ("(1,234)", -1234.0),
        ("1,234-", -1234.0),
        ("5 -", -5.0),
        ("5+", 5.0),
        ("\u{a4}5", 5.0),
        ("5\u{a4}", 5.0),
        ("-\u{a4}5", -5.0),
        (" 5 ", 5.0),
        ("1e5", 100_000.0),
        ("1,000e3", 1_000_000.0),
        ("0.5", 0.5),
    ];
    for &(text, expected) in accepted {
        assert_eq!(parse_numeric_text(text), Some(expected), "{text:?}");
    }
    for text in [
        "1,5",
        "1,23",
        "1,2345",
        ",123",
        "1.234,5",
        "1,234.5,6",
        "1,,234",
        "(5",
        "5)",
        "( 5)",
        "(-5)",
        "- 5",
        "12pt",
        "",
        "abc",
        "5--",
    ] {
        assert_eq!(parse_numeric_text(text), None, "{text:?}");
    }
    assert_eq!(parse_numeric_text("1e999"), Some(f64::INFINITY));
}

#[test]
fn long_grouped_numerals_still_parse() {
    let groups = vec!["123"; 60].join(",");
    let text = format!("1,{groups}");
    let value = parse_numeric_text(&text).expect("well-formed groups");
    assert!(value > 1e180 && value.is_finite());
}

#[test]
fn styles_compose_as_in_dotnet() {
    assert_eq!(
        parse_f64("1,5", Styles::FLOAT_THOUSANDS),
        Some(15.0),
        ".NET does not check group sizes"
    );
    assert_eq!(parse_f64("1,5", Styles::FLOAT), None);
    assert_eq!(parse_i64("1.0", Styles::INTEGER), None);
    assert_eq!(parse_i64("(5)", Styles::ANY), Some(-5));
}

#[test]
fn decimals_round_half_away_from_zero() {
    let cases: &[(i64, i64, u8, &str)] = &[
        (914_400, 360_000, 2, "2.54"),
        (-254, 100, 2, "-2.54"),
        (1, 3, 2, "0.33"),
        (2, 3, 2, "0.67"),
        (5, 1000, 2, "0.01"),
        (-5, 1000, 2, "-0.01"),
        (-4, 1000, 2, "-0"),
        (4, 1000, 2, "0"),
        (0, 5, 2, "0"),
        (0, -5, 2, "0"),
        (100, 1, 2, "100"),
        (1050, 100, 2, "10.5"),
        (265, 240, 4, "1.1042"),
        (3, -2, 0, "-2"),
        (i64::MIN, 1, 18, "-9223372036854775808"),
        (i64::MAX, 7, 18, "1317624576693539401"),
        (1, 7, 18, "0.142857142857142857"),
    ];
    for &(numerator, denominator, decimals, expected) in cases {
        let decimal = Decimal::from_ratio(numerator, denominator, decimals).expect("valid");
        assert_eq!(decimal.to_string(), expected, "{numerator}/{denominator}");
    }
    assert_eq!(Decimal::from_ratio(1, 0, 2), None);
    assert_eq!(Decimal::from_ratio(1, 1, 19), None);
    let negative_zero = Decimal::from_ratio(-1, 1000, 2).expect("valid");
    assert!(negative_zero.is_zero() && negative_zero.is_negative());
}

#[test]
fn doubles_print_as_dotnet_prints_them() {
    let cases: &[(f64, &str)] = &[
        (0.0, "0"),
        (-0.0, "-0"),
        (1.5, "1.5"),
        (-2.5, "-2.5"),
        (100.0, "100"),
        (1e14, "100000000000000"),
        (123_456_789_012_345.0, "123456789012345"),
        (1e15, "1E+15"),
        (1_234_567_890_123_456.0, "1.234567890123456E+15"),
        (123_456_789_012_345_680_000.0, "1.2345678901234568E+20"),
        (0.0001, "0.0001"),
        (0.000_123, "0.000123"),
        (0.00001, "1E-05"),
        (1.5e-7, "1.5E-07"),
        (1e300, "1E+300"),
        (5e-324, "5E-324"),
        (f64::MAX, "1.7976931348623157E+308"),
        (0.1 + 0.2, "0.30000000000000004"),
        (33.33, "33.33"),
        (40.0, "40"),
        (f64::NAN, "NaN"),
        (f64::INFINITY, "Infinity"),
        (f64::NEG_INFINITY, "-Infinity"),
    ];
    for &(value, expected) in cases {
        assert_eq!(DotNetDouble(value).to_string(), expected, "{value:e}");
    }
}

#[test]
fn printed_doubles_parse_back_exactly() {
    let mut state = 0x9E37_79B9_7F4A_7C15_u64;
    for _ in 0..20_000 {
        state = state.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = state;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        let value = f64::from_bits(z ^ (z >> 31));
        if !value.is_finite() {
            continue;
        }
        let text = DotNetDouble(value).to_string();
        assert_eq!(parse_numeric_text(&text), Some(value), "{text}");
    }
}

/// `Decimal::from_ratio` rounds in 64 bits when the scaled numerator fits
/// (record 0262): the result and its text are those of the 128-bit
/// computation, kept here as the reference.
#[test]
fn decimals_from_ratios_match_the_128_bit_reference() {
    fn reference(numerator: i64, denominator: i64, decimals: u8) -> String {
        let negative = numerator != 0 && (numerator < 0) != (denominator < 0);
        let scaled = u128::from(numerator.unsigned_abs()) * 10_u128.pow(u32::from(decimals));
        let denominator = u128::from(denominator.unsigned_abs());
        let (quotient, remainder) = (scaled / denominator, scaled % denominator);
        let mut magnitude = if remainder * 2 >= denominator {
            quotient + 1
        } else {
            quotient
        };
        let mut scale = decimals;
        while scale > 0 && magnitude % 10 == 0 {
            magnitude /= 10;
            scale -= 1;
        }
        let unit = 10_u128.pow(u32::from(scale));
        let mut text = String::new();
        if negative {
            text.push('-');
        }
        text.push_str(&(magnitude / unit).to_string());
        if scale > 0 {
            text.push_str(&format!(
                ".{:0width$}",
                magnitude % unit,
                width = usize::from(scale)
            ));
        }
        text
    }
    let mut state = 3_u64;
    let mut next = || {
        state = state.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = state;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    };
    let mut numerators = vec![
        0,
        1,
        -1,
        5,
        -5,
        127,
        3_599,
        3_600,
        914_400,
        i64::MAX,
        i64::MIN,
        i64::MAX / 100,
        i64::MAX / 100 + 1,
        -(i64::MAX / 100) - 1,
    ];
    let mut denominators = vec![1, -1, 2, 3, 7, 127, 240, 3_600, 12_700, 360_000, 914_400];
    denominators.extend([i64::MAX, i64::MIN, 1 << 62, (1 << 62) + 1]);
    for _ in 0..400 {
        numerators.push(next() as i64);
        numerators.push((next() % 100_000_000) as i64 - 50_000_000);
    }
    for _ in 0..40 {
        denominators.push((next() as i64) | 1);
        denominators.push((next() % 1_000_000) as i64 + 1);
    }
    for &numerator in &numerators {
        for &denominator in &denominators {
            for decimals in [0_u8, 1, 2, 4, 18] {
                let decimal =
                    Decimal::from_ratio(numerator, denominator, decimals).expect("a denominator");
                assert_eq!(
                    decimal.to_string(),
                    reference(numerator, denominator, decimals),
                    "{numerator} / {denominator} to {decimals} decimals"
                );
            }
        }
    }
}
