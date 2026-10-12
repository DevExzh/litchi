//! Lengths: OfficeCLI `EmuConverter.ParseEmu`, `ParseEmuAsInt`,
//! `ParseLineWidth`, `ParseHelpers.ParseFontSize` and the `Format*`
//! renderings.

use super::super::{ErrorKind, Length, LengthContext, LengthStyle, What};

fn emu(text: &str) -> i64 {
    Length::parse_in(text, LengthContext::Emu)
        .unwrap_or_else(|error| panic!("{text:?}: {error}"))
        .emu()
}

fn emu_in(text: &str, context: LengthContext) -> Result<i64, ErrorKind> {
    Length::parse_in(text, context)
        .map(Length::emu)
        .map_err(|error| error.kind())
}

#[test]
fn every_unit_reads_one_inch() {
    for text in [
        "1in",
        "2.54cm",
        "25.4mm",
        "6pc",
        "72pt",
        "96px",
        "101.6Q",
        "914400emu",
        "914400",
        "9.144e5",
        "9.144E+5",
    ] {
        assert_eq!(emu(text), 914_400, "{text}");
    }
}

#[test]
fn unit_sizes_follow_officecli() {
    let cases: &[(&str, i64)] = &[
        ("1mm", 36_000),
        ("1cm", 360_000),
        ("1in", 914_400),
        ("1pc", 152_400),
        ("1pt", 12_700),
        ("1px", 9_525),
        ("1Q", 9_000),
        ("1q", 9_000),
        ("1emu", 1),
        ("1", 1),
    ];
    for &(text, expected) in cases {
        assert_eq!(emu(text), expected, "{text}");
    }
}

#[test]
fn units_ignore_case_and_white_space() {
    let cases: &[(&str, i64)] = &[
        ("1IN", 914_400),
        ("2.54CM", 914_400),
        ("1Pt", 12_700),
        ("1PX", 9_525),
        ("1Mm", 36_000),
        ("1EMU", 1),
        (" 1in ", 914_400),
        ("1 in", 914_400),
        ("\t2cm\n", 720_000),
        ("\u{3000}1in\u{3000}", 914_400),
        ("+1cm", 360_000),
        ("1 emu", 1),
    ];
    for &(text, expected) in cases {
        assert_eq!(emu(text), expected, "{text:?}");
    }
    // `NumberStyles.Float` admits only ASCII white space inside.
    assert_eq!(
        emu_in("1\u{a0}in", LengthContext::Emu),
        Err(ErrorKind::BadNumber { unit: "in" })
    );
}

#[test]
fn negative_positions_are_allowed() {
    assert_eq!(emu("-1in"), -914_400);
    assert_eq!(emu("-1"), -1);
    assert_eq!(emu("-2.54cm"), -914_400);
}

#[test]
fn rounding_is_to_nearest_emu_ties_to_even() {
    // .NET `Math.Round`: 4762.5 -> 4762, 14287.5 -> 14288, 23812.5 -> 23812.
    assert_eq!(emu("0.5px"), 4_762);
    assert_eq!(emu("1.5px"), 14_288);
    assert_eq!(emu("2.5px"), 23_812);
    assert_eq!(emu("-0.5px"), -4_762);
    assert_eq!(emu("2.5e0"), 2);
    assert_eq!(emu("3.5e0"), 4);
    assert_eq!(emu("-2.5e0"), -2);
    assert_eq!(emu("1e-1"), 0);
}

#[test]
fn bare_decimals_are_not_emu() {
    let error = Length::parse_in("1.5", LengthContext::Emu).unwrap_err();
    assert_eq!(error.kind(), ErrorKind::Syntax);
    assert_eq!(error.what(), What::Length);
    assert_eq!(
        error.suggestion().as_deref(),
        Some("add a unit such as 'cm' or 'pt'; a bare number is a whole number of EMU")
    );
    for text in [
        "0.5",
        "abc",
        "1 2",
        "0x10",
        "1,000",
        "NaN",
        "Infinity",
        "9223372036854775808",
    ] {
        assert_eq!(
            emu_in(text, LengthContext::Emu),
            Err(ErrorKind::Syntax),
            "{text}"
        );
    }
}

#[test]
fn unit_suffix_errors() {
    for (text, unit) in [
        ("cm", "cm"),
        (" cm", "cm"),
        ("emu", "emu"),
        ("Q", "Q"),
        ("q", "Q"),
        ("  pt", "pt"),
    ] {
        assert_eq!(
            emu_in(text, LengthContext::Emu),
            Err(ErrorKind::MissingNumber { unit }),
            "{text:?}"
        );
    }
    for (text, unit) in [
        ("abcm", "cm"),
        ("Berlin", "in"),
        ("1,000cm", "cm"),
        ("1e999cm", "cm"),
        ("1.5emu", "emu"),
        ("1e3emu", "emu"),
        ("--1pt", "pt"),
    ] {
        assert_eq!(
            emu_in(text, LengthContext::Emu),
            Err(ErrorKind::BadNumber { unit }),
            "{text:?}"
        );
    }
    let error = Length::parse("Berlin").unwrap_err();
    assert_eq!(
        error.to_string(),
        "invalid length: the component before the 'in' unit is not a finite number"
    );
}

#[test]
fn font_and_viewport_units_are_refused() {
    for (text, unit) in [
        ("2em", "em"),
        ("1.5rem", "rem"),
        ("3ex", "ex"),
        ("50vw", "vw"),
        ("50vh", "vh"),
        ("2EM", "em"),
    ] {
        assert_eq!(
            emu_in(text, LengthContext::Emu),
            Err(ErrorKind::UnsupportedUnit { unit }),
            "{text}"
        );
    }
    let error = Length::parse("2em").unwrap_err();
    assert_eq!(
        error.suggestion().as_deref(),
        Some("use cm, mm, in, pt, pc, px, Q or emu")
    );
}

#[test]
fn non_finite_and_overflowing_values() {
    assert_eq!(
        emu_in("1e999", LengthContext::Emu),
        Err(ErrorKind::NotFinite)
    );
    assert_eq!(emu_in("1e19", LengthContext::Emu), Err(ErrorKind::Overflow));
    assert_eq!(
        emu_in("-1e19", LengthContext::Emu),
        Err(ErrorKind::Overflow)
    );
    assert_eq!(
        emu_in("9.2233720368547758e18", LengthContext::Emu),
        Err(ErrorKind::Overflow)
    );
    assert_eq!(emu("9.2233720368547748e18"), 9_223_372_036_854_774_784);
    assert_eq!(emu("9223372036854775807"), i64::MAX);
    assert_eq!(emu("-9223372036854775808"), i64::MIN);
    assert_eq!(
        emu_in("1e15in", LengthContext::Emu),
        Err(ErrorKind::Overflow)
    );
    assert_eq!(emu_in("", LengthContext::Emu), Err(ErrorKind::Empty));
    assert_eq!(emu_in(" \t", LengthContext::Emu), Err(ErrorKind::Empty));
}

#[test]
fn extents_are_non_negative_32_bit() {
    assert_eq!(
        emu_in("2147483647", LengthContext::Extent),
        Ok(2_147_483_647)
    );
    assert_eq!(emu_in("1in", LengthContext::Extent), Ok(914_400));
    assert_eq!(emu_in("0", LengthContext::Extent), Ok(0));
    assert_eq!(
        emu_in("2147483648", LengthContext::Extent),
        Err(ErrorKind::TooLarge)
    );
    assert_eq!(
        emu_in("-1", LengthContext::Extent),
        Err(ErrorKind::Negative)
    );
    assert_eq!(
        emu_in("-1cm", LengthContext::Extent),
        Err(ErrorKind::Negative)
    );
    assert_eq!(emu_in("1.5", LengthContext::Extent), Err(ErrorKind::Syntax));
}

#[test]
fn line_widths_read_bare_numbers_as_points() {
    let cases: &[(&str, i64)] = &[
        ("1", 12_700),
        ("0.5", 6_350),
        ("1.5", 19_050),
        ("0", 0),
        (" 2 ", 25_400),
        ("1e2", 1_270_000),
        ("2pt", 25_400),
        ("1cm", 360_000),
        ("12700emu", 12_700),
        ("169093pt", 2_147_481_100),
    ];
    for &(text, expected) in cases {
        assert_eq!(
            emu_in(text, LengthContext::LineWidth),
            Ok(expected),
            "{text}"
        );
    }
    assert_eq!(
        emu_in("-1", LengthContext::LineWidth),
        Err(ErrorKind::Negative)
    );
    assert_eq!(
        emu_in("-1pt", LengthContext::LineWidth),
        Err(ErrorKind::Negative)
    );
    assert_eq!(
        emu_in("169094pt", LengthContext::LineWidth),
        Err(ErrorKind::TooLarge)
    );
    assert_eq!(
        emu_in("1e999", LengthContext::LineWidth),
        Err(ErrorKind::NotFinite)
    );
    assert_eq!(emu_in("", LengthContext::LineWidth), Err(ErrorKind::Empty));
    assert_eq!(
        emu_in("2em", LengthContext::LineWidth),
        Err(ErrorKind::UnsupportedUnit { unit: "em" })
    );
    let error = Length::parse_in("-1", LengthContext::LineWidth).unwrap_err();
    assert_eq!(error.what(), What::LineWidth);
    assert_eq!(
        error.to_string(),
        "invalid line width: it must not be negative"
    );
}

#[test]
fn font_sizes_are_points_from_half_to_4000() {
    let cases: &[(&str, i64)] = &[
        ("11", 139_700),
        ("10.5", 133_350),
        ("10.5pt", 133_350),
        ("10.5 pt", 133_350),
        (" 14PT ", 177_800),
        ("0.5", 6_350),
        ("4000", 50_800_000),
        ("1e1", 127_000),
        ("10.25", 130_175),
    ];
    for &(text, expected) in cases {
        assert_eq!(
            emu_in(text, LengthContext::FontSize),
            Ok(expected),
            "{text:?}"
        );
    }
    assert_eq!(
        emu_in("0.49", LengthContext::FontSize),
        Err(ErrorKind::TooSmall)
    );
    assert_eq!(
        emu_in("0", LengthContext::FontSize),
        Err(ErrorKind::Negative)
    );
    assert_eq!(
        emu_in("-1", LengthContext::FontSize),
        Err(ErrorKind::Negative)
    );
    assert_eq!(
        emu_in("4000.01", LengthContext::FontSize),
        Err(ErrorKind::TooLarge)
    );
    assert_eq!(
        emu_in("10,5", LengthContext::FontSize),
        Err(ErrorKind::Comma)
    );
    assert_eq!(
        emu_in("1cm", LengthContext::FontSize),
        Err(ErrorKind::Syntax)
    );
    assert_eq!(emu_in("pt", LengthContext::FontSize), Err(ErrorKind::Empty));
    assert_eq!(
        emu_in("1e999", LengthContext::FontSize),
        Err(ErrorKind::NotFinite)
    );
    let error = Length::parse_in("10,5", LengthContext::FontSize).unwrap_err();
    assert_eq!(
        error.to_string(),
        "invalid font size: use '.' as the decimal separator, not ','"
    );
}

fn shown(emu: i64, style: LengthStyle) -> String {
    Length::from_emu(emu).display(style).to_string()
}

#[test]
fn auto_style_is_format_emu() {
    let cases: &[(i64, &str)] = &[
        (0, "0cm"),
        (914_400, "2.54cm"),
        (360_000, "1cm"),
        (12_700, "1pt"),
        (12_192_000, "960pt"),
        (6_858_000, "19.05cm"),
        (1_800, "1800emu"),
        (-914_400, "-2.54cm"),
        (9_144, "0.72pt"),
        (1, "1emu"),
        (127, "0.01pt"),
        (3_600, "0.01cm"),
        (63, "63emu"),
        (i64::MAX, "9223372036854775807emu"),
        (i64::MIN, "-9223372036854775808emu"),
    ];
    for &(value, expected) in cases {
        assert_eq!(shown(value, LengthStyle::Auto), expected, "{value}");
        assert_eq!(Length::from_emu(value).to_string(), expected);
    }
}

#[test]
fn points_first_style_is_format_emu_pt_first() {
    let cases: &[(i64, &str)] = &[
        (0, "0pt"),
        (914_400, "72pt"),
        (360_000, "1cm"),
        (6_858_000, "540pt"),
        (1_800, "1800emu"),
    ];
    for &(value, expected) in cases {
        assert_eq!(shown(value, LengthStyle::PointsFirst), expected, "{value}");
    }
}

#[test]
fn lossy_style_is_format_emu_lossy() {
    let cases: &[(i64, &str)] = &[
        (0, "0cm"),
        (914_400, "2.54cm"),
        (2_640_000, "7.33cm"),
        (1_800, "0.01cm"),
        (1_799, "0cm"),
        (-1_799, "-0cm"),
        (12_700, "1pt"),
    ];
    for &(value, expected) in cases {
        assert_eq!(shown(value, LengthStyle::Lossy), expected, "{value}");
    }
}

#[test]
fn line_width_style_is_format_line_width() {
    let cases: &[(i64, &str)] = &[
        (0, "0pt"),
        (12_700, "1pt"),
        (180_000, "14.17pt"),
        (63, "63emu"),
        (64, "0.01pt"),
        (126, "0.01pt"),
        (1, "1emu"),
        (-64, "-0.01pt"),
        (-63, "-63emu"),
        (6_350, "0.5pt"),
    ];
    for &(value, expected) in cases {
        assert_eq!(shown(value, LengthStyle::LineWidth), expected, "{value}");
    }
}

#[test]
fn points_style_prints_font_sizes_and_spacing() {
    assert_eq!(shown(139_700, LengthStyle::Points), "11pt");
    assert_eq!(shown(133_350, LengthStyle::Points), "10.5pt");
    assert_eq!(shown(130_175, LengthStyle::Points), "10.25pt");
    assert_eq!(shown(0, LengthStyle::Points), "0pt");
    assert_eq!(shown(-63, LengthStyle::Points), "-0pt");
    assert_eq!(
        Length::from_half_points(21)
            .display(LengthStyle::Points)
            .to_string(),
        "10.5pt"
    );
    assert_eq!(
        Length::from_centipoints(1025)
            .display(LengthStyle::Points)
            .to_string(),
        "10.25pt"
    );
    assert_eq!(
        Length::from_twips(240)
            .display(LengthStyle::Points)
            .to_string(),
        "12pt"
    );
    assert_eq!(
        Length::from_twips(1)
            .display(LengthStyle::Points)
            .to_string(),
        "0.05pt"
    );
    assert_eq!(
        Length::from_points(-3)
            .display(LengthStyle::Points)
            .to_string(),
        "-3pt"
    );
}

#[test]
fn pairs_share_one_exact_unit() {
    let pair = |a, b| {
        let (first, second) = Length::display_pair(Length::from_emu(a), Length::from_emu(b));
        (first.to_string(), second.to_string())
    };
    assert_eq!(
        pair(12_192_000, 6_858_000),
        ("960pt".to_owned(), "540pt".to_owned())
    );
    assert_eq!(
        pair(7_560_000, 10_692_000),
        ("21cm".to_owned(), "29.7cm".to_owned())
    );
    assert_eq!(pair(0, 914_400), ("0cm".to_owned(), "2.54cm".to_owned()));
    assert_eq!(
        pair(1_800, 914_400),
        ("1800emu".to_owned(), "2.54cm".to_owned())
    );
}

#[test]
fn constructors_and_accessors() {
    assert_eq!(Length::from_points(1).emu(), 12_700);
    assert_eq!(
        Length::from_points(i32::MIN).emu(),
        i64::from(i32::MIN) * 12_700
    );
    assert_eq!(Length::from_centipoints(1).emu(), 127);
    assert_eq!(Length::from_half_points(1).emu(), 6_350);
    assert_eq!(Length::from_twips(1).emu(), 635);
    assert_eq!(Length::ZERO, Length::default());
    assert!((Length::from_emu(6_350).points() - 0.5).abs() < f64::EPSILON);
    assert_eq!("2.54cm".parse::<Length>().expect("parses").emu(), 914_400);
    let error = "2.5".parse::<Length>().unwrap_err();
    assert_eq!(error.code(), crate::Code::InvalidValue);
    assert!(Length::from_emu(1) < Length::from_emu(2));
}

#[test]
fn word_spacing_uses_unsigned_twips_and_rounds_at_the_leaf_unit() {
    let parse = |value| {
        Length::parse_in(value, LengthContext::WordSpacing)
            .unwrap()
            .emu()
    };
    assert_eq!(parse("20"), 12_700);
    assert_eq!(parse("1pt"), 12_700);
    assert_eq!(parse("2.54cm"), 914_400);
    assert_eq!(parse("1in"), 914_400);
    assert_eq!(parse("1.5"), 1_270);
    assert_eq!(parse("2.5"), 1_270);
    assert_eq!(parse("4294967295"), i64::from(u32::MAX) * 635);
    for input in ["-1", "4294967296", "NaN", "inf", "1emu", "1mm"] {
        assert!(Length::parse_in(input, LengthContext::WordSpacing).is_err());
    }
}

#[cfg(feature = "rtf")]
#[test]
fn checked_twips_rounds_signed_boundaries_and_refuses_i32_overflow() {
    let emu = |value: i64| Length::from_emu(value).checked_twips();
    assert_eq!(emu(317), Some(0));
    assert_eq!(emu(318), Some(1));
    assert_eq!(emu(-317), Some(0));
    assert_eq!(emu(-318), Some(-1));

    let maximum = Length::from_twips(i32::MAX).emu();
    assert_eq!(emu(maximum), Some(i32::MAX));
    assert_eq!(emu(maximum + 635), None);
}
