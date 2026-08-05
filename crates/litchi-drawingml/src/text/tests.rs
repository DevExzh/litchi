//! Regression tests for the shared DrawingML text primitives.

use super::*;
use std::mem::size_of;

#[test]
fn closed_domains_round_trip_and_reject_unknown_tokens() {
    for value in [
        Anchor::Top,
        Anchor::Center,
        Anchor::Bottom,
        Anchor::Justified,
        Anchor::Distributed,
    ] {
        assert_eq!(value.token().parse::<Anchor>().unwrap(), value);
    }
    for value in [
        Direction::Horizontal,
        Direction::Vertical,
        Direction::Vertical270,
        Direction::WordArtVertical,
        Direction::EastAsianVertical,
        Direction::MongolianVertical,
        Direction::WordArtVerticalRtl,
    ] {
        assert_eq!(value.token().parse::<Direction>().unwrap(), value);
    }
    for value in [Wrap::Square, Wrap::None] {
        assert_eq!(value.token().parse::<Wrap>().unwrap(), value);
    }
    assert!("middle".parse::<Anchor>().is_err());
    assert!("diagonal".parse::<Direction>().is_err());
    assert!("tight".parse::<Wrap>().is_err());
}

#[test]
fn underline_codecs_are_lossless_in_both_dialects() {
    let values = [
        Underline::None,
        Underline::Words,
        Underline::Single,
        Underline::Double,
        Underline::Heavy,
        Underline::Dotted,
        Underline::DottedHeavy,
        Underline::Dash,
        Underline::DashHeavy,
        Underline::DashLong,
        Underline::DashLongHeavy,
        Underline::DotDash,
        Underline::DotDashHeavy,
        Underline::DotDotDash,
        Underline::DotDotDashHeavy,
        Underline::Wavy,
        Underline::WavyHeavy,
        Underline::WavyDouble,
    ];
    for value in values {
        assert_eq!(Underline::from_dml(value.dml()).unwrap(), value);
        assert_eq!(Underline::from_wml(value.wml()).unwrap(), value);
    }
    assert!(Underline::from_dml("single").is_err());
    assert!(Underline::from_wml("sng").is_err());
}

#[test]
fn bounded_values_reject_invalid_authoring_and_xml() {
    assert_eq!(Columns::new(16).unwrap().get(), 16);
    assert!(Columns::new(0).is_err());
    assert!("17".parse::<Columns>().is_err());
    assert_eq!(TextSize::new(100).unwrap().get(), 100);
    assert_eq!(TextSize::new(400_000).unwrap().get(), 400_000);
    assert!(TextSize::new(99).is_err());
    assert!("400001".parse::<TextSize>().is_err());
    assert_eq!(
        Coordinate32::try_from(i64::from(i32::MIN))
            .unwrap()
            .as_emu(),
        Some(i32::MIN)
    );
    assert!(Coordinate32::try_from(i64::from(i32::MAX) + 1).is_err());
    assert_eq!(
        "1.25cm".parse::<Coordinate32>().unwrap().to_string(),
        "1.25cm"
    );
}

#[test]
fn booleans_are_dialect_exact() {
    for (token, expected) in [("1", true), ("true", true), ("0", false), ("false", false)] {
        assert_eq!(parse_bool(token).unwrap(), expected);
    }
    assert!(parse_bool("on").is_err());
    assert!(parse_on_off("on").unwrap());
    assert!(!parse_on_off("off").unwrap());
    assert!(parse_on_off("yes").is_err());
}

#[test]
fn common_values_remain_cache_friendly() {
    assert_eq!(size_of::<Anchor>(), 1);
    assert_eq!(size_of::<Direction>(), 1);
    assert_eq!(size_of::<Wrap>(), 1);
    assert_eq!(size_of::<Autofit>(), 1);
    assert_eq!(size_of::<Underline>(), 1);
    assert_eq!(size_of::<Columns>(), 1);
    assert_eq!(size_of::<TextSize>(), 4);
}
