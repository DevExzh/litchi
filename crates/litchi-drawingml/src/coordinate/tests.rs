//! Conformance and representation tests for DrawingML coordinate domains.

use super::*;
use std::mem::size_of;

#[test]
fn unqualified_coordinates_are_canonical_and_bounded() {
    assert_eq!(Coordinate::parse(" +00042 ").unwrap().as_emu(), Some(42));
    assert_eq!(Coordinate::parse("-0").unwrap(), Coordinate::ZERO);
    assert_eq!(
        Coordinate::emu(MIN_EMU).unwrap().to_string(),
        MIN_EMU.to_string()
    );
    assert_eq!(
        Coordinate::emu(MAX_EMU).unwrap().to_string(),
        MAX_EMU.to_string()
    );
    assert!(matches!(
        Coordinate::emu(MIN_EMU - 1),
        Err(ParseError::OutOfRange { .. })
    ));
    assert!(matches!(
        Coordinate::emu(MAX_EMU + 1),
        Err(ParseError::OutOfRange { .. })
    ));
}

#[test]
fn universal_measures_cover_every_unit_without_floating_point() {
    for (source, canonical, unit) in [
        ("001.2500mm", "1.25mm", Unit::Mm),
        ("-2cm", "-2cm", Unit::Cm),
        ("3in", "3in", Unit::Inch),
        ("4pt", "4pt", Unit::Pt),
        ("5pc", "5pc", Unit::Pc),
        ("6pi", "6pi", Unit::Pi),
    ] {
        let coordinate = Coordinate::parse(source).unwrap();
        assert_eq!(coordinate.to_string(), canonical);
        assert_eq!(coordinate.unit(), Some(unit));
    }
    assert_eq!(Coordinate::parse("-0.000mm").unwrap().to_string(), "0mm");
    assert_eq!(
        Coordinate::measure("1.25", Unit::Cm).unwrap().number(),
        Some("1.25")
    );
}

#[test]
fn measures_retain_long_decimal_lexicals_without_float_rounding() {
    let coordinate =
        Coordinate::parse("00000000000012345678901234567890.12345678901234567890123456789mm")
            .unwrap();
    assert_eq!(
        coordinate.to_string(),
        "12345678901234567890.12345678901234567890123456789mm"
    );
    assert_eq!(
        coordinate.number(),
        Some("12345678901234567890.12345678901234567890123456789")
    );
}

#[test]
fn malformed_or_unbounded_spellings_are_rejected() {
    for invalid in [
        "", ".1mm", "1.mm", "+1mm", "1e2mm", "1 px", "1MM", "--1", "1.0",
    ] {
        assert!(Coordinate::parse(invalid).is_err(), "accepted {invalid:?}");
    }
    assert!(matches!(
        Coordinate::parse(&format!("{}mm", "1".repeat(MAX_BYTES))),
        Err(ParseError::TooLong { .. })
    ));
    let oversized = "1".repeat(MAX_BYTES + 1);
    assert_eq!(
        oversized.parse::<Coordinate>(),
        Err(ParseError::TooLong {
            len: MAX_BYTES + 1,
            max: MAX_BYTES,
        })
    );
    assert_eq!(
        Coordinate::try_from(oversized),
        Err(ParseError::TooLong {
            len: MAX_BYTES + 1,
            max: MAX_BYTES,
        })
    );
    assert!("".parse::<Unit>().is_err());
}

#[test]
fn extents_are_nonnegative_bounded_and_representation_compact() {
    assert_eq!(size_of::<Extent>(), size_of::<i64>());
    assert_eq!(Extent::ZERO.as_emu(), 0);
    assert_eq!(Extent::emu(0).unwrap(), Extent::ZERO);
    assert_eq!(Extent::emu(1).unwrap().as_emu(), 1);
    assert_eq!(
        Extent::emu(MAX_EMU).unwrap().to_string(),
        MAX_EMU.to_string()
    );
    assert_eq!(
        Extent::emu(-1),
        Err(ParseError::ExtentOutOfRange { value: -1 })
    );
    assert!(matches!(
        Extent::emu(MAX_EMU + 1),
        Err(ParseError::ExtentOutOfRange { .. })
    ));
}

#[test]
fn extents_accept_only_the_exact_integer_lexical_space() {
    assert_eq!(Extent::parse(" +00042 ").unwrap().as_emu(), 42);
    assert_eq!(Extent::parse("-0").unwrap(), Extent::ZERO);
    assert_eq!(Extent::from(7_u32).as_emu(), 7);
    assert_eq!(i64::from(Extent::try_from(9_i64).unwrap()), 9);
    assert_eq!(Extent::parse("1cm"), Err(ParseError::InvalidExtent));
    assert_eq!(
        ParseError::InvalidExtent.to_string(),
        format!("DrawingML extent must be an integer between 0 and {MAX_EMU}")
    );
    for invalid in ["-1", "0mm", "1.25cm", "1.0", "1e2"] {
        assert!(Extent::parse(invalid).is_err(), "accepted {invalid:?}");
    }
}
