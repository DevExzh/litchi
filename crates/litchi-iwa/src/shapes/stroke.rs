//! Native protobuf conversion and copy-on-write mutation for shape strokes.

mod native;
mod style;

pub use litchi_iwa_common::shape::stroke::{
    Cap, Join, LineStyle, MiterLimit, Pattern, Stroke, Width,
};
pub(crate) use native::{empty_stroke_archive, stroke_from_native, stroke_to_native};
pub(crate) use style::{reset_shape_stroke, set_shape_stroke, shape_stroke};

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::{RgbColorSpace, RgbaColor};
    use native::pattern_to_native;

    #[test]
    fn standard_patterns_match_real_app_archives() {
        for (pattern, expected) in [
            (Pattern::Solid, &[][..]),
            (Pattern::ShortDash, &[1.0, 1.0][..]),
            (Pattern::MediumDash, &[2.0, 2.0][..]),
            (Pattern::LongDash, &[6.0, 6.0][..]),
            (Pattern::RoundedDash, &[0.001, 2.0][..]),
        ] {
            let native = pattern_to_native(pattern);
            let count = native.count.unwrap() as usize;
            assert_eq!(&native.pattern[..count], expected);
        }
    }

    #[test]
    fn typed_stroke_round_trips_through_native_archive() {
        let stroke = Stroke::new(
            RgbaColor::new(0.1, 0.2, 0.3, 0.8, RgbColorSpace::DisplayP3).unwrap(),
            Width::new(3.5).unwrap(),
            Pattern::RoundedDash,
        );
        assert_eq!(
            stroke_from_native(&stroke_to_native(stroke)).unwrap(),
            Some(stroke)
        );
    }

    #[test]
    fn invalid_scalar_values_are_rejected() {
        assert!(Width::new(0.0).is_err());
        assert!(Width::new(f32::NAN).is_err());
        assert!(MiterLimit::new(f32::INFINITY).is_err());
        assert!(RgbaColor::new(-0.1, 0.0, 0.0, 1.0, RgbColorSpace::Srgb).is_err());
    }
}
