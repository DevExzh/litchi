//! Native protobuf conversion and copy-on-write mutation for shape shadows.

mod native;
mod style;

pub use litchi_iwa_common::shape::shadow::{
    Angle, Appearance, BlurRadius, Contact, Curve, Curved, Drop, Offset, Opacity, Perspective,
    Shadow,
};
pub(crate) use native::{shadow_from_native, shadow_to_native};
pub(crate) use style::{reset_shape_shadow, set_shape_shadow, shape_shadow};

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::RgbaColor;

    fn appearance(radius: u32, offset: f32, opacity: f32) -> Appearance {
        Appearance::new(
            RgbaColor::black(),
            BlurRadius::from_points(radius).unwrap(),
            Offset::from_points(offset).unwrap(),
            Opacity::new(opacity).unwrap(),
        )
    }

    #[test]
    fn all_native_shadow_families_round_trip() {
        for shadow in [
            Shadow::Disabled,
            Shadow::Drop(Drop::new(
                appearance(7, 11.0, 0.42),
                Angle::from_degrees(135.0).unwrap(),
            )),
            Shadow::Contact(Contact::new(
                appearance(18, 6.0, 0.58),
                Perspective::from_degrees(23.0).unwrap(),
            )),
            Shadow::Curved(Curved::new(
                appearance(15, 4.0, 0.73),
                Angle::from_degrees(310.0).unwrap(),
                Curve::new(0.2).unwrap(),
            )),
        ] {
            assert_eq!(
                shadow_from_native(&shadow_to_native(shadow)).unwrap(),
                shadow
            );
        }
    }

    #[test]
    fn inspector_angles_map_to_native_counterclockwise_values() {
        let angle = Angle::from_degrees(135.0).unwrap();
        assert_eq!(native::angle_to_native(angle), 225.0);
        assert_eq!(native::angle_from_native(225.0).unwrap(), angle);
    }
}
