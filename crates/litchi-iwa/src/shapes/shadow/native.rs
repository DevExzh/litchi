//! Native protobuf conversion for shape shadows.

use crate::protobuf::tsd;
use crate::{Error, Result};

use super::super::color::{color_from_native, color_to_native};
use super::{
    ShapeContactShadow, ShapeCurvedShadow, ShapeDropShadow, ShapeShadow, ShapeShadowAngle,
    ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowCurve, ShapeShadowOffset,
    ShapeShadowOpacity, ShapeShadowPerspective,
};

const FULL_TURN_DEGREES: f32 = 360.0;
const DEFAULT_NATIVE_ANGLE_DEGREES: f32 = 315.0;
const DEFAULT_OFFSET_POINTS: f32 = 5.0;
const DEFAULT_BLUR_RADIUS_POINTS: i32 = 1;
const DEFAULT_OPACITY: f32 = 1.0;
const DEFAULT_CONTACT_HEIGHT: f32 = 0.2;
const DEFAULT_CURVE: f32 = 0.6;

pub(crate) fn shadow_from_native(native: &tsd::ShadowArchive) -> Result<ShapeShadow> {
    if native == &tsd::ShadowArchive::default() {
        return Ok(ShapeShadow::Disabled);
    }
    if native.is_enabled == Some(false) {
        let mut disabled = *native;
        disabled.is_enabled = None;
        if disabled == tsd::ShadowArchive::default() {
            return Ok(ShapeShadow::Disabled);
        }
        return Err(Error::InvalidFormat(
            "disabled iWork shadow unexpectedly retains enabled-shadow properties".to_owned(),
        ));
    }

    let appearance = appearance_from_native(native)?;
    let kind = tsd::shadow_archive::ShadowType::try_from(
        native
            .r#type
            .unwrap_or(tsd::shadow_archive::ShadowType::TsdDropShadow as i32),
    )
    .map_err(|_| Error::InvalidFormat("iWork shape uses an unknown shadow type".to_owned()))?;
    match kind {
        tsd::shadow_archive::ShadowType::TsdDropShadow => {
            if native.contact_shadow.is_some() || native.curved_shadow.is_some() {
                return Err(conflicting_shadow_error());
            }
            Ok(ShapeShadow::Drop(ShapeDropShadow::new(
                appearance,
                angle_from_native(native.angle.unwrap_or(DEFAULT_NATIVE_ANGLE_DEGREES))?,
            )))
        },
        tsd::shadow_archive::ShadowType::TsdContactShadow => {
            if native.drop_shadow.is_some() || native.curved_shadow.is_some() {
                return Err(conflicting_shadow_error());
            }
            let contact = native.contact_shadow.as_ref();
            let perspective = ShapeShadowPerspective::from_native_height(
                contact
                    .and_then(|archive| archive.height)
                    .unwrap_or(DEFAULT_CONTACT_HEIGHT),
            )?;
            let contact_offset = ShapeShadowOffset::from_points(
                contact.and_then(|archive| archive.offset).unwrap_or(0.0),
            )?;
            Ok(ShapeShadow::Contact(
                ShapeContactShadow::new(appearance, perspective)
                    .with_contact_offset(contact_offset),
            ))
        },
        tsd::shadow_archive::ShadowType::TsdCurvedShadow => {
            if native.drop_shadow.is_some() || native.contact_shadow.is_some() {
                return Err(conflicting_shadow_error());
            }
            let curve = native
                .curved_shadow
                .as_ref()
                .and_then(|archive| archive.curve)
                .unwrap_or(DEFAULT_CURVE);
            Ok(ShapeShadow::Curved(ShapeCurvedShadow::new(
                appearance,
                angle_from_native(native.angle.unwrap_or(DEFAULT_NATIVE_ANGLE_DEGREES))?,
                ShapeShadowCurve::new(curve)?,
            )))
        },
    }
}

pub(crate) fn shadow_to_native(shadow: ShapeShadow) -> tsd::ShadowArchive {
    match shadow {
        ShapeShadow::Disabled => tsd::ShadowArchive::default(),
        ShapeShadow::Drop(drop) => tsd::ShadowArchive {
            angle: Some(angle_to_native(drop.angle())),
            r#type: Some(tsd::shadow_archive::ShadowType::TsdDropShadow as i32),
            ..appearance_to_native(drop.appearance())
        },
        ShapeShadow::Contact(contact) => tsd::ShadowArchive {
            angle: Some(0.0),
            r#type: Some(tsd::shadow_archive::ShadowType::TsdContactShadow as i32),
            contact_shadow: Some(tsd::ContactShadowArchive {
                height: Some(contact.perspective().native_height()),
                offset: (contact.contact_offset() != ShapeShadowOffset::ZERO)
                    .then(|| contact.contact_offset().points()),
            }),
            ..appearance_to_native(contact.appearance())
        },
        ShapeShadow::Curved(curved) => tsd::ShadowArchive {
            angle: Some(angle_to_native(curved.angle())),
            r#type: Some(tsd::shadow_archive::ShadowType::TsdCurvedShadow as i32),
            curved_shadow: Some(tsd::CurvedShadowArchive {
                curve: Some(curved.curve().get()),
            }),
            ..appearance_to_native(curved.appearance())
        },
    }
}

fn appearance_from_native(native: &tsd::ShadowArchive) -> Result<ShapeShadowAppearance> {
    let radius = native.radius.unwrap_or(DEFAULT_BLUR_RADIUS_POINTS);
    let radius = u32::try_from(radius)
        .map_err(|_| Error::InvalidFormat("iWork shadow has a negative blur radius".to_owned()))?;
    Ok(ShapeShadowAppearance::new(
        native
            .color
            .as_ref()
            .map(color_from_native)
            .transpose()?
            .unwrap_or_else(super::super::RgbaColor::black),
        ShapeShadowBlurRadius::from_points(radius)?,
        ShapeShadowOffset::from_points(native.offset.unwrap_or(DEFAULT_OFFSET_POINTS))?,
        ShapeShadowOpacity::new(native.opacity.unwrap_or(DEFAULT_OPACITY))?,
    ))
}

fn appearance_to_native(appearance: ShapeShadowAppearance) -> tsd::ShadowArchive {
    tsd::ShadowArchive {
        color: Some(color_to_native(appearance.color())),
        offset: Some(appearance.offset().points()),
        radius: Some(appearance.blur_radius().points() as i32),
        opacity: Some(appearance.opacity().get()),
        is_enabled: Some(true),
        ..Default::default()
    }
}

fn angle_from_native(native_degrees: f32) -> Result<ShapeShadowAngle> {
    if !native_degrees.is_finite() || !(0.0..FULL_TURN_DEGREES).contains(&native_degrees) {
        return Err(Error::InvalidFormat(
            "iWork shadow has a non-canonical native angle".to_owned(),
        ));
    }
    ShapeShadowAngle::from_degrees(if native_degrees == 0.0 {
        0.0
    } else {
        FULL_TURN_DEGREES - native_degrees
    })
}

fn angle_to_native(angle: ShapeShadowAngle) -> f32 {
    if angle == ShapeShadowAngle::ZERO {
        0.0
    } else {
        FULL_TURN_DEGREES - angle.degrees()
    }
}

fn conflicting_shadow_error() -> Error {
    Error::InvalidFormat("iWork shape combines mutually exclusive shadow families".to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::RgbaColor;

    fn appearance(radius: u32, offset: f32, opacity: f32) -> ShapeShadowAppearance {
        ShapeShadowAppearance::new(
            RgbaColor::black(),
            ShapeShadowBlurRadius::from_points(radius).unwrap(),
            ShapeShadowOffset::from_points(offset).unwrap(),
            ShapeShadowOpacity::new(opacity).unwrap(),
        )
    }

    #[test]
    fn all_native_shadow_families_round_trip() {
        for shadow in [
            ShapeShadow::Disabled,
            ShapeShadow::Drop(ShapeDropShadow::new(
                appearance(7, 11.0, 0.42),
                ShapeShadowAngle::from_degrees(135.0).unwrap(),
            )),
            ShapeShadow::Contact(ShapeContactShadow::new(
                appearance(18, 6.0, 0.58),
                ShapeShadowPerspective::from_degrees(23.0).unwrap(),
            )),
            ShapeShadow::Curved(ShapeCurvedShadow::new(
                appearance(15, 4.0, 0.73),
                ShapeShadowAngle::from_degrees(310.0).unwrap(),
                ShapeShadowCurve::new(0.2).unwrap(),
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
        let angle = ShapeShadowAngle::from_degrees(135.0).unwrap();
        assert_eq!(angle_to_native(angle), 225.0);
        assert_eq!(angle_from_native(225.0).unwrap(), angle);
    }
}
