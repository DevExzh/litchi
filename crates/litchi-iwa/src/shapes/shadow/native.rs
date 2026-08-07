//! Native protobuf conversion for shape shadows.

use crate::protobuf::tsd;
use crate::{Error, Result};

use super::super::color::{RgbaColor, color_from_native, color_to_native};
use super::{
    Angle, Appearance, BlurRadius, Contact, Curve, Curved, Drop, Offset, Opacity, Perspective,
    Shadow,
};

const FULL_TURN_DEGREES: f32 = 360.0;
const DEFAULT_NATIVE_ANGLE_DEGREES: f32 = 315.0;
const DEFAULT_OFFSET_POINTS: f32 = 5.0;
const DEFAULT_BLUR_RADIUS_POINTS: i32 = 1;
const DEFAULT_OPACITY: f32 = 1.0;
const DEFAULT_CONTACT_HEIGHT: f32 = 0.2;
const DEFAULT_CURVE: f32 = 0.6;

pub(crate) fn shadow_from_native(native: &tsd::ShadowArchive) -> Result<Shadow> {
    if native == &tsd::ShadowArchive::default() {
        return Ok(Shadow::Disabled);
    }
    if native.is_enabled == Some(false) {
        let mut disabled = *native;
        disabled.is_enabled = None;
        if disabled == tsd::ShadowArchive::default() {
            return Ok(Shadow::Disabled);
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
            Ok(Shadow::Drop(Drop::new(
                appearance,
                angle_from_native(native.angle.unwrap_or(DEFAULT_NATIVE_ANGLE_DEGREES))?,
            )))
        },
        tsd::shadow_archive::ShadowType::TsdContactShadow => {
            if native.drop_shadow.is_some() || native.curved_shadow.is_some() {
                return Err(conflicting_shadow_error());
            }
            let contact = native.contact_shadow.as_ref();
            let perspective = Perspective::from_height(
                contact
                    .and_then(|archive| archive.height)
                    .unwrap_or(DEFAULT_CONTACT_HEIGHT),
            )?;
            let contact_offset =
                Offset::from_points(contact.and_then(|archive| archive.offset).unwrap_or(0.0))?;
            Ok(Shadow::Contact(
                Contact::new(appearance, perspective).with_contact_offset(contact_offset),
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
            Ok(Shadow::Curved(Curved::new(
                appearance,
                angle_from_native(native.angle.unwrap_or(DEFAULT_NATIVE_ANGLE_DEGREES))?,
                Curve::new(curve)?,
            )))
        },
    }
}

pub(crate) fn shadow_to_native(shadow: Shadow) -> tsd::ShadowArchive {
    match shadow {
        Shadow::Disabled => tsd::ShadowArchive::default(),
        Shadow::Drop(drop) => tsd::ShadowArchive {
            angle: Some(angle_to_native(drop.angle())),
            r#type: Some(tsd::shadow_archive::ShadowType::TsdDropShadow as i32),
            ..appearance_to_native(drop.appearance())
        },
        Shadow::Contact(contact) => tsd::ShadowArchive {
            angle: Some(0.0),
            r#type: Some(tsd::shadow_archive::ShadowType::TsdContactShadow as i32),
            contact_shadow: Some(tsd::ContactShadowArchive {
                height: Some(contact.perspective().height()),
                offset: (contact.contact_offset() != Offset::ZERO)
                    .then(|| contact.contact_offset().points()),
            }),
            ..appearance_to_native(contact.appearance())
        },
        Shadow::Curved(curved) => tsd::ShadowArchive {
            angle: Some(angle_to_native(curved.angle())),
            r#type: Some(tsd::shadow_archive::ShadowType::TsdCurvedShadow as i32),
            curved_shadow: Some(tsd::CurvedShadowArchive {
                curve: Some(curved.curve().get()),
            }),
            ..appearance_to_native(curved.appearance())
        },
    }
}

fn appearance_from_native(native: &tsd::ShadowArchive) -> Result<Appearance> {
    let radius = native.radius.unwrap_or(DEFAULT_BLUR_RADIUS_POINTS);
    let radius = u32::try_from(radius)
        .map_err(|_| Error::InvalidFormat("iWork shadow has a negative blur radius".to_owned()))?;
    Ok(Appearance::new(
        native
            .color
            .as_ref()
            .map(color_from_native)
            .transpose()?
            .unwrap_or_else(RgbaColor::black),
        BlurRadius::from_points(radius)?,
        Offset::from_points(native.offset.unwrap_or(DEFAULT_OFFSET_POINTS))?,
        Opacity::new(native.opacity.unwrap_or(DEFAULT_OPACITY))?,
    ))
}

fn appearance_to_native(appearance: Appearance) -> tsd::ShadowArchive {
    tsd::ShadowArchive {
        color: Some(color_to_native(appearance.color())),
        offset: Some(appearance.offset().points()),
        radius: Some(appearance.blur_radius().points() as i32),
        opacity: Some(appearance.opacity().get()),
        is_enabled: Some(true),
        ..Default::default()
    }
}

pub(super) fn angle_from_native(native_degrees: f32) -> Result<Angle> {
    if !native_degrees.is_finite() || !(0.0..FULL_TURN_DEGREES).contains(&native_degrees) {
        return Err(Error::InvalidFormat(
            "iWork shadow has a non-canonical native angle".to_owned(),
        ));
    }
    Ok(Angle::from_degrees(if native_degrees == 0.0 {
        0.0
    } else {
        FULL_TURN_DEGREES - native_degrees
    })?)
}

pub(super) fn angle_to_native(angle: Angle) -> f32 {
    if angle == Angle::ZERO {
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
        assert_eq!(angle_to_native(angle), 225.0);
        assert_eq!(angle_from_native(225.0).unwrap(), angle);
    }
}
