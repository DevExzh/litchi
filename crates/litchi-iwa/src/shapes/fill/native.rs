//! Native protobuf conversion for standard shape fills.

use crate::protobuf::tsd;
use crate::{Error, Result};

use super::super::color::{color_from_native, color_to_native};
use super::{
    ShapeFill, ShapeGradient, ShapeGradientAngle, ShapeGradientKind, ShapeGradientOpacity,
    ShapeGradientStop, ShapeGradientStopMidpoint, ShapeGradientStopPosition,
};

pub(super) fn fill_from_native(fill: &tsd::FillArchive) -> Result<ShapeFill> {
    match (
        fill.color.as_ref(),
        fill.gradient.as_ref(),
        fill.image.as_ref(),
    ) {
        (None, None, None) => Ok(ShapeFill::None),
        (Some(color), None, None) => Ok(ShapeFill::Solid(color_from_native(color)?)),
        (None, Some(gradient), None) => Ok(ShapeFill::Gradient(gradient_from_native(gradient)?)),
        _ => Err(Error::InvalidFormat(
            "native iWork shape fill combines mutually exclusive fill representations".to_owned(),
        )),
    }
}

pub(super) fn fill_to_native(fill: &ShapeFill) -> tsd::FillArchive {
    match fill {
        ShapeFill::None => tsd::FillArchive::default(),
        ShapeFill::Solid(color) => tsd::FillArchive {
            color: Some(color_to_native(*color)),
            ..Default::default()
        },
        ShapeFill::Gradient(gradient) => tsd::FillArchive {
            gradient: Some(gradient_to_native(gradient)),
            ..Default::default()
        },
    }
}

fn gradient_from_native(native: &tsd::GradientArchive) -> Result<ShapeGradient> {
    if native.transformgradient.is_some() {
        return Err(Error::InvalidFormat(
            "transformed iWork shape gradients are not supported".to_owned(),
        ));
    }
    let kind = match native
        .r#type
        .and_then(|value| tsd::gradient_archive::GradientType::try_from(value).ok())
    {
        Some(tsd::gradient_archive::GradientType::Linear) => ShapeGradientKind::Linear,
        Some(tsd::gradient_archive::GradientType::Radial) => ShapeGradientKind::Radial,
        None => {
            return Err(Error::InvalidFormat(
                "native iWork shape gradient has no recognized geometry".to_owned(),
            ));
        },
    };
    let angle = native
        .anglegradient
        .as_ref()
        .and_then(|angle| angle.gradientangle)
        .ok_or_else(|| {
            Error::InvalidFormat("native iWork shape gradient has no angle".to_owned())
        })?;
    let mut stops = Vec::with_capacity(native.stops.len());
    for stop in &native.stops {
        let color = stop.color.as_ref().ok_or_else(|| {
            Error::InvalidFormat("native iWork shape gradient stop has no color".to_owned())
        })?;
        stops.push(ShapeGradientStop::new(
            color_from_native(color)?,
            ShapeGradientStopPosition::new(stop.fraction.ok_or_else(|| {
                Error::InvalidFormat("native iWork shape gradient stop has no position".to_owned())
            })?)?,
            ShapeGradientStopMidpoint::new(stop.inflection.ok_or_else(|| {
                Error::InvalidFormat("native iWork shape gradient stop has no midpoint".to_owned())
            })?)?,
        ));
    }
    ShapeGradient::from_native_parts(
        kind,
        stops,
        ShapeGradientOpacity::new(native.opacity.ok_or_else(|| {
            Error::InvalidFormat("native iWork shape gradient has no opacity".to_owned())
        })?)?,
        native.advanced_gradient.ok_or_else(|| {
            Error::InvalidFormat("native iWork shape gradient has no editing mode".to_owned())
        })?,
        ShapeGradientAngle::from_radians(angle)?,
    )
}

fn gradient_to_native(gradient: &ShapeGradient) -> tsd::GradientArchive {
    tsd::GradientArchive {
        r#type: Some(match gradient.kind() {
            ShapeGradientKind::Linear => tsd::gradient_archive::GradientType::Linear as i32,
            ShapeGradientKind::Radial => tsd::gradient_archive::GradientType::Radial as i32,
        }),
        stops: gradient
            .stops()
            .iter()
            .map(|stop| tsd::gradient_archive::GradientStop {
                color: Some(color_to_native(stop.color())),
                fraction: Some(stop.position().get()),
                inflection: Some(stop.midpoint().get()),
            })
            .collect(),
        opacity: Some(gradient.opacity().get()),
        advanced_gradient: Some(gradient.is_advanced()),
        anglegradient: Some(tsd::AngleGradientArchive {
            gradientangle: Some(gradient.angle().radians()),
        }),
        transformgradient: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::{RgbColorSpace, RgbaColor};

    #[test]
    fn standard_fills_round_trip_through_native_archives() {
        let start = RgbaColor::new(0.2, 0.4, 0.8, 0.75, RgbColorSpace::DisplayP3).unwrap();
        let end = RgbaColor::new(0.9, 0.3, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
        for fill in [
            ShapeFill::None,
            ShapeFill::Solid(start),
            ShapeFill::Gradient(ShapeGradient::linear(
                start,
                end,
                ShapeGradientAngle::from_degrees(270.0).unwrap(),
            )),
            ShapeFill::Gradient(
                ShapeGradient::advanced(
                    ShapeGradientKind::Radial,
                    vec![
                        ShapeGradientStop::new(
                            start,
                            ShapeGradientStopPosition::START,
                            ShapeGradientStopMidpoint::new(0.35).unwrap(),
                        ),
                        ShapeGradientStop::new(
                            end,
                            ShapeGradientStopPosition::END,
                            ShapeGradientStopMidpoint::CENTER,
                        ),
                    ],
                    ShapeGradientOpacity::new(0.8).unwrap(),
                    ShapeGradientAngle::from_degrees(315.0).unwrap(),
                )
                .unwrap(),
            ),
        ] {
            assert_eq!(fill_from_native(&fill_to_native(&fill)).unwrap(), fill);
        }
    }
}
