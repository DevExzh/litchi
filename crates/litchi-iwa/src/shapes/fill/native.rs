//! Native protobuf conversion for standard shape fills.

use crate::protobuf::tsd;
use crate::{Error, Result};

use super::super::color::{color_from_native, color_to_native};
use super::ShapeFill;

pub(super) fn fill_from_native(fill: &tsd::FillArchive) -> Result<ShapeFill> {
    match (
        fill.color.as_ref(),
        fill.gradient.as_ref(),
        fill.image.as_ref(),
    ) {
        (None, None, None) => Ok(ShapeFill::None),
        (Some(color), None, None) => Ok(ShapeFill::Solid(color_from_native(color)?)),
        _ => Err(Error::InvalidFormat(
            "gradient and image iWork fills are not standard solid shape fills".to_owned(),
        )),
    }
}

pub(super) fn fill_to_native(fill: ShapeFill) -> tsd::FillArchive {
    match fill {
        ShapeFill::None => tsd::FillArchive::default(),
        ShapeFill::Solid(color) => tsd::FillArchive {
            color: Some(color_to_native(color)),
            ..Default::default()
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::{RgbColorSpace, RgbaColor};

    #[test]
    fn standard_fills_round_trip_through_native_archives() {
        for fill in [
            ShapeFill::None,
            ShapeFill::Solid(
                RgbaColor::new(0.2, 0.4, 0.8, 0.75, RgbColorSpace::DisplayP3).unwrap(),
            ),
        ] {
            assert_eq!(fill_from_native(&fill_to_native(fill)).unwrap(), fill);
        }
    }
}
