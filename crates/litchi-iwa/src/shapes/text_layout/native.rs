//! Native protobuf conversion for frame-level shape text layout.

use crate::protobuf::tswp;
use crate::{Error, Result};

use super::{ShapeTextAutoSize, ShapeTextInset, ShapeTextInsets, ShapeTextVerticalAlignment};

pub(super) fn vertical_alignment_from_native(value: i32) -> Result<ShapeTextVerticalAlignment> {
    match tswp::shape_style_properties_archive::VerticalAlignmentType::try_from(value) {
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignTop) => {
            Ok(ShapeTextVerticalAlignment::Top)
        },
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignMiddle) => {
            Ok(ShapeTextVerticalAlignment::Middle)
        },
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignBottom) => {
            Ok(ShapeTextVerticalAlignment::Bottom)
        },
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignJustify) => {
            Ok(ShapeTextVerticalAlignment::Justified)
        },
        Err(_) => Err(Error::InvalidFormat(
            "iWork shape uses an unknown vertical text alignment".to_owned(),
        )),
    }
}

pub(super) const fn vertical_alignment_to_native(value: ShapeTextVerticalAlignment) -> i32 {
    match value {
        ShapeTextVerticalAlignment::Top => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignTop as i32
        },
        ShapeTextVerticalAlignment::Middle => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignMiddle as i32
        },
        ShapeTextVerticalAlignment::Bottom => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignBottom as i32
        },
        ShapeTextVerticalAlignment::Justified => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignJustify as i32
        },
    }
}

pub(super) fn insets_from_native(native: &tswp::PaddingArchive) -> Result<ShapeTextInsets> {
    Ok(ShapeTextInsets::new(
        ShapeTextInset::from_points(native.left.unwrap_or(0.0))?,
        ShapeTextInset::from_points(native.top.unwrap_or(0.0))?,
        ShapeTextInset::from_points(native.right.unwrap_or(0.0))?,
        ShapeTextInset::from_points(native.bottom.unwrap_or(0.0))?,
    ))
}

pub(super) fn insets_to_native(insets: ShapeTextInsets) -> tswp::PaddingArchive {
    tswp::PaddingArchive {
        left: Some(insets.left().points()),
        top: Some(insets.top().points()),
        right: Some(insets.right().points()),
        bottom: Some(insets.bottom().points()),
    }
}

pub(super) const fn auto_size_from_native(value: bool) -> ShapeTextAutoSize {
    if value {
        ShapeTextAutoSize::ShrinkToFit
    } else {
        ShapeTextAutoSize::Fixed
    }
}

pub(super) const fn auto_size_to_native(value: ShapeTextAutoSize) -> bool {
    matches!(value, ShapeTextAutoSize::ShrinkToFit)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_text_layout_values_round_trip() {
        for alignment in [
            ShapeTextVerticalAlignment::Top,
            ShapeTextVerticalAlignment::Middle,
            ShapeTextVerticalAlignment::Bottom,
            ShapeTextVerticalAlignment::Justified,
        ] {
            assert_eq!(
                vertical_alignment_from_native(vertical_alignment_to_native(alignment)).unwrap(),
                alignment
            );
        }
        let insets = ShapeTextInsets::new(
            ShapeTextInset::from_points(1.0).unwrap(),
            ShapeTextInset::from_points(2.0).unwrap(),
            ShapeTextInset::from_points(3.0).unwrap(),
            ShapeTextInset::from_points(4.0).unwrap(),
        );
        assert_eq!(
            insets_from_native(&insets_to_native(insets)).unwrap(),
            insets
        );
    }

    #[test]
    fn malformed_native_layout_values_are_rejected() {
        assert!(vertical_alignment_from_native(i32::MAX).is_err());
        assert!(
            insets_from_native(&tswp::PaddingArchive {
                left: Some(f32::NAN),
                ..Default::default()
            })
            .is_err()
        );
        assert!(
            insets_from_native(&tswp::PaddingArchive {
                bottom: Some(-1.0),
                ..Default::default()
            })
            .is_err()
        );
    }
}
