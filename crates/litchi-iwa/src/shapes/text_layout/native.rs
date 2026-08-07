//! Native protobuf conversion for frame-level shape text layout.

use crate::protobuf::tswp;
use crate::{Error, Result};

use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, VerticalAlignment};

pub(super) fn vertical_alignment_from_native(value: i32) -> Result<VerticalAlignment> {
    match tswp::shape_style_properties_archive::VerticalAlignmentType::try_from(value) {
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignTop) => {
            Ok(VerticalAlignment::Top)
        },
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignMiddle) => {
            Ok(VerticalAlignment::Middle)
        },
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignBottom) => {
            Ok(VerticalAlignment::Bottom)
        },
        Ok(tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignJustify) => {
            Ok(VerticalAlignment::Justified)
        },
        Err(_) => Err(Error::InvalidFormat(
            "iWork shape uses an unknown vertical text alignment".to_owned(),
        )),
    }
}

pub(super) const fn vertical_alignment_to_native(value: VerticalAlignment) -> i32 {
    match value {
        VerticalAlignment::Top => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignTop as i32
        },
        VerticalAlignment::Middle => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignMiddle as i32
        },
        VerticalAlignment::Bottom => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignBottom as i32
        },
        VerticalAlignment::Justified => {
            tswp::shape_style_properties_archive::VerticalAlignmentType::KFrameAlignJustify as i32
        },
    }
}

pub(super) fn insets_from_native(native: &tswp::PaddingArchive) -> Result<Insets> {
    Ok(Insets::new(
        Inset::from_points(native.left.unwrap_or(0.0))?,
        Inset::from_points(native.top.unwrap_or(0.0))?,
        Inset::from_points(native.right.unwrap_or(0.0))?,
        Inset::from_points(native.bottom.unwrap_or(0.0))?,
    ))
}

pub(super) fn insets_to_native(insets: Insets) -> tswp::PaddingArchive {
    tswp::PaddingArchive {
        left: Some(insets.left().points()),
        top: Some(insets.top().points()),
        right: Some(insets.right().points()),
        bottom: Some(insets.bottom().points()),
    }
}

pub(super) const fn auto_size_from_native(value: bool) -> AutoSize {
    if value {
        AutoSize::ShrinkToFit
    } else {
        AutoSize::Fixed
    }
}

pub(super) const fn auto_size_to_native(value: AutoSize) -> bool {
    matches!(value, AutoSize::ShrinkToFit)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_common::text::layout::Layout;

    #[test]
    fn native_text_layout_values_round_trip() {
        for alignment in [
            VerticalAlignment::Top,
            VerticalAlignment::Middle,
            VerticalAlignment::Bottom,
            VerticalAlignment::Justified,
        ] {
            assert_eq!(
                vertical_alignment_from_native(vertical_alignment_to_native(alignment)).unwrap(),
                alignment
            );
        }
        let insets = Insets::new(
            Inset::from_points(1.0).unwrap(),
            Inset::from_points(2.0).unwrap(),
            Inset::from_points(3.0).unwrap(),
            Inset::from_points(4.0).unwrap(),
        );
        assert_eq!(
            insets_from_native(&insets_to_native(insets)).unwrap(),
            insets
        );
        assert_eq!(Layout::default().auto_size(), AutoSize::Fixed);
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
