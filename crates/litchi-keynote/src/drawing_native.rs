//! Field pass between the stored drawing values of
//! `litchi_iwa_protos::shape_style_codec::native` (decoders `color`,
//! `stroke`, `fill`; encoders `encode_stroke`, `encode_fill`) and their
//! semantic form in `litchi_iwa_common::shape::native` (`color`, `stroke`,
//! `stop`, `gradient`, `image`, `fill`; `stored_stroke`, `stored_fill`).
//!
//! The crate graph keeps the two crates apart, so this plain 1:1 pass is
//! repeated, identically, in `litchi-pages`, `litchi-numbers` and
//! `litchi-keynote` (record 0375; W3 may fold it). The unit test compares the
//! three copies.

use litchi_iwa_common::color::Rgba;
use litchi_iwa_common::shape::fill::{Fill, Gradient, Image};
use litchi_iwa_common::shape::native::{self as rules, Error, StoredColor, StoredPattern};
use litchi_iwa_common::shape::stroke::Stroke;
use litchi_iwa_protos::shape_style_codec::native as stored;

/// A stored `TSP.Color`.
pub fn color(value: &stored::Color) -> Result<Rgba, Error> {
    rules::color(
        value.model,
        value.rgb_space,
        value.rgba,
        value.has_other_channels,
    )
}

fn optional_color(value: Option<&stored::Color>) -> Result<Option<Rgba>, Error> {
    value.map(color).transpose()
}

/// A stored `TSD.StrokeArchive` (`None` for a hidden stroke).
pub fn stroke(value: &stored::Stroke) -> Result<Option<Stroke>, Error> {
    rules::stroke(
        value.color.as_ref().map(color),
        value.width,
        value.cap,
        value.join,
        value.miter_limit,
        value.pattern.as_ref().map(|pattern| StoredPattern {
            kind: pattern.kind,
            count: pattern.count,
            lengths: &pattern.lengths,
        }),
        value.is_special,
    )
}

/// A stored `TSD.GradientArchive`.
pub fn gradient(value: &stored::Gradient) -> Result<Gradient, Error> {
    let stops = value
        .stops
        .iter()
        .map(|stop| {
            rules::stop(
                optional_color(stop.color.as_ref())?,
                stop.fraction,
                stop.inflection,
            )
        })
        .collect::<Result<Vec<_>, Error>>()?;
    rules::gradient(
        value.kind,
        stops,
        value.opacity,
        value.advanced,
        value.angle,
        value.is_transformed,
    )
}

/// A stored `TSD.ImageFillArchive`.
pub fn image(value: &stored::ImageFill) -> Result<Image, Error> {
    rules::image(
        value.data,
        value.technique,
        value.fill_size,
        optional_color(value.tint.as_ref())?,
        optional_color(value.reference_color.as_ref())?,
        value.generic,
        value.is_legacy,
    )
}

/// A stored `TSD.FillArchive`.
pub fn fill(value: &stored::Fill) -> Result<Fill, Error> {
    rules::fill(
        optional_color(value.color.as_ref())?,
        value.gradient.as_ref().map(gradient).transpose()?,
        value.image.as_ref().map(image).transpose()?,
    )
}

fn to_stored_color(value: StoredColor) -> stored::Color {
    stored::Color {
        model: Some(value.model),
        rgba: value.rgba.map(Some),
        rgb_space: Some(value.rgb_space),
        has_other_channels: false,
    }
}

/// The encoded `TSD.StrokeArchive` of a stroke (`None`: the hidden stroke).
pub fn encode_stroke(value: Option<Stroke>) -> Vec<u8> {
    let value = rules::stored_stroke(value);
    stored::encode_stroke(&stored::Stroke {
        color: Some(to_stored_color(value.color)),
        width: Some(value.width),
        cap: Some(value.cap),
        join: Some(value.join),
        miter_limit: Some(value.miter_limit),
        pattern: Some(stored::StrokePattern {
            kind: Some(value.pattern_kind),
            phase: Some(0.0),
            count: Some(value.pattern_count),
            lengths: value.pattern_lengths.to_vec(),
        }),
        is_special: false,
    })
}

/// The encoded `TSD.FillArchive` of a fill.
pub fn encode_fill(value: &Fill) -> Vec<u8> {
    let value = rules::stored_fill(value);
    stored::encode_fill(&stored::Fill {
        color: value.color.map(to_stored_color),
        gradient: value.gradient.map(|gradient| stored::Gradient {
            kind: Some(gradient.kind),
            stops: gradient
                .stops
                .iter()
                .map(|stop| stored::GradientStop {
                    color: Some(to_stored_color(stop.color)),
                    fraction: Some(stop.fraction),
                    inflection: Some(stop.inflection),
                })
                .collect(),
            opacity: Some(gradient.opacity),
            advanced: Some(gradient.advanced),
            angle: Some(gradient.angle),
            is_transformed: false,
        }),
        image: value.image.map(|image| stored::ImageFill {
            data: Some(image.data),
            technique: Some(image.technique),
            tint: image.tint.map(to_stored_color),
            fill_size: Some(image.fill_size),
            generic: Some(image.generic),
            reference_color: image.reference_color.map(to_stored_color),
            is_legacy: false,
        }),
    })
}

#[cfg(test)]
mod tests {
    use litchi_iwa_common::shape::fill::Fill;
    use litchi_iwa_protos::chart_style_codec::Limits;
    use litchi_iwa_protos::shape_style_codec::native as stored;

    const COPIES: [&str; 3] = [
        include_str!("../../litchi-pages/src/drawing_native.rs"),
        include_str!("../../litchi-numbers/src/drawing_native.rs"),
        include_str!("../../litchi-keynote/src/drawing_native.rs"),
    ];

    #[test]
    fn the_three_copies_are_identical() {
        assert_eq!(COPIES[0], COPIES[1]);
        assert_eq!(COPIES[1], COPIES[2]);
    }

    #[test]
    fn a_stored_color_fill_round_trips() {
        // TSD.FillArchive { color: TSP.Color { model: rgb, r 1, g 0, b 0, a 1, rgbspace srgb } }
        let color = [
            0x08, 0x01, 0x1d, 0x00, 0x00, 0x80, 0x3f, 0x25, 0x00, 0x00, 0x00, 0x00, 0x2d, 0x00,
            0x00, 0x00, 0x00, 0x35, 0x00, 0x00, 0x80, 0x3f, 0x60, 0x01,
        ];
        let mut data = vec![0x0a, u8::try_from(color.len()).unwrap()];
        data.extend_from_slice(&color);
        let decoded = stored::fill(&data, Limits::DEFAULT).unwrap();
        let Fill::Solid(red) = super::fill(&decoded).unwrap() else {
            panic!("not a solid fill");
        };
        assert_eq!((red.red(), red.green(), red.alpha()), (1.0, 0.0, 1.0));
        assert_eq!(
            super::fill(&stored::fill(&[], Limits::DEFAULT).unwrap()),
            Ok(Fill::None)
        );
    }
}
