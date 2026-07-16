//! Native protobuf conversion for standard iWork strokes.

use crate::protobuf::{tsd, tsp};
use crate::{Error, Result};

use super::{
    RgbColorSpace, RgbaColor, ShapeStroke, StrokeCap, StrokeJoin, StrokeMiterLimit, StrokePattern,
    StrokeWidth,
};

const DEFAULT_MITER_LIMIT: f32 = 4.0;
const SHORT_DASH_LENGTH: f32 = 1.0;
const MEDIUM_DASH_LENGTH: f32 = 2.0;
const LONG_DASH_LENGTH: f32 = 6.0;
const ROUNDED_DOT_LENGTH: f32 = 0.001;
const ROUNDED_DOT_GAP: f32 = 2.0;
const NATIVE_PATTERN_CAPACITY: usize = 6;

pub(super) fn stroke_from_native(stroke: &tsd::StrokeArchive) -> Result<Option<ShapeStroke>> {
    if stroke.smart_stroke.is_some() || stroke.frame.is_some() || stroke.patterned_stroke.is_some()
    {
        return Err(Error::InvalidFormat(
            "smart, frame, and patterned iWork strokes are not standard shape strokes".to_owned(),
        ));
    }
    let pattern = stroke
        .pattern
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("native iWork stroke has no pattern".to_owned()))?;
    let pattern_type = tsd::stroke_pattern_archive::StrokePatternType::try_from(
        pattern.r#type.unwrap_or_default(),
    )
    .map_err(|_| {
        Error::InvalidFormat("native iWork stroke has an unknown pattern type".to_owned())
    })?;
    if pattern_type == tsd::stroke_pattern_archive::StrokePatternType::TsdEmptyPattern {
        return Ok(None);
    }
    let typed_pattern = match pattern_type {
        tsd::stroke_pattern_archive::StrokePatternType::TsdSolidPattern => StrokePattern::Solid,
        tsd::stroke_pattern_archive::StrokePatternType::TsdPattern => {
            let used = usize::try_from(pattern.count.unwrap_or_default()).map_err(|_| {
                Error::InvalidFormat("native iWork stroke pattern count is too large".to_owned())
            })?;
            let values = pattern.pattern.get(..used).ok_or_else(|| {
                Error::InvalidFormat("native iWork stroke pattern is truncated".to_owned())
            })?;
            match values {
                [SHORT_DASH_LENGTH, SHORT_DASH_LENGTH] => StrokePattern::ShortDash,
                [MEDIUM_DASH_LENGTH, MEDIUM_DASH_LENGTH] => StrokePattern::MediumDash,
                [LONG_DASH_LENGTH, LONG_DASH_LENGTH] => StrokePattern::LongDash,
                [ROUNDED_DOT_LENGTH, ROUNDED_DOT_GAP] => StrokePattern::RoundedDash,
                _ => {
                    return Err(Error::InvalidFormat(format!(
                        "unsupported native iWork stroke pattern {values:?}"
                    )));
                },
            }
        },
        tsd::stroke_pattern_archive::StrokePatternType::TsdEmptyPattern => unreachable!(),
    };
    let color = color_from_native(
        stroke
            .color
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("native iWork stroke has no color".to_owned()))?,
    )?;
    let width = StrokeWidth::new(
        stroke
            .width
            .ok_or_else(|| Error::InvalidFormat("native iWork stroke has no width".to_owned()))?,
    )?;
    let cap = match tsd::stroke_archive::LineCap::try_from(stroke.cap.unwrap_or_default()) {
        Ok(tsd::stroke_archive::LineCap::ButtCap) => StrokeCap::Butt,
        Ok(tsd::stroke_archive::LineCap::RoundCap) => StrokeCap::Round,
        Ok(tsd::stroke_archive::LineCap::SquareCap) => StrokeCap::Square,
        Err(_) => {
            return Err(Error::InvalidFormat(
                "unknown native iWork stroke cap".to_owned(),
            ));
        },
    };
    let join = match tsd::LineJoin::try_from(stroke.join.unwrap_or_default()) {
        Ok(tsd::LineJoin::MiterJoin) => StrokeJoin::Miter,
        Ok(tsd::LineJoin::RoundJoin) => StrokeJoin::Round,
        Ok(tsd::LineJoin::BevelJoin) => StrokeJoin::Bevel,
        Err(_) => {
            return Err(Error::InvalidFormat(
                "unknown native iWork stroke join".to_owned(),
            ));
        },
    };
    let miter_limit = StrokeMiterLimit::new(stroke.miter_limit.unwrap_or(DEFAULT_MITER_LIMIT))?;
    Ok(Some(ShapeStroke {
        color,
        width,
        pattern: typed_pattern,
        cap,
        join,
        miter_limit,
    }))
}

pub(super) fn stroke_to_native(stroke: ShapeStroke) -> tsd::StrokeArchive {
    tsd::StrokeArchive {
        color: Some(color_to_native(stroke.color)),
        width: Some(stroke.width.points()),
        cap: Some(match stroke.cap {
            StrokeCap::Butt => tsd::stroke_archive::LineCap::ButtCap as i32,
            StrokeCap::Round => tsd::stroke_archive::LineCap::RoundCap as i32,
            StrokeCap::Square => tsd::stroke_archive::LineCap::SquareCap as i32,
        }),
        join: Some(match stroke.join {
            StrokeJoin::Miter => tsd::LineJoin::MiterJoin as i32,
            StrokeJoin::Round => tsd::LineJoin::RoundJoin as i32,
            StrokeJoin::Bevel => tsd::LineJoin::BevelJoin as i32,
        }),
        miter_limit: Some(stroke.miter_limit.ratio()),
        pattern: Some(pattern_to_native(stroke.pattern)),
        ..Default::default()
    }
}

pub(super) fn empty_stroke_archive() -> tsd::StrokeArchive {
    let mut stroke = stroke_to_native(ShapeStroke::new(
        RgbaColor::black(),
        StrokeWidth(1.0),
        StrokePattern::Solid,
    ));
    stroke.pattern = Some(tsd::StrokePatternArchive {
        r#type: Some(tsd::stroke_pattern_archive::StrokePatternType::TsdEmptyPattern as i32),
        phase: Some(0.0),
        count: Some(0),
        pattern: vec![0.0; NATIVE_PATTERN_CAPACITY],
    });
    stroke
}

pub(super) fn pattern_to_native(pattern: StrokePattern) -> tsd::StrokePatternArchive {
    let (pattern_type, count, first, second) = match pattern {
        StrokePattern::Solid => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdSolidPattern,
            0,
            0.0,
            0.0,
        ),
        StrokePattern::ShortDash => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdPattern,
            2,
            SHORT_DASH_LENGTH,
            SHORT_DASH_LENGTH,
        ),
        StrokePattern::MediumDash => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdPattern,
            2,
            MEDIUM_DASH_LENGTH,
            MEDIUM_DASH_LENGTH,
        ),
        StrokePattern::LongDash => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdPattern,
            2,
            LONG_DASH_LENGTH,
            LONG_DASH_LENGTH,
        ),
        StrokePattern::RoundedDash => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdPattern,
            2,
            ROUNDED_DOT_LENGTH,
            ROUNDED_DOT_GAP,
        ),
    };
    let mut values = vec![0.0; NATIVE_PATTERN_CAPACITY];
    values[0] = first;
    values[1] = second;
    tsd::StrokePatternArchive {
        r#type: Some(pattern_type as i32),
        phase: Some(0.0),
        count: Some(count),
        pattern: values,
    }
}

fn color_from_native(color: &tsp::Color) -> Result<RgbaColor> {
    if color.model != tsp::color::ColorModel::Rgb as i32
        || color.c.is_some()
        || color.m.is_some()
        || color.y.is_some()
        || color.k.is_some()
        || color.w.is_some()
    {
        return Err(Error::InvalidFormat(
            "native iWork stroke color is not RGB".to_owned(),
        ));
    }
    let color_space =
        match tsp::color::RgbColorSpace::try_from(color.rgbspace.ok_or_else(|| {
            Error::InvalidFormat("native iWork stroke color has no RGB color space".to_owned())
        })?) {
            Ok(tsp::color::RgbColorSpace::Srgb) => RgbColorSpace::Srgb,
            Ok(tsp::color::RgbColorSpace::P3) => RgbColorSpace::DisplayP3,
            Err(_) => {
                return Err(Error::InvalidFormat(
                    "native iWork stroke uses an unknown RGB color space".to_owned(),
                ));
            },
        };
    RgbaColor::new(
        color
            .r
            .ok_or_else(|| Error::InvalidFormat("stroke color has no red channel".to_owned()))?,
        color
            .g
            .ok_or_else(|| Error::InvalidFormat("stroke color has no green channel".to_owned()))?,
        color
            .b
            .ok_or_else(|| Error::InvalidFormat("stroke color has no blue channel".to_owned()))?,
        color.a.unwrap_or(1.0),
        color_space,
    )
}

fn color_to_native(color: RgbaColor) -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(color.red()),
        g: Some(color.green()),
        b: Some(color.blue()),
        rgbspace: Some(match color.color_space() {
            RgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
            RgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
        }),
        a: Some(color.alpha()),
        ..Default::default()
    }
}
