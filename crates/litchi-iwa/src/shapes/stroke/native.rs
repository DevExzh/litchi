//! Native protobuf conversion for standard iWork strokes.

use crate::protobuf::tsd;
use crate::{Error, Result};

use super::super::color::{RgbaColor, color_from_native, color_to_native};
use super::{Cap, Join, MiterLimit, Pattern, Stroke, Width};

const DEFAULT_MITER_LIMIT: f32 = 4.0;
const SHORT_DASH_LENGTH: f32 = 1.0;
const MEDIUM_DASH_LENGTH: f32 = 2.0;
const LONG_DASH_LENGTH: f32 = 6.0;
const ROUNDED_DOT_LENGTH: f32 = 0.001;
const ROUNDED_DOT_GAP: f32 = 2.0;
const NATIVE_PATTERN_CAPACITY: usize = 6;

pub(crate) fn stroke_from_native(stroke: &tsd::StrokeArchive) -> Result<Option<Stroke>> {
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
        tsd::stroke_pattern_archive::StrokePatternType::TsdSolidPattern => Pattern::Solid,
        tsd::stroke_pattern_archive::StrokePatternType::TsdPattern => {
            let used = usize::try_from(pattern.count.unwrap_or_default()).map_err(|_| {
                Error::InvalidFormat("native iWork stroke pattern count is too large".to_owned())
            })?;
            let values = pattern.pattern.get(..used).ok_or_else(|| {
                Error::InvalidFormat("native iWork stroke pattern is truncated".to_owned())
            })?;
            match values {
                [SHORT_DASH_LENGTH, SHORT_DASH_LENGTH] => Pattern::ShortDash,
                [MEDIUM_DASH_LENGTH, MEDIUM_DASH_LENGTH] => Pattern::MediumDash,
                [LONG_DASH_LENGTH, LONG_DASH_LENGTH] => Pattern::LongDash,
                [ROUNDED_DOT_LENGTH, ROUNDED_DOT_GAP] => Pattern::RoundedDash,
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
    let width = Width::new(
        stroke
            .width
            .ok_or_else(|| Error::InvalidFormat("native iWork stroke has no width".to_owned()))?,
    )?;
    let cap = match tsd::stroke_archive::LineCap::try_from(stroke.cap.unwrap_or_default()) {
        Ok(tsd::stroke_archive::LineCap::ButtCap) => Cap::Butt,
        Ok(tsd::stroke_archive::LineCap::RoundCap) => Cap::Round,
        Ok(tsd::stroke_archive::LineCap::SquareCap) => Cap::Square,
        Err(_) => {
            return Err(Error::InvalidFormat(
                "unknown native iWork stroke cap".to_owned(),
            ));
        },
    };
    let join = match tsd::LineJoin::try_from(stroke.join.unwrap_or_default()) {
        Ok(tsd::LineJoin::MiterJoin) => Join::Miter,
        Ok(tsd::LineJoin::RoundJoin) => Join::Round,
        Ok(tsd::LineJoin::BevelJoin) => Join::Bevel,
        Err(_) => {
            return Err(Error::InvalidFormat(
                "unknown native iWork stroke join".to_owned(),
            ));
        },
    };
    let miter_limit = MiterLimit::new(stroke.miter_limit.unwrap_or(DEFAULT_MITER_LIMIT))?;
    Ok(Some(Stroke {
        color,
        width,
        pattern: typed_pattern,
        cap,
        join,
        miter_limit,
    }))
}

pub(crate) fn stroke_to_native(stroke: Stroke) -> tsd::StrokeArchive {
    tsd::StrokeArchive {
        color: Some(color_to_native(stroke.color)),
        width: Some(stroke.width.points()),
        cap: Some(match stroke.cap {
            Cap::Butt => tsd::stroke_archive::LineCap::ButtCap as i32,
            Cap::Round => tsd::stroke_archive::LineCap::RoundCap as i32,
            Cap::Square => tsd::stroke_archive::LineCap::SquareCap as i32,
        }),
        join: Some(match stroke.join {
            Join::Miter => tsd::LineJoin::MiterJoin as i32,
            Join::Round => tsd::LineJoin::RoundJoin as i32,
            Join::Bevel => tsd::LineJoin::BevelJoin as i32,
        }),
        miter_limit: Some(stroke.miter_limit.ratio()),
        pattern: Some(pattern_to_native(stroke.pattern)),
        ..Default::default()
    }
}

pub(crate) fn empty_stroke_archive() -> tsd::StrokeArchive {
    let mut stroke = stroke_to_native(Stroke::new(RgbaColor::black(), Width::ONE, Pattern::Solid));
    stroke.pattern = Some(tsd::StrokePatternArchive {
        r#type: Some(tsd::stroke_pattern_archive::StrokePatternType::TsdEmptyPattern as i32),
        phase: Some(0.0),
        count: Some(0),
        pattern: vec![0.0; NATIVE_PATTERN_CAPACITY],
    });
    stroke
}

pub(super) fn pattern_to_native(pattern: Pattern) -> tsd::StrokePatternArchive {
    let (pattern_type, count, first, second) = match pattern {
        Pattern::Solid => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdSolidPattern,
            0,
            0.0,
            0.0,
        ),
        Pattern::ShortDash => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdPattern,
            2,
            SHORT_DASH_LENGTH,
            SHORT_DASH_LENGTH,
        ),
        Pattern::MediumDash => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdPattern,
            2,
            MEDIUM_DASH_LENGTH,
            MEDIUM_DASH_LENGTH,
        ),
        Pattern::LongDash => (
            tsd::stroke_pattern_archive::StrokePatternType::TsdPattern,
            2,
            LONG_DASH_LENGTH,
            LONG_DASH_LENGTH,
        ),
        Pattern::RoundedDash => (
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
