//! Semantic reads of the character and paragraph properties of one
//! `TSWP.ParagraphStyleArchive` (no inheritance).

use litchi_iwa_common::color::Rgba;
use litchi_iwa_common::shape::stroke::{Cap, Join, Stroke};
use litchi_iwa_text::appearance::{Background, Outline, ParagraphBackground, Shadow};
use litchi_iwa_text::character::TextCapitalization;
use litchi_iwa_text::font::{Font, Name};
use litchi_iwa_text::paragraph::border::{Offset as BorderOffset, Sides as BorderSides};
use litchi_iwa_text::paragraph::format::{Border, Borders};

use crate::native::{self, NativeError, NativeResult};
use crate::values::{
    capitalization_from_native, color_from_native, color_to_native, shadow_from_native,
    stroke_from_native, stroke_to_native,
};

fn invalid<T>(reason: &'static str) -> NativeResult<T> {
    Err(NativeError::Invalid(reason))
}

pub(super) fn text_color(properties: &native::CharacterProperties) -> NativeResult<Option<Rgba>> {
    if properties.font_color_null == Some(true) || properties.tsd_fill_null == Some(true) {
        if properties.font_color.is_some() || properties.tsd_fill.is_some() {
            return invalid("native iWork text color is both null and populated");
        }
        return Ok(Some(Rgba::black()));
    }
    let legacy = properties
        .font_color
        .as_ref()
        .map(color_from_native)
        .transpose()?;
    let drawing = properties
        .tsd_fill
        .as_ref()
        .map(|fill| {
            if fill.has_gradient || fill.has_image {
                return invalid("native iWork text fill is not a solid color");
            }
            let Some(color) = fill.color.as_ref() else {
                return invalid("native iWork text fill has no color");
            };
            color_from_native(color)
        })
        .transpose()?;
    match (legacy, drawing) {
        (Some(legacy), Some(drawing)) if legacy != drawing => {
            invalid("native iWork legacy and drawing text colors disagree")
        },
        (Some(color), _) | (_, Some(color)) => Ok(Some(color)),
        (None, None) => Ok(None),
    }
}

pub(super) fn text_font(properties: &native::CharacterProperties) -> NativeResult<Option<Font>> {
    if properties.font_name_null == Some(true) {
        if properties.font_name.is_some() {
            return invalid("native iWork text font is both null and populated");
        }
        return Ok(Some(Font::Default));
    }
    if properties.font_name_null == Some(false) && properties.font_name.is_none() {
        return invalid("native iWork text font has a false null marker without a name");
    }
    properties
        .font_name
        .as_deref()
        .map(|name| {
            Name::new(name)
                .map(Font::Named)
                .map_err(|_| NativeError::Invalid("native iWork font name"))
        })
        .transpose()
}

pub(super) fn text_outline(
    properties: &native::CharacterProperties,
) -> NativeResult<Option<Outline>> {
    if properties.tsd_stroke_null == Some(true) {
        if properties.tsd_stroke.is_some() {
            return invalid("native iWork text outline is both null and populated");
        }
        return Ok(Some(Outline::None));
    }
    if properties.tsd_stroke_null == Some(false) && properties.tsd_stroke.is_none() {
        return invalid("native iWork text outline has a false null marker without a stroke");
    }
    properties
        .tsd_stroke
        .as_ref()
        .map(|stroke| {
            stroke_from_native(stroke).map(|outline| outline.map_or(Outline::None, Outline::Stroke))
        })
        .transpose()
}

pub(super) fn text_shadow(
    properties: &native::CharacterProperties,
) -> NativeResult<Option<Shadow>> {
    if properties.shadow_null == Some(true) {
        if properties.shadow.is_some() {
            return invalid("native iWork text shadow is both null and populated");
        }
        return Ok(Some(Shadow::None));
    }
    if properties.shadow_null == Some(false) && properties.shadow.is_none() {
        return invalid("native iWork text shadow has a false null marker without a shadow");
    }
    properties
        .shadow
        .as_ref()
        .map(|shadow| {
            shadow_from_native(shadow).and_then(|shadow| {
                Shadow::from_shape_shadow(shadow)
                    .map_err(|_| NativeError::Invalid("text appearance supports only drop shadows"))
            })
        })
        .transpose()
}

pub(super) fn text_background(
    properties: &native::CharacterProperties,
) -> NativeResult<Option<Background>> {
    if properties.background_color_null == Some(true) {
        if properties.background_color.is_some() {
            return invalid("native iWork text background is both null and populated");
        }
        return Ok(Some(Background::None));
    }
    if properties.background_color_null == Some(false) && properties.background_color.is_none() {
        return invalid("native iWork text background has a false null marker without a color");
    }
    properties
        .background_color
        .as_ref()
        .map(|color| color_from_native(color).map(Background::Color))
        .transpose()
}

pub(super) fn capitalization(
    properties: &native::CharacterProperties,
) -> NativeResult<Option<TextCapitalization>> {
    let Some(value) = properties.capitalization else {
        if properties.capitalization_uses_linguistics.is_some() {
            return invalid("native iWork linguistic capitalization has no capitalization type");
        }
        return Ok(None);
    };
    capitalization_from_native(value, properties.capitalization_uses_linguistics).map(Some)
}

pub(super) fn paragraph_background(
    properties: &native::ParagraphProperties,
) -> NativeResult<Option<ParagraphBackground>> {
    if properties.fill_null == Some(true) {
        if properties.fill.is_some() {
            return invalid("native iWork paragraph background is both null and populated");
        }
        return Ok(Some(ParagraphBackground::None));
    }
    if properties.fill_null == Some(false) && properties.fill.is_none() {
        return invalid("native iWork paragraph background has a false null marker");
    }
    properties
        .fill
        .as_ref()
        .map(|color| color_from_native(color).map(ParagraphBackground::Color))
        .transpose()
}

const LEGACY_BORDER_TOP: i32 = 1;
const LEGACY_BORDER_BOTTOM: i32 = 2;
const LEGACY_BORDER_ALL: i32 = 4;
const LEGACY_BORDER_LEFT: i32 = 8;
const LEGACY_BORDER_RIGHT: i32 = 16;
const PARAGRAPH_BORDER_ALL_BITS: u8 = 0b1111;
const DEFAULT_PARAGRAPH_BORDER_OFFSET_POINTS: f32 = 6.0;

pub(super) fn paragraph_borders(
    properties: &native::ParagraphProperties,
) -> NativeResult<Option<Borders>> {
    let has_border_fields = properties.deprecated_borders.is_some()
        || properties.historical_rule_offset_null.is_some()
        || properties.historical_rule_offset.is_some()
        || properties.rule_width.is_some()
        || properties.stroke_null.is_some()
        || properties.stroke.is_some()
        || properties.border_positions.is_some()
        || properties.rounded_corners.is_some();
    if !has_border_fields {
        return Ok(None);
    }
    if properties.historical_rule_offset_null == Some(true)
        && properties.historical_rule_offset.is_some()
    {
        return invalid("native paragraph-border offset is both null and populated");
    }
    if properties.historical_rule_offset_null == Some(false)
        && properties.historical_rule_offset.is_none()
    {
        return invalid("native paragraph-border offset has a false null marker without a point");
    }
    if properties.stroke_null == Some(true) && properties.stroke.is_some() {
        return invalid("native paragraph-border stroke is both null and populated");
    }
    if properties.stroke_null == Some(false) && properties.stroke.is_none() {
        return invalid("native paragraph-border stroke has a false null marker without a stroke");
    }
    if properties
        .rule_width
        .is_some_and(|width| !width.is_finite() || width < 0.0)
    {
        return invalid("native legacy paragraph-border width is not finite and nonnegative");
    }
    let legacy_sides = properties
        .deprecated_borders
        .map(sides_from_legacy)
        .transpose()?;
    let modern_sides = properties
        .border_positions
        .map(sides_from_bits)
        .transpose()?;
    if legacy_sides
        .zip(modern_sides)
        .is_some_and(|(legacy, modern)| legacy != modern)
    {
        return invalid("native paragraph-border side encodings disagree");
    }
    let sides = modern_sides.or(legacy_sides).unwrap_or(BorderSides::NONE);
    let stroke = properties
        .stroke
        .as_ref()
        .map(stroke_from_native)
        .transpose()?
        .flatten();
    if sides.is_empty() {
        if stroke.is_some() {
            return invalid("native paragraph border has a visible stroke without sides");
        }
        return Ok(Some(Borders::None));
    }
    if properties.stroke_null == Some(true) {
        return invalid("native paragraph border has sides but an explicit null stroke");
    }
    let Some(stroke) = stroke else {
        return invalid("native paragraph border has sides but no visible stroke");
    };
    if stroke.cap != Cap::Round || stroke.join != Join::Round {
        return invalid("native paragraph border does not use round stroke geometry");
    }
    let offset = match properties.historical_rule_offset.as_ref() {
        Some(point) => {
            #[allow(clippy::float_cmp, reason = "the host requires exactly equal axes")]
            let equal = point.x == point.y;
            if !point.x.is_finite() || !point.y.is_finite() || !equal {
                return invalid("native paragraph-border offset must contain equal finite axes");
            }
            BorderOffset::from_points(point.x + DEFAULT_PARAGRAPH_BORDER_OFFSET_POINTS)
                .map_err(|_| NativeError::Invalid("invalid native paragraph-border offset"))?
        },
        None => BorderOffset::DEFAULT,
    };
    let rounded_corners = properties.rounded_corners.unwrap_or(false) && sides == BorderSides::ALL;
    Border::new(
        stroke.color,
        stroke.width,
        stroke.pattern,
        sides,
        offset,
        rounded_corners,
    )
    .map(|border| Some(Borders::Bordered(border)))
    .map_err(|_| NativeError::Invalid("native paragraph border"))
}

/// The native border fields one `Borders` value writes.
#[derive(Default)]
pub(super) struct NativeBorders {
    pub(super) deprecated_borders: Option<i32>,
    pub(super) historical_rule_offset: Option<native::Point>,
    pub(super) stroke_null: Option<bool>,
    pub(super) stroke: Option<native::Stroke>,
    pub(super) positions: Option<i32>,
    pub(super) rounded_corners: Option<bool>,
}

pub(super) fn borders_to_native(borders: Option<Borders>) -> NativeBorders {
    match borders {
        None => NativeBorders::default(),
        Some(Borders::None) => NativeBorders {
            deprecated_borders: Some(0),
            stroke_null: Some(true),
            positions: Some(sides_bits(BorderSides::NONE)),
            rounded_corners: Some(false),
            ..NativeBorders::default()
        },
        Some(Borders::Bordered(border)) => {
            let offset = (border.offset() != BorderOffset::DEFAULT).then(|| {
                let inset = border.offset().points() - DEFAULT_PARAGRAPH_BORDER_OFFSET_POINTS;
                native::Point { x: inset, y: inset }
            });
            NativeBorders {
                deprecated_borders: Some(legacy_value(border.sides())),
                historical_rule_offset: offset,
                stroke: Some(stroke_to_native(border_stroke(border))),
                positions: Some(sides_bits(border.sides())),
                rounded_corners: Some(border.has_rounded_corners()),
                ..NativeBorders::default()
            }
        },
    }
}

fn border_stroke(border: Border) -> Stroke {
    Stroke::new(border.color(), border.width(), border.pattern())
        .with_cap(Cap::Round)
        .with_join(Join::Round)
}

fn sides_bits(sides: BorderSides) -> i32 {
    i32::from(sides.contains(BorderSides::TOP))
        | (i32::from(sides.contains(BorderSides::BOTTOM)) << 1)
        | (i32::from(sides.contains(BorderSides::LEFT)) << 2)
        | (i32::from(sides.contains(BorderSides::RIGHT)) << 3)
}

fn sides_from_bits(bits: i32) -> NativeResult<BorderSides> {
    let Ok(bits) = u8::try_from(bits) else {
        return invalid("native paragraph-border sides are negative");
    };
    if bits & !PARAGRAPH_BORDER_ALL_BITS != 0 {
        return invalid("native paragraph-border sides contain unknown bits");
    }
    Ok(BorderSides::from_flags([
        bits & 1 != 0,
        bits & 2 != 0,
        bits & 4 != 0,
        bits & 8 != 0,
    ]))
}

fn legacy_value(sides: BorderSides) -> i32 {
    if sides == BorderSides::ALL {
        return LEGACY_BORDER_ALL;
    }
    i32::from(sides.contains(BorderSides::TOP)) * LEGACY_BORDER_TOP
        + i32::from(sides.contains(BorderSides::BOTTOM)) * LEGACY_BORDER_BOTTOM
        + i32::from(sides.contains(BorderSides::LEFT)) * LEGACY_BORDER_LEFT
        + i32::from(sides.contains(BorderSides::RIGHT)) * LEGACY_BORDER_RIGHT
}

fn sides_from_legacy(value: i32) -> NativeResult<BorderSides> {
    if value == LEGACY_BORDER_ALL {
        return Ok(BorderSides::ALL);
    }
    let known = LEGACY_BORDER_TOP | LEGACY_BORDER_BOTTOM | LEGACY_BORDER_LEFT | LEGACY_BORDER_RIGHT;
    if value < 0 || value & !known != 0 {
        return invalid("native paragraph border has an unknown legacy side value");
    }
    Ok(BorderSides::from_flags([
        value & LEGACY_BORDER_TOP != 0,
        value & LEGACY_BORDER_BOTTOM != 0,
        value & LEGACY_BORDER_LEFT != 0,
        value & LEGACY_BORDER_RIGHT != 0,
    ]))
}

/// The canonical native color of a text property (legacy and drawing fill
/// colors are written together).
pub(super) fn native_color(color: Rgba) -> native::Color {
    color_to_native(color)
}
