//! Conversions between the private native models and the archive-free
//! semantic values of `litchi-iwa-text` and `litchi-iwa-common`.
//!
//! These are the host's (`litchi-iwa`) native adapters, ported unchanged in
//! behaviour: every value the host accepts is accepted, every value it
//! rejects is rejected, and every value it writes is written with the same
//! fields.

use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::shadow::{
    Angle, Appearance, BlurRadius, Contact, Curve, Curved, Drop, Offset, Opacity, Perspective,
    Shadow as ShapeShadow,
};
use litchi_iwa_common::shape::stroke::{Cap, Join, MiterLimit, Pattern, Stroke, Width};
use litchi_iwa_text::character::{
    TextCapitalization, TextCharacterSpacing, TextLigatures, TextScript, TextStrikethrough,
    TextUnderline,
};
use litchi_iwa_text::paragraph::direction::WritingDirection;
use litchi_iwa_text::paragraph::flow::Hyphenation;
use litchi_iwa_text::paragraph::format::{
    Alignment, LineSpacing, LineSpacingMultiple, LineSpacingPoints,
};
use litchi_iwa_text::paragraph::tabs::{
    Alignment as TabAlignment, DecimalCharacter, Leader, Position, Stop, Stops,
};

use crate::native::{self, NativeError, NativeResult};

fn invalid<T>(reason: &'static str) -> NativeResult<T> {
    Err(NativeError::Invalid(reason))
}

const COLOR_MODEL_RGB: i32 = 1;
const RGB_SPACE_SRGB: i32 = 1;
const RGB_SPACE_P3: i32 = 2;

pub(crate) fn color_from_native(color: &native::Color) -> NativeResult<Rgba> {
    if color.model != COLOR_MODEL_RGB
        || color.c.is_some()
        || color.m.is_some()
        || color.y.is_some()
        || color.k.is_some()
        || color.w.is_some()
    {
        return invalid("native iWork color is not RGB");
    }
    let color_space = match color.rgbspace {
        Some(RGB_SPACE_SRGB) => RgbColorSpace::Srgb,
        Some(RGB_SPACE_P3) => RgbColorSpace::DisplayP3,
        Some(_) => return invalid("native iWork color uses an unknown RGB color space"),
        None => return invalid("native iWork color has no RGB color space"),
    };
    let (Some(red), Some(green), Some(blue)) = (color.r, color.g, color.b) else {
        return invalid("native iWork color is missing a channel");
    };
    Rgba::new(red, green, blue, color.a.unwrap_or(1.0), color_space)
        .map_err(|_| NativeError::Invalid("native iWork color channel is out of range"))
}

pub(crate) fn color_to_native(color: Rgba) -> native::Color {
    native::Color {
        model: COLOR_MODEL_RGB,
        r: Some(color.red()),
        g: Some(color.green()),
        b: Some(color.blue()),
        rgbspace: Some(match color.color_space() {
            RgbColorSpace::Srgb => RGB_SPACE_SRGB,
            RgbColorSpace::DisplayP3 => RGB_SPACE_P3,
        }),
        a: Some(color.alpha()),
        ..native::Color::default()
    }
}

const DEFAULT_MITER_LIMIT: f32 = 4.0;
const SHORT_DASH_LENGTH: f32 = 1.0;
const MEDIUM_DASH_LENGTH: f32 = 2.0;
const LONG_DASH_LENGTH: f32 = 6.0;
const ROUNDED_DOT_LENGTH: f32 = 0.001;
const ROUNDED_DOT_GAP: f32 = 2.0;
const NATIVE_PATTERN_CAPACITY: usize = 6;
const PATTERN_DASHED: i32 = 0;
const PATTERN_SOLID: i32 = 1;
const PATTERN_EMPTY: i32 = 2;

/// A standard stroke, `None` for the empty pattern.
pub(crate) fn stroke_from_native(stroke: &native::Stroke) -> NativeResult<Option<Stroke>> {
    if stroke.has_special {
        return invalid("smart, frame, and patterned iWork strokes are not standard strokes");
    }
    let Some(pattern) = stroke.pattern.as_ref() else {
        return invalid("native iWork stroke has no pattern");
    };
    let typed_pattern = match pattern.kind.unwrap_or_default() {
        PATTERN_EMPTY => return Ok(None),
        PATTERN_SOLID => Pattern::Solid,
        PATTERN_DASHED => {
            let used = usize::try_from(pattern.count.unwrap_or_default())
                .map_err(|_| NativeError::Invalid("native iWork stroke pattern is too large"))?;
            let Some(values) = pattern.pattern.get(..used) else {
                return invalid("native iWork stroke pattern is truncated");
            };
            #[allow(clippy::float_cmp, reason = "native dash presets are exact constants")]
            match values {
                [SHORT_DASH_LENGTH, SHORT_DASH_LENGTH] => Pattern::ShortDash,
                [MEDIUM_DASH_LENGTH, MEDIUM_DASH_LENGTH] => Pattern::MediumDash,
                [LONG_DASH_LENGTH, LONG_DASH_LENGTH] => Pattern::LongDash,
                [ROUNDED_DOT_LENGTH, ROUNDED_DOT_GAP] => Pattern::RoundedDash,
                _ => return invalid("unsupported native iWork stroke pattern"),
            }
        },
        _ => return invalid("native iWork stroke has an unknown pattern type"),
    };
    let Some(color) = stroke.color.as_ref() else {
        return invalid("native iWork stroke has no color");
    };
    let color = color_from_native(color)?;
    let Some(width) = stroke.width else {
        return invalid("native iWork stroke has no width");
    };
    let width = Width::new(width).map_err(|_| NativeError::Invalid("native iWork stroke width"))?;
    let cap = match stroke.cap.unwrap_or_default() {
        0 => Cap::Butt,
        1 => Cap::Round,
        2 => Cap::Square,
        _ => return invalid("unknown native iWork stroke cap"),
    };
    let join = match stroke.join.unwrap_or_default() {
        0 => Join::Miter,
        1 => Join::Round,
        2 => Join::Bevel,
        _ => return invalid("unknown native iWork stroke join"),
    };
    let miter_limit = MiterLimit::new(stroke.miter_limit.unwrap_or(DEFAULT_MITER_LIMIT))
        .map_err(|_| NativeError::Invalid("native iWork stroke miter limit"))?;
    Ok(Some(Stroke {
        color,
        width,
        pattern: typed_pattern,
        cap,
        join,
        miter_limit,
    }))
}

pub(crate) fn stroke_to_native(stroke: Stroke) -> native::Stroke {
    native::Stroke {
        color: Some(color_to_native(stroke.color)),
        width: Some(stroke.width.points()),
        cap: Some(match stroke.cap {
            Cap::Butt => 0,
            Cap::Round => 1,
            Cap::Square => 2,
        }),
        join: Some(match stroke.join {
            Join::Miter => 0,
            Join::Round => 1,
            Join::Bevel => 2,
        }),
        miter_limit: Some(stroke.miter_limit.ratio()),
        pattern: Some(pattern_to_native(stroke.pattern)),
        has_special: false,
    }
}

fn pattern_to_native(pattern: Pattern) -> native::StrokePattern {
    let (kind, count, first, second) = match pattern {
        Pattern::Solid => (PATTERN_SOLID, 0, 0.0, 0.0),
        Pattern::ShortDash => (PATTERN_DASHED, 2, SHORT_DASH_LENGTH, SHORT_DASH_LENGTH),
        Pattern::MediumDash => (PATTERN_DASHED, 2, MEDIUM_DASH_LENGTH, MEDIUM_DASH_LENGTH),
        Pattern::LongDash => (PATTERN_DASHED, 2, LONG_DASH_LENGTH, LONG_DASH_LENGTH),
        Pattern::RoundedDash => (PATTERN_DASHED, 2, ROUNDED_DOT_LENGTH, ROUNDED_DOT_GAP),
    };
    let mut values = vec![0.0; NATIVE_PATTERN_CAPACITY];
    values[0] = first;
    values[1] = second;
    native::StrokePattern {
        kind: Some(kind),
        phase: Some(0.0),
        count: Some(count),
        pattern: values,
    }
}

const FULL_TURN_DEGREES: f32 = 360.0;
const DEFAULT_NATIVE_ANGLE_DEGREES: f32 = 315.0;
const DEFAULT_OFFSET_POINTS: f32 = 5.0;
const DEFAULT_BLUR_RADIUS_POINTS: i32 = 1;
const DEFAULT_OPACITY: f32 = 1.0;
const DEFAULT_CONTACT_HEIGHT: f32 = 0.2;
const DEFAULT_CURVE: f32 = 0.6;
const SHADOW_DROP: i32 = 0;
const SHADOW_CONTACT: i32 = 1;
const SHADOW_CURVED: i32 = 2;

pub(crate) fn shadow_from_native(shadow: &native::Shadow) -> NativeResult<ShapeShadow> {
    if shadow == &native::Shadow::default() {
        return Ok(ShapeShadow::Disabled);
    }
    if shadow.is_enabled == Some(false) {
        let mut disabled = *shadow;
        disabled.is_enabled = None;
        if disabled == native::Shadow::default() {
            return Ok(ShapeShadow::Disabled);
        }
        return invalid("disabled iWork shadow retains enabled-shadow properties");
    }
    let radius = u32::try_from(shadow.radius.unwrap_or(DEFAULT_BLUR_RADIUS_POINTS))
        .map_err(|_| NativeError::Invalid("iWork shadow has a negative blur radius"))?;
    let appearance = Appearance::new(
        shadow
            .color
            .as_ref()
            .map(color_from_native)
            .transpose()?
            .unwrap_or_else(Rgba::black),
        BlurRadius::from_points(radius)
            .map_err(|_| NativeError::Invalid("iWork shadow blur radius"))?,
        Offset::from_points(shadow.offset.unwrap_or(DEFAULT_OFFSET_POINTS))
            .map_err(|_| NativeError::Invalid("iWork shadow offset"))?,
        Opacity::new(shadow.opacity.unwrap_or(DEFAULT_OPACITY))
            .map_err(|_| NativeError::Invalid("iWork shadow opacity"))?,
    );
    match shadow.kind.unwrap_or(SHADOW_DROP) {
        SHADOW_DROP => {
            if shadow.contact.is_some() || shadow.curved.is_some() {
                return invalid("iWork shadow combines mutually exclusive families");
            }
            Ok(ShapeShadow::Drop(Drop::new(
                appearance,
                angle_from_native(shadow.angle.unwrap_or(DEFAULT_NATIVE_ANGLE_DEGREES))?,
            )))
        },
        SHADOW_CONTACT => {
            if shadow.drop || shadow.curved.is_some() {
                return invalid("iWork shadow combines mutually exclusive families");
            }
            let (height, offset) = shadow.contact.unwrap_or((None, None));
            let perspective = Perspective::from_height(height.unwrap_or(DEFAULT_CONTACT_HEIGHT))
                .map_err(|_| NativeError::Invalid("iWork contact shadow height"))?;
            let contact_offset = Offset::from_points(offset.unwrap_or(0.0))
                .map_err(|_| NativeError::Invalid("iWork contact shadow offset"))?;
            Ok(ShapeShadow::Contact(
                Contact::new(appearance, perspective).with_contact_offset(contact_offset),
            ))
        },
        SHADOW_CURVED => {
            if shadow.drop || shadow.contact.is_some() {
                return invalid("iWork shadow combines mutually exclusive families");
            }
            let curve = shadow.curved.flatten().unwrap_or(DEFAULT_CURVE);
            Ok(ShapeShadow::Curved(Curved::new(
                appearance,
                angle_from_native(shadow.angle.unwrap_or(DEFAULT_NATIVE_ANGLE_DEGREES))?,
                Curve::new(curve).map_err(|_| NativeError::Invalid("iWork shadow curve"))?,
            )))
        },
        _ => invalid("iWork shape uses an unknown shadow type"),
    }
}

#[allow(
    clippy::cast_possible_wrap,
    reason = "the host writes the bounded blur radius as a native int32"
)]
pub(crate) fn shadow_to_native(shadow: ShapeShadow) -> native::Shadow {
    let appearance_to_native = |appearance: Appearance| native::Shadow {
        color: Some(color_to_native(appearance.color())),
        offset: Some(appearance.offset().points()),
        radius: Some(appearance.blur_radius().points() as i32),
        opacity: Some(appearance.opacity().get()),
        is_enabled: Some(true),
        ..native::Shadow::default()
    };
    match shadow {
        ShapeShadow::Disabled => native::Shadow::default(),
        ShapeShadow::Drop(drop) => native::Shadow {
            angle: Some(angle_to_native(drop.angle())),
            kind: Some(SHADOW_DROP),
            ..appearance_to_native(drop.appearance())
        },
        ShapeShadow::Contact(contact) => native::Shadow {
            angle: Some(0.0),
            kind: Some(SHADOW_CONTACT),
            contact: Some((
                Some(contact.perspective().height()),
                (contact.contact_offset() != Offset::ZERO)
                    .then(|| contact.contact_offset().points()),
            )),
            ..appearance_to_native(contact.appearance())
        },
        ShapeShadow::Curved(curved) => native::Shadow {
            angle: Some(angle_to_native(curved.angle())),
            kind: Some(SHADOW_CURVED),
            curved: Some(Some(curved.curve().get())),
            ..appearance_to_native(curved.appearance())
        },
    }
}

fn angle_from_native(native_degrees: f32) -> NativeResult<Angle> {
    if !native_degrees.is_finite() || !(0.0..FULL_TURN_DEGREES).contains(&native_degrees) {
        return invalid("iWork shadow has a non-canonical native angle");
    }
    #[allow(clippy::float_cmp, reason = "native zero is an exact sentinel")]
    let degrees = if native_degrees == 0.0 {
        0.0
    } else {
        FULL_TURN_DEGREES - native_degrees
    };
    Angle::from_degrees(degrees).map_err(|_| NativeError::Invalid("iWork shadow angle"))
}

fn angle_to_native(angle: Angle) -> f32 {
    if angle == Angle::ZERO {
        0.0
    } else {
        FULL_TURN_DEGREES - angle.degrees()
    }
}

pub(crate) fn underline_from_native(value: i32) -> NativeResult<TextUnderline> {
    match value {
        0 => Ok(TextUnderline::None),
        1 => Ok(TextUnderline::Single),
        2 => Ok(TextUnderline::Double),
        3 => Ok(TextUnderline::Wavy),
        _ => invalid("unsupported native iWork underline type"),
    }
}

pub(crate) const fn underline_to_native(value: TextUnderline) -> i32 {
    match value {
        TextUnderline::None => 0,
        TextUnderline::Single => 1,
        TextUnderline::Double => 2,
        TextUnderline::Wavy => 3,
    }
}

pub(crate) fn strikethrough_from_native(value: i32) -> NativeResult<TextStrikethrough> {
    match value {
        0 => Ok(TextStrikethrough::None),
        1 => Ok(TextStrikethrough::Single),
        2 => Ok(TextStrikethrough::Double),
        3 => Ok(TextStrikethrough::Triple),
        _ => invalid("unsupported native iWork strikethrough type"),
    }
}

pub(crate) const fn strikethrough_to_native(value: TextStrikethrough) -> i32 {
    match value {
        TextStrikethrough::None => 0,
        TextStrikethrough::Single => 1,
        TextStrikethrough::Double => 2,
        TextStrikethrough::Triple => 3,
    }
}

pub(crate) fn script_from_native(value: i32) -> NativeResult<TextScript> {
    match value {
        0 => Ok(TextScript::Normal),
        1 => Ok(TextScript::Superscript),
        2 => Ok(TextScript::Subscript),
        _ => invalid("unsupported native iWork text script"),
    }
}

pub(crate) const fn script_to_native(value: TextScript) -> i32 {
    match value {
        TextScript::Normal => 0,
        TextScript::Superscript => 1,
        TextScript::Subscript => 2,
    }
}

pub(crate) fn ligatures_from_native(value: i32) -> NativeResult<TextLigatures> {
    match value {
        0 => Ok(TextLigatures::RequiredOnly),
        1 => Ok(TextLigatures::Standard),
        2 => Ok(TextLigatures::All),
        _ => invalid("unsupported native iWork ligature policy"),
    }
}

pub(crate) const fn ligatures_to_native(value: TextLigatures) -> i32 {
    match value {
        TextLigatures::RequiredOnly => 0,
        TextLigatures::Standard => 1,
        TextLigatures::All => 2,
    }
}

pub(crate) const fn capitalization_to_native(value: TextCapitalization) -> i32 {
    match value {
        TextCapitalization::None => 0,
        TextCapitalization::AllCaps => 1,
        TextCapitalization::SmallCaps => 2,
        TextCapitalization::TitleCase | TextCapitalization::StartCase => 3,
    }
}

pub(crate) const fn capitalization_uses_linguistics(value: TextCapitalization) -> Option<bool> {
    match value {
        TextCapitalization::TitleCase => Some(true),
        _ => None,
    }
}

pub(crate) const fn capitalization_override_count(value: TextCapitalization) -> u32 {
    match value {
        TextCapitalization::TitleCase => 2,
        _ => 1,
    }
}

pub(crate) fn capitalization_from_native(
    value: i32,
    uses_linguistics: Option<bool>,
) -> NativeResult<TextCapitalization> {
    match (value, uses_linguistics.unwrap_or(false)) {
        (0, false) => Ok(TextCapitalization::None),
        (1, false) => Ok(TextCapitalization::AllCaps),
        (2, false) => Ok(TextCapitalization::SmallCaps),
        (3, true) => Ok(TextCapitalization::TitleCase),
        (3, false) => Ok(TextCapitalization::StartCase),
        (0..=2, true) => invalid("native iWork linguistic capitalization is not title case"),
        _ => invalid("unsupported native iWork capitalization type"),
    }
}

pub(crate) fn character_spacing_to_native(value: TextCharacterSpacing) -> f32 {
    value.percent() / 100.0
}

pub(crate) fn character_spacing_from_native(ratio: f32) -> NativeResult<TextCharacterSpacing> {
    if !ratio.is_finite()
        || !(TextCharacterSpacing::MINIMUM_PERCENT / 100.0
            ..=TextCharacterSpacing::MAXIMUM_PERCENT / 100.0)
            .contains(&ratio)
    {
        return invalid("native iWork character spacing must be between -0.4 and 4.0");
    }
    TextCharacterSpacing::from_percent(ratio * 100.0)
        .map_err(|_| NativeError::Invalid("native iWork character spacing"))
}

pub(crate) fn writing_direction_from_native(value: i32) -> NativeResult<WritingDirection> {
    match value {
        -1 => Ok(WritingDirection::Natural),
        0 => Ok(WritingDirection::LeftToRight),
        1 => Ok(WritingDirection::RightToLeft),
        _ => invalid("unsupported native iWork paragraph writing direction"),
    }
}

pub(crate) const fn writing_direction_to_native(value: WritingDirection) -> i32 {
    match value {
        WritingDirection::Natural => -1,
        WritingDirection::LeftToRight => 0,
        WritingDirection::RightToLeft => 1,
    }
}

pub(crate) const fn hyphenation_from_native(value: bool) -> Hyphenation {
    if value {
        Hyphenation::Automatic
    } else {
        Hyphenation::Prevented
    }
}

pub(crate) const fn hyphenation_to_native(value: Hyphenation) -> bool {
    matches!(value, Hyphenation::Automatic)
}

pub(crate) fn alignment_from_native(value: i32) -> NativeResult<Alignment> {
    match value {
        0 => Ok(Alignment::Natural),
        1 => Ok(Alignment::Right),
        2 => Ok(Alignment::Center),
        3 => Ok(Alignment::Justified),
        4 => Ok(Alignment::Left),
        _ => invalid("unsupported native iWork paragraph alignment"),
    }
}

pub(crate) const fn alignment_to_native(value: Alignment) -> i32 {
    match value {
        Alignment::Natural => 0,
        Alignment::Right => 1,
        Alignment::Center => 2,
        Alignment::Justified => 3,
        Alignment::Left => 4,
    }
}

const RELATIVE_LINE_SPACING_MODE: i32 = 0;
const MINIMUM_LINE_SPACING_MODE: i32 = 1;
const EXACT_LINE_SPACING_MODE: i32 = 2;
const MAXIMUM_LINE_SPACING_MODE: i32 = 3;
const BETWEEN_LINE_SPACING_MODE: i32 = 4;

pub(crate) fn line_spacing_to_native(spacing: LineSpacing) -> native::LineSpacing {
    let points = |mode: i32, points: LineSpacingPoints| native::LineSpacing {
        mode: Some(mode),
        amount: Some(points.points()),
        baseline_rule: None,
    };
    match spacing {
        LineSpacing::Relative(multiple) => native::LineSpacing {
            mode: None,
            amount: Some(multiple.get()),
            baseline_rule: None,
        },
        LineSpacing::AtLeast(value) => points(MINIMUM_LINE_SPACING_MODE, value),
        LineSpacing::Exactly(value) => points(EXACT_LINE_SPACING_MODE, value),
        LineSpacing::Maximum(value) => points(MAXIMUM_LINE_SPACING_MODE, value),
        LineSpacing::Between(value) => points(BETWEEN_LINE_SPACING_MODE, value),
    }
}

pub(crate) fn line_spacing_from_native(spacing: &native::LineSpacing) -> NativeResult<LineSpacing> {
    if spacing.baseline_rule.is_some() {
        return invalid("unsupported native iWork line-spacing baseline rule");
    }
    let mode = spacing.mode.unwrap_or(RELATIVE_LINE_SPACING_MODE);
    if mode == RELATIVE_LINE_SPACING_MODE {
        return LineSpacingMultiple::new(spacing.amount.unwrap_or(1.0))
            .map(LineSpacing::Relative)
            .map_err(|_| NativeError::Invalid("native iWork line-spacing multiple"));
    }
    let Some(amount) = spacing.amount else {
        return invalid("native iWork line-spacing mode has no amount");
    };
    let points = LineSpacingPoints::from_points(amount)
        .map_err(|_| NativeError::Invalid("native iWork line-spacing points"))?;
    match mode {
        MINIMUM_LINE_SPACING_MODE => Ok(LineSpacing::AtLeast(points)),
        EXACT_LINE_SPACING_MODE => Ok(LineSpacing::Exactly(points)),
        MAXIMUM_LINE_SPACING_MODE => Ok(LineSpacing::Maximum(points)),
        BETWEEN_LINE_SPACING_MODE => Ok(LineSpacing::Between(points)),
        _ => invalid("unsupported native iWork line-spacing mode"),
    }
}

fn tab_alignment_from_native(value: i32) -> NativeResult<TabAlignment> {
    match value {
        0 => Ok(TabAlignment::Left),
        1 => Ok(TabAlignment::Center),
        2 => Ok(TabAlignment::Right),
        3 => Ok(TabAlignment::Decimal),
        _ => invalid("unsupported native iWork tab-stop alignment"),
    }
}

const fn tab_alignment_to_native(value: TabAlignment) -> i32 {
    match value {
        TabAlignment::Left => 0,
        TabAlignment::Center => 1,
        TabAlignment::Right => 2,
        TabAlignment::Decimal => 3,
    }
}

pub(crate) fn tabs_from_native(archive: &native::Tabs) -> NativeResult<Stops> {
    let mut stops = Vec::new();
    native::reserve(&mut stops, archive.tabs.len())?;
    for tab in &archive.tabs {
        let Some(position) = tab.position else {
            return invalid("native iWork tab stop has no position");
        };
        let alignment = tab_alignment_from_native(tab.alignment.unwrap_or(0))?;
        let leader = tab
            .leader
            .as_deref()
            .map(Leader::new)
            .transpose()
            .map_err(|_| NativeError::Invalid("native iWork tab leader"))?;
        stops.push(Stop {
            position: Position::from_points(position)
                .map_err(|_| NativeError::Invalid("native iWork tab position"))?,
            alignment,
            leader,
        });
    }
    Stops::new(stops).map_err(|_| NativeError::Invalid("native iWork tab stops"))
}

pub(crate) fn tabs_to_native(stops: &Stops) -> native::Tabs {
    native::Tabs {
        tabs: stops
            .as_slice()
            .iter()
            .map(|stop| native::Tab {
                position: Some(stop.position.points()),
                alignment: (stop.alignment != TabAlignment::Left)
                    .then(|| tab_alignment_to_native(stop.alignment)),
                leader: stop
                    .leader
                    .as_ref()
                    .map(|leader| leader.as_str().to_owned()),
            })
            .collect(),
    }
}

pub(crate) fn decimal_character_from_native(value: &str) -> NativeResult<DecimalCharacter> {
    let mut characters = value.chars();
    let Some(character) = characters.next() else {
        return invalid("native iWork decimal-tab character is empty");
    };
    if characters.next().is_some() {
        return invalid("native iWork decimal-tab character has several scalars");
    }
    DecimalCharacter::new(character)
        .map_err(|_| NativeError::Invalid("native iWork decimal-tab character"))
}
