//! The managed overrides of a paragraph-style variation: recognised only
//! when the variation holds exactly the fields the engine writes, and
//! written canonically.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::color::Rgba;
use litchi_iwa_text::appearance::{Background, Outline, ParagraphBackground, Shadow};
use litchi_iwa_text::character::{
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextLigatures, TextPointSize,
    TextScript, TextStrikethrough, TextUnderline,
};
use litchi_iwa_text::font::Font;
use litchi_iwa_text::paragraph::direction::WritingDirection;
use litchi_iwa_text::paragraph::flow::Hyphenation;
use litchi_iwa_text::paragraph::format::{
    Alignment, Borders, IndentPoints, LineSpacing, SpacingPoints,
};
use litchi_iwa_text::paragraph::style::{ParagraphFollowingStyle, raw};
use litchi_iwa_text::paragraph::tabs::{
    Alignment as TabAlignment, DecimalCharacter, DefaultInterval, Stops,
};

use super::ManagedFields;
use super::convert;
use crate::native::{
    self, Budget, NativeError, NativeResult, has_canonical_bool_field, has_exact_fields,
    required_payload,
};
use crate::values;

const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_CHARACTER_PROPERTIES_FIELD: u32 = 11;
const STYLE_PARAGRAPH_PROPERTIES_FIELD: u32 = 12;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const CHARACTER_BOLD_FIELD: u32 = 1;
const CHARACTER_ITALIC_FIELD: u32 = 2;
const CHARACTER_FONT_SIZE_FIELD: u32 = 3;
const CHARACTER_FONT_NAME_NULL_FIELD: u32 = 4;
const CHARACTER_FONT_NAME_FIELD: u32 = 5;
const CHARACTER_FONT_COLOR_FIELD: u32 = 7;
const CHARACTER_SCRIPT_FIELD: u32 = 10;
const CHARACTER_UNDERLINE_FIELD: u32 = 11;
const CHARACTER_STRIKETHROUGH_FIELD: u32 = 12;
const CHARACTER_CAPITALIZATION_FIELD: u32 = 13;
const CHARACTER_BASELINE_SHIFT_FIELD: u32 = 14;
const CHARACTER_LIGATURES_FIELD: u32 = 16;
const CHARACTER_SHADOW_NULL_FIELD: u32 = 20;
const CHARACTER_SHADOW_FIELD: u32 = 21;
const CHARACTER_BACKGROUND_COLOR_NULL_FIELD: u32 = 25;
const CHARACTER_BACKGROUND_COLOR_FIELD: u32 = 26;
const CHARACTER_TRACKING_FIELD: u32 = 27;
const CHARACTER_WRITING_DIRECTION_FIELD: u32 = 35;
const CHARACTER_CAPITALIZATION_LINGUISTICS_FIELD: u32 = 41;
const CHARACTER_DRAWING_STROKE_NULL_FIELD: u32 = 43;
const CHARACTER_DRAWING_STROKE_FIELD: u32 = 44;
const CHARACTER_DRAWING_FILL_FIELD: u32 = 46;
const DRAWING_FILL_COLOR_FIELD: u32 = 1;
const COLOR_FIELDS: [u32; 6] = [1, 3, 4, 5, 6, 12];
const PARAGRAPH_ALIGNMENT_FIELD: u32 = 1;
const PARAGRAPH_DECIMAL_TAB_NULL_FIELD: u32 = 2;
const PARAGRAPH_DECIMAL_TAB_FIELD: u32 = 3;
const PARAGRAPH_DEFAULT_TAB_INTERVAL_FIELD: u32 = 4;
const PARAGRAPH_FILL_NULL_FIELD: u32 = 5;
const PARAGRAPH_FILL_FIELD: u32 = 6;
const PARAGRAPH_FIRST_LINE_INDENT_FIELD: u32 = 7;
const PARAGRAPH_HYPHENATE_FIELD: u32 = 8;
const PARAGRAPH_KEEP_LINES_TOGETHER_FIELD: u32 = 9;
const PARAGRAPH_KEEP_WITH_NEXT_FIELD: u32 = 10;
const PARAGRAPH_LEFT_INDENT_FIELD: u32 = 11;
const PARAGRAPH_LINE_SPACING_FIELD: u32 = 13;
const PARAGRAPH_PAGE_BREAK_BEFORE_FIELD: u32 = 14;
const PARAGRAPH_DEPRECATED_BORDERS_FIELD: u32 = 15;
const PARAGRAPH_BORDER_OFFSET_NULL_FIELD: u32 = 16;
const PARAGRAPH_BORDER_OFFSET_FIELD: u32 = 17;
const PARAGRAPH_LEGACY_RULE_WIDTH_FIELD: u32 = 18;
const PARAGRAPH_RIGHT_INDENT_FIELD: u32 = 19;
const PARAGRAPH_SPACE_AFTER_FIELD: u32 = 20;
const PARAGRAPH_SPACE_BEFORE_FIELD: u32 = 21;
const PARAGRAPH_TABS_FIELD: u32 = 25;
const PARAGRAPH_WIDOW_CONTROL_FIELD: u32 = 26;
const PARAGRAPH_BORDER_STROKE_NULL_FIELD: u32 = 31;
const PARAGRAPH_BORDER_STROKE_FIELD: u32 = 32;
const PARAGRAPH_FOLLOWING_STYLE_NULL_FIELD: u32 = 41;
const PARAGRAPH_FOLLOWING_STYLE_FIELD: u32 = 42;
const PARAGRAPH_BORDER_POSITIONS_FIELD: u32 = 45;
const PARAGRAPH_BORDER_ROUNDED_CORNERS_FIELD: u32 = 46;
const LINE_SPACING_MODE_FIELD: u32 = 1;
const LINE_SPACING_AMOUNT_FIELD: u32 = 2;
const TABS_TAB_FIELD: u32 = 1;
const TAB_POSITION_FIELD: u32 = 1;
const TAB_ALIGNMENT_FIELD: u32 = 2;
const TAB_LEADER_FIELD: u32 = 3;

/// The overrides a managed paragraph-style variation holds.
#[derive(Debug, Clone, PartialEq, Default)]
pub(super) struct Overrides {
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) point_size: Option<TextPointSize>,
    pub(super) font: Option<Font>,
    pub(super) font_color: Option<Rgba>,
    pub(super) capitalization: Option<TextCapitalization>,
    pub(super) script: Option<TextScript>,
    pub(super) baseline_shift: Option<TextBaselineShift>,
    pub(super) character_spacing: Option<TextCharacterSpacing>,
    pub(super) ligatures: Option<TextLigatures>,
    pub(super) outline: Option<Outline>,
    pub(super) shadow: Option<Shadow>,
    pub(super) background: Option<Background>,
    pub(super) paragraph_background: Option<ParagraphBackground>,
    pub(super) paragraph_borders: Option<Borders>,
    pub(super) hyphenation: Option<Hyphenation>,
    pub(super) keep_lines_together: Option<bool>,
    pub(super) keep_with_next: Option<bool>,
    pub(super) start_on_new_page: Option<bool>,
    pub(super) prevent_widow_orphan_lines: Option<bool>,
    pub(super) writing_direction: Option<WritingDirection>,
    pub(super) following_style: Option<ParagraphFollowingStyle>,
    pub(super) underline: Option<TextUnderline>,
    pub(super) strikethrough: Option<TextStrikethrough>,
    pub(super) alignment: Option<Alignment>,
    pub(super) line_spacing: Option<LineSpacing>,
    pub(super) space_before: Option<SpacingPoints>,
    pub(super) space_after: Option<SpacingPoints>,
    pub(super) first_line_indent: Option<IndentPoints>,
    pub(super) left_indent: Option<IndentPoints>,
    pub(super) right_indent: Option<IndentPoints>,
    pub(super) decimal_tab_character: Option<DecimalCharacter>,
    pub(super) default_tab_interval: Option<DefaultInterval>,
    pub(super) tab_stops: Option<Stops>,
}

impl Overrides {
    pub(super) fn count(&self) -> u32 {
        u32::from(self.bold.is_some())
            + u32::from(self.italic.is_some())
            + u32::from(self.point_size.is_some())
            + u32::from(self.font.is_some())
            + u32::from(self.font_color.is_some())
            + self
                .capitalization
                .map_or(0, values::capitalization_override_count)
            + u32::from(self.script.is_some())
            + u32::from(self.baseline_shift.is_some())
            + u32::from(self.character_spacing.is_some())
            + u32::from(self.ligatures.is_some())
            + u32::from(self.outline.is_some())
            + u32::from(self.shadow.is_some())
            + u32::from(self.background.is_some())
            + u32::from(self.paragraph_background.is_some())
            + self.paragraph_borders.map_or(0, |_| 4)
            + u32::from(self.hyphenation.is_some())
            + u32::from(self.keep_lines_together.is_some())
            + u32::from(self.keep_with_next.is_some())
            + u32::from(self.start_on_new_page.is_some())
            + u32::from(self.prevent_widow_orphan_lines.is_some())
            + u32::from(self.writing_direction.is_some())
            + u32::from(self.following_style.is_some())
            + u32::from(self.underline.is_some())
            + u32::from(self.strikethrough.is_some())
            + u32::from(self.alignment.is_some())
            + u32::from(self.line_spacing.is_some())
            + u32::from(self.space_before.is_some())
            + u32::from(self.space_after.is_some())
            + u32::from(self.first_line_indent.is_some())
            + u32::from(self.left_indent.is_some())
            + u32::from(self.right_indent.is_some())
            + u32::from(self.decimal_tab_character.is_some())
            + u32::from(self.default_tab_interval.is_some())
            + u32::from(self.tab_stops.is_some())
    }

    pub(super) fn is_empty(&self) -> bool {
        self.count() == 0
    }

    pub(super) fn is_tab_defaults_only(&self) -> bool {
        !self.is_empty()
            && self.count()
                == u32::from(self.decimal_tab_character.is_some())
                    + u32::from(self.default_tab_interval.is_some())
    }
}

fn invalid<T>(reason: &'static str) -> NativeResult<T> {
    Err(NativeError::Invalid(reason))
}

fn transpose<T, E>(value: Option<Result<T, E>>, reason: &'static str) -> NativeResult<Option<T>> {
    value.transpose().map_err(|_| NativeError::Invalid(reason))
}

/// The overrides of `style` when its payload `raw` is a managed variation:
/// a variation whose fields are exactly the canonical fields of its
/// overrides. Anything else (a named style, foreign or unknown fields, a
/// non-canonical encoding) is `None`, and the engine derives a child
/// variation instead of rewriting it.
#[allow(
    clippy::too_many_lines,
    reason = "a single field-by-field port of the host recogniser"
)]
pub(super) fn direct_overrides(
    style: &native::ParagraphStyle,
    raw: &[u8],
    budget: &mut Budget,
) -> NativeResult<Option<Overrides>> {
    let Some(character) = style.character.as_ref() else {
        return Ok(None);
    };
    let Some(properties) = style.paragraph.as_ref() else {
        return Ok(None);
    };
    let bold = character.bold;
    let italic = character.italic;
    let point_size = transpose(
        character.font_size.map(TextPointSize::from_points),
        "native iWork point size",
    )?;
    let font = convert::text_font(character)?;
    let font_field = font.as_ref().map(|font| match font {
        Font::Default => CHARACTER_FONT_NAME_NULL_FIELD,
        Font::Named(_) => CHARACTER_FONT_NAME_FIELD,
    });
    let font_color = convert::text_color(character)?;
    let capitalization = convert::capitalization(character)?;
    let script = character
        .superscript
        .map(values::script_from_native)
        .transpose()?;
    let baseline_shift = transpose(
        character.baseline_shift.map(TextBaselineShift::from_points),
        "native iWork baseline shift",
    )?;
    let character_spacing = character
        .tracking
        .map(values::character_spacing_from_native)
        .transpose()?;
    let ligatures = character
        .ligatures
        .map(values::ligatures_from_native)
        .transpose()?;
    let outline = convert::text_outline(character)?;
    let shadow = convert::text_shadow(character)?;
    let background = convert::text_background(character)?;
    let paragraph_background = convert::paragraph_background(properties)?;
    let paragraph_borders = convert::paragraph_borders(properties)?;
    let hyphenation = properties.hyphenate.map(values::hyphenation_from_native);
    let keep_lines_together = properties.keep_lines_together;
    let keep_with_next = properties.keep_with_next;
    let start_on_new_page = properties.page_break_before;
    let prevent_widow_orphan_lines = properties.widow_control;
    let writing_direction = character
        .writing_direction
        .map(values::writing_direction_from_native)
        .transpose()?;
    let following_style = if properties.following_style_null == Some(true) {
        if properties.following_style.is_some() {
            return invalid("native iWork following paragraph style is both null and populated");
        }
        Some(ParagraphFollowingStyle::Same)
    } else {
        properties
            .following_style
            .map(|reference| {
                raw::from_native_id(reference.identifier)
                    .map(ParagraphFollowingStyle::Named)
                    .map_err(|_| NativeError::Invalid("native iWork following paragraph style"))
            })
            .transpose()?
    };
    let underline = character
        .underline
        .map(values::underline_from_native)
        .transpose()?;
    let strikethrough = character
        .strikethru
        .map(values::strikethrough_from_native)
        .transpose()?;
    let alignment = properties
        .alignment
        .map(values::alignment_from_native)
        .transpose()?;
    let line_spacing = properties
        .line_spacing
        .as_ref()
        .map(values::line_spacing_from_native)
        .transpose()?;
    let space_before = transpose(
        properties.space_before.map(SpacingPoints::from_points),
        "native iWork paragraph spacing",
    )?;
    let space_after = transpose(
        properties.space_after.map(SpacingPoints::from_points),
        "native iWork paragraph spacing",
    )?;
    let first_line_indent = transpose(
        properties.first_line_indent.map(IndentPoints::from_points),
        "native iWork paragraph indent",
    )?;
    let left_indent = transpose(
        properties.left_indent.map(IndentPoints::from_points),
        "native iWork paragraph indent",
    )?;
    let right_indent = transpose(
        properties.right_indent.map(IndentPoints::from_points),
        "native iWork paragraph indent",
    )?;
    let decimal_tab_character = if properties.decimal_tab_null == Some(true) {
        if properties.decimal_tab.is_some() {
            return invalid("native iWork decimal tab is both null and populated");
        }
        Some(DecimalCharacter::default())
    } else {
        properties
            .decimal_tab
            .as_deref()
            .map(values::decimal_character_from_native)
            .transpose()?
    };
    let default_tab_interval = transpose(
        properties
            .default_tab_stops
            .map(DefaultInterval::from_points),
        "native iWork default tab interval",
    )?;
    let tab_stops = properties
        .tabs
        .as_ref()
        .map(values::tabs_from_native)
        .transpose()?;
    let overrides = Overrides {
        bold,
        italic,
        point_size,
        font,
        font_color,
        capitalization,
        script,
        baseline_shift,
        character_spacing,
        ligatures,
        outline,
        shadow,
        background,
        paragraph_background,
        paragraph_borders,
        hyphenation,
        keep_lines_together,
        keep_with_next,
        start_on_new_page,
        prevent_widow_orphan_lines,
        writing_direction,
        following_style,
        underline,
        strikethrough,
        alignment,
        line_spacing,
        space_before,
        space_after,
        first_line_indent,
        left_indent,
        right_indent,
        decimal_tab_character,
        default_tab_interval,
        tab_stops,
    };
    let semantic = !overrides.is_empty()
        && style.override_count == Some(overrides.count())
        && style.base.name.is_none()
        && style.base.style_identifier.is_none()
        && style.base.parent.is_some()
        && style.base.is_variation == Some(true)
        && style.base.stylesheet.is_some();
    if !semantic {
        return Ok(None);
    }

    let super_raw = required_payload(raw, STYLE_SUPER_FIELD, budget)?;
    let character_raw = required_payload(raw, STYLE_CHARACTER_PROPERTIES_FIELD, budget)?;
    let paragraph_raw = required_payload(raw, STYLE_PARAGRAPH_PROPERTIES_FIELD, budget)?;
    let mut character_fields = Vec::with_capacity(16);
    if bold.is_some() {
        character_fields.push(CHARACTER_BOLD_FIELD);
    }
    if italic.is_some() {
        character_fields.push(CHARACTER_ITALIC_FIELD);
    }
    if point_size.is_some() {
        character_fields.push(CHARACTER_FONT_SIZE_FIELD);
    }
    if let Some(field) = font_field {
        character_fields.push(field);
    }
    if font_color.is_some() {
        character_fields.push(CHARACTER_FONT_COLOR_FIELD);
        character_fields.push(CHARACTER_DRAWING_FILL_FIELD);
        let legacy = required_payload(character_raw, CHARACTER_FONT_COLOR_FIELD, budget)?;
        let fill = required_payload(character_raw, CHARACTER_DRAWING_FILL_FIELD, budget)?;
        let drawing = required_payload(fill, DRAWING_FILL_COLOR_FIELD, budget)?;
        if !has_exact_fields(legacy, &COLOR_FIELDS, budget)?
            || !has_exact_fields(fill, &[DRAWING_FILL_COLOR_FIELD], budget)?
            || !has_exact_fields(drawing, &COLOR_FIELDS, budget)?
        {
            return Ok(None);
        }
    }
    if let Some(capitalization) = capitalization {
        character_fields.push(CHARACTER_CAPITALIZATION_FIELD);
        if values::capitalization_uses_linguistics(capitalization).is_some() {
            character_fields.push(CHARACTER_CAPITALIZATION_LINGUISTICS_FIELD);
        }
    }
    if script.is_some() {
        character_fields.push(CHARACTER_SCRIPT_FIELD);
    }
    if baseline_shift.is_some() {
        character_fields.push(CHARACTER_BASELINE_SHIFT_FIELD);
    }
    if character_spacing.is_some() {
        character_fields.push(CHARACTER_TRACKING_FIELD);
    }
    if ligatures.is_some() {
        character_fields.push(CHARACTER_LIGATURES_FIELD);
    }
    if let Some(outline) = outline {
        character_fields.push(match outline {
            Outline::None => CHARACTER_DRAWING_STROKE_NULL_FIELD,
            Outline::Stroke(_) => CHARACTER_DRAWING_STROKE_FIELD,
        });
    }
    if let Some(shadow) = shadow {
        character_fields.push(match shadow {
            Shadow::None => CHARACTER_SHADOW_NULL_FIELD,
            Shadow::Drop(_) => CHARACTER_SHADOW_FIELD,
        });
    }
    if let Some(background) = background {
        let field = match background {
            Background::None => CHARACTER_BACKGROUND_COLOR_NULL_FIELD,
            Background::Color(_) => CHARACTER_BACKGROUND_COLOR_FIELD,
        };
        character_fields.push(field);
        if matches!(background, Background::Color(_)) {
            let color = required_payload(character_raw, field, budget)?;
            if !has_exact_fields(color, &COLOR_FIELDS, budget)? {
                return Ok(None);
            }
        }
    }
    if underline.is_some() {
        character_fields.push(CHARACTER_UNDERLINE_FIELD);
    }
    if strikethrough.is_some() {
        character_fields.push(CHARACTER_STRIKETHROUGH_FIELD);
    }
    let mut paragraph_fields = Vec::with_capacity(24);
    if alignment.is_some() {
        paragraph_fields.push(PARAGRAPH_ALIGNMENT_FIELD);
    }
    if writing_direction.is_some() {
        character_fields.push(CHARACTER_WRITING_DIRECTION_FIELD);
    }
    if let Some(background) = paragraph_background {
        let field = match background {
            ParagraphBackground::None => PARAGRAPH_FILL_NULL_FIELD,
            ParagraphBackground::Color(_) => PARAGRAPH_FILL_FIELD,
        };
        paragraph_fields.push(field);
        if matches!(background, ParagraphBackground::Color(_)) {
            let color = required_payload(paragraph_raw, field, budget)?;
            if !has_exact_fields(color, &COLOR_FIELDS, budget)? {
                return Ok(None);
            }
        }
    }
    if let Some(borders) = paragraph_borders {
        paragraph_fields.push(PARAGRAPH_DEPRECATED_BORDERS_FIELD);
        if properties.historical_rule_offset_null.is_some() {
            paragraph_fields.push(PARAGRAPH_BORDER_OFFSET_NULL_FIELD);
        }
        if properties.historical_rule_offset.is_some() {
            paragraph_fields.push(PARAGRAPH_BORDER_OFFSET_FIELD);
        }
        if properties.rule_width.is_some() {
            paragraph_fields.push(PARAGRAPH_LEGACY_RULE_WIDTH_FIELD);
        }
        paragraph_fields.push(match borders {
            Borders::None => PARAGRAPH_BORDER_STROKE_NULL_FIELD,
            Borders::Bordered(_) => PARAGRAPH_BORDER_STROKE_FIELD,
        });
        paragraph_fields.push(PARAGRAPH_BORDER_POSITIONS_FIELD);
        paragraph_fields.push(PARAGRAPH_BORDER_ROUNDED_CORNERS_FIELD);
    }
    for (field, present) in [
        (PARAGRAPH_HYPHENATE_FIELD, hyphenation.is_some()),
        (
            PARAGRAPH_KEEP_LINES_TOGETHER_FIELD,
            keep_lines_together.is_some(),
        ),
        (PARAGRAPH_KEEP_WITH_NEXT_FIELD, keep_with_next.is_some()),
        (
            PARAGRAPH_PAGE_BREAK_BEFORE_FIELD,
            start_on_new_page.is_some(),
        ),
        (
            PARAGRAPH_WIDOW_CONTROL_FIELD,
            prevent_widow_orphan_lines.is_some(),
        ),
    ] {
        if present {
            paragraph_fields.push(field);
            if !has_canonical_bool_field(paragraph_raw, field, budget)? {
                return Ok(None);
            }
        }
    }
    if let Some(spacing) = line_spacing {
        paragraph_fields.push(PARAGRAPH_LINE_SPACING_FIELD);
        let spacing_raw = required_payload(paragraph_raw, PARAGRAPH_LINE_SPACING_FIELD, budget)?;
        let expected: &[u32] = match spacing {
            LineSpacing::Relative(_) => &[LINE_SPACING_AMOUNT_FIELD],
            _ => &[LINE_SPACING_MODE_FIELD, LINE_SPACING_AMOUNT_FIELD],
        };
        if !has_exact_fields(spacing_raw, expected, budget)? {
            return Ok(None);
        }
    }
    if space_after.is_some() {
        paragraph_fields.push(PARAGRAPH_SPACE_AFTER_FIELD);
    }
    if space_before.is_some() {
        paragraph_fields.push(PARAGRAPH_SPACE_BEFORE_FIELD);
    }
    if first_line_indent.is_some() {
        paragraph_fields.push(PARAGRAPH_FIRST_LINE_INDENT_FIELD);
    }
    if left_indent.is_some() {
        paragraph_fields.push(PARAGRAPH_LEFT_INDENT_FIELD);
    }
    if right_indent.is_some() {
        paragraph_fields.push(PARAGRAPH_RIGHT_INDENT_FIELD);
    }
    if let Some(following_style) = following_style {
        match following_style {
            ParagraphFollowingStyle::Same => {
                paragraph_fields.push(PARAGRAPH_FOLLOWING_STYLE_NULL_FIELD);
                if !has_canonical_bool_field(
                    paragraph_raw,
                    PARAGRAPH_FOLLOWING_STYLE_NULL_FIELD,
                    budget,
                )? {
                    return Ok(None);
                }
            },
            ParagraphFollowingStyle::Named(_) => {
                paragraph_fields.push(PARAGRAPH_FOLLOWING_STYLE_FIELD);
                let reference =
                    required_payload(paragraph_raw, PARAGRAPH_FOLLOWING_STYLE_FIELD, budget)?;
                if !has_exact_fields(reference, &[1], budget)? {
                    return Ok(None);
                }
            },
        }
    }
    if decimal_tab_character.is_some() {
        if properties.decimal_tab_null == Some(true) {
            paragraph_fields.push(PARAGRAPH_DECIMAL_TAB_NULL_FIELD);
            if !has_canonical_bool_field(paragraph_raw, PARAGRAPH_DECIMAL_TAB_NULL_FIELD, budget)? {
                return Ok(None);
            }
        } else {
            paragraph_fields.push(PARAGRAPH_DECIMAL_TAB_FIELD);
        }
    }
    if default_tab_interval.is_some() {
        paragraph_fields.push(PARAGRAPH_DEFAULT_TAB_INTERVAL_FIELD);
    }
    if let Some(stops) = overrides.tab_stops.as_ref() {
        paragraph_fields.push(PARAGRAPH_TABS_FIELD);
        let tabs_raw = required_payload(paragraph_raw, PARAGRAPH_TABS_FIELD, budget)?;
        if !has_canonical_tabs(tabs_raw, stops, budget)? {
            return Ok(None);
        }
    }
    let exact = has_exact_fields(
        raw,
        &[
            STYLE_SUPER_FIELD,
            STYLE_OVERRIDE_COUNT_FIELD,
            STYLE_CHARACTER_PROPERTIES_FIELD,
            STYLE_PARAGRAPH_PROPERTIES_FIELD,
        ],
        budget,
    )? && has_exact_fields(
        super_raw,
        &[
            STYLE_PARENT_FIELD,
            STYLE_VARIATION_FIELD,
            STYLE_STYLESHEET_FIELD,
        ],
        budget,
    )? && has_exact_fields(character_raw, &character_fields, budget)?
        && has_exact_fields(paragraph_raw, &paragraph_fields, budget)?;
    Ok(exact.then_some(overrides))
}

fn has_canonical_tabs(data: &[u8], stops: &Stops, budget: &mut Budget) -> NativeResult<bool> {
    let mut payloads = Vec::new();
    let mut other = false;
    native::fields(data, 0, budget, |field, _| {
        if field.number == TABS_TAB_FIELD {
            native::reserve(&mut payloads, 1)?;
            payloads.push(field.bytes()?);
        } else {
            other = true;
        }
        Ok(())
    })?;
    if other || payloads.len() != stops.as_slice().len() {
        return Ok(false);
    }
    for (payload, stop) in payloads.into_iter().zip(stops.as_slice()) {
        let mut expected = vec![TAB_POSITION_FIELD];
        if stop.alignment != TabAlignment::Left {
            expected.push(TAB_ALIGNMENT_FIELD);
        }
        if stop.leader.is_some() {
            expected.push(TAB_LEADER_FIELD);
        }
        if !has_exact_fields(payload, &expected, budget)? {
            return Ok(false);
        }
    }
    Ok(true)
}

/// The canonical payload of a variation of `parent` holding `overrides`,
/// its object references and its managed field paths.
pub(super) fn variation(
    parent: u64,
    stylesheet: u64,
    overrides: &Overrides,
    limits: WireLimits,
) -> NativeResult<(Vec<u8>, Vec<u64>, ManagedFields)> {
    if overrides.is_empty() {
        return invalid("an iWork paragraph-style variation must contain an override");
    }
    let (font_name_null, font_name) = match &overrides.font {
        Some(Font::Default) => (Some(true), None),
        Some(Font::Named(name)) => (None, Some(name.as_str().to_owned())),
        None => (None, None),
    };
    let borders = convert::borders_to_native(overrides.paragraph_borders);
    let style = native::ParagraphStyle {
        base: native::StyleBase {
            parent: Some(native::Reference { identifier: parent }),
            is_variation: Some(true),
            stylesheet: Some(native::Reference {
                identifier: stylesheet,
            }),
            ..native::StyleBase::default()
        },
        override_count: Some(overrides.count()),
        character: Some(native::CharacterProperties {
            bold: overrides.bold,
            italic: overrides.italic,
            font_size: overrides.point_size.map(TextPointSize::points),
            font_name_null,
            font_name,
            font_color: overrides.font_color.map(convert::native_color),
            capitalization: overrides
                .capitalization
                .map(values::capitalization_to_native),
            superscript: overrides.script.map(values::script_to_native),
            baseline_shift: overrides.baseline_shift.map(TextBaselineShift::points),
            tracking: overrides
                .character_spacing
                .map(values::character_spacing_to_native),
            ligatures: overrides.ligatures.map(values::ligatures_to_native),
            tsd_stroke_null: matches!(overrides.outline, Some(Outline::None)).then_some(true),
            tsd_stroke: overrides.outline.and_then(|outline| match outline {
                Outline::None => None,
                Outline::Stroke(stroke) => Some(values::stroke_to_native(stroke)),
            }),
            shadow_null: matches!(overrides.shadow, Some(Shadow::None)).then_some(true),
            shadow: overrides.shadow.and_then(|shadow| match shadow {
                Shadow::None => None,
                Shadow::Drop(_) => Some(values::shadow_to_native(shadow.into_shape_shadow())),
            }),
            background_color_null: matches!(overrides.background, Some(Background::None))
                .then_some(true),
            background_color: overrides
                .background
                .and_then(|background| match background {
                    Background::None => None,
                    Background::Color(color) => Some(convert::native_color(color)),
                }),
            underline: overrides.underline.map(values::underline_to_native),
            strikethru: overrides.strikethrough.map(values::strikethrough_to_native),
            tsd_fill: overrides.font_color.map(|color| native::Fill {
                color: Some(convert::native_color(color)),
                ..native::Fill::default()
            }),
            capitalization_uses_linguistics: overrides
                .capitalization
                .and_then(values::capitalization_uses_linguistics),
            writing_direction: overrides
                .writing_direction
                .map(values::writing_direction_to_native),
            ..native::CharacterProperties::default()
        }),
        paragraph: Some(native::ParagraphProperties {
            alignment: overrides.alignment.map(values::alignment_to_native),
            fill_null: matches!(
                overrides.paragraph_background,
                Some(ParagraphBackground::None)
            )
            .then_some(true),
            fill: overrides
                .paragraph_background
                .and_then(|background| match background {
                    ParagraphBackground::None => None,
                    ParagraphBackground::Color(color) => Some(convert::native_color(color)),
                }),
            deprecated_borders: borders.deprecated_borders,
            historical_rule_offset: borders.historical_rule_offset,
            stroke_null: borders.stroke_null,
            stroke: borders.stroke,
            border_positions: borders.positions,
            rounded_corners: borders.rounded_corners,
            hyphenate: overrides.hyphenation.map(values::hyphenation_to_native),
            keep_lines_together: overrides.keep_lines_together,
            keep_with_next: overrides.keep_with_next,
            page_break_before: overrides.start_on_new_page,
            widow_control: overrides.prevent_widow_orphan_lines,
            following_style_null: matches!(
                overrides.following_style,
                Some(ParagraphFollowingStyle::Same)
            )
            .then_some(true),
            following_style: overrides
                .following_style
                .and_then(|following| match following {
                    ParagraphFollowingStyle::Same => None,
                    ParagraphFollowingStyle::Named(identifier) => Some(native::Reference {
                        identifier: raw::native_id(identifier),
                    }),
                }),
            line_spacing: overrides.line_spacing.map(values::line_spacing_to_native),
            space_before: overrides.space_before.map(SpacingPoints::points),
            space_after: overrides.space_after.map(SpacingPoints::points),
            first_line_indent: overrides.first_line_indent.map(IndentPoints::points),
            left_indent: overrides.left_indent.map(IndentPoints::points),
            right_indent: overrides.right_indent.map(IndentPoints::points),
            decimal_tab: overrides
                .decimal_tab_character
                .map(|character| character.character().to_string()),
            default_tab_stops: overrides.default_tab_interval.map(DefaultInterval::points),
            tabs: overrides.tab_stops.as_ref().map(values::tabs_to_native),
            ..native::ParagraphProperties::default()
        }),
    };
    let payload = style.encode(limits)?;
    let mut references = vec![parent];
    if let Some(ParagraphFollowingStyle::Named(identifier)) = overrides.following_style {
        references.push(raw::native_id(identifier));
    }
    let managed = ManagedFields {
        text_fill: overrides.font_color.is_some(),
        text_stroke: matches!(overrides.outline, Some(Outline::Stroke(_))),
        text_shadow: matches!(overrides.shadow, Some(Shadow::Drop(_))),
        text_background: matches!(overrides.background, Some(Background::Color(_))),
    };
    Ok((payload, references, managed))
}
