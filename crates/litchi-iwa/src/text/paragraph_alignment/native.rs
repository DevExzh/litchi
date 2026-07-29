//! Native paragraph-style inheritance, minimal variations, and ownership checks.

mod inheritance;
mod tabs;

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::{tsd, tsp, tss, tswp};
use crate::shapes::{
    RgbaColor, StrokeCap, StrokeJoin, color_from_native, color_to_native, shadow_from_native,
    shadow_to_native, stroke_from_native, stroke_to_native,
};
use crate::wire::{parse_wire_fields, repeated_length_delimited_payloads};
use crate::{Error, IWorkPackage, Result};

use super::super::font::{TextFont, TextFontName};
use super::super::paragraph_direction::ParagraphWritingDirection;
use super::super::paragraph_flow::{ParagraphFlow, ParagraphHyphenation};
use super::super::paragraph_tabs::ParagraphTabStops;
use super::super::style::{
    ParagraphBackground, ParagraphBorder, ParagraphBorderOffset, ParagraphBorderSides,
    ParagraphBorders, ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing,
    ParagraphLineSpacingMultiple, ParagraphLineSpacingPoints, ParagraphSpacing,
    ParagraphSpacingPoints, TextAlignment, TextBackground, TextBaselineShift, TextCapitalization,
    TextCharacterSpacing, TextDecorations, TextLigatures, TextOutline, TextPointSize, TextScript,
    TextShadow, TextStrikethrough, TextStyle, TextUnderline,
};
use super::super::style_registry::object_archive_name;

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

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
const CHARACTER_DRAWING_STROKE_NULL_FIELD: u32 = 43;
const CHARACTER_DRAWING_STROKE_FIELD: u32 = 44;
const CHARACTER_DRAWING_FILL_FIELD: u32 = 46;
const CHARACTER_CAPITALIZATION_LINGUISTICS_FIELD: u32 = 41;
const CHARACTER_WRITING_DIRECTION_FIELD: u32 = 35;
const DRAWING_FILL_COLOR_FIELD: u32 = 1;
const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_RGB_SPACE_FIELD: u32 = 12;
const PARAGRAPH_ALIGNMENT_FIELD: u32 = 1;
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
const PARAGRAPH_BORDER_POSITIONS_FIELD: u32 = 45;
const PARAGRAPH_BORDER_ROUNDED_CORNERS_FIELD: u32 = 46;
const LINE_SPACING_MODE_FIELD: u32 = 1;
const LINE_SPACING_AMOUNT_FIELD: u32 = 2;
const LEGACY_BORDER_TOP: i32 = 1;
const LEGACY_BORDER_BOTTOM: i32 = 2;
const LEGACY_BORDER_ALL: i32 = 4;
const LEGACY_BORDER_LEFT: i32 = 8;
const LEGACY_BORDER_RIGHT: i32 = 16;
const PROTOBUF_FALSE_BYTE: u8 = 0;
const PROTOBUF_TRUE_BYTE: u8 = 1;

const RELATIVE_LINE_SPACING_MODE: i32 = 0;
const MINIMUM_LINE_SPACING_MODE: i32 = 1;
const EXACT_LINE_SPACING_MODE: i32 = 2;
const MAXIMUM_LINE_SPACING_MODE: i32 = 3;
const BETWEEN_LINE_SPACING_MODE: i32 = 4;

pub(crate) struct ParagraphStyleLocation {
    pub(crate) archive_name: String,
    pub(crate) message: RawMessage,
    pub(crate) style: tswp::ParagraphStyleArchive,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub(crate) struct ParagraphStyleOverrides {
    pub(crate) bold: Option<bool>,
    pub(crate) italic: Option<bool>,
    pub(crate) point_size: Option<TextPointSize>,
    pub(crate) font: Option<TextFont>,
    pub(crate) font_color: Option<RgbaColor>,
    pub(crate) capitalization: Option<TextCapitalization>,
    pub(crate) script: Option<TextScript>,
    pub(crate) baseline_shift: Option<TextBaselineShift>,
    pub(crate) character_spacing: Option<TextCharacterSpacing>,
    pub(crate) ligatures: Option<TextLigatures>,
    pub(crate) outline: Option<TextOutline>,
    pub(crate) shadow: Option<TextShadow>,
    pub(crate) background: Option<TextBackground>,
    pub(crate) paragraph_background: Option<ParagraphBackground>,
    pub(crate) paragraph_borders: Option<ParagraphBorders>,
    pub(crate) hyphenation: Option<ParagraphHyphenation>,
    pub(crate) keep_lines_together: Option<bool>,
    pub(crate) keep_with_next: Option<bool>,
    pub(crate) start_on_new_page: Option<bool>,
    pub(crate) prevent_widow_orphan_lines: Option<bool>,
    pub(crate) writing_direction: Option<ParagraphWritingDirection>,
    pub(crate) underline: Option<TextUnderline>,
    pub(crate) strikethrough: Option<TextStrikethrough>,
    pub(crate) alignment: Option<TextAlignment>,
    pub(crate) line_spacing: Option<ParagraphLineSpacing>,
    pub(crate) space_before: Option<ParagraphSpacingPoints>,
    pub(crate) space_after: Option<ParagraphSpacingPoints>,
    pub(crate) first_line_indent: Option<ParagraphIndentPoints>,
    pub(crate) left_indent: Option<ParagraphIndentPoints>,
    pub(crate) right_indent: Option<ParagraphIndentPoints>,
    pub(crate) tab_stops: Option<ParagraphTabStops>,
}

#[derive(Default)]
struct NativeParagraphBorderFields {
    deprecated_borders: Option<i32>,
    historical_rule_offset: Option<tsp::Point>,
    stroke_null: Option<bool>,
    stroke: Option<tsd::StrokeArchive>,
    positions: Option<i32>,
    rounded_corners: Option<bool>,
}

impl ParagraphStyleOverrides {
    pub(super) fn count(&self) -> u32 {
        u32::from(self.bold.is_some())
            + u32::from(self.italic.is_some())
            + u32::from(self.point_size.is_some())
            + u32::from(self.font.is_some())
            + u32::from(self.font_color.is_some())
            + self
                .capitalization
                .map_or(0, TextCapitalization::native_override_count)
            + u32::from(self.script.is_some())
            + u32::from(self.baseline_shift.is_some())
            + u32::from(self.character_spacing.is_some())
            + u32::from(self.ligatures.is_some())
            + u32::from(self.outline.is_some())
            + u32::from(self.shadow.is_some())
            + u32::from(self.background.is_some())
            + u32::from(self.paragraph_background.is_some())
            + self
                .paragraph_borders
                .map_or(0, ParagraphBorders::native_override_count)
            + u32::from(self.hyphenation.is_some())
            + u32::from(self.keep_lines_together.is_some())
            + u32::from(self.keep_with_next.is_some())
            + u32::from(self.start_on_new_page.is_some())
            + u32::from(self.prevent_widow_orphan_lines.is_some())
            + u32::from(self.writing_direction.is_some())
            + u32::from(self.underline.is_some())
            + u32::from(self.strikethrough.is_some())
            + u32::from(self.alignment.is_some())
            + u32::from(self.line_spacing.is_some())
            + u32::from(self.space_before.is_some())
            + u32::from(self.space_after.is_some())
            + u32::from(self.first_line_indent.is_some())
            + u32::from(self.left_indent.is_some())
            + u32::from(self.right_indent.is_some())
            + u32::from(self.tab_stops.is_some())
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.count() == 0
    }

    pub(crate) fn is_chart_font_format_only(&self) -> bool {
        (self.font.is_some()
            || self.bold.is_some()
            || self.italic.is_some()
            || self.point_size.is_some())
            && self.font_color.is_none()
            && self.capitalization.is_none()
            && self.script.is_none()
            && self.baseline_shift.is_none()
            && self.character_spacing.is_none()
            && self.ligatures.is_none()
            && self.outline.is_none()
            && self.shadow.is_none()
            && self.background.is_none()
            && self.paragraph_background.is_none()
            && self.paragraph_borders.is_none()
            && self.hyphenation.is_none()
            && self.keep_lines_together.is_none()
            && self.keep_with_next.is_none()
            && self.start_on_new_page.is_none()
            && self.prevent_widow_orphan_lines.is_none()
            && self.writing_direction.is_none()
            && self.underline.is_none()
            && self.strikethrough.is_none()
            && self.alignment.is_none()
            && self.line_spacing.is_none()
            && self.space_before.is_none()
            && self.space_after.is_none()
            && self.first_line_indent.is_none()
            && self.left_indent.is_none()
            && self.right_indent.is_none()
            && self.tab_stops.is_none()
    }
}

pub(crate) fn locate_style(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<ParagraphStyleLocation> {
    let archive_name = object_archive_name(package, style_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(style_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork paragraph style {style_id} is missing"))
    })?;
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
        .filter_map(|message| {
            tswp::ParagraphStyleArchive::decode(message.data.as_slice())
                .ok()
                .map(|style| (message.clone(), style))
        })
        .collect::<Vec<_>>();
    let [(message, style)] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {style_id} must have exactly one paragraph-style payload"
        )));
    };
    Ok(ParagraphStyleLocation {
        archive_name,
        message: message.clone(),
        style: style.clone(),
    })
}

pub(crate) fn inherited_alignment(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextAlignment> {
    inheritance::alignment(package, first_style_id)
}

pub(crate) fn inherited_text_style(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextStyle> {
    inheritance::text_style(package, first_style_id)
}

pub(crate) fn inherited_text_font(package: &IWorkPackage, first_style_id: u64) -> Result<TextFont> {
    inheritance::text_font(package, first_style_id)
}

pub(crate) fn inherited_text_decorations(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextDecorations> {
    inheritance::text_decorations(package, first_style_id)
}

pub(crate) fn inherited_text_color(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<RgbaColor> {
    inheritance::text_color(package, first_style_id)
}

pub(crate) fn inherited_text_capitalization(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextCapitalization> {
    inheritance::text_capitalization(package, first_style_id)
}

pub(crate) fn inherited_text_script(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextScript> {
    inheritance::text_script(package, first_style_id)
}

pub(crate) fn inherited_text_baseline_shift(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextBaselineShift> {
    inheritance::text_baseline_shift(package, first_style_id)
}

pub(crate) fn inherited_text_character_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextCharacterSpacing> {
    inheritance::text_character_spacing(package, first_style_id)
}

pub(crate) fn inherited_text_ligatures(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextLigatures> {
    inheritance::text_ligatures(package, first_style_id)
}

pub(crate) fn inherited_text_outline(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextOutline> {
    inheritance::text_outline(package, first_style_id)
}

pub(crate) fn inherited_text_shadow(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextShadow> {
    inheritance::text_shadow(package, first_style_id)
}

pub(crate) fn inherited_text_background(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextBackground> {
    inheritance::text_background(package, first_style_id)
}

pub(crate) fn inherited_paragraph_background(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphBackground> {
    inheritance::paragraph_background(package, first_style_id)
}

pub(crate) fn inherited_paragraph_borders(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphBorders> {
    inheritance::paragraph_borders(package, first_style_id)
}

pub(crate) fn inherited_paragraph_flow(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphFlow> {
    inheritance::paragraph_flow(package, first_style_id)
}

pub(crate) fn inherited_paragraph_writing_direction(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphWritingDirection> {
    inheritance::paragraph_writing_direction(package, first_style_id)
}

pub(crate) fn inherited_line_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphLineSpacing> {
    inheritance::line_spacing(package, first_style_id)
}

pub(crate) fn inherited_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphSpacing> {
    inheritance::spacing(package, first_style_id)
}

pub(crate) fn inherited_indents(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphIndents> {
    inheritance::indents(package, first_style_id)
}

pub(crate) fn inherited_tab_stops(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphTabStops> {
    inheritance::tab_stops(package, first_style_id)
}

pub(crate) fn direct_overrides(
    style: &tswp::ParagraphStyleArchive,
    raw: &[u8],
) -> Result<Option<ParagraphStyleOverrides>> {
    let Some(character_properties) = style.char_properties.as_ref() else {
        return Ok(None);
    };
    let Some(properties) = style.para_properties.as_ref() else {
        return Ok(None);
    };
    let bold = character_properties.bold;
    let italic = character_properties.italic;
    let point_size = character_properties
        .font_size
        .map(TextPointSize::from_points)
        .transpose()?;
    let font = text_font_from_character(character_properties)?;
    let font_is_some = font.is_some();
    let font_field = font.as_ref().map(|font| match font {
        TextFont::Default => CHARACTER_FONT_NAME_NULL_FIELD,
        TextFont::Named(_) => CHARACTER_FONT_NAME_FIELD,
    });
    let font_color = text_color_from_character(character_properties)?;
    let capitalization = capitalization_from_character(character_properties)?;
    let script = character_properties
        .superscript
        .map(TextScript::from_native_value)
        .transpose()?;
    let baseline_shift = character_properties
        .baseline_shift
        .map(TextBaselineShift::from_points)
        .transpose()?;
    let character_spacing = character_properties
        .tracking
        .map(TextCharacterSpacing::from_native_ratio)
        .transpose()?;
    let ligatures = character_properties
        .ligatures
        .map(TextLigatures::from_native_value)
        .transpose()?;
    let outline = text_outline_from_character(character_properties)?;
    let shadow = text_shadow_from_character(character_properties)?;
    let background = text_background_from_character(character_properties)?;
    let paragraph_background = paragraph_background_from_properties(properties)?;
    let paragraph_borders = paragraph_borders_from_properties(properties)?;
    let hyphenation = properties
        .hyphenate
        .map(ParagraphHyphenation::from_native_value);
    let keep_lines_together = properties.keep_lines_together;
    let keep_with_next = properties.keep_with_next;
    let start_on_new_page = properties.page_break_before;
    let prevent_widow_orphan_lines = properties.widow_control;
    let writing_direction = character_properties
        .writing_direction
        .map(ParagraphWritingDirection::from_native_value)
        .transpose()?;
    let underline = character_properties
        .underline
        .map(TextUnderline::from_native_value)
        .transpose()?;
    let strikethrough = character_properties
        .strikethru
        .map(TextStrikethrough::from_native_value)
        .transpose()?;
    let alignment = properties
        .alignment
        .map(TextAlignment::from_native_value)
        .transpose()?;
    let line_spacing = properties
        .line_spacing
        .as_ref()
        .map(line_spacing_from_archive)
        .transpose()?;
    let space_before = properties
        .space_before
        .map(ParagraphSpacingPoints::from_points)
        .transpose()?;
    let space_after = properties
        .space_after
        .map(ParagraphSpacingPoints::from_points)
        .transpose()?;
    let first_line_indent = properties
        .first_line_indent
        .map(ParagraphIndentPoints::from_points)
        .transpose()?;
    let left_indent = properties
        .left_indent
        .map(ParagraphIndentPoints::from_points)
        .transpose()?;
    let right_indent = properties
        .right_indent
        .map(ParagraphIndentPoints::from_points)
        .transpose()?;
    let tab_stops = properties
        .tabs
        .as_ref()
        .map(tabs::from_archive)
        .transpose()?;
    let overrides = ParagraphStyleOverrides {
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
        underline,
        strikethrough,
        alignment,
        line_spacing,
        space_before,
        space_after,
        first_line_indent,
        left_indent,
        right_indent,
        tab_stops,
    };
    let mut remaining = properties.clone();
    remaining.alignment = None;
    if paragraph_background.is_some() {
        remaining.fill_null = None;
        remaining.fill = None;
    }
    if paragraph_borders.is_some() {
        remaining.deprecated_borders = None;
        remaining.historical_rule_offset_null = None;
        remaining.historical_rule_offset = None;
        remaining.rule_width = None;
        remaining.stroke_null = None;
        remaining.stroke = None;
        remaining.border_positions = None;
        remaining.rounded_corners = None;
    }
    remaining.hyphenate = None;
    remaining.keep_lines_together = None;
    remaining.keep_with_next = None;
    remaining.page_break_before = None;
    remaining.widow_control = None;
    remaining.line_spacing = None;
    remaining.space_before = None;
    remaining.space_after = None;
    remaining.first_line_indent = None;
    remaining.left_indent = None;
    remaining.right_indent = None;
    remaining.tabs = None;
    let mut remaining_character = character_properties.clone();
    remaining_character.bold = None;
    remaining_character.italic = None;
    remaining_character.font_size = None;
    if font_is_some {
        remaining_character.font_name_null = None;
        remaining_character.font_name = None;
    }
    remaining_character.font_color = None;
    remaining_character.tsd_fill = None;
    if capitalization.is_some() {
        remaining_character.capitalization = None;
        remaining_character.capitalization_uses_linguistics = None;
    }
    remaining_character.superscript = None;
    remaining_character.baseline_shift = None;
    remaining_character.tracking = None;
    remaining_character.ligatures = None;
    if outline.is_some() {
        remaining_character.tsd_stroke_null = None;
        remaining_character.tsd_stroke = None;
    }
    if shadow.is_some() {
        remaining_character.shadow_null = None;
        remaining_character.shadow = None;
    }
    if background.is_some() {
        remaining_character.background_color_null = None;
        remaining_character.background_color = None;
    }
    remaining_character.underline = None;
    remaining_character.strikethru = None;
    if writing_direction.is_some() {
        remaining_character.writing_direction = None;
    }
    let semantic = !overrides.is_empty()
        && remaining == tswp::ParagraphStylePropertiesArchive::default()
        && remaining_character == tswp::CharacterStylePropertiesArchive::default()
        && style.override_count == Some(overrides.count())
        && style.super_.name.is_none()
        && style.super_.style_identifier.is_none()
        && style.super_.parent.is_some()
        && style.super_.is_variation == Some(true)
        && style.super_.stylesheet.is_some();
    if !semantic {
        return Ok(None);
    }

    let super_raw = required_payload(raw, STYLE_SUPER_FIELD, "paragraph style")?;
    let character_raw = required_payload(
        raw,
        STYLE_CHARACTER_PROPERTIES_FIELD,
        "paragraph character properties",
    )?;
    let paragraph_raw = required_payload(
        raw,
        STYLE_PARAGRAPH_PROPERTIES_FIELD,
        "paragraph properties",
    )?;
    let mut character_fields = Vec::with_capacity(15);
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
        let legacy_color_raw = required_payload(
            character_raw,
            CHARACTER_FONT_COLOR_FIELD,
            "paragraph font color",
        )?;
        let drawing_fill_raw = required_payload(
            character_raw,
            CHARACTER_DRAWING_FILL_FIELD,
            "paragraph text fill",
        )?;
        let drawing_color_raw = required_payload(
            drawing_fill_raw,
            DRAWING_FILL_COLOR_FIELD,
            "paragraph text fill color",
        )?;
        if !has_canonical_color_wire(legacy_color_raw)?
            || !has_exact_fields(drawing_fill_raw, &[DRAWING_FILL_COLOR_FIELD])?
            || !has_canonical_color_wire(drawing_color_raw)?
        {
            return Ok(None);
        }
    }
    if let Some(capitalization) = capitalization {
        character_fields.push(CHARACTER_CAPITALIZATION_FIELD);
        if capitalization.uses_linguistics().is_some() {
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
            TextOutline::None => CHARACTER_DRAWING_STROKE_NULL_FIELD,
            TextOutline::Stroke(_) => CHARACTER_DRAWING_STROKE_FIELD,
        });
    }
    if let Some(shadow) = shadow {
        character_fields.push(match shadow {
            TextShadow::None => CHARACTER_SHADOW_NULL_FIELD,
            TextShadow::Drop(_) => CHARACTER_SHADOW_FIELD,
        });
    }
    if let Some(background) = background {
        let field = match background {
            TextBackground::None => CHARACTER_BACKGROUND_COLOR_NULL_FIELD,
            TextBackground::Color(_) => CHARACTER_BACKGROUND_COLOR_FIELD,
        };
        character_fields.push(field);
        if matches!(background, TextBackground::Color(_)) {
            let color_raw = required_payload(character_raw, field, "paragraph text background")?;
            if !has_canonical_color_wire(color_raw)? {
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
    let mut paragraph_fields = Vec::with_capacity(22);
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
            let color_raw = required_payload(paragraph_raw, field, "paragraph background")?;
            if !has_canonical_color_wire(color_raw)? {
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
            ParagraphBorders::None => PARAGRAPH_BORDER_STROKE_NULL_FIELD,
            ParagraphBorders::Bordered(_) => PARAGRAPH_BORDER_STROKE_FIELD,
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
            if !has_canonical_bool_field(paragraph_raw, field)? {
                return Ok(None);
            }
        }
    }
    if line_spacing.is_some() {
        paragraph_fields.push(PARAGRAPH_LINE_SPACING_FIELD);
        let line_spacing_raw = required_payload(
            paragraph_raw,
            PARAGRAPH_LINE_SPACING_FIELD,
            "paragraph line spacing",
        )?;
        let expected = match line_spacing {
            Some(ParagraphLineSpacing::Relative(_)) => vec![LINE_SPACING_AMOUNT_FIELD],
            Some(_) => vec![LINE_SPACING_MODE_FIELD, LINE_SPACING_AMOUNT_FIELD],
            None => Vec::new(),
        };
        if !has_exact_fields(line_spacing_raw, &expected)? {
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
    if let Some(stops) = overrides.tab_stops.as_ref() {
        paragraph_fields.push(PARAGRAPH_TABS_FIELD);
        let tabs_raw = required_payload(paragraph_raw, PARAGRAPH_TABS_FIELD, "paragraph tabs")?;
        if !tabs::has_canonical_wire(tabs_raw, stops)? {
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
    )? && has_exact_fields(
        super_raw,
        &[
            STYLE_PARENT_FIELD,
            STYLE_VARIATION_FIELD,
            STYLE_STYLESHEET_FIELD,
        ],
    )? && has_exact_fields(character_raw, &character_fields)?
        && has_exact_fields(paragraph_raw, &paragraph_fields)?;
    Ok(exact.then_some(overrides))
}

pub(crate) fn variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    overrides: ParagraphStyleOverrides,
) -> Result<ArchiveObject> {
    if overrides.is_empty() {
        return Err(Error::InvalidFormat(
            "an iWork paragraph-style variation must contain an override".to_owned(),
        ));
    }
    let override_count = overrides.count();
    let (font_name_null, font_name) = match overrides.font {
        Some(TextFont::Default) => (Some(true), None),
        Some(TextFont::Named(name)) => (None, Some(name.into_string())),
        None => (None, None),
    };
    let native_borders = paragraph_borders_to_native(overrides.paragraph_borders);
    let data = tswp::ParagraphStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(override_count),
        char_properties: Some(tswp::CharacterStylePropertiesArchive {
            bold: overrides.bold,
            italic: overrides.italic,
            font_size: overrides.point_size.map(TextPointSize::points),
            font_name_null,
            font_name,
            font_color: overrides.font_color.map(color_to_native),
            capitalization: overrides
                .capitalization
                .map(TextCapitalization::native_value),
            superscript: overrides.script.map(TextScript::native_value),
            baseline_shift: overrides.baseline_shift.map(TextBaselineShift::points),
            tracking: overrides
                .character_spacing
                .map(TextCharacterSpacing::native_ratio),
            ligatures: overrides.ligatures.map(TextLigatures::native_value),
            tsd_stroke_null: matches!(overrides.outline, Some(TextOutline::None)).then_some(true),
            tsd_stroke: overrides.outline.and_then(|outline| match outline {
                TextOutline::None => None,
                TextOutline::Stroke(stroke) => Some(stroke_to_native(stroke)),
            }),
            shadow_null: matches!(overrides.shadow, Some(TextShadow::None)).then_some(true),
            shadow: overrides
                .shadow
                .map(TextShadow::into_shape_shadow)
                .and_then(|shadow| match shadow {
                    crate::shapes::ShapeShadow::Disabled => None,
                    enabled => Some(shadow_to_native(enabled)),
                }),
            background_color_null: matches!(overrides.background, Some(TextBackground::None))
                .then_some(true),
            background_color: overrides
                .background
                .and_then(|background| match background {
                    TextBackground::None => None,
                    TextBackground::Color(color) => Some(color_to_native(color)),
                }),
            underline: overrides.underline.map(TextUnderline::native_value),
            strikethru: overrides.strikethrough.map(TextStrikethrough::native_value),
            tsd_fill: overrides.font_color.map(|color| tsd::FillArchive {
                color: Some(color_to_native(color)),
                ..Default::default()
            }),
            capitalization_uses_linguistics: overrides
                .capitalization
                .and_then(TextCapitalization::uses_linguistics),
            writing_direction: overrides
                .writing_direction
                .map(ParagraphWritingDirection::native_value),
            ..Default::default()
        }),
        para_properties: Some(tswp::ParagraphStylePropertiesArchive {
            alignment: overrides.alignment.map(TextAlignment::native_value),
            fill_null: matches!(
                overrides.paragraph_background,
                Some(ParagraphBackground::None)
            )
            .then_some(true),
            fill: overrides
                .paragraph_background
                .and_then(|background| match background {
                    ParagraphBackground::None => None,
                    ParagraphBackground::Color(color) => Some(color_to_native(color)),
                }),
            deprecated_borders: native_borders.deprecated_borders,
            historical_rule_offset: native_borders.historical_rule_offset,
            stroke_null: native_borders.stroke_null,
            stroke: native_borders.stroke,
            border_positions: native_borders.positions,
            rounded_corners: native_borders.rounded_corners,
            hyphenate: overrides
                .hyphenation
                .map(ParagraphHyphenation::native_value),
            keep_lines_together: overrides.keep_lines_together,
            keep_with_next: overrides.keep_with_next,
            page_break_before: overrides.start_on_new_page,
            widow_control: overrides.prevent_widow_orphan_lines,
            line_spacing: overrides.line_spacing.map(line_spacing_archive),
            space_before: overrides.space_before.map(ParagraphSpacingPoints::points),
            space_after: overrides.space_after.map(ParagraphSpacingPoints::points),
            first_line_indent: overrides
                .first_line_indent
                .map(ParagraphIndentPoints::points),
            left_indent: overrides.left_indent.map(ParagraphIndentPoints::points),
            right_indent: overrides.right_indent.map(ParagraphIndentPoints::points),
            tabs: overrides.tab_stops.as_ref().map(tabs::archive),
            ..Default::default()
        }),
    }
    .encode_to_vec();
    tswp::ParagraphStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    object.archive_info.message_infos[0]
        .object_references
        .push(parent_style_id);
    if overrides.font_color.is_some() {
        object.archive_info.message_infos[0]
            .field_infos
            .push(text_fill_field_info());
    }
    if matches!(overrides.outline, Some(TextOutline::Stroke(_))) {
        object.archive_info.message_infos[0]
            .field_infos
            .push(text_stroke_field_info());
    }
    if matches!(overrides.shadow, Some(TextShadow::Drop(_))) {
        object.archive_info.message_infos[0]
            .field_infos
            .push(text_shadow_field_info());
    }
    if matches!(overrides.background, Some(TextBackground::Color(_))) {
        object.archive_info.message_infos[0]
            .field_infos
            .push(text_background_field_info());
    }
    Ok(object)
}

pub(crate) fn replace_variation(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_id: u64,
    mut replacement: ArchiveObject,
) -> Result<()> {
    let has_text_fill_info = replacement.archive_info.message_infos[0]
        .field_infos
        .iter()
        .any(is_text_fill_field_info);
    let has_text_stroke_info = replacement.archive_info.message_infos[0]
        .field_infos
        .iter()
        .any(is_text_stroke_field_info);
    let has_text_shadow_info = replacement.archive_info.message_infos[0]
        .field_infos
        .iter()
        .any(is_text_shadow_field_info);
    let has_text_background_info = replacement.archive_info.message_infos[0]
        .field_infos
        .iter()
        .any(is_text_background_field_info);
    let message = replacement.messages.pop().ok_or_else(|| {
        Error::InvalidFormat("replacement paragraph style has no payload".to_owned())
    })?;
    if !replacement.messages.is_empty() {
        return Err(Error::InvalidFormat(
            "replacement paragraph style has multiple payloads".to_owned(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork paragraph style {style_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, candidate)| candidate.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {style_id} must have exactly one payload"
            )));
        };
        object.replace_message(*index, message)?;
        let info = &mut object.archive_info.message_infos[*index];
        info.field_infos.retain(|field| {
            !is_text_fill_field_info(field)
                && !is_text_stroke_field_info(field)
                && !is_text_shadow_field_info(field)
                && !is_text_background_field_info(field)
        });
        if has_text_fill_info {
            info.field_infos.push(text_fill_field_info());
        }
        if has_text_stroke_info {
            info.field_infos.push(text_stroke_field_info());
        }
        if has_text_shadow_info {
            info.field_infos.push(text_shadow_field_info());
        }
        if has_text_background_info {
            info.field_infos.push(text_background_field_info());
        }
        Ok(())
    })
}

pub(super) fn text_color_from_character(
    properties: &tswp::CharacterStylePropertiesArchive,
) -> Result<Option<RgbaColor>> {
    if properties.font_color_null == Some(true) || properties.tsd_fill_null == Some(true) {
        if properties.font_color.is_some() || properties.tsd_fill.is_some() {
            return Err(Error::InvalidFormat(
                "native iWork text color is both null and populated".to_owned(),
            ));
        }
        return Ok(Some(RgbaColor::black()));
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
            if fill.gradient.is_some() || fill.image.is_some() {
                return Err(Error::InvalidFormat(
                    "native iWork text fill is not a solid color".to_owned(),
                ));
            }
            let color = fill.color.as_ref().ok_or_else(|| {
                Error::InvalidFormat("native iWork text fill has no color".to_owned())
            })?;
            color_from_native(color)
        })
        .transpose()?;
    match (legacy, drawing) {
        (Some(legacy), Some(drawing)) if legacy != drawing => Err(Error::InvalidFormat(
            "native iWork legacy and drawing text colors disagree".to_owned(),
        )),
        (Some(color), _) | (_, Some(color)) => Ok(Some(color)),
        (None, None) => Ok(None),
    }
}

pub(super) fn text_font_from_character(
    properties: &tswp::CharacterStylePropertiesArchive,
) -> Result<Option<TextFont>> {
    if properties.font_name_null == Some(true) {
        if properties.font_name.is_some() {
            return Err(Error::InvalidFormat(
                "native iWork text font is both null and populated".to_owned(),
            ));
        }
        return Ok(Some(TextFont::Default));
    }
    if properties.font_name_null == Some(false) && properties.font_name.is_none() {
        return Err(Error::InvalidFormat(
            "native iWork text font has a false null marker without a name".to_owned(),
        ));
    }
    properties
        .font_name
        .as_deref()
        .map(|name| TextFontName::new(name).map(TextFont::Named))
        .transpose()
}

pub(super) fn text_outline_from_character(
    properties: &tswp::CharacterStylePropertiesArchive,
) -> Result<Option<TextOutline>> {
    if properties.tsd_stroke_null == Some(true) {
        if properties.tsd_stroke.is_some() {
            return Err(Error::InvalidFormat(
                "native iWork text outline is both null and populated".to_owned(),
            ));
        }
        return Ok(Some(TextOutline::None));
    }
    if properties.tsd_stroke_null == Some(false) && properties.tsd_stroke.is_none() {
        return Err(Error::InvalidFormat(
            "native iWork text outline has a false null marker without a stroke".to_owned(),
        ));
    }
    properties
        .tsd_stroke
        .as_ref()
        .map(|stroke| {
            stroke_from_native(stroke)
                .map(|outline| outline.map_or(TextOutline::None, TextOutline::Stroke))
        })
        .transpose()
}

pub(super) fn text_shadow_from_character(
    properties: &tswp::CharacterStylePropertiesArchive,
) -> Result<Option<TextShadow>> {
    if properties.shadow_null == Some(true) {
        if properties.shadow.is_some() {
            return Err(Error::InvalidFormat(
                "native iWork text shadow is both null and populated".to_owned(),
            ));
        }
        return Ok(Some(TextShadow::None));
    }
    if properties.shadow_null == Some(false) && properties.shadow.is_none() {
        return Err(Error::InvalidFormat(
            "native iWork text shadow has a false null marker without a shadow".to_owned(),
        ));
    }
    properties
        .shadow
        .as_ref()
        .map(|shadow| shadow_from_native(shadow).and_then(TextShadow::from_shape_shadow))
        .transpose()
}

pub(super) fn text_background_from_character(
    properties: &tswp::CharacterStylePropertiesArchive,
) -> Result<Option<TextBackground>> {
    if properties.background_color_null == Some(true) {
        if properties.background_color.is_some() {
            return Err(Error::InvalidFormat(
                "native iWork text background is both null and populated".to_owned(),
            ));
        }
        return Ok(Some(TextBackground::None));
    }
    if properties.background_color_null == Some(false) && properties.background_color.is_none() {
        return Err(Error::InvalidFormat(
            "native iWork text background has a false null marker without a color".to_owned(),
        ));
    }
    properties
        .background_color
        .as_ref()
        .map(|color| color_from_native(color).map(TextBackground::Color))
        .transpose()
}

pub(super) fn paragraph_background_from_properties(
    properties: &tswp::ParagraphStylePropertiesArchive,
) -> Result<Option<ParagraphBackground>> {
    if properties.fill_null == Some(true) {
        if properties.fill.is_some() {
            return Err(Error::InvalidFormat(
                "native iWork paragraph background is both null and populated".to_owned(),
            ));
        }
        return Ok(Some(ParagraphBackground::None));
    }
    if properties.fill_null == Some(false) && properties.fill.is_none() {
        return Err(Error::InvalidFormat(
            "native iWork paragraph background has a false null marker without a color".to_owned(),
        ));
    }
    properties
        .fill
        .as_ref()
        .map(|color| color_from_native(color).map(ParagraphBackground::Color))
        .transpose()
}

pub(super) fn paragraph_borders_from_properties(
    properties: &tswp::ParagraphStylePropertiesArchive,
) -> Result<Option<ParagraphBorders>> {
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
        return Err(Error::InvalidFormat(
            "native paragraph-border offset is both null and populated".to_owned(),
        ));
    }
    if properties.historical_rule_offset_null == Some(false)
        && properties.historical_rule_offset.is_none()
    {
        return Err(Error::InvalidFormat(
            "native paragraph-border offset has a false null marker without a point".to_owned(),
        ));
    }
    if properties.stroke_null == Some(true) && properties.stroke.is_some() {
        return Err(Error::InvalidFormat(
            "native paragraph-border stroke is both null and populated".to_owned(),
        ));
    }
    if properties.stroke_null == Some(false) && properties.stroke.is_none() {
        return Err(Error::InvalidFormat(
            "native paragraph-border stroke has a false null marker without a stroke".to_owned(),
        ));
    }
    if properties
        .rule_width
        .is_some_and(|width| !width.is_finite() || width < 0.0)
    {
        return Err(Error::InvalidFormat(
            "native legacy paragraph-border width is not finite and nonnegative".to_owned(),
        ));
    }

    let legacy_sides = properties
        .deprecated_borders
        .map(paragraph_border_sides_from_legacy)
        .transpose()?;
    let modern_sides = properties
        .border_positions
        .map(ParagraphBorderSides::from_native_bits)
        .transpose()?;
    if legacy_sides
        .zip(modern_sides)
        .is_some_and(|(legacy, modern)| legacy != modern)
    {
        return Err(Error::InvalidFormat(
            "native paragraph-border side encodings disagree".to_owned(),
        ));
    }
    let sides = modern_sides
        .or(legacy_sides)
        .unwrap_or(ParagraphBorderSides::NONE);
    let native_stroke = properties
        .stroke
        .as_ref()
        .map(stroke_from_native)
        .transpose()?
        .flatten();
    if sides.is_empty() {
        if native_stroke.is_some() {
            return Err(Error::InvalidFormat(
                "native paragraph border has a visible stroke without sides".to_owned(),
            ));
        }
        return Ok(Some(ParagraphBorders::None));
    }
    if properties.stroke_null == Some(true) {
        return Err(Error::InvalidFormat(
            "native paragraph border has sides but an explicit null stroke".to_owned(),
        ));
    }
    let stroke = native_stroke.ok_or_else(|| {
        Error::InvalidFormat("native paragraph border has sides but no visible stroke".to_owned())
    })?;
    if stroke.cap != StrokeCap::Round || stroke.join != StrokeJoin::Round {
        return Err(Error::InvalidFormat(
            "native paragraph border does not use app-standard round stroke geometry".to_owned(),
        ));
    }
    let offset = match properties.historical_rule_offset.as_ref() {
        Some(point) => {
            if !point.x.is_finite() || !point.y.is_finite() || point.x != point.y {
                return Err(Error::InvalidFormat(
                    "native paragraph-border offset must contain equal finite axes".to_owned(),
                ));
            }
            ParagraphBorderOffset::from_native_inset(point.x)?
        },
        None => ParagraphBorderOffset::DEFAULT,
    };
    let rounded_corners =
        properties.rounded_corners.unwrap_or(false) && sides == ParagraphBorderSides::ALL;
    Ok(Some(ParagraphBorders::Bordered(ParagraphBorder::new(
        stroke.color,
        stroke.width,
        stroke.pattern,
        sides,
        offset,
        rounded_corners,
    )?)))
}

fn paragraph_borders_to_native(borders: Option<ParagraphBorders>) -> NativeParagraphBorderFields {
    match borders {
        None => NativeParagraphBorderFields::default(),
        Some(ParagraphBorders::None) => NativeParagraphBorderFields {
            deprecated_borders: Some(0),
            stroke_null: Some(true),
            positions: Some(ParagraphBorderSides::NONE.native_bits()),
            rounded_corners: Some(false),
            ..Default::default()
        },
        Some(ParagraphBorders::Bordered(border)) => {
            let offset = (border.offset() != ParagraphBorderOffset::DEFAULT).then(|| {
                let inset = border.offset().native_inset();
                tsp::Point { x: inset, y: inset }
            });
            NativeParagraphBorderFields {
                deprecated_borders: Some(paragraph_border_legacy_value(border.sides())),
                historical_rule_offset: offset,
                stroke: Some(stroke_to_native(border.native_stroke())),
                positions: Some(border.sides().native_bits()),
                rounded_corners: Some(border.has_rounded_corners()),
                ..Default::default()
            }
        },
    }
}

fn paragraph_border_legacy_value(sides: ParagraphBorderSides) -> i32 {
    if sides == ParagraphBorderSides::ALL {
        return LEGACY_BORDER_ALL;
    }
    i32::from(sides.contains(ParagraphBorderSides::TOP)) * LEGACY_BORDER_TOP
        + i32::from(sides.contains(ParagraphBorderSides::BOTTOM)) * LEGACY_BORDER_BOTTOM
        + i32::from(sides.contains(ParagraphBorderSides::LEFT)) * LEGACY_BORDER_LEFT
        + i32::from(sides.contains(ParagraphBorderSides::RIGHT)) * LEGACY_BORDER_RIGHT
}

fn paragraph_border_sides_from_legacy(value: i32) -> Result<ParagraphBorderSides> {
    if value == LEGACY_BORDER_ALL {
        return Ok(ParagraphBorderSides::ALL);
    }
    let known = LEGACY_BORDER_TOP | LEGACY_BORDER_BOTTOM | LEGACY_BORDER_LEFT | LEGACY_BORDER_RIGHT;
    if value < 0 || value & !known != 0 {
        return Err(Error::InvalidFormat(
            "native paragraph border has an unknown legacy side value".to_owned(),
        ));
    }
    Ok(ParagraphBorderSides::new(
        value & LEGACY_BORDER_TOP != 0,
        value & LEGACY_BORDER_BOTTOM != 0,
        value & LEGACY_BORDER_LEFT != 0,
        value & LEGACY_BORDER_RIGHT != 0,
    ))
}

pub(super) fn capitalization_from_character(
    properties: &tswp::CharacterStylePropertiesArchive,
) -> Result<Option<TextCapitalization>> {
    let Some(value) = properties.capitalization else {
        if properties.capitalization_uses_linguistics.is_some() {
            return Err(Error::InvalidFormat(
                "native iWork linguistic capitalization has no capitalization type".to_owned(),
            ));
        }
        return Ok(None);
    };
    TextCapitalization::from_native_value(value, properties.capitalization_uses_linguistics)
        .map(Some)
}

fn has_canonical_color_wire(raw: &[u8]) -> Result<bool> {
    has_exact_fields(
        raw,
        &[
            COLOR_MODEL_FIELD,
            COLOR_RED_FIELD,
            COLOR_GREEN_FIELD,
            COLOR_BLUE_FIELD,
            COLOR_ALPHA_FIELD,
            COLOR_RGB_SPACE_FIELD,
        ],
    )
}

fn text_fill_field_info() -> tsp::FieldInfo {
    tsp::FieldInfo {
        path: tsp::FieldPath {
            path: vec![
                STYLE_CHARACTER_PROPERTIES_FIELD,
                CHARACTER_DRAWING_FILL_FIELD,
            ],
        },
        r#type: Some(tsp::field_info::Type::Message as i32),
        unknown_field_rule: Some(tsp::field_info::UnknownFieldRule::IgnoreAndPreserve as i32),
        ..Default::default()
    }
}

fn text_stroke_field_info() -> tsp::FieldInfo {
    tsp::FieldInfo {
        path: tsp::FieldPath {
            path: vec![
                STYLE_CHARACTER_PROPERTIES_FIELD,
                CHARACTER_DRAWING_STROKE_FIELD,
            ],
        },
        r#type: Some(tsp::field_info::Type::Message as i32),
        unknown_field_rule: Some(tsp::field_info::UnknownFieldRule::IgnoreAndPreserve as i32),
        ..Default::default()
    }
}

fn text_shadow_field_info() -> tsp::FieldInfo {
    tsp::FieldInfo {
        path: tsp::FieldPath {
            path: vec![STYLE_CHARACTER_PROPERTIES_FIELD, CHARACTER_SHADOW_FIELD],
        },
        r#type: Some(tsp::field_info::Type::Message as i32),
        unknown_field_rule: Some(tsp::field_info::UnknownFieldRule::IgnoreAndPreserve as i32),
        ..Default::default()
    }
}

fn text_background_field_info() -> tsp::FieldInfo {
    tsp::FieldInfo {
        path: tsp::FieldPath {
            path: vec![
                STYLE_CHARACTER_PROPERTIES_FIELD,
                CHARACTER_BACKGROUND_COLOR_FIELD,
            ],
        },
        r#type: Some(tsp::field_info::Type::Message as i32),
        unknown_field_rule: Some(tsp::field_info::UnknownFieldRule::IgnoreAndPreserve as i32),
        ..Default::default()
    }
}

fn is_text_fill_field_info(field: &tsp::FieldInfo) -> bool {
    field.path.path
        == [
            STYLE_CHARACTER_PROPERTIES_FIELD,
            CHARACTER_DRAWING_FILL_FIELD,
        ]
}

fn is_text_stroke_field_info(field: &tsp::FieldInfo) -> bool {
    field.path.path
        == [
            STYLE_CHARACTER_PROPERTIES_FIELD,
            CHARACTER_DRAWING_STROKE_FIELD,
        ]
}

fn is_text_shadow_field_info(field: &tsp::FieldInfo) -> bool {
    field.path.path == [STYLE_CHARACTER_PROPERTIES_FIELD, CHARACTER_SHADOW_FIELD]
}

fn is_text_background_field_info(field: &tsp::FieldInfo) -> bool {
    field.path.path
        == [
            STYLE_CHARACTER_PROPERTIES_FIELD,
            CHARACTER_BACKGROUND_COLOR_FIELD,
        ]
}

pub(super) fn is_exclusive(package: &IWorkPackage, style_id: u64) -> Result<bool> {
    let mut storage_references = 0usize;
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            for message in &object.messages {
                if STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && let Ok(storage) = tswp::StorageArchive::decode(message.data.as_slice())
                {
                    storage_references += storage
                        .table_para_style
                        .iter()
                        .flat_map(|table| &table.entries)
                        .filter(|entry| {
                            entry
                                .object
                                .as_ref()
                                .is_some_and(|reference| reference.identifier == style_id)
                        })
                        .count();
                }
                if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                    && let Ok(style) = tswp::ParagraphStyleArchive::decode(message.data.as_slice())
                    && style
                        .super_
                        .parent
                        .as_ref()
                        .is_some_and(|parent| parent.identifier == style_id)
                {
                    return Ok(false);
                }
            }
        }
    }
    Ok(storage_references == 1)
}

pub(crate) fn parent_style_id(style: &tswp::ParagraphStyleArchive, style_id: u64) -> Result<u64> {
    style
        .super_
        .parent
        .as_ref()
        .map(|parent| parent.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph-style variation {style_id} has no parent"
            ))
        })
}

pub(crate) fn stylesheet_id(style: &tswp::ParagraphStyleArchive, style_id: u64) -> Result<u64> {
    style
        .super_
        .stylesheet
        .as_ref()
        .map(|stylesheet| stylesheet.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {style_id} has no stylesheet"
            ))
        })
}

fn line_spacing_archive(spacing: ParagraphLineSpacing) -> tswp::LineSpacingArchive {
    match spacing {
        ParagraphLineSpacing::Relative(multiple) => tswp::LineSpacingArchive {
            amount: Some(multiple.get()),
            ..Default::default()
        },
        ParagraphLineSpacing::AtLeast(points) => point_spacing(MINIMUM_LINE_SPACING_MODE, points),
        ParagraphLineSpacing::Exactly(points) => point_spacing(EXACT_LINE_SPACING_MODE, points),
        ParagraphLineSpacing::Maximum(points) => point_spacing(MAXIMUM_LINE_SPACING_MODE, points),
        ParagraphLineSpacing::Between(points) => point_spacing(BETWEEN_LINE_SPACING_MODE, points),
    }
}

fn point_spacing(mode: i32, points: ParagraphLineSpacingPoints) -> tswp::LineSpacingArchive {
    tswp::LineSpacingArchive {
        mode: Some(mode),
        amount: Some(points.points()),
        ..Default::default()
    }
}

fn line_spacing_from_archive(spacing: &tswp::LineSpacingArchive) -> Result<ParagraphLineSpacing> {
    if spacing.baseline_rule.is_some() {
        return Err(Error::InvalidFormat(
            "unsupported native iWork line-spacing baseline rule".to_owned(),
        ));
    }
    let mode = spacing.mode.unwrap_or(RELATIVE_LINE_SPACING_MODE);
    if mode == RELATIVE_LINE_SPACING_MODE {
        let multiple = spacing.amount.unwrap_or(1.0);
        return Ok(ParagraphLineSpacing::Relative(
            ParagraphLineSpacingMultiple::new(multiple)?,
        ));
    }
    let points = spacing.amount.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "native iWork line-spacing mode {mode} has no amount"
        ))
    })?;
    let points = ParagraphLineSpacingPoints::from_points(points)?;
    match mode {
        MINIMUM_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::AtLeast(points)),
        EXACT_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::Exactly(points)),
        MAXIMUM_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::Maximum(points)),
        BETWEEN_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::Between(points)),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported native iWork line-spacing mode {mode}"
        ))),
    }
}

fn has_exact_fields(data: &[u8], expected: &[u32]) -> Result<bool> {
    let mut actual = parse_wire_fields(data)?
        .into_iter()
        .map(|field| field.number)
        .collect::<Vec<_>>();
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    Ok(actual == expected)
}

fn has_canonical_bool_field(data: &[u8], number: u32) -> Result<bool> {
    let fields = parse_wire_fields(data)?
        .into_iter()
        .filter(|field| field.number == number)
        .collect::<Vec<_>>();
    let [field] = fields.as_slice() else {
        return Ok(false);
    };
    Ok(field.wire_type == 0
        && matches!(
            &data[field.key_end..field.end],
            [PROTOBUF_FALSE_BYTE] | [PROTOBUF_TRUE_BYTE]
        ))
}

fn required_payload<'a>(data: &'a [u8], field: u32, context: &str) -> Result<&'a [u8]> {
    let payloads = repeated_length_delimited_payloads(data, field)?;
    let [payload] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain field {field} exactly once"
        )));
    };
    Ok(payload)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
