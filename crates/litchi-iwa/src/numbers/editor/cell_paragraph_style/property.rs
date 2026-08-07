//! Typed paragraph properties supported by the native table-cell style graph.

use crate::shapes::RgbaColor;
use crate::text::paragraph_alignment::native::{
    ParagraphStyleOverrides, inherited_alignment, inherited_indents, inherited_line_spacing,
    inherited_spacing, inherited_tab_stops, inherited_text_background,
    inherited_text_baseline_shift, inherited_text_capitalization, inherited_text_character_spacing,
    inherited_text_color, inherited_text_decorations, inherited_text_font,
    inherited_text_ligatures, inherited_text_outline, inherited_text_script, inherited_text_shadow,
    inherited_text_style,
};
use crate::text::{
    Alignment, Background, Indents, LineSpacing, Outline, ParagraphTabStops, Shadow, Spacing,
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations, TextFont,
    TextLigatures, TextScript, TextStyle,
};
use crate::{Error, IWorkPackage, Result};

#[derive(Debug, Clone, PartialEq)]
pub(super) enum CellParagraphProperty {
    Alignment(Alignment),
    Background(Background),
    BaselineShift(TextBaselineShift),
    Capitalization(TextCapitalization),
    CharacterSpacing(TextCharacterSpacing),
    Color(RgbaColor),
    Decorations(TextDecorations),
    Font(TextFont),
    Indents(Indents),
    Ligatures(TextLigatures),
    LineSpacing(LineSpacing),
    Outline(Outline),
    Script(TextScript),
    Shadow(Shadow),
    Spacing(Spacing),
    TabStops(ParagraphTabStops),
    TextStyle(TextStyle),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CellParagraphPropertyKind {
    Alignment,
    Background,
    BaselineShift,
    Capitalization,
    CharacterSpacing,
    Color,
    Decorations,
    Font,
    Indents,
    Ligatures,
    LineSpacing,
    Outline,
    Script,
    Shadow,
    Spacing,
    TabStops,
    TextStyle,
}

impl CellParagraphProperty {
    pub(super) const fn kind(&self) -> CellParagraphPropertyKind {
        match self {
            Self::Alignment(_) => CellParagraphPropertyKind::Alignment,
            Self::Background(_) => CellParagraphPropertyKind::Background,
            Self::BaselineShift(_) => CellParagraphPropertyKind::BaselineShift,
            Self::Capitalization(_) => CellParagraphPropertyKind::Capitalization,
            Self::CharacterSpacing(_) => CellParagraphPropertyKind::CharacterSpacing,
            Self::Color(_) => CellParagraphPropertyKind::Color,
            Self::Decorations(_) => CellParagraphPropertyKind::Decorations,
            Self::Font(_) => CellParagraphPropertyKind::Font,
            Self::Indents(_) => CellParagraphPropertyKind::Indents,
            Self::Ligatures(_) => CellParagraphPropertyKind::Ligatures,
            Self::LineSpacing(_) => CellParagraphPropertyKind::LineSpacing,
            Self::Outline(_) => CellParagraphPropertyKind::Outline,
            Self::Script(_) => CellParagraphPropertyKind::Script,
            Self::Shadow(_) => CellParagraphPropertyKind::Shadow,
            Self::Spacing(_) => CellParagraphPropertyKind::Spacing,
            Self::TabStops(_) => CellParagraphPropertyKind::TabStops,
            Self::TextStyle(_) => CellParagraphPropertyKind::TextStyle,
        }
    }

    pub(super) fn inherited(
        package: &IWorkPackage,
        style_id: u64,
        kind: CellParagraphPropertyKind,
    ) -> Result<Self> {
        match kind {
            CellParagraphPropertyKind::Alignment => {
                inherited_alignment(package, style_id).map(Self::Alignment)
            },
            CellParagraphPropertyKind::Background => {
                inherited_text_background(package, style_id).map(Self::Background)
            },
            CellParagraphPropertyKind::BaselineShift => {
                inherited_text_baseline_shift(package, style_id).map(Self::BaselineShift)
            },
            CellParagraphPropertyKind::Capitalization => {
                inherited_text_capitalization(package, style_id).map(Self::Capitalization)
            },
            CellParagraphPropertyKind::CharacterSpacing => {
                inherited_text_character_spacing(package, style_id).map(Self::CharacterSpacing)
            },
            CellParagraphPropertyKind::Color => {
                inherited_text_color(package, style_id).map(Self::Color)
            },
            CellParagraphPropertyKind::Decorations => {
                inherited_text_decorations(package, style_id).map(Self::Decorations)
            },
            CellParagraphPropertyKind::Font => {
                inherited_text_font(package, style_id).map(Self::Font)
            },
            CellParagraphPropertyKind::Indents => {
                inherited_indents(package, style_id).map(Self::Indents)
            },
            CellParagraphPropertyKind::Ligatures => {
                inherited_text_ligatures(package, style_id).map(Self::Ligatures)
            },
            CellParagraphPropertyKind::LineSpacing => {
                inherited_line_spacing(package, style_id).map(Self::LineSpacing)
            },
            CellParagraphPropertyKind::Outline => {
                inherited_text_outline(package, style_id).map(Self::Outline)
            },
            CellParagraphPropertyKind::Script => {
                inherited_text_script(package, style_id).map(Self::Script)
            },
            CellParagraphPropertyKind::Shadow => {
                inherited_text_shadow(package, style_id).map(Self::Shadow)
            },
            CellParagraphPropertyKind::Spacing => {
                inherited_spacing(package, style_id).map(Self::Spacing)
            },
            CellParagraphPropertyKind::TabStops => {
                inherited_tab_stops(package, style_id).map(Self::TabStops)
            },
            CellParagraphPropertyKind::TextStyle => {
                inherited_text_style(package, style_id).map(Self::TextStyle)
            },
        }
    }

    pub(super) fn apply_to(
        &self,
        overrides: &mut ParagraphStyleOverrides,
        inherited: &Self,
    ) -> Result<()> {
        match (self, inherited) {
            (Self::Alignment(value), Self::Alignment(_)) => {
                overrides.alignment = Some(*value);
            },
            (Self::Background(value), Self::Background(inherited)) => {
                overrides.background = (value != inherited).then_some(*value);
            },
            (Self::BaselineShift(value), Self::BaselineShift(inherited)) => {
                overrides.baseline_shift = (value != inherited).then_some(*value);
            },
            (Self::Capitalization(value), Self::Capitalization(inherited)) => {
                overrides.capitalization = (value != inherited).then_some(*value);
            },
            (Self::CharacterSpacing(value), Self::CharacterSpacing(inherited)) => {
                overrides.character_spacing = (value != inherited).then_some(*value);
            },
            (Self::Color(value), Self::Color(inherited)) => {
                overrides.font_color = (value != inherited).then_some(*value);
            },
            (Self::Decorations(value), Self::Decorations(inherited)) => {
                overrides.underline =
                    (value.underline != inherited.underline).then_some(value.underline);
                overrides.strikethrough =
                    (value.strikethrough != inherited.strikethrough).then_some(value.strikethrough);
            },
            (Self::Font(value), Self::Font(inherited)) => {
                overrides.font = (value != inherited).then(|| value.clone());
            },
            (Self::Indents(value), Self::Indents(inherited)) => {
                overrides.first_line_indent =
                    (value.first_line != inherited.first_line).then_some(value.first_line);
                overrides.left_indent = (value.left != inherited.left).then_some(value.left);
                overrides.right_indent = (value.right != inherited.right).then_some(value.right);
            },
            (Self::Ligatures(value), Self::Ligatures(inherited)) => {
                overrides.ligatures = (value != inherited).then_some(*value);
            },
            (Self::LineSpacing(value), Self::LineSpacing(inherited)) => {
                overrides.line_spacing = (value != inherited).then_some(*value);
            },
            (Self::Outline(value), Self::Outline(inherited)) => {
                overrides.outline = (value != inherited).then_some(*value);
            },
            (Self::Script(value), Self::Script(inherited)) => {
                overrides.script = (value != inherited).then_some(*value);
            },
            (Self::Shadow(value), Self::Shadow(inherited)) => {
                overrides.shadow = (value != inherited).then_some(*value);
            },
            (Self::Spacing(value), Self::Spacing(inherited)) => {
                overrides.space_before = (value.before != inherited.before).then_some(value.before);
                overrides.space_after = (value.after != inherited.after).then_some(value.after);
            },
            (Self::TabStops(value), Self::TabStops(inherited)) => {
                overrides.tab_stops = (value != inherited).then(|| value.clone());
            },
            (Self::TextStyle(value), Self::TextStyle(inherited)) => {
                overrides.point_size =
                    (value.point_size != inherited.point_size).then_some(value.point_size);
                overrides.bold = (value.bold != inherited.bold).then_some(value.bold);
                overrides.italic = (value.italic != inherited.italic).then_some(value.italic);
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "iWork table-cell paragraph property kind mismatch".to_owned(),
                ));
            },
        }
        Ok(())
    }
}

impl CellParagraphPropertyKind {
    pub(super) const fn name(self) -> &'static str {
        match self {
            Self::Alignment => "text alignment",
            Self::Background => "text background",
            Self::BaselineShift => "baseline shift",
            Self::Capitalization => "capitalization",
            Self::CharacterSpacing => "character spacing",
            Self::Color => "text color",
            Self::Decorations => "text decorations",
            Self::Font => "font",
            Self::Indents => "paragraph indents",
            Self::Ligatures => "ligatures",
            Self::LineSpacing => "paragraph line spacing",
            Self::Outline => "text outline",
            Self::Script => "baseline script",
            Self::Shadow => "text shadow",
            Self::Spacing => "paragraph spacing",
            Self::TabStops => "paragraph tab stops",
            Self::TextStyle => "character formatting",
        }
    }

    pub(super) fn has_direct(self, overrides: &ParagraphStyleOverrides) -> bool {
        match self {
            Self::Alignment => overrides.alignment.is_some(),
            Self::Background => overrides.background.is_some(),
            Self::BaselineShift => overrides.baseline_shift.is_some(),
            Self::Capitalization => overrides.capitalization.is_some(),
            Self::CharacterSpacing => overrides.character_spacing.is_some(),
            Self::Color => overrides.font_color.is_some(),
            Self::Decorations => overrides.underline.is_some() || overrides.strikethrough.is_some(),
            Self::Font => overrides.font.is_some(),
            Self::Indents => {
                overrides.first_line_indent.is_some()
                    || overrides.left_indent.is_some()
                    || overrides.right_indent.is_some()
            },
            Self::Ligatures => overrides.ligatures.is_some(),
            Self::LineSpacing => overrides.line_spacing.is_some(),
            Self::Outline => overrides.outline.is_some(),
            Self::Script => overrides.script.is_some(),
            Self::Shadow => overrides.shadow.is_some(),
            Self::Spacing => overrides.space_before.is_some() || overrides.space_after.is_some(),
            Self::TabStops => overrides.tab_stops.is_some(),
            Self::TextStyle => {
                overrides.point_size.is_some()
                    || overrides.bold.is_some()
                    || overrides.italic.is_some()
            },
        }
    }

    pub(super) fn clear(self, overrides: &mut ParagraphStyleOverrides) {
        match self {
            Self::Alignment => overrides.alignment = None,
            Self::Background => overrides.background = None,
            Self::BaselineShift => overrides.baseline_shift = None,
            Self::Capitalization => overrides.capitalization = None,
            Self::CharacterSpacing => overrides.character_spacing = None,
            Self::Color => overrides.font_color = None,
            Self::Decorations => {
                overrides.underline = None;
                overrides.strikethrough = None;
            },
            Self::Font => overrides.font = None,
            Self::Indents => {
                overrides.first_line_indent = None;
                overrides.left_indent = None;
                overrides.right_indent = None;
            },
            Self::Ligatures => overrides.ligatures = None,
            Self::LineSpacing => overrides.line_spacing = None,
            Self::Outline => overrides.outline = None,
            Self::Script => overrides.script = None,
            Self::Shadow => overrides.shadow = None,
            Self::Spacing => {
                overrides.space_before = None;
                overrides.space_after = None;
            },
            Self::TabStops => overrides.tab_stops = None,
            Self::TextStyle => {
                overrides.point_size = None;
                overrides.bold = None;
                overrides.italic = None;
            },
        }
    }
}
