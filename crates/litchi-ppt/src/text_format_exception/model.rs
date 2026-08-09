//! Semantic text-format exception values for legacy `PowerPoint` text.

use crate::slide_show_settings::ColorIndex;
use crate::text_run::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphTabStop, ParagraphTextDirection,
};

// CFStyle bits (MS-PPT 2.9.16).
const CF_STYLE_BOLD: u16 = 0x0001;
const CF_STYLE_ITALIC: u16 = 0x0002;
const CF_STYLE_UNDERLINE: u16 = 0x0004;
const CF_STYLE_SHADOW: u16 = 0x0010;
const CF_STYLE_FEHINT: u16 = 0x0020;
const CF_STYLE_KUMI: u16 = 0x0080;
const CF_STYLE_EMBOSS: u16 = 0x0200;
const CF_STYLE_PP9RT_SHIFT: u16 = 10;
const CF_STYLE_PP9RT_MASK: u16 = 0x000F;

/// A `CFStyle` character-style bitfield (MS-PPT 2.9.16). The raw bits are
/// preserved because the `unused` bits are undefined.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct CFStyle {
    pub(super) bits: u16,
}

impl CFStyle {
    /// Whether the characters are bold.
    #[must_use]
    pub const fn bold(&self) -> bool {
        self.bits & CF_STYLE_BOLD != 0
    }
    /// Whether the characters are italicized.
    #[must_use]
    pub const fn italic(&self) -> bool {
        self.bits & CF_STYLE_ITALIC != 0
    }
    /// Whether the characters are underlined.
    #[must_use]
    pub const fn underline(&self) -> bool {
        self.bits & CF_STYLE_UNDERLINE != 0
    }
    /// Whether the characters have a shadow effect.
    #[must_use]
    pub const fn shadow(&self) -> bool {
        self.bits & CF_STYLE_SHADOW != 0
    }
    /// Whether the characters originated from double-byte input.
    #[must_use]
    pub const fn fehint(&self) -> bool {
        self.bits & CF_STYLE_FEHINT != 0
    }
    /// Whether Kumimoji are used for vertical text.
    #[must_use]
    pub const fn kumi(&self) -> bool {
        self.bits & CF_STYLE_KUMI != 0
    }
    /// Whether the characters are embossed.
    #[must_use]
    pub const fn emboss(&self) -> bool {
        self.bits & CF_STYLE_EMBOSS != 0
    }
    /// The four-bit `StyleTextProp9Atom` run grouping (`pp9rt`).
    #[must_use]
    pub const fn pp9_run_group(&self) -> u8 {
        ((self.bits >> CF_STYLE_PP9RT_SHIFT) & CF_STYLE_PP9RT_MASK) as u8
    }
}

/// A parsed `TextCFException` structure (MS-PPT 2.9.14) with character-level
/// style and formatting defaults. The `CFMasks` value is preserved verbatim
/// so that undefined `unused` bits round-trip exactly.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextCFException {
    pub(super) masks: u32,
    pub(super) font_style: Option<CFStyle>,
    pub(super) font_ref: Option<u16>,
    pub(super) old_east_asian_font_ref: Option<u16>,
    pub(super) ansi_font_ref: Option<u16>,
    pub(super) symbol_font_ref: Option<u16>,
    pub(super) font_size: Option<i16>,
    pub(super) color: Option<ColorIndex>,
    pub(super) position: Option<i16>,
}

impl TextCFException {
    /// The raw `CFMasks` value (MS-PPT 2.9.15).
    #[must_use]
    pub const fn masks(&self) -> u32 {
        self.masks
    }
    /// The `CFStyle` character style, when present.
    #[must_use]
    pub const fn font_style(&self) -> Option<CFStyle> {
        self.font_style
    }
    /// Zero-based index of the font in the `FontCollectionContainer`,
    /// when present (MS-PPT 2.2.10).
    #[must_use]
    pub const fn font_ref(&self) -> Option<u16> {
        self.font_ref
    }
    /// Zero-based index of an East Asian font, when present.
    #[must_use]
    pub const fn old_east_asian_font_ref(&self) -> Option<u16> {
        self.old_east_asian_font_ref
    }
    /// Zero-based index of an ANSI font, when present.
    #[must_use]
    pub const fn ansi_font_ref(&self) -> Option<u16> {
        self.ansi_font_ref
    }
    /// Zero-based index of a symbol font, when present.
    #[must_use]
    pub const fn symbol_font_ref(&self) -> Option<u16> {
        self.symbol_font_ref
    }
    /// Font size in points, when present; within 1..=4000 (MS-PPT 2.9.14).
    #[must_use]
    pub const fn font_size(&self) -> Option<i16> {
        self.font_size
    }
    /// Text color, when present (MS-PPT 2.12.2).
    #[must_use]
    pub const fn color(&self) -> Option<ColorIndex> {
        self.color
    }
    /// Baseline position as a percentage of line height, when present;
    /// within -100..=100 (MS-PPT 2.9.14).
    #[must_use]
    pub const fn position(&self) -> Option<i16> {
        self.position
    }
}

/// `BulletFlags` bullet-property validity bits (MS-PPT 2.9.22).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "each bool mirrors one independent validity bit of the MS-PPT `BulletFlags` bitfield; they are spec flags, not a state machine"
)]
pub struct BulletFlags {
    pub(super) has_bullet: bool,
    pub(super) bullet_has_font: bool,
    pub(super) bullet_has_color: bool,
    pub(super) bullet_has_size: bool,
}

impl BulletFlags {
    /// Whether a bullet exists.
    #[must_use]
    pub const fn has_bullet(&self) -> bool {
        self.has_bullet
    }
    /// Whether the bullet has a font.
    #[must_use]
    pub const fn bullet_has_font(&self) -> bool {
        self.bullet_has_font
    }
    /// Whether the bullet has a color.
    #[must_use]
    pub const fn bullet_has_color(&self) -> bool {
        self.bullet_has_color
    }
    /// Whether the bullet has a size.
    #[must_use]
    pub const fn bullet_has_size(&self) -> bool {
        self.bullet_has_size
    }
}

/// `PFWrapFlags` line-breaking settings (MS-PPT 2.9.25).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct WrapFlags {
    pub(super) char_wrap: bool,
    pub(super) word_wrap: bool,
    pub(super) overflow: bool,
}

impl WrapFlags {
    /// Whether the paragraph follows the East Asian kinsoku line-breaking
    /// settings.
    #[must_use]
    pub const fn char_wrap(&self) -> bool {
        self.char_wrap
    }
    /// Whether text wraps only at word breaks.
    #[must_use]
    pub const fn word_wrap(&self) -> bool {
        self.word_wrap
    }
    /// Whether hanging punctuation is allowed for East Asian text.
    #[must_use]
    pub const fn overflow(&self) -> bool {
        self.overflow
    }
}

/// A parsed `TextPFException` structure (MS-PPT 2.9.20) with paragraph-level
/// formatting defaults. The `PFMasks` value is preserved verbatim so that the
/// undefined `unused` bit round-trips exactly.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextPFException {
    pub(super) masks: u32,
    pub(super) bullet_flags: Option<BulletFlags>,
    pub(super) bullet_char: Option<u16>,
    pub(super) bullet_font_ref: Option<u16>,
    pub(super) bullet_size: Option<i16>,
    pub(super) bullet_color: Option<ColorIndex>,
    pub(super) text_alignment: Option<ParagraphAlignment>,
    pub(super) line_spacing: Option<i16>,
    pub(super) space_before: Option<i16>,
    pub(super) space_after: Option<i16>,
    pub(super) left_margin: Option<i16>,
    pub(super) indent: Option<i16>,
    pub(super) default_tab_size: Option<i16>,
    pub(super) tab_stops: Option<Vec<ParagraphTabStop>>,
    pub(super) font_align: Option<ParagraphFontAlignment>,
    pub(super) wrap_flags: Option<WrapFlags>,
    pub(super) text_direction: Option<ParagraphTextDirection>,
}

impl TextPFException {
    /// The raw `PFMasks` value (MS-PPT 2.9.21).
    #[must_use]
    pub const fn masks(&self) -> u32 {
        self.masks
    }
    /// Bullet validity flags, when present.
    #[must_use]
    pub const fn bullet_flags(&self) -> Option<BulletFlags> {
        self.bullet_flags
    }
    /// UTF-16 code unit displayed as the bullet, when present; never NUL
    /// (MS-PPT 2.9.20).
    #[must_use]
    pub const fn bullet_char(&self) -> Option<u16> {
        self.bullet_char
    }
    /// Zero-based index of the bullet font, when present.
    #[must_use]
    pub const fn bullet_font_ref(&self) -> Option<u16> {
        self.bullet_font_ref
    }
    /// `BulletSize` of the bullet, when present (MS-PPT 2.2.3).
    #[must_use]
    pub const fn bullet_size(&self) -> Option<i16> {
        self.bullet_size
    }
    /// Bullet color, when present (MS-PPT 2.12.2).
    #[must_use]
    pub const fn bullet_color(&self) -> Option<ColorIndex> {
        self.bullet_color
    }
    /// Paragraph alignment, when present (MS-PPT 2.13.27).
    #[must_use]
    pub const fn text_alignment(&self) -> Option<ParagraphAlignment> {
        self.text_alignment
    }
    /// `ParaSpacing` between lines, when present (MS-PPT 2.2.20).
    #[must_use]
    pub const fn line_spacing(&self) -> Option<i16> {
        self.line_spacing
    }
    /// `ParaSpacing` before the paragraph, when present.
    #[must_use]
    pub const fn space_before(&self) -> Option<i16> {
        self.space_before
    }
    /// `ParaSpacing` after the paragraph, when present.
    #[must_use]
    pub const fn space_after(&self) -> Option<i16> {
        self.space_after
    }
    /// Left margin in master units, when present (MS-PPT 2.2.15).
    #[must_use]
    pub const fn left_margin(&self) -> Option<i16> {
        self.left_margin
    }
    /// Paragraph indentation in master units, when present.
    #[must_use]
    pub const fn indent(&self) -> Option<i16> {
        self.indent
    }
    /// Default tab size in master units, when present (MS-PPT 2.2.29).
    #[must_use]
    pub const fn default_tab_size(&self) -> Option<i16> {
        self.default_tab_size
    }
    /// Tab stops, when present (MS-PPT 2.9.23).
    #[must_use]
    pub fn tab_stops(&self) -> Option<&[ParagraphTabStop]> {
        self.tab_stops.as_deref()
    }
    /// Font alignment, when present (MS-PPT 2.13.31).
    #[must_use]
    pub const fn font_align(&self) -> Option<ParagraphFontAlignment> {
        self.font_align
    }
    /// Line-breaking settings, when present (MS-PPT 2.9.25).
    #[must_use]
    pub const fn wrap_flags(&self) -> Option<WrapFlags> {
        self.wrap_flags
    }
    /// Text direction, when present (MS-PPT 2.13.30).
    #[must_use]
    pub const fn text_direction(&self) -> Option<ParagraphTextDirection> {
        self.text_direction
    }
}
