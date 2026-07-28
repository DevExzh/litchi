//! `TextCFExceptionAtom` and `TextPFExceptionAtom` document-level text
//! formatting defaults (MS-PPT 2.9.13/2.9.14 and 2.9.19/2.9.20).
//!
//! Both atoms live in the `DocumentTextInfoContainer` (MS-PPT 2.9.1) and
//! specify the default character-level and paragraph-level formatting for
//! text in the document. They are inert: font references are never resolved
//! and formatting is never applied.

use super::records::record::PptRecord;
use super::slide_show_settings::PowerPointColorIndex;
use super::text_run::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphTabAlignment, ParagraphTabStop,
    ParagraphTextDirection,
};
use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};

/// `RT_TextCharFormatExceptionAtom` record type (MS-PPT 2.9.13).
const TEXT_CF_EXCEPTION_TYPE: u16 = 0x0FA4;
/// `RT_TextParagraphFormatExceptionAtom` record type (MS-PPT 2.9.19).
const TEXT_PF_EXCEPTION_TYPE: u16 = 0x0FA5;

// CFMasks bits (MS-PPT 2.9.15).
const CF_MASK_BOLD: u32 = 0x0000_0001;
const CF_MASK_ITALIC: u32 = 0x0000_0002;
const CF_MASK_UNDERLINE: u32 = 0x0000_0004;
const CF_MASK_SHADOW: u32 = 0x0000_0010;
const CF_MASK_FEHINT: u32 = 0x0000_0020;
const CF_MASK_KUMI: u32 = 0x0000_0080;
const CF_MASK_EMBOSS: u32 = 0x0000_0200;
const CF_MASK_HAS_STYLE: u32 = 0x0000_3C00;
const CF_MASK_TYPEFACE: u32 = 0x0001_0000;
const CF_MASK_SIZE: u32 = 0x0002_0000;
const CF_MASK_COLOR: u32 = 0x0004_0000;
const CF_MASK_POSITION: u32 = 0x0008_0000;
const CF_MASK_OLD_EA_TYPEFACE: u32 = 0x0020_0000;
const CF_MASK_ANSI_TYPEFACE: u32 = 0x0040_0000;
const CF_MASK_SYMBOL_TYPEFACE: u32 = 0x0080_0000;
/// Mask bits that must be zero in a `TextCFException`: `pp10ext`,
/// `newEATypeface`, `csTypeface`, `pp11ext`, and `reserved` (MS-PPT 2.9.14).
/// The `unused1`..`unused4` bits are undefined and ignored.
const CF_MASK_FORBIDDEN: u32 = 0xF800_0000
    | 0x0400_0000 // pp11ext
    | 0x0200_0000 // csTypeface
    | 0x0100_0000 // newEATypeface
    | 0x0010_0000; // pp10ext
/// Mask bits that gate the presence of the `fontStyle` field (MS-PPT 2.9.14).
const CF_MASK_STYLE_PRESENCE: u32 = CF_MASK_BOLD
    | CF_MASK_ITALIC
    | CF_MASK_UNDERLINE
    | CF_MASK_SHADOW
    | CF_MASK_FEHINT
    | CF_MASK_KUMI
    | CF_MASK_EMBOSS
    | CF_MASK_HAS_STYLE;

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

/// Smallest valid `fontSize` value, in points (MS-PPT 2.9.14).
const MIN_FONT_SIZE: i16 = 1;
/// Largest valid `fontSize` value, in points (MS-PPT 2.9.14).
const MAX_FONT_SIZE: i16 = 4000;
/// Largest absolute `position` percentage (MS-PPT 2.9.14).
const MAX_POSITION_PERCENT: i16 = 100;

// PFMasks bits (MS-PPT 2.9.21).
const PF_MASK_HAS_BULLET: u32 = 0x0000_0001;
const PF_MASK_BULLET_HAS_FONT: u32 = 0x0000_0002;
const PF_MASK_BULLET_HAS_COLOR: u32 = 0x0000_0004;
const PF_MASK_BULLET_HAS_SIZE: u32 = 0x0000_0008;
const PF_MASK_BULLET_FONT: u32 = 0x0000_0010;
const PF_MASK_BULLET_COLOR: u32 = 0x0000_0020;
const PF_MASK_BULLET_SIZE: u32 = 0x0000_0040;
const PF_MASK_BULLET_CHAR: u32 = 0x0000_0080;
const PF_MASK_LEFT_MARGIN: u32 = 0x0000_0100;
const PF_MASK_INDENT: u32 = 0x0000_0400;
const PF_MASK_ALIGN: u32 = 0x0000_0800;
const PF_MASK_LINE_SPACING: u32 = 0x0000_1000;
const PF_MASK_SPACE_BEFORE: u32 = 0x0000_2000;
const PF_MASK_SPACE_AFTER: u32 = 0x0000_4000;
const PF_MASK_DEFAULT_TAB_SIZE: u32 = 0x0000_8000;
const PF_MASK_FONT_ALIGN: u32 = 0x0001_0000;
const PF_MASK_CHAR_WRAP: u32 = 0x0002_0000;
const PF_MASK_WORD_WRAP: u32 = 0x0004_0000;
const PF_MASK_OVERFLOW: u32 = 0x0008_0000;
const PF_MASK_TAB_STOPS: u32 = 0x0010_0000;
const PF_MASK_TEXT_DIRECTION: u32 = 0x0020_0000;
/// Mask bits that must be zero in a `TextPFException`: `reserved1`,
/// `bulletBlip`, `bulletScheme`, `bulletHasScheme`, and `reserved2`
/// (MS-PPT 2.9.20). The `unused` bit is undefined and ignored.
const PF_MASK_FORBIDDEN: u32 = 0xFFC0_0000;
/// Mask bits that gate the presence of the `bulletFlags` field (MS-PPT 2.9.20).
const PF_MASK_BULLET_FLAGS_PRESENCE: u32 = PF_MASK_HAS_BULLET
    | PF_MASK_BULLET_HAS_FONT
    | PF_MASK_BULLET_HAS_COLOR
    | PF_MASK_BULLET_HAS_SIZE;
/// Mask bits that gate the presence of the `wrapFlags` field (MS-PPT 2.9.20).
const PF_MASK_WRAP_FLAGS_PRESENCE: u32 =
    PF_MASK_CHAR_WRAP | PF_MASK_WORD_WRAP | PF_MASK_OVERFLOW;

// BulletFlags bits (MS-PPT 2.9.22).
const BULLET_HAS_BULLET: u16 = 0x0001;
const BULLET_HAS_FONT: u16 = 0x0002;
const BULLET_HAS_COLOR: u16 = 0x0004;
const BULLET_HAS_SIZE: u16 = 0x0008;
/// `BulletFlags.reserved` bits: must be zero.
const BULLET_FLAGS_RESERVED: u16 = 0xFFF0;

// PFWrapFlags bits (MS-PPT 2.9.25).
const WRAP_CHAR_WRAP: u16 = 0x0001;
const WRAP_WORD_WRAP: u16 = 0x0002;
const WRAP_OVERFLOW: u16 = 0x0004;
/// `PFWrapFlags.reserved` bits: must be zero.
const WRAP_FLAGS_RESERVED: u16 = 0xFFF8;

/// Size in bytes of one `TabStop` (MS-PPT 2.9.24).
const TAB_STOP_LEN: usize = 4;
/// Largest `ParaSpacing` percentage; larger positive values are invalid
/// (MS-PPT 2.2.20).
const MAX_PARA_SPACING_PERCENT: i16 = 13200;

fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().expect("length checked"))
}

fn read_i16(data: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes(data[offset..offset + 2].try_into().expect("length checked"))
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

fn require_bytes(data: &[u8], offset: usize, needed: usize, field: &str) -> Result<()> {
    if data.len() < offset.saturating_add(needed) {
        return Err(corrupted(format!("{field} is truncated")));
    }
    Ok(())
}

fn parse_alignment(value: u16) -> Result<ParagraphAlignment> {
    match value {
        0 => Ok(ParagraphAlignment::Left),
        1 => Ok(ParagraphAlignment::Center),
        2 => Ok(ParagraphAlignment::Right),
        3 => Ok(ParagraphAlignment::Justify),
        4 => Ok(ParagraphAlignment::Distributed),
        5 => Ok(ParagraphAlignment::ThaiDistributed),
        6 => Ok(ParagraphAlignment::JustifyLow),
        _ => Err(corrupted("invalid TextAlignmentEnum value")),
    }
}

const fn alignment_raw(alignment: ParagraphAlignment) -> u16 {
    match alignment {
        ParagraphAlignment::Left => 0,
        ParagraphAlignment::Center => 1,
        ParagraphAlignment::Right => 2,
        ParagraphAlignment::Justify => 3,
        ParagraphAlignment::Distributed => 4,
        ParagraphAlignment::ThaiDistributed => 5,
        ParagraphAlignment::JustifyLow => 6,
    }
}

fn parse_font_alignment(value: u16) -> Result<ParagraphFontAlignment> {
    match value {
        0 => Ok(ParagraphFontAlignment::Roman),
        1 => Ok(ParagraphFontAlignment::Hanging),
        2 => Ok(ParagraphFontAlignment::Center),
        3 => Ok(ParagraphFontAlignment::UpholdFixed),
        _ => Err(corrupted("invalid TextFontAlignmentEnum value")),
    }
}

const fn font_alignment_raw(alignment: ParagraphFontAlignment) -> u16 {
    match alignment {
        ParagraphFontAlignment::Roman => 0,
        ParagraphFontAlignment::Hanging => 1,
        ParagraphFontAlignment::Center => 2,
        ParagraphFontAlignment::UpholdFixed => 3,
    }
}

fn parse_text_direction(value: u16) -> Result<ParagraphTextDirection> {
    match value {
        0 => Ok(ParagraphTextDirection::LeftToRight),
        1 => Ok(ParagraphTextDirection::RightToLeft),
        _ => Err(corrupted("invalid TextDirectionEnum value")),
    }
}

const fn text_direction_raw(direction: ParagraphTextDirection) -> u16 {
    match direction {
        ParagraphTextDirection::LeftToRight => 0,
        ParagraphTextDirection::RightToLeft => 1,
    }
}

fn parse_tab_alignment(value: u16) -> Result<ParagraphTabAlignment> {
    match value {
        0 => Ok(ParagraphTabAlignment::Left),
        1 => Ok(ParagraphTabAlignment::Center),
        2 => Ok(ParagraphTabAlignment::Right),
        3 => Ok(ParagraphTabAlignment::Decimal),
        _ => Err(corrupted("invalid TextTabTypeEnum value")),
    }
}

const fn tab_alignment_raw(alignment: ParagraphTabAlignment) -> u16 {
    match alignment {
        ParagraphTabAlignment::Left => 0,
        ParagraphTabAlignment::Center => 1,
        ParagraphTabAlignment::Right => 2,
        ParagraphTabAlignment::Decimal => 3,
    }
}

/// A `CFStyle` character-style bitfield (MS-PPT 2.9.16). The raw bits are
/// preserved because the `unused` bits are undefined.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointCFStyle {
    bits: u16,
}

impl PowerPointCFStyle {
    /// Whether the characters are bold.
    pub const fn bold(&self) -> bool {
        self.bits & CF_STYLE_BOLD != 0
    }
    /// Whether the characters are italicized.
    pub const fn italic(&self) -> bool {
        self.bits & CF_STYLE_ITALIC != 0
    }
    /// Whether the characters are underlined.
    pub const fn underline(&self) -> bool {
        self.bits & CF_STYLE_UNDERLINE != 0
    }
    /// Whether the characters have a shadow effect.
    pub const fn shadow(&self) -> bool {
        self.bits & CF_STYLE_SHADOW != 0
    }
    /// Whether the characters originated from double-byte input.
    pub const fn fehint(&self) -> bool {
        self.bits & CF_STYLE_FEHINT != 0
    }
    /// Whether Kumimoji are used for vertical text.
    pub const fn kumi(&self) -> bool {
        self.bits & CF_STYLE_KUMI != 0
    }
    /// Whether the characters are embossed.
    pub const fn emboss(&self) -> bool {
        self.bits & CF_STYLE_EMBOSS != 0
    }
    /// The four-bit `StyleTextProp9Atom` run grouping (`pp9rt`).
    pub const fn pp9_run_group(&self) -> u8 {
        ((self.bits >> CF_STYLE_PP9RT_SHIFT) & CF_STYLE_PP9RT_MASK) as u8
    }
}

/// A parsed `TextCFException` structure (MS-PPT 2.9.14) with character-level
/// style and formatting defaults. The `CFMasks` value is preserved verbatim
/// so that undefined `unused` bits round-trip exactly.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointTextCFException {
    masks: u32,
    font_style: Option<PowerPointCFStyle>,
    font_ref: Option<u16>,
    old_east_asian_font_ref: Option<u16>,
    ansi_font_ref: Option<u16>,
    symbol_font_ref: Option<u16>,
    font_size: Option<i16>,
    color: Option<PowerPointColorIndex>,
    position: Option<i16>,
}

impl PowerPointTextCFException {
    /// The raw `CFMasks` value (MS-PPT 2.9.15).
    pub const fn masks(&self) -> u32 {
        self.masks
    }
    /// The `CFStyle` character style, when present.
    pub const fn font_style(&self) -> Option<PowerPointCFStyle> {
        self.font_style
    }
    /// Zero-based index of the font in the `FontCollectionContainer`,
    /// when present (MS-PPT 2.2.10).
    pub const fn font_ref(&self) -> Option<u16> {
        self.font_ref
    }
    /// Zero-based index of an East Asian font, when present.
    pub const fn old_east_asian_font_ref(&self) -> Option<u16> {
        self.old_east_asian_font_ref
    }
    /// Zero-based index of an ANSI font, when present.
    pub const fn ansi_font_ref(&self) -> Option<u16> {
        self.ansi_font_ref
    }
    /// Zero-based index of a symbol font, when present.
    pub const fn symbol_font_ref(&self) -> Option<u16> {
        self.symbol_font_ref
    }
    /// Font size in points, when present; within 1..=4000 (MS-PPT 2.9.14).
    pub const fn font_size(&self) -> Option<i16> {
        self.font_size
    }
    /// Text color, when present (MS-PPT 2.12.2).
    pub const fn color(&self) -> Option<PowerPointColorIndex> {
        self.color
    }
    /// Baseline position as a percentage of line height, when present;
    /// within -100..=100 (MS-PPT 2.9.14).
    pub const fn position(&self) -> Option<i16> {
        self.position
    }

    /// Parse one `TextCFException` from the start of `data`.
    ///
    /// Returns the decoded structure and the number of bytes consumed.
    pub(crate) fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "TextCFException masks")?;
        let masks = read_u32(data, 0);
        if masks & CF_MASK_FORBIDDEN != 0 {
            return Err(corrupted("TextCFException has forbidden mask bits set"));
        }
        let mut offset = 4usize;

        let font_style = if masks & CF_MASK_STYLE_PRESENCE != 0 {
            require_bytes(data, offset, 2, "TextCFException fontStyle")?;
            let style = PowerPointCFStyle {
                bits: read_u16(data, offset),
            };
            offset += 2;
            Some(style)
        } else {
            None
        };
        let font_ref = if masks & CF_MASK_TYPEFACE != 0 {
            require_bytes(data, offset, 2, "TextCFException fontRef")?;
            let value = read_u16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let old_east_asian_font_ref = if masks & CF_MASK_OLD_EA_TYPEFACE != 0 {
            require_bytes(data, offset, 2, "TextCFException oldEAFontRef")?;
            let value = read_u16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let ansi_font_ref = if masks & CF_MASK_ANSI_TYPEFACE != 0 {
            require_bytes(data, offset, 2, "TextCFException ansiFontRef")?;
            let value = read_u16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let symbol_font_ref = if masks & CF_MASK_SYMBOL_TYPEFACE != 0 {
            require_bytes(data, offset, 2, "TextCFException symbolFontRef")?;
            let value = read_u16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let font_size = if masks & CF_MASK_SIZE != 0 {
            require_bytes(data, offset, 2, "TextCFException fontSize")?;
            let value = read_i16(data, offset);
            offset += 2;
            if !(MIN_FONT_SIZE..=MAX_FONT_SIZE).contains(&value) {
                return Err(corrupted("TextCFException fontSize is out of range"));
            }
            Some(value)
        } else {
            None
        };
        let color = if masks & CF_MASK_COLOR != 0 {
            require_bytes(data, offset, 4, "TextCFException color")?;
            let value = PowerPointColorIndex::parse_bytes(&data[offset..offset + 4])?;
            offset += 4;
            Some(value)
        } else {
            None
        };
        let position = if masks & CF_MASK_POSITION != 0 {
            require_bytes(data, offset, 2, "TextCFException position")?;
            let value = read_i16(data, offset);
            offset += 2;
            if value.abs() > MAX_POSITION_PERCENT {
                return Err(corrupted("TextCFException position is out of range"));
            }
            Some(value)
        } else {
            None
        };

        Ok((
            Self {
                masks,
                font_style,
                font_ref,
                old_east_asian_font_ref,
                ansi_font_ref,
                symbol_font_ref,
                font_size,
                color,
                position,
            },
            offset,
        ))
    }

    /// Serialize the structure without any record header.
    fn to_payload(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(16);
        data.extend_from_slice(&self.masks.to_le_bytes());
        if let Some(font_style) = self.font_style {
            data.extend_from_slice(&font_style.bits.to_le_bytes());
        }
        if let Some(font_ref) = self.font_ref {
            data.extend_from_slice(&font_ref.to_le_bytes());
        }
        if let Some(font_ref) = self.old_east_asian_font_ref {
            data.extend_from_slice(&font_ref.to_le_bytes());
        }
        if let Some(font_ref) = self.ansi_font_ref {
            data.extend_from_slice(&font_ref.to_le_bytes());
        }
        if let Some(font_ref) = self.symbol_font_ref {
            data.extend_from_slice(&font_ref.to_le_bytes());
        }
        if let Some(font_size) = self.font_size {
            data.extend_from_slice(&font_size.to_le_bytes());
        }
        if let Some(color) = self.color {
            data.extend_from_slice(&color.to_bytes());
        }
        if let Some(position) = self.position {
            data.extend_from_slice(&position.to_le_bytes());
        }
        data
    }

    /// Parse a complete `TextCFExceptionAtom` record (MS-PPT 2.9.13).
    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::TxCFStyleAtom
            || record.record_type_raw != TEXT_CF_EXCEPTION_TYPE
            || record.version != 0
            || record.instance != 0
        {
            return Err(corrupted("TextCFExceptionAtom has an invalid header"));
        }
        Self::parse(&record.data)
    }

    /// Parse the whole payload of a `TextCFExceptionAtom`.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let (exception, consumed) = Self::parse_prefix(data)?;
        if consumed != data.len() {
            return Err(corrupted("TextCFExceptionAtom has trailing bytes"));
        }
        Ok(exception)
    }

    /// Serialize the complete `TextCFExceptionAtom`, including its header.
    pub fn to_bytes(&self) -> Vec<u8> {
        let payload = self.to_payload();
        let mut data = Vec::with_capacity(8 + payload.len());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&TEXT_CF_EXCEPTION_TYPE.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(&payload);
        data
    }
}

/// `BulletFlags` bullet-property validity bits (MS-PPT 2.9.22).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointBulletFlags {
    has_bullet: bool,
    bullet_has_font: bool,
    bullet_has_color: bool,
    bullet_has_size: bool,
}

impl PowerPointBulletFlags {
    /// Whether a bullet exists.
    pub const fn has_bullet(&self) -> bool {
        self.has_bullet
    }
    /// Whether the bullet has a font.
    pub const fn bullet_has_font(&self) -> bool {
        self.bullet_has_font
    }
    /// Whether the bullet has a color.
    pub const fn bullet_has_color(&self) -> bool {
        self.bullet_has_color
    }
    /// Whether the bullet has a size.
    pub const fn bullet_has_size(&self) -> bool {
        self.bullet_has_size
    }

    fn from_bits(bits: u16) -> Result<Self> {
        if bits & BULLET_FLAGS_RESERVED != 0 {
            return Err(corrupted("BulletFlags has reserved bits set"));
        }
        Ok(Self {
            has_bullet: bits & BULLET_HAS_BULLET != 0,
            bullet_has_font: bits & BULLET_HAS_FONT != 0,
            bullet_has_color: bits & BULLET_HAS_COLOR != 0,
            bullet_has_size: bits & BULLET_HAS_SIZE != 0,
        })
    }

    fn to_bits(self) -> u16 {
        u16::from(self.has_bullet)
            | u16::from(self.bullet_has_font) << 1
            | u16::from(self.bullet_has_color) << 2
            | u16::from(self.bullet_has_size) << 3
    }
}

/// `PFWrapFlags` line-breaking settings (MS-PPT 2.9.25).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointWrapFlags {
    char_wrap: bool,
    word_wrap: bool,
    overflow: bool,
}

impl PowerPointWrapFlags {
    /// Whether the paragraph follows the East Asian kinsoku line-breaking
    /// settings.
    pub const fn char_wrap(&self) -> bool {
        self.char_wrap
    }
    /// Whether text wraps only at word breaks.
    pub const fn word_wrap(&self) -> bool {
        self.word_wrap
    }
    /// Whether hanging punctuation is allowed for East Asian text.
    pub const fn overflow(&self) -> bool {
        self.overflow
    }

    fn from_bits(bits: u16) -> Result<Self> {
        if bits & WRAP_FLAGS_RESERVED != 0 {
            return Err(corrupted("PFWrapFlags has reserved bits set"));
        }
        Ok(Self {
            char_wrap: bits & WRAP_CHAR_WRAP != 0,
            word_wrap: bits & WRAP_WORD_WRAP != 0,
            overflow: bits & WRAP_OVERFLOW != 0,
        })
    }

    fn to_bits(self) -> u16 {
        u16::from(self.char_wrap)
            | u16::from(self.word_wrap) << 1
            | u16::from(self.overflow) << 2
    }
}

/// A parsed `TextPFException` structure (MS-PPT 2.9.20) with paragraph-level
/// formatting defaults. The `PFMasks` value is preserved verbatim so that the
/// undefined `unused` bit round-trips exactly.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointTextPFException {
    masks: u32,
    bullet_flags: Option<PowerPointBulletFlags>,
    bullet_char: Option<u16>,
    bullet_font_ref: Option<u16>,
    bullet_size: Option<i16>,
    bullet_color: Option<PowerPointColorIndex>,
    text_alignment: Option<ParagraphAlignment>,
    line_spacing: Option<i16>,
    space_before: Option<i16>,
    space_after: Option<i16>,
    left_margin: Option<i16>,
    indent: Option<i16>,
    default_tab_size: Option<i16>,
    tab_stops: Option<Vec<ParagraphTabStop>>,
    font_align: Option<ParagraphFontAlignment>,
    wrap_flags: Option<PowerPointWrapFlags>,
    text_direction: Option<ParagraphTextDirection>,
}

impl PowerPointTextPFException {
    /// The raw `PFMasks` value (MS-PPT 2.9.21).
    pub const fn masks(&self) -> u32 {
        self.masks
    }
    /// Bullet validity flags, when present.
    pub const fn bullet_flags(&self) -> Option<PowerPointBulletFlags> {
        self.bullet_flags
    }
    /// UTF-16 code unit displayed as the bullet, when present; never NUL
    /// (MS-PPT 2.9.20).
    pub const fn bullet_char(&self) -> Option<u16> {
        self.bullet_char
    }
    /// Zero-based index of the bullet font, when present.
    pub const fn bullet_font_ref(&self) -> Option<u16> {
        self.bullet_font_ref
    }
    /// `BulletSize` of the bullet, when present (MS-PPT 2.2.3).
    pub const fn bullet_size(&self) -> Option<i16> {
        self.bullet_size
    }
    /// Bullet color, when present (MS-PPT 2.12.2).
    pub const fn bullet_color(&self) -> Option<PowerPointColorIndex> {
        self.bullet_color
    }
    /// Paragraph alignment, when present (MS-PPT 2.13.27).
    pub const fn text_alignment(&self) -> Option<ParagraphAlignment> {
        self.text_alignment
    }
    /// `ParaSpacing` between lines, when present (MS-PPT 2.2.20).
    pub const fn line_spacing(&self) -> Option<i16> {
        self.line_spacing
    }
    /// `ParaSpacing` before the paragraph, when present.
    pub const fn space_before(&self) -> Option<i16> {
        self.space_before
    }
    /// `ParaSpacing` after the paragraph, when present.
    pub const fn space_after(&self) -> Option<i16> {
        self.space_after
    }
    /// Left margin in master units, when present (MS-PPT 2.2.15).
    pub const fn left_margin(&self) -> Option<i16> {
        self.left_margin
    }
    /// Paragraph indentation in master units, when present.
    pub const fn indent(&self) -> Option<i16> {
        self.indent
    }
    /// Default tab size in master units, when present (MS-PPT 2.2.29).
    pub const fn default_tab_size(&self) -> Option<i16> {
        self.default_tab_size
    }
    /// Tab stops, when present (MS-PPT 2.9.23).
    pub fn tab_stops(&self) -> Option<&[ParagraphTabStop]> {
        self.tab_stops.as_deref()
    }
    /// Font alignment, when present (MS-PPT 2.13.31).
    pub const fn font_align(&self) -> Option<ParagraphFontAlignment> {
        self.font_align
    }
    /// Line-breaking settings, when present (MS-PPT 2.9.25).
    pub const fn wrap_flags(&self) -> Option<PowerPointWrapFlags> {
        self.wrap_flags
    }
    /// Text direction, when present (MS-PPT 2.13.30).
    pub const fn text_direction(&self) -> Option<ParagraphTextDirection> {
        self.text_direction
    }

    /// Parse one `TextPFException` from the start of `data`.
    ///
    /// Returns the decoded structure and the number of bytes consumed.
    pub(crate) fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "TextPFException masks")?;
        let masks = read_u32(data, 0);
        if masks & PF_MASK_FORBIDDEN != 0 {
            return Err(corrupted("TextPFException has forbidden mask bits set"));
        }
        let mut offset = 4usize;

        let bullet_flags = if masks & PF_MASK_BULLET_FLAGS_PRESENCE != 0 {
            require_bytes(data, offset, 2, "TextPFException bulletFlags")?;
            let flags = PowerPointBulletFlags::from_bits(read_u16(data, offset))?;
            offset += 2;
            Some(flags)
        } else {
            None
        };
        let bullet_char = if masks & PF_MASK_BULLET_CHAR != 0 {
            require_bytes(data, offset, 2, "TextPFException bulletChar")?;
            let value = read_u16(data, offset);
            offset += 2;
            if value == 0 {
                return Err(corrupted("TextPFException bulletChar must not be NUL"));
            }
            Some(value)
        } else {
            None
        };
        let bullet_font_ref = if masks & PF_MASK_BULLET_FONT != 0 {
            require_bytes(data, offset, 2, "TextPFException bulletFontRef")?;
            let value = read_u16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let bullet_size = if masks & PF_MASK_BULLET_SIZE != 0 {
            require_bytes(data, offset, 2, "TextPFException bulletSize")?;
            let value = read_i16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let bullet_color = if masks & PF_MASK_BULLET_COLOR != 0 {
            require_bytes(data, offset, 4, "TextPFException bulletColor")?;
            let value = PowerPointColorIndex::parse_bytes(&data[offset..offset + 4])?;
            offset += 4;
            Some(value)
        } else {
            None
        };
        let text_alignment = if masks & PF_MASK_ALIGN != 0 {
            require_bytes(data, offset, 2, "TextPFException textAlignment")?;
            let value = parse_alignment(read_u16(data, offset))?;
            offset += 2;
            Some(value)
        } else {
            None
        };
        let line_spacing = if masks & PF_MASK_LINE_SPACING != 0 {
            let value = parse_para_spacing(data, &mut offset, "TextPFException lineSpacing")?;
            Some(value)
        } else {
            None
        };
        let space_before = if masks & PF_MASK_SPACE_BEFORE != 0 {
            let value = parse_para_spacing(data, &mut offset, "TextPFException spaceBefore")?;
            Some(value)
        } else {
            None
        };
        let space_after = if masks & PF_MASK_SPACE_AFTER != 0 {
            let value = parse_para_spacing(data, &mut offset, "TextPFException spaceAfter")?;
            Some(value)
        } else {
            None
        };
        let left_margin = if masks & PF_MASK_LEFT_MARGIN != 0 {
            require_bytes(data, offset, 2, "TextPFException leftMargin")?;
            let value = read_i16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let indent = if masks & PF_MASK_INDENT != 0 {
            require_bytes(data, offset, 2, "TextPFException indent")?;
            let value = read_i16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let default_tab_size = if masks & PF_MASK_DEFAULT_TAB_SIZE != 0 {
            require_bytes(data, offset, 2, "TextPFException defaultTabSize")?;
            let value = read_i16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let tab_stops = if masks & PF_MASK_TAB_STOPS != 0 {
            require_bytes(data, offset, 2, "TextPFException tabStops count")?;
            let count = usize::from(read_u16(data, offset));
            offset += 2;
            let remaining = (data.len() - offset) / TAB_STOP_LEN;
            if count > remaining {
                return Err(corrupted("TextPFException tabStops are truncated"));
            }
            let mut stops = Vec::with_capacity(count);
            for _ in 0..count {
                let position = read_i16(data, offset);
                let alignment = parse_tab_alignment(read_u16(data, offset + 2))?;
                offset += TAB_STOP_LEN;
                stops.push(ParagraphTabStop {
                    position,
                    alignment,
                });
            }
            Some(stops)
        } else {
            None
        };
        let font_align = if masks & PF_MASK_FONT_ALIGN != 0 {
            require_bytes(data, offset, 2, "TextPFException fontAlign")?;
            let value = parse_font_alignment(read_u16(data, offset))?;
            offset += 2;
            Some(value)
        } else {
            None
        };
        let wrap_flags = if masks & PF_MASK_WRAP_FLAGS_PRESENCE != 0 {
            require_bytes(data, offset, 2, "TextPFException wrapFlags")?;
            let flags = PowerPointWrapFlags::from_bits(read_u16(data, offset))?;
            offset += 2;
            Some(flags)
        } else {
            None
        };
        let text_direction = if masks & PF_MASK_TEXT_DIRECTION != 0 {
            require_bytes(data, offset, 2, "TextPFException textDirection")?;
            let value = parse_text_direction(read_u16(data, offset))?;
            offset += 2;
            Some(value)
        } else {
            None
        };

        Ok((
            Self {
                masks,
                bullet_flags,
                bullet_char,
                bullet_font_ref,
                bullet_size,
                bullet_color,
                text_alignment,
                line_spacing,
                space_before,
                space_after,
                left_margin,
                indent,
                default_tab_size,
                tab_stops,
                font_align,
                wrap_flags,
                text_direction,
            },
            offset,
        ))
    }

    /// Serialize the structure without any record header or reserved field.
    fn to_payload(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(16);
        data.extend_from_slice(&self.masks.to_le_bytes());
        if let Some(bullet_flags) = self.bullet_flags {
            data.extend_from_slice(&bullet_flags.to_bits().to_le_bytes());
        }
        if let Some(bullet_char) = self.bullet_char {
            data.extend_from_slice(&bullet_char.to_le_bytes());
        }
        if let Some(bullet_font_ref) = self.bullet_font_ref {
            data.extend_from_slice(&bullet_font_ref.to_le_bytes());
        }
        if let Some(bullet_size) = self.bullet_size {
            data.extend_from_slice(&bullet_size.to_le_bytes());
        }
        if let Some(bullet_color) = self.bullet_color {
            data.extend_from_slice(&bullet_color.to_bytes());
        }
        if let Some(text_alignment) = self.text_alignment {
            data.extend_from_slice(&alignment_raw(text_alignment).to_le_bytes());
        }
        if let Some(line_spacing) = self.line_spacing {
            data.extend_from_slice(&line_spacing.to_le_bytes());
        }
        if let Some(space_before) = self.space_before {
            data.extend_from_slice(&space_before.to_le_bytes());
        }
        if let Some(space_after) = self.space_after {
            data.extend_from_slice(&space_after.to_le_bytes());
        }
        if let Some(left_margin) = self.left_margin {
            data.extend_from_slice(&left_margin.to_le_bytes());
        }
        if let Some(indent) = self.indent {
            data.extend_from_slice(&indent.to_le_bytes());
        }
        if let Some(default_tab_size) = self.default_tab_size {
            data.extend_from_slice(&default_tab_size.to_le_bytes());
        }
        if let Some(tab_stops) = &self.tab_stops {
            data.extend_from_slice(&(tab_stops.len() as u16).to_le_bytes());
            for stop in tab_stops {
                data.extend_from_slice(&stop.position.to_le_bytes());
                data.extend_from_slice(&tab_alignment_raw(stop.alignment).to_le_bytes());
            }
        }
        if let Some(font_align) = self.font_align {
            data.extend_from_slice(&font_alignment_raw(font_align).to_le_bytes());
        }
        if let Some(wrap_flags) = self.wrap_flags {
            data.extend_from_slice(&wrap_flags.to_bits().to_le_bytes());
        }
        if let Some(text_direction) = self.text_direction {
            data.extend_from_slice(&text_direction_raw(text_direction).to_le_bytes());
        }
        data
    }

    /// Parse a complete `TextPFExceptionAtom` record (MS-PPT 2.9.19).
    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::TxPFStyleAtom
            || record.record_type_raw != TEXT_PF_EXCEPTION_TYPE
            || record.version != 0
            || record.instance != 0
        {
            return Err(corrupted("TextPFExceptionAtom has an invalid header"));
        }
        Self::parse(&record.data)
    }

    /// Parse the whole payload of a `TextPFExceptionAtom`, including its
    /// two reserved leading bytes, which must be zero (MS-PPT 2.9.19).
    pub fn parse(data: &[u8]) -> Result<Self> {
        require_bytes(data, 0, 2, "TextPFExceptionAtom reserved")?;
        if read_u16(data, 0) != 0 {
            return Err(corrupted("TextPFExceptionAtom reserved bytes must be zero"));
        }
        let (exception, consumed) = Self::parse_prefix(&data[2..])?;
        if consumed != data.len() - 2 {
            return Err(corrupted("TextPFExceptionAtom has trailing bytes"));
        }
        Ok(exception)
    }

    /// Serialize the complete `TextPFExceptionAtom`, including its header.
    pub fn to_bytes(&self) -> Vec<u8> {
        let payload = self.to_payload();
        let mut data = Vec::with_capacity(8 + 2 + payload.len());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&TEXT_PF_EXCEPTION_TYPE.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32 + 2).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&payload);
        data
    }
}

fn parse_para_spacing(data: &[u8], offset: &mut usize, field: &str) -> Result<i16> {
    require_bytes(data, *offset, 2, field)?;
    let value = read_i16(data, *offset);
    *offset += 2;
    if value > MAX_PARA_SPACING_PERCENT {
        return Err(corrupted(format!("{field} is out of range")));
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::slide_show_settings::PowerPointColorIndexKind;

    fn cf_record(data: &[u8]) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::TxCFStyleAtom,
            record_type_raw: TEXT_CF_EXCEPTION_TYPE,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    fn pf_record(data: &[u8]) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::TxPFStyleAtom,
            record_type_raw: TEXT_PF_EXCEPTION_TYPE,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    /// Mask: bold | italic | emboss | fHasStyle(3) | typeface | ansiTypeface
    /// | size | color | position.
    fn sample_cf_payload() -> Vec<u8> {
        let masks = CF_MASK_BOLD
            | CF_MASK_ITALIC
            | CF_MASK_EMBOSS
            | (3 << 10)
            | CF_MASK_TYPEFACE
            | CF_MASK_ANSI_TYPEFACE
            | CF_MASK_SIZE
            | CF_MASK_COLOR
            | CF_MASK_POSITION;
        let mut data = Vec::new();
        data.extend_from_slice(&masks.to_le_bytes());
        // fontStyle: bold + italic set, pp9rt = 3.
        data.extend_from_slice(&0x0C03u16.to_le_bytes());
        data.extend_from_slice(&7u16.to_le_bytes()); // fontRef
        data.extend_from_slice(&11u16.to_le_bytes()); // ansiFontRef
        data.extend_from_slice(&2400i16.to_le_bytes()); // fontSize
        data.extend_from_slice(&[0x12, 0x34, 0x56, 0xFE]); // sRGB color
        data.extend_from_slice(&(-30i16).to_le_bytes()); // position
        data
    }

    #[test]
    fn parses_cf_exception_and_round_trips() {
        let payload = sample_cf_payload();
        let parsed = PowerPointTextCFException::parse_record(&cf_record(&payload)).unwrap();
        assert_eq!(
            parsed.masks(),
            CF_MASK_BOLD
                | CF_MASK_ITALIC
                | CF_MASK_EMBOSS
                | (3 << 10)
                | CF_MASK_TYPEFACE
                | CF_MASK_ANSI_TYPEFACE
                | CF_MASK_SIZE
                | CF_MASK_COLOR
                | CF_MASK_POSITION
        );
        let style = parsed.font_style().unwrap();
        assert!(style.bold());
        assert!(style.italic());
        assert!(!style.underline());
        assert!(!style.emboss());
        assert_eq!(style.pp9_run_group(), 3);
        assert_eq!(parsed.font_ref(), Some(7));
        assert_eq!(parsed.old_east_asian_font_ref(), None);
        assert_eq!(parsed.ansi_font_ref(), Some(11));
        assert_eq!(parsed.symbol_font_ref(), None);
        assert_eq!(parsed.font_size(), Some(2400));
        let color = parsed.color().unwrap();
        assert_eq!(color.red, 0x12);
        assert_eq!(color.green, 0x34);
        assert_eq!(color.blue, 0x56);
        assert_eq!(color.kind, PowerPointColorIndexKind::Srgb);
        assert_eq!(parsed.position(), Some(-30));

        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn parses_empty_cf_exception_and_round_trips() {
        let payload = 0u32.to_le_bytes().to_vec();
        let parsed = PowerPointTextCFException::parse_record(&cf_record(&payload)).unwrap();
        assert_eq!(parsed.masks(), 0);
        assert_eq!(parsed.font_style(), None);
        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn cf_style_presence_follows_fhas_style_only() {
        // Only fHasStyle = 2 set: fontStyle exists, pp9rt read from CFStyle.
        let mut payload = Vec::new();
        payload.extend_from_slice(&(2u32 << 10).to_le_bytes());
        payload.extend_from_slice(&0x1400u16.to_le_bytes()); // pp9rt = 5
        let parsed = PowerPointTextCFException::parse_record(&cf_record(&payload)).unwrap();
        let style = parsed.font_style().unwrap();
        assert!(!style.bold());
        assert_eq!(style.pp9_run_group(), 5);
        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn rejects_malformed_cf_exception() {
        // Wrong record type.
        assert!(PowerPointTextCFException::parse_record(&pf_record(&[])).is_err());
        // Truncated masks.
        assert!(PowerPointTextCFException::parse_record(&cf_record(&[1, 0])).is_err());
        // Forbidden pp10ext mask bit.
        assert!(
            PowerPointTextCFException::parse_record(&cf_record(&0x0010_0000u32.to_le_bytes()))
                .is_err()
        );
        // Forbidden reserved mask bits.
        assert!(
            PowerPointTextCFException::parse_record(&cf_record(&0x8000_0000u32.to_le_bytes()))
                .is_err()
        );
        // Truncated fontStyle.
        assert!(
            PowerPointTextCFException::parse_record(&cf_record(&CF_MASK_BOLD.to_le_bytes()))
                .is_err()
        );
        // fontSize below the minimum.
        let mut payload = CF_MASK_SIZE.to_le_bytes().to_vec();
        payload.extend_from_slice(&0i16.to_le_bytes());
        assert!(PowerPointTextCFException::parse_record(&cf_record(&payload)).is_err());
        // fontSize above the maximum.
        let mut payload = CF_MASK_SIZE.to_le_bytes().to_vec();
        payload.extend_from_slice(&4001i16.to_le_bytes());
        assert!(PowerPointTextCFException::parse_record(&cf_record(&payload)).is_err());
        // position above the maximum.
        let mut payload = CF_MASK_POSITION.to_le_bytes().to_vec();
        payload.extend_from_slice(&101i16.to_le_bytes());
        assert!(PowerPointTextCFException::parse_record(&cf_record(&payload)).is_err());
        // Invalid color index byte.
        let mut payload = CF_MASK_COLOR.to_le_bytes().to_vec();
        payload.extend_from_slice(&[0, 0, 0, 0x08]);
        assert!(PowerPointTextCFException::parse_record(&cf_record(&payload)).is_err());
        // Trailing bytes after a complete structure.
        let mut payload = 0u32.to_le_bytes().to_vec();
        payload.push(0);
        assert!(PowerPointTextCFException::parse_record(&cf_record(&payload)).is_err());
        // Nonzero instance.
        let mut record = cf_record(&0u32.to_le_bytes());
        record.instance = 1;
        assert!(PowerPointTextCFException::parse_record(&record).is_err());
    }

    /// Mask: hasBullet | bulletHasFont | bulletChar | bulletFont | bulletSize
    /// | align | lineSpacing | leftMargin | defaultTabSize | tabStops
    /// | fontAlign | charWrap | wordWrap | textDirection.
    fn sample_pf_payload() -> Vec<u8> {
        let masks = PF_MASK_HAS_BULLET
            | PF_MASK_BULLET_HAS_FONT
            | PF_MASK_BULLET_CHAR
            | PF_MASK_BULLET_FONT
            | PF_MASK_BULLET_SIZE
            | PF_MASK_ALIGN
            | PF_MASK_LINE_SPACING
            | PF_MASK_LEFT_MARGIN
            | PF_MASK_DEFAULT_TAB_SIZE
            | PF_MASK_TAB_STOPS
            | PF_MASK_FONT_ALIGN
            | PF_MASK_CHAR_WRAP
            | PF_MASK_WORD_WRAP
            | PF_MASK_TEXT_DIRECTION;
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // reserved
        data.extend_from_slice(&masks.to_le_bytes());
        data.extend_from_slice(&0x0003u16.to_le_bytes()); // bulletFlags
        data.extend_from_slice(&0x2022u16.to_le_bytes()); // bulletChar U+2022
        data.extend_from_slice(&2u16.to_le_bytes()); // bulletFontRef
        data.extend_from_slice(&(-1200i16).to_le_bytes()); // bulletSize, points
        data.extend_from_slice(&1u16.to_le_bytes()); // align center
        data.extend_from_slice(&150i16.to_le_bytes()); // lineSpacing, percent
        data.extend_from_slice(&288i16.to_le_bytes()); // leftMargin
        data.extend_from_slice(&720i16.to_le_bytes()); // defaultTabSize
        data.extend_from_slice(&2u16.to_le_bytes()); // two tab stops
        data.extend_from_slice(&100i16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes()); // left tab
        data.extend_from_slice(&(-40i16).to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes()); // decimal tab
        data.extend_from_slice(&3u16.to_le_bytes()); // fontAlign upholdFixed
        data.extend_from_slice(&0x0003u16.to_le_bytes()); // wrapFlags
        data.extend_from_slice(&1u16.to_le_bytes()); // right-to-left
        data
    }

    #[test]
    fn parses_pf_exception_and_round_trips() {
        let payload = sample_pf_payload();
        let parsed = PowerPointTextPFException::parse_record(&pf_record(&payload)).unwrap();

        let flags = parsed.bullet_flags().unwrap();
        assert!(flags.has_bullet());
        assert!(flags.bullet_has_font());
        assert!(!flags.bullet_has_color());
        assert!(!flags.bullet_has_size());
        assert_eq!(parsed.bullet_char(), Some(0x2022));
        assert_eq!(parsed.bullet_font_ref(), Some(2));
        assert_eq!(parsed.bullet_size(), Some(-1200));
        assert_eq!(parsed.bullet_color(), None);
        assert_eq!(parsed.text_alignment(), Some(ParagraphAlignment::Center));
        assert_eq!(parsed.line_spacing(), Some(150));
        assert_eq!(parsed.space_before(), None);
        assert_eq!(parsed.left_margin(), Some(288));
        assert_eq!(parsed.indent(), None);
        assert_eq!(parsed.default_tab_size(), Some(720));
        let stops = parsed.tab_stops().unwrap();
        assert_eq!(stops.len(), 2);
        assert_eq!(stops[0].position, 100);
        assert_eq!(stops[0].alignment, ParagraphTabAlignment::Left);
        assert_eq!(stops[1].position, -40);
        assert_eq!(stops[1].alignment, ParagraphTabAlignment::Decimal);
        assert_eq!(parsed.font_align(), Some(ParagraphFontAlignment::UpholdFixed));
        let wrap = parsed.wrap_flags().unwrap();
        assert!(wrap.char_wrap());
        assert!(wrap.word_wrap());
        assert!(!wrap.overflow());
        assert_eq!(
            parsed.text_direction(),
            Some(ParagraphTextDirection::RightToLeft)
        );

        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn parses_empty_pf_exception_and_round_trips() {
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&0u32.to_le_bytes());
        let parsed = PowerPointTextPFException::parse_record(&pf_record(&payload)).unwrap();
        assert_eq!(parsed.masks(), 0);
        assert_eq!(parsed.bullet_flags(), None);
        assert_eq!(parsed.tab_stops(), None);
        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn rejects_malformed_pf_exception() {
        // Wrong record type.
        assert!(PowerPointTextPFException::parse_record(&cf_record(&[])).is_err());
        // Truncated reserved field.
        assert!(PowerPointTextPFException::parse_record(&pf_record(&[0])).is_err());
        // Nonzero reserved field.
        assert!(PowerPointTextPFException::parse_record(&pf_record(&[1, 0])).is_err());
        // Forbidden bulletBlip mask bit.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&0x0080_0000u32.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // BulletFlags reserved bits.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_HAS_BULLET.to_le_bytes());
        payload.extend_from_slice(&0x0010u16.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // NUL bullet character.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_BULLET_CHAR.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // Invalid alignment value.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_ALIGN.to_le_bytes());
        payload.extend_from_slice(&7u16.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // lineSpacing above the maximum percentage.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_LINE_SPACING.to_le_bytes());
        payload.extend_from_slice(&13201i16.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // Truncated tab stops.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_TAB_STOPS.to_le_bytes());
        payload.extend_from_slice(&2u16.to_le_bytes());
        payload.extend_from_slice(&[0; 4]);
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // Invalid tab-stop type.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_TAB_STOPS.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0i16.to_le_bytes());
        payload.extend_from_slice(&4u16.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // PFWrapFlags reserved bits.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_CHAR_WRAP.to_le_bytes());
        payload.extend_from_slice(&0x0008u16.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // Invalid text direction.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&PF_MASK_TEXT_DIRECTION.to_le_bytes());
        payload.extend_from_slice(&2u16.to_le_bytes());
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // Trailing bytes after a complete structure.
        let mut payload = 0u16.to_le_bytes().to_vec();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.push(0);
        assert!(PowerPointTextPFException::parse_record(&pf_record(&payload)).is_err());
        // Nonzero version.
        let mut record = pf_record(&[0; 6]);
        record.version = 0xF;
        assert!(PowerPointTextPFException::parse_record(&record).is_err());
    }
}
