//! MS-PPT record decoding and encoding for text-format exceptions.

use super::model::{BulletFlags, CFStyle, TextCFException, TextPFException, WrapFlags};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::record::Record;
use crate::slide_show_settings::ColorIndex;
use crate::text_run::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphTabAlignment, ParagraphTabStop,
    ParagraphTextDirection,
};

/// `RT_TextCharFormatExceptionAtom` record type (MS-PPT 2.9.13).
pub(super) const TEXT_CF_EXCEPTION_TYPE: u16 = 0x0FA4;
/// `RT_TextParagraphFormatExceptionAtom` record type (MS-PPT 2.9.19).
pub(super) const TEXT_PF_EXCEPTION_TYPE: u16 = 0x0FA5;

// CFMasks bits (MS-PPT 2.9.15).
pub(super) const CF_MASK_BOLD: u32 = 0x0000_0001;
pub(super) const CF_MASK_ITALIC: u32 = 0x0000_0002;
pub(super) const CF_MASK_UNDERLINE: u32 = 0x0000_0004;
pub(super) const CF_MASK_SHADOW: u32 = 0x0000_0010;
pub(super) const CF_MASK_FEHINT: u32 = 0x0000_0020;
pub(super) const CF_MASK_KUMI: u32 = 0x0000_0080;
pub(super) const CF_MASK_EMBOSS: u32 = 0x0000_0200;
pub(super) const CF_MASK_HAS_STYLE: u32 = 0x0000_3C00;
pub(super) const CF_MASK_TYPEFACE: u32 = 0x0001_0000;
pub(super) const CF_MASK_SIZE: u32 = 0x0002_0000;
pub(super) const CF_MASK_COLOR: u32 = 0x0004_0000;
pub(super) const CF_MASK_POSITION: u32 = 0x0008_0000;
pub(super) const CF_MASK_OLD_EA_TYPEFACE: u32 = 0x0020_0000;
pub(super) const CF_MASK_ANSI_TYPEFACE: u32 = 0x0040_0000;
pub(super) const CF_MASK_SYMBOL_TYPEFACE: u32 = 0x0080_0000;
/// Mask bits that must be zero in a `TextCFException`: `pp10ext`,
/// `newEATypeface`, `csTypeface`, `pp11ext`, and `reserved` (MS-PPT 2.9.14).
/// The `unused1`..`unused4` bits are undefined and ignored.
pub(super) const CF_MASK_FORBIDDEN: u32 = 0xF800_0000
    | 0x0400_0000 // pp11ext
    | 0x0200_0000 // csTypeface
    | 0x0100_0000 // newEATypeface
    | 0x0010_0000; // pp10ext
/// Mask bits that gate the presence of the `fontStyle` field (MS-PPT 2.9.14).
pub(super) const CF_MASK_STYLE_PRESENCE: u32 = CF_MASK_BOLD
    | CF_MASK_ITALIC
    | CF_MASK_UNDERLINE
    | CF_MASK_SHADOW
    | CF_MASK_FEHINT
    | CF_MASK_KUMI
    | CF_MASK_EMBOSS
    | CF_MASK_HAS_STYLE;

/// Smallest valid `fontSize` value, in points (MS-PPT 2.9.14).
pub(super) const MIN_FONT_SIZE: i16 = 1;
/// Largest valid `fontSize` value, in points (MS-PPT 2.9.14).
pub(super) const MAX_FONT_SIZE: i16 = 4000;
/// Largest absolute `position` percentage (MS-PPT 2.9.14).
pub(super) const MAX_POSITION_PERCENT: i16 = 100;

// PFMasks bits (MS-PPT 2.9.21).
pub(super) const PF_MASK_HAS_BULLET: u32 = 0x0000_0001;
pub(super) const PF_MASK_BULLET_HAS_FONT: u32 = 0x0000_0002;
pub(super) const PF_MASK_BULLET_HAS_COLOR: u32 = 0x0000_0004;
pub(super) const PF_MASK_BULLET_HAS_SIZE: u32 = 0x0000_0008;
pub(super) const PF_MASK_BULLET_FONT: u32 = 0x0000_0010;
pub(super) const PF_MASK_BULLET_COLOR: u32 = 0x0000_0020;
pub(super) const PF_MASK_BULLET_SIZE: u32 = 0x0000_0040;
pub(super) const PF_MASK_BULLET_CHAR: u32 = 0x0000_0080;
pub(super) const PF_MASK_LEFT_MARGIN: u32 = 0x0000_0100;
pub(super) const PF_MASK_INDENT: u32 = 0x0000_0400;
pub(super) const PF_MASK_ALIGN: u32 = 0x0000_0800;
pub(super) const PF_MASK_LINE_SPACING: u32 = 0x0000_1000;
pub(super) const PF_MASK_SPACE_BEFORE: u32 = 0x0000_2000;
pub(super) const PF_MASK_SPACE_AFTER: u32 = 0x0000_4000;
pub(super) const PF_MASK_DEFAULT_TAB_SIZE: u32 = 0x0000_8000;
pub(super) const PF_MASK_FONT_ALIGN: u32 = 0x0001_0000;
pub(super) const PF_MASK_CHAR_WRAP: u32 = 0x0002_0000;
pub(super) const PF_MASK_WORD_WRAP: u32 = 0x0004_0000;
pub(super) const PF_MASK_OVERFLOW: u32 = 0x0008_0000;
pub(super) const PF_MASK_TAB_STOPS: u32 = 0x0010_0000;
pub(super) const PF_MASK_TEXT_DIRECTION: u32 = 0x0020_0000;
/// Mask bits that must be zero in a `TextPFException`: `reserved1`,
/// `bulletBlip`, `bulletScheme`, `bulletHasScheme`, and `reserved2`
/// (MS-PPT 2.9.20). The `unused` bit is undefined and ignored.
pub(super) const PF_MASK_FORBIDDEN: u32 = 0xFFC0_0000;
/// Mask bits that gate the presence of the `bulletFlags` field (MS-PPT 2.9.20).
pub(super) const PF_MASK_BULLET_FLAGS_PRESENCE: u32 = PF_MASK_HAS_BULLET
    | PF_MASK_BULLET_HAS_FONT
    | PF_MASK_BULLET_HAS_COLOR
    | PF_MASK_BULLET_HAS_SIZE;
/// Mask bits that gate the presence of the `wrapFlags` field (MS-PPT 2.9.20).
pub(super) const PF_MASK_WRAP_FLAGS_PRESENCE: u32 =
    PF_MASK_CHAR_WRAP | PF_MASK_WORD_WRAP | PF_MASK_OVERFLOW;

// BulletFlags bits (MS-PPT 2.9.22).
pub(super) const BULLET_HAS_BULLET: u16 = 0x0001;
pub(super) const BULLET_HAS_FONT: u16 = 0x0002;
pub(super) const BULLET_HAS_COLOR: u16 = 0x0004;
pub(super) const BULLET_HAS_SIZE: u16 = 0x0008;
/// `BulletFlags.reserved` bits: must be zero.
pub(super) const BULLET_FLAGS_RESERVED: u16 = 0xFFF0;

// PFWrapFlags bits (MS-PPT 2.9.25).
pub(super) const WRAP_CHAR_WRAP: u16 = 0x0001;
pub(super) const WRAP_WORD_WRAP: u16 = 0x0002;
pub(super) const WRAP_OVERFLOW: u16 = 0x0004;
/// `PFWrapFlags.reserved` bits: must be zero.
pub(super) const WRAP_FLAGS_RESERVED: u16 = 0xFFF8;

/// Size in bytes of one `TabStop` (MS-PPT 2.9.24).
pub(super) const TAB_STOP_LEN: usize = 4;
/// Largest `ParaSpacing` percentage; larger positive values are invalid
/// (MS-PPT 2.2.20).
pub(super) const MAX_PARA_SPACING_PERCENT: i16 = 13200;

impl TextCFException {
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
            let style = CFStyle {
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
            let value = ColorIndex::parse_bytes(&data[offset..offset + 4])?;
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
    pub(crate) fn to_payload(&self) -> Vec<u8> {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_record(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::TxCFStyleAtom
            || record.record_type_raw != TEXT_CF_EXCEPTION_TYPE
            || record.version != 0
            || record.instance != 0
        {
            return Err(corrupted("TextCFExceptionAtom has an invalid header"));
        }
        Self::parse(&record.data)
    }

    /// Parse the whole payload of a `TextCFExceptionAtom`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let (exception, consumed) = Self::parse_prefix(data)?;
        if consumed != data.len() {
            return Err(corrupted("TextCFExceptionAtom has trailing bytes"));
        }
        Ok(exception)
    }

    /// Serialize the complete `TextCFExceptionAtom`, including its header.
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "a `TextCFException` payload contains only fixed-size fields, so its length is bounded far below `u32::MAX`"
    )]
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

impl BulletFlags {
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

impl WrapFlags {
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
        u16::from(self.char_wrap) | u16::from(self.word_wrap) << 1 | u16::from(self.overflow) << 2
    }
}

impl TextPFException {
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
            let flags = BulletFlags::from_bits(read_u16(data, offset))?;
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
            let value = ColorIndex::parse_bytes(&data[offset..offset + 4])?;
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
            let flags = WrapFlags::from_bits(read_u16(data, offset))?;
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
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the tab-stop count is parsed from a `u16` length field and `TextPFException` fields are crate-private, so it cannot exceed `u16::MAX`"
    )]
    pub(crate) fn to_payload(&self) -> Vec<u8> {
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
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_record(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::TxPFStyleAtom
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
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
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
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "a `TextPFException` payload holds at most `u16::MAX` four-byte tab stops plus fixed-size fields, so its length is bounded far below `u32::MAX`"
    )]
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

fn corrupted(message: impl Into<String>) -> Error {
    Error::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_i16(data: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
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
