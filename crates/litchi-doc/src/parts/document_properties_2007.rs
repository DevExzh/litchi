//! Typed, lossless Word 2007 Document Properties extension.

use super::document_properties_97::DopExtensionError;
use super::document_properties_2003::Dop2003;

const DOP2003_SIZE: usize = 616;
const DOP2007_SIZE: usize = 674;
const EXTENSION_SIZE: usize = DOP2007_SIZE - DOP2003_SIZE;
const SETTINGS_FLAGS: usize = 4;
const EMPTY_FIELDS: std::ops::Range<usize> = 8..24;
const MATH_PROPERTIES: usize = 24;
const RESERVED_SETTINGS_MASK: u32 = 0xffff_f81c;
const MAX_MATH_MARGIN_TWIPS: i32 = 31_680;

/// Placement of a binary operator relative to a wrapped equation line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathBreakPlacement {
    Before,
    After,
    Repeat,
}

impl MathBreakPlacement {
    fn parse(raw: u32) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::Before),
            1 => Ok(Self::After),
            2 => Ok(Self::Repeat),
            _ => Err(DopExtensionError::new(
                "invalid DopMth binary-operator break placement",
            )),
        }
    }

    const fn raw(self) -> u32 {
        self as u32
    }
}

/// Sign placement when subtraction is repeated across a wrapped equation line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathSubtractionBreak {
    MinusMinus,
    PlusMinus,
    MinusPlus,
}

impl MathSubtractionBreak {
    fn parse(raw: u32) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::MinusMinus),
            1 => Ok(Self::PlusMinus),
            2 => Ok(Self::MinusPlus),
            _ => Err(DopExtensionError::new(
                "invalid DopMth subtraction break placement",
            )),
        }
    }

    const fn raw(self) -> u32 {
        self as u32
    }
}

/// Default justification of display equations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathJustification {
    CenteredAsGroup,
    Center,
    Left,
    Right,
}

impl MathJustification {
    fn parse(raw: u32) -> Result<Self, DopExtensionError> {
        match raw {
            1 => Ok(Self::CenteredAsGroup),
            2 => Ok(Self::Center),
            3 => Ok(Self::Left),
            4 => Ok(Self::Right),
            _ => Err(DopExtensionError::new(
                "invalid DopMth default math justification",
            )),
        }
    }

    const fn raw(self) -> u32 {
        match self {
            Self::CenteredAsGroup => 1,
            Self::Center => 2,
            Self::Left => 3,
            Self::Right => 4,
        }
    }
}

/// Document-wide equation settings (`DopMth`, MS-DOC 2.7.17).
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "the independent Boolean fields are exact DopMth wire flags"
)]
pub struct DopMth {
    raw: [u8; 34],
    pub break_placement: MathBreakPlacement,
    pub subtraction_break: MathSubtractionBreak,
    pub justification: MathJustification,
    pub small_fractions: bool,
    pub integral_limits_above_below: bool,
    pub nary_limits_above_below: bool,
    pub wrapped_line_align_left: bool,
    pub use_display_defaults: bool,
    /// Index into the document font table used for newly inserted equations.
    pub math_font_index: u16,
    pub left_margin_twips: i32,
    pub right_margin_twips: i32,
    pub wrapped_indent_twips: i32,
}

impl DopMth {
    pub const BYTE_LEN: usize = 34;

    /// Parses one exact 34-byte `DopMth` record.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] for truncation, reserved values, invalid
    /// fixed fields, or distances outside the specification domain.
    pub fn parse(data: &[u8]) -> Result<Self, DopExtensionError> {
        let record = data
            .get(..Self::BYTE_LEN)
            .ok_or_else(|| DopExtensionError::new("DopMth is shorter than 34 bytes"))?;
        let flags = le_u32(record, 0);
        if flags & 0xffff_e000 != 0 {
            return Err(DopExtensionError::new(
                "DopMth reserved high bits are nonzero",
            ));
        }
        for (offset, expected) in [(14, 120), (18, 120), (22, 0), (26, 0)] {
            if le_i32(record, offset) != expected {
                return Err(DopExtensionError::new(format!(
                    "DopMth fixed field at offset {offset} is not {expected}"
                )));
            }
        }
        let left_margin_twips = checked_math_distance(record, 6, "left margin")?;
        let right_margin_twips = checked_math_distance(record, 10, "right margin")?;
        let wrapped_indent_twips = checked_math_distance(record, 30, "wrapped indent")?;
        let mut raw = [0u8; Self::BYTE_LEN];
        raw.copy_from_slice(record);
        Ok(Self {
            raw,
            break_placement: MathBreakPlacement::parse(flags & 0x03)?,
            subtraction_break: MathSubtractionBreak::parse((flags >> 2) & 0x03)?,
            justification: MathJustification::parse((flags >> 4) & 0x07)?,
            small_fractions: flags & (1 << 8) != 0,
            integral_limits_above_below: flags & (1 << 9) != 0,
            nary_limits_above_below: flags & (1 << 10) != 0,
            wrapped_line_align_left: flags & (1 << 11) != 0,
            use_display_defaults: flags & (1 << 12) != 0,
            math_font_index: le_u16(record, 4),
            left_margin_twips,
            right_margin_twips,
            wrapped_indent_twips,
        })
    }

    /// Validates the font index when the font-table slot count is available.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when the stored index is outside the
    /// caller-supplied font table.
    pub fn validate_font_index(&self, font_count: usize) -> Result<(), DopExtensionError> {
        if usize::from(self.math_font_index) >= font_count {
            Err(DopExtensionError::new(format!(
                "DopMth math font index {} exceeds font table",
                self.math_font_index
            )))
        } else {
            Ok(())
        }
    }

    /// Writes this record while retaining its undefined bit.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when `target` is too short or a distance
    /// was changed outside the specification domain.
    pub fn write_into(mut self, target: &mut [u8]) -> Result<(), DopExtensionError> {
        let output = target
            .get_mut(..Self::BYTE_LEN)
            .ok_or_else(|| DopExtensionError::new("DopMth target is shorter than 34 bytes"))?;
        validate_math_distance(self.left_margin_twips, "left margin")?;
        validate_math_distance(self.right_margin_twips, "right margin")?;
        validate_math_distance(self.wrapped_indent_twips, "wrapped indent")?;

        // Bit 7 is undefined rather than reserved, so retain it losslessly.
        let mut flags = le_u32(&self.raw, 0) & (1 << 7);
        flags |= self.break_placement.raw();
        flags |= self.subtraction_break.raw() << 2;
        flags |= self.justification.raw() << 4;
        flags |= u32::from(self.small_fractions) << 8;
        flags |= u32::from(self.integral_limits_above_below) << 9;
        flags |= u32::from(self.nary_limits_above_below) << 10;
        flags |= u32::from(self.wrapped_line_align_left) << 11;
        flags |= u32::from(self.use_display_defaults) << 12;
        put_u32(&mut self.raw, 0, flags);
        put_u16(&mut self.raw, 4, self.math_font_index);
        put_i32(&mut self.raw, 6, self.left_margin_twips);
        put_i32(&mut self.raw, 10, self.right_margin_twips);
        put_i32(&mut self.raw, 14, 120);
        put_i32(&mut self.raw, 18, 120);
        put_i32(&mut self.raw, 22, 0);
        put_i32(&mut self.raw, 26, 0);
        put_i32(&mut self.raw, 30, self.wrapped_indent_twips);
        output.copy_from_slice(&self.raw);
        Ok(())
    }
}

/// Style-pane sorting method stored by Word 2007.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleSortMethod {
    Name,
    ApplicationDefault,
    Font,
    BasedOn,
    StyleType,
}

impl StyleSortMethod {
    fn parse(raw: u32) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::Name),
            1 => Ok(Self::ApplicationDefault),
            2 => Ok(Self::Font),
            3 => Ok(Self::BasedOn),
            4 => Ok(Self::StyleType),
            _ => Err(DopExtensionError::new(format!(
                "invalid Dop2007 style sorting method {raw}"
            ))),
        }
    }

    const fn raw(self) -> u32 {
        self as u32
    }
}

/// Typed, lossless Word 2007 DOP extension.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "the independent Boolean fields are exact Dop2007 wire flags"
)]
pub struct Dop2007 {
    raw: [u8; EXTENSION_SIZE],
    pub track_formatting: bool,
    pub track_moves: bool,
    pub style_sort_method: StyleSortMethod,
    pub reading_mode_actual_pages: bool,
    pub auto_compress_pictures: bool,
    pub math: DopMth,
}

impl Dop2007 {
    /// Parses the complete Word 2007 DOP generation.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when this or any earlier DOP prefix
    /// violates its specification-defined grammar.
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP2007_SIZE {
            return Err(DopExtensionError::new("Dop2007 is shorter than 674 bytes"));
        }
        Dop2003::parse(dop)?;
        let extension = &dop[DOP2003_SIZE..DOP2007_SIZE];
        let flags = le_u32(extension, SETTINGS_FLAGS);
        if flags & RESERVED_SETTINGS_MASK != 0 {
            return Err(DopExtensionError::new(
                "Dop2007 reserved settings bits are nonzero",
            ));
        }
        if extension[EMPTY_FIELDS].iter().any(|byte| *byte != 0) {
            return Err(DopExtensionError::new(
                "Dop2007 fixed empty fields are nonzero",
            ));
        }
        let math = DopMth::parse(&extension[MATH_PROPERTIES..])?;
        let mut raw = [0u8; EXTENSION_SIZE];
        raw.copy_from_slice(extension);
        Ok(Self {
            raw,
            track_formatting: flags & 1 != 0,
            track_moves: flags & 2 != 0,
            style_sort_method: StyleSortMethod::parse((flags >> 5) & 0x0f)?,
            reading_mode_actual_pages: flags & (1 << 9) != 0,
            auto_compress_pictures: flags & (1 << 10) != 0,
            math,
        })
    }

    /// Writes the Word 2007 extension without normalizing older DOP bytes.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when the target is too short or the
    /// nested math settings are invalid.
    pub fn write_into(mut self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP2007_SIZE {
            return Err(DopExtensionError::new(
                "Dop2007 target is shorter than 674 bytes",
            ));
        }
        let mut flags = u32::from(self.track_formatting);
        flags |= u32::from(self.track_moves) << 1;
        flags |= self.style_sort_method.raw() << 5;
        flags |= u32::from(self.reading_mode_actual_pages) << 9;
        flags |= u32::from(self.auto_compress_pictures) << 10;
        put_u32(&mut self.raw, SETTINGS_FLAGS, flags);
        self.raw[EMPTY_FIELDS].fill(0);
        self.math
            .write_into(&mut self.raw[MATH_PROPERTIES..MATH_PROPERTIES + DopMth::BYTE_LEN])?;
        dop[DOP2003_SIZE..DOP2007_SIZE].copy_from_slice(&self.raw);
        Ok(())
    }
}

fn checked_math_distance(data: &[u8], offset: usize, name: &str) -> Result<i32, DopExtensionError> {
    let value = le_i32(data, offset);
    validate_math_distance(value, name)?;
    Ok(value)
}

fn validate_math_distance(value: i32, name: &str) -> Result<(), DopExtensionError> {
    if (0..=MAX_MATH_MARGIN_TWIPS).contains(&value) {
        Ok(())
    } else {
        Err(DopExtensionError::new(format!(
            "DopMth {name} {value} is outside 0..={MAX_MATH_MARGIN_TWIPS} twips"
        )))
    }
}

fn le_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn le_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn le_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn put_u16(data: &mut [u8], offset: usize, value: u16) {
    data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
}

fn put_u32(data: &mut [u8], offset: usize, value: u32) {
    data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn put_i32(data: &mut [u8], offset: usize, value: i32) {
    data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_dop2007() -> Vec<u8> {
        let mut dop = vec![0u8; DOP2007_SIZE];
        dop[0x190..0x19a].copy_from_slice(&[0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81]);
        let math = DOP2003_SIZE + MATH_PROPERTIES;
        put_u32(&mut dop, math, 1 << 11 | 1 << 12 | 1 << 4);
        put_i32(&mut dop, math + 14, 120);
        put_i32(&mut dop, math + 18, 120);
        put_i32(&mut dop, math + 30, 1440);
        dop
    }

    #[test]
    fn parses_and_writes_word_2007_and_math_settings_losslessly() {
        let mut dop = valid_dop2007();
        put_u32(
            &mut dop,
            DOP2003_SIZE + SETTINGS_FLAGS,
            1 | 2 | (4 << 5) | (1 << 9) | (1 << 10),
        );
        let math = DOP2003_SIZE + MATH_PROPERTIES;
        put_u32(
            &mut dop,
            math,
            2 | (1 << 2) | (3 << 4) | (1 << 7) | (1 << 8) | (1 << 12),
        );
        put_u16(&mut dop, math + 4, 7);
        put_i32(&mut dop, math + 6, 720);
        put_i32(&mut dop, math + 10, 360);

        let parsed = Dop2007::parse(&dop).unwrap();
        assert_eq!(parsed.style_sort_method, StyleSortMethod::StyleType);
        assert_eq!(parsed.math.break_placement, MathBreakPlacement::Repeat);
        assert_eq!(parsed.math.justification, MathJustification::Left);
        assert_eq!(parsed.math.left_margin_twips, 720);
        let mut output = dop.clone();
        parsed.write_into(&mut output).unwrap();
        assert_eq!(output, dop);
    }

    #[test]
    fn writes_typed_mutations_and_preserves_undefined_bits() {
        let mut dop = valid_dop2007();
        dop[DOP2003_SIZE..DOP2003_SIZE + 4].copy_from_slice(&[1, 2, 3, 4]);
        let math = DOP2003_SIZE + MATH_PROPERTIES;
        let math_flags = le_u32(&dop, math) | (1 << 7);
        put_u32(&mut dop, math, math_flags);
        let mut parsed = Dop2007::parse(&dop).unwrap();
        parsed.track_moves = true;
        parsed.math.right_margin_twips = 480;
        let mut output = dop.clone();
        parsed.write_into(&mut output).unwrap();
        assert_eq!(&output[DOP2003_SIZE..DOP2003_SIZE + 4], &[1, 2, 3, 4]);
        assert_ne!(le_u32(&output, math) & (1 << 7), 0);
        assert_eq!(
            Dop2007::parse(&output).unwrap().math.right_margin_twips,
            480
        );
    }

    #[test]
    fn rejects_reserved_values_and_invalid_math_domains() {
        let mut settings = valid_dop2007();
        put_u32(&mut settings, DOP2003_SIZE + SETTINGS_FLAGS, 1 << 2);
        assert!(Dop2007::parse(&settings).is_err());

        let mut justification = valid_dop2007();
        let math = DOP2003_SIZE + MATH_PROPERTIES;
        put_u32(&mut justification, math, 0);
        assert!(Dop2007::parse(&justification).is_err());

        let mut margin = valid_dop2007();
        put_i32(&mut margin, math + 6, -1);
        assert!(Dop2007::parse(&margin).is_err());

        let mut fixed = valid_dop2007();
        put_i32(&mut fixed, math + 14, 119);
        assert!(Dop2007::parse(&fixed).is_err());
    }
}
