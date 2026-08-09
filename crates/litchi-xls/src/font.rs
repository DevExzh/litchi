//! BIFF8 workbook font table support.

use super::leniency::{FormattingDefect, ToleranceLog};
use super::{Color, Error, Palette, Result};

/// MS-XLS 2.4.122 `Font` record type.
pub(crate) const FONT_RECORD_TYPE: u16 = 0x0031;
const FONT_FIXED_LENGTH: usize = 14;
const MAX_FONT_NAME_LENGTH: usize = 31;
/// Smallest `Font.cch` MS-XLS permits; zero is the tolerated defect.
const MIN_FONT_NAME_LENGTH: usize = 1;
/// Largest `Font.bFamily` value defined by MS-XLS 2.4.122.
const MAX_FONT_FAMILY: u8 = 5;

/// Vertical positioning applied to a font.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FontEscapement {
    Normal,
    Superscript,
    Subscript,
}

/// Underline style applied to a font.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FontUnderline {
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

/// Windows logical font family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FontFamily {
    NotApplicable,
    Roman,
    Swiss,
    Modern,
    Script,
    Decorative,
}

/// Windows character set associated with a BIFF8 font.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FontCharset {
    Ansi,
    Default,
    Symbol,
    Mac,
    ShiftJis,
    Korean,
    Johab,
    Gb2312,
    ChineseBig5,
    Greek,
    Turkish,
    Vietnamese,
    Hebrew,
    Arabic,
    Baltic,
    Russian,
    Thai,
    EastEurope,
    Oem,
}

/// Font and font-formatting information from a global `Font` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    index: u16,
    name: String,
    height_twips: u16,
    color_index: u16,
    weight: u16,
    italic: bool,
    strikeout: bool,
    outline: bool,
    shadow: bool,
    condensed: bool,
    extended: bool,
    escapement: FontEscapement,
    underline: FontUnderline,
    family: FontFamily,
    charset: FontCharset,
}

impl Font {
    #[must_use]
    pub fn index(&self) -> u16 {
        self.index
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn height_twips(&self) -> u16 {
        self.height_twips
    }

    #[must_use]
    pub fn color_index(&self) -> u16 {
        self.color_index
    }

    #[must_use]
    pub fn color(&self, palette: &Palette) -> Option<Color> {
        palette.color(self.color_index)
    }

    #[must_use]
    pub fn weight(&self) -> u16 {
        self.weight
    }

    #[must_use]
    pub fn is_bold(&self) -> bool {
        self.weight >= 700
    }

    #[must_use]
    pub fn is_italic(&self) -> bool {
        self.italic
    }

    #[must_use]
    pub fn is_struck_out(&self) -> bool {
        self.strikeout
    }

    #[must_use]
    pub fn is_outline(&self) -> bool {
        self.outline
    }

    #[must_use]
    pub fn has_shadow(&self) -> bool {
        self.shadow
    }

    #[must_use]
    pub fn is_condensed(&self) -> bool {
        self.condensed
    }

    #[must_use]
    pub fn is_extended(&self) -> bool {
        self.extended
    }

    #[must_use]
    pub fn escapement(&self) -> FontEscapement {
        self.escapement
    }

    #[must_use]
    pub fn underline(&self) -> FontUnderline {
        self.underline
    }

    #[must_use]
    pub fn family(&self) -> FontFamily {
        self.family
    }

    #[must_use]
    pub fn charset(&self) -> FontCharset {
        self.charset
    }

    /// Parse a `Font` record under an explicit leniency policy.
    ///
    /// Under [`super::Leniency::TolerateFormattingDefects`] an out-of-range
    /// `bFamily` degrades to [`FontFamily::NotApplicable`] and a zero `cch`
    /// yields an empty name; both are recorded in `tolerance`. Every other
    /// deviation — including a payload whose length disagrees with `cch` — stays
    /// a hard error, because that is a framing defect rather than a cosmetic one.
    pub(crate) fn parse_record(
        index: u16,
        data: &[u8],
        tolerance: &mut ToleranceLog,
    ) -> Result<Self> {
        validate_font_index(index)?;
        if data.len() < FONT_FIXED_LENGTH + 2 {
            return Err(invalid(format!(
                "Font record has {} bytes; expected at least {}",
                data.len(),
                FONT_FIXED_LENGTH + 2
            )));
        }

        let height_twips = u16::from_le_bytes([data[0], data[1]]);
        if height_twips != 0 && !(20..=8191).contains(&height_twips) {
            return Err(invalid(format!(
                "Font height is {height_twips} twips; expected zero or 20..=8191"
            )));
        }
        let flags = u16::from_le_bytes([data[2], data[3]]);

        let color_index = u16::from_le_bytes([data[4], data[5]]);
        if !valid_color_index(color_index) {
            return Err(invalid(format!(
                "Font color index {color_index:#06x} is not a valid Icv"
            )));
        }
        let weight = u16::from_le_bytes([data[6], data[7]]);
        if weight != 0 && !(100..=1000).contains(&weight) {
            return Err(invalid(format!(
                "Font weight is {weight}; expected zero or 100..=1000"
            )));
        }
        let escapement = match u16::from_le_bytes([data[8], data[9]]) {
            0 => FontEscapement::Normal,
            1 => FontEscapement::Superscript,
            2 => FontEscapement::Subscript,
            value => return Err(invalid(format!("Font escapement {value} is invalid"))),
        };
        let underline = match data[10] {
            0x00 => FontUnderline::None,
            0x01 => FontUnderline::Single,
            0x02 => FontUnderline::Double,
            0x21 => FontUnderline::SingleAccounting,
            0x22 => FontUnderline::DoubleAccounting,
            value => return Err(invalid(format!("Font underline {value:#04x} is invalid"))),
        };
        let family = match data[11] {
            0 => FontFamily::NotApplicable,
            1 => FontFamily::Roman,
            2 => FontFamily::Swiss,
            3 => FontFamily::Modern,
            4 => FontFamily::Script,
            5 => FontFamily::Decorative,
            value => {
                tolerance.tolerate(
                    FormattingDefect::FontFamily,
                    u32::from(index),
                    u32::from(value),
                    || {
                        invalid(format!(
                            "Font family {value} is invalid; expected 0..={MAX_FONT_FAMILY}"
                        ))
                    },
                )?;
                FontFamily::NotApplicable
            },
        };
        let charset = match data[12] {
            0x00 => FontCharset::Ansi,
            0x01 => FontCharset::Default,
            0x02 => FontCharset::Symbol,
            0x4d => FontCharset::Mac,
            0x80 => FontCharset::ShiftJis,
            0x81 => FontCharset::Korean,
            0x82 => FontCharset::Johab,
            0x86 => FontCharset::Gb2312,
            0x88 => FontCharset::ChineseBig5,
            0xa1 => FontCharset::Greek,
            0xa2 => FontCharset::Turkish,
            0xa3 => FontCharset::Vietnamese,
            0xb1 => FontCharset::Hebrew,
            0xb2 => FontCharset::Arabic,
            0xba => FontCharset::Baltic,
            0xcc => FontCharset::Russian,
            0xdd => FontCharset::Thai,
            0xee => FontCharset::EastEurope,
            0xff => FontCharset::Oem,
            value => {
                return Err(invalid(format!(
                    "Font character set {value:#04x} is invalid"
                )));
            },
        };

        let character_count = usize::from(data[14]);
        if character_count == 0 {
            // A nameless font is cosmetic: the record still carries every
            // metric, and a renderer substitutes its own default face exactly
            // as it would for a name it does not have installed.
            tolerance.tolerate(FormattingDefect::FontNameEmpty, u32::from(index), 0, || {
                invalid(format!(
                    "Font name has 0 characters; expected {MIN_FONT_NAME_LENGTH}..={MAX_FONT_NAME_LENGTH}"
                ))
            })?;
        } else if character_count > MAX_FONT_NAME_LENGTH {
            return Err(invalid(format!(
                "Font name has {character_count} characters; expected {MIN_FONT_NAME_LENGTH}..={MAX_FONT_NAME_LENGTH}"
            )));
        }
        let string_flags = data[15];
        let wide = string_flags & 0x01 != 0;
        let character_bytes = character_count
            .checked_mul(if wide { 2 } else { 1 })
            .ok_or_else(|| invalid("Font name byte length overflows"))?;
        let expected_length = FONT_FIXED_LENGTH
            .checked_add(2)
            .and_then(|value| value.checked_add(character_bytes))
            .ok_or_else(|| invalid("Font record length overflows"))?;
        if data.len() != expected_length {
            return Err(invalid(format!(
                "Font record has {} bytes; expected {expected_length}",
                data.len()
            )));
        }

        let name_bytes = &data[16..];
        let name = if wide {
            char::decode_utf16(
                name_bytes
                    .chunks_exact(2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
            )
            .collect::<Result<String, _>>()
            .map_err(|error| invalid(format!("Font name is invalid UTF-16: {error}")))?
        } else {
            name_bytes.iter().map(|&value| char::from(value)).collect()
        };
        if name.contains('\0') {
            return Err(invalid("Font name contains a null character"));
        }

        Ok(Self {
            index,
            name,
            height_twips,
            color_index,
            weight,
            italic: flags & 0x0002 != 0,
            strikeout: flags & 0x0008 != 0,
            outline: flags & 0x0010 != 0,
            shadow: flags & 0x0020 != 0,
            condensed: flags & 0x0040 != 0,
            extended: flags & 0x0080 != 0,
            escapement,
            underline,
            family,
            charset,
        })
    }
}

pub(crate) fn logical_font_index(physical_index: usize) -> Result<u16> {
    let logical_index = if physical_index < 4 {
        physical_index
    } else {
        physical_index
            .checked_add(1)
            .ok_or_else(|| invalid("Font index overflows"))?
    };
    let logical_index = u16::try_from(logical_index)
        .map_err(|_error| invalid("Font index does not fit in BIFF8 FontIndex"))?;
    validate_font_index(logical_index)?;
    Ok(logical_index)
}

pub(crate) fn validate_font_index(index: u16) -> Result<()> {
    if index == 4 || index > 1022 {
        return Err(invalid(format!("Font logical index {index} is invalid")));
    }
    Ok(())
}

pub(crate) fn validate_font_table(fonts: &[Font]) -> Result<()> {
    if fonts.len() < 4 {
        return Err(invalid(format!(
            "workbook has {} Font records; expected at least four",
            fonts.len()
        )));
    }
    for (physical_index, font) in fonts.iter().enumerate() {
        let expected_index = logical_font_index(physical_index)?;
        if font.index != expected_index {
            return Err(invalid(format!(
                "Font record {physical_index} has logical index {}; expected {expected_index}",
                font.index,
            )));
        }
    }
    Ok(())
}

fn valid_color_index(index: u16) -> bool {
    matches!(index, 0x0000..=0x0041 | 0x004d..=0x004f | 0x0051 | 0x7fff)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::leniency::Leniency;

    /// Strict-mode shim: these tests exercise the default reject-everything
    /// policy, which the production reader threads through a tolerance log.
    fn parse_record(index: u16, data: &[u8]) -> Result<Font> {
        Font::parse_record(index, data, &mut ToleranceLog::new(Leniency::Strict))
    }

    fn font_record(weight: u16, italic: bool, color_index: u16, name: &str) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&200u16.to_le_bytes());
        data.extend_from_slice(&(if italic { 0x0002u16 } else { 0 }).to_le_bytes());
        data.extend_from_slice(&color_index.to_le_bytes());
        data.extend_from_slice(&weight.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.push(0);
        data.push(0);
        data.push(0);
        data.push(0xdf);
        data.push(name.encode_utf16().count() as u8);
        data.push(1);
        for unit in name.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    fn required_prefix() -> Vec<Font> {
        [(400, false), (700, false), (400, true), (700, true)]
            .into_iter()
            .enumerate()
            .map(|(physical, (weight, italic))| {
                parse_record(
                    logical_font_index(physical).unwrap(),
                    &font_record(weight, italic, 0x7fff, "Arial"),
                )
                .unwrap()
            })
            .collect()
    }

    #[test]
    fn parses_normative_normal_and_bold_blue_fonts() {
        let normal = parse_record(0, &font_record(400, false, 0x7fff, "Arial")).unwrap();
        assert_eq!(normal.name(), "Arial");
        assert_eq!(normal.height_twips(), 200);
        assert!(!normal.is_bold());
        assert!(!normal.is_italic());

        let bold = parse_record(5, &font_record(700, false, 0x000c, "Arial")).unwrap();
        assert!(bold.is_bold());
        assert_eq!(bold.color_index(), 0x000c);
    }

    #[test]
    fn parses_all_font_attributes_and_unicode_name() {
        let mut data = font_record(700, true, 0x000c, "ＭＳ ゴシック");
        data[2..4].copy_from_slice(&0x00fau16.to_le_bytes());
        data[8..10].copy_from_slice(&1u16.to_le_bytes());
        data[10] = 0x22;
        data[11] = 3;
        data[12] = 0xdd;
        data[13] = 0xff;
        data[15] |= 0xfe;

        let font = parse_record(5, &data).unwrap();
        assert_eq!(font.name(), "ＭＳ ゴシック");
        assert!(font.is_bold());
        assert!(font.is_italic());
        assert!(font.is_struck_out());
        assert!(font.is_outline());
        assert!(font.has_shadow());
        assert!(font.is_condensed());
        assert!(font.is_extended());
        assert_eq!(font.escapement(), FontEscapement::Superscript);
        assert_eq!(font.underline(), FontUnderline::DoubleAccounting);
        assert_eq!(font.family(), FontFamily::Modern);
        assert_eq!(font.charset(), FontCharset::Thai);
        assert!(font.color(&Palette::default()).is_some());
    }

    #[test]
    fn accepts_zero_weight_and_ignores_reserved_producer_bits() {
        let mut data = font_record(0, false, 0x7fff, "Arial");
        data[2..4].copy_from_slice(&0xff05u16.to_le_bytes());
        data[13] = 0xff;
        data[15] = 0xff;
        parse_record(0, &data).unwrap();
    }

    #[test]
    fn skips_reserved_logical_index_four() {
        assert_eq!(logical_font_index(0).unwrap(), 0);
        assert_eq!(logical_font_index(3).unwrap(), 3);
        assert_eq!(logical_font_index(4).unwrap(), 5);
        assert!(parse_record(4, &font_record(400, false, 0x7fff, "Arial")).is_err());
    }

    #[test]
    fn validates_required_default_font_prefix() {
        let fonts = required_prefix();
        validate_font_table(&fonts).unwrap();

        assert!(validate_font_table(&fonts[..3]).is_err());
        let mut wrong = fonts;
        wrong[1].index = 5;
        assert!(validate_font_table(&wrong).is_err());
    }

    #[test]
    fn rejects_invalid_scalar_fields() {
        for (offset, bytes) in [
            (0, 19u16.to_le_bytes()),
            (4, 0x0042u16.to_le_bytes()),
            (6, 99u16.to_le_bytes()),
            (8, 3u16.to_le_bytes()),
        ] {
            let mut data = font_record(400, false, 0x7fff, "Arial");
            data[offset..offset + 2].copy_from_slice(&bytes);
            assert!(parse_record(0, &data).is_err());
        }

        for (offset, value) in [(10, 3), (11, 6), (12, 3), (12, 0xde)] {
            let mut data = font_record(400, false, 0x7fff, "Arial");
            data[offset] = value;
            assert!(parse_record(0, &data).is_err());
        }
    }

    #[test]
    fn rejects_invalid_name_lengths_and_truncation() {
        let mut empty_name = font_record(400, false, 0x7fff, "A");
        empty_name[14] = 0;
        assert!(parse_record(0, &empty_name).is_err());

        let mut long_name = font_record(400, false, 0x7fff, "Arial");
        long_name[14] = 32;
        assert!(parse_record(0, &long_name).is_err());

        let mut truncated = font_record(400, false, 0x7fff, "Arial");
        truncated.pop();
        assert!(parse_record(0, &truncated).is_err());

        assert!(parse_record(0, &font_record(400, false, 0x7fff, "A\0B")).is_err());
    }

    #[test]
    fn accepts_compressed_font_names_from_nonconformant_producers() {
        let mut compressed = font_record(400, false, 0x7fff, "Arial");
        compressed[15] = 0;
        compressed.truncate(16);
        compressed.extend_from_slice(b"Arial");

        let font = parse_record(0, &compressed).unwrap();
        assert_eq!(font.name(), "Arial");
    }
}
