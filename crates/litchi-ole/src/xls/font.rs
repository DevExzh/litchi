//! BIFF8 workbook font table support.

use super::{XlsColor, XlsError, XlsPalette, XlsResult};

const FONT_FIXED_LENGTH: usize = 14;
const MAX_FONT_NAME_LENGTH: usize = 31;

/// Vertical positioning applied to a font.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsFontEscapement {
    Normal,
    Superscript,
    Subscript,
}

/// Underline style applied to a font.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsFontUnderline {
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

/// Windows logical font family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsFontFamily {
    NotApplicable,
    Roman,
    Swiss,
    Modern,
    Script,
    Decorative,
}

/// Windows character set associated with a BIFF8 font.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsFontCharset {
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
pub struct XlsFont {
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
    escapement: XlsFontEscapement,
    underline: XlsFontUnderline,
    family: XlsFontFamily,
    charset: XlsFontCharset,
}

impl XlsFont {
    pub fn index(&self) -> u16 {
        self.index
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn height_twips(&self) -> u16 {
        self.height_twips
    }

    pub fn color_index(&self) -> u16 {
        self.color_index
    }

    pub fn color(&self, palette: &XlsPalette) -> Option<XlsColor> {
        palette.color(self.color_index)
    }

    pub fn weight(&self) -> u16 {
        self.weight
    }

    pub fn is_bold(&self) -> bool {
        self.weight >= 700
    }

    pub fn is_italic(&self) -> bool {
        self.italic
    }

    pub fn is_struck_out(&self) -> bool {
        self.strikeout
    }

    pub fn is_outline(&self) -> bool {
        self.outline
    }

    pub fn has_shadow(&self) -> bool {
        self.shadow
    }

    pub fn is_condensed(&self) -> bool {
        self.condensed
    }

    pub fn is_extended(&self) -> bool {
        self.extended
    }

    pub fn escapement(&self) -> XlsFontEscapement {
        self.escapement
    }

    pub fn underline(&self) -> XlsFontUnderline {
        self.underline
    }

    pub fn family(&self) -> XlsFontFamily {
        self.family
    }

    pub fn charset(&self) -> XlsFontCharset {
        self.charset
    }

    pub(crate) fn parse_record(index: u16, data: &[u8]) -> XlsResult<Self> {
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
            0 => XlsFontEscapement::Normal,
            1 => XlsFontEscapement::Superscript,
            2 => XlsFontEscapement::Subscript,
            value => return Err(invalid(format!("Font escapement {value} is invalid"))),
        };
        let underline = match data[10] {
            0x00 => XlsFontUnderline::None,
            0x01 => XlsFontUnderline::Single,
            0x02 => XlsFontUnderline::Double,
            0x21 => XlsFontUnderline::SingleAccounting,
            0x22 => XlsFontUnderline::DoubleAccounting,
            value => return Err(invalid(format!("Font underline {value:#04x} is invalid"))),
        };
        let family = match data[11] {
            0 => XlsFontFamily::NotApplicable,
            1 => XlsFontFamily::Roman,
            2 => XlsFontFamily::Swiss,
            3 => XlsFontFamily::Modern,
            4 => XlsFontFamily::Script,
            5 => XlsFontFamily::Decorative,
            value => return Err(invalid(format!("Font family {value} is invalid"))),
        };
        let charset = match data[12] {
            0x00 => XlsFontCharset::Ansi,
            0x01 => XlsFontCharset::Default,
            0x02 => XlsFontCharset::Symbol,
            0x4d => XlsFontCharset::Mac,
            0x80 => XlsFontCharset::ShiftJis,
            0x81 => XlsFontCharset::Korean,
            0x82 => XlsFontCharset::Johab,
            0x86 => XlsFontCharset::Gb2312,
            0x88 => XlsFontCharset::ChineseBig5,
            0xa1 => XlsFontCharset::Greek,
            0xa2 => XlsFontCharset::Turkish,
            0xa3 => XlsFontCharset::Vietnamese,
            0xb1 => XlsFontCharset::Hebrew,
            0xb2 => XlsFontCharset::Arabic,
            0xba => XlsFontCharset::Baltic,
            0xcc => XlsFontCharset::Russian,
            0xdd => XlsFontCharset::Thai,
            0xee => XlsFontCharset::EastEurope,
            0xff => XlsFontCharset::Oem,
            value => return Err(invalid(format!("Font character set {value:#04x} is invalid"))),
        };

        let character_count = usize::from(data[14]);
        if !(1..=MAX_FONT_NAME_LENGTH).contains(&character_count) {
            return Err(invalid(format!(
                "Font name has {character_count} characters; expected 1..={MAX_FONT_NAME_LENGTH}"
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

pub(crate) fn logical_font_index(physical_index: usize) -> XlsResult<u16> {
    let logical_index = if physical_index < 4 {
        physical_index
    } else {
        physical_index
            .checked_add(1)
            .ok_or_else(|| invalid("Font index overflows"))?
    };
    let logical_index = u16::try_from(logical_index)
        .map_err(|_| invalid("Font index does not fit in BIFF8 FontIndex"))?;
    validate_font_index(logical_index)?;
    Ok(logical_index)
}

pub(crate) fn validate_font_index(index: u16) -> XlsResult<()> {
    if index == 4 || index > 1022 {
        return Err(invalid(format!("Font logical index {index} is invalid")));
    }
    Ok(())
}

pub(crate) fn validate_font_table(fonts: &[XlsFont]) -> XlsResult<()> {
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

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidData(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

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

    fn required_prefix() -> Vec<XlsFont> {
        [(400, false), (700, false), (400, true), (700, true)]
            .into_iter()
            .enumerate()
            .map(|(physical, (weight, italic))| {
                XlsFont::parse_record(
                    logical_font_index(physical).unwrap(),
                    &font_record(weight, italic, 0x7fff, "Arial"),
                )
                .unwrap()
            })
            .collect()
    }

    #[test]
    fn parses_normative_normal_and_bold_blue_fonts() {
        let normal = XlsFont::parse_record(0, &font_record(400, false, 0x7fff, "Arial"))
            .unwrap();
        assert_eq!(normal.name(), "Arial");
        assert_eq!(normal.height_twips(), 200);
        assert!(!normal.is_bold());
        assert!(!normal.is_italic());

        let bold = XlsFont::parse_record(5, &font_record(700, false, 0x000c, "Arial"))
            .unwrap();
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

        let font = XlsFont::parse_record(5, &data).unwrap();
        assert_eq!(font.name(), "ＭＳ ゴシック");
        assert!(font.is_bold());
        assert!(font.is_italic());
        assert!(font.is_struck_out());
        assert!(font.is_outline());
        assert!(font.has_shadow());
        assert!(font.is_condensed());
        assert!(font.is_extended());
        assert_eq!(font.escapement(), XlsFontEscapement::Superscript);
        assert_eq!(font.underline(), XlsFontUnderline::DoubleAccounting);
        assert_eq!(font.family(), XlsFontFamily::Modern);
        assert_eq!(font.charset(), XlsFontCharset::Thai);
        assert!(font.color(&XlsPalette::default()).is_some());
    }

    #[test]
    fn accepts_zero_weight_and_ignores_reserved_producer_bits() {
        let mut data = font_record(0, false, 0x7fff, "Arial");
        data[2..4].copy_from_slice(&0xff05u16.to_le_bytes());
        data[13] = 0xff;
        data[15] = 0xff;
        XlsFont::parse_record(0, &data).unwrap();
    }

    #[test]
    fn skips_reserved_logical_index_four() {
        assert_eq!(logical_font_index(0).unwrap(), 0);
        assert_eq!(logical_font_index(3).unwrap(), 3);
        assert_eq!(logical_font_index(4).unwrap(), 5);
        assert!(XlsFont::parse_record(4, &font_record(400, false, 0x7fff, "Arial")).is_err());
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
            assert!(XlsFont::parse_record(0, &data).is_err());
        }

        for (offset, value) in [(10, 3), (11, 6), (12, 3), (12, 0xde)] {
            let mut data = font_record(400, false, 0x7fff, "Arial");
            data[offset] = value;
            assert!(XlsFont::parse_record(0, &data).is_err());
        }
    }

    #[test]
    fn rejects_invalid_name_lengths_and_truncation() {
        let mut empty_name = font_record(400, false, 0x7fff, "A");
        empty_name[14] = 0;
        assert!(XlsFont::parse_record(0, &empty_name).is_err());

        let mut long_name = font_record(400, false, 0x7fff, "Arial");
        long_name[14] = 32;
        assert!(XlsFont::parse_record(0, &long_name).is_err());

        let mut truncated = font_record(400, false, 0x7fff, "Arial");
        truncated.pop();
        assert!(XlsFont::parse_record(0, &truncated).is_err());

        assert!(XlsFont::parse_record(0, &font_record(400, false, 0x7fff, "A\0B")).is_err());
    }

    #[test]
    fn accepts_compressed_font_names_from_nonconformant_producers() {
        let mut compressed = font_record(400, false, 0x7fff, "Arial");
        compressed[15] = 0;
        compressed.truncate(16);
        compressed.extend_from_slice(b"Arial");

        let font = XlsFont::parse_record(0, &compressed).unwrap();
        assert_eq!(font.name(), "Arial");
    }
}
