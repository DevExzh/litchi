/// Numbering and list structures parser for Word binary format.
///
/// Based on Apache POI's ListTables and LibreOffice's implementation.
/// Lists in DOC files are defined by:
/// - List Format Override (LFO) structures
/// - List Format (LF) structures  
/// - List Level Format (LVL) structures
use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;
use litchi_core::binary;

/// Number format for list levels
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum NumberFormat {
    /// Arabic numerals (1, 2, 3...)
    Arabic = 0,
    /// Uppercase Roman (I, II, III...)
    UpperRoman = 1,
    /// Lowercase Roman (i, ii, iii...)
    LowerRoman = 2,
    /// Uppercase letters (A, B, C...)
    UpperLetter = 3,
    /// Lowercase letters (a, b, c...)
    LowerLetter = 4,
    /// Ordinal numbers (1st, 2nd, 3rd...)
    Ordinal = 5,
    /// Cardinal text (One, Two, Three...)
    CardinalText = 6,
    /// Ordinal text (First, Second, Third...)
    OrdinalText = 7,
    Hex = 8,
    Chicago = 9,
    IdeographDigital = 10,
    JapaneseCounting = 11,
    Aiueo = 12,
    Iroha = 13,
    DecimalFullWidth = 14,
    DecimalHalfWidth = 15,
    JapaneseLegal = 16,
    JapaneseDigitalTenThousand = 17,
    DecimalEnclosedCircle = 18,
    DecimalFullWidth2 = 19,
    AiueoFullWidth = 20,
    IrohaFullWidth = 21,
    DecimalZero = 22,
    /// Bullet
    Bullet = 23,
    Ganada = 24,
    Chosung = 25,
    DecimalEnclosedFullstop = 26,
    DecimalEnclosedParen = 27,
    DecimalEnclosedCircleChinese = 28,
    IdeographEnclosedCircle = 29,
    IdeographTraditional = 30,
    IdeographZodiac = 31,
    IdeographZodiacTraditional = 32,
    TaiwaneseCounting = 33,
    IdeographLegalTraditional = 34,
    TaiwaneseCountingThousand = 35,
    TaiwaneseDigital = 36,
    ChineseCounting = 37,
    ChineseLegalSimplified = 38,
    ChineseCountingThousand = 39,
    DecimalChinese = 40,
    KoreanDigital = 41,
    KoreanCounting = 42,
    KoreanLegal = 43,
    KoreanDigital2 = 44,
    Hebrew1 = 45,
    ArabicAlpha = 46,
    Hebrew2 = 47,
    ArabicAbjad = 48,
    HindiVowels = 49,
    HindiConsonants = 50,
    HindiNumbers = 51,
    HindiCounting = 52,
    ThaiLetters = 53,
    ThaiNumbers = 54,
    ThaiCounting = 55,
    VietnameseCounting = 56,
    NumberInDash = 57,
    RussianLower = 58,
    RussianUpper = 59,
    /// No numbering
    None = 255,
}

impl NumberFormat {
    /// Writer-compatible name for decimal numbering.
    #[allow(non_upper_case_globals)]
    pub const Decimal: Self = Self::Arabic;
}

impl TryFrom<u8> for NumberFormat {
    type Error = u8;

    fn try_from(value: u8) -> std::result::Result<Self, Self::Error> {
        const VALUES: [NumberFormat; 60] = [
            NumberFormat::Arabic,
            NumberFormat::UpperRoman,
            NumberFormat::LowerRoman,
            NumberFormat::UpperLetter,
            NumberFormat::LowerLetter,
            NumberFormat::Ordinal,
            NumberFormat::CardinalText,
            NumberFormat::OrdinalText,
            NumberFormat::Hex,
            NumberFormat::Chicago,
            NumberFormat::IdeographDigital,
            NumberFormat::JapaneseCounting,
            NumberFormat::Aiueo,
            NumberFormat::Iroha,
            NumberFormat::DecimalFullWidth,
            NumberFormat::DecimalHalfWidth,
            NumberFormat::JapaneseLegal,
            NumberFormat::JapaneseDigitalTenThousand,
            NumberFormat::DecimalEnclosedCircle,
            NumberFormat::DecimalFullWidth2,
            NumberFormat::AiueoFullWidth,
            NumberFormat::IrohaFullWidth,
            NumberFormat::DecimalZero,
            NumberFormat::Bullet,
            NumberFormat::Ganada,
            NumberFormat::Chosung,
            NumberFormat::DecimalEnclosedFullstop,
            NumberFormat::DecimalEnclosedParen,
            NumberFormat::DecimalEnclosedCircleChinese,
            NumberFormat::IdeographEnclosedCircle,
            NumberFormat::IdeographTraditional,
            NumberFormat::IdeographZodiac,
            NumberFormat::IdeographZodiacTraditional,
            NumberFormat::TaiwaneseCounting,
            NumberFormat::IdeographLegalTraditional,
            NumberFormat::TaiwaneseCountingThousand,
            NumberFormat::TaiwaneseDigital,
            NumberFormat::ChineseCounting,
            NumberFormat::ChineseLegalSimplified,
            NumberFormat::ChineseCountingThousand,
            NumberFormat::DecimalChinese,
            NumberFormat::KoreanDigital,
            NumberFormat::KoreanCounting,
            NumberFormat::KoreanLegal,
            NumberFormat::KoreanDigital2,
            NumberFormat::Hebrew1,
            NumberFormat::ArabicAlpha,
            NumberFormat::Hebrew2,
            NumberFormat::ArabicAbjad,
            NumberFormat::HindiVowels,
            NumberFormat::HindiConsonants,
            NumberFormat::HindiNumbers,
            NumberFormat::HindiCounting,
            NumberFormat::ThaiLetters,
            NumberFormat::ThaiNumbers,
            NumberFormat::ThaiCounting,
            NumberFormat::VietnameseCounting,
            NumberFormat::NumberInDash,
            NumberFormat::RussianLower,
            NumberFormat::RussianUpper,
        ];
        match value {
            0..=59 => Ok(VALUES[usize::from(value)]),
            255 => Ok(NumberFormat::None),
            invalid => Err(invalid),
        }
    }
}

/// Alignment for list numbers/bullets
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListAlignment {
    Left = 0,
    Center = 1,
    Right = 2,
}

impl TryFrom<u8> for ListAlignment {
    type Error = u8;

    fn try_from(value: u8) -> std::result::Result<Self, Self::Error> {
        match value {
            0 => Ok(ListAlignment::Left),
            1 => Ok(ListAlignment::Center),
            2 => Ok(ListAlignment::Right),
            invalid => Err(invalid),
        }
    }
}

/// List level format (LVLF structure)
#[derive(Debug, Clone)]
pub struct ListLevel {
    /// Start-at value
    pub start_at: u32,
    /// Number format
    pub number_format: NumberFormat,
    /// Alignment
    pub alignment: ListAlignment,
    /// Level number (0-8)
    pub level: u8,
    /// Follow character after number (tab, space, nothing)
    pub follow_char: u8,
    /// Indentation in twips
    pub indent_left: i32,
    /// Hanging indent in twips
    pub indent_hanging: i32,
    /// Number text (format string with placeholders)
    pub number_text: String,
}

impl ListLevel {
    /// Parse a complete LVL structure.
    pub fn from_bytes(data: &[u8], level: u8) -> Result<Self> {
        Self::parse_with_size(data, level).map(|(level, _)| level)
    }

    fn parse_with_size(data: &[u8], level: u8) -> Result<(Self, usize)> {
        if data.len() < 28 {
            return Err(DocError::InvalidFormat("LVLF too short".to_string()));
        }
        if level > 8 {
            return Err(DocError::InvalidFormat(
                "list level index exceeds 8".to_string(),
            ));
        }

        let start_at = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read start_at: {}", e)))?;
        let number_format = NumberFormat::try_from(data[4]).map_err(|invalid| {
            DocError::InvalidFormat(format!("LVLF has invalid MSONFC value {invalid:#04x}"))
        })?;
        if matches!(
            number_format,
            NumberFormat::Hex
                | NumberFormat::Chicago
                | NumberFormat::DecimalHalfWidth
                | NumberFormat::DecimalFullWidth2
        ) {
            return Err(DocError::InvalidFormat(format!(
                "LVLF forbids MSONFC value {:#04x}",
                number_format as u8
            )));
        }
        if number_format != NumberFormat::Bullet
            && number_format != NumberFormat::None
            && start_at > 0x7FFF
        {
            return Err(DocError::InvalidFormat(format!(
                "LVLF start value {start_at} exceeds 32767"
            )));
        }
        let alignment = ListAlignment::try_from(data[5] & 0x03).map_err(|invalid| {
            DocError::InvalidFormat(format!("LVLF has invalid alignment {invalid}"))
        })?;
        let follow_char = data[15];
        if follow_char > 2 {
            return Err(DocError::InvalidFormat(format!(
                "LVLF has invalid follow character {follow_char}"
            )));
        }
        let indent_left = binary::read_i32_le(data, 16)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read indent_left: {}", e)))?;
        let cb_chpx = data[24] as usize;
        let cb_papx = data[25] as usize;
        let text_offset = 28usize
            .checked_add(cb_papx)
            .and_then(|offset| offset.checked_add(cb_chpx))
            .ok_or_else(|| DocError::InvalidFormat("LVL size overflows".to_string()))?;
        let cch_end = text_offset
            .checked_add(2)
            .ok_or_else(|| DocError::InvalidFormat("LVL XST offset overflows".to_string()))?;
        if cch_end > data.len() {
            return Err(DocError::InvalidFormat(
                "LVL is missing its XST length".to_string(),
            ));
        }
        let text_len = binary::read_u16_le(data, text_offset)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read XST length: {e}")))?
            as usize;
        let text_bytes_len = text_len
            .checked_mul(2)
            .ok_or_else(|| DocError::InvalidFormat("LVL XST size overflows".to_string()))?;
        let total_size = cch_end
            .checked_add(text_bytes_len)
            .ok_or_else(|| DocError::InvalidFormat("LVL size overflows".to_string()))?;
        if total_size > data.len() {
            return Err(DocError::InvalidFormat(
                "LVL XST extends beyond the table stream".to_string(),
            ));
        }

        let text_bytes = &data[cch_end..total_size];
        let mut text_units = Vec::with_capacity(text_len);
        for chunk in text_bytes.chunks_exact(2) {
            text_units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        let mut number_text = String::new();
        for decoded in std::char::decode_utf16(text_units) {
            match decoded.unwrap_or(char::REPLACEMENT_CHARACTER) {
                placeholder @ '\0'..='\u{8}' => {
                    number_text.push('%');
                    number_text.push(char::from(b'1' + placeholder as u8));
                },
                ch => number_text.push(ch),
            }
        }

        Ok((
            Self {
                start_at,
                number_format,
                alignment,
                level,
                follow_char,
                indent_left,
                indent_hanging: 0,
                number_text,
            },
            total_size,
        ))
    }

    /// Check if this is a bullet list
    pub fn is_bullet(&self) -> bool {
        self.number_format == NumberFormat::Bullet
    }

    /// Check if this is a numbered list
    pub fn is_numbered(&self) -> bool {
        !self.is_bullet() && self.number_format != NumberFormat::None
    }
}

/// List structure (LST - List Structure)
#[derive(Debug, Clone)]
pub struct ListStructure {
    /// List ID (lsid)
    pub list_id: u32,
    /// Template ID (tplc)
    pub template_id: u32,
    /// Simple list flag
    pub is_simple: bool,
    /// List levels (up to 9 levels)
    pub levels: Vec<ListLevel>,
}

impl ListStructure {
    /// Parse a list structure from LST
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() < 28 {
            return Err(DocError::InvalidFormat("LST too short".to_string()));
        }

        let list_id = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read list_id: {}", e)))?;
        let template_id = binary::read_u32_le(data, 4)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read template_id: {}", e)))?;

        // Flags byte at offset 26
        let flags = data[26];
        let is_simple = (flags & 0x01) != 0;

        Ok(Self {
            list_id,
            template_id,
            is_simple,
            levels: Vec::new(),
        })
    }

    /// Get a specific level
    pub fn level(&self, level: u8) -> Option<&ListLevel> {
        self.levels.get(level as usize)
    }
}

/// List Format Override (LFO structure)
#[derive(Debug, Clone)]
pub struct ListFormatOverride {
    /// List ID this override applies to
    pub list_id: u32,
    /// Override count
    pub override_count: u8,
    /// LFO ID (used to reference this from paragraphs)
    pub lfo_id: u32,
}

impl ListFormatOverride {
    /// Parse an LFO structure (16 bytes).
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        Self::from_bytes_with_id(data, 1)
    }

    fn from_bytes_with_id(data: &[u8], lfo_id: u32) -> Result<Self> {
        if data.len() < 16 {
            return Err(DocError::InvalidFormat("LFO too short".to_string()));
        }

        let list_id = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read list_id: {}", e)))?;
        let override_count = data[12];
        if override_count > 9 {
            return Err(DocError::InvalidFormat(
                "LFO override count exceeds 9".to_string(),
            ));
        }

        Ok(Self {
            list_id,
            override_count,
            lfo_id,
        })
    }
}

/// List tables parser
pub struct ListTables {
    /// All list structures
    list_structures: Vec<ListStructure>,
    /// All list format overrides
    list_overrides: Vec<ListFormatOverride>,
}

impl ListTables {
    /// Parse list tables from the table stream
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block
    /// * `table_stream` - The table stream (0Table or 1Table)
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut list_structures = Vec::new();
        let mut list_overrides = Vec::new();

        // Parse PlfLst (List Table) - FibRgFcLcb97 index 73.
        if let Some((offset, length)) = fib.get_table_pointer(73)
            && length > 0
            && (offset as usize) < table_stream.len()
        {
            let offset = offset as usize;
            let header_end = offset.checked_add(length as usize).ok_or_else(|| {
                DocError::InvalidFormat("PlfLst table range overflows".to_string())
            })?;
            if header_end > table_stream.len() {
                return Err(DocError::InvalidFormat(
                    "PlfLst header extends beyond the table stream".to_string(),
                ));
            }
            let level_end = fib
                .get_table_pointer(74)
                .map(|(lfo_offset, _)| lfo_offset as usize)
                .filter(|&lfo_offset| lfo_offset >= header_end)
                .unwrap_or(table_stream.len());
            if level_end > table_stream.len() {
                return Err(DocError::InvalidFormat(
                    "PlfLst level range extends beyond the table stream".to_string(),
                ));
            }

            list_structures = Self::parse_plflst(
                &table_stream[offset..header_end],
                &table_stream[header_end..level_end],
            )?;
        }

        // Parse PlfLfo (List Format Override Table) - FibRgFcLcb97 index 74.
        if let Some((offset, length)) = fib.get_table_pointer(74)
            && length > 0
            && (offset as usize) < table_stream.len()
        {
            let plf_data = &table_stream[offset as usize..];
            let plf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

            list_overrides = Self::parse_plflfo(&plf_data[..plf_len])?;
        }

        Ok(Self {
            list_structures,
            list_overrides,
        })
    }

    /// Parse PlfLst (List Table)
    fn parse_plflst(header_data: &[u8], level_data: &[u8]) -> Result<Vec<ListStructure>> {
        if header_data.len() < 2 {
            return Err(DocError::InvalidFormat("PlfLst is too short".to_string()));
        }

        let count = binary::read_u16_le(header_data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read count: {}", e)))?
            as usize;
        let expected_header_len = 2usize
            .checked_add(count.checked_mul(28).ok_or_else(|| {
                DocError::InvalidFormat("PlfLst structure count overflows".to_string())
            })?)
            .ok_or_else(|| DocError::InvalidFormat("PlfLst size overflows".to_string()))?;
        if header_data.len() != expected_header_len {
            return Err(DocError::InvalidFormat(format!(
                "PlfLst header length is {}, expected {expected_header_len}",
                header_data.len()
            )));
        }

        let mut structures = Vec::with_capacity(count);
        for index in 0..count {
            let offset = 2 + index * 28;
            structures.push(ListStructure::from_bytes(
                &header_data[offset..offset + 28],
            )?);
        }

        let mut level_offset = 0usize;
        for structure in &mut structures {
            let level_count = if structure.is_simple { 1 } else { 9 };
            structure.levels.reserve(level_count);
            for level in 0..level_count {
                let (parsed, size) = ListLevel::parse_with_size(
                    level_data.get(level_offset..).ok_or_else(|| {
                        DocError::InvalidFormat("PlfLst LVL offset is invalid".to_string())
                    })?,
                    level as u8,
                )?;
                level_offset = level_offset.checked_add(size).ok_or_else(|| {
                    DocError::InvalidFormat("PlfLst LVL size overflows".to_string())
                })?;
                structure.levels.push(parsed);
            }
        }

        Ok(structures)
    }

    /// Parse PlfLfo (List Format Override Table)
    fn parse_plflfo(data: &[u8]) -> Result<Vec<ListFormatOverride>> {
        if data.len() < 4 {
            return Ok(Vec::new());
        }

        let count = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read count: {}", e)))?
            as usize;
        let mut overrides = Vec::with_capacity(count);
        let lfo_bytes = count
            .checked_mul(16)
            .ok_or_else(|| DocError::InvalidFormat("PlfLfo count overflows".to_string()))?;
        let lfo_data_start = 4usize
            .checked_add(lfo_bytes)
            .ok_or_else(|| DocError::InvalidFormat("PlfLfo size overflows".to_string()))?;
        if lfo_data_start > data.len() {
            return Err(DocError::InvalidFormat(
                "PlfLfo LFO array is truncated".to_string(),
            ));
        }
        let mut offset = 4;

        for index in 0..count {
            overrides.push(ListFormatOverride::from_bytes_with_id(
                &data[offset..offset + 16],
                u32::try_from(index + 1)
                    .map_err(|_| DocError::InvalidFormat("PlfLfo index exceeds u32".to_string()))?,
            )?);
            offset += 16;
        }

        let mut data_offset = lfo_data_start;
        for lfo in &overrides {
            data_offset = data_offset.checked_add(4).ok_or_else(|| {
                DocError::InvalidFormat("PlfLfo LFOData size overflows".to_string())
            })?;
            if data_offset > data.len() {
                return Err(DocError::InvalidFormat(
                    "PlfLfo LFOData array is truncated".to_string(),
                ));
            }
            for _ in 0..lfo.override_count {
                let base_end = data_offset
                    .checked_add(8)
                    .ok_or_else(|| DocError::InvalidFormat("LFOLVL size overflows".to_string()))?;
                if base_end > data.len() {
                    return Err(DocError::InvalidFormat("LFOLVL is truncated".to_string()));
                }
                let flags = binary::read_u32_le(data, data_offset + 4).map_err(|e| {
                    DocError::InvalidFormat(format!("Failed to read LFOLVL flags: {e}"))
                })?;
                data_offset = base_end;
                if flags & 0x20 != 0 {
                    let (_, size) = ListLevel::parse_with_size(&data[data_offset..], 0)?;
                    data_offset = data_offset.checked_add(size).ok_or_else(|| {
                        DocError::InvalidFormat("LFOLVL formatting size overflows".to_string())
                    })?;
                }
            }
        }
        if data_offset != data.len() {
            return Err(DocError::InvalidFormat(format!(
                "PlfLfo has {} trailing bytes",
                data.len() - data_offset
            )));
        }

        Ok(overrides)
    }

    /// Get all list structures
    pub fn structures(&self) -> &[ListStructure] {
        &self.list_structures
    }

    /// Get all list format overrides
    pub fn overrides(&self) -> &[ListFormatOverride] {
        &self.list_overrides
    }

    /// Find a list structure by ID
    pub fn find_structure(&self, list_id: u32) -> Option<&ListStructure> {
        self.list_structures
            .iter()
            .find(|lst| lst.list_id == list_id)
    }

    /// Find a list override by LFO ID
    pub fn find_override(&self, lfo_id: u32) -> Option<&ListFormatOverride> {
        self.list_overrides.iter().find(|lfo| lfo.lfo_id == lfo_id)
    }

    /// Get the list structure for a given LFO ID
    pub fn get_list_for_lfo(&self, lfo_id: u32) -> Option<&ListStructure> {
        self.find_override(lfo_id)
            .and_then(|lfo| self.find_structure(lfo.list_id))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_format() {
        assert_eq!(NumberFormat::try_from(0), Ok(NumberFormat::Arabic));
        assert_eq!(NumberFormat::try_from(23), Ok(NumberFormat::Bullet));
        assert_eq!(NumberFormat::try_from(255), Ok(NumberFormat::None));
    }

    #[test]
    fn test_list_alignment() {
        assert_eq!(ListAlignment::try_from(0), Ok(ListAlignment::Left));
        assert_eq!(ListAlignment::try_from(1), Ok(ListAlignment::Center));
        assert_eq!(ListAlignment::try_from(2), Ok(ListAlignment::Right));
        assert_eq!(ListAlignment::try_from(3), Err(3));
    }

    #[test]
    fn test_number_format_all_variants() {
        for value in 0..=59 {
            assert_eq!(NumberFormat::try_from(value).unwrap() as u8, value);
        }
        assert_eq!(NumberFormat::try_from(255), Ok(NumberFormat::None));
    }

    #[test]
    fn test_number_format_rejects_unknown() {
        for value in 60..=254 {
            assert_eq!(NumberFormat::try_from(value), Err(value));
        }
    }

    #[test]
    fn test_number_format_clone() {
        let fmt = NumberFormat::Bullet;
        let cloned = fmt;
        assert_eq!(fmt, cloned);
    }

    #[test]
    fn test_number_format_copy() {
        let fmt = NumberFormat::UpperRoman;
        let copied = fmt;
        assert_eq!(fmt, copied);
    }

    #[test]
    fn test_number_format_debug() {
        let fmt = NumberFormat::Arabic;
        let debug_str = format!("{:?}", fmt);
        assert!(debug_str.contains("Arabic"));
    }

    #[test]
    fn test_number_format_equality() {
        assert_eq!(NumberFormat::Arabic, NumberFormat::Arabic);
        assert_ne!(NumberFormat::Arabic, NumberFormat::Bullet);
    }

    #[test]
    fn test_list_alignment_rejects_unknown() {
        assert_eq!(ListAlignment::try_from(3), Err(3));
        assert_eq!(ListAlignment::try_from(100), Err(100));
    }

    #[test]
    fn test_list_alignment_clone() {
        let align = ListAlignment::Center;
        let cloned = align;
        assert_eq!(align, cloned);
    }

    #[test]
    fn test_list_level_creation() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "%1.".to_string(),
        };

        assert_eq!(level.start_at, 1);
        assert_eq!(level.number_format, NumberFormat::Arabic);
        assert_eq!(level.alignment, ListAlignment::Left);
        assert_eq!(level.level, 0);
        assert_eq!(level.indent_left, 720);
        assert_eq!(level.indent_hanging, 360);
        assert_eq!(level.number_text, "%1.");
        assert!(level.is_numbered());
        assert!(!level.is_bullet());
    }

    #[test]
    fn test_list_level_bullet() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Bullet,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "\u{2022}".to_string(),
        };

        assert!(level.is_bullet());
        assert!(!level.is_numbered());
    }

    #[test]
    fn test_list_level_none() {
        let level = ListLevel {
            start_at: 0,
            number_format: NumberFormat::None,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 0,
            indent_hanging: 0,
            number_text: String::new(),
        };

        assert!(!level.is_bullet());
        assert!(!level.is_numbered());
    }

    #[test]
    fn test_list_level_clone() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::LowerRoman,
            alignment: ListAlignment::Right,
            level: 2,
            follow_char: 1,
            indent_left: 1440,
            indent_hanging: 720,
            number_text: "(%2)".to_string(),
        };
        let cloned = level.clone();

        assert_eq!(cloned.start_at, level.start_at);
        assert_eq!(cloned.number_format, level.number_format);
        assert_eq!(cloned.alignment, level.alignment);
        assert_eq!(cloned.level, level.level);
        assert_eq!(cloned.number_text, level.number_text);
    }

    #[test]
    fn test_list_level_debug() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "%1.".to_string(),
        };
        let debug_str = format!("{:?}", level);
        assert!(debug_str.contains("ListLevel"));
        assert!(debug_str.contains("Arabic"));
    }

    #[test]
    fn test_list_structure_creation() {
        let levels = vec![ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "%1.".to_string(),
        }];

        let lst = ListStructure {
            list_id: 12345,
            template_id: 67890,
            is_simple: false,
            levels,
        };

        assert_eq!(lst.list_id, 12345);
        assert_eq!(lst.template_id, 67890);
        assert!(!lst.is_simple);
        assert_eq!(lst.levels.len(), 1);
    }

    #[test]
    fn test_list_structure_simple() {
        let lst = ListStructure {
            list_id: 1,
            template_id: 1,
            is_simple: true,
            levels: Vec::new(),
        };

        assert!(lst.is_simple);
    }

    #[test]
    fn test_list_structure_level_accessor() {
        let levels = vec![
            ListLevel {
                start_at: 1,
                number_format: NumberFormat::Arabic,
                alignment: ListAlignment::Left,
                level: 0,
                follow_char: 0,
                indent_left: 720,
                indent_hanging: 360,
                number_text: "%1.".to_string(),
            },
            ListLevel {
                start_at: 1,
                number_format: NumberFormat::LowerLetter,
                alignment: ListAlignment::Left,
                level: 1,
                follow_char: 0,
                indent_left: 1440,
                indent_hanging: 360,
                number_text: "%1.%2.".to_string(),
            },
        ];

        let lst = ListStructure {
            list_id: 1,
            template_id: 1,
            is_simple: false,
            levels,
        };

        assert!(lst.level(0).is_some());
        assert!(lst.level(1).is_some());
        assert!(lst.level(2).is_none());
        assert_eq!(lst.level(0).unwrap().number_format, NumberFormat::Arabic);
        assert_eq!(
            lst.level(1).unwrap().number_format,
            NumberFormat::LowerLetter
        );
    }

    #[test]
    fn test_list_structure_clone() {
        let lst = ListStructure {
            list_id: 100,
            template_id: 200,
            is_simple: false,
            levels: vec![ListLevel {
                start_at: 1,
                number_format: NumberFormat::Bullet,
                alignment: ListAlignment::Left,
                level: 0,
                follow_char: 0,
                indent_left: 720,
                indent_hanging: 360,
                number_text: "\u{2022}".to_string(),
            }],
        };
        let cloned = lst.clone();

        assert_eq!(cloned.list_id, lst.list_id);
        assert_eq!(cloned.template_id, lst.template_id);
        assert_eq!(cloned.levels.len(), lst.levels.len());
    }

    #[test]
    fn test_list_structure_debug() {
        let lst = ListStructure {
            list_id: 1,
            template_id: 2,
            is_simple: false,
            levels: Vec::new(),
        };
        let debug_str = format!("{:?}", lst);
        assert!(debug_str.contains("ListStructure"));
    }

    #[test]
    fn test_list_format_override_creation() {
        let lfo = ListFormatOverride {
            list_id: 12345,
            override_count: 1,
            lfo_id: 1,
        };

        assert_eq!(lfo.list_id, 12345);
        assert_eq!(lfo.override_count, 1);
        assert_eq!(lfo.lfo_id, 1);
    }

    #[test]
    fn test_list_format_override_clone() {
        let lfo = ListFormatOverride {
            list_id: 100,
            override_count: 2,
            lfo_id: 5,
        };
        let cloned = lfo.clone();

        assert_eq!(cloned.list_id, lfo.list_id);
        assert_eq!(cloned.override_count, lfo.override_count);
        assert_eq!(cloned.lfo_id, lfo.lfo_id);
    }

    #[test]
    fn test_list_format_override_debug() {
        let lfo = ListFormatOverride {
            list_id: 1,
            override_count: 0,
            lfo_id: 1,
        };
        let debug_str = format!("{:?}", lfo);
        assert!(debug_str.contains("ListFormatOverride"));
    }

    #[test]
    fn test_list_tables_empty() {
        let tables = ListTables {
            list_structures: Vec::new(),
            list_overrides: Vec::new(),
        };

        assert!(tables.structures().is_empty());
        assert!(tables.overrides().is_empty());
    }

    #[test]
    fn test_list_tables_with_data() {
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 1,
                template_id: 1,
                is_simple: false,
                levels: Vec::new(),
            }],
            list_overrides: vec![ListFormatOverride {
                list_id: 1,
                override_count: 0,
                lfo_id: 1,
            }],
        };

        assert_eq!(tables.structures().len(), 1);
        assert_eq!(tables.overrides().len(), 1);
    }

    #[test]
    fn test_list_tables_find_structure() {
        let tables = ListTables {
            list_structures: vec![
                ListStructure {
                    list_id: 100,
                    template_id: 1,
                    is_simple: false,
                    levels: Vec::new(),
                },
                ListStructure {
                    list_id: 200,
                    template_id: 2,
                    is_simple: true,
                    levels: Vec::new(),
                },
            ],
            list_overrides: Vec::new(),
        };

        assert!(tables.find_structure(100).is_some());
        assert!(tables.find_structure(200).is_some());
        assert!(tables.find_structure(999).is_none());
    }

    #[test]
    fn test_list_tables_find_override() {
        let tables = ListTables {
            list_structures: Vec::new(),
            list_overrides: vec![
                ListFormatOverride {
                    list_id: 1,
                    override_count: 0,
                    lfo_id: 10,
                },
                ListFormatOverride {
                    list_id: 2,
                    override_count: 1,
                    lfo_id: 20,
                },
            ],
        };

        assert!(tables.find_override(10).is_some());
        assert!(tables.find_override(20).is_some());
        assert!(tables.find_override(999).is_none());
    }

    #[test]
    fn test_list_tables_get_list_for_lfo() {
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 100,
                template_id: 1,
                is_simple: false,
                levels: Vec::new(),
            }],
            list_overrides: vec![ListFormatOverride {
                list_id: 100,
                override_count: 0,
                lfo_id: 1,
            }],
        };

        let lst = tables.get_list_for_lfo(1);
        assert!(lst.is_some());
        assert_eq!(lst.unwrap().list_id, 100);

        assert!(tables.get_list_for_lfo(999).is_none());
    }

    #[test]
    fn test_list_tables_get_list_for_lfo_no_override() {
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 100,
                template_id: 1,
                is_simple: false,
                levels: Vec::new(),
            }],
            list_overrides: Vec::new(),
        };

        assert!(tables.get_list_for_lfo(1).is_none());
    }

    #[test]
    fn test_list_level_from_bytes_too_short() {
        let data = vec![0u8; 10];
        let result = ListLevel::from_bytes(&data, 0);
        assert!(result.is_err());
    }

    #[test]
    fn test_list_level_from_bytes_minimal() {
        // A minimal LVL is a 28-byte LVLF plus an empty two-byte XST.
        let mut data = vec![0u8; 30];
        // start_at at offset 0
        data[0] = 1; // start_at = 1
        // number_format at offset 4
        data[4] = 0; // Arabic
        // alignment at offset 5
        data[5] = 0; // Left
        // follow_char at offset 15
        data[15] = 0;
        // dxaIndentSav at offset 16
        data[16] = 0xD0; // 720 in little-endian
        data[17] = 0x02;

        let result = ListLevel::from_bytes(&data, 0);
        assert!(result.is_ok());

        let level = result.unwrap();
        assert_eq!(level.start_at, 1);
        assert_eq!(level.number_format, NumberFormat::Arabic);
        assert_eq!(level.alignment, ListAlignment::Left);
        assert_eq!(level.level, 0);
        assert_eq!(level.indent_left, 720);
        assert_eq!(level.indent_hanging, 0);
        assert_eq!(level.number_text, "");
    }

    #[test]
    fn list_level_preserves_exotic_msonfc_and_rejects_reserved_values() {
        let mut data = vec![0u8; 30];
        data[4] = NumberFormat::RussianUpper as u8;
        let level = ListLevel::from_bytes(&data, 0).unwrap();
        assert_eq!(level.number_format, NumberFormat::RussianUpper);

        data[4] = 0x3C;
        assert!(ListLevel::from_bytes(&data, 0).is_err());
        data[4] = NumberFormat::Hex as u8;
        assert!(ListLevel::from_bytes(&data, 0).is_err());
        data[4] = NumberFormat::Arabic as u8;
        data[5] = 3;
        assert!(ListLevel::from_bytes(&data, 0).is_err());
        data[5] = 0;
        data[15] = 3;
        assert!(ListLevel::from_bytes(&data, 0).is_err());
        data[15] = 0;
        data[..4].copy_from_slice(&32_768u32.to_le_bytes());
        assert!(ListLevel::from_bytes(&data, 0).is_err());
    }

    #[test]
    fn test_list_level_from_bytes_bullet() {
        let mut data = vec![0u8; 32];
        data[4] = 23; // Bullet format
        data[28..30].copy_from_slice(&1u16.to_le_bytes());
        data[30..32].copy_from_slice(&0x2022u16.to_le_bytes());

        let level = ListLevel::from_bytes(&data, 0).unwrap();
        assert!(level.is_bullet());
        assert!(!level.is_numbered());
    }

    #[test]
    fn test_list_level_from_bytes_with_text() {
        let mut data = vec![0u8; 34];
        // Fixed part
        data[0] = 1; // start_at
        data[4] = 0; // Arabic
        data[5] = 0; // Left
        data[15] = 0; // follow_char
        data[28..30].copy_from_slice(&2u16.to_le_bytes());
        data[30..32].copy_from_slice(&0u16.to_le_bytes()); // level 0 placeholder
        data[32..34].copy_from_slice(&('.' as u16).to_le_bytes());

        let level = ListLevel::from_bytes(&data, 0).unwrap();
        assert_eq!(level.number_text, "%1.");
    }

    #[test]
    fn test_list_structure_from_bytes_too_short() {
        let data = vec![0u8; 10];
        let result = ListStructure::from_bytes(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_list_structure_from_bytes_minimal() {
        let mut data = vec![0u8; 28];
        // list_id at offset 0
        data[0] = 0x39; // 57 in little-endian
        data[1] = 0x00;
        data[2] = 0x00;
        data[3] = 0x00;
        // template_id at offset 4
        data[4] = 0x30; // 48 in little-endian
        data[5] = 0x00;
        // flags at offset 26 - simple flag
        data[26] = 0x01; // is_simple = true

        let result = ListStructure::from_bytes(&data);
        assert!(result.is_ok());

        let lst = result.unwrap();
        assert_eq!(lst.list_id, 57);
        assert_eq!(lst.template_id, 48);
        assert!(lst.is_simple);
        assert!(lst.levels.is_empty());
    }

    #[test]
    fn test_list_format_override_from_bytes_too_short() {
        let data = vec![0u8; 5];
        let result = ListFormatOverride::from_bytes(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_list_format_override_from_bytes_valid() {
        let mut data = vec![0u8; 16];
        // list_id at offset 0
        data[0] = 0x39;
        data[1] = 0x00;
        data[2] = 0x00;
        data[3] = 0x00;
        // override_count at offset 12
        data[12] = 2;

        let result = ListFormatOverride::from_bytes(&data);
        assert!(result.is_ok());

        let lfo = result.unwrap();
        assert_eq!(lfo.list_id, 57);
        assert_eq!(lfo.override_count, 2);
    }

    #[test]
    fn test_numbering_with_unicode_number_text() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Bullet,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "\u{2022} \u{25ba} \u{2192}".to_string(), // bullet, pointer, arrow
        };

        assert_eq!(level.number_text, "\u{2022} \u{25ba} \u{2192}");
    }

    #[test]
    fn test_list_level_negative_indent() {
        let mut data = vec![0u8; 30];
        // dxaIndentSav at offset 16 (signed 32-bit)
        data[16] = 0xF0; // -16 in little-endian two's complement
        data[17] = 0xFF;
        data[18] = 0xFF;
        data[19] = 0xFF;

        let level = ListLevel::from_bytes(&data, 0).unwrap();
        assert_eq!(level.indent_left, -16);
    }

    #[test]
    fn parses_split_plflst_header_and_level_array() {
        let mut writer = crate::doc::writer::numbering::NumberingWriter::new();
        let mut list = crate::doc::writer::numbering::ListStructure::new(42);
        let mut first = crate::doc::writer::numbering::ListLevel::new(
            3,
            crate::doc::writer::numbering::NumberFormat::Decimal,
        );
        first.number_text = "%1.😀".to_string();
        list.add_level(first);
        list.add_level(crate::doc::writer::numbering::ListLevel::new(
            1,
            crate::doc::writer::numbering::NumberFormat::LowerLetter,
        ));
        writer.add_list(list);
        let (header, levels) = writer.build_plflst().unwrap();

        let parsed = ListTables::parse_plflst(&header, &levels).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].list_id, 42);
        assert!(!parsed[0].is_simple);
        assert_eq!(parsed[0].levels.len(), 9);
        assert_eq!(parsed[0].levels[0].start_at, 3);
        assert_eq!(parsed[0].levels[0].number_text, "%1.😀");
        assert_eq!(parsed[0].levels[1].number_format, NumberFormat::LowerLetter);
    }

    #[test]
    fn parses_parallel_lfo_and_lfo_data_arrays() {
        let mut writer = crate::doc::writer::numbering::NumberingWriter::new();
        writer.add_override(crate::doc::writer::numbering::ListFormatOverride::new(
            100, 1,
        ));
        writer.add_override(crate::doc::writer::numbering::ListFormatOverride::new(
            200, 2,
        ));

        let parsed = ListTables::parse_plflfo(&writer.build_plflfo()).unwrap();
        assert_eq!(parsed.len(), 2);
        assert_eq!((parsed[0].list_id, parsed[0].lfo_id), (100, 1));
        assert_eq!((parsed[1].list_id, parsed[1].lfo_id), (200, 2));
    }

    #[test]
    fn rejects_truncated_list_tables() {
        assert!(ListTables::parse_plflst(&[1, 0], &[]).is_err());
        assert!(ListTables::parse_plflst(&[0, 0, 0], &[]).is_err());

        let mut truncated_lfo = vec![0u8; 20];
        truncated_lfo[..4].copy_from_slice(&1u32.to_le_bytes());
        assert!(ListTables::parse_plflfo(&truncated_lfo).is_err());
    }
}
