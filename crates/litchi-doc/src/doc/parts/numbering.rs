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

/// Character emitted after a list label (`LVLF.ixchFollow`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum ListFollowCharacter {
    #[default]
    Tab = 0,
    Space = 1,
    Nothing = 2,
}

impl TryFrom<u8> for ListFollowCharacter {
    type Error = u8;

    fn try_from(value: u8) -> std::result::Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Tab),
            1 => Ok(Self::Space),
            2 => Ok(Self::Nothing),
            invalid => Err(invalid),
        }
    }
}

/// Opaque HTML-compatibility flags (`grfhic`).
///
/// The individual bits are application hints. Keeping the byte typed but
/// opaque lets readers and writers preserve values they do not interpret.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct HtmlCompatibilityFlags(u8);

impl HtmlCompatibilityFlags {
    pub const fn from_raw(raw: u8) -> Self {
        Self(raw)
    }

    pub const fn raw(self) -> u8 {
        self.0
    }
}

/// Valid paragraph-style index linked to an LSTF level.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ListStyleIndex(u16);

impl ListStyleIndex {
    pub fn new(index: u16) -> std::result::Result<Self, u16> {
        if index < 0x0FFF {
            Ok(Self(index))
        } else {
            Err(index)
        }
    }

    pub const fn get(self) -> u16 {
        self.0
    }
}

/// Field represented by an LFO (`ibstFltAutoNum`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum AutomaticNumberingField {
    #[default]
    None = 0x00,
    AutoNumberLegal = 0xFC,
    AutoNumberOutline = 0xFD,
    AutoNumber = 0xFE,
    NoneLegacy = 0xFF,
}

impl TryFrom<u8> for AutomaticNumberingField {
    type Error = u8;

    fn try_from(value: u8) -> std::result::Result<Self, Self::Error> {
        match value {
            0x00 => Ok(Self::None),
            0xFC => Ok(Self::AutoNumberLegal),
            0xFD => Ok(Self::AutoNumberOutline),
            0xFE => Ok(Self::AutoNumber),
            0xFF => Ok(Self::NoneLegacy),
            invalid => Err(invalid),
        }
    }
}

/// Lossless LSTF metadata not represented by [`ListStructure`].
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ListStructureMetadata {
    pub style_links: [Option<ListStyleIndex>; 9],
    pub automatic_numbering: bool,
    pub hybrid: bool,
    /// Ignored/reserved LSTF flag bits, retained for round trips.
    pub ignored_flags: u8,
    pub html_compatibility: HtmlCompatibilityFlags,
}

/// Lossless LVLF/LVL metadata and property payloads.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ListLevelMetadata {
    pub legal_numbering: bool,
    pub no_restart: bool,
    pub saved_indent: Option<i32>,
    /// Preserved `dxaIndentSav` value when `fIndentSav` is clear.
    pub ignored_saved_indent: i32,
    pub converted: bool,
    pub tentative: bool,
    pub ignored_flags: u8,
    pub placeholder_positions: [u8; 9],
    pub follow_character: ListFollowCharacter,
    pub unused_value: u32,
    pub restart_limit: Option<u8>,
    /// Preserved `ilvlRestartLim` value when `fNoRestart` is clear.
    pub ignored_restart_limit: u8,
    pub html_compatibility: HtmlCompatibilityFlags,
    pub paragraph_properties: Vec<u8>,
    pub number_properties: Vec<u8>,
}

/// Lossless LFOLVL flag metadata.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ListLevelOverrideMetadata {
    /// Preserved `iStartAt` value when `fStartAt` is clear.
    pub unused_start_at: u32,
    pub html_compatibility: HtmlCompatibilityFlags,
    /// Bits declared unused by the specification, retained verbatim.
    pub ignored_flags: u32,
    pub formatting: Option<ListLevelMetadata>,
}

/// Lossless LFO and corresponding LFOData metadata.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ListFormatOverrideMetadata {
    pub unused1: u32,
    pub unused2: u32,
    pub field: AutomaticNumberingField,
    pub html_compatibility: HtmlCompatibilityFlags,
    pub unused3: u8,
    pub first_paragraph_cp: Option<u32>,
    pub levels: Vec<ListLevelOverrideMetadata>,
}

/// Metadata arrays aligned with `ListTables::structures()` and `overrides()`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ListTablesMetadata {
    definitions: Vec<ListStructureMetadata>,
    levels: Vec<Vec<ListLevelMetadata>>,
    overrides: Vec<ListFormatOverrideMetadata>,
}

impl ListTablesMetadata {
    pub fn definition(&self, index: usize) -> Option<&ListStructureMetadata> {
        self.definitions.get(index)
    }

    pub fn level(&self, definition: usize, level: u8) -> Option<&ListLevelMetadata> {
        self.levels.get(definition)?.get(usize::from(level))
    }

    pub fn format_override(&self, index: usize) -> Option<&ListFormatOverrideMetadata> {
        self.overrides.get(index)
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

/// Level override carried by one `LFOLVL` structure ([MS-DOC] 2.9.133).
#[derive(Debug, Clone, Default)]
pub struct ListLevelOverride {
    /// Zero-based list level this override applies to (`iLvl`, 0..=8).
    pub level: u8,
    /// Start-at value overriding `lvlf.iStartAt` when `fStartAt` is set
    /// without `fFormatting`.
    pub start_at: Option<u32>,
    /// Complete replacement level formatting when `fFormatting` is set.
    pub format: Option<ListLevel>,
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
    /// Parsed `LFOLVL` level overrides for this LFO.
    pub level_overrides: Vec<ListLevelOverride>,
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
            level_overrides: Vec::new(),
        })
    }

    /// Get the `LFOLVL` override for a zero-based list level, if any.
    pub fn level_override(&self, level: u8) -> Option<&ListLevelOverride> {
        self.level_overrides
            .iter()
            .find(|lfolvl| lfolvl.level == level)
    }
}

/// List tables parser
pub struct ListTables {
    /// All list structures
    list_structures: Vec<ListStructure>,
    /// All list format overrides
    list_overrides: Vec<ListFormatOverride>,
    /// Lossless metadata aligned with the convenience structures above.
    metadata: ListTablesMetadata,
}

/// Borrowed resolution of a paragraph's `sprmPIlfo`/`sprmPIlvl` list binding.
///
/// This keeps the definition and override structures available to callers and
/// avoids cloning an `LVL` merely to inspect its effective formatting. A
/// start-at-only `LFOLVL` override is reported separately because all other
/// formatting continues to come from `base_level` in that case.
#[derive(Debug, Clone, Copy)]
pub struct ParagraphListBinding<'a> {
    /// One-based index into `PlfLfo`.
    pub lfo_id: u32,
    /// Zero-based list level from `sprmPIlvl`.
    pub level: u8,
    /// Whether a negative `sprmPIlfo` requests preservation of paragraph indents.
    pub preserve_indents: bool,
    /// The `LSTF` selected by the LFO's list identifier.
    pub definition: &'a ListStructure,
    /// The paragraph's selected `LFO`.
    pub format_override: &'a ListFormatOverride,
    /// The definition's level before applying an `LFOLVL`.
    pub base_level: &'a ListLevel,
    /// The level-specific `LFOLVL`, when present.
    pub level_override: Option<&'a ListLevelOverride>,
}

impl<'a> ParagraphListBinding<'a> {
    /// Effective formatting-bearing level without allocating.
    ///
    /// A formatting override replaces the complete base `LVL`; a start-at-only
    /// override leaves this reference pointing at the base level.
    pub fn effective_level(&self) -> &'a ListLevel {
        self.level_override
            .and_then(|level| level.format.as_ref())
            .unwrap_or(self.base_level)
    }

    /// Effective starting number after applying a start-at-only override.
    pub fn effective_start_at(&self) -> u32 {
        self.level_override
            .and_then(|level| level.start_at)
            .unwrap_or_else(|| self.effective_level().start_at)
    }

    /// Whether this binding replaces the complete base `LVL` formatting.
    pub fn has_formatting_override(&self) -> bool {
        self.level_override
            .is_some_and(|level| level.format.is_some())
    }

    /// Whether this binding overrides only the starting number.
    pub fn has_start_at_override(&self) -> bool {
        self.level_override
            .is_some_and(|level| level.start_at.is_some())
    }
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

        let metadata = Self::parse_metadata(fib, table_stream, &list_structures, &list_overrides)?;

        Ok(Self {
            list_structures,
            list_overrides,
            metadata,
        })
    }

    fn parse_metadata(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        structures: &[ListStructure],
        overrides: &[ListFormatOverride],
    ) -> Result<ListTablesMetadata> {
        let mut metadata = ListTablesMetadata::default();
        if let Some((offset, length)) = fib.get_table_pointer(73).filter(|(_, length)| *length > 0)
        {
            let start = usize::try_from(offset)
                .map_err(|_| DocError::InvalidFormat("PlfLst offset exceeds usize".to_string()))?;
            let header_end = start.checked_add(length as usize).ok_or_else(|| {
                DocError::InvalidFormat("PlfLst metadata range overflows".to_string())
            })?;
            let level_end = fib
                .get_table_pointer(74)
                .map(|(offset, _)| offset as usize)
                .filter(|&offset| offset >= header_end)
                .unwrap_or(table_stream.len());
            let header = table_stream.get(start..header_end).ok_or_else(|| {
                DocError::InvalidFormat("PlfLst metadata header is truncated".to_string())
            })?;
            let levels = table_stream.get(header_end..level_end).ok_or_else(|| {
                DocError::InvalidFormat("PlfLst metadata levels are truncated".to_string())
            })?;
            let mut level_offset = 0usize;
            for (index, structure) in structures.iter().enumerate() {
                let lstf = &header[2 + index * 28..2 + (index + 1) * 28];
                let mut definition = ListStructureMetadata::default();
                for style in 0..9 {
                    let raw = u16::from_le_bytes([lstf[8 + style * 2], lstf[9 + style * 2]]);
                    definition.style_links[style] = match raw {
                        0x0FFF | 0xFFFF => None,
                        value => Some(ListStyleIndex::new(value).map_err(|invalid| {
                            DocError::InvalidFormat(format!(
                                "LSTF has invalid linked style index {invalid:#06x}"
                            ))
                        })?),
                    };
                }
                let flags = lstf[26];
                definition.automatic_numbering = flags & 0x04 != 0;
                definition.hybrid = flags & 0x10 != 0;
                definition.ignored_flags = flags & 0xEA;
                definition.html_compatibility = HtmlCompatibilityFlags::from_raw(lstf[27]);
                metadata.definitions.push(definition);

                let mut level_metadata = Vec::with_capacity(structure.levels.len());
                for level in &structure.levels {
                    let (parsed, size) = Self::parse_level_metadata(
                        levels.get(level_offset..).ok_or_else(|| {
                            DocError::InvalidFormat("LVL metadata offset is invalid".to_string())
                        })?,
                        level.level,
                        level.number_format,
                    )?;
                    level_offset = level_offset.checked_add(size).ok_or_else(|| {
                        DocError::InvalidFormat("LVL metadata size overflows".to_string())
                    })?;
                    level_metadata.push(parsed);
                }
                metadata.levels.push(level_metadata);
            }
        }

        if let Some((offset, length)) = fib.get_table_pointer(74).filter(|(_, length)| *length > 0)
        {
            let start = offset as usize;
            let data = table_stream
                .get(start..start.saturating_add(length as usize))
                .ok_or_else(|| {
                    DocError::InvalidFormat("PlfLfo metadata is truncated".to_string())
                })?;
            let mut data_offset = 4usize
                .checked_add(overrides.len().checked_mul(16).ok_or_else(|| {
                    DocError::InvalidFormat("PlfLfo metadata count overflows".to_string())
                })?)
                .ok_or_else(|| DocError::InvalidFormat("PlfLfo metadata overflows".to_string()))?;
            for (index, lfo) in overrides.iter().enumerate() {
                let raw = &data[4 + index * 16..4 + (index + 1) * 16];
                let first_cp = binary::read_u32_le(data, data_offset).map_err(|e| {
                    DocError::InvalidFormat(format!("Failed to read LFOData CP: {e}"))
                })?;
                data_offset += 4;
                let mut parsed = ListFormatOverrideMetadata {
                    unused1: u32::from_le_bytes(raw[4..8].try_into().expect("LFO unused1")),
                    unused2: u32::from_le_bytes(raw[8..12].try_into().expect("LFO unused2")),
                    field: AutomaticNumberingField::try_from(raw[13]).map_err(|invalid| {
                        DocError::InvalidFormat(format!(
                            "LFO has invalid automatic-number field {invalid:#04x}"
                        ))
                    })?,
                    html_compatibility: HtmlCompatibilityFlags::from_raw(raw[14]),
                    unused3: raw[15],
                    first_paragraph_cp: (first_cp != u32::MAX).then_some(first_cp),
                    levels: Vec::with_capacity(lfo.level_overrides.len()),
                };
                for level_override in &lfo.level_overrides {
                    let flags = binary::read_u32_le(data, data_offset + 4).map_err(|e| {
                        DocError::InvalidFormat(format!("Failed to read LFOLVL flags: {e}"))
                    })?;
                    data_offset += 8;
                    let formatting = if let Some(level) = level_override.format.as_ref() {
                        let (metadata, size) = Self::parse_level_metadata(
                            &data[data_offset..],
                            level.level,
                            level.number_format,
                        )?;
                        data_offset += size;
                        Some(metadata)
                    } else {
                        None
                    };
                    parsed.levels.push(ListLevelOverrideMetadata {
                        unused_start_at: if flags & 0x10 == 0 {
                            binary::read_u32_le(data, data_offset - 8).map_err(|e| {
                                DocError::InvalidFormat(format!(
                                    "Failed to read LFOLVL ignored start: {e}"
                                ))
                            })?
                        } else {
                            0
                        },
                        html_compatibility: HtmlCompatibilityFlags::from_raw(
                            ((flags >> 6) & 0xFF) as u8,
                        ),
                        ignored_flags: flags & 0xFFFF_C000,
                        formatting,
                    });
                }
                metadata.overrides.push(parsed);
            }
        }
        Ok(metadata)
    }

    fn parse_level_metadata(
        data: &[u8],
        level: u8,
        number_format: NumberFormat,
    ) -> Result<(ListLevelMetadata, usize)> {
        if data.len() < 30 {
            return Err(DocError::InvalidFormat(
                "LVL metadata is truncated".to_string(),
            ));
        }
        let flags = data[5];
        let cb_chpx = usize::from(data[24]);
        let cb_papx = usize::from(data[25]);
        let text_offset = 28usize
            .checked_add(cb_papx)
            .and_then(|value| value.checked_add(cb_chpx))
            .ok_or_else(|| DocError::InvalidFormat("LVL metadata size overflows".to_string()))?;
        let text_len = usize::from(binary::read_u16_le(data, text_offset).map_err(|e| {
            DocError::InvalidFormat(format!("Failed to read LVL metadata XST: {e}"))
        })?);
        let total = text_offset
            .checked_add(2)
            .and_then(|value| value.checked_add(text_len.checked_mul(2)?))
            .ok_or_else(|| DocError::InvalidFormat("LVL metadata XST overflows".to_string()))?;
        if total > data.len() {
            return Err(DocError::InvalidFormat(
                "LVL metadata XST is truncated".to_string(),
            ));
        }
        let placeholders: [u8; 9] = data[6..15].try_into().expect("LVLF placeholders");
        for position in placeholders.into_iter().filter(|position| *position != 0) {
            if usize::from(position) > text_len {
                return Err(DocError::InvalidFormat(format!(
                    "LVLF placeholder position {position} exceeds XST length {text_len}"
                )));
            }
            let offset = text_offset + 2 + (usize::from(position) - 1) * 2;
            let placeholder = u16::from_le_bytes([data[offset], data[offset + 1]]);
            if placeholder > u16::from(level) {
                return Err(DocError::InvalidFormat(format!(
                    "LVL placeholder level {placeholder} exceeds level {level}"
                )));
            }
        }
        if number_format == NumberFormat::Bullet
            && (text_len != 1 || placeholders.iter().any(|position| *position != 0))
        {
            return Err(DocError::InvalidFormat(
                "bullet LVL must contain one character and no placeholders".to_string(),
            ));
        }
        let no_restart = flags & 0x08 != 0;
        let restart = data[26];
        if no_restart && restart > level {
            return Err(DocError::InvalidFormat(format!(
                "LVLF restart limit {restart} exceeds level {level}"
            )));
        }
        Ok((
            ListLevelMetadata {
                legal_numbering: flags & 0x04 != 0,
                no_restart,
                saved_indent: (flags & 0x10 != 0).then(|| {
                    i32::from_le_bytes(data[16..20].try_into().expect("LVLF saved indent"))
                }),
                ignored_saved_indent: if flags & 0x10 == 0 {
                    i32::from_le_bytes(data[16..20].try_into().expect("LVLF ignored indent"))
                } else {
                    0
                },
                converted: flags & 0x20 != 0,
                tentative: flags & 0x80 != 0,
                ignored_flags: flags & 0x40,
                placeholder_positions: placeholders,
                follow_character: ListFollowCharacter::try_from(data[15]).map_err(|invalid| {
                    DocError::InvalidFormat(format!("LVLF has invalid follow character {invalid}"))
                })?,
                unused_value: u32::from_le_bytes(data[20..24].try_into().expect("LVLF unused2")),
                restart_limit: no_restart.then_some(restart),
                ignored_restart_limit: if no_restart { 0 } else { restart },
                html_compatibility: HtmlCompatibilityFlags::from_raw(data[27]),
                paragraph_properties: data[28..28 + cb_papx].to_vec(),
                number_properties: data[28 + cb_papx..text_offset].to_vec(),
            },
            total,
        ))
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

        /// Mask for the zero-based `iLvl` field of an `LFOLVL`.
        const LFOLVL_ILVL_MASK: u32 = 0x0F;
        /// `fStartAt` flag of an `LFOLVL`.
        const LFOLVL_F_START_AT: u32 = 0x10;
        /// `fFormatting` flag of an `LFOLVL`.
        const LFOLVL_F_FORMATTING: u32 = 0x20;
        /// Maximum permitted `iStartAt` override value ([MS-DOC] 2.9.133).
        const LFOLVL_MAX_START_AT: u32 = 0x7FFF;

        let mut data_offset = lfo_data_start;
        for lfo in &mut overrides {
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
                let start_at = binary::read_u32_le(data, data_offset).map_err(|e| {
                    DocError::InvalidFormat(format!("Failed to read LFOLVL iStartAt: {e}"))
                })?;
                let flags = binary::read_u32_le(data, data_offset + 4).map_err(|e| {
                    DocError::InvalidFormat(format!("Failed to read LFOLVL flags: {e}"))
                })?;
                data_offset = base_end;
                let level = (flags & LFOLVL_ILVL_MASK) as u8;
                if level > 8 {
                    return Err(DocError::InvalidFormat(format!(
                        "LFOLVL has invalid iLvl {level}"
                    )));
                }
                let overrides_start_at = flags & LFOLVL_F_START_AT != 0;
                let overrides_formatting = flags & LFOLVL_F_FORMATTING != 0;
                let mut level_override = ListLevelOverride {
                    level,
                    start_at: None,
                    format: None,
                };
                if overrides_formatting {
                    let (parsed, size) = ListLevel::parse_with_size(&data[data_offset..], level)?;
                    data_offset = data_offset.checked_add(size).ok_or_else(|| {
                        DocError::InvalidFormat("LFOLVL formatting size overflows".to_string())
                    })?;
                    level_override.format = Some(parsed);
                } else if overrides_start_at {
                    // iStartAt is only meaningful when fFormatting is clear.
                    if start_at > LFOLVL_MAX_START_AT {
                        return Err(DocError::InvalidFormat(format!(
                            "LFOLVL start value {start_at} exceeds 32767"
                        )));
                    }
                    level_override.start_at = Some(start_at);
                }
                lfo.level_overrides.push(level_override);
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

    /// Lossless typed metadata for the list tables.
    pub fn metadata(&self) -> &ListTablesMetadata {
        &self.metadata
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

    /// Resolve the signed paragraph list reference and level without cloning.
    ///
    /// `sprmPIlfo` is one-based. Negative values select the same absolute LFO
    /// index while requesting that paragraph indents be preserved. Level 12 is
    /// the specification's "skip numbering" sentinel and does not bind a list.
    pub fn bind_paragraph(&self, signed_lfo: i16, level: u8) -> Option<ParagraphListBinding<'_>> {
        if signed_lfo == 0 || signed_lfo == i16::MIN || level > 8 {
            return None;
        }
        let lfo_id = u32::from(signed_lfo.unsigned_abs());
        let format_override = self.find_override(lfo_id)?;
        let definition = self.find_structure(format_override.list_id)?;
        let base_level = definition.level(level)?;
        Some(ParagraphListBinding {
            lfo_id,
            level,
            preserve_indents: signed_lfo.is_negative(),
            definition,
            format_override,
            base_level,
            level_override: format_override.level_override(level),
        })
    }

    /// Resolve the effective level formatting for an LFO ID and zero-based
    /// level, applying any `LFOLVL` start-at or formatting overrides.
    pub fn resolve_level(&self, lfo_id: u32, level: u8) -> Option<ListLevel> {
        let signed_lfo = i16::try_from(lfo_id).ok()?;
        let binding = self.bind_paragraph(signed_lfo, level)?;
        let mut resolved = binding.effective_level().clone();
        resolved.start_at = binding.effective_start_at();
        Some(resolved)
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
            level_overrides: Vec::new(),
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
            level_overrides: Vec::new(),
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
            level_overrides: Vec::new(),
        };
        let debug_str = format!("{:?}", lfo);
        assert!(debug_str.contains("ListFormatOverride"));
    }

    #[test]
    fn test_list_tables_empty() {
        let tables = ListTables {
            list_structures: Vec::new(),
            list_overrides: Vec::new(),
            metadata: ListTablesMetadata::default(),
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
                level_overrides: Vec::new(),
            }],
            metadata: ListTablesMetadata::default(),
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
            metadata: ListTablesMetadata::default(),
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
                    level_overrides: Vec::new(),
                },
                ListFormatOverride {
                    list_id: 2,
                    override_count: 1,
                    lfo_id: 20,
                    level_overrides: Vec::new(),
                },
            ],
            metadata: ListTablesMetadata::default(),
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
                level_overrides: Vec::new(),
            }],
            metadata: ListTablesMetadata::default(),
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
            metadata: ListTablesMetadata::default(),
        };

        assert!(tables.get_list_for_lfo(1).is_none());
    }

    #[test]
    fn paragraph_binding_borrows_base_level_and_applies_start_override() {
        let base = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 0,
            indent_hanging: 0,
            number_text: "%1.".to_string(),
        };
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 42,
                template_id: 42,
                is_simple: true,
                levels: vec![base],
            }],
            list_overrides: vec![ListFormatOverride {
                list_id: 42,
                override_count: 1,
                lfo_id: 1,
                level_overrides: vec![ListLevelOverride {
                    level: 0,
                    start_at: Some(7),
                    format: None,
                }],
            }],
            metadata: ListTablesMetadata::default(),
        };

        let binding = tables.bind_paragraph(-1, 0).unwrap();
        assert!(binding.preserve_indents);
        assert_eq!(binding.definition.list_id, 42);
        assert_eq!(binding.format_override.lfo_id, 1);
        assert!(std::ptr::eq(binding.effective_level(), binding.base_level));
        assert_eq!(binding.effective_start_at(), 7);
        assert!(binding.has_start_at_override());
        assert!(!binding.has_formatting_override());
    }

    #[test]
    fn paragraph_binding_borrows_formatting_override_and_rejects_sentinels() {
        let level = |start_at, text: &str| ListLevel {
            start_at,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 0,
            indent_hanging: 0,
            number_text: text.to_string(),
        };
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 9,
                template_id: 9,
                is_simple: true,
                levels: vec![level(1, "%1.")],
            }],
            list_overrides: vec![ListFormatOverride {
                list_id: 9,
                override_count: 1,
                lfo_id: 1,
                level_overrides: vec![ListLevelOverride {
                    level: 0,
                    start_at: None,
                    format: Some(level(3, "(%1)")),
                }],
            }],
            metadata: ListTablesMetadata::default(),
        };

        let binding = tables.bind_paragraph(1, 0).unwrap();
        let replacement = binding.level_override.unwrap().format.as_ref().unwrap();
        assert!(std::ptr::eq(binding.effective_level(), replacement));
        assert_eq!(binding.effective_start_at(), 3);
        assert!(binding.has_formatting_override());
        assert!(tables.bind_paragraph(0, 0).is_none());
        assert!(tables.bind_paragraph(i16::MIN, 0).is_none());
        assert!(tables.bind_paragraph(1, 12).is_none());
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
        let mut first = crate::doc::writer::numbering::ListLevel::new(3, NumberFormat::Decimal);
        first.number_text = "%1.😀".to_string();
        list.add_level(first);
        list.add_level(crate::doc::writer::numbering::ListLevel::new(
            1,
            NumberFormat::LowerLetter,
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
