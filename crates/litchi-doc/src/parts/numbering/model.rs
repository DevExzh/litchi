use super::validation;
/// Numbering and list structures parser for Word binary format.
///
/// Based on Apache POI's ListTables and LibreOffice's implementation.
/// Lists in DOC files are defined by:
/// - List Format Override (LFO) structures
/// - List Format (LF) structures
/// - List Level Format (LVL) structures
use crate::package::{Error as PackageError, Result};
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
    pub(super) definitions: Vec<ListStructureMetadata>,
    pub(super) levels: Vec<Vec<ListLevelMetadata>>,
    pub(super) overrides: Vec<ListFormatOverrideMetadata>,
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

    pub(super) fn parse_with_size(data: &[u8], level: u8) -> Result<(Self, usize)> {
        if data.len() < 28 {
            return Err(PackageError::InvalidFormat("LVLF too short".to_string()));
        }
        validation::level(level)?;

        let start_at = binary::read_u32_le(data, 0)
            .map_err(|e| PackageError::InvalidFormat(format!("Failed to read start_at: {}", e)))?;
        let number_format = validation::number_format(data[4])?;
        validation::start_at(number_format, start_at)?;
        let alignment = validation::alignment(data[5] & 0x03)?;
        let follow_char = data[15];
        validation::follow_character(follow_char)?;
        let indent_left = binary::read_i32_le(data, 16).map_err(|e| {
            PackageError::InvalidFormat(format!("Failed to read indent_left: {}", e))
        })?;
        let cb_chpx = data[24] as usize;
        let cb_papx = data[25] as usize;
        let text_offset = 28usize
            .checked_add(cb_papx)
            .and_then(|offset| offset.checked_add(cb_chpx))
            .ok_or_else(|| PackageError::InvalidFormat("LVL size overflows".to_string()))?;
        let cch_end = text_offset
            .checked_add(2)
            .ok_or_else(|| PackageError::InvalidFormat("LVL XST offset overflows".to_string()))?;
        if cch_end > data.len() {
            return Err(PackageError::InvalidFormat(
                "LVL is missing its XST length".to_string(),
            ));
        }
        let text_len = binary::read_u16_le(data, text_offset)
            .map_err(|e| PackageError::InvalidFormat(format!("Failed to read XST length: {e}")))?
            as usize;
        let text_bytes_len = text_len
            .checked_mul(2)
            .ok_or_else(|| PackageError::InvalidFormat("LVL XST size overflows".to_string()))?;
        let total_size = cch_end
            .checked_add(text_bytes_len)
            .ok_or_else(|| PackageError::InvalidFormat("LVL size overflows".to_string()))?;
        if total_size > data.len() {
            return Err(PackageError::InvalidFormat(
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
            return Err(PackageError::InvalidFormat("LST too short".to_string()));
        }

        let list_id = binary::read_u32_le(data, 0)
            .map_err(|e| PackageError::InvalidFormat(format!("Failed to read list_id: {}", e)))?;
        let template_id = binary::read_u32_le(data, 4).map_err(|e| {
            PackageError::InvalidFormat(format!("Failed to read template_id: {}", e))
        })?;

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

    pub(super) fn from_bytes_with_id(data: &[u8], lfo_id: u32) -> Result<Self> {
        if data.len() < 16 {
            return Err(PackageError::InvalidFormat("LFO too short".to_string()));
        }

        let list_id = binary::read_u32_le(data, 0)
            .map_err(|e| PackageError::InvalidFormat(format!("Failed to read list_id: {}", e)))?;
        let override_count = data[12];
        validation::override_count(override_count)?;

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
    pub(super) list_structures: Vec<ListStructure>,
    /// All list format overrides
    pub(super) list_overrides: Vec<ListFormatOverride>,
    /// Lossless metadata aligned with the convenience structures above.
    pub(super) metadata: ListTablesMetadata,
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
