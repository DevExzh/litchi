//! Typed Word 95/97 Document Properties extensions.
//!
//! The codecs retain the original extension bytes so fields which are undefined
//! or belong to later DOP versions survive a parse/write cycle unchanged.

use super::numbering::NumberFormat;
use std::fmt;

const DOP_BASE_SIZE: usize = 84;
const DOP95_SIZE: usize = 88;
const DOP97_SIZE: usize = 500;
const COPTS60_OFFSET: usize = 8;
const DOP97_EXTENSION_SIZE: usize = DOP97_SIZE - DOP_BASE_SIZE;

const ADT: usize = 4;
const TYPOGRAPHY: usize = 6;
const TYPOGRAPHY_SIZE: usize = 310;
const GRID: usize = 316;
const GRID_SIZE: usize = 10;
const DOCINFO5: usize = 326;
const AUTOSUMMARY: usize = 330;
const CCH_WS: usize = 342;
const CCH_WS_WITH_SUBDOCS: usize = 346;
const EVENTS: usize = 350;
const VIRUS_INFO: usize = 354;
const LIST_CACHE_CP: usize = 388;
const LAST_BULLET_LFO: usize = 392;
const LAST_NUMBER_LFO: usize = 394;
const CDBC: usize = 396;
const CDBC_WITH_SUBDOCS: usize = 400;
const FOOTNOTE_FORMAT: usize = 408;
const ENDNOTE_FORMAT: usize = 410;
const PAGE_ZOOM_FONT: usize = 412;
const PAGE_DISPLAY_WIDTH: usize = 414;

const ALLOWED_EVENT_FLAGS: u32 = 0x0000_7f3f;

/// A structural or value error in a Word 95/97 DOP extension.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DopExtensionError(String);

impl DopExtensionError {
    pub(crate) fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for DopExtensionError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for DopExtensionError {}

/// Word 95 compatibility flags. The low word is the mirrored `Copts60`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct CompatibilityOptions80(u32);

impl CompatibilityOptions80 {
    #[must_use]
    pub const fn from_raw(raw: u32) -> Self {
        Self(raw)
    }

    #[must_use]
    pub const fn raw(self) -> u32 {
        self.0
    }

    #[must_use]
    pub const fn copts60(self) -> u16 {
        self.0 as u16
    }

    #[must_use]
    pub const fn option(self, index: u8) -> bool {
        index < 16 && self.0 & (1u32 << (16 + index)) != 0
    }

    pub fn set_option(&mut self, index: u8, enabled: bool) -> Result<(), DopExtensionError> {
        if index >= 16 {
            return Err(DopExtensionError::new("Copts80 option index exceeds 15"));
        }
        let mask = 1u32 << (16 + index);
        self.0 = if enabled {
            self.0 | mask
        } else {
            self.0 & !mask
        };
        Ok(())
    }
}

/// Typed Word 95 DOP extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Dop95 {
    compatibility: CompatibilityOptions80,
}

impl Dop95 {
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP95_SIZE {
            return Err(DopExtensionError::new("Dop95 is shorter than 88 bytes"));
        }
        let base = le_u16(dop, COPTS60_OFFSET);
        let compatibility = CompatibilityOptions80::from_raw(le_u32(dop, DOP_BASE_SIZE));
        if compatibility.copts60() != base {
            return Err(DopExtensionError::new(
                "Copts80.copts60 does not mirror DopBase.copts60",
            ));
        }
        Ok(Self { compatibility })
    }

    #[must_use]
    pub const fn compatibility(self) -> CompatibilityOptions80 {
        self.compatibility
    }

    pub fn write_into(self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP95_SIZE {
            return Err(DopExtensionError::new(
                "Dop95 target is shorter than 88 bytes",
            ));
        }
        if self.compatibility.copts60() != le_u16(dop, COPTS60_OFFSET) {
            return Err(DopExtensionError::new(
                "Copts80.copts60 does not mirror DopBase.copts60",
            ));
        }
        put_u32(dop, DOP_BASE_SIZE, self.compatibility.raw());
        Ok(())
    }
}

/// Word 97 document classification (`adt`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentClassification {
    NotSpecified,
    Letter,
    Email,
}

impl DocumentClassification {
    fn parse(raw: u16) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::NotSpecified),
            1 => Ok(Self::Letter),
            2 => Ok(Self::Email),
            _ => Err(DopExtensionError::new(format!(
                "invalid Dop97 adt value {raw}"
            ))),
        }
    }

    const fn raw(self) -> u16 {
        match self {
            Self::NotSpecified => 0,
            Self::Letter => 1,
            Self::Email => 2,
        }
    }
}

/// Heading depth stored in `lvlDop`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OutlineDisplayLevel {
    ThroughHeading(u8),
    AllLevels,
}

impl OutlineDisplayLevel {
    pub fn through_heading(level: u8) -> Result<Self, DopExtensionError> {
        if level <= 9 {
            Ok(Self::ThroughHeading(level))
        } else {
            Err(DopExtensionError::new("outline heading level exceeds 9"))
        }
    }

    fn parse(raw: u8) -> Result<Self, DopExtensionError> {
        match raw {
            0..=9 => Ok(Self::ThroughHeading(raw)),
            15 => Ok(Self::AllLevels),
            _ => Err(DopExtensionError::new(format!(
                "invalid Dop97 outline display level {raw}"
            ))),
        }
    }

    const fn raw(self) -> u8 {
        match self {
            Self::ThroughHeading(level) => level,
            Self::AllLevels => 15,
        }
    }
}

/// Character-level East Asian whitespace compression.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TypographyJustification {
    DoNotCompress,
    CompressPunctuation,
    CompressPunctuationAndKana,
}

impl TypographyJustification {
    fn parse(raw: u16) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::DoNotCompress),
            1 => Ok(Self::CompressPunctuation),
            2 => Ok(Self::CompressPunctuationAndKana),
            _ => Err(DopExtensionError::new("invalid typography justification")),
        }
    }
}

/// East Asian line-breaking rule set.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum KinsokuLevel {
    Normal,
    JapaneseLevelTwo,
    Custom,
}

impl KinsokuLevel {
    fn parse(raw: u16) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::Normal),
            1 => Ok(Self::JapaneseLevelTwo),
            2 => Ok(Self::Custom),
            _ => Err(DopExtensionError::new("invalid typography kinsoku level")),
        }
    }
}

/// Language whose custom kinsoku character arrays apply.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomKinsokuLanguage {
    None,
    Japanese,
    ChineseSimplified,
    Korean,
    ChineseTraditional,
}

impl CustomKinsokuLanguage {
    fn parse(raw: u16) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::None),
            1 => Ok(Self::Japanese),
            2 => Ok(Self::ChineseSimplified),
            3 => Ok(Self::Korean),
            4 => Ok(Self::ChineseTraditional),
            _ => Err(DopExtensionError::new("invalid custom kinsoku language")),
        }
    }
}

/// Validated Word 97 East Asian typography state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentTypography {
    raw: [u8; TYPOGRAPHY_SIZE],
    pub kern_punctuation: bool,
    pub justification: TypographyJustification,
    pub kinsoku_level: KinsokuLevel,
    pub print_two_on_one: bool,
    pub custom_language: CustomKinsokuLanguage,
    pub japanese_use_level_two: bool,
    pub cannot_start_line: String,
    pub cannot_end_line: String,
}

impl DocumentTypography {
    fn parse(data: &[u8]) -> Result<Self, DopExtensionError> {
        let mut raw = [0u8; TYPOGRAPHY_SIZE];
        raw.copy_from_slice(data);
        let flags = le_u16(data, 0);
        if flags & 0xf800 != 0 {
            return Err(DopExtensionError::new(
                "DopTypography reserved bits are nonzero",
            ));
        }
        let following = le_u16(data, 2) as usize;
        let leading = le_u16(data, 4) as usize;
        if following > 100 || leading > 50 {
            return Err(DopExtensionError::new(
                "DopTypography punctuation count is out of range",
            ));
        }
        Ok(Self {
            raw,
            kern_punctuation: flags & 1 != 0,
            justification: TypographyJustification::parse((flags >> 1) & 3)?,
            kinsoku_level: KinsokuLevel::parse((flags >> 3) & 3)?,
            print_two_on_one: flags & 0x20 != 0,
            custom_language: CustomKinsokuLanguage::parse((flags >> 7) & 7)?,
            japanese_use_level_two: flags & 0x400 != 0,
            cannot_start_line: decode_utf16(&data[6..208], following)?,
            cannot_end_line: decode_utf16(&data[208..310], leading)?,
        })
    }

    pub fn set_cannot_start_line(&mut self, value: &str) -> Result<(), DopExtensionError> {
        write_utf16(
            &mut self.raw[6..208],
            &mut self.cannot_start_line,
            value,
            100,
        )
    }

    pub fn set_cannot_end_line(&mut self, value: &str) -> Result<(), DopExtensionError> {
        write_utf16(
            &mut self.raw[208..310],
            &mut self.cannot_end_line,
            value,
            50,
        )
    }

    fn encode(mut self) -> [u8; TYPOGRAPHY_SIZE] {
        let mut flags = le_u16(&self.raw, 0) & 0x40;
        flags |= u16::from(self.kern_punctuation);
        flags |= match self.justification {
            TypographyJustification::DoNotCompress => 0,
            TypographyJustification::CompressPunctuation => 1 << 1,
            TypographyJustification::CompressPunctuationAndKana => 2 << 1,
        };
        flags |= match self.kinsoku_level {
            KinsokuLevel::Normal => 0,
            KinsokuLevel::JapaneseLevelTwo => 1 << 3,
            KinsokuLevel::Custom => 2 << 3,
        };
        flags |= u16::from(self.print_two_on_one) << 5;
        flags |= match self.custom_language {
            CustomKinsokuLanguage::None => 0,
            CustomKinsokuLanguage::Japanese => 1 << 7,
            CustomKinsokuLanguage::ChineseSimplified => 2 << 7,
            CustomKinsokuLanguage::Korean => 3 << 7,
            CustomKinsokuLanguage::ChineseTraditional => 4 << 7,
        };
        flags |= u16::from(self.japanese_use_level_two) << 10;
        put_u16(&mut self.raw, 0, flags);
        put_u16(
            &mut self.raw,
            2,
            self.cannot_start_line.encode_utf16().count() as u16,
        );
        put_u16(
            &mut self.raw,
            4,
            self.cannot_end_line.encode_utf16().count() as u16,
        );
        self.raw
    }
}

/// Word 97 drawing-grid settings, in twips and grid-unit multiples.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DrawingGrid {
    pub horizontal_origin: u16,
    pub vertical_origin: u16,
    pub horizontal_spacing: u16,
    pub vertical_spacing: u16,
    pub vertical_display_every: u8,
    pub horizontal_display_every: u8,
    pub follow_margins: bool,
    unused_bit: bool,
}

impl DrawingGrid {
    fn parse(data: &[u8]) -> Result<Self, DopExtensionError> {
        let display = le_u16(data, 8);
        let vertical_display_every = (display & 0x7f) as u8;
        let horizontal_display_every = ((display >> 8) & 0x7f) as u8;
        if vertical_display_every == 0 || horizontal_display_every == 0 {
            return Err(DopExtensionError::new(
                "drawing-grid display multiples must be positive",
            ));
        }
        Ok(Self {
            horizontal_origin: le_u16(data, 0),
            vertical_origin: le_u16(data, 2),
            horizontal_spacing: le_u16(data, 4),
            vertical_spacing: le_u16(data, 6),
            vertical_display_every,
            horizontal_display_every,
            follow_margins: display & 0x8000 != 0,
            unused_bit: display & 0x80 != 0,
        })
    }

    pub(crate) fn encode(self) -> [u8; GRID_SIZE] {
        let mut data = [0u8; GRID_SIZE];
        put_u16(&mut data, 0, self.horizontal_origin);
        put_u16(&mut data, 2, self.vertical_origin);
        put_u16(&mut data, 4, self.horizontal_spacing);
        put_u16(&mut data, 6, self.vertical_spacing);
        let display = u16::from(self.vertical_display_every)
            | (u16::from(self.unused_bit) << 7)
            | (u16::from(self.horizontal_display_every) << 8)
            | (u16::from(self.follow_margins) << 15);
        put_u16(&mut data, 8, display);
        data
    }
}

impl Default for DrawingGrid {
    fn default() -> Self {
        Self {
            horizontal_origin: 1701,
            vertical_origin: 1984,
            horizontal_spacing: 180,
            vertical_spacing: 180,
            vertical_display_every: 1,
            horizontal_display_every: 1,
            follow_margins: true,
            unused_bit: false,
        }
    }
}

/// `AutoSummary` display mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoSummaryView {
    Highlight,
    HideNonSummary,
    InsertAtTop,
    NewDocument,
}

/// Passive `AutoSummary` state. No summarization is performed by this codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AutoSummaryState {
    pub valid: bool,
    pub view_active: bool,
    pub view: AutoSummaryView,
    pub update_properties: bool,
    pub desired_level: u16,
    pub highest_level: i32,
    pub current_level: i32,
}

impl AutoSummaryState {
    fn parse(data: &[u8]) -> Result<Self, DopExtensionError> {
        let flags = le_u16(data, 0);
        if flags & 0xffe0 != 0 {
            return Err(DopExtensionError::new("Asumyi reserved bits are nonzero"));
        }
        let view = match (flags >> 2) & 3 {
            0 => AutoSummaryView::Highlight,
            1 => AutoSummaryView::HideNonSummary,
            2 => AutoSummaryView::InsertAtTop,
            _ => AutoSummaryView::NewDocument,
        };
        Ok(Self {
            valid: flags & 1 != 0,
            view_active: flags & 2 != 0,
            view,
            update_properties: flags & 0x10 != 0,
            desired_level: le_u16(data, 2),
            highest_level: le_i32(data, 4),
            current_level: le_i32(data, 8),
        })
    }

    fn encode(self) -> [u8; 12] {
        let mut data = [0u8; 12];
        let mut flags = u16::from(self.valid) | (u16::from(self.view_active) << 1);
        flags |= match self.view {
            AutoSummaryView::Highlight => 0,
            AutoSummaryView::HideNonSummary => 1 << 2,
            AutoSummaryView::InsertAtTop => 2 << 2,
            AutoSummaryView::NewDocument => 3 << 2,
        };
        flags |= u16::from(self.update_properties) << 4;
        put_u16(&mut data, 0, flags);
        put_u16(&mut data, 2, self.desired_level);
        put_i32(&mut data, 4, self.highest_level);
        put_i32(&mut data, 8, self.current_level);
        data
    }
}

/// Validated document-event mask. It describes metadata only and never executes code.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentEventFlags(u32);

impl DocumentEventFlags {
    pub fn from_raw(raw: u32) -> Result<Self, DopExtensionError> {
        if raw & !ALLOWED_EVENT_FLAGS != 0 {
            Err(DopExtensionError::new(
                "Dop97 document-event mask has reserved bits",
            ))
        } else {
            Ok(Self(raw))
        }
    }

    #[must_use]
    pub const fn raw(self) -> u32 {
        self.0
    }

    #[must_use]
    pub const fn contains(self, flag: u32) -> bool {
        self.0 & flag != 0
    }
}

/// Passive virus-prompt/session metadata. The session key is never used to run macros.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MacroSecurityMetadata {
    pub prompted_on_load: bool,
    pub load_was_safe: bool,
    pub session_key: u32,
}

impl MacroSecurityMetadata {
    fn parse(raw: u32) -> Self {
        Self {
            prompted_on_load: raw & 1 != 0,
            load_was_safe: raw & 2 != 0,
            session_key: raw >> 2,
        }
    }

    fn raw(self) -> Result<u32, DopExtensionError> {
        if self.session_key >= (1 << 30) {
            return Err(DopExtensionError::new("macro session key exceeds 30 bits"));
        }
        Ok((self.session_key << 2)
            | u32::from(self.prompted_on_load)
            | (u32::from(self.load_was_safe) << 1))
    }
}

/// Typed, lossless Word 97 DOP extension.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dop97 {
    raw: [u8; DOP97_EXTENSION_SIZE],
    pub compatibility: CompatibilityOptions80,
    pub classification: DocumentClassification,
    pub typography: DocumentTypography,
    pub grid: DrawingGrid,
    pub outline_level: OutlineDisplayLevel,
    pub grammar_all_done: bool,
    pub grammar_all_clean: bool,
    pub subset_fonts: bool,
    pub disk_lvc_invalid: bool,
    pub snap_border: bool,
    pub include_header: bool,
    pub include_footer: bool,
    pub auto_summary: AutoSummaryState,
    pub character_count_with_spaces: i32,
    pub character_count_with_spaces_and_subdocs: i32,
    pub event_flags: DocumentEventFlags,
    pub macro_security: MacroSecurityMetadata,
    pub list_cache_main_doc_cp: u32,
    pub last_bullet_lfo: u16,
    pub last_number_lfo: u16,
    pub double_byte_character_count: i32,
    pub double_byte_character_count_with_subdocs: i32,
    pub footnote_number_format: NumberFormat,
    pub endnote_number_format: NumberFormat,
    pub page_zoom_font_half_points: u16,
    pub page_display_width: u16,
}

impl Dop97 {
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP97_SIZE {
            return Err(DopExtensionError::new("Dop97 is shorter than 500 bytes"));
        }
        Dop95::parse(dop)?;
        let extension = &dop[DOP_BASE_SIZE..DOP97_SIZE];
        let mut raw = [0u8; DOP97_EXTENSION_SIZE];
        raw.copy_from_slice(extension);
        let docinfo5 = le_u16(extension, DOCINFO5);
        let outline_level = OutlineDisplayLevel::parse(((docinfo5 >> 1) & 0xf) as u8)?;
        let footnote_number_format = parse_number_format(extension, FOOTNOTE_FORMAT)?;
        let endnote_number_format = parse_number_format(extension, ENDNOTE_FORMAT)?;
        Ok(Self {
            raw,
            compatibility: CompatibilityOptions80::from_raw(le_u32(extension, 0)),
            classification: DocumentClassification::parse(le_u16(extension, ADT))?,
            typography: DocumentTypography::parse(
                &extension[TYPOGRAPHY..TYPOGRAPHY + TYPOGRAPHY_SIZE],
            )?,
            grid: DrawingGrid::parse(&extension[GRID..GRID + GRID_SIZE])?,
            outline_level,
            grammar_all_done: docinfo5 & 0x20 != 0,
            grammar_all_clean: docinfo5 & 0x40 != 0,
            subset_fonts: docinfo5 & 0x80 != 0,
            disk_lvc_invalid: docinfo5 & 0x400 != 0,
            snap_border: docinfo5 & 0x800 != 0,
            include_header: docinfo5 & 0x1000 != 0,
            include_footer: docinfo5 & 0x2000 != 0,
            auto_summary: AutoSummaryState::parse(&extension[AUTOSUMMARY..AUTOSUMMARY + 12])?,
            character_count_with_spaces: le_i32(extension, CCH_WS),
            character_count_with_spaces_and_subdocs: le_i32(extension, CCH_WS_WITH_SUBDOCS),
            event_flags: DocumentEventFlags::from_raw(le_u32(extension, EVENTS))?,
            macro_security: MacroSecurityMetadata::parse(le_u32(extension, VIRUS_INFO)),
            list_cache_main_doc_cp: le_u32(extension, LIST_CACHE_CP),
            last_bullet_lfo: le_u16(extension, LAST_BULLET_LFO),
            last_number_lfo: le_u16(extension, LAST_NUMBER_LFO),
            double_byte_character_count: le_i32(extension, CDBC),
            double_byte_character_count_with_subdocs: le_i32(extension, CDBC_WITH_SUBDOCS),
            footnote_number_format,
            endnote_number_format,
            page_zoom_font_half_points: le_u16(extension, PAGE_ZOOM_FONT),
            page_display_width: le_u16(extension, PAGE_DISPLAY_WIDTH),
        })
    }

    /// Checks LFO indices once the document's list-table size is known.
    pub fn validate_list_indices(&self, lfo_count: usize) -> Result<(), DopExtensionError> {
        for (name, value) in [
            ("last bullet", self.last_bullet_lfo),
            ("last number", self.last_number_lfo),
        ] {
            if value != 0xffff && usize::from(value) >= lfo_count {
                return Err(DopExtensionError::new(format!(
                    "Dop97 {name} LFO index {value} exceeds list table"
                )));
            }
        }
        Ok(())
    }

    /// Writes the typed fields while preserving undefined bytes and later-version data.
    pub fn write_into(mut self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP97_SIZE {
            return Err(DopExtensionError::new(
                "Dop97 target is shorter than 500 bytes",
            ));
        }
        if self.compatibility.copts60() != le_u16(dop, COPTS60_OFFSET) {
            return Err(DopExtensionError::new(
                "Copts80.copts60 does not mirror DopBase.copts60",
            ));
        }
        if self.grid.vertical_display_every == 0
            || self.grid.vertical_display_every > 127
            || self.grid.horizontal_display_every == 0
            || self.grid.horizontal_display_every > 127
        {
            return Err(DopExtensionError::new(
                "drawing-grid display multiple is out of range",
            ));
        }
        put_u32(&mut self.raw, 0, self.compatibility.raw());
        put_u16(&mut self.raw, ADT, self.classification.raw());
        self.raw[TYPOGRAPHY..TYPOGRAPHY + TYPOGRAPHY_SIZE]
            .copy_from_slice(&self.typography.encode());
        self.raw[GRID..GRID + GRID_SIZE].copy_from_slice(&self.grid.encode());
        let old_docinfo5 = le_u16(&self.raw, DOCINFO5);
        let mut docinfo5 = old_docinfo5 & 0xc301;
        docinfo5 |= u16::from(self.outline_level.raw()) << 1;
        docinfo5 |= u16::from(self.grammar_all_done) << 5;
        docinfo5 |= u16::from(self.grammar_all_clean) << 6;
        docinfo5 |= u16::from(self.subset_fonts) << 7;
        docinfo5 |= u16::from(self.disk_lvc_invalid) << 10;
        docinfo5 |= u16::from(self.snap_border) << 11;
        docinfo5 |= u16::from(self.include_header) << 12;
        docinfo5 |= u16::from(self.include_footer) << 13;
        put_u16(&mut self.raw, DOCINFO5, docinfo5);
        self.raw[AUTOSUMMARY..AUTOSUMMARY + 12].copy_from_slice(&self.auto_summary.encode());
        put_i32(&mut self.raw, CCH_WS, self.character_count_with_spaces);
        put_i32(
            &mut self.raw,
            CCH_WS_WITH_SUBDOCS,
            self.character_count_with_spaces_and_subdocs,
        );
        put_u32(&mut self.raw, EVENTS, self.event_flags.raw());
        put_u32(&mut self.raw, VIRUS_INFO, self.macro_security.raw()?);
        put_u32(&mut self.raw, LIST_CACHE_CP, self.list_cache_main_doc_cp);
        put_u16(&mut self.raw, LAST_BULLET_LFO, self.last_bullet_lfo);
        put_u16(&mut self.raw, LAST_NUMBER_LFO, self.last_number_lfo);
        put_i32(&mut self.raw, CDBC, self.double_byte_character_count);
        put_i32(
            &mut self.raw,
            CDBC_WITH_SUBDOCS,
            self.double_byte_character_count_with_subdocs,
        );
        put_number_format(&mut self.raw, FOOTNOTE_FORMAT, self.footnote_number_format);
        put_number_format(&mut self.raw, ENDNOTE_FORMAT, self.endnote_number_format);
        put_u16(
            &mut self.raw,
            PAGE_ZOOM_FONT,
            self.page_zoom_font_half_points,
        );
        put_u16(&mut self.raw, PAGE_DISPLAY_WIDTH, self.page_display_width);
        dop[DOP_BASE_SIZE..DOP97_SIZE].copy_from_slice(&self.raw);
        Ok(())
    }
}

fn parse_number_format(data: &[u8], offset: usize) -> Result<NumberFormat, DopExtensionError> {
    let raw = le_u16(data, offset);
    if raw > u16::from(u8::MAX) {
        return Err(DopExtensionError::new(format!(
            "invalid 16-bit MSONFC {raw:#06x}"
        )));
    }
    NumberFormat::try_from(raw as u8)
        .map_err(|_| DopExtensionError::new(format!("unknown MSONFC {raw:#04x}")))
}

fn put_number_format(data: &mut [u8], offset: usize, value: NumberFormat) {
    put_u16(data, offset, u16::from(value as u8));
}

fn decode_utf16(data: &[u8], count: usize) -> Result<String, DopExtensionError> {
    let units = data
        .chunks_exact(2)
        .take(count)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units)
        .map_err(|_| DopExtensionError::new("DopTypography contains invalid UTF-16"))
}

fn write_utf16(
    target: &mut [u8],
    stored: &mut String,
    value: &str,
    maximum: usize,
) -> Result<(), DopExtensionError> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > maximum {
        return Err(DopExtensionError::new(
            "kinsoku punctuation string is too long",
        ));
    }
    target.fill(0);
    for (slot, unit) in target.chunks_exact_mut(2).zip(units) {
        slot.copy_from_slice(&unit.to_le_bytes());
    }
    stored.clear();
    stored.push_str(value);
    Ok(())
}

fn le_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn le_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(
        data[offset..offset + 4]
            .try_into()
            .expect("fixed-width slice"),
    )
}

fn le_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes(
        data[offset..offset + 4]
            .try_into()
            .expect("fixed-width slice"),
    )
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

    fn valid_dop97() -> Vec<u8> {
        let mut dop = vec![0u8; DOP97_SIZE];
        let grid = DrawingGrid::default().encode();
        dop[DOP_BASE_SIZE + GRID..DOP_BASE_SIZE + GRID + GRID_SIZE].copy_from_slice(&grid);
        dop
    }

    #[test]
    fn parses_and_writes_losslessly() {
        let mut dop = valid_dop97();
        dop[86] = 0x34;
        dop[87] = 0x12;
        dop[DOP_BASE_SIZE + 328] = 0xaa;
        dop[DOP_BASE_SIZE + 329] = 0x55;
        dop[DOP_BASE_SIZE + 358] = 0x7e;
        dop[DOP_BASE_SIZE + 404] = 0x6d;
        let parsed = Dop97::parse(&dop).unwrap();
        let mut output = dop.clone();
        parsed.write_into(&mut output).unwrap();
        assert_eq!(output, dop);
    }

    #[test]
    fn exposes_typed_fields_and_inert_security_metadata() {
        let mut dop = valid_dop97();
        put_u16(&mut dop, DOP_BASE_SIZE + ADT, 2);
        put_u16(&mut dop, DOP_BASE_SIZE + DOCINFO5, (4 << 1) | 0x20 | 0x1000);
        put_u32(&mut dop, DOP_BASE_SIZE + EVENTS, 0x1001);
        put_u32(&mut dop, DOP_BASE_SIZE + VIRUS_INFO, (0x1234 << 2) | 3);
        let parsed = Dop97::parse(&dop).unwrap();
        assert_eq!(parsed.classification, DocumentClassification::Email);
        assert_eq!(parsed.outline_level, OutlineDisplayLevel::ThroughHeading(4));
        assert!(parsed.grammar_all_done);
        assert!(parsed.include_header);
        assert_eq!(parsed.event_flags.raw(), 0x1001);
        assert_eq!(parsed.macro_security.session_key, 0x1234);
        assert!(parsed.macro_security.prompted_on_load);
        assert!(parsed.macro_security.load_was_safe);
    }

    #[test]
    fn rejects_must_constraint_violations() {
        let cases: &[(usize, &[u8])] = &[
            (DOP_BASE_SIZE, &[1, 0]),
            (DOP_BASE_SIZE + ADT, &[3, 0]),
            (DOP_BASE_SIZE + TYPOGRAPHY, &[0, 0xf8]),
            (DOP_BASE_SIZE + DOCINFO5, &[20, 0]),
            (DOP_BASE_SIZE + AUTOSUMMARY, &[0x20, 0]),
            (DOP_BASE_SIZE + EVENTS, &[0x40, 0, 0, 0]),
            (DOP_BASE_SIZE + FOOTNOTE_FORMAT, &[0x3c, 0]),
        ];
        for &(offset, bytes) in cases {
            let mut dop = valid_dop97();
            dop[offset..offset + bytes.len()].copy_from_slice(bytes);
            assert!(Dop97::parse(&dop).is_err(), "offset {offset}");
        }
    }

    #[test]
    fn validates_context_dependent_list_indices() {
        let mut dop = valid_dop97();
        put_u16(&mut dop, DOP_BASE_SIZE + LAST_BULLET_LFO, 2);
        put_u16(&mut dop, DOP_BASE_SIZE + LAST_NUMBER_LFO, 0xffff);
        let parsed = Dop97::parse(&dop).unwrap();
        assert!(parsed.validate_list_indices(2).is_err());
        assert!(parsed.validate_list_indices(3).is_ok());
    }

    #[test]
    fn typography_strings_round_trip_as_utf16() {
        let mut dop = valid_dop97();
        let mut parsed = Dop97::parse(&dop).unwrap();
        parsed.typography.set_cannot_start_line("。😀").unwrap();
        parsed.typography.set_cannot_end_line("（").unwrap();
        parsed.write_into(&mut dop).unwrap();
        let reparsed = Dop97::parse(&dop).unwrap();
        assert_eq!(reparsed.typography.cannot_start_line, "。😀");
        assert_eq!(reparsed.typography.cannot_end_line, "（");
    }
}
