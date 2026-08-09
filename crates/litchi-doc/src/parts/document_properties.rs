//! Typed legacy Word Document Properties (`Dop`) metadata.

use super::super::package::{Error as PackageError, Result};
use super::document_properties_97::{Dop95, Dop97, DopExtensionError, DrawingGrid};
use super::document_properties_2000::Dop2000;
use super::document_properties_2002::Dop2002;
use super::document_properties_2003::Dop2003;
use super::document_properties_2007::Dop2007;
use super::document_properties_2010::Dop2010;
use super::document_properties_2013::Dop2013;
use super::fib::FileInformationBlock;

const FIB_INDEX: usize = 31;
const BASE_SIZE: usize = 84;
const WORD97_SIZE: usize = 500;
const WORD2002_SIZE: usize = 594;
const WORD97_GRID_OFFSET: usize = 0x190;
const WORD97_GRID_SIZE: usize = 10;
const DOCINFO5_OFFSET: usize = 0x19A;
const FACTOID_FLAGS_OFFSET: usize = 0x224;
const INCLUDE_HEADER_MASK: u16 = 0x1000;
const INCLUDE_FOOTER_MASK: u16 = 0x2000;
const EMBED_FACTOIDS_MASK: u16 = 0x0008;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn u16_at(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn i16_at(data: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes([data[offset], data[offset + 1]])
}

fn u32_at(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn i32_at(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn set_flag(byte: &mut u8, mask: u8, value: bool) {
    if value {
        *byte |= mask;
    } else {
        *byte &= !mask;
    }
}

/// The on-disk DOP generation identified by its specification-defined size.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentPropertyVersion {
    Base,
    Word95,
    Word97,
    Word2000,
    Word2002,
    Word2003,
    Word2007,
    Word2010,
    Word2013,
}

impl DocumentPropertyVersion {
    pub const fn byte_len(self) -> usize {
        match self {
            Self::Base => 84,
            Self::Word95 => 88,
            Self::Word97 => 500,
            Self::Word2000 => 544,
            Self::Word2002 => 594,
            Self::Word2003 => 616,
            Self::Word2007 => 674,
            Self::Word2010 => 690,
            Self::Word2013 => 694,
        }
    }

    fn from_len(length: usize) -> Result<Self> {
        match length {
            84 => Ok(Self::Base),
            88 => Ok(Self::Word95),
            500 => Ok(Self::Word97),
            544 => Ok(Self::Word2000),
            594 => Ok(Self::Word2002),
            616 => Ok(Self::Word2003),
            674 => Ok(Self::Word2007),
            690 => Ok(Self::Word2010),
            694 => Ok(Self::Word2013),
            _ => Err(corrupted(format!(
                "Dop has non-standard length {length}; expected 84, 88, 500, 544, 594, 616, 674, 690, or 694 bytes"
            ))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FootnotePlacement {
    SectionEnd,
    PageBottom,
    BeneathText,
}

impl FootnotePlacement {
    fn from_bits(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::SectionEnd),
            1 => Ok(Self::PageBottom),
            2 => Ok(Self::BeneathText),
            _ => Err(corrupted("DopBase.fpc has reserved value 3")),
        }
    }

    const fn bits(self) -> u8 {
        match self {
            Self::SectionEnd => 0,
            Self::PageBottom => 1,
            Self::BeneathText => 2,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteNumberingRestart {
    Continuous,
    EachSection,
    EachPage,
}

impl NoteNumberingRestart {
    fn from_bits(value: u16, field: &str) -> Result<Self> {
        match value {
            0 => Ok(Self::Continuous),
            1 => Ok(Self::EachSection),
            2 => Ok(Self::EachPage),
            _ => Err(corrupted(format!("{field} has reserved value 3"))),
        }
    }

    const fn bits(self) -> u16 {
        match self {
            Self::Continuous => 0,
            Self::EachSection => 1,
            Self::EachPage => 2,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EndnotePlacement {
    SectionEnd,
    DocumentEnd,
}

impl EndnotePlacement {
    fn from_bits(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::SectionEnd),
            3 => Ok(Self::DocumentEnd),
            _ => Err(corrupted(format!("DopBase.epc has reserved value {value}"))),
        }
    }

    const fn bits(self) -> u16 {
        match self {
            Self::SectionEnd => 0,
            Self::DocumentEnd => 3,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SavedViewKind {
    Default,
    Print,
    Outline,
    MasterPages,
    Normal,
    Web,
}

impl SavedViewKind {
    fn from_bits(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::Default),
            1 => Ok(Self::Print),
            2 => Ok(Self::Outline),
            3 => Ok(Self::MasterPages),
            4 => Ok(Self::Normal),
            5 => Ok(Self::Web),
            _ => Err(corrupted(format!(
                "DopBase.wvkoSaved has reserved value {value}"
            ))),
        }
    }

    const fn bits(self) -> u16 {
        self as u16
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SavedZoomKind {
    Default,
    FullPage,
    BestFit,
    TextFit,
}

impl SavedZoomKind {
    fn from_bits(value: u16) -> Self {
        match value {
            0 => Self::Default,
            1 => Self::FullPage,
            2 => Self::BestFit,
            _ => Self::TextFit,
        }
    }

    const fn bits(self) -> u16 {
        self as u16
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SavedView {
    pub kind: SavedViewKind,
    /// `None` represents the application's default zoom percentage.
    pub zoom_percent: Option<u16>,
    pub zoom_kind: SavedZoomKind,
    pub gutter_at_top: bool,
}

/// A packed Word `DTTM`; zero denotes an omitted timestamp.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DocumentTimestamp(u32);

impl DocumentTimestamp {
    pub fn from_raw(raw: u32) -> Result<Self> {
        if raw != 0 {
            let minute = raw & 0x3F;
            let hour = (raw >> 6) & 0x1F;
            let day = (raw >> 11) & 0x1F;
            let month = (raw >> 16) & 0x0F;
            let weekday = (raw >> 29) & 0x07;
            if minute > 59
                || hour > 23
                || !(1..=31).contains(&day)
                || !(1..=12).contains(&month)
                || weekday > 6
            {
                return Err(corrupted(format!("invalid DopBase DTTM {raw:#010x}")));
            }
        }
        Ok(Self(raw))
    }

    pub const fn raw(self) -> u32 {
        self.0
    }
    pub const fn is_omitted(self) -> bool {
        self.0 == 0
    }
    pub fn minute(self) -> Option<u8> {
        (!self.is_omitted()).then_some((self.0 & 0x3F) as u8)
    }
    pub fn hour(self) -> Option<u8> {
        (!self.is_omitted()).then_some(((self.0 >> 6) & 0x1F) as u8)
    }
    pub fn day(self) -> Option<u8> {
        (!self.is_omitted()).then_some(((self.0 >> 11) & 0x1F) as u8)
    }
    pub fn month(self) -> Option<u8> {
        (!self.is_omitted()).then_some(((self.0 >> 16) & 0x0F) as u8)
    }
    pub fn year(self) -> Option<u16> {
        (!self.is_omitted()).then_some((((self.0 >> 20) & 0x1FF) as u16) + 1900)
    }
    pub fn weekday(self) -> Option<u8> {
        (!self.is_omitted()).then_some(((self.0 >> 29) & 0x07) as u8)
    }
}

/// Word 6 compatibility flags retained in every later DOP generation.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct CompatibilityOptions60(u16);

impl CompatibilityOptions60 {
    pub const fn from_raw(raw: u16) -> Self {
        Self(raw)
    }
    pub const fn raw(self) -> u16 {
        self.0
    }
    pub const fn no_tab_for_indent(self) -> bool {
        self.0 & 0x0001 != 0
    }
    pub const fn no_space_raise_lower(self) -> bool {
        self.0 & 0x0002 != 0
    }
    pub const fn suppress_spacing_after_page_break(self) -> bool {
        self.0 & 0x0004 != 0
    }
    pub const fn wrap_trailing_spaces(self) -> bool {
        self.0 & 0x0008 != 0
    }
    pub const fn map_print_text_color(self) -> bool {
        self.0 & 0x0010 != 0
    }
    pub const fn no_column_balance(self) -> bool {
        self.0 & 0x0020 != 0
    }
    pub const fn convert_mail_merge_escapes(self) -> bool {
        self.0 & 0x0040 != 0
    }
    pub const fn suppress_top_spacing(self) -> bool {
        self.0 & 0x0080 != 0
    }
    pub const fn original_word_table_rules(self) -> bool {
        self.0 & 0x0100 != 0
    }
    pub const fn transparent_metafiles(self) -> bool {
        self.0 & 0x0200 != 0
    }
    pub const fn show_breaks_in_frames(self) -> bool {
        self.0 & 0x0400 != 0
    }
    pub const fn swap_borders_on_facing_pages(self) -> bool {
        self.0 & 0x0800 != 0
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ProtectionSettings {
    pub comments_or_read_only: bool,
    pub form_fields: bool,
    pub tracked_revisions: bool,
    pub track_revisions: bool,
    pub vba_project_locked: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentStatistics {
    pub words: i32,
    pub characters: i32,
    pub pages: i16,
    pub paragraphs: i32,
    pub lines: i32,
    pub words_with_subdocuments: i32,
    pub characters_with_subdocuments: i32,
    pub pages_with_subdocuments: i16,
    pub paragraphs_with_subdocuments: i32,
    pub lines_with_subdocuments: i32,
}

/// The 84-byte portion common to every DOP generation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentPropertiesBase {
    raw: [u8; BASE_SIZE],
}

impl Default for DocumentPropertiesBase {
    fn default() -> Self {
        Self {
            raw: [0; BASE_SIZE],
        }
    }
}

impl DocumentPropertiesBase {
    pub fn parse(data: &[u8]) -> Result<Self> {
        let bytes = data
            .get(..BASE_SIZE)
            .ok_or_else(|| corrupted("DopBase is shorter than 84 bytes"))?;
        let mut raw = [0u8; BASE_SIZE];
        raw.copy_from_slice(bytes);
        let value = Self { raw };
        value.validate()?;
        Ok(value)
    }

    pub fn as_bytes(&self) -> &[u8; BASE_SIZE] {
        &self.raw
    }

    pub fn facing_pages(&self) -> bool {
        self.raw[0] & 0x01 != 0
    }
    pub fn mail_merge_main_document(&self) -> bool {
        self.raw[0] & 0x04 != 0
    }
    pub fn footnote_placement(&self) -> FootnotePlacement {
        FootnotePlacement::from_bits((self.raw[0] >> 5) & 0x03).expect("validated DopBase fpc")
    }
    pub fn footnote_numbering_restart(&self) -> NoteNumberingRestart {
        NoteNumberingRestart::from_bits(u16_at(&self.raw, 2) & 0x03, "DopBase.rncFtn")
            .expect("validated DopBase rncFtn")
    }
    pub fn footnote_start(&self) -> u16 {
        u16_at(&self.raw, 2) >> 2
    }
    pub fn spelling_complete(&self) -> bool {
        self.raw[4] & 0x40 != 0
    }
    pub fn spelling_clean(&self) -> bool {
        self.raw[4] & 0x80 != 0
    }
    pub fn hide_spelling_errors(&self) -> bool {
        self.raw[5] & 0x01 != 0
    }
    pub fn hide_grammar_errors(&self) -> bool {
        self.raw[5] & 0x02 != 0
    }
    pub fn labels_document(&self) -> bool {
        self.raw[5] & 0x04 != 0
    }
    pub fn hyphenate_capitals(&self) -> bool {
        self.raw[5] & 0x08 != 0
    }
    pub fn auto_hyphenation(&self) -> bool {
        self.raw[5] & 0x10 != 0
    }
    pub fn update_styles_from_template(&self) -> bool {
        self.raw[5] & 0x40 != 0
    }
    pub fn exact_statistics(&self) -> bool {
        self.raw[6] & 0x02 != 0
    }
    pub fn mirror_margins(&self) -> bool {
        self.raw[6] & 0x20 != 0
    }
    pub fn word97_compatibility_mode(&self) -> bool {
        self.raw[6] & 0x40 != 0
    }
    pub fn show_revision_markup(&self) -> bool {
        self.raw[7] & 0x08 != 0
    }
    pub fn print_revision_markup(&self) -> bool {
        self.raw[7] & 0x10 != 0
    }
    pub fn embed_true_type_fonts(&self) -> bool {
        self.raw[7] & 0x80 != 0
    }
    pub fn protection(&self) -> ProtectionSettings {
        ProtectionSettings {
            comments_or_read_only: self.raw[6] & 0x10 != 0,
            form_fields: self.raw[7] & 0x02 != 0,
            tracked_revisions: self.raw[7] & 0x40 != 0,
            track_revisions: self.raw[5] & 0x80 != 0,
            vba_project_locked: self.raw[7] & 0x20 != 0,
        }
    }
    pub fn compatibility_options(&self) -> CompatibilityOptions60 {
        CompatibilityOptions60::from_raw(u16_at(&self.raw, 8))
    }
    pub fn default_tab_stop_twips(&self) -> u16 {
        u16_at(&self.raw, 10)
    }
    pub fn web_code_page(&self) -> u16 {
        u16_at(&self.raw, 12)
    }
    pub fn hyphenation_zone_twips(&self) -> u16 {
        u16_at(&self.raw, 14)
    }
    pub fn consecutive_hyphen_limit(&self) -> u16 {
        u16_at(&self.raw, 16)
    }
    pub fn created_at(&self) -> DocumentTimestamp {
        DocumentTimestamp(u32_at(&self.raw, 20))
    }
    pub fn revised_at(&self) -> DocumentTimestamp {
        DocumentTimestamp(u32_at(&self.raw, 24))
    }
    pub fn last_printed_at(&self) -> DocumentTimestamp {
        DocumentTimestamp(u32_at(&self.raw, 28))
    }
    pub fn revision_count(&self) -> u16 {
        u16_at(&self.raw, 32)
    }
    pub fn editing_minutes(&self) -> i32 {
        i32_at(&self.raw, 34)
    }
    pub fn statistics(&self) -> DocumentStatistics {
        DocumentStatistics {
            words: i32_at(&self.raw, 38),
            characters: i32_at(&self.raw, 42),
            pages: i16_at(&self.raw, 46),
            paragraphs: i32_at(&self.raw, 48),
            lines: i32_at(&self.raw, 56),
            words_with_subdocuments: i32_at(&self.raw, 60),
            characters_with_subdocuments: i32_at(&self.raw, 64),
            pages_with_subdocuments: i16_at(&self.raw, 68),
            paragraphs_with_subdocuments: i32_at(&self.raw, 70),
            lines_with_subdocuments: i32_at(&self.raw, 74),
        }
    }
    pub fn endnote_numbering_restart(&self) -> NoteNumberingRestart {
        NoteNumberingRestart::from_bits(u16_at(&self.raw, 52) & 0x03, "DopBase.rncEdn")
            .expect("validated DopBase rncEdn")
    }
    pub fn endnote_start(&self) -> u16 {
        u16_at(&self.raw, 52) >> 2
    }
    pub fn endnote_placement(&self) -> EndnotePlacement {
        EndnotePlacement::from_bits(u16_at(&self.raw, 54) & 0x03).expect("validated DopBase epc")
    }
    pub fn print_form_data_only(&self) -> bool {
        u16_at(&self.raw, 54) & 0x0400 != 0
    }
    pub fn save_form_data_only(&self) -> bool {
        u16_at(&self.raw, 54) & 0x0800 != 0
    }
    pub fn shade_form_fields(&self) -> bool {
        u16_at(&self.raw, 54) & 0x1000 != 0
    }
    pub fn shade_merge_fields(&self) -> bool {
        u16_at(&self.raw, 54) & 0x2000 != 0
    }
    pub fn include_subdocuments_in_statistics(&self) -> bool {
        u16_at(&self.raw, 54) & 0x8000 != 0
    }
    pub fn protection_key(&self) -> i32 {
        i32_at(&self.raw, 78)
    }
    pub fn saved_view(&self) -> SavedView {
        let raw = u16_at(&self.raw, 82);
        let percent = (raw >> 3) & 0x01FF;
        SavedView {
            kind: SavedViewKind::from_bits(raw & 0x07).expect("validated DopBase view"),
            zoom_percent: (percent != 0).then_some(percent),
            zoom_kind: SavedZoomKind::from_bits((raw >> 12) & 0x03),
            gutter_at_top: raw & 0x8000 != 0,
        }
    }

    pub fn set_facing_pages(&mut self, value: bool) {
        set_flag(&mut self.raw[0], 0x01, value);
    }
    pub fn set_footnote_placement(&mut self, value: FootnotePlacement) {
        self.raw[0] = (self.raw[0] & !0x60) | (value.bits() << 5);
    }
    pub fn set_footnote_numbering(
        &mut self,
        restart: NoteNumberingRestart,
        start: u16,
    ) -> Result<()> {
        if start > 0x3FFF {
            return Err(corrupted("footnote start exceeds 14 bits"));
        }
        self.raw[2..4].copy_from_slice(&((start << 2) | restart.bits()).to_le_bytes());
        Ok(())
    }
    pub fn set_endnote_numbering(
        &mut self,
        restart: NoteNumberingRestart,
        start: u16,
    ) -> Result<()> {
        if start > 0x3FFF {
            return Err(corrupted("endnote start exceeds 14 bits"));
        }
        self.raw[52..54].copy_from_slice(&((start << 2) | restart.bits()).to_le_bytes());
        Ok(())
    }
    pub fn set_endnote_placement(&mut self, value: EndnotePlacement) {
        let current = u16_at(&self.raw, 54);
        self.raw[54..56].copy_from_slice(&((current & !0x0003) | value.bits()).to_le_bytes());
    }
    pub fn set_saved_view(&mut self, value: SavedView) -> Result<()> {
        let percent = value.zoom_percent.unwrap_or(0);
        if percent != 0 && !(10..=500).contains(&percent) {
            return Err(corrupted(
                "saved zoom percentage must be absent or 10..=500",
            ));
        }
        let raw = value.kind.bits()
            | (percent << 3)
            | (value.zoom_kind.bits() << 12)
            | if value.gutter_at_top { 0x8000 } else { 0 };
        self.raw[82..84].copy_from_slice(&raw.to_le_bytes());
        Ok(())
    }
    pub fn set_protection(&mut self, value: ProtectionSettings) -> Result<()> {
        validate_protection(value, self.raw[5] & 0x20 != 0)?;
        set_flag(&mut self.raw[6], 0x10, value.comments_or_read_only);
        set_flag(&mut self.raw[7], 0x02, value.form_fields);
        set_flag(&mut self.raw[7], 0x40, value.tracked_revisions);
        set_flag(&mut self.raw[5], 0x80, value.track_revisions);
        set_flag(&mut self.raw[7], 0x20, value.vba_project_locked);
        Ok(())
    }

    fn validate(&self) -> Result<()> {
        FootnotePlacement::from_bits((self.raw[0] >> 5) & 0x03)?;
        NoteNumberingRestart::from_bits(u16_at(&self.raw, 2) & 0x03, "DopBase.rncFtn")?;
        NoteNumberingRestart::from_bits(u16_at(&self.raw, 52) & 0x03, "DopBase.rncEdn")?;
        EndnotePlacement::from_bits(u16_at(&self.raw, 54) & 0x03)?;
        if u16_at(&self.raw, 18) != 0 {
            return Err(corrupted("DopBase.wSpare2 must be zero"));
        }
        if u16_at(&self.raw, 32) > 0x7FFF {
            return Err(corrupted("DopBase.nRevision exceeds 0x7FFF"));
        }
        if u16_at(&self.raw, 54) & 0x4000 != 0 {
            return Err(corrupted("DopBase reserved form-data bit must be zero"));
        }
        for offset in [20, 24, 28] {
            DocumentTimestamp::from_raw(u32_at(&self.raw, offset))?;
        }
        validate_protection(self.protection(), self.raw[5] & 0x20 != 0)?;
        let view = u16_at(&self.raw, 82);
        SavedViewKind::from_bits(view & 0x07)?;
        let percent = (view >> 3) & 0x01FF;
        if percent != 0 && !(10..=500).contains(&percent) {
            return Err(corrupted("DopBase.pctWwdSaved must be zero or 10..=500"));
        }
        Ok(())
    }
}

fn validate_protection(value: ProtectionSettings, form_no_fields: bool) -> Result<()> {
    let restriction_count = u8::from(value.comments_or_read_only)
        + u8::from(value.form_fields)
        + u8::from(value.tracked_revisions);
    if restriction_count > 1 {
        return Err(corrupted(
            "DopBase contains mutually exclusive document protection modes",
        ));
    }
    if value.tracked_revisions && !value.track_revisions {
        return Err(corrupted("DopBase.fLockRev requires fRevMarking"));
    }
    if form_no_fields && !value.form_fields {
        return Err(corrupted("DopBase.fFormNoFields requires fProtEnabled"));
    }
    Ok(())
}

/// Complete versioned DOP with typed base fields and lossless later-version bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentProperties {
    version: DocumentPropertyVersion,
    base: DocumentPropertiesBase,
    extension: Vec<u8>,
}

/// The specification-defined typed payload for an exact DOP generation.
///
/// Parsing this view is explicit so callers can preserve and inspect real-world
/// documents whose later extension contains a producer defect without making
/// the rest of the document inaccessible.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Versioned {
    Base,
    Word95(Dop95),
    Word97(Box<Dop97>),
    Word2000(Dop2000),
    Word2002(Dop2002),
    Word2003(Dop2003),
    Word2007(Dop2007),
    Word2010(Dop2010),
    Word2013(Dop2013),
}

impl DocumentProperties {
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let start = usize::try_from(offset).map_err(|_| corrupted("Dop offset is too large"))?;
        let length = usize::try_from(length).map_err(|_| corrupted("Dop length is too large"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("Dop range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("Dop extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        let version = DocumentPropertyVersion::from_len(data.len())?;
        let base = DocumentPropertiesBase::parse(data)?;
        Ok(Self {
            version,
            base,
            extension: data[BASE_SIZE..].to_vec(),
        })
    }

    pub fn version(&self) -> DocumentPropertyVersion {
        self.version
    }
    pub fn base(&self) -> &DocumentPropertiesBase {
        &self.base
    }
    pub fn base_mut(&mut self) -> &mut DocumentPropertiesBase {
        &mut self.base
    }
    pub fn extension_bytes(&self) -> &[u8] {
        &self.extension
    }

    /// Parses the complete version-specific DOP extension into its typed form.
    ///
    /// # Errors
    ///
    /// Returns [`DopExtensionError`] when a fixed field, reserved bit, or
    /// version-specific value violates the selected DOP grammar.
    pub fn versioned(&self) -> std::result::Result<Versioned, DopExtensionError> {
        let data = self
            .to_bytes()
            .map_err(|error| DopExtensionError::new(error.to_string()))?;
        match self.version {
            DocumentPropertyVersion::Base => Ok(Versioned::Base),
            DocumentPropertyVersion::Word95 => Dop95::parse(&data).map(Versioned::Word95),
            DocumentPropertyVersion::Word97 => {
                Dop97::parse(&data).map(Box::new).map(Versioned::Word97)
            },
            DocumentPropertyVersion::Word2000 => Dop2000::parse(&data).map(Versioned::Word2000),
            DocumentPropertyVersion::Word2002 => Dop2002::parse(&data).map(Versioned::Word2002),
            DocumentPropertyVersion::Word2003 => Dop2003::parse(&data).map(Versioned::Word2003),
            DocumentPropertyVersion::Word2007 => Dop2007::parse(&data).map(Versioned::Word2007),
            DocumentPropertyVersion::Word2010 => Dop2010::parse(&data).map(Versioned::Word2010),
            DocumentPropertyVersion::Word2013 => Dop2013::parse(&data).map(Versioned::Word2013),
        }
    }
    pub fn includes_headers(&self) -> Option<bool> {
        self.absolute_u16(DOCINFO5_OFFSET)
            .map(|value| value & INCLUDE_HEADER_MASK != 0)
    }
    pub fn includes_footers(&self) -> Option<bool> {
        self.absolute_u16(DOCINFO5_OFFSET)
            .map(|value| value & INCLUDE_FOOTER_MASK != 0)
    }
    /// Whether Word should preserve embedded smart-tag/factoid metadata.
    pub fn embeds_factoids(&self) -> Option<bool> {
        self.absolute_u16(FACTOID_FLAGS_OFFSET)
            .map(|value| value & EMBED_FACTOIDS_MASK != 0)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.base.validate()?;
        if self.extension.len() + BASE_SIZE != self.version.byte_len() {
            return Err(corrupted("Dop version and extension length disagree"));
        }
        let mut data = Vec::with_capacity(self.version.byte_len());
        data.extend_from_slice(self.base.as_bytes());
        data.extend_from_slice(&self.extension);
        Ok(data)
    }

    fn absolute_u16(&self, offset: usize) -> Option<u16> {
        let extension_offset = offset.checked_sub(BASE_SIZE)?;
        self.extension
            .get(extension_offset..extension_offset + 2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    pub(crate) fn word97_writer_bytes(
        facing_pages: bool,
        include_headers: bool,
        include_footers: bool,
    ) -> Vec<u8> {
        let mut base = DocumentPropertiesBase::default();
        // Retain Word's historical widow-control default and bottom-page footnotes.
        base.raw[0] = 0x22;
        base.set_facing_pages(facing_pages);
        let mut data = vec![0u8; WORD97_SIZE];
        data[..BASE_SIZE].copy_from_slice(base.as_bytes());
        data[WORD97_GRID_OFFSET..WORD97_GRID_OFFSET + WORD97_GRID_SIZE]
            .copy_from_slice(&DrawingGrid::default().encode());
        let mut docinfo5 = 0u16;
        if include_headers {
            docinfo5 |= INCLUDE_HEADER_MASK;
        }
        if include_footers {
            docinfo5 |= INCLUDE_FOOTER_MASK;
        }
        data[DOCINFO5_OFFSET..DOCINFO5_OFFSET + 2].copy_from_slice(&docinfo5.to_le_bytes());
        data
    }

    pub(crate) fn writer_bytes(
        facing_pages: bool,
        include_headers: bool,
        include_footers: bool,
        embed_factoids: bool,
    ) -> Vec<u8> {
        let mut data = Self::word97_writer_bytes(facing_pages, include_headers, include_footers);
        if embed_factoids {
            data.resize(WORD2002_SIZE, 0);
            data[FACTOID_FLAGS_OFFSET..FACTOID_FLAGS_OFFSET + 2]
                .copy_from_slice(&EMBED_FACTOIDS_MASK.to_le_bytes());
        }
        data
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn packed_timestamp(year: u16, month: u8, day: u8, hour: u8, minute: u8, weekday: u8) -> u32 {
        u32::from(minute)
            | (u32::from(hour) << 6)
            | (u32::from(day) << 11)
            | (u32::from(month) << 16)
            | (u32::from(year - 1900) << 20)
            | (u32::from(weekday) << 29)
    }

    #[test]
    fn parses_typed_base_fields_and_round_trips_extensions_losslessly() {
        let mut data = DocumentProperties::word97_writer_bytes(true, true, false);
        data[2..4].copy_from_slice(&((7u16 << 2) | 1).to_le_bytes());
        data[5] |= 0x10;
        data[8..10].copy_from_slice(&0x0801u16.to_le_bytes());
        data[10..12].copy_from_slice(&720u16.to_le_bytes());
        data[20..24].copy_from_slice(&packed_timestamp(2026, 7, 18, 9, 30, 6).to_le_bytes());
        data[38..42].copy_from_slice(&123i32.to_le_bytes());
        data[52..54].copy_from_slice(&((4u16 << 2) | 2).to_le_bytes());
        data[54..56].copy_from_slice(&0x8003u16.to_le_bytes());
        data[82..84].copy_from_slice(&(1u16 | (125 << 3) | (2 << 12)).to_le_bytes());

        let dop = DocumentProperties::parse_bytes(&data).unwrap();
        assert_eq!(dop.version(), DocumentPropertyVersion::Word97);
        assert!(dop.base().facing_pages());
        assert_eq!(
            dop.base().footnote_placement(),
            FootnotePlacement::PageBottom
        );
        assert_eq!(
            dop.base().footnote_numbering_restart(),
            NoteNumberingRestart::EachSection
        );
        assert_eq!(dop.base().footnote_start(), 7);
        assert!(dop.base().auto_hyphenation());
        assert!(dop.base().compatibility_options().no_tab_for_indent());
        assert!(
            dop.base()
                .compatibility_options()
                .swap_borders_on_facing_pages()
        );
        assert_eq!(dop.base().created_at().year(), Some(2026));
        assert_eq!(dop.base().statistics().words, 123);
        assert_eq!(
            dop.base().endnote_placement(),
            EndnotePlacement::DocumentEnd
        );
        assert_eq!(dop.base().saved_view().zoom_percent, Some(125));
        assert_eq!(dop.includes_headers(), Some(true));
        assert_eq!(dop.includes_footers(), Some(false));
        assert_eq!(dop.to_bytes().unwrap(), data);
    }

    #[test]
    fn typed_mutation_validates_note_view_and_protection_domains() {
        let mut base = DocumentPropertiesBase::default();
        base.set_footnote_numbering(NoteNumberingRestart::EachPage, 12)
            .unwrap();
        base.set_endnote_numbering(NoteNumberingRestart::Continuous, 3)
            .unwrap();
        base.set_endnote_placement(EndnotePlacement::DocumentEnd);
        base.set_saved_view(SavedView {
            kind: SavedViewKind::Web,
            zoom_percent: Some(110),
            zoom_kind: SavedZoomKind::BestFit,
            gutter_at_top: true,
        })
        .unwrap();
        base.set_protection(ProtectionSettings {
            track_revisions: true,
            tracked_revisions: true,
            ..ProtectionSettings::default()
        })
        .unwrap();
        assert_eq!(base.endnote_placement(), EndnotePlacement::DocumentEnd);
        assert_eq!(base.saved_view().kind, SavedViewKind::Web);
        assert!(
            base.set_saved_view(SavedView {
                kind: SavedViewKind::Print,
                zoom_percent: Some(9),
                zoom_kind: SavedZoomKind::Default,
                gutter_at_top: false,
            })
            .is_err()
        );
        assert!(
            base.set_protection(ProtectionSettings {
                comments_or_read_only: true,
                form_fields: true,
                ..ProtectionSettings::default()
            })
            .is_err()
        );
        assert!(
            base.set_footnote_numbering(NoteNumberingRestart::Continuous, 0x4000)
                .is_err()
        );
    }

    #[test]
    fn rejects_reserved_values_invalid_dates_and_nonstandard_lengths() {
        let valid = DocumentProperties::word97_writer_bytes(false, false, false);
        for (offset, value) in [(0usize, 0x60u8), (2, 0x03), (54, 0x01)] {
            let mut data = valid.clone();
            data[offset] = value;
            assert!(DocumentProperties::parse_bytes(&data).is_err());
        }
        let mut data = valid.clone();
        data[18] = 1;
        assert!(DocumentProperties::parse_bytes(&data).is_err());
        let mut data = valid.clone();
        data[20..24].copy_from_slice(&packed_timestamp(2026, 13, 1, 0, 0, 1).to_le_bytes());
        assert!(DocumentProperties::parse_bytes(&data).is_err());
        let mut data = valid;
        data[82..84].copy_from_slice(&(1u16 | (9 << 3)).to_le_bytes());
        assert!(DocumentProperties::parse_bytes(&data).is_err());
        assert!(DocumentProperties::parse_bytes(&vec![0; 499]).is_err());
    }

    #[test]
    fn dispatches_word_2013_to_the_complete_typed_extension() {
        let mut data = vec![0u8; DocumentPropertyVersion::Word2013.byte_len()];
        // Required Dop97 typography defaults used by the versioned prefix.
        data[0x190..0x19a].copy_from_slice(&[0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81]);
        // DopMth defaults: centered-as-group, wrapped-left, display defaults.
        data[640..644].copy_from_slice(&(1u32 << 4 | 1 << 11 | 1 << 12).to_le_bytes());
        data[654..658].copy_from_slice(&120i32.to_le_bytes());
        data[658..662].copy_from_slice(&120i32.to_le_bytes());
        data[674..678].copy_from_slice(&1u32.to_le_bytes());
        data[690..694].copy_from_slice(&1u32.to_le_bytes());

        let properties = DocumentProperties::parse_bytes(&data).unwrap();
        match properties.versioned().unwrap() {
            Versioned::Word2013(value) => {
                assert!(value.chart_tracking_reference_based);
            },
            other => panic!("expected Word2013 properties, got {other:?}"),
        }
    }
}
