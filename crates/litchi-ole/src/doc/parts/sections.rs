//! Section range, layout, and property-revision parsing for Word 97+ documents.

use super::fib::FileInformationBlock;
use super::numbering::NumberFormat;
use super::revisions::RevisionAuthorTable;
use crate::doc::package::{DocError, Result};
use crate::doc::revision::{SectionRevisionMark, decode_dttm};
use crate::doc::section::{
    ChapterNumberSeparator, DocSection, LineNumberRestart, NoteNumberRestart, PageOrientation,
    SectionBehavior, SectionBreakKind, SectionColumn, SectionColumnLayout, SectionFootnotePosition,
    SectionLineNumbering, SectionMargins, SectionNoteSettings, SectionPageBorder,
    SectionPageBorderApplyTo, SectionPageBorderArt, SectionPageBorderColor, SectionPageBorderDepth,
    SectionPageBorderOffsetFrom, SectionPageBorderStyle, SectionPageBorders, SectionPageGrid,
    SectionPageGridMode, SectionPageLayout, SectionPageNumbering, SectionPaperSettings,
    SectionProtection, SectionTextFlow, SectionVerticalJustification, VerticalMargin,
};
use crate::sprm::{Sprm, parse_sprms};

const MAX_COLUMNS: usize = 44;
const MAX_NON_NEGATIVE_TWIPS: u16 = 31_680;
const MAX_VERTICAL_MARGIN_TWIPS: i16 = 31_665;

/// Parsed section ranges, layout, and property revision marks in document order.
#[derive(Debug, Clone, Default)]
pub struct SectionsTable {
    sections: Vec<DocSection>,
    revisions: Vec<SectionRevisionMark>,
}

impl SectionsTable {
    /// Parse the section PLC and every referenced SEPX once.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<Self> {
        let Some((offset, length)) = fib.get_table_pointer(6) else {
            return Ok(Self::default());
        };
        if length == 0 {
            return Ok(Self::default());
        }
        let start = usize::try_from(offset)
            .map_err(|_| DocError::Corrupted("PlcfSed offset is too large".to_string()))?;
        let length = usize::try_from(length)
            .map_err(|_| DocError::Corrupted("PlcfSed length is too large".to_string()))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| DocError::Corrupted("PlcfSed range overflows".to_string()))?;
        let data = table_stream.get(start..end).ok_or_else(|| {
            DocError::Corrupted("PlcfSed extends beyond the table stream".to_string())
        })?;
        Self::parse_data(data, word_document, authors)
    }

    fn parse_data(
        data: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<Self> {
        if data.len() < 20 || (data.len() - 4) % 16 != 0 {
            return Err(DocError::Corrupted(
                "PlcfSed does not contain complete CP and SED arrays".to_string(),
            ));
        }
        let section_count = (data.len() - 4) / 16;
        let sed_offset = (section_count + 1)
            .checked_mul(4)
            .ok_or_else(|| DocError::Corrupted("PlcfSed SED offset overflows".to_string()))?;
        let mut cps = Vec::with_capacity(section_count + 1);
        for index in 0..=section_count {
            cps.push(read_u32(data, index * 4, "PlcfSed CP")?);
        }
        if cps.windows(2).any(|pair| pair[0] >= pair[1]) {
            return Err(DocError::Corrupted(
                "PlcfSed character positions must be strictly increasing".to_string(),
            ));
        }

        let mut sections = Vec::with_capacity(section_count);
        let mut revisions = Vec::new();
        for index in 0..section_count {
            let record_offset = sed_offset
                .checked_add(index.checked_mul(12).ok_or_else(|| {
                    DocError::Corrupted("PlcfSed record offset overflows".to_string())
                })?)
                .ok_or_else(|| {
                    DocError::Corrupted("PlcfSed record offset overflows".to_string())
                })?;
            let fc_sepx = read_i32(data, record_offset + 2, "Sed.fcSepx")?;
            let mut properties = SectionProperties::default();
            let mut revision = None;

            if fc_sepx != -1 {
                if fc_sepx < 0 {
                    return Err(DocError::Corrupted(
                        "Sed.fcSepx contains an invalid negative offset".to_string(),
                    ));
                }
                let sepx_offset = usize::try_from(fc_sepx)
                    .map_err(|_| DocError::Corrupted("Sed.fcSepx is too large".to_string()))?;
                let size_end = sepx_offset
                    .checked_add(2)
                    .ok_or_else(|| DocError::Corrupted("SEPX size range overflows".to_string()))?;
                let size_bytes = word_document.get(sepx_offset..size_end).ok_or_else(|| {
                    DocError::Corrupted("SEPX size extends beyond WordDocument".to_string())
                })?;
                let grpprl_len = i16::from_le_bytes([size_bytes[0], size_bytes[1]]);
                if grpprl_len < 0 {
                    return Err(DocError::Corrupted(
                        "SEPX contains a negative grpprl size".to_string(),
                    ));
                }
                let grpprl_len = usize::from(grpprl_len as u16);
                let grpprl_end = size_end
                    .checked_add(grpprl_len)
                    .ok_or_else(|| DocError::Corrupted("SEPX range overflows".to_string()))?;
                let grpprl = word_document.get(size_end..grpprl_end).ok_or_else(|| {
                    DocError::Corrupted("SEPX extends beyond WordDocument".to_string())
                })?;
                let sprms = parse_sprms(grpprl);
                let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
                if consumed != grpprl.len() {
                    return Err(DocError::Corrupted(
                        "SEPX does not contain a whole number of SPRMs".to_string(),
                    ));
                }

                for sprm in &sprms {
                    if sprm.opcode == 0xD243 {
                        revision = parse_revision(sprm, cps[index], cps[index + 1], authors)?;
                    } else {
                        properties.apply(sprm)?;
                    }
                }
            }

            sections.push(properties.finish(cps[index], cps[index + 1])?);
            if let Some(revision) = revision {
                revisions.push(revision);
            }
        }
        Ok(Self {
            sections,
            revisions,
        })
    }

    /// Sections in main-document character-position order.
    pub fn sections(&self) -> &[DocSection] {
        &self.sections
    }

    /// Find the section containing `cp` using half-open section ranges.
    pub fn section_at_cp(&self, cp: u32) -> Option<&DocSection> {
        let index = self
            .sections
            .partition_point(|section| section.end_cp <= cp);
        self.sections
            .get(index)
            .filter(|section| section.start_cp <= cp && cp < section.end_cp)
    }

    /// Section property revision marks in document order.
    pub fn revisions(&self) -> &[SectionRevisionMark] {
        &self.revisions
    }
}

#[derive(Debug, Clone)]
struct SectionProperties {
    break_kind: SectionBreakKind,
    page: SectionPageLayout,
    column_count: u8,
    evenly_spaced: bool,
    column_spacing_twips: u16,
    line_between: bool,
    column_widths: [Option<u16>; MAX_COLUMNS],
    column_spacings: [Option<u16>; MAX_COLUMNS],
    page_numbering: SectionPageNumbering,
    line_numbering: SectionLineNumbering,
    notes: SectionNoteSettings,
    behavior: SectionBehavior,
    paper: SectionPaperSettings,
    page_borders: SectionPageBorders,
    page_grid: SectionPageGrid,
    text_flow: SectionTextFlow,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
            break_kind: SectionBreakKind::NewPage,
            page: SectionPageLayout {
                width_twips: 12_240,
                height_twips: 15_840,
                orientation: PageOrientation::Portrait,
                margins: SectionMargins {
                    left_twips: 1_800,
                    right_twips: 1_800,
                    top: VerticalMargin::Minimum(1_440),
                    bottom: VerticalMargin::Minimum(1_440),
                    gutter_twips: 0,
                },
                header_distance_twips: 720,
                footer_distance_twips: 720,
                vertical_justification: SectionVerticalJustification::Top,
            },
            column_count: 1,
            evenly_spaced: true,
            column_spacing_twips: 720,
            line_between: false,
            column_widths: [None; MAX_COLUMNS],
            column_spacings: [None; MAX_COLUMNS],
            page_numbering: SectionPageNumbering::default(),
            line_numbering: SectionLineNumbering::default(),
            notes: SectionNoteSettings::default(),
            behavior: SectionBehavior::default(),
            paper: SectionPaperSettings::default(),
            page_borders: SectionPageBorders::default(),
            page_grid: SectionPageGrid::default(),
            text_flow: SectionTextFlow::default(),
        }
    }
}

impl SectionProperties {
    fn apply(&mut self, sprm: &Sprm) -> Result<()> {
        match sprm.opcode {
            0x3000 => {
                self.page_numbering.chapter_separator = match byte_operand(sprm, "sprmScnsPgn")? {
                    0 => ChapterNumberSeparator::Hyphen,
                    1 => ChapterNumberSeparator::Period,
                    2 => ChapterNumberSeparator::Colon,
                    3 => ChapterNumberSeparator::EmDash,
                    4 => ChapterNumberSeparator::EnDash,
                    _ => return corrupted("sprmScnsPgn contains an invalid CNS value"),
                };
            },
            0x3001 => {
                self.page_numbering.chapter_heading_level =
                    match byte_operand(sprm, "sprmSiHeadingPgn")? {
                        0 => None,
                        level @ 1..=9 => Some(level),
                        _ => return corrupted("sprmSiHeadingPgn must be in [0, 9]"),
                    };
            },
            0xF203 => {
                let (index, value) = column_operand(sprm, "sprmSDxaColWidth")?;
                if !(718..=MAX_NON_NEGATIVE_TWIPS).contains(&value) {
                    return corrupted("sprmSDxaColWidth must be in [718, 31680]");
                }
                self.column_widths[index] = Some(value);
            },
            0xF204 => {
                let (index, value) = column_operand(sprm, "sprmSDxaColSpacing")?;
                if value > MAX_NON_NEGATIVE_TWIPS {
                    return corrupted("sprmSDxaColSpacing must be at most 31680");
                }
                self.column_spacings[index] = Some(value);
            },
            0x3005 => self.evenly_spaced = bool_operand(sprm, "sprmSFEvenlySpaced")?,
            0x3006 => {
                self.behavior.protection = if bool_operand(sprm, "sprmSFProtected")? {
                    SectionProtection::Unprotected
                } else {
                    SectionProtection::Protected
                };
            },
            0x5007 => {
                self.paper.first_page_source = Some(word_operand(sprm, "sprmSDmBinFirst")?);
            },
            0x5008 => {
                self.paper.other_page_source = Some(word_operand(sprm, "sprmSDmBinOther")?);
            },
            0x3009 => {
                self.break_kind = match byte_operand(sprm, "sprmSBkc")? {
                    0 => SectionBreakKind::Continuous,
                    1 => SectionBreakKind::NewColumn,
                    2 => SectionBreakKind::NewPage,
                    3 => SectionBreakKind::EvenPage,
                    4 => SectionBreakKind::OddPage,
                    _ => return corrupted("sprmSBkc contains an invalid section break kind"),
                };
            },
            0x300A => {
                self.behavior.different_first_page = bool_operand(sprm, "sprmSFTitlePage")?;
            },
            0x500B => {
                let count_minus_one = word_operand(sprm, "sprmSCcolumns")?;
                if count_minus_one > 43 {
                    return corrupted("sprmSCcolumns cannot specify more than 44 columns");
                }
                self.column_count = (count_minus_one + 1) as u8;
            },
            0x900C => {
                self.column_spacing_twips = non_negative_twips(sprm, "sprmSDxaColumns")?;
            },
            0x300E => {
                self.page_numbering.number_format = number_format8(sprm, "sprmSNfcPgn")?;
            },
            0x3011 => {
                self.page_numbering.restart = bool_operand(sprm, "sprmSFPgnRestart")?;
            },
            0x3012 => {
                self.notes.show_endnotes_at_section_end = bool_operand(sprm, "sprmSFEndnote")?;
            },
            0x3013 => {
                self.line_numbering.restart = match byte_operand(sprm, "sprmSLnc")? {
                    0 => LineNumberRestart::EachPage,
                    1 => LineNumberRestart::EachSection,
                    2 => LineNumberRestart::Continuous,
                    _ => return corrupted("sprmSLnc contains an invalid SLncOperand"),
                };
            },
            0x5015 => {
                let interval = word_operand(sprm, "sprmSNLnnMod")?;
                if interval > 100 {
                    return corrupted("sprmSNLnnMod must be in [0, 100]");
                }
                self.line_numbering.interval = interval;
            },
            0x9016 => {
                self.line_numbering.distance_twips = non_negative_twips(sprm, "sprmSDxaLnn")?;
            },
            0xB017 => {
                self.page.header_distance_twips =
                    bounded_word(sprm, "sprmSDyaHdrTop", MAX_NON_NEGATIVE_TWIPS)?;
            },
            0xB018 => {
                self.page.footer_distance_twips =
                    bounded_word(sprm, "sprmSDyaHdrBottom", MAX_NON_NEGATIVE_TWIPS)?;
            },
            0x3019 => self.line_between = bool_operand(sprm, "sprmSLBetween")?,
            0x301A => {
                self.page.vertical_justification = match byte_operand(sprm, "sprmSVjc")? {
                    0 => SectionVerticalJustification::Top,
                    1 => SectionVerticalJustification::Center,
                    2 => SectionVerticalJustification::Justified,
                    3 => SectionVerticalJustification::Bottom,
                    _ => return corrupted("sprmSVjc contains an invalid Vjc value"),
                };
            },
            0x501B => {
                let start_minus_one = word_operand(sprm, "sprmSLnnMin")?;
                if start_minus_one > 32_766 {
                    return corrupted("sprmSLnnMin must be at most 32766");
                }
                self.line_numbering.start_at = start_minus_one + 1;
            },
            0x501C => {
                let start = word_operand(sprm, "sprmSPgnStart97")?;
                if start > 32_766 {
                    return corrupted("sprmSPgnStart97 must be at most 32766");
                }
                self.page_numbering.start_at = u32::from(start);
            },
            0x301D => {
                self.page.orientation = match byte_operand(sprm, "sprmSBOrientation")? {
                    1 => PageOrientation::Portrait,
                    2 => PageOrientation::Landscape,
                    _ => return corrupted("sprmSBOrientation contains an invalid orientation"),
                };
            },
            0xB01F => {
                self.page.width_twips = page_dimension(sprm, "sprmSXaPage")?;
            },
            0xB020 => {
                self.page.height_twips = page_dimension(sprm, "sprmSYaPage")?;
            },
            0xB021 => {
                self.page.margins.left_twips =
                    bounded_word(sprm, "sprmSDxaLeft", MAX_NON_NEGATIVE_TWIPS)?;
            },
            0xB022 => {
                self.page.margins.right_twips =
                    bounded_word(sprm, "sprmSDxaRight", MAX_NON_NEGATIVE_TWIPS)?;
            },
            0x9023 => {
                self.page.margins.top = vertical_margin(sprm, "sprmSDyaTop")?;
            },
            0x9024 => {
                self.page.margins.bottom = vertical_margin(sprm, "sprmSDyaBottom")?;
            },
            0xB025 => {
                self.page.margins.gutter_twips = word_operand(sprm, "sprmSDzaGutter")?;
            },
            0x5026 => {
                self.paper.requested_paper_kind = Some(word_operand(sprm, "sprmSDmPaperReq")?);
            },
            0x3228 => {
                self.behavior.right_to_left = bool_operand(sprm, "sprmSFBiDi")?;
            },
            0x322A => {
                self.behavior.right_to_left_gutter = bool_operand(sprm, "sprmSFRTLGutter")?;
            },
            0x702B => {
                self.page_borders.top = page_border_operand(sprm, "sprmSBrcTop80")?;
            },
            0x702C => {
                self.page_borders.left = page_border_operand(sprm, "sprmSBrcLeft80")?;
            },
            0x702D => {
                self.page_borders.bottom = page_border_operand(sprm, "sprmSBrcBottom80")?;
            },
            0x702E => {
                self.page_borders.right = page_border_operand(sprm, "sprmSBrcRight80")?;
            },
            0x522F => {
                apply_page_border_properties(&mut self.page_borders, sprm)?;
            },
            0x7030 => {
                let value = dword_operand(sprm, "sprmSDxtCharSpace")? as i32;
                if !(-670_925..=6_488_064).contains(&value) {
                    return corrupted("sprmSDxtCharSpace must be in [-670925, 6488064]");
                }
                self.page_grid.character_pitch_adjustment = value;
            },
            0x9031 => {
                let value = word_operand(sprm, "sprmSDyaLinePitch")?;
                if !(1..=MAX_NON_NEGATIVE_TWIPS).contains(&value) {
                    return corrupted("sprmSDyaLinePitch must be in [1, 31680]");
                }
                self.page_grid.line_pitch_twips = Some(value);
            },
            0x5032 => {
                self.page_grid.mode = match word_operand(sprm, "sprmSClm")? {
                    0 => SectionPageGridMode::Disabled,
                    1 => SectionPageGridMode::CharactersAndLines,
                    2 => SectionPageGridMode::LinesOnly,
                    3 => SectionPageGridMode::EnforceCharacterGrid,
                    _ => return corrupted("sprmSClm contains an invalid SClmOperand value"),
                };
            },
            0x5033 => {
                self.text_flow = match word_operand(sprm, "sprmSTextFlow")? {
                    0 => SectionTextFlow::HorizontalNonAsian,
                    1 => SectionTextFlow::TopToBottomAsian,
                    2 => SectionTextFlow::BottomToTop,
                    3 => SectionTextFlow::TopToBottomNonAsian,
                    4 => SectionTextFlow::HorizontalAsian,
                    5 => SectionTextFlow::VerticalNonAsian,
                    _ => return corrupted("sprmSTextFlow contains an invalid MSOTXFL value"),
                };
            },
            0x3239 => {
                self.behavior.preserve_properties_for_revision = bool_operand(sprm, "sprmSWall")?;
            },
            0x703A => {
                self.behavior.revision_save_id = Some(dword_operand(sprm, "sprmSRsid")?);
            },
            0x303B => {
                self.notes.footnote_position = match byte_operand(sprm, "sprmSFpc")? {
                    1 => SectionFootnotePosition::BottomOfPage,
                    2 => SectionFootnotePosition::BeneathText,
                    _ => return corrupted("sprmSFpc contains an invalid SFpcOperand"),
                };
            },
            0x303C => {
                self.notes.footnote_restart = note_restart(sprm, "sprmSRncFtn", true)?;
            },
            0x303E => {
                self.notes.endnote_restart = note_restart(sprm, "sprmSRncEdn", false)?;
            },
            0x503F => {
                let value = word_operand(sprm, "sprmSNFtn")?;
                if value > 16_383 {
                    return corrupted("sprmSNFtn must be at most 16383");
                }
                self.notes.footnote_offset_operand = value;
            },
            0x5040 => {
                self.notes.footnote_number_format = number_format16(sprm, "sprmSNfcFtnRef")?;
            },
            0x5041 => {
                let value = word_operand(sprm, "sprmSNEdn")?;
                if value > 16_383 {
                    return corrupted("sprmSNEdn must be at most 16383");
                }
                self.notes.endnote_offset_operand = value;
            },
            0x5042 => {
                self.notes.endnote_number_format = number_format16(sprm, "sprmSNfcEdnRef")?;
            },
            0x7044 => {
                let start = dword_operand(sprm, "sprmSPgnStart")?;
                if start > 2_147_483_646 {
                    return corrupted("sprmSPgnStart must be at most 2147483646");
                }
                self.page_numbering.start_at = start;
            },
            _ => {},
        }
        Ok(())
    }

    fn finish(self, start_cp: u32, end_cp: u32) -> Result<DocSection> {
        if self.page_grid.mode != SectionPageGridMode::Disabled
            && self.page_grid.line_pitch_twips.is_none()
        {
            return corrupted("an enabled section document grid requires sprmSDyaLinePitch");
        }
        let columns = if self.evenly_spaced {
            SectionColumnLayout::even(
                self.column_count,
                self.column_spacing_twips,
                self.line_between,
            )
            .map_err(|error| DocError::Corrupted(error.to_string()))?
        } else {
            let count = usize::from(self.column_count);
            if self.column_widths[count..].iter().any(Option::is_some)
                || self.column_spacings[count.saturating_sub(1)..]
                    .iter()
                    .any(Option::is_some)
            {
                return corrupted(
                    "unequal-column operand index is outside the section column count",
                );
            }
            let mut columns = Vec::with_capacity(count);
            for index in 0..count {
                let width_twips = self.column_widths[index].ok_or_else(|| {
                    DocError::Corrupted(format!(
                        "unequal column {index} is missing sprmSDxaColWidth"
                    ))
                })?;
                let spacing_after_twips = if index + 1 < count {
                    Some(self.column_spacings[index].ok_or_else(|| {
                        DocError::Corrupted(format!(
                            "unequal column {index} is missing sprmSDxaColSpacing"
                        ))
                    })?)
                } else {
                    None
                };
                columns.push(SectionColumn {
                    width_twips,
                    spacing_after_twips,
                });
            }
            SectionColumnLayout::unequal(columns, self.line_between)
                .map_err(|error| DocError::Corrupted(error.to_string()))?
        };
        Ok(DocSection {
            start_cp,
            end_cp,
            break_kind: self.break_kind,
            page: self.page,
            columns,
            page_numbering: self.page_numbering,
            line_numbering: self.line_numbering,
            notes: self.notes,
            behavior: self.behavior,
            paper: self.paper,
            page_borders: self.page_borders,
            page_grid: self.page_grid,
            text_flow: self.text_flow,
        })
    }
}

fn page_border_operand(sprm: &Sprm, name: &str) -> Result<Option<SectionPageBorder>> {
    let operand = sprm.operand_bytes();
    if operand.len() != 4 {
        return corrupted(&format!("{name} operand must contain exactly 4 bytes"));
    }
    let style = match operand[1] {
        0x00 => None,
        0x01 => Some(SectionPageBorderStyle::Single),
        0x03 => Some(SectionPageBorderStyle::Double),
        0x05 => Some(SectionPageBorderStyle::Thick),
        0x06 => Some(SectionPageBorderStyle::Dotted),
        0x07 => Some(SectionPageBorderStyle::Dashed),
        0x08 => Some(SectionPageBorderStyle::DotDash),
        0x09 => Some(SectionPageBorderStyle::DotDotDash),
        0x0A => Some(SectionPageBorderStyle::Triple),
        0x0B => Some(SectionPageBorderStyle::ThinThickSmallGap),
        0x0C => Some(SectionPageBorderStyle::ThickThinSmallGap),
        0x0D => Some(SectionPageBorderStyle::ThinThickThinSmallGap),
        0x0E => Some(SectionPageBorderStyle::ThinThickMediumGap),
        0x0F => Some(SectionPageBorderStyle::ThickThinMediumGap),
        0x10 => Some(SectionPageBorderStyle::ThinThickThinMediumGap),
        0x11 => Some(SectionPageBorderStyle::ThinThickLargeGap),
        0x12 => Some(SectionPageBorderStyle::ThickThinLargeGap),
        0x13 => Some(SectionPageBorderStyle::ThinThickThinLargeGap),
        0x14 => Some(SectionPageBorderStyle::Wave),
        0x15 => Some(SectionPageBorderStyle::DoubleWave),
        0x16 => Some(SectionPageBorderStyle::DashSmallGap),
        0x17 => Some(SectionPageBorderStyle::DashDotStroked),
        0x18 => Some(SectionPageBorderStyle::ThreeDEmboss),
        0x19 => Some(SectionPageBorderStyle::ThreeDEngrave),
        code @ 0x40..=0xE3 => Some(SectionPageBorderStyle::Art(
            SectionPageBorderArt::try_from(code).expect("validated page-border art range"),
        )),
        invalid => {
            return corrupted(&format!(
                "{name} contains invalid Brc80 border type {invalid:#04x}"
            ));
        },
    };
    let color = match operand[2] {
        0x00 => SectionPageBorderColor::Automatic,
        0x01 => SectionPageBorderColor::Black,
        0x02 => SectionPageBorderColor::Blue,
        0x03 => SectionPageBorderColor::Cyan,
        0x04 => SectionPageBorderColor::Green,
        0x05 => SectionPageBorderColor::Magenta,
        0x06 => SectionPageBorderColor::Red,
        0x07 => SectionPageBorderColor::Yellow,
        0x08 => SectionPageBorderColor::White,
        0x09 => SectionPageBorderColor::DarkBlue,
        0x0A => SectionPageBorderColor::DarkCyan,
        0x0B => SectionPageBorderColor::DarkGreen,
        0x0C => SectionPageBorderColor::DarkMagenta,
        0x0D => SectionPageBorderColor::DarkRed,
        0x0E => SectionPageBorderColor::DarkYellow,
        0x0F => SectionPageBorderColor::DarkGray,
        0x10 => SectionPageBorderColor::LightGray,
        invalid => {
            return corrupted(&format!(
                "{name} contains invalid Ico color index {invalid:#04x}"
            ));
        },
    };
    let Some(style) = style else {
        return Ok(None);
    };
    let effects = operand[3];
    Ok(Some(SectionPageBorder {
        style,
        width_eighth_points: operand[0],
        color,
        spacing_points: effects & 0x1F,
        shadow: effects & 0x20 != 0,
        frame: effects & 0x40 != 0,
    }))
}

fn apply_page_border_properties(borders: &mut SectionPageBorders, sprm: &Sprm) -> Result<()> {
    let operand = sprm.operand_bytes();
    if operand.len() != 2 {
        return corrupted("sprmSPgbProp operand must contain exactly 2 bytes");
    }
    if operand[1] != 0 {
        return corrupted("sprmSPgbProp reserved byte must be zero");
    }
    borders.apply_to = match operand[0] & 0x07 {
        0 => SectionPageBorderApplyTo::AllPages,
        1 => SectionPageBorderApplyTo::FirstPage,
        2 => SectionPageBorderApplyTo::AllButFirstPage,
        _ => return corrupted("sprmSPgbProp contains an invalid PgbApplyTo value"),
    };
    borders.depth = match (operand[0] >> 3) & 0x03 {
        0 => SectionPageBorderDepth::InFront,
        1 => SectionPageBorderDepth::Behind,
        _ => return corrupted("sprmSPgbProp contains an invalid PgbPageDepth value"),
    };
    borders.offset_from = match (operand[0] >> 5) & 0x07 {
        0 => SectionPageBorderOffsetFrom::Text,
        1 => SectionPageBorderOffsetFrom::PageEdge,
        _ => return corrupted("sprmSPgbProp contains an invalid PgbOffsetFrom value"),
    };
    Ok(())
}

fn parse_revision(
    sprm: &Sprm,
    start: u32,
    end: u32,
    authors: &RevisionAuthorTable,
) -> Result<Option<SectionRevisionMark>> {
    let operand = sprm.operand_bytes();
    if operand.len() != 7 {
        return corrupted("sprmSPropRMark operand must contain exactly 7 bytes");
    }
    match operand[0] {
        0 => Ok(None),
        1 => {
            let signed_author = i16::from_le_bytes([operand[1], operand[2]]);
            let author_index = u16::try_from(signed_author).map_err(|_| {
                DocError::Corrupted("sprmSPropRMark author index is negative".to_string())
            })?;
            let author = authors.get(author_index).ok_or_else(|| {
                DocError::Corrupted("sprmSPropRMark author index is outside SttbfRMark".to_string())
            })?;
            let timestamp = u32::from_le_bytes(operand[3..7].try_into().unwrap());
            Ok(Some(SectionRevisionMark {
                start,
                end,
                author_index,
                author: author.to_string(),
                timestamp: decode_dttm(timestamp)?,
            }))
        },
        _ => corrupted("sprmSPropRMark must begin with a Boolean8 value"),
    }
}

fn byte_operand(sprm: &Sprm, name: &str) -> Result<u8> {
    if sprm.operand_bytes().len() != 1 {
        return corrupted(&format!("{name} operand must contain exactly 1 byte"));
    }
    Ok(sprm.operand_bytes()[0])
}

fn bool_operand(sprm: &Sprm, name: &str) -> Result<bool> {
    match byte_operand(sprm, name)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => corrupted(&format!("{name} must contain a Boolean8 value")),
    }
}

fn word_operand(sprm: &Sprm, name: &str) -> Result<u16> {
    if sprm.operand_bytes().len() != 2 {
        return corrupted(&format!("{name} operand must contain exactly 2 bytes"));
    }
    Ok(u16::from_le_bytes(sprm.operand_bytes().try_into().unwrap()))
}

fn dword_operand(sprm: &Sprm, name: &str) -> Result<u32> {
    if sprm.operand_bytes().len() != 4 {
        return corrupted(&format!("{name} operand must contain exactly 4 bytes"));
    }
    Ok(u32::from_le_bytes(sprm.operand_bytes().try_into().unwrap()))
}

fn number_format8(sprm: &Sprm, name: &str) -> Result<NumberFormat> {
    let raw = byte_operand(sprm, name)?;
    NumberFormat::try_from(raw)
        .map_err(|_| DocError::Corrupted(format!("{name} contains unknown MSONFC {raw:#04x}")))
}

fn number_format16(sprm: &Sprm, name: &str) -> Result<NumberFormat> {
    let raw = word_operand(sprm, name)?;
    let value = u8::try_from(raw).map_err(|_| {
        DocError::Corrupted(format!("{name} contains invalid 16-bit MSONFC {raw:#06x}"))
    })?;
    NumberFormat::try_from(value)
        .map_err(|_| DocError::Corrupted(format!("{name} contains unknown MSONFC {raw:#04x}")))
}

fn note_restart(sprm: &Sprm, name: &str, allow_each_page: bool) -> Result<NoteNumberRestart> {
    match byte_operand(sprm, name)? {
        0 => Ok(NoteNumberRestart::Continuous),
        1 => Ok(NoteNumberRestart::EachSection),
        2 if allow_each_page => Ok(NoteNumberRestart::EachPage),
        _ => corrupted(&format!("{name} contains an invalid Rnc value")),
    }
}

fn signed_word_operand(sprm: &Sprm, name: &str) -> Result<i16> {
    Ok(word_operand(sprm, name)? as i16)
}

fn bounded_word(sprm: &Sprm, name: &str, maximum: u16) -> Result<u16> {
    let value = word_operand(sprm, name)?;
    if value > maximum {
        return corrupted(&format!("{name} must be at most {maximum}"));
    }
    Ok(value)
}

fn non_negative_twips(sprm: &Sprm, name: &str) -> Result<u16> {
    let value = signed_word_operand(sprm, name)?;
    if value < 0 {
        return corrupted(&format!("{name} cannot be negative"));
    }
    Ok(value as u16)
}

fn page_dimension(sprm: &Sprm, name: &str) -> Result<u16> {
    let value = word_operand(sprm, name)?;
    if !(144..=MAX_NON_NEGATIVE_TWIPS).contains(&value) {
        return corrupted(&format!("{name} must be in [144, 31680]"));
    }
    Ok(value)
}

fn vertical_margin(sprm: &Sprm, name: &str) -> Result<VerticalMargin> {
    let value = signed_word_operand(sprm, name)?;
    if !(-MAX_VERTICAL_MARGIN_TWIPS..=MAX_VERTICAL_MARGIN_TWIPS).contains(&value) {
        return corrupted(&format!("{name} must be in [-31665, 31665]"));
    }
    Ok(VerticalMargin::from_signed_twips(value))
}

fn column_operand(sprm: &Sprm, name: &str) -> Result<(usize, u16)> {
    let operand = sprm.operand_bytes();
    if operand.len() != 3 {
        return corrupted(&format!("{name} operand must contain exactly 3 bytes"));
    }
    let index = usize::from(operand[0]);
    if index >= MAX_COLUMNS {
        return corrupted(&format!("{name} column index must be at most 43"));
    }
    Ok((index, u16::from_le_bytes([operand[1], operand[2]])))
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(DocError::Corrupted(message.to_string()))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| DocError::Corrupted(format!("{field} range overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| DocError::Corrupted(format!("{field} is truncated")))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    Ok(read_u32(data, offset, field)? as i32)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::{DocOpenOptions, Package};
    use std::path::{Path, PathBuf};

    fn fixed_sprm(opcode: u16, operand: &[u8]) -> Vec<u8> {
        let mut bytes = opcode.to_le_bytes().to_vec();
        bytes.extend_from_slice(operand);
        bytes
    }

    fn variable_sprm(opcode: u16, operand: &[u8]) -> Vec<u8> {
        let mut bytes = opcode.to_le_bytes().to_vec();
        bytes.push(operand.len() as u8);
        bytes.extend_from_slice(operand);
        bytes
    }

    fn build_section_data(grpprls: &[Option<Vec<u8>>]) -> (Vec<u8>, Vec<u8>) {
        let mut data = Vec::new();
        for cp in 0..=grpprls.len() {
            data.extend_from_slice(&((cp * 10) as u32).to_le_bytes());
        }
        let mut word_document = vec![0; 8];
        for grpprl in grpprls {
            data.extend_from_slice(&0u16.to_le_bytes());
            let fc_sepx = if let Some(grpprl) = grpprl {
                let offset = word_document.len() as i32;
                word_document.extend_from_slice(&(grpprl.len() as i16).to_le_bytes());
                word_document.extend_from_slice(grpprl);
                offset
            } else {
                -1
            };
            data.extend_from_slice(&fc_sepx.to_le_bytes());
            data.extend_from_slice(&0u16.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
        }
        (data, word_document)
    }

    fn parse_synthetic(grpprls: &[Option<Vec<u8>>]) -> Result<SectionsTable> {
        let (data, word_document) = build_section_data(grpprls);
        SectionsTable::parse_data(
            &data,
            &word_document,
            &RevisionAuthorTable::from_authors(&["Unknown", "Editor"]),
        )
    }

    #[test]
    fn parses_defaults_all_breaks_and_binary_lookup() {
        let mut grpprls = Vec::new();
        for value in 0u8..=4 {
            grpprls.push(Some(fixed_sprm(0x3009, &[value])));
        }
        grpprls.push(None);
        let parsed = parse_synthetic(&grpprls).unwrap();
        assert_eq!(parsed.sections().len(), 6);
        assert_eq!(
            parsed.sections()[0].break_kind,
            SectionBreakKind::Continuous
        );
        assert_eq!(parsed.sections()[1].break_kind, SectionBreakKind::NewColumn);
        assert_eq!(parsed.sections()[2].break_kind, SectionBreakKind::NewPage);
        assert_eq!(parsed.sections()[3].break_kind, SectionBreakKind::EvenPage);
        assert_eq!(parsed.sections()[4].break_kind, SectionBreakKind::OddPage);
        let defaults = &parsed.sections()[5];
        assert_eq!(defaults.break_kind, SectionBreakKind::NewPage);
        assert_eq!(
            (defaults.page.width_twips, defaults.page.height_twips),
            (12_240, 15_840)
        );
        assert_eq!(defaults.page.orientation, PageOrientation::Portrait);
        assert_eq!(defaults.page.margins.left_twips, 1_800);
        assert_eq!(defaults.columns.count(), 1);
        assert_eq!(parsed.section_at_cp(0).unwrap().start_cp, 0);
        assert_eq!(parsed.section_at_cp(10).unwrap().start_cp, 10);
        assert_eq!(parsed.section_at_cp(59).unwrap().end_cp, 60);
        assert!(parsed.section_at_cp(60).is_none());
    }

    #[test]
    fn parses_orientation_signed_margins_and_later_wins() {
        let mut grpprl = Vec::new();
        grpprl.extend(fixed_sprm(0x301D, &[1]));
        grpprl.extend(fixed_sprm(0x301D, &[2]));
        grpprl.extend(fixed_sprm(0xB01F, &15_840u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0xB020, &12_240u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0xB021, &1_000u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0xB022, &1_100u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x9023, &(-1_200i16).to_le_bytes()));
        grpprl.extend(fixed_sprm(0x9024, &1_300i16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0xB025, &200u16.to_le_bytes()));
        let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
        let page = parsed.sections()[0].page;
        assert_eq!(page.orientation, PageOrientation::Landscape);
        assert_eq!((page.width_twips, page.height_twips), (15_840, 12_240));
        assert_eq!(page.margins.top, VerticalMargin::Fixed(1_200));
        assert_eq!(page.margins.top.signed_twips(), -1_200);
        assert_eq!(page.margins.bottom, VerticalMargin::Minimum(1_300));
        assert_eq!(page.margins.gutter_twips, 200);
    }

    #[test]
    fn parses_even_and_unequal_columns() {
        let mut even = Vec::new();
        even.extend(fixed_sprm(0x500B, &2u16.to_le_bytes()));
        even.extend(fixed_sprm(0x900C, &900i16.to_le_bytes()));
        even.extend(fixed_sprm(0x3019, &[1]));

        let mut unequal = Vec::new();
        unequal.extend(fixed_sprm(0x500B, &2u16.to_le_bytes()));
        unequal.extend(fixed_sprm(0x3005, &[0]));
        for (index, width) in [1_000u16, 1_100, 1_200].into_iter().enumerate() {
            let mut operand = vec![index as u8];
            operand.extend_from_slice(&width.to_le_bytes());
            unequal.extend(fixed_sprm(0xF203, &operand));
        }
        for (index, spacing) in [100u16, 200].into_iter().enumerate() {
            let mut operand = vec![index as u8];
            operand.extend_from_slice(&spacing.to_le_bytes());
            unequal.extend(fixed_sprm(0xF204, &operand));
        }
        let parsed = parse_synthetic(&[Some(even), Some(unequal)]).unwrap();
        assert_eq!(
            parsed.sections()[0].columns,
            SectionColumnLayout::Even {
                count: 3,
                spacing_twips: 900,
                line_between: true,
            }
        );
        let SectionColumnLayout::Unequal { columns, .. } = &parsed.sections()[1].columns else {
            panic!("expected unequal columns");
        };
        assert_eq!(columns.len(), 3);
        assert_eq!(
            columns[0],
            SectionColumn {
                width_twips: 1_000,
                spacing_after_twips: Some(100)
            }
        );
        assert_eq!(
            columns[2],
            SectionColumn {
                width_twips: 1_200,
                spacing_after_twips: None
            }
        );
    }

    #[test]
    fn unequal_columns_are_order_independent_and_later_indexed_values_win() {
        let mut grpprl = Vec::new();
        for (index, width) in [1_000u16, 1_100, 1_200].into_iter().enumerate() {
            let mut operand = vec![index as u8];
            operand.extend_from_slice(&width.to_le_bytes());
            grpprl.extend(fixed_sprm(0xF203, &operand));
        }
        for (index, spacing) in [100u16, 200].into_iter().enumerate() {
            let mut operand = vec![index as u8];
            operand.extend_from_slice(&spacing.to_le_bytes());
            grpprl.extend(fixed_sprm(0xF204, &operand));
        }
        let mut replacement = vec![0];
        replacement.extend_from_slice(&1_500u16.to_le_bytes());
        grpprl.extend(fixed_sprm(0xF203, &replacement));
        grpprl.extend(fixed_sprm(0x3005, &[0]));
        grpprl.extend(fixed_sprm(0x500B, &2u16.to_le_bytes()));

        let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
        let SectionColumnLayout::Unequal { columns, .. } = &parsed.sections()[0].columns else {
            panic!("expected unequal columns");
        };
        assert_eq!(columns[0].width_twips, 1_500);
        assert_eq!(columns[0].spacing_after_twips, Some(100));
        assert_eq!(columns[2].width_twips, 1_200);
    }

    #[test]
    fn rejects_out_of_count_and_final_column_spacing_operands() {
        let invalid = |extra_index: u8| {
            let mut grpprl = Vec::new();
            grpprl.extend(fixed_sprm(0x3005, &[0]));
            grpprl.extend(fixed_sprm(0x500B, &1u16.to_le_bytes()));
            for index in 0..2u8 {
                let mut operand = vec![index];
                operand.extend_from_slice(&1_000u16.to_le_bytes());
                grpprl.extend(fixed_sprm(0xF203, &operand));
            }
            let mut spacing = vec![0];
            spacing.extend_from_slice(&100u16.to_le_bytes());
            grpprl.extend(fixed_sprm(0xF204, &spacing));
            let mut extra = vec![extra_index];
            extra.extend_from_slice(&200u16.to_le_bytes());
            grpprl.extend(fixed_sprm(0xF204, &extra));
            parse_synthetic(&[Some(grpprl)])
        };
        assert!(invalid(1).is_err());
        assert!(invalid(2).is_err());
    }

    #[test]
    fn parses_section_page_line_and_note_numbering() {
        let mut grpprl = Vec::new();
        grpprl.extend(fixed_sprm(0x3000, &[4]));
        grpprl.extend(fixed_sprm(0x3001, &[3]));
        grpprl.extend(fixed_sprm(0x300E, &[2]));
        grpprl.extend(fixed_sprm(0x3011, &[1]));
        grpprl.extend(fixed_sprm(0x501C, &123u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x7044, &70_000u32.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x3012, &[0]));
        grpprl.extend(fixed_sprm(0x3013, &[2]));
        grpprl.extend(fixed_sprm(0x5015, &3u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x9016, &360u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x501B, &6u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x301A, &[3]));
        grpprl.extend(fixed_sprm(0x303B, &[2]));
        grpprl.extend(fixed_sprm(0x303C, &[2]));
        grpprl.extend(fixed_sprm(0x303E, &[1]));
        grpprl.extend(fixed_sprm(0x503F, &6u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5040, &3u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5041, &8u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5042, &4u16.to_le_bytes()));

        let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
        let section = &parsed.sections()[0];
        assert_eq!(
            section.page_numbering.chapter_separator,
            ChapterNumberSeparator::EnDash
        );
        assert_eq!(section.page_numbering.chapter_heading_level, Some(3));
        assert_eq!(
            section.page_numbering.number_format,
            NumberFormat::LowerRoman
        );
        assert!(section.page_numbering.restart);
        assert_eq!(section.page_numbering.start_at, 70_000);
        assert_eq!(section.line_numbering.interval, 3);
        assert_eq!(
            section.line_numbering.restart,
            LineNumberRestart::Continuous
        );
        assert_eq!(section.line_numbering.distance_twips, 360);
        assert_eq!(section.line_numbering.start_at, 7);
        assert_eq!(
            section.page.vertical_justification,
            SectionVerticalJustification::Bottom
        );
        assert!(!section.notes.show_endnotes_at_section_end);
        assert_eq!(
            section.notes.footnote_position,
            SectionFootnotePosition::BeneathText
        );
        assert_eq!(section.notes.footnote_restart, NoteNumberRestart::EachPage);
        assert_eq!(
            section.notes.endnote_restart,
            NoteNumberRestart::EachSection
        );
        assert_eq!(section.notes.footnote_offset_operand, 6);
        assert_eq!(
            section.notes.footnote_number_format,
            NumberFormat::UpperLetter
        );
        assert_eq!(section.notes.endnote_offset_operand, 8);
        assert_eq!(
            section.notes.endnote_number_format,
            NumberFormat::LowerLetter
        );
    }

    #[test]
    fn rejects_invalid_section_numbering_operands() {
        for grpprl in [
            fixed_sprm(0x3000, &[5]),
            fixed_sprm(0x3001, &[10]),
            fixed_sprm(0x300E, &[60]),
            fixed_sprm(0x3013, &[3]),
            fixed_sprm(0x5015, &101u16.to_le_bytes()),
            fixed_sprm(0x501B, &32_767u16.to_le_bytes()),
            fixed_sprm(0x303B, &[0]),
            fixed_sprm(0x303E, &[2]),
            fixed_sprm(0x503F, &16_384u16.to_le_bytes()),
            fixed_sprm(0x5040, &0x0100u16.to_le_bytes()),
            fixed_sprm(0x7044, &2_147_483_647u32.to_le_bytes()),
        ] {
            assert!(parse_synthetic(&[Some(grpprl)]).is_err());
        }
    }

    #[test]
    fn parses_section_behavior_and_paper_settings() {
        let mut grpprl = Vec::new();
        grpprl.extend(fixed_sprm(0x3006, &[0]));
        grpprl.extend(fixed_sprm(0x3006, &[1]));
        grpprl.extend(fixed_sprm(0x5007, &7u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5008, &9u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x300A, &[1]));
        grpprl.extend(fixed_sprm(0x5026, &42u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x3228, &[1]));
        grpprl.extend(fixed_sprm(0x322A, &[1]));
        grpprl.extend(fixed_sprm(0x3239, &[1]));
        grpprl.extend(fixed_sprm(0x703A, &0xAABB_CCDDu32.to_le_bytes()));
        let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
        let section = &parsed.sections()[0];
        assert_eq!(section.behavior.protection, SectionProtection::Unprotected);
        assert!(section.behavior.different_first_page);
        assert!(section.behavior.right_to_left);
        assert!(section.behavior.right_to_left_gutter);
        assert!(section.behavior.preserve_properties_for_revision);
        assert_eq!(section.behavior.revision_save_id, Some(0xAABB_CCDD));
        assert_eq!(section.paper.first_page_source, Some(7));
        assert_eq!(section.paper.other_page_source, Some(9));
        assert_eq!(section.paper.requested_paper_kind, Some(42));
    }

    #[test]
    fn rejects_invalid_section_behavior_booleans() {
        for opcode in [0x3006, 0x300A, 0x3228, 0x322A, 0x3239] {
            assert!(parse_synthetic(&[Some(fixed_sprm(opcode, &[2]))]).is_err());
        }
    }

    #[test]
    fn retains_section_property_revisions() {
        let timestamp: u32 = 30 | (14 << 6) | (12 << 11) | (7 << 16) | (126 << 20) | (1 << 29);
        let mut operand = vec![1];
        operand.extend_from_slice(&1i16.to_le_bytes());
        operand.extend_from_slice(&timestamp.to_le_bytes());
        let parsed = parse_synthetic(&[Some(variable_sprm(0xD243, &operand))]).unwrap();
        let revision = &parsed.revisions()[0];
        assert_eq!((revision.start, revision.end), (0, 10));
        assert_eq!(revision.author_index, 1);
        assert_eq!(revision.author, "Editor");
        assert_eq!(revision.timestamp.unwrap().year, 2026);
    }

    #[test]
    fn rejects_malformed_tables_properties_and_columns() {
        let authors = RevisionAuthorTable::from_authors(&["Unknown"]);
        assert!(SectionsTable::parse_data(&[0; 19], &[], &authors).is_err());
        let (mut duplicate_cps, word_document) = build_section_data(&[None]);
        duplicate_cps[4..8].copy_from_slice(&0u32.to_le_bytes());
        assert!(SectionsTable::parse_data(&duplicate_cps, &word_document, &authors).is_err());

        for grpprl in [
            fixed_sprm(0x3009, &[5]),
            fixed_sprm(0x301D, &[0]),
            fixed_sprm(0xB01F, &143u16.to_le_bytes()),
            fixed_sprm(0x500B, &44u16.to_le_bytes()),
            fixed_sprm(0x9023, &i16::MIN.to_le_bytes()),
        ] {
            assert!(parse_synthetic(&[Some(grpprl)]).is_err());
        }

        let mut incomplete = Vec::new();
        incomplete.extend(fixed_sprm(0x500B, &1u16.to_le_bytes()));
        incomplete.extend(fixed_sprm(0x3005, &[0]));
        assert!(parse_synthetic(&[Some(incomplete)]).is_err());

        let (data, _) = build_section_data(&[Some(Vec::new())]);
        assert!(SectionsTable::parse_data(&data, &[0; 9], &authors).is_err());
    }

    fn poi_fixture(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/poi/test-data/document")
            .join(name)
    }

    #[test]
    fn opens_poi_bug53453_section_layout() {
        let mut package = Package::open(poi_fixture("Bug53453Section.doc")).unwrap();
        let document = package.document().unwrap();
        assert_eq!(document.sections().len(), 2);
        for section in document.sections() {
            assert_eq!(section.page.margins.left_twips, 1_440);
            assert_eq!(section.page.margins.right_twips, 1_440);
            assert_eq!(section.page.margins.top, VerticalMargin::Minimum(1_440));
            assert_eq!(section.page.margins.bottom, VerticalMargin::Minimum(1_440));
        }
        assert_eq!(document.sections()[0].columns.count(), 1);
        assert_eq!(document.sections()[1].columns.count(), 3);
    }

    #[test]
    fn exposes_layout_after_doc_decryption() {
        for (name, password) in [
            ("password_tika_binaryrc4.doc", "tika"),
            ("password_password_cryptoapi.doc", "password"),
        ] {
            let mut package = Package::open(poi_fixture(name)).unwrap();
            let document = package
                .document_with_options(DocOpenOptions {
                    password: Some(password),
                })
                .unwrap();
            assert!(!document.sections().is_empty());
            assert!(document.sections().iter().all(|section| {
                (144..=31_680).contains(&section.page.width_twips)
                    && (144..=31_680).contains(&section.page.height_twips)
            }));
        }
    }

    #[test]
    fn page_border_defaults_and_later_operands_win() {
        let defaults = parse_synthetic(&[None]).unwrap();
        assert_eq!(
            defaults.sections()[0].page_borders,
            SectionPageBorders::default()
        );

        let mut grpprl = Vec::new();
        grpprl.extend(fixed_sprm(0x702B, &[4, 0x06, 1, 0]));
        grpprl.extend(fixed_sprm(0x702B, &[8, 0x01, 6, 3]));
        grpprl.extend(fixed_sprm(0x702C, &[16, 0x03, 0, 0xA4]));
        grpprl.extend(fixed_sprm(0x702D, &[24, 0x40, 16, 0x45]));
        grpprl.extend(fixed_sprm(0x702E, &[0, 0, 0, 0]));
        grpprl.extend(fixed_sprm(0x522F, &[0, 0]));
        grpprl.extend(fixed_sprm(0x522F, &[0x2A, 0]));
        let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
        let borders = parsed.sections()[0].page_borders;

        assert_eq!(
            borders.top,
            Some(SectionPageBorder {
                style: SectionPageBorderStyle::Single,
                width_eighth_points: 8,
                color: SectionPageBorderColor::Red,
                spacing_points: 3,
                shadow: false,
                frame: false,
            })
        );
        assert_eq!(borders.left.unwrap().spacing_points, 4);
        assert!(borders.left.unwrap().shadow);
        let SectionPageBorderStyle::Art(art) = borders.bottom.unwrap().style else {
            panic!("expected an art page border");
        };
        assert_eq!(art.code(), 0x40);
        assert!(borders.bottom.unwrap().frame);
        assert_eq!(borders.right, None);
        assert_eq!(borders.apply_to, SectionPageBorderApplyTo::AllButFirstPage);
        assert_eq!(borders.depth, SectionPageBorderDepth::Behind);
        assert_eq!(borders.offset_from, SectionPageBorderOffsetFrom::PageEdge);
    }

    #[test]
    fn rejects_malformed_page_border_operands() {
        for border in [
            fixed_sprm(0x702B, &[8, 0x02, 0, 0]),
            fixed_sprm(0x702B, &[8, 0x1A, 0, 0]),
            fixed_sprm(0x702B, &[8, 0x1B, 0, 0]),
            fixed_sprm(0x702B, &[8, 0xE4, 0, 0]),
            fixed_sprm(0x702B, &[8, 0x01, 0x11, 0]),
            fixed_sprm(0x702B, &[8, 0x00, 0x11, 0]),
        ] {
            assert!(parse_synthetic(&[Some(border)]).is_err());
        }

        for properties in [
            fixed_sprm(0x522F, &[0x03, 0]),
            fixed_sprm(0x522F, &[0x10, 0]),
            fixed_sprm(0x522F, &[0x40, 0]),
            fixed_sprm(0x522F, &[0, 1]),
            fixed_sprm(0x522F, &[0]),
        ] {
            assert!(parse_synthetic(&[Some(properties)]).is_err());
        }
    }

    #[test]
    fn parses_page_grid_text_flow_defaults_and_later_overrides() {
        let defaults = parse_synthetic(&[None]).unwrap();
        assert_eq!(defaults.sections()[0].page_grid, SectionPageGrid::default());
        assert_eq!(
            defaults.sections()[0].text_flow,
            SectionTextFlow::HorizontalNonAsian
        );

        let mut grpprl = Vec::new();
        grpprl.extend(fixed_sprm(0x7030, &(-670_925i32).to_le_bytes()));
        grpprl.extend(fixed_sprm(0x7030, &6_144i32.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x9031, &360u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x9031, &480u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5032, &1u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5032, &3u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5033, &0u16.to_le_bytes()));
        grpprl.extend(fixed_sprm(0x5033, &5u16.to_le_bytes()));
        let parsed = parse_synthetic(&[Some(grpprl)]).unwrap();
        assert_eq!(
            parsed.sections()[0].page_grid,
            SectionPageGrid {
                mode: SectionPageGridMode::EnforceCharacterGrid,
                character_pitch_adjustment: 6_144,
                line_pitch_twips: Some(480),
            }
        );
        assert_eq!(
            parsed.sections()[0].text_flow,
            SectionTextFlow::VerticalNonAsian
        );

        let mut disabled = Vec::new();
        disabled.extend(fixed_sprm(0x7030, &1_024i32.to_le_bytes()));
        disabled.extend(fixed_sprm(0x9031, &240u16.to_le_bytes()));
        disabled.extend(fixed_sprm(0x5032, &2u16.to_le_bytes()));
        disabled.extend(fixed_sprm(0x5032, &0u16.to_le_bytes()));
        let parsed = parse_synthetic(&[Some(disabled)]).unwrap();
        assert_eq!(
            parsed.sections()[0].page_grid.mode,
            SectionPageGridMode::Disabled
        );
        assert_eq!(
            parsed.sections()[0].page_grid.character_pitch_adjustment,
            1_024
        );
        assert_eq!(parsed.sections()[0].page_grid.line_pitch_twips, Some(240));
    }

    #[test]
    fn rejects_malformed_page_grid_and_text_flow_operands() {
        for grpprl in [
            fixed_sprm(0x7030, &(-670_926i32).to_le_bytes()),
            fixed_sprm(0x7030, &6_488_065i32.to_le_bytes()),
            fixed_sprm(0x9031, &0u16.to_le_bytes()),
            fixed_sprm(0x9031, &31_681u16.to_le_bytes()),
            fixed_sprm(0x5032, &4u16.to_le_bytes()),
            fixed_sprm(0x5033, &6u16.to_le_bytes()),
            fixed_sprm(0x7030, &[0, 0, 0]),
            fixed_sprm(0x9031, &[1]),
        ] {
            assert!(parse_synthetic(&[Some(grpprl)]).is_err());
        }

        assert!(parse_synthetic(&[Some(fixed_sprm(0x5032, &1u16.to_le_bytes()))]).is_err());

        let mut later_enabled = Vec::new();
        later_enabled.extend(fixed_sprm(0x5032, &0u16.to_le_bytes()));
        later_enabled.extend(fixed_sprm(0x5032, &2u16.to_le_bytes()));
        assert!(parse_synthetic(&[Some(later_enabled)]).is_err());
    }
}
