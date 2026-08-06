//! Section range, layout, and property-revision parsing for Word 97+ documents.

use crate::package::{Error as PackageError, Result};
use crate::parts::numbering::NumberFormat;
use crate::parts::revisions::RevisionAuthorTable;
use crate::revision::{SectionRevisionMark, decode_dttm};
use crate::section::borders::{self, Borders};
use crate::section::columns::{Column, Layout};
use crate::section::{
    Behavior, BreakKind, ChapterNumberSeparator, FootnotePosition, LineNumberRestart,
    LineNumbering, Margins, NoteNumberRestart, NoteSettings, PageGrid, PageGridMode, PageLayout,
    PageNumbering, PageOrientation, PaperSettings, Protection, Section, TextFlow,
    VerticalJustification, VerticalMargin,
};
use crate::sprm::Sprm;

const MAX_NON_NEGATIVE_TWIPS: u16 = 31_680;
const MAX_VERTICAL_MARGIN_TWIPS: i16 = 31_665;

/// Parsed section ranges, layout, and property revision marks in document order.
#[derive(Debug, Clone, Default)]
pub struct SectionsTable {
    pub(super) sections: Vec<Section>,
    pub(super) revisions: Vec<SectionRevisionMark>,
}

impl SectionsTable {
    /// Sections in main-document character-position order.
    pub fn sections(&self) -> &[Section] {
        &self.sections
    }

    /// Find the section containing `cp` using half-open section ranges.
    pub fn section_at_cp(&self, cp: u32) -> Option<&Section> {
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
pub(super) struct Properties {
    break_kind: BreakKind,
    page: PageLayout,
    column_count: u8,
    evenly_spaced: bool,
    column_spacing_twips: u16,
    line_between: bool,
    column_widths: [Option<u16>; Layout::MAX_COLUMNS],
    column_spacings: [Option<u16>; Layout::MAX_COLUMNS],
    page_numbering: PageNumbering,
    line_numbering: LineNumbering,
    notes: NoteSettings,
    behavior: Behavior,
    paper: PaperSettings,
    page_borders: Borders,
    page_grid: PageGrid,
    text_flow: TextFlow,
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            break_kind: BreakKind::NewPage,
            page: PageLayout {
                width_twips: 12_240,
                height_twips: 15_840,
                orientation: PageOrientation::Portrait,
                margins: Margins {
                    left_twips: 1_800,
                    right_twips: 1_800,
                    top: VerticalMargin::Minimum(1_440),
                    bottom: VerticalMargin::Minimum(1_440),
                    gutter_twips: 0,
                },
                header_distance_twips: 720,
                footer_distance_twips: 720,
                vertical_justification: VerticalJustification::Top,
            },
            column_count: 1,
            evenly_spaced: true,
            column_spacing_twips: 720,
            line_between: false,
            column_widths: [None; Layout::MAX_COLUMNS],
            column_spacings: [None; Layout::MAX_COLUMNS],
            page_numbering: PageNumbering::default(),
            line_numbering: LineNumbering::default(),
            notes: NoteSettings::default(),
            behavior: Behavior::default(),
            paper: PaperSettings::default(),
            page_borders: Borders::default(),
            page_grid: PageGrid::default(),
            text_flow: TextFlow::default(),
        }
    }
}

impl Properties {
    pub(super) fn apply(&mut self, sprm: &Sprm) -> Result<()> {
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
                    Protection::Unprotected
                } else {
                    Protection::Protected
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
                    0 => BreakKind::Continuous,
                    1 => BreakKind::NewColumn,
                    2 => BreakKind::NewPage,
                    3 => BreakKind::EvenPage,
                    4 => BreakKind::OddPage,
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
                    0 => VerticalJustification::Top,
                    1 => VerticalJustification::Center,
                    2 => VerticalJustification::Justified,
                    3 => VerticalJustification::Bottom,
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
                self.page_borders.top = borders::decode_brc80(sprm, "sprmSBrcTop80")?;
            },
            0x702C => {
                self.page_borders.left = borders::decode_brc80(sprm, "sprmSBrcLeft80")?;
            },
            0x702D => {
                self.page_borders.bottom = borders::decode_brc80(sprm, "sprmSBrcBottom80")?;
            },
            0x702E => {
                self.page_borders.right = borders::decode_brc80(sprm, "sprmSBrcRight80")?;
            },
            0x522F => {
                borders::decode_pgb_prop(&mut self.page_borders, sprm)?;
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
                    0 => PageGridMode::Disabled,
                    1 => PageGridMode::CharactersAndLines,
                    2 => PageGridMode::LinesOnly,
                    3 => PageGridMode::EnforceCharacterGrid,
                    _ => return corrupted("sprmSClm contains an invalid SClmOperand value"),
                };
            },
            0x5033 => {
                self.text_flow = match word_operand(sprm, "sprmSTextFlow")? {
                    0 => TextFlow::HorizontalNonAsian,
                    1 => TextFlow::TopToBottomAsian,
                    2 => TextFlow::BottomToTop,
                    3 => TextFlow::TopToBottomNonAsian,
                    4 => TextFlow::HorizontalAsian,
                    5 => TextFlow::VerticalNonAsian,
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
                    1 => FootnotePosition::BottomOfPage,
                    2 => FootnotePosition::BeneathText,
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

    pub(super) fn finish(self, start_cp: u32, end_cp: u32) -> Result<Section> {
        if self.page_grid.mode != PageGridMode::Disabled
            && self.page_grid.line_pitch_twips.is_none()
        {
            return corrupted("an enabled section document grid requires sprmSDyaLinePitch");
        }
        let columns = if self.evenly_spaced {
            Layout::even(
                self.column_count,
                self.column_spacing_twips,
                self.line_between,
            )
            .map_err(|error| PackageError::Corrupted(error.to_string()))?
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
                    PackageError::Corrupted(format!(
                        "unequal column {index} is missing sprmSDxaColWidth"
                    ))
                })?;
                let spacing_after_twips = if index + 1 < count {
                    Some(self.column_spacings[index].ok_or_else(|| {
                        PackageError::Corrupted(format!(
                            "unequal column {index} is missing sprmSDxaColSpacing"
                        ))
                    })?)
                } else {
                    None
                };
                columns.push(
                    Column::from_parts(index, width_twips, spacing_after_twips)
                        .map_err(|error| PackageError::Corrupted(error.to_string()))?,
                );
            }
            Layout::unequal(columns, self.line_between)
                .map_err(|error| PackageError::Corrupted(error.to_string()))?
        };
        Ok(Section {
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

pub(super) fn parse_revision(
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
                PackageError::Corrupted("sprmSPropRMark author index is negative".to_string())
            })?;
            let author = authors.get(author_index).ok_or_else(|| {
                PackageError::Corrupted(
                    "sprmSPropRMark author index is outside SttbfRMark".to_string(),
                )
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
        .map_err(|_| PackageError::Corrupted(format!("{name} contains unknown MSONFC {raw:#04x}")))
}

fn number_format16(sprm: &Sprm, name: &str) -> Result<NumberFormat> {
    let raw = word_operand(sprm, name)?;
    let value = u8::try_from(raw).map_err(|_| {
        PackageError::Corrupted(format!("{name} contains invalid 16-bit MSONFC {raw:#06x}"))
    })?;
    NumberFormat::try_from(value)
        .map_err(|_| PackageError::Corrupted(format!("{name} contains unknown MSONFC {raw:#04x}")))
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
    if index >= Layout::MAX_COLUMNS {
        return corrupted(&format!("{name} column index must be at most 43"));
    }
    Ok((index, u16::from_le_bytes([operand[1], operand[2]])))
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(PackageError::Corrupted(message.to_string()))
}

pub(super) fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| PackageError::Corrupted(format!("{field} range overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| PackageError::Corrupted(format!("{field} is truncated")))?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

pub(super) fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    Ok(read_u32(data, offset, field)? as i32)
}
