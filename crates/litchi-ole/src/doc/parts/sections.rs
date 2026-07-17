//! Section range, layout, and property-revision parsing for Word 97+ documents.

use super::fib::FileInformationBlock;
use super::revisions::RevisionAuthorTable;
use crate::doc::package::{DocError, Result};
use crate::doc::revision::{SectionRevisionMark, decode_dttm};
use crate::doc::section::{
    DocSection, PageOrientation, SectionBreakKind, SectionColumn, SectionColumnLayout,
    SectionMargins, SectionPageLayout, VerticalMargin,
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
            },
            column_count: 1,
            evenly_spaced: true,
            column_spacing_twips: 720,
            line_between: false,
            column_widths: [None; MAX_COLUMNS],
            column_spacings: [None; MAX_COLUMNS],
        }
    }
}

impl SectionProperties {
    fn apply(&mut self, sprm: &Sprm) -> Result<()> {
        match sprm.opcode {
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
            0xB017 => {
                self.page.header_distance_twips =
                    bounded_word(sprm, "sprmSDyaHdrTop", MAX_NON_NEGATIVE_TWIPS)?;
            },
            0xB018 => {
                self.page.footer_distance_twips =
                    bounded_word(sprm, "sprmSDyaHdrBottom", MAX_NON_NEGATIVE_TWIPS)?;
            },
            0x3019 => self.line_between = bool_operand(sprm, "sprmSLBetween")?,
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
            _ => {},
        }
        Ok(())
    }

    fn finish(self, start_cp: u32, end_cp: u32) -> Result<DocSection> {
        let columns = if self.evenly_spaced {
            SectionColumnLayout::Even {
                count: self.column_count,
                spacing_twips: self.column_spacing_twips,
                line_between: self.line_between,
            }
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
            SectionColumnLayout::Unequal {
                columns,
                line_between: self.line_between,
            }
        };
        Ok(DocSection {
            start_cp,
            end_cp,
            break_kind: self.break_kind,
            page: self.page,
            columns,
        })
    }
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
}
