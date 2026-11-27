//! Binary routing and SEPX/PlcSed decoding for the section owner.

use super::model::{Properties, SectionsTable, parse_revision, read_i32, read_u32};
use super::validation;
use crate::leniency::ToleranceReport;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use crate::parts::revisions::RevisionAuthorTable;
use crate::sprm::parse_sprms;

/// Decode the section PLC routed by the FIB.
pub(super) fn parse(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    word_document: &[u8],
    authors: &RevisionAuthorTable,
    tolerance: &mut ToleranceReport,
) -> Result<SectionsTable> {
    let Some((offset, length)) = fib.get_table_pointer(6) else {
        return Ok(SectionsTable::default());
    };
    if length == 0 {
        return Ok(SectionsTable::default());
    }
    let start = usize::try_from(offset)
        .map_err(|_| PackageError::Corrupted("PlcfSed offset is too large".to_string()))?;
    let length = usize::try_from(length)
        .map_err(|_| PackageError::Corrupted("PlcfSed length is too large".to_string()))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| PackageError::Corrupted("PlcfSed range overflows".to_string()))?;
    let data = table_stream.get(start..end).ok_or_else(|| {
        PackageError::Corrupted("PlcfSed extends beyond the table stream".to_string())
    })?;
    parse_data(
        data,
        word_document,
        authors,
        tolerance,
        Some(fib.get_main_doc_range().1),
    )
}

impl SectionsTable {
    /// Parse the section PLC and every referenced SEPX once.
    ///
    /// Word-compatible per-SPRM repairs are applied but their record is
    /// discarded by this convenience parser.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<Self> {
        Self::parse_recording(fib, table_stream, word_document, authors).map(|(table, _)| table)
    }

    /// Parse the section PLC, returning the repairs
    /// [`crate::SectionDefect`] lists alongside the table.
    pub(crate) fn parse_recording(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<(Self, ToleranceReport)> {
        let mut tolerance = ToleranceReport::default();
        let table = parse(fib, table_stream, word_document, authors, &mut tolerance)?;
        Ok((table, tolerance))
    }

    #[cfg(test)]
    pub(super) fn parse_data(
        data: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<Self> {
        parse_data(
            data,
            word_document,
            authors,
            &mut ToleranceReport::default(),
            None,
        )
    }

    #[cfg(test)]
    pub(super) fn parse_data_recording(
        data: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<(Self, ToleranceReport)> {
        let mut tolerance = ToleranceReport::default();
        let table = parse_data(data, word_document, authors, &mut tolerance, None)?;
        Ok((table, tolerance))
    }
}

fn parse_data(
    data: &[u8],
    word_document: &[u8],
    authors: &RevisionAuthorTable,
    tolerance: &mut ToleranceReport,
    main_end: Option<u32>,
) -> Result<SectionsTable> {
    if data.len() < 20 || !(data.len() - 4).is_multiple_of(16) {
        return Err(PackageError::Corrupted(
            "PlcfSed does not contain complete CP and SED arrays".to_string(),
        ));
    }
    let section_count = (data.len() - 4) / 16;
    let sed_offset = (section_count + 1)
        .checked_mul(4)
        .ok_or_else(|| PackageError::Corrupted("PlcfSed SED offset overflows".to_string()))?;
    let mut cps = Vec::with_capacity(section_count + 1);
    for index in 0..=section_count {
        cps.push(read_u32(data, index * 4, "PlcfSed CP")?);
    }
    validation::character_positions(&cps)?;
    if let Some(main_end) = main_end {
        // MS-DOC 2.8.26 allows the stored sentinel to extend into later
        // document parts. Those parts are not members of the last section.
        if cps[section_count] < main_end
            || cps[..section_count]
                .iter()
                .any(|&start| start > main_end || (start == main_end && main_end != 0))
        {
            return Err(PackageError::Corrupted(
                "PlcfSed does not delimit sections of the main document".into(),
            ));
        }
        cps[section_count] = main_end;
    }

    let mut sections = Vec::with_capacity(section_count);
    let mut revisions = Vec::new();
    for index in 0..section_count {
        let record_offset = sed_offset
            .checked_add(index.checked_mul(12).ok_or_else(|| {
                PackageError::Corrupted("PlcfSed record offset overflows".to_string())
            })?)
            .ok_or_else(|| {
                PackageError::Corrupted("PlcfSed record offset overflows".to_string())
            })?;
        let fc_sepx = read_i32(data, record_offset + 2, "Sed.fcSepx")?;
        let mut properties = Properties::default();
        let mut revision = None;

        if fc_sepx != -1 {
            if fc_sepx < 0 {
                return Err(PackageError::Corrupted(
                    "Sed.fcSepx contains an invalid negative offset".to_string(),
                ));
            }
            let sepx_offset = usize::try_from(fc_sepx)
                .map_err(|_| PackageError::Corrupted("Sed.fcSepx is too large".to_string()))?;
            let size_end = sepx_offset
                .checked_add(2)
                .ok_or_else(|| PackageError::Corrupted("SEPX size range overflows".to_string()))?;
            let size_bytes = word_document.get(sepx_offset..size_end).ok_or_else(|| {
                PackageError::Corrupted("SEPX size extends beyond WordDocument".to_string())
            })?;
            let grpprl_len = i16::from_le_bytes([size_bytes[0], size_bytes[1]]);
            if grpprl_len < 0 {
                return Err(PackageError::Corrupted(
                    "SEPX contains a negative grpprl size".to_string(),
                ));
            }
            let grpprl_len = usize::from(grpprl_len as u16);
            let grpprl_end = size_end
                .checked_add(grpprl_len)
                .ok_or_else(|| PackageError::Corrupted("SEPX range overflows".to_string()))?;
            let grpprl = word_document.get(size_end..grpprl_end).ok_or_else(|| {
                PackageError::Corrupted("SEPX extends beyond WordDocument".to_string())
            })?;
            let sprms = parse_sprms(grpprl)?;
            let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
            if consumed != grpprl.len() {
                return Err(PackageError::Corrupted(
                    "SEPX does not contain a whole number of SPRMs".to_string(),
                ));
            }

            for sprm in &sprms {
                if sprm.opcode == 0xD243 {
                    revision =
                        parse_revision(sprm, cps[index], cps[index + 1], authors, tolerance)?;
                } else {
                    properties.apply(sprm, tolerance)?;
                }
            }
        }

        sections.push(properties.finish(cps[index], cps[index + 1])?);
        if let Some(revision) = revision {
            revisions.push(revision);
        }
    }
    Ok(SectionsTable {
        sections,
        revisions,
    })
}

#[cfg(test)]
mod main_story_tests {
    use super::*;

    #[test]
    fn section_and_revision_ranges_share_main_end_without_mutating_wire_data() {
        // One section, a trailing sentinel in secondary text, and a revision.
        let mut data = [0u32.to_le_bytes(), 30u32.to_le_bytes()].concat();
        data.extend_from_slice(&[0, 0]);
        data.extend_from_slice(&8i32.to_le_bytes());
        data.extend_from_slice(&[0; 6]);
        let mut word = vec![0; 8];
        let sprms = [0x43, 0xd2, 7, 1, 0, 0, 0, 0, 0, 0];
        word.extend_from_slice(&(sprms.len() as i16).to_le_bytes());
        word.extend_from_slice(&sprms);
        let authors = RevisionAuthorTable::from_authors(&["Editor"]);
        let original = data.clone();
        let parsed = parse_data(
            &data,
            &word,
            &authors,
            &mut ToleranceReport::default(),
            Some(20),
        )
        .unwrap();
        assert_eq!(
            (parsed.sections()[0].start_cp, parsed.sections()[0].end_cp),
            (0, 20)
        );
        assert_eq!(
            (parsed.revisions()[0].start, parsed.revisions()[0].end),
            (0, 20)
        );
        assert!(parsed.section_at_cp(20).is_none());
        assert_eq!(data, original);
        assert!(
            parse_data(
                &data,
                &word,
                &authors,
                &mut ToleranceReport::default(),
                Some(31)
            )
            .is_err()
        );
        // A nonfinal start at/past the main-story end cannot be clamped away.
        let mut two = [0u32.to_le_bytes(), 20u32.to_le_bytes(), 30u32.to_le_bytes()].concat();
        for _ in 0..2 {
            two.extend_from_slice(&[0, 0, 255, 255, 255, 255, 0, 0, 0, 0, 0, 0]);
        }
        for end in [19, 20] {
            assert!(
                parse_data(
                    &two,
                    &[],
                    &authors,
                    &mut ToleranceReport::default(),
                    Some(end)
                )
                .is_err()
            );
        }
    }
}
