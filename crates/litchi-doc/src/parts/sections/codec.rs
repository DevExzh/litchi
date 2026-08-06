//! Binary routing and SEPX/PlcSed decoding for the section owner.

use super::model::{Properties, SectionsTable, parse_revision, read_i32, read_u32};
use super::validation;
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
    parse_data(data, word_document, authors)
}

impl SectionsTable {
    /// Parse the section PLC and every referenced SEPX once.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<Self> {
        parse(fib, table_stream, word_document, authors)
    }

    #[cfg(test)]
    pub(super) fn parse_data(
        data: &[u8],
        word_document: &[u8],
        authors: &RevisionAuthorTable,
    ) -> Result<Self> {
        parse_data(data, word_document, authors)
    }
}

fn parse_data(
    data: &[u8],
    word_document: &[u8],
    authors: &RevisionAuthorTable,
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
    Ok(SectionsTable {
        sections,
        revisions,
    })
}
