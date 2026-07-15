//! Section property-revision parsing for Word 97+ documents.

use super::fib::FileInformationBlock;
use super::revisions::RevisionAuthorTable;
use crate::doc::package::{DocError, Result};
use crate::doc::revision::{SectionRevisionMark, decode_dttm};
use crate::sprm::parse_sprms;

/// Parsed section property revision marks in document order.
#[derive(Debug, Clone, Default)]
pub struct SectionRevisionsTable {
    revisions: Vec<SectionRevisionMark>,
}

impl SectionRevisionsTable {
    /// Parse the section PLC and each referenced SEPX.
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
        let sed_offset = (section_count + 1) * 4;
        let mut cps = Vec::with_capacity(section_count + 1);
        for index in 0..=section_count {
            cps.push(read_u32(data, index * 4, "PlcfSed CP")?);
        }
        if cps.windows(2).any(|pair| pair[0] >= pair[1]) {
            return Err(DocError::Corrupted(
                "PlcfSed character positions must be strictly increasing".to_string(),
            ));
        }

        let mut revisions = Vec::new();
        for index in 0..section_count {
            let record_offset = sed_offset + index * 12;
            let fc_sepx = read_i32(data, record_offset + 2, "Sed.fcSepx")?;
            if fc_sepx == -1 {
                continue;
            }
            if fc_sepx < 0 {
                return Err(DocError::Corrupted(
                    "Sed.fcSepx contains an invalid negative offset".to_string(),
                ));
            }
            let sepx_offset = usize::try_from(fc_sepx)
                .map_err(|_| DocError::Corrupted("Sed.fcSepx is too large".to_string()))?;
            let size_bytes = word_document
                .get(sepx_offset..sepx_offset + 2)
                .ok_or_else(|| {
                    DocError::Corrupted("SEPX size extends beyond WordDocument".to_string())
                })?;
            let grpprl_len = i16::from_le_bytes([size_bytes[0], size_bytes[1]]);
            if grpprl_len < 0 {
                return Err(DocError::Corrupted(
                    "SEPX contains a negative grpprl size".to_string(),
                ));
            }
            let grpprl_len =
                usize::try_from(grpprl_len).expect("non-negative i16 always fits in usize");
            let grpprl_start = sepx_offset + 2;
            let grpprl_end = grpprl_start
                .checked_add(grpprl_len)
                .ok_or_else(|| DocError::Corrupted("SEPX range overflows".to_string()))?;
            let grpprl = word_document.get(grpprl_start..grpprl_end).ok_or_else(|| {
                DocError::Corrupted("SEPX extends beyond WordDocument".to_string())
            })?;
            let sprms = parse_sprms(grpprl);
            let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
            if consumed != grpprl.len() {
                return Err(DocError::Corrupted(
                    "SEPX does not contain a whole number of SPRMs".to_string(),
                ));
            }

            let mut revision = None;
            for sprm in sprms.iter().filter(|sprm| sprm.opcode == 0xD243) {
                let operand = sprm.operand_bytes();
                if operand.len() != 7 {
                    return Err(DocError::Corrupted(
                        "sprmSPropRMark operand must contain exactly 7 bytes".to_string(),
                    ));
                }
                match operand[0] {
                    0 => revision = None,
                    1 => {
                        let signed_author = i16::from_le_bytes([operand[1], operand[2]]);
                        let author_index = u16::try_from(signed_author).map_err(|_| {
                            DocError::Corrupted(
                                "sprmSPropRMark author index is negative".to_string(),
                            )
                        })?;
                        let author = authors.get(author_index).ok_or_else(|| {
                            DocError::Corrupted(
                                "sprmSPropRMark author index is outside SttbfRMark".to_string(),
                            )
                        })?;
                        let packed_timestamp =
                            u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
                        revision = Some(SectionRevisionMark {
                            start: cps[index],
                            end: cps[index + 1],
                            author_index,
                            author: author.to_string(),
                            timestamp: decode_dttm(packed_timestamp)?,
                        });
                    },
                    _ => {
                        return Err(DocError::Corrupted(
                            "sprmSPropRMark must begin with a Boolean8 value".to_string(),
                        ));
                    },
                }
            }
            if let Some(revision) = revision {
                revisions.push(revision);
            }
        }
        Ok(Self { revisions })
    }

    /// Section property revision marks in document order.
    pub fn revisions(&self) -> &[SectionRevisionMark] {
        &self.revisions
    }
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| DocError::Corrupted(format!("{field} is truncated")))?;
    Ok(u32::from_le_bytes(
        bytes.try_into().expect("four-byte slice"),
    ))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    Ok(read_u32(data, offset, field)? as i32)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn section_data(fc_sepx: i32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&10u32.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&fc_sepx.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data
    }

    #[test]
    fn parses_section_property_revision_strictly() {
        let authors = RevisionAuthorTable::from_authors(&["Unknown", "Editor"]);
        let timestamp: u32 = 30 | (14 << 6) | (12 << 11) | (7 << 16) | (126 << 20) | (1 << 29);
        let mut word_document = vec![0; 32];
        let mut grpprl = Vec::new();
        grpprl.extend_from_slice(&0xD243u16.to_le_bytes());
        grpprl.push(7);
        grpprl.push(1);
        grpprl.extend_from_slice(&1i16.to_le_bytes());
        grpprl.extend_from_slice(&timestamp.to_le_bytes());
        word_document[8..10].copy_from_slice(&(grpprl.len() as u16).to_le_bytes());
        word_document[10..10 + grpprl.len()].copy_from_slice(&grpprl);

        let parsed =
            SectionRevisionsTable::parse_data(&section_data(8), &word_document, &authors).unwrap();
        let revision = &parsed.revisions()[0];
        assert_eq!((revision.start, revision.end), (0, 10));
        assert_eq!(revision.author_index, 1);
        assert_eq!(revision.author, "Editor");
        assert_eq!(revision.timestamp.unwrap().year, 2026);

        let mut invalid_flag = word_document.clone();
        invalid_flag[13] = 2;
        assert!(
            SectionRevisionsTable::parse_data(&section_data(8), &invalid_flag, &authors,).is_err()
        );

        let mut invalid_author = word_document;
        invalid_author[14..16].copy_from_slice(&(-1i16).to_le_bytes());
        assert!(
            SectionRevisionsTable::parse_data(&section_data(8), &invalid_author, &authors,)
                .is_err()
        );
    }

    #[test]
    fn rejects_malformed_section_tables_and_sepx_records() {
        let authors = RevisionAuthorTable::from_authors(&["Unknown"]);
        assert!(SectionRevisionsTable::parse_data(&[0; 19], &[], &authors).is_err());

        let mut duplicate_cps = section_data(-1);
        duplicate_cps[4..8].copy_from_slice(&0u32.to_le_bytes());
        assert!(SectionRevisionsTable::parse_data(&duplicate_cps, &[], &authors).is_err());

        assert!(SectionRevisionsTable::parse_data(&section_data(4), &[0; 5], &authors,).is_err());
    }
}
