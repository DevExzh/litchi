//! Legacy Word revision-save identifier table (`PLRSID`, MS-DOC 2.9.203).
//!
//! The table lists the revision-save identifiers (RSIDs) assigned in the
//! document. It is parsed as inert metadata: identifiers are never used to
//! drive revision behavior.

use super::fib::FileInformationBlock;
use crate::leniency::{RsidDefect, ToleranceReport};
use crate::package::{Error as PackageError, Result};

/// Table-pointer index of `fcPlrsid`/`lcbPlrsid`.
const PLRSID: usize = 113;
/// Fixed `PLRSID` header size in bytes: six 4-byte fields.
const HEADER_LEN: usize = 24;
/// `cbRsidInFile`: the mandated size of one RSID.
const CB_RSID_IN_FILE: u32 = 4;
/// `cbHeadExtraInFile`: the mandated header extension size.
const CB_HEAD_EXTRA_IN_FILE: u32 = 8;
/// `reserved1`: the mandated marker value.
const RESERVED1: u32 = 229;
/// Size in bytes of one RSID element.
const RSID_LEN: usize = 4;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

/// The revision-save identifiers assigned in a document (MS-DOC 2.9.203).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentRsids {
    ids: Vec<u32>,
}

impl DocumentRsids {
    /// The identifiers in table order.
    #[must_use]
    pub fn ids(&self) -> &[u32] {
        &self.ids
    }

    /// Whether an identifier is present in the table.
    #[must_use]
    pub fn contains(&self, id: u32) -> bool {
        self.ids.contains(&id)
    }

    /// Parse the `PLRSID` from the table stream, or `None` when the document
    /// predates revision-save identifiers.
    ///
    /// Word-compatible repairs are applied but their record is discarded; use
    /// [`Self::parse_recording`] to enumerate them.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<DocumentRsids>> {
        Self::parse_recording(fib, table_stream).map(|(table, _)| table)
    }

    /// Parse the `PLRSID`, returning the repairs [`RsidDefect`] lists
    /// alongside the table.
    pub(crate) fn parse_recording(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<(Option<DocumentRsids>, ToleranceReport)> {
        let mut tolerance = ToleranceReport::default();
        let Some((offset, length)) = fib.get_table_pointer(PLRSID) else {
            return Ok((None, tolerance));
        };
        if length == 0 {
            return Ok((None, tolerance));
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("PLRSID offset does not fit in memory"))?;
        let end = start
            .checked_add(
                usize::try_from(length)
                    .map_err(|_| corrupted("PLRSID length does not fit in memory"))?,
            )
            .ok_or_else(|| corrupted("PLRSID range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("PLRSID extends past the table stream"))?;
        let table = Self::parse_bytes(data, &mut tolerance)?;
        Ok((Some(table), tolerance))
    }

    /// Parse one `PLRSID` structure.
    ///
    /// `cbRsidInFile` and `cbHeadExtraInFile` determine the element layout,
    /// so a value other than the mandated one leaves the table unreadable
    /// (the caller degrades it to absent). `reserved1` describes nothing;
    /// like Word, a value other than 229 is ignored and recorded.
    fn parse_bytes(data: &[u8], tolerance: &mut ToleranceReport) -> Result<DocumentRsids> {
        if data.len() < HEADER_LEN {
            return Err(corrupted("PLRSID header is truncated"));
        }
        let read_u32 = |offset: usize| {
            u32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
        };
        let count = usize::try_from(read_u32(0))
            .map_err(|_| corrupted("PLRSID identifier count does not fit in memory"))?;
        if read_u32(4) != CB_RSID_IN_FILE {
            return Err(corrupted("PLRSID cbRsidInFile is not 4"));
        }
        if read_u32(8) != CB_HEAD_EXTRA_IN_FILE {
            return Err(corrupted("PLRSID cbHeadExtraInFile is not 8"));
        }
        let reserved1 = read_u32(12);
        if reserved1 != RESERVED1 {
            tolerance.record_wide(RsidDefect::ReservedMarker, reserved1);
        }
        let expected = HEADER_LEN
            .checked_add(
                count
                    .checked_mul(RSID_LEN)
                    .ok_or_else(|| corrupted("PLRSID identifier count overflows"))?,
            )
            .ok_or_else(|| corrupted("PLRSID identifier count overflows"))?;
        if data.len() != expected {
            return Err(corrupted("PLRSID identifier count does not match its size"));
        }
        let ids = data[HEADER_LEN..]
            .as_chunks::<RSID_LEN>()
            .0
            .iter()
            .map(|chunk| u32::from_le_bytes(*chunk))
            .collect();
        Ok(DocumentRsids { ids })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plrsid(ids: &[u32]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&(ids.len() as u32).to_le_bytes());
        data.extend_from_slice(&CB_RSID_IN_FILE.to_le_bytes());
        data.extend_from_slice(&CB_HEAD_EXTRA_IN_FILE.to_le_bytes());
        data.extend_from_slice(&RESERVED1.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        for id in ids {
            data.extend_from_slice(&id.to_le_bytes());
        }
        data
    }

    #[test]
    fn parses_identifier_table() {
        let mut tolerance = ToleranceReport::default();
        let parsed =
            DocumentRsids::parse_bytes(&plrsid(&[0x0042_0001, 0x00FF_AA55]), &mut tolerance)
                .unwrap();
        assert!(tolerance.is_clean());
        assert_eq!(parsed.ids(), &[0x0042_0001, 0x00FF_AA55]);
        assert!(parsed.contains(0x00FF_AA55));
        assert!(!parsed.contains(1));
    }

    /// MS-DOC 2.9.203's `reserved1` constant is a writer rule over an inert
    /// lookup table: a different value is ignored and recorded.
    #[test]
    fn ignores_and_records_a_wrong_reserved_marker() {
        use crate::leniency::{Defect, ToleratedDefect};

        let mut wrong_marker = plrsid(&[7]);
        wrong_marker[12..16].copy_from_slice(&0x1234u32.to_le_bytes());
        let mut tolerance = ToleranceReport::default();
        let parsed = DocumentRsids::parse_bytes(&wrong_marker, &mut tolerance).unwrap();
        assert_eq!(parsed.ids(), &[7]);
        assert_eq!(
            tolerance.defects(),
            &[ToleratedDefect {
                defect: Defect::Rsid(RsidDefect::ReservedMarker),
                context: 0x1234,
            }]
        );
    }

    #[test]
    fn rejects_malformed_tables() {
        // Truncated header.
        assert!(
            DocumentRsids::parse_bytes(&plrsid(&[1])[..20], &mut ToleranceReport::default())
                .is_err()
        );
        // Wrong cbRsidInFile.
        let mut wrong_size = plrsid(&[1]);
        wrong_size[4] = 8;
        assert!(DocumentRsids::parse_bytes(&wrong_size, &mut ToleranceReport::default()).is_err());
        // Wrong cbHeadExtraInFile.
        let mut wrong_extra = plrsid(&[1]);
        wrong_extra[8] = 4;
        assert!(DocumentRsids::parse_bytes(&wrong_extra, &mut ToleranceReport::default()).is_err());
        // Count disagreeing with the table size.
        let mut wrong_count = plrsid(&[1]);
        wrong_count[0] = 2;
        assert!(DocumentRsids::parse_bytes(&wrong_count, &mut ToleranceReport::default()).is_err());
    }
}
