//! Legacy Word revision-save identifier table (`PLRSID`, MS-DOC 2.9.203).
//!
//! The table lists the revision-save identifiers (RSIDs) assigned in the
//! document. It is parsed as inert metadata: identifiers are never used to
//! drive revision behavior.

use super::fib::FileInformationBlock;
use crate::doc::package::{DocError, Result};

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

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

/// The revision-save identifiers assigned in a document (MS-DOC 2.9.203).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentRsids {
    ids: Vec<u32>,
}

impl DocumentRsids {
    /// The identifiers in table order.
    pub fn ids(&self) -> &[u32] {
        &self.ids
    }

    /// Whether an identifier is present in the table.
    pub fn contains(&self, id: u32) -> bool {
        self.ids.contains(&id)
    }

    /// Parse the `PLRSID` from the table stream, or `None` when the document
    /// predates revision-save identifiers.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentRsids>> {
        let Some((offset, length)) = fib.get_table_pointer(PLRSID) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("PLRSID offset does not fit in memory"))?;
        let end = start
            .checked_add(usize::try_from(length).map_err(|_| {
                corrupted("PLRSID length does not fit in memory")
            })?)
            .ok_or_else(|| corrupted("PLRSID range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("PLRSID extends past the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one `PLRSID` structure.
    fn parse_bytes(data: &[u8]) -> Result<DocumentRsids> {
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
        if read_u32(12) != RESERVED1 {
            return Err(corrupted("PLRSID reserved1 is not 229"));
        }
        let expected = HEADER_LEN
            .checked_add(count.checked_mul(RSID_LEN).ok_or_else(|| {
                corrupted("PLRSID identifier count overflows")
            })?)
            .ok_or_else(|| corrupted("PLRSID identifier count overflows"))?;
        if data.len() != expected {
            return Err(corrupted("PLRSID identifier count does not match its size"));
        }
        let ids = data[HEADER_LEN..]
            .chunks_exact(RSID_LEN)
            .map(|chunk| u32::from_le_bytes(chunk.try_into().expect("length checked")))
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
        let parsed = DocumentRsids::parse_bytes(&plrsid(&[0x0042_0001, 0x00FF_AA55])).unwrap();
        assert_eq!(parsed.ids(), &[0x0042_0001, 0x00FF_AA55]);
        assert!(parsed.contains(0x00FF_AA55));
        assert!(!parsed.contains(1));
    }

    #[test]
    fn rejects_malformed_tables() {
        // Truncated header.
        assert!(DocumentRsids::parse_bytes(&plrsid(&[1])[..20]).is_err());
        // Wrong cbRsidInFile.
        let mut wrong_size = plrsid(&[1]);
        wrong_size[4] = 8;
        assert!(DocumentRsids::parse_bytes(&wrong_size).is_err());
        // Wrong cbHeadExtraInFile.
        let mut wrong_extra = plrsid(&[1]);
        wrong_extra[8] = 4;
        assert!(DocumentRsids::parse_bytes(&wrong_extra).is_err());
        // Wrong reserved1 marker.
        let mut wrong_marker = plrsid(&[1]);
        wrong_marker[12] = 0;
        assert!(DocumentRsids::parse_bytes(&wrong_marker).is_err());
        // Count disagreeing with the table size.
        let mut wrong_count = plrsid(&[1]);
        wrong_count[0] = 2;
        assert!(DocumentRsids::parse_bytes(&wrong_count).is_err());
    }
}
