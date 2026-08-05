//! Word AutoSummary priority ranges (`PlcfAsumy`, MS-DOC 2.8.4).
//!
//! The PLCF assigns every main-document text range a priority for automatic
//! summarization. It is parsed as inert metadata: no summary is generated.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;
use crate::plcf::Plcf;

/// Table-pointer index of `fcPlcfAsumy`/`lcbPlcfAsumy`.
const PLCF_ASUMY: usize = 89;
/// Size in bytes of one `ASUMY` element (MS-DOC 2.9.3).
const ASUMY_SIZE: usize = 4;
/// Implementation limit on the number of AutoSummary ranges.
const MAX_ASUMY_ENTRIES: usize = 1_000_000;

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

/// One main-document text range and its AutoSummary priority.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AutoSummaryRange {
    /// First character position of the range in the main document.
    start: u32,
    /// Character position immediately after the range.
    end: u32,
    /// AutoSummary priority: smaller values mark more important text.
    level: u32,
}

impl AutoSummaryRange {
    /// First character position of the range.
    pub fn start(&self) -> u32 {
        self.start
    }

    /// Character position immediately after the range.
    pub fn end(&self) -> u32 {
        self.end
    }

    /// AutoSummary priority of the range. A smaller number implies greater
    /// importance of the text range to the summary.
    pub fn level(&self) -> u32 {
        self.level
    }
}

/// A document's AutoSummary priority ranges (MS-DOC 2.8.4).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentAutoSummary {
    ranges: Vec<AutoSummaryRange>,
}

impl DocumentAutoSummary {
    /// The priority ranges in character-position order.
    pub fn ranges(&self) -> &[AutoSummaryRange] {
        &self.ranges
    }

    /// Parse the `PlcfAsumy` from the table stream, or `None` when the
    /// document carries no AutoSummary priorities.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentAutoSummary>> {
        let Some((offset, length)) = fib.get_table_pointer(PLCF_ASUMY) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("PlcfAsumy offset does not fit in memory"))?;
        let end = start
            .checked_add(
                usize::try_from(length)
                    .map_err(|_| corrupted("PlcfAsumy length does not fit in memory"))?,
            )
            .ok_or_else(|| corrupted("PlcfAsumy range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("PlcfAsumy extends past the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete `PlcfAsumy` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<DocumentAutoSummary> {
        let plcf = Plcf::parse(data, ASUMY_SIZE)
            .filter(|plcf| plcf.count() <= MAX_ASUMY_ENTRIES)
            .ok_or_else(|| corrupted("PlcfAsumy is malformed"))?;
        // `Plcf` tolerates trailing bytes; a whole PlcfAsumy must not
        // contain any.
        let expected = 4 * (plcf.count() + 1) + plcf.count() * ASUMY_SIZE;
        if expected != data.len() {
            return Err(corrupted("PlcfAsumy contains trailing bytes"));
        }

        let mut ranges = Vec::with_capacity(plcf.count());
        let mut previous_end = None;
        for index in 0..plcf.count() {
            let (start, end) = plcf
                .range(index)
                .ok_or_else(|| corrupted("PlcfAsumy is missing a CP range"))?;
            if start > end || previous_end.is_some_and(|previous| previous > start) {
                return Err(corrupted("PlcfAsumy has non-monotonic CPs"));
            }
            previous_end = Some(end);
            let raw_level = plcf
                .property(index)
                .map(|bytes| i32::from_le_bytes(bytes.try_into().expect("element size checked")))
                .ok_or_else(|| corrupted("PlcfAsumy is missing an ASUMY element"))?;
            if raw_level <= 0 {
                return Err(corrupted("PlcfAsumy level is not greater than zero"));
            }
            ranges.push(AutoSummaryRange {
                start,
                end,
                level: raw_level as u32,
            });
        }
        Ok(DocumentAutoSummary { ranges })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plcf_asumy(cps: &[u32], levels: &[i32]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        for level in levels {
            data.extend_from_slice(&level.to_le_bytes());
        }
        data
    }

    #[test]
    fn parses_priority_ranges() {
        let data = plcf_asumy(&[0, 120, 480, 1024], &[1, 3, 2]);
        let parsed = DocumentAutoSummary::parse_bytes(&data).unwrap();
        assert_eq!(parsed.ranges().len(), 3);
        assert_eq!(parsed.ranges()[0].start(), 0);
        assert_eq!(parsed.ranges()[0].end(), 120);
        assert_eq!(parsed.ranges()[0].level(), 1);
        assert_eq!(parsed.ranges()[2].level(), 2);
    }

    #[test]
    fn rejects_malformed_tables() {
        // Level of zero or negative violates ASUMY.lLevel.
        assert!(DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0, 10], &[0])).is_err());
        assert!(DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0, 10], &[-1])).is_err());
        // Non-monotonic CPs.
        assert!(DocumentAutoSummary::parse_bytes(&plcf_asumy(&[10, 0], &[1])).is_err());
        // Byte count not divisible into a whole PLCF.
        assert!(DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0, 10], &[1])[..9]).is_err());
    }
}
