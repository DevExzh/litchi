//! Word AutoSummary priority ranges (`PlcfAsumy`, MS-DOC 2.8.4).
//!
//! The PLCF assigns document-text ranges a priority for automatic
//! summarization. It is parsed as inert metadata: no summary is generated.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use crate::plcf::Plcf;

/// Table-pointer index of `fcPlcfAsumy`/`lcbPlcfAsumy`.
const PLCF_ASUMY: usize = 89;
/// Size in bytes of one `ASUMY` element (MS-DOC 2.9.3).
const ASUMY_SIZE: usize = 4;
/// Implementation limit on the number of AutoSummary ranges.
const MAX_ASUMY_ENTRIES: usize = 1_000_000;
/// A CP is valid only when it is less than `0x7FFF_FFFF` (MS-DOC 2.2.1).
const MAX_CP: u32 = 0x7FFF_FFFE;
/// `ASUMY.lLevel` is a positive signed 32-bit integer (MS-DOC 2.9.3).
const MAX_LEVEL: u32 = i32::MAX as u32;
/// A `PlcfAsumy` has one terminal CP in addition to one CP per element.
const MAX_ASUMY_TABLE_BYTES: usize = 4 + MAX_ASUMY_ENTRIES * (4 + ASUMY_SIZE);

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

/// One document-text range and its AutoSummary priority.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AutoSummaryRange {
    /// First character position of the range in the document parts.
    start: u32,
    /// Character position immediately after the range.
    end: u32,
    /// AutoSummary priority: smaller values mark more important text.
    level: u32,
}

impl AutoSummaryRange {
    /// Create one checked AutoSummary range.
    ///
    /// The range is half-open. Its CPs and priority are checked against the
    /// domains of `PlcfAsumy`; in particular, empty ranges and values outside
    /// the signed wire domains are rejected.
    pub fn new(start: u32, end: u32, level: u32) -> Result<Self> {
        let range = Self { start, end, level };
        validation::range(&range)?;
        Ok(range)
    }

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

    /// Return this range with a checked replacement priority.
    pub fn with_level(mut self, level: u32) -> Result<Self> {
        self.set_level(level)?;
        Ok(self)
    }

    /// Replace the priority after checking the `ASUMY.lLevel` wire domain.
    pub fn set_level(&mut self, level: u32) -> Result<()> {
        if !(1..=MAX_LEVEL).contains(&level) {
            return Err(corrupted(
                "ASUMY lLevel must be a positive signed 32-bit integer",
            ));
        }
        self.level = level;
        Ok(())
    }
}

/// A document's AutoSummary priority ranges (MS-DOC 2.8.4).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentAutoSummary {
    ranges: Vec<AutoSummaryRange>,
}

impl DocumentAutoSummary {
    /// Create an empty authoring table.
    pub const fn new() -> Self {
        Self { ranges: Vec::new() }
    }

    /// Create a table from already materialized ranges after validating the
    /// complete PLC invariant.
    pub fn try_new(ranges: Vec<AutoSummaryRange>) -> Result<Self> {
        validation::ranges(&ranges)?;
        Ok(Self { ranges })
    }

    /// The priority ranges in character-position order.
    pub fn ranges(&self) -> &[AutoSummaryRange] {
        &self.ranges
    }

    /// Number of authored priority ranges.
    pub fn len(&self) -> usize {
        self.ranges.len()
    }

    /// Whether this table has no priority ranges.
    pub fn is_empty(&self) -> bool {
        self.ranges.is_empty()
    }

    /// Append one range whose start continues the current terminal CP.
    pub fn push(&mut self, range: AutoSummaryRange) -> Result<()> {
        validation::range(&range)?;
        if self
            .ranges
            .last()
            .is_some_and(|previous| previous.end != range.start)
        {
            return Err(corrupted(
                "PlcfAsumy range must start at the current terminal CP",
            ));
        }
        if self.ranges.len() >= MAX_ASUMY_ENTRIES {
            return Err(corrupted("PlcfAsumy exceeds the one-million-entry cap"));
        }
        self.ranges.push(range);
        Ok(())
    }

    /// Insert one range while preserving the surrounding CP boundaries.
    pub fn insert(&mut self, index: usize, range: AutoSummaryRange) -> Result<()> {
        if index > self.ranges.len() {
            return Err(corrupted("PlcfAsumy insertion index is out of bounds"));
        }
        validation::range(&range)?;
        if index > 0 && self.ranges[index - 1].end != range.start {
            return Err(corrupted(
                "PlcfAsumy inserted range does not continue its predecessor",
            ));
        }
        if index < self.ranges.len() && range.end != self.ranges[index].start {
            return Err(corrupted(
                "PlcfAsumy inserted range does not reach its successor",
            ));
        }
        if self.ranges.len() >= MAX_ASUMY_ENTRIES {
            return Err(corrupted("PlcfAsumy exceeds the one-million-entry cap"));
        }
        self.ranges.insert(index, range);
        Ok(())
    }

    /// Replace one range without changing either neighboring CP boundary.
    pub fn replace(&mut self, index: usize, range: AutoSummaryRange) -> Result<()> {
        if index >= self.ranges.len() {
            return Err(corrupted("PlcfAsumy replacement index is out of bounds"));
        }
        validation::range(&range)?;
        if index > 0 && self.ranges[index - 1].end != range.start {
            return Err(corrupted(
                "PlcfAsumy replacement does not continue its predecessor",
            ));
        }
        if index + 1 < self.ranges.len() && range.end != self.ranges[index + 1].start {
            return Err(corrupted(
                "PlcfAsumy replacement does not reach its successor",
            ));
        }
        self.ranges[index] = range;
        Ok(())
    }

    /// Remove one range when the resulting PLC remains representable.
    ///
    /// Removing an interior range would leave a CP gap with no ASUMY value;
    /// that edit is rejected atomically. Removing the first or final range is
    /// representable because a PLC may begin later or terminate earlier.
    pub fn remove(&mut self, index: usize) -> Result<AutoSummaryRange> {
        if index >= self.ranges.len() {
            return Err(corrupted("PlcfAsumy removal index is out of bounds"));
        }
        if index > 0 && index + 1 < self.ranges.len() {
            return Err(corrupted(
                "cannot remove an interior PlcfAsumy range without a replacement",
            ));
        }
        Ok(self.ranges.remove(index))
    }

    /// Remove all priority ranges, leaving a valid empty PLC authoring model.
    pub fn clear(&mut self) {
        self.ranges.clear();
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
        if data.len() > MAX_ASUMY_TABLE_BYTES {
            return Err(corrupted(
                "PlcfAsumy exceeds the one-million-entry size cap",
            ));
        }
        let plcf = Plcf::parse(data, ASUMY_SIZE)
            .filter(|plcf| plcf.count() <= MAX_ASUMY_ENTRIES)
            .ok_or_else(|| corrupted("PlcfAsumy is malformed"))?;
        // `Plcf` tolerates trailing bytes; a whole PlcfAsumy must not
        // contain any.
        let expected = plcf
            .count()
            .checked_add(1)
            .and_then(|count| count.checked_mul(4))
            .and_then(|positions| {
                plcf.count()
                    .checked_mul(ASUMY_SIZE)
                    .and_then(|properties| positions.checked_add(properties))
            })
            .ok_or_else(|| corrupted("PlcfAsumy size overflows"))?;
        if expected != data.len() {
            return Err(corrupted("PlcfAsumy contains trailing bytes"));
        }

        let mut ranges = Vec::with_capacity(plcf.count());
        let mut previous_end = None;
        for index in 0..plcf.count() {
            let (start, end) = plcf
                .range(index)
                .ok_or_else(|| corrupted("PlcfAsumy is missing a CP range"))?;
            if start > MAX_CP || end > MAX_CP || start >= end {
                return Err(corrupted("PlcfAsumy has non-monotonic CPs"));
            }
            if previous_end.is_some_and(|previous| previous != start) {
                return Err(corrupted("PlcfAsumy has non-monotonic CPs"));
            }
            let raw_level = plcf
                .property(index)
                .map(|bytes| match bytes {
                    [a, b, c, d] => i32::from_le_bytes([*a, *b, *c, *d]),
                    _ => 0,
                })
                .ok_or_else(|| corrupted("PlcfAsumy is missing an ASUMY element"))?;
            if raw_level <= 0 {
                return Err(corrupted("PlcfAsumy level is not greater than zero"));
            }
            ranges.push(AutoSummaryRange {
                start,
                end,
                level: raw_level as u32,
            });
            previous_end = Some(end);
        }
        Self::try_new(ranges)
    }

    /// Serialize the complete `PlcfAsumy` payload.
    ///
    /// This writes only the priority PLC. It does not run or embed any
    /// AutoSummary generation algorithm.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validation::ranges(&self.ranges)?;
        codec::encode(&self.ranges)
    }
}

mod validation {
    use super::{AutoSummaryRange, MAX_ASUMY_ENTRIES, MAX_CP, MAX_LEVEL, Result, corrupted};

    pub(super) fn range(range: &AutoSummaryRange) -> Result<()> {
        if range.start > MAX_CP || range.end > MAX_CP {
            return Err(corrupted("PlcfAsumy CP is outside the MS-DOC CP domain"));
        }
        if range.start >= range.end {
            return Err(corrupted(
                "PlcfAsumy ranges must have strictly increasing CPs",
            ));
        }
        if !(1..=MAX_LEVEL).contains(&range.level) {
            return Err(corrupted(
                "ASUMY lLevel must be a positive signed 32-bit integer",
            ));
        }
        Ok(())
    }

    pub(super) fn ranges(ranges: &[AutoSummaryRange]) -> Result<()> {
        if ranges.len() > MAX_ASUMY_ENTRIES {
            return Err(corrupted("PlcfAsumy exceeds the one-million-entry cap"));
        }
        let mut previous_end = None;
        for item in ranges {
            range(item)?;
            if previous_end.is_some_and(|previous| previous != item.start) {
                return Err(corrupted(
                    "PlcfAsumy ranges must form one contiguous CP sequence",
                ));
            }
            previous_end = Some(item.end);
        }
        Ok(())
    }
}

mod codec {
    use super::{AutoSummaryRange, Result};

    pub(super) fn encode(ranges: &[AutoSummaryRange]) -> Result<Vec<u8>> {
        let size = ranges
            .len()
            .checked_mul(4 + super::ASUMY_SIZE)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| super::corrupted("PlcfAsumy serialized size overflows"))?;
        let mut data = Vec::with_capacity(size);
        for range in ranges {
            data.extend_from_slice(&range.start.to_le_bytes());
        }
        data.extend_from_slice(&ranges.last().map_or(0, |range| range.end).to_le_bytes());
        for range in ranges {
            data.extend_from_slice(&(range.level as i32).to_le_bytes());
        }
        Ok(data)
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
    fn serializes_positions_before_asumy_elements() {
        let summary = DocumentAutoSummary::try_new(vec![
            AutoSummaryRange::new(0, 10, 7).unwrap(),
            AutoSummaryRange::new(10, 20, 3).unwrap(),
        ])
        .unwrap();
        assert_eq!(
            summary.to_bytes().unwrap(),
            plcf_asumy(&[0, 10, 20], &[7, 3])
        );
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
