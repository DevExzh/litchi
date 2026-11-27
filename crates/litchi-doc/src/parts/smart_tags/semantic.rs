//! Semantic ownership and fixed FIB/table topology for legacy smart tags.

use super::{
    FACTOID_DATA, FileInformationBlock, PLCF_BKF_FACTOID, PLCF_BKL_FACTOID, PLCF_FACTOID,
    STTBF_BKMK_FACTOID,
};
use crate::package::{Error as PackageError, Result};
use std::sync::Arc;

/// One FIB range owned by a legacy smart-tag table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableRange {
    /// Byte offset in the selected table stream.
    pub offset: u32,
    /// Byte length addressed by the FIB.
    pub length: u32,
}

impl TableRange {
    /// Exclusive end offset, if it fits in the FIB's 32-bit address space.
    #[must_use]
    pub fn end(self) -> Option<u32> {
        self.offset.checked_add(self.length)
    }

    pub(super) fn as_usize(self, stream_len: usize) -> Result<std::ops::Range<usize>> {
        let start = usize::try_from(self.offset)
            .map_err(|_| corrupted("smart-tag table offset exceeds usize"))?;
        let length = usize::try_from(self.length)
            .map_err(|_| corrupted("smart-tag table length exceeds usize"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("smart-tag table range overflows usize"))?;
        if end > stream_len {
            return Err(corrupted("smart-tag table range exceeds the table stream"));
        }
        Ok(start..end)
    }
}

/// The five FIB-addressed structures owned by this facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableKind {
    /// `SttbfBkmkFactoid` / `fcSttbfBkmkFactoid`.
    BookmarkInfo,
    /// `PlcfBkfFactoid` / `fcPlcfBkfFactoid`.
    BookmarkStarts,
    /// `PlcfBklFactoid` / `fcPlcfBklFactoid`.
    BookmarkEnds,
    /// `FactoidData` / `fcFactoidData`.
    PropertyBags,
    /// `Plcffactoid` / `fcPlcffactoid`.
    Recognizer,
}

impl TableKind {
    /// The corresponding `FibRgFcLcb` array index.
    #[must_use]
    pub const fn fib_index(self) -> usize {
        match self {
            Self::BookmarkInfo => STTBF_BKMK_FACTOID,
            Self::BookmarkStarts => PLCF_BKF_FACTOID,
            Self::BookmarkEnds => PLCF_BKL_FACTOID,
            Self::PropertyBags => FACTOID_DATA,
            Self::Recognizer => PLCF_FACTOID,
        }
    }

    pub(super) const ALL: [Self; 5] = [
        Self::BookmarkInfo,
        Self::BookmarkStarts,
        Self::BookmarkEnds,
        Self::PropertyBags,
        Self::Recognizer,
    ];
}

/// The immutable FIB/table layout captured by a [`super::Snapshot`].
///
/// Smart-tag edits never relocate a table and never rewrite a FIB pointer.
/// A candidate whose encoded property-bag payload does not fit its original
/// range is rejected instead of shifting unrelated table bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Topology {
    table_stream_len: usize,
    pointers: Arc<[Option<TableRange>]>,
    document_end: u32,
    smart_tag_ranges: [Option<TableRange>; 5],
}

impl Topology {
    /// Length of the selected table stream at snapshot time.
    #[must_use]
    pub const fn table_stream_len(&self) -> usize {
        self.table_stream_len
    }

    /// Number of FIB pointer pairs retained in this topology.
    #[must_use]
    pub fn pointer_count(&self) -> usize {
        self.pointers.len()
    }

    /// Returns a copy of one FIB-addressed range, including unrelated tables.
    #[must_use]
    pub fn pointer(&self, index: usize) -> Option<TableRange> {
        self.pointers.get(index).copied().flatten()
    }

    /// Returns the range owned by one smart-tag structure.
    #[must_use]
    pub fn range(&self, kind: TableKind) -> Option<TableRange> {
        self.smart_tag_ranges[smart_index(kind)]
    }

    /// Character-position limit captured from `FibRgLw97`.
    #[must_use]
    pub const fn document_end(&self) -> u32 {
        self.document_end
    }

    pub(super) fn capture(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        if table_stream.len() > usize::try_from(u32::MAX).unwrap_or(usize::MAX) {
            return Err(corrupted(
                "smart-tag table stream exceeds the FIB address space",
            ));
        }
        let count = fib
            .table_pointer_count()
            .ok_or_else(|| corrupted("WordDocument FIB table-pointer array is truncated"))?;
        let mut pointers = Vec::with_capacity(count);
        for index in 0..count {
            let pointer = fib.get_table_pointer(index).and_then(|(offset, length)| {
                (length != 0).then_some(TableRange { offset, length })
            });
            pointers.push(pointer);
        }

        let smart_tag_ranges =
            TableKind::ALL.map(|kind| pointers.get(kind.fib_index()).copied().flatten());
        for (index, range) in smart_tag_ranges.iter().enumerate() {
            let Some(range) = range else {
                continue;
            };
            range.as_usize(table_stream.len())?;
            for (other_index, other) in pointers.iter().enumerate() {
                if TableKind::ALL[index].fib_index() == other_index {
                    continue;
                }
                let Some(other) = other else {
                    continue;
                };
                if ranges_overlap(*range, *other) {
                    return Err(corrupted(
                        "smart-tag FIB ranges overlap an unrelated table range",
                    ));
                }
            }
        }

        let document_end = fib
            .get_document_parts_end()
            .ok_or_else(|| corrupted("document-part character counts overflow"))?;
        Ok(Self {
            table_stream_len: table_stream.len(),
            pointers: pointers.into(),
            document_end,
            smart_tag_ranges,
        })
    }

    pub(super) fn range_bytes<'a>(
        &self,
        table_stream: &'a [u8],
        kind: TableKind,
    ) -> Result<Option<&'a [u8]>> {
        self.range(kind)
            .map(|range| {
                range
                    .as_usize(table_stream.len())
                    .map(|range| &table_stream[range])
            })
            .transpose()
    }
}

fn smart_index(kind: TableKind) -> usize {
    match kind {
        TableKind::BookmarkInfo => 0,
        TableKind::BookmarkStarts => 1,
        TableKind::BookmarkEnds => 2,
        TableKind::PropertyBags => 3,
        TableKind::Recognizer => 4,
    }
}

fn ranges_overlap(left: TableRange, right: TableRange) -> bool {
    let Some(left_end) = left.end() else {
        return true;
    };
    let Some(right_end) = right.end() else {
        return true;
    };
    left.offset < right_end && right.offset < left_end
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
