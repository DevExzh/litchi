//! Semantic stream-planning values shared by validation and BIFF encoding.

use std::collections::HashMap;

use crate::{Error, Result};

use super::super::worksheet::WritableWorksheet;

/// Result of generating the workbook: the Workbook stream plus any pivot
/// cache storage streams that must be placed in `_SX_DB_CUR/nnnn`.
pub(crate) struct WorkbookStreams {
    /// The main Workbook BIFF stream.
    pub workbook: Vec<u8>,
    /// Optional Office Toolbars (`XCB`) stream.
    pub toolbar: Option<Vec<u8>>,
    /// Pivot cache streams: `(stream_id, data)`.  Each goes into
    /// `_SX_DB_CUR/{stream_id:04X}`.
    pub pivot_caches: Vec<(u16, Vec<u8>)>,
}

#[derive(Clone, Copy)]
pub(super) struct PivotCacheIdentity {
    /// Zero-based index used by SXVIEW.iCache.
    pub(super) cache_index: u16,
    /// One-based identifier used by `SXStreamID` and `_SX_DB_CUR/nnnn`.
    pub(super) stream_id: u16,
}

pub(super) fn lookup_shared_string_index(
    shared_strings: &[String],
    string_map: &HashMap<String, u32>,
    value: &str,
) -> Result<u32> {
    let index = string_map.get(value).copied().ok_or_else(|| {
        Error::InvalidData(format!(
            "string cell value {value:?} is missing from the shared string table"
        ))
    })?;
    let table_index = usize::try_from(index).map_err(|_error| {
        Error::InvalidData(format!(
            "shared string index {index} for value {value:?} cannot be represented"
        ))
    })?;
    match shared_strings.get(table_index) {
        Some(entry) if entry == value => Ok(index),
        Some(_) => Err(Error::InvalidData(format!(
            "shared string index {index} for value {value:?} does not match the shared string table"
        ))),
        None => Err(Error::InvalidData(format!(
            "shared string index {index} for value {value:?} is outside the shared string table"
        ))),
    }
}

pub(super) fn stage_pivot_cache_identities(
    worksheets: &[WritableWorksheet],
) -> Result<Vec<Vec<PivotCacheIdentity>>> {
    let pivot_count = worksheets.iter().try_fold(0usize, |count, worksheet| {
        count
            .checked_add(worksheet.pivot_tables.len())
            .ok_or_else(|| Error::InvalidData("PivotTable cache count overflow".to_string()))
    })?;
    if pivot_count > usize::from(u16::MAX) {
        return Err(Error::InvalidData(format!(
            "PivotTable cache count {pivot_count} exceeds the BIFF8 limit of {}",
            u16::MAX
        )));
    }

    let mut next_index = 0usize;
    worksheets
        .iter()
        .map(|worksheet| {
            (0..worksheet.pivot_tables.len())
                .map(|_| {
                    let cache_index = u16::try_from(next_index).map_err(|_error| {
                        Error::InvalidData("PivotTable cache index overflow".to_string())
                    })?;
                    let stream_id = cache_index.checked_add(1).ok_or_else(|| {
                        Error::InvalidData("PivotTable cache stream ID overflow".to_string())
                    })?;
                    next_index = next_index.checked_add(1).ok_or_else(|| {
                        Error::InvalidData("PivotTable cache index overflow".to_string())
                    })?;
                    Ok(PivotCacheIdentity {
                        cache_index,
                        stream_id,
                    })
                })
                .collect()
        })
        .collect()
}
