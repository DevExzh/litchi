//! Input validation for `PivotTable` cache writers.

use super::model::PivotCacheFieldInfo;
use crate::{Error, Result};

/// Validate the packed shared-index shape of one SXDBB row and return its
/// encoded payload length.
pub(super) fn validate_sxdbb_inputs(
    fields: &[PivotCacheFieldInfo<'_>],
    indices: &[u16],
) -> Result<usize> {
    let expected = fields
        .iter()
        .filter(|field| field.is_source_field && !field.items.is_empty())
        .count();
    if indices.len() != expected {
        return Err(Error::InvalidData(
            "PivotCache row shared-index cardinality mismatch".to_string(),
        ));
    }

    fields
        .iter()
        .filter(|field| field.is_source_field && !field.items.is_empty())
        .try_fold(0usize, |size, field| {
            size.checked_add(if field.items.len() >= 0x100 { 2 } else { 1 })
        })
        .ok_or_else(|| Error::InvalidData("SXINDEXLIST size overflow".to_string()))
}

/// Validate one packed SXDBB item index without changing its wire encoding.
pub(super) fn validate_sxdbb_index(index: u16, use_16bit: bool) -> Result<()> {
    if use_16bit {
        Ok(())
    } else {
        u8::try_from(index)
            .map(|_| ())
            .map_err(|_error| Error::InvalidData("SXINDEXLIST 8-bit index overflow".to_string()))
    }
}
