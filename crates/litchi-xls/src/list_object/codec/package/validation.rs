//! Validation for future-record insertion points in the List12 sequence.

use super::super::super::model::ListObject;
use super::super::super::{AUTO_FILTER12_RECORD_TYPE, invalid};
use crate::Result;

pub(super) fn validate_future_insertion(table: &ListObject) -> Result<()> {
    if table
        .opaque_future_records
        .iter()
        .any(|future| future.after_list12_count == 0 || future.after_list12_count > 3)
    {
        return Err(invalid(
            AUTO_FILTER12_RECORD_TYPE,
            "opaque table future-record insertion point is invalid",
        ));
    }
    Ok(())
}
