//! Small transaction-state guards shared by the semantic facade.

use crate::error::{RemoveBlock, Result};

use super::Edit;

/// Reject edits that cannot be combined with a pending worksheet removal.
pub(super) fn no_removal(edit: &Edit, part: &str) -> Result<()> {
    if edit.removed.is_empty() {
        Ok(())
    } else {
        Err(edit.remove_block(RemoveBlock::MixedEdit, part))
    }
}
