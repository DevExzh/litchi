//! Low-level handoffs for already-indexed XLS container sources.
//!
//! The ordinary [`crate::SourceBackedWorkbook`] API intentionally exposes no
//! CFB container types. Facades that already own a validated positional CFB
//! catalog may use this module to transfer that catalog without rebuilding it.

use crate::{SourceBackedError, SourceBackedLimits, SourceBackedWorkbook};
use litchi_cfb::SharedOleFile;
use std::sync::Arc;

/// Constructs a source-backed XLS owner from an already-indexed CFB catalog.
///
/// The handoff derives the retained positional source directly from `cfb`,
/// so callers cannot pair the catalog with unrelated bytes. It captures and
/// validates the catalog's source identity before parsing workbook globals,
/// and every subsequent source-backed operation remains fenced against source
/// replacement or in-place mutation.
///
/// This is a low-level integration boundary for container-owning facades.
/// Ordinary XLS callers should use [`SourceBackedWorkbook::from_read_at`]
/// instead.
pub fn source_backed_workbook_from_shared_ole_file(
    cfb: Arc<SharedOleFile>,
    limits: SourceBackedLimits,
) -> Result<SourceBackedWorkbook, SourceBackedError> {
    SourceBackedWorkbook::from_shared_ole_file_with_limits(cfb, limits)
}
