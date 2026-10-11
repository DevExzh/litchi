//! Low-level handoffs for already-indexed XLS container sources.
//!
//! The ordinary [`crate::SourceBackedWorkbook`] API intentionally exposes no
//! CFB container types. Facades that already own a validated positional CFB
//! catalog may use this module to transfer that catalog without rebuilding it.

use crate::{
    Result, SourceBackedError, SourceBackedLimits, SourceBackedOptions, SourceBackedWorkbook,
    SourceLimits,
};
use litchi_cfb::{OleFile, SharedOleFile};
use std::io::{Read, Seek};
use std::sync::Arc;

/// Validate an already-parsed CFB container against an explicit XLS physical
/// source policy before any XML, Workbook, or other stream payload is read.
///
/// This helper is intended for container-owning facades that already parsed
/// the CFB catalog for format detection. Adoption cannot undo allocations
/// made by that parse, but it preserves the selected policy ordering for the
/// XLS payload owner.
pub fn preflight_ole_file<R: Read + Seek>(
    ole_file: &OleFile<R>,
    source_limits: SourceLimits,
) -> Result<()> {
    crate::workbook::package::preflight_ole_file(ole_file, source_limits)
}

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
    SourceBackedWorkbook::from_shared_ole_file_with_options(
        cfb,
        SourceBackedOptions::default().with_limits(limits),
    )
}
