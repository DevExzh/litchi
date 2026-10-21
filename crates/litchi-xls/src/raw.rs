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

/// The legacy generic XLS reader.
///
/// This low-level reader owns an eager BIFF projection over a caller-provided
/// `Read + Seek` source. New package ownership APIs live beside this type;
/// callers that need the legacy eager projection should use this explicit
/// `raw` namespace.
pub use crate::workbook::model::Reader;

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
        SourceBackedOptions::default().limits(limits),
    )
}

/// Admit explicitly transformed source bytes under the original package's
/// retained edit limits. The caller is responsible for authorizing and
/// validating the transformation (for example, signature removal).
///
/// # Errors
/// Returns the typed package or cell-values error if admission fails.
pub(crate) fn cell_values_snapshot_from_bytes(
    package: &crate::Package,
    bytes: Vec<u8>,
) -> crate::package::Result<crate::cell_values::Snapshot> {
    let (limits, cfb_limits) = package.edit_profile(
        litchi_ole_common::object::Limits::default(),
        litchi_cfb::OleFileLimits::default(),
    )?;
    crate::cell_values::Snapshot::from_bytes_with_cfb(bytes, limits, cfb_limits)
        .map_err(crate::package::Error::Reader)
}

/// Reopen explicitly transformed bytes as a primary package under the source's
/// retained read and edit policies. The caller must authorize and validate the
/// transformation, such as removing a compound signature after explicit consent.
///
/// This handoff preserves eager versus positional read policy and checks source
/// freshness around admission. It does not authorize decryption or re-encryption.
///
/// # Errors
/// Returns typed encryption-policy, source, admission or projection failures.
pub fn package_from_replacement_bytes(
    source: &crate::Package,
    bytes: Vec<u8>,
) -> crate::package::edit::Result<crate::Package> {
    crate::package::edit::package_from_replacement_bytes(source, bytes)
}
