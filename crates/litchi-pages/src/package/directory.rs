//! App-authored Pages package directories (IWORK_API_REQUESTS.md request 16).
//!
//! The capture, conversion and bundle publication are shared with the
//! Keynote and Numbers owners in [`litchi_iwa_archive::bundle`] (record
//! 0306); this module maps the Pages limits and errors onto it.

use std::path::Path;

use litchi_iwa_archive::bundle::{CaptureError, CaptureLimitKind, CaptureLimits};

use super::{PackageError, PackageResult};
use crate::{LimitKind, SourceLimitKind, SourceLimits};

/// Capture the bundle at `root` and return its single-file package bytes.
///
/// # Errors
///
/// [`PackageError::InvalidFormat`] for a bundle the shared capture refuses
/// (no `Index.zip` nor `Index/` folder, or both, a symbolic link, a special
/// file, a non-UTF-8 name or nesting deeper than 16 levels);
/// [`PackageError::Limit`] when a figure exceeds `limits`;
/// [`PackageError::Io`] for a read failure; archive errors when `Index.zip`
/// is malformed or the reassembly fails.
pub(super) fn single_file_bytes(root: &Path, limits: SourceLimits) -> PackageResult<Vec<u8>> {
    let capture = CaptureLimits::new(
        limits.max_members(),
        limits.max_entry_bytes(),
        limits.max_input_bytes().min(limits.max_aggregate_bytes()),
    );
    litchi_iwa_archive::bundle::capture(root, capture, limits.archive()?).map_err(capture_error)
}

/// The Pages error of a shared capture failure.
fn capture_error(error: CaptureError) -> PackageError {
    match error {
        CaptureError::Invalid(message) => PackageError::InvalidFormat(message.to_owned()),
        CaptureError::Limit {
            kind,
            observed,
            maximum,
        } => PackageError::Limit {
            kind: LimitKind::Source(match kind {
                CaptureLimitKind::Members => SourceLimitKind::Members,
                CaptureLimitKind::EntryBytes => SourceLimitKind::EntryBytes,
                _ => SourceLimitKind::InputBytes,
            }),
            observed,
            maximum,
        },
        CaptureError::Allocation { amount } => PackageError::Allocation { amount },
        CaptureError::Io(error) => PackageError::Io(error),
        CaptureError::Archive(error) => PackageError::Archive(error),
        _ => PackageError::InvalidFormat("the package directory cannot be captured".to_owned()),
    }
}

/// The archive profile that admits a converted bundle of `length` bytes.
pub(super) fn converted_limits(
    limits: litchi_iwa_archive::Limits,
    length: usize,
) -> PackageResult<litchi_iwa_archive::Limits> {
    Ok(litchi_iwa_archive::bundle::converted_limits(
        limits, length,
    )?)
}

/// Save `catalog` as a package directory at `path` through the archive
/// layer's directory publication boundary (a sibling staging directory,
/// synchronized, then swapped in).
pub(super) fn save_bundle(
    catalog: &litchi_iwa_archive::SourceCatalog,
    path: &Path,
) -> Result<(), super::SaveError> {
    litchi_iwa_archive::bundle::save(
        catalog.package(),
        catalog.limits(),
        path,
        super::SaveError::Bundle,
        super::SaveError::Bundling,
    )
}
