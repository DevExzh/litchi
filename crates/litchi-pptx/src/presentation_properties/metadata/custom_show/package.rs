//! OPC ownership and atomic publication for presentation custom shows.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::model::List;
use super::transaction::{Commit, Patch, Snapshot};
use super::wire;
use crate::{Error, Result};

/// Load the validated custom-show list, returning an empty list when absent.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage) -> Result<List> {
    Ok(load_snapshot(package)?.list().clone())
}

/// Atomically store a typed custom-show list.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn store(package: &mut OpcPackage, value: &List) -> Result<()> {
    let snapshot = load_snapshot(package)?;
    let mut transaction = snapshot.edit();
    transaction.replace(value.clone())?;
    let commit = transaction.commit()?;
    apply_commit(package, commit)?;
    Ok(())
}

/// Capture the owning presentation XML and all relationship context.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_snapshot(package: &OpcPackage) -> Result<Snapshot> {
    let presentation = package.main_document_part()?;
    let located = wire::locate(package)?;
    Snapshot::from_located(
        presentation.partname().to_string(),
        presentation.content_type().to_owned(),
        presentation.blob_arc(),
        located,
    )
}

/// Apply a previously committed patch after a complete source check.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let current = load_snapshot(package)?;
    if !current.same_source(patch.before()) {
        return Err(invalid("custom-show PresentationML source is stale"));
    }
    if patch.is_empty() {
        return Ok(current);
    }
    if current.presentation_part_name() != patch.after().presentation_part_name()
        || current.presentation_content_type() != patch.after().presentation_content_type()
        || current.relationships != patch.after().relationships
    {
        return Err(invalid("custom-show patch changes presentation topology"));
    }

    let target = PackURI::new(patch.after().presentation_part_name()).map_err(Error::Uri)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&target)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate)?;
    if !result.same_source(patch.after()) {
        return Err(invalid(
            "published custom-show source differs from the committed snapshot",
        ));
    }
    candidate.unsign();
    *package = candidate;
    Ok(result)
}

/// Apply a committed edit atomically.
///
/// # Errors
///
/// Returns an error if the operation fails.
#[allow(
    clippy::needless_pass_by_value,
    reason = "public API consumes the commit to signal it has been applied"
)]
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

/// Remove all typed custom shows while retaining opaque custom-show XML.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn remove(package: &mut OpcPackage) -> Result<Option<List>> {
    let before = load_snapshot(package)?;
    if before.list().is_empty() {
        return Ok(None);
    }
    let previous = before.list().clone();
    let mut transaction = before.edit();
    transaction.replace(List::new())?;
    let commit = transaction.commit()?;
    apply_commit(package, commit)?;
    Ok(Some(previous))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
