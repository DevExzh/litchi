//! OPC ownership and atomic publication for presentation custom shows.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::Limits;
use super::model::List;
use super::transaction::{Change, Prepared, Snapshot};
use super::wire;
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

/// Load the validated custom-show list, returning an empty list when absent.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, limits: Limits) -> Result<List> {
    Ok(load_snapshot(package, limits)?.list().clone())
}

/// Atomically store a typed custom-show list.
///
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub(crate) fn store(package: &mut OpcPackage, value: &List) -> Result<()> {
    let snapshot = load_snapshot(package, Limits::from_source_limits(package.read_limits()))?;
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
pub fn load_snapshot(package: &OpcPackage, limits: Limits) -> Result<Snapshot> {
    let presentation = package.main_document_part()?;
    let located = wire::locate(package, limits)?;
    Snapshot::from_located(
        presentation.partname().to_string(),
        presentation.content_type().to_owned(),
        presentation.blob_arc(),
        located,
        limits,
    )
}

/// Apply a previously committed patch after a complete source check.
///
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn apply_patch(package: &mut OpcPackage, patch: &Change) -> Result<Snapshot> {
    let limits = Limits::from_source_limits(package.read_limits());
    let current = load_snapshot(package, limits)?;
    if !current.same_source(patch.before()) {
        return Err(invalid("custom-show PresentationML source is stale"));
    }
    if patch.is_noop() {
        return Ok(current);
    }
    if current.presentation_part_name() != patch.after().presentation_part_name()
        || current.presentation_content_type() != patch.after().presentation_content_type()
        || current.relationships != patch.after().relationships
    {
        return Err(invalid("custom-show patch changes presentation topology"));
    }

    let target = PackURI::new(patch.after().presentation_part_name()).map_err(Error::Uri)?;
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&target)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate, limits)?;
    if !result.same_source(patch.after()) {
        return Err(invalid(
            "published custom-show source differs from the committed snapshot",
        ));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(result)
}

/// Apply a committed edit atomically.
///
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
#[allow(
    clippy::needless_pass_by_value,
    reason = "public API consumes the commit to signal it has been applied"
)]
pub(crate) fn apply_commit(package: &mut OpcPackage, commit: Prepared) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

/// Remove all typed custom shows while retaining opaque custom-show XML.
///
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn remove(package: &mut OpcPackage) -> Result<Option<List>> {
    let before = load_snapshot(package, Limits::from_source_limits(package.read_limits()))?;
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
