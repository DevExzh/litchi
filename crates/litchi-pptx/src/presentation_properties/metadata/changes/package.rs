//! Changes Information OPC ownership and source-checked publication.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::codec;
use super::codec::Limits;
use super::model::Part as ChangesPart;
use super::transaction::{Change, Prepared, Snapshot};
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

/// Read the unique Changes Information part after validating its OPC graph.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, limits: Limits) -> Result<Option<ChangesPart>> {
    codec::load(package, limits)
}

/// Add a new Changes Information owner atomically.
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
pub(crate) fn store(package: &mut OpcPackage, value: &ChangesPart, limits: Limits) -> Result<()> {
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    codec::store(&mut candidate, value, limits)?;
    if codec::load(&candidate, limits)?.is_none() {
        return Err(invalid("stored Changes Information part cannot be read"));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(())
}

/// Capture the typed Changes Information owner and its exact source context.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_snapshot(package: &OpcPackage, limits: Limits) -> Result<Option<Snapshot>> {
    let Some(value) = codec::load(package, limits)? else {
        return Ok(None);
    };
    let presentation = package.main_document_part()?;
    let relationship = presentation
        .rels()
        .get(&value.relationship_id)
        .ok_or_else(|| invalid("Changes Information relationship disappeared during load"))?;
    let target = relationship.target_ref().to_owned();
    let target_name = PackURI::new(&value.part_name).map_err(Error::Uri)?;
    let part = package.get_part(&target_name)?;
    Snapshot::from_wire(
        presentation.partname().to_string(),
        presentation.content_type().to_owned(),
        value.relationship_id,
        target,
        value.part_name,
        part.blob_arc(),
        value.changes_information,
        limits,
    )
    .map(Some)
}

/// Apply an already committed Changes Information patch atomically.
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
    patch.apply(package)
}

/// Apply a committed Changes Information edit atomically.
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

/// Remove the Changes Information owner and its Presentation relationship.
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
pub(crate) fn remove(package: &mut OpcPackage, limits: Limits) -> Result<Option<ChangesPart>> {
    let mut candidate = package.clone();
    let Some(value) = codec::load(&candidate, limits)? else {
        return Ok(None);
    };
    refuse_signed_change(package)?;
    let presentation_name = candidate.main_document_part()?.partname().clone();
    let target = PackURI::new(&value.part_name).map_err(Error::Uri)?;
    let actual_target = candidate
        .iter_parts()
        .find(|part| {
            part.partname()
                .as_str()
                .eq_ignore_ascii_case(target.as_str())
        })
        .map(|part| part.partname().clone())
        .ok_or_else(|| invalid("Changes Information target disappeared during removal"))?;
    let presentation = candidate.get_part_mut(&presentation_name)?;
    if presentation
        .rels_mut()
        .remove(&value.relationship_id)
        .is_none()
    {
        return Err(invalid(
            "Changes Information relationship disappeared during removal",
        ));
    }
    if !candidate.remove_part(&actual_target) {
        return Err(invalid(
            "Changes Information part disappeared during removal",
        ));
    }
    if codec::load(&candidate, limits)?.is_some() {
        return Err(invalid("Changes Information removal left an owner behind"));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(Some(value))
}

/// Replace only the already-owned Changes Information blob.
pub(crate) fn replace_snapshot(
    package: &mut OpcPackage,
    after: &Snapshot,
    limits: Limits,
) -> Result<Snapshot> {
    let current = load_snapshot(package, limits)?
        .ok_or_else(|| invalid("cannot publish Changes Information into an absent owner"))?;
    if current.part_name() != after.part_name()
        || current.relationship_id() != after.relationship_id()
        || current.presentation_part_name() != after.presentation_part_name()
    {
        return Err(invalid(
            "Changes Information patch changes package topology",
        ));
    }
    let target = PackURI::new(after.part_name()).map_err(Error::Uri)?;
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&target)?
        .set_blob_shared(Arc::clone(after.source_arc()));
    let resulting = load_snapshot(&candidate, limits)?
        .ok_or_else(|| invalid("published Changes Information part cannot be read"))?;
    if !resulting.same_source(after) {
        return Err(invalid(
            "published Changes Information source differs from the commit",
        ));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(resulting)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
