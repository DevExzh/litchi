//! Presentation-part ownership and atomic guide publication.

use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

use super::Limits;
use super::model::Guides;
use super::transaction::{Change, Prepared, Snapshot};
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

/// Read the typed extended-guide value from the package's main presentation.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, limits: Limits) -> Result<Guides> {
    Ok(load_snapshot(package, limits)?.guides().clone())
}

/// Capture the typed guide value and exact owning presentation source.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_snapshot(package: &OpcPackage, limits: Limits) -> Result<Snapshot> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation.content_type())?;
    Snapshot::from_wire(
        presentation.partname().to_string(),
        presentation.content_type().to_owned(),
        presentation.blob_arc(),
        limits,
    )
}

/// Replace the typed guide value while preserving unrelated presentation XML.
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
pub(crate) fn store(package: &mut OpcPackage, value: &Guides) -> Result<()> {
    let source = load_snapshot(package, Limits::from_source_limits(package.read_limits()))?;
    let mut edit = source.edit();
    edit.replace(value.clone())?;
    apply_commit(package, edit.commit()?)?;
    Ok(())
}

/// Remove both guide lists and return their previous typed value.
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
pub(crate) fn remove(package: &mut OpcPackage) -> Result<Guides> {
    let source = load_snapshot(package, Limits::from_source_limits(package.read_limits()))?;
    let previous = source.guides().clone();
    let mut edit = source.edit();
    edit.replace(Guides::default())?;
    apply_commit(package, edit.commit()?)?;
    Ok(previous)
}

/// Apply a committed guide patch atomically after a complete source check.
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
        return Err(invalid("extended-guide source is stale"));
    }
    if patch.is_noop() {
        return Ok(current);
    }
    if patch.before().presentation_part_name() != patch.after().presentation_part_name()
        || patch.before().presentation_content_type() != patch.after().presentation_content_type()
    {
        return Err(invalid(
            "extended-guide patch changes presentation ownership",
        ));
    }
    let presentation_name =
        PackURI::new(patch.before().presentation_part_name()).map_err(Error::Uri)?;
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&presentation_name)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate, limits)?;
    if !result.same_source(patch.after()) {
        return Err(invalid(
            "published extended-guide source differs from the commit",
        ));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(result)
}

/// Apply a committed edit and return its validated post-publication snapshot.
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
pub(crate) fn apply_commit(package: &mut OpcPackage, commit: Prepared) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

fn require_presentation(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: ct::PML_PRESENTATION_MAIN.to_owned(),
            actual: content_type.to_owned(),
        })
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
