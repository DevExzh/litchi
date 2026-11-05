//! Presentation-structure OPC facade.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::super::Limits;
pub(crate) use super::codec::{
    add_custom_show, add_custom_show_slide, add_section, add_section_slide, remove_custom_show,
    remove_custom_show_slide, remove_section, remove_section_slide, reorder_custom_show_slides,
    reorder_custom_shows, reorder_section_slides, reorder_sections, store, update_custom_show,
    update_section,
};
pub use super::codec::{find_custom_show, find_section, load};

use super::transaction::{Change, Prepared, Snapshot};
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

/// Capture the validated presentation graph and its exact source context.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_snapshot(package: &OpcPackage, limits: Limits) -> Result<Snapshot> {
    let graph = load(package, limits)?;
    let presentation = package.main_document_part()?;
    Snapshot::from_wire(
        presentation.partname().to_string(),
        presentation.content_type().to_owned(),
        presentation.blob_arc(),
        Snapshot::capture_relationships(presentation),
        graph,
        limits,
    )
}

/// Apply a committed presentation-structure patch atomically.
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
        return Err(invalid("presentation structure source is stale"));
    }
    if patch.is_noop() {
        return Ok(current);
    }

    let presentation_name = package.main_document_part()?.partname().clone();
    if presentation_name.as_str() != patch.after().presentation_part_name() {
        return Err(invalid(
            "presentation structure patch changes the owning part",
        ));
    }
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&presentation_name)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate, limits)?;
    if !result.same_source(patch.after()) || result.graph() != patch.after().graph() {
        return Err(invalid(
            "published presentation structure differs from the commit",
        ));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(result)
}

/// Apply a committed presentation-structure edit atomically.
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

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
