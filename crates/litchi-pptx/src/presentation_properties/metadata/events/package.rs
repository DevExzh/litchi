//! Slide-show event OPC ownership and atomic publication.

use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

use super::codec::{self, Limits, LoadLimits};
use super::model::{Draft, Event};
use super::transaction::{Change, Prepared, Snapshot};
use super::validation::validate_events;
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

/// Discover the typed event records on one slide without executing them.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(
    package: &OpcPackage,
    slide_part_name: &PackURI,
    limits: Limits,
) -> Result<Option<Vec<Event>>> {
    load_snapshot(package, slide_part_name, limits)
        .map(|value| value.map(|snapshot| snapshot.events().to_vec()))
}

/// Capture one slide's event list and exact source context.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_snapshot(
    package: &OpcPackage,
    slide_part_name: &PackURI,
    limits: Limits,
) -> Result<Option<Snapshot>> {
    let slide = package.get_part(slide_part_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_owned(),
            actual: slide.content_type().to_owned(),
        });
    }
    let Some(located) = codec::locate(0, slide.blob(), limits, &mut LoadLimits::default())? else {
        return Ok(None);
    };
    Snapshot::from_located(
        slide_part_name.to_string(),
        slide.blob_arc(),
        located,
        limits,
    )
    .map(Some)
}

/// Add a new slide-show event extension atomically.
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
pub(crate) fn store(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    events: &[Draft],
    limits: Limits,
) -> Result<()> {
    validate_events(events, limits)?;
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    codec::store(&mut candidate, slide_part_name, events, limits)?;
    if load_snapshot(&candidate, slide_part_name, limits)?.is_none() {
        return Err(invalid("stored slide-show events cannot be read back"));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(())
}

/// Remove only the MS-PPTX show-event extension, retaining neighboring opaque
/// extension blocks and all other slide XML.
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
pub(crate) fn remove(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    limits: Limits,
) -> Result<Option<Snapshot>> {
    let Some(before) = load_snapshot(package, slide_part_name, limits)? else {
        return Ok(None);
    };
    let located = codec::locate(0, before.source_xml(), limits, &mut LoadLimits::default())?
        .ok_or_else(|| invalid("slide-show event source disappeared during removal"))?;
    let updated = codec::remove_extension(before.source_xml(), &located, limits)?;
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    candidate.get_part_mut(slide_part_name)?.set_blob(updated);
    if load_snapshot(&candidate, slide_part_name, limits)?.is_some() {
        return Err(invalid("slide-show event removal left an event extension"));
    }
    settle_unsigned_change(&mut candidate)?;
    *package = candidate;
    Ok(Some(before))
}

/// Apply a committed source-checked event patch atomically.
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
    let slide_name = PackURI::new(patch.before().slide_part_name()).map_err(Error::Uri)?;
    let current = load_snapshot(package, &slide_name, patch.before().limits)?
        .ok_or_else(|| invalid("slide-show event source is absent"))?;
    if !current.same_source(patch.before()) {
        return Err(invalid("slide-show event source is stale"));
    }
    if patch.is_noop() {
        return Ok(current);
    }
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&slide_name)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate, &slide_name, patch.after().limits)?
        .ok_or_else(|| invalid("published slide-show event list cannot be read"))?;
    if !result.same_source(patch.after()) {
        return Err(invalid(
            "published slide-show event source differs from the patch",
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

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
