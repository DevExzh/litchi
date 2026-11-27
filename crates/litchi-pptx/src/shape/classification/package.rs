//! Package-context discovery and atomic publication for shape classification.

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

use super::codec;
use super::model::{Outcome, Snapshot};
use super::transaction;
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

pub(crate) fn load_with_limits<'k>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
    shape_limits: crate::shape::Limits,
) -> Result<Option<Snapshot>> {
    let part = package.get_part(owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    let key = key.into();
    let span = crate::tag::shape::selected_raw_span_with_limits(part.blob(), key, shape_limits)?;
    let source = codec::read_in_owner(part.blob(), span)?;
    Ok(source.snapshot)
}

pub(crate) fn put_with_limits<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
    outcome: Outcome,
    shape_limits: crate::shape::Limits,
) -> Result<Option<Snapshot>> {
    let key = key.into();
    let (owner_name, owner_blob, span, source, shape_blob) = {
        let part = package.get_part(owner)?;
        crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
        let span =
            crate::tag::shape::selected_raw_span_with_limits(part.blob(), key, shape_limits)?;
        let shape_blob = part
            .blob()
            .get(span.clone())
            .ok_or_else(|| Error::Invalid("classification shape span is invalid".into()))?
            .to_vec();
        (
            part.partname().clone(),
            part.blob().to_vec(),
            span.clone(),
            codec::read_in_owner(part.blob(), span)?,
            shape_blob,
        )
    };
    let previous = source.snapshot.clone();
    let max_output_bytes = shape_limits.output_bytes();
    let staged_shape = transaction::set(&shape_blob, &source, outcome, max_output_bytes)?;
    let staged_owner = codec::replace(&owner_blob, span.clone(), &staged_shape, max_output_bytes)?;
    validate_staged(&staged_owner, key, outcome, shape_limits)?;
    // An exact no-op returns before the package is borrowed mutably, so it
    // keeps the exact source, and with it a signature, byte for byte.
    if staged_owner == owner_blob {
        return Ok(previous);
    }
    refuse_signed_change(package)?;

    let part = package.get_part_mut(&owner_name)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != owner_blob {
        return Err(Error::PatchConflict);
    }
    part.set_blob(staged_owner);
    settle_unsigned_change(package)?;
    Ok(previous)
}

pub(crate) fn remove_with_limits<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
    shape_limits: crate::shape::Limits,
) -> Result<Option<Snapshot>> {
    let key = key.into();
    let (owner_name, owner_blob, span, source, shape_blob) = {
        let part = package.get_part(owner)?;
        crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
        let span =
            crate::tag::shape::selected_raw_span_with_limits(part.blob(), key, shape_limits)?;
        let shape_blob = part
            .blob()
            .get(span.clone())
            .ok_or_else(|| Error::Invalid("classification shape span is invalid".into()))?
            .to_vec();
        (
            part.partname().clone(),
            part.blob().to_vec(),
            span.clone(),
            codec::read_in_owner(part.blob(), span)?,
            shape_blob,
        )
    };
    let Some(previous) = source.snapshot.clone() else {
        return Ok(None);
    };
    let max_output_bytes = shape_limits.output_bytes();
    let Some(staged_shape) = transaction::remove(&shape_blob, &source, max_output_bytes)? else {
        return Ok(None);
    };
    let staged_owner = codec::replace(&owner_blob, span, &staged_shape, max_output_bytes)?;
    validate_removed(&staged_owner, key, shape_limits)?;
    // An exact no-op returns before the package is borrowed mutably, so it
    // keeps the exact source, and with it a signature, byte for byte.
    if staged_owner == owner_blob {
        return Ok(Some(previous));
    }
    refuse_signed_change(package)?;

    let part = package.get_part_mut(&owner_name)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != owner_blob {
        return Err(Error::PatchConflict);
    }
    part.set_blob(staged_owner);
    settle_unsigned_change(package)?;
    Ok(Some(previous))
}

fn validate_staged(
    owner: &[u8],
    key: crate::shape::Key<'_>,
    expected: Outcome,
    shape_limits: crate::shape::Limits,
) -> Result<()> {
    let span = crate::tag::shape::selected_raw_span_with_limits(owner, key, shape_limits)?;
    let source = codec::read_in_owner(owner, span)?;
    if source.snapshot.as_ref().and_then(Snapshot::outcome) != Some(expected) {
        return Err(Error::Invalid(
            "staged classification did not round-trip its typed outcome".into(),
        ));
    }
    Ok(())
}

fn validate_removed(
    owner: &[u8],
    key: crate::shape::Key<'_>,
    shape_limits: crate::shape::Limits,
) -> Result<()> {
    let span = crate::tag::shape::selected_raw_span_with_limits(owner, key, shape_limits)?;
    let source = codec::read_in_owner(owner, span)?;
    if source.snapshot.is_some() {
        return Err(Error::Invalid(
            "staged classification still contains a typed outcome".into(),
        ));
    }
    Ok(())
}
