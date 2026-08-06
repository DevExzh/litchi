//! Package-context discovery and atomic publication for shape classification.

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

use super::codec;
use super::model::{Outcome, Snapshot};
use super::transaction;
use crate::{Error, Result};

/// Read a classification snapshot from one selected shape in a slide part.
pub(crate) fn load<'k>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
) -> Result<Option<Snapshot>> {
    let part = package.get_part(owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    let key = key.into();
    let span = crate::tag::shape::selected_raw_span(part.blob(), key)?;
    let source = codec::read(
        part.blob()
            .get(span)
            .ok_or_else(|| Error::Invalid("classification shape span is invalid".into()))?,
    )?;
    Ok(source.snapshot)
}

/// Replace or create the typed outcome in one shape transaction.
pub(crate) fn put<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
    outcome: Outcome,
) -> Result<Option<Snapshot>> {
    let key = key.into();
    let (owner_name, owner_blob, span, source, shape_blob) = {
        let part = package.get_part(owner)?;
        crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
        let span = crate::tag::shape::selected_raw_span(part.blob(), key)?;
        let shape_blob = part
            .blob()
            .get(span.clone())
            .ok_or_else(|| Error::Invalid("classification shape span is invalid".into()))?
            .to_vec();
        (
            part.partname().clone(),
            part.blob().to_vec(),
            span,
            codec::read(&shape_blob)?,
            shape_blob,
        )
    };
    let previous = source.snapshot.clone();
    let staged_shape = transaction::set(&shape_blob, &source, outcome)?;
    let staged_owner = codec::replace(&owner_blob, span.clone(), &staged_shape)?;
    validate_staged(&staged_owner, key, outcome)?;

    let part = package.get_part_mut(&owner_name)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != &owner_blob {
        return Err(Error::UnsafeEdit {
            operation: "put_classification",
            reason: "the selected slide changed during classification staging",
        });
    }
    if staged_owner == owner_blob {
        return Ok(previous);
    }
    part.set_blob(staged_owner);
    package.unsign();
    Ok(previous)
}

/// Remove the typed classification element while retaining unrelated entries.
pub(crate) fn remove<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
) -> Result<Option<Snapshot>> {
    let key = key.into();
    let (owner_name, owner_blob, span, source, shape_blob) = {
        let part = package.get_part(owner)?;
        crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
        let span = crate::tag::shape::selected_raw_span(part.blob(), key)?;
        let shape_blob = part
            .blob()
            .get(span.clone())
            .ok_or_else(|| Error::Invalid("classification shape span is invalid".into()))?
            .to_vec();
        (
            part.partname().clone(),
            part.blob().to_vec(),
            span,
            codec::read(&shape_blob)?,
            shape_blob,
        )
    };
    let Some(previous) = source.snapshot.clone() else {
        return Ok(None);
    };
    let Some(staged_shape) = transaction::remove(&shape_blob, &source)? else {
        return Ok(None);
    };
    let staged_owner = codec::replace(&owner_blob, span, &staged_shape)?;
    validate_removed(&staged_owner, key)?;

    let part = package.get_part_mut(&owner_name)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != &owner_blob {
        return Err(Error::UnsafeEdit {
            operation: "remove_classification",
            reason: "the selected slide changed during classification staging",
        });
    }
    if staged_owner != owner_blob {
        part.set_blob(staged_owner);
        package.unsign();
    }
    Ok(Some(previous))
}

fn validate_staged(owner: &[u8], key: crate::shape::Key<'_>, expected: Outcome) -> Result<()> {
    let span = crate::tag::shape::selected_raw_span(owner, key)?;
    let shape = owner
        .get(span)
        .ok_or_else(|| Error::Invalid("staged classification shape span is invalid".into()))?;
    let source = codec::read(shape)?;
    if source.snapshot.as_ref().and_then(Snapshot::outcome) != Some(expected) {
        return Err(Error::Invalid(
            "staged classification did not round-trip its typed outcome".into(),
        ));
    }
    Ok(())
}

fn validate_removed(owner: &[u8], key: crate::shape::Key<'_>) -> Result<()> {
    let span = crate::tag::shape::selected_raw_span(owner, key)?;
    let shape = owner
        .get(span)
        .ok_or_else(|| Error::Invalid("staged classification shape span is invalid".into()))?;
    let source = codec::read(shape)?;
    if source.snapshot.is_some() {
        return Err(Error::Invalid(
            "staged classification still contains a typed outcome".into(),
        ));
    }
    Ok(())
}
