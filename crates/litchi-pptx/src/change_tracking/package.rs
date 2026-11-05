//! Slide graph discovery and atomic change-tracking publication.

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

use super::Snapshot;
use super::transaction::{Change, Prepared};
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

pub(crate) fn load_with_limits(
    package: &OpcPackage,
    owner: &PackURI,
    shape_limits: crate::shape::Limits,
) -> Result<Snapshot> {
    let part = package.get_part(owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    Snapshot::from_source_with_limits(part.partname().clone(), part.blob().to_vec(), shape_limits)
}

pub(crate) fn apply_commit_with_limits(
    package: &mut OpcPackage,
    commit: Prepared,
    shape_limits: crate::shape::Limits,
) -> Result<Snapshot> {
    let (snapshot, patch) = commit.into_parts();
    if patch.after() != snapshot.state() {
        return Err(Error::Invalid(
            "change-tracking commit differs from its candidate snapshot".into(),
        ));
    }
    apply_patch_with_limits(package, &patch, shape_limits)
}

pub(crate) fn apply_patch_with_limits(
    package: &mut OpcPackage,
    patch: &Change,
    shape_limits: crate::shape::Limits,
) -> Result<Snapshot> {
    let part = package.get_part(&patch.owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != patch.source {
        return Err(Error::PatchConflict);
    }
    let snapshot =
        Snapshot::from_source_with_limits(patch.owner.clone(), patch.target.clone(), shape_limits)?;
    if snapshot.state() != patch.after() {
        return Err(Error::Invalid(
            "change-tracking patch target differs from its semantic result".into(),
        ));
    }
    if !patch.is_noop() {
        refuse_signed_change(package)?;
        let mutable_part = package.get_part_mut(&patch.owner)?;
        crate::parts::validate_content_type(mutable_part, ct::PML_SLIDE)?;
        if mutable_part.blob() != patch.source {
            return Err(Error::PatchConflict);
        }
        mutable_part.set_blob(patch.target.clone());
        settle_unsigned_change(package)?;
    }
    Ok(snapshot)
}
