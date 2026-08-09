//! Slide graph discovery and atomic change-tracking publication.

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

use super::{Commit, Patch, Snapshot};
use crate::{Error, Result};

pub(crate) fn load(package: &OpcPackage, owner: &PackURI) -> Result<Snapshot> {
    let part = package.get_part(owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    Snapshot::from_source(part.partname().clone(), part.blob().to_vec())
}

pub(crate) fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    let (snapshot, patch) = commit.into_parts();
    if patch.after() != snapshot.state() {
        return Err(Error::Invalid(
            "change-tracking commit differs from its candidate snapshot".into(),
        ));
    }
    apply_patch(package, &patch)
}

pub(crate) fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let part = package.get_part(&patch.owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != patch.source {
        return Err(Error::UnsafeEdit {
            operation: "apply_change_tracking_patch",
            reason: "the selected slide no longer matches the patch source",
        });
    }
    let snapshot = Snapshot::from_source(patch.owner.clone(), patch.target.clone())?;
    if snapshot.state() != patch.after() {
        return Err(Error::Invalid(
            "change-tracking patch target differs from its semantic result".into(),
        ));
    }
    if patch.is_changed() {
        let mutable_part = package.get_part_mut(&patch.owner)?;
        crate::parts::validate_content_type(mutable_part, ct::PML_SLIDE)?;
        if mutable_part.blob() != patch.source {
            return Err(Error::UnsafeEdit {
                operation: "apply_change_tracking_patch",
                reason: "the selected slide changed during patch validation",
            });
        }
        mutable_part.set_blob(patch.target.clone());
        package.unsign();
    }
    Ok(snapshot)
}
