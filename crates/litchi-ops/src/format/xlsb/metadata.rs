//! XLSB document metadata on the workbook root (record 0253).
//!
//! A set on `/` with OfficeCLI's document-property keys commits one leaf
//! [`litchi_xlsb::metadata`] transaction. The transaction history keeps the
//! canonical package images on both sides, as the structural patches do.

use litchi_core::{FileFormat, Metadata};
use litchi_xlsb as xlsb;

use super::write::{Commit, package_from_editor};
use super::{Xlsb, XlsbOwner};
use crate::error::Result;
use crate::format::metadata::{self, Assignments};
use crate::options::OpenOptions;
use crate::props::Props;
use crate::target::TargetArg;

/// The canonical package images around one metadata patch.
#[derive(Debug, Clone)]
pub(crate) struct Images {
    pub(super) before: Vec<u8>,
    pub(super) after: Vec<u8>,
    pub(super) patch: xlsb::metadata::Patch,
}

impl Images {
    pub(super) fn retained_bytes(&self) -> usize {
        self.before.len().saturating_add(self.after.len())
    }
}

fn native(error: xlsb::metadata::Error) -> crate::Error {
    metadata::native_error(error, FileFormat::Xlsb)
}

/// Commit `update` to `package`'s metadata.
fn commit(
    package: &xlsb::Package,
    update: impl FnOnce(&mut Metadata),
) -> Result<(xlsb::Package, xlsb::metadata::Patch, Metadata)> {
    let mut edit = package.edit_metadata().map_err(native)?;
    edit.update(update).map_err(native)?;
    let (after, patch) = edit.commit().map_err(native)?.into_parts();
    let metadata = after.metadata().map_err(native)?;
    Ok((after, patch, metadata))
}

impl Xlsb {
    /// Set document properties on the root path; `None` for other targets.
    pub(super) fn try_set_metadata(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Commit>> {
        if !metadata::is_root(target) {
            return Ok(None);
        }
        let assignments = metadata::admit(props, FileFormat::Xlsb)?;
        self.set_metadata(&assignments).map(Some)
    }

    fn set_metadata(&self, assignments: &Assignments) -> Result<Commit> {
        let source = package_from_editor(self.editor()?);
        let (target, patch, metadata) = commit(&source, |value| assignments.apply(value))?;
        let before = source.to_bytes()?;
        let (after, after_bytes) = if patch.is_noop() {
            (self.clone(), before.clone())
        } else {
            let bytes = target.to_bytes()?;
            (self.adopt_metadata(target, metadata)?, bytes)
        };
        self.check_source()?;
        Ok(Commit::from_metadata(
            after,
            Images {
                before,
                after: after_bytes,
                patch,
            },
        ))
    }

    /// [`Self::adopt_package`] with the committed metadata.
    fn adopt_metadata(&self, package: xlsb::Package, metadata: Metadata) -> Result<Self> {
        let options = OpenOptions::new()
            .retain_source(self.0.retain_source)
            .ooxml_limits(self.0.ooxml_limits);
        let workbook = package.into_workbook()?;
        let signature = match self.signature() {
            super::Signature::RemovalAuthorized => super::Signature::Unsigned,
            signature => signature,
        };
        Ok(Self::with_owner(
            XlsbOwner::Eager(Box::new(workbook)),
            metadata,
            self.0.metadata_repairs.clone(),
            &options,
        )
        .with_origin(self.origin())
        .with_signature(signature))
    }
}

/// Create-time metadata: commit `metadata` to a fresh package.
pub(super) fn create(package: xlsb::Package, metadata: &Metadata) -> Result<Xlsb> {
    let (package, _, _) = commit(&package, |value| *value = metadata.clone())?;
    Xlsb::from_package(package, &OpenOptions::new())
}
