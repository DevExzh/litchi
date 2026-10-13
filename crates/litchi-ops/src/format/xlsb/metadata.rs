//! XLSB document metadata on the workbook root (record 0253).
//!
//! A set on `/` with OfficeCLI's document-property keys commits one leaf
//! [`litchi_xlsb::metadata`] transaction. The transaction history keeps the
//! canonical package images on both sides, as the structural patches do.

use std::sync::OnceLock;

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
///
/// The images serialize on first use: a one-shot set never needs them, and
/// transaction history reads them once. A serialization failure yields an
/// image that equals no package image, so history reports a change rather
/// than hiding one.
#[derive(Debug, Clone)]
pub(crate) struct Images {
    before_package: xlsb::Package,
    after_file: Option<Xlsb>,
    before: OnceLock<Vec<u8>>,
    after: OnceLock<Vec<u8>>,
    pub(super) patch: xlsb::metadata::Patch,
}

impl Images {
    fn image(package: &xlsb::Package) -> Vec<u8> {
        package
            .to_bytes()
            .unwrap_or_else(|_| b"<unserializable metadata image>".to_vec())
    }

    pub(super) fn before(&self) -> &[u8] {
        self.before
            .get_or_init(|| Self::image(&self.before_package))
    }

    pub(super) fn after(&self) -> &[u8] {
        match &self.after_file {
            Some(file) => self.after.get_or_init(|| {
                let mut bytes = Vec::new();
                match file.write_to(&mut bytes) {
                    Ok(()) => bytes,
                    Err(_) => b"<unserializable metadata image>".to_vec(),
                }
            }),
            None => self.before(),
        }
    }

    pub(super) fn retained_bytes(&self) -> usize {
        self.before().len().saturating_add(self.after().len())
    }
}

fn native(error: xlsb::metadata::Error) -> crate::Error {
    metadata::native_error(error, FileFormat::Xlsb)
}

/// Commit `update` to `package`'s metadata.
fn commit(
    package: &xlsb::Package,
    update: impl FnOnce(&mut Metadata),
    custom: Option<&Assignments>,
) -> Result<(xlsb::Package, xlsb::metadata::Patch, Metadata)> {
    let mut edit = package.edit_metadata().map_err(native)?;
    edit.update(update).map_err(native)?;
    if let Some(assignments) = custom.filter(|assignments| assignments.has_custom()) {
        let props = assignments.apply_custom(edit.custom_props(), FileFormat::Xlsb)?;
        edit.set_custom_props(props);
    }
    // The staged value is the committed one; no second read of the parts.
    let metadata = edit.metadata().map_err(native)?;
    let (after, patch) = edit.commit().map_err(native)?.into_parts();
    Ok((after, patch, metadata))
}

impl Xlsb {
    /// Set document properties on `/metadata`; `None` for other targets.
    pub(super) fn try_set_metadata(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Commit>> {
        if !metadata::is_target(target) {
            return Ok(None);
        }
        let assignments = metadata::admit(props, FileFormat::Xlsb)?;
        self.set_metadata(&assignments).map(Some)
    }

    fn set_metadata(&self, assignments: &Assignments) -> Result<Commit> {
        let source = package_from_editor(self.editor()?);
        let (target, patch, metadata) =
            commit(&source, |value| assignments.apply(value), Some(assignments))?;
        let (after, after_file) = if patch.is_noop() {
            (self.clone(), None)
        } else {
            let after = self.adopt_metadata(target.into_workbook()?, metadata);
            (after.clone(), Some(after))
        };
        self.check_source()?;
        Ok(Commit::from_metadata(
            after,
            Images {
                before_package: source,
                after_file,
                before: OnceLock::new(),
                after: OnceLock::new(),
                patch,
            },
        ))
    }

    /// [`Self::adopt_package`] with the committed metadata.
    fn adopt_metadata(&self, workbook: xlsb::Workbook, metadata: Metadata) -> Self {
        let options = OpenOptions::new()
            .retain_source(self.0.retain_source)
            .ooxml_limits(self.0.ooxml_limits);
        let signature = match self.signature() {
            super::Signature::RemovalAuthorized => super::Signature::Unsigned,
            signature => signature,
        };
        Self::with_owner(
            XlsbOwner::Eager(Box::new(workbook)),
            metadata,
            self.0.metadata_repairs.clone(),
            &options,
        )
        .with_origin(self.origin())
        .with_signature(signature)
    }
}

/// Create-time metadata: commit `metadata` to a fresh package.
pub(super) fn create(package: xlsb::Package, metadata: &Metadata) -> Result<Xlsb> {
    let (package, _, _) = commit(&package, |value| *value = metadata.clone(), None)?;
    Xlsb::from_package(package, &OpenOptions::new())
}
