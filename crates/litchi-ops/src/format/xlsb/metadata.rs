//! XLSB document metadata on the workbook root (records 0253, 0261).
//!
//! A set on `/` with OfficeCLI's document-property keys commits one leaf
//! [`litchi_xlsb::metadata`] transaction. A source-backed handle commits the
//! leaf's source transaction ([`litchi_xlsb::metadata::source`]): the
//! property parts are planned over the retained source and the handle stays
//! source-backed, with no workbook materialization. The transaction history
//! keeps the canonical package images on both sides, as the structural
//! patches do.

use std::sync::OnceLock;

use litchi_core::{FileFormat, Metadata};
use litchi_xlsb as xlsb;

use super::write::Commit;
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
    before_file: Xlsb,
    after_file: Option<Xlsb>,
    before: OnceLock<Vec<u8>>,
    after: OnceLock<Vec<u8>>,
    pub(super) patch: xlsb::metadata::Patch,
}

impl Images {
    fn image(file: &Xlsb) -> Vec<u8> {
        let mut bytes = Vec::new();
        match file.write_to(&mut bytes) {
            Ok(()) => bytes,
            Err(_) => b"<unserializable metadata image>".to_vec(),
        }
    }

    pub(super) fn before(&self) -> &[u8] {
        self.before.get_or_init(|| Self::image(&self.before_file))
    }

    pub(super) fn after(&self) -> &[u8] {
        match &self.after_file {
            Some(file) => self.after.get_or_init(|| Self::image(file)),
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
    edit: std::result::Result<xlsb::metadata::Edit, xlsb::metadata::Error>,
    update: impl FnOnce(&mut Metadata),
    custom: Option<&Assignments>,
) -> Result<(xlsb::Package, xlsb::metadata::Patch, Metadata)> {
    let mut edit = edit.map_err(native)?;
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
        if let (XlsbOwner::Source(source), super::Signature::Unsigned) =
            (&self.0.owner, self.signature())
        {
            if self.0.retain_source {
                return self.set_source_metadata(source, assignments);
            }
        }
        // Stage on the workbook itself: no copy into a package first
        // (record 0281).
        let editor = self.editor()?;
        let (target, patch, metadata) = commit(
            editor.edit_metadata(),
            |value| assignments.apply(value),
            Some(assignments),
        )?;
        drop(editor);
        let (after, after_file) = if patch.is_noop() {
            (self.clone(), None)
        } else {
            let after = self.adopt_metadata(target, metadata)?;
            (after.clone(), Some(after))
        };
        self.check_source()?;
        Ok(Commit::from_metadata(
            after,
            Images {
                before_file: self.clone(),
                after_file,
                before: OnceLock::new(),
                after: OnceLock::new(),
                patch,
            },
        ))
    }

    /// The source-backed set: the leaf plans the property parts over the
    /// retained source, and the result stays a source-backed handle that
    /// shares the catalog and every worksheet cache with this one.
    fn set_source_metadata(
        &self,
        source: &super::source::SourceWorkbook,
        assignments: &Assignments,
    ) -> Result<Commit> {
        self.check_source()?;
        let mut edit = source.workbook().edit_metadata().map_err(native)?;
        edit.update(|value| assignments.apply(value))
            .map_err(native)?;
        if assignments.has_custom() {
            let props = assignments.apply_custom(edit.custom_props(), FileFormat::Xlsb)?;
            edit.set_custom_props(props);
        }
        // The staged value is the committed one; no second read of the parts.
        let metadata = edit.metadata().map_err(native)?;
        let (workbook, patch) = edit.commit().map_err(native)?.into_parts();
        let (after, after_file) = if patch.is_noop() {
            (self.clone(), None)
        } else {
            let options = OpenOptions::new()
                .retain_source(self.0.retain_source)
                .ooxml_limits(self.0.ooxml_limits);
            let after = Self::with_owner(
                XlsbOwner::Source(source.with_workbook(workbook)),
                metadata,
                self.0.metadata_repairs.clone(),
                &options,
            )
            .with_origin(self.origin())
            .with_signature(self.signature());
            (after.clone(), Some(after))
        };
        self.check_source()?;
        Ok(Commit::from_metadata(
            after,
            Images {
                before_file: self.clone(),
                after_file,
                before: OnceLock::new(),
                after: OnceLock::new(),
                patch,
            },
        ))
    }

    /// [`Self::adopt_package`] with the committed metadata.
    fn adopt_metadata(&self, package: xlsb::Package, metadata: Metadata) -> Result<Self> {
        let options = OpenOptions::new()
            .retain_source(self.0.retain_source)
            .ooxml_limits(self.0.ooxml_limits);
        let signature = match self.signature() {
            super::Signature::RemovalAuthorized => super::Signature::Unsigned,
            signature => signature,
        };
        Ok(Self::with_owner(
            super::eager_package(package)?,
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
    let (package, _, _) = commit(
        package.edit_metadata(),
        |value| *value = metadata.clone(),
        None,
    )?;
    Xlsb::from_package(package, &OpenOptions::new())
}

#[cfg(test)]
#[path = "metadata_tests.rs"]
mod tests;
