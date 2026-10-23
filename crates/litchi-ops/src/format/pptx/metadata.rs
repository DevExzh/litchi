//! PPTX document metadata on the presentation root (record 0253).
//!
//! A set on `/` with OfficeCLI's document-property keys commits one leaf
//! [`litchi_pptx::metadata`] transaction on the eager package. A
//! source-backed owner is materialized once for the commit: the leaf's
//! source owner has no package-level part topology edit yet. The new handle
//! caches the committed metadata, so reads of `/` observe it.

use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Metadata};

use super::write::{Commit, Proof};
use super::{Editable, Pptx, PptxOwner, PptxState};
use crate::error::Result;
use crate::format::metadata::{self, Assignments};
use crate::format::signature::Opened;
use crate::options::OpenOptions;
use crate::props::Props;
use crate::target::TargetArg;
use crate::{opc, pptx};

fn native(error: pptx::metadata::Error) -> crate::Error {
    metadata::native_error(error, FileFormat::Pptx)
}

/// Commit `update` to `package`'s metadata.
fn commit(
    package: &pptx::Package,
    update: impl FnOnce(&mut Metadata),
    custom: Option<&Assignments>,
) -> Result<(pptx::Package, pptx::metadata::Patch, Metadata)> {
    let mut edit = package.edit_metadata().map_err(native)?;
    edit.update(update).map_err(native)?;
    if let Some(assignments) = custom.filter(|assignments| assignments.has_custom()) {
        let props = assignments.apply_custom(edit.custom_props(), FileFormat::Pptx)?;
        edit.set_custom_props(props);
    }
    // The staged value is the committed one; no second read of the parts.
    let metadata = edit.metadata().map_err(native)?;
    let (after, patch) = edit.commit().map_err(native)?.into_parts();
    Ok((after, patch, metadata))
}

impl Pptx {
    /// The eager package an [`Editable`] commits metadata on. An unrevised
    /// source owner adopts its source graph without re-serializing; a
    /// revised one publishes its revision first. A source-backed handle
    /// keeps that package once built (record 0262): the handle is
    /// immutable, so every later metadata set on it commits on the same
    /// before state, and the leaf's source graph adoption, not the edit,
    /// was most of a set's cost (record 0253 §7).
    fn metadata_owner(&self, source: &Editable) -> Result<Arc<pptx::Package>> {
        match source {
            Editable::Eager(package) => Ok(Arc::clone(package)),
            Editable::Source(package) => {
                if let Some(cached) = self.0.metadata_package.get() {
                    return Ok(Arc::clone(cached));
                }
                let graph = match (&self.0.revision, &self.0.owner) {
                    (None, PptxOwner::Source(presentation)) => presentation
                        .source_backed_package()
                        .to_opc_package_preserving_source()?,
                    _ => {
                        let mut bytes = Vec::new();
                        package.write_to(&mut bytes)?;
                        opc::OpcPackage::from_bytes(&bytes, self.0.ooxml_limits)?
                    },
                };
                let built = Arc::new(pptx::Package::from_opc_package(graph)?);
                Ok(Arc::clone(self.0.metadata_package.get_or_init(|| built)))
            },
        }
    }
}

impl Pptx {
    /// Set document properties on `/metadata`, or on `/` when a property
    /// is a metadata key (OfficeCLI's spelling, record 0338); `None` for
    /// other targets.
    pub(super) fn try_set_metadata(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Commit>> {
        let root = metadata::is_root(target) && metadata::names_metadata(props);
        if !(metadata::is_target(target) || root) {
            return Ok(None);
        }
        let assignments = metadata::admit(props, FileFormat::Pptx)?;
        self.set_metadata(&assignments).map(Some)
    }

    fn set_metadata(&self, assignments: &Assignments) -> Result<Commit> {
        let source = self.editable_owner()?;
        let package = self.metadata_owner(&source)?;
        let (after_package, patch, metadata) = commit(
            &package,
            |value| assignments.apply(value),
            Some(assignments),
        )?;
        if patch.is_noop() {
            return Ok(Commit {
                after: self.clone(),
                target: source.clone(),
                source,
                proof: Proof::Metadata(patch),
                operations: 0,
            });
        }
        let after_package = Arc::new(after_package);
        let operations = patch.changes().len();
        Ok(Commit {
            after: self.adopt_metadata(Arc::clone(&after_package), metadata),
            source: Editable::Eager(package),
            target: Editable::Eager(after_package),
            proof: Proof::Metadata(patch),
            operations,
        })
    }

    /// Adopt an eager metadata revision with its committed metadata cached.
    fn adopt_metadata(&self, package: Arc<pptx::Package>, metadata: Metadata) -> Self {
        let source_guard = match &self.0.owner {
            PptxOwner::Source(source) => Some(source.clone()),
            PptxOwner::Eager(_) => self.0.source_guard.clone(),
        };
        Self(Opened::new(PptxState {
            owner: PptxOwner::Eager(package),
            source_guard,
            metadata: OnceLock::from(Some(metadata).filter(Metadata::has_data)),
            retain_source: self.0.retain_source,
            ooxml_limits: self.0.ooxml_limits,
            eager_catalog: OnceLock::new(),
            masters: OnceLock::new(),
            related: super::owner::Related::default(),
            scenes: Arc::clone(&self.0.scenes),
            metadata_package: OnceLock::new(),
            promotion: OnceLock::new(),
            revision: None,
            unsigned_promotion: OnceLock::new(),
            admitted: OnceLock::new(),
            cells: std::sync::Mutex::new(None),
            notes: std::sync::Mutex::new(None),
            spare: std::sync::Mutex::new(None),
        }))
        .with_origin(self.origin())
    }
}

/// Create-time metadata: commit `metadata` to a fresh presentation.
pub(super) fn create(package: pptx::Package, metadata: &Metadata) -> Result<Pptx> {
    let (package, _, _) = commit(&package, |value| *value = metadata.clone(), None)?;
    Pptx::from_package(package, &OpenOptions::new())
}
