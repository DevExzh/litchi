//! PPTX document metadata on the presentation root (record 0253).
//!
//! A set on `/` with OfficeCLI's document-property keys commits one leaf
//! [`litchi_pptx::metadata`] transaction on the eager package. A
//! source-backed owner is materialized once for the commit: the leaf's
//! source owner has no package-level part topology edit yet. The new handle
//! caches the committed metadata, so reads of `/` observe it.

use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Metadata};

use super::write::{Commit, Proof, native_snapshot_budget};
use super::{Editable, Pptx, PptxOwner, PptxState};
use crate::error::Result;
use crate::format::metadata::{self, Assignments};
use crate::format::signature::Opened;
use crate::options::OpenOptions;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::{opc, pptx};

fn native(error: pptx::metadata::Error) -> crate::Error {
    metadata::native_error(error, FileFormat::Pptx)
}

/// Commit `update` to `package`'s metadata.
fn commit(
    package: &pptx::Package,
    update: impl FnOnce(&mut Metadata),
    custom: Option<&Assignments>,
    max_patch_bytes: usize,
) -> Result<(pptx::Package, pptx::metadata::Patch, Metadata)> {
    let mut edit = package
        .edit_metadata_with_options(
            pptx::metadata::EditOptions::default().max_retained_bytes(max_patch_bytes),
        )
        .map_err(native)?;
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
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let root = metadata::is_root(target) && metadata::names_metadata(props);
        if !(metadata::is_target(target) || root) {
            return Ok(None);
        }
        // The native metadata transaction treats an empty update as a
        // genuine no-op and retains the complete property-family scope for
        // replay.  Preserve that proof at the facade boundary instead of
        // turning an empty `/metadata` setter into an InvalidArgument.  Root
        // metadata shorthand still reaches this route only when it names a
        // property, so other empty mutation targets keep their ordinary
        // admission behavior.
        let assignments = if props.is_empty() {
            Assignments::default()
        } else {
            metadata::admit(props, FileFormat::Pptx)?
        };
        self.set_metadata(&assignments, props.len(), limits)
            .map(Some)
    }

    fn set_metadata(
        &self,
        assignments: &Assignments,
        authored_operations: usize,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let unsupported_property = assignments.unsupported_property();
        // A signed metadata setter is staged through the unsigned private
        // owner so the publication fence can distinguish a genuine no-op
        // from a changed write.  Changed candidates are still refused by
        // `set_with_limits`; exact no-ops are re-rooted on the signed owner.
        let source = self.editable_owner_for_write()?;
        let package = self.metadata_owner(&source)?;
        let (after_package, patch, metadata) = commit(
            &package,
            |value| assignments.apply(value),
            Some(assignments),
            native_snapshot_budget(limits),
        )
        .map_err(|error| metadata::with_unsupported_property(error, unsupported_property))?;
        if patch.is_noop() {
            return Ok(Commit::new(
                self.clone(),
                source.clone(),
                source,
                Proof::Metadata(patch),
                authored_operations,
            ));
        }
        let after_package = Arc::new(after_package);
        Ok(Commit::new(
            self.adopt_metadata(Arc::clone(&after_package), metadata),
            Editable::Eager(package),
            Editable::Eager(after_package),
            Proof::Metadata(patch),
            authored_operations,
        ))
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
            cells: std::sync::Mutex::new(None),
            notes: std::sync::Mutex::new(None),
        }))
        .with_origin(self.origin())
    }
}

/// Create-time metadata: commit `metadata` to a fresh presentation.
pub(super) fn create(package: pptx::Package, metadata: &Metadata) -> Result<Pptx> {
    let unsupported_property = metadata::unsupported_field(metadata);
    let (package, _, _) = commit(
        &package,
        |value| *value = metadata.clone(),
        None,
        litchi_pptx::opened::Limits::default().max_patch_bytes(),
    )
    .map_err(|error| metadata::with_unsupported_property(error, unsupported_property))?;
    Pptx::from_package(package, &OpenOptions::new())
}
