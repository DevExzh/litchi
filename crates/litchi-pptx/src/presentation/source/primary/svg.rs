//! Immutable SVG picture edits of a source-backed [`Package`] (record
//! 0356a): replacing an embedded SVG picture's SVG and raster fallback, and
//! attaching or detaching the native SVG of a raster picture.
//!
//! [`Package::edit_svg_images`] starts an [`Edit`]; [`Edit::replace`],
//! [`Edit::attach`] and [`Edit::detach`] stage one picture change each.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_svg_images`] replays a patch onto a package holding its
//! before state. The source package never changes.
//!
//! The committed package stays source-backed: writing it
//! ([`Package::write_to`], [`Package::save`]) streams the source with only
//! the picture's media parts and owning slide relationships changed, raw
//! copying every other member, so memory stays bounded by the changed
//! parts. Reading or editing the committed package materializes the
//! published artifact once. A picture is selected by the zero-based
//! position of its slide and of the picture on it, as the other
//! source-backed edits select slides. A no-op keeps every byte; a changed
//! edit of a signed source is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
//! before anything is staged.
//!
//! ```no_run
//! use litchi_core::Position;
//! use litchi_pptx::source::{Package, svg};
//!
//! # fn main() -> litchi_pptx::Result<()> {
//! let package = Package::open("deck.pptx")?;
//! let mut edit = package.edit_svg_images()?;
//! let replacement = svg::Replacement::new(b"<svg/>".to_vec(), b"\x89PNG".to_vec());
//! edit.replace(Position::new(0), Position::new(0), replacement)?;
//! let (updated, patch) = edit.commit()?.into_parts();
//! updated.save("updated.pptx")?;
//! let restored = updated.apply_svg_images(&patch.inverse())?.into_package();
//! # let _ = restored;
//! # Ok(())
//! # }
//! ```

use std::sync::Arc;

use litchi_core::Position;

use super::Package;
use crate::presentation::source::{SourceBackedSvgAttachmentPatch, SourceBackedSvgPatch};

pub use crate::Error;
use crate::Result;
pub use crate::opened::LimitKind;
pub use crate::presentation::source::{
    AttachedSvg, SvgAttachment as Attachment, SvgAttachmentReplacement as AttachmentReplacement,
    SvgImage as Image, SvgReplacement as Replacement,
};
pub use litchi_core::edit::Diagnostics;

/// The result of an SVG picture edit's `commit`: the new source-backed
/// package, the patch and the diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// One staged picture change.
#[derive(Clone)]
enum Step {
    Replace(SourceBackedSvgPatch),
    Attachment(SourceBackedSvgAttachmentPatch),
}

impl Step {
    fn is_noop(&self) -> bool {
        match self {
            Self::Replace(patch) => patch.is_noop(),
            Self::Attachment(patch) => patch.is_noop(),
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Replace(patch) => Self::Replace(patch.inverse()),
            Self::Attachment(patch) => Self::Attachment(patch.inverse()),
        }
    }

    /// Check the step against `package`'s current state and return the
    /// committed package whose output publishes it.
    fn publish(&self, package: &Package) -> Result<Package> {
        let base = physical(package)?;
        let owner = Arc::clone(&base.owner);
        // A changed write of a signed source is refused before any output,
        // as the source publisher would refuse it.
        if !self.is_noop() && owner.package.has_signature_infrastructure() {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        match self {
            Self::Replace(patch) => {
                let (slide, image) = patch.location();
                let current = owner.svg_image_edit(slide, image)?.source().clone();
                patch.apply(&current)?;
                if patch.is_noop() {
                    return Ok(base);
                }
                let commit = patch.to_commit();
                Ok(base.deferred(move |writer| {
                    (*owner)
                        .clone()
                        .publish_svg_commit_to_stream(writer, &commit)
                        .map(drop)
                }))
            },
            Self::Attachment(patch) => {
                let (slide, image) = patch.location();
                let current = owner.svg_attachment_snapshot(slide, image)?;
                patch.apply(&current)?;
                if patch.is_noop() {
                    return Ok(base);
                }
                let commit = patch.to_commit();
                Ok(base.deferred(move |writer| {
                    (*owner)
                        .clone()
                        .publish_svg_attachment_commit_to_stream(writer, &commit)
                        .map(drop)
                }))
            },
        }
    }
}

/// The exact, reversible change of an SVG picture edit: the exact before
/// and after state of each picture it changed, in order.
#[derive(Clone, Default)]
pub struct Patch {
    steps: Vec<Step>,
    /// The packages the patch was committed between, so its inverse applied
    /// to the commit's own package restores the exact source package.
    restore: Option<(Package, Package)>,
}

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.steps.iter().all(Step::is_noop)
    }

    /// The patch that restores the before state.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            steps: self.steps.iter().rev().map(Step::inverse).collect(),
            restore: self
                .restore
                .as_ref()
                .map(|(before, after)| (after.clone(), before.clone())),
        }
    }
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("steps", &self.steps.len())
            .field("noop", &self.is_noop())
            .finish()
    }
}

/// A staged SVG picture edit of one source-backed [`Package`]. Dropping it
/// discards it.
pub struct Edit<'a> {
    source: &'a Package,
    current: Package,
    steps: Vec<Step>,
}

impl std::fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("steps", &self.steps.len())
            .finish_non_exhaustive()
    }
}

/// `package` as a source-backed package with no staged state: itself when
/// it is pristine, else its published artifact reopened (record 0356a).
pub(super) fn physical(package: &Package) -> Result<Package> {
    let effective = package.effective()?;
    if effective.is_pristine() {
        return Ok(effective.clone());
    }
    let mut options = crate::OpenOptions::default()
        .source_limits(effective.owner.limits)
        .cache_limits(effective.owner.package.cache_limits());
    options.semantic_limits = effective.owner.semantic_limits.clone();
    let options = match effective.owner.package.execution_context() {
        Some(context) => options.execution_context(context),
        None => options,
    };
    Package::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(effective.to_bytes()?)),
        options,
    )
}

impl Package {
    /// Start an SVG picture edit.
    ///
    /// # Errors
    ///
    /// Returns the source's execution or version error.
    pub fn edit_svg_images(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            source: self,
            current: self.effective()?.clone(),
            steps: Vec::new(),
        })
    }

    /// Apply an SVG picture patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package does not hold the
    /// patch's before state. A picture patch is exact-source: it replays onto
    /// a package over the same retained source, and its inverse applied to
    /// the commit's own package restores the package it was made on.
    pub fn apply_svg_images(&self, patch: &Patch) -> Result<Commit> {
        if let Some((before, after)) = &patch.restore
            && self.same_content(before)?
        {
            return Ok(Commit::new(
                after.clone(),
                patch.clone(),
                Diagnostics::new(!patch.is_noop()),
            ));
        }
        let mut current = self.clone();
        for step in &patch.steps {
            current = step.publish(&current)?;
        }
        Ok(Commit::new(
            current,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}

impl Edit<'_> {
    /// The staged state of the embedded SVG picture `image` on the slide
    /// `slide`.
    ///
    /// # Errors
    ///
    /// Returns an error for a position out of range, or a picture that is not
    /// a direct embedded SVG picture with an internal raster fallback.
    pub fn image(&self, slide: impl Into<Position>, image: impl Into<Position>) -> Result<Image> {
        let base = physical(&self.current)?;
        Ok(base
            .owner
            .svg_image_edit(slide.into().get(), image.into().get())?
            .source()
            .clone())
    }

    /// The staged SVG attachment state of the raster picture `image` on the
    /// slide `slide`.
    ///
    /// # Errors
    ///
    /// Returns an error for a position out of range, or a picture that is not
    /// a direct raster picture.
    pub fn attachment(
        &self,
        slide: impl Into<Position>,
        image: impl Into<Position>,
    ) -> Result<Attachment> {
        let base = physical(&self.current)?;
        base.owner
            .svg_attachment_snapshot(slide.into().get(), image.into().get())
    }

    /// Replace the SVG and raster fallback payloads of the embedded SVG
    /// picture `image` on the slide `slide`. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a position out of range, a picture that is not
    /// a direct embedded SVG picture, invalid media part names, empty or
    /// oversized payloads.
    pub fn replace(
        &mut self,
        slide: impl Into<Position>,
        image: impl Into<Position>,
        replacement: Replacement,
    ) -> Result<bool> {
        let base = physical(&self.current)?;
        let mut edit = base
            .owner
            .svg_image_edit(slide.into().get(), image.into().get())?;
        edit.replace(replacement)?;
        let patch = edit.commit_checked()?.patch().clone();
        self.stage(Step::Replace(patch))
    }

    /// Attach a native SVG to the raster picture `image` on the slide
    /// `slide`, or replace its attached SVG. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a position out of range, a picture that is not a
    /// direct raster picture, a linked SVG, or an invalid or oversized
    /// payload.
    pub fn attach(
        &mut self,
        slide: impl Into<Position>,
        image: impl Into<Position>,
        replacement: &AttachmentReplacement,
    ) -> Result<bool> {
        let base = physical(&self.current)?;
        let mut edit = base
            .owner
            .edit_svg_attachment(slide.into().get(), image.into().get())?;
        edit.attach(replacement)?;
        let patch = edit.commit_checked()?.patch().clone();
        self.stage(Step::Attachment(patch))
    }

    /// Detach the native SVG of the picture `image` on the slide `slide`,
    /// keeping its raster. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a position out of range or a picture that is not
    /// a direct raster picture.
    pub fn detach(
        &mut self,
        slide: impl Into<Position>,
        image: impl Into<Position>,
    ) -> Result<bool> {
        let base = physical(&self.current)?;
        let mut edit = base
            .owner
            .edit_svg_attachment(slide.into().get(), image.into().get())?;
        edit.detach()?;
        let patch = edit.commit_checked()?.patch().clone();
        self.stage(Step::Attachment(patch))
    }

    fn stage(&mut self, step: Step) -> Result<bool> {
        if step.is_noop() {
            return Ok(false);
        }
        self.current = step.publish(&self.current)?;
        self.steps.push(step);
        Ok(true)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.current = self.source.clone();
        self.steps.clear();
    }

    /// Publish the staged changes to a new package.
    ///
    /// # Errors
    ///
    /// Infallible today; the `Result` keeps the family's commit shape.
    pub fn commit(self) -> Result<Commit> {
        let changed = !self.steps.iter().all(Step::is_noop);
        let package = if changed {
            self.current
        } else {
            self.source.clone()
        };
        let patch = Patch {
            steps: self.steps,
            restore: Some((self.source.clone(), package.clone())),
        };
        Ok(Commit::new(package, patch, Diagnostics::new(changed)))
    }
}
