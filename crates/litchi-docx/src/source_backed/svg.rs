//! The SVG attachment family of a source-backed package (record 0355a):
//! attach an SVG to, or detach it from, direct main-story pictures, keeping
//! their raster fallback.
//!
//! [`Package::edit_svg`](crate::source_backed::Package::edit_svg) stages an
//! [`Edit`] over selected pictures whose commit returns a new source-backed
//! package; writing it adds or removes the SVG parts and relationships and
//! rewrites the main story.

use crate::Result;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::{
    Package, PictureSelector, SourceBackedSvgAttachmentBatchEdit as RecordEdit,
    SourceBackedSvgAttachmentBatchPatch as RecordPatch,
    SourceBackedSvgAttachmentBatchSnapshot as RecordSnapshot, SourceBackedSvgAttachmentReplacement,
    SvgInput,
};

/// The batch record patch the family publishes, with its staged operation
/// count.
#[derive(Clone)]
pub(crate) struct SvgChange(RecordPatch, usize);

impl SvgChange {
    /// The commit's content-free facts.
    fn diagnostics(&self) -> Diagnostics {
        Diagnostics::new(self.1, self.0.source().len(), !self.0.is_noop())
    }
}

impl Change for SvgChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        self.0.is_noop()
    }

    fn inverse(&self) -> Self {
        Self(self.0.inverse(), self.1)
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        let (current, target) = package.svg_batch_target(&self.0)?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        let base = package.clone();
        package
            .deferred(move |writer, policy| base.write_svg_batch(writer, &current, &target, policy))
    }
}

/// Content-free facts about an SVG commit: staged operations, selected
/// pictures and whether it changed anything.
pub use crate::source_backed::SvgAttachmentCommitDiagnostics as Diagnostics;

source_family!("SVG attachment", SvgChange, crate::Error, diagnostics Diagnostics);

impl From<RecordPatch> for Patch {
    /// Replay a batch record patch onto a package holding its before state.
    /// Its diagnostics count one operation per changed picture.
    fn from(patch: RecordPatch) -> Self {
        let operations = patch.changed_pictures();
        Self {
            change: SvgChange(patch, operations),
            restore: None,
        }
    }
}

/// The budgets of this family: the package's open source limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, _formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match *self {}
    }
}

/// A staged SVG change of selected pictures. The source is never changed:
/// [`Self::commit`] returns a new package.
pub struct Edit<'a> {
    package: &'a Package,
    inner: RecordEdit<'a>,
}

impl std::fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        self.package
    }

    /// The selected pictures before the edit.
    #[must_use]
    pub const fn source(&self) -> &RecordSnapshot {
        self.inner.source()
    }

    /// The selected pictures with the staged changes.
    #[must_use]
    pub const fn projected(&self) -> &RecordSnapshot {
        self.inner.projected()
    }

    /// Attach an SVG payload to a selected picture; returns whether it
    /// changed anything.
    ///
    /// # Errors
    ///
    /// Returns an error for an unselected picture or an invalid payload.
    pub fn attach_svg<'b>(
        &mut self,
        selector: impl Into<PictureSelector>,
        input: impl Into<SvgInput<'b>>,
    ) -> Result<bool> {
        self.inner.attach_svg(selector, input)
    }

    /// Attach with explicit identities after all checks.
    ///
    /// # Errors
    ///
    /// As [`Self::attach_svg`].
    pub fn attach(
        &mut self,
        selector: impl Into<PictureSelector>,
        request: &SourceBackedSvgAttachmentReplacement,
    ) -> Result<bool> {
        self.inner.attach(selector, request)
    }

    /// Detach a selected picture's SVG, keeping its raster.
    ///
    /// # Errors
    ///
    /// Returns an error for an unselected picture.
    pub fn detach_svg(&mut self, selector: impl Into<PictureSelector>) -> Result<bool> {
        self.inner.detach_svg(selector)
    }

    /// Commit the staged changes as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed closure, graph, signature or source error before any
    /// output.
    pub fn commit(self) -> Result<Commit> {
        let commit = self.inner.commit()?;
        let operations = commit.diagnostics().operations();
        commit_change(self.package, SvgChange(commit.patch().clone(), operations))
    }
}

impl Package {
    /// Stage one atomic SVG change spanning `selectors`. The drawing and
    /// relationship inventory are read once.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, drawing or limit error.
    pub fn edit_svg<I, S>(&self, selectors: I) -> Result<Edit<'_>>
    where
        I: IntoIterator<Item = S>,
        S: Into<PictureSelector>,
    {
        Ok(Edit {
            package: self,
            inner: self.edit_svg_attachments(selectors)?,
        })
    }

    /// Replay an SVG patch onto this package.
    ///
    /// # Errors
    ///
    /// Returns an error for a package that does not hold the before state.
    pub fn apply_svg(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
