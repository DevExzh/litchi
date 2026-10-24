//! The SVG attachment family of a source-backed package (record 0355a):
//! attach an SVG to, or detach it from, direct main-story pictures, keeping
//! their raster fallback.
//!
//! [`Package::edit_svg`](crate::source_backed::Package::edit_svg) stages an
//! [`Edit`] over selected pictures whose commit returns a new source-backed
//! package; writing it adds or removes the SVG parts and relationships and
//! rewrites the main story.

use std::cell::{RefCell, RefMut};

use crate::Result;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::svg_lifecycle::{BatchEdit as RecordEdit, BatchPatch as RecordPatch};
use crate::source_backed::{Package, svg_lifecycle};

pub use svg_lifecycle::{
    Attachment, AttachmentReplacement, Input, OwnerState, Picture, PictureSelector, PictureSource,
    PictureView, Raster, RasterSource, Resource, ResourceSource, Snapshot,
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
        let stream_base = base.clone();
        let stream_current = current.clone();
        let stream_target = target.clone();
        package.deferred(
            move |writer, policy| base.write_svg_batch(writer, &current, &target, policy),
            move |target_sink| {
                stream_base.stream_svg_batch(target_sink, &stream_current, &stream_target)
            },
        )
    }
}

/// Content-free facts about an SVG commit: staged operations, selected
/// pictures and whether it changed anything.
pub use svg_lifecycle::BatchDiagnostics as Diagnostics;

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
    selectors: Vec<PictureSelector>,
    inner: RefCell<Option<RecordEdit<'a>>>,
}

impl std::fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn materialize(&self) -> Result<RecordEdit<'a>> {
        self.package.edit_svg_record(self.selectors.clone())
    }

    fn ensure_inner(&self) -> Result<()> {
        let mut inner = self.inner.borrow_mut();
        if inner.is_none() {
            *inner = Some(self.materialize()?);
        }
        Ok(())
    }

    fn inner_mut(&self) -> Result<RefMut<'_, RecordEdit<'a>>> {
        self.ensure_inner()?;
        RefMut::filter_map(self.inner.borrow_mut(), Option::as_mut)
            .map_err(|_| crate::Error::internal("svg.edit", "SVG edit materializer is unavailable"))
    }

    fn into_inner(mut self) -> Result<RecordEdit<'a>> {
        if self.inner.get_mut().is_none() {
            let inner = self.materialize()?;
            *self.inner.get_mut() = Some(inner);
        }
        self.inner.into_inner().ok_or_else(|| {
            crate::Error::internal("svg.edit", "SVG edit materializer is unavailable")
        })
    }

    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        self.package
    }

    /// The selected pictures before the edit.
    pub fn source(&self) -> Result<Snapshot> {
        self.ensure_inner()?;
        Ok(self
            .inner
            .borrow()
            .as_ref()
            .ok_or_else(|| crate::Error::internal("svg.edit", "SVG source is unavailable"))?
            .source()
            .clone())
    }

    /// The selected pictures with the staged changes.
    pub fn projected(&self) -> Result<Snapshot> {
        self.ensure_inner()?;
        Ok(self
            .inner
            .borrow()
            .as_ref()
            .ok_or_else(|| crate::Error::internal("svg.edit", "SVG projection is unavailable"))?
            .projected()
            .clone())
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
        input: impl Into<Input<'b>>,
    ) -> Result<bool> {
        self.inner_mut()?.attach_svg(selector, input)
    }

    /// Attach with explicit identities after all checks.
    ///
    /// # Errors
    ///
    /// As [`Self::attach_svg`].
    pub fn attach(
        &mut self,
        selector: impl Into<PictureSelector>,
        request: &AttachmentReplacement,
    ) -> Result<bool> {
        self.inner_mut()?.attach(selector, request)
    }

    /// Detach a selected picture's SVG, keeping its raster.
    ///
    /// # Errors
    ///
    /// Returns an error for an unselected picture.
    pub fn detach_svg(&mut self, selector: impl Into<PictureSelector>) -> Result<bool> {
        self.inner_mut()?.detach_svg(selector)
    }

    /// Commit the staged changes as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed closure, graph, signature or source error before any
    /// output.
    pub fn commit(self) -> Result<Commit> {
        let this = self;
        let package = this.package;
        let commit = this.into_inner()?.commit()?;
        let operations = commit.diagnostics().operations();
        commit_change(package, SvgChange(commit.patch().clone(), operations))
    }
}

impl Package {
    /// Stage one atomic SVG change spanning `selectors`. The selectors are
    /// retained without reading drawing payloads; the drawing and relationship
    /// inventory are read once by the first operation, projection or commit.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, drawing or limit error.
    pub fn edit_svg<I, S>(&self, selectors: I) -> Result<Edit<'_>>
    where
        I: IntoIterator<Item = S>,
        S: Into<PictureSelector>,
    {
        // Reading the retained policy must not materialize a pending output
        // from a prior source-backed commit.  The record constructor below
        // performs the physical source check and inventory capture only when
        // the first operation, projection or commit needs it.
        let selector_limit = self.read_limits().max_total_relationships();
        let mut selected = Vec::new();
        for selector in selectors {
            if selected.len() >= selector_limit {
                return Err(crate::Error::Invalid(
                    "SVG lifecycle selector count exceeds the bounded batch limit".into(),
                ));
            }
            selected
                .try_reserve(1)
                .map_err(|source| crate::Error::Allocation {
                    resource: "DOCX SVG lifecycle selectors",
                    source,
                })?;
            selected.push(selector.into());
        }
        Ok(Edit {
            package: self,
            selectors: selected,
            inner: RefCell::new(None),
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
