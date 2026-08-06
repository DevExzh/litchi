//! Borrowed typed presentation facade.

use litchi_opc::OpcPackage;

use crate::Result;
use crate::parts::{PresentationPart, SlideReference};
use crate::slide::{Key, Slide, SlideLayout, SlideMaster};

use super::embedded;
use super::package;

/// Borrowed semantic view of one PresentationML package graph.
pub struct Presentation<'a> {
    pub(super) package: &'a OpcPackage,
    pub(super) part: PresentationPart<'a>,
}

impl<'a> Presentation<'a> {
    /// Construct a view from a validated main part and its package.
    pub fn new(part: PresentationPart<'a>, package: &'a OpcPackage) -> Self {
        Self { package, part }
    }

    /// The underlying OPC package.
    #[inline]
    pub fn package(&self) -> &'a OpcPackage {
        self.package
    }

    /// The borrowed main-document part view.
    #[inline]
    pub fn part(&self) -> &PresentationPart<'a> {
        &self.part
    }

    /// The ordered, low-level slide references.
    pub fn slide_references(&self) -> Result<Vec<SlideReference>> {
        package::slide_references(&self.part)
    }

    /// Number of slides in the ordered presentation graph.
    pub fn slide_count(&self) -> Result<usize> {
        package::slide_count(&self.part)
    }

    /// Presentation slide size in EMUs.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        package::slide_size(&self.part)
    }

    /// Load the optional DrawingML table-style catalog owned by this
    /// presentation. The catalog remains a detached value so callers can
    /// inspect it without holding a mutable package borrow.
    pub fn styles(&self) -> Result<Option<crate::table::style::List>> {
        package::styles(self.package)
    }

    /// Load the complete inert speaker-notes graph, when present.
    pub fn notes(&self) -> Result<Option<crate::notes::Graph>> {
        package::notes(self.package, &self.part)
    }

    /// Discover the optional opaque VBA project relationship owned by this
    /// presentation. The binary payload is never decoded or executed.
    pub fn vba(&self) -> Result<Option<embedded::vba::Project>> {
        package::vba(self.package, &self.part)
    }

    /// Load all inert, opaque PresentationML content parts in slide order.
    ///
    /// The content-part anchor, relationship metadata, and referenced
    /// payload bytes are retained without interpreting or executing the
    /// external vocabulary stored by the producer.
    pub fn content_parts(&self) -> Result<Vec<embedded::content_parts::ContentPart>> {
        package::content_parts(self.package, &self.part)
    }

    /// Discover inert hyperlinks owned by the presentation's slides.
    ///
    /// Each result contains the zero-based slide position and a typed target.
    /// Relationship targets and inline actions are parsed as values only;
    /// they are never followed, opened, or executed.
    pub fn hyperlinks(&self) -> Result<Vec<(usize, crate::hyperlinks::Hyperlink)>> {
        package::hyperlinks(self.package, &self.part)
    }

    /// Resolve one ordered slide by zero-based index.
    pub fn slide(&self, index: usize) -> Result<Option<Slide<'a>>> {
        package::slide(self.package, &self.part, index)
    }

    /// Resolve a slide by checked index or exact producer-visible name.
    pub fn find_slide<'k>(&self, key: impl Into<Key<'k>>) -> Result<Option<Slide<'a>>> {
        package::find_slide(self.package, &self.part, key.into())
    }

    /// Resolve all slides in presentation order.
    pub fn slides(&self) -> Result<Vec<Slide<'a>>> {
        package::slides(self.package, &self.part)
    }

    /// Resolve the slide masters declared by `p:sldMasterIdLst` in XML order.
    pub fn slide_masters(&self) -> Result<Vec<SlideMaster<'a>>> {
        package::slide_masters(self.package, &self.part)
    }

    /// Resolve all layouts reachable from all presentation masters.
    pub fn slide_layouts(&self) -> Result<Vec<SlideLayout<'a>>> {
        package::slide_layouts(self.package, &self.part)
    }

    /// Flatten all slide text in presentation order.
    pub fn text(&self) -> Result<String> {
        package::text(self.package, &self.part)
    }

    /// Load the slide-library synchronization metadata reachable from this
    /// presentation's slide graph.
    pub fn slide_sync(
        &self,
    ) -> Result<Vec<crate::presentation_properties::metadata::slide_sync::Part>> {
        package::slide_sync(self.package)
    }
}
