//! Borrowed typed presentation facade.

use litchi_core::{TextOutputError, TextOutputOptions, TextOutputReport};
use litchi_opc::OpcPackage;
use std::io::Write;
use std::sync::OnceLock;

use crate::Result;
use crate::parts::{PresentationPart, SlideReference};
use crate::slide::{Key, Slide, SlideLayout, SlideMaster};

use super::embedded;
use super::package;

/// Borrowed semantic view of one `PresentationML` package graph.
pub struct Presentation<'a> {
    pub(super) package: &'a OpcPackage,
    pub(super) part: PresentationPart<'a>,
    /// Memo of [`PresentationPart::slide_references`]: the ordered catalog as
    /// parsed from the main part, before any package-graph validation.
    ///
    /// The package and the main part are borrowed for `'a`, so neither the
    /// part's bytes nor the graph they name can change while this view lives.
    /// Filling the memo therefore preserves every value and every refusal: a
    /// later call cannot observe a different catalog than the first call did.
    catalog: OnceLock<Vec<SlideReference>>,
    /// Set once [`super::package::validate_slide_catalog`] has accepted
    /// `catalog`. Only success is memoized; a refusal is recomputed, so a
    /// failing catalog returns the same typed error every time.
    catalog_validated: OnceLock<()>,
}

impl<'a> Presentation<'a> {
    /// Construct a view from a validated main part and its package.
    #[must_use]
    pub fn new(part: PresentationPart<'a>, package: &'a OpcPackage) -> Self {
        Self {
            package,
            part,
            catalog: OnceLock::new(),
            catalog_validated: OnceLock::new(),
        }
    }

    /// The ordered slide catalog parsed from the immutable main part.
    ///
    /// This is the borrowed form of [`PresentationPart::slide_references`],
    /// parsed at most once per view. Callers that additionally need the
    /// package-graph checks use [`Self::validated_catalog`].
    ///
    /// # Errors
    ///
    /// Returns the same error [`PresentationPart::slide_references`] returns.
    pub(super) fn catalog(&self) -> Result<&[SlideReference]> {
        if let Some(catalog) = self.catalog.get() {
            return Ok(catalog);
        }
        let parsed = self.part.slide_references()?;
        Ok(self.catalog.get_or_init(|| parsed))
    }

    /// The ordered slide catalog after the package-graph validation
    /// [`Self::slide_references`] and [`Self::slide_count`] have always run.
    ///
    /// # Errors
    ///
    /// Returns an error if the catalog cannot be parsed or a slide
    /// relationship is missing, external, mistyped, or resolves to a part
    /// with the wrong content type.
    pub(super) fn validated_catalog(&self) -> Result<&[SlideReference]> {
        let catalog = self.catalog()?;
        if self.catalog_validated.get().is_none() {
            package::validate_slide_catalog(self.package, &self.part, catalog)?;
            let _unset = self.catalog_validated.set(());
        }
        Ok(catalog)
    }

    /// The underlying OPC package.
    #[inline]
    #[must_use]
    pub fn package(&self) -> &'a OpcPackage {
        self.package
    }

    /// The borrowed main-document part view.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &PresentationPart<'a> {
        &self.part
    }

    /// The ordered, low-level slide references.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_references(&self) -> Result<Vec<SlideReference>> {
        Ok(self.validated_catalog()?.to_vec())
    }

    /// Number of slides in the ordered presentation graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_count(&self) -> Result<usize> {
        Ok(self.validated_catalog()?.len())
    }

    /// Presentation slide size in EMUs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        package::slide_size(&self.part)
    }

    /// Load the optional `DrawingML` table-style catalog owned by this
    /// presentation. The catalog remains a detached value so callers can
    /// inspect it without holding a mutable package borrow.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn styles(&self) -> Result<Option<crate::table::style::List>> {
        package::styles(self.package)
    }

    /// Load the complete inert speaker-notes graph, when present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn notes(&self) -> Result<Option<crate::notes::Graph>> {
        package::notes(self.package, &self.part)
    }

    /// Discover the optional opaque VBA project relationship owned by this
    /// presentation. The binary payload is never decoded or executed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn vba(&self) -> Result<Option<embedded::vba::Project>> {
        package::vba(self.package, &self.part)
    }

    /// Load all inert, opaque `PresentationML` content parts in slide order.
    ///
    /// The content-part anchor, relationship metadata, and referenced
    /// payload bytes are retained without interpreting or executing the
    /// external vocabulary stored by the producer.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn content_parts(&self) -> Result<Vec<embedded::content_parts::ContentPart>> {
        package::content_parts(self)
    }

    /// Discover inert hyperlinks owned by the presentation's slides.
    ///
    /// Each result contains the zero-based slide position and a typed target.
    /// Relationship targets and inline actions are parsed as values only;
    /// they are never followed, opened, or executed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn hyperlinks(&self) -> Result<Vec<(usize, crate::hyperlinks::Hyperlink)>> {
        package::hyperlinks(self)
    }

    /// Resolve one ordered slide by zero-based index.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide(&self, index: usize) -> Result<Option<Slide<'a>>> {
        package::slide(self, index)
    }

    /// Resolve a slide by checked index or exact producer-visible name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn find_slide<'k>(&self, key: impl Into<Key<'k>>) -> Result<Option<Slide<'a>>> {
        package::find_slide(self, key.into())
    }

    /// Resolve all slides in presentation order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slides(&self) -> Result<Vec<Slide<'a>>> {
        package::slides(self)
    }

    /// Resolve the slide masters declared by `p:sldMasterIdLst` in XML order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_masters(&self) -> Result<Vec<SlideMaster<'a>>> {
        package::slide_masters(self.package, &self.part)
    }

    /// Resolve all layouts reachable from all presentation masters.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_layouts(&self) -> Result<Vec<SlideLayout<'a>>> {
        package::slide_layouts(self.package, &self.part)
    }

    /// Flatten all slide text in presentation order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(&self) -> Result<String> {
        package::text(self)
    }

    /// Stream one semantic text object per slide into a caller-owned sink.
    ///
    /// The complete slide relationship graph is validated before the first
    /// sink byte is written. Each selected slide is then parsed independently,
    /// retaining only one bounded slide text value at a time. The bounded
    /// relationship-metadata preflight is not slide-payload aggregation. For
    /// parity with [`Self::text`], use `"\n"` for both separators, use
    /// `"\n"` as the paragraph separator, and exclude empty objects.
    ///
    /// # Errors
    ///
    /// Returns a typed document, resource-limit, or sink error with exact
    /// partial-output progress.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<crate::Error>> {
        package::write_text_to(self, output, options)
    }

    /// Load the slide-library synchronization metadata reachable from this
    /// presentation's slide graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_sync(
        &self,
    ) -> Result<Vec<crate::presentation_properties::metadata::slide_sync::Part>> {
        package::slide_sync(self.package)
    }
}
