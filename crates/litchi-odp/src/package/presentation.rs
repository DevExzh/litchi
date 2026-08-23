//! Main Presentation structure and implementation.

use crate::codec::Parser;
use crate::core::{OwnedPackage, family::Package};
use crate::model::{Reference, Settings, Slide, declaration, page_layout, page_metadata, settings};
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::{constants::ODF_PRESENTATION, core::PreparedPackage};
use std::path::Path;

const BODY_MARKER: &str = "<office:presentation";

/// An `OpenDocument` presentation (.odp).
///
/// This struct represents a complete ODP presentation and provides methods to access
/// its slides and metadata.
///
/// # Examples
///
/// ```no_run
/// use litchi_odp::Presentation;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut presentation = Presentation::open("slides.odp")?;
///
/// // Get slide count
/// println!("Slides: {}", presentation.slide_count()?);
///
/// // Access slides
/// let slides = presentation.slides()?;
/// for slide in slides {
///     println!("Slide {}: {}", slide.index() + 1, slide.text()?);
/// }
/// # Ok(())
/// # }
/// ```
pub struct Presentation {
    package: Package,
}

impl Presentation {
    /// Open an ODP presentation from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .odp file
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid ODP file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::from_bytes(std::fs::read(path.as_ref())?)
    }

    /// Open a password-encrypted ODP presentation.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn open_with_password<P: AsRef<Path>>(
        path: P,
        password: impl Into<String>,
    ) -> Result<Self> {
        Package::open_with_password(path, password, ODF_PRESENTATION, BODY_MARKER, "ODP")
            .map(|package| Self { package })
    }

    /// Create a Presentation from a byte buffer.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Complete ODP file contents as bytes
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes do not represent a valid ODP file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let bytes = std::fs::read("slides.odp")?;
    /// let presentation = Presentation::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        // Let the prepared detector perform the local MIME probe exactly
        // once.  A matching ODP result transfers its indexed archive; a
        // different ODF family still follows the historical package owner so
        // its MIME error is reported at the same boundary; rejected probes
        // recover the original allocation for the ordinary parser.
        match litchi_odf_common::detect::prepared_or_original(bytes) {
            Ok(prepared) if prepared.format() == litchi_core::detection::FileFormat::Odp => {
                Self::from_prepared_package(prepared)
            },
            Ok(prepared) => Self::from_owned_package(prepared.into_package()),
            Err(bytes) => Package::from_bytes(bytes, ODF_PRESENTATION, BODY_MARKER, "ODP")
                .map(|package| Self { package }),
        }
    }

    /// Adopt the indexed package retained by smart ODF detection.
    ///
    /// The concrete ODP MIME and body contracts remain checked at this
    /// boundary while the detector-owned ZIP index is transferred unchanged.
    pub fn from_prepared_package(prepared: PreparedPackage) -> Result<Self> {
        if prepared.format() != litchi_core::detection::FileFormat::Odp {
            return Err(Error::InvalidFormat(
                "prepared ODF package is not an ODP family document".to_string(),
            ));
        }
        Self::from_owned_package(prepared.into_package())
    }

    /// Alias for [`Self::from_prepared_package`].
    #[inline]
    pub fn from_prepared(prepared: PreparedPackage) -> Result<Self> {
        Self::from_prepared_package(prepared)
    }

    /// Adopt an already validated archive without copying its package bytes.
    pub(crate) fn from_owned_package(package: OwnedPackage) -> Result<Self> {
        Package::from_owned_package(package, ODF_PRESENTATION, BODY_MARKER, "ODP")
            .map(|validated| Self { package: validated })
    }

    /// Return the identity of the archive index retained by smart detection.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.package.package().prepared_index_identity()
    }

    /// Create a presentation from password-encrypted ODP bytes.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Package::from_bytes_with_password(bytes, password, ODF_PRESENTATION, BODY_MARKER, "ODP")
            .map(|package| Self { package })
    }

    /// Create an ODP presentation from raw bytes (ZIP archive data).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been validated during format detection. It avoids double-parsing.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn from_archive_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Borrow the validated `content.xml` snapshot without reparsing it.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Discover embedded charts in presentation drawing order.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn charts(&self) -> Result<crate::charts::Inventory<'_>> {
        self.charts_with(crate::charts::Limits::default())
    }

    /// Discover embedded charts with an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn charts_with(
        &self,
        limits: crate::charts::Limits,
    ) -> Result<crate::charts::Inventory<'_>> {
        crate::charts::inventory(&self.package, limits)
    }

    /// Select one embedded chart by exact frame name or checked position.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn chart<'a, S>(&self, selector: S) -> Result<Option<crate::charts::Chart>>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        self.charts()?
            .get(selector)
            .map(Option::<&crate::charts::Chart>::cloned)
    }

    /// Create an immutable, exact-source embedded-chart snapshot.
    ///
    /// # Errors
    /// Returns an error when the package or chart inventory is malformed or exceeds its limits.
    pub fn chart_snapshot(&self) -> Result<crate::charts::Snapshot> {
        self.chart_snapshot_with(crate::charts::Limits::default())
    }

    /// Create an immutable embedded-chart snapshot under an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the package or chart inventory is malformed or exceeds its limits.
    pub fn chart_snapshot_with(
        &self,
        limits: crate::charts::Limits,
    ) -> Result<crate::charts::Snapshot> {
        crate::charts::Snapshot::from_owned_package(self.package.package().clone(), limits)
    }

    /// Borrow the optional validated `styles.xml` snapshot without reparsing it.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Borrow the validated archive for package-level snapshot edits within
    /// the crate's semantic owner modules.
    pub(crate) fn owned_package(&self) -> &OwnedPackage {
        self.package.package()
    }

    /// Get the number of slides in the presentation.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn slide_count(&self) -> Result<usize> {
        let slides = self.slides()?;
        Ok(slides.len())
    }

    /// Get all slides in the presentation.
    ///
    /// Returns a vector of `Slide` objects representing all slides in the document.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn slides(&self) -> Result<Vec<Slide>> {
        Parser::parse_slides_with_styles(self.package.content_xml(), self.package.styles_xml())
    }

    /// Inspect inert slide-show settings and ordered custom shows.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn settings(&self) -> Result<Option<Settings>> {
        settings::parse(self.package.content_xml())
    }

    /// Inspect inert header, footer, date-time, and page-binding declarations.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn declarations(&self) -> Result<declaration::Collection> {
        declaration::parse(self.package.content_xml())
    }

    /// Inspect static page names, IDs, and layout/master references.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn pages(&self) -> Result<page_metadata::Collection> {
        page_metadata::parse(self.package.content_xml())
    }

    /// Inspect slide- and shape-anchored ODF annotations in document order.
    ///
    /// Rich annotation bodies are shared with the other ODF family crates;
    /// this facade adds only the presentation-specific page/shape anchor.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn annotations(&self) -> Result<Vec<crate::annotation::Info>> {
        crate::annotation::annotations(self.package.content_xml())
    }

    /// Find a uniquely named slide or shape annotation.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn find_annotation(&self, name: &str) -> Result<Option<crate::annotation::Info>> {
        crate::annotation::find(self.package.content_xml(), name)
    }

    /// Add an annotation to a page or uniquely named shape atomically.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn add_annotation(
        &mut self,
        anchor: &crate::annotation::Anchor,
        annotation: &crate::annotation::Annotation,
    ) -> Result<usize> {
        let (bytes, index) = crate::annotation::add(
            self.package.package(),
            self.package.content_xml(),
            anchor,
            annotation,
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(index)
    }

    /// Replace one annotation body while retaining its existing anchor.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn replace_annotation(
        &mut self,
        index: usize,
        annotation: &crate::annotation::Annotation,
    ) -> Result<()> {
        let bytes = crate::annotation::replace(
            self.package.package(),
            self.package.content_xml(),
            index,
            annotation,
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Remove one annotation atomically.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn remove_annotation(&mut self, index: usize) -> Result<()> {
        let bytes =
            crate::annotation::remove(self.package.package(), self.package.content_xml(), index)?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Inspect named presentation page layouts and their typed placeholders.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn layouts(&self) -> Result<page_layout::Collection> {
        match self.package.styles_xml() {
            Some(styles) => page_layout::parse(styles),
            None => Ok(page_layout::Collection::default()),
        }
    }

    /// Inspect named drawing fill-image resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_fill_images(
        &self,
    ) -> Result<litchi_odf_common::drawing::resources::fill_image::Collection> {
        self.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::fill_image::Collection::default()),
            litchi_odf_common::drawing::resources::fill_image::parse_drawing_fill_images,
        )
    }

    /// Inspect named legacy and SVG drawing gradients without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_gradients(
        &self,
    ) -> Result<litchi_odf_common::drawing::resources::gradient::Collection> {
        self.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::gradient::Collection::default()),
            litchi_odf_common::drawing::resources::gradient::parse_drawing_gradients,
        )
    }

    /// Inspect named drawing hatch resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_hatches(
        &self,
    ) -> Result<litchi_odf_common::drawing::resources::hatch::Collection> {
        self.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::hatch::Collection::default()),
            litchi_odf_common::drawing::resources::hatch::parse_drawing_hatches,
        )
    }

    /// Inspect named drawing marker resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_markers(
        &self,
    ) -> Result<litchi_odf_common::drawing::resources::marker::Collection> {
        self.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::marker::Collection::default()),
            litchi_odf_common::drawing::resources::marker::parse_drawing_markers,
        )
    }

    /// Inspect named drawing opacity resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_opacities(
        &self,
    ) -> Result<litchi_odf_common::drawing::resources::opacity::Collection> {
        self.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::opacity::Collection::default()),
            litchi_odf_common::drawing::resources::opacity::parse_drawing_opacities,
        )
    }

    /// Inspect named drawing stroke-dash resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<litchi_odf_common::drawing::resources::stroke_dash::Collection> {
        self.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::stroke_dash::Collection::default()),
            litchi_odf_common::drawing::resources::stroke_dash::parse_drawing_stroke_dashes,
        )
    }

    /// Get a slide by index.
    ///
    /// Returns `Some(slide)` if a slide exists at the given index, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `index` - 0-based index of the slide
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn slide(&self, index: usize) -> Result<Option<Slide>> {
        Parser::parse_slide_with_styles_at(
            self.package.content_xml(),
            self.package.styles_xml(),
            index,
        )
    }

    /// Read a package-contained media payload without fetching external URLs.
    ///
    /// Returns `None` for external links, fragment links, unsafe paths, and
    /// package-relative references whose payload is absent.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn media_data(&self, media: &Reference) -> Result<Option<Vec<u8>>> {
        let Some(path) = media.package_path() else {
            return Ok(None);
        };
        let package = self.package.package().package()?;
        if !package.has_file(path) {
            return Ok(None);
        }
        package.get_file(path).map(Some)
    }

    /// Discover inert embedded objects in content and styles document order.
    ///
    /// This includes regular objects, OLE payloads, applets, plugins, and
    /// floating frames. Linked targets are classified but never fetched;
    /// applets and plugins are never loaded or executed.
    ///
    /// # Errors
    /// Returns an error when XML, package paths, inline payloads, or the bounded inventory is
    /// malformed or exceeds its configured safety ceilings.
    pub fn embedded_objects(&self) -> Result<Vec<crate::embedded::Object>> {
        let package = self.package.package().package()?;
        litchi_odf_common::embedded::scan_package(
            self.package.content_xml(),
            self.package.styles_xml(),
            &package,
        )
    }

    /// Extract all text content from the presentation.
    ///
    /// Returns text from all slides, separated by double newlines.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn text(&self) -> Result<String> {
        let slides = self.slides()?;
        let mut all_text = Vec::new();

        for slide in slides {
            let text = slide.all_text();
            if !text.is_empty() {
                all_text.push(text);
            }
        }

        Ok(all_text.join("\n\n"))
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn metadata(&self) -> Result<Metadata> {
        Ok(self.package.metadata().cloned().unwrap_or_default())
    }

    /// Read all inert RDF metadata graphs in package order.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn rdf_graphs(&self) -> Result<Vec<crate::rdf::Graph>> {
        crate::rdf::graphs(self.package.package())
    }

    /// Add a graph and atomically replace this snapshot with the rebuilt package.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let (bytes, path) = crate::rdf::add_graph(self.package.package(), preferred_path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(path)
    }

    /// Replace one complete RDF graph and atomically publish the result.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let bytes = crate::rdf::replace_graph(self.package.package(), path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one RDF graph after validating that no remaining graph references it.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = crate::rdf::remove_graph(self.package.package(), path)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Append one triple to an existing graph and return its committed index.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<usize> {
        let index = self
            .rdf_graphs()?
            .into_iter()
            .find(|graph| graph.path == path)
            .ok_or_else(|| Error::InvalidFormat(format!("RDF graph '{path}' was not found")))?
            .triples
            .len();
        let (bytes, _) = crate::rdf::add_triple(self.package.package(), path, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    /// Replace one triple while preserving its description subject.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<()> {
        let bytes = crate::rdf::replace_triple(self.package.package(), path, index, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one triple from a graph.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = crate::rdf::remove_triple(self.package.package(), path, index)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Move one triple within its RDF description.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = crate::rdf::move_triple(self.package.package(), path, from, to)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Get the complete format-specific `OpenDocument` metadata model.
    /// Create an immutable, source-bound editing snapshot.
    ///
    /// Transactions never mutate this presentation. Publication returns a new
    /// snapshot and an exact-source-checked reversible patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the retained package exceeds editing limits or cannot be reparsed.
    pub fn snapshot(&self) -> Result<crate::edit::Snapshot> {
        crate::edit::Snapshot::from_owned_package(self.package.package().clone())
    }

    /// Save the presentation to a new file.
    ///
    /// This method saves the current presentation state to a new file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("input.odp")?;
    /// presentation.save("output.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Note
    ///
    /// Use [`Self::snapshot`] for source-checked edits, or [`crate::Builder`]
    /// for detached construction.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the presentation to bytes.
    ///
    /// This method serializes the presentation to an ODF-compliant ZIP archive.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// let bytes = presentation.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.package.as_bytes().to_vec())
    }
}
