//! Main Presentation structure and implementation.

use crate::codec::Parser;
use crate::core::{OwnedPackage, family::Package};
use crate::model::{Reference, Settings, Slide, declaration, page_layout, page_metadata, settings};
use litchi_core::{
    Error, Metadata, Result as CoreResult, SequentialTextWriter, TextOutputError,
    TextOutputOptions, TextOutputReport,
};
use litchi_odf_common::{XmlStreamLimits, constants::ODF_PRESENTATION, core::PreparedPackage};
use std::io::Write;
use std::sync::{Arc, OnceLock};

const BODY_MARKER: &str = "<office:presentation";

type FamilyResult<T> = Result<T, litchi_odf_common::Error>;

/// The read-only semantic projection of an `OpenDocument` presentation.
///
/// A `Presentation` is obtained from [`crate::Package::presentation`]; it is
/// owned, immutable and cheap to clone (one reference-count increment). Every
/// state verb (open, edit, apply, save) lives on [`crate::Package`].
///
/// # Examples
///
/// ```no_run
/// use litchi_odp::Package;
///
/// # fn main() -> std::result::Result<(), litchi_odp::Error> {
/// let presentation = Package::open("slides.odp")?.presentation()?;
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
#[derive(Clone)]
pub struct Presentation {
    inner: Arc<State>,
}

struct State {
    package: Package,
    /// Package repairs followed by the slide scan's content repairs,
    /// computed by the first [`Presentation::tolerance_report`] or
    /// [`Presentation::slides`].
    tolerance: OnceLock<litchi_odf_common::leniency::ToleranceReport>,
}

impl std::fmt::Debug for Presentation {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Presentation")
            .field("package_bytes", &self.inner.package.as_bytes().len())
            .finish_non_exhaustive()
    }
}

impl Presentation {
    fn with_package(package: Package) -> Self {
        Self {
            inner: Arc::new(State {
                package,
                tolerance: OnceLock::new(),
            }),
        }
    }

    /// Publish the merged report from one complete slide scan.
    fn publish_tolerance(&self, content: &litchi_odf_common::leniency::ToleranceReport) {
        if self.inner.tolerance.get().is_none() {
            let mut merged = self.inner.package.tolerance_report().clone();
            merged.append(content);
            let _ = self.inner.tolerance.set(merged);
        }
    }

    /// Reparse complete ODP bytes produced inside the crate.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> CoreResult<Self> {
        // Let the prepared detector perform the local MIME probe exactly
        // once.  A matching ODP result transfers its indexed archive; a
        // different ODF family still follows the historical package owner so
        // its MIME error is reported at the same boundary; rejected probes
        // recover the original allocation for the ordinary parser.
        match litchi_odf_common::detect::prepared_or_original(
            bytes,
            litchi_odf_common::SourceLimits::default(),
        ) {
            Ok(prepared) if prepared.format() == litchi_core::detection::FileFormat::Odp => {
                Self::from_prepared_package(prepared)
            },
            Ok(prepared) => Self::from_owned_package(prepared.into_package()),
            Err(bytes) => Package::from_bytes(bytes, ODF_PRESENTATION, BODY_MARKER, "ODP")
                .map(Self::with_package),
        }
    }

    /// Adopt the indexed package retained by smart ODF detection.
    ///
    /// The concrete ODP MIME and body contracts remain checked at this
    /// boundary while the detector-owned ZIP index is transferred unchanged.
    pub(crate) fn from_prepared_package(prepared: PreparedPackage) -> CoreResult<Self> {
        if prepared.format() != litchi_core::detection::FileFormat::Odp {
            return Err(Error::UnexpectedFormat {
                detected: prepared.format(),
            });
        }
        Self::from_owned_package(prepared.into_package())
    }

    /// Adopt an already validated archive without copying its package bytes.
    pub(crate) fn from_owned_package(package: OwnedPackage) -> CoreResult<Self> {
        Package::from_owned_package(package, ODF_PRESENTATION, BODY_MARKER, "ODP")
            .map(Self::with_package)
    }

    /// Copy the exact package bytes (edit pipeline only).
    pub(crate) fn to_bytes(&self) -> CoreResult<Vec<u8>> {
        Ok(self.inner.package.as_bytes().to_vec())
    }

    /// Return the identity of the archive index retained by smart detection.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.inner.package.package().prepared_index_identity()
    }

    /// Borrow the validated `content.xml` snapshot without reparsing it.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.inner.package.content_xml()
    }

    /// Everything a tolerant read repaired while this presentation was opened.
    ///
    /// An empty report means the document conformed; see
    /// [`litchi_odf_common::leniency::ToleranceReport`]. Package repairs come
    /// first, then the content repairs of the slide scan (LENIENCY_AUDIT
    /// F9.10). Slides are parsed lazily, so the first call to this method
    /// (or to [`Self::slides`]) runs that scan once.
    #[inline]
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        self.inner.tolerance.get_or_init(|| {
            let mut merged = self.inner.package.tolerance_report().clone();
            // Content repairs (LENIENCY_AUDIT F9.10) are found by the slide
            // scan. A scan that fails outright reports its error through
            // `slides()`; the report then lists the package repairs only.
            if let Ok((_slides, content)) = Parser::parse_slides_with_report(
                self.inner.package.content_xml(),
                self.inner.package.styles_xml(),
            ) {
                merged.append(&content);
            }
            merged
        })
    }

    /// Discover embedded charts in presentation drawing order.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn charts(&self) -> FamilyResult<crate::charts::Inventory<'_>> {
        self.charts_with(crate::charts::Limits::default())
    }

    /// Discover embedded charts with an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn charts_with(
        &self,
        limits: crate::charts::Limits,
    ) -> FamilyResult<crate::charts::Inventory<'_>> {
        Ok(crate::charts::inventory(&self.inner.package, limits)?)
    }

    /// Select one embedded chart by exact frame name or checked position.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn chart<'a, S>(&self, selector: S) -> FamilyResult<Option<crate::charts::Chart>>
    where
        S: Into<crate::charts::Selector<'a>>,
    {
        Ok(self
            .charts()?
            .get(selector)
            .map(Option::<&crate::charts::Chart>::cloned)?)
    }

    /// Create an immutable, exact-source embedded-chart snapshot.
    ///
    /// # Errors
    /// Returns an error when the package or chart inventory is malformed or exceeds its limits.
    pub fn chart_snapshot(&self) -> FamilyResult<crate::charts::Snapshot> {
        self.chart_snapshot_with(crate::charts::Limits::default())
    }

    /// Create an immutable embedded-chart snapshot under an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the package or chart inventory is malformed or exceeds its limits.
    pub fn chart_snapshot_with(
        &self,
        limits: crate::charts::Limits,
    ) -> FamilyResult<crate::charts::Snapshot> {
        Ok(crate::charts::Snapshot::from_owned_package(
            self.inner.package.package().clone(),
            limits,
        )?)
    }

    /// Borrow the optional validated `styles.xml` snapshot without reparsing it.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.inner.package.styles_xml()
    }

    /// Borrow the validated archive for package-level snapshot edits within
    /// the crate's semantic owner modules.
    pub(crate) fn owned_package(&self) -> &OwnedPackage {
        self.inner.package.package()
    }

    /// Get the number of slides in the presentation.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn slide_count(&self) -> FamilyResult<usize> {
        let slides = self.slides()?;
        Ok(slides.len())
    }

    /// Get all slides in the presentation.
    ///
    /// Returns a vector of `Slide` objects representing all slides in the document.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn slides(&self) -> FamilyResult<Vec<Slide>> {
        let (slides, content) = Parser::parse_slides_with_report(
            self.inner.package.content_xml(),
            self.inner.package.styles_xml(),
        )?;
        self.publish_tolerance(&content);
        Ok(slides)
    }

    /// Inspect inert slide-show settings and ordered custom shows.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn settings(&self) -> FamilyResult<Option<Settings>> {
        Ok(settings::parse(self.inner.package.content_xml())?)
    }

    /// Read the auxiliary projections and retained source skeleton required to
    /// stage an edit in one traversal.
    pub(crate) fn staging(&self) -> CoreResult<crate::model::staging::Staged> {
        crate::model::staging::parse_with_content_source(self.inner.package.content_xml())
    }

    /// Inspect inert header, footer, date-time, and page-binding declarations.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn declarations(&self) -> FamilyResult<declaration::Collection> {
        Ok(declaration::parse(self.inner.package.content_xml())?)
    }

    /// Inspect all typed `draw:layer-set` declarations lazily.
    ///
    /// The inventory is inert: layer visibility and protection are metadata
    /// values and no rendering or access-control behavior is executed.
    pub fn layers(&self) -> FamilyResult<crate::model::LayerInventory> {
        Ok(crate::model::layer::inventory(
            self.inner.package.content_xml(),
            self.inner.package.styles_xml(),
        )?)
    }

    /// Inspect static page names, IDs, and layout/master references.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn pages(&self) -> FamilyResult<page_metadata::Collection> {
        Ok(page_metadata::parse(self.inner.package.content_xml())?)
    }

    /// Inspect slide- and shape-anchored ODF annotations in document order.
    ///
    /// Rich annotation bodies are shared with the other ODF family crates;
    /// this facade adds only the presentation-specific page/shape anchor.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn annotations(&self) -> FamilyResult<Vec<crate::annotation::Info>> {
        Ok(crate::annotation::annotations(
            self.inner.package.content_xml(),
            &self.inner.package.xml_limits_value(),
        )?)
    }

    /// Find a uniquely named slide or shape annotation.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn find_annotation(&self, name: &str) -> FamilyResult<Option<crate::annotation::Info>> {
        Ok(crate::annotation::find(
            self.inner.package.content_xml(),
            name,
            &self.inner.package.xml_limits_value(),
        )?)
    }

    /// Add an annotation to a page or uniquely named shape atomically.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub(crate) fn add_annotation(
        &mut self,
        anchor: &crate::annotation::Anchor,
        annotation: &crate::annotation::Annotation,
    ) -> CoreResult<usize> {
        let (bytes, index) = crate::annotation::add(
            self.inner.package.package(),
            self.inner.package.content_xml(),
            anchor,
            annotation,
            &self.inner.package.xml_limits_value(),
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(index)
    }

    /// Replace one annotation body while retaining its existing anchor.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub(crate) fn replace_annotation(
        &mut self,
        index: usize,
        annotation: &crate::annotation::Annotation,
    ) -> CoreResult<()> {
        let bytes = crate::annotation::replace(
            self.inner.package.package(),
            self.inner.package.content_xml(),
            index,
            annotation,
            &self.inner.package.xml_limits_value(),
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Remove one annotation atomically.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub(crate) fn remove_annotation(&mut self, index: usize) -> CoreResult<()> {
        let bytes = crate::annotation::remove(
            self.inner.package.package(),
            self.inner.package.content_xml(),
            index,
            &self.inner.package.xml_limits_value(),
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Inspect named presentation page layouts and their typed placeholders.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn layouts(&self) -> FamilyResult<page_layout::Collection> {
        match self.inner.package.styles_xml() {
            Some(styles) => Ok(page_layout::parse(styles)?),
            None => Ok(page_layout::Collection::default()),
        }
    }

    /// Inspect named drawing fill-image resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_fill_images(
        &self,
    ) -> FamilyResult<litchi_odf_common::drawing::resources::fill_image::Collection> {
        self.inner.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::fill_image::Collection::default()),
            |styles| {
                litchi_odf_common::drawing::resources::fill_image::parse_drawing_fill_images(
                    styles,
                    XmlStreamLimits::owner_default_value(),
                )
            },
        )
    }

    /// Inspect named legacy and SVG drawing gradients without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_gradients(
        &self,
    ) -> FamilyResult<litchi_odf_common::drawing::resources::gradient::Collection> {
        self.inner.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::gradient::Collection::default()),
            |styles| {
                litchi_odf_common::drawing::resources::gradient::parse_drawing_gradients(
                    styles,
                    XmlStreamLimits::owner_default_value(),
                )
            },
        )
    }

    /// Inspect named drawing hatch resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_hatches(
        &self,
    ) -> FamilyResult<litchi_odf_common::drawing::resources::hatch::Collection> {
        self.inner.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::hatch::Collection::default()),
            |styles| {
                litchi_odf_common::drawing::resources::hatch::parse_drawing_hatches(
                    styles,
                    XmlStreamLimits::owner_default_value(),
                )
            },
        )
    }

    /// Inspect named drawing marker resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_markers(
        &self,
    ) -> FamilyResult<litchi_odf_common::drawing::resources::marker::Collection> {
        self.inner.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::marker::Collection::default()),
            |styles| {
                litchi_odf_common::drawing::resources::marker::parse_drawing_markers(
                    styles,
                    XmlStreamLimits::owner_default_value(),
                )
            },
        )
    }

    /// Inspect named drawing opacity resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_opacities(
        &self,
    ) -> FamilyResult<litchi_odf_common::drawing::resources::opacity::Collection> {
        self.inner.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::opacity::Collection::default()),
            |styles| {
                litchi_odf_common::drawing::resources::opacity::parse_drawing_opacities(
                    styles,
                    XmlStreamLimits::owner_default_value(),
                )
            },
        )
    }

    /// Inspect named drawing stroke-dash resources without resolving style use sites.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> FamilyResult<litchi_odf_common::drawing::resources::stroke_dash::Collection> {
        self.inner.package.styles_xml().map_or_else(
            || Ok(litchi_odf_common::drawing::resources::stroke_dash::Collection::default()),
            |styles| {
                litchi_odf_common::drawing::resources::stroke_dash::parse_drawing_stroke_dashes(
                    styles,
                    XmlStreamLimits::owner_default_value(),
                )
            },
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
    pub fn slide(&self, index: usize) -> FamilyResult<Option<Slide>> {
        Ok(Parser::parse_slide_with_styles_at(
            self.inner.package.content_xml(),
            self.inner.package.styles_xml(),
            index,
        )?)
    }

    /// Read a package-contained media payload without fetching external URLs.
    ///
    /// Returns `None` for external links, fragment links, unsafe paths, and
    /// package-relative references whose payload is absent.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn media_data(&self, media: &Reference) -> FamilyResult<Option<Vec<u8>>> {
        let Some(path) = media.package_path() else {
            return Ok(None);
        };
        let package = self.inner.package.package().package()?;
        if !package.has_file(path) {
            return Ok(None);
        }
        Ok(package.get_file(path).map(Some)?)
    }

    /// Read one package member (an embedded object's part), or `None` when
    /// the package has no member at `path` (record 0328).
    ///
    /// # Errors
    /// Returns an error when the package or the member cannot be read.
    pub fn member_data(&self, path: &str) -> FamilyResult<Option<Vec<u8>>> {
        let package = self.inner.package.package().package()?;
        if !package.has_file(path) {
            return Ok(None);
        }
        Ok(package.get_file(path).map(Some)?)
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
    pub fn embedded_objects(&self) -> FamilyResult<Vec<crate::embedded::Object>> {
        let package = self.inner.package.package().package()?;
        Ok(litchi_odf_common::embedded::scan_package(
            self.inner.package.content_xml(),
            self.inner.package.styles_xml(),
            &package,
        )?)
    }

    /// Extract all text content from the presentation.
    ///
    /// Returns text from all slides, separated by double newlines.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn text(&self) -> FamilyResult<String> {
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

    /// Stream one semantic text object per slide into `output` in document order.
    ///
    /// The output may be partial when parsing, limits, or the sink fail. This
    /// method does not flush or roll back the sink and provides no cancellation
    /// hook. The paragraph separator is used for every retained paragraph,
    /// title/body, and shape-text boundary within a slide; the slide separator
    /// is used between slide objects. Use [`TextOutputOptions::paragraph_separator`]
    /// and [`TextOutputOptions::slide_separator`] to select them.
    /// Semantic output parity is defined for documents accepted by the ordinary
    /// parser; the fused sink validates only its bounded XML/text contexts and
    /// does not promise full parser validation or error-precedence parity. Use
    /// `with_empty_objects(false)` to match [`Self::text`] filtering of empty
    /// slides.
    ///
    /// # Errors
    /// Returns a typed error containing exact partial-output progress.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> Result<TextOutputReport, TextOutputError<Error>> {
        let paragraph_separator = options.paragraph_separator();
        let mut writer = SequentialTextWriter::new(output, options);
        Parser::write_slides_text_to(
            self.inner.package.content_xml(),
            &mut writer,
            paragraph_separator,
        )?;
        Ok(writer.finish())
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn metadata(&self) -> FamilyResult<Metadata> {
        Ok(self.inner.package.metadata().cloned().unwrap_or_default())
    }

    /// Read all inert RDF metadata graphs in package order.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn rdf_graphs(&self) -> FamilyResult<Vec<crate::rdf::Graph>> {
        Ok(crate::rdf::graphs(self.inner.package.package())?)
    }
}

#[cfg(test)]
mod tests {
    use super::Presentation;
    use crate::package::SourceBackedPresentation;
    use std::path::Path;
    use std::sync::Arc;

    /// A conforming ODP opens through every facade with a present but empty
    /// tolerance report: nothing was tolerated, and nothing was recorded.
    #[test]
    fn conformant_fixture_opens_with_an_empty_tolerance_report() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/odf/odp/cellspan.odp");
        let bytes = std::fs::read(&path).expect("test fixture must be readable");

        // Owned open and the prepared-detection handoff.
        let presentation = Presentation::from_bytes(bytes.clone()).expect("fixture must open");
        assert!(presentation.tolerance_report().is_clean());
        assert_eq!(presentation.tolerance_report().total(), 0);
        let prepared =
            litchi_odf_common::detect::prepared(bytes, litchi_odf_common::SourceLimits::default())
                .expect("fixture must be detected");
        let presentation =
            Presentation::from_prepared_package(prepared).expect("fixture must open");
        assert!(presentation.tolerance_report().is_clean());

        // Path-based open and the positional source-backed facade.
        let presentation = crate::Package::open(&path)
            .expect("fixture must open")
            .presentation()
            .expect("fixture must open");
        assert!(presentation.tolerance_report().is_clean());
        let source = Arc::new(litchi_core::OwnedSource::new(
            std::fs::read(&path).expect("test fixture must be readable"),
        ));
        let source_backed =
            SourceBackedPresentation::from_read_at(source).expect("fixture must open");
        assert!(source_backed.tolerance_report().is_clean());
    }
}
