use crate::error::{OoxmlError, Result};
use crate::pptx::actions::{ActionLoadLimits, PptxActionSetting, load_slide_action_settings};
use crate::pptx::controls::{ControlLoadLimits, PptxSlideControl, load_slide_controls};
use crate::pptx::handout::HandoutMaster;
use crate::pptx::ink::{InkLoadLimits, PptxInkAnnotation, load_slide_ink_annotations};
use crate::pptx::laser::{LaserLoadLimits, PptxLaserTrace, load_slide_laser_traces};
use crate::pptx::namespace::is_presentationml_name;
use crate::pptx::ole::{OleLoadLimits, PptxOleObject, load_slide_ole_objects};
use crate::pptx::package::STALE_NOTES_REASON;
use crate::pptx::parts::{
    NotesSize, PresentationCustomerDataList, PresentationDefaultTextStyle,
    PresentationKinsokuSettings, PresentationMetadata, PresentationModificationVerifier,
    PresentationPart, PresentationPhotoAlbum, SlideMasterPart, SlidePart, SlideSize,
};
use crate::pptx::show_events::{PptxSlideShowEvent, ShowEventLoadLimits, load_slide_show_events};
use crate::pptx::slide::{Slide, SlideMaster};
use litchi_ooxml_common::ribbon;
use litchi_ooxml_common::web;
use litchi_ooxml_common::xml::{is_drawingml_name, unqualified_attribute_value};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::Part;
/// Main presentation object - the high-level API for working with presentations.
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

const STRICT_SLIDE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const STRICT_SLIDE_MASTER_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";
const STRICT_HANDOUT_MASTER_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/handoutMaster";

fn resolve_presentation_relationship<'a>(
    presentation_part: &dyn Part,
    package: &'a OpcPackage,
    relationship_id: &str,
    relationship_types: &[&str],
    relationship_label: &str,
    expected_content_type: &str,
) -> Result<&'a dyn Part> {
    let relationship = presentation_part
        .rels()
        .get(relationship_id)
        .ok_or_else(|| {
            OoxmlError::InvalidRelationship(format!(
                "presentation references missing {relationship_label} relationship '{relationship_id}'"
            ))
        })?;
    if !relationship_types
        .iter()
        .any(|relationship_type| relationship.reltype() == *relationship_type)
    {
        return Err(OoxmlError::InvalidRelationship(format!(
            "relationship '{relationship_id}' is not a {relationship_label} relationship"
        )));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidRelationship(format!(
            "{relationship_label} relationship '{relationship_id}' must be internal"
        )));
    }

    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!(
            "invalid {relationship_label} relationship '{relationship_id}': {error}"
        ))
    })?;
    let target = package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "{relationship_label} relationship '{relationship_id}' targets missing part '{}': {error}",
            part_name.as_str()
        ))
    })?;
    if target.content_type() != expected_content_type {
        return Err(OoxmlError::InvalidContentType {
            expected: expected_content_type.to_string(),
            got: target.content_type().to_string(),
        });
    }

    Ok(target)
}

fn validate_handout_master_root(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if is_presentationml_name(&namespace, element.name(), b"handoutMaster") {
                    return Ok(());
                }
                return Err(OoxmlError::InvalidFormat(
                    "handout-master part must have a PresentationML handoutMaster root".to_string(),
                ));
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "handout-master part is missing its PresentationML root".to_string(),
                ));
            },
            _ => {},
        }
    }
}

/// A chart part discovered on a presentation slide.
///
/// The chart XML is parsed as inert document data. Its relationship and part
/// identities are retained so callers can identify the source part precisely.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxChart {
    slide_index: usize,
    relationship_id: String,
    part_name: PackURI,
    info: crate::pptx::parts::ChartInfo,
}

impl PptxChart {
    /// Return the zero-based index of the slide that owns this chart.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the relationship ID from the owning slide to this chart part.
    #[inline]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the absolute OPC part name of this chart.
    #[inline]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    /// Return basic parsed chart metadata.
    #[inline]
    pub fn info(&self) -> &crate::pptx::parts::ChartInfo {
        &self.info
    }
}

/// A PowerPoint presentation.
///
/// This is the main high-level API for working with presentation content,
/// following the python-pptx interface design.
///
/// Not intended to be constructed directly. Use `Package::presentation()` to
/// access a presentation.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::pptx::Package;
///
/// let pkg = Package::open("presentation.pptx")?;
/// let pres = pkg.presentation()?;
///
/// // Get presentation dimensions
/// if let (Some(width), Some(height)) = (pres.slide_width()?, pres.slide_height()?) {
///     println!("Slide size: {}x{} EMUs", width, height);
/// }
///
/// // Access slides
/// for slide in pres.slides()? {
///     println!("Slide: {}", slide.text()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Presentation<'a> {
    /// The underlying presentation part
    part: PresentationPart<'a>,
    /// Reference to the OPC package for accessing related parts
    package: &'a OpcPackage,
    /// Whether package-layer notes match any pending legacy-writer state.
    notes_current: bool,
}

impl<'a> Presentation<'a> {
    /// Create a new Presentation.
    ///
    /// This is typically called internally by `Package::presentation()`.
    #[inline]
    pub(crate) fn new(
        part: PresentationPart<'a>,
        package: &'a OpcPackage,
        notes_current: bool,
    ) -> Self {
        Self {
            part,
            package,
            notes_current,
        }
    }

    /// Get the number of slides in the presentation.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// println!("Slide count: {}", pres.slide_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide_count(&self) -> Result<usize> {
        self.part.slide_count()
    }

    /// Get stable presentation slide identifiers in document order.
    ///
    /// These identifiers are used by features such as custom slide shows and
    /// PowerPoint 2010 sections.
    pub fn slide_ids(&self) -> Result<Vec<u32>> {
        self.part.slide_ids()
    }

    /// Get the slide width in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    /// 1 EMU = 1/914400 inch = 1/36000 mm
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// if let Some(width) = pres.slide_width()? {
    ///     let inches = width as f64 / 914400.0;
    ///     println!("Slide width: {:.2} inches", inches);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide_width(&self) -> Result<Option<i64>> {
        self.part.slide_width()
    }

    /// Get the slide height in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// if let Some(height) = pres.slide_height()? {
    ///     let inches = height as f64 / 914400.0;
    ///     println!("Slide height: {:.2} inches", inches);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide_height(&self) -> Result<Option<i64>> {
        self.part.slide_height()
    }

    /// Get all slides in the presentation.
    ///
    /// Returns a vector of `Slide` objects in presentation order.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (idx, slide) in pres.slides()?.iter().enumerate() {
    ///     println!("Slide {}: {}", idx + 1, slide.name()?);
    ///     println!("  Text: {}", slide.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slides(&self) -> Result<Vec<Slide<'a>>> {
        let slide_rids = self.part.slide_rids()?;
        let mut slides = Vec::with_capacity(slide_rids.len());

        let pres_part = self.part.part();

        for rid in slide_rids {
            let slide = resolve_presentation_relationship(
                pres_part,
                self.package,
                &rid,
                &[rt::SLIDE, STRICT_SLIDE_RELATIONSHIP_TYPE],
                "slide",
                ct::PML_SLIDE,
            )?;
            let slide_part = SlidePart::from_part(slide)?;
            slides.push(Slide::with_package(slide_part, self.package));
        }

        Ok(slides)
    }

    /// Get all slide masters in the presentation.
    ///
    /// Returns a vector of `SlideMaster` objects.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (idx, master) in pres.slide_masters()?.iter().enumerate() {
    ///     println!("Master {}: {}", idx + 1, master.name()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide_masters(&self) -> Result<Vec<SlideMaster<'a>>> {
        let master_rids = self.part.slide_master_rids()?;
        let mut masters = Vec::with_capacity(master_rids.len());

        let pres_part = self.part.part();

        for rid in master_rids {
            let master = resolve_presentation_relationship(
                pres_part,
                self.package,
                &rid,
                &[rt::SLIDE_MASTER, STRICT_SLIDE_MASTER_RELATIONSHIP_TYPE],
                "slide-master",
                ct::PML_SLIDE_MASTER,
            )?;
            let master_part = SlideMasterPart::from_part(master)?;
            masters.push(SlideMaster::with_package(master_part, self.package));
        }

        Ok(masters)
    }

    /// Get access to the underlying presentation part.
    ///
    /// This provides lower-level access to the presentation XML.
    #[inline]
    pub fn part(&self) -> &PresentationPart<'a> {
        &self.part
    }

    /// Get access to the underlying OPC package.
    #[inline]
    pub fn package(&self) -> &'a OpcPackage {
        self.package
    }

    // ========================================================================
    // Slide Size Manipulation
    // ========================================================================

    /// Get the slide dimensions as a tuple (width, height) in EMUs.
    ///
    /// Returns None if either dimension is not defined.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// if let Some((width, height)) = pres.slide_size()? {
    ///     println!("Slide size: {} x {} EMUs", width, height);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide_size(&self) -> Result<Option<(i64, i64)>> {
        match (self.slide_width()?, self.slide_height()?) {
            (Some(w), Some(h)) => Ok(Some((w, h))),
            _ => Ok(None),
        }
    }

    /// Get the presentation slide surface dimensions and declared size type.
    ///
    /// This preserves the declared type value in addition to the dimensions
    /// returned by slide_size.
    pub fn slide_size_metadata(&self) -> Result<Option<SlideSize>> {
        self.part.slide_size()
    }

    /// Get the notes and handout surface dimensions.
    pub fn notes_size(&self) -> Result<Option<NotesSize>> {
        self.part.notes_size()
    }

    /// Get the root-level presentation behavior and document settings.
    pub fn metadata(&self) -> Result<PresentationMetadata> {
        self.part.metadata()
    }

    /// Get the presentation-wide default text-style inventory.
    pub fn default_text_style(&self) -> Result<Option<PresentationDefaultTextStyle>> {
        self.part.default_text_style()
    }

    /// Get the presentation-wide East Asian line-breaking settings.
    ///
    /// Returns None when the presentation does not declare kinsoku settings.
    pub fn kinsoku_settings(&self) -> Result<Option<PresentationKinsokuSettings>> {
        self.part.kinsoku_settings()
    }

    /// Get the presentation-wide photo-album defaults.
    ///
    /// Returns None when the presentation is not declared as a photo album.
    pub fn photo_album(&self) -> Result<Option<PresentationPhotoAlbum>> {
        self.part.photo_album()
    }

    /// Get typed PowerPoint 2013 slide and notes guide extensions.
    ///
    /// Guide metadata and unknown extensions remain inert document data. This
    /// accessor does not resolve or open extension targets.
    pub fn extended_guides(
        &self,
    ) -> Result<crate::pptx::extended_guides::PresentationExtendedGuides> {
        self.part.extended_guides()
    }

    /// Get the presentation-level customer-data relationship references.
    ///
    /// Returns None when the presentation does not declare customer data.
    pub fn customer_data(&self) -> Result<Option<PresentationCustomerDataList>> {
        self.part.customer_data()
    }

    /// Get validated presentation-level embedded-font metadata and resources.
    ///
    /// Font programs are returned only as inert stored bytes. This does not
    /// parse, load, install, render, or execute a font program.
    pub fn embedded_fonts(&self) -> Result<Option<crate::pptx::PresentationEmbeddedFonts>> {
        crate::pptx::load_embedded_fonts(self.package)
    }

    /// Get the validated list of custom slide shows.
    ///
    /// Each show retains its stable ID and the presentation slide IDs it
    /// contains. The presentation slide relationships are validated before the
    /// list is returned.
    pub fn custom_shows(&self) -> Result<crate::pptx::CustomShowList> {
        Ok(crate::pptx::load_presentation_structure(self.package)?.custom_shows)
    }

    /// Get the root-level password-verification metadata.
    ///
    /// Returns None when the presentation has no modification verifier.
    pub fn modification_verifier(&self) -> Result<Option<PresentationModificationVerifier>> {
        self.part.modification_verifier()
    }

    /// Get the typed PowerPoint 2010 section list.
    ///
    /// Section membership is expressed using stable presentation slide IDs.
    pub fn sections(&self) -> Result<crate::pptx::sections::SectionList> {
        self.part.sections()
    }

    /// Get the relationship ID of the declared smart-tags data.
    ///
    /// Returns None when the presentation does not declare smart tags.
    pub fn smart_tags_relationship_id(&self) -> Result<Option<String>> {
        self.part.smart_tags_relationship_id()
    }

    /// Get the relationship ID of the declared handout master.
    ///
    /// Returns None when the presentation does not declare a handout master.
    pub fn handout_master_relationship_id(&self) -> Result<Option<String>> {
        self.part.handout_master_relationship_id()
    }

    /// Resolve and parse the declared handout master.
    ///
    /// Returns None when the presentation does not declare a handout master.
    /// The relationship must be internal, use a handout-master relationship
    /// type, and target a PresentationML handout-master part.
    pub fn handout_master(&self) -> Result<Option<HandoutMaster>> {
        let Some(relationship_id) = self.handout_master_relationship_id()? else {
            return Ok(None);
        };
        let handout_master = resolve_presentation_relationship(
            self.part.part(),
            self.package,
            &relationship_id,
            &[rt::HANDOUT_MASTER, STRICT_HANDOUT_MASTER_RELATIONSHIP_TYPE],
            "handout-master",
            ct::PML_HANDOUT_MASTER,
        )?;
        validate_handout_master_root(handout_master.blob())?;
        let xml = std::str::from_utf8(handout_master.blob()).map_err(|error| {
            OoxmlError::InvalidFormat(format!(
                "handout-master relationship '{relationship_id}' targets non-UTF-8 XML: {error}"
            ))
        })?;
        Ok(Some(HandoutMaster::parse_xml(xml)?))
    }

    /// Load the complete validated notes-master and notes-slide graph.
    ///
    /// Returns None when the presentation has no notes graph. Notes and theme
    /// resources are returned as inert stored data.
    pub fn notes(&self) -> Result<Option<litchi_pptx::notes::Graph>> {
        if !self.notes_current {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: "notes",
                reason: STALE_NOTES_REASON,
            });
        }
        Ok(litchi_pptx::notes::load(
            self.package,
            self.part.part().partname(),
        )?)
    }

    /// Discover attached VBA-project metadata without inspecting its payload.
    ///
    /// The project binary remains opaque and inert. This validates only its
    /// declared package relationship graph and content type.
    pub fn vba(&self) -> Result<Option<crate::pptx::VbaProject>> {
        crate::pptx::vba_project::discover_vba_project(self.package, self.part.part())
    }

    /// Load persisted Office Add-in task-pane metadata without activation.
    ///
    /// Add-ins, manifests, catalog entries, and linked content are never
    /// located, opened, fetched, or executed.
    pub fn task_panes(&self) -> Result<Option<web::Panes>> {
        Ok(web::load(self.package)?)
    }

    /// Read the fixed legacy and modern Ribbon slots for this presentation.
    ///
    /// Custom UI XML remains opaque inert data. Callback names, macros,
    /// commands, and linked content are never invoked or resolved. Use
    /// [`ribbon::Set::effective`] for modern-first precedence.
    pub fn ribbon(&self) -> Result<ribbon::Set<'_>> {
        Ok(ribbon::load(self.package)?)
    }

    // ========================================================================
    // Slide Access by Index
    // ========================================================================

    /// Get a specific slide by index.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the slide
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// if let Some(slide) = pres.slide(0)? {
    ///     println!("First slide: {}", slide.name()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide(&self, index: usize) -> Result<Option<Slide<'a>>> {
        let slide_rids = self.part.slide_rids()?;
        let Some(relationship_id) = slide_rids.get(index) else {
            return Ok(None);
        };
        let slide = resolve_presentation_relationship(
            self.part.part(),
            self.package,
            relationship_id,
            &[rt::SLIDE, STRICT_SLIDE_RELATIONSHIP_TYPE],
            "slide",
            ct::PML_SLIDE,
        )?;
        let slide_part = SlidePart::from_part(slide)?;
        Ok(Some(Slide::with_package(slide_part, self.package)))
    }

    // ========================================================================
    // Presentation-level Text Search
    // ========================================================================

    /// Search for text across all slides.
    ///
    /// Returns a vector of (slide_index, shape_index) tuples indicating
    /// where the search text was found.
    ///
    /// # Arguments
    /// * `query` - Text to search for (case-sensitive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// let results = pres.find_text("important")?;
    /// for (slide_idx, shape_idx) in results {
    ///     println!("Found in slide {} shape {}", slide_idx, shape_idx);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn find_text(&self, query: &str) -> Result<Vec<(usize, usize)>> {
        let mut results = Vec::new();

        for (slide_idx, slide) in self.slides()?.iter().enumerate() {
            let shape_matches = slide.find_text(query)?;
            for shape_idx in shape_matches {
                results.push((slide_idx, shape_idx));
            }
        }

        Ok(results)
    }

    // ========================================================================
    // Placeholder Management
    // ========================================================================

    /// Get all placeholders from a specific slide.
    ///
    /// Placeholders are special shapes on slides that define content areas,
    /// such as title, body text, charts, etc.
    ///
    /// # Arguments
    /// * `slide_index` - Zero-based index of the slide
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// if let Some(placeholders) = pres.get_placeholders(0)? {
    ///     println!("Slide has {} placeholders", placeholders.len());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_placeholders(&self, slide_index: usize) -> Result<Option<Vec<String>>> {
        if let Some(slide) = self.slide(slide_index)? {
            // Get the slide's placeholder inventory.
            let placeholders: Vec<String> = slide
                .placeholders()?
                .filter_map(|shape| {
                    shape
                        .placeholder()
                        .map(|placeholder| placeholder.kind().unwrap_or("obj").to_owned())
                })
                .collect();

            Ok(Some(placeholders))
        } else {
            Ok(None)
        }
    }

    // ========================================================================
    // Slide Statistics
    // ========================================================================

    /// Get statistics about all slides in the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, shape_count, text_length)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (idx, shape_count, text_len) in pres.slide_statistics()? {
    ///     println!("Slide {}: {} shapes, {} chars", idx, shape_count, text_len);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide_statistics(&self) -> Result<Vec<(usize, usize, usize)>> {
        let mut stats = Vec::new();

        for (idx, slide) in self.slides()?.iter().enumerate() {
            let shape_count = slide.shape_count()?;
            let text = slide.text()?;
            stats.push((idx, shape_count, text.len()));
        }

        Ok(stats)
    }

    /// Get the total number of shapes across all slides.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// println!("Total shapes: {}", pres.total_shape_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn total_shape_count(&self) -> Result<usize> {
        let mut total = 0;
        for slide in self.slides()? {
            total += slide.shape_count()?;
        }
        Ok(total)
    }

    /// Extract all text from the presentation.
    ///
    /// Returns concatenated text from all slides, separated by newlines.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// let all_text = pres.all_text()?;
    /// println!("Presentation text:\n{}", all_text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn all_text(&self) -> Result<String> {
        let mut texts = Vec::new();

        for slide in self.slides()? {
            let text = slide.text()?;
            if !text.is_empty() {
                texts.push(text);
            }
        }

        Ok(texts.join("\n\n"))
    }

    // ========================================================================
    // Advanced Features - Comments
    // ========================================================================

    /// Load the typed, validated legacy comment graph.
    ///
    /// Returns None when the presentation does not contain legacy comments.
    /// Comment text and extension payloads remain inert document data.
    pub fn comments(&self) -> Result<Option<crate::pptx::PresentationComments>> {
        crate::pptx::load_presentation_comments(self.package)
    }

    /// Load the typed, validated modern comment graph.
    ///
    /// Author and comment XML remain inert document data. This accessor never
    /// resolves identities or executes embedded payloads.
    pub fn modern_comments(
        &self,
    ) -> Result<crate::pptx::modern_comment_authors::ModernCommentGraph> {
        crate::pptx::modern_comment_authors::load_modern_comment_graph(self.package)
    }

    /// Get all comments from the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, comment).
    /// Returns an empty vector when the presentation has no legacy comments.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (slide_idx, comment) in pres.get_comments()? {
    ///     println!("Slide {}: {}", slide_idx, comment.text);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_comments(&self) -> Result<Vec<(usize, crate::pptx::parts::Comment)>> {
        let Some(comments) = self.comments()? else {
            return Ok(Vec::new());
        };
        let slide_indices = self
            .slides()?
            .into_iter()
            .enumerate()
            .map(|(slide_index, slide)| (slide.part().part().partname().to_string(), slide_index))
            .collect::<std::collections::HashMap<_, _>>();
        let mut all_comments = Vec::new();
        for slide_comments in comments.slides {
            let slide_index = slide_indices
                .get(&slide_comments.slide_part_name)
                .copied()
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!(
                        "comment graph references undeclared slide part '{}'",
                        slide_comments.slide_part_name
                    ))
                })?;
            all_comments.extend(slide_comments.comments.into_iter().map(|comment| {
                (
                    slide_index,
                    crate::pptx::parts::Comment {
                        author_id: comment.author_id,
                        text: comment.text,
                        x: comment.x,
                        y: comment.y,
                        datetime: comment.date_time,
                        index: Some(comment.index),
                    },
                )
            }));
        }
        all_comments.sort_by_key(|(slide_index, _)| *slide_index);
        Ok(all_comments)
    }

    /// Get all comment authors from the presentation.
    ///
    /// Returns a vector of comment authors if the commentAuthors.xml part exists.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for author in pres.get_comment_authors()? {
    ///     println!("Author: {} ({})", author.name, author.initials);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_comment_authors(&self) -> Result<Vec<crate::pptx::parts::CommentAuthor>> {
        let Some(comments) = self.comments()? else {
            return Ok(Vec::new());
        };
        Ok(comments
            .authors
            .into_iter()
            .map(|author| {
                crate::pptx::parts::CommentAuthor::new(author.id, author.name, author.initials)
            })
            .collect())
    }

    // ========================================================================
    // Advanced Features - Themes
    // ========================================================================

    /// Get all themes from the presentation.
    ///
    /// Returns a vector of themes. Each slide master typically has an associated theme.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for theme in pres.get_themes()? {
    ///     println!("Theme: {}", theme.name);
    ///     if let Some(major) = &theme.major_font {
    ///         println!("  Major font: {}", major.typeface);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_themes(&self) -> Result<Vec<crate::pptx::parts::Theme>> {
        let masters = self.slide_masters()?;
        masters.iter().map(SlideMaster::theme).collect()
    }

    // ========================================================================
    // Advanced Features - Charts
    // ========================================================================

    /// Discover all native chart parts reachable from presentation slides.
    ///
    /// Each result retains the owning slide, relationship ID, and part name in
    /// addition to basic inert chart metadata. Both transitional and strict
    /// chart relationships are accepted.
    pub fn charts(&self) -> Result<Vec<PptxChart>> {
        use crate::pptx::parts::ChartPart;

        let mut charts = Vec::new();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            let slide_part = slide.part().part();

            for relationship in slide_part.rels().iter() {
                if !matches!(relationship.reltype(), rt::CHART | rt::STRICT_CHART) {
                    continue;
                }
                if relationship.is_external() {
                    return Err(OoxmlError::InvalidRelationship(format!(
                        "chart relationship '{}' on slide {slide_index} must be internal",
                        relationship.r_id()
                    )));
                }

                let part_name = relationship.target_partname().map_err(|error| {
                    OoxmlError::InvalidRelationship(format!(
                        "invalid chart relationship '{}' on slide {slide_index}: {error}",
                        relationship.r_id()
                    ))
                })?;
                let part = self.package.get_part(&part_name).map_err(|error| {
                    OoxmlError::PartNotFound(format!(
                        "chart relationship '{}' on slide {slide_index} targets missing part '{}': {error}",
                        relationship.r_id(),
                        part_name.as_str()
                    ))
                })?;
                if part.content_type() != ct::DML_CHART {
                    return Err(OoxmlError::InvalidContentType {
                        expected: ct::DML_CHART.to_string(),
                        got: part.content_type().to_string(),
                    });
                }

                let info = ChartPart::from_part(part)?.chart_info()?;
                charts.push(PptxChart {
                    slide_index,
                    relationship_id: relationship.r_id().to_string(),
                    part_name,
                    info,
                });
            }
        }

        charts.sort_unstable_by(|left, right| {
            left.slide_index
                .cmp(&right.slide_index)
                .then_with(|| left.relationship_id.cmp(&right.relationship_id))
        });
        Ok(charts)
    }

    /// Discover inert InkML annotation content parts on presentation slides.
    ///
    /// Results retain slide and OPC relationship identity together with
    /// bounded stored-trace counts. Ink is never rendered, recognized,
    /// interpreted, or executed.
    pub fn ink_annotations(&self) -> Result<Vec<PptxInkAnnotation>> {
        let mut annotations = Vec::new();
        let mut limits = InkLoadLimits::default();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            annotations.extend(load_slide_ink_annotations(
                self.package,
                slide_index,
                slide.part().part(),
                &mut limits,
            )?);
        }

        Ok(annotations)
    }

    /// Discover persisted laser-pointer traces from presentation slides.
    ///
    /// Trace points are retained as bounded inert data. They are never
    /// replayed, rendered, interpolated, modified, or executed.
    pub fn laser_traces(&self) -> Result<Vec<PptxLaserTrace>> {
        let mut traces = Vec::new();
        let mut limits = LaserLoadLimits::default();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            traces.extend(load_slide_laser_traces(
                slide_index,
                slide.part().part(),
                &mut limits,
            )?);
        }

        Ok(traces)
    }

    /// Discover persisted slide-show event records from presentation slides.
    ///
    /// Event records remain inert historical metadata. This never replays
    /// triggers, seeks media, opens targets, or changes slide-show state.
    pub fn show_events(&self) -> Result<Vec<PptxSlideShowEvent>> {
        let mut events = Vec::new();
        let mut limits = ShowEventLoadLimits::default();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            events.extend(load_slide_show_events(
                slide_index,
                slide.part().part(),
                &mut limits,
            )?);
        }

        Ok(events)
    }

    /// Discover bounded, inert click and hover action settings on slides.
    ///
    /// Declared targets remain stored metadata only. This never follows links,
    /// opens files or presentations, runs macros or programs, plays media, or
    /// controls a slide show.
    pub fn action_settings(&self) -> Result<Vec<PptxActionSetting>> {
        let mut settings = Vec::new();
        let mut limits = ActionLoadLimits::default();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            settings.extend(load_slide_action_settings(
                self.package,
                slide_index,
                slide.part().part(),
                &mut limits,
            )?);
        }

        Ok(settings)
    }

    /// Discover bounded, inert OLE object shapes and declared payload targets.
    ///
    /// This never parses, opens, activates, renders, or executes an embedded
    /// object or package payload.
    pub fn ole_objects(&self) -> Result<Vec<PptxOleObject>> {
        let mut objects = Vec::new();
        let mut limits = OleLoadLimits::default();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            objects.extend(load_slide_ole_objects(
                self.package,
                slide_index,
                slide.part().part(),
                &mut limits,
            )?);
        }

        Ok(objects)
    }

    /// Discover bounded, inert slide controls (ActiveX/OCX) and their
    /// resolved controls-part descriptors.
    ///
    /// This never instantiates a control, resolves a CLSID, decodes binary
    /// control state, executes a macro, or follows an external relationship.
    pub fn controls(&self) -> Result<Vec<PptxSlideControl>> {
        let mut controls = Vec::new();
        let mut limits = ControlLoadLimits::default();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            controls.extend(load_slide_controls(
                self.package,
                slide_index,
                slide.part().part(),
                &mut limits,
            )?);
        }

        Ok(controls)
    }

    /// Load the presentation's WebVTT caption tracks.
    ///
    /// Internal tracks are parsed as bounded inert text. External targets are
    /// retained as document metadata and are never fetched.
    pub fn caption_tracks(&self) -> Result<Vec<crate::pptx::tracks::PresentationTrack>> {
        crate::pptx::tracks::load_presentation_tracks(self.package)
    }

    /// Load typed presentation-view settings, if the package contains them.
    ///
    /// View settings are returned as stored document data only; this does not
    /// alter the application's display state or follow outline-slide targets.
    pub fn view_properties(&self) -> Result<Option<crate::pptx::view_properties::ViewProperties>> {
        crate::pptx::view_properties::load_from_package(self.package)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Load the presentation's typed, bounded table-style catalog.
    pub fn styles(&self) -> Result<Option<litchi_pptx::table::style::List>> {
        Ok(litchi_pptx::table::style::load(self.package)?)
    }

    /// Load typed presentation settings, if the package contains them.
    ///
    /// Declared HTML publishing targets remain inert metadata and are never
    /// opened, fetched, or otherwise activated.
    pub fn presentation_properties(
        &self,
    ) -> Result<Option<crate::pptx::presentation_properties::PresentationProperties>> {
        crate::pptx::presentation_properties::load_from_package(self.package)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Load the PowerPoint Revision Information part, if present.
    ///
    /// Revision extension XML remains inert metadata and is never executed or
    /// used to resolve relationships.
    pub fn revision_information(
        &self,
    ) -> Result<Option<crate::pptx::revision_information::RevisionInformationPart>> {
        crate::pptx::revision_information::load_revision_information(self.package)
    }

    /// Load the PowerPoint Changes Information part, if present.
    ///
    /// Nested change descriptors remain inert XML and are never executed or
    /// used to resolve relationships.
    pub fn changes_information(
        &self,
    ) -> Result<Option<crate::pptx::changes_information::ChangesInformationPart>> {
        crate::pptx::changes_information::load_changes_information(self.package)
    }

    /// Get basic chart information from the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, chart_info).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (slide_idx, chart) in pres.get_charts()? {
    ///     println!("Slide {}: Chart type {:?}", slide_idx, chart.chart_type);
    ///     if let Some(title) = &chart.title {
    ///         println!("  Title: {}", title);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_charts(&self) -> Result<Vec<(usize, crate::pptx::parts::ChartInfo)>> {
        Ok(self
            .charts()?
            .into_iter()
            .map(|chart| (chart.slide_index, chart.info))
            .collect())
    }

    // ========================================================================
    // Advanced Features - Tables
    // ========================================================================

    /// Get all tables from the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, shape_index).
    /// The shape at the specified index contains a table.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (slide_idx, shape_idx) in pres.get_tables()? {
    ///     println!("Table found at slide {} shape {}", slide_idx, shape_idx);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_tables(&self) -> Result<Vec<(usize, usize)>> {
        let mut all_tables = Vec::new();

        for (slide_idx, slide) in self.slides()?.iter().enumerate() {
            let shapes = slide.shapes()?;
            for (shape_idx, shape) in shapes.iter().enumerate() {
                if matches!(shape, litchi_pptx::shape::Shape::Table(_)) {
                    all_tables.push((slide_idx, shape_idx));
                }
            }
        }

        Ok(all_tables)
    }

    // ========================================================================
    // Advanced Features - Hyperlinks
    // ========================================================================

    /// Get all hyperlinks from the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, hyperlink).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (slide_idx, hyperlink) in pres.get_hyperlinks()? {
    ///     if hyperlink.is_external() {
    ///         if let Some(tooltip) = hyperlink.tooltip() {
    ///             println!("Slide {}: External link: {}", slide_idx, tooltip);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_hyperlinks(&self) -> Result<Vec<(usize, crate::pptx::Hyperlink)>> {
        use crate::pptx::Hyperlink;

        let mut all_hyperlinks = Vec::new();

        // Iterate through all slides to find hyperlinks
        let slides = self.slides()?;
        for (slide_idx, slide) in slides.iter().enumerate() {
            let slide_part = slide.part().part();
            let rels = slide_part.rels();

            // Look for transitional and strict hyperlink relationships.
            for rel in rels
                .iter()
                .filter(|rel| matches!(rel.reltype(), rt::HYPERLINK | rt::STRICT_HYPERLINK))
            {
                let target = rel.target_ref();
                if target.is_empty() {
                    return Err(OoxmlError::InvalidRelationship(format!(
                        "hyperlink relationship '{}' on slide {slide_idx} has an empty target",
                        rel.r_id()
                    )));
                }
                all_hyperlinks.push((slide_idx, Hyperlink::from_xml(target, None)?));
            }

            // Also parse inline hyperlinks from slide XML (internal slide links)
            let slide_xml = slide_part.blob();
            for hyperlink in Self::parse_inline_hyperlinks(slide_xml)? {
                all_hyperlinks.push((slide_idx, hyperlink));
            }
        }

        Ok(all_hyperlinks)
    }

    /// Parse inline hyperlinks from slide XML.
    fn parse_inline_hyperlinks(xml: &[u8]) -> Result<Vec<crate::pptx::Hyperlink>> {
        use crate::pptx::Hyperlink;

        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(true);
        let mut hyperlinks = Vec::new();
        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(element) | Event::Empty(element)
                    if is_drawingml_name(&namespace, element.name(), b"hlinkClick") =>
                {
                    let action = unqualified_attribute_value(&element, b"action", decoder)?;
                    let tooltip = unqualified_attribute_value(&element, b"tooltip", decoder)?;
                    if let Some(action) = action {
                        if action.is_empty() {
                            return Err(OoxmlError::InvalidFormat(
                                "inline hyperlink action cannot be empty".to_string(),
                            ));
                        }
                        hyperlinks.push(Hyperlink::from_xml(&action, tooltip)?);
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(hyperlinks)
    }

    // ========================================================================
    // Advanced Features - Sections
    // ========================================================================

    /// Get all sections from the presentation.
    ///
    /// Sections are used to organize slides into logical groups.
    /// Returns a vector of tuples: (section_name, slide_indices).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (section_name, slide_indices) in pres.get_sections()? {
    ///     println!("Section '{}': {} slides", section_name, slide_indices.len());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_sections(&self) -> Result<Vec<(String, Vec<usize>)>> {
        let slide_ids = self.part.slide_ids()?;
        self.sections()?
            .sections()
            .iter()
            .map(|section| {
                let slide_indices = section
                    .slide_ids
                    .iter()
                    .map(|slide_id| {
                        slide_ids
                            .iter()
                            .position(|id| id == slide_id)
                            .ok_or_else(|| {
                                OoxmlError::InvalidFormat(format!(
                                    "PowerPoint section references undeclared slide ID {slide_id}"
                                ))
                            })
                    })
                    .collect::<Result<Vec<_>>>()?;
                Ok((section.name.clone().unwrap_or_default(), slide_indices))
            })
            .collect()
    }

    // ========================================================================
    // Notes
    // ========================================================================

    /// Get all notes from the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, notes_text).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (slide_idx, notes) in pres.get_notes()? {
    ///     println!("Slide {}: {}", slide_idx, notes);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_notes(&self) -> Result<Vec<(usize, String)>> {
        let mut all_notes = Vec::new();

        for (slide_idx, slide) in self.slides()?.iter().enumerate() {
            if let Some(notes) = slide.notes()? {
                all_notes.push((slide_idx, notes));
            }
        }

        Ok(all_notes)
    }
}

#[cfg(test)]
mod tests {
    // Tests will be added as implementation progresses
}
