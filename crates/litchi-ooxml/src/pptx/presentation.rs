/// Main presentation object - the high-level API for working with presentations.
use crate::error::{OoxmlError, Result};
use crate::pptx::actions::{ActionLoadLimits, PptxActionSetting, load_slide_action_settings};
use crate::pptx::ink::{InkLoadLimits, PptxInkAnnotation, load_slide_ink_annotations};
use crate::pptx::laser::{LaserLoadLimits, PptxLaserTrace, load_slide_laser_traces};
use crate::pptx::ole::{OleLoadLimits, PptxOleObject, load_slide_ole_objects};
use crate::pptx::parts::{PresentationPart, SlideMasterPart, SlidePart};
use crate::pptx::show_events::{
    PptxSlideShowEvent, ShowEventLoadLimits, load_slide_show_events,
};
use crate::pptx::slide::{Slide, SlideMaster};
use crate::pptx::tags::{SlideTagList, TagList};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;

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

/// A programmable tag-list part discovered on a presentation slide.
///
/// Tag names and values are retained only as inert document strings. They are
/// never interpreted as XML, paths, commands, or relationship targets.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxTagList {
    slide_index: usize,
    tag_list_index: usize,
    value: SlideTagList,
}

impl PptxTagList {
    /// Return the zero-based index of the slide that owns this tag-list part.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this tag-list on its slide.
    #[inline]
    pub fn tag_list_index(&self) -> usize {
        self.tag_list_index
    }

    /// Return the relationship ID from the owning slide to this tag-list part.
    #[inline]
    pub fn relationship_id(&self) -> &str {
        self.value.relationship_id()
    }

    /// Return the absolute OPC part name of this tag-list part.
    #[inline]
    pub fn part_name(&self) -> &str {
        self.value.part_name()
    }

    /// Return the parsed inert programmable tags.
    #[inline]
    pub fn tag_list(&self) -> &TagList {
        self.value.tag_list()
    }

    /// Return the underlying slide-scoped tag-list value.
    #[inline]
    pub fn as_slide_tag_list(&self) -> &SlideTagList {
        &self.value
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
}

impl<'a> Presentation<'a> {
    /// Create a new Presentation.
    ///
    /// This is typically called internally by `Package::presentation()`.
    #[inline]
    pub(crate) fn new(part: PresentationPart<'a>, package: &'a OpcPackage) -> Self {
        Self { part, package }
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
            // Get the target reference from the relationship
            let target_ref = pres_part.target_ref(&rid)?;

            // Resolve the target partname and get the part from the package
            let base_uri = pres_part.partname().base_uri();
            let target_partname = PackURI::from_rel_ref(base_uri, target_ref)
                .map_err(crate::error::OoxmlError::InvalidFormat)?;
            let related_part = self.package.get_part(&target_partname)?;

            let slide_part = SlidePart::from_part(related_part)?;
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
            // Get the target reference from the relationship
            let target_ref = pres_part.target_ref(&rid)?;

            // Resolve the target partname and get the part from the package
            let base_uri = pres_part.partname().base_uri();
            let target_partname = PackURI::from_rel_ref(base_uri, target_ref)
                .map_err(crate::error::OoxmlError::InvalidFormat)?;
            let related_part = self.package.get_part(&target_partname)?;

            let master_part = SlideMasterPart::from_part(related_part)?;
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
        Ok(self.slides()?.into_iter().nth(index))
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
            // Get shapes and filter for placeholders
            let shapes = slide.shapes()?;
            let placeholders: Vec<String> = shapes
                .iter()
                .filter(|s| s.is_placeholder())
                .filter_map(|s| s.placeholder_type().ok())
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

    /// Get all comments from the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, comment).
    /// Returns empty vector if no comments are found or comment authors are not available.
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
        use crate::pptx::parts::CommentsPart;

        let mut all_comments = Vec::new();

        // Iterate through all slides to find comments
        let slides = self.slides()?;
        for (slide_idx, slide) in slides.iter().enumerate() {
            let slide_part = slide.part().part();
            let rels = slide_part.rels();

            // Look for comments relationship
            for rel in rels.iter() {
                if rel.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
                {
                    // Get the comments part
                    let base_uri = slide_part.partname().base_uri();
                    let comments_partname = PackURI::from_rel_ref(base_uri, rel.target_ref())
                        .map_err(crate::error::OoxmlError::InvalidFormat)?;

                    if let Ok(comments_part) = self.package.get_part(&comments_partname) {
                        let comments_part = CommentsPart::from_part(comments_part)?;
                        let comments = comments_part.comments()?;

                        for comment in comments {
                            all_comments.push((slide_idx, comment));
                        }
                    }
                }
            }
        }

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
        use crate::pptx::parts::CommentAuthorsPart;

        let pres_part = self.part.part();
        let rels = pres_part.rels();

        // Look for comment authors relationship
        for rel in rels.iter() {
            if rel.reltype()
                == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
            {
                let base_uri = pres_part.partname().base_uri();
                let authors_partname = PackURI::from_rel_ref(base_uri, rel.target_ref())
                    .map_err(crate::error::OoxmlError::InvalidFormat)?;

                if let Ok(authors_part) = self.package.get_part(&authors_partname) {
                    let authors_part = CommentAuthorsPart::from_part(authors_part)?;
                    return authors_part.authors();
                }
            }
        }

        Ok(Vec::new())
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

    /// Discover programmable tag-list parts reachable from presentation slides.
    ///
    /// Tag names and values remain inert document strings. This never follows
    /// values as paths or relationships, evaluates markup, or executes commands.
    pub fn tag_lists(&self) -> Result<Vec<PptxTagList>> {
        let mut tag_lists = Vec::new();

        for (slide_index, slide) in self.slides()?.iter().enumerate() {
            tag_lists.extend(slide.tag_lists()?.into_iter().enumerate().map(
                |(tag_list_index, value)| PptxTagList {
                    slide_index,
                    tag_list_index,
                    value,
                },
            ));
        }

        Ok(tag_lists)
    }

    /// Load typed presentation-view settings, if the package contains them.
    ///
    /// View settings are returned as stored document data only; this does not
    /// alter the application's display state or follow outline-slide targets.
    pub fn view_properties(&self) -> Result<Option<crate::pptx::view_properties::ViewProperties>> {
        crate::pptx::view_properties::load_from_package(self.package)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
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
                if shape.has_table() {
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

            // Look for hyperlink relationships
            for rel in rels.iter() {
                if rel.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                {
                    // External hyperlink
                    let target = rel.target_ref();
                    if let Ok(hyperlink) = Hyperlink::from_xml(target, None) {
                        all_hyperlinks.push((slide_idx, hyperlink));
                    }
                }
            }

            // Also parse inline hyperlinks from slide XML (internal slide links)
            let slide_xml = slide_part.blob();
            if let Ok(inline_links) = Self::parse_inline_hyperlinks(slide_xml) {
                for hyperlink in inline_links {
                    all_hyperlinks.push((slide_idx, hyperlink));
                }
            }
        }

        Ok(all_hyperlinks)
    }

    /// Parse inline hyperlinks from slide XML.
    fn parse_inline_hyperlinks(xml: &[u8]) -> Result<Vec<crate::pptx::Hyperlink>> {
        use crate::pptx::Hyperlink;
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut hyperlinks = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                    if e.local_name().as_ref() == b"hlinkClick" =>
                {
                    let mut action = None;
                    let mut tooltip = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"action" => {
                                action =
                                    std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
                            },
                            b"tooltip" => {
                                tooltip =
                                    std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
                            },
                            _ => {},
                        }
                    }

                    if let Some(action_str) = action
                        && let Ok(hyperlink) = Hyperlink::from_xml(&action_str, tooltip)
                    {
                        hyperlinks.push(hyperlink);
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
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
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let xml = self.part.part().blob();
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut sections = Vec::new();
        let mut current_section: Option<(String, usize)> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == b"section" {
                        let mut name = String::new();
                        let mut id = 0;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    name = std::str::from_utf8(&attr.value)
                                        .map(|s| s.to_string())
                                        .unwrap_or_default();
                                },
                                b"id" => {
                                    id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                },
                                _ => {},
                            }
                        }

                        if !name.is_empty() {
                            current_section = Some((name, id));
                        }
                    } else if e.local_name().as_ref() == b"sldId" && current_section.is_some() {
                        // This slide belongs to the current section
                        // We'll need to track slide IDs and map them to indices
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"section"
                        && let Some((name, _id)) = current_section.take()
                    {
                        // For now, we'll create empty section entries
                        // A full implementation would track slide IDs
                        sections.push((name, Vec::new()));
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        Ok(sections)
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
