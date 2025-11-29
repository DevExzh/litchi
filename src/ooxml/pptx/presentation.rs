/// Main presentation object - the high-level API for working with presentations.
use crate::ooxml::error::Result;
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::packuri::PackURI;
use crate::ooxml::pptx::parts::{PresentationPart, SlideMasterPart, SlidePart};
use crate::ooxml::pptx::slide::{Slide, SlideMaster};

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
/// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
                .map_err(crate::ooxml::error::OoxmlError::InvalidFormat)?;
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
    /// use litchi::ooxml::pptx::Package;
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
                .map_err(crate::ooxml::error::OoxmlError::InvalidFormat)?;
            let related_part = self.package.get_part(&target_partname)?;

            let master_part = SlideMasterPart::from_part(related_part)?;
            masters.push(SlideMaster::new(master_part));
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (slide_idx, comment) in pres.get_comments()? {
    ///     println!("Slide {}: {}", slide_idx, comment.text);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_comments(&self) -> Result<Vec<(usize, crate::ooxml::pptx::parts::Comment)>> {
        use crate::ooxml::pptx::parts::CommentsPart;

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
                        .map_err(crate::ooxml::error::OoxmlError::InvalidFormat)?;

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
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for author in pres.get_comment_authors()? {
    ///     println!("Author: {} ({})", author.name, author.initials);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_comment_authors(&self) -> Result<Vec<crate::ooxml::pptx::parts::CommentAuthor>> {
        use crate::ooxml::pptx::parts::CommentAuthorsPart;

        let pres_part = self.part.part();
        let rels = pres_part.rels();

        // Look for comment authors relationship
        for rel in rels.iter() {
            if rel.reltype()
                == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
            {
                let base_uri = pres_part.partname().base_uri();
                let authors_partname = PackURI::from_rel_ref(base_uri, rel.target_ref())
                    .map_err(crate::ooxml::error::OoxmlError::InvalidFormat)?;

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
    /// use litchi::ooxml::pptx::Package;
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
    pub fn get_themes(&self) -> Result<Vec<crate::ooxml::pptx::parts::Theme>> {
        use crate::ooxml::pptx::parts::ThemePart;

        let mut themes = Vec::new();

        // Get themes from slide masters
        let masters = self.slide_masters()?;
        for master in masters {
            let master_part = master.part().part();
            let rels = master_part.rels();

            // Look for theme relationship
            for rel in rels.iter() {
                if rel.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                {
                    let base_uri = master_part.partname().base_uri();
                    let theme_partname = PackURI::from_rel_ref(base_uri, rel.target_ref())
                        .map_err(crate::ooxml::error::OoxmlError::InvalidFormat)?;

                    if let Ok(theme_part) = self.package.get_part(&theme_partname) {
                        let theme_part = ThemePart::from_part(theme_part)?;
                        themes.push(theme_part.theme()?);
                    }
                }
            }
        }

        Ok(themes)
    }

    // ========================================================================
    // Advanced Features - Charts
    // ========================================================================

    /// Get all charts from the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, chart_info).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
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
    pub fn get_charts(&self) -> Result<Vec<(usize, crate::ooxml::pptx::parts::ChartInfo)>> {
        use crate::ooxml::pptx::parts::ChartPart;

        let mut all_charts = Vec::new();

        // Iterate through all slides to find charts
        let slides = self.slides()?;
        for (slide_idx, slide) in slides.iter().enumerate() {
            let slide_part = slide.part().part();
            let rels = slide_part.rels();

            // Look for chart relationships
            for rel in rels.iter() {
                if rel.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                {
                    let base_uri = slide_part.partname().base_uri();
                    let chart_partname = PackURI::from_rel_ref(base_uri, rel.target_ref())
                        .map_err(crate::ooxml::error::OoxmlError::InvalidFormat)?;

                    if let Ok(chart_part) = self.package.get_part(&chart_partname) {
                        let chart_part = ChartPart::from_part(chart_part)?;
                        let chart_info = chart_part.chart_info()?;
                        all_charts.push((slide_idx, chart_info));
                    }
                }
            }
        }

        Ok(all_charts)
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for (slide_idx, hyperlink) in pres.get_hyperlinks()? {
    ///     if hyperlink.is_external() {
    ///         println!("Slide {}: External link to {}", slide_idx, hyperlink.target());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_hyperlinks(&self) -> Result<Vec<(usize, crate::ooxml::pptx::Hyperlink)>> {
        use crate::ooxml::pptx::Hyperlink;

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
    fn parse_inline_hyperlinks(xml: &[u8]) -> Result<Vec<crate::ooxml::pptx::Hyperlink>> {
        use crate::ooxml::pptx::Hyperlink;
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut hyperlinks = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == b"hlinkClick" {
                        let mut action = None;
                        let mut tooltip = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"action" => {
                                    action = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(|s| s.to_string());
                                },
                                b"tooltip" => {
                                    tooltip = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(|s| s.to_string());
                                },
                                _ => {},
                            }
                        }

                        if let Some(action_str) = action
                            && let Ok(hyperlink) = Hyperlink::from_xml(&action_str, tooltip)
                        {
                            hyperlinks.push(hyperlink);
                        }
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
