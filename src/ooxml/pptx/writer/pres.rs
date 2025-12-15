/// Presentation writer for PPTX.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::pptx::customshow::{CustomShow, CustomShowList};
use crate::ooxml::pptx::handout::HandoutMaster;
use crate::ooxml::pptx::protection::PresentationProtection;
use crate::ooxml::pptx::sections::{Section, SectionList};
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;

// Import shared format types
use super::super::format::ImageFormat;
use super::slide::MutableSlide;

// ============================================================================
// Chart and SmartArt Parts Storage
// ============================================================================

/// Stores all parts needed for a chart embedded in the presentation.
///
/// Charts in PPTX consist of:
/// - Chart XML defining the chart type, series, and axes
/// - Optional colors and style XML for theming
/// - Embedded Excel workbook containing the source data
#[derive(Debug, Clone)]
pub struct ChartParts {
    /// Chart XML content (chartN.xml)
    pub chart_xml: String,
    /// Chart colors XML (colorsN.xml) - optional styling
    pub colors_xml: Option<String>,
    /// Chart style XML (styleN.xml) - optional styling
    pub style_xml: Option<String>,
    /// Embedded Excel workbook bytes containing chart data
    pub excel_data: Vec<u8>,
}

/// Stores all parts needed for a SmartArt diagram.
///
/// SmartArt diagrams require five separate XML parts:
/// - Data: The actual content/text of the diagram nodes
/// - Layout: How the nodes are arranged
/// - Style: Visual styling (3D effects, shadows, etc.)
/// - Colors: Color scheme for the diagram
/// - Drawing: Pre-rendered shapes for display (required for PowerPoint compatibility)
#[derive(Debug, Clone)]
pub struct SmartArtParts {
    /// Diagram data XML (dataN.xml) - contains node text and structure
    pub data_xml: String,
    /// Diagram layout XML (layoutN.xml) - defines arrangement
    pub layout_xml: String,
    /// Diagram quick style XML (quickStyleN.xml) - visual effects
    pub style_xml: String,
    /// Diagram colors XML (colorsN.xml) - color scheme
    pub colors_xml: String,
    /// The SmartArt data for generating drawing XML during save
    pub(crate) smartart: crate::ooxml::pptx::smartart::SmartArt,
}

/// A mutable PowerPoint presentation for writing and modification.
///
/// Provides methods to add and modify slides, set dimensions, and configure presentation settings.
#[derive(Debug)]
pub struct MutablePresentation {
    /// Slides in the presentation
    pub(crate) slides: Vec<MutableSlide>,
    /// Slide width in EMUs (English Metric Units, 914400 EMU = 1 inch)
    slide_width: i64,
    /// Slide height in EMUs
    slide_height: i64,
    /// Sections for organizing slides
    sections: SectionList,
    /// Custom slide shows
    custom_shows: CustomShowList,
    /// Presentation protection settings
    protection: PresentationProtection,
    /// Handout master for printing settings
    handout_master: Option<HandoutMaster>,
    /// Whether the presentation has been modified
    modified: bool,
    /// Charts in the presentation (chart_idx -> parts)
    pub(crate) charts: HashMap<u32, ChartParts>,
    /// SmartArt diagrams in the presentation (diagram_idx -> parts)
    pub(crate) smartarts: HashMap<u32, SmartArtParts>,
    /// Next chart index for unique naming (chart1.xml, chart2.xml, etc.)
    pub(crate) next_chart_idx: u32,
    /// Next SmartArt index for unique naming
    pub(crate) next_smartart_idx: u32,
}

impl MutablePresentation {
    /// Create a new empty presentation with default dimensions.
    ///
    /// Default size is 10" x 7.5" (standard 4:3 aspect ratio).
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            slide_width: 9144000,  // 10 inches
            slide_height: 6858000, // 7.5 inches
            sections: SectionList::new(),
            custom_shows: CustomShowList::new(),
            protection: PresentationProtection::new(),
            handout_master: None,
            modified: false,
            charts: HashMap::new(),
            smartarts: HashMap::new(),
            next_chart_idx: 1,
            next_smartart_idx: 1,
        }
    }

    /// Add a new slide to the presentation.
    pub fn add_slide(&mut self) -> Result<&mut MutableSlide> {
        let slide_id = (self.slides.len() + 256) as u32;
        let slide = MutableSlide::new(slide_id);
        self.slides.push(slide);
        self.modified = true;
        Ok(self.slides.last_mut().unwrap())
    }

    /// Get the number of slides.
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Get a mutable reference to a slide by index (0-based).
    pub fn slide_mut(&mut self, index: usize) -> Option<&mut MutableSlide> {
        self.slides.get_mut(index)
    }

    /// Get the slide width in EMUs.
    pub fn slide_width(&self) -> i64 {
        self.slide_width
    }

    /// Set the slide width in EMUs.
    pub fn set_slide_width(&mut self, width: i64) {
        self.slide_width = width;
        self.modified = true;
    }

    /// Get the slide height in EMUs.
    pub fn slide_height(&self) -> i64 {
        self.slide_height
    }

    /// Set the slide height in EMUs.
    pub fn set_slide_height(&mut self, height: i64) {
        self.slide_height = height;
        self.modified = true;
    }

    /// Check if the presentation has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified || self.slides.iter().any(|s| s.is_modified())
    }

    // ========================================================================
    // Slide Manipulation
    // ========================================================================

    /// Delete a slide by index.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the slide to delete
    ///
    /// # Returns
    /// * `Ok(())` if the slide was successfully deleted
    /// * `Err` if the index is out of bounds
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.add_slide().unwrap();
    /// pres.add_slide().unwrap();
    /// assert_eq!(pres.slide_count(), 2);
    ///
    /// pres.delete_slide(0).unwrap();
    /// assert_eq!(pres.slide_count(), 1);
    /// ```
    pub fn delete_slide(&mut self, index: usize) -> Result<()> {
        if index >= self.slides.len() {
            return Err(OoxmlError::InvalidFormat(format!(
                "Slide index {} out of bounds (max: {})",
                index,
                self.slides.len() - 1
            )));
        }

        self.slides.remove(index);
        self.modified = true;
        Ok(())
    }

    /// Duplicate a slide by index.
    ///
    /// Creates a copy of the slide at the specified index and appends it
    /// to the end of the presentation.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the slide to duplicate
    ///
    /// # Returns
    /// * `Ok(usize)` - Index of the newly created duplicate slide
    /// * `Err` if the index is out of bounds
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    /// slide.set_title("Original");
    ///
    /// let new_index = pres.duplicate_slide(0).unwrap();
    /// assert_eq!(new_index, 1);
    /// assert_eq!(pres.slide_count(), 2);
    /// ```
    pub fn duplicate_slide(&mut self, index: usize) -> Result<usize> {
        if index >= self.slides.len() {
            return Err(OoxmlError::InvalidFormat(format!(
                "Slide index {} out of bounds (max: {})",
                index,
                self.slides.len() - 1
            )));
        }

        // Clone the slide
        let slide_to_duplicate = &self.slides[index];
        let mut new_slide = slide_to_duplicate.clone();

        // Assign a new slide ID
        let new_slide_id = (self.slides.len() + 256) as u32;
        new_slide.set_slide_id(new_slide_id);

        self.slides.push(new_slide);
        self.modified = true;
        Ok(self.slides.len() - 1)
    }

    /// Move a slide from one position to another.
    ///
    /// # Arguments
    /// * `from_index` - Current zero-based index of the slide
    /// * `to_index` - Target zero-based index for the slide
    ///
    /// # Returns
    /// * `Ok(())` if the slide was successfully moved
    /// * `Err` if either index is out of bounds
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.add_slide().unwrap().set_title("First");
    /// pres.add_slide().unwrap().set_title("Second");
    /// pres.add_slide().unwrap().set_title("Third");
    ///
    /// // Move the first slide to the end
    /// pres.move_slide(0, 2).unwrap();
    /// assert_eq!(pres.slide_mut(2).unwrap().title(), Some("First"));
    /// ```
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<()> {
        if from_index >= self.slides.len() {
            return Err(OoxmlError::InvalidFormat(format!(
                "Source index {} out of bounds (max: {})",
                from_index,
                self.slides.len() - 1
            )));
        }

        if to_index >= self.slides.len() {
            return Err(OoxmlError::InvalidFormat(format!(
                "Target index {} out of bounds (max: {})",
                to_index,
                self.slides.len() - 1
            )));
        }

        if from_index == to_index {
            return Ok(());
        }

        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        self.modified = true;
        Ok(())
    }

    /// Get all slides as an immutable slice.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.add_slide().unwrap();
    /// pres.add_slide().unwrap();
    ///
    /// for (i, slide) in pres.slides().iter().enumerate() {
    ///     println!("Slide {}: {:?}", i, slide.title());
    /// }
    /// ```
    pub fn slides(&self) -> &[MutableSlide] {
        &self.slides
    }

    // ========================================================================
    // Slide Size Manipulation
    // ========================================================================

    /// Set slide dimensions (width and height) in EMUs.
    ///
    /// # Arguments
    /// * `width` - Slide width in EMUs
    /// * `height` - Slide height in EMUs
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// // Set to 16:9 aspect ratio (10" x 5.625")
    /// pres.set_slide_size(9144000, 5143500);
    /// assert_eq!(pres.slide_width(), 9144000);
    /// assert_eq!(pres.slide_height(), 5143500);
    /// ```
    pub fn set_slide_size(&mut self, width: i64, height: i64) {
        self.slide_width = width;
        self.slide_height = height;
        self.modified = true;
    }

    /// Get slide dimensions as a tuple (width, height) in EMUs.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let pres = MutablePresentation::new();
    /// let (width, height) = pres.slide_size();
    /// println!("Slide size: {}x{} EMUs", width, height);
    /// ```
    pub fn slide_size(&self) -> (i64, i64) {
        (self.slide_width, self.slide_height)
    }

    /// Set slide size to standard 4:3 aspect ratio (10" x 7.5").
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.set_standard_slide_size();
    /// assert_eq!(pres.slide_size(), (9144000, 6858000));
    /// ```
    pub fn set_standard_slide_size(&mut self) {
        self.set_slide_size(9144000, 6858000); // 10" x 7.5"
    }

    /// Set slide size to widescreen 16:9 aspect ratio (10" x 5.625").
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.set_widescreen_slide_size();
    /// assert_eq!(pres.slide_size(), (9144000, 5143500));
    /// ```
    pub fn set_widescreen_slide_size(&mut self) {
        self.set_slide_size(9144000, 5143500); // 10" x 5.625"
    }

    // ========================================================================
    // Sections
    // ========================================================================

    /// Add a section to organize slides.
    ///
    /// # Arguments
    /// * `name` - Display name for the section
    /// * `slide_ids` - IDs of slides to include in this section
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.add_slide().unwrap();
    /// pres.add_slide().unwrap();
    /// pres.add_section("Introduction", vec![256, 257]);
    /// ```
    pub fn add_section(&mut self, name: impl Into<String>, slide_ids: Vec<u32>) {
        let section =
            Section::new(name, crate::common::id::generate_guid_braced()).with_slides(slide_ids);
        self.sections.add_section(section);
        self.modified = true;
    }

    /// Get the sections in this presentation.
    pub fn sections(&self) -> &SectionList {
        &self.sections
    }

    /// Get the number of sections.
    pub fn section_count(&self) -> usize {
        self.sections.len()
    }

    // ========================================================================
    // Custom Slide Shows
    // ========================================================================

    /// Create a custom slide show.
    ///
    /// Custom shows allow presenting a subset of slides in a specific order.
    ///
    /// # Arguments
    /// * `name` - Display name for the custom show
    /// * `slide_ids` - IDs of slides to include (in presentation order)
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.add_slide().unwrap();
    /// pres.add_slide().unwrap();
    /// pres.add_slide().unwrap();
    /// pres.create_custom_show("Executive Summary", vec![256, 258]); // Skip slide 257
    /// ```
    pub fn create_custom_show(&mut self, name: impl Into<String>, slide_ids: Vec<u32>) {
        self.custom_shows.create(name, slide_ids);
        self.modified = true;
    }

    /// Get the custom shows in this presentation.
    pub fn custom_shows(&self) -> &CustomShowList {
        &self.custom_shows
    }

    /// Get a custom show by name.
    pub fn get_custom_show(&self, name: &str) -> Option<&CustomShow> {
        self.custom_shows.get_by_name(name)
    }

    /// Remove a custom show by name.
    pub fn remove_custom_show(&mut self, name: &str) -> bool {
        let result = self.custom_shows.remove_by_name(name).is_some();
        if result {
            self.modified = true;
        }
        result
    }

    // ========================================================================
    // Protection
    // ========================================================================

    /// Set presentation protection settings.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{MutablePresentation, PresentationProtection};
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.set_protection(
    ///     PresentationProtection::new()
    ///         .with_read_only_recommended(true)
    ///         .with_structure_protection(true)
    /// );
    /// ```
    pub fn set_protection(&mut self, protection: PresentationProtection) {
        self.protection = protection;
        self.modified = true;
    }

    /// Get the protection settings.
    pub fn protection(&self) -> &PresentationProtection {
        &self.protection
    }

    /// Get mutable protection settings.
    pub fn protection_mut(&mut self) -> &mut PresentationProtection {
        self.modified = true;
        &mut self.protection
    }

    /// Set read-only recommended flag.
    pub fn set_read_only_recommended(&mut self, value: bool) {
        self.protection.read_only_recommended = value;
        self.modified = true;
    }

    // ========================================================================
    // Handout Master
    // ========================================================================

    /// Set handout master settings for printing.
    ///
    /// The handout master defines the layout for printed handouts that show
    /// multiple slides per page.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{MutablePresentation, HandoutMaster, HandoutLayout};
    ///
    /// let mut pres = MutablePresentation::new();
    /// pres.set_handout_master(
    ///     HandoutMaster::new()
    ///         .with_layout(HandoutLayout::SixSlides)
    ///         .with_header("My Presentation")
    ///         .with_footer("Confidential")
    ///         .with_slide_numbers()
    /// );
    /// ```
    pub fn set_handout_master(&mut self, handout_master: HandoutMaster) {
        self.handout_master = Some(handout_master);
        self.modified = true;
    }

    /// Get the handout master settings.
    pub fn handout_master(&self) -> Option<&HandoutMaster> {
        self.handout_master.as_ref()
    }

    /// Remove handout master settings (use default).
    pub fn remove_handout_master(&mut self) {
        self.handout_master = None;
        self.modified = true;
    }

    /// Check if a custom handout master is set.
    pub fn has_handout_master(&self) -> bool {
        self.handout_master.is_some()
    }

    // ========================================================================
    // Internal Methods
    // ========================================================================

    /// Collect all images from all slides in the presentation.
    pub(crate) fn collect_all_images(&self) -> Vec<(usize, &[u8], ImageFormat)> {
        let mut all_images = Vec::new();

        for (slide_index, slide) in self.slides.iter().enumerate() {
            for (image_data, image_format) in slide.collect_images() {
                all_images.push((slide_index, image_data, image_format));
            }
        }

        all_images
    }

    /// Collect all background images from all slides in the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, image_data, image_format).
    pub(crate) fn collect_all_background_images(&self) -> Vec<(usize, &[u8], ImageFormat)> {
        let mut background_images = Vec::new();

        for (slide_index, slide) in self.slides.iter().enumerate() {
            if let Some((image_data, image_format)) = slide.get_background_image() {
                background_images.push((slide_index, image_data, image_format));
            }
        }

        background_images
    }

    /// Collect all media (audio/video) from all slides in the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, media_index_in_slide, media_data, media_format).
    pub(crate) fn collect_all_media(
        &self,
    ) -> Vec<(usize, usize, &[u8], crate::ooxml::pptx::media::MediaFormat)> {
        let mut all_media = Vec::new();

        for (slide_index, slide) in self.slides.iter().enumerate() {
            for (media_index, (media_data, media_format)) in
                slide.collect_media().iter().enumerate()
            {
                all_media.push((slide_index, media_index, *media_data, *media_format));
            }
        }

        all_media
    }

    /// Collect all comments from all slides in the presentation.
    ///
    /// Returns a vector of tuples: (slide_index, comments_slice).
    pub(crate) fn collect_all_comments(
        &self,
    ) -> Vec<(usize, &[crate::ooxml::pptx::parts::Comment])> {
        let mut all_comments = Vec::new();

        for (slide_index, slide) in self.slides.iter().enumerate() {
            if !slide.comments().is_empty() {
                all_comments.push((slide_index, slide.comments()));
            }
        }

        all_comments
    }

    /// Check if any slide has comments.
    #[allow(dead_code)] // Public API for future use
    pub(crate) fn has_comments(&self) -> bool {
        self.slides.iter().any(|s| !s.comments().is_empty())
    }

    /// Check if any slide has media (audio/video).
    #[allow(dead_code)] // Public API for future use
    pub(crate) fn has_media(&self) -> bool {
        self.slides.iter().any(|s| !s.media().is_empty())
    }

    // ========================================================================
    // Chart and SmartArt Management
    // ========================================================================

    /// Register chart parts and return the chart index.
    ///
    /// This method generates all the XML and Excel data needed for a chart
    /// and stores them for later serialization when the presentation is saved.
    ///
    /// # Arguments
    /// * `chart_data` - The chart data to register
    ///
    /// # Returns
    /// * `Ok(u32)` - The unique chart index (1, 2, 3, ...)
    /// * `Err` if chart generation fails
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::{Package, ChartData, ChartSeries, ChartType};
    ///
    /// let mut pkg = Package::new().unwrap();
    /// let pres = pkg.presentation_mut().unwrap();
    ///
    /// let chart = ChartData::new(ChartType::Column, 914400, 1600000, 7315200, 4000000)
    ///     .with_title("Sales")
    ///     .add_series(ChartSeries::new("2024").with_values(vec![100.0, 150.0, 200.0]));
    ///
    /// let chart_idx = pres.add_chart_parts(&chart).unwrap();
    /// ```
    pub fn add_chart_parts(
        &mut self,
        chart_data: &crate::ooxml::pptx::parts::chart::ChartData,
    ) -> Result<u32> {
        use super::excel_embed::generate_chart_excel_data;
        use crate::ooxml::pptx::parts::chart::generate_chart_xml;

        // Generate chart XML
        let chart_xml = generate_chart_xml(chart_data);

        // Generate embedded Excel data
        let excel_data = generate_chart_excel_data(chart_data)?;

        // Allocate unique index
        let chart_idx = self.next_chart_idx;
        self.next_chart_idx += 1;

        // Store the parts
        self.charts.insert(
            chart_idx,
            ChartParts {
                chart_xml,
                colors_xml: None, // Optional, not generated for now
                style_xml: None,  // Optional, not generated for now
                excel_data,
            },
        );

        self.modified = true;
        Ok(chart_idx)
    }

    /// Register SmartArt parts and return the diagram index.
    ///
    /// This method generates all four XML parts needed for a SmartArt diagram
    /// and stores them for later serialization when the presentation is saved.
    ///
    /// # Arguments
    /// * `smartart` - The SmartArt diagram to register
    ///
    /// # Returns
    /// * `Ok(u32)` - The unique diagram index (1, 2, 3, ...)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::{Package, SmartArtBuilder, DiagramType};
    ///
    /// let mut pkg = Package::new().unwrap();
    /// let pres = pkg.presentation_mut().unwrap();
    ///
    /// let smartart = SmartArtBuilder::new(DiagramType::List)
    ///     .add_items(vec!["Item 1", "Item 2", "Item 3"])
    ///     .build();
    ///
    /// let diagram_idx = pres.add_smartart_parts(&smartart).unwrap();
    /// ```
    pub fn add_smartart_parts(
        &mut self,
        smartart: &crate::ooxml::pptx::smartart::SmartArt,
    ) -> Result<u32> {
        use crate::ooxml::pptx::smartart::{
            generate_smartart_colors_xml, generate_smartart_data_xml, generate_smartart_layout_xml,
            generate_smartart_quickstyle_xml,
        };

        // Generate all four diagram parts
        let data_xml = generate_smartart_data_xml(smartart);
        let layout_xml = generate_smartart_layout_xml(smartart);
        let style_xml = generate_smartart_quickstyle_xml();
        let colors_xml = generate_smartart_colors_xml();

        // Allocate unique index
        let diagram_idx = self.next_smartart_idx;
        self.next_smartart_idx += 1;

        // Store the parts (including SmartArt data for drawing generation during save)
        self.smartarts.insert(
            diagram_idx,
            SmartArtParts {
                data_xml,
                layout_xml,
                style_xml,
                colors_xml,
                smartart: smartart.clone(),
            },
        );

        self.modified = true;
        Ok(diagram_idx)
    }

    /// Get the number of charts in the presentation.
    #[allow(dead_code)] // Public API for future use
    pub fn chart_count(&self) -> usize {
        self.charts.len()
    }

    /// Get the number of SmartArt diagrams in the presentation.
    #[allow(dead_code)] // Public API for future use
    pub fn smartart_count(&self) -> usize {
        self.smartarts.len()
    }

    /// Generate presentation.xml content.
    pub fn generate_presentation_xml(&self) -> Result<String> {
        self.generate_presentation_xml_with_rels(None, None, None)
    }

    /// Generate presentation.xml content with actual relationship IDs.
    ///
    /// # Arguments
    /// * `slide_rel_ids` - Optional vector of relationship IDs for slides (e.g., ["rId5", "rId6", ...])
    ///   If None, will generate default IDs starting at rId2
    /// * `notes_master_rel_id` - Optional relationship ID for notes master
    /// * `handout_rel_id` - Optional relationship ID for handout master
    pub(crate) fn generate_presentation_xml_with_rels(
        &self,
        slide_rel_ids: Option<&[String]>,
        notes_master_rel_id: Option<&str>,
        handout_rel_id: Option<&str>,
    ) -> Result<String> {
        let mut xml = String::with_capacity(2048);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#);

        // Write slide master ID list (required)
        xml.push_str("<p:sldMasterIdLst>");
        xml.push_str(r#"<p:sldMasterId id="2147483648" r:id="rId1"/>"#);
        xml.push_str("</p:sldMasterIdLst>");

        // Write notes master ID list if present (MUST come before handoutMasterIdLst per OOXML spec)
        if let Some(rel_id) = notes_master_rel_id {
            xml.push_str("<p:notesMasterIdLst>");
            xml.push_str(&format!(r#"<p:notesMasterId r:id="{}"/>"#, rel_id));
            xml.push_str("</p:notesMasterIdLst>");
        }

        // Write handout master ID list if present (MUST come after notesMasterIdLst per OOXML spec)
        if let Some(rel_id) = handout_rel_id {
            xml.push_str("<p:handoutMasterIdLst>");
            xml.push_str(&format!(r#"<p:handoutMasterId r:id="{}"/>"#, rel_id));
            xml.push_str("</p:handoutMasterIdLst>");
        }

        // Build slide ID to relationship ID mapping for custom shows
        let mut slide_id_to_rel_id = std::collections::HashMap::new();

        // Write slide ID list
        if !self.slides.is_empty() {
            xml.push_str("<p:sldIdLst>");
            for (index, slide) in self.slides.iter().enumerate() {
                let rel_id = if let Some(ids) = slide_rel_ids {
                    ids.get(index)
                        .map(|s| s.as_str())
                        .unwrap_or("rId2")
                        .to_string()
                } else {
                    // Default behavior: calculate ID (starts at rId2)
                    format!("rId{}", index + 2)
                };

                // Store mapping for custom shows
                slide_id_to_rel_id.insert(slide.slide_id(), rel_id.clone());

                write!(
                    xml,
                    r#"<p:sldId id="{}" r:id="{}"/>"#,
                    slide.slide_id(),
                    rel_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }
            xml.push_str("</p:sldIdLst>");
        }

        // Write slide size
        write!(
            xml,
            r#"<p:sldSz cx="{}" cy="{}"/>"#,
            self.slide_width, self.slide_height
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;

        xml.push_str("<p:notesSz cx=\"6858000\" cy=\"9144000\"/>");

        // Add custom slide shows if any exist (MUST come before defaultTextStyle per XSD schema)
        if !self.custom_shows.is_empty() {
            xml.push_str(&self.custom_shows.to_xml_with_rel_map(&slide_id_to_rel_id));
        }

        // Add default text style (required for proper text rendering)
        xml.push_str(r#"<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>"#);

        // Add 9 levels of paragraph properties as per OOXML spec
        for level in 1..=9 {
            let margin = (level - 1) * 457200;
            write!(
                xml,
                r#"<a:lvl{}pPr marL="{}" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">"#,
                level, margin
            )
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;

            xml.push_str(r#"<a:defRPr sz="1800" kern="1200">"#);
            xml.push_str(r#"<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>"#);
            xml.push_str(r#"<a:latin typeface="+mn-lt"/>"#);
            xml.push_str(r#"<a:ea typeface="+mn-ea"/>"#);
            xml.push_str(r#"<a:cs typeface="+mn-cs"/>"#);
            xml.push_str("</a:defRPr>");

            write!(xml, "</a:lvl{}pPr>", level).map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }

        xml.push_str("</p:defaultTextStyle>");

        // Add protection settings if any protection is enabled (modifyVerifier comes after defaultTextStyle)
        if self.protection.is_protected() {
            xml.push_str(&self.protection.to_xml());
        }

        // Add sections via extension list if any exist (extLst comes last)
        if !self.sections.is_empty() {
            xml.push_str(&self.sections.to_xml()?);
        }

        xml.push_str("</p:presentation>");

        Ok(xml)
    }
}

impl Default for MutablePresentation {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_presentation() {
        let pres = MutablePresentation::new();
        assert_eq!(pres.slide_count(), 0);
        assert_eq!(pres.slide_width(), 9144000);
        assert_eq!(pres.slide_height(), 6858000);
    }

    #[test]
    fn test_add_slide() {
        let mut pres = MutablePresentation::new();
        let _slide = pres.add_slide().unwrap();
        assert_eq!(pres.slide_count(), 1);
        assert!(pres.is_modified());
    }

    #[test]
    fn test_slide_title() {
        let mut pres = MutablePresentation::new();
        let slide = pres.add_slide().unwrap();
        slide.set_title("Test Title");
        assert_eq!(slide.title(), Some("Test Title"));
    }

    #[test]
    fn test_add_text_box() {
        let mut pres = MutablePresentation::new();
        let slide = pres.add_slide().unwrap();
        slide.add_text_box("Hello", 100, 100, 500, 200);
        assert_eq!(slide.shape_count(), 1);
    }

    #[test]
    fn test_xml_generation() {
        let mut pres = MutablePresentation::new();
        pres.add_slide().unwrap().set_title("Test");

        let xml = pres.generate_presentation_xml().unwrap();
        assert!(xml.contains("<p:presentation"));
        assert!(xml.contains("<p:sldIdLst>"));

        let slide_xml = pres.slides[0].to_xml().unwrap();
        assert!(slide_xml.contains("<p:sld"));
        assert!(slide_xml.contains("Test"));
    }

    #[test]
    fn test_add_chart_parts() {
        use crate::ooxml::pptx::parts::chart::{ChartData, ChartSeries, ChartType};

        let mut pres = MutablePresentation::new();

        let chart = ChartData::new(ChartType::Bar, 0, 0, 100, 100)
            .add_series(ChartSeries::new("Test").with_values(vec![1.0, 2.0]));

        let idx1 = pres.add_chart_parts(&chart).unwrap();
        let idx2 = pres.add_chart_parts(&chart).unwrap();
        let idx3 = pres.add_chart_parts(&chart).unwrap();

        assert_eq!(idx1, 1);
        assert_eq!(idx2, 2);
        assert_eq!(idx3, 3);
        assert_eq!(pres.chart_count(), 3);
    }

    #[test]
    fn test_add_smartart_parts() {
        use crate::ooxml::pptx::smartart::{DiagramType, SmartArtBuilder};

        let mut pres = MutablePresentation::new();

        let smartart = SmartArtBuilder::new(DiagramType::List)
            .add_item("Item 1")
            .add_item("Item 2")
            .build();

        let idx1 = pres.add_smartart_parts(&smartart).unwrap();
        let idx2 = pres.add_smartart_parts(&smartart).unwrap();

        assert_eq!(idx1, 1);
        assert_eq!(idx2, 2);
        assert_eq!(pres.smartart_count(), 2);
    }

    /// **Feature: charts-smartart-integration, Property 6: Unique chart indices across presentation**
    /// **Validates: Requirements 4.1**
    #[cfg(test)]
    mod property_tests {
        use super::*;
        use crate::ooxml::pptx::parts::chart::{ChartData, ChartSeries, ChartType};
        use crate::ooxml::pptx::smartart::{DiagramType, SmartArtBuilder};
        use proptest::prelude::*;
        use std::collections::HashSet;

        proptest! {
            #![proptest_config(ProptestConfig::with_cases(100))]

            /// **Feature: charts-smartart-integration, Property 6: Unique chart indices**
            /// **Validates: Requirements 4.1**
            #[test]
            fn prop_unique_chart_indices(num_charts in 1usize..20) {
                let mut pres = MutablePresentation::new();
                let mut indices = HashSet::new();

                let chart = ChartData::new(ChartType::Bar, 0, 0, 100, 100)
                    .add_series(ChartSeries::new("Test").with_values(vec![1.0]));

                for _ in 0..num_charts {
                    let idx = pres.add_chart_parts(&chart).unwrap();
                    prop_assert!(
                        indices.insert(idx),
                        "Duplicate chart index: {}",
                        idx
                    );
                }

                prop_assert_eq!(indices.len(), num_charts);
            }

            /// **Feature: charts-smartart-integration, Property 7: Unique SmartArt indices**
            /// **Validates: Requirements 4.2**
            #[test]
            fn prop_unique_smartart_indices(num_diagrams in 1usize..20) {
                let mut pres = MutablePresentation::new();
                let mut indices = HashSet::new();

                let smartart = SmartArtBuilder::new(DiagramType::List)
                    .add_item("Item")
                    .build();

                for _ in 0..num_diagrams {
                    let idx = pres.add_smartart_parts(&smartart).unwrap();
                    prop_assert!(
                        indices.insert(idx),
                        "Duplicate SmartArt index: {}",
                        idx
                    );
                }

                prop_assert_eq!(indices.len(), num_diagrams);
            }

            /// **Feature: charts-smartart-integration, Property 1: Chart shape creation preserves data**
            /// **Validates: Requirements 1.1**
            #[test]
            fn prop_chart_shape_preserves_position(
                x in 0i64..10000000i64,
                y in 0i64..10000000i64,
                width in 100i64..10000000i64,
                height in 100i64..10000000i64,
            ) {
                let mut pres = MutablePresentation::new();
                let slide = pres.add_slide().unwrap();

                let chart = ChartData::new(ChartType::Bar, x, y, width, height)
                    .add_series(ChartSeries::new("Test").with_values(vec![1.0]));

                let chart_idx = pres.add_chart_parts(&chart).unwrap();
                let slide = pres.slide_mut(0).unwrap();
                let shape_id = slide.add_chart_shape(chart_idx, x, y, width, height);

                // Shape ID should be valid (positive)
                prop_assert!(shape_id > 0, "Invalid shape ID: {}", shape_id);

                // Slide should have one shape
                prop_assert_eq!(slide.shape_count(), 1);
            }

            /// **Feature: charts-smartart-integration, Property 2: SmartArt shape creation preserves data**
            /// **Validates: Requirements 2.1**
            #[test]
            fn prop_smartart_shape_preserves_position(
                x in 0i64..10000000i64,
                y in 0i64..10000000i64,
                width in 100i64..10000000i64,
                height in 100i64..10000000i64,
            ) {
                let mut pres = MutablePresentation::new();
                let slide = pres.add_slide().unwrap();

                let smartart = SmartArtBuilder::new(DiagramType::List)
                    .add_item("Item")
                    .build();

                let diagram_idx = pres.add_smartart_parts(&smartart).unwrap();
                let slide = pres.slide_mut(0).unwrap();
                let shape_id = slide.add_smartart_shape(diagram_idx, x, y, width, height);

                // Shape ID should be valid (positive)
                prop_assert!(shape_id > 0, "Invalid shape ID: {}", shape_id);

                // Slide should have one shape
                prop_assert_eq!(slide.shape_count(), 1);
            }

            /// **Feature: charts-smartart-integration, Property 10: Slide modified flag is set**
            /// **Validates: Requirements 5.3**
            #[test]
            fn prop_modified_flag_set_on_chart_add(num_charts in 1usize..5) {
                let mut pres = MutablePresentation::new();
                pres.add_slide().unwrap();

                let chart = ChartData::new(ChartType::Bar, 0, 0, 100, 100)
                    .add_series(ChartSeries::new("Test").with_values(vec![1.0]));

                for _ in 0..num_charts {
                    let chart_idx = pres.add_chart_parts(&chart).unwrap();
                    let slide = pres.slide_mut(0).unwrap();
                    slide.add_chart_shape(chart_idx, 0, 0, 100, 100);
                }

                // Presentation should be modified
                prop_assert!(pres.is_modified());

                // Slide should be modified
                let slide = pres.slide_mut(0).unwrap();
                prop_assert!(slide.is_modified());
            }

            /// **Feature: charts-smartart-integration, Property 8: Relationship IDs are unique per slide**
            /// **Validates: Requirements 3.3, 4.3**
            ///
            /// This test verifies that when multiple charts and SmartArt are added to the same slide,
            /// all relationship IDs assigned are unique within that slide's relationship set.
            #[test]
            fn prop_unique_relationship_ids_per_slide(
                num_charts in 0usize..5,
                num_smartarts in 0usize..5,
            ) {
                // Skip if no elements to add
                prop_assume!(num_charts > 0 || num_smartarts > 0);

                let mut pres = MutablePresentation::new();
                pres.add_slide().unwrap();

                // Add charts
                let chart = ChartData::new(ChartType::Bar, 0, 0, 100, 100)
                    .add_series(ChartSeries::new("Test").with_values(vec![1.0]));

                let mut chart_indices = Vec::new();
                for _ in 0..num_charts {
                    let chart_idx = pres.add_chart_parts(&chart).unwrap();
                    chart_indices.push(chart_idx);
                    let slide = pres.slide_mut(0).unwrap();
                    slide.add_chart_shape(chart_idx, 0, 0, 100, 100);
                }

                // Add SmartArt
                let smartart = SmartArtBuilder::new(DiagramType::List)
                    .add_item("Item")
                    .build();

                let mut smartart_indices = Vec::new();
                for _ in 0..num_smartarts {
                    let diagram_idx = pres.add_smartart_parts(&smartart).unwrap();
                    smartart_indices.push(diagram_idx);
                    let slide = pres.slide_mut(0).unwrap();
                    slide.add_smartart_shape(diagram_idx, 0, 0, 100, 100);
                }

                // Simulate relationship ID assignment (as done in package.rs during save)
                // Each chart gets 1 relationship ID, each SmartArt gets 4 relationship IDs
                let mut rel_mapper = crate::ooxml::pptx::writer::relmap::RelationshipMapper::new();
                let mut next_rel_id = 2u32; // Start at 2 (1 is for slide layout)

                // Assign chart relationship IDs
                for chart_idx in &chart_indices {
                    let rid = format!("rId{}", next_rel_id);
                    next_rel_id += 1;
                    rel_mapper.add_chart(0, *chart_idx, rid);
                }

                // Assign SmartArt relationship IDs (4 per SmartArt)
                for diagram_idx in &smartart_indices {
                    let data_rid = format!("rId{}", next_rel_id);
                    next_rel_id += 1;
                    let layout_rid = format!("rId{}", next_rel_id);
                    next_rel_id += 1;
                    let style_rid = format!("rId{}", next_rel_id);
                    next_rel_id += 1;
                    let colors_rid = format!("rId{}", next_rel_id);
                    next_rel_id += 1;
                    rel_mapper.add_smartart(0, *diagram_idx, data_rid, layout_rid, style_rid, colors_rid);
                }

                // Collect all relationship IDs and verify uniqueness
                let mut all_rel_ids = HashSet::new();

                for chart_idx in &chart_indices {
                    if let Some(rid) = rel_mapper.get_chart_id(0, *chart_idx) {
                        prop_assert!(
                            all_rel_ids.insert(rid.to_string()),
                            "Duplicate chart relationship ID: {}",
                            rid
                        );
                    }
                }

                for diagram_idx in &smartart_indices {
                    if let Some((data_rid, layout_rid, style_rid, colors_rid)) =
                        rel_mapper.get_smartart_ids(0, *diagram_idx)
                    {
                        prop_assert!(
                            all_rel_ids.insert(data_rid.to_string()),
                            "Duplicate SmartArt data relationship ID: {}",
                            data_rid
                        );
                        prop_assert!(
                            all_rel_ids.insert(layout_rid.to_string()),
                            "Duplicate SmartArt layout relationship ID: {}",
                            layout_rid
                        );
                        prop_assert!(
                            all_rel_ids.insert(style_rid.to_string()),
                            "Duplicate SmartArt style relationship ID: {}",
                            style_rid
                        );
                        prop_assert!(
                            all_rel_ids.insert(colors_rid.to_string()),
                            "Duplicate SmartArt colors relationship ID: {}",
                            colors_rid
                        );
                    }
                }

                // Verify we have the expected number of unique relationship IDs
                let expected_count = num_charts + (num_smartarts * 4);
                prop_assert_eq!(
                    all_rel_ids.len(),
                    expected_count,
                    "Expected {} unique relationship IDs, got {}",
                    expected_count,
                    all_rel_ids.len()
                );
            }
        }
    }
}
