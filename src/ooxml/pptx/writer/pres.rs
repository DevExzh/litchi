/// Presentation writer for PPTX.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
use super::super::format::ImageFormat;
use super::slide::MutableSlide;

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
    /// Whether the presentation has been modified
    modified: bool,
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
            modified: false,
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

    /// Generate presentation.xml content.
    pub fn generate_presentation_xml(&self) -> Result<String> {
        self.generate_presentation_xml_with_rels(None)
    }

    /// Generate presentation.xml content with actual relationship IDs.
    ///
    /// # Arguments
    /// * `slide_rel_ids` - Optional vector of relationship IDs for slides (e.g., ["rId5", "rId6", ...])
    ///   If None, will generate default IDs starting at rId2
    pub(crate) fn generate_presentation_xml_with_rels(
        &self,
        slide_rel_ids: Option<&[String]>,
    ) -> Result<String> {
        let mut xml = String::with_capacity(2048);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#);

        // Write slide master ID list
        xml.push_str("<p:sldMasterIdLst>");
        xml.push_str(r#"<p:sldMasterId id="2147483648" r:id="rId1"/>"#);
        xml.push_str("</p:sldMasterIdLst>");

        // Write slide ID list
        if !self.slides.is_empty() {
            xml.push_str("<p:sldIdLst>");
            for (index, slide) in self.slides.iter().enumerate() {
                if let Some(ids) = slide_rel_ids {
                    let rel_id = ids.get(index).map(|s| s.as_str()).unwrap_or("rId2");
                    write!(
                        xml,
                        r#"<p:sldId id="{}" r:id="{}"/>"#,
                        slide.slide_id(),
                        rel_id
                    )
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                } else {
                    // Default behavior: calculate ID (starts at rId2)
                    // This is used for testing or when explicit IDs aren't needed
                    write!(
                        xml,
                        r#"<p:sldId id="{}" r:id="rId{}"/>"#,
                        slide.slide_id(),
                        index + 2
                    )
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
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
}
