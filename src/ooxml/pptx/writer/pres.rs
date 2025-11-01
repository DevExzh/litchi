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

    /// Generate presentation.xml content.
    pub fn generate_presentation_xml(&self) -> Result<String> {
        self.generate_presentation_xml_with_rels(None)
    }

    /// Generate presentation.xml content with actual relationship IDs.
    ///
    /// # Arguments
    /// * `slide_rel_ids` - Optional vector of relationship IDs for slides (e.g., ["rId5", "rId6", ...])
    ///   If None, will calculate IDs assuming slides start at rId2
    pub(crate) fn generate_presentation_xml_with_rels(
        &self,
        slide_rel_ids: Option<&[String]>,
    ) -> Result<String> {
        let mut xml = String::with_capacity(2048);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);

        // Write slide master ID list
        xml.push_str("<p:sldMasterIdLst>");
        xml.push_str(r#"<p:sldMasterId id="2147483648" r:id="rId1"/>"#);
        xml.push_str("</p:sldMasterIdLst>");

        // Write slide ID list
        if !self.slides.is_empty() {
            xml.push_str("<p:sldIdLst>");
            for (index, slide) in self.slides.iter().enumerate() {
                let rel_id = if let Some(ids) = slide_rel_ids {
                    ids.get(index).map(|s| s.as_str()).unwrap_or("rId2") // Fallback, shouldn't happen
                } else {
                    // Legacy behavior: calculate ID (starts at rId2)
                    return Err(OoxmlError::Xml(
                        "Slide relationship IDs must be provided".to_string(),
                    ));
                };

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
