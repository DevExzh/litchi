/// Slide types and implementation for PPTX presentations.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
use super::super::format::ImageFormat;
use super::shape::MutableShape;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// A mutable slide in a presentation.
#[derive(Debug, Clone)]
pub struct MutableSlide {
    /// Slide ID (unique identifier)
    pub(crate) slide_id: u32,
    /// Slide title (stored in title placeholder)
    pub(crate) title: Option<String>,
    /// Shapes on the slide
    pub(crate) shapes: Vec<MutableShape>,
    /// Speaker notes for the slide
    pub(crate) notes: Option<String>,
    /// Slide transition effect
    pub(crate) transition: Option<crate::ooxml::pptx::transitions::SlideTransition>,
    /// Slide background
    pub(crate) background: Option<crate::ooxml::pptx::backgrounds::SlideBackground>,
    /// Whether the slide has been modified
    pub(crate) modified: bool,
}

impl MutableSlide {
    /// Create a new empty slide.
    pub(crate) fn new(slide_id: u32) -> Self {
        Self {
            slide_id,
            title: None,
            shapes: Vec::new(),
            notes: None,
            transition: None,
            background: None,
            modified: false,
        }
    }

    /// Get the slide ID.
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Set the slide ID.
    ///
    /// This is used internally when duplicating slides to assign new IDs.
    pub(crate) fn set_slide_id(&mut self, slide_id: u32) {
        self.slide_id = slide_id;
        self.modified = true;
    }

    /// Set the slide title.
    pub fn set_title(&mut self, title: &str) {
        self.title = Some(title.to_string());
        self.modified = true;
    }

    /// Get the slide title.
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Set speaker notes for the slide.
    pub fn set_notes(&mut self, notes: &str) {
        self.notes = Some(notes.to_string());
        self.modified = true;
    }

    /// Get the speaker notes for the slide.
    pub fn notes(&self) -> Option<&str> {
        self.notes.as_deref()
    }

    /// Check if the slide has speaker notes.
    pub fn has_notes(&self) -> bool {
        self.notes.is_some()
    }

    /// Set a transition effect for the slide.
    ///
    /// # Arguments
    /// * `transition` - The transition configuration
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{MutablePresentation, TransitionType, TransitionSpeed, SlideTransition};
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    ///
    /// // Add a fade transition
    /// let transition = SlideTransition::new(TransitionType::Fade)
    ///     .with_speed(TransitionSpeed::Fast)
    ///     .with_advance_after_ms(3000);
    /// slide.set_transition(transition);
    /// ```
    pub fn set_transition(&mut self, transition: crate::ooxml::pptx::transitions::SlideTransition) {
        self.transition = Some(transition);
        self.modified = true;
    }

    /// Get the transition effect for the slide.
    ///
    /// Returns `None` if no transition is set.
    pub fn transition(&self) -> Option<&crate::ooxml::pptx::transitions::SlideTransition> {
        self.transition.as_ref()
    }

    /// Remove the transition effect from the slide.
    pub fn remove_transition(&mut self) {
        self.transition = None;
        self.modified = true;
    }

    /// Set a background for the slide.
    ///
    /// # Arguments
    /// * `background` - The background configuration
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{MutablePresentation, SlideBackground};
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    ///
    /// // Set a solid blue background
    /// slide.set_background(SlideBackground::solid("4472C4"));
    /// ```
    pub fn set_background(&mut self, background: crate::ooxml::pptx::backgrounds::SlideBackground) {
        self.background = Some(background);
        self.modified = true;
    }

    /// Get the background for the slide.
    ///
    /// Returns `None` if no background is set.
    pub fn background(&self) -> Option<&crate::ooxml::pptx::backgrounds::SlideBackground> {
        self.background.as_ref()
    }

    /// Remove the background from the slide (use master background).
    pub fn remove_background(&mut self) {
        self.background = None;
        self.modified = true;
    }

    /// Add a text box to the slide.
    pub fn add_text_box(&mut self, text: &str, x: i64, y: i64, width: i64, height: i64) {
        let shape_id = (self.shapes.len() + 2) as u32;
        let shape = MutableShape::new_text_box(shape_id, text.to_string(), x, y, width, height);
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add a rectangle to the slide.
    pub fn add_rectangle(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) {
        let shape_id = (self.shapes.len() + 2) as u32;
        let shape = MutableShape::new_rectangle(shape_id, x, y, width, height, fill_color);
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add an ellipse (circle/oval) to the slide.
    pub fn add_ellipse(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) {
        let shape_id = (self.shapes.len() + 2) as u32;
        let shape = MutableShape::new_ellipse(shape_id, x, y, width, height, fill_color);
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add a picture to the slide from a file.
    pub fn add_picture(
        &mut self,
        image_path: &str,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> Result<()> {
        use std::fs;

        let data = fs::read(image_path).map_err(OoxmlError::IoError)?;

        let format = ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| OoxmlError::InvalidFormat("Unknown image format".to_string()))?;

        let shape_id = (self.shapes.len() + 2) as u32;
        let description = format!("Picture from {}", image_path);
        let shape =
            MutableShape::new_picture(shape_id, data, format, x, y, width, height, description)?;
        self.shapes.push(shape);
        self.modified = true;

        Ok(())
    }

    /// Add a picture to the slide from bytes.
    pub fn add_picture_from_bytes(
        &mut self,
        data: Vec<u8>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        description: Option<String>,
    ) -> Result<()> {
        let format = ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| OoxmlError::InvalidFormat("Unknown image format".to_string()))?;

        let shape_id = (self.shapes.len() + 2) as u32;
        let desc = description.unwrap_or_else(|| "Picture".to_string());
        let shape = MutableShape::new_picture(shape_id, data, format, x, y, width, height, desc)?;
        self.shapes.push(shape);
        self.modified = true;

        Ok(())
    }

    /// Get the number of shapes on the slide.
    pub fn shape_count(&self) -> usize {
        self.shapes.len()
    }

    /// Check if the slide has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified
    }

    /// Collect all images from this slide (from shapes only, not background).
    pub(crate) fn collect_images(&self) -> Vec<(&[u8], ImageFormat)> {
        let mut images = Vec::new();

        for shape in &self.shapes {
            if let Some((data, format)) = shape.get_image_data() {
                images.push((data, format));
            }
        }

        images
    }

    /// Get the background image if this slide has a picture background.
    ///
    /// Returns `Some((image_data, format))` if the background is a picture,
    /// otherwise returns `None`.
    pub(crate) fn get_background_image(&self) -> Option<(&[u8], ImageFormat)> {
        self.background
            .as_ref()
            .and_then(|bg| bg.get_image_data())
            .map(|(data, &format)| (data, format))
    }

    /// Generate slide XML content.
    #[allow(dead_code)] // Public API but not used in the current implementation
    pub(crate) fn to_xml(&self) -> Result<String> {
        self.to_xml_with_rels(None, None)
    }

    /// Generate slide XML content with relationship IDs from the mapper.
    ///
    /// # Arguments
    /// * `slide_index` - The index of this slide (used to look up relationships)
    /// * `rel_mapper` - The relationship mapper containing actual relationship IDs
    pub(crate) fn to_xml_with_rels(
        &self,
        slide_index: Option<usize>,
        rel_mapper: Option<&crate::ooxml::pptx::writer::relmap::RelationshipMapper>,
    ) -> Result<String> {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);

        xml.push_str(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
        );
        xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#);
        xml.push_str(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        xml.push_str("<p:cSld>");

        // Add background if present (must come BEFORE spTree per OOXML spec)
        if let Some(ref background) = self.background {
            // For picture backgrounds, we need to get the relationship ID
            let bg_rel_id = if background.get_image_data().is_some() {
                // Get actual relationship ID from mapper
                slide_index.and_then(|si| rel_mapper.and_then(|rm| rm.get_background_id(si)))
            } else {
                None
            };
            xml.push_str(&background.to_xml(bg_rel_id)?);
        }

        xml.push_str("<p:spTree>");

        // Write group shape properties (required)
        xml.push_str("<p:nvGrpSpPr>");
        xml.push_str(r#"<p:cNvPr id="1" name=""/>"#);
        xml.push_str("<p:cNvGrpSpPr/>");
        xml.push_str("<p:nvPr/>");
        xml.push_str("</p:nvGrpSpPr>");
        xml.push_str("<p:grpSpPr>");
        xml.push_str("<a:xfrm>");
        xml.push_str(r#"<a:off x="0" y="0"/>"#);
        xml.push_str(r#"<a:ext cx="0" cy="0"/>"#);
        xml.push_str(r#"<a:chOff x="0" y="0"/>"#);
        xml.push_str(r#"<a:chExt cx="0" cy="0"/>"#);
        xml.push_str("</a:xfrm>");
        xml.push_str("</p:grpSpPr>");

        // Write title placeholder if title is set
        if let Some(ref title) = self.title {
            self.write_title_shape(&mut xml, title)?;
        }

        // Write shapes with relationship IDs
        let mut image_counter = 0;
        for shape in &self.shapes {
            // Check if this shape is an image and get its relationship ID
            let rel_id = if shape.get_image_data().is_some() {
                let rid = slide_index
                    .and_then(|si| rel_mapper.and_then(|rm| rm.get_image_id(si, image_counter)));
                image_counter += 1;
                rid
            } else {
                None
            };
            shape.to_xml(&mut xml, rel_id)?;
        }

        xml.push_str("</p:spTree>");
        xml.push_str("</p:cSld>");

        xml.push_str(r#"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"#);

        // Add transition if present
        if let Some(ref transition) = self.transition {
            xml.push_str(&transition.to_xml()?);
        }

        xml.push_str("</p:sld>");

        Ok(xml)
    }

    /// Generate notes slide XML content.
    pub(crate) fn generate_notes_xml(&self) -> Option<Result<String>> {
        let notes_text = self.notes.as_ref()?;

        let mut xml = String::with_capacity(2048);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);

        xml.push_str(
            r#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
        );
        xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#);
        xml.push_str(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        xml.push_str("<p:cSld>");
        xml.push_str("<p:spTree>");

        // Group shape properties
        xml.push_str("<p:nvGrpSpPr>");
        xml.push_str(r#"<p:cNvPr id="1" name=""/>"#);
        xml.push_str("<p:cNvGrpSpPr/>");
        xml.push_str("<p:nvPr/>");
        xml.push_str("</p:nvGrpSpPr>");
        xml.push_str("<p:grpSpPr>");
        xml.push_str("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>");
        xml.push_str("<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm>");
        xml.push_str("</p:grpSpPr>");

        // Notes text shape
        xml.push_str("<p:sp>");
        xml.push_str("<p:nvSpPr>");
        xml.push_str(r#"<p:cNvPr id="2" name="Notes Placeholder"/>"#);
        xml.push_str("<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>");
        xml.push_str("<p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr>");
        xml.push_str("</p:nvSpPr>");

        xml.push_str("<p:spPr/>");

        xml.push_str("<p:txBody>");
        xml.push_str("<a:bodyPr/>");
        xml.push_str("<a:lstStyle/>");
        xml.push_str("<a:p>");
        xml.push_str("<a:r>");
        xml.push_str("<a:rPr lang=\"en-US\" dirty=\"0\"/>");
        if let Err(e) = write!(xml, "<a:t>{}</a:t>", escape_xml(notes_text)) {
            return Some(Err(OoxmlError::Xml(e.to_string())));
        }
        xml.push_str("</a:r>");
        xml.push_str("</a:p>");
        xml.push_str("</p:txBody>");
        xml.push_str("</p:sp>");

        xml.push_str("</p:spTree>");
        xml.push_str("</p:cSld>");
        xml.push_str(r#"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"#);
        xml.push_str("</p:notes>");

        Some(Ok(xml))
    }

    /// Write the title placeholder shape.
    fn write_title_shape(&self, xml: &mut String, title: &str) -> Result<()> {
        xml.push_str("<p:sp>");
        xml.push_str("<p:nvSpPr>");
        xml.push_str(r#"<p:cNvPr id="1" name="Title 1"/>"#);
        xml.push_str("<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>");
        xml.push_str(r#"<p:nvPr><p:ph type="ctrTitle"/></p:nvPr>"#);
        xml.push_str("</p:nvSpPr>");

        xml.push_str("<p:spPr/>");

        xml.push_str("<p:txBody>");
        xml.push_str("<a:bodyPr/>");
        xml.push_str("<a:lstStyle/>");
        xml.push_str("<a:p>");
        xml.push_str("<a:r>");
        xml.push_str("<a:rPr lang=\"en-US\" dirty=\"0\" smtClean=\"0\"/>");
        write!(xml, "<a:t>{}</a:t>", escape_xml(title))
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</a:r>");
        xml.push_str("</a:p>");
        xml.push_str("</p:txBody>");

        xml.push_str("</p:sp>");

        Ok(())
    }
}
