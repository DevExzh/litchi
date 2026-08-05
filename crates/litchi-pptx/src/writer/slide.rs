//! Mutable slide authoring and XML generation.

use crate::backgrounds::SlideBackground;
use crate::format::TextFormat;
use crate::transition::Transition;
use crate::{Error, Result};

use super::shape::{MutableShape, escape_xml};

/// Mutable slide state owned by [`super::MutablePresentation`].
#[derive(Debug, Clone)]
pub struct MutableSlide {
    pub(crate) slide_id: u32,
    pub(crate) title: Option<String>,
    pub(crate) shapes: Vec<MutableShape>,
    pub(crate) notes: Option<String>,
    pub(crate) transition: Option<Transition>,
    pub(crate) background: Option<SlideBackground>,
    pub(crate) modified: bool,
}

impl MutableSlide {
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

    /// The stable slide ID emitted into `p:sldId@id`.
    #[inline]
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    pub(crate) fn set_slide_id(&mut self, slide_id: u32) {
        self.slide_id = slide_id;
        self.modified = true;
    }

    /// Set the slide title.
    pub fn set_title(&mut self, title: &str) {
        self.title = Some(title.to_string());
        self.modified = true;
    }

    /// Return the slide title.
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Set inert speaker-notes text in the mutable model.
    pub fn set_notes(&mut self, notes: &str) {
        self.notes = Some(notes.to_string());
        self.modified = true;
    }

    /// Return mutable-model speaker notes.
    pub fn notes(&self) -> Option<&str> {
        self.notes.as_deref()
    }

    /// Whether the mutable model contains speaker notes.
    pub fn has_notes(&self) -> bool {
        self.notes.is_some()
    }

    /// Remove speaker notes from the mutable model.
    pub fn clear_notes(&mut self) -> bool {
        let removed = self.notes.take().is_some();
        self.modified |= removed;
        removed
    }

    /// Set the canonical transition value.
    pub fn set_transition(&mut self, transition: Transition) {
        self.transition = Some(transition);
        self.modified = true;
    }

    /// Borrow the canonical transition value.
    pub fn transition(&self) -> Option<&Transition> {
        self.transition.as_ref()
    }

    /// Remove the transition value.
    pub fn remove_transition(&mut self) {
        self.modified |= self.transition.take().is_some();
    }

    /// Set a package-independent background value.
    pub fn set_background(&mut self, background: SlideBackground) {
        self.background = Some(background);
        self.modified = true;
    }

    /// Borrow the package-independent background value.
    pub fn background(&self) -> Option<&SlideBackground> {
        self.background.as_ref()
    }

    /// Remove the background value.
    pub fn remove_background(&mut self) {
        self.modified |= self.background.take().is_some();
    }

    /// Add a text box and return it for formatting.
    pub fn add_text_box(
        &mut self,
        text: &str,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut MutableShape {
        let shape_id = (self.shapes.len() as u32).saturating_add(3);
        self.shapes.push(MutableShape::new_text_box(
            shape_id,
            text.to_string(),
            x,
            y,
            width,
            height,
        ));
        self.modified = true;
        self.shapes.last_mut().expect("shape was just pushed")
    }

    /// Add a filled or unfilled rectangle.
    pub fn add_rectangle(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) {
        let shape_id = (self.shapes.len() as u32).saturating_add(3);
        self.shapes.push(MutableShape::new_rectangle(
            shape_id, x, y, width, height, fill_color,
        ));
        self.modified = true;
    }

    /// Add a filled or unfilled ellipse.
    pub fn add_ellipse(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) {
        let shape_id = (self.shapes.len() as u32).saturating_add(3);
        self.shapes.push(MutableShape::new_ellipse(
            shape_id, x, y, width, height, fill_color,
        ));
        self.modified = true;
    }

    /// Borrow authored shapes in source order.
    pub fn shapes(&self) -> &[MutableShape] {
        &self.shapes
    }

    /// Number of authored shapes, excluding the title convenience shape.
    pub fn shape_count(&self) -> usize {
        self.shapes.len()
    }

    /// Whether this slide needs managed publication.
    pub fn is_modified(&self) -> bool {
        self.modified || self.shapes.iter().any(MutableShape::is_modified)
    }

    pub(crate) fn mark_clean(&mut self) {
        self.modified = false;
        for shape in &mut self.shapes {
            shape.mark_clean();
        }
    }

    /// Generate one complete slide part.
    pub fn generate_slide_xml(&self) -> Result<String> {
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide {}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
            self.slide_id
        );
        if let Some(background) = &self.background {
            xml.push_str(&background.to_xml(None)?);
        }
        if let Some(title) = &self.title {
            let title =
                MutableShape::new_text_box(2, title.clone(), 914400, 457200, 7315200, 914400)
                    .set_text_format(TextFormat::default())
                    .to_xml()?;
            xml.push_str(&title);
        }
        for shape in &self.shapes {
            xml.push_str(&shape.to_xml()?);
        }
        xml.push_str("</p:spTree></p:cSld>");
        if let Some(transition) = &self.transition {
            crate::transition::write_to(transition, &mut xml)?;
        }
        xml.push_str("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
        Ok(xml)
    }

    /// Alias used by package materialization code.
    pub fn to_xml(&self) -> Result<String> {
        self.generate_slide_xml()
    }

    /// Add a text box using a checked `TextFormat` value.
    pub fn add_formatted_text_box(
        &mut self,
        text: &str,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        format: TextFormat,
    ) -> &mut MutableShape {
        self.add_text_box(text, x, y, width, height)
            .set_text_format(format)
    }

    #[allow(dead_code)]
    fn ensure_positive_bounds(&self) -> Result<()> {
        if self.shapes.len() > u32::MAX as usize {
            return Err(Error::Limit {
                resource: "mutable slide shapes",
                limit: u32::MAX as usize,
            });
        }
        Ok(())
    }
}

#[allow(dead_code)]
fn _escape_is_kept_in_this_layer(value: &str) -> String {
    escape_xml(value)
}
