//! Mutable slide authoring and XML generation.

use crate::backgrounds::SlideBackground;
use crate::format::TextFormat;
use crate::shape::designer::{Limits as DesignerLimits, Tags};
use crate::transition::Transition;
use crate::{Error, Result};

use super::shape::{MutableShape, escape_xml};

#[cfg(feature = "automatic-fonts")]
use litchi_fonts::{CollectGlyphs, GlyphMap};

/// Mutable slide state owned by [`super::MutablePresentation`].
#[derive(Debug, Clone)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`MutableSlide` is the established public API name; renaming it would break downstream crates"
)]
pub struct MutableSlide {
    pub(crate) slide_id: u32,
    pub(crate) title: Option<String>,
    pub(crate) shapes: Vec<MutableShape>,
    pub(crate) notes: Option<String>,
    pub(crate) transition: Option<Transition>,
    pub(crate) background: Option<SlideBackground>,
    designer_tags: Option<Tags>,
    designer_limits: DesignerLimits,
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
            designer_tags: None,
            designer_limits: DesignerLimits::default(),
            modified: false,
        }
    }

    /// The stable slide ID emitted into `p:sldId@id`.
    #[inline]
    #[must_use]
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Borrow the optional Designer tags owned by this stable slide ID.
    #[inline]
    #[must_use]
    pub fn designer_tags(&self) -> Option<&Tags> {
        self.designer_tags.as_ref()
    }

    /// Set Designer tags under safe default resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_designer_tags(&mut self, tags: Tags) -> Result<&mut Self> {
        self.set_designer_tags_with_limits(tags, DesignerLimits::default())
    }

    /// Set Designer tags under caller-supplied resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_designer_tags_with_limits(
        &mut self,
        tags: Tags,
        limits: DesignerLimits,
    ) -> Result<&mut Self> {
        // Validate the complete wire value before changing mutable state.
        crate::shape::designer::write_tags(&tags, limits)?;
        // Limits configure validation for future materialization. They are not
        // serialized presentation state and therefore must not make a writer
        // publication dirty by themselves.
        let changed = self.designer_tags.as_ref() != Some(&tags);
        self.designer_tags = Some(tags);
        self.designer_limits = limits;
        self.modified |= changed;
        Ok(self)
    }

    /// Remove Designer tags from this slide ID.
    pub fn clear_designer_tags(&mut self) -> bool {
        let removed = self.designer_tags.take().is_some();
        self.modified |= removed;
        removed
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
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Set inert speaker-notes text in the mutable model.
    pub fn set_notes(&mut self, notes: &str) {
        self.notes = Some(notes.to_string());
        self.modified = true;
    }

    /// Return mutable-model speaker notes.
    #[must_use]
    pub fn notes(&self) -> Option<&str> {
        self.notes.as_deref()
    }

    /// Whether the mutable model contains speaker notes.
    #[must_use]
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
    #[must_use]
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
    #[must_use]
    pub fn background(&self) -> Option<&SlideBackground> {
        self.background.as_ref()
    }

    /// Remove the background value.
    pub fn remove_background(&mut self) {
        self.modified |= self.background.take().is_some();
    }

    /// Add a text box and return it for formatting.
    ///
    /// # Panics
    ///
    /// Never panics; the returned shape was just pushed onto the shape list.
    #[allow(
        clippy::expect_used,
        reason = "the shape list cannot be empty immediately after pushing a shape"
    )]
    pub fn add_text_box(
        &mut self,
        text: &str,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut MutableShape {
        let shape_id = u32::try_from(self.shapes.len())
            .unwrap_or(u32::MAX)
            .saturating_add(3);
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
        self.push_rectangle(x, y, width, height, fill_color);
    }

    /// Add a rectangle with inert Designer drawing properties.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_rectangle_with_designer_properties(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
        properties: crate::shape::designer::DrawingProperties,
    ) -> Result<&mut MutableShape> {
        self.add_rectangle_with_designer_properties_and_limits(
            x,
            y,
            width,
            height,
            fill_color,
            properties,
            DesignerLimits::default(),
        )
    }

    /// Add a rectangle with inert Designer drawing properties and explicit bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_rectangle_with_designer_properties_and_limits(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
        properties: crate::shape::designer::DrawingProperties,
        limits: DesignerLimits,
    ) -> Result<&mut MutableShape> {
        // Reject an invalid value before adding a shape so the operation is
        // atomic from the caller's perspective.
        crate::shape::designer::write_properties(&properties, None, limits)?;
        self.push_rectangle(x, y, width, height, fill_color)
            .set_designer_properties_with_limits(properties, limits)
    }

    #[allow(
        clippy::expect_used,
        reason = "the shape list cannot be empty immediately after pushing a shape"
    )]
    fn push_rectangle(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) -> &mut MutableShape {
        let shape_id = u32::try_from(self.shapes.len())
            .unwrap_or(u32::MAX)
            .saturating_add(3);
        self.shapes.push(MutableShape::new_rectangle(
            shape_id, x, y, width, height, fill_color,
        ));
        self.modified = true;
        self.shapes.last_mut().expect("shape was just pushed")
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
        self.push_ellipse(x, y, width, height, fill_color);
    }

    /// Add an ellipse with inert Designer drawing properties.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_ellipse_with_designer_properties(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
        properties: crate::shape::designer::DrawingProperties,
    ) -> Result<&mut MutableShape> {
        self.add_ellipse_with_designer_properties_and_limits(
            x,
            y,
            width,
            height,
            fill_color,
            properties,
            DesignerLimits::default(),
        )
    }

    /// Add an ellipse with inert Designer drawing properties and explicit bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_ellipse_with_designer_properties_and_limits(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
        properties: crate::shape::designer::DrawingProperties,
        limits: DesignerLimits,
    ) -> Result<&mut MutableShape> {
        // Reject an invalid value before adding a shape so the operation is
        // atomic from the caller's perspective.
        crate::shape::designer::write_properties(&properties, None, limits)?;
        self.push_ellipse(x, y, width, height, fill_color)
            .set_designer_properties_with_limits(properties, limits)
    }

    #[allow(
        clippy::expect_used,
        reason = "the shape list cannot be empty immediately after pushing a shape"
    )]
    fn push_ellipse(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) -> &mut MutableShape {
        let shape_id = u32::try_from(self.shapes.len())
            .unwrap_or(u32::MAX)
            .saturating_add(3);
        self.shapes.push(MutableShape::new_ellipse(
            shape_id, x, y, width, height, fill_color,
        ));
        self.modified = true;
        self.shapes.last_mut().expect("shape was just pushed")
    }

    /// Borrow authored shapes in source order.
    #[must_use]
    pub fn shapes(&self) -> &[MutableShape] {
        &self.shapes
    }

    /// Number of authored shapes, excluding the title convenience shape.
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn generate_slide_xml(&self) -> Result<String> {
        let designer = self.preflight_designer()?;
        self.generate_slide_xml_with(&designer)
    }

    pub(crate) fn preflight_designer(&self) -> Result<PreparedDesigner> {
        let mut shape_properties = Vec::new();
        shape_properties
            .try_reserve_exact(self.shapes.len())
            .map_err(|source| Error::Allocation {
                resource: "mutable slide Designer properties",
                source,
            })?;
        for shape in &self.shapes {
            shape_properties.push(shape.preflight_designer_properties()?);
        }
        let tags = self
            .designer_tags
            .as_ref()
            .map(|tags| {
                crate::shape::designer::write_tags(tags, self.designer_limits).and_then(|bytes| {
                    String::from_utf8(bytes).map_err(|_err| {
                        Error::Invalid("Designer serializer produced non-UTF-8 XML".into())
                    })
                })
            })
            .transpose()?;
        Ok(PreparedDesigner {
            shape_properties,
            tags,
        })
    }

    pub(crate) fn generate_slide_xml_with(&self, designer: &PreparedDesigner) -> Result<String> {
        if designer.shape_properties.len() != self.shapes.len() {
            return Err(Error::Invalid(
                "precompiled Designer shape count does not match slide shape count".into(),
            ));
        }
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide {}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
            self.slide_id
        );
        if let Some(background) = &self.background {
            xml.push_str(&background.to_xml(None)?);
        }
        if let Some(title) = &self.title {
            let title_xml =
                MutableShape::new_text_box(2, title.clone(), 914_400, 457_200, 7_315_200, 914_400)
                    .set_text_format(TextFormat::default())
                    .to_xml()?;
            xml.push_str(&title_xml);
        }
        for (shape, properties) in self.shapes.iter().zip(&designer.shape_properties) {
            xml.push_str(&shape.to_xml_with_designer(properties.as_deref())?);
        }
        xml.push_str("</p:spTree></p:cSld>");
        if let Some(transition) = &self.transition {
            crate::transition::write_to(transition, &mut xml)?;
        }
        xml.push_str("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
        Ok(xml)
    }

    /// Alias used by package materialization code.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
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

    #[allow(dead_code, reason = "guard retained for the staged writer API")]
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

#[derive(Debug)]
pub(crate) struct PreparedDesigner {
    pub(crate) shape_properties: Vec<Option<String>>,
    pub(crate) tags: Option<String>,
}

#[cfg(feature = "automatic-fonts")]
impl CollectGlyphs for MutableSlide {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut glyphs = GlyphMap::new();
        for shape in &self.shapes {
            for (request, used) in shape.collect_glyphs() {
                *glyphs.entry(request).or_default() |= used;
            }
        }
        glyphs
    }
}

#[allow(dead_code, reason = "documents that XML escaping stays in this layer")]
fn _escape_is_kept_in_this_layer(value: &str) -> String {
    escape_xml(value)
}
