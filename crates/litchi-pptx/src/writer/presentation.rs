//! Mutable presentation orchestration.

use crate::error::{Error, Result};
use crate::shape::designer::TAGS_EXTENSION_URI;

use super::shape::escape_xml;
use super::slide::{MutableSlide, PreparedDesigner as PreparedSlideDesigner};

#[cfg(feature = "automatic-fonts")]
use litchi_fonts::{CollectGlyphs, GlyphMap};

/// First legal slide ID used by the writer.
pub const FIRST_SLIDE_ID: u32 = 256;

/// Mutable presentation model for new-package authoring.
#[derive(Debug, Clone)]
pub struct MutablePresentation {
    pub(crate) slides: Vec<MutableSlide>,
    pub(crate) slide_width: i64,
    pub(crate) slide_height: i64,
    pub(crate) modified: bool,
}

impl Default for MutablePresentation {
    fn default() -> Self {
        Self::new()
    }
}

impl MutablePresentation {
    /// Create an empty 4:3 presentation model.
    #[must_use]
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            slide_width: 9_144_000,
            slide_height: 6_858_000,
            modified: false,
        }
    }

    fn allocate_slide_id(&self) -> Result<u32> {
        self.slides
            .iter()
            .map(MutableSlide::slide_id)
            .max()
            .map_or(Ok(FIRST_SLIDE_ID), |id| {
                id.checked_add(1)
                    .ok_or_else(|| Error::Invalid("presentation slide ID overflow".into()))
            })
    }

    /// Append an empty slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_slide(&mut self) -> Result<&mut MutableSlide> {
        let id = self.allocate_slide_id()?;
        let index = self.slides.len();
        self.slides.push(MutableSlide::new(id));
        self.modified = true;
        self.slides
            .get_mut(index)
            .ok_or(Error::SlideIndexOutOfBounds { index, len: index })
    }

    /// Insert an empty slide at an ordered position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn insert_slide(&mut self, index: usize) -> Result<&mut MutableSlide> {
        if index > self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index,
                len: self.slides.len(),
            });
        }
        let id = self.allocate_slide_id()?;
        self.slides.insert(index, MutableSlide::new(id));
        self.modified = true;
        let len = self.slides.len();
        self.slides
            .get_mut(index)
            .ok_or(Error::SlideIndexOutOfBounds { index, len })
    }

    /// Number of authored slides.
    #[inline]
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Borrow one mutable slide by index.
    pub fn slide_mut(&mut self, index: usize) -> Option<&mut MutableSlide> {
        self.slides.get_mut(index)
    }

    /// Borrow all authored slides.
    #[must_use]
    pub fn slides(&self) -> &[MutableSlide] {
        &self.slides
    }

    /// Delete one slide by index.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn delete_slide(&mut self, index: usize) -> Result<()> {
        if index >= self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index,
                len: self.slides.len(),
            });
        }
        self.slides.remove(index);
        self.modified = true;
        Ok(())
    }

    /// Duplicate a slide and append it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn duplicate_slide(&mut self, index: usize) -> Result<usize> {
        if index >= self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index,
                len: self.slides.len(),
            });
        }
        let mut duplicate = self.slides[index].clone();
        duplicate.set_slide_id(self.allocate_slide_id()?);
        self.slides.push(duplicate);
        self.modified = true;
        Ok(self.slides.len() - 1)
    }

    /// Duplicate a slide at an ordered position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn insert_duplicate_slide(&mut self, index: usize, position: usize) -> Result<usize> {
        if index >= self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index,
                len: self.slides.len(),
            });
        }
        if position > self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index: position,
                len: self.slides.len(),
            });
        }
        let mut duplicate = self.slides[index].clone();
        duplicate.set_slide_id(self.allocate_slide_id()?);
        self.slides.insert(position, duplicate);
        self.modified = true;
        Ok(position)
    }

    /// Move one slide within the ordered list.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<()> {
        if from_index >= self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index: from_index,
                len: self.slides.len(),
            });
        }
        if to_index >= self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index: to_index,
                len: self.slides.len(),
            });
        }
        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        self.modified |= from_index != to_index;
        Ok(())
    }

    /// Current slide width in EMUs.
    #[must_use]
    pub fn slide_width(&self) -> i64 {
        self.slide_width
    }

    /// Current slide height in EMUs.
    #[must_use]
    pub fn slide_height(&self) -> i64 {
        self.slide_height
    }

    /// Set slide dimensions in EMUs.
    pub fn set_slide_size(&mut self, width: i64, height: i64) {
        self.slide_width = width;
        self.slide_height = height;
        self.modified = true;
    }

    /// Set the standard 4:3 slide dimensions.
    pub fn set_standard_slide_size(&mut self) {
        self.set_slide_size(9_144_000, 6_858_000);
    }

    /// Set the standard 16:9 slide dimensions.
    pub fn set_widescreen_slide_size(&mut self) {
        self.set_slide_size(9_144_000, 5_143_500);
    }

    /// Whether any mutable presentation state needs publication.
    pub fn is_modified(&self) -> bool {
        self.modified || self.slides.iter().any(MutableSlide::is_modified)
    }

    pub(crate) fn mark_clean(&mut self) {
        self.modified = false;
        for slide in &mut self.slides {
            slide.mark_clean();
        }
    }

    /// Generate presentation XML using conventional `rId2..` slide links.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn generate_presentation_xml(&self) -> Result<String> {
        let relationship_ids: Vec<_> = (0..self.slides.len())
            .map(|index| format!("rId{}", index.saturating_add(2)))
            .collect();
        self.generate_presentation_xml_with(&relationship_ids)
    }

    pub(crate) fn preflight_designer(&self) -> Result<PreparedDesigner> {
        let mut slides = Vec::new();
        slides
            .try_reserve_exact(self.slides.len())
            .map_err(|source| Error::Allocation {
                resource: "mutable presentation Designer metadata",
                source,
            })?;
        for slide in &self.slides {
            slides.push(slide.preflight_designer()?);
        }
        Ok(PreparedDesigner { slides })
    }

    pub(crate) fn generate_presentation_xml_with(
        &self,
        relationship_ids: &[String],
    ) -> Result<String> {
        let designer = self.preflight_designer()?;
        self.generate_presentation_xml_with_designer(relationship_ids, &designer)
    }

    pub(crate) fn generate_presentation_xml_with_designer(
        &self,
        relationship_ids: &[String],
        designer: &PreparedDesigner,
    ) -> Result<String> {
        if relationship_ids.len() != self.slides.len() {
            return Err(Error::Invalid(
                "presentation relationship count does not match slide count".into(),
            ));
        }
        if designer.slides.len() != self.slides.len() {
            return Err(Error::Invalid(
                "precompiled Designer slide count does not match presentation slide count".into(),
            ));
        }
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rIdNotesMaster"/></p:notesMasterIdLst><p:sldIdLst>"#,
        );
        for ((slide, relationship_id), prepared) in self
            .slides
            .iter()
            .zip(relationship_ids)
            .zip(&designer.slides)
        {
            xml.push_str("<p:sldId id=\"");
            xml.push_str(&slide.slide_id().to_string());
            xml.push_str("\" r:id=\"");
            xml.push_str(&escape_xml(relationship_id));
            if let Some(tags) = prepared.tags.as_deref() {
                xml.push_str("\"><p:extLst><p:ext uri=\"");
                xml.push_str(TAGS_EXTENSION_URI);
                xml.push_str("\">");
                xml.push_str(tags);
                xml.push_str("</p:ext></p:extLst></p:sldId>");
            } else {
                xml.push_str("\"/>");
            }
        }
        xml.push_str("</p:sldIdLst><p:sldSz cx=\"");
        xml.push_str(&self.slide_width.to_string());
        xml.push_str("\" cy=\"");
        xml.push_str(&self.slide_height.to_string());
        xml.push_str("\" type=\"screen4x3\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/><p:defaultTextStyle/></p:presentation>");
        Ok(xml)
    }
}

#[derive(Debug)]
pub(crate) struct PreparedDesigner {
    slides: Vec<PreparedSlideDesigner>,
}

impl PreparedDesigner {
    pub(crate) fn slides(&self) -> &[PreparedSlideDesigner] {
        &self.slides
    }
}

#[cfg(feature = "automatic-fonts")]
impl CollectGlyphs for MutablePresentation {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut glyphs = GlyphMap::new();
        for slide in &self.slides {
            for (request, used) in slide.collect_glyphs() {
                *glyphs.entry(request).or_default() |= used;
            }
        }
        glyphs
    }
}
