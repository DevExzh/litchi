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
            slides.push(Slide::new(slide_part));
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

    // TODO: Apache POI features not yet implemented:
    // - Slide manipulation: delete_slide(), duplicate_slide(), move_slide()
    // - Charts: add_chart(), get_charts(), update_chart()
    // - Tables (reading/writing): get_tables(), add_table(), modify_table()
    // - SmartArt: add_smartart(), get_smartart()
    // - Audio/Video: add_audio(), add_video(), get_media()
    // - Animations: add_animation(), get_animations(), set_animation_timing()
    // - Transitions: set_slide_transition(), get_slide_transition()
    // - Comments: add_comment(), get_comments(), delete_comment()
    // - Notes (advanced): get_notes_master(), set_notes_master()
    // - Handout master: get_handout_master(), set_handout_master()
    // - Custom slide shows: add_custom_show(), get_custom_shows()
    // - Slide layout manipulation: get_slide_layouts(), apply_layout()
    // - Master slides manipulation: add_slide_master(), get_all_slide_masters()
    // - Themes (advanced): apply_theme(), get_theme(), modify_theme()
    // - Hyperlinks: add_hyperlink(), get_hyperlinks(), remove_hyperlink()
    // - Group shapes: group_shapes(), ungroup_shapes(), get_grouped_shapes()
    // - Shape formatting: set_shape_fill(), set_shape_line(), get_shape_effects()
    // - Text formatting (rich): set_text_format(), get_text_format(), apply_style()
    // - Placeholder management: get_placeholders(), fill_placeholder()
    // - Slide backgrounds: set_background(), get_background(), remove_background()
    // - Slide size: set_slide_size(), get_slide_size()
    // - Presentation protection: protect_presentation(), unprotect_presentation()
    // - Custom properties: set_custom_property(), get_custom_properties()
    // - Sections: add_section(), get_sections(), rename_section(), delete_section()
    // - Slide timings: set_slide_timing(), get_slide_timing()
}

#[cfg(test)]
mod tests {
    // Tests will be added as implementation progresses
}
