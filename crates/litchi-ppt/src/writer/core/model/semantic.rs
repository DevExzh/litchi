//! Contextual authoring operations for the PPT semantic writer model.

#[cfg(feature = "encryption")]
use super::WriterEncryption;
use super::{
    HyperlinkCollection, Pictures, ShapeProperties, ShapeType, TextAlignment, WritableShape,
    WritableSlide, WriteError, Writer, WriterModifyPassword,
};
use crate::animation::AnimationInfo;
#[cfg(feature = "encryption")]
use crate::encryption::{EncryptionProfile, validate_writer_password};
use crate::font::{
    EmbeddedFont, Facet as FontFacet, Font, FontCollection, FontCollections, FontEmbeddingFlags,
    Scope as FontScope,
};
use crate::header_footer::{
    HeaderFooter, HeaderFooterParent, HeaderFooterParentOrdinal, HeaderFooterScope,
};
use crate::modify_password::{ModifyPassword, validate_value as validate_modify_password};
use crate::transition::TransitionInfo;
use crate::writer::blip::{Id as PictureId, Kind as PictureKind};
use crate::writer::chart::Chart;
use crate::writer::comments::SlideComment;
use crate::writer::core::codec::sound_collection_error;
use crate::writer::custom_shows::CustomShow;
use crate::writer::escher::FreeformGeometry;
use crate::writer::hyperlink::Hyperlink;
use crate::writer::notes::NotesPage;
use crate::writer::shape_style::{ArrowStyle, FillStyle, LineStyleConfig, ShapeStyle};
use crate::writer::slide_timing::SlideTiming;
use crate::writer::smart_tags::{SmartTagDefinition, SmartTagIndex};
use crate::writer::table::{PositionedTable, Table};
use crate::writer::text_format::{FontEntity, Paragraph, TextAlign};
use litchi_core::unit::pt_to_emu_i32;
use std::collections::{BTreeMap, HashMap};
use zeroize::Zeroizing;

impl Writer {
    /// Create a new PPT writer with standard 4:3 slide dimensions
    #[must_use]
    pub fn new() -> Self {
        Self::with_dimensions(9_144_000, 6_858_000) // 10" x 7.5" in EMUs
    }

    /// Create a new PPT writer with widescreen 16:9 dimensions
    #[must_use]
    pub fn new_widescreen() -> Self {
        Self::with_dimensions(9_144_000, 5_143_500) // 10" x 5.625" in EMUs
    }

    /// Create a new PPT writer with custom dimensions
    ///
    /// # Arguments
    ///
    /// * `width` - Slide width in EMUs (914400 EMUs = 1 inch)
    /// * `height` - Slide height in EMUs
    #[must_use]
    pub fn with_dimensions(width: i32, height: i32) -> Self {
        Self {
            slides: Vec::new(),
            properties: HashMap::new(),
            slide_width: width,
            slide_height: height,
            blip_store: Pictures::new(),
            hyperlinks: HyperlinkCollection::new(),
            fonts: default_font_collections(),
            sound_resources: BTreeMap::new(),
            next_sound_resource_id: 21,
            custom_shows: Vec::new(),
            smart_tags: Vec::new(),
            slide_view_info: None,
            notes_view_info: None,
            presentation_header_footer: None,
            notes_and_handouts_header_footer: None,
            main_master_header_footer: None,
            #[cfg(feature = "encryption")]
            encryption: None,
            modify_password: None,
            vba_project: None,
        }
    }

    /// Protect the generated presentation with `CryptoAPI` password-to-open encryption.
    ///
    /// Validation is atomic: invalid input leaves any previous setting unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "encryption")]
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: EncryptionProfile,
    ) -> Result<(), WriteError> {
        let protected_password = Zeroizing::new(password.into());
        validate_writer_password(profile, protected_password.as_str())
            .map_err(WriteError::InvalidData)?;
        self.encryption = Some(WriterEncryption {
            profile,
            password: protected_password,
        });
        Ok(())
    }

    /// Remove password-to-open protection and wipe the stored password.
    #[cfg(feature = "encryption")]
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured encryption profile without exposing the password.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub fn encryption_profile(&self) -> Option<EncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Configure a complete inert VBA project with safe limits and zlib storage.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<(), WriteError> {
        self.set_vba_with(
            project,
            &litchi_vba::Limits::default(),
            crate::VbaProjectCompression::Zlib,
        )
    }

    /// Configure a complete inert VBA project with explicit limits and storage.
    ///
    /// The inner CFB and optional outer zlib stream are fully serialized before
    /// writer state changes. Module source is never compiled or executed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
        compression: crate::VbaProjectCompression,
    ) -> Result<(), WriteError> {
        let payload = project.finish(limits)?;
        self.put_vba_with(payload, compression)
    }

    /// Configure an already validated inert VBA project using zlib storage.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) -> Result<(), WriteError> {
        self.put_vba_with(payload, crate::VbaProjectCompression::Zlib)
    }

    /// Configure an already validated inert VBA project with explicit storage.
    ///
    /// Import standalone CFB bytes through [`litchi_vba::Payload::read`] first.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn put_vba_with(
        &mut self,
        payload: litchi_vba::Payload,
        compression: crate::VbaProjectCompression,
    ) -> Result<(), WriteError> {
        use std::io::Write;

        let cfb = payload.into_bytes();
        let storage = match compression {
            crate::VbaProjectCompression::Uncompressed => {
                crate::embedded::storage::Storage::uncompressed(
                    crate::embedded::storage::Kind::VbaProject,
                    cfb,
                )
                .map_err(|error| WriteError::InvalidData(error.to_string()))?
            },
            crate::VbaProjectCompression::Zlib => {
                let uncompressed_len = u32::try_from(cfb.len()).map_err(|_err| {
                    WriteError::InvalidData(
                        "PowerPoint VBA project CFB exceeds the 32-bit size limit".to_string(),
                    )
                })?;
                let mut encoder =
                    flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
                encoder.write_all(&cfb)?;
                let data = encoder.finish()?;
                crate::embedded::storage::Storage::compressed(
                    crate::embedded::storage::Kind::VbaProject,
                    uncompressed_len,
                    data,
                )
                .map_err(|error| WriteError::InvalidData(error.to_string()))?
            },
        };
        storage
            .to_record_bytes()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        self.vba_project = Some(storage);
        Ok(())
    }

    /// Remove the configured VBA project and its persisted metadata.
    #[cfg(feature = "vba-inspection")]
    pub fn clear_vba(&mut self) {
        self.vba_project = None;
    }

    /// Whether a complete VBA project is configured for output.
    #[cfg(feature = "vba-inspection")]
    #[must_use]
    pub fn has_vba(&self) -> bool {
        self.vba_project.is_some()
    }

    /// Set the inert password required by `PowerPoint` to modify the presentation.
    ///
    /// The secret is stored in zeroizing memory. Password-to-open encryption
    /// must also be configured before the presentation can be written.
    /// Validation is atomic and does not replace an existing valid value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_modify_password(&mut self, password: impl Into<String>) -> Result<(), WriteError> {
        let protected_password = Zeroizing::new(password.into());
        validate_modify_password(protected_password.as_str())
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        self.modify_password = Some(WriterModifyPassword {
            password: protected_password,
        });
        Ok(())
    }

    /// Remove the modify-password atom and wipe the stored secret.
    pub fn clear_modify_password(&mut self) {
        self.modify_password = None;
    }

    /// Add a document-wide `PowerPoint` 11 smart tag and return its zero-based index.
    ///
    /// The returned index can be attached to one or more rich-text runs with
    /// [`crate::writer::text_format::TextRun::with_smart_tag`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_smart_tag(
        &mut self,
        definition: SmartTagDefinition,
    ) -> Result<SmartTagIndex, WriteError> {
        crate::writer::smart_tags::validate_definition(&definition)?;
        let index = u32::try_from(self.smart_tags.len()).map_err(|_err| {
            WriteError::InvalidData("PowerPoint smart-tag count exceeds u32".to_string())
        })?;
        self.smart_tags.push(definition);
        Ok(SmartTagIndex::new(index))
    }

    /// Return the number of document-wide smart tags.
    #[must_use]
    pub fn smart_tag_count(&self) -> usize {
        self.smart_tags.len()
    }

    /// Return the configured value through the redacted typed password model.
    #[must_use]
    pub fn modify_password(&self) -> Option<ModifyPassword> {
        self.modify_password
            .as_ref()
            .and_then(|value| ModifyPassword::new(value.password.as_str()).ok())
    }
}

impl Writer {
    /// Add a new blank slide
    ///
    /// # Returns
    ///
    /// * `Result<usize, WriteError>` - Slide index or error
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_slide(&mut self) -> Result<usize, WriteError> {
        let index = self.slides.len();
        self.slides.push(WritableSlide {
            shapes: Vec::new(),
            tables: Vec::new(),
            charts: Vec::new(),
            notes: None,
            notes_page: None,
            comments: Vec::new(),
            timing: None,
            transition: None,
            header_footer: None,
        });
        Ok(index)
    }

    /// Delete a slide
    ///
    /// # Arguments
    ///
    /// * `index` - Slide index (0-based)
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn delete_slide(&mut self, index: usize) -> Result<(), WriteError> {
        if index >= self.slides.len() {
            return Err(WriteError::InvalidData(format!(
                "Slide {index} does not exist"
            )));
        }
        self.slides.remove(index);
        self.reindex_slide_header_footers();
        Ok(())
    }

    /// Move a slide to a new position
    ///
    /// # Arguments
    ///
    /// * `from_index` - Current slide index
    /// * `to_index` - New slide index
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<(), WriteError> {
        if from_index >= self.slides.len() || to_index >= self.slides.len() {
            return Err(WriteError::InvalidData("Invalid slide index".to_string()));
        }

        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        self.reindex_slide_header_footers();
        Ok(())
    }

    fn validated_header_footer(
        mut value: HeaderFooter,
        scope: HeaderFooterScope,
    ) -> Result<HeaderFooter, WriteError> {
        value.scope = scope;
        value
            .to_record_bytes()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        Ok(value)
    }

    fn reindex_slide_header_footers(&mut self) {
        for (index, slide) in self.slides.iter_mut().enumerate() {
            if let Some(value) = &mut slide.header_footer {
                value.scope = HeaderFooterScope::Local {
                    parent: HeaderFooterParent::Slide,
                    parent_ordinal: HeaderFooterParentOrdinal::new(index),
                };
            }
        }
    }
}

impl Writer {
    /// Set presentation-wide header/footer defaults for ordinary slides.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_presentation_header_footer(
        &mut self,
        value: HeaderFooter,
    ) -> Result<(), WriteError> {
        let validated =
            Self::validated_header_footer(value, HeaderFooterScope::PresentationSlides)?;
        self.presentation_header_footer = Some(validated);
        Ok(())
    }

    /// Remove presentation-wide header/footer defaults for ordinary slides.
    pub fn clear_presentation_header_footer(&mut self) {
        self.presentation_header_footer = None;
    }

    /// Return presentation-wide header/footer defaults for ordinary slides.
    #[must_use]
    pub fn presentation_header_footer(&self) -> Option<&HeaderFooter> {
        self.presentation_header_footer.as_ref()
    }

    /// Set presentation-wide header/footer defaults for notes pages and handouts.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_notes_and_handouts_header_footer(
        &mut self,
        value: HeaderFooter,
    ) -> Result<(), WriteError> {
        let validated = Self::validated_header_footer(value, HeaderFooterScope::NotesAndHandouts)?;
        self.notes_and_handouts_header_footer = Some(validated);
        Ok(())
    }

    /// Remove presentation-wide header/footer defaults for notes pages and handouts.
    pub fn clear_notes_and_handouts_header_footer(&mut self) {
        self.notes_and_handouts_header_footer = None;
    }

    /// Return presentation-wide header/footer defaults for notes pages and handouts.
    #[must_use]
    pub fn notes_and_handouts_header_footer(&self) -> Option<&HeaderFooter> {
        self.notes_and_handouts_header_footer.as_ref()
    }

    /// Set the header/footer defaults attached directly to the main master.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_main_master_header_footer(&mut self, value: HeaderFooter) -> Result<(), WriteError> {
        let validated = Self::validated_header_footer(
            value,
            HeaderFooterScope::Local {
                parent: HeaderFooterParent::MainMaster,
                parent_ordinal: HeaderFooterParentOrdinal::new(0),
            },
        )?;
        self.main_master_header_footer = Some(validated);
        Ok(())
    }

    /// Remove the header/footer defaults attached directly to the main master.
    pub fn clear_main_master_header_footer(&mut self) {
        self.main_master_header_footer = None;
    }

    /// Return the header/footer defaults attached directly to the main master.
    #[must_use]
    pub fn main_master_header_footer(&self) -> Option<&HeaderFooter> {
        self.main_master_header_footer.as_ref()
    }

    /// Set a header/footer override attached directly to one slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_slide_header_footer(
        &mut self,
        slide: usize,
        value: HeaderFooter,
    ) -> Result<(), WriteError> {
        if slide >= self.slides.len() {
            return Err(WriteError::InvalidData(format!(
                "Slide {slide} does not exist"
            )));
        }
        let validated = Self::validated_header_footer(
            value,
            HeaderFooterScope::Local {
                parent: HeaderFooterParent::Slide,
                parent_ordinal: HeaderFooterParentOrdinal::new(slide),
            },
        )?;
        self.slides[slide].header_footer = Some(validated);
        Ok(())
    }

    /// Remove a header/footer override attached directly to one slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn clear_slide_header_footer(&mut self, slide: usize) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        slide_data.header_footer = None;
        Ok(())
    }

    /// Return the header/footer override attached directly to one slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_header_footer(&self, slide: usize) -> Result<Option<&HeaderFooter>, WriteError> {
        self.slides
            .get(slide)
            .map(|slide_data| slide_data.header_footer.as_ref())
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))
    }

    /// Add a text box to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x` - X position (in points, 72 points = 1 inch)
    /// * `y` - Y position (in points)
    /// * `width` - Width (in points)
    /// * `height` - Height (in points)
    /// * `text` - Text content
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_textbox(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        text: &str,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::TextBox,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                text: Some(text.to_string()),
                alignment: TextAlignment::Left,
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a text box with rich formatting
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x` - X position (in points)
    /// * `y` - Y position (in points)
    /// * `width` - Width (in points)
    /// * `height` - Height (in points)
    /// * `paragraphs` - Rich text paragraphs with formatting
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_rich_textbox(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        paragraphs: Vec<Paragraph>,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::TextBox,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                text: None,
                paragraphs: Some(paragraphs),
                alignment: TextAlignment::Left,
                fill: Some(FillStyle::none()),
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a rectangle shape to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x` - X position (in points)
    /// * `y` - Y position (in points)
    /// * `width` - Width (in points)
    /// * `height` - Height (in points)
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_rectangle(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
    ) -> Result<(), WriteError> {
        self.add_shape(slide, ShapeType::Rectangle, x, y, width, height)
    }

    /// Add an ellipse (oval) shape to a slide
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_ellipse(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
    ) -> Result<(), WriteError> {
        self.add_shape(slide, ShapeType::Ellipse, x, y, width, height)
    }

    /// Add a custom/freeform shape to a slide.
    ///
    /// The anchor is specified in points. Geometry vertices and segment words
    /// use the internal coordinate space declared by `geometry`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::too_many_arguments,
        reason = "parameters mirror the shape's distinct placement and style properties"
    )]
    pub fn add_freeform(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        geometry: FreeformGeometry,
    ) -> Result<(), WriteError> {
        self.add_styled_freeform(slide, x, y, width, height, geometry, ShapeStyle::default())
    }

    /// Add a styled custom/freeform shape to a slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::too_many_arguments,
        reason = "parameters mirror the shape's distinct placement and style properties"
    )]
    pub fn add_styled_freeform(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        geometry: FreeformGeometry,
        style: ShapeStyle,
    ) -> Result<(), WriteError> {
        geometry.validate()?;
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        slide_data.shapes.push(WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Freeform,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                freeform_geometry: Some(geometry),
                fill: Some(style.fill),
                line: Some(style.line),
                shadow: Some(style.shadow),
                ..Default::default()
            },
            animation_info: None,
        });
        Ok(())
    }

    /// Add a line to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x1`, `y1` - Start point (in points)
    /// * `x2`, `y2` - End point (in points)
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_line(
        &mut self,
        slide: usize,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let x = x1.min(x2);
        let y = y1.min(y2);
        let width = (x2 - x1).abs();
        let height = (y2 - y1).abs();

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Line,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                fill: Some(FillStyle::none()),
                line: Some(LineStyleConfig::default_line()),
                flip_h: x2 < x1,
                flip_v: y2 < y1,
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add an arrow line to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x1`, `y1` - Start point (in points)
    /// * `x2`, `y2` - End point (arrow head location, in points)
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_arrow_line(
        &mut self,
        slide: usize,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let x = x1.min(x2);
        let y = y1.min(y2);
        let width = (x2 - x1).abs();
        let height = (y2 - y1).abs();

        let mut line_style = LineStyleConfig::default_line();
        line_style.end_arrow = ArrowStyle::Triangle;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Line,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                fill: Some(FillStyle::none()),
                line: Some(line_style),
                flip_h: x2 < x1,
                flip_v: y2 < y1,
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a generic shape to a slide
    fn add_shape(
        &mut self,
        slide: usize,
        shape_type: ShapeType,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a styled shape to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `shape_type` - Type of shape
    /// * `x`, `y` - Position (in points)
    /// * `width`, `height` - Size (in points)
    /// * `style` - Visual style (fill, line, shadow)
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::too_many_arguments,
        reason = "parameters mirror the shape's distinct placement and style properties"
    )]
    pub fn add_styled_shape(
        &mut self,
        slide: usize,
        shape_type: ShapeType,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        style: ShapeStyle,
    ) -> Result<(), WriteError> {
        if shape_type == ShapeType::Freeform {
            return Err(WriteError::InvalidData(
                "freeform shapes require explicit geometry; use add_styled_freeform".to_string(),
            ));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                fill: Some(style.fill),
                line: Some(style.line),
                shadow: Some(style.shadow),
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a table to a slide
    ///
    /// The table is emitted as an `OfficeArt` table group (group shape with
    /// one rectangle cell shape per grid position), readable through the
    /// table extraction APIs after save/reopen.
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x`, `y` - Position of the table's top-left corner (in points)
    /// * `table` - Table grid, cell texts, and dimensions (see [`Table`])
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_table(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        table: Table,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        slide_data.tables.push(PositionedTable {
            x: pt_to_emu_i32(x),
            y: pt_to_emu_i32(y),
            table,
        });
        Ok(())
    }

    /// Validate a native-chart request and refuse incomplete binary authoring.
    ///
    /// No presentation state is changed. A structurally valid request returns
    /// [`litchi_ograph::Error::UnsupportedAuthoring`] through
    /// [`WriteError::Graph`] until the complete Office-compatible BIFF chart
    /// grammar is implemented. Invalid chart definitions, frames, or slide
    /// indexes continue to return [`WriteError::InvalidData`].
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x`, `y` - Position (in points)
    /// * `width`, `height` - Size (in points)
    /// * `chart` - The chart definition (kind, title, categories, series)
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "callers pass an owned chart definition, which binary chart authoring will consume once implemented"
    )]
    pub fn add_chart(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        chart: Chart,
    ) -> Result<(), WriteError> {
        if width <= 0 || height <= 0 {
            return Err(WriteError::InvalidData(
                "chart frame dimensions must be positive".to_string(),
            ));
        }
        let x_emu = pt_to_emu_i32(x);
        let y_emu = pt_to_emu_i32(y);
        let width_emu = pt_to_emu_i32(width);
        let height_emu = pt_to_emu_i32(height);
        x_emu.checked_add(width_emu).ok_or_else(|| {
            WriteError::InvalidData("chart frame horizontal extent is too large".to_string())
        })?;
        y_emu.checked_add(height_emu).ok_or_else(|| {
            WriteError::InvalidData("chart frame vertical extent is too large".to_string())
        })?;
        chart.validate()?;
        let total: usize = self.slides.iter().map(|entry| entry.charts.len()).sum();
        if total >= crate::writer::chart::MAX_CHART_OBJECTS {
            return Err(WriteError::InvalidData(format!(
                "presentation exceeds {} chart objects",
                crate::writer::chart::MAX_CHART_OBJECTS
            )));
        }
        self.slides
            .get(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        Err(litchi_ograph::Error::UnsupportedAuthoring {
            reason: "PPT chart creation requires the complete Office-compatible BIFF chart grammar",
        }
        .into())
    }
}

impl Writer {
    /// Add a picture to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x`, `y` - Position (in points)
    /// * `width`, `height` - Size (in points)
    /// * `image_data` - Raw image bytes (JPEG, PNG, etc.)
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_picture(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        image_data: Vec<u8>,
    ) -> Result<(), WriteError> {
        if slide >= self.slides.len() {
            return Err(WriteError::InvalidData(format!(
                "Slide {slide} does not exist"
            )));
        }
        // Add picture to BLIP store
        let blip_id = self.add_picture_data(image_data)?;

        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Picture,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                picture_index: Some(blip_id),
                fill: Some(FillStyle::picture(blip_id)),
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a picture with explicit type
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::too_many_arguments,
        reason = "parameters mirror the shape's distinct placement and style properties"
    )]
    pub fn add_picture_as(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        image_data: Vec<u8>,
        kind: PictureKind,
    ) -> Result<(), WriteError> {
        if slide >= self.slides.len() {
            return Err(WriteError::InvalidData(format!(
                "Slide {slide} does not exist"
            )));
        }
        let blip_id = self.add_picture_data_as(image_data, kind)?;

        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Picture,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                picture_index: Some(blip_id),
                fill: Some(FillStyle::picture(blip_id)),
                ..Default::default()
            },
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Register picture data in the BLIP store and return its 1-based index.
    ///
    /// Use the returned index with [`FillStyle::picture`] to create a picture
    /// or texture fill without adding a picture-frame shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_picture_data(&mut self, image_data: Vec<u8>) -> Result<PictureId, WriteError> {
        self.blip_store.add(image_data).map_err(WriteError::Io)
    }

    /// Register explicitly typed picture data and return its checked ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_picture_data_as(
        &mut self,
        image_data: Vec<u8>,
        kind: PictureKind,
    ) -> Result<PictureId, WriteError> {
        self.blip_store
            .add_as(image_data, kind)
            .map_err(WriteError::Io)
    }

    /// Add a hyperlink and return its ID
    ///
    /// The returned ID can be used with [`Self::set_last_shape_hyperlink`] to attach
    /// the hyperlink to a shape.
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) -> u32 {
        self.hyperlinks.add(hyperlink)
    }
}

impl Writer {
    /// Set the rotation of the last shape on a slide, in degrees.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_last_shape_rotation(
        &mut self,
        slide: usize,
        degrees: f32,
    ) -> Result<(), WriteError> {
        if !degrees.is_finite() {
            return Err(WriteError::InvalidData(
                "shape rotation must be finite".to_string(),
            ));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        shape.properties.rotation = degrees;
        Ok(())
    }

    /// Set one of the ten `OfficeArt` adjustment values on the last shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_last_shape_adjustment(
        &mut self,
        slide: usize,
        index: usize,
        value: i32,
    ) -> Result<(), WriteError> {
        if index >= 10 {
            return Err(WriteError::InvalidData(
                "shape adjustment index must be in the range 0..10".to_string(),
            ));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        if shape.properties.adjust_values.len() <= index {
            shape.properties.adjust_values.resize(index + 1, 0);
        }
        shape.properties.adjust_values[index] = value;
        Ok(())
    }

    /// Set horizontal alignment for all text in the last shape on a slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_last_shape_text_alignment(
        &mut self,
        slide: usize,
        alignment: TextAlignment,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        if shape.properties.text.is_none() && shape.properties.paragraphs.is_none() {
            return Err(WriteError::InvalidData(
                "Last shape has no text".to_string(),
            ));
        }
        shape.properties.alignment = alignment;
        if let Some(paragraphs) = &mut shape.properties.paragraphs {
            let text_align = TextAlign::from(alignment);
            for paragraph in paragraphs {
                paragraph.alignment = text_align;
            }
        }
        Ok(())
    }

    /// Add a font to the base collection and return its ordinal.
    ///
    /// This compatibility method refuses invalid or over-capacity additions and
    /// returns `u16::MAX`. New code should use [`Self::add_font_checked`] so the
    /// validation error is retained.
    pub fn add_font(&mut self, font: FontEntity) -> u16 {
        self.add_font_checked(font).unwrap_or(u16::MAX)
    }

    /// Atomically add a legacy writer font to the base collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_font_checked(&mut self, font: FontEntity) -> Result<u16, WriteError> {
        self.add_font_model(FontScope::Base, font.into())
    }

    /// Atomically add a typed font to the selected collection.
    ///
    /// The collection is serialized before publication, enforcing the 129-font
    /// limit, exact UTF-16 face-name rules, reserved authoring bits, facet order,
    /// and aggregate embedded-font limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_font_model(&mut self, scope: FontScope, font: Font) -> Result<u16, WriteError> {
        let mut candidate = self
            .fonts
            .collection(scope)
            .cloned()
            .unwrap_or_else(|| FontCollection::new(scope));
        if candidate.len() >= 129 {
            return Err(WriteError::InvalidData(
                "PowerPoint font collection exceeds the 129-font format limit".to_string(),
            ));
        }
        let index = candidate
            .try_push(font)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        candidate
            .to_record_bytes()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        match scope {
            FontScope::Base => self.fonts.base = Some(candidate),
            FontScope::International => self.fonts.international = Some(candidate),
        }
        Ok(index)
    }

    /// Atomically add or replace one inert EOT facet.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_embedded_font(
        &mut self,
        scope: FontScope,
        index: u16,
        facet: FontFacet,
        data: impl Into<crate::font::SharedFontData>,
    ) -> Result<Option<EmbeddedFont>, WriteError> {
        let mut candidate = self.fonts.collection(scope).cloned().ok_or_else(|| {
            WriteError::InvalidData(format!("PowerPoint {scope:?} font collection is absent"))
        })?;
        let replaced = candidate
            .set_facet(index, facet, data)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        candidate
            .to_record_bytes()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        match scope {
            FontScope::Base => self.fonts.base = Some(candidate),
            FontScope::International => self.fonts.international = Some(candidate),
        }
        Ok(replaced)
    }

    /// Prepare and atomically publish one inert `PowerPoint` EOT font facet.
    ///
    /// The facet is derived from `font.style`. This validates the actual
    /// OpenType `OS/2.fsType`, enforces the explicit document intent and EOT
    /// bounds, and never installs, loads, renders, or executes the font.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "fonts")]
    pub fn set_prepared_font(
        &mut self,
        scope: FontScope,
        index: u16,
        font: &mut crate::font::PreparedFont,
        intent: crate::font::EotIntent,
        limits: crate::font::EotLimits,
    ) -> Result<Option<EmbeddedFont>, WriteError> {
        self.set_prepared_font_with_limits(
            scope,
            index,
            font,
            intent,
            limits,
            crate::font::Limits::default(),
        )
    }

    /// Prepare and atomically publish an inert EOT facet under explicit PPT
    /// record and aggregate-font limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "fonts")]
    pub fn set_prepared_font_with_limits(
        &mut self,
        scope: FontScope,
        index: u16,
        font: &mut crate::font::PreparedFont,
        intent: crate::font::EotIntent,
        eot_limits: crate::font::EotLimits,
        ppt_limits: crate::font::Limits,
    ) -> Result<Option<EmbeddedFont>, WriteError> {
        let facet = crate::font::prepared::facet_for_style(font.style);
        let subsetted = font.subsetted;
        self.fonts
            .validate_with_limits(ppt_limits)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        let mut collection = self.fonts.collection(scope).cloned().ok_or_else(|| {
            WriteError::InvalidData(format!("PowerPoint {scope:?} font collection is absent"))
        })?;
        collection
            .to_record_bytes_with_limits(ppt_limits)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        let current = collection.get(index).ok_or_else(|| {
            WriteError::InvalidData(format!("unknown PowerPoint font ordinal {index}"))
        })?;
        if current
            .embedded_fonts
            .iter()
            .any(|embedded| embedded.style != facet as u8)
            && current.embedded_subset != subsetted
        {
            return Err(WriteError::InvalidData(
                "one PowerPoint face cannot mix subsetted and complete embedded facets".into(),
            ));
        }

        let data = litchi_fonts::embedding::powerpoint::encode(font, intent, eot_limits).map_err(
            |error| WriteError::InvalidData(format!("font preparation failed: {error}")),
        )?;
        let replaced = crate::font::prepared::stage_facet(&mut collection, index, facet, data);
        let staged = collection.get_mut(index).ok_or_else(|| {
            WriteError::InvalidData(format!("unknown PowerPoint font ordinal {index}"))
        })?;
        staged.embedded_subset = subsetted;
        if subsetted {
            staged.font_flags |= 1;
        } else {
            staged.font_flags &= !1;
        }
        let mut candidate = self.fonts.clone();
        match scope {
            FontScope::Base => candidate.base = Some(collection),
            FontScope::International => candidate.international = Some(collection),
        }
        if let Err(error) = candidate.validate_with_limits(ppt_limits) {
            if let Err(restore) = crate::font::prepared::restore_staged(
                font,
                &mut candidate,
                scope,
                index,
                facet,
                eot_limits,
            ) {
                return Err(WriteError::InvalidData(format!(
                    "prepared-font rollback invariant failed: {restore}"
                )));
            }
            return Err(WriteError::InvalidData(error.to_string()));
        }
        let Some(staged_collection) = candidate.collection(scope) else {
            return Err(WriteError::InvalidData(
                "staged font collection is missing after publication".to_string(),
            ));
        };
        let serialized = staged_collection.to_record_bytes_with_limits(ppt_limits);
        if let Err(error) = serialized {
            if let Err(restore) = crate::font::prepared::restore_staged(
                font,
                &mut candidate,
                scope,
                index,
                facet,
                eot_limits,
            ) {
                return Err(WriteError::InvalidData(format!(
                    "prepared-font rollback invariant failed: {restore}"
                )));
            }
            return Err(WriteError::InvalidData(error.to_string()));
        }
        self.fonts = candidate;
        Ok(replaced)
    }

    /// Set or clear the `PowerPoint` 10 document-wide embedding flags.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_font_embedding_flags(
        &mut self,
        flags: Option<FontEmbeddingFlags>,
    ) -> Result<(), WriteError> {
        let mut candidate = self.fonts.clone();
        candidate.embedding_flags = flags;
        candidate
            .powerpoint10_records()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        self.fonts = candidate;
        Ok(())
    }

    /// Return the complete inert font catalog configured for fresh output.
    #[must_use]
    pub const fn font_collections(&self) -> &FontCollections {
        &self.fonts
    }

    /// Set slide notes (simple text)
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `notes` - Notes text
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_slide_notes(&mut self, slide: usize, notes: &str) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        slide_data.notes = Some(notes.to_string());
        Ok(())
    }

    /// Set rich notes page for a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `notes_page` - Full notes page with formatting
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_notes_page(
        &mut self,
        slide: usize,
        notes_page: NotesPage,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        slide_data.notes_page = Some(notes_page);
        Ok(())
    }

    /// Set animation for a specific shape
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `shape_index` - Shape index on the slide (0-based)
    /// * `animation` - Animation info to attach to the shape
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_shape_animation(
        &mut self,
        slide: usize,
        shape_index: usize,
        animation: AnimationInfo,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;

        let shape = slide_data.shapes.get_mut(shape_index).ok_or_else(|| {
            WriteError::InvalidData(format!(
                "Shape {shape_index} does not exist on slide {slide}"
            ))
        })?;

        shape.animation_info = Some(animation);
        Ok(())
    }

    /// Get number of pictures in the presentation
    #[must_use]
    pub fn picture_count(&self) -> usize {
        self.blip_store.len()
    }

    /// Get number of hyperlinks in the presentation
    #[must_use]
    pub fn hyperlink_count(&self) -> usize {
        self.hyperlinks.len()
    }

    /// Get number of fonts in the presentation
    #[must_use]
    pub fn font_count(&self) -> usize {
        self.fonts.base_font_count()
    }

    /// Get the number of `PowerPoint` 10 international fonts.
    pub fn international_font_count(&self) -> usize {
        self.fonts
            .international
            .as_ref()
            .map_or(0, FontCollection::len)
    }

    /// Register an exact embedded WAV or AIFF resource for interactions or animations.
    ///
    /// Validation is atomic. The returned non-zero writer-local ID can be
    /// passed to [`crate::Interaction::with_sound_reference`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_embedded_sound(
        &mut self,
        name: impl Into<String>,
        data: Vec<u8>,
    ) -> Result<std::num::NonZeroU32, WriteError> {
        if self.sound_resources.len()
            >= crate::writer::sound_collection::SoundCollectionLimits::default().max_sounds
        {
            return Err(WriteError::InvalidData(
                "sound collection exceeds the configured sound count".to_string(),
            ));
        }
        let next = self
            .next_sound_resource_id
            .checked_add(1)
            .ok_or_else(|| WriteError::InvalidData("sound resource ID overflow".to_string()))?;
        let Some(id) = std::num::NonZeroU32::new(self.next_sound_resource_id) else {
            return Err(WriteError::InvalidData(
                "sound resource ID must be non-zero".to_string(),
            ));
        };
        let sound_type = crate::animation::SoundType::Embedded {
            name: name.into(),
            data,
        };
        let mut validator = crate::writer::sound_collection::SoundCollectionBuilder::new(
            crate::writer::sound_collection::SoundCollectionLimits::default(),
        );
        validator
            .register(id.get(), &sound_type)
            .map_err(sound_collection_error)?;

        self.sound_resources.insert(id.get(), sound_type);
        self.next_sound_resource_id = next;
        Ok(id)
    }

    /// Atomically replace one explicitly registered embedded sound.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_embedded_sound(
        &mut self,
        id: std::num::NonZeroU32,
        name: impl Into<String>,
        data: Vec<u8>,
    ) -> Result<(), WriteError> {
        if !self.sound_resources.contains_key(&id.get()) {
            return Err(WriteError::InvalidData(format!(
                "sound resource {id} does not exist"
            )));
        }
        let sound_type = crate::animation::SoundType::Embedded {
            name: name.into(),
            data,
        };
        let mut validator = crate::writer::sound_collection::SoundCollectionBuilder::new(
            crate::writer::sound_collection::SoundCollectionLimits::default(),
        );
        validator
            .register(id.get(), &sound_type)
            .map_err(sound_collection_error)?;
        self.sound_resources.insert(id.get(), sound_type);
        Ok(())
    }

    /// Remove one explicit sound resource.
    ///
    /// Any remaining reference to this ID causes serialization to fail rather
    /// than producing a dangling `SoundIdRef`.
    pub fn remove_embedded_sound(&mut self, id: std::num::NonZeroU32) -> bool {
        self.sound_resources.remove(&id.get()).is_some()
    }

    /// Number of explicitly registered embedded sound resources.
    #[must_use]
    pub fn embedded_sound_count(&self) -> usize {
        self.sound_resources.len()
    }

    /// Add a comment to a slide.
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index (0-based)
    /// * `comment` - The comment to add
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_comment(&mut self, slide: usize, comment: SlideComment) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        slide_data.comments.push(comment);
        Ok(())
    }

    /// Set per-slide timing (auto-advance, hidden, etc.).
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index (0-based)
    /// * `timing` - Timing configuration
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_slide_timing(
        &mut self,
        slide: usize,
        timing: SlideTiming,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        slide_data.timing = Some(timing);
        Ok(())
    }

    /// Set the slide transition effect.
    ///
    /// The transition owns the slide's `SSSlideInfoAtom` record (MS-PPT 2.6.6),
    /// so per-slide timing set via [`Writer::set_slide_timing`] is emitted only
    /// for slides without a transition; a hidden-slide flag is preserved either
    /// way.
    ///
    /// Transition sounds are refused: the create-side writer only builds the
    /// `SoundCollection` from shape and animation references, so a transition
    /// `soundIdRef` would dangle.
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index (0-based)
    /// * `transition` - Transition configuration
    ///
    /// # Errors
    ///
    /// Returns an error if the slide does not exist or the transition carries a
    /// sound.
    pub fn set_slide_transition(
        &mut self,
        slide: usize,
        transition: TransitionInfo,
    ) -> Result<(), WriteError> {
        if transition.sound.is_some() {
            return Err(WriteError::InvalidData(
                "transition sounds are not supported by the create-side writer".to_string(),
            ));
        }
        if crate::transition::encode_visual(
            transition.transition_type,
            transition.direction,
            transition.speed,
        )
        .is_none()
        {
            return Err(WriteError::InvalidData(
                "transition type/direction is not representable by [MS-PPT] 2.6.6".to_string(),
            ));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {slide} does not exist")))?;
        slide_data.transition = Some(transition);
        Ok(())
    }

    /// Add a custom slide show (named show).
    ///
    /// # Arguments
    ///
    /// * `show` - Custom show definition with name and slide indices
    pub fn add_custom_show(&mut self, show: CustomShow) {
        self.custom_shows.push(show);
    }

    /// Get the number of custom shows.
    #[must_use]
    pub fn custom_show_count(&self) -> usize {
        self.custom_shows.len()
    }
}

impl Writer {
    /// Set a presentation property
    ///
    /// # Arguments
    ///
    /// * `name` - Property name (e.g., "Title", "Author", "Subject")
    /// * `value` - Property value
    pub fn set_property(&mut self, name: &str, value: &str) {
        self.properties.insert(name.to_string(), value.to_string());
    }

    /// Get the number of slides
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }
}

impl Default for Writer {
    fn default() -> Self {
        Self::new()
    }
}

fn default_font_collections() -> FontCollections {
    let mut base = FontCollection::new(FontScope::Base);
    #[allow(
        clippy::expect_used,
        reason = "the built-in Arial font satisfies every collection invariant and a fresh collection has capacity"
    )]
    base.try_push(FontEntity::arial().into())
        .expect("the built-in Arial font is valid");
    FontCollections {
        base: Some(base),
        international: None,
        embedding_flags: None,
    }
}
