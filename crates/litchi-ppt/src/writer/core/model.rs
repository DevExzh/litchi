//! Semantic state and authoring operations for the legacy PPT writer.

use super::super::blip::{Id as PictureId, Kind as PictureKind, Pictures};
use super::super::chart::{Chart, ChartPlan, PositionedChart};
use super::super::comments::SlideComment;
use super::super::custom_shows::CustomShow;
use super::super::escher::FreeformGeometry;
use super::super::hyperlink::{Hyperlink, HyperlinkCollection};
use super::super::notes::NotesPage;
use super::super::persist::PersistPtrBuilder;
use super::super::records::{RecordBuilder, record_type};
use super::super::shape_style::{ArrowStyle, FillStyle, LineStyleConfig, ShadowStyle, ShapeStyle};
use super::super::slide_timing::SlideTiming;
use super::super::smart_tags::{SmartTagDefinition, SmartTagIndex};
use super::super::spec::Tag10;
use super::super::table::{PositionedTable, Table};
use super::super::text_format::{FontEntity, Paragraph, TextAlign};
use super::codec::{interaction_for_hyperlink, shape_text_unit_count, sound_collection_error};
use crate::animation::AnimationInfo;
use crate::encryption::{
    EncryptionProfile, WriterEncryptionMaterial, prepare_writer_encryption,
    validate_writer_password,
};
use crate::header_footer::{
    HeaderFooter, HeaderFooterParent, HeaderFooterParentOrdinal, HeaderFooterScope,
};
use crate::modify_password::{ModifyPassword, validate_value as validate_modify_password};
use crate::view_info::{SlideViewInfo, ViewKind};
use litchi_core::unit::pt_to_emu_i32;
use std::collections::{BTreeMap, HashMap};
use zeroize::Zeroizing;

/// Error type for PPT writing
#[derive(Debug)]
pub enum WriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(litchi_cfb::OleError),
    /// MS-OVBA project authoring error
    Vba(litchi_vba::Error),
    /// Host-neutral Office Graph authoring error.
    Graph(litchi_ograph::Error),
}
impl From<std::io::Error> for WriteError {
    fn from(err: std::io::Error) -> Self {
        WriteError::Io(err)
    }
}

impl From<litchi_cfb::OleError> for WriteError {
    fn from(err: litchi_cfb::OleError) -> Self {
        WriteError::Ole(err)
    }
}

impl From<litchi_vba::Error> for WriteError {
    fn from(err: litchi_vba::Error) -> Self {
        WriteError::Vba(err)
    }
}

impl From<litchi_ograph::Error> for WriteError {
    fn from(err: litchi_ograph::Error) -> Self {
        Self::Graph(err)
    }
}

impl std::fmt::Display for WriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            WriteError::Io(e) => write!(f, "I/O error: {}", e),
            WriteError::InvalidData(s) => write!(f, "Invalid data: {}", s),
            WriteError::Ole(e) => write!(f, "OLE error: {}", e),
            WriteError::Vba(e) => write!(f, "VBA project error: {}", e),
            WriteError::Graph(e) => write!(f, "Office Graph error: {e}"),
        }
    }
}

impl std::error::Error for WriteError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Ole(error) => Some(error),
            Self::Vba(error) => Some(error),
            Self::Graph(error) => Some(error),
            Self::InvalidData(_) => None,
        }
    }
}

/// Shape type (legacy - use ShapeKind from shapes module for new code)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    /// Rectangle
    Rectangle,
    /// Text box
    TextBox,
    /// Placeholder
    Placeholder,
    /// Line
    Line,
    /// Ellipse
    Ellipse,
    /// Rounded rectangle
    RoundRectangle,
    /// Diamond
    Diamond,
    /// Triangle
    Triangle,
    /// Arrow (block arrow shape)
    Arrow,
    /// Star
    Star,
    /// Heart
    Heart,
    /// Picture frame
    Picture,
    /// Custom/freeform shape with explicit OfficeArt geometry
    Freeform,
}

/// Text alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlignment {
    /// Left-aligned
    Left,
    /// Center-aligned
    Center,
    /// Right-aligned
    Right,
    /// Justified
    Justify,
}

impl From<TextAlignment> for TextAlign {
    fn from(value: TextAlignment) -> Self {
        match value {
            TextAlignment::Left => Self::Left,
            TextAlignment::Center => Self::Center,
            TextAlignment::Right => Self::Right,
            TextAlignment::Justify => Self::Justify,
        }
    }
}

/// Shape properties (extended with styling support)
#[derive(Debug, Clone)]
pub struct ShapeProperties {
    /// Shape type
    pub shape_type: ShapeType,
    /// X position (in EMUs - English Metric Units, 914400 EMUs = 1 inch)
    pub x: i32,
    /// Y position (in EMUs)
    pub y: i32,
    /// Width (in EMUs)
    pub width: i32,
    /// Height (in EMUs)
    pub height: i32,
    /// Text content (if applicable)
    pub text: Option<String>,
    /// Rich text paragraphs (alternative to plain text)
    pub paragraphs: Option<Vec<Paragraph>>,
    /// Text alignment
    pub alignment: TextAlignment,
    /// Fill style
    pub fill: Option<FillStyle>,
    /// Line style
    pub line: Option<LineStyleConfig>,
    /// Shadow style
    pub shadow: Option<ShadowStyle>,
    /// Rotation in degrees
    pub rotation: f32,
    /// Shape-specific adjustment values, at indices 0 through 9.
    pub adjust_values: Vec<i32>,
    /// Flip horizontal
    pub flip_h: bool,
    /// Flip vertical
    pub flip_v: bool,
    /// Picture BLIP index (for Picture type)
    pub picture_index: Option<PictureId>,
    /// Explicit geometry for a custom/freeform shape.
    pub freeform_geometry: Option<FreeformGeometry>,
    /// Hyperlink attached to shape
    pub hyperlink_id: Option<u32>,
    /// Typed click and mouse-over actions attached to the shape.
    pub interactions: Vec<crate::Interaction>,
    /// Typed actions attached to UTF-16 ranges in the shape text.
    pub text_interactions: Vec<crate::TextInteraction>,
}

/// Represents a shape on a slide
#[derive(Debug, Clone)]
pub(super) struct WritableShape {
    /// Shape properties
    pub(super) properties: ShapeProperties,
    /// Animation info for this shape
    pub(super) animation_info: Option<AnimationInfo>,
}

impl Default for ShapeProperties {
    fn default() -> Self {
        Self {
            shape_type: ShapeType::Rectangle,
            x: 0,
            y: 0,
            width: 914400,  // 1 inch
            height: 914400, // 1 inch
            text: None,
            paragraphs: None,
            alignment: TextAlignment::Left,
            fill: None,
            line: None,
            shadow: None,
            rotation: 0.0,
            adjust_values: Vec::new(),
            flip_h: false,
            flip_v: false,
            picture_index: None,
            freeform_geometry: None,
            hyperlink_id: None,
            interactions: Vec::new(),
            text_interactions: Vec::new(),
        }
    }
}

/// Represents a slide
#[derive(Debug, Clone)]
pub(super) struct WritableSlide {
    /// Shapes on this slide
    pub(super) shapes: Vec<WritableShape>,
    /// Tables on this slide
    pub(super) tables: Vec<PositionedTable>,
    /// Native charts on this slide
    pub(super) charts: Vec<PositionedChart>,
    /// Slide notes text (simple)
    pub(super) notes: Option<String>,
    /// Rich notes page
    pub(super) notes_page: Option<NotesPage>,
    /// Slide comments
    pub(super) comments: Vec<SlideComment>,
    /// Per-slide timing (auto-advance, hidden, etc.)
    pub(super) timing: Option<SlideTiming>,
    /// Optional header/footer override attached directly to this slide.
    pub(super) header_footer: Option<HeaderFooter>,
}

pub(super) struct SerializedHeaderFooters {
    pub(super) presentation_slides: Option<Vec<u8>>,
    pub(super) notes_and_handouts: Option<Vec<u8>>,
    pub(super) main_master: Option<Vec<u8>>,
    pub(super) slides: Vec<Option<Vec<u8>>>,
}

impl WritableSlide {
    /// Number of OfficeArt shapes in this slide's drawing, including the
    /// group patriarch, the background shape, and every table group/cell.
    pub(super) fn escher_shape_count(&self) -> u32 {
        let table_shapes: u32 = self.tables.iter().map(|t| t.table.shape_count()).sum();
        2 + self.shapes.len() as u32 + table_shapes + self.charts.len() as u32
    }
}
pub struct Writer {
    /// Slides in the presentation
    pub(super) slides: Vec<WritableSlide>,
    /// Presentation properties
    pub(super) properties: HashMap<String, String>,
    /// Slide width in EMUs (default: Letter size)
    pub(super) slide_width: i32,
    /// Slide height in EMUs (default: Letter size)
    pub(super) slide_height: i32,
    /// Picture/BLIP storage
    pub(super) blip_store: Pictures,
    /// Hyperlink collection
    pub(super) hyperlinks: HyperlinkCollection,
    /// Font collection
    pub(super) fonts: Vec<FontEntity>,
    /// Explicit embedded sound resources keyed by writer-local IDs.
    pub(super) sound_resources: BTreeMap<u32, crate::animation::SoundType>,
    /// Next writer-local sound ID; built-in catalog IDs occupy 1 through 20.
    pub(super) next_sound_resource_id: u32,
    /// Custom slide shows (named shows)
    pub(super) custom_shows: Vec<CustomShow>,
    /// Document-wide PowerPoint 11 smart-tag property bags.
    pub(super) smart_tags: Vec<SmartTagDefinition>,
    /// Optional typed override for the slide editing view.
    pub(super) slide_view_info: Option<SlideViewInfo>,
    /// Optional typed notes editing view.
    pub(super) notes_view_info: Option<SlideViewInfo>,
    /// Presentation-wide defaults for ordinary slides.
    pub(super) presentation_header_footer: Option<HeaderFooter>,
    /// Presentation-wide defaults for notes pages and handouts.
    pub(super) notes_and_handouts_header_footer: Option<HeaderFooter>,
    /// Header/footer defaults attached directly to the main master.
    pub(super) main_master_header_footer: Option<HeaderFooter>,
    /// Password-to-open settings, including a password wiped on replacement or drop.
    encryption: Option<WriterEncryption>,
    /// Inert modify password, wiped on replacement, clear, or drop.
    modify_password: Option<WriterModifyPassword>,
    /// Standalone CFB project wrapped for a persisted `VbaProjectStg`.
    pub(super) vba_project: Option<crate::embedded::storage::Storage>,
}

struct WriterEncryption {
    profile: EncryptionProfile,
    password: Zeroizing<String>,
}

struct WriterModifyPassword {
    password: Zeroizing<String>,
}

impl std::fmt::Debug for WriterModifyPassword {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("WriterModifyPassword")
            .field("utf16_units", &self.password.encode_utf16().count())
            .field("password", &"[REDACTED]")
            .finish()
    }
}

impl Writer {
    /// Create a new PPT writer with standard 4:3 slide dimensions
    pub fn new() -> Self {
        Self::with_dimensions(9144000, 6858000) // 10" x 7.5" in EMUs
    }

    /// Create a new PPT writer with widescreen 16:9 dimensions
    pub fn new_widescreen() -> Self {
        Self::with_dimensions(9144000, 5143500) // 10" x 5.625" in EMUs
    }

    /// Create a new PPT writer with custom dimensions
    ///
    /// # Arguments
    ///
    /// * `width` - Slide width in EMUs (914400 EMUs = 1 inch)
    /// * `height` - Slide height in EMUs
    pub fn with_dimensions(width: i32, height: i32) -> Self {
        Self {
            slides: Vec::new(),
            properties: HashMap::new(),
            slide_width: width,
            slide_height: height,
            blip_store: Pictures::new(),
            hyperlinks: HyperlinkCollection::new(),
            fonts: vec![FontEntity::arial()], // Default font
            sound_resources: BTreeMap::new(),
            next_sound_resource_id: 21,
            custom_shows: Vec::new(),
            smart_tags: Vec::new(),
            slide_view_info: None,
            notes_view_info: None,
            presentation_header_footer: None,
            notes_and_handouts_header_footer: None,
            main_master_header_footer: None,
            encryption: None,
            modify_password: None,
            vba_project: None,
        }
    }

    /// Protect the generated presentation with CryptoAPI password-to-open encryption.
    ///
    /// Validation is atomic: invalid input leaves any previous setting unchanged.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: EncryptionProfile,
    ) -> Result<(), WriteError> {
        let password = Zeroizing::new(password.into());
        validate_writer_password(profile, password.as_str()).map_err(WriteError::InvalidData)?;
        self.encryption = Some(WriterEncryption { profile, password });
        Ok(())
    }

    /// Remove password-to-open protection and wipe the stored password.
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured encryption profile without exposing the password.
    pub fn encryption_profile(&self) -> Option<EncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Configure a complete inert VBA project with safe limits and zlib storage.
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
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) -> Result<(), WriteError> {
        self.put_vba_with(payload, crate::VbaProjectCompression::Zlib)
    }

    /// Configure an already validated inert VBA project with explicit storage.
    ///
    /// Import standalone CFB bytes through [`litchi_vba::Payload::read`] first.
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
                let uncompressed_len = u32::try_from(cfb.len()).map_err(|_| {
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
    pub fn clear_vba(&mut self) {
        self.vba_project = None;
    }

    /// Whether a complete VBA project is configured for output.
    pub fn has_vba(&self) -> bool {
        self.vba_project.is_some()
    }

    /// Set the inert password required by PowerPoint to modify the presentation.
    ///
    /// The secret is stored in zeroizing memory. Password-to-open encryption
    /// must also be configured before the presentation can be written.
    /// Validation is atomic and does not replace an existing valid value.
    pub fn set_modify_password(&mut self, password: impl Into<String>) -> Result<(), WriteError> {
        let password = Zeroizing::new(password.into());
        validate_modify_password(password.as_str())
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        self.modify_password = Some(WriterModifyPassword { password });
        Ok(())
    }

    /// Remove the modify-password atom and wipe the stored secret.
    pub fn clear_modify_password(&mut self) {
        self.modify_password = None;
    }

    /// Add a document-wide PowerPoint 11 smart tag and return its zero-based index.
    ///
    /// The returned index can be attached to one or more rich-text runs with
    /// [`super::super::text_format::TextRun::with_smart_tag`].
    pub fn add_smart_tag(
        &mut self,
        definition: SmartTagDefinition,
    ) -> Result<SmartTagIndex, WriteError> {
        super::super::smart_tags::validate_definition(&definition)?;
        let index = u32::try_from(self.smart_tags.len()).map_err(|_| {
            WriteError::InvalidData("PowerPoint smart-tag count exceeds u32".to_string())
        })?;
        self.smart_tags.push(definition);
        Ok(SmartTagIndex::new(index))
    }

    /// Return the number of document-wide smart tags.
    pub fn smart_tag_count(&self) -> usize {
        self.smart_tags.len()
    }

    /// Return the configured value through the redacted typed password model.
    pub fn modify_password(&self) -> Option<ModifyPassword> {
        self.modify_password.as_ref().map(|value| {
            ModifyPassword::new(value.password.as_str())
                .expect("stored modify password was validated before assignment")
        })
    }

    pub(super) fn validate_encryption(&self) -> Result<(), WriteError> {
        if self.modify_password.is_some() && self.encryption.is_none() {
            return Err(WriteError::InvalidData(
                "PowerPoint modify-password output requires password-to-open encryption"
                    .to_string(),
            ));
        }
        if let Some(value) = &self.encryption {
            validate_writer_password(value.profile, value.password.as_str())
                .map_err(WriteError::InvalidData)?;
        }
        if let Some(value) = &self.modify_password {
            validate_modify_password(value.password.as_str())
                .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        }
        Ok(())
    }

    pub(super) fn validate_smart_tag_references(&self) -> Result<(), WriteError> {
        let smart_tag_count = u32::try_from(self.smart_tags.len()).map_err(|_| {
            WriteError::InvalidData("PowerPoint smart-tag count exceeds u32".to_string())
        })?;
        for run in self
            .slides
            .iter()
            .flat_map(|slide| &slide.shapes)
            .filter_map(|shape| shape.properties.paragraphs.as_ref())
            .flatten()
            .flat_map(|paragraph| &paragraph.runs)
            .filter(|run| !run.smart_tag_indices.is_empty())
        {
            if run.text.is_empty() {
                return Err(WriteError::InvalidData(
                    "PowerPoint smart tags cannot be attached to an empty text run".to_string(),
                ));
            }
            for index in &run.smart_tag_indices {
                if index.as_u32() >= smart_tag_count {
                    return Err(WriteError::InvalidData(format!(
                        "PowerPoint text run references missing smart tag {}",
                        index.as_u32()
                    )));
                }
            }
        }
        Ok(())
    }

    pub(super) fn build_modify_password_programmable_tag(
        &self,
    ) -> Result<Option<Vec<u8>>, WriteError> {
        let Some(value) = &self.modify_password else {
            return Ok(None);
        };
        validate_modify_password(value.password.as_str())
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;

        let mut atom = RecordBuilder::new(0x00, 3, record_type::CSTRING);
        let atom_data = value
            .password
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        atom.write_data(&atom_data);
        let mut blob = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
        blob.write_child(&atom.build()?);
        let mut binary_tag = RecordBuilder::new(0x0f, 0, record_type::PROG_BINARY_TAG);
        let mut name = RecordBuilder::new(0x00, 0, record_type::CSTRING);
        name.write_data(&Tag10::to_bytes());
        binary_tag.write_child(&name.build()?);
        binary_tag.write_child(&blob.build()?);
        let mut tags = RecordBuilder::new(0x0f, 0, record_type::PROG_TAGS);
        tags.write_child(&binary_tag.build()?);
        Ok(Some(tags.build()?))
    }

    pub(super) fn prepare_encryption(
        &self,
    ) -> Result<Option<WriterEncryptionMaterial>, WriteError> {
        self.encryption
            .as_ref()
            .map(|value| prepare_writer_encryption(value.profile, value.password.as_str()))
            .transpose()
            .map_err(WriteError::InvalidData)
    }

    /// Add a new blank slide
    ///
    /// # Returns
    ///
    /// * `Result<usize, WriteError>` - Slide index or error
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
            header_footer: None,
        });
        Ok(index)
    }

    /// Delete a slide
    ///
    /// # Arguments
    ///
    /// * `index` - Slide index (0-based)
    pub fn delete_slide(&mut self, index: usize) -> Result<(), WriteError> {
        if index >= self.slides.len() {
            return Err(WriteError::InvalidData(format!(
                "Slide {} does not exist",
                index
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

    pub(super) fn serialize_header_footers(&self) -> Result<SerializedHeaderFooters, WriteError> {
        let serialize = |value: Option<&HeaderFooter>, scope| {
            value
                .map(|value| {
                    let mut value = value.clone();
                    value.scope = scope;
                    value
                        .to_record_bytes()
                        .map_err(|error| WriteError::InvalidData(error.to_string()))
                })
                .transpose()
        };
        let presentation_slides = serialize(
            self.presentation_header_footer.as_ref(),
            HeaderFooterScope::PresentationSlides,
        )?;
        let notes_and_handouts = serialize(
            self.notes_and_handouts_header_footer.as_ref(),
            HeaderFooterScope::NotesAndHandouts,
        )?;
        let main_master = serialize(
            self.main_master_header_footer.as_ref(),
            HeaderFooterScope::Local {
                parent: HeaderFooterParent::MainMaster,
                parent_ordinal: HeaderFooterParentOrdinal::new(0),
            },
        )?;
        let slides = self
            .slides
            .iter()
            .enumerate()
            .map(|(index, slide)| {
                serialize(
                    slide.header_footer.as_ref(),
                    HeaderFooterScope::Local {
                        parent: HeaderFooterParent::Slide,
                        parent_ordinal: HeaderFooterParentOrdinal::new(index),
                    },
                )
            })
            .collect::<Result<Vec<_>, _>>()?;
        Ok(SerializedHeaderFooters {
            presentation_slides,
            notes_and_handouts,
            main_master,
            slides,
        })
    }

    /// Set presentation-wide header/footer defaults for ordinary slides.
    pub fn set_presentation_header_footer(
        &mut self,
        value: HeaderFooter,
    ) -> Result<(), WriteError> {
        let value = Self::validated_header_footer(value, HeaderFooterScope::PresentationSlides)?;
        self.presentation_header_footer = Some(value);
        Ok(())
    }

    /// Remove presentation-wide header/footer defaults for ordinary slides.
    pub fn clear_presentation_header_footer(&mut self) {
        self.presentation_header_footer = None;
    }

    /// Return presentation-wide header/footer defaults for ordinary slides.
    pub fn presentation_header_footer(&self) -> Option<&HeaderFooter> {
        self.presentation_header_footer.as_ref()
    }

    /// Set presentation-wide header/footer defaults for notes pages and handouts.
    pub fn set_notes_and_handouts_header_footer(
        &mut self,
        value: HeaderFooter,
    ) -> Result<(), WriteError> {
        let value = Self::validated_header_footer(value, HeaderFooterScope::NotesAndHandouts)?;
        self.notes_and_handouts_header_footer = Some(value);
        Ok(())
    }

    /// Remove presentation-wide header/footer defaults for notes pages and handouts.
    pub fn clear_notes_and_handouts_header_footer(&mut self) {
        self.notes_and_handouts_header_footer = None;
    }

    /// Return presentation-wide header/footer defaults for notes pages and handouts.
    pub fn notes_and_handouts_header_footer(&self) -> Option<&HeaderFooter> {
        self.notes_and_handouts_header_footer.as_ref()
    }

    /// Set the header/footer defaults attached directly to the main master.
    pub fn set_main_master_header_footer(&mut self, value: HeaderFooter) -> Result<(), WriteError> {
        let value = Self::validated_header_footer(
            value,
            HeaderFooterScope::Local {
                parent: HeaderFooterParent::MainMaster,
                parent_ordinal: HeaderFooterParentOrdinal::new(0),
            },
        )?;
        self.main_master_header_footer = Some(value);
        Ok(())
    }

    /// Remove the header/footer defaults attached directly to the main master.
    pub fn clear_main_master_header_footer(&mut self) {
        self.main_master_header_footer = None;
    }

    /// Return the header/footer defaults attached directly to the main master.
    pub fn main_master_header_footer(&self) -> Option<&HeaderFooter> {
        self.main_master_header_footer.as_ref()
    }

    /// Set a header/footer override attached directly to one slide.
    pub fn set_slide_header_footer(
        &mut self,
        slide: usize,
        value: HeaderFooter,
    ) -> Result<(), WriteError> {
        if slide >= self.slides.len() {
            return Err(WriteError::InvalidData(format!(
                "Slide {} does not exist",
                slide
            )));
        }
        let value = Self::validated_header_footer(
            value,
            HeaderFooterScope::Local {
                parent: HeaderFooterParent::Slide,
                parent_ordinal: HeaderFooterParentOrdinal::new(slide),
            },
        )?;
        self.slides[slide].header_footer = Some(value);
        Ok(())
    }

    /// Remove a header/footer override attached directly to one slide.
    pub fn clear_slide_header_footer(&mut self, slide: usize) -> Result<(), WriteError> {
        let slide = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        slide.header_footer = None;
        Ok(())
    }

    /// Return the header/footer override attached directly to one slide.
    pub fn slide_header_footer(&self, slide: usize) -> Result<Option<&HeaderFooter>, WriteError> {
        self.slides
            .get(slide)
            .map(|slide| slide.header_footer.as_ref())
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))
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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    #[allow(clippy::too_many_arguments)]
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
    #[allow(clippy::too_many_arguments)]
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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    #[allow(clippy::too_many_arguments)]
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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    /// The table is emitted as an OfficeArt table group (group shape with
    /// one rectangle cell shape per grid position), readable through the
    /// table extraction APIs after save/reopen.
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x`, `y` - Position of the table's top-left corner (in points)
    /// * `table` - Table grid, cell texts, and dimensions (see [`Table`])
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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
        let x = pt_to_emu_i32(x);
        let y = pt_to_emu_i32(y);
        let width = pt_to_emu_i32(width);
        let height = pt_to_emu_i32(height);
        x.checked_add(width).ok_or_else(|| {
            WriteError::InvalidData("chart frame horizontal extent is too large".to_string())
        })?;
        y.checked_add(height).ok_or_else(|| {
            WriteError::InvalidData("chart frame vertical extent is too large".to_string())
        })?;
        chart.validate()?;
        let total: usize = self.slides.iter().map(|slide| slide.charts.len()).sum();
        if total >= super::super::chart::MAX_CHART_OBJECTS {
            return Err(WriteError::InvalidData(format!(
                "presentation exceeds {} chart objects",
                super::super::chart::MAX_CHART_OBJECTS
            )));
        }
        self.slides
            .get(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        Err(litchi_ograph::Error::UnsupportedAuthoring {
            reason: "PPT chart creation requires the complete Office-compatible BIFF chart grammar",
        }
        .into())
    }

    /// Assign external-object and persist identifiers to every chart.
    ///
    /// Chart object identifiers continue above the hyperlink identifier seed
    /// because both share the `ExObjId` namespace ([MS-PPT] 2.10.1).
    pub(super) fn plan_charts(
        &self,
        persist_builder: &mut PersistPtrBuilder,
    ) -> Result<Vec<ChartPlan>, WriteError> {
        let total: usize = self.slides.iter().map(|slide| slide.charts.len()).sum();
        if total > super::super::chart::MAX_CHART_OBJECTS {
            return Err(WriteError::InvalidData(format!(
                "presentation exceeds {} chart objects",
                super::super::chart::MAX_CHART_OBJECTS
            )));
        }
        let mut next_id = self.hyperlinks.id_seed();
        let mut plans = Vec::with_capacity(total);
        for (slide_index, slide) in self.slides.iter().enumerate() {
            for chart_index in 0..slide.charts.len() {
                next_id = next_id.checked_add(1).ok_or_else(|| {
                    WriteError::InvalidData("external-object ID space exhausted".to_string())
                })?;
                plans.push(ChartPlan {
                    slide: slide_index,
                    chart: chart_index,
                    ex_obj_id: next_id,
                    persist_id: persist_builder.allocate_id(),
                });
            }
        }
        Ok(plans)
    }

    /// Add a picture to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x`, `y` - Position (in points)
    /// * `width`, `height` - Size (in points)
    /// * `image_data` - Raw image bytes (JPEG, PNG, etc.)
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
                "Slide {} does not exist",
                slide
            )));
        }
        // Add picture to BLIP store
        let blip_id = self.add_picture_data(image_data)?;

        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    #[allow(clippy::too_many_arguments)]
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
                "Slide {} does not exist",
                slide
            )));
        }
        let blip_id = self.add_picture_data_as(image_data, kind)?;

        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    pub fn add_picture_data(&mut self, image_data: Vec<u8>) -> Result<PictureId, WriteError> {
        self.blip_store.add(image_data).map_err(WriteError::Io)
    }

    /// Register explicitly typed picture data and return its checked ID.
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

    /// Attach a hyperlink to the last shape added on a slide
    pub fn set_last_shape_hyperlink(
        &mut self,
        slide: usize,
        hyperlink_id: u32,
    ) -> Result<(), WriteError> {
        let interaction =
            interaction_for_hyperlink(hyperlink_id, &self.hyperlinks).ok_or_else(|| {
                WriteError::InvalidData(format!("Hyperlink {hyperlink_id} does not exist"))
            })?;
        self.set_last_shape_interaction(slide, interaction)
    }

    /// Add or replace one typed click or mouse-over action on the last shape.
    ///
    /// Validation is atomic. Hyperlink references must identify an existing
    /// writer hyperlink, and a shape can carry at most one action per trigger.
    pub fn set_last_shape_interaction(
        &mut self,
        slide: usize,
        interaction: crate::Interaction,
    ) -> Result<(), WriteError> {
        self.set_last_shape_interaction_with_limits(
            slide,
            interaction,
            crate::InteractionLimits::default(),
        )
    }

    /// Add or replace one shape action with explicit record and name limits.
    pub fn set_last_shape_interaction_with_limits(
        &mut self,
        slide: usize,
        interaction: crate::Interaction,
        limits: crate::InteractionLimits,
    ) -> Result<(), WriteError> {
        interaction
            .validate_with_limits(limits)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        if interaction.hyperlink_id != 0 && self.hyperlinks.get(interaction.hyperlink_id).is_none()
        {
            return Err(WriteError::InvalidData(format!(
                "Hyperlink {} does not exist",
                interaction.hyperlink_id
            )));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        if let Some(shape) = slide_data.shapes.last_mut() {
            shape.properties.hyperlink_id = None;
            if let Some(existing) = shape
                .properties
                .interactions
                .iter_mut()
                .find(|existing| existing.trigger == interaction.trigger)
            {
                *existing = interaction;
            } else {
                shape.properties.interactions.push(interaction);
                shape.properties.interactions.sort_by_key(|interaction| {
                    match interaction.trigger {
                        crate::InteractionTrigger::Click => 0,
                        crate::InteractionTrigger::MouseOver => 1,
                    }
                });
            }
            Ok(())
        } else {
            Err(WriteError::InvalidData("No shapes on slide".to_string()))
        }
    }

    /// Remove one trigger from the last shape, returning whether it was present.
    pub fn clear_last_shape_interaction(
        &mut self,
        slide: usize,
        trigger: crate::InteractionTrigger,
    ) -> Result<bool, WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        if trigger == crate::InteractionTrigger::Click {
            shape.properties.hyperlink_id = None;
        }
        let old_len = shape.properties.interactions.len();
        shape
            .properties
            .interactions
            .retain(|interaction| interaction.trigger != trigger);
        Ok(shape.properties.interactions.len() != old_len)
    }

    /// Attach a registered hyperlink to one UTF-16 range in the last shape's text.
    pub fn set_last_shape_text_hyperlink(
        &mut self,
        slide: usize,
        range: crate::TextRange,
        hyperlink_id: u32,
    ) -> Result<(), WriteError> {
        let interaction =
            interaction_for_hyperlink(hyperlink_id, &self.hyperlinks).ok_or_else(|| {
                WriteError::InvalidData(format!("Hyperlink {hyperlink_id} does not exist"))
            })?;
        self.set_last_shape_text_interaction(
            slide,
            crate::TextInteraction::new(range, interaction)
                .map_err(|error| WriteError::InvalidData(error.to_string()))?,
        )
    }

    /// Add or replace one trigger/range pair on the last shape's text.
    ///
    /// Text positions are UTF-16 code units. Validation occurs before mutation.
    pub fn set_last_shape_text_interaction(
        &mut self,
        slide: usize,
        interaction: crate::TextInteraction,
    ) -> Result<(), WriteError> {
        self.set_last_shape_text_interaction_with_limits(
            slide,
            interaction,
            crate::TextInteractionLimits::default(),
        )
    }

    /// Add or replace a text action with explicit resource limits.
    pub fn set_last_shape_text_interaction_with_limits(
        &mut self,
        slide: usize,
        interaction: crate::TextInteraction,
        limits: crate::TextInteractionLimits,
    ) -> Result<(), WriteError> {
        if interaction.interaction.hyperlink_id != 0
            && self
                .hyperlinks
                .get(interaction.interaction.hyperlink_id)
                .is_none()
        {
            return Err(WriteError::InvalidData(format!(
                "Hyperlink {} does not exist",
                interaction.interaction.hyperlink_id
            )));
        }
        let slide_data = self
            .slides
            .get(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        let text_units = shape_text_unit_count(&shape.properties)?;
        interaction
            .validate_for_text(text_units, limits)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        let replace_index = shape
            .properties
            .text_interactions
            .iter()
            .position(|existing| {
                existing.range == interaction.range
                    && existing.interaction.trigger == interaction.interaction.trigger
            });
        let prospective_len = shape
            .properties
            .text_interactions
            .len()
            .checked_add(usize::from(replace_index.is_none()))
            .ok_or_else(|| {
                WriteError::InvalidData("Shape text interaction count overflows".to_string())
            })?;
        if prospective_len > limits.max_interactions {
            return Err(WriteError::InvalidData(
                "Shape exceeds the configured text interaction count".to_string(),
            ));
        }

        let shape = self
            .slides
            .get_mut(slide)
            .and_then(|slide| slide.shapes.last_mut())
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        if let Some(index) = replace_index {
            shape.properties.text_interactions[index] = interaction;
        } else {
            shape.properties.text_interactions.push(interaction);
            shape.properties.text_interactions.sort_by_key(|value| {
                (
                    value.range.begin(),
                    value.range.end(),
                    match value.interaction.trigger {
                        crate::InteractionTrigger::Click => 0,
                        crate::InteractionTrigger::MouseOver => 1,
                    },
                )
            });
        }
        Ok(())
    }

    /// Remove one trigger/range pair from the last shape.
    pub fn clear_last_shape_text_interaction(
        &mut self,
        slide: usize,
        range: crate::TextRange,
        trigger: crate::InteractionTrigger,
    ) -> Result<bool, WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        let old_len = shape.properties.text_interactions.len();
        shape
            .properties
            .text_interactions
            .retain(|value| value.range != range || value.interaction.trigger != trigger);
        Ok(shape.properties.text_interactions.len() != old_len)
    }

    /// Set the rotation of the last shape on a slide, in degrees.
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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        shape.properties.rotation = degrees;
        Ok(())
    }

    /// Set one of the ten OfficeArt adjustment values on the last shape.
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
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
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
    pub fn set_last_shape_text_alignment(
        &mut self,
        slide: usize,
        alignment: TextAlignment,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
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
            let alignment = TextAlign::from(alignment);
            for paragraph in paragraphs {
                paragraph.alignment = alignment;
            }
        }
        Ok(())
    }

    /// Add a font to the font collection and return its index
    pub fn add_font(&mut self, font: FontEntity) -> u16 {
        let index = self.fonts.len() as u16;
        self.fonts.push(font);
        index
    }

    /// Set slide notes (simple text)
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `notes` - Notes text
    pub fn set_slide_notes(&mut self, slide: usize, notes: &str) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        slide_data.notes = Some(notes.to_string());
        Ok(())
    }

    /// Set rich notes page for a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `notes_page` - Full notes page with formatting
    pub fn set_notes_page(
        &mut self,
        slide: usize,
        notes_page: NotesPage,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    pub fn set_shape_animation(
        &mut self,
        slide: usize,
        shape_index: usize,
        animation: AnimationInfo,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        let shape = slide_data.shapes.get_mut(shape_index).ok_or_else(|| {
            WriteError::InvalidData(format!(
                "Shape {} does not exist on slide {}",
                shape_index, slide
            ))
        })?;

        shape.animation_info = Some(animation);
        Ok(())
    }

    /// Get number of pictures in the presentation
    pub fn picture_count(&self) -> usize {
        self.blip_store.len()
    }

    /// Get number of hyperlinks in the presentation
    pub fn hyperlink_count(&self) -> usize {
        self.hyperlinks.len()
    }

    /// Get number of fonts in the presentation
    pub fn font_count(&self) -> usize {
        self.fonts.len()
    }

    /// Register an exact embedded WAV or AIFF resource for interactions or animations.
    ///
    /// Validation is atomic. The returned non-zero writer-local ID can be
    /// passed to [`crate::Interaction::with_sound_reference`].
    pub fn add_embedded_sound(
        &mut self,
        name: impl Into<String>,
        data: Vec<u8>,
    ) -> Result<std::num::NonZeroU32, WriteError> {
        if self.sound_resources.len()
            >= super::super::sound_collection::SoundCollectionLimits::default().max_sounds
        {
            return Err(WriteError::InvalidData(
                "sound collection exceeds the configured sound count".to_string(),
            ));
        }
        let next = self
            .next_sound_resource_id
            .checked_add(1)
            .ok_or_else(|| WriteError::InvalidData("sound resource ID overflow".to_string()))?;
        let id = std::num::NonZeroU32::new(self.next_sound_resource_id)
            .expect("writer sound IDs start above zero");
        let sound_type = crate::animation::SoundType::Embedded {
            name: name.into(),
            data,
        };
        let mut validator = super::super::sound_collection::SoundCollectionBuilder::new(
            super::super::sound_collection::SoundCollectionLimits::default(),
        );
        validator
            .register(id.get(), &sound_type)
            .map_err(sound_collection_error)?;

        self.sound_resources.insert(id.get(), sound_type);
        self.next_sound_resource_id = next;
        Ok(id)
    }

    /// Atomically replace one explicitly registered embedded sound.
    pub fn replace_embedded_sound(
        &mut self,
        id: std::num::NonZeroU32,
        name: impl Into<String>,
        data: Vec<u8>,
    ) -> Result<(), WriteError> {
        if !self.sound_resources.contains_key(&id.get()) {
            return Err(WriteError::InvalidData(format!(
                "sound resource {} does not exist",
                id
            )));
        }
        let sound_type = crate::animation::SoundType::Embedded {
            name: name.into(),
            data,
        };
        let mut validator = super::super::sound_collection::SoundCollectionBuilder::new(
            super::super::sound_collection::SoundCollectionLimits::default(),
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
    pub fn embedded_sound_count(&self) -> usize {
        self.sound_resources.len()
    }

    /// Add a comment to a slide.
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index (0-based)
    /// * `comment` - The comment to add
    pub fn add_comment(&mut self, slide: usize, comment: SlideComment) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        slide_data.comments.push(comment);
        Ok(())
    }

    /// Set per-slide timing (auto-advance, hidden, etc.).
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index (0-based)
    /// * `timing` - Timing configuration
    pub fn set_slide_timing(
        &mut self,
        slide: usize,
        timing: SlideTiming,
    ) -> Result<(), WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        slide_data.timing = Some(timing);
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
    pub fn custom_show_count(&self) -> usize {
        self.custom_shows.len()
    }

    fn validate_view_info_kind(view: &SlideViewInfo, expected: ViewKind) -> Result<(), WriteError> {
        view.to_bytes()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        if view.kind() != expected {
            return Err(WriteError::InvalidData(format!(
                "editing-view kind {:?} does not match {:?}",
                view.kind(),
                expected
            )));
        }
        Ok(())
    }

    /// Set the presentation's slide editing-view preferences, zoom, and guides.
    pub fn set_slide_view_info(&mut self, view: SlideViewInfo) -> Result<(), WriteError> {
        Self::validate_view_info_kind(&view, ViewKind::Slide)?;
        self.slide_view_info = Some(view);
        Ok(())
    }

    /// Restore the writer's canonical default slide editing view.
    pub fn clear_slide_view_info(&mut self) {
        self.slide_view_info = None;
    }

    /// Return the explicit slide editing-view override, if present.
    pub fn slide_view_info(&self) -> Option<&SlideViewInfo> {
        self.slide_view_info.as_ref()
    }

    /// Set the presentation's notes editing-view preferences, zoom, and guides.
    pub fn set_notes_view_info(&mut self, view: SlideViewInfo) -> Result<(), WriteError> {
        Self::validate_view_info_kind(&view, ViewKind::Notes)?;
        self.notes_view_info = Some(view);
        Ok(())
    }

    /// Remove the optional notes editing-view record.
    pub fn clear_notes_view_info(&mut self) {
        self.notes_view_info = None;
    }

    /// Return the explicit notes editing-view, if present.
    pub fn notes_view_info(&self) -> Option<&SlideViewInfo> {
        self.notes_view_info.as_ref()
    }

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
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }
}

impl Default for Writer {
    fn default() -> Self {
        Self::new()
    }
}
