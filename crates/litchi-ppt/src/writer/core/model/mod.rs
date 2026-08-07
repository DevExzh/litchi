//! Layered semantic state for the legacy PowerPoint writer core.
//!
//! This facade owns the typed writer state and caller-facing value models.
//! Contextual authoring operations live in the semantic module, validation and
//! atomic mutation gates live in the validation module, and binary snapshot
//! planning lives in the records module. The surrounding core codec and
//! package layers consume these typed snapshots without reaching into
//! format-neutral wire details.

mod records;
mod semantic;
mod validation;

use super::super::blip::{Id as PictureId, Pictures};
use super::super::chart::PositionedChart;
use super::super::comments::SlideComment;
use super::super::custom_shows::CustomShow;
use super::super::escher::FreeformGeometry;
use super::super::hyperlink::HyperlinkCollection;
use super::super::notes::NotesPage;
use super::super::shape_style::{FillStyle, LineStyleConfig, ShadowStyle};
use super::super::slide_timing::SlideTiming;
use super::super::smart_tags::SmartTagDefinition;
use super::super::table::PositionedTable;
use super::super::text_format::{Paragraph, TextAlign};
use crate::animation::AnimationInfo;
#[cfg(feature = "encryption")]
use crate::encryption::EncryptionProfile;
use crate::header_footer::HeaderFooter;
use crate::view_info::SlideViewInfo;
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
    #[cfg(feature = "vba-inspection")]
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

#[cfg(feature = "vba-inspection")]
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
            #[cfg(feature = "vba-inspection")]
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
            #[cfg(feature = "vba-inspection")]
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
    /// Typed base and PowerPoint 10 font collections.
    pub(super) fonts: crate::font::FontCollections,
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
    #[cfg(feature = "encryption")]
    encryption: Option<WriterEncryption>,
    /// Inert modify password, wiped on replacement, clear, or drop.
    modify_password: Option<WriterModifyPassword>,
    /// Standalone CFB project wrapped for a persisted `VbaProjectStg`.
    pub(super) vba_project: Option<crate::embedded::storage::Storage>,
}

#[cfg(feature = "encryption")]
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
