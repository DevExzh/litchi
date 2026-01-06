//! PPT file writer implementation
//!
//! This module provides functionality to create and modify Microsoft PowerPoint presentations
//! in the legacy binary format (.ppt files) using OLE2 structured storage.
//!
//! # Architecture
//!
//! The writer generates PPT record structures including:
//! - DocumentContainer - the root record container
//! - SlideListWithText - contains all slides
//! - MainMaster - master slide definitions
//! - Escher drawing records - for shapes and drawing objects
//! - PersistPtr - mapping system for record locations
//!
//! # PPT File Format Overview
//!
//! PPT files use a record-based binary format where:
//! 1. Everything is a record (atom or container)
//! 2. Containers hold other records
//! 3. Atoms contain actual data
//! 4. Escher (Office Drawing) format for shapes
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ole::ppt::PptWriter;
//!
//! let mut writer = PptWriter::new();
//!
//! // Add a slide
//! let slide = writer.add_slide()?;
//!
//! // Add a text box
//! writer.add_textbox(slide, 100, 100, 400, 200, "Hello, World!")?;
//!
//! // Save the presentation
//! writer.save("output.ppt")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use super::blip::{BlipStoreBuilder, BlipType};
use super::escher::{
    UserShapeData, create_dg_container_with_shapes, create_dgg_container,
    shape_type as escher_shape_type,
};
use super::hyperlink::{Hyperlink, HyperlinkCollection};
use super::master_drawing::build_master_ppdrawing;
use super::notes::{NotesContainerBuilder, NotesPage};
use super::persist::{PersistPtrBuilder, UserEditAtom};
use super::records::{
    RecordBuilder, create_docinfo_list_container_minimal, create_document_atom,
    create_end_document, create_environment_minimal, create_main_master_container,
    create_slide_list_with_text_master, record_type, wrap_dg_into_ppdrawing,
    wrap_dgg_into_ppdrawing_group,
};
use super::shape_style::{ArrowStyle, FillStyle, LineStyleConfig, ShadowStyle, ShapeStyle};
#[allow(unused_imports)]
use super::shapes::ShapeKind;
use super::spec::{BinaryTagData, ColorScheme, Ppt10Tag, SlideLayoutType, slide_flags};
use super::text_format::{FontEntity, Paragraph};
use crate::common::unit::pt_to_emu_i32;
use crate::ole::writer::OleWriter;
use std::collections::HashMap;

/// Error type for PPT writing
#[derive(Debug)]
pub enum PptWriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(crate::ole::OleError),
}

/// Build a minimal, valid Current User stream referencing the given UserEditAtom offset.
fn build_current_user_stream(offset_to_current_edit: u32) -> Vec<u8> {
    // Build per Apache POI CurrentUserAtom:
    // [0..3]   atomHeader = {0x00,0x00,0xF6,0x0F}
    // [4..7]   atomSize = 20 + 4 + lenAsciiUser (we use 0) => 24
    // [8..11]  details size = 20
    // [12..15] headerToken (unencrypted) = 0xE391C05F (bytes {95,-64,-111,-29})
    // [16..19] offsetToCurrentEdit
    // [20..21] lenUserName (ANSI) = 0
    // [22..23] docFinalVersion = 0x03F4
    // [24]     docMajorNo = 3
    // [25]     docMinorNo = 0
    // [26..27] reserved = 0
    // [28..31] releaseVersion = 8
    // [32..]   unicode username (2*len) (none)
    let mut s = Vec::with_capacity(32);
    // atomHeader
    s.extend_from_slice(&[0x00, 0x00, 0xF6, 0x0F]);
    // atomSize (20 + 4 + lenAsciiUsername)
    s.extend_from_slice(&24u32.to_le_bytes());
    // details size (20)
    s.extend_from_slice(&20u32.to_le_bytes());
    // headerToken (unencrypted)
    s.extend_from_slice(&0xE391C05Fu32.to_le_bytes());
    // current edit offset
    s.extend_from_slice(&offset_to_current_edit.to_le_bytes());
    // username length (ANSI)
    s.extend_from_slice(&0u16.to_le_bytes());
    // doc final version
    s.extend_from_slice(&0x03F4u16.to_le_bytes());
    // major/minor
    s.push(3u8);
    s.push(0u8);
    // reserved
    s.extend_from_slice(&[0u8; 2]);
    // release version
    s.extend_from_slice(&8u32.to_le_bytes());
    // no username
    s
}

fn build_summary_information_stream() -> Vec<u8> {
    let mut s = Vec::new();
    s.extend_from_slice(&0xFFFEu16.to_le_bytes());
    s.extend_from_slice(&0u16.to_le_bytes());
    s.extend_from_slice(&0u32.to_le_bytes());
    s.extend_from_slice(&[0u8; 16]);
    s.extend_from_slice(&1u32.to_le_bytes());
    let fmtid: [u8; 16] = [
        0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3,
        0xD9,
    ];
    s.extend_from_slice(&fmtid);
    let section_offset = 48u32;
    s.extend_from_slice(&section_offset.to_le_bytes());
    let mut section = Vec::new();
    section.extend_from_slice(&0u32.to_le_bytes());
    section.extend_from_slice(&1u32.to_le_bytes());
    section.extend_from_slice(&1u32.to_le_bytes());
    section.extend_from_slice(&16u32.to_le_bytes());
    section.extend_from_slice(&2u16.to_le_bytes());
    section.extend_from_slice(&0u16.to_le_bytes());
    section.extend_from_slice(&(1252i16).to_le_bytes());
    section.extend_from_slice(&0i16.to_le_bytes());
    let size = section.len() as u32;
    section[0..4].copy_from_slice(&size.to_le_bytes());
    s.extend_from_slice(&section);
    s
}

fn build_document_summary_information_stream() -> Vec<u8> {
    let mut s = Vec::new();
    s.extend_from_slice(&0xFFFEu16.to_le_bytes());
    s.extend_from_slice(&0u16.to_le_bytes());
    s.extend_from_slice(&0u32.to_le_bytes());
    s.extend_from_slice(&[0u8; 16]);
    s.extend_from_slice(&1u32.to_le_bytes());
    let fmtid: [u8; 16] = [
        0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9,
        0xAE,
    ];
    s.extend_from_slice(&fmtid);
    let section_offset = 48u32;
    s.extend_from_slice(&section_offset.to_le_bytes());
    let mut section = Vec::new();
    section.extend_from_slice(&0u32.to_le_bytes());
    section.extend_from_slice(&0u32.to_le_bytes());
    let size = section.len() as u32;
    section[0..4].copy_from_slice(&size.to_le_bytes());
    s.extend_from_slice(&section);
    s
}

impl From<std::io::Error> for PptWriteError {
    fn from(err: std::io::Error) -> Self {
        PptWriteError::Io(err)
    }
}

impl From<crate::ole::OleError> for PptWriteError {
    fn from(err: crate::ole::OleError) -> Self {
        PptWriteError::Ole(err)
    }
}

impl std::fmt::Display for PptWriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PptWriteError::Io(e) => write!(f, "I/O error: {}", e),
            PptWriteError::InvalidData(s) => write!(f, "Invalid data: {}", s),
            PptWriteError::Ole(e) => write!(f, "OLE error: {}", e),
        }
    }
}

impl std::error::Error for PptWriteError {}

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
    /// Flip horizontal
    pub flip_h: bool,
    /// Flip vertical
    pub flip_v: bool,
    /// Picture BLIP index (for Picture type)
    pub picture_index: Option<u32>,
    /// Hyperlink attached to shape
    pub hyperlink_id: Option<u32>,
}

/// Represents a shape on a slide
#[derive(Debug, Clone)]
struct WritableShape {
    /// Shape properties
    properties: ShapeProperties,
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
            flip_h: false,
            flip_v: false,
            picture_index: None,
            hyperlink_id: None,
        }
    }
}

/// Represents a slide
#[derive(Debug, Clone)]
struct WritableSlide {
    /// Shapes on this slide
    shapes: Vec<WritableShape>,
    /// Slide notes text (simple)
    notes: Option<String>,
    /// Rich notes page
    notes_page: Option<NotesPage>,
}

/// Convert ShapeType to Escher MSOSPT value
fn shape_type_to_escher(shape_type: ShapeType) -> u16 {
    match shape_type {
        ShapeType::Rectangle => escher_shape_type::RECTANGLE,
        ShapeType::TextBox => escher_shape_type::TEXT_BOX,
        ShapeType::Placeholder => escher_shape_type::RECTANGLE,
        ShapeType::Line => escher_shape_type::LINE,
        ShapeType::Ellipse => escher_shape_type::ELLIPSE,
        ShapeType::RoundRectangle => escher_shape_type::ROUND_RECTANGLE,
        ShapeType::Diamond => escher_shape_type::DIAMOND,
        ShapeType::Triangle => 5, // TRIANGLE
        ShapeType::Arrow => 13,   // ARROW
        ShapeType::Star => 12,    // STAR
        ShapeType::Heart => 74,   // HEART
        ShapeType::Picture => 75, // FRAME (PictureFrame) per POI HSLFPictureShape
    }
}

/// Convert WritableShape to UserShapeData for Escher serialization
fn convert_shape_to_escher(
    shape: &WritableShape,
    hyperlinks: &HyperlinkCollection,
) -> UserShapeData {
    let props = &shape.properties;

    // Extract fill properties from FillStyle
    let (fill_color, fill_type, fill_opacity, fill_back_color, fill_angle) = props
        .fill
        .as_ref()
        .map_or((None, None, None, None, None), |fill| {
            if !fill.enabled {
                return (None, None, None, None, None);
            }

            let color = Some(fill.color.to_rgbx());
            let fill_type = Some(fill.fill_type as u32);

            // Opacity: convert 0-100 to 0-65536
            let opacity = if fill.opacity < 100 {
                Some(((fill.opacity as u32) * 65536) / 100)
            } else {
                None
            };

            // Back color for gradients
            let back_color = fill.back_color.as_ref().map(|c| c.to_rgbx());

            // Gradient angle (degrees * 65536)
            // Per Apache POI HSLFFill.java: "Zero degrees represents a vertical vector from bottom to top"
            // Standard: 0° = horizontal right, 90° = vertical up
            // PPT format: 0° = vertical up, so we need: PPT_angle = 90 - user_angle
            let angle = fill.gradient_angle.map(|a| ((90 - a) as i32) * 65536);

            (color, fill_type, opacity, back_color, angle)
        });

    // Extract line color, width, dash style, and arrows from LineStyleConfig
    let (line_color, line_width, line_dash_style, line_start_arrow, line_end_arrow) = props
        .line
        .as_ref()
        .map_or((None, None, None, None, None), |line| {
            if line.width > 0 && line.enabled {
                let dash = match line.dash {
                    super::shape_style::LineDashStyle::Solid => None,
                    _ => Some(line.dash as u32),
                };
                let start_arrow = if line.start_arrow != super::shape_style::ArrowStyle::None {
                    Some(line.start_arrow as u32)
                } else {
                    None
                };
                let end_arrow = if line.end_arrow != super::shape_style::ArrowStyle::None {
                    Some(line.end_arrow as u32)
                } else {
                    None
                };
                (
                    Some(line.color.to_rgbx()),
                    Some(line.width as i32),
                    dash,
                    start_arrow,
                    end_arrow,
                )
            } else {
                (None, None, None, None, None)
            }
        });

    // Extract shadow properties from ShadowStyle
    let (has_shadow, shadow_color, shadow_offset_x, shadow_offset_y, shadow_opacity, shadow_type) =
        props
            .shadow
            .as_ref()
            .map_or((false, None, None, None, None, None), |shadow| {
                if !shadow.enabled {
                    (false, None, None, None, None, None)
                } else {
                    (
                        true,
                        Some(shadow.color.to_rgbx()),
                        Some(shadow.offset_x),
                        Some(shadow.offset_y),
                        Some(((shadow.opacity as u32) * 65536) / 100),
                        Some(shadow.shadow_type as u32),
                    )
                }
            });

    // Get text content - prefer paragraphs with formatting
    let paragraphs = props.paragraphs.clone();
    let text = if paragraphs.is_some() {
        None // Don't use plain text if paragraphs are available
    } else {
        props.text.clone()
    };

    UserShapeData {
        shape_type: shape_type_to_escher(props.shape_type),
        x: props.x,
        y: props.y,
        width: props.width,
        height: props.height,
        fill_color,
        fill_type,
        fill_opacity,
        fill_back_color,
        fill_angle,
        line_color,
        line_width,
        line_dash_style,
        line_start_arrow,
        line_end_arrow,
        text,
        paragraphs,
        text_type: 4,           // OTHER for regular shapes
        placeholder_type: None, // Not a placeholder for regular shapes
        has_shadow,
        flip_h: props.flip_h,
        flip_v: props.flip_v,
        hyperlink_id: props.hyperlink_id,
        hyperlink_action: get_hyperlink_info(props.hyperlink_id, hyperlinks).0,
        hyperlink_jump: get_hyperlink_info(props.hyperlink_id, hyperlinks).1,
        hyperlink_type: get_hyperlink_info(props.hyperlink_id, hyperlinks).2,
        picture_index: props.picture_index,
        shadow_color,
        shadow_offset_x,
        shadow_offset_y,
        shadow_opacity,
        shadow_type,
    }
}

/// Get hyperlink interactive info values based on hyperlink target
/// Returns (action, jump, hyperlink_type)
fn get_hyperlink_info(hyperlink_id: Option<u32>, hyperlinks: &HyperlinkCollection) -> (u8, u8, u8) {
    use super::hyperlink::HyperlinkTarget;

    // Defaults for URL links: ACTION_HYPERLINK=4, JUMP_NONE=0, LINK_Url=8
    let Some(id) = hyperlink_id else {
        return (4, 0, 8);
    };

    let Some(hyperlink) = hyperlinks.get(id) else {
        return (4, 0, 8);
    };

    // Per POI HSLFHyperlink:
    // - URL/File links: action=ACTION_HYPERLINK(4), jump=JUMP_NONE(0), hyperlinkType=LINK_Url(8)
    // - Slide number: action=ACTION_HYPERLINK(4), jump=JUMP_NONE(0), hyperlinkType=LINK_SlideNumber(3)
    // - Next/Prev/First/Last: action=ACTION_JUMP(3), jump=varies, hyperlinkType=varies
    match &hyperlink.target {
        HyperlinkTarget::Url(_) | HyperlinkTarget::File(_) => (4, 0, 8), // ACTION_HYPERLINK, JUMP_NONE, LINK_Url
        HyperlinkTarget::Slide(_) => (4, 0, 3), // ACTION_HYPERLINK (not JUMP!), JUMP_NONE, LINK_SlideNumber
        HyperlinkTarget::NextSlide => (3, 1, 1), // ACTION_JUMP, JUMP_NEXTSLIDE, LINK_NextSlide
        HyperlinkTarget::PrevSlide => (3, 2, 2), // ACTION_JUMP, JUMP_PREVIOUSSLIDE, LINK_PreviousSlide
        HyperlinkTarget::FirstSlide => (3, 3, 3), // ACTION_JUMP, JUMP_FIRSTSLIDE, LINK_FirstSlide
        HyperlinkTarget::LastSlide => (3, 4, 4), // ACTION_JUMP, JUMP_LASTSLIDE, LINK_LastSlide
        HyperlinkTarget::EndShow => (3, 6, 0xFF), // ACTION_JUMP, JUMP_ENDSHOW, LINK_NULL
        HyperlinkTarget::CustomShow(_) => (7, 0, 5), // ACTION_CUSTOMSHOW, JUMP_NONE, LINK_CustomShow
    }
}

/// PPT file writer
///
/// Provides methods to create and modify PPT files with full support for:
/// - Shapes with fill, line, and shadow styling
/// - Rich text formatting (bold, italic, colors, sizes)
/// - Pictures/images
/// - Hyperlinks
/// - Speaker notes
pub struct PptWriter {
    /// Slides in the presentation
    slides: Vec<WritableSlide>,
    /// Presentation properties
    properties: HashMap<String, String>,
    /// Slide width in EMUs (default: Letter size)
    slide_width: i32,
    /// Slide height in EMUs (default: Letter size)
    slide_height: i32,
    /// Picture/BLIP storage
    blip_store: BlipStoreBuilder,
    /// Hyperlink collection
    hyperlinks: HyperlinkCollection,
    /// Font collection
    fonts: Vec<FontEntity>,
}

impl PptWriter {
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
            blip_store: BlipStoreBuilder::new(),
            hyperlinks: HyperlinkCollection::new(),
            fonts: vec![FontEntity::arial()], // Default font
        }
    }

    /// Add a new blank slide
    ///
    /// # Returns
    ///
    /// * `Result<usize, PptWriteError>` - Slide index or error
    pub fn add_slide(&mut self) -> Result<usize, PptWriteError> {
        let index = self.slides.len();
        self.slides.push(WritableSlide {
            shapes: Vec::new(),
            notes: None,
            notes_page: None,
        });
        Ok(index)
    }

    /// Delete a slide
    ///
    /// # Arguments
    ///
    /// * `index` - Slide index (0-based)
    pub fn delete_slide(&mut self, index: usize) -> Result<(), PptWriteError> {
        if index >= self.slides.len() {
            return Err(PptWriteError::InvalidData(format!(
                "Slide {} does not exist",
                index
            )));
        }
        self.slides.remove(index);
        Ok(())
    }

    /// Move a slide to a new position
    ///
    /// # Arguments
    ///
    /// * `from_index` - Current slide index
    /// * `to_index` - New slide index
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<(), PptWriteError> {
        if from_index >= self.slides.len() || to_index >= self.slides.len() {
            return Err(PptWriteError::InvalidData(
                "Invalid slide index".to_string(),
            ));
        }

        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        Ok(())
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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    ) -> Result<(), PptWriteError> {
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
    ) -> Result<(), PptWriteError> {
        self.add_shape(slide, ShapeType::Ellipse, x, y, width, height)
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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                ..Default::default()
            },
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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
        };

        slide_data.shapes.push(shape);
        Ok(())
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
    ) -> Result<(), PptWriteError> {
        // Add picture to BLIP store
        let blip_index = self.blip_store.add_picture(image_data);

        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Picture,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                picture_index: Some(blip_index),
                fill: Some(FillStyle::picture(blip_index)),
                ..Default::default()
            },
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a picture with explicit type
    #[allow(clippy::too_many_arguments)]
    pub fn add_picture_with_type(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        image_data: Vec<u8>,
        blip_type: BlipType,
    ) -> Result<(), PptWriteError> {
        let blip_index = self.blip_store.add_picture_with_type(image_data, blip_type);

        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Picture,
                x: pt_to_emu_i32(x),
                y: pt_to_emu_i32(y),
                width: pt_to_emu_i32(width),
                height: pt_to_emu_i32(height),
                picture_index: Some(blip_index),
                fill: Some(FillStyle::picture(blip_index)),
                ..Default::default()
            },
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a hyperlink and return its ID
    ///
    /// The returned ID can be used with `add_shape_hyperlink` to attach
    /// the hyperlink to a shape.
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) -> u32 {
        self.hyperlinks.add(hyperlink)
    }

    /// Attach a hyperlink to the last shape added on a slide
    pub fn set_last_shape_hyperlink(
        &mut self,
        slide: usize,
        hyperlink_id: u32,
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        if let Some(shape) = slide_data.shapes.last_mut() {
            shape.properties.hyperlink_id = Some(hyperlink_id);
            Ok(())
        } else {
            Err(PptWriteError::InvalidData("No shapes on slide".to_string()))
        }
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
    pub fn set_slide_notes(&mut self, slide: usize, notes: &str) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        slide_data.notes_page = Some(notes_page);
        Ok(())
    }

    /// Get number of pictures in the presentation
    pub fn picture_count(&self) -> usize {
        self.blip_store.count()
    }

    /// Get number of hyperlinks in the presentation
    pub fn hyperlink_count(&self) -> usize {
        self.hyperlinks.len()
    }

    /// Get number of fonts in the presentation
    pub fn font_count(&self) -> usize {
        self.fonts.len()
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

    /// Save the presentation to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), PptWriteError>` - Success or error
    ///
    /// # Implementation
    ///
    /// This generates a complete PowerPoint 97-2003 binary file conforming to MS-PPT specification:
    /// - PPT record structures - [MS-PPT] Section 2.3
    /// - Escher drawing containers - [MS-ODRAW] Section 2.2
    /// - PersistPtr directory - [MS-PPT] Section 2.4.16
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<(), PptWriteError> {
        // 1) We'll write DocumentContainer at stream offset 0
        let mut ppt_stream = Vec::new();
        let mut persist_builder = PersistPtrBuilder::new();

        // Allocate a persist ID for the Document itself and set its offset to 0
        let doc_persist_id = persist_builder.allocate_id();
        persist_builder.set_offset(doc_persist_id, 0);
        // Allocate persist ID for MainMaster (top-level record written after Document)
        let master_persist_id = persist_builder.allocate_id();

        // 2) Build DocumentContainer
        let mut doc_container = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

        // 2.1) DocumentAtom
        let doc_atom = create_document_atom(
            self.slide_width as u32,
            self.slide_height as u32,
            self.slides.len() as u32,
            0,
            0,
        )?;
        doc_container.write_child(&doc_atom);

        // 2.2) Environment (with FontCollection)
        let env = create_environment_minimal()?;
        doc_container.write_child(&env);

        // 2.3) PPDrawingGroup wrapping Dgg Escher
        // Calculate per-slide shape counts (group + background + user shapes)
        let master_shapes = 6u32;
        let slide_shape_counts: Vec<u32> = self
            .slides
            .iter()
            .map(|s| 2 + s.shapes.len() as u32) // 2 for group+background, plus user shapes
            .collect();
        // Build DggContainer with BStore if pictures are present
        let dgg = if !self.blip_store.is_empty() {
            let bstore = self.blip_store.build().map_err(PptWriteError::Io)?;
            super::escher::create_dgg_container_with_blips(
                master_shapes,
                &slide_shape_counts,
                &bstore,
            )?
        } else {
            create_dgg_container(master_shapes, &slide_shape_counts)?
        };
        let pp_dgg = wrap_dgg_into_ppdrawing_group(&dgg)?;
        doc_container.write_child(&pp_dgg);

        // 2.3.1) SlideListWithText for masters (instance=1) referencing MainMaster
        let master_entries = vec![(master_persist_id, 0x8000_0000u32)];
        let slwt_master = create_slide_list_with_text_master(&master_entries)?;
        doc_container.write_child(&slwt_master);

        // 2.4) DocInfo List (0x07D0) before SlideListWithText (slides), per POI empty_textbox.ppt
        let docinfo = create_docinfo_list_container_minimal()?;
        doc_container.write_child(&docinfo);

        // 2.5) SlideListWithText (SLIDES) referencing each slide by (persist id ref, slide identifier)
        let mut slide_persist_ids = Vec::with_capacity(self.slides.len());
        let mut slwt_entries = Vec::with_capacity(self.slides.len());
        for (i, _slide) in self.slides.iter().enumerate() {
            let pid = persist_builder.allocate_id();
            slide_persist_ids.push(pid);
            let slide_identifier = 256u32 + (i as u32);
            slwt_entries.push((pid, slide_identifier));
        }
        if !slwt_entries.is_empty() {
            use super::records::create_slide_list_with_text_slides;
            let slwt = create_slide_list_with_text_slides(&slwt_entries)?;
            doc_container.write_child(&slwt);
        }

        // 2.5.1) Pre-allocate notes persist IDs and build SlideListWithText for notes
        // Per POI: Notes' SlidePersistAtom.slideIdentifier must match Slide's slideIdentifier
        // This is how POI matches notes to slides in findNotesSlides/findSlides
        let mut notes_persist_ids: Vec<Option<u32>> = vec![None; self.slides.len()];
        let mut notes_slwt_entries = Vec::new();
        for (i, slide) in self.slides.iter().enumerate() {
            let has_notes =
                slide.notes.as_ref().is_some_and(|n| !n.is_empty()) || slide.notes_page.is_some();
            if has_notes {
                let notes_pid = persist_builder.allocate_id();
                notes_persist_ids[i] = Some(notes_pid);
                // Use SAME slideIdentifier as the slide (256 + i) for matching!
                let slide_identifier = 256u32 + (i as u32);
                notes_slwt_entries.push((notes_pid, slide_identifier));
            }
        }
        if !notes_slwt_entries.is_empty() {
            use super::records::create_slide_list_with_text_notes;
            let slwt_notes = create_slide_list_with_text_notes(&notes_slwt_entries)?;
            doc_container.write_child(&slwt_notes);
        }

        // 2.5.2) ExObjList for hyperlinks (if any)
        let ex_obj_list = self.hyperlinks.build_ex_obj_list()?;
        if !ex_obj_list.is_empty() {
            doc_container.write_child(&ex_obj_list);
        }

        // 2.6) EndDocument
        let end_doc = create_end_document()?;
        doc_container.write_child(&end_doc);

        // Finalize DocumentContainer and write to stream (offset 0)
        let doc_bytes = doc_container.build()?;
        ppt_stream.extend_from_slice(&doc_bytes);

        // 3) MainMaster then Slides (top-level after DocumentContainer)
        // 3.1) Write MainMaster using dynamically built PPDrawing (includes all placeholders)
        let master_ppdrawing = build_master_ppdrawing();
        let master_container = create_main_master_container(&master_ppdrawing)?;
        let master_offset = ppt_stream.len() as u32;
        persist_builder.set_offset(master_persist_id, master_offset);
        ppt_stream.extend_from_slice(&master_container);

        // 3.2) Slides
        for (i, slide) in self.slides.iter().enumerate() {
            // drawing_id for slides starts from 2 (1 is used by MainMaster)
            let drawing_id = (i as u32) + 2;
            let slide_identifier = 256u32 + (i as u32);

            // Build Slide container with SlideAtom
            let mut slide_container = RecordBuilder::new(0x0F, 0, record_type::SLIDE);
            // SlideAtom (MS-PPT 2.4.7)
            let mut slide_atom = RecordBuilder::new(0x02, 0, record_type::SLIDE_ATOM);
            let mut atom_data = Vec::with_capacity(24);
            // SSlideLayoutAtom: geometry + placeholder types
            atom_data.extend_from_slice(&(SlideLayoutType::Blank as u32).to_le_bytes());
            atom_data.extend_from_slice(&[0u8; 8]); // rgPlaceholderTypes
            // masterIdRef (0x80000000 = reference to master)
            atom_data.extend_from_slice(&0x8000_0000u32.to_le_bytes());
            // notesIdRef: Per POI, this equals NotesAtom.slideID = slideIdentifier
            // Set to the slide's own identifier if notes exist, 0 otherwise
            let notes_id_ref = if notes_persist_ids[i].is_some() {
                slide_identifier // Same value as NotesAtom.slideID
            } else {
                0
            };
            atom_data.extend_from_slice(&notes_id_ref.to_le_bytes());
            // slideFlags: follow master objects/scheme/background
            atom_data.extend_from_slice(&slide_flags::DEFAULT.to_le_bytes());
            atom_data.extend_from_slice(&0u16.to_le_bytes()); // reserved
            slide_atom.write_data(&atom_data);
            slide_container.write_child(&slide_atom.build()?);

            // PPDrawing with Escher DgContainer (including user shapes)
            let escher_shapes: Vec<UserShapeData> = slide
                .shapes
                .iter()
                .map(|s| convert_shape_to_escher(s, &self.hyperlinks))
                .collect();
            let dg = create_dg_container_with_shapes(drawing_id, &escher_shapes)?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // ProgTags with PPT10 binary tag (PowerPoint 2002+ features)
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Ppt10Tag::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&BinaryTagData::SLIDE.to_bytes());
            prog_bin.write_child(&bin.build()?);
            prog_tags.write_child(&prog_bin.build()?);
            slide_container.write_child(&prog_tags.build()?);

            // Compute this slide's offset in the stream: current top-level length
            let slide_offset = ppt_stream.len() as u32;

            // Track persist pointer (allocate new persist id per slide)
            let persist_id = slide_persist_ids[i];
            persist_builder.set_offset(persist_id, slide_offset);

            // Append slide as top-level record
            let slide_bytes = slide_container.build()?;
            ppt_stream.extend_from_slice(&slide_bytes);
        }

        // 3.3) Notes containers for slides with notes
        for (i, slide) in self.slides.iter().enumerate() {
            if let Some(notes_pid) = notes_persist_ids[i] {
                let notes_offset = ppt_stream.len() as u32;
                persist_builder.set_offset(notes_pid, notes_offset);

                // Per POI: NotesAtom.slideID = slideIdentifier (same as slide's identifier)
                // This equals SlideAtom.notesID and Notes' SlidePersistAtom.slideIdentifier
                let slide_identifier = 256u32 + (i as u32);
                let notes_page = if let Some(page) = &slide.notes_page {
                    let mut page = page.clone();
                    page.slide_id_ref = slide_identifier;
                    page
                } else if let Some(text) = &slide.notes {
                    NotesPage::simple(slide_identifier, text)
                } else {
                    continue;
                };

                // Build notes container (drawing_id continues after slides)
                let notes_drawing_id = (self.slides.len() as u32) + 2 + (i as u32);
                let notes_builder = NotesContainerBuilder::new(notes_page, notes_drawing_id);
                let notes_bytes = notes_builder.build().map_err(std::io::Error::other)?;
                ppt_stream.extend_from_slice(&notes_bytes);
            }
        }

        // 4) PersistPtrIncrementalBlock (6002) then single UserEditAtom
        let persist_dir_offset = ppt_stream.len() as u32;
        let persist_dir_block = persist_builder.generate_record();
        ppt_stream.extend_from_slice(&persist_dir_block);

        let user_edit = UserEditAtom::new_minimal(
            persist_dir_offset,
            doc_persist_id,
            persist_builder.persist_id_seed(),
            self.slides.len() as u32,
        );
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        // 5) Build Current User and property streams
        let current_user = build_current_user_stream(user_edit_offset);
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

        // 6) Write OLE streams
        let mut ole_writer = OleWriter::new();
        // Set root CLSID to PowerPoint V8
        ole_writer.set_root_clsid([
            0x10, 0x8D, 0x81, 0x64, 0x9B, 0x4F, 0xCF, 0x11, 0x86, 0xEA, 0x00, 0xAA, 0x00, 0xB9,
            0x29, 0xE8,
        ]);
        ole_writer.create_stream(&["PowerPoint Document"], &ppt_stream)?;
        ole_writer.create_stream(&["Current User"], &current_user)?;
        ole_writer.create_stream(&["\u{0005}SummaryInformation"], &summary_info)?;
        ole_writer.create_stream(&["\u{0005}DocumentSummaryInformation"], &doc_summary)?;

        // Pictures stream (per POI: separate stream for BLIP data)
        if !self.blip_store.is_empty() {
            let pictures_stream = self
                .blip_store
                .build_pictures_stream()
                .map_err(PptWriteError::Io)?;
            ole_writer.create_stream(&["Pictures"], &pictures_stream)?;
        }

        ole_writer.save(path)?;

        Ok(())
    }

    /// Write presentation to an in-memory buffer
    ///
    /// # Arguments
    ///
    /// * `writer` - Output writer (must support Write + Seek)
    ///
    /// # Returns
    ///
    /// * `Result<(), PptWriteError>` - Success or error
    pub fn write_to<W: std::io::Write + std::io::Seek>(
        &mut self,
        writer: &mut W,
    ) -> Result<(), PptWriteError> {
        // Same logic as save(), but writing to provided writer
        let mut ppt_stream = Vec::new();
        let mut persist_builder = PersistPtrBuilder::new();

        let doc_persist_id = persist_builder.allocate_id();
        persist_builder.set_offset(doc_persist_id, 0);
        // Allocate persist ID for MainMaster
        let master_persist_id = persist_builder.allocate_id();

        let mut doc_container = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

        let doc_atom = create_document_atom(
            self.slide_width as u32,
            self.slide_height as u32,
            self.slides.len() as u32,
            0,
            0,
        )?;
        doc_container.write_child(&doc_atom);
        // 2.2) Environment (with FontCollection)
        let env = create_environment_minimal()?;
        doc_container.write_child(&env);

        // 2.3) PPDrawingGroup wrapping Dgg Escher
        // Calculate per-slide shape counts (group + background + user shapes)
        let master_shapes = 6u32;
        let slide_shape_counts: Vec<u32> = self
            .slides
            .iter()
            .map(|s| 2 + s.shapes.len() as u32)
            .collect();
        // Build DggContainer with BStore if pictures are present
        let dgg = if !self.blip_store.is_empty() {
            let bstore = self.blip_store.build().map_err(PptWriteError::Io)?;
            super::escher::create_dgg_container_with_blips(
                master_shapes,
                &slide_shape_counts,
                &bstore,
            )?
        } else {
            create_dgg_container(master_shapes, &slide_shape_counts)?
        };
        let pp_dgg = wrap_dgg_into_ppdrawing_group(&dgg)?;
        doc_container.write_child(&pp_dgg);

        // 2.3.1) SlideListWithText for masters (instance=1)
        let master_entries = vec![(master_persist_id, 0x8000_0000u32)];
        let slwt_master = create_slide_list_with_text_master(&master_entries)?;
        doc_container.write_child(&slwt_master);

        // DocInfo List before SlideListWithText (slides), matching POI empty_textbox.ppt
        let docinfo = create_docinfo_list_container_minimal()?;
        doc_container.write_child(&docinfo);

        // SlideListWithText (SLIDES) for non-empty presentations
        let mut slide_persist_ids = Vec::with_capacity(self.slides.len());
        let mut slwt_entries = Vec::with_capacity(self.slides.len());
        for (i, _slide) in self.slides.iter().enumerate() {
            let pid = persist_builder.allocate_id();
            slide_persist_ids.push(pid);
            let slide_identifier = 256u32 + (i as u32);
            slwt_entries.push((pid, slide_identifier));
        }
        if !slwt_entries.is_empty() {
            use super::records::create_slide_list_with_text_slides;
            let slwt = create_slide_list_with_text_slides(&slwt_entries)?;
            doc_container.write_child(&slwt);
        }

        // ExObjList for hyperlinks (if any)
        let ex_obj_list = self.hyperlinks.build_ex_obj_list()?;
        if !ex_obj_list.is_empty() {
            doc_container.write_child(&ex_obj_list);
        }

        let end_doc = create_end_document()?;
        doc_container.write_child(&end_doc);

        // Write finalized DocumentContainer
        let doc_bytes = doc_container.build()?;
        ppt_stream.extend_from_slice(&doc_bytes);

        // Then write MainMaster and slides as top-level records
        // MainMaster using dynamically built PPDrawing (includes all placeholders)
        let master_ppdrawing = build_master_ppdrawing();
        let master_container = create_main_master_container(&master_ppdrawing)?;
        let master_offset = ppt_stream.len() as u32;
        persist_builder.set_offset(master_persist_id, master_offset);
        ppt_stream.extend_from_slice(&master_container);

        // Slides
        for (i, slide) in self.slides.iter().enumerate() {
            let drawing_id = (i as u32) + 2; // 1 reserved for master

            let mut slide_container = RecordBuilder::new(0x0F, 0, record_type::SLIDE);
            // SlideAtom (MS-PPT 2.4.7)
            let mut slide_atom = RecordBuilder::new(0x02, 0, record_type::SLIDE_ATOM);
            let mut atom_data = Vec::with_capacity(24);
            atom_data.extend_from_slice(&(SlideLayoutType::Blank as u32).to_le_bytes());
            atom_data.extend_from_slice(&[0u8; 8]); // rgPlaceholderTypes
            atom_data.extend_from_slice(&0x8000_0000u32.to_le_bytes()); // masterIdRef
            atom_data.extend_from_slice(&0u32.to_le_bytes()); // notesIdRef
            atom_data.extend_from_slice(&slide_flags::DEFAULT.to_le_bytes());
            atom_data.extend_from_slice(&0u16.to_le_bytes()); // reserved
            slide_atom.write_data(&atom_data);
            slide_container.write_child(&slide_atom.build()?);

            // PPDrawing with Escher DgContainer (including user shapes)
            let escher_shapes: Vec<UserShapeData> = slide
                .shapes
                .iter()
                .map(|s| convert_shape_to_escher(s, &self.hyperlinks))
                .collect();
            let dg = create_dg_container_with_shapes(drawing_id, &escher_shapes)?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // ProgTags with PPT10 binary tag
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Ppt10Tag::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&BinaryTagData::SLIDE.to_bytes());
            prog_bin.write_child(&bin.build()?);
            prog_tags.write_child(&prog_bin.build()?);
            slide_container.write_child(&prog_tags.build()?);

            let slide_offset = ppt_stream.len() as u32;
            let persist_id = slide_persist_ids[i];
            persist_builder.set_offset(persist_id, slide_offset);

            let slide_bytes = slide_container.build()?;
            ppt_stream.extend_from_slice(&slide_bytes);
        }

        // 3.3) Notes containers - DISABLED for testing
        // Notes need more work - SlideListWithText instance=2, proper linking

        // PersistPtrHolder and UserEditAtom
        let persist_dir_offset = ppt_stream.len() as u32;
        let persist_dir_block = persist_builder.generate_record();
        ppt_stream.extend_from_slice(&persist_dir_block);

        let user_edit = UserEditAtom::new_minimal(
            persist_dir_offset,
            doc_persist_id,
            persist_builder.persist_id_seed(),
            self.slides.len() as u32,
        );
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        let current_user = build_current_user_stream(user_edit_offset);
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

        let mut ole_writer = OleWriter::new();
        ole_writer.set_root_clsid([
            0x10, 0x8D, 0x81, 0x64, 0x9B, 0x4F, 0xCF, 0x11, 0x86, 0xEA, 0x00, 0xAA, 0x00, 0xB9,
            0x29, 0xE8,
        ]);
        ole_writer.create_stream(&["PowerPoint Document"], &ppt_stream)?;
        ole_writer.create_stream(&["Current User"], &current_user)?;
        ole_writer.create_stream(&["\u{0005}SummaryInformation"], &summary_info)?;
        ole_writer.create_stream(&["\u{0005}DocumentSummaryInformation"], &doc_summary)?;

        // Pictures stream (per POI: separate stream for BLIP data)
        if !self.blip_store.is_empty() {
            let pictures_stream = self
                .blip_store
                .build_pictures_stream()
                .map_err(PptWriteError::Io)?;
            ole_writer.create_stream(&["Pictures"], &pictures_stream)?;
        }

        ole_writer.write_to(writer)?;

        Ok(())
    }

    // Helper methods for PPT writer:
    // The following are implemented via the modular components:
    // - Generating PPT record headers and containers
    // - Building Escher drawing records (DggContainer, DgContainer, etc.)
    // - Creating shape records (ClientData, ClientAnchor, etc.)
    // - Building text run records (TextCharsAtom, TextBytesAtom)
    // - Generating PersistPtr directory
    // - Creating CurrentUser stream
    // - Building SlideAtom and NotesAtom structures
    // - Managing master slides and layouts
    //
    // For production use, the PPTX writer is fully implemented and recommended.
}

impl Default for PptWriter {
    fn default() -> Self {
        Self::new()
    }
}

// Implementation deferred - PPT record generation functions:
// These would be needed for full PPT binary format support:
// - write_record_header() - Record header with version, type, instance, length
// - write_document_container() - DocumentContainer record
// - write_slide_container() - Slide record
// - write_drawing_container() - Drawing (Escher) container
// - write_shape_container() - Shape container (spContainer)
// - write_text_box() - Text box Escher record
// - write_client_data() - ClientData record linking to text
//
// Recommendation: Use the PPTX writer (fully implemented) for production use.
// - write_persist_directory() - PersistPtr directory

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_writer() {
        let writer = PptWriter::new();
        assert_eq!(writer.slides.len(), 0);
        assert_eq!(writer.slide_width, 9144000);
        assert_eq!(writer.slide_height, 6858000);
    }

    #[test]
    fn test_create_widescreen() {
        let writer = PptWriter::new_widescreen();
        assert_eq!(writer.slide_width, 9144000);
        assert_eq!(writer.slide_height, 5143500);
    }

    #[test]
    fn test_add_slide() {
        let mut writer = PptWriter::new();
        let idx = writer.add_slide().unwrap();
        assert_eq!(idx, 0);
        assert_eq!(writer.slides.len(), 1);
    }

    #[test]
    fn test_add_textbox() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer.add_textbox(slide, 10, 10, 100, 50, "Test").unwrap();
        assert_eq!(writer.slides[0].shapes.len(), 1);
    }

    #[test]
    fn test_delete_slide() {
        let mut writer = PptWriter::new();
        writer.add_slide().unwrap();
        writer.add_slide().unwrap();
        writer.delete_slide(0).unwrap();
        assert_eq!(writer.slides.len(), 1);
    }
}
