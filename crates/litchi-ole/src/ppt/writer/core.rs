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
//! use litchi_ole::ppt::PptWriter;
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
use super::comments::SlideComment;
use super::custom_shows::CustomShow;
use super::escher::{
    FreeformGeometry, UserShapeData, create_dg_container_with_tables, create_dgg_container,
    shape_type as escher_shape_type,
};
use super::hyperlink::{Hyperlink, HyperlinkCollection};
use super::master_drawing::build_master_ppdrawing;
use super::notes::{NotesContainerBuilder, NotesPage};
use super::persist::{PersistPtrBuilder, UserEditAtom};
use super::records::{
    RecordBuilder, create_docinfo_list_container, create_document_atom,
    create_end_document, create_environment_minimal, create_main_master_container,
    create_slide_list_with_text_master, record_type, wrap_dg_into_ppdrawing,
    wrap_dgg_into_ppdrawing_group,
};
use super::shape_style::{
    ArrowStyle, FillStyle, LineCapStyle, LineJoinStyle, LineStyle, LineStyleConfig, ShadowStyle,
    ShapeStyle,
};
#[allow(unused_imports)]
use super::shapes::ShapeKind;
use super::slide_timing::SlideTiming;
use super::spec::{BinaryTagData, ColorScheme, Ppt10Tag, SlideLayoutType, slide_flags};
use super::table::{PositionedTable, Table};
use super::text_format::{FontEntity, Paragraph, TextAlign};
use crate::ppt::animation::AnimationInfo;
use crate::ppt::encryption::{
    PptEncryptionProfile, WriterEncryptionMaterial, encrypt_pictures_for_write,
    encrypt_powerpoint_document_for_write, prepare_writer_encryption,
    validate_writer_password,
};
use crate::ppt::modify_password::{
    PowerPointModifyPassword, validate_value as validate_modify_password,
};
use crate::ppt::header_footer::{
    PowerPointHeaderFooter, PowerPointHeaderFooterParent,
    PowerPointHeaderFooterParentOrdinal, PowerPointHeaderFooterScope,
};
use crate::ppt::view_info::{PowerPointSlideViewInfo, PowerPointViewKind};
use litchi_cfb::writer::OleWriter;
use litchi_core::unit::pt_to_emu_i32;
use std::collections::HashMap;
use zeroize::Zeroizing;

/// Error type for PPT writing
#[derive(Debug)]
pub enum PptWriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(crate::OleError),
}

/// Build a minimal, valid Current User stream referencing the given UserEditAtom offset.
fn build_current_user_stream(offset_to_current_edit: u32, encrypted: bool) -> Vec<u8> {
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
    // headerToken
    let token: u32 = if encrypted { 0xF3D1_C4DF } else { 0xE391_C05F };
    s.extend_from_slice(&token.to_le_bytes());
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

impl From<crate::OleError> for PptWriteError {
    fn from(err: crate::OleError) -> Self {
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
    pub picture_index: Option<u32>,
    /// Explicit geometry for a custom/freeform shape.
    pub freeform_geometry: Option<FreeformGeometry>,
    /// Hyperlink attached to shape
    pub hyperlink_id: Option<u32>,
}

/// Represents a shape on a slide
#[derive(Debug, Clone)]
struct WritableShape {
    /// Shape properties
    properties: ShapeProperties,
    /// Animation info for this shape
    animation_info: Option<AnimationInfo>,
}

#[derive(Default)]
struct ConvertedLineProperties {
    color: Option<u32>,
    width: Option<i32>,
    opacity: Option<u32>,
    style: Option<u32>,
    dash_style: Option<u32>,
    start_arrow: Option<u32>,
    end_arrow: Option<u32>,
    start_arrow_width: Option<u32>,
    start_arrow_length: Option<u32>,
    end_arrow_width: Option<u32>,
    end_arrow_length: Option<u32>,
    join_style: Option<u32>,
    end_cap_style: Option<u32>,
}

fn convert_line_properties(line: Option<&LineStyleConfig>) -> ConvertedLineProperties {
    let Some(line) = line.filter(|line| line.enabled && line.width > 0) else {
        return ConvertedLineProperties::default();
    };
    let has_start_arrow = line.start_arrow != ArrowStyle::None;
    let has_end_arrow = line.end_arrow != ArrowStyle::None;
    ConvertedLineProperties {
        color: Some(line.color.to_rgbx()),
        width: Some(line.width as i32),
        opacity: (line.opacity < 100).then(|| (u32::from(line.opacity) * 65536) / 100),
        style: (line.style != LineStyle::Simple).then_some(line.style as u32),
        dash_style: (line.dash != super::shape_style::LineDashStyle::Solid)
            .then_some(line.dash as u32),
        start_arrow: has_start_arrow.then_some(line.start_arrow as u32),
        end_arrow: has_end_arrow.then_some(line.end_arrow as u32),
        start_arrow_width: has_start_arrow.then_some(line.start_arrow_width as u32),
        start_arrow_length: has_start_arrow.then_some(line.start_arrow_length as u32),
        end_arrow_width: has_end_arrow.then_some(line.end_arrow_width as u32),
        end_arrow_length: has_end_arrow.then_some(line.end_arrow_length as u32),
        join_style: (line.join != LineJoinStyle::Miter).then_some(line.join as u32),
        end_cap_style: (line.cap != LineCapStyle::Round).then_some(line.cap as u32),
    }
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
        }
    }
}

/// Represents a slide
#[derive(Debug, Clone)]
struct WritableSlide {
    /// Shapes on this slide
    shapes: Vec<WritableShape>,
    /// Tables on this slide
    tables: Vec<PositionedTable>,
    /// Slide notes text (simple)
    notes: Option<String>,
    /// Rich notes page
    notes_page: Option<NotesPage>,
    /// Slide comments
    comments: Vec<SlideComment>,
    /// Per-slide timing (auto-advance, hidden, etc.)
    timing: Option<SlideTiming>,
    /// Optional header/footer override attached directly to this slide.
    header_footer: Option<PowerPointHeaderFooter>,
}

struct SerializedHeaderFooters {
    presentation_slides: Option<Vec<u8>>,
    notes_and_handouts: Option<Vec<u8>>,
    main_master: Option<Vec<u8>>,
    slides: Vec<Option<Vec<u8>>>,
}

impl WritableSlide {
    /// Number of OfficeArt shapes in this slide's drawing, including the
    /// group patriarch, the background shape, and every table group/cell.
    fn escher_shape_count(&self) -> u32 {
        let table_shapes: u32 = self.tables.iter().map(|t| t.table.shape_count()).sum();
        2 + self.shapes.len() as u32 + table_shapes
    }
}

fn append_child_to_built_container(
    container: &mut Vec<u8>,
    child: &[u8],
) -> Result<(), PptWriteError> {
    if container.len() < 8 {
        return Err(PptWriteError::InvalidData(
            "PPT container is missing its record header".to_string(),
        ));
    }
    let stored_len = u32::from_le_bytes([
        container[4],
        container[5],
        container[6],
        container[7],
    ]) as usize;
    if stored_len != container.len() - 8 {
        return Err(PptWriteError::InvalidData(
            "PPT container length does not match its payload".to_string(),
        ));
    }
    let new_len = stored_len
        .checked_add(child.len())
        .and_then(|len| u32::try_from(len).ok())
        .ok_or_else(|| PptWriteError::InvalidData("PPT container is too large".to_string()))?;
    container.extend_from_slice(child);
    container[4..8].copy_from_slice(&new_len.to_le_bytes());
    Ok(())
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
        ShapeType::Freeform => escher_shape_type::NOT_PRIMITIVE,
    }
}

/// Convert WritableShape to UserShapeData for Escher serialization
fn convert_shape_to_escher(
    shape: &WritableShape,
    hyperlinks: &HyperlinkCollection,
) -> UserShapeData {
    let props = &shape.properties;

    // Extract fill properties from FillStyle
    let (fill_color, fill_type, fill_opacity, fill_back_color, fill_angle, fill_blip_index) = props
        .fill
        .as_ref()
        .map_or((None, None, None, None, None, None), |fill| {
            if !fill.enabled {
                return (None, None, None, None, None, None);
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

            (
                color,
                fill_type,
                opacity,
                back_color,
                angle,
                fill.picture_index,
            )
        });

    let line = convert_line_properties(props.line.as_ref());

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
    let paragraphs = props.paragraphs.clone().or_else(|| {
        if props.alignment == TextAlignment::Left {
            None
        } else {
            props
                .text
                .as_ref()
                .map(|text| vec![Paragraph::new(text.clone()).align(props.alignment.into())])
        }
    });
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
        fill_blip_index,
        line_color: line.color,
        line_width: line.width,
        line_opacity: line.opacity,
        line_style: line.style,
        line_dash_style: line.dash_style,
        line_start_arrow: line.start_arrow,
        line_end_arrow: line.end_arrow,
        line_start_arrow_width: line.start_arrow_width,
        line_start_arrow_length: line.start_arrow_length,
        line_end_arrow_width: line.end_arrow_width,
        line_end_arrow_length: line.end_arrow_length,
        line_join_style: line.join_style,
        line_end_cap_style: line.end_cap_style,
        text,
        paragraphs,
        text_type: 4,           // OTHER for regular shapes
        placeholder_type: None, // Not a placeholder for regular shapes
        has_shadow,
        flip_h: props.flip_h,
        flip_v: props.flip_v,
        rotation: shape_rotation_to_fixed(props.rotation),
        adjust_values: props.adjust_values.clone(),
        hyperlink_id: props.hyperlink_id,
        hyperlink_action: get_hyperlink_info(props.hyperlink_id, hyperlinks).0,
        hyperlink_jump: get_hyperlink_info(props.hyperlink_id, hyperlinks).1,
        hyperlink_type: get_hyperlink_info(props.hyperlink_id, hyperlinks).2,
        picture_index: props.picture_index,
        freeform_geometry: props.freeform_geometry.clone(),
        animation_info: shape.animation_info.clone(),
        shadow_color,
        shadow_offset_x,
        shadow_offset_y,
        shadow_opacity,
        shadow_type,
    }
}

fn shape_rotation_to_fixed(degrees: f32) -> Option<i32> {
    if !degrees.is_finite() || degrees.abs() <= 0.001 {
        return None;
    }
    Some(((f64::from(degrees) % 360.0) * 65536.0).round() as i32)
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
    // - Slide number: action=ACTION_HYPERLINK(4), jump=JUMP_NONE(0), hyperlinkType=LINK_SlideNumber(7)
    // - Next/Prev/First/Last: action=ACTION_JUMP(3), jump=varies, hyperlinkType=varies
    match &hyperlink.target {
        HyperlinkTarget::Url(_) | HyperlinkTarget::File(_) => (4, 0, 8), // ACTION_HYPERLINK, JUMP_NONE, LINK_Url
        HyperlinkTarget::Slide(_) => (4, 0, 7), // ACTION_HYPERLINK, JUMP_NONE, LINK_SlideNumber
        HyperlinkTarget::NextSlide => (3, 1, 0), // ACTION_JUMP, JUMP_NEXTSLIDE, LINK_NextSlide
        HyperlinkTarget::PrevSlide => (3, 2, 1), // ACTION_JUMP, JUMP_PREVIOUSSLIDE, LINK_PreviousSlide
        HyperlinkTarget::FirstSlide => (3, 3, 2), // ACTION_JUMP, JUMP_FIRSTSLIDE, LINK_FirstSlide
        HyperlinkTarget::LastSlide => (3, 4, 3), // ACTION_JUMP, JUMP_LASTSLIDE, LINK_LastSlide
        HyperlinkTarget::EndShow => (3, 6, 0xFF), // ACTION_JUMP, JUMP_ENDSHOW, LINK_NULL
        HyperlinkTarget::CustomShow(_) => (7, 0, 6), // ACTION_CUSTOMSHOW, JUMP_NONE, LINK_CustomShow
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
    /// Custom slide shows (named shows)
    custom_shows: Vec<CustomShow>,
    /// Optional typed override for the slide editing view.
    slide_view_info: Option<PowerPointSlideViewInfo>,
    /// Optional typed notes editing view.
    notes_view_info: Option<PowerPointSlideViewInfo>,
    /// Presentation-wide defaults for ordinary slides.
    presentation_header_footer: Option<PowerPointHeaderFooter>,
    /// Presentation-wide defaults for notes pages and handouts.
    notes_and_handouts_header_footer: Option<PowerPointHeaderFooter>,
    /// Header/footer defaults attached directly to the main master.
    main_master_header_footer: Option<PowerPointHeaderFooter>,
    /// Password-to-open settings, including a password wiped on replacement or drop.
    encryption: Option<PptWriterEncryption>,
    /// Inert modify password, wiped on replacement, clear, or drop.
    modify_password: Option<PptWriterModifyPassword>,
}

struct PptWriterEncryption {
    profile: PptEncryptionProfile,
    password: Zeroizing<String>,
}

struct PptWriterModifyPassword {
    password: Zeroizing<String>,
}

impl std::fmt::Debug for PptWriterModifyPassword {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PptWriterModifyPassword")
            .field("utf16_units", &self.password.encode_utf16().count())
            .field("password", &"[REDACTED]")
            .finish()
    }
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
            custom_shows: Vec::new(),
            slide_view_info: None,
            notes_view_info: None,
            presentation_header_footer: None,
            notes_and_handouts_header_footer: None,
            main_master_header_footer: None,
            encryption: None,
            modify_password: None,
        }
    }

    /// Protect the generated presentation with CryptoAPI password-to-open encryption.
    ///
    /// Validation is atomic: invalid input leaves any previous setting unchanged.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: PptEncryptionProfile,
    ) -> Result<(), PptWriteError> {
        let password = Zeroizing::new(password.into());
        validate_writer_password(profile, password.as_str())
            .map_err(PptWriteError::InvalidData)?;
        self.encryption = Some(PptWriterEncryption { profile, password });
        Ok(())
    }

    /// Remove password-to-open protection and wipe the stored password.
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured encryption profile without exposing the password.
    pub fn encryption_profile(&self) -> Option<PptEncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Set the inert password required by PowerPoint to modify the presentation.
    ///
    /// The secret is stored in zeroizing memory. Password-to-open encryption
    /// must also be configured before the presentation can be written.
    /// Validation is atomic and does not replace an existing valid value.
    pub fn set_modify_password(
        &mut self,
        password: impl Into<String>,
    ) -> Result<(), PptWriteError> {
        let password = Zeroizing::new(password.into());
        validate_modify_password(password.as_str())
            .map_err(|error| PptWriteError::InvalidData(error.to_string()))?;
        self.modify_password = Some(PptWriterModifyPassword { password });
        Ok(())
    }

    /// Remove the modify-password atom and wipe the stored secret.
    pub fn clear_modify_password(&mut self) {
        self.modify_password = None;
    }

    /// Return the configured value through the redacted typed password model.
    pub fn modify_password(&self) -> Option<PowerPointModifyPassword> {
        self.modify_password.as_ref().map(|value| {
            PowerPointModifyPassword::new(value.password.as_str())
                .expect("stored modify password was validated before assignment")
        })
    }

    fn validate_encryption(&self) -> Result<(), PptWriteError> {
        if self.modify_password.is_some() && self.encryption.is_none() {
            return Err(PptWriteError::InvalidData(
                "PowerPoint modify-password output requires password-to-open encryption"
                    .to_string(),
            ));
        }
        if let Some(value) = &self.encryption {
            validate_writer_password(value.profile, value.password.as_str())
                .map_err(PptWriteError::InvalidData)?;
        }
        if let Some(value) = &self.modify_password {
            validate_modify_password(value.password.as_str())
                .map_err(|error| PptWriteError::InvalidData(error.to_string()))?;
        }
        Ok(())
    }

    fn build_modify_password_programmable_tag(&self) -> Result<Option<Vec<u8>>, PptWriteError> {
        let Some(value) = &self.modify_password else {
            return Ok(None);
        };
        validate_modify_password(value.password.as_str())
            .map_err(|error| PptWriteError::InvalidData(error.to_string()))?;

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
        name.write_data(&Ppt10Tag::to_bytes());
        binary_tag.write_child(&name.build()?);
        binary_tag.write_child(&blob.build()?);
        let mut tags = RecordBuilder::new(0x0f, 0, record_type::PROG_TAGS);
        tags.write_child(&binary_tag.build()?);
        Ok(Some(tags.build()?))
    }

    fn prepare_encryption(&self) -> Result<Option<WriterEncryptionMaterial>, PptWriteError> {
        self.encryption
            .as_ref()
            .map(|value| prepare_writer_encryption(value.profile, value.password.as_str()))
            .transpose()
            .map_err(PptWriteError::InvalidData)
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
            tables: Vec::new(),
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
    pub fn delete_slide(&mut self, index: usize) -> Result<(), PptWriteError> {
        if index >= self.slides.len() {
            return Err(PptWriteError::InvalidData(format!(
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
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<(), PptWriteError> {
        if from_index >= self.slides.len() || to_index >= self.slides.len() {
            return Err(PptWriteError::InvalidData(
                "Invalid slide index".to_string(),
            ));
        }

        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        self.reindex_slide_header_footers();
        Ok(())
    }

    fn validated_header_footer(
        mut value: PowerPointHeaderFooter,
        scope: PowerPointHeaderFooterScope,
    ) -> Result<PowerPointHeaderFooter, PptWriteError> {
        value.scope = scope;
        value
            .to_record_bytes()
            .map_err(|error| PptWriteError::InvalidData(error.to_string()))?;
        Ok(value)
    }

    fn reindex_slide_header_footers(&mut self) {
        for (index, slide) in self.slides.iter_mut().enumerate() {
            if let Some(value) = &mut slide.header_footer {
                value.scope = PowerPointHeaderFooterScope::Local {
                    parent: PowerPointHeaderFooterParent::Slide,
                    parent_ordinal: PowerPointHeaderFooterParentOrdinal::new(index),
                };
            }
        }
    }

    fn serialize_header_footers(&self) -> Result<SerializedHeaderFooters, PptWriteError> {
        let serialize = |value: Option<&PowerPointHeaderFooter>, scope| {
            value
                .map(|value| {
                    let mut value = value.clone();
                    value.scope = scope;
                    value
                        .to_record_bytes()
                        .map_err(|error| PptWriteError::InvalidData(error.to_string()))
                })
                .transpose()
        };
        let presentation_slides = serialize(
            self.presentation_header_footer.as_ref(),
            PowerPointHeaderFooterScope::PresentationSlides,
        )?;
        let notes_and_handouts = serialize(
            self.notes_and_handouts_header_footer.as_ref(),
            PowerPointHeaderFooterScope::NotesAndHandouts,
        )?;
        let main_master = serialize(
            self.main_master_header_footer.as_ref(),
            PowerPointHeaderFooterScope::Local {
                parent: PowerPointHeaderFooterParent::MainMaster,
                parent_ordinal: PowerPointHeaderFooterParentOrdinal::new(0),
            },
        )?;
        let slides = self
            .slides
            .iter()
            .enumerate()
            .map(|(index, slide)| {
                serialize(
                    slide.header_footer.as_ref(),
                    PowerPointHeaderFooterScope::Local {
                        parent: PowerPointHeaderFooterParent::Slide,
                        parent_ordinal: PowerPointHeaderFooterParentOrdinal::new(index),
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
        value: PowerPointHeaderFooter,
    ) -> Result<(), PptWriteError> {
        let value = Self::validated_header_footer(
            value,
            PowerPointHeaderFooterScope::PresentationSlides,
        )?;
        self.presentation_header_footer = Some(value);
        Ok(())
    }

    /// Remove presentation-wide header/footer defaults for ordinary slides.
    pub fn clear_presentation_header_footer(&mut self) {
        self.presentation_header_footer = None;
    }

    /// Return presentation-wide header/footer defaults for ordinary slides.
    pub fn presentation_header_footer(&self) -> Option<&PowerPointHeaderFooter> {
        self.presentation_header_footer.as_ref()
    }

    /// Set presentation-wide header/footer defaults for notes pages and handouts.
    pub fn set_notes_and_handouts_header_footer(
        &mut self,
        value: PowerPointHeaderFooter,
    ) -> Result<(), PptWriteError> {
        let value = Self::validated_header_footer(
            value,
            PowerPointHeaderFooterScope::NotesAndHandouts,
        )?;
        self.notes_and_handouts_header_footer = Some(value);
        Ok(())
    }

    /// Remove presentation-wide header/footer defaults for notes pages and handouts.
    pub fn clear_notes_and_handouts_header_footer(&mut self) {
        self.notes_and_handouts_header_footer = None;
    }

    /// Return presentation-wide header/footer defaults for notes pages and handouts.
    pub fn notes_and_handouts_header_footer(&self) -> Option<&PowerPointHeaderFooter> {
        self.notes_and_handouts_header_footer.as_ref()
    }

    /// Set the header/footer defaults attached directly to the main master.
    pub fn set_main_master_header_footer(
        &mut self,
        value: PowerPointHeaderFooter,
    ) -> Result<(), PptWriteError> {
        let value = Self::validated_header_footer(
            value,
            PowerPointHeaderFooterScope::Local {
                parent: PowerPointHeaderFooterParent::MainMaster,
                parent_ordinal: PowerPointHeaderFooterParentOrdinal::new(0),
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
    pub fn main_master_header_footer(&self) -> Option<&PowerPointHeaderFooter> {
        self.main_master_header_footer.as_ref()
    }

    /// Set a header/footer override attached directly to one slide.
    pub fn set_slide_header_footer(
        &mut self,
        slide: usize,
        value: PowerPointHeaderFooter,
    ) -> Result<(), PptWriteError> {
        if slide >= self.slides.len() {
            return Err(PptWriteError::InvalidData(format!(
                "Slide {} does not exist",
                slide
            )));
        }
        let value = Self::validated_header_footer(
            value,
            PowerPointHeaderFooterScope::Local {
                parent: PowerPointHeaderFooterParent::Slide,
                parent_ordinal: PowerPointHeaderFooterParentOrdinal::new(slide),
            },
        )?;
        self.slides[slide].header_footer = Some(value);
        Ok(())
    }

    /// Remove a header/footer override attached directly to one slide.
    pub fn clear_slide_header_footer(&mut self, slide: usize) -> Result<(), PptWriteError> {
        let slide = self.slides.get_mut(slide).ok_or_else(|| {
            PptWriteError::InvalidData(format!("Slide {} does not exist", slide))
        })?;
        slide.header_footer = None;
        Ok(())
    }

    /// Return the header/footer override attached directly to one slide.
    pub fn slide_header_footer(
        &self,
        slide: usize,
    ) -> Result<Option<&PowerPointHeaderFooter>, PptWriteError> {
        self.slides
            .get(slide)
            .map(|slide| slide.header_footer.as_ref())
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))
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
    ) -> Result<(), PptWriteError> {
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
    ) -> Result<(), PptWriteError> {
        geometry.validate()?;
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

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
    ) -> Result<(), PptWriteError> {
        if shape_type == ShapeType::Freeform {
            return Err(PptWriteError::InvalidData(
                "freeform shapes require explicit geometry; use add_styled_freeform".to_string(),
            ));
        }
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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        slide_data.tables.push(PositionedTable {
            x: pt_to_emu_i32(x),
            y: pt_to_emu_i32(y),
            table,
        });
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
        if slide >= self.slides.len() {
            return Err(PptWriteError::InvalidData(format!(
                "Slide {} does not exist",
                slide
            )));
        }
        // Add picture to BLIP store
        let blip_index = self.add_picture_data(image_data);

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
            animation_info: None,
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
        if slide >= self.slides.len() {
            return Err(PptWriteError::InvalidData(format!(
                "Slide {} does not exist",
                slide
            )));
        }
        let blip_index = self.add_picture_data_with_type(image_data, blip_type);

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
            animation_info: None,
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Register picture data in the BLIP store and return its 1-based index.
    ///
    /// Use the returned index with [`FillStyle::picture`] to create a picture
    /// or texture fill without adding a picture-frame shape.
    pub fn add_picture_data(&mut self, image_data: Vec<u8>) -> u32 {
        self.blip_store.add_picture(image_data)
    }

    /// Register explicitly typed picture data and return its 1-based BLIP index.
    pub fn add_picture_data_with_type(&mut self, image_data: Vec<u8>, blip_type: BlipType) -> u32 {
        self.blip_store.add_picture_with_type(image_data, blip_type)
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

    /// Set the rotation of the last shape on a slide, in degrees.
    pub fn set_last_shape_rotation(
        &mut self,
        slide: usize,
        degrees: f32,
    ) -> Result<(), PptWriteError> {
        if !degrees.is_finite() {
            return Err(PptWriteError::InvalidData(
                "shape rotation must be finite".to_string(),
            ));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| PptWriteError::InvalidData("No shapes on slide".to_string()))?;
        shape.properties.rotation = degrees;
        Ok(())
    }

    /// Set one of the ten OfficeArt adjustment values on the last shape.
    pub fn set_last_shape_adjustment(
        &mut self,
        slide: usize,
        index: usize,
        value: i32,
    ) -> Result<(), PptWriteError> {
        if index >= 10 {
            return Err(PptWriteError::InvalidData(
                "shape adjustment index must be in the range 0..10".to_string(),
            ));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| PptWriteError::InvalidData("No shapes on slide".to_string()))?;
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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| PptWriteError::InvalidData("No shapes on slide".to_string()))?;
        if shape.properties.text.is_none() && shape.properties.paragraphs.is_none() {
            return Err(PptWriteError::InvalidData(
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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        let shape = slide_data.shapes.get_mut(shape_index).ok_or_else(|| {
            PptWriteError::InvalidData(format!(
                "Shape {} does not exist on slide {}",
                shape_index, slide
            ))
        })?;

        shape.animation_info = Some(animation);
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

    /// Add a comment to a slide.
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index (0-based)
    /// * `comment` - The comment to add
    pub fn add_comment(
        &mut self,
        slide: usize,
        comment: SlideComment,
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
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
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
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

    fn validate_view_info_kind(
        view: &PowerPointSlideViewInfo,
        expected: PowerPointViewKind,
    ) -> Result<(), PptWriteError> {
        view.to_bytes()
            .map_err(|error| PptWriteError::InvalidData(error.to_string()))?;
        if view.kind() != expected {
            return Err(PptWriteError::InvalidData(format!(
                "editing-view kind {:?} does not match {:?}",
                view.kind(),
                expected
            )));
        }
        Ok(())
    }

    /// Set the presentation's slide editing-view preferences, zoom, and guides.
    pub fn set_slide_view_info(
        &mut self,
        view: PowerPointSlideViewInfo,
    ) -> Result<(), PptWriteError> {
        Self::validate_view_info_kind(&view, PowerPointViewKind::Slide)?;
        self.slide_view_info = Some(view);
        Ok(())
    }

    /// Restore the writer's canonical default slide editing view.
    pub fn clear_slide_view_info(&mut self) {
        self.slide_view_info = None;
    }

    /// Return the explicit slide editing-view override, if present.
    pub fn slide_view_info(&self) -> Option<&PowerPointSlideViewInfo> {
        self.slide_view_info.as_ref()
    }

    /// Set the presentation's notes editing-view preferences, zoom, and guides.
    pub fn set_notes_view_info(
        &mut self,
        view: PowerPointSlideViewInfo,
    ) -> Result<(), PptWriteError> {
        Self::validate_view_info_kind(&view, PowerPointViewKind::Notes)?;
        self.notes_view_info = Some(view);
        Ok(())
    }

    /// Remove the optional notes editing-view record.
    pub fn clear_notes_view_info(&mut self) {
        self.notes_view_info = None;
    }

    /// Return the explicit notes editing-view, if present.
    pub fn notes_view_info(&self) -> Option<&PowerPointSlideViewInfo> {
        self.notes_view_info.as_ref()
    }

    fn build_docinfo_list(&self) -> Result<Vec<u8>, PptWriteError> {
        Ok(create_docinfo_list_container(
            self.slide_view_info.as_ref(),
            self.notes_view_info.as_ref(),
        )?)
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
        self.validate_encryption()?;
        let modify_password_tag = self.build_modify_password_programmable_tag()?;
        let header_footers = self.serialize_header_footers()?;
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
            .map(|s| s.escher_shape_count()) // 2 for group+background, plus user shapes and tables
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
        let docinfo = self.build_docinfo_list()?;
        doc_container.write_child(&docinfo);

        if let Some(value) = &header_footers.presentation_slides {
            doc_container.write_child(value);
        }
        if let Some(value) = &header_footers.notes_and_handouts {
            doc_container.write_child(value);
        }

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

        // 2.5.3) SoundCollection with embedded WAV data
        // Collect all sound IDs referenced by animations
        let mut sound_ids = std::collections::HashSet::new();
        for slide in &self.slides {
            for shape in &slide.shapes {
                if let Some(ref anim_info) = shape.animation_info
                    && let Some(ref build_list) = anim_info.build_list
                {
                    for build in &build_list.builds {
                        if let Some(ref sound) = build.sound {
                            sound_ids.insert(sound.sound_ref);
                        }
                    }
                }
            }
        }

        // Build SoundCollection with actual WAV data and get ID→ref mapping
        let (sound_collection, sound_id_mapping) = super::build_sound_collection(&sound_ids)?;
        if !sound_collection.is_empty() {
            doc_container.write_child(&sound_collection);
        }

        // Remap soundRef in animations to match CString instance 2 in SoundCollection
        for slide in &mut self.slides {
            for shape in &mut slide.shapes {
                if let Some(ref mut anim_info) = shape.animation_info
                    && let Some(ref mut build_list) = anim_info.build_list
                {
                    for build in &mut build_list.builds {
                        if let Some(ref mut sound) = build.sound
                            && let Some(&mapped_ref) = sound_id_mapping.get(&sound.sound_ref)
                        {
                            sound.sound_ref = mapped_ref;
                        }
                    }
                }
            }
        }

        // 2.5.4) NamedShows (custom slide shows) in Document container
        if !self.custom_shows.is_empty() {
            let named_shows = super::custom_shows::build_named_shows(&self.custom_shows)?;
            if !named_shows.is_empty() {
                doc_container.write_child(&named_shows);
            }
        }

        if let Some(value) = &modify_password_tag {
            doc_container.write_child(value);
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
        let mut master_container = create_main_master_container(&master_ppdrawing)?;
        if let Some(value) = &header_footers.main_master {
            append_child_to_built_container(&mut master_container, value)?;
        }
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
            let dg = create_dg_container_with_tables(drawing_id, &escher_shapes, &slide.tables)?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // SSSlideInfoAtom for per-slide timing (if set and no transition handles it)
            if let Some(ref timing) = slide.timing {
                let timing_record = super::slide_timing::build_slide_timing(timing)?;
                slide_container.write_child(&timing_record);
            }

            if let Some(value) = &header_footers.slides[i] {
                slide_container.write_child(value);
            }

            // ProgTags with PPT10 binary tag (PowerPoint 2002+ features)
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Ppt10Tag::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            // BinaryTagData: slide defaults + comments
            let comment_bytes = super::comments::build_slide_comments(&slide.comments)?;
            let mut tag_data = BinaryTagData::SLIDE.to_bytes().to_vec();
            tag_data.extend_from_slice(&comment_bytes);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&tag_data);
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

        let mut pictures_stream = if self.blip_store.is_empty() {
            None
        } else {
            Some(
                self.blip_store
                    .build_pictures_stream()
                    .map_err(PptWriteError::Io)?,
            )
        };
        let encryption = self.prepare_encryption()?;
        let encryption_session_id = if let Some(encryption) = &encryption {
            let persist_id = persist_builder.allocate_id();
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                PptWriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(persist_id, offset);
            ppt_stream.extend_from_slice(&encryption.session_record);
            Some(persist_id)
        } else {
            None
        };

        // 4) PersistPtrIncrementalBlock (6002) then single UserEditAtom
        let persist_dir_offset = ppt_stream.len() as u32;
        let persist_dir_block = persist_builder.generate_record();
        ppt_stream.extend_from_slice(&persist_dir_block);

        let mut user_edit = UserEditAtom::new_minimal(
            persist_dir_offset,
            doc_persist_id,
            persist_builder.persist_id_seed(),
            self.slides.len() as u32,
        );
        if let Some(session_id) = encryption_session_id {
            user_edit = user_edit.with_encryption_session(session_id);
        }
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        // 5) Build Current User and property streams
        let current_user = build_current_user_stream(user_edit_offset, encryption.is_some());
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

        if let (Some(encryption), Some(session_id)) = (&encryption, encryption_session_id) {
            encrypt_powerpoint_document_for_write(
                &mut ppt_stream,
                persist_dir_offset as usize,
                user_edit_offset as usize,
                session_id,
                &encryption.crypto,
            )
            .map_err(PptWriteError::InvalidData)?;
            if let Some(pictures) = &mut pictures_stream {
                encrypt_pictures_for_write(pictures, &encryption.crypto)
                    .map_err(PptWriteError::InvalidData)?;
            }
        }

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
        if let Some(pictures_stream) = &pictures_stream {
            ole_writer.create_stream(&["Pictures"], pictures_stream)?;
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
        self.validate_encryption()?;
        let modify_password_tag = self.build_modify_password_programmable_tag()?;
        let header_footers = self.serialize_header_footers()?;
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
            .map(|s| s.escher_shape_count())
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
        let docinfo = self.build_docinfo_list()?;
        doc_container.write_child(&docinfo);

        if let Some(value) = &header_footers.presentation_slides {
            doc_container.write_child(value);
        }
        if let Some(value) = &header_footers.notes_and_handouts {
            doc_container.write_child(value);
        }

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

        // SoundCollection with embedded WAV data
        let mut sound_ids = std::collections::HashSet::new();
        for slide in &self.slides {
            for shape in &slide.shapes {
                if let Some(ref anim_info) = shape.animation_info
                    && let Some(ref build_list) = anim_info.build_list
                {
                    for build in &build_list.builds {
                        if let Some(ref sound) = build.sound {
                            sound_ids.insert(sound.sound_ref);
                        }
                    }
                }
            }
        }

        // Build SoundCollection with actual WAV data and get ID→ref mapping
        let (sound_collection, sound_id_mapping) = super::build_sound_collection(&sound_ids)?;
        if !sound_collection.is_empty() {
            doc_container.write_child(&sound_collection);
        }

        // Remap soundRef in animations to match CString instance 2 in SoundCollection
        for slide in &mut self.slides {
            for shape in &mut slide.shapes {
                if let Some(ref mut anim_info) = shape.animation_info
                    && let Some(ref mut build_list) = anim_info.build_list
                {
                    for build in &mut build_list.builds {
                        if let Some(ref mut sound) = build.sound
                            && let Some(&mapped_ref) = sound_id_mapping.get(&sound.sound_ref)
                        {
                            sound.sound_ref = mapped_ref;
                        }
                    }
                }
            }
        }

        // NamedShows (custom slide shows) in Document container
        if !self.custom_shows.is_empty() {
            let named_shows = super::custom_shows::build_named_shows(&self.custom_shows)?;
            if !named_shows.is_empty() {
                doc_container.write_child(&named_shows);
            }
        }


        if let Some(value) = &modify_password_tag {
            doc_container.write_child(value);
        }

        let end_doc = create_end_document()?;
        doc_container.write_child(&end_doc);

        // Write finalized DocumentContainer
        let doc_bytes = doc_container.build()?;
        ppt_stream.extend_from_slice(&doc_bytes);

        // Then write MainMaster and slides as top-level records
        // MainMaster using dynamically built PPDrawing (includes all placeholders)
        let master_ppdrawing = build_master_ppdrawing();
        let mut master_container = create_main_master_container(&master_ppdrawing)?;
        if let Some(value) = &header_footers.main_master {
            append_child_to_built_container(&mut master_container, value)?;
        }
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
            let dg = create_dg_container_with_tables(drawing_id, &escher_shapes, &slide.tables)?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // SSSlideInfoAtom for per-slide timing (if set)
            if let Some(ref timing) = slide.timing {
                let timing_record = super::slide_timing::build_slide_timing(timing)?;
                slide_container.write_child(&timing_record);
            }

            if let Some(value) = &header_footers.slides[i] {
                slide_container.write_child(value);
            }

            // ProgTags with PPT10 binary tag
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Ppt10Tag::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            // BinaryTagData: slide defaults + comments
            let comment_bytes = super::comments::build_slide_comments(&slide.comments)?;
            let mut tag_data = BinaryTagData::SLIDE.to_bytes().to_vec();
            tag_data.extend_from_slice(&comment_bytes);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&tag_data);
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

        let mut pictures_stream = if self.blip_store.is_empty() {
            None
        } else {
            Some(
                self.blip_store
                    .build_pictures_stream()
                    .map_err(PptWriteError::Io)?,
            )
        };
        let encryption = self.prepare_encryption()?;
        let encryption_session_id = if let Some(encryption) = &encryption {
            let persist_id = persist_builder.allocate_id();
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                PptWriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(persist_id, offset);
            ppt_stream.extend_from_slice(&encryption.session_record);
            Some(persist_id)
        } else {
            None
        };

        // PersistPtrHolder and UserEditAtom
        let persist_dir_offset = ppt_stream.len() as u32;
        let persist_dir_block = persist_builder.generate_record();
        ppt_stream.extend_from_slice(&persist_dir_block);

        let mut user_edit = UserEditAtom::new_minimal(
            persist_dir_offset,
            doc_persist_id,
            persist_builder.persist_id_seed(),
            self.slides.len() as u32,
        );
        if let Some(session_id) = encryption_session_id {
            user_edit = user_edit.with_encryption_session(session_id);
        }
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        let current_user = build_current_user_stream(user_edit_offset, encryption.is_some());
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

        if let (Some(encryption), Some(session_id)) = (&encryption, encryption_session_id) {
            encrypt_powerpoint_document_for_write(
                &mut ppt_stream,
                persist_dir_offset as usize,
                user_edit_offset as usize,
                session_id,
                &encryption.crypto,
            )
            .map_err(PptWriteError::InvalidData)?;
            if let Some(pictures) = &mut pictures_stream {
                encrypt_pictures_for_write(pictures, &encryption.crypto)
                    .map_err(PptWriteError::InvalidData)?;
            }
        }

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
        if let Some(pictures_stream) = &pictures_stream {
            ole_writer.create_stream(&["Pictures"], pictures_stream)?;
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
    use crate::ppt::shapes::geometry::{GeometryRect, ShapePathType};
    use crate::ppt::writer::shape_style::{ArrowSize, ShapeColor};
    use std::io::Cursor;

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
    fn test_add_multiple_slides() {
        let mut writer = PptWriter::new();
        let idx1 = writer.add_slide().unwrap();
        let idx2 = writer.add_slide().unwrap();
        let idx3 = writer.add_slide().unwrap();
        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(idx3, 2);
        assert_eq!(writer.slide_count(), 3);
    }

    #[test]
    fn test_add_and_write_freeform_shape() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        let geometry = FreeformGeometry::new(
            GeometryRect::new(0, 0, 21600, 21600),
            ShapePathType::Complex,
            vec![(0, 0), (10800, 21600), (21600, 0)],
            vec![0x4000, 0x0001, 0x0001, 0x8000],
        );

        writer
            .add_freeform(slide, 10, 20, 300, 200, geometry)
            .unwrap();
        assert_eq!(writer.slides[slide].shapes.len(), 1);
        assert_eq!(
            writer.slides[slide].shapes[0].properties.shape_type,
            ShapeType::Freeform
        );

        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        assert!(!output.into_inner().is_empty());
    }

    #[test]
    fn test_rejects_empty_freeform_geometry() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        let geometry = FreeformGeometry::new(
            GeometryRect::new(0, 0, 21600, 21600),
            ShapePathType::Complex,
            Vec::new(),
            vec![0x8000],
        );

        assert!(
            writer
                .add_freeform(slide, 0, 0, 100, 100, geometry)
                .is_err()
        );
        assert!(writer.slides[slide].shapes.is_empty());
    }

    #[test]
    fn test_generic_styled_shape_rejects_geometryless_freeform() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();

        let result = writer.add_styled_shape(
            slide,
            ShapeType::Freeform,
            0,
            0,
            100,
            100,
            ShapeStyle::default(),
        );

        assert!(result.is_err());
        assert!(writer.slides[slide].shapes.is_empty());
    }

    #[test]
    fn test_add_textbox() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer.add_textbox(slide, 10, 10, 100, 50, "Test").unwrap();
        assert_eq!(writer.slides[0].shapes.len(), 1);
    }

    #[test]
    fn test_plain_text_alignment_and_rotation_reach_escher_shape() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_textbox(slide, 10, 20, 300, 100, "Centered")
            .unwrap();
        writer
            .set_last_shape_text_alignment(slide, TextAlignment::Center)
            .unwrap();
        writer.set_last_shape_rotation(slide, 450.5).unwrap();

        let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);
        assert!(shape.text.is_none());
        let paragraphs = shape.paragraphs.as_ref().expect("formatted paragraph");
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].alignment, TextAlign::Center);
        assert_eq!(shape.rotation, Some((90 * 65536) + 32768));
    }

    #[test]
    fn test_alignment_setter_updates_rich_paragraphs() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_rich_textbox(
                slide,
                0,
                0,
                200,
                100,
                vec![Paragraph::new("One"), Paragraph::new("Two")],
            )
            .unwrap();
        writer
            .set_last_shape_text_alignment(slide, TextAlignment::Justify)
            .unwrap();

        let paragraphs = writer.slides[slide].shapes[0]
            .properties
            .paragraphs
            .as_ref()
            .unwrap();
        assert!(
            paragraphs
                .iter()
                .all(|paragraph| paragraph.alignment == TextAlign::Justify)
        );
    }

    #[test]
    fn test_rotation_setter_rejects_non_finite_values() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer.add_rectangle(slide, 0, 0, 100, 100).unwrap();

        assert!(writer.set_last_shape_rotation(slide, f32::NAN).is_err());
        assert_eq!(writer.slides[slide].shapes[0].properties.rotation, 0.0);
        assert!(
            writer
                .set_last_shape_text_alignment(slide, TextAlignment::Center)
                .is_err()
        );
    }

    #[test]
    fn test_shape_adjustment_setter_preserves_sparse_positions() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_styled_shape(
                slide,
                ShapeType::Arrow,
                0,
                0,
                200,
                100,
                ShapeStyle::default(),
            )
            .unwrap();

        writer.set_last_shape_adjustment(slide, 3, -42).unwrap();
        assert_eq!(
            writer.slides[slide].shapes[0].properties.adjust_values,
            [0, 0, 0, -42]
        );
        let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);
        assert_eq!(shape.adjust_values, [0, 0, 0, -42]);

        assert!(writer.set_last_shape_adjustment(slide, 10, 7).is_err());
        assert_eq!(
            writer.slides[slide].shapes[0].properties.adjust_values,
            [0, 0, 0, -42]
        );
    }

    #[test]
    fn test_add_textbox_invalid_slide() {
        let mut writer = PptWriter::new();
        let result = writer.add_textbox(0, 10, 10, 100, 50, "Test");
        assert!(result.is_err());
    }

    #[test]
    fn test_delete_slide() {
        let mut writer = PptWriter::new();
        writer.add_slide().unwrap();
        writer.add_slide().unwrap();
        writer.delete_slide(0).unwrap();
        assert_eq!(writer.slides.len(), 1);
    }

    #[test]
    fn test_delete_invalid_slide() {
        let mut writer = PptWriter::new();
        let result = writer.delete_slide(0);
        assert!(result.is_err());
    }

    #[test]
    fn test_add_styled_shape() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        let style = ShapeStyle::solid_no_line(ShapeColor::RED);
        writer
            .add_styled_shape(slide, ShapeType::Rectangle, 10, 10, 100, 50, style)
            .unwrap();
        assert_eq!(writer.slides[0].shapes.len(), 1);
    }

    #[test]
    fn test_add_rectangle() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer.add_rectangle(slide, 10, 10, 100, 50).unwrap();
        assert_eq!(writer.slides[0].shapes.len(), 1);
    }

    #[test]
    fn test_add_arrow_line() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer.add_arrow_line(slide, 0, 0, 100, 100).unwrap();
        assert_eq!(writer.slides[0].shapes.len(), 1);
    }

    #[test]
    fn test_set_slide_notes() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .set_slide_notes(slide, "These are speaker notes")
            .unwrap();
        assert_eq!(
            writer.slides[0].notes,
            Some("These are speaker notes".to_string())
        );
    }

    #[test]
    fn test_add_font() {
        let mut writer = PptWriter::new();
        // PptWriter::new() already adds Arial as default font at index 0
        let font = FontEntity::times_new_roman();
        let idx = writer.add_font(font);
        assert_eq!(idx, 1); // Second font at index 1
        assert_eq!(writer.font_count(), 2);
    }

    #[test]
    fn test_add_multiple_fonts() {
        let mut writer = PptWriter::new();
        // PptWriter::new() already adds Arial as default font at index 0
        let idx1 = writer.add_font(FontEntity::arial()); // Returns 1
        let idx2 = writer.add_font(FontEntity::times_new_roman()); // Returns 2
        let idx3 = writer.add_font(FontEntity::new("Calibri")); // Returns 3
        assert_eq!(idx1, 1);
        assert_eq!(idx2, 2);
        assert_eq!(idx3, 3);
        assert_eq!(writer.font_count(), 4); // Arial (default) + 3 added
    }

    #[test]
    fn test_set_property() {
        let mut writer = PptWriter::new();
        writer.set_property("Title", "My Presentation");
        writer.set_property("Author", "Test Author");
        assert_eq!(
            writer.properties.get("Title"),
            Some(&"My Presentation".to_string())
        );
        assert_eq!(
            writer.properties.get("Author"),
            Some(&"Test Author".to_string())
        );
    }

    #[test]
    fn test_hyperlink_collection() {
        let mut writer = PptWriter::new();
        let link = Hyperlink::url("https://example.com").with_display_text("Example");
        let id = writer.add_hyperlink(link);
        assert_eq!(id, 1);
        assert_eq!(writer.hyperlink_count(), 1);
        assert!(writer.hyperlinks.get(1).is_some());
    }

    #[test]
    fn maps_writer_hyperlinks_to_spec_link_targets() {
        let mut links = HyperlinkCollection::new();
        let slide = links.add(Hyperlink::slide(2));
        assert_eq!(get_hyperlink_info(Some(slide), &links), (4, 0, 7));

        let next = links.add(Hyperlink::next_slide());
        assert_eq!(get_hyperlink_info(Some(next), &links), (3, 1, 0));

        let custom = links.add(Hyperlink {
            id: 0,
            display_text: None,
            target: crate::ppt::writer::hyperlink::HyperlinkTarget::CustomShow("Demo".to_string()),
            target_frame: None,
        });
        assert_eq!(get_hyperlink_info(Some(custom), &links), (7, 0, 6));
    }

    #[test]
    fn test_add_comment() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        let comment = SlideComment::new("John Doe", "Great slide!", 100, 50);
        writer.add_comment(slide, comment).unwrap();
        assert_eq!(writer.slides[0].comments.len(), 1);
        assert_eq!(writer.slides[0].comments[0].author, "John Doe");
    }

    #[test]
    fn test_add_multiple_comments() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_comment(slide, SlideComment::new("Alice", "First", 10, 10))
            .unwrap();
        writer
            .add_comment(slide, SlideComment::new("Bob", "Second", 20, 20))
            .unwrap();
        assert_eq!(writer.slides[0].comments.len(), 2);
    }

    #[test]
    fn test_add_comment_invalid_slide() {
        let mut writer = PptWriter::new();
        let comment = SlideComment::new("John", "Test", 0, 0);
        let result = writer.add_comment(0, comment);
        assert!(result.is_err());
    }

    #[test]
    fn test_shape_properties() {
        let props = ShapeProperties {
            shape_type: ShapeType::Rectangle,
            x: 100,
            y: 200,
            width: 300,
            height: 400,
            text: Some("Hello".to_string()),
            paragraphs: None,
            alignment: TextAlignment::Center,
            fill: None,
            line: None,
            shadow: None,
            rotation: 45.0,
            adjust_values: Vec::new(),
            flip_h: true,
            flip_v: false,
            hyperlink_id: None,
            picture_index: None,
            freeform_geometry: None,
        };
        assert_eq!(props.x, 100);
        assert_eq!(props.y, 200);
        assert_eq!(props.width, 300);
        assert_eq!(props.height, 400);
        assert!(props.flip_h);
        assert!(!props.flip_v);
    }

    #[test]
    fn test_slide_count() {
        let mut writer = PptWriter::new();
        assert_eq!(writer.slide_count(), 0);
        writer.add_slide().unwrap();
        assert_eq!(writer.slide_count(), 1);
        writer.add_slide().unwrap();
        assert_eq!(writer.slide_count(), 2);
    }

    #[test]
    fn test_default_writer() {
        let writer: PptWriter = Default::default();
        assert_eq!(writer.slide_count(), 0);
        assert_eq!(writer.slide_width, 9144000);
        assert_eq!(writer.slide_height, 6858000);
    }

    #[test]
    fn test_ppt_write_error_display() {
        let io_err = PptWriteError::Io(std::io::Error::other("test error"));
        let err_str = format!("{}", io_err);
        assert!(err_str.contains("I/O error"));

        let data_err = PptWriteError::InvalidData("bad data".to_string());
        let err_str = format!("{}", data_err);
        assert!(err_str.contains("Invalid data"));
    }

    #[test]
    fn test_text_alignment_conversions() {
        assert_eq!(TextAlignment::Left as u8, 0);
        assert_eq!(TextAlignment::Center as u8, 1);
        assert_eq!(TextAlignment::Right as u8, 2);
        assert_eq!(TextAlignment::Justify as u8, 3);
    }

    #[test]
    fn test_slide_layout_types() {
        use super::super::spec::SlideLayoutType;
        assert_eq!(SlideLayoutType::TitleSlide as u32, 0);
        assert_eq!(SlideLayoutType::TitleBody as u32, 1);
        assert_eq!(SlideLayoutType::MasterTitle as u32, 2);
        assert_eq!(SlideLayoutType::TitleOnly as u32, 7);
        assert_eq!(SlideLayoutType::Blank as u32, 13);
    }

    #[test]
    fn test_shape_type_variants() {
        let types = vec![
            ShapeType::Rectangle,
            ShapeType::TextBox,
            ShapeType::Placeholder,
            ShapeType::Line,
            ShapeType::Ellipse,
            ShapeType::RoundRectangle,
            ShapeType::Diamond,
            ShapeType::Triangle,
            ShapeType::Arrow,
            ShapeType::Star,
            ShapeType::Heart,
            ShapeType::Picture,
        ];
        for shape_type in types {
            let _ = format!("{:?}", shape_type);
        }
    }

    #[test]
    fn test_write_to_memory() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_textbox(slide, 100, 100, 400, 200, "Hello, World!")
            .unwrap();

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
        assert!(!buffer.get_ref().is_empty());
    }

    #[test]
    fn test_write_empty_presentation() {
        let mut writer = PptWriter::new();
        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
        assert!(!buffer.get_ref().is_empty());
    }

    #[test]
    fn test_write_multiple_slides() {
        let mut writer = PptWriter::new();

        let slide1 = writer.add_slide().unwrap();
        writer
            .add_textbox(slide1, 100, 100, 400, 100, "Slide 1")
            .unwrap();

        let slide2 = writer.add_slide().unwrap();
        writer.add_ellipse(slide2, 100, 100, 200, 150).unwrap();

        let slide3 = writer.add_slide().unwrap();
        writer.add_line(slide3, 0, 0, 500, 500).unwrap();

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
        assert!(!buffer.get_ref().is_empty());
    }

    #[test]
    fn test_presentation_with_hyperlink() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();

        writer
            .add_textbox(slide, 100, 100, 300, 100, "Click here")
            .unwrap();

        let link = Hyperlink::url("https://example.com");
        let link_id = writer.add_hyperlink(link);
        writer.set_last_shape_hyperlink(slide, link_id).unwrap();

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
    }

    #[test]
    fn test_presentation_with_comments() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_textbox(slide, 100, 100, 400, 200, "Content")
            .unwrap();

        let comment = SlideComment::new("Reviewer", "Please update this", 150, 150);
        writer.add_comment(slide, comment).unwrap();

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
    }

    #[test]
    fn test_presentation_with_notes() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_textbox(slide, 100, 100, 400, 200, "Title")
            .unwrap();
        writer
            .set_slide_notes(slide, "These are speaker notes for this slide")
            .unwrap();

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
    }

    #[test]
    fn test_presentation_with_multiple_shapes() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();

        writer.add_rectangle(slide, 50, 50, 100, 100).unwrap();
        writer.add_line(slide, 50, 200, 300, 200).unwrap();
        writer
            .add_textbox(slide, 50, 300, 300, 100, "Text box content")
            .unwrap();

        assert_eq!(writer.slides[0].shapes.len(), 3);

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
    }

    #[test]
    fn test_custom_show_support() {
        let mut writer = PptWriter::new();
        writer.add_slide().unwrap();
        writer.add_slide().unwrap();
        writer.add_slide().unwrap();

        let custom_show = CustomShow::new("Important Slides", &[0, 2]);
        writer.add_custom_show(custom_show);

        assert_eq!(writer.custom_show_count(), 1);
    }

    #[test]
    fn test_multiple_custom_shows() {
        let mut writer = PptWriter::new();
        for _ in 0..5 {
            writer.add_slide().unwrap();
        }

        writer.add_custom_show(CustomShow::new("First Show", &[0usize, 1]));
        writer.add_custom_show(CustomShow::new("Second Show", &[2usize, 3, 4]));

        assert_eq!(writer.custom_show_count(), 2);
    }

    #[test]
    fn test_widescreen_write() {
        let mut writer = PptWriter::new_widescreen();
        let slide = writer.add_slide().unwrap();
        writer
            .add_textbox(slide, 100, 100, 800, 100, "Widescreen")
            .unwrap();

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
    }

    #[test]
    fn test_shape_with_styling() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();

        let style = ShapeStyle::new()
            .with_fill(FillStyle::solid_rgb(255, 0, 0))
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                2.0,
            ));

        writer
            .add_styled_shape(slide, ShapeType::Rectangle, 100, 100, 200, 150, style)
            .unwrap();

        let mut buffer = Cursor::new(Vec::new());
        let result = writer.write_to(&mut buffer);
        assert!(result.is_ok());
    }

    #[test]
    fn test_extended_line_style_reaches_escher_shape() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        let mut line = LineStyleConfig::with_color_and_width(ShapeColor::RED, 2.0);
        line.opacity = 50;
        line.style = LineStyle::Triple;
        line.cap = LineCapStyle::Flat;
        line.join = LineJoinStyle::Round;
        line.start_arrow = ArrowStyle::Triangle;
        line.start_arrow_width = ArrowSize::Small;
        line.start_arrow_length = ArrowSize::Large;
        line.end_arrow = ArrowStyle::Open;
        line.end_arrow_width = ArrowSize::Large;
        line.end_arrow_length = ArrowSize::Small;
        let style = ShapeStyle::new().with_line(line);

        writer
            .add_styled_shape(slide, ShapeType::Rectangle, 10, 10, 100, 100, style)
            .unwrap();
        let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);

        assert_eq!(shape.line_opacity, Some(32768));
        assert_eq!(shape.line_style, Some(LineStyle::Triple as u32));
        assert_eq!(shape.line_end_cap_style, Some(LineCapStyle::Flat as u32));
        assert_eq!(shape.line_join_style, Some(LineJoinStyle::Round as u32));
        assert_eq!(shape.line_start_arrow_width, Some(ArrowSize::Small as u32));
        assert_eq!(shape.line_start_arrow_length, Some(ArrowSize::Large as u32));
        assert_eq!(shape.line_end_arrow_width, Some(ArrowSize::Large as u32));
        assert_eq!(shape.line_end_arrow_length, Some(ArrowSize::Small as u32));

        let default_line = convert_line_properties(Some(&LineStyleConfig::default_line()));
        assert_eq!(default_line.opacity, None);
        assert_eq!(default_line.style, None);
        assert_eq!(default_line.end_cap_style, None);
        assert_eq!(default_line.join_style, None);
    }

    #[test]
    fn test_picture_fill_registers_and_serializes_blip_reference() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        let blip_index =
            writer.add_picture_data_with_type(vec![0x89, b'P', b'N', b'G'], BlipType::Png);
        let style = ShapeStyle::new().with_fill(FillStyle::picture(blip_index));
        writer
            .add_styled_shape(slide, ShapeType::Rectangle, 10, 10, 100, 100, style)
            .unwrap();

        let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);
        assert_eq!(shape.fill_type, Some(3));
        assert_eq!(shape.fill_blip_index, Some(1));
        assert_eq!(writer.picture_count(), 1);

        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        assert!(!output.into_inner().is_empty());
    }

    #[test]
    fn test_invalid_slide_picture_does_not_mutate_blip_store() {
        let mut writer = PptWriter::new();

        assert!(
            writer
                .add_picture(7, 0, 0, 100, 100, vec![0x89, b'P', b'N', b'G'])
                .is_err()
        );
        assert_eq!(writer.picture_count(), 0);
    }

    #[test]
    fn test_invalid_operations() {
        let mut writer = PptWriter::new();

        // Try to add shape to non-existent slide
        let result = writer.add_rectangle(0, 0, 0, 100, 100);
        assert!(result.is_err());

        // Try to add textbox to non-existent slide
        let result = writer.add_textbox(5, 10, 10, 100, 50, "Test");
        assert!(result.is_err());

        // Try to set notes on non-existent slide
        let result = writer.set_slide_notes(0, "Notes");
        assert!(result.is_err());
    }

    #[test]
    fn test_internal_slide_data() {
        let mut writer = PptWriter::new();
        let slide_idx = writer.add_slide().unwrap();

        // Verify slide was created with correct defaults
        let slide = &writer.slides[slide_idx];
        assert!(slide.shapes.is_empty());
        assert!(slide.notes.is_none());
        assert!(slide.comments.is_empty());
    }

    #[test]
    fn test_slide_persist_tracking() {
        let mut writer = PptWriter::new();

        let idx1 = writer.add_slide().unwrap();
        let idx2 = writer.add_slide().unwrap();

        // Each slide gets a persist ID assigned during writing
        // (We can't check this directly, but we verify the structure)
        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
    }
}
