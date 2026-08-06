//! PPT-specific semantic OfficeArt properties and typed shape inputs.
//!
//! This layer owns the PPT vocabulary and validated data that the wire
//! encoder consumes. Shared OfficeArt records remain owned by the
//! litchi-odraw substrate exposed through the facade.

use std::io::Write;
use zerocopy::IntoBytes as _;
use zerocopy_derive::{Immutable, IntoBytes, KnownLayout};

use crate::shapes::geometry::{GeometryRect, ShapePathType};

use litchi_odraw::write::{Header, Property, shape_type};

use super::Error;

// =============================================================================
// Property IDs (MS-ODRAW 2.3.1)
// =============================================================================

/// Escher property IDs
pub(crate) mod prop_id {
    // Transform group
    pub(crate) const ROTATION: u16 = 0x0004;

    // Geometry group
    pub(crate) const GEOM_LEFT: u16 = 0x0140;
    pub(crate) const GEOM_TOP: u16 = 0x0141;
    pub(crate) const GEOM_RIGHT: u16 = 0x0142;
    pub(crate) const GEOM_BOTTOM: u16 = 0x0143;
    pub(crate) const SHAPE_PATH: u16 = 0x0144;
    pub(crate) const VERTICES: u16 = 0x0145;
    pub(crate) const SEGMENT_INFO: u16 = 0x0146;
    pub(crate) const ADJUST_VALUE: u16 = 0x0147;

    // Fill style (MS-ODRAW section 2.3.7)
    pub(crate) const FILL_TYPE: u16 = 0x0180;
    pub(crate) const FILL_COLOR: u16 = 0x0181;
    pub(crate) const FILL_OPACITY: u16 = 0x0182;
    pub(crate) const FILL_BACK_COLOR: u16 = 0x0183;
    pub(crate) const FILL_BLIP: u16 = 0x4186;
    pub(crate) const FILL_ANGLE: u16 = 0x018B; // fillAngle for gradients (degrees * 65536)
    pub(crate) const FILL_RECT_RIGHT: u16 = 0x0193; // fillRectRight per MS-ODRAW
    pub(crate) const FILL_RECT_BOTTOM: u16 = 0x0194; // fillRectBottom per MS-ODRAW
    pub(crate) const NO_FILL_HIT_TEST: u16 = 0x01BF;

    // Line style
    pub(crate) const LINE_COLOR: u16 = 0x01C0;
    pub(crate) const LINE_OPACITY: u16 = 0x01C1;
    pub(crate) const LINE_WIDTH: u16 = 0x01CB;
    pub(crate) const LINE_STYLE: u16 = 0x01CD;
    pub(crate) const LINE_DASH_STYLE: u16 = 0x01CE;
    pub(crate) const LINE_START_ARROW: u16 = 0x01D0;
    pub(crate) const LINE_END_ARROW: u16 = 0x01D1;
    pub(crate) const LINE_START_ARROW_WIDTH: u16 = 0x01D2;
    pub(crate) const LINE_START_ARROW_LENGTH: u16 = 0x01D3;
    pub(crate) const LINE_END_ARROW_WIDTH: u16 = 0x01D4;
    pub(crate) const LINE_END_ARROW_LENGTH: u16 = 0x01D5;
    pub(crate) const LINE_JOIN_STYLE: u16 = 0x01D6;
    pub(crate) const LINE_END_CAP_STYLE: u16 = 0x01D7;
    pub(crate) const LINE_BLIP: u16 = 0x41C5;
    pub(crate) const LINE_STYLE_BOOL: u16 = 0x01FF;

    // Shadow style
    pub(crate) const SHADOW_TYPE: u16 = 0x0200;
    pub(crate) const SHADOW_COLOR: u16 = 0x0201;
    pub(crate) const SHADOW_OPACITY: u16 = 0x0204;
    pub(crate) const SHADOW_OFFSET_X: u16 = 0x0205;
    pub(crate) const SHADOW_OFFSET_Y: u16 = 0x0206;
    pub(crate) const SHADOW_BOOL: u16 = 0x023F; // shadowObscured

    // Shape
    pub(crate) const BW_MODE: u16 = 0x0304;
    pub(crate) const BACKGROUND_SHAPE: u16 = 0x033F;
}

// =============================================================================
// PPT-Specific Property Values
// =============================================================================

/// PPT-specific property values (extends shared prop_value)
pub(crate) mod ppt_prop_value {
    pub(crate) use litchi_odraw::write::prop_value::*;

    /// Background fill color
    pub(crate) const BG_FILL_COLOR: u32 = 134_217_728; // 0x0800_0000
    pub(crate) const BG_FILL_BACK_COLOR: u32 = 134_217_733; // 0x0800_0005

    /// Slide dimensions (EMUs)
    pub(crate) const SLIDE_WIDTH_EMU: u32 = 10_064_750; // 914400 * 11
    pub(crate) const SLIDE_HEIGHT_EMU: u32 = 7_778_750; // 914400 * 8.5

    /// No fill hit test value
    pub(crate) const NO_FILL_HIT_TEST: u32 = 1_179_666; // 0x0012_0012
    pub(crate) const NO_LINE_DRAW_DASH: u32 = 524_288; // 0x0008_0000

    /// Black and white mode
    pub(crate) const BW_MODE_AUTO: u32 = 9;

    /// Background shape flag
    pub(crate) const BACKGROUND_SHAPE: u32 = 65_537; // 0x0001_0001

    /// Reserved cluster cspidCur
    pub(crate) const RESERVED_CSPID_CUR: u32 = 4;

    /// POI master shape count
    pub(crate) const POI_MASTER_SHAPE_COUNT: u32 = 6;
    pub(crate) const POI_SPID_MAX: u32 = 3076;
}

// =============================================================================
// Split Menu Colors
// =============================================================================

/// Split menu colors structure
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub(crate) struct SplitMenuColors {
    pub(crate) fill_color: u32,
    pub(crate) line_color: u32,
    pub(crate) shadow_color: u32,
    pub(crate) color_3d: u32,
}

impl SplitMenuColors {
    pub(crate) const DEFAULT: Self = Self {
        fill_color: crate::officeart_wire::prop_value::SCHEME_FILL,
        line_color: crate::officeart_wire::prop_value::SCHEME_LINE,
        shadow_color: crate::officeart_wire::prop_value::SCHEME_SHADOW,
        color_3d: 0x1000_00F7,
    };
}

// =============================================================================
// Header Version Constants
// =============================================================================

/// Escher record header versions
pub(crate) mod header_version {
    pub(crate) const CONTAINER: u8 = 0x0F;
    pub(crate) const SIMPLE: u8 = 0x00;
    pub(crate) const SPGR: u8 = 0x01;
    pub(crate) const SP: u8 = 0x02;
    pub(crate) const OPT: u8 = 0x03;
    pub(crate) const DG: u8 = 0x00; // instance is drawing_id
}

/// Escher record header (8 bytes) - builder-friendly version
#[derive(Debug, Clone)]
pub(crate) struct EscherHeader {
    /// Version (4 bits)
    pub(crate) version: u8,
    /// Instance (12 bits)
    pub(crate) instance: u16,
    /// Record type
    pub(crate) record_type: u16,
    /// Length
    pub(crate) length: u32,
}

impl EscherHeader {
    /// Create a new Escher header
    pub(crate) fn new(version: u8, instance: u16, record_type: u16, length: u32) -> Self {
        Self {
            version: version & 0x0F,
            instance: instance & 0x0FFF,
            record_type,
            length,
        }
    }

    /// Write header to writer
    pub(crate) fn write<W: Write>(&self, writer: &mut W) -> Result<(), Error> {
        let raw = Header::new(self.version, self.instance, self.record_type, self.length);
        writer.write_all(raw.as_bytes())?;
        Ok(())
    }
}

// =============================================================================
// Drawing Group (EscherDgg) - MS-ODRAW 2.2.12
// =============================================================================

/// File ID cluster entry
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub(crate) struct FileIdCluster {
    /// Drawing group ID
    pub(crate) dgid: u32,
    /// Next available shape ID in cluster
    pub(crate) cspid_cur: u32,
}

impl FileIdCluster {
    pub(crate) const fn new(dgid: u32, cspid_cur: u32) -> Self {
        Self { dgid, cspid_cur }
    }

    pub(crate) const fn reserved() -> Self {
        Self {
            dgid: 0,
            cspid_cur: ppt_prop_value::RESERVED_CSPID_CUR,
        }
    }
}

/// Drawing group header (without clusters)
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub(crate) struct EscherDggHeader {
    /// Maximum shape ID
    pub(crate) spid_max: u32,
    /// Number of clusters + 1
    pub(crate) cidcl: u32,
    /// Number of shapes saved
    pub(crate) csp_saved: u32,
    /// Number of drawings saved
    pub(crate) cdg_saved: u32,
}

// =============================================================================
// Drawing (EscherDg) - MS-ODRAW 2.2.14
// =============================================================================

/// Drawing record data
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub(crate) struct EscherDgData {
    /// Number of shapes in this drawing
    pub(crate) csp: u32,
    /// Next available shape ID
    pub(crate) spid_cur: u32,
}

impl EscherDgData {
    pub(crate) fn new(shape_count: u32, drawing_id: u32) -> Self {
        Self {
            csp: shape_count,
            spid_cur: (drawing_id << 10) + shape_count,
        }
    }
}

// =============================================================================
// Shape Group (EscherSpgr) - MS-ODRAW 2.2.38
// =============================================================================

/// Shape group bounding rectangle
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub(crate) struct EscherSpgrData {
    pub(crate) left: i32,
    pub(crate) top: i32,
    pub(crate) right: i32,
    pub(crate) bottom: i32,
}

impl EscherSpgrData {
    pub(crate) const ZERO: Self = Self {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
    };
}

/// Default drawing group properties (8 properties = 48 bytes)
pub(crate) const DGG_DEFAULT_PROPERTIES: [Property; 8] = [
    Property::new(prop_id::FILL_COLOR, ppt_prop_value::SCHEME_FILL),
    Property::new(prop_id::FILL_BACK_COLOR, ppt_prop_value::SCHEME_FILL_BACK),
    Property::new(prop_id::FILL_BLIP, 0),
    Property::new(
        prop_id::NO_FILL_HIT_TEST,
        ppt_prop_value::LINE_STYLE_DEFAULT,
    ),
    Property::new(prop_id::LINE_COLOR, ppt_prop_value::SCHEME_LINE),
    Property::new(prop_id::LINE_BLIP, 0),
    Property::new(
        prop_id::LINE_STYLE_BOOL,
        ppt_prop_value::LINE_STYLE_BOOL_DEFAULT,
    ),
    Property::new(prop_id::SHADOW_COLOR, ppt_prop_value::SCHEME_SHADOW),
];

/// Background shape properties (8 properties = 48 bytes)
pub(crate) const BG_SHAPE_PROPERTIES: [Property; 8] = [
    Property::new(prop_id::FILL_COLOR, ppt_prop_value::BG_FILL_COLOR),
    Property::new(prop_id::FILL_BACK_COLOR, ppt_prop_value::BG_FILL_BACK_COLOR),
    Property::new(prop_id::FILL_RECT_RIGHT, ppt_prop_value::SLIDE_WIDTH_EMU),
    Property::new(prop_id::FILL_RECT_BOTTOM, ppt_prop_value::SLIDE_HEIGHT_EMU),
    Property::new(prop_id::NO_FILL_HIT_TEST, ppt_prop_value::NO_FILL_HIT_TEST),
    Property::new(prop_id::LINE_STYLE_BOOL, ppt_prop_value::NO_LINE_DRAW_DASH),
    Property::new(prop_id::BW_MODE, ppt_prop_value::BW_MODE_AUTO),
    Property::new(prop_id::BACKGROUND_SHAPE, ppt_prop_value::BACKGROUND_SHAPE),
];

/// Child anchor with full coordinates (16 bytes)
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub(crate) struct ChildAnchor {
    pub(crate) left: i32,
    pub(crate) top: i32,
    pub(crate) right: i32,
    pub(crate) bottom: i32,
}

// =============================================================================
// User Shape Building
// =============================================================================

/// Owned custom/freeform OfficeArt geometry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FreeformGeometry {
    coordinate_space: GeometryRect,
    path_type: ShapePathType,
    vertices: Vec<(i32, i32)>,
    segment_info: Vec<u16>,
}

impl FreeformGeometry {
    /// Create freeform geometry in its internal OfficeArt coordinate space.
    pub fn new(
        coordinate_space: GeometryRect,
        path_type: ShapePathType,
        vertices: Vec<(i32, i32)>,
        segment_info: Vec<u16>,
    ) -> Self {
        Self {
            coordinate_space,
            path_type,
            vertices,
            segment_info,
        }
    }

    /// Return the internal geometry coordinate space.
    pub const fn coordinate_space(&self) -> GeometryRect {
        self.coordinate_space
    }

    /// Return how the vertices and segment array define the path.
    pub const fn path_type(&self) -> ShapePathType {
        self.path_type
    }

    /// Return the freeform vertices.
    pub fn vertices(&self) -> &[(i32, i32)] {
        &self.vertices
    }

    /// Return raw MSOPATHINFO words.
    pub fn segment_info(&self) -> &[u16] {
        &self.segment_info
    }

    pub(crate) fn validate(&self) -> Result<(u16, u16), Error> {
        if self.vertices.is_empty() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "freeform geometry requires at least one vertex",
            ));
        }
        if self.segment_info.is_empty() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "freeform geometry requires segment information",
            ));
        }
        if self.path_type == ShapePathType::Unknown {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "freeform geometry requires a known shape path type",
            ));
        }
        let vertex_count = u16::try_from(self.vertices.len()).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "freeform geometry exceeds 65535 vertices",
            )
        })?;
        let segment_count = u16::try_from(self.segment_info.len()).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "freeform geometry exceeds 65535 segment entries",
            )
        })?;
        Ok((vertex_count, segment_count))
    }

    pub(super) fn encode_arrays(&self) -> Result<(Vec<u8>, Vec<u8>), Error> {
        let (vertex_count, segment_count) = self.validate()?;

        let mut vertices = Vec::with_capacity(6 + self.vertices.len() * 8);
        vertices.extend_from_slice(&vertex_count.to_le_bytes());
        vertices.extend_from_slice(&vertex_count.to_le_bytes());
        vertices.extend_from_slice(&8u16.to_le_bytes());
        for &(x, y) in &self.vertices {
            vertices.extend_from_slice(&x.to_le_bytes());
            vertices.extend_from_slice(&y.to_le_bytes());
        }

        let mut segments = Vec::with_capacity(6 + self.segment_info.len() * 2);
        segments.extend_from_slice(&segment_count.to_le_bytes());
        segments.extend_from_slice(&segment_count.to_le_bytes());
        segments.extend_from_slice(&2u16.to_le_bytes());
        for &segment in &self.segment_info {
            segments.extend_from_slice(&segment.to_le_bytes());
        }

        Ok((vertices, segments))
    }
}

/// Shape data for building user shapes
#[derive(Debug, Clone)]
pub(crate) struct UserShapeData {
    /// Shape type (Escher MSOSPT value)
    pub shape_type: u16,
    /// Position and size in EMUs
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    /// Fill color (RGB, None = no fill)
    pub fill_color: Option<u32>,
    /// Fill type (0=solid, 4=shade/gradient, 5=shadecenter, etc.)
    pub fill_type: Option<u32>,
    /// Fill opacity (0-65536, 65536 = 100%)
    pub fill_opacity: Option<u32>,
    /// Fill back color (for gradients)
    pub fill_back_color: Option<u32>,
    /// Fill gradient angle (in degrees * 65536)
    pub fill_angle: Option<i32>,
    /// Fill BLIP index for pattern, texture, or picture fills.
    pub fill_blip_index: Option<u32>,
    /// Line color (RGB, None = no line)
    pub line_color: Option<u32>,
    /// Line width in EMUs (12700 = 1pt)
    pub line_width: Option<i32>,
    /// Line opacity (0-65536, 65536 = 100%).
    pub line_opacity: Option<u32>,
    /// Compound line style.
    pub line_style: Option<u32>,
    /// Line dash style (0=solid, 1=dash, 2=dot, etc.)
    pub line_dash_style: Option<u32>,
    /// Line start arrow style
    pub line_start_arrow: Option<u32>,
    /// Line end arrow style
    pub line_end_arrow: Option<u32>,
    /// Line start arrow width.
    pub line_start_arrow_width: Option<u32>,
    /// Line start arrow length.
    pub line_start_arrow_length: Option<u32>,
    /// Line end arrow width.
    pub line_end_arrow_width: Option<u32>,
    /// Line end arrow length.
    pub line_end_arrow_length: Option<u32>,
    /// Line join style.
    pub line_join_style: Option<u32>,
    /// Line end-cap style.
    pub line_end_cap_style: Option<u32>,
    /// Text content (simple string, ignored if paragraphs set)
    pub text: Option<String>,
    /// Rich text paragraphs (with formatting)
    pub paragraphs: Option<Vec<crate::writer::text_format::Paragraph>>,
    /// PowerPoint 11 document smart-tag indices for each rich-text run.
    pub smart_tag_runs: Option<Vec<Vec<u32>>>,
    /// Text type for TextHeaderAtom (0=Title, 1=Body, 2=Notes, 4=Other)
    pub text_type: u32,
    /// Placeholder type for notes/master shapes (None = not a placeholder)
    pub placeholder_type: Option<u8>,
    /// Shadow enabled
    pub has_shadow: bool,
    /// Flip horizontal
    pub flip_h: bool,
    /// Flip vertical
    pub flip_v: bool,
    /// Shape rotation in signed 16.16 fixed-point degrees.
    pub rotation: Option<i32>,
    /// Shape-specific adjustment values, at indices 0 through 9.
    pub adjust_values: Vec<i32>,
    /// Hyperlink ID (reference to ExObjList)
    pub hyperlink_id: Option<u32>,
    /// Hyperlink action type (for InteractiveInfoAtom)
    pub hyperlink_action: u8,
    /// Hyperlink jump type (for InteractiveInfoAtom)
    pub hyperlink_jump: u8,
    /// Hyperlink type (for InteractiveInfoAtom)
    pub hyperlink_type: u8,
    /// Typed click and mouse-over interactions.
    ///
    /// These take precedence over the legacy single-hyperlink fields for the
    /// corresponding trigger.
    pub interactions: Vec<crate::Interaction>,
    /// Typed actions attached to UTF-16 ranges in this shape's text.
    pub text_interactions: Vec<crate::TextInteraction>,
    /// Picture BLIP index (for picture frames)
    pub picture_index: Option<u32>,
    /// Explicit custom/freeform geometry.
    pub freeform_geometry: Option<FreeformGeometry>,
    /// Animation info for this shape
    pub animation_info: Option<crate::animation::AnimationInfo>,
    /// Shadow color (RGB format)
    pub shadow_color: Option<u32>,
    /// Shadow X offset in EMUs
    pub shadow_offset_x: Option<i32>,
    /// Shadow Y offset in EMUs
    pub shadow_offset_y: Option<i32>,
    /// Shadow opacity (0-65536)
    pub shadow_opacity: Option<u32>,
    /// Shadow type
    pub shadow_type: Option<u32>,
}

impl Default for UserShapeData {
    fn default() -> Self {
        Self {
            shape_type: shape_type::RECTANGLE,
            x: 0,
            y: 0,
            width: 914400, // 1 inch
            height: 914400,
            fill_color: None,
            fill_type: None,
            fill_opacity: None,
            fill_back_color: None,
            fill_angle: None,
            fill_blip_index: None,
            line_color: None,
            line_width: None,
            line_opacity: None,
            line_style: None,
            line_dash_style: None,
            line_start_arrow: None,
            line_end_arrow: None,
            line_start_arrow_width: None,
            line_start_arrow_length: None,
            line_end_arrow_width: None,
            line_end_arrow_length: None,
            line_join_style: None,
            line_end_cap_style: None,
            text: None,
            paragraphs: None,
            smart_tag_runs: None,
            text_type: 4,           // OTHER by default
            placeholder_type: None, // Not a placeholder by default
            has_shadow: false,
            flip_h: false,
            flip_v: false,
            rotation: None,
            adjust_values: Vec::new(),
            hyperlink_id: None,
            hyperlink_action: 4, // ACTION_HYPERLINK
            hyperlink_jump: 0,   // JUMP_NONE
            hyperlink_type: 8,   // LINK_Url
            interactions: Vec::new(),
            text_interactions: Vec::new(),
            picture_index: None,
            freeform_geometry: None,
            animation_info: None,
            shadow_color: None,
            shadow_offset_x: None,
            shadow_offset_y: None,
            shadow_opacity: None,
            shadow_type: None,
        }
    }
}
