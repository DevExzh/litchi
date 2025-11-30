//! Escher (Office Drawing) record generation for PPT files
//!
//! Escher is the binary format for drawing objects (shapes, connectors, pictures)
//! shared across Office applications.
//!
//! Based on Microsoft's "[MS-ODRAW]" specification and Apache POI's EscherRecord classes.

use bitflags::bitflags;
use std::io::Write;
use zerocopy::IntoBytes;
use zerocopy_derive::*;

/// Error type for PPT operations
pub type PptError = std::io::Error;

// =============================================================================
// Escher Record Types (MS-ODRAW 2.1.1)
// =============================================================================

/// Escher record types
pub mod record_type {
    /// Drawing group container
    pub const DGG_CONTAINER: u16 = 0xF000;
    /// BLIP store container
    pub const BSTORE_CONTAINER: u16 = 0xF001;
    /// Drawing container
    pub const DG_CONTAINER: u16 = 0xF002;
    /// Shape group container
    pub const SPGR_CONTAINER: u16 = 0xF003;
    /// Shape container
    pub const SP_CONTAINER: u16 = 0xF004;
    /// Drawing group record
    pub const DGG: u16 = 0xF006;
    /// Drawing record
    pub const DG: u16 = 0xF008;
    /// Shape group coordinates
    pub const SPGR: u16 = 0xF009;
    /// Shape record
    pub const SP: u16 = 0xF00A;
    /// Property table
    pub const OPT: u16 = 0xF00B;
    /// Client anchor
    pub const CLIENT_ANCHOR: u16 = 0xF010;
    /// Client data
    pub const CLIENT_DATA: u16 = 0xF011;
    /// Split menu colors
    pub const SPLIT_MENU_COLORS: u16 = 0xF11E;
}

// =============================================================================
// Shape Types (MS-ODRAW 2.4.6)
// =============================================================================

/// Shape types (MSOSPT values)
pub mod shape_type {
    pub const NOT_PRIMITIVE: u16 = 0;
    pub const RECTANGLE: u16 = 1;
    pub const ROUND_RECTANGLE: u16 = 2;
    pub const ELLIPSE: u16 = 3;
    pub const DIAMOND: u16 = 4;
    pub const LINE: u16 = 20;
    pub const TEXT_BOX: u16 = 202;
}

// =============================================================================
// Shape Flags (MS-ODRAW 2.2.40)
// =============================================================================

bitflags! {
    /// Shape flags for EscherSpRecord (MS-ODRAW 2.2.40)
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub struct ShapeFlags: u32 {
        /// Shape is a group
        const GROUP = 0x0001;
        /// Shape is a child of a group
        const CHILD = 0x0002;
        /// Shape is the topmost group (patriarch)
        const PATRIARCH = 0x0004;
        /// Shape has been deleted
        const DELETED = 0x0008;
        /// Shape is an OLE object
        const OLE_SHAPE = 0x0010;
        /// Shape has a valid master
        const HAVE_MASTER = 0x0020;
        /// Shape is flipped horizontally
        const FLIP_H = 0x0040;
        /// Shape is flipped vertically
        const FLIP_V = 0x0080;
        /// Shape is a connector
        const CONNECTOR = 0x0100;
        /// Shape has an anchor
        const HAVE_ANCHOR = 0x0200;
        /// Shape is a background shape
        const BACKGROUND = 0x0400;
        /// Shape has a shape type property
        const HAVE_SPT = 0x0800;
    }
}

// =============================================================================
// Property IDs (MS-ODRAW 2.3.1)
// =============================================================================

/// Escher property IDs
pub mod prop_id {
    // Protection group
    pub const LOCK_ROTATION: u16 = 0x0077;
    pub const LOCK_ASPECT_RATIO: u16 = 0x0078;
    pub const LOCK_POSITION: u16 = 0x0079;
    pub const LOCK_AGAINST_SELECT: u16 = 0x007A;
    pub const LOCK_CROPPING: u16 = 0x007B;
    pub const LOCK_VERTICES: u16 = 0x007C;
    pub const LOCK_TEXT: u16 = 0x007D;
    pub const LOCK_ADJUST_HANDLES: u16 = 0x007E;
    pub const LOCK_AGGR: u16 = 0x007F;

    // Transform group
    pub const ADJUST_VALUE: u16 = 0x0080;
    pub const ADJUST2_VALUE: u16 = 0x0081;

    // Fill style (MS-ODRAW section 2.3.7)
    pub const FILL_TYPE: u16 = 0x0180;
    pub const FILL_COLOR: u16 = 0x0181;
    pub const FILL_OPACITY: u16 = 0x0182;
    pub const FILL_BACK_COLOR: u16 = 0x0183;
    pub const FILL_BACK_OPACITY: u16 = 0x0184;
    pub const FILL_BLIP: u16 = 0x4186;
    pub const FILL_WIDTH: u16 = 0x0187; // fillWidth for pattern fills
    pub const FILL_HEIGHT: u16 = 0x0188; // fillHeight for pattern fills
    pub const FILL_ANGLE: u16 = 0x018A; // fillAngle for gradients (degrees * 65536)
    pub const FILL_FOCUS: u16 = 0x018B; // fillFocus for gradients (-100 to 100)
    pub const FILL_RECT_RIGHT: u16 = 0x0193; // fillRectRight per MS-ODRAW
    pub const FILL_RECT_BOTTOM: u16 = 0x0194; // fillRectBottom per MS-ODRAW
    pub const NO_FILL_HIT_TEST: u16 = 0x01BF;

    // Line style
    pub const LINE_COLOR: u16 = 0x01C0;
    pub const LINE_OPACITY: u16 = 0x01C1;
    pub const LINE_BACK_COLOR: u16 = 0x01C2;
    pub const LINE_WIDTH: u16 = 0x01CB;
    pub const LINE_STYLE: u16 = 0x01CD;
    pub const LINE_DASH_STYLE: u16 = 0x01CE;
    pub const LINE_START_ARROW: u16 = 0x01D0;
    pub const LINE_END_ARROW: u16 = 0x01D1;
    pub const LINE_START_ARROW_WIDTH: u16 = 0x01D2;
    pub const LINE_START_ARROW_LENGTH: u16 = 0x01D3;
    pub const LINE_END_ARROW_WIDTH: u16 = 0x01D4;
    pub const LINE_END_ARROW_LENGTH: u16 = 0x01D5;
    pub const LINE_BLIP: u16 = 0x41C5;
    pub const LINE_STYLE_BOOL: u16 = 0x01FF;

    // Shadow style
    pub const SHADOW_TYPE: u16 = 0x0200;
    pub const SHADOW_COLOR: u16 = 0x0201;
    pub const SHADOW_OPACITY: u16 = 0x0204;
    pub const SHADOW_OFFSET_X: u16 = 0x0205;
    pub const SHADOW_OFFSET_Y: u16 = 0x0206;
    pub const SHADOW_BOOL: u16 = 0x023F; // shadowObscured

    // Shape
    pub const BW_MODE: u16 = 0x0304;
    pub const SHAPE_BOOL: u16 = 0x01FF;
    pub const BACKGROUND_SHAPE: u16 = 0x017F; // fBackground per MS-ODRAW
}

// =============================================================================
// Property Values (scheme colors, etc.)
// =============================================================================

/// Common property values
pub mod prop_value {
    /// Scheme color flag (OR with scheme index)
    pub const SCHEME_COLOR: u32 = 0x0800_0000;

    /// Scheme color indices
    pub const SCHEME_FILL: u32 = SCHEME_COLOR | 0x04;
    pub const SCHEME_FILL_BACK: u32 = SCHEME_COLOR;
    pub const SCHEME_LINE: u32 = SCHEME_COLOR | 0x01;
    pub const SCHEME_SHADOW: u32 = SCHEME_COLOR | 0x02;

    /// Line style boolean properties
    pub const LINE_STYLE_DEFAULT: u32 = 0x0010_0010;

    /// Shape boolean properties
    pub const SHAPE_BOOL_DEFAULT: u32 = 0x0008_0008;

    /// Background fill color
    pub const BG_FILL_COLOR: u32 = 134_217_728; // 0x0800_0000
    pub const BG_FILL_BACK_COLOR: u32 = 134_217_733; // 0x0800_0005

    /// Slide dimensions (EMUs)
    pub const SLIDE_WIDTH_EMU: u32 = 10_064_750; // 914400 * 11
    pub const SLIDE_HEIGHT_EMU: u32 = 7_778_750; // 914400 * 8.5

    /// No fill hit test value
    pub const NO_FILL_HIT_TEST: u32 = 1_179_666; // 0x0012_0012
    pub const NO_LINE_DRAW_DASH: u32 = 524_288; // 0x0008_0000

    /// Black and white mode
    pub const BW_MODE_AUTO: u32 = 9;

    /// Background shape flag
    pub const BACKGROUND_SHAPE: u32 = 65_537; // 0x0001_0001

    /// Reserved cluster cspidCur
    pub const RESERVED_CSPID_CUR: u32 = 4;

    /// POI master shape count
    pub const POI_MASTER_SHAPE_COUNT: u32 = 6;
    pub const POI_SPID_MAX: u32 = 3076;
}

// =============================================================================
// Split Menu Colors
// =============================================================================

/// Split menu colors structure
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct SplitMenuColors {
    pub fill_color: u32,
    pub line_color: u32,
    pub shadow_color: u32,
    pub color_3d: u32,
}

impl SplitMenuColors {
    pub const DEFAULT: Self = Self {
        fill_color: prop_value::SCHEME_FILL,
        line_color: prop_value::SCHEME_LINE,
        shadow_color: prop_value::SCHEME_SHADOW,
        color_3d: 0x1000_00F7,
    };
}

// =============================================================================
// Escher Record Header (MS-ODRAW 2.2.1)
// =============================================================================

/// Escher record header versions
pub mod header_version {
    pub const CONTAINER: u8 = 0x0F;
    pub const SIMPLE: u8 = 0x00;
    pub const SPGR: u8 = 0x01;
    pub const SP: u8 = 0x02;
    pub const OPT: u8 = 0x03;
    pub const DG: u8 = 0x00; // instance is drawing_id
}

/// Raw Escher record header (8 bytes) - zerocopy compatible
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct EscherRecordHeader {
    /// Version (4 bits) | Instance (12 bits)
    pub ver_inst: u16,
    /// Record type
    pub rec_type: u16,
    /// Record length (not including header)
    pub length: u32,
}

impl EscherRecordHeader {
    /// Create a new header with version, instance, type, and length
    pub const fn new(version: u8, instance: u16, rec_type: u16, length: u32) -> Self {
        let ver_inst = (version as u16 & 0x0F) | ((instance & 0x0FFF) << 4);
        Self {
            ver_inst,
            rec_type,
            length,
        }
    }

    /// Create a container header
    pub const fn container(rec_type: u16, length: u32) -> Self {
        Self::new(header_version::CONTAINER, 0, rec_type, length)
    }
}

/// Escher record header (8 bytes) - builder-friendly version
#[derive(Debug, Clone)]
pub struct EscherHeader {
    /// Version (4 bits)
    pub version: u8,
    /// Instance (12 bits)
    pub instance: u16,
    /// Record type
    pub record_type: u16,
    /// Length
    pub length: u32,
}

impl EscherHeader {
    /// Create a new Escher header
    pub fn new(version: u8, instance: u16, record_type: u16, length: u32) -> Self {
        Self {
            version: version & 0x0F,
            instance: instance & 0x0FFF,
            record_type,
            length,
        }
    }

    /// Write header to writer
    pub fn write<W: Write>(&self, writer: &mut W) -> Result<(), PptError> {
        let raw =
            EscherRecordHeader::new(self.version, self.instance, self.record_type, self.length);
        writer.write_all(raw.as_bytes())?;
        Ok(())
    }
}

// =============================================================================
// Drawing Group (EscherDgg) - MS-ODRAW 2.2.12
// =============================================================================

/// File ID cluster entry
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct FileIdCluster {
    /// Drawing group ID
    pub dgid: u32,
    /// Next available shape ID in cluster
    pub cspid_cur: u32,
}

impl FileIdCluster {
    pub const fn new(dgid: u32, cspid_cur: u32) -> Self {
        Self { dgid, cspid_cur }
    }

    pub const fn reserved() -> Self {
        Self {
            dgid: 0,
            cspid_cur: prop_value::RESERVED_CSPID_CUR,
        }
    }
}

/// Drawing group header (without clusters)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct EscherDggHeader {
    /// Maximum shape ID
    pub spid_max: u32,
    /// Number of clusters + 1
    pub cidcl: u32,
    /// Number of shapes saved
    pub csp_saved: u32,
    /// Number of drawings saved
    pub cdg_saved: u32,
}

// =============================================================================
// Drawing (EscherDg) - MS-ODRAW 2.2.14
// =============================================================================

/// Drawing record data
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct EscherDgData {
    /// Number of shapes in this drawing
    pub csp: u32,
    /// Next available shape ID
    pub spid_cur: u32,
}

impl EscherDgData {
    pub fn new(shape_count: u32, drawing_id: u32) -> Self {
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
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct EscherSpgrData {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

impl EscherSpgrData {
    pub const ZERO: Self = Self {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
    };
}

// =============================================================================
// Shape (EscherSp) - MS-ODRAW 2.2.40
// =============================================================================

/// Shape record data
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct EscherSpData {
    /// Shape ID
    pub spid: u32,
    /// Shape flags
    pub flags: u32,
}

impl EscherSpData {
    pub const fn new(spid: u32, flags: u32) -> Self {
        Self { spid, flags }
    }

    pub const fn with_flags(spid: u32, flags: ShapeFlags) -> Self {
        Self {
            spid,
            flags: flags.bits(),
        }
    }

    pub const fn group_patriarch(spid: u32) -> Self {
        Self {
            spid,
            flags: ShapeFlags::GROUP.bits() | ShapeFlags::PATRIARCH.bits(),
        }
    }

    pub const fn background(spid: u32) -> Self {
        Self {
            spid,
            flags: ShapeFlags::BACKGROUND.bits() | ShapeFlags::HAVE_SPT.bits(),
        }
    }
}

// =============================================================================
// Property Entry (EscherOpt) - MS-ODRAW 2.3.1
// =============================================================================

/// Single property entry (6 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct EscherProperty {
    /// Property ID (with flags in high bits)
    pub prop_id: u16,
    /// Property value
    pub value: u32,
}

impl EscherProperty {
    pub const fn new(prop_id: u16, value: u32) -> Self {
        Self { prop_id, value }
    }
}

/// Default drawing group properties (8 properties = 48 bytes)
pub const DGG_DEFAULT_PROPERTIES: [EscherProperty; 8] = [
    EscherProperty::new(prop_id::FILL_COLOR, prop_value::SCHEME_FILL),
    EscherProperty::new(prop_id::FILL_BACK_COLOR, prop_value::SCHEME_FILL_BACK),
    EscherProperty::new(prop_id::FILL_BLIP, 0),
    EscherProperty::new(prop_id::NO_FILL_HIT_TEST, prop_value::LINE_STYLE_DEFAULT),
    EscherProperty::new(prop_id::LINE_COLOR, prop_value::SCHEME_LINE),
    EscherProperty::new(prop_id::LINE_BLIP, 0),
    EscherProperty::new(prop_id::SHAPE_BOOL, prop_value::SHAPE_BOOL_DEFAULT),
    EscherProperty::new(prop_id::SHADOW_COLOR, prop_value::SCHEME_SHADOW),
];

/// Background shape properties (8 properties = 48 bytes)
pub const BG_SHAPE_PROPERTIES: [EscherProperty; 8] = [
    EscherProperty::new(prop_id::FILL_COLOR, prop_value::BG_FILL_COLOR),
    EscherProperty::new(prop_id::FILL_BACK_COLOR, prop_value::BG_FILL_BACK_COLOR),
    EscherProperty::new(prop_id::FILL_RECT_RIGHT, prop_value::SLIDE_WIDTH_EMU),
    EscherProperty::new(prop_id::FILL_RECT_BOTTOM, prop_value::SLIDE_HEIGHT_EMU),
    EscherProperty::new(prop_id::NO_FILL_HIT_TEST, prop_value::NO_FILL_HIT_TEST),
    EscherProperty::new(prop_id::LINE_STYLE_BOOL, prop_value::NO_LINE_DRAW_DASH),
    EscherProperty::new(prop_id::BW_MODE, prop_value::BW_MODE_AUTO),
    EscherProperty::new(prop_id::BACKGROUND_SHAPE, prop_value::BACKGROUND_SHAPE),
];

// =============================================================================
// Client Anchor - MS-ODRAW 2.2.46
// =============================================================================

/// Client anchor (8 bytes for PPT)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct ClientAnchor {
    pub left: u16,
    pub top: u16,
    pub right: u16,
    pub bottom: u16,
}

impl ClientAnchor {
    pub const ZERO: Self = Self {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
    };
}

// =============================================================================
// PPT Record Types
// =============================================================================

/// PPT-specific record types embedded in Escher
pub mod ppt_record_type {
    /// OEPlaceholderAtom
    pub const OE_PLACEHOLDER_ATOM: u16 = 0x0BC3;
}

/// OEPlaceholderAtom data (8 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct OEPlaceholderAtom {
    /// Placement ID
    pub placement_id: u32,
    /// Placeholder type
    pub placeholder_type: u8,
    /// Placeholder size
    pub placeholder_size: u8,
    /// Unused
    pub unused: u16,
}

impl OEPlaceholderAtom {
    pub const BACKGROUND: Self = Self {
        placement_id: 0,
        placeholder_type: 0,
        placeholder_size: 0,
        unused: 0,
    };
}

/// Escher record builder
pub struct EscherBuilder {
    header: EscherHeader,
    data: Vec<u8>,
}

impl EscherBuilder {
    /// Create a new Escher record builder
    pub fn new(version: u8, instance: u16, record_type: u16) -> Self {
        Self {
            header: EscherHeader::new(version, instance, record_type, 0),
            data: Vec::new(),
        }
    }

    /// Add data to the record
    pub fn add_data(&mut self, data: &[u8]) {
        self.data.extend_from_slice(data);
        self.header.length = self.data.len() as u32;
    }

    /// Build the complete record
    pub fn build(&self) -> Result<Vec<u8>, PptError> {
        let mut record = Vec::new();
        self.header.write(&mut record)?;
        record.extend_from_slice(&self.data);
        Ok(record)
    }
}

/// Create a DggContainer (Drawing Group Container) per MS-ODRAW
///
/// # Arguments
/// * `master_shapes` - Number of shapes in the master (6 for POI template)
/// * `slide_shape_counts` - Shape count for each slide (including group+background, so user_shapes+2)
pub fn create_dgg_container(
    master_shapes: u32,
    slide_shape_counts: &[u32],
) -> Result<Vec<u8>, PptError> {
    let mut container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::DGG_CONTAINER);

    // Total drawings = 1 (master) + number of slides
    let drawing_count = (slide_shape_counts.len() as u32) + 1;

    // Calculate total shapes: master + sum of all slide shapes
    let total_slide_shapes: u32 = slide_shape_counts.iter().sum();
    let csp_saved = master_shapes + total_slide_shapes;

    // POI uses cidcl=4 (3 clusters) even for 1 drawing
    let num_clusters = std::cmp::max(3, drawing_count as usize);
    let cidcl = (num_clusters + 1) as u32;

    // spidMax: Calculate based on highest drawing ID * 1024 + shapes in that drawing
    let max_slide_shapes = slide_shape_counts.iter().max().copied().unwrap_or(2);
    let spid_max = if drawing_count == 1 && master_shapes == prop_value::POI_MASTER_SHAPE_COUNT {
        prop_value::POI_SPID_MAX
    } else {
        drawing_count * 1024 + max_slide_shapes
    };

    // Build EscherDgg record (OfficeArtFDGGBlock)
    let mut dgg = EscherBuilder::new(header_version::SIMPLE, 0, record_type::DGG);
    let mut dgg_data = Vec::with_capacity(16 + num_clusters * 8);

    // Write header using zerocopy struct
    let header = EscherDggHeader {
        spid_max,
        cidcl,
        csp_saved,
        cdg_saved: drawing_count,
    };
    dgg_data.extend_from_slice(header.as_bytes());

    // FileIdClusters: each drawing gets its own cluster
    // dg_id 1 = master, dg_id 2+ = slides
    for dg_id in 1..=drawing_count {
        let cspid_cur = if dg_id == 1 {
            master_shapes + 1
        } else {
            let slide_idx = (dg_id - 2) as usize;
            slide_shape_counts.get(slide_idx).copied().unwrap_or(2) + 1
        };
        let cluster = FileIdCluster::new(dg_id, cspid_cur);
        dgg_data.extend_from_slice(cluster.as_bytes());
    }

    // Add reserved cluster slots to match POI
    for _ in drawing_count..num_clusters as u32 {
        dgg_data.extend_from_slice(FileIdCluster::reserved().as_bytes());
    }
    dgg.add_data(&dgg_data);
    container.add_data(&dgg.build()?);

    // NOTE: BStore container goes here if pictures are present
    // Call create_dgg_container_with_blips() instead if you have pictures

    // Add EscherOpt with default properties using const array
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        DGG_DEFAULT_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &DGG_DEFAULT_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    container.add_data(&opt.build()?);

    // Add SplitMenuColors using zerocopy struct
    let mut colors = EscherBuilder::new(header_version::SIMPLE, 4, record_type::SPLIT_MENU_COLORS);
    colors.add_data(SplitMenuColors::DEFAULT.as_bytes());
    container.add_data(&colors.build()?);

    container.build()
}

/// Create a DggContainer with BStore for pictures
///
/// Same as `create_dgg_container` but includes a BStoreContainer for pictures.
/// The bstore_blob should be the raw bytes from `BlipStoreBuilder::build()`.
pub fn create_dgg_container_with_blips(
    master_shapes: u32,
    slide_shape_counts: &[u32],
    bstore_blob: &[u8],
) -> Result<Vec<u8>, PptError> {
    let mut container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::DGG_CONTAINER);

    // Total drawings = 1 (master) + number of slides
    let drawing_count = (slide_shape_counts.len() as u32) + 1;

    // Calculate total shapes: master + sum of all slide shapes
    let total_slide_shapes: u32 = slide_shape_counts.iter().sum();
    let csp_saved = master_shapes + total_slide_shapes;

    // POI uses cidcl=4 (3 clusters) even for 1 drawing
    let num_clusters = std::cmp::max(3, drawing_count as usize);
    let cidcl = (num_clusters + 1) as u32;

    // spidMax: Calculate based on highest drawing ID * 1024 + shapes in that drawing
    let max_slide_shapes = slide_shape_counts.iter().max().copied().unwrap_or(2);
    let spid_max = if drawing_count == 1 && master_shapes == prop_value::POI_MASTER_SHAPE_COUNT {
        prop_value::POI_SPID_MAX
    } else {
        drawing_count * 1024 + max_slide_shapes
    };

    // Build EscherDgg record (OfficeArtFDGGBlock)
    let mut dgg = EscherBuilder::new(header_version::SIMPLE, 0, record_type::DGG);
    let mut dgg_data = Vec::with_capacity(16 + num_clusters * 8);

    let header = EscherDggHeader {
        spid_max,
        cidcl,
        csp_saved,
        cdg_saved: drawing_count,
    };
    dgg_data.extend_from_slice(header.as_bytes());

    for dg_id in 1..=drawing_count {
        let cspid_cur = if dg_id == 1 {
            master_shapes + 1
        } else {
            let slide_idx = (dg_id - 2) as usize;
            slide_shape_counts.get(slide_idx).copied().unwrap_or(2) + 1
        };
        let cluster = FileIdCluster::new(dg_id, cspid_cur);
        dgg_data.extend_from_slice(cluster.as_bytes());
    }

    for _ in drawing_count..num_clusters as u32 {
        dgg_data.extend_from_slice(FileIdCluster::reserved().as_bytes());
    }
    dgg.add_data(&dgg_data);
    container.add_data(&dgg.build()?);

    // BStore container (if not empty)
    if !bstore_blob.is_empty() {
        container.add_data(bstore_blob);
    }

    // Add EscherOpt with default properties
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        DGG_DEFAULT_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &DGG_DEFAULT_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    container.add_data(&opt.build()?);

    // Add SplitMenuColors
    let mut colors = EscherBuilder::new(header_version::SIMPLE, 4, record_type::SPLIT_MENU_COLORS);
    colors.add_data(SplitMenuColors::DEFAULT.as_bytes());
    container.add_data(&colors.build()?);

    container.build()
}

/// Create a DgContainer (Drawing Container) for a slide
pub fn create_dg_container(drawing_id: u32, shape_count: u32) -> Result<Vec<u8>, PptError> {
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);

    // Add DG record (instance = drawing_id)
    let mut dg = EscherBuilder::new(header_version::DG, drawing_id as u16, record_type::DG);
    let total_shapes = shape_count.saturating_add(2); // +group +background
    let dg_data = EscherDgData::new(total_shapes, drawing_id);
    dg.add_data(dg_data.as_bytes());
    container.add_data(&dg.build()?);

    // SpgrContainer
    let mut spgr_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SPGR_CONTAINER);

    // Group patriarch SpContainer
    let mut group_sp_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    // Spgr (group bbox = all zeros)
    let mut spgr = EscherBuilder::new(header_version::SPGR, 0, record_type::SPGR);
    spgr.add_data(EscherSpgrData::ZERO.as_bytes());
    group_sp_container.add_data(&spgr.build()?);

    // Sp (group shape)
    let group_spid = drawing_id << 10;
    let mut sp = EscherBuilder::new(
        header_version::SP,
        shape_type::NOT_PRIMITIVE,
        record_type::SP,
    );
    sp.add_data(EscherSpData::group_patriarch(group_spid).as_bytes());
    group_sp_container.add_data(&sp.build()?);

    spgr_container.add_data(&group_sp_container.build()?);

    // Add SpgrContainer to DgContainer first
    container.add_data(&spgr_container.build()?);

    // Background shape container - added to DgContainer OUTSIDE SpgrContainer (per POI)
    let mut bg_sp_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    // Background EscherSp
    let bg_spid = group_spid + 1;
    let mut bg_sp = EscherBuilder::new(header_version::SP, shape_type::RECTANGLE, record_type::SP);
    bg_sp.add_data(EscherSpData::background(bg_spid).as_bytes());
    bg_sp_container.add_data(&bg_sp.build()?);

    // Background EscherOpt properties (per POI PPDrawing.create())
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        BG_SHAPE_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &BG_SHAPE_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    bg_sp_container.add_data(&opt.build()?);

    // NOTE: Per POI's PPDrawing.create(), background SpContainer has NO ClientAnchor or ClientData
    // Only Sp + Opt records are present

    // Add background SpContainer to DgContainer (NOT to SpgrContainer)
    container.add_data(&bg_sp_container.build()?);

    container.build()
}

/// Create a shape container (spContainer)
pub fn create_shape_container(
    shape_id: u32,
    stype: u16,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
) -> Result<Vec<u8>, PptError> {
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    // Add SP (Shape) record
    let mut sp = EscherBuilder::new(header_version::SP, stype, record_type::SP);
    let sp_data =
        EscherSpData::with_flags(shape_id, ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT);
    sp.add_data(sp_data.as_bytes());
    container.add_data(&sp.build()?);

    // Add client anchor (position and size)
    let mut anchor = EscherBuilder::new(header_version::SIMPLE, 0, record_type::CLIENT_ANCHOR);
    // Extended anchor format with position info
    let anchor_data = ChildAnchor {
        left: x,
        top: y,
        right: x + width,
        bottom: y + height,
    };
    anchor.add_data(anchor_data.as_bytes());
    container.add_data(&anchor.build()?);

    container.build()
}

/// Child anchor with full coordinates (16 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct ChildAnchor {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

// =============================================================================
// User Shape Building
// =============================================================================

/// Shape data for building user shapes
#[derive(Debug, Clone)]
pub struct UserShapeData {
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
    /// Line color (RGB, None = no line)
    pub line_color: Option<u32>,
    /// Line width in EMUs (12700 = 1pt)
    pub line_width: Option<i32>,
    /// Line dash style (0=solid, 1=dash, 2=dot, etc.)
    pub line_dash_style: Option<u32>,
    /// Line start arrow style
    pub line_start_arrow: Option<u32>,
    /// Line end arrow style
    pub line_end_arrow: Option<u32>,
    /// Text content (simple string, ignored if paragraphs set)
    pub text: Option<String>,
    /// Rich text paragraphs (with formatting)
    pub paragraphs: Option<Vec<super::text_format::Paragraph>>,
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
    /// Hyperlink ID (reference to ExObjList)
    pub hyperlink_id: Option<u32>,
    /// Hyperlink action type (for InteractiveInfoAtom)
    pub hyperlink_action: u8,
    /// Hyperlink jump type (for InteractiveInfoAtom)
    pub hyperlink_jump: u8,
    /// Hyperlink type (for InteractiveInfoAtom)
    pub hyperlink_type: u8,
    /// Picture BLIP index (1-based, for picture shapes)
    pub picture_index: Option<u32>,
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
            line_color: Some(0x000000), // Black line by default
            line_width: Some(12700),    // 1pt
            line_dash_style: None,
            line_start_arrow: None,
            line_end_arrow: None,
            text: None,
            paragraphs: None,
            text_type: 4,           // OTHER by default
            placeholder_type: None, // Not a placeholder by default
            has_shadow: false,
            flip_h: false,
            flip_v: false,
            hyperlink_id: None,
            hyperlink_action: 4, // ACTION_HYPERLINK
            hyperlink_jump: 0,   // JUMP_NONE
            hyperlink_type: 8,   // LINK_Url
            picture_index: None, // Not a picture by default
        }
    }
}

/// Create a DgContainer with user shapes
pub fn create_dg_container_with_shapes(
    drawing_id: u32,
    shapes: &[UserShapeData],
) -> Result<Vec<u8>, PptError> {
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);

    // Total shapes = group + background + user shapes
    let total_shapes = (shapes.len() as u32).saturating_add(2);

    // Add DG record
    let mut dg = EscherBuilder::new(header_version::DG, drawing_id as u16, record_type::DG);
    let dg_data = EscherDgData::new(total_shapes, drawing_id);
    dg.add_data(dg_data.as_bytes());
    container.add_data(&dg.build()?);

    // SpgrContainer
    let mut spgr_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SPGR_CONTAINER);

    // Group patriarch SpContainer
    let mut group_sp_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut spgr = EscherBuilder::new(header_version::SPGR, 0, record_type::SPGR);
    spgr.add_data(EscherSpgrData::ZERO.as_bytes());
    group_sp_container.add_data(&spgr.build()?);

    let group_spid = drawing_id << 10;
    let mut sp = EscherBuilder::new(
        header_version::SP,
        shape_type::NOT_PRIMITIVE,
        record_type::SP,
    );
    sp.add_data(EscherSpData::group_patriarch(group_spid).as_bytes());
    group_sp_container.add_data(&sp.build()?);

    spgr_container.add_data(&group_sp_container.build()?);

    // User shapes go INSIDE SpgrContainer (after group patriarch)
    let bg_spid = group_spid + 1;
    for (i, shape) in shapes.iter().enumerate() {
        let shape_spid = bg_spid + 1 + (i as u32);
        let sp_container = create_user_shape_container(shape_spid, shape)?;
        spgr_container.add_data(&sp_container);
    }

    // Add SpgrContainer to DgContainer
    container.add_data(&spgr_container.build()?);

    // Background shape container - added to DgContainer OUTSIDE SpgrContainer (per POI)
    let mut bg_sp_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut bg_sp = EscherBuilder::new(header_version::SP, shape_type::RECTANGLE, record_type::SP);
    bg_sp.add_data(EscherSpData::background(bg_spid).as_bytes());
    bg_sp_container.add_data(&bg_sp.build()?);

    let mut opt = EscherBuilder::new(
        header_version::OPT,
        BG_SHAPE_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &BG_SHAPE_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    bg_sp_container.add_data(&opt.build()?);

    // NOTE: Per POI's PPDrawing.create(), background SpContainer has NO ClientAnchor or ClientData
    // Only Sp + Opt records are present

    // Add background SpContainer to DgContainer (NOT to SpgrContainer)
    container.add_data(&bg_sp_container.build()?);

    container.build()
}

/// Convert EMU to master units (1/576 inch)
/// EMU = 914400 per inch, Master = 576 per inch
/// master = emu * 576 / 914400 = emu / 1588.0
fn emu_to_master(emu: i32) -> i16 {
    (emu as f64 / 1588.0).round() as i16
}

/// Create a user shape SpContainer
fn create_user_shape_container(shape_id: u32, shape: &UserShapeData) -> Result<Vec<u8>, PptError> {
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    // Shape flags
    let mut flags = ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT;
    if shape.flip_h {
        flags |= ShapeFlags::FLIP_H;
    }
    if shape.flip_v {
        flags |= ShapeFlags::FLIP_V;
    }

    // SP record
    let mut sp = EscherBuilder::new(header_version::SP, shape.shape_type, record_type::SP);
    sp.add_data(EscherSpData::with_flags(shape_id, flags).as_bytes());
    container.add_data(&sp.build()?);

    // OPT record with shape properties (sorted by property number, not full ID)
    // Per POI: sort by getPropertyNumber() which masks out flags (id & 0x3FFF)
    let mut properties = build_shape_properties(shape);
    properties.sort_by_key(|p| p.prop_id & 0x3FFF);
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        properties.len() as u16,
        record_type::OPT,
    );
    for prop in &properties {
        opt.add_data(prop.as_bytes());
    }
    container.add_data(&opt.build()?);

    // ClientAnchor with position/size (8-byte short format for PPT top-level shapes)
    // POI uses: flag(y1), col1(x1), dx1(x2), row1(y2) - all shorts in master units
    let mut anchor = EscherBuilder::new(header_version::SIMPLE, 0, record_type::CLIENT_ANCHOR);
    let x1 = emu_to_master(shape.x);
    let y1 = emu_to_master(shape.y);
    let x2 = emu_to_master(shape.x + shape.width);
    let y2 = emu_to_master(shape.y + shape.height);
    // Short record format: 8 bytes (4 shorts)
    anchor.add_data(&y1.to_le_bytes()); // flag/top
    anchor.add_data(&x1.to_le_bytes()); // col1/left
    anchor.add_data(&x2.to_le_bytes()); // dx1/right
    anchor.add_data(&y2.to_le_bytes()); // row1/bottom
    container.add_data(&anchor.build()?);

    // ClientData with OEPlaceholderAtom for placeholders OR InteractiveInfo for hyperlinks
    // MUST come BEFORE ClientTextbox per POI (addChildBefore(clientData, EscherTextboxRecord.RECORD_ID))
    if let Some(placeholder_type) = shape.placeholder_type {
        let client_data = build_client_data_with_placeholder(placeholder_type)?;
        container.add_data(&client_data);
    } else if let Some(hyperlink_id) = shape.hyperlink_id {
        let client_data = build_client_data_with_hyperlink(
            hyperlink_id,
            shape.hyperlink_action,
            shape.hyperlink_jump,
            shape.hyperlink_type,
        )?;
        container.add_data(&client_data);
    }

    // ClientTextBox if text present (prefer paragraphs with formatting over plain text)
    if let Some(paragraphs) = &shape.paragraphs {
        if !paragraphs.is_empty() {
            let textbox = build_client_textbox_formatted(paragraphs, shape.text_type)?;
            container.add_data(&textbox);
        }
    } else if let Some(text) = &shape.text {
        let textbox = build_client_textbox(text, shape.text_type)?;
        container.add_data(&textbox);
    }

    container.build()
}

/// Build ClientData record with InteractiveInfo for hyperlink
/// Per POI HSLFShape.getClientData(): clientData.setOptions((short)15) => version=0xF, instance=0
fn build_client_data_with_hyperlink(
    hyperlink_id: u32,
    action: u8,
    jump: u8,
    hyperlink_type: u8,
) -> Result<Vec<u8>, PptError> {
    // Build InteractiveInfoAtom manually (16 bytes of data per POI)
    // Offset 0-3: soundRef = 0
    // Offset 4-7: hyperlinkID
    // Offset 8: action
    // Offset 9: oleVerb = 0
    // Offset 10: jump
    // Offset 11: flags = 0
    // Offset 12: hyperlinkType
    // Offset 13-15: reserved = 0
    let mut atom_data = [0u8; 16];
    atom_data[4..8].copy_from_slice(&hyperlink_id.to_le_bytes());
    atom_data[8] = action;
    atom_data[10] = jump;
    atom_data[12] = hyperlink_type;

    // InteractiveInfoAtom PPT record header (type 4083, 16 bytes data)
    let mut info_atom = Vec::with_capacity(24);
    info_atom.extend_from_slice(&[0x00, 0x00]); // version=0, instance=0
    info_atom.extend_from_slice(&4083u16.to_le_bytes()); // RT_InteractiveInfoAtom
    info_atom.extend_from_slice(&16u32.to_le_bytes()); // length
    info_atom.extend_from_slice(&atom_data);

    // InteractiveInfo container PPT record (type 4082)
    let mut info_container = Vec::with_capacity(32);
    info_container.extend_from_slice(&[0x0F, 0x00]); // version=F (container), instance=0
    info_container.extend_from_slice(&4082u16.to_le_bytes()); // RT_InteractiveInfo
    info_container.extend_from_slice(&(info_atom.len() as u32).to_le_bytes()); // length
    info_container.extend_from_slice(&info_atom);

    // ClientData Escher record (0xF011) wrapping InteractiveInfo PPT record
    // POI uses options=15 (0x000F) => version=0xF, instance=0 (container version!)
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&info_container);

    client_data.build()
}

/// Build ClientData record with OEPlaceholderAtom for placeholder shapes
/// Per POI HSLFSimpleShape - placeholders have OEPlaceholderAtom in ClientData
fn build_client_data_with_placeholder(placeholder_type: u8) -> Result<Vec<u8>, PptError> {
    use super::records::RecordBuilder;

    // OEPlaceholderAtom (type 0x0BC3 = 3011)
    // Structure: position (4 bytes), placeholderType (1 byte), size (1 byte), unused (2 bytes)
    let mut oe_atom = RecordBuilder::new(0x00, 0, ppt_record_type::OE_PLACEHOLDER_ATOM);
    oe_atom.write_data(&0u32.to_le_bytes()); // position = 0
    oe_atom.write_data(&[placeholder_type]); // placeholder type (12 = NotesBody per MS-PPT)
    oe_atom.write_data(&[0x00]); // size = full
    oe_atom.write_data(&[0x00, 0x00]); // unused
    let oe_bytes = oe_atom.build()?;

    // ClientData Escher record (0xF011) wrapping OEPlaceholderAtom
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&oe_bytes);

    client_data.build()
}

/// Build shape properties for OPT record
/// Based on Apache POI HSLFTextBox.createSpContainer() defaults
fn build_shape_properties(shape: &UserShapeData) -> Vec<EscherProperty> {
    let mut props = Vec::with_capacity(16);

    // Picture shapes have special handling - BLIP reference only, no fill/line
    if shape.picture_index.is_some() {
        // PROTECTION__LOCKAGAINSTGROUPING (0x007F) = 0x800080 per POI
        props.push(EscherProperty::new(0x007F, 0x0080_0080));
        // BLIP__BLIPTODISPLAY (0x4104) - with isBlipId flag (0x4000 + 0x0104)
        props.push(EscherProperty::new(0x4104, shape.picture_index.unwrap()));
        // No fill for pictures (picture IS the fill)
        props.push(EscherProperty::new(prop_id::NO_FILL_HIT_TEST, 0x0010_0000));
        // No line for pictures
        props.push(EscherProperty::new(prop_id::LINE_STYLE_BOOL, 0x0008_0000));
        return props;
    }

    // Fill properties
    if let Some(fill_color) = shape.fill_color {
        // Fill type (0=solid, 4=shade/gradient)
        if let Some(fill_type) = shape.fill_type {
            props.push(EscherProperty::new(prop_id::FILL_TYPE, fill_type));
        }

        // Fill color
        props.push(EscherProperty::new(prop_id::FILL_COLOR, fill_color));

        // Fill opacity
        if let Some(opacity) = shape.fill_opacity {
            props.push(EscherProperty::new(prop_id::FILL_OPACITY, opacity));
        }

        // Back color (for gradients)
        if let Some(back_color) = shape.fill_back_color {
            props.push(EscherProperty::new(prop_id::FILL_BACK_COLOR, back_color));
        } else {
            props.push(EscherProperty::new(prop_id::FILL_BACK_COLOR, 0x0800_0000)); // scheme bg
        }

        // Gradient angle (for gradient fills)
        if let Some(angle) = shape.fill_angle {
            props.push(EscherProperty::new(prop_id::FILL_ANGLE, angle as u32));
        }

        // Fill boolean: filled = true (0x00010001 per POI)
        props.push(EscherProperty::new(prop_id::NO_FILL_HIT_TEST, 0x0001_0001));
    } else {
        // Default: scheme fill colors with no-fill flag
        props.push(EscherProperty::new(prop_id::FILL_COLOR, 0x0800_0004)); // scheme fill
        props.push(EscherProperty::new(prop_id::FILL_BACK_COLOR, 0x0800_0000));
        props.push(EscherProperty::new(prop_id::NO_FILL_HIT_TEST, 0x0010_0000)); // no fill
    }

    // Line properties (based on POI HSLFSimpleShape)
    if let Some(line_color) = shape.line_color {
        props.push(EscherProperty::new(prop_id::LINE_COLOR, line_color));
        if let Some(width) = shape.line_width {
            props.push(EscherProperty::new(prop_id::LINE_WIDTH, width as u32));
        }
        // Line dash style
        if let Some(dash) = shape.line_dash_style {
            props.push(EscherProperty::new(prop_id::LINE_DASH_STYLE, dash));
        }
        // Line start arrow
        if let Some(arrow) = shape.line_start_arrow {
            props.push(EscherProperty::new(prop_id::LINE_START_ARROW, arrow));
            props.push(EscherProperty::new(prop_id::LINE_START_ARROW_WIDTH, 1)); // Medium
            props.push(EscherProperty::new(prop_id::LINE_START_ARROW_LENGTH, 1)); // Medium
        }
        // Line end arrow
        if let Some(arrow) = shape.line_end_arrow {
            props.push(EscherProperty::new(prop_id::LINE_END_ARROW, arrow));
            props.push(EscherProperty::new(prop_id::LINE_END_ARROW_WIDTH, 1)); // Medium
            props.push(EscherProperty::new(prop_id::LINE_END_ARROW_LENGTH, 1)); // Medium
        }
        // Enable line: 0x180018 = line visible
        props.push(EscherProperty::new(prop_id::LINE_STYLE_BOOL, 0x0018_0018));
    } else {
        // No line: POI uses 0x80000 for no line
        props.push(EscherProperty::new(prop_id::LINE_COLOR, 0x0800_0001)); // scheme line
        props.push(EscherProperty::new(prop_id::LINE_STYLE_BOOL, 0x0008_0000));
    }

    // Shadow properties
    props.push(EscherProperty::new(prop_id::SHADOW_COLOR, 0x0800_0002)); // scheme shadow
    if shape.has_shadow {
        // Enable shadow: offset and boolean
        props.push(EscherProperty::new(prop_id::SHADOW_OFFSET_X, 25400)); // 2pt offset
        props.push(EscherProperty::new(prop_id::SHADOW_OFFSET_Y, 25400));
        props.push(EscherProperty::new(prop_id::SHADOW_BOOL, 0x0003_0003)); // shadow on
    }

    props
}

/// Build ClientTextBox record with plain text content (no formatting)
/// Based on Apache POI EscherTextboxWrapper and HSLFTextShape
/// text_type: 0=Title, 1=Body, 2=Notes, 4=Other
fn build_client_textbox(text: &str, text_type: u32) -> Result<Vec<u8>, PptError> {
    use super::records::{RecordBuilder, record_type as ppt_rt};

    let mut result = Vec::new();
    let mut ppt_content = Vec::new();

    // TextHeaderAtom (type=3999): textType from parameter
    let mut text_header = RecordBuilder::new(0, 0, ppt_rt::TEXT_HEADER_ATOM);
    text_header.write_data(&text_type.to_le_bytes());
    ppt_content.extend_from_slice(&text_header.build()?);

    // TextBytesAtom (type=4008) for ASCII or TextCharsAtom (type=4000) for Unicode
    let is_ascii = text.is_ascii();
    if is_ascii {
        let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_BYTES_ATOM);
        text_atom.write_data(text.as_bytes());
        ppt_content.extend_from_slice(&text_atom.build()?);
    } else {
        let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_CHARS_ATOM);
        for ch in text.encode_utf16() {
            text_atom.write_data(&ch.to_le_bytes());
        }
        ppt_content.extend_from_slice(&text_atom.build()?);
    }

    // StyleTextPropAtom with no formatting
    let char_count = text.chars().count() as u32 + 1;
    let mut style_atom = RecordBuilder::new(0, 0, ppt_rt::STYLE_TEXT_PROP_ATOM);
    style_atom.write_data(&char_count.to_le_bytes()); // para char count
    style_atom.write_data(&0u16.to_le_bytes()); // indent
    style_atom.write_data(&0u32.to_le_bytes()); // para mask
    style_atom.write_data(&char_count.to_le_bytes()); // char count
    style_atom.write_data(&0u32.to_le_bytes()); // char mask
    ppt_content.extend_from_slice(&style_atom.build()?);

    let header = EscherRecordHeader::new(0x0F, 0, 0xF00D, ppt_content.len() as u32);
    result.extend_from_slice(header.as_bytes());
    result.extend_from_slice(&ppt_content);

    Ok(result)
}

/// Build ClientTextBox record with rich text formatting (paragraphs with runs)
/// text_type: 0=Title, 1=Body, 2=Notes, 4=Other
fn build_client_textbox_formatted(
    paragraphs: &[super::text_format::Paragraph],
    text_type: u32,
) -> Result<Vec<u8>, PptError> {
    use super::records::{RecordBuilder, record_type as ppt_rt};
    use super::text_format::TextPropsBuilder;

    let mut result = Vec::new();
    let mut ppt_content = Vec::new();

    // TextHeaderAtom (type=3999): textType from parameter
    let mut text_header = RecordBuilder::new(0, 0, ppt_rt::TEXT_HEADER_ATOM);
    text_header.write_data(&text_type.to_le_bytes());
    ppt_content.extend_from_slice(&text_header.build()?);

    // Build text content from paragraphs
    let mut builder = TextPropsBuilder::new();
    for para in paragraphs {
        builder.add_paragraph(para.clone());
    }

    // Use TextCharsAtom (UTF-16) since we might have unicode
    let text_chars = builder.build_text_chars();
    let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_CHARS_ATOM);
    text_atom.write_data(&text_chars);
    ppt_content.extend_from_slice(&text_atom.build()?);

    // StyleTextPropAtom with full formatting
    let style_data = builder.build_style_text_prop();
    let mut style_atom = RecordBuilder::new(0, 0, ppt_rt::STYLE_TEXT_PROP_ATOM);
    style_atom.write_data(&style_data);
    ppt_content.extend_from_slice(&style_atom.build()?);

    let header = EscherRecordHeader::new(0x0F, 0, 0xF00D, ppt_content.len() as u32);
    result.extend_from_slice(header.as_bytes());
    result.extend_from_slice(&ppt_content);

    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escher_header() {
        let header = EscherHeader::new(0x0F, 0, record_type::DGG_CONTAINER, 100);
        assert_eq!(header.version, 0x0F);
    }
}
