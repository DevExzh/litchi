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

    // Fill style
    pub const FILL_TYPE: u16 = 0x0180;
    pub const FILL_COLOR: u16 = 0x0181;
    pub const FILL_OPACITY: u16 = 0x0182;
    pub const FILL_BACK_COLOR: u16 = 0x0183;
    pub const FILL_BACK_OPACITY: u16 = 0x0184;
    pub const FILL_BLIP: u16 = 0x4186;
    pub const FILL_RECT_RIGHT: u16 = 0x0193;
    pub const FILL_RECT_BOTTOM: u16 = 0x0194;
    pub const NO_FILL_HIT_TEST: u16 = 0x01BF;

    // Line style
    pub const LINE_COLOR: u16 = 0x01C0;
    pub const LINE_OPACITY: u16 = 0x01C1;
    pub const LINE_BACK_COLOR: u16 = 0x01C2;
    pub const LINE_WIDTH: u16 = 0x01CB;
    pub const LINE_STYLE: u16 = 0x01CD;
    pub const LINE_DASH_STYLE: u16 = 0x01CE;
    pub const LINE_BLIP: u16 = 0x41C5;
    pub const LINE_STYLE_BOOL: u16 = 0x01FF;

    // Shadow style
    pub const SHADOW_TYPE: u16 = 0x0200;
    pub const SHADOW_COLOR: u16 = 0x0201;
    pub const SHADOW_OPACITY: u16 = 0x0204;

    // Shape
    pub const BW_MODE: u16 = 0x0304;
    pub const SHAPE_BOOL: u16 = 0x01FF;
    pub const BACKGROUND_SHAPE: u16 = 0x033F;
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
/// * `drawing_count` - Number of drawings (master + slides)
/// * `master_shapes` - Number of shapes in the master (6 for POI template)
/// * `slide_shapes` - Number of shapes per slide (typically 2: group + background)
pub fn create_dgg_container(
    drawing_count: u32,
    master_shapes: u32,
    slide_shapes: u32,
) -> Result<Vec<u8>, PptError> {
    let mut container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::DGG_CONTAINER);

    // Calculate total shapes: master has master_shapes, each slide has slide_shapes
    let slide_count = drawing_count.saturating_sub(1);
    let csp_saved = master_shapes + slide_count * slide_shapes;

    // POI uses cidcl=4 (3 clusters) even for 1 drawing
    let num_clusters = std::cmp::max(3, drawing_count as usize);
    let cidcl = (num_clusters + 1) as u32;

    // spidMax: POI's empty.ppt uses exactly 3076 = 3*1024 + 4
    let spid_max = if drawing_count == 1 && master_shapes == prop_value::POI_MASTER_SHAPE_COUNT {
        prop_value::POI_SPID_MAX
    } else if drawing_count == 1 {
        3 * 1024 + csp_saved
    } else {
        drawing_count * 1024 + slide_shapes
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
    for dg_id in 1..=drawing_count {
        let cspid_cur = if dg_id == 1 {
            master_shapes + 1
        } else {
            slide_shapes + 1
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

    // Background shape container
    let mut bg_sp_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    // Background EscherSp
    let bg_spid = group_spid + 1;
    let mut bg_sp = EscherBuilder::new(header_version::SP, shape_type::RECTANGLE, record_type::SP);
    bg_sp.add_data(EscherSpData::background(bg_spid).as_bytes());
    bg_sp_container.add_data(&bg_sp.build()?);

    // Background EscherOpt properties
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        BG_SHAPE_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &BG_SHAPE_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    bg_sp_container.add_data(&opt.build()?);

    // ClientAnchor (8 bytes zeros)
    let mut client_anchor =
        EscherBuilder::new(header_version::SIMPLE, 0, record_type::CLIENT_ANCHOR);
    client_anchor.add_data(ClientAnchor::ZERO.as_bytes());
    bg_sp_container.add_data(&client_anchor.build()?);

    // ClientData with nested OEPlaceholderAtom
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    // PPT record header for OEPlaceholderAtom
    let placeholder_header = EscherRecordHeader::new(0, 0, ppt_record_type::OE_PLACEHOLDER_ATOM, 8);
    client_data.add_data(placeholder_header.as_bytes());
    client_data.add_data(OEPlaceholderAtom::BACKGROUND.as_bytes());
    bg_sp_container.add_data(&client_data.build()?);

    spgr_container.add_data(&bg_sp_container.build()?);
    container.add_data(&spgr_container.build()?);

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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escher_header() {
        let header = EscherHeader::new(0x0F, 0, record_type::DGG_CONTAINER, 100);
        assert_eq!(header.version, 0x0F);
    }
}
