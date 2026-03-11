//! PPT-specific Escher (Office Drawing) record generation.
//!
//! Re-exports shared Escher writer functionality and adds PPT-specific extensions.

use std::io::Write;
use zerocopy::IntoBytes;
use zerocopy_derive::*;

use crate::common::unit::emu_i32_to_ppt_master_i16_round;

/// Error type for PPT operations
pub type PptError = std::io::Error;

// Re-export shared Escher writer functionality
pub use crate::ole::escher::writer::{EscherProperty, EscherRecordHeader, EscherSpData};
pub use crate::ole::escher::writer::{ShapeFlags, record_type, shape_type};

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
    pub const FILL_ANGLE: u16 = 0x0189; // fillAngle for gradients (degrees * 65536)
    pub const FILL_FOCUS: u16 = 0x018A; // fillFocus for gradients (-100 to 100)
    pub const FILL_SHADE_TYPE: u16 = 0x018C; // fillShadeType (0=linear, 1=gamma, etc.)
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
// PPT-Specific Property Values
// =============================================================================

/// PPT-specific property values (extends shared prop_value)
pub mod ppt_prop_value {
    pub use crate::ole::escher::writer::prop_value::*;

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
        fill_color: crate::ole::escher::writer::prop_value::SCHEME_FILL,
        line_color: crate::ole::escher::writer::prop_value::SCHEME_LINE,
        shadow_color: crate::ole::escher::writer::prop_value::SCHEME_SHADOW,
        color_3d: 0x1000_00F7,
    };
}

// =============================================================================
// Header Version Constants
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
            cspid_cur: ppt_prop_value::RESERVED_CSPID_CUR,
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
// PPT-Specific EscherSpData Extensions
// =============================================================================

impl EscherSpData {
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

// Re-use EscherProperty from shared module

/// Default drawing group properties (8 properties = 48 bytes)
pub const DGG_DEFAULT_PROPERTIES: [EscherProperty; 8] = [
    EscherProperty::new(prop_id::FILL_COLOR, ppt_prop_value::SCHEME_FILL),
    EscherProperty::new(prop_id::FILL_BACK_COLOR, ppt_prop_value::SCHEME_FILL_BACK),
    EscherProperty::new(prop_id::FILL_BLIP, 0),
    EscherProperty::new(
        prop_id::NO_FILL_HIT_TEST,
        ppt_prop_value::LINE_STYLE_DEFAULT,
    ),
    EscherProperty::new(prop_id::LINE_COLOR, ppt_prop_value::SCHEME_LINE),
    EscherProperty::new(prop_id::LINE_BLIP, 0),
    EscherProperty::new(prop_id::SHAPE_BOOL, ppt_prop_value::SHAPE_BOOL_DEFAULT),
    EscherProperty::new(prop_id::SHADOW_COLOR, ppt_prop_value::SCHEME_SHADOW),
];

/// Background shape properties (8 properties = 48 bytes)
pub const BG_SHAPE_PROPERTIES: [EscherProperty; 8] = [
    EscherProperty::new(prop_id::FILL_COLOR, ppt_prop_value::BG_FILL_COLOR),
    EscherProperty::new(prop_id::FILL_BACK_COLOR, ppt_prop_value::BG_FILL_BACK_COLOR),
    EscherProperty::new(prop_id::FILL_RECT_RIGHT, ppt_prop_value::SLIDE_WIDTH_EMU),
    EscherProperty::new(prop_id::FILL_RECT_BOTTOM, ppt_prop_value::SLIDE_HEIGHT_EMU),
    EscherProperty::new(prop_id::NO_FILL_HIT_TEST, ppt_prop_value::NO_FILL_HIT_TEST),
    EscherProperty::new(prop_id::LINE_STYLE_BOOL, ppt_prop_value::NO_LINE_DRAW_DASH),
    EscherProperty::new(prop_id::BW_MODE, ppt_prop_value::BW_MODE_AUTO),
    EscherProperty::new(prop_id::BACKGROUND_SHAPE, ppt_prop_value::BACKGROUND_SHAPE),
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
    let spid_max = if drawing_count == 1 && master_shapes == ppt_prop_value::POI_MASTER_SHAPE_COUNT
    {
        ppt_prop_value::POI_SPID_MAX
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
    let spid_max = if drawing_count == 1 && master_shapes == ppt_prop_value::POI_MASTER_SHAPE_COUNT
    {
        ppt_prop_value::POI_SPID_MAX
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
    /// Picture BLIP index (for picture frames)
    pub picture_index: Option<u32>,
    /// Animation info for this shape
    pub animation_info: Option<crate::ole::ppt::animation::AnimationInfo>,
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
            line_color: None,
            line_width: None,
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
            picture_index: None,
            animation_info: None,
            shadow_color: None,
            shadow_offset_x: None,
            shadow_offset_y: None,
            shadow_opacity: None,
            shadow_type: None,
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
    let x1 = emu_i32_to_ppt_master_i16_round(shape.x);
    let y1 = emu_i32_to_ppt_master_i16_round(shape.y);
    let x2 = emu_i32_to_ppt_master_i16_round(shape.x + shape.width);
    let y2 = emu_i32_to_ppt_master_i16_round(shape.y + shape.height);
    // Short record format: 8 bytes (4 shorts)
    anchor.add_data(&y1.to_le_bytes()); // flag/top
    anchor.add_data(&x1.to_le_bytes()); // col1/left
    anchor.add_data(&x2.to_le_bytes()); // dx1/right
    anchor.add_data(&y2.to_le_bytes()); // row1/bottom
    container.add_data(&anchor.build()?);

    // ClientData with animation, placeholders, or hyperlinks
    // MUST come BEFORE ClientTextbox per POI (addChildBefore(clientData, EscherTextboxRecord.RECORD_ID))
    if let Some(ref animation_info) = shape.animation_info {
        // Animation takes priority - write AnimationInfo to ClientData
        let client_data = build_client_data_with_animation(animation_info)?;
        container.add_data(&client_data);
    } else if let Some(placeholder_type) = shape.placeholder_type {
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

/// Build ClientData record with AnimationInfo.
///
/// Per LibreOffice reference files, animation sounds only need AnimationInfo in ClientData.
/// InteractiveInfo with action=6 (MEDIA) is for movie/media objects, NOT animation sounds.
/// The reference `sound.ppt` has AnimationInfo WITHOUT InteractiveInfo in its ClientData.
fn build_client_data_with_animation(
    animation_info: &crate::ole::ppt::animation::AnimationInfo,
) -> Result<Vec<u8>, PptError> {
    use crate::ole::ppt::animation::writer::write_animation_info;

    // Write AnimationInfo container (contains AnimationInfoAtom with soundRef)
    let (animation_bytes, _sound_ref) = write_animation_info(animation_info);

    // ClientData Escher record (0xF011) wrapping AnimationInfo only
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&animation_bytes);
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
    if let Some(picture_index) = shape.picture_index {
        // PROTECTION__LOCKAGAINSTGROUPING (0x007F) = 0x800080 per POI
        props.push(EscherProperty::new(0x007F, 0x0080_0080));
        // BLIP__BLIPTODISPLAY (0x4104) - with isBlipId flag (0x4000 + 0x0104)
        props.push(EscherProperty::new(0x4104, picture_index));
        // No fill for pictures (picture IS the fill)
        props.push(EscherProperty::new(prop_id::NO_FILL_HIT_TEST, 0x0010_0000));
        // No line for pictures
        props.push(EscherProperty::new(prop_id::LINE_STYLE_BOOL, 0x0008_0000));
        return props;
    }

    // Fill properties
    if let Some(fill_color) = shape.fill_color {
        // Fill type (0=solid, 4=shade/gradient) - MUST be first
        if let Some(fill_type) = shape.fill_type {
            props.push(EscherProperty::new(prop_id::FILL_TYPE, fill_type));
        }

        // Fill color
        props.push(EscherProperty::new(prop_id::FILL_COLOR, fill_color));

        // Back color (for gradients) - before angle
        if let Some(back_color) = shape.fill_back_color {
            props.push(EscherProperty::new(prop_id::FILL_BACK_COLOR, back_color));
        }

        // Gradient angle (for gradient fills) - MUST be before opacity
        if let Some(angle) = shape.fill_angle {
            props.push(EscherProperty::new(prop_id::FILL_ANGLE, angle as u32));
        }

        // Fill opacity (after angle)
        if let Some(opacity) = shape.fill_opacity {
            props.push(EscherProperty::new(prop_id::FILL_OPACITY, opacity));
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
    if shape.has_shadow {
        // Shadow type
        if let Some(shadow_type) = shape.shadow_type {
            props.push(EscherProperty::new(prop_id::SHADOW_TYPE, shadow_type));
        }

        // Shadow color
        let shadow_color = shape.shadow_color.unwrap_or(0x0800_0002); // default: scheme shadow
        props.push(EscherProperty::new(prop_id::SHADOW_COLOR, shadow_color));

        // Shadow offsets
        let offset_x = shape.shadow_offset_x.unwrap_or(25400) as u32; // default: 2pt
        let offset_y = shape.shadow_offset_y.unwrap_or(25400) as u32; // default: 2pt
        props.push(EscherProperty::new(prop_id::SHADOW_OFFSET_X, offset_x));
        props.push(EscherProperty::new(prop_id::SHADOW_OFFSET_Y, offset_y));

        // Shadow opacity
        if let Some(opacity) = shape.shadow_opacity {
            props.push(EscherProperty::new(prop_id::SHADOW_OPACITY, opacity));
        }

        // Enable shadow boolean
        props.push(EscherProperty::new(prop_id::SHADOW_BOOL, 0x0003_0003)); // shadow on
    } else {
        // No shadow - still set scheme color for consistency
        props.push(EscherProperty::new(prop_id::SHADOW_COLOR, 0x0800_0002));
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
    use super::super::shapes::shape_type;
    use super::*;
    use crate::ole::ppt::writer::text_format::{Paragraph, TextRun};

    #[test]
    fn test_escher_header() {
        let header = EscherHeader::new(0x0F, 5, record_type::DG_CONTAINER, 100);
        assert_eq!(header.version, 0x0F);
        assert_eq!(header.instance, 5);
        assert_eq!(header.record_type, record_type::DG_CONTAINER);
        assert_eq!(header.length, 100);
    }

    #[test]
    fn test_escher_record_header() {
        let header = EscherRecordHeader::new(0x0F, 0, record_type::DG_CONTAINER, 100);
        // Fields are ver_inst, rec_type, length - copy to locals to avoid unaligned access
        let ver_inst = header.ver_inst;
        let rec_type = header.rec_type;
        let length = header.length;
        assert_eq!(ver_inst, 0x000F); // version 0x0F in low 4 bits
        assert_eq!(rec_type, record_type::DG_CONTAINER);
        assert_eq!(length, 100);
    }

    #[test]
    fn test_escher_record_header_as_bytes() {
        let header = EscherRecordHeader::new(0x0F, 1, record_type::SP_CONTAINER, 50);
        let bytes = header.as_bytes();
        assert_eq!(bytes.len(), 8);

        // Verify byte content directly
        // ver_inst = (0x0F & 0x0F) | ((1 & 0x0FFF) << 4) = 0x000F | 0x0010 = 0x001F
        let ver_inst = u16::from_le_bytes([bytes[0], bytes[1]]);
        assert_eq!(ver_inst, 0x001F);

        let rec_type = u16::from_le_bytes([bytes[2], bytes[3]]);
        assert_eq!(rec_type, record_type::SP_CONTAINER);

        let length = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
        assert_eq!(length, 50);
    }

    #[test]
    fn test_escher_builder_basic() {
        let mut builder =
            EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);
        builder.add_data(&[1, 2, 3, 4]);

        let result = builder.build();
        assert!(result.is_ok());
        let data = result.unwrap();
        assert!(data.len() >= 12); // 8 bytes header + 4 bytes data
    }

    #[test]
    fn test_escher_builder_empty() {
        let builder = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);
        let result = builder.build();
        assert!(result.is_ok());
        let data = result.unwrap();
        assert_eq!(data.len(), 8); // Just header
    }

    #[test]
    fn test_escher_dg_data() {
        let dg_data = EscherDgData::new(10, 1);
        let bytes = dg_data.as_bytes();
        assert_eq!(bytes.len(), 8);

        // Verify byte content directly - fields are csp and spid_cur
        let csp = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
        let spid_cur = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
        assert_eq!(csp, 10);
        // spid_cur = (drawing_id << 10) + shape_count = (1 << 10) + 10 = 1024 + 10 = 1034
        assert_eq!(spid_cur, 1034);
    }

    #[test]
    fn test_escher_sp_data() {
        let sp_data =
            EscherSpData::with_flags(0x0401, ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT);
        let bytes = sp_data.as_bytes();
        assert_eq!(bytes.len(), 8);

        // Verify byte content directly - field is spid not sp_id
        let spid = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
        let flags = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
        assert_eq!(spid, 0x0401);
        assert_eq!(
            flags,
            (ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT).bits()
        );
    }

    #[test]
    fn test_escher_spgr_data() {
        // Construct using struct literal since there's no new() method
        let data = EscherSpgrData {
            left: 0,
            top: 0,
            right: 1000,
            bottom: 1000,
        };
        let bytes = data.as_bytes();
        assert_eq!(bytes.len(), 16);

        // Verify byte content directly
        let left = i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
        let top = i32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
        let right = i32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]);
        let bottom = i32::from_le_bytes([bytes[12], bytes[13], bytes[14], bytes[15]]);
        assert_eq!(left, 0);
        assert_eq!(top, 0);
        assert_eq!(right, 1000);
        assert_eq!(bottom, 1000);
    }

    #[test]
    fn test_escher_property() {
        let prop = EscherProperty::new(0x0181, 0x00FF0000);
        let bytes = prop.as_bytes();
        assert_eq!(bytes.len(), 6);

        // Verify byte content directly
        let prop_id = u16::from_le_bytes([bytes[0], bytes[1]]);
        let value = u32::from_le_bytes([bytes[2], bytes[3], bytes[4], bytes[5]]);
        assert_eq!(prop_id, 0x0181);
        assert_eq!(value, 0x00FF0000);
    }

    #[test]
    fn test_shape_flags() {
        let flags = ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT;
        let value: u32 = flags.bits();
        assert_eq!(value, 0x0A00);

        let flags2 = ShapeFlags::FLIP_H | ShapeFlags::FLIP_V;
        let value2: u32 = flags2.bits();
        assert_eq!(value2, 0x00C0);
    }

    #[test]
    fn test_user_shape_data_default() {
        let shape = UserShapeData::default();
        assert_eq!(shape.shape_type, shape_type::RECTANGLE);
        assert_eq!(shape.x, 0);
        assert_eq!(shape.y, 0);
        assert_eq!(shape.width, 914400); // 1 inch in EMUs
        assert_eq!(shape.height, 914400);
        assert!(!shape.has_shadow);
        assert!(!shape.flip_h);
        assert!(!shape.flip_v);
    }

    #[test]
    fn test_create_dgg_container() {
        let container = create_dgg_container(5, &[3, 4, 5]);
        assert!(container.is_ok());
        let data = container.unwrap();
        assert!(!data.is_empty());
        assert!(data.len() > 20);
    }

    #[test]
    fn test_create_dgg_container_empty() {
        let container = create_dgg_container(0, &[]);
        assert!(container.is_ok());
        let data = container.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_create_dgg_container_many_slides() {
        let slide_counts = vec![1u32; 100];
        let container = create_dgg_container(5, &slide_counts);
        assert!(container.is_ok());
    }

    #[test]
    fn test_create_dg_container_with_shapes_empty() {
        let container = create_dg_container_with_shapes(1, &[]);
        assert!(container.is_ok());
        let data = container.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_create_dg_container_with_shapes_single() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            x: 100000,
            y: 100000,
            width: 500000,
            height: 300000,
            text: Some("Test".to_string()),
            ..Default::default()
        };
        let container = create_dg_container_with_shapes(1, &[shape]);
        assert!(container.is_ok());
        let data = container.unwrap();
        assert!(!data.is_empty());
        assert!(data.len() > 50);
    }

    #[test]
    fn test_create_dg_container_with_shapes_multiple() {
        let shapes = vec![
            UserShapeData {
                shape_type: shape_type::RECTANGLE,
                x: 0,
                y: 0,
                width: 100000,
                height: 100000,
                ..Default::default()
            },
            UserShapeData {
                shape_type: shape_type::ELLIPSE,
                x: 200000,
                y: 200000,
                width: 100000,
                height: 100000,
                ..Default::default()
            },
            UserShapeData {
                shape_type: shape_type::LINE,
                x: 0,
                y: 300000,
                width: 300000,
                height: 0,
                ..Default::default()
            },
        ];
        let container = create_dg_container_with_shapes(1, &shapes);
        assert!(container.is_ok());
    }

    #[test]
    fn test_create_shape_container() {
        let container = create_shape_container(
            0x0401,
            shape_type::RECTANGLE,
            100000,
            100000,
            500000,
            300000,
        );
        assert!(container.is_ok());
        let data = container.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_shape_properties_rectangle() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            fill_color: Some(0x00FF0000),
            line_color: Some(0x00000000),
            ..Default::default()
        };
        let props = build_shape_properties(&shape);
        assert!(!props.is_empty());
        // Should have fill and line properties
        assert!(props.len() >= 4);
    }

    #[test]
    fn test_build_shape_properties_no_fill() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            fill_color: None,
            ..Default::default()
        };
        let props = build_shape_properties(&shape);
        // Should have scheme fill with no-fill flag
        let has_no_fill = props.iter().any(|p| p.prop_id == prop_id::NO_FILL_HIT_TEST);
        assert!(has_no_fill);
    }

    #[test]
    fn test_build_shape_properties_with_shadow() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            has_shadow: true,
            shadow_color: Some(0x00808080),
            ..Default::default()
        };
        let props = build_shape_properties(&shape);
        // Should have shadow properties
        let has_shadow_prop = props.iter().any(|p| p.prop_id == prop_id::SHADOW_BOOL);
        assert!(has_shadow_prop);
    }

    #[test]
    fn test_build_shape_properties_picture() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            picture_index: Some(1),
            ..Default::default()
        };
        let props = build_shape_properties(&shape);
        // Should have BLIP property
        let has_blip = props.iter().any(|p| p.prop_id == 0x4104);
        assert!(has_blip);
    }

    #[test]
    fn test_build_shape_properties_with_arrows() {
        let shape = UserShapeData {
            shape_type: shape_type::LINE,
            line_color: Some(0x00000000),
            line_end_arrow: Some(1), // Triangle arrow
            ..Default::default()
        };
        let props = build_shape_properties(&shape);
        // Should have arrow properties
        let has_arrow = props.iter().any(|p| p.prop_id == prop_id::LINE_END_ARROW);
        assert!(has_arrow);
    }

    #[test]
    fn test_build_shape_properties_gradient_fill() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            fill_color: Some(0x00FF0000),
            fill_type: Some(4), // Shade/gradient
            fill_back_color: Some(0x0000FF00),
            fill_angle: Some(0),
            ..Default::default()
        };
        let props = build_shape_properties(&shape);
        // Should have fill type and back color
        let has_fill_type = props.iter().any(|p| p.prop_id == prop_id::FILL_TYPE);
        let has_back_color = props.iter().any(|p| p.prop_id == prop_id::FILL_BACK_COLOR);
        assert!(has_fill_type);
        assert!(has_back_color);
    }

    #[test]
    fn test_client_textbox_plain_ascii() {
        let textbox = build_client_textbox("Hello World", 4);
        assert!(textbox.is_ok());
        let data = textbox.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_client_textbox_unicode() {
        let textbox = build_client_textbox("Hello 世界 🌍", 4);
        assert!(textbox.is_ok());
        let data = textbox.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_client_textbox_empty() {
        let textbox = build_client_textbox("", 4);
        assert!(textbox.is_ok());
    }

    #[test]
    fn test_client_textbox_formatted() {
        let paragraphs = vec![
            Paragraph::new("First paragraph"),
            Paragraph::with_runs(vec![
                TextRun::new("Bold text").bold(),
                TextRun::new(" and "),
                TextRun::new("italic").italic(),
            ]),
        ];
        let textbox = build_client_textbox_formatted(&paragraphs, 1);
        assert!(textbox.is_ok());
        let data = textbox.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_client_data_with_hyperlink() {
        let client_data = build_client_data_with_hyperlink(1, 4, 0, 8);
        assert!(client_data.is_ok());
        let data = client_data.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_client_data_with_placeholder() {
        let client_data = build_client_data_with_placeholder(6); // NotesBody
        assert!(client_data.is_ok());
        let data = client_data.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_shape_type_constants() {
        // Verify all shape type constants are correctly defined
        assert_eq!(shape_type::NOT_PRIMITIVE, 0);
        assert_eq!(shape_type::RECTANGLE, 1);
        assert_eq!(shape_type::ROUND_RECTANGLE, 2);
        assert_eq!(shape_type::ELLIPSE, 3);
        assert_eq!(shape_type::DIAMOND, 4);
        assert_eq!(shape_type::ISOCELES_TRIANGLE, 5);
        assert_eq!(shape_type::RIGHT_TRIANGLE, 6);
        assert_eq!(shape_type::PARALLELOGRAM, 7);
        assert_eq!(shape_type::TRAPEZOID, 8);
        assert_eq!(shape_type::HEXAGON, 9);
        assert_eq!(shape_type::OCTAGON, 10);
        assert_eq!(shape_type::PLUS, 11);
        assert_eq!(shape_type::STAR, 12);
        assert_eq!(shape_type::ARROW, 13);
        assert_eq!(shape_type::THICK_ARROW, 14);
        assert_eq!(shape_type::LINE, 20);
        assert_eq!(shape_type::TEXT_BOX, 202);
    }

    #[test]
    fn test_prop_id_constants() {
        // Verify property ID constants
        assert_eq!(prop_id::FILL_TYPE, 0x0180);
        assert_eq!(prop_id::FILL_COLOR, 0x0181);
        assert_eq!(prop_id::FILL_OPACITY, 0x0182);
        assert_eq!(prop_id::FILL_BACK_COLOR, 0x0183);
        assert_eq!(prop_id::LINE_COLOR, 0x01C0);
        assert_eq!(prop_id::LINE_WIDTH, 0x01CB);
        assert_eq!(prop_id::LINE_START_ARROW, 0x01D0);
        assert_eq!(prop_id::LINE_END_ARROW, 0x01D1);
        assert_eq!(prop_id::SHADOW_TYPE, 0x0200);
        assert_eq!(prop_id::SHADOW_COLOR, 0x0201);
        assert_eq!(prop_id::NO_FILL_HIT_TEST, 0x01BF);
        assert_eq!(prop_id::LINE_STYLE_BOOL, 0x01FF);
    }

    #[test]
    fn test_record_type_constants() {
        assert_eq!(record_type::DGG_CONTAINER, 0xF000);
        assert_eq!(record_type::DGG, 0xF006);
        assert_eq!(record_type::DG_CONTAINER, 0xF002);
        assert_eq!(record_type::DG, 0xF008);
        assert_eq!(record_type::SPGR_CONTAINER, 0xF003);
        assert_eq!(record_type::SP_CONTAINER, 0xF004);
        assert_eq!(record_type::SP, 0xF00A);
        assert_eq!(record_type::SPGR, 0xF009);
        assert_eq!(record_type::OPT, 0xF00B);
        assert_eq!(record_type::CLIENT_ANCHOR, 0xF010);
        assert_eq!(record_type::CLIENT_DATA, 0xF011);
    }

    #[test]
    fn test_header_version_constants() {
        assert_eq!(header_version::CONTAINER, 0x0F);
        // DGG doesn't exist in header_version - the DGG record type is different from header version
        assert_eq!(header_version::DG, 0x00);
        assert_eq!(header_version::SPGR, 0x01);
        assert_eq!(header_version::SP, 0x02);
    }

    #[test]
    fn test_escher_header_as_bytes() {
        let header = EscherHeader::new(0x0F, 5, record_type::DG_CONTAINER, 100);
        let mut buf = Vec::new();
        header.write(&mut buf).unwrap();
        assert_eq!(buf.len(), 8);

        // Verify byte content directly
        // version in low 4 bits, instance in high 12 bits
        let ver_inst = u16::from_le_bytes([buf[0], buf[1]]);
        assert_eq!(ver_inst, 0x005F); // 0x0F | (5 << 4) = 0x0F | 0x50 = 0x5F

        let rec_type = u16::from_le_bytes([buf[2], buf[3]]);
        assert_eq!(rec_type, record_type::DG_CONTAINER);

        let length = u32::from_le_bytes([buf[4], buf[5], buf[6], buf[7]]);
        assert_eq!(length, 100);
    }

    #[test]
    fn test_create_dg_container_with_flip_flags() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            x: 100000,
            y: 100000,
            width: 500000,
            height: 300000,
            flip_h: true,
            flip_v: true,
            ..Default::default()
        };
        let container = create_dg_container_with_shapes(1, &[shape]);
        assert!(container.is_ok());
    }

    #[test]
    fn test_create_dg_container_with_dash_style() {
        let shape = UserShapeData {
            shape_type: shape_type::LINE,
            line_color: Some(0x00000000),
            line_dash_style: Some(1), // Dash
            ..Default::default()
        };
        let container = create_dg_container_with_shapes(1, &[shape]);
        assert!(container.is_ok());
    }

    #[test]
    fn test_shape_properties_with_line_width() {
        let shape = UserShapeData {
            shape_type: shape_type::RECTANGLE,
            line_color: Some(0x00000000),
            line_width: Some(25400), // 2pt in EMUs
            ..Default::default()
        };
        let props = build_shape_properties(&shape);
        let has_width = props.iter().any(|p| p.prop_id == prop_id::LINE_WIDTH);
        assert!(has_width);
    }

    #[test]
    fn test_multiple_paragraphs_textbox() {
        let paragraphs = vec![
            Paragraph::new("First paragraph with some text"),
            Paragraph::new("Second paragraph").center(),
            Paragraph::new("Third paragraph").right(),
        ];
        let textbox = build_client_textbox_formatted(&paragraphs, 1);
        assert!(textbox.is_ok());
        let data = textbox.unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_text_with_multiple_runs() {
        let runs = vec![
            TextRun::new("Normal "),
            TextRun::new("bold ").bold(),
            TextRun::new("italic").italic(),
            TextRun::new(" "),
            TextRun::new("underline").underline(),
        ];
        let para = Paragraph::with_runs(runs);
        let textbox = build_client_textbox_formatted(&[para], 4);
        assert!(textbox.is_ok());
    }
}
