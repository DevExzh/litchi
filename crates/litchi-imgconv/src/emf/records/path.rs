/// EMF Path Operation Records
use zerocopy::{FromBytes, IntoBytes};

// Path operations - these records have no additional data beyond type/size

/// EMR_BEGINPATH
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrBeginPath {
    pub record_type: u32,
    pub record_size: u32,
}

/// EMR_ENDPATH
pub type EmrEndPath = EmrBeginPath;

/// EMR_CLOSEFIGURE
pub type EmrCloseFigure = EmrBeginPath;

/// EMR_FILLPATH
pub type EmrFillPath = EmrBeginPath;

/// EMR_STROKEPATH
pub type EmrStrokePath = EmrBeginPath;

/// EMR_STROKEANDFILLPATH
pub type EmrStrokeAndFillPath = EmrBeginPath;

/// EMR_FLATTENPATH
pub type EmrFlattenPath = EmrBeginPath;

/// EMR_WIDENPATH
pub type EmrWidenPath = EmrBeginPath;

/// EMR_ABORTPATH
pub type EmrAbortPath = EmrBeginPath;

/// Region combine mode for EMR_SELECTCLIPPATH / EMR_EXTSELECTCLIPRGN
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum RegionMode {
    And = 1,  // Intersection
    Or = 2,   // Union
    Xor = 3,  // XOR
    Diff = 4, // Difference (clip region - new region)
    Copy = 5, // Replace with new region
}

/// EMR_SELECTCLIPPATH
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSelectClipPath {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}
