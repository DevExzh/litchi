//! Custom `DrawingML` geometry (`a:custGeom`) and its XML codecs.
//!
//! This module owns the format-neutral ECMA-376 `CT_CustomGeometry2D` model:
//! guide formulas, adjustment handles, connection sites, text rectangles,
//! path primitives, and bounded validation. DOCX, PPTX, XLSX, and XLSB keep
//! their surrounding shape, anchor, package, and relationship semantics in
//! their own crates and consume this module for the shared geometry subtree.
//!
//! Specification anchors: [MS-ODRAWXML] `CT_CustomGeometry2D`, `coordinates`,
//! coordinate simple types, and path primitives (structures 2.13, 2.18, 2.20; schema
//! appendix 5.1); [MS-OI29500] OPC part/relationship boundaries (Part 1
//! sections 12–13). This crate never resolves package relationships.

mod formula;
mod model;
pub mod reader;
mod validation;
pub mod writer;

#[cfg(test)]
mod tests;

pub use formula::Formula;
pub use model::{
    AdjustHandle, AdjustValue, ConnectionSite, CustomGeometry, Guide, Path, PathCommand,
    PathFillMode, Point, PolarAdjustHandle, Rectangle, XyAdjustHandle,
};
pub use validation::{validate_custom_geometry, validate_parsed_custom_geometry};

pub(super) use model::{
    MAX_ADJUST_HANDLES, MAX_ANGLE, MAX_CONNECTION_SITES, MAX_COORDINATE, MAX_GEOMETRY_GUIDES,
    MAX_GEOMETRY_PATHS, MAX_GUIDE_NAME_BYTES, MAX_PATH_COMMANDS, MAX_POSITIVE_COORDINATE,
    MIN_ANGLE, MIN_COORDINATE, normalize_xsd_token,
};
