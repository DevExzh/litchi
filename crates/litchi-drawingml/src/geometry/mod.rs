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

/// Caller-selected bounds for a raw custom-geometry subtree.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[must_use]
pub struct Limits {
    pub max_adjust_handles: usize,
    pub max_connection_sites: usize,
    pub max_guides: usize,
    pub max_paths: usize,
    pub max_path_commands: usize,
    pub max_output_bytes: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_adjust_handles: MAX_ADJUST_HANDLES,
        max_connection_sites: MAX_CONNECTION_SITES,
        max_guides: MAX_GEOMETRY_GUIDES,
        max_paths: MAX_GEOMETRY_PATHS,
        max_path_commands: MAX_PATH_COMMANDS,
        max_output_bytes: 1 << 20,
    };

    #[must_use]
    pub const fn new(
        max_adjust_handles: usize,
        max_connection_sites: usize,
        max_guides: usize,
        max_paths: usize,
        max_path_commands: usize,
        max_output_bytes: usize,
    ) -> Self {
        Self {
            max_adjust_handles,
            max_connection_sites,
            max_guides,
            max_paths,
            max_path_commands,
            max_output_bytes,
        }
    }

    pub fn validate(self) -> crate::Result<Self> {
        for (kind, value, maximum) in [
            (
                crate::LimitKind::GeometrySites,
                self.max_adjust_handles,
                MAX_ADJUST_HANDLES,
            ),
            (
                crate::LimitKind::GeometrySites,
                self.max_connection_sites,
                MAX_CONNECTION_SITES,
            ),
            (
                crate::LimitKind::GeometryGuides,
                self.max_guides,
                MAX_GEOMETRY_GUIDES,
            ),
            (
                crate::LimitKind::GeometryPaths,
                self.max_paths,
                MAX_GEOMETRY_PATHS,
            ),
            (
                crate::LimitKind::GeometryPathCommands,
                self.max_path_commands,
                MAX_PATH_COMMANDS,
            ),
            (
                crate::LimitKind::OutputBytes,
                self.max_output_bytes,
                1 << 20,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(crate::Error::invalid_limit(kind, value, maximum));
            }
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

pub(super) use model::{
    MAX_ADJUST_HANDLES, MAX_ANGLE, MAX_CONNECTION_SITES, MAX_COORDINATE, MAX_GEOMETRY_GUIDES,
    MAX_GEOMETRY_PATHS, MAX_GUIDE_NAME_BYTES, MAX_PATH_COMMANDS, MAX_POSITIVE_COORDINATE,
    MIN_ANGLE, MIN_COORDINATE, normalize_xsd_token,
};
