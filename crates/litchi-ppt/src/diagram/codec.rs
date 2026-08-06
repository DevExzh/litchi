//! Composition of the existing MS-PPT and MS-ODRAW parsers.

use crate::package::{Error, Result};
use crate::records::Record;

use super::model::{Diagram, Inventory, Limits, ShapeRef};
use super::validation::{collect_associated, diagram_builds, validate_shape_ids};

/// Parse a native diagram inventory from a BuildList and its PPDrawing body.
pub fn parse(build_list: &Record, drawing: &[u8]) -> Result<Inventory<'_>> {
    parse_with_limits(build_list, drawing, Limits::default())
}

/// Parse a native diagram inventory with explicit resource ceilings.
pub fn parse_with_limits(
    build_list: &Record,
    drawing: &[u8],
    limits: Limits,
) -> Result<Inventory<'_>> {
    let builds = diagram_builds(build_list, limits)?;
    let drawing = crate::odraw::parse_drawing(drawing)?;
    validate_shape_ids(&drawing)?;

    let mut diagrams = Vec::new();
    diagrams
        .try_reserve(builds.len())
        .map_err(|_| Error::Corrupted("unable to allocate diagram inventory".to_string()))?;
    for build in builds {
        let root = ShapeRef::new(build.shape_id());
        let root_shape = root.resolve(&drawing).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "diagram build ({}, {}) references missing OfficeArt shape {}",
                build.build_id(),
                build.shape_id(),
                root.id()
            ))
        })?;
        let (shapes, payloads) = collect_associated(root_shape, limits)?;
        diagrams.push(Diagram::new(build, shapes, payloads));
    }
    Ok(Inventory::new(drawing, diagrams))
}

/// Parse one exact serialized BuildList record and a PPDrawing body.
pub fn parse_bytes(build_list: &[u8], drawing: &[u8]) -> Result<Inventory<'_>> {
    parse_bytes_with_limits(build_list, drawing, Limits::default())
}

/// Parse one exact serialized BuildList record with explicit limits.
pub fn parse_bytes_with_limits(
    build_list: &[u8],
    drawing: &[u8],
    limits: Limits,
) -> Result<Inventory<'_>> {
    let (record, consumed) = Record::parse_strict(build_list, 0)?;
    if consumed != build_list.len() {
        return Err(Error::Corrupted("BuildList has trailing bytes".to_string()));
    }
    parse_with_limits(&record, drawing, limits)
}
