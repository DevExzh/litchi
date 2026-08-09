//! Composition of the existing MS-PPT and MS-ODRAW parsers.

use crate::package::{Error, Result};
use crate::records::Record;

use super::model::{Build, Diagram, EditLimits, Inventory, Limits, LocatedBuild, ShapeRef};
use super::validation::{
    collect_associated, diagram_builds, validate_edit_build_list, validate_shape_ids,
};
use crate::animation::diagram_build::{self, BuildType};

/// Fixed offsets inside one validated `DiagramBuildContainer` record.
///
/// The transaction edits only these two fields and copies every other source
/// byte unchanged.  The constants mirror the fixed records in [MS-PPT]
/// §§2.8.13–2.8.14.
const SHAPE_ID_OFFSET: usize = 24;
const MODE_OFFSET: usize = 40;

/// Parse a native diagram inventory from a `BuildList` and its `PPDrawing` body.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse<'data>(build_list: &Record, drawing: &'data [u8]) -> Result<Inventory<'data>> {
    parse_with_limits(build_list, drawing, Limits::default())
}

/// Parse a native diagram inventory with explicit resource ceilings.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_with_limits<'data>(
    build_list: &Record,
    drawing: &'data [u8],
    limits: Limits,
) -> Result<Inventory<'data>> {
    let builds = diagram_builds(build_list, limits)?;
    let parsed_drawing = crate::odraw::parse_drawing(drawing)?;
    validate_shape_ids(&parsed_drawing)?;

    let mut diagrams = Vec::new();
    diagrams
        .try_reserve(builds.len())
        .map_err(|_err| Error::Corrupted("unable to allocate diagram inventory".to_string()))?;
    for build in builds {
        let root = ShapeRef::new(build.shape_id());
        let root_shape = root.resolve(&parsed_drawing).ok_or_else(|| {
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
    Ok(Inventory::new(parsed_drawing, diagrams))
}

/// Parse one exact serialized `BuildList` record and a `PPDrawing` body.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_bytes<'data>(build_list: &[u8], drawing: &'data [u8]) -> Result<Inventory<'data>> {
    parse_bytes_with_limits(build_list, drawing, Limits::default())
}

/// Parse one exact serialized `BuildList` record with explicit limits.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_bytes_with_limits<'data>(
    build_list: &[u8],
    drawing: &'data [u8],
    limits: Limits,
) -> Result<Inventory<'data>> {
    let (record, consumed) = Record::parse_strict(build_list, 0)?;
    if consumed != build_list.len() {
        return Err(Error::Corrupted("BuildList has trailing bytes".to_string()));
    }
    parse_with_limits(&record, drawing, limits)
}

/// Parse typed diagram builds and retain their source offsets without
/// rebuilding the surrounding `BuildList`.
pub(super) fn parse_entries(
    bytes: &[u8],
    shape_ids: &[u32],
    limits: EditLimits,
) -> Result<Vec<LocatedBuild>> {
    let (record, consumed) = Record::parse_strict(bytes, 0)?;
    if consumed != bytes.len() {
        return Err(Error::Corrupted("BuildList has trailing bytes".to_string()));
    }
    validate_edit_build_list(&record, limits)?;

    let mut entries = Vec::new();
    entries
        .try_reserve(record.children.len().min(limits.max_diagrams))
        .map_err(|_err| {
            Error::Corrupted("unable to allocate diagram transaction state".to_string())
        })?;
    let mut offset = 8usize;
    for child in &record.children {
        let child_len = 8usize
            .checked_add(child.data.len())
            .ok_or_else(|| Error::Corrupted("BuildList child length overflow".to_string()))?;
        if child.record_type == crate::consts::RecordType::DiagramBuild {
            let container = diagram_build::parse_record(child)?;
            let shape_id = container.build().shape_id_ref;
            if shape_ids.binary_search(&shape_id).is_err() {
                return Err(Error::InvalidFormat(format!(
                    "diagram build ({}, {}) references missing OfficeArt shape {}",
                    container.build().build_id,
                    shape_id,
                    shape_id
                )));
            }
            let build = Build::new(container);
            if entries
                .iter()
                .any(|entry: &LocatedBuild| entry.build.id() == build.id())
            {
                return Err(Error::InvalidFormat(format!(
                    "duplicate native diagram identity ({}, {})",
                    build.build_id(),
                    build.shape_id()
                )));
            }
            if entries.len() == limits.max_diagrams {
                return Err(Error::InvalidFormat(
                    "diagram build inventory exceeds configured transaction limit".to_string(),
                ));
            }
            if child_len != diagram_build::Container::RECORD_LEN {
                return Err(Error::Corrupted(
                    "DiagramBuild record length is not fixed by MS-PPT".to_string(),
                ));
            }
            entries.push(LocatedBuild { offset, build });
        }
        offset = offset
            .checked_add(child_len)
            .ok_or_else(|| Error::Corrupted("BuildList offset overflow".to_string()))?;
    }
    if offset != bytes.len() {
        return Err(Error::Corrupted(
            "BuildList child offsets do not cover the source".to_string(),
        ));
    }
    Ok(entries)
}

/// Rewrite only the fixed-width diagram build mode field.
pub(super) fn rewrite_mode(bytes: &mut [u8], offset: usize, mode: BuildType) -> Result<()> {
    let start = offset
        .checked_add(MODE_OFFSET)
        .ok_or_else(|| Error::Corrupted("diagram build mode offset overflow".to_string()))?;
    let end = start
        .checked_add(4)
        .ok_or_else(|| Error::Corrupted("diagram build mode extent overflow".to_string()))?;
    let field = bytes.get_mut(start..end).ok_or_else(|| {
        Error::Corrupted("diagram build mode offset is out of bounds".to_string())
    })?;
    field.copy_from_slice(&mode.raw().to_le_bytes());
    Ok(())
}

/// Rewrite only the fixed-width `BuildAtom` `shapeIdRef` field.
pub(super) fn rewrite_shape_id(bytes: &mut [u8], offset: usize, shape_id: u32) -> Result<()> {
    let start = offset
        .checked_add(SHAPE_ID_OFFSET)
        .ok_or_else(|| Error::Corrupted("diagram shape identity offset overflow".to_string()))?;
    let end = start
        .checked_add(4)
        .ok_or_else(|| Error::Corrupted("diagram shape identity extent overflow".to_string()))?;
    let field = bytes.get_mut(start..end).ok_or_else(|| {
        Error::Corrupted("diagram shape identity offset is out of bounds".to_string())
    })?;
    field.copy_from_slice(&shape_id.to_le_bytes());
    Ok(())
}
