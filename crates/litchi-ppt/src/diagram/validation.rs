//! Context validation for the diagram inventory facade.

use std::collections::HashSet;

use crate::animation::diagram_build;
use crate::consts::RecordType;
use crate::odraw::{Drawing, Shape};
use crate::package::{Error, Result};
use crate::records::Record;
use litchi_odraw::shape::Flags as ShapeFlags;

use super::model::{Build, EditLimits, Id, Limits, Payload, PayloadKind, ShapeRef};

pub(super) fn diagram_builds(record: &Record, limits: Limits) -> Result<Vec<Build>> {
    validate_build_list(record)?;
    let mut builds = Vec::new();
    builds
        .try_reserve(record.children.len().min(limits.max_diagrams))
        .map_err(|_| allocation("diagram build inventory"))?;
    let mut identities = HashSet::new();

    for child in &record.children {
        if child.record_type != RecordType::DiagramBuild {
            continue;
        }
        let parsed = diagram_build::parse_record(child)?;
        let id = Id::new(parsed.build().build_id, parsed.build().shape_id_ref);
        if !identities.insert(id) {
            return Err(Error::InvalidFormat(format!(
                "duplicate native diagram identity ({}, {})",
                id.build_id(),
                id.shape_id(),
            )));
        }
        if builds.len() == limits.max_diagrams {
            return Err(limit("diagram build inventory", limits.max_diagrams));
        }
        builds.push(Build::new(parsed));
    }
    Ok(builds)
}

pub(super) fn validate_build_list(record: &Record) -> Result<()> {
    if record.record_type != RecordType::BuildList
        || record.record_type_raw != RecordType::BuildList.as_u16()
    {
        return Err(Error::InvalidFormat(
            "diagram inventory requires a BuildList record".to_string(),
        ));
    }
    if record.version != 0 || record.instance != 0 {
        return Err(Error::Corrupted(
            "BuildList requires record version 0 and instance 0".to_string(),
        ));
    }
    if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
        return Err(Error::Corrupted(
            "BuildList record length does not match its payload".to_string(),
        ));
    }

    let mut encoded_len = 0usize;
    for child in &record.children {
        let child_type = child.record_type;
        if !matches!(
            child_type,
            RecordType::ParaBuild | RecordType::ChartBuild | RecordType::DiagramBuild
        ) {
            return Err(Error::InvalidFormat(format!(
                "BuildList contains unsupported child {child_type:?}"
            )));
        }
        if usize::try_from(child.data_length).ok() != Some(child.data.len()) {
            return Err(Error::Corrupted(
                "BuildList child length does not match its payload".to_string(),
            ));
        }
        encoded_len = encoded_len
            .checked_add(8)
            .and_then(|value| value.checked_add(child.data.len()))
            .ok_or_else(|| Error::Corrupted("BuildList child length overflow".to_string()))?;
    }
    if encoded_len != record.data.len() {
        return Err(Error::Corrupted(
            "BuildList children do not cover its payload".to_string(),
        ));
    }
    Ok(())
}

/// Validate the structural envelope needed by the source-preserving editor.
///
/// Unlike the read-only inventory validator, this deliberately does not
/// reject future or host-specific BuildList children.  They are outside the
/// semantic edit surface but remain byte-for-byte inert in the transaction.
pub(super) fn validate_edit_build_list(record: &Record, limits: EditLimits) -> Result<()> {
    if record.record_type != RecordType::BuildList
        || record.record_type_raw != RecordType::BuildList.as_u16()
    {
        return Err(Error::InvalidFormat(
            "diagram transaction requires a BuildList record".to_string(),
        ));
    }
    if record.instance != 0 || !matches!(record.version, 0 | 0x0F) {
        return Err(Error::Corrupted(
            "BuildList requires record instance 0 and version 0 or 0xF".to_string(),
        ));
    }
    if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
        return Err(Error::Corrupted(
            "BuildList record length does not match its payload".to_string(),
        ));
    }
    if record.data.len() > limits.max_build_list_bytes.saturating_sub(8) {
        return Err(Error::InvalidFormat(
            "BuildList exceeds the configured transaction byte limit".to_string(),
        ));
    }

    let mut encoded_len = 0usize;
    for child in &record.children {
        if usize::try_from(child.data_length).ok() != Some(child.data.len()) {
            return Err(Error::Corrupted(
                "BuildList child length does not match its payload".to_string(),
            ));
        }
        encoded_len = encoded_len
            .checked_add(8)
            .and_then(|value| value.checked_add(child.data.len()))
            .ok_or_else(|| Error::Corrupted("BuildList child length overflow".to_string()))?;
    }
    if encoded_len != record.data.len() {
        return Err(Error::Corrupted(
            "BuildList children do not cover its payload".to_string(),
        ));
    }
    Ok(())
}

/// Parse and validate the bounded shape graph used by checked shape edits.
pub(super) fn shape_ids(drawing: &Drawing<'_>, maximum: usize) -> Result<Vec<u32>> {
    let mut ids = HashSet::new();
    ids.try_reserve(drawing.shapes().len().min(maximum))
        .map_err(|_| allocation("diagram shape index"))?;
    for shape in drawing.shapes() {
        collect_shape_ids_bounded(shape, &mut ids, maximum)?;
    }
    let mut ordered = ids.into_iter().collect::<Vec<_>>();
    ordered.sort_unstable();
    Ok(ordered)
}

fn collect_shape_ids_bounded(
    shape: &Shape<'_>,
    ids: &mut HashSet<u32>,
    maximum: usize,
) -> Result<()> {
    if !ids.insert(shape.id()) {
        return Err(Error::InvalidFormat(format!(
            "duplicate OfficeArt shape identifier {}",
            shape.id()
        )));
    }
    if ids.len() > maximum {
        return Err(limit("diagram shape index", maximum));
    }
    for child in shape.children() {
        collect_shape_ids_bounded(child, ids, maximum)?;
    }
    Ok(())
}

pub(super) fn validate_shape_ids(drawing: &Drawing<'_>) -> Result<()> {
    let mut ids = HashSet::new();
    for shape in drawing.shapes() {
        collect_shape_ids(shape, &mut ids)?;
    }
    Ok(())
}

fn collect_shape_ids(shape: &Shape<'_>, ids: &mut HashSet<u32>) -> Result<()> {
    if !ids.insert(shape.id()) {
        return Err(Error::InvalidFormat(format!(
            "duplicate OfficeArt shape identifier {}",
            shape.id()
        )));
    }
    for child in shape.children() {
        collect_shape_ids(child, ids)?;
    }
    Ok(())
}

pub(super) fn collect_associated<'data>(
    root: &Shape<'data>,
    limits: Limits,
) -> Result<(Vec<ShapeRef>, Vec<Payload<'data>>)> {
    let mut shapes = Vec::new();
    let mut payloads = Vec::new();
    collect_shape(root, limits, &mut shapes, &mut payloads)?;
    Ok((shapes, payloads))
}

fn collect_shape<'data>(
    shape: &Shape<'data>,
    limits: Limits,
    shapes: &mut Vec<ShapeRef>,
    payloads: &mut Vec<Payload<'data>>,
) -> Result<()> {
    if shapes.len() == limits.max_shapes_per_diagram {
        return Err(limit("associated shapes", limits.max_shapes_per_diagram));
    }
    shapes
        .try_reserve(1)
        .map_err(|_| allocation("associated shape references"))?;
    shapes.push(ShapeRef::new(shape.id()));

    add_payload(
        shape.id(),
        PayloadKind::Shape,
        shape.meta().record().clone(),
        limits,
        payloads,
    )?;
    if shape.flags().contains(ShapeFlags::GROUP) {
        add_payload(
            shape.id(),
            PayloadKind::Group,
            shape.container().record().clone(),
            limits,
            payloads,
        )?;
    }
    if let Some(record) = shape.client_data() {
        add_payload(
            shape.id(),
            PayloadKind::ClientData,
            record.clone(),
            limits,
            payloads,
        )?;
    }
    if let Some(record) = shape.textbox() {
        add_payload(
            shape.id(),
            PayloadKind::Textbox,
            record.clone(),
            limits,
            payloads,
        )?;
    }
    if let Some(record) = shape.client_anchor() {
        add_payload(
            shape.id(),
            PayloadKind::Anchor,
            record.clone(),
            limits,
            payloads,
        )?;
    }

    for child in shape.children() {
        collect_shape(child, limits, shapes, payloads)?;
    }
    Ok(())
}

fn add_payload<'data>(
    shape_id: u32,
    kind: PayloadKind,
    record: litchi_odraw::Record<'data>,
    limits: Limits,
    payloads: &mut Vec<Payload<'data>>,
) -> Result<()> {
    if payloads.len() == limits.max_payloads_per_diagram {
        return Err(limit(
            "diagram payload references",
            limits.max_payloads_per_diagram,
        ));
    }
    payloads
        .try_reserve(1)
        .map_err(|_| allocation("diagram payload references"))?;
    payloads.push(Payload::new(shape_id, kind, record));
    Ok(())
}

fn limit(resource: &str, maximum: usize) -> Error {
    Error::InvalidFormat(format!("{resource} exceeds configured limit {maximum}"))
}

fn allocation(resource: &str) -> Error {
    Error::Corrupted(format!("unable to allocate {resource}"))
}
