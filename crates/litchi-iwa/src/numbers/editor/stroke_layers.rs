//! Native Numbers cell-border layer maintenance during table-axis edits.
//!
//! A stroke sidecar owns four sets of sparse border layers. Layers on the
//! edited axis carry an axis index; perpendicular layers carry runs over the
//! edited axis. This module keeps existing borders attached to their original
//! cells when the editor inserts or removes a blank row or column.

use super::*;

const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const STROKE_LAYER_MESSAGE_TYPE: u32 = 6_306;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum StrokeAxis {
    Row,
    Column,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum AxisEdit {
    Insert { index: u32 },
    Delete { index: u32 },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct TableDimensions {
    rows: u32,
    columns: u32,
}

impl TableDimensions {
    fn inserted(self, axis: StrokeAxis) -> Result<Self> {
        match axis {
            StrokeAxis::Row => Ok(Self {
                rows: self
                    .rows
                    .checked_add(1)
                    .ok_or_else(|| Error::ParseError("Numbers row count overflow".to_owned()))?,
                ..self
            }),
            StrokeAxis::Column => Ok(Self {
                columns: self
                    .columns
                    .checked_add(1)
                    .ok_or_else(|| Error::ParseError("Numbers column count overflow".to_owned()))?,
                ..self
            }),
        }
    }

    fn deleted(self, axis: StrokeAxis) -> Result<Self> {
        match axis {
            StrokeAxis::Row => Ok(Self {
                rows: self
                    .rows
                    .checked_sub(1)
                    .ok_or_else(|| Error::ParseError("Numbers row count underflow".to_owned()))?,
                ..self
            }),
            StrokeAxis::Column => Ok(Self {
                columns: self.columns.checked_sub(1).ok_or_else(|| {
                    Error::ParseError("Numbers column count underflow".to_owned())
                })?,
                ..self
            }),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct StrokeLayerMutation {
    axis: StrokeAxis,
    edit: AxisEdit,
    previous: TableDimensions,
    current: TableDimensions,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum LayerSide {
    Left,
    Right,
    Top,
    Bottom,
}

impl LayerSide {
    const ALL: [Self; 4] = [Self::Left, Self::Right, Self::Top, Self::Bottom];

    const fn field_number(self) -> u32 {
        match self {
            Self::Left => 4,
            Self::Right => 5,
            Self::Top => 6,
            Self::Bottom => 7,
        }
    }

    const fn fixed_axis(self) -> StrokeAxis {
        match self {
            Self::Left | Self::Right => StrokeAxis::Column,
            Self::Top | Self::Bottom => StrokeAxis::Row,
        }
    }

    fn references(self, sidecar: &tst::StrokeSidecarArchive) -> &[tsp::Reference] {
        match self {
            Self::Left => &sidecar.left_column_stroke_layers,
            Self::Right => &sidecar.right_column_stroke_layers,
            Self::Top => &sidecar.top_row_stroke_layers,
            Self::Bottom => &sidecar.bottom_row_stroke_layers,
        }
    }

    fn references_mut(self, sidecar: &mut tst::StrokeSidecarArchive) -> &mut Vec<tsp::Reference> {
        match self {
            Self::Left => &mut sidecar.left_column_stroke_layers,
            Self::Right => &mut sidecar.right_column_stroke_layers,
            Self::Top => &mut sidecar.top_row_stroke_layers,
            Self::Bottom => &mut sidecar.bottom_row_stroke_layers,
        }
    }
}

#[derive(Debug)]
struct LayerUpdate {
    identifier: u64,
    archive_name: String,
    message_index: usize,
    previous: tst::StrokeLayerArchive,
    current: tst::StrokeLayerArchive,
    run_sources: Vec<usize>,
}

#[derive(Debug)]
struct PlannedLayer {
    keep: bool,
    update: Option<LayerUpdate>,
}

/// Shift explicit border layers for one inserted table axis.
pub(super) fn insert(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar_id: u64,
    axis: StrokeAxis,
    index: usize,
    rows: u32,
    columns: u32,
) -> Result<()> {
    let index = u32::try_from(index)
        .map_err(|_| Error::ParseError("Numbers stroke-layer index exceeds u32".to_owned()))?;
    let previous = TableDimensions { rows, columns };
    let edited_length = axis_length(axis, previous);
    if index > edited_length {
        return Err(Error::InvalidFormat(format!(
            "Numbers stroke-layer insertion index {index} exceeds {axis:?} count {edited_length}"
        )));
    }
    mutate(
        package,
        locations,
        sidecar_id,
        StrokeLayerMutation {
            axis,
            edit: AxisEdit::Insert { index },
            previous,
            current: previous.inserted(axis)?,
        },
    )
}

/// Shift explicit border layers for one deleted table axis.
pub(super) fn delete(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar_id: u64,
    axis: StrokeAxis,
    index: usize,
    rows: u32,
    columns: u32,
) -> Result<()> {
    let index = u32::try_from(index)
        .map_err(|_| Error::ParseError("Numbers stroke-layer index exceeds u32".to_owned()))?;
    let previous = TableDimensions { rows, columns };
    let edited_length = axis_length(axis, previous);
    if index >= edited_length {
        return Err(Error::InvalidFormat(format!(
            "Numbers stroke-layer deletion index {index} exceeds {axis:?} count {edited_length}"
        )));
    }
    mutate(
        package,
        locations,
        sidecar_id,
        StrokeLayerMutation {
            axis,
            edit: AxisEdit::Delete { index },
            previous,
            current: previous.deleted(axis)?,
        },
    )
}

fn mutate(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar_id: u64,
    mutation: StrokeLayerMutation,
) -> Result<()> {
    let sidecar_archive = locations.get(&sidecar_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers stroke sidecar {sidecar_id} is missing"))
    })?;
    let (sidecar_message_index, previous) = {
        let archive = package.archive(sidecar_archive)?;
        let object = archive.object(sidecar_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers stroke sidecar {sidecar_id} is missing"))
        })?;
        let message_index = unique_message_index(
            object,
            STROKE_SIDECAR_MESSAGE_TYPE,
            |data| tst::StrokeSidecarArchive::decode(data).is_ok(),
            "stroke sidecar",
        )?;
        (
            message_index,
            tst::StrokeSidecarArchive::decode(object.messages[message_index].data.as_slice())?,
        )
    };
    validate_sidecar_dimensions(&previous, mutation.previous)?;

    let mut current = previous.clone();
    current.row_count = Some(mutation.current.rows);
    current.column_count = Some(mutation.current.columns);
    let mut updates = Vec::new();
    let mut seen_layers = HashSet::new();

    for side in LayerSide::ALL {
        let mut references = Vec::with_capacity(side.references(&previous).len());
        for reference in side.references(&previous) {
            if !seen_layers.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers stroke sidecar {sidecar_id} references layer {} more than once",
                    reference.identifier
                )));
            }
            let plan = plan_layer(package, locations, reference.identifier, side, mutation)?;
            if plan.keep {
                references.push(*reference);
            }
            if let Some(update) = plan.update {
                updates.push(update);
            }
        }
        *side.references_mut(&mut current) = references;
    }

    for update in &updates {
        package.update_archive(&update.archive_name, |archive| {
            let object = archive.object_mut(update.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers stroke layer {} is missing",
                    update.identifier
                ))
            })?;
            let original = object.messages.get(update.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers stroke layer {} payload is missing",
                    update.identifier
                ))
            })?;
            let decoded = tst::StrokeLayerArchive::decode(original.data.as_slice())?;
            if decoded != update.previous {
                return Err(Error::InvalidFormat(format!(
                    "Numbers stroke layer {} changed before mutation",
                    update.identifier
                )));
            }
            let data = rewrite_stroke_layer_wire(
                original.data.as_slice(),
                &update.previous,
                &update.current,
                &update.run_sources,
            )?;
            object.replace_message(
                update.message_index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            Ok(())
        })?;
    }

    package.update_archive(sidecar_archive, |archive| {
        let object = archive.object_mut(sidecar_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers stroke sidecar {sidecar_id} is missing"))
        })?;
        let original = object.messages.get(sidecar_message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers stroke sidecar {sidecar_id} payload is missing"
            ))
        })?;
        let decoded = tst::StrokeSidecarArchive::decode(original.data.as_slice())?;
        if decoded != previous {
            return Err(Error::InvalidFormat(format!(
                "Numbers stroke sidecar {sidecar_id} changed before mutation"
            )));
        }
        let data = rewrite_stroke_sidecar_wire(original.data.as_slice(), &previous, &current)?;
        object.replace_message(
            sidecar_message_index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let previous_layers = referenced_layer_identifiers(&previous);
        let retained = referenced_layer_identifiers(&current);
        let message_info = object
            .archive_info
            .message_infos
            .get_mut(sidecar_message_index)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers stroke sidecar {sidecar_id} has no matching message metadata"
                ))
            })?;
        message_info.object_references.retain(|identifier| {
            !previous_layers.contains(identifier) || retained.contains(identifier)
        });
        Ok(())
    })
}

fn plan_layer(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    side: LayerSide,
    mutation: StrokeLayerMutation,
) -> Result<PlannedLayer> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers stroke layer {identifier} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers stroke layer {identifier} is missing"))
    })?;
    let message_index = unique_message_index(
        object,
        STROKE_LAYER_MESSAGE_TYPE,
        |data| tst::StrokeLayerArchive::decode(data).is_ok(),
        "stroke layer",
    )?;
    let previous = tst::StrokeLayerArchive::decode(object.messages[message_index].data.as_slice())?;
    let fixed_length = axis_length(side.fixed_axis(), mutation.previous);
    let traversal_length = axis_length(opposite_axis(side.fixed_axis()), mutation.previous);
    let layer_index = previous.row_column_index.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers stroke layer {identifier} has no row/column index"
        ))
    })?;
    if layer_index >= fixed_length {
        return Err(Error::InvalidFormat(format!(
            "Numbers stroke layer {identifier} index {layer_index} exceeds its axis length {fixed_length}"
        )));
    }

    let (keep, current, run_sources) = if side.fixed_axis() == mutation.axis {
        let Some(index) = shifted_layer_index(layer_index, mutation.edit)? else {
            return Ok(PlannedLayer {
                keep: false,
                update: None,
            });
        };
        let mut current = previous.clone();
        current.row_column_index = Some(index);
        (true, current, (0..previous.stroke_runs.len()).collect())
    } else {
        let (runs, run_sources) =
            transform_runs(&previous.stroke_runs, mutation.edit, traversal_length)?;
        if runs.is_empty() && !previous.stroke_runs.is_empty() {
            return Ok(PlannedLayer {
                keep: false,
                update: None,
            });
        }
        let mut current = previous.clone();
        current.stroke_runs = runs;
        (true, current, run_sources)
    };

    let update = (current != previous).then_some(LayerUpdate {
        identifier,
        archive_name: archive_name.clone(),
        message_index,
        previous,
        current,
        run_sources,
    });
    Ok(PlannedLayer { keep, update })
}

fn transform_runs(
    runs: &[tst::stroke_layer_archive::StrokeRunArchive],
    edit: AxisEdit,
    length: u32,
) -> Result<(Vec<tst::stroke_layer_archive::StrokeRunArchive>, Vec<usize>)> {
    let mut transformed = Vec::with_capacity(runs.len().saturating_add(1));
    let mut sources = Vec::with_capacity(runs.len().saturating_add(1));
    for (source, run) in runs.iter().enumerate() {
        let (start, end) = validated_run_range(run, length)?;
        match edit {
            AxisEdit::Insert { index } if index <= start => {
                let mut shifted = run.clone();
                shifted.origin = Some(checked_i32(indexed_add(start, 1)?)?);
                transformed.push(shifted);
                sources.push(source);
            },
            AxisEdit::Insert { index } if index >= end => {
                transformed.push(run.clone());
                sources.push(source);
            },
            AxisEdit::Insert { index } => {
                let mut before = run.clone();
                before.length = Some(index - start);
                transformed.push(before);
                sources.push(source);

                let mut after = run.clone();
                after.origin = Some(checked_i32(indexed_add(index, 1)?)?);
                after.length = Some(end - index);
                transformed.push(after);
                sources.push(source);
            },
            AxisEdit::Delete { index } if index < start => {
                let mut shifted = run.clone();
                shifted.origin = Some(checked_i32(start - 1)?);
                transformed.push(shifted);
                sources.push(source);
            },
            AxisEdit::Delete { index } if index >= end => {
                transformed.push(run.clone());
                sources.push(source);
            },
            AxisEdit::Delete { .. } if end - start == 1 => {},
            AxisEdit::Delete { .. } => {
                let mut shortened = run.clone();
                shortened.length = Some(end - start - 1);
                transformed.push(shortened);
                sources.push(source);
            },
        }
    }
    Ok((transformed, sources))
}

fn shifted_layer_index(index: u32, edit: AxisEdit) -> Result<Option<u32>> {
    match edit {
        AxisEdit::Insert { index: insertion } if index >= insertion => {
            Ok(Some(indexed_add(index, 1)?))
        },
        AxisEdit::Insert { .. } => Ok(Some(index)),
        AxisEdit::Delete { index: deletion } if index == deletion => Ok(None),
        AxisEdit::Delete { index: deletion } if index > deletion => Ok(Some(index - 1)),
        AxisEdit::Delete { .. } => Ok(Some(index)),
    }
}

fn validated_run_range(
    run: &tst::stroke_layer_archive::StrokeRunArchive,
    length: u32,
) -> Result<(u32, u32)> {
    let origin = run
        .origin
        .ok_or_else(|| Error::InvalidFormat("Numbers stroke run has no origin".to_owned()))?;
    let start = u32::try_from(origin).map_err(|_| {
        Error::InvalidFormat(format!("Numbers stroke run has negative origin {origin}"))
    })?;
    let run_length = run
        .length
        .ok_or_else(|| Error::InvalidFormat("Numbers stroke run has no length".to_owned()))?;
    if run_length == 0 {
        return Err(Error::InvalidFormat(
            "Numbers stroke run has zero length".to_owned(),
        ));
    }
    let end = start
        .checked_add(run_length)
        .ok_or_else(|| Error::InvalidFormat("Numbers stroke run range overflows u32".to_owned()))?;
    if end > length {
        return Err(Error::InvalidFormat(format!(
            "Numbers stroke run {start}..{end} exceeds axis length {length}"
        )));
    }
    Ok((start, end))
}

fn rewrite_stroke_layer_wire(
    original: &[u8],
    previous: &tst::StrokeLayerArchive,
    current: &tst::StrokeLayerArchive,
    run_sources: &[usize],
) -> Result<Vec<u8>> {
    if current.stroke_runs.len() != run_sources.len() {
        return Err(Error::InvalidFormat(
            "Numbers stroke-layer run source count is inconsistent".to_owned(),
        ));
    }
    let raw_runs = repeated_length_delimited_payloads(original, 2)?;
    if raw_runs.len() != previous.stroke_runs.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers stroke layer has {} raw runs but {} decoded runs",
            raw_runs.len(),
            previous.stroke_runs.len()
        )));
    }
    for (raw, expected) in raw_runs.iter().zip(&previous.stroke_runs) {
        if tst::stroke_layer_archive::StrokeRunArchive::decode(*raw)? != *expected {
            return Err(Error::InvalidFormat(
                "Numbers stroke-layer run changed during mutation".to_owned(),
            ));
        }
    }
    let replacements = current
        .stroke_runs
        .iter()
        .zip(run_sources)
        .map(|(run, &source)| {
            let previous_run = previous.stroke_runs.get(source).ok_or_else(|| {
                Error::InvalidFormat("Numbers stroke-layer run source is missing".to_owned())
            })?;
            rewrite_stroke_run_wire(raw_runs[source], previous_run, run)
        })
        .collect::<Result<Vec<_>>>()?;
    let mut data = rewrite_repeated_length_delimited_fields(original, 2, &replacements)?;
    data = patch_varint_field(
        &data,
        1,
        previous.row_column_index.is_some(),
        current.row_column_index.map(u64::from),
    )?;
    if tst::StrokeLayerArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers stroke-layer wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_stroke_run_wire(
    original: &[u8],
    previous: &tst::stroke_layer_archive::StrokeRunArchive,
    current: &tst::stroke_layer_archive::StrokeRunArchive,
) -> Result<Vec<u8>> {
    let mut expected = current.clone();
    expected.origin = previous.origin;
    expected.length = previous.length;
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers stroke run changed outside its coordinates".to_owned(),
        ));
    }
    let mut data = patch_varint_field(
        original,
        1,
        previous.origin.is_some(),
        current.origin.map(int32_wire_value),
    )?;
    data = patch_varint_field(
        &data,
        2,
        previous.length.is_some(),
        current.length.map(u64::from),
    )?;
    if tst::stroke_layer_archive::StrokeRunArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers stroke-run wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_stroke_sidecar_wire(
    original: &[u8],
    previous: &tst::StrokeSidecarArchive,
    current: &tst::StrokeSidecarArchive,
) -> Result<Vec<u8>> {
    let mut expected = current.clone();
    expected.column_count = previous.column_count;
    expected.row_count = previous.row_count;
    for side in LayerSide::ALL {
        *side.references_mut(&mut expected) = side.references(previous).to_vec();
    }
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers stroke sidecar changed outside its dimensions and layer references".to_owned(),
        ));
    }
    let mut data = patch_varint_field(
        original,
        2,
        previous.column_count.is_some(),
        current.column_count.map(u64::from),
    )?;
    data = patch_varint_field(
        &data,
        3,
        previous.row_count.is_some(),
        current.row_count.map(u64::from),
    )?;
    for side in LayerSide::ALL {
        data = rewrite_reference_list(
            &data,
            side.field_number(),
            side.references(previous),
            side.references(current),
        )?;
    }
    if tst::StrokeSidecarArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers stroke-sidecar wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_reference_list(
    data: &[u8],
    field_number: u32,
    previous: &[tsp::Reference],
    current: &[tsp::Reference],
) -> Result<Vec<u8>> {
    let raw_references = repeated_length_delimited_payloads(data, field_number)?;
    if raw_references.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers stroke sidecar field {field_number} has {} raw references but {} decoded references",
            raw_references.len(),
            previous.len()
        )));
    }
    let mut raw_by_identifier = HashMap::with_capacity(previous.len());
    for (raw, expected) in raw_references.iter().zip(previous) {
        if tsp::Reference::decode(*raw)? != *expected {
            return Err(Error::InvalidFormat(format!(
                "Numbers stroke sidecar field {field_number} changed during mutation"
            )));
        }
        if raw_by_identifier
            .insert(expected.identifier, *raw)
            .is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers stroke sidecar field {field_number} has duplicate layer reference {}",
                expected.identifier
            )));
        }
    }
    let replacements = current
        .iter()
        .map(|reference| {
            let raw = raw_by_identifier
                .get(&reference.identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers stroke sidecar introduced layer reference {}",
                        reference.identifier
                    ))
                })?;
            if tsp::Reference::decode(*raw)? != *reference {
                return Err(Error::InvalidFormat(format!(
                    "Numbers stroke sidecar changed layer reference {}",
                    reference.identifier
                )));
            }
            Ok((*raw).to_vec())
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

fn validate_sidecar_dimensions(
    sidecar: &tst::StrokeSidecarArchive,
    dimensions: TableDimensions,
) -> Result<()> {
    if sidecar
        .row_count
        .is_some_and(|count| count != dimensions.rows)
        || sidecar
            .column_count
            .is_some_and(|count| count != dimensions.columns)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers stroke sidecar dimensions {:?}x{:?} do not match table {}x{}",
            sidecar.row_count, sidecar.column_count, dimensions.rows, dimensions.columns
        )));
    }
    Ok(())
}

fn referenced_layer_identifiers(sidecar: &tst::StrokeSidecarArchive) -> HashSet<u64> {
    LayerSide::ALL
        .into_iter()
        .flat_map(|side| side.references(sidecar))
        .map(|reference| reference.identifier)
        .collect()
}

fn unique_message_index(
    object: &ArchiveObject,
    expected_type: u32,
    matches: impl Fn(&[u8]) -> bool,
    description: &str,
) -> Result<usize> {
    let typed_indexes = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| {
            (message.type_ == expected_type && matches(message.data.as_slice())).then_some(index)
        })
        .collect::<Vec<_>>();
    match typed_indexes.as_slice() {
        [index] => Ok(*index),
        [] if object
            .messages
            .iter()
            .any(|message| message.type_ == expected_type) =>
        {
            Err(Error::InvalidFormat(format!(
                "Object {:?} has no decodable Numbers {description} payload with message type {expected_type}",
                object.archive_info.identifier
            )))
        },
        [] => unique_decoded_message_index(object, matches, description),
        _ => Err(Error::InvalidFormat(format!(
            "Object {:?} has multiple Numbers {description} payloads with message type {expected_type}",
            object.archive_info.identifier
        ))),
    }
}

fn unique_decoded_message_index(
    object: &ArchiveObject,
    matches: impl Fn(&[u8]) -> bool,
    description: &str,
) -> Result<usize> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| matches(message.data.as_slice()).then_some(index))
        .collect::<Vec<_>>();
    match indexes.as_slice() {
        [index] => Ok(*index),
        [] => Err(Error::InvalidFormat(format!(
            "Object {:?} has no Numbers {description} payload",
            object.archive_info.identifier
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "Object {:?} has multiple Numbers {description} payloads",
            object.archive_info.identifier
        ))),
    }
}

const fn axis_length(axis: StrokeAxis, dimensions: TableDimensions) -> u32 {
    match axis {
        StrokeAxis::Row => dimensions.rows,
        StrokeAxis::Column => dimensions.columns,
    }
}

const fn opposite_axis(axis: StrokeAxis) -> StrokeAxis {
    match axis {
        StrokeAxis::Row => StrokeAxis::Column,
        StrokeAxis::Column => StrokeAxis::Row,
    }
}

fn indexed_add(value: u32, increment: u32) -> Result<u32> {
    value
        .checked_add(increment)
        .ok_or_else(|| Error::ParseError("Numbers stroke-layer coordinate overflow".to_owned()))
}

fn checked_i32(value: u32) -> Result<i32> {
    i32::try_from(value)
        .map_err(|_| Error::ParseError("Numbers stroke-layer coordinate exceeds i32".to_owned()))
}

const fn int32_wire_value(value: i32) -> u64 {
    value as i64 as u64
}
