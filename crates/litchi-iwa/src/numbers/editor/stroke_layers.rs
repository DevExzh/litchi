//! Native Numbers cell-border layer maintenance during table-axis edits.
//!
//! A stroke sidecar owns four sets of sparse border layers. Layers on the
//! edited axis carry an axis index; perpendicular layers carry runs over the
//! edited axis. This module keeps existing borders attached to their original
//! cells when the editor inserts or removes a blank row or column.

use super::table::cell::{BorderSide, Borders};
use super::*;
use crate::package_metadata::{next_object_identifier, set_package_last_object_identifier};
use crate::shapes::{empty_stroke_archive, stroke_from_native, stroke_to_native};
use crate::wire::patch_length_delimited_field;

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

    const fn from_public(side: BorderSide) -> Self {
        match side {
            BorderSide::Left => Self::Left,
            BorderSide::Right => Self::Right,
            BorderSide::Top => Self::Top,
            BorderSide::Bottom => Self::Bottom,
        }
    }

    const fn public(self) -> BorderSide {
        match self {
            Self::Left => BorderSide::Left,
            Self::Right => BorderSide::Right,
            Self::Top => BorderSide::Top,
            Self::Bottom => BorderSide::Bottom,
        }
    }

    const fn coordinate(self, row: u32, column: u32) -> (u32, u32) {
        match self.fixed_axis() {
            StrokeAxis::Row => (row, column),
            StrokeAxis::Column => (column, row),
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
    run_sources: Vec<Option<usize>>,
}

#[derive(Debug)]
struct PlannedLayer {
    keep: bool,
    update: Option<LayerUpdate>,
}

pub(crate) fn cell_borders(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Borders> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let (row, column, dimensions) = validated_cell_coordinates(&descriptor.model, row, column)?;
    let sidecar_id = descriptor
        .model
        .stroke_sidecar
        .as_ref()
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table {table_id} has no cell-border sidecar"))
        })?
        .identifier;
    let locations = object_locations(package)?;
    let sidecar = decoded_sidecar(package, &locations, sidecar_id)?;
    validate_sidecar_dimensions(&sidecar, dimensions)?;

    let mut borders = Borders::default();
    let mut seen_layers = HashSet::new();
    for side in LayerSide::ALL {
        let (fixed_index, traversal_index) = side.coordinate(row, column);
        let mut selected: Option<(u32, usize, tsd::StrokeArchive)> = None;
        let mut encounter = 0;
        for reference in side.references(&sidecar) {
            if !seen_layers.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "iWork stroke sidecar {sidecar_id} references layer {} more than once",
                    reference.identifier
                )));
            }
            let layer = decoded_layer(package, &locations, reference.identifier)?;
            if layer.row_column_index != Some(fixed_index) {
                continue;
            }
            for (run_index, run) in layer.stroke_runs.iter().enumerate() {
                encounter += 1;
                let (start, end) = validated_run_range(
                    run,
                    axis_length(opposite_axis(side.fixed_axis()), dimensions),
                )?;
                if !(start..end).contains(&traversal_index) {
                    continue;
                }
                let stroke = run.stroke.as_ref().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork stroke layer {} run {run_index} has no stroke",
                        reference.identifier
                    ))
                })?;
                let candidate = (run.order.unwrap_or_default(), encounter, stroke.clone());
                if selected
                    .as_ref()
                    .is_none_or(|(order, index, _)| (candidate.0, candidate.1) > (*order, *index))
                {
                    selected = Some(candidate);
                }
            }
        }
        borders.set(
            side.public(),
            selected
                .map(|(_, _, stroke)| stroke_from_native(&stroke))
                .transpose()?
                .flatten(),
        );
    }
    Ok(borders)
}

pub(crate) fn set_cell_border(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    public_side: BorderSide,
    stroke: Option<crate::shapes::ShapeStroke>,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let (row, column, dimensions) = validated_cell_coordinates(&descriptor.model, row, column)?;
    let sidecar_id = descriptor
        .model
        .stroke_sidecar
        .as_ref()
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table {table_id} has no cell-border sidecar"))
        })?
        .identifier;
    let locations = object_locations(package)?;
    let sidecar_archive = locations.get(&sidecar_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork stroke sidecar {sidecar_id} is missing"))
    })?;
    let (sidecar_message_index, previous) =
        decoded_sidecar_with_index(package, sidecar_archive, sidecar_id)?;
    validate_sidecar_dimensions(&previous, dimensions)?;

    let side = LayerSide::from_public(public_side);
    let (fixed_index, traversal_index) = side.coordinate(row, column);
    let next_order = highest_order(package, &locations, &previous)?
        .max(previous.max_order.unwrap_or_default())
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork cell-border order overflow".to_owned()))?;
    let run = tst::stroke_layer_archive::StrokeRunArchive {
        origin: Some(checked_i32(traversal_index)?),
        length: Some(1),
        stroke: Some(stroke.map_or_else(empty_stroke_archive, stroke_to_native)),
        order: Some(next_order),
    };

    let mut current = previous.clone();
    current.max_order = Some(next_order);
    let existing_layer =
        find_layer_at_index(package, &locations, side.references(&previous), fixed_index)?;
    if let Some(mut update) = existing_layer {
        let traversal_length = axis_length(opposite_axis(side.fixed_axis()), dimensions);
        let mut effective = None;
        for (index, candidate) in update.current.stroke_runs.iter().enumerate() {
            let (start, end) = validated_run_range(candidate, traversal_length)?;
            if (start..end).contains(&traversal_index) {
                effective = effective.max(Some((candidate.order.unwrap_or_default(), index)));
            }
        }
        if let Some((_, index)) = effective
            && update.current.stroke_runs[index].origin == run.origin
            && update.current.stroke_runs[index].length == Some(1)
        {
            update.current.stroke_runs[index].stroke = run.stroke;
            update.current.stroke_runs[index].order = run.order;
        } else {
            update.current.stroke_runs.push(run);
            update.run_sources.push(None);
        }
        apply_layer_updates(package, &[update])?;
    } else {
        let identifier = next_object_identifier(package)?;
        let layer = tst::StrokeLayerArchive {
            row_column_index: Some(fixed_index),
            stroke_runs: vec![run],
        };
        let object = ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: STROKE_LAYER_MESSAGE_TYPE,
                data: layer.encode_to_vec(),
            }],
        )?;
        package.update_archive(sidecar_archive, |archive| archive.insert_object(object))?;
        side.references_mut(&mut current).push(tsp::Reference {
            identifier,
            ..Default::default()
        });
        set_package_last_object_identifier(package, identifier)?;
    }

    package.update_archive(sidecar_archive, |archive| {
        let object = archive.object_mut(sidecar_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork stroke sidecar {sidecar_id} is missing"))
        })?;
        let original = object.messages.get(sidecar_message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork stroke sidecar {sidecar_id} payload is missing"
            ))
        })?;
        if tst::StrokeSidecarArchive::decode(original.data.as_slice())? != previous {
            return Err(Error::InvalidFormat(format!(
                "iWork stroke sidecar {sidecar_id} changed before mutation"
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
        if let Some(reference) = side
            .references(&current)
            .iter()
            .find(|reference| !side.references(&previous).contains(reference))
        {
            add_message_object_reference(
                object,
                sidecar_message_index,
                side.references(&previous)
                    .first()
                    .map_or(reference.identifier, |existing| existing.identifier),
                reference.identifier,
            );
        }
        Ok(())
    })
}

fn validated_cell_coordinates(
    model: &TableModelArchive,
    row: usize,
    column: usize,
) -> Result<(u32, u32, TableDimensions)> {
    if row >= model.number_of_rows as usize || column >= model.number_of_columns as usize {
        return Err(Error::ParseError(format!(
            "Cell ({row}, {column}) is outside iWork table {:?} dimensions {}x{}",
            model.table_name, model.number_of_rows, model.number_of_columns
        )));
    }
    Ok((
        u32::try_from(row)
            .map_err(|_| Error::ParseError("iWork table row exceeds u32".to_owned()))?,
        u32::try_from(column)
            .map_err(|_| Error::ParseError("iWork table column exceeds u32".to_owned()))?,
        TableDimensions {
            rows: model.number_of_rows,
            columns: model.number_of_columns,
        },
    ))
}

fn decoded_sidecar(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar_id: u64,
) -> Result<tst::StrokeSidecarArchive> {
    let archive_name = locations.get(&sidecar_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork stroke sidecar {sidecar_id} is missing"))
    })?;
    decoded_sidecar_with_index(package, archive_name, sidecar_id).map(|(_, sidecar)| sidecar)
}

fn decoded_sidecar_with_index(
    package: &IWorkPackage,
    archive_name: &str,
    sidecar_id: u64,
) -> Result<(usize, tst::StrokeSidecarArchive)> {
    let archive = package.archive(archive_name)?;
    let object = archive.object(sidecar_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork stroke sidecar {sidecar_id} is missing"))
    })?;
    let index = unique_message_index(
        object,
        STROKE_SIDECAR_MESSAGE_TYPE,
        |data| tst::StrokeSidecarArchive::decode(data).is_ok(),
        "stroke sidecar",
    )?;
    Ok((
        index,
        tst::StrokeSidecarArchive::decode(object.messages[index].data.as_slice())?,
    ))
}

fn decoded_layer(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
) -> Result<tst::StrokeLayerArchive> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork stroke layer {identifier} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork stroke layer {identifier} is missing"))
    })?;
    let index = unique_message_index(
        object,
        STROKE_LAYER_MESSAGE_TYPE,
        |data| tst::StrokeLayerArchive::decode(data).is_ok(),
        "stroke layer",
    )?;
    tst::StrokeLayerArchive::decode(object.messages[index].data.as_slice()).map_err(Into::into)
}

fn highest_order(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar: &tst::StrokeSidecarArchive,
) -> Result<u32> {
    let mut highest = 0;
    let mut seen = HashSet::new();
    for side in LayerSide::ALL {
        for reference in side.references(sidecar) {
            if !seen.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "iWork stroke sidecar references layer {} more than once",
                    reference.identifier
                )));
            }
            for run in decoded_layer(package, locations, reference.identifier)?.stroke_runs {
                highest = highest.max(run.order.unwrap_or_default());
            }
        }
    }
    Ok(highest)
}

fn find_layer_at_index(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    references: &[tsp::Reference],
    fixed_index: u32,
) -> Result<Option<LayerUpdate>> {
    let mut found = None;
    for reference in references {
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork stroke layer {} is missing",
                reference.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork stroke layer {} is missing",
                reference.identifier
            ))
        })?;
        let message_index = unique_message_index(
            object,
            STROKE_LAYER_MESSAGE_TYPE,
            |data| tst::StrokeLayerArchive::decode(data).is_ok(),
            "stroke layer",
        )?;
        let previous =
            tst::StrokeLayerArchive::decode(object.messages[message_index].data.as_slice())?;
        if previous.row_column_index != Some(fixed_index) {
            continue;
        }
        if found.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork cell-border side has multiple layers at index {fixed_index}"
            )));
        }
        found = Some(LayerUpdate {
            identifier: reference.identifier,
            archive_name: archive_name.clone(),
            message_index,
            run_sources: (0..previous.stroke_runs.len()).map(Some).collect(),
            current: previous.clone(),
            previous,
        });
    }
    Ok(found)
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

/// Validate every explicit border layer before a table row permutation.
pub(super) fn validate_row_reorder(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar_id: u64,
    rows: u32,
    columns: u32,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<()> {
    planned_row_reorder(
        package,
        locations,
        sidecar_id,
        TableDimensions { rows, columns },
        body_start,
        destinations_by_source,
    )
    .map(drop)
}

/// Keep explicit cell borders attached to their cells during a row permutation.
pub(super) fn reorder_rows(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar_id: u64,
    rows: u32,
    columns: u32,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<()> {
    let updates = planned_row_reorder(
        package,
        locations,
        sidecar_id,
        TableDimensions { rows, columns },
        body_start,
        destinations_by_source,
    )?;
    apply_layer_updates(package, &updates)
}

fn planned_row_reorder(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    sidecar_id: u64,
    dimensions: TableDimensions,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<Vec<LayerUpdate>> {
    let sidecar_archive = locations.get(&sidecar_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers stroke sidecar {sidecar_id} is missing"))
    })?;
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
    let sidecar =
        tst::StrokeSidecarArchive::decode(object.messages[message_index].data.as_slice())?;
    validate_sidecar_dimensions(&sidecar, dimensions)?;

    let mut updates = Vec::new();
    let mut seen_layers = HashSet::new();
    for side in LayerSide::ALL {
        for reference in side.references(&sidecar) {
            if !seen_layers.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers stroke sidecar {sidecar_id} references layer {} more than once",
                    reference.identifier
                )));
            }
            if let Some(update) = plan_reordered_layer(
                package,
                locations,
                reference.identifier,
                side,
                dimensions,
                body_start,
                destinations_by_source,
            )? {
                updates.push(update);
            }
        }
    }
    Ok(updates)
}

fn plan_reordered_layer(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    side: LayerSide,
    dimensions: TableDimensions,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<Option<LayerUpdate>> {
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
    let fixed_length = axis_length(side.fixed_axis(), dimensions);
    let traversal_length = axis_length(opposite_axis(side.fixed_axis()), dimensions);
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
    for run in &previous.stroke_runs {
        validated_run_range(run, traversal_length)?;
    }

    let (current, run_sources) = if side.fixed_axis() == StrokeAxis::Row {
        let mut current = previous.clone();
        current.row_column_index = Some(relocated_row(
            layer_index,
            body_start,
            destinations_by_source,
        )?);
        (current, (0..previous.stroke_runs.len()).map(Some).collect())
    } else {
        let (runs, run_sources) = reorder_runs(
            &previous.stroke_runs,
            dimensions.rows,
            body_start,
            destinations_by_source,
        )?;
        let mut current = previous.clone();
        current.stroke_runs = runs;
        (current, run_sources)
    };

    Ok((current != previous).then_some(LayerUpdate {
        identifier,
        archive_name: archive_name.clone(),
        message_index,
        previous,
        current,
        run_sources,
    }))
}

fn reorder_runs(
    runs: &[tst::stroke_layer_archive::StrokeRunArchive],
    rows: u32,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<(
    Vec<tst::stroke_layer_archive::StrokeRunArchive>,
    Vec<Option<usize>>,
)> {
    let mut reordered = Vec::with_capacity(runs.len());
    let mut sources = Vec::with_capacity(runs.len());
    for (source, run) in runs.iter().enumerate() {
        let (start, end) = validated_run_range(run, rows)?;
        let mut destinations = (start..end)
            .map(|row| relocated_row(row, body_start, destinations_by_source))
            .collect::<Result<Vec<_>>>()?;
        destinations.sort_unstable();
        let Some(&first) = destinations.first() else {
            continue;
        };
        let mut range_start = first;
        let mut previous = first;
        for destination in destinations.into_iter().skip(1) {
            if previous.checked_add(1) == Some(destination) {
                previous = destination;
                continue;
            }
            push_reordered_run(
                &mut reordered,
                &mut sources,
                run,
                source,
                range_start,
                previous,
            )?;
            range_start = destination;
            previous = destination;
        }
        push_reordered_run(
            &mut reordered,
            &mut sources,
            run,
            source,
            range_start,
            previous,
        )?;
    }

    Ok((reordered, sources))
}

fn push_reordered_run(
    runs: &mut Vec<tst::stroke_layer_archive::StrokeRunArchive>,
    sources: &mut Vec<Option<usize>>,
    template: &tst::stroke_layer_archive::StrokeRunArchive,
    source: usize,
    start: u32,
    end: u32,
) -> Result<()> {
    let mut run = template.clone();
    run.origin = Some(checked_i32(start)?);
    run.length = Some(
        end.checked_sub(start)
            .and_then(|length| length.checked_add(1))
            .ok_or_else(|| Error::ParseError("Numbers stroke run length overflow".to_owned()))?,
    );
    runs.push(run);
    sources.push(Some(source));
    Ok(())
}

fn relocated_row(row: u32, body_start: usize, destinations_by_source: &[usize]) -> Result<u32> {
    let row = usize::try_from(row)
        .map_err(|_| Error::ParseError("Numbers stroke row exceeds usize".to_owned()))?;
    let Some(source_offset) = row.checked_sub(body_start) else {
        return u32::try_from(row)
            .map_err(|_| Error::ParseError("Numbers stroke row exceeds u32".to_owned()));
    };
    let Some(destination) = destinations_by_source.get(source_offset) else {
        return u32::try_from(row)
            .map_err(|_| Error::ParseError("Numbers stroke row exceeds u32".to_owned()));
    };
    u32::try_from(*destination)
        .map_err(|_| Error::ParseError("Numbers stroke row destination exceeds u32".to_owned()))
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

    apply_layer_updates(package, &updates)?;

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

fn apply_layer_updates(package: &mut IWorkPackage, updates: &[LayerUpdate]) -> Result<()> {
    for update in updates {
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
    Ok(())
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
        (
            true,
            current,
            (0..previous.stroke_runs.len()).map(Some).collect(),
        )
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
) -> Result<(
    Vec<tst::stroke_layer_archive::StrokeRunArchive>,
    Vec<Option<usize>>,
)> {
    let mut transformed = Vec::with_capacity(runs.len().saturating_add(1));
    let mut sources = Vec::with_capacity(runs.len().saturating_add(1));
    for (source, run) in runs.iter().enumerate() {
        let (start, end) = validated_run_range(run, length)?;
        match edit {
            AxisEdit::Insert { index } if index <= start => {
                let mut shifted = run.clone();
                shifted.origin = Some(checked_i32(indexed_add(start, 1)?)?);
                transformed.push(shifted);
                sources.push(Some(source));
            },
            AxisEdit::Insert { index } if index >= end => {
                transformed.push(run.clone());
                sources.push(Some(source));
            },
            AxisEdit::Insert { index } => {
                let mut before = run.clone();
                before.length = Some(index - start);
                transformed.push(before);
                sources.push(Some(source));

                let mut after = run.clone();
                after.origin = Some(checked_i32(indexed_add(index, 1)?)?);
                after.length = Some(end - index);
                transformed.push(after);
                sources.push(Some(source));
            },
            AxisEdit::Delete { index } if index < start => {
                let mut shifted = run.clone();
                shifted.origin = Some(checked_i32(start - 1)?);
                transformed.push(shifted);
                sources.push(Some(source));
            },
            AxisEdit::Delete { index } if index >= end => {
                transformed.push(run.clone());
                sources.push(Some(source));
            },
            AxisEdit::Delete { .. } if end - start == 1 => {},
            AxisEdit::Delete { .. } => {
                let mut shortened = run.clone();
                shortened.length = Some(end - start - 1);
                transformed.push(shortened);
                sources.push(Some(source));
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
    run_sources: &[Option<usize>],
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
            let Some(source) = source else {
                return Ok(run.encode_to_vec());
            };
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
    data = patch_length_delimited_field(
        &data,
        3,
        previous.stroke.is_some(),
        current
            .stroke
            .as_ref()
            .map(Message::encode_to_vec)
            .as_deref(),
    )?;
    data = patch_varint_field(
        &data,
        4,
        previous.order.is_some(),
        current.order.map(u64::from),
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
    expected.max_order = previous.max_order;
    expected.column_count = previous.column_count;
    expected.row_count = previous.row_count;
    for side in LayerSide::ALL {
        *side.references_mut(&mut expected) = side.references(previous).to_vec();
    }
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers stroke sidecar changed outside its order, dimensions, and layer references"
                .to_owned(),
        ));
    }
    let mut data = patch_varint_field(
        original,
        1,
        previous.max_order.is_some(),
        current.max_order.map(u64::from),
    )?;
    data = patch_varint_field(
        &data,
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
            if let Some(raw) = raw_by_identifier.get(&reference.identifier) {
                if tsp::Reference::decode(*raw)? != *reference {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers stroke sidecar changed layer reference {}",
                        reference.identifier
                    )));
                }
                Ok((*raw).to_vec())
            } else {
                Ok(reference.encode_to_vec())
            }
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
