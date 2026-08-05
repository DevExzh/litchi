//! Inline chart-grid and drawable-geometry codecs.

use super::*;

#[derive(Debug, Clone, Copy)]
enum GridAxis {
    Row,
    Column,
}

impl GridAxis {
    const fn label(self) -> &'static str {
        match self {
            Self::Row => "row",
            Self::Column => "column",
        }
    }

    const fn identifier_offset(self) -> u64 {
        match self {
            Self::Row => 0,
            Self::Column => 1_u64 << 47,
        }
    }
}

pub(crate) fn chart_grid(seed: u64, data: ChartData) -> Result<tsch::ChartGridArchive> {
    let (row_names, column_names, values) = data.into_parts();
    let row_id_map = (0..row_names.len())
        .map(|index| grid_id_entry(seed, index, GridAxis::Row))
        .collect::<Result<_>>()?;
    let column_id_map = (0..column_names.len())
        .map(|index| grid_id_entry(seed, index, GridAxis::Column))
        .collect::<Result<_>>()?;
    Ok(tsch::ChartGridArchive {
        row_name: row_names,
        column_name: column_names,
        grid_row: values
            .into_iter()
            .map(|row| tsch::GridRow {
                value: row
                    .into_iter()
                    .map(|numeric_value| tsch::GridValue {
                        numeric_value,
                        ..Default::default()
                    })
                    .collect(),
            })
            .collect(),
        id_map: Some(tsch::chart_grid_archive::ChartGridRowColumnIdMap {
            row_id_map,
            column_id_map,
        }),
    })
}

pub(crate) fn chart_data(
    application: &str,
    drawable_object_id: u64,
    chart: &tsch::ChartArchive,
) -> Result<ChartData> {
    let grid = chart.grid.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{application} chart {drawable_object_id} has no inline grid"
        ))
    })?;
    ChartData::new(
        grid.row_name.clone(),
        grid.column_name.clone(),
        grid.grid_row
            .iter()
            .map(|row| row.value.iter().map(|value| value.numeric_value).collect())
            .collect(),
    )
}

pub(crate) fn drawable_geometry(
    application: &str,
    drawable_object_id: u64,
    drawable: &tsd::DrawableArchive,
) -> Result<DrawableGeometry> {
    let geometry = drawable.geometry.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{application} chart {drawable_object_id} has no geometry"
        ))
    })?;
    DrawableGeometry {
        position: geometry.position.map(|point| DrawablePoint {
            x: point.x,
            y: point.y,
        }),
        size: geometry.size.map(|size| DrawableSize {
            width: size.width,
            height: size.height,
        }),
        flags: geometry.flags,
        angle: geometry.angle,
    }
    .validate()
}

pub(crate) fn geometry_archive(geometry: DrawableGeometry) -> Result<tsd::GeometryArchive> {
    geometry.validate()?;
    Ok(tsd::GeometryArchive {
        position: geometry.position.map(|point| tsp::Point {
            x: point.x,
            y: point.y,
        }),
        size: geometry.size.map(|size| tsp::Size {
            width: size.width,
            height: size.height,
        }),
        flags: geometry.flags,
        angle: geometry.angle,
    })
}

pub(crate) fn chart_geometry(
    application: &str,
    position: DrawablePoint,
    size: DrawableSize,
) -> Result<DrawableGeometry> {
    if !position.x.is_finite()
        || !position.y.is_finite()
        || !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::ParseError(format!(
            "{application} chart position must be finite and dimensions must be finite and positive"
        )));
    }
    DrawableGeometry {
        position: Some(position),
        size: Some(size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_ROTATION_DEGREES),
    }
    .validate()
}

pub(crate) fn require_creatable_kind(kind: Kind) -> Result<()> {
    if kind == Kind::Undefined || kind.is_unsupported() {
        return Err(Error::ParseError(
            "chart kind must be a supported concrete iWork kind".to_owned(),
        ));
    }
    Ok(())
}

pub(crate) fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn grid_id_entry(
    seed: u64,
    index: usize,
    axis: GridAxis,
) -> Result<tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry> {
    let index_u32 = u32::try_from(index)
        .map_err(|_| Error::ParseError(format!("chart {} index exceeds u32", axis.label())))?;
    let offset = u64::try_from(index)
        .map_err(|_| Error::ParseError(format!("chart {} index exceeds u64", axis.label())))?;
    Ok(
        tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
            unique_id: deterministic_uuid(
                seed.wrapping_add(axis.identifier_offset())
                    .wrapping_add(offset),
            ),
            index: index_u32,
        },
    )
}
