//! Inline chart-grid and drawable-geometry codecs.

use std::mem::size_of;

use litchi_iwa_common::WireLimits;

use super::*;

const MAX_CHART_DATA_CELLS: usize = 1_000_000;
const MAX_CHART_DATA_LABEL_COUNT: usize = 1_000_000;
const MAX_CHART_DATA_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_CHART_DATA_RETAINED_BYTES: usize = WireLimits::MAX_OUTPUT_BYTES;

/// Encode one fresh grid through the shared bounded Buffa writer.
///
/// The host keeps `ChartData` as its compatibility model, while the wire
/// owner emits labels, repeated rows, scalar values, and deterministic ID-map
/// entries directly into one bounded output allocation. No generated
/// `ChartGridArchive`, `GridRow`, or `GridValue` tree is materialized here.
pub(crate) fn chart_grid_bytes(
    seed: u64,
    data: &ChartData,
) -> Result<litchi_iwa_protos::chart_grid_creation_codec::EncodeOutput> {
    let request = litchi_iwa_protos::chart_grid_creation_codec::ChartGridCreationRequest::new(
        data.row_names(),
        data.column_names(),
        data.values(),
        seed,
    );
    let options =
        litchi_iwa_protos::chart_grid_creation_codec::EncodeOptions::for_request(&request);
    litchi_iwa_protos::chart_grid_creation_codec::encode_chart_grid(request, options)
        .map_err(|error| Error::InvalidFormat(format!("chart grid encoding failed: {error}")))
}

/// Decode one modern chart's inline grid through the shared bounded borrowed
/// codec, then materialize the host compatibility model.
pub(crate) fn chart_data_from_source(
    application: &str,
    drawable_object_id: u64,
    source: &[u8],
) -> Result<ChartData> {
    let source_bytes = source.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    let options = litchi_iwa_protos::chart_data_codec::DecodeOptions::new(
        source_bytes,
        source_bytes
            .saturating_mul(8)
            .clamp(1, WireLimits::MAX_FIELDS),
        source_bytes
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
        source_bytes.min(MAX_CHART_DATA_CELLS),
        source_bytes.min(MAX_CHART_DATA_LABEL_COUNT),
        source_bytes.min(MAX_CHART_DATA_TEXT_BYTES),
    );
    let (snapshot, report) = litchi_iwa_protos::chart_data_codec::decode_modern_with_report(
        source, &options,
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "{application} chart {drawable_object_id} has invalid inline grid: {error}"
        ))
    })?;

    let row_count = snapshot.row_count();
    let column_count = snapshot.column_count();
    let cell_count = row_count.checked_mul(column_count).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{application} chart {drawable_object_id} inline grid cell count overflow"
        ))
    })?;
    if cell_count > MAX_CHART_DATA_CELLS {
        return Err(chart_data_limit(
            litchi_iwa_common::LimitKind::MaterializedCells,
            cell_count,
            MAX_CHART_DATA_CELLS,
        ));
    }

    let label_count = row_count.checked_add(column_count).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{application} chart {drawable_object_id} inline grid label count overflow"
        ))
    })?;
    if label_count > MAX_CHART_DATA_LABEL_COUNT {
        return Err(chart_data_limit(
            litchi_iwa_common::LimitKind::Fields,
            label_count,
            MAX_CHART_DATA_LABEL_COUNT,
        ));
    }

    // The strict pass already counted every validated UTF-8 label byte. Reuse
    // that report so materialization does not rescan the borrowed wire view
    // before its work budget has been precharged.
    let text_bytes = report.text_bytes();
    if text_bytes > MAX_CHART_DATA_TEXT_BYTES {
        return Err(chart_data_limit(
            litchi_iwa_common::LimitKind::InputBytes,
            text_bytes,
            MAX_CHART_DATA_TEXT_BYTES,
        ));
    }

    let materialized_bytes = materialized_chart_data_bytes(
        label_count,
        row_count,
        cell_count,
        text_bytes,
    )
    .ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{application} chart {drawable_object_id} inline grid materialized size overflow"
        ))
    })?;
    if materialized_bytes > MAX_CHART_DATA_RETAINED_BYTES {
        return Err(chart_data_limit(
            litchi_iwa_common::LimitKind::OutputBytes,
            materialized_bytes,
            MAX_CHART_DATA_RETAINED_BYTES,
        ));
    }
    let materialization_work = snapshot
        .grid_source()
        .len()
        .checked_mul(10)
        .and_then(|work| work.checked_add(text_bytes))
        .and_then(|work| {
            cell_count
                .checked_mul(size_of::<Option<f64>>())
                .and_then(|cells| work.checked_add(cells))
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{application} chart {drawable_object_id} inline grid materialization work overflow"
            ))
        })?;
    let total_work = report
        .work_bytes()
        .checked_add(materialization_work)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{application} chart {drawable_object_id} inline grid work overflow"
            ))
        })?;
    if total_work > options.max_work_bytes() {
        return Err(chart_data_limit(
            litchi_iwa_common::LimitKind::RewriteWork,
            total_work,
            options.max_work_bytes(),
        ));
    }

    let row_names = own_chart_labels(snapshot.row_labels().iter(), row_count, "chart row labels")?;
    let column_names = own_chart_labels(
        snapshot.column_labels().iter(),
        column_count,
        "chart column labels",
    )?;
    let values = own_chart_values(snapshot.rows(), row_count, column_count)?;
    ChartData::new(row_names, column_names, values).map_err(|error| {
        Error::InvalidFormat(format!(
            "{application} chart {drawable_object_id} has invalid inline grid dimensions: {error}"
        ))
    })
}

fn chart_data_limit(kind: litchi_iwa_common::LimitKind, observed: usize, limit: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
        kind,
        observed,
        limit,
    })
}

fn materialized_chart_data_bytes(
    label_count: usize,
    row_count: usize,
    cell_count: usize,
    text_bytes: usize,
) -> Option<usize> {
    size_of::<ChartData>()
        .checked_add(label_count.checked_mul(size_of::<String>())?)?
        .checked_add(row_count.checked_mul(size_of::<Vec<Option<f64>>>())?)?
        .checked_add(cell_count.checked_mul(size_of::<Option<f64>>())?)?
        .checked_add(text_bytes)
}

fn own_chart_labels<'source>(
    labels: impl Iterator<Item = &'source str>,
    count: usize,
    resource: &'static str,
) -> Result<Vec<String>> {
    let mut owned = Vec::new();
    owned.try_reserve_exact(count).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: count,
        })
    })?;
    for label in labels {
        let mut value = String::new();
        value.try_reserve_exact(label.len()).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "chart label text",
                amount: label.len(),
            })
        })?;
        value.push_str(label);
        owned.push(value);
    }
    if owned.len() != count {
        return Err(Error::InvalidFormat(
            "chart label count changed during materialization".to_owned(),
        ));
    }
    Ok(owned)
}

fn own_chart_values(
    rows: litchi_iwa_protos::chart_data_codec::GridRows<'_>,
    row_count: usize,
    column_count: usize,
) -> Result<Vec<Vec<Option<f64>>>> {
    let mut owned = Vec::new();
    owned.try_reserve_exact(row_count).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "chart value rows",
            amount: row_count,
        })
    })?;
    for row in rows.iter() {
        let mut values = Vec::new();
        values.try_reserve_exact(column_count).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "chart value cells",
                amount: column_count,
            })
        })?;
        for value in row.values() {
            values.push(value);
        }
        if values.len() != column_count {
            return Err(Error::InvalidFormat(
                "chart value row count changed during materialization".to_owned(),
            ));
        }
        owned.push(values);
    }
    if owned.len() != row_count {
        return Err(Error::InvalidFormat(
            "chart value row count changed during materialization".to_owned(),
        ));
    }
    Ok(owned)
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
