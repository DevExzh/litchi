//! Numbers sheet chart graph discovery and validation.

use super::*;

pub(super) struct SheetChartGraph {
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) info: NumbersSheetChartInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) private_preset_id: Option<u64>,
}

pub(super) fn chart_graph(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<SheetChartGraph> {
    let (archive_name, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own chart {drawable_object_id} exactly once"
        )));
    }
    let locations = object_locations(editor.package())?;
    if locations.get(&drawable_object_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers chart {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not exactly one chart"
        )));
    };
    let chart = IWorkChartArchive::decode(&message.data)?;
    let drawable = chart.drawable.super_.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no drawable payload"
        ))
    })?;
    if drawable
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(sheet_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} does not name sheet {sheet_id} as its parent"
        )));
    }
    let payload = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no chart payload"
        ))
    })?;
    let caption_id = required_chart_reference(
        drawable_object_id,
        drawable.caption.as_ref(),
        "caption stand-in",
    )?;
    let title_id = required_chart_reference(
        drawable_object_id,
        drawable.title.as_ref(),
        "title stand-in",
    )?;
    let mediator_id =
        required_chart_reference(drawable_object_id, payload.mediator.as_ref(), "mediator")?;
    let preset_id = payload
        .preset
        .as_ref()
        .map(|reference| reference.identifier);
    let mut object_ids = vec![drawable_object_id, caption_id, title_id, mediator_id];
    let mut local_styles = Vec::new();
    local_styles.extend(
        payload
            .preset
            .map(|reference| (reference.identifier, CHART_PRESET_MESSAGE_TYPE, "preset")),
    );
    local_styles.extend(payload.chart_style.map(|reference| {
        (
            reference.identifier,
            CHART_STYLE_MESSAGE_TYPE,
            "chart style",
        )
    }));
    local_styles.extend(payload.chart_non_style.map(|reference| {
        (
            reference.identifier,
            CHART_NON_STYLE_MESSAGE_TYPE,
            "chart non-style",
        )
    }));
    local_styles.extend(payload.legend_style.map(|reference| {
        (
            reference.identifier,
            LEGEND_STYLE_MESSAGE_TYPE,
            "legend style",
        )
    }));
    local_styles.extend(payload.legend_non_style.map(|reference| {
        (
            reference.identifier,
            LEGEND_NON_STYLE_MESSAGE_TYPE,
            "legend non-style",
        )
    }));
    local_styles.extend(payload.value_axis_styles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_STYLE_MESSAGE_TYPE,
            "value-axis style",
        )
    }));
    local_styles.extend(payload.value_axis_nonstyles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_NON_STYLE_MESSAGE_TYPE,
            "value-axis non-style",
        )
    }));
    local_styles.extend(payload.category_axis_styles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_STYLE_MESSAGE_TYPE,
            "category-axis style",
        )
    }));
    local_styles.extend(payload.category_axis_nonstyles.iter().map(|reference| {
        (
            reference.identifier,
            AXIS_NON_STYLE_MESSAGE_TYPE,
            "category-axis non-style",
        )
    }));
    local_styles.extend(payload.series_theme_styles.iter().map(|reference| {
        (
            reference.identifier,
            SERIES_STYLE_MESSAGE_TYPE,
            "series style",
        )
    }));
    for (identifier, message_type, label) in local_styles {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            continue;
        }
        let style = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers chart {label} {identifier} is missing"))
        })?;
        if style
            .messages
            .iter()
            .filter(|message| message.type_ == message_type)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers chart {label} {identifier} must have exactly one expected payload"
            )));
        }
        object_ids.push(identifier);
    }
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} aliases private objects"
        )));
    }
    for (identifier, message_type, label) in [
        (caption_id, STANDIN_MESSAGE_TYPE, "caption stand-in"),
        (title_id, STANDIN_MESSAGE_TYPE, "title stand-in"),
        (mediator_id, CHART_MEDIATOR_MESSAGE_TYPE, "mediator"),
    ] {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers chart {label} {identifier} is outside {archive_name}"
            )));
        }
        let private = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers chart {label} {identifier} is missing"))
        })?;
        if private
            .messages
            .iter()
            .filter(|message| message.type_ == message_type)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers chart {label} {identifier} must have exactly one expected payload"
            )));
        }
    }
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    if !registered.is_empty() && uuid_object_ids.len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers component {component_id} UUID map does not cover chart {drawable_object_id}"
        )));
    }
    let private_preset_id = preset_id.filter(|identifier| {
        locations.get(identifier).map(String::as_str) == Some(archive_name.as_str())
    });
    Ok(SheetChartGraph {
        archive_name,
        component_id,
        info: NumbersSheetChartInfo {
            sheet_id,
            drawable_object_id,
            kind: ChartKind::from_raw(
                payload
                    .chart_type
                    .unwrap_or(tsch::ChartType::UndefinedChartType as i32),
            ),
            direction: ChartSeriesDirection::from_raw(
                payload
                    .series_direction
                    .unwrap_or(tsch::SeriesDirection::Unknown as i32),
            ),
            data: chart_data(drawable_object_id, payload)?,
            geometry: drawable_geometry(drawable_object_id, drawable)?,
        },
        object_ids,
        uuid_object_ids,
        private_preset_id,
    })
}

fn chart_data(drawable_object_id: u64, chart: &tsch::ChartArchive) -> Result<ChartData> {
    let grid = chart.grid.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no inline grid"
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

fn drawable_geometry(
    drawable_object_id: u64,
    drawable: &tsd::DrawableArchive,
) -> Result<DrawableGeometry> {
    let geometry = drawable.geometry.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no geometry"
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

fn required_chart_reference(
    drawable_object_id: u64,
    reference: Option<&tsp::Reference>,
    label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers chart {drawable_object_id} has no {label}"))
        })
}
