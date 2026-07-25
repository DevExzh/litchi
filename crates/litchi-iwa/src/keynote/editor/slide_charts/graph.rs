//! Keynote slide chart graph discovery and validation.

use super::*;
use crate::image_caption::{DrawableCaptionKind, drawable_caption_slot};

pub(super) struct SlideChartGraph {
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) info: KeynoteSlideChartInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) private_preset_id: Option<u64>,
}

pub(super) fn chart_graph(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideChartGraph> {
    let graph = ObjectGraph::read(editor.package())?;
    let context = text_box_create::text_box_context(&graph, slide_index)?;
    for (name, references) in [
        ("owned_drawables", &context.slide.owned_drawables),
        ("drawables_z_order", &context.slide.drawables_z_order),
    ] {
        if references
            .iter()
            .filter(|reference| reference.identifier == drawable_object_id)
            .count()
            != 1
        {
            return Err(Error::ParseError(format!(
                "Keynote slide {} {name} does not own chart {drawable_object_id} exactly once",
                context.slide_id
            )));
        }
    }
    let archive_name = graph.archive_name(context.slide_id)?.to_owned();
    if graph.archive_name(drawable_object_id)? != archive_name {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} is outside slide component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Keynote chart {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Keynote drawable {drawable_object_id} is not exactly one chart"
        )));
    };
    let chart = IWorkChartArchive::decode(&message.data)?;
    let drawable = chart.drawable.super_.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} has no drawable payload"
        ))
    })?;
    if drawable
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(context.slide_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} does not name slide {} as its parent",
            context.slide_id
        )));
    }
    let reference_line_objects = chart_reference_line_objects(&chart)?;
    let payload = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} has no chart payload"
        ))
    })?;
    if payload
        .mediator
        .as_ref()
        .is_some_and(|reference| reference.identifier != 0)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote inline chart {drawable_object_id} unexpectedly has a Numbers mediator"
        )));
    }
    let caption = drawable_caption_slot(
        editor.package(),
        drawable_object_id,
        drawable.caption.as_ref(),
        DrawableCaptionKind::Caption,
        "Keynote chart",
    )?;
    let title_id = required_chart_reference(
        drawable_object_id,
        drawable.title.as_ref(),
        "title stand-in",
    )?;
    let preset_id = payload
        .preset
        .as_ref()
        .map(|reference| reference.identifier);
    let mut object_ids = vec![drawable_object_id];
    object_ids.extend(&caption.object_ids);
    object_ids.push(title_id);
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
    local_styles.extend(
        payload
            .series_private_styles
            .as_ref()
            .into_iter()
            .flat_map(|sparse| {
                sparse.entries.iter().map(|entry| {
                    (
                        entry.reference.identifier,
                        SERIES_STYLE_MESSAGE_TYPE,
                        "private series style",
                    )
                })
            }),
    );
    local_styles.extend(reference_line_objects);
    local_styles.extend(
        payload
            .series_non_styles
            .as_ref()
            .into_iter()
            .flat_map(|sparse| {
                sparse.entries.iter().map(|entry| {
                    (
                        entry.reference.identifier,
                        SERIES_NON_STYLE_MESSAGE_TYPE,
                        "series non-style",
                    )
                })
            }),
    );
    for (identifier, message_type, label) in local_styles {
        if graph.archive_name(identifier)? != archive_name {
            continue;
        }
        let style = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote chart {label} {identifier} is missing"))
        })?;
        if style
            .messages
            .iter()
            .filter(|message| message.type_ == message_type)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote chart {label} {identifier} must have exactly one expected payload"
            )));
        }
        object_ids.push(identifier);
    }
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} aliases private objects"
        )));
    }
    if graph.archive_name(title_id)? != archive_name {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart title stand-in {title_id} is outside {archive_name}"
        )));
    }
    let title = archive.object(title_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote chart title stand-in {title_id} is missing"
        ))
    })?;
    if title
        .messages
        .iter()
        .filter(|message| message.type_ == STANDIN_MESSAGE_TYPE)
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote chart title stand-in {title_id} must have exactly one expected payload"
        )));
    }
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    // App-created native captions can leave part of their chart graph out of
    // the component UUID map. Placeholder-only graphs retain the strict
    // source-built invariant, while native caption graphs keep their actual
    // registered subset for safe duplication and removal.
    if !registered.is_empty()
        && caption.storage_id.is_none()
        && uuid_object_ids.len() != object_ids.len()
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide UUID map does not cover chart {drawable_object_id}"
        )));
    }
    let private_preset_id = preset_id.filter(|identifier| {
        graph
            .archive_name(*identifier)
            .is_ok_and(|name| name == archive_name)
    });
    Ok(SlideChartGraph {
        archive_name,
        component_id,
        info: KeynoteSlideChartInfo {
            slide_index,
            slide_id: context.slide_id,
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
            data: chart_data("Keynote", drawable_object_id, payload)?,
            geometry: drawable_geometry("Keynote", drawable_object_id, drawable)?,
            arrangement: ChartArrangement::new(
                drawable.locked.unwrap_or(false),
                drawable.aspect_ratio_locked.unwrap_or(false),
            ),
        },
        object_ids,
        uuid_object_ids,
        private_preset_id,
    })
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
            Error::InvalidFormat(format!("Keynote chart {drawable_object_id} has no {label}"))
        })
}
