//! Strict discovery and construction of body-anchored Pages chart graphs.

use super::*;
use crate::image_caption::{DrawableCaptionKind, drawable_caption_slot};
use crate::package_metadata::component_identifier_for_object_uuid;

const DRAWABLE_Z_ORDER_MESSAGE_TYPE: u32 = 10_015;
const ATTACHMENT_HORIZONTAL_OFFSET_FIELD: u32 = 3;
const ATTACHMENT_VERTICAL_OFFSET_FIELD: u32 = 5;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum HorizontalAnchorBasis {
    BodyMargin = 0,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum VerticalAnchorBasis {
    Page = 1,
}

pub(super) struct BodyChartGraph {
    pub(super) archive_name: String,
    pub(super) archive_groups: Vec<BodyChartArchiveGroup>,
    pub(super) attachment_id: u64,
    pub(super) component_id: u64,
    pub(super) info: PagesBodyChartInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) private_preset_id: Option<u64>,
}

pub(super) struct BodyChartArchiveGroup {
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) object_ids: Vec<u64>,
    pub(super) style_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
}

pub(super) fn body_chart_infos(editor: &PagesEditor) -> Result<Vec<PagesBodyChartInfo>> {
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let mut charts = Vec::new();
    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
    {
        let Some(attachment_reference) = entry.object else {
            continue;
        };
        if !object_has_message_type(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )? {
            continue;
        }
        let attachment: DrawableAttachmentArchive = decode_typed_package_object(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            "TSWP.DrawableAttachmentArchive",
        )?;
        let Some(drawable) = attachment.drawable else {
            continue;
        };
        if !object_has_message_type(editor.package(), drawable.identifier, CHART_MESSAGE_TYPE)? {
            continue;
        }
        let graph = body_chart_graph(editor, drawable.identifier)?;
        if graph.info.anchor_character_index != entry.character_index {
            return Err(Error::InvalidFormat(format!(
                "Pages chart {} attachment index changed during discovery",
                drawable.identifier
            )));
        }
        charts.push(graph.info);
    }
    Ok(charts)
}

pub(super) fn body_chart_graph(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<BodyChartGraph> {
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let mut attachments = Vec::new();
    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
    {
        let Some(reference) = entry.object else {
            continue;
        };
        if !object_has_message_type(
            editor.package(),
            reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )? {
            continue;
        }
        let attachment: DrawableAttachmentArchive = decode_typed_package_object(
            editor.package(),
            reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            "TSWP.DrawableAttachmentArchive",
        )?;
        if attachment
            .drawable
            .is_some_and(|drawable| drawable.identifier == drawable_object_id)
        {
            attachments.push((entry.character_index, reference.identifier));
        }
    }
    let [(anchor_character_index, attachment_id)] = attachments.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} has {} body attachments; expected one",
            attachments.len()
        )));
    };
    let body_units = editor.body_text()?.encode_utf16().collect::<Vec<_>>();
    if body_units.get(*anchor_character_index as usize) != Some(&0xfffc) {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} attachment is not backed by an object-replacement character"
        )));
    }

    let document = root_document(editor.package())?;
    let z_order_id = document.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".into())
    })?;
    let z_order: tp::DrawablesZOrderArchive = decode_typed_package_object(
        editor.package(),
        z_order_id.identifier,
        DRAWABLE_Z_ORDER_MESSAGE_TYPE,
        "TP.DrawablesZOrderArchive",
    )?;
    if z_order
        .drawables
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Pages drawable z-order does not own chart {drawable_object_id} exactly once"
        )));
    }

    let archive_name = find_object_archive(editor.package(), drawable_object_id)?;
    if find_object_archive(editor.package(), *attachment_id)? != archive_name {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} and attachment {attachment_id} are in different components"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages chart {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages drawable {drawable_object_id} is not exactly one chart"
        )));
    };
    let chart = IWorkChartArchive::decode(&message.data)?;
    let drawable = chart.drawable.super_.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} has no drawable payload"
        ))
    })?;
    if drawable
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(editor.body_storage_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} does not name body storage {} as its parent",
            editor.body_storage_id
        )));
    }
    let reference_line_objects = chart_reference_line_objects(&chart)?;
    let payload = chart.chart.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} has no chart payload"
        ))
    })?;
    if payload
        .mediator
        .as_ref()
        .is_some_and(|reference| reference.identifier != 0)
    {
        return Err(Error::InvalidFormat(format!(
            "Pages inline chart {drawable_object_id} unexpectedly has a Numbers mediator"
        )));
    }
    let caption = drawable_caption_slot(
        editor.package(),
        drawable_object_id,
        drawable.caption.as_ref(),
        DrawableCaptionKind::Caption,
        "Pages chart",
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
    let private_preset_id = payload
        .owned_preset
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0);
    if private_preset_id.is_some() && private_preset_id != preset_id {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} has inconsistent owned and active presets"
        )));
    }

    let mut object_ids = vec![drawable_object_id];
    object_ids.extend(&caption.object_ids);
    object_ids.push(title_id);
    let mut style_ids = HashSet::new();
    if private_preset_id.is_some() {
        let reference_owner_counts = chart_reference_owner_counts(editor.package())?;
        if let Some(preset_id) = private_preset_id {
            if find_object_archive(editor.package(), preset_id)? != archive_name {
                return Err(Error::InvalidFormat(format!(
                    "Pages private chart preset {preset_id} is outside {archive_name}"
                )));
            }
            let preset = archive.object(preset_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Pages chart preset {preset_id} is missing"))
            })?;
            if preset
                .messages
                .iter()
                .filter(|message| message.type_ == CHART_PRESET_MESSAGE_TYPE)
                .count()
                != 1
                || reference_owner_counts.get(&preset_id) != Some(&1)
            {
                return Err(Error::InvalidFormat(format!(
                    "Pages private chart preset {preset_id} must have one payload and one chart owner"
                )));
            }
            object_ids.push(preset_id);
        }
        let mut private_styles = Vec::new();
        private_styles.extend(payload.chart_style.map(|reference| {
            (
                reference.identifier,
                CHART_STYLE_MESSAGE_TYPE,
                "chart style",
            )
        }));
        private_styles.extend(payload.chart_non_style.map(|reference| {
            (
                reference.identifier,
                CHART_NON_STYLE_MESSAGE_TYPE,
                "chart non-style",
            )
        }));
        private_styles.extend(payload.legend_style.map(|reference| {
            (
                reference.identifier,
                LEGEND_STYLE_MESSAGE_TYPE,
                "legend style",
            )
        }));
        private_styles.extend(payload.legend_non_style.map(|reference| {
            (
                reference.identifier,
                LEGEND_NON_STYLE_MESSAGE_TYPE,
                "legend non-style",
            )
        }));
        private_styles.extend(payload.value_axis_styles.iter().map(|reference| {
            (
                reference.identifier,
                AXIS_STYLE_MESSAGE_TYPE,
                "value-axis style",
            )
        }));
        private_styles.extend(payload.value_axis_nonstyles.iter().map(|reference| {
            (
                reference.identifier,
                AXIS_NON_STYLE_MESSAGE_TYPE,
                "value-axis non-style",
            )
        }));
        private_styles.extend(payload.category_axis_styles.iter().map(|reference| {
            (
                reference.identifier,
                AXIS_STYLE_MESSAGE_TYPE,
                "category-axis style",
            )
        }));
        private_styles.extend(payload.category_axis_nonstyles.iter().map(|reference| {
            (
                reference.identifier,
                AXIS_NON_STYLE_MESSAGE_TYPE,
                "category-axis non-style",
            )
        }));
        private_styles.extend(payload.series_theme_styles.iter().map(|reference| {
            (
                reference.identifier,
                SERIES_STYLE_MESSAGE_TYPE,
                "series style",
            )
        }));
        private_styles.extend(payload.series_private_styles.as_ref().into_iter().flat_map(
            |sparse| {
                sparse.entries.iter().map(|entry| {
                    (
                        entry.reference.identifier,
                        SERIES_STYLE_MESSAGE_TYPE,
                        "private series style",
                    )
                })
            },
        ));
        private_styles.extend(
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
        private_styles.extend(reference_line_objects);
        let mut unique_styles = HashMap::new();
        for (identifier, message_type, label) in private_styles {
            if let Some((existing_type, existing_label)) =
                unique_styles.insert(identifier, (message_type, label))
            {
                if existing_type != message_type {
                    return Err(Error::InvalidFormat(format!(
                        "Pages chart object {identifier} is both {existing_label} and {label}"
                    )));
                }
                continue;
            }
            let style_archive_name = find_object_archive(editor.package(), identifier)?;
            let style_archive = editor.package().archive(&style_archive_name)?;
            let style = style_archive.object(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!("Pages chart {label} {identifier} is missing"))
            })?;
            if style
                .messages
                .iter()
                .filter(|message| message.type_ == message_type)
                .count()
                != 1
            {
                return Err(Error::InvalidFormat(format!(
                    "Pages chart {label} {identifier} must have exactly one expected payload"
                )));
            }
            if reference_owner_counts.get(&identifier) == Some(&1) {
                object_ids.push(identifier);
                if matches!(
                    message_type,
                    CHART_STYLE_MESSAGE_TYPE
                        | CHART_NON_STYLE_MESSAGE_TYPE
                        | LEGEND_STYLE_MESSAGE_TYPE
                        | LEGEND_NON_STYLE_MESSAGE_TYPE
                        | AXIS_STYLE_MESSAGE_TYPE
                        | AXIS_NON_STYLE_MESSAGE_TYPE
                        | SERIES_STYLE_MESSAGE_TYPE
                        | SERIES_NON_STYLE_MESSAGE_TYPE
                ) {
                    style_ids.insert(identifier);
                }
            }
        }
    }
    if find_object_archive(editor.package(), title_id)? != archive_name {
        return Err(Error::InvalidFormat(format!(
            "Pages chart title stand-in {title_id} is outside {archive_name}"
        )));
    }
    let title = archive.object(title_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages chart title stand-in {title_id} is missing"))
    })?;
    if title
        .messages
        .iter()
        .filter(|message| message.type_ == STANDIN_MESSAGE_TYPE)
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Pages chart title stand-in {title_id} must have exactly one stand-in payload"
        )));
    }
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} aliases private objects"
        )));
    }
    object_ids.push(*attachment_id);

    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages chart component {archive_name} is not registered"
            ))
        })?;
    let mut archive_groups: Vec<BodyChartArchiveGroup> = Vec::new();
    let mut registered_uuid_count = 0usize;
    for &identifier in &object_ids {
        let object_archive_name = find_object_archive(editor.package(), identifier)?;
        if !style_ids.contains(&identifier) && object_archive_name != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Pages private chart object {identifier} is outside {archive_name}"
            )));
        }
        let object_component_id =
            component_identifier_for_entry(editor.package(), &object_archive_name)?.ok_or_else(
                || {
                    Error::InvalidFormat(format!(
                        "Pages chart object component {object_archive_name} is not registered"
                    ))
                },
            )?;
        let group_index = archive_groups
            .iter()
            .position(|group| group.archive_name == object_archive_name)
            .unwrap_or_else(|| {
                archive_groups.push(BodyChartArchiveGroup {
                    archive_name: object_archive_name.clone(),
                    component_id: object_component_id,
                    object_ids: Vec::new(),
                    style_ids: Vec::new(),
                    uuid_object_ids: Vec::new(),
                });
                archive_groups.len() - 1
            });
        let group = &mut archive_groups[group_index];
        if group.component_id != object_component_id {
            return Err(Error::InvalidFormat(format!(
                "Pages chart archive {} resolves to inconsistent components",
                group.archive_name
            )));
        }
        group.object_ids.push(identifier);
        if style_ids.contains(&identifier) {
            group.style_ids.push(identifier);
        }
        if identifier == *attachment_id {
            continue;
        }
        if let Some(uuid_component_id) =
            component_identifier_for_object_uuid(editor.package(), identifier)?
        {
            if uuid_component_id != object_component_id {
                return Err(Error::InvalidFormat(format!(
                    "Pages chart object {identifier} is stored in component {object_component_id} but registered in component {uuid_component_id}"
                )));
            }
            group.uuid_object_ids.push(identifier);
            registered_uuid_count = registered_uuid_count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Pages chart UUID count overflow".to_owned())
            })?;
        }
    }
    // App-created native captions can leave part of their chart graph out of
    // the component UUID map. Placeholder-only graphs retain the strict
    // source-built invariant, while native caption graphs keep their actual
    // registered subset for safe duplication and removal.
    if private_preset_id.is_some()
        && caption.storage_id.is_none()
        && registered_uuid_count + 1 != object_ids.len()
    {
        return Err(Error::InvalidFormat(format!(
            "Pages component UUID map does not cover private chart {drawable_object_id}"
        )));
    }
    let theme_id = document
        .theme
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| Error::InvalidFormat("Pages document has no theme".into()))?;
    let theme = chart_theme_context(editor.package(), theme_id)?;
    for group in &archive_groups {
        if !group.style_ids.is_empty() {
            validate_chart_styles_registered(
                editor.package(),
                theme.stylesheet_id,
                &group.archive_name,
                &group.style_ids,
            )?;
        }
    }

    Ok(BodyChartGraph {
        archive_name,
        archive_groups,
        attachment_id: *attachment_id,
        component_id,
        info: PagesBodyChartInfo {
            anchor_character_index: *anchor_character_index,
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
            data: chart_data("Pages", drawable_object_id, payload)?,
            geometry: drawable_geometry("Pages", drawable_object_id, drawable)?,
            arrangement: ChartArrangement::new(
                drawable.locked.unwrap_or(false),
                drawable.aspect_ratio_locked.unwrap_or(false),
            ),
        },
        object_ids,
        private_preset_id,
    })
}

fn chart_reference_owner_counts(package: &IWorkPackage) -> Result<HashMap<u64, usize>> {
    let mut counts = HashMap::new();
    for archive_name in package.iwa_entry_names() {
        for object in &package.archive(archive_name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
            {
                let chart = IWorkChartArchive::decode(message.data.as_slice())?;
                for identifier in chart.typed_reference_identifiers()? {
                    let count = counts.entry(identifier).or_insert(0usize);
                    *count = count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Pages chart owner count overflow".to_owned())
                    })?;
                }
            }
        }
    }
    Ok(counts)
}

pub(super) fn chart_attachment_object(
    attachment_id: u64,
    drawable_id: u64,
    position: DrawablePoint,
    left_margin: f32,
) -> Result<ArchiveObject> {
    if !left_margin.is_finite() {
        return Err(Error::ParseError("Pages left margin must be finite".into()));
    }
    let attachment = DrawableAttachmentArchive {
        drawable: Some(reference(drawable_id)),
        h_offset_type: Some(HorizontalAnchorBasis::BodyMargin as u32),
        h_offset: Some(position.x - left_margin),
        v_offset_type: Some(VerticalAnchorBasis::Page as u32),
        v_offset: Some(position.y),
    };
    let mut object = ArchiveObject::new(
        attachment_id,
        vec![RawMessage {
            type_: DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            data: attachment.encode_to_vec(),
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(drawable_id);
    Ok(object)
}

pub(super) fn set_chart_attachment_position(
    package: &mut IWorkPackage,
    archive_name: &str,
    attachment_id: u64,
    position: DrawablePoint,
    left_margin: f32,
) -> Result<()> {
    if !position.x.is_finite() || !position.y.is_finite() || !left_margin.is_finite() {
        return Err(Error::ParseError(
            "Pages chart attachment position and left margin must be finite".into(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(attachment_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages chart attachment {attachment_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == DRAWABLE_ATTACHMENT_MESSAGE_TYPE).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages chart attachment {attachment_id} must have exactly one payload"
            )));
        };
        let data = patch_fixed32_field(
            &object.messages[*message_index].data,
            ATTACHMENT_HORIZONTAL_OFFSET_FIELD,
            true,
            Some((position.x - left_margin).to_bits()),
        )?;
        let data = patch_fixed32_field(
            &data,
            ATTACHMENT_VERTICAL_OFFSET_FIELD,
            true,
            Some(position.y.to_bits()),
        )?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
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
            Error::InvalidFormat(format!("Pages chart {drawable_object_id} has no {label}"))
        })
}

fn object_has_message_type(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
) -> Result<bool> {
    let archive_name = find_object_archive(package, identifier)?;
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Pages object {identifier} is missing")))?;
    Ok(object
        .messages
        .iter()
        .any(|message| message.type_ == message_type))
}
