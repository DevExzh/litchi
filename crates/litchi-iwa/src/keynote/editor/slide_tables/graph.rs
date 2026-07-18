//! Slide ownership and native table graph discovery.

use prost::Message;

use super::*;
use crate::protobuf::tst::{TableInfoArchive, TableModelArchive};

#[derive(Debug, Clone)]
pub(super) struct SlideTableGraph {
    pub(super) info: KeynoteSlideTableInfo,
    pub(super) slide_archive: String,
    pub(super) slide_component_id: u64,
}

pub(super) fn require_table_model(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<KeynoteSlideTableInfo> {
    editor
        .slide_tables(slide_index)?
        .into_iter()
        .find(|table| table.model_object_id == model_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote table model {model_object_id} is not owned by slide {slide_index}"
            ))
        })
}

pub(super) fn slide_table_graph(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideTableGraph> {
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
                "Keynote slide {} {name} does not own table {drawable_object_id} exactly once",
                context.slide_id
            )));
        }
    }
    let table_info = graph.decode_type::<TableInfoArchive>(
        drawable_object_id,
        TABLE_INFO_MESSAGE_TYPE,
        "TableInfoArchive",
    )?;
    if table_info
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(context.slide_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table {drawable_object_id} does not name slide {} as its parent",
            context.slide_id
        )));
    }
    let model_id = table_info.table_model.identifier;
    let model = decode_table_model(&graph, model_id)?;
    let slide_archive = graph.archive_name(context.slide_id)?.to_owned();
    let slide_component_id = component_identifier_for_entry(editor.package(), &slide_archive)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {slide_archive} is not registered"
            ))
        })?;
    Ok(SlideTableGraph {
        info: KeynoteSlideTableInfo {
            slide_index,
            slide_id: context.slide_id,
            drawable_object_id,
            model_object_id: model_id,
            name: model.table_name,
            rows: model.number_of_rows as usize,
            columns: model.number_of_columns as usize,
            geometry: crate::shapes::geometry_from_drawable(&table_info.super_)?,
        },
        slide_archive,
        slide_component_id,
    })
}

fn decode_table_model(graph: &ObjectGraph, model_id: u64) -> Result<TableModelArchive> {
    let messages = graph.objects.get(&model_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Keynote table model {model_id} is missing"))
    })?;
    let models = messages
        .iter()
        .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
        .filter_map(|message| TableModelArchive::decode(message.data.as_slice()).ok())
        .collect::<Vec<_>>();
    let [model] = models.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote table model {model_id} must contain exactly one table-model payload"
        )));
    };
    Ok(model.clone())
}

pub(super) fn table_template(package: &IWorkPackage) -> Result<(u64, u64)> {
    let graph = ObjectGraph::read(package)?;
    let mut candidates = graph.objects.keys().copied().collect::<Vec<_>>();
    candidates.sort_unstable();
    for info_id in candidates {
        let Some(messages) = graph.objects.get(&info_id) else {
            continue;
        };
        for message in messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        {
            let Ok(info) = TableInfoArchive::decode(message.data.as_slice()) else {
                continue;
            };
            let model_id = info.table_model.identifier;
            if model_id != 0 && decode_table_model(&graph, model_id).is_ok() {
                return Ok((info_id, model_id));
            }
        }
    }
    Err(Error::InvalidFormat(
        "Keynote package has no native table creation template".to_owned(),
    ))
}
