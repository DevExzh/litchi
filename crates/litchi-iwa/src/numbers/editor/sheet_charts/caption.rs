//! Native caption CRUD for Numbers sheet charts.

use super::*;
use crate::charts::caption_edge::{chart_caption_identifier, rewrite_chart_caption_identifier};
use crate::image_caption::{
    CaptionObjectIds, CaptionThemeStyle, DrawableCaptionKind, DrawableCaptionSlot, caption_objects,
    componentized_caption_objects, drawable_caption_slot, insert_componentized_caption_style,
    replace_object_reference, standin_caption_object,
};

const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ChartCaptionGraphLayout {
    Inline,
    Componentized,
}

impl NumbersEditor {
    /// Read the native caption attached to one sheet chart.
    pub fn sheet_chart_caption(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<String>> {
        sheet_chart_caption(self, sheet_id, drawable_object_id)
    }

    /// Create or replace the native caption attached to one sheet chart.
    pub fn set_sheet_chart_caption(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        caption: &str,
    ) -> Result<()> {
        set_sheet_chart_caption(self, sheet_id, drawable_object_id, caption)
    }

    /// Remove the native caption attached to one sheet chart.
    ///
    /// Returns whether a caption was present. Native iWork removal preserves
    /// the prior caption graph for undo history and attaches a fresh empty
    /// stand-in.
    pub fn remove_sheet_chart_caption(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        remove_sheet_chart_caption(self, sheet_id, drawable_object_id)
    }
}

fn sheet_chart_caption(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<Option<String>> {
    let slot = sheet_chart_caption_slot(editor, sheet_id, drawable_object_id)?;
    slot.storage_id
        .map(|storage_id| {
            IWorkTextEditor::from_package(editor.package.clone())
                .storage(crate::text::native_storage_id(storage_id)?)
                .map(|storage| storage.storage.into_text())
        })
        .transpose()
}

pub(super) fn sheet_chart_caption_slot(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<DrawableCaptionSlot> {
    let source = chart_graph(editor, sheet_id, drawable_object_id)?;
    let archive = editor.package.archive(&source.archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers chart {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    let caption_identifier =
        chart_caption_identifier(editor.package.limits(), message.data.as_slice())?;
    drawable_caption_slot(
        &editor.package,
        drawable_object_id,
        caption_identifier,
        DrawableCaptionKind::Caption,
        "Numbers chart",
    )
}

fn set_sheet_chart_caption(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    text: &str,
) -> Result<()> {
    let source = chart_graph(editor, sheet_id, drawable_object_id)?;
    let slot = sheet_chart_caption_slot(editor, sheet_id, drawable_object_id)?;
    let expected = Some(text.to_owned());
    let staged = if let Some(storage_id) = slot.storage_id {
        let mut text_editor = IWorkTextEditor::from_package(editor.package.clone());
        text_editor.set_text(crate::text::native_storage_id(storage_id)?, text)?;
        text_editor.into_package()
    } else {
        let (theme, language) = sheet_chart_caption_theme(editor)?;
        let drawable_width = source
            .info
            .geometry
            .size
            .ok_or_else(|| Error::InvalidFormat("Numbers chart has no displayed size".to_owned()))?
            .width;
        let ids = CaptionObjectIds::allocate(next_object_identifier(&editor.package)?)?;
        let layout = chart_caption_graph_layout(&editor.package, &source.archive_name)?;
        let mut staged = editor.package.clone();
        insert_sheet_chart_caption(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            slot.reference_id,
            drawable_width,
            text,
            theme,
            language.as_deref(),
            ids,
            layout,
        )?;
        match layout {
            ChartCaptionGraphLayout::Inline => {
                add_component_object_uuids(&mut staged, source.component_id, &ids.all())?;
            },
            ChartCaptionGraphLayout::Componentized => {
                add_component_object_uuids(
                    &mut staged,
                    source.component_id,
                    &ids.component_objects(),
                )?;
            },
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;
        staged
    };
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_caption(sheet_id, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Numbers chart caption update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_sheet_chart_caption(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<bool> {
    let source = chart_graph(editor, sheet_id, drawable_object_id)?;
    let slot = sheet_chart_caption_slot(editor, sheet_id, drawable_object_id)?;
    if slot.storage_id.is_none() {
        return Ok(false);
    }
    let standin_id = next_object_identifier(&editor.package)?;
    let mut staged = editor.package.clone();
    insert_sheet_chart_caption_standin(
        &mut staged,
        &source.archive_name,
        drawable_object_id,
        slot.reference_id,
        standin_id,
    )?;
    add_component_object_uuids(&mut staged, source.component_id, &[standin_id])?;
    set_package_last_object_identifier(&mut staged, standin_id)?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified
        .sheet_chart_caption(sheet_id, drawable_object_id)?
        .is_some()
    {
        return Err(Error::InvalidFormat(
            "Numbers chart caption removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}

fn sheet_chart_caption_theme(
    editor: &NumbersEditor,
) -> Result<(CaptionThemeStyle, Option<String>)> {
    let document = numbers_document(&editor.package)?;
    let theme_id = document.theme.identifier;
    let stylesheet_id = document.stylesheet.identifier;
    let locations = object_locations(&editor.package)?;
    let archive_name = locations
        .get(&theme_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers theme {theme_id} is missing")))?;
    let archive = editor.package.archive(archive_name)?;
    let object = archive
        .object(theme_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers theme {theme_id} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBERS_THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers theme {theme_id} must have exactly one theme payload"
        )));
    };
    let paragraph_style_id = IWorkThemeArchive::decode(message.data.as_slice())?
        .extensions
        .application
        .ok_or_else(|| Error::InvalidFormat("Numbers theme has no application presets".to_owned()))?
        .caption_style_presets
        .into_iter()
        .next()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers theme has no caption style preset".to_owned())
        })?;
    for (identifier, label) in [
        (stylesheet_id, "stylesheet"),
        (paragraph_style_id, "caption paragraph style"),
    ] {
        if !locations.contains_key(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {label} object {identifier} is missing"
            )));
        }
    }
    Ok((
        CaptionThemeStyle {
            stylesheet_id,
            paragraph_style_id,
        },
        document.super_.document_language,
    ))
}

fn insert_sheet_chart_caption(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    old_reference_id: u64,
    drawable_width: f32,
    text: &str,
    theme: CaptionThemeStyle,
    language: Option<&str>,
    ids: CaptionObjectIds,
    layout: ChartCaptionGraphLayout,
) -> Result<()> {
    match layout {
        ChartCaptionGraphLayout::Inline => {
            let objects = caption_objects(
                ids,
                drawable_object_id,
                drawable_width,
                text,
                DrawableCaptionKind::Caption,
                theme,
                language,
            )?;
            let codec_limits = package.limits();
            package.update_archive(archive_name, |archive| {
                for object in objects {
                    archive.insert_object(object)?;
                }
                retarget_sheet_chart_caption_edge(
                    codec_limits,
                    archive,
                    drawable_object_id,
                    old_reference_id,
                    ids.info,
                )
            })
        },
        ChartCaptionGraphLayout::Componentized => {
            let [style, info, storage, placement] = componentized_caption_objects(
                ids,
                drawable_object_id,
                drawable_width,
                text,
                DrawableCaptionKind::Caption,
                theme,
                language,
            )?;
            insert_componentized_caption_style(package, archive_name, theme.stylesheet_id, style)?;
            let codec_limits = package.limits();
            package.update_archive(archive_name, |archive| {
                for object in [info, storage, placement] {
                    archive.insert_object(object)?;
                }
                retarget_sheet_chart_caption_edge(
                    codec_limits,
                    archive,
                    drawable_object_id,
                    old_reference_id,
                    ids.info,
                )
            })
        },
    }
}

fn chart_caption_graph_layout(
    package: &IWorkPackage,
    chart_archive_name: &str,
) -> Result<ChartCaptionGraphLayout> {
    let componentized = package
        .archive(chart_archive_name)?
        .objects
        .iter()
        .any(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE)
        });
    Ok(if componentized {
        ChartCaptionGraphLayout::Componentized
    } else {
        ChartCaptionGraphLayout::Inline
    })
}

fn insert_sheet_chart_caption_standin(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    old_reference_id: u64,
    standin_id: u64,
) -> Result<()> {
    let standin = standin_caption_object(standin_id)?;
    let codec_limits = package.limits();
    package.update_archive(archive_name, |archive| {
        archive.insert_object(standin)?;
        retarget_sheet_chart_caption_edge(
            codec_limits,
            archive,
            drawable_object_id,
            old_reference_id,
            standin_id,
        )
    })
}

fn retarget_sheet_chart_caption_edge(
    limits: crate::package::PackageLimits,
    archive: &mut crate::archive::Archive,
    drawable_object_id: u64,
    old_reference_id: u64,
    replacement_id: u64,
) -> Result<()> {
    let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers chart {drawable_object_id} is missing"))
    })?;
    let message_indexes = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| (message.type_ == CHART_MESSAGE_TYPE).then_some(index))
        .collect::<Vec<_>>();
    let [message_index] = message_indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    let original = object.messages[*message_index].data.as_slice();
    let current_reference_id = chart_caption_identifier(limits, original)?;
    if current_reference_id != Some(old_reference_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} caption reference changed unexpectedly"
        )));
    }
    let data = rewrite_chart_caption_identifier(limits, original, replacement_id)?;
    let actual_reference_id = chart_caption_identifier(limits, data.as_slice())?;
    if actual_reference_id != Some(replacement_id) {
        return Err(Error::InvalidFormat(
            "Numbers chart caption reference patch failed validation".to_owned(),
        ));
    }
    object.replace_message(
        *message_index,
        RawMessage {
            type_: CHART_MESSAGE_TYPE,
            data,
        },
    )?;
    replace_object_reference(
        &mut object.archive_info.message_infos[*message_index].object_references,
        old_reference_id,
        replacement_id,
    );
    Ok(())
}
