//! Native caption CRUD for Numbers sheet charts.

use super::*;
use crate::image_caption::{
    CaptionObjectIds, CaptionThemeStyle, DrawableCaptionKind, DrawableCaptionSlot, caption_objects,
    drawable_caption_slot, patch_drawable_caption_reference, replace_object_reference,
    standin_caption_object,
};
use crate::wire::transform_length_delimited_field;

/// `TSCH.ChartDrawableArchive` embeds its `TSD.DrawableArchive` in field one.
const CHART_DRAWABLE_SUPER_FIELD: u32 = 1;

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
                .storage(storage_id)
                .map(|storage| storage.text)
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
    let chart = IWorkChartArchive::decode(message.data.as_slice())?;
    let drawable = chart.drawable.super_.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} has no drawable payload"
        ))
    })?;
    drawable_caption_slot(
        &editor.package,
        drawable_object_id,
        drawable.caption.as_ref(),
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
        text_editor.set_text(storage_id, text)?;
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
        )?;
        add_component_object_uuids(&mut staged, source.component_id, &ids.all())?;
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
) -> Result<()> {
    let objects = caption_objects(
        ids,
        drawable_object_id,
        drawable_width,
        text,
        DrawableCaptionKind::Caption,
        theme,
        language,
    )?;
    package.update_archive(archive_name, |archive| {
        for object in objects {
            archive.insert_object(object)?;
        }
        replace_sheet_chart_caption_reference(
            archive,
            drawable_object_id,
            old_reference_id,
            ids.info,
        )
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
    package.update_archive(archive_name, |archive| {
        archive.insert_object(standin)?;
        replace_sheet_chart_caption_reference(
            archive,
            drawable_object_id,
            old_reference_id,
            standin_id,
        )
    })
}

fn replace_sheet_chart_caption_reference(
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
    let chart = IWorkChartArchive::decode(original)?;
    let current_reference_id = chart
        .drawable
        .super_
        .as_ref()
        .and_then(|drawable| drawable.caption.as_ref())
        .map(|reference| reference.identifier);
    if current_reference_id != Some(old_reference_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers chart {drawable_object_id} caption reference changed unexpectedly"
        )));
    }
    let data =
        transform_length_delimited_field(original, CHART_DRAWABLE_SUPER_FIELD, |drawable| {
            patch_drawable_caption_reference(drawable, DrawableCaptionKind::Caption, replacement_id)
        })?;
    let actual_reference_id = IWorkChartArchive::decode(data.as_slice())?
        .drawable
        .super_
        .as_ref()
        .and_then(|drawable| drawable.caption.as_ref())
        .map(|reference| reference.identifier);
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
