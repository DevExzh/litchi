//! Native caption CRUD for Pages body charts.

use super::*;
use crate::charts::caption_edge::{chart_caption_identifier, rewrite_chart_caption_identifier};
use crate::image_caption::{
    CaptionObjectIds, CaptionThemeStyle, DrawableCaptionKind, DrawableCaptionSlot, caption_objects,
    drawable_caption_slot, standin_caption_object,
};
use crate::package::PackageLimits;
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};

impl PagesEditor {
    /// Read the native caption attached to one body chart.
    pub fn body_chart_caption(&self, drawable_object_id: u64) -> Result<Option<String>> {
        body_chart_caption(self, drawable_object_id)
    }

    /// Create or replace the native caption attached to one body chart.
    pub fn set_body_chart_caption(&mut self, drawable_object_id: u64, caption: &str) -> Result<()> {
        set_body_chart_caption(self, drawable_object_id, caption)
    }

    /// Remove the native caption attached to one body chart.
    ///
    /// Returns whether a caption was present. Native iWork removal preserves
    /// the prior caption graph for undo history and attaches a fresh empty
    /// stand-in.
    pub fn remove_body_chart_caption(&mut self, drawable_object_id: u64) -> Result<bool> {
        remove_body_chart_caption(self, drawable_object_id)
    }
}

fn body_chart_caption(editor: &PagesEditor, drawable_object_id: u64) -> Result<Option<String>> {
    let slot = body_chart_caption_slot(editor, drawable_object_id)?;
    slot.storage_id
        .map(|storage_id| {
            IWorkTextEditor::from_package(editor.package().clone())
                .storage(crate::text::native_storage_id(storage_id)?)
                .map(|storage| storage.storage.into_text())
        })
        .transpose()
}

fn body_chart_caption_slot(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<DrawableCaptionSlot> {
    let source = body_chart_graph(editor, drawable_object_id)?;
    let archive = editor.package().archive(&source.archive_name)?;
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
            "Pages chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    let caption_identifier =
        chart_caption_identifier(editor.package().limits(), message.data.as_slice())?;
    drawable_caption_slot(
        editor.package(),
        drawable_object_id,
        caption_identifier,
        DrawableCaptionKind::Caption,
        "Pages chart",
    )
}

fn set_body_chart_caption(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    text: &str,
) -> Result<()> {
    let source = body_chart_graph(editor, drawable_object_id)?;
    let slot = body_chart_caption_slot(editor, drawable_object_id)?;
    let expected = Some(text.to_owned());
    let staged = if let Some(storage_id) = slot.storage_id {
        let mut text_editor = IWorkTextEditor::from_package(editor.package().clone());
        text_editor.set_text(crate::text::native_storage_id(storage_id)?, text)?;
        text_editor.into_package()
    } else {
        let (theme, language) = body_chart_caption_theme(editor)?;
        let drawable_width = source
            .info
            .geometry
            .size
            .ok_or_else(|| Error::InvalidFormat("Pages chart has no displayed size".to_owned()))?
            .width;
        let ids = CaptionObjectIds::allocate(next_object_identifier(editor.package())?)?;
        let mut staged = editor.package().clone();
        insert_body_chart_caption(
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
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_caption(drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Pages chart caption update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_body_chart_caption(editor: &mut PagesEditor, drawable_object_id: u64) -> Result<bool> {
    let source = body_chart_graph(editor, drawable_object_id)?;
    let slot = body_chart_caption_slot(editor, drawable_object_id)?;
    if slot.storage_id.is_none() {
        return Ok(false);
    }
    let standin_id = next_object_identifier(editor.package())?;
    let mut staged = editor.package().clone();
    insert_body_chart_caption_standin(
        &mut staged,
        &source.archive_name,
        drawable_object_id,
        slot.reference_id,
        standin_id,
    )?;
    add_component_object_uuids(&mut staged, source.component_id, &[standin_id])?;
    set_package_last_object_identifier(&mut staged, standin_id)?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_caption(drawable_object_id)?.is_some() {
        return Err(Error::InvalidFormat(
            "Pages chart caption removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}

fn body_chart_caption_theme(editor: &PagesEditor) -> Result<(CaptionThemeStyle, Option<String>)> {
    let root = root_document(editor.package())?;
    let theme_id = root
        .theme
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("Pages document has no theme".to_owned()))?
        .identifier;
    let archive_name = find_object_archive(editor.package(), theme_id)?;
    let archive = editor.package().archive(&archive_name)?;
    let object = archive
        .object(theme_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Pages theme {theme_id} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == PAGES_THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages theme {theme_id} must have exactly one theme payload"
        )));
    };
    let theme = IWorkThemeArchive::decode(message.data.as_slice())?;
    let stylesheet_id = theme
        .base
        .document_stylesheet
        .ok_or_else(|| Error::InvalidFormat("Pages theme has no stylesheet".to_owned()))?
        .identifier;
    let paragraph_style_id = theme
        .extensions
        .application
        .ok_or_else(|| Error::InvalidFormat("Pages theme has no application presets".to_owned()))?
        .caption_style_presets
        .into_iter()
        .next()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Pages theme has no caption style preset".to_owned())
        })?;
    find_object_archive(editor.package(), stylesheet_id)?;
    find_object_archive(editor.package(), paragraph_style_id)?;
    Ok((
        CaptionThemeStyle {
            stylesheet_id,
            paragraph_style_id,
        },
        root.super_.document_language,
    ))
}

fn insert_body_chart_caption(
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
    let limits = package.limits();
    package.update_archive(archive_name, |archive| {
        for object in objects {
            archive.insert_object(object)?;
        }
        retarget_body_chart_caption_edge(
            limits,
            archive,
            drawable_object_id,
            old_reference_id,
            ids.info,
        )
    })
}

fn insert_body_chart_caption_standin(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    old_reference_id: u64,
    standin_id: u64,
) -> Result<()> {
    let standin = standin_caption_object(standin_id)?;
    let limits = package.limits();
    package.update_archive(archive_name, |archive| {
        archive.insert_object(standin)?;
        retarget_body_chart_caption_edge(
            limits,
            archive,
            drawable_object_id,
            old_reference_id,
            standin_id,
        )
    })
}

fn retarget_body_chart_caption_edge(
    limits: PackageLimits,
    archive: &mut crate::archive::Archive,
    drawable_object_id: u64,
    old_reference_id: u64,
    replacement_id: u64,
) -> Result<()> {
    let source_object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages chart {drawable_object_id} is missing"))
    })?;
    let message_indexes = source_object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| (message.type_ == CHART_MESSAGE_TYPE).then_some(index))
        .collect::<Vec<_>>();
    let [message_index] = message_indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    let message_info = source_object
        .archive_info
        .message_infos
        .get(*message_index)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages chart {drawable_object_id} has no metadata for chart payload"
            ))
        })?;
    if old_reference_id == replacement_id || replacement_id == 0 {
        return Err(Error::InvalidFormat(
            "Pages chart caption reference transition is invalid".to_owned(),
        ));
    }
    if message_info.data_references.contains(&old_reference_id)
        || message_info.data_references.contains(&replacement_id)
        || message_info
            .object_references
            .iter()
            .filter(|identifier| **identifier == old_reference_id)
            .count()
            != 1
        || message_info.object_references.contains(&replacement_id)
    {
        return Err(Error::InvalidFormat(
            "Pages chart caption reference metadata is ambiguous".to_owned(),
        ));
    }

    struct OwnedFieldTransition {
        path: Vec<u32>,
        before: Vec<u64>,
        after: Vec<u64>,
    }
    let mut field_states = Vec::new();
    field_states
        .try_reserve_exact(message_info.field_infos.len())
        .map_err(|_| Error::InvalidFormat("Pages chart FieldInfo allocation failed".to_owned()))?;
    let mut field_edges = 0usize;
    for field in &message_info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|identifier| **identifier == old_reference_id)
            .count();
        if field.data_references.contains(&old_reference_id)
            || field.data_references.contains(&replacement_id)
            || field.object_references.contains(&replacement_id)
        {
            return Err(Error::InvalidFormat(
                "Pages chart caption FieldInfo references are ambiguous".to_owned(),
            ));
        }
        if count != 0
            && (count != 1
                || !matches!(
                    field.path.path.as_slice(),
                    [11, 1] | [1, 11, 1] | [1, 1, 11, 1]
                ))
        {
            return Err(Error::InvalidFormat(
                "Pages chart caption FieldInfo path is unsupported".to_owned(),
            ));
        }
        if count != 0 {
            field_edges = field_edges.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Pages chart FieldInfo count overflow".to_owned())
            })?;
            if field_edges > 1 {
                return Err(Error::InvalidFormat(
                    "Pages chart caption FieldInfo reference is duplicated".to_owned(),
                ));
            }
        }
        let mut after = field.object_references.clone();
        after.retain(|identifier| *identifier != old_reference_id);
        if count != 0 {
            after.push(replacement_id);
        }
        field_states.push(OwnedFieldTransition {
            path: field.path.path.clone(),
            before: field.object_references.clone(),
            after,
        });
    }

    let original = source_object.messages[*message_index].data.as_slice();
    let current_reference_id = chart_caption_identifier(limits, original)?;
    if current_reference_id != Some(old_reference_id) {
        return Err(Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} caption reference changed unexpectedly"
        )));
    }
    let data = rewrite_chart_caption_identifier(limits, original, replacement_id)?;
    let actual_reference_id = chart_caption_identifier(limits, data.as_slice())?;
    if actual_reference_id != Some(replacement_id) {
        return Err(Error::InvalidFormat(
            "Pages chart caption reference patch failed validation".to_owned(),
        ));
    }

    let aggregate_before = message_info.object_references.clone();
    let mut aggregate_after = aggregate_before.clone();
    aggregate_after.retain(|identifier| *identifier != old_reference_id);
    aggregate_after.push(replacement_id);
    let fields = field_states
        .iter()
        .enumerate()
        .map(|(field_info_index, field)| FieldObjectReferenceTransition {
            field_info_index,
            expected_path: field.path.as_slice(),
            before: field.before.as_slice(),
            after: field.after.as_slice(),
        })
        .collect::<Vec<_>>();
    let transition = ObjectReferenceTransition {
        aggregate_before: aggregate_before.as_slice(),
        aggregate_after: aggregate_after.as_slice(),
        fields: fields.as_slice(),
    };
    let mut rewritten_object = source_object.clone();
    rewritten_object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            *message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data,
            },
            transition,
            limits.archive_limits(),
        )?;
    *archive.object_mut(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages chart {drawable_object_id} is missing"))
    })? = rewritten_object;
    Ok(())
}
