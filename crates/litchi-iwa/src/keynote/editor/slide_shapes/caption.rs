//! Native title and caption CRUD for Keynote slide shapes.

use std::collections::HashSet;

use prost::Message;

use super::*;
use crate::DrawableTitleCaption;
use crate::IWorkThemeArchive;
use crate::image_caption::{
    CAPTION_INFO_MESSAGE_TYPE, CAPTION_PLACEMENT_MESSAGE_TYPE, CaptionObjectIds, CaptionThemeStyle,
    DrawableCaptionKind, SHAPE_STYLE_MESSAGE_TYPE, STANDIN_CAPTION_MESSAGE_TYPE,
    STORAGE_MESSAGE_TYPE, caption_objects, patch_shape_info_caption_reference,
    replace_object_reference, standin_caption_object,
};
use crate::protobuf::tswp;

const KEYNOTE_THEME_MESSAGE_TYPE: u32 = 10;

#[derive(Debug, Clone)]
pub(super) struct SlideShapeCaptionSlot {
    pub(super) reference_id: u64,
    pub(super) storage_id: Option<u64>,
    pub(super) object_ids: Vec<u64>,
}

impl KeynoteEditor {
    /// Read the native title and caption attached to one ordinary slide shape.
    pub fn slide_shape_title_caption(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableTitleCaption> {
        slide_shape_title_caption(self, slide_index, drawable_object_id)
    }

    /// Create or replace one ordinary slide shape's native title.
    pub fn set_slide_shape_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        title: &str,
    ) -> Result<()> {
        set_slide_shape_caption(
            self,
            slide_index,
            drawable_object_id,
            title,
            DrawableCaptionKind::Title,
        )
    }

    /// Remove one ordinary slide shape's native title.
    ///
    /// Returns whether a title was present. Native iWork removal preserves the
    /// prior title graph for undo history and attaches a fresh empty stand-in.
    pub fn remove_slide_shape_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        remove_slide_shape_caption(
            self,
            slide_index,
            drawable_object_id,
            DrawableCaptionKind::Title,
        )
    }

    /// Create or replace one ordinary slide shape's native caption.
    pub fn set_slide_shape_caption(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        caption: &str,
    ) -> Result<()> {
        set_slide_shape_caption(
            self,
            slide_index,
            drawable_object_id,
            caption,
            DrawableCaptionKind::Caption,
        )
    }

    /// Remove one ordinary slide shape's native caption.
    ///
    /// Returns whether a caption was present. Native iWork removal preserves
    /// the prior caption graph for undo history and attaches a fresh empty
    /// stand-in.
    pub fn remove_slide_shape_caption(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        remove_slide_shape_caption(
            self,
            slide_index,
            drawable_object_id,
            DrawableCaptionKind::Caption,
        )
    }
}

pub(super) fn slide_shape_title_caption(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<DrawableTitleCaption> {
    let title = slide_shape_caption_slot(
        editor,
        slide_index,
        drawable_object_id,
        DrawableCaptionKind::Title,
    )?;
    let caption = slide_shape_caption_slot(
        editor,
        slide_index,
        drawable_object_id,
        DrawableCaptionKind::Caption,
    )?;
    let graph = ObjectGraph::read(editor.package())?;
    Ok(DrawableTitleCaption {
        title: title
            .storage_id
            .map(|storage_id| graph.storage_text(storage_id))
            .transpose()?,
        caption: caption
            .storage_id
            .map(|storage_id| graph.storage_text(storage_id))
            .transpose()?,
    })
}

pub(super) fn slide_shape_caption_slot(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    kind: DrawableCaptionKind,
) -> Result<SlideShapeCaptionSlot> {
    shape_graph(editor, slide_index, drawable_object_id)?;
    let graph = ObjectGraph::read(editor.package())?;
    let shape: tswp::ShapeInfoArchive = graph.decode_type(
        drawable_object_id,
        SHAPE_MESSAGE_TYPE,
        "TSWP.ShapeInfoArchive",
    )?;
    let reference = match kind {
        DrawableCaptionKind::Caption => shape.super_.super_.caption,
        DrawableCaptionKind::Title => shape.super_.super_.title,
    };
    caption_slot_from_reference(&graph, drawable_object_id, reference, kind)
}

#[allow(deprecated)]
pub(super) fn caption_slot_from_reference(
    graph: &ObjectGraph,
    drawable_object_id: u64,
    reference: Option<tsp::Reference>,
    kind: DrawableCaptionKind,
) -> Result<SlideShapeCaptionSlot> {
    let reference_id = reference
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote shape {drawable_object_id} has no native title/caption reference"
            ))
        })?
        .identifier;
    let drawable_archive_name = graph.archive_name(drawable_object_id)?.to_owned();
    if graph.archive_name(reference_id)? != drawable_archive_name.as_str() {
        return Err(Error::InvalidFormat(format!(
            "Keynote title/caption object {reference_id} is outside shape {drawable_object_id}'s component"
        )));
    }
    let messages = graph.objects.get(&reference_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote title/caption object {reference_id} is missing"
        ))
    })?;
    let caption_messages = messages
        .iter()
        .filter(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if caption_messages.is_empty() {
        require_exact_message_count(
            graph,
            reference_id,
            STANDIN_CAPTION_MESSAGE_TYPE,
            "stand-in caption",
        )?;
        return Ok(SlideShapeCaptionSlot {
            reference_id,
            storage_id: None,
            object_ids: vec![reference_id],
        });
    }
    let [message] = caption_messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote title/caption object {reference_id} has multiple CaptionInfo payloads"
        )));
    };
    let info = crate::protobuf::tsa::CaptionInfoArchive::decode(message.data.as_slice())?;
    if info.child_info_kind != Some(kind.native_kind()) {
        return Err(Error::InvalidFormat(format!(
            "Keynote title/caption object {reference_id} has the wrong native kind"
        )));
    }
    if info
        .super_
        .super_
        .super_
        .parent
        .as_ref()
        .map(|parent| parent.identifier)
        != Some(drawable_object_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote title/caption object {reference_id} has the wrong parent drawable"
        )));
    }
    let storage_id = required_caption_reference(
        reference_id,
        info.super_.owned_storage.as_ref(),
        "text storage",
    )?;
    if info
        .super_
        .deprecated_storage
        .as_ref()
        .map(|storage| storage.identifier)
        != Some(storage_id)
        || info.super_.is_text_box != Some(true)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote title/caption object {reference_id} has inconsistent text storage"
        )));
    }
    let style_id = required_caption_reference(
        reference_id,
        info.super_.super_.style.as_ref(),
        "shape style",
    )?;
    let placement_id =
        required_caption_reference(reference_id, info.placement.as_ref(), "placement")?;
    for (identifier, message_type, label) in [
        (style_id, SHAPE_STYLE_MESSAGE_TYPE, "shape style"),
        (storage_id, STORAGE_MESSAGE_TYPE, "text storage"),
        (placement_id, CAPTION_PLACEMENT_MESSAGE_TYPE, "placement"),
    ] {
        require_exact_message_count(graph, identifier, message_type, label)?;
    }
    for (identifier, label) in [(storage_id, "text storage"), (placement_id, "placement")] {
        if graph.archive_name(identifier)? != drawable_archive_name.as_str() {
            return Err(Error::InvalidFormat(format!(
                "Keynote title/caption {label} {identifier} is outside shape {drawable_object_id}'s component"
            )));
        }
    }
    let mut object_ids = vec![reference_id];
    if graph.archive_name(style_id)? == drawable_archive_name.as_str() {
        object_ids.push(style_id);
    }
    object_ids.extend([storage_id, placement_id]);
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Keynote title/caption object {reference_id} aliases its private graph"
        )));
    }
    Ok(SlideShapeCaptionSlot {
        reference_id,
        storage_id: Some(storage_id),
        object_ids,
    })
}

fn set_slide_shape_caption(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    text: &str,
    kind: DrawableCaptionKind,
) -> Result<()> {
    let source = shape_graph(editor, slide_index, drawable_object_id)?;
    let slot = slide_shape_caption_slot(editor, slide_index, drawable_object_id, kind)?;
    let mut expected = slide_shape_title_caption(editor, slide_index, drawable_object_id)?;
    match kind {
        DrawableCaptionKind::Caption => expected.caption = Some(text.to_owned()),
        DrawableCaptionKind::Title => expected.title = Some(text.to_owned()),
    }
    let staged = if let Some(storage_id) = slot.storage_id {
        let mut text_editor = IWorkTextEditor::from_package(editor.package().clone());
        text_editor.set_text(storage_id, text)?;
        text_editor.into_package()
    } else {
        let (theme, language) = slide_shape_caption_theme(editor, slide_index)?;
        let drawable_width = source
            .info
            .geometry
            .size
            .ok_or_else(|| Error::InvalidFormat("Keynote shape has no displayed size".to_owned()))?
            .width;
        let ids = CaptionObjectIds::allocate(next_object_identifier(editor.package())?)?;
        let mut staged = editor.package().clone();
        insert_slide_shape_caption(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            slot.reference_id,
            drawable_width,
            text,
            kind,
            theme,
            language.as_deref(),
            ids,
        )?;
        add_component_object_uuids(&mut staged, source.component_id, &ids.all())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;
        staged
    };
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_shape_title_caption(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote shape title/caption update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_slide_shape_caption(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    kind: DrawableCaptionKind,
) -> Result<bool> {
    let source = shape_graph(editor, slide_index, drawable_object_id)?;
    let slot = slide_shape_caption_slot(editor, slide_index, drawable_object_id, kind)?;
    if slot.storage_id.is_none() {
        return Ok(false);
    }
    let mut expected = slide_shape_title_caption(editor, slide_index, drawable_object_id)?;
    match kind {
        DrawableCaptionKind::Caption => expected.caption = None,
        DrawableCaptionKind::Title => expected.title = None,
    }
    let standin_id = next_object_identifier(editor.package())?;
    let mut staged = editor.package().clone();
    insert_slide_shape_caption_standin(
        &mut staged,
        &source.archive_name,
        drawable_object_id,
        slot.reference_id,
        kind,
        standin_id,
    )?;
    add_component_object_uuids(&mut staged, source.component_id, &[standin_id])?;
    set_package_last_object_identifier(&mut staged, standin_id)?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_shape_title_caption(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote shape title/caption removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}

fn insert_slide_shape_caption(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    old_reference_id: u64,
    drawable_width: f32,
    text: &str,
    kind: DrawableCaptionKind,
    theme: CaptionThemeStyle,
    language: Option<&str>,
    ids: CaptionObjectIds,
) -> Result<()> {
    let objects = caption_objects(
        ids,
        drawable_object_id,
        drawable_width,
        text,
        kind,
        theme,
        language,
    )?;
    package.update_archive(archive_name, |archive| {
        for object in objects {
            archive.insert_object(object)?;
        }
        replace_slide_shape_caption_reference(
            archive,
            drawable_object_id,
            old_reference_id,
            ids.info,
            kind,
        )
    })
}

fn insert_slide_shape_caption_standin(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    old_reference_id: u64,
    kind: DrawableCaptionKind,
    standin_id: u64,
) -> Result<()> {
    let standin = standin_caption_object(standin_id)?;
    package.update_archive(archive_name, |archive| {
        archive.insert_object(standin)?;
        replace_slide_shape_caption_reference(
            archive,
            drawable_object_id,
            old_reference_id,
            standin_id,
            kind,
        )
    })
}

fn slide_shape_caption_theme(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<(CaptionThemeStyle, Option<String>)> {
    let graph = ObjectGraph::read(editor.package())?;
    let context = super::super::text_box_create::text_box_context(&graph, slide_index)?;
    let theme = IWorkThemeArchive::decode(graph.message_data_type(
        context.theme_id,
        KEYNOTE_THEME_MESSAGE_TYPE,
        "KN.ThemeArchive",
    )?)?;
    let paragraph_style_id = theme
        .extensions
        .application
        .ok_or_else(|| Error::InvalidFormat("Keynote theme has no application presets".to_owned()))?
        .caption_style_presets
        .into_iter()
        .next()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote theme has no caption style preset".to_owned())
        })?;
    if !graph.objects.contains_key(&paragraph_style_id) {
        return Err(Error::InvalidFormat(format!(
            "Keynote caption paragraph style {paragraph_style_id} is missing"
        )));
    }
    Ok((
        CaptionThemeStyle {
            stylesheet_id: context.stylesheet_id,
            paragraph_style_id,
        },
        context.language,
    ))
}

fn replace_slide_shape_caption_reference(
    archive: &mut Archive,
    drawable_object_id: u64,
    old_reference_id: u64,
    replacement_id: u64,
    kind: DrawableCaptionKind,
) -> Result<()> {
    let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote shape object {drawable_object_id} is missing"
        ))
    })?;
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == SHAPE_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} must have exactly one ShapeInfo payload"
        )));
    };
    let original = object.messages[*message_index].data.as_slice();
    let current = tswp::ShapeInfoArchive::decode(original)?;
    let current_reference_id = match kind {
        DrawableCaptionKind::Caption => current.super_.super_.caption,
        DrawableCaptionKind::Title => current.super_.super_.title,
    }
    .map(|reference| reference.identifier);
    if current_reference_id != Some(old_reference_id) {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} title/caption reference changed unexpectedly"
        )));
    }
    let data = patch_shape_info_caption_reference(original, kind, replacement_id)?;
    object.replace_message(
        *message_index,
        RawMessage {
            type_: SHAPE_MESSAGE_TYPE,
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

fn required_caption_reference(
    reference_id: u64,
    reference: Option<&tsp::Reference>,
    label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote title/caption object {reference_id} has no {label} reference"
            ))
        })
}

fn require_exact_message_count(
    graph: &ObjectGraph,
    object_id: u64,
    message_type: u32,
    label: &str,
) -> Result<()> {
    let messages = graph.objects.get(&object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote title/caption {label} {object_id} is missing"
        ))
    })?;
    if messages
        .iter()
        .filter(|message| message.type_ == message_type)
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote title/caption {label} {object_id} must have exactly one expected payload"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize, ShapePreset};

    const POSITION: DrawablePoint = DrawablePoint { x: 360.0, y: 420.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 480.0,
        height: 240.0,
    };

    #[test]
    fn scratch_presentation_supports_native_shape_title_caption_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Shape labels")
            .build()
            .unwrap();
        let shape = editor
            .add_slide_shape(
                0,
                "Quarterly trend",
                POSITION,
                SIZE,
                ShapePreset::RightArrow,
            )
            .unwrap();
        assert_eq!(
            editor
                .slide_shape_title_caption(0, shape.drawable_object_id)
                .unwrap(),
            DrawableTitleCaption::default()
        );

        editor
            .set_slide_shape_title(0, shape.drawable_object_id, "Trend title")
            .unwrap();
        editor
            .set_slide_shape_caption(0, shape.drawable_object_id, "Trend caption")
            .unwrap();
        let expected = DrawableTitleCaption {
            title: Some("Trend title".to_owned()),
            caption: Some("Trend caption".to_owned()),
        };
        assert_eq!(
            editor
                .slide_shape_title_caption(0, shape.drawable_object_id)
                .unwrap(),
            expected
        );

        let duplicate = editor
            .duplicate_slide_shape(0, shape.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor
                .slide_shape_title_caption(0, duplicate.drawable_object_id)
                .unwrap(),
            expected
        );

        editor
            .set_slide_shape_title(0, shape.drawable_object_id, "Updated trend title")
            .unwrap();
        assert!(
            editor
                .remove_slide_shape_caption(0, shape.drawable_object_id)
                .unwrap()
        );
        assert!(
            !editor
                .remove_slide_shape_caption(0, shape.drawable_object_id)
                .unwrap()
        );
        assert!(
            editor
                .remove_slide_shape_title(0, shape.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .slide_shape_title_caption(0, shape.drawable_object_id)
                .unwrap(),
            DrawableTitleCaption::default()
        );

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_title_caption(0, duplicate.drawable_object_id)
                .unwrap(),
            expected
        );
        editor = reopened;
        editor
            .remove_slide_shape(0, duplicate.drawable_object_id)
            .unwrap();
        assert!(
            editor
                .slide_shapes(0)
                .unwrap()
                .iter()
                .all(|item| item.drawable_object_id != duplicate.drawable_object_id)
        );
    }
}
