//! Native title and caption CRUD for Keynote slide movies.

use prost::Message;

use super::*;
use crate::DrawableTitleCaption;
use crate::image_caption::{
    CAPTION_INFO_MESSAGE_TYPE, CaptionObjectIds, CaptionThemeStyle, DrawableCaptionKind,
    caption_objects, patch_drawable_caption_reference, replace_object_reference,
    standin_caption_object,
};
use crate::wire::transform_length_delimited_field;

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;

#[derive(Debug, Clone, Copy)]
pub(super) struct MovieCaptionSlot {
    pub(super) reference_id: u64,
    pub(super) storage_id: Option<u64>,
}

impl KeynoteEditor {
    /// Read the native title and caption attached to one ordinary slide movie.
    pub fn slide_movie_title_caption(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableTitleCaption> {
        movie_title_caption(self, slide_index, drawable_object_id)
    }

    /// Create or replace one ordinary slide movie's native title.
    pub fn set_slide_movie_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        title: &str,
    ) -> Result<()> {
        set_slide_movie_caption(
            self,
            slide_index,
            drawable_object_id,
            title,
            DrawableCaptionKind::Title,
        )
    }

    /// Remove one ordinary slide movie's native title.
    ///
    /// Returns whether a title was present. Native iWork removal preserves the
    /// prior title graph for undo history and attaches a fresh empty stand-in.
    pub fn remove_slide_movie_title(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        remove_slide_movie_caption(
            self,
            slide_index,
            drawable_object_id,
            DrawableCaptionKind::Title,
        )
    }

    /// Create or replace one ordinary slide movie's native caption.
    pub fn set_slide_movie_caption(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        caption: &str,
    ) -> Result<()> {
        set_slide_movie_caption(
            self,
            slide_index,
            drawable_object_id,
            caption,
            DrawableCaptionKind::Caption,
        )
    }

    /// Remove one ordinary slide movie's native caption.
    ///
    /// Returns whether a caption was present. Native iWork removal preserves
    /// the prior caption graph for undo history and attaches a fresh empty
    /// stand-in.
    pub fn remove_slide_movie_caption(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        remove_slide_movie_caption(
            self,
            slide_index,
            drawable_object_id,
            DrawableCaptionKind::Caption,
        )
    }
}

fn movie_title_caption(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<DrawableTitleCaption> {
    editor.require_file_movie(slide_index, drawable_object_id)?;
    let graph = ObjectGraph::read(editor.package())?;
    let movie: tsd::MovieArchive =
        graph.decode_type(drawable_object_id, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    Ok(DrawableTitleCaption {
        title: movie_caption_slot_from_reference(
            &graph,
            movie.super_.title,
            DrawableCaptionKind::Title,
        )?
        .storage_id
        .map(|storage_id| graph.storage_text(storage_id))
        .transpose()?,
        caption: movie_caption_slot_from_reference(
            &graph,
            movie.super_.caption,
            DrawableCaptionKind::Caption,
        )?
        .storage_id
        .map(|storage_id| graph.storage_text(storage_id))
        .transpose()?,
    })
}

fn movie_caption_slot(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    kind: DrawableCaptionKind,
) -> Result<MovieCaptionSlot> {
    editor.require_file_movie(slide_index, drawable_object_id)?;
    let graph = ObjectGraph::read(editor.package())?;
    let movie: tsd::MovieArchive =
        graph.decode_type(drawable_object_id, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    let reference = match kind {
        DrawableCaptionKind::Caption => movie.super_.caption,
        DrawableCaptionKind::Title => movie.super_.title,
    };
    movie_caption_slot_from_reference(&graph, reference, kind)
}

fn movie_caption_slot_from_reference(
    graph: &ObjectGraph,
    reference: Option<tsp::Reference>,
    kind: DrawableCaptionKind,
) -> Result<MovieCaptionSlot> {
    let reference_id = reference
        .ok_or_else(|| {
            Error::InvalidFormat("Keynote movie has no title/caption reference".to_owned())
        })?
        .identifier;
    let messages = graph.objects.get(&reference_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote movie title/caption object {reference_id} is missing"
        ))
    })?;
    if messages
        .iter()
        .any(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
    {
        let info: crate::protobuf::tsa::CaptionInfoArchive = graph.decode_type(
            reference_id,
            CAPTION_INFO_MESSAGE_TYPE,
            "TSA.CaptionInfoArchive",
        )?;
        if info.child_info_kind != Some(kind.native_kind()) {
            return Err(Error::InvalidFormat(format!(
                "Keynote movie title/caption object {reference_id} has the wrong native kind"
            )));
        }
        let storage_id = info
            .super_
            .owned_storage
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote movie title/caption object {reference_id} has no text storage"
                ))
            })?
            .identifier;
        return Ok(MovieCaptionSlot {
            reference_id,
            storage_id: Some(storage_id),
        });
    }
    graph.decode_type::<tsd::StandinCaptionArchive>(
        reference_id,
        STANDIN_CAPTION_MESSAGE_TYPE,
        "TSD.StandinCaptionArchive",
    )?;
    Ok(MovieCaptionSlot {
        reference_id,
        storage_id: None,
    })
}

fn set_slide_movie_caption(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    text: &str,
    kind: DrawableCaptionKind,
) -> Result<()> {
    let source = editor.require_file_movie(slide_index, drawable_object_id)?;
    let slot = movie_caption_slot(editor, slide_index, drawable_object_id, kind)?;
    let mut expected = movie_title_caption(editor, slide_index, drawable_object_id)?;
    match kind {
        DrawableCaptionKind::Caption => expected.caption = Some(text.to_owned()),
        DrawableCaptionKind::Title => expected.title = Some(text.to_owned()),
    }
    let staged = if let Some(storage_id) = slot.storage_id {
        let mut text_editor = IWorkTextEditor::from_package(editor.package().clone());
        text_editor.set_text(crate::text::native_storage_id(storage_id)?, text)?;
        text_editor.into_package()
    } else {
        let context = movie_creation_context(editor, slide_index)?;
        let drawable_width = source
            .info
            .geometry
            .size
            .ok_or_else(|| Error::InvalidFormat("Keynote movie has no displayed size".to_owned()))?
            .width;
        let ids = CaptionObjectIds::allocate(next_object_identifier(editor.package())?)?;
        let mut staged = editor.package().clone();
        insert_slide_movie_caption(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            slot.reference_id,
            drawable_width,
            text,
            kind,
            context.caption_theme,
            context.language.as_deref(),
            ids,
        )?;
        add_component_object_uuids(&mut staged, context.component_id, &ids.all())?;
        set_package_last_object_identifier(&mut staged, ids.last())?;
        staged
    };
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_movie_title_caption(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote movie title/caption update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn remove_slide_movie_caption(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    kind: DrawableCaptionKind,
) -> Result<bool> {
    let source = editor.require_file_movie(slide_index, drawable_object_id)?;
    let slot = movie_caption_slot(editor, slide_index, drawable_object_id, kind)?;
    if slot.storage_id.is_none() {
        return Ok(false);
    }
    let mut expected = movie_title_caption(editor, slide_index, drawable_object_id)?;
    match kind {
        DrawableCaptionKind::Caption => expected.caption = None,
        DrawableCaptionKind::Title => expected.title = None,
    }
    let standin_id = next_object_identifier(editor.package())?;
    let mut staged = editor.package().clone();
    insert_slide_movie_caption_standin(
        &mut staged,
        &source.archive_name,
        drawable_object_id,
        slot.reference_id,
        kind,
        standin_id,
    )?;
    let component =
        component_identifier_for_entry(&staged, &source.archive_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {} is not registered",
                source.archive_name
            ))
        })?;
    add_component_object_uuids(&mut staged, component, &[standin_id])?;
    set_package_last_object_identifier(&mut staged, standin_id)?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_movie_title_caption(slide_index, drawable_object_id)? != expected {
        return Err(Error::InvalidFormat(
            "Keynote movie title/caption removal failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(true)
}

fn insert_slide_movie_caption(
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
        replace_slide_movie_caption_reference(
            archive,
            drawable_object_id,
            old_reference_id,
            ids.info,
            kind,
        )
    })
}

fn insert_slide_movie_caption_standin(
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
        replace_slide_movie_caption_reference(
            archive,
            drawable_object_id,
            old_reference_id,
            standin_id,
            kind,
        )
    })
}

fn replace_slide_movie_caption_reference(
    archive: &mut crate::archive::Archive,
    drawable_object_id: u64,
    old_reference_id: u64,
    replacement_id: u64,
    kind: DrawableCaptionKind,
) -> Result<()> {
    let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote movie object {drawable_object_id} is missing"
        ))
    })?;
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote movie {drawable_object_id} must have exactly one MovieArchive payload"
        )));
    };
    let original = object.messages[*message_index].data.as_slice();
    let current = tsd::MovieArchive::decode(original)?;
    let current_reference_id = match kind {
        DrawableCaptionKind::Caption => current.super_.caption,
        DrawableCaptionKind::Title => current.super_.title,
    }
    .map(|reference| reference.identifier);
    if current_reference_id != Some(old_reference_id) {
        return Err(Error::InvalidFormat(format!(
            "Keynote movie {drawable_object_id} title/caption reference changed unexpectedly"
        )));
    }
    let data = transform_length_delimited_field(original, 1, |drawable| {
        patch_drawable_caption_reference(drawable, kind, replacement_id)
    })?;
    object.replace_message(
        *message_index,
        RawMessage {
            type_: MOVIE_MESSAGE_TYPE,
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
