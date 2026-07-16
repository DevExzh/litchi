//! Body-anchored audio CRUD for Pages documents.

use std::time::Duration;

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::package_metadata::{add_component_external_reference, component_identifier_for_entry};
use crate::shapes::DrawablePoint;

mod graph;

use graph::*;

/// One audio-only media control anchored to the Pages body text flow.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesAudioInfo {
    pub drawable_object_id: u64,
    /// UTF-16 index of the object-replacement character in the body text.
    pub anchor_character_index: u32,
    pub audio_data_identifier: u64,
    /// Center point of Pages' zero-size audio control, in document points.
    pub position: DrawablePoint,
    pub duration: Duration,
}

/// Typed placement and playback metadata for a newly created Pages audio clip.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PagesAudioOptions {
    /// Center point of Pages' zero-size audio control, in document points.
    pub position: DrawablePoint,
    /// Playable duration of the source audio.
    pub duration: Duration,
}

impl PagesAudioOptions {
    pub const fn new(position: DrawablePoint, duration: Duration) -> Self {
        Self { position, duration }
    }
}

/// Result of removing one body-anchored Pages audio clip and its private graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedPagesAudio {
    pub audio: PagesAudioInfo,
    /// Assets culled because the removed clip held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

impl PagesEditor {
    /// List audio-only media controls anchored to the body in text-flow order.
    pub fn body_audio(&self) -> Result<Vec<PagesAudioInfo>> {
        body_audio_infos(self)
    }

    /// Add an independently editable audio clip at a UTF-16 body position.
    ///
    /// The audio archive, title/caption stand-ins, body attachment,
    /// object-replacement character, z-order, style relationship, UUIDs,
    /// component data reference, and `Data/*` asset are constructed directly
    /// from typed values. No source drawable or package is copied.
    pub fn add_body_audio(
        &mut self,
        anchor_character_index: usize,
        preferred_filename: &str,
        data: &[u8],
        options: PagesAudioOptions,
    ) -> Result<PagesAudioInfo> {
        let (geometry, duration_seconds) = audio_creation_values(options)?;
        let root = root_document(self.package())?;
        let style_id = audio_style_id(self.package(), &root)?;
        let first_identifier = next_object_identifier(self.package())?;
        let (creates_z_order, z_order_id) = if let Some(z_order) = &root.drawables_zorder {
            (false, z_order.identifier)
        } else {
            (true, first_identifier)
        };
        let graph_first_identifier = first_identifier
            .checked_add(u64::from(creates_z_order))
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let ids = AudioObjectIds::allocate(graph_first_identifier)?;
        let archive_name = find_object_archive(self.package(), self.body_storage_id)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let asset = media.insert_unreferenced(preferred_filename, data)?;
        if asset.media_type != crate::MediaType::Audio {
            return Err(Error::ParseError(format!(
                "Pages body audio requires audio data, not {}",
                asset.media_type.name()
            )));
        }
        let mut staged = media.into_package();
        if creates_z_order {
            text_box_create::create_drawable_z_order(&mut staged, &archive_name, z_order_id)?;
        }
        let objects = audio_objects(
            ids,
            self.body_storage_id,
            style_id,
            asset.data_identifier,
            geometry,
            duration_seconds,
            root.left_margin.unwrap_or_default(),
        )?;
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;

        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id,
            anchor_character_index,
            ids.attachment,
        )?;
        patch_pages_zorder(&mut staged, None, Some(ids.drawable))?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &ids.uuid_objects())?;
        add_component_data_reference(
            &mut staged,
            DOCUMENT_OBJECT_ID,
            asset.data_identifier,
            ids.drawable,
        )?;
        let style_archive = find_object_archive(&staged, style_id)?;
        if let Some(stylesheet_component) = component_identifier_for_entry(&staged, &style_archive)?
            && stylesheet_component != DOCUMENT_OBJECT_ID
        {
            add_component_external_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                stylesheet_component,
                style_id,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .body_audio()?
            .into_iter()
            .find(|audio| audio.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages audio creation failed validation".to_owned())
            })?;
        let created_graph = body_audio_graph(&verified, ids.drawable)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        if created.anchor_character_index != expected_anchor
            || created.audio_data_identifier != asset.data_identifier
            || created.position != options.position
            || created.duration.as_secs_f32() != duration_seconds
            || created_graph.object_ids != ids.all()
            || verified.extract_media(asset.data_identifier)? != data
        {
            return Err(Error::InvalidFormat(
                "Pages audio creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read the center position of one body-anchored audio control.
    pub fn body_audio_position(&self, drawable_object_id: u64) -> Result<DrawablePoint> {
        Ok(body_audio_graph(self, drawable_object_id)?.info.position)
    }

    /// Move a body-anchored audio control while preserving opaque media fields.
    pub fn set_body_audio_position(
        &mut self,
        drawable_object_id: u64,
        position: DrawablePoint,
    ) -> Result<()> {
        let source = body_audio_graph(self, drawable_object_id)?;
        let geometry = DrawableGeometry {
            position: Some(position),
            ..source.geometry
        }
        .validate()?;
        let left_margin = root_document(self.package())?
            .left_margin
            .unwrap_or_default();
        let mut staged = self.package().clone();
        set_audio_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        set_audio_attachment_position(
            &mut staged,
            &source.archive_name,
            source.attachment_id,
            position,
            left_margin,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_audio_position(drawable_object_id)? != position {
            return Err(Error::InvalidFormat(
                "Pages audio position update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace the bytes referenced by one body-anchored audio clip.
    pub fn replace_body_audio_data(
        &mut self,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = body_audio_graph(self, drawable_object_id)?;
        self.replace_media(source.info.audio_data_identifier, replacement)
    }

    /// Remove body audio, its attachment/private graph, and unshared asset.
    pub fn remove_body_audio(&mut self, drawable_object_id: u64) -> Result<RemovedPagesAudio> {
        let source = body_audio_graph(self, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = source.info.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(drawable_object_id), None)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                data_identifier,
                object_identifier,
            )?;
        }
        for identifier in &source.object_ids {
            let object_archive = find_object_archive(&staged, *identifier)?;
            staged.update_archive(&object_archive, |archive| {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages audio object {identifier} is missing from {object_archive}"
                    ))
                })?;
                Ok(())
            })?;
        }
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Pages audio object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let mut media = IWorkMediaEditor::from_package(staged)?;
        let mut removed_data_identifiers = Vec::new();
        let data_identifiers = source
            .data_references
            .iter()
            .map(|(data, _)| *data)
            .collect::<HashSet<_>>();
        for identifier in data_identifiers {
            if media
                .asset(identifier)
                .is_some_and(|asset| !asset.is_referenced())
            {
                media.remove_unreferenced(identifier)?;
                removed_data_identifiers.push(identifier);
            }
        }
        removed_data_identifiers.sort_unstable();
        let staged = media.into_package();
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let remaining_assets = verified.media_assets()?;
        if verified
            .body_audio()?
            .iter()
            .any(|audio| audio.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_assets
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Pages audio deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedPagesAudio {
            audio: source.info,
            removed_data_identifiers,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-pages-audio";
    const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFFreplacement-pages-audio";
    const POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };

    fn options() -> PagesAudioOptions {
        PagesAudioOptions::new(POSITION, Duration::from_millis(1_375))
    }

    #[test]
    fn scratch_document_supports_audio_crud_without_a_source_package() {
        let mut editor = PagesEditor::create_with_text("Audio notes").unwrap();
        let anchor = "Audio notes".encode_utf16().count();
        assert!(editor.body_audio().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());

        let created = editor
            .add_body_audio(anchor, "audio.aiff", AUDIO, options())
            .unwrap();
        assert_eq!(created.anchor_character_index, anchor as u32);
        assert_eq!(created.position, POSITION);
        assert_eq!(created.duration, Duration::from_millis(1_375));
        assert_eq!(editor.body_text().unwrap(), "Audio notes\u{fffc}");
        assert_eq!(
            editor.extract_media(created.audio_data_identifier).unwrap(),
            AUDIO
        );
        assert!(editor.body_movies().unwrap().is_empty());

        let roundtripped = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.body_audio().unwrap(),
            std::slice::from_ref(&created)
        );

        let moved = DrawablePoint { x: 300.0, y: 360.0 };
        editor
            .set_body_audio_position(created.drawable_object_id, moved)
            .unwrap();
        assert_eq!(
            editor
                .body_audio_position(created.drawable_object_id)
                .unwrap(),
            moved
        );
        assert_eq!(
            editor
                .replace_body_audio_data(created.drawable_object_id, REPLACEMENT_AUDIO)
                .unwrap(),
            AUDIO
        );
        editor
            .set_drawable_comment(created.drawable_object_id, "Remove after review")
            .unwrap();

        let removed = editor
            .remove_body_audio(created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.audio.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [created.audio_data_identifier]
        );
        assert_eq!(editor.body_text().unwrap(), "Audio notes");
        assert!(editor.body_audio().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_audio_creation_and_cross_type_edits_are_transactional() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();
        for result in [
            editor.add_body_audio(4, "payload.bin", b"not audio", options()),
            editor.add_body_audio(
                4,
                "audio.aiff",
                AUDIO,
                PagesAudioOptions::new(POSITION, Duration::ZERO),
            ),
            editor.add_body_audio(5, "audio.aiff", AUDIO, options()),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }

        let movie = editor
            .add_body_movie(
                4,
                "movie.mov",
                b"\0\0\0\x18ftypqt  source-built-pages-movie",
                "poster.png",
                b"\x89PNG\r\n\x1a\nsource-built-pages-poster",
                crate::pages::PagesMovieOptions::new(
                    POSITION,
                    crate::shapes::DrawableSize {
                        width: 320.0,
                        height: 180.0,
                    },
                    Duration::from_secs(1),
                ),
            )
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_audio_position(movie.drawable_object_id, POSITION)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .replace_body_audio_data(movie.drawable_object_id, AUDIO)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn removing_earlier_audio_shifts_and_preserves_later_attachments() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let first = editor
            .add_body_audio(4, "first.aiff", AUDIO, options())
            .unwrap();
        let second = editor
            .add_body_audio(5, "second.aiff", REPLACEMENT_AUDIO, options())
            .unwrap();

        editor.remove_body_audio(first.drawable_object_id).unwrap();
        let remaining = editor.body_audio().unwrap();
        assert_eq!(remaining.len(), 1);
        assert_eq!(remaining[0].drawable_object_id, second.drawable_object_id);
        assert_eq!(remaining[0].anchor_character_index, 4);
        assert_eq!(editor.body_text().unwrap(), "Body\u{fffc}");
        assert_eq!(
            editor.extract_media(second.audio_data_identifier).unwrap(),
            REPLACEMENT_AUDIO
        );
        PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }
}
