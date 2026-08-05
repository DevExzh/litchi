//! Body-anchored audio CRUD for Pages documents.

use std::collections::HashMap;
use std::time::Duration;

use litchi_iwa_common::media::Type as MediaType;

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::media_playback::replace_movie_playback_settings;
use crate::package_metadata::{add_component_external_reference, component_identifier_for_entry};
use crate::shapes::{DrawablePoint, DrawableProperties, offset_drawable_geometry};
use litchi_iwa_common::media::playback::MediaPlaybackSettings;

mod graph;

use graph::*;

/// Pages offsets duplicated audio controls farther than ordinary body drawables.
const AUDIO_DUPLICATE_OFFSET: f32 = 30.0;

/// One audio-only media control anchored to the Pages body text flow.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesAudioInfo {
    pub drawable_object_id: u64,
    /// UTF-16 index of the object-replacement character in the body text.
    pub anchor_character_index: u32,
    pub audio_data_identifier: u64,
    /// Center point of Pages' zero-size audio control, in document points.
    pub position: DrawablePoint,
    /// Shared drawable metadata, including accessibility description and lock state.
    pub properties: DrawableProperties,
    /// Trim, poster, repeat, and volume settings.
    pub playback: MediaPlaybackSettings,
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
        if asset.media_type != MediaType::Audio {
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

    /// Read shared drawable properties for one body-anchored audio control.
    pub fn body_audio_properties(&self, drawable_object_id: u64) -> Result<DrawableProperties> {
        Ok(body_audio_graph(self, drawable_object_id)?.info.properties)
    }

    /// Update audio accessibility, hyperlink, and lock properties.
    ///
    /// The typed update retains unknown native media fields and supports both
    /// clearing a property with `None` and encoding explicit boolean defaults.
    pub fn set_body_audio_properties(
        &mut self,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = body_audio_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_audio_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_audio_properties(drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Pages audio properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read trim, poster, repeat, and volume settings for one body audio clip.
    pub fn body_audio_playback_settings(
        &self,
        drawable_object_id: u64,
    ) -> Result<MediaPlaybackSettings> {
        Ok(body_audio_graph(self, drawable_object_id)?.info.playback)
    }

    /// Update playback settings while retaining unrelated and unknown audio fields.
    pub fn set_body_audio_playback_settings(
        &mut self,
        drawable_object_id: u64,
        settings: MediaPlaybackSettings,
    ) -> Result<()> {
        let source = body_audio_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        let expected = replace_movie_playback_settings(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            "Pages audio",
            settings,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_audio_playback_settings(drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Pages audio playback update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate one body audio control at a UTF-16 body position.
    ///
    /// The audio, title/caption stand-ins, and body attachment receive fresh
    /// identifiers and UUIDs while retaining the source's style and unknown
    /// protobuf fields. The clone is offset using Pages' native duplicate
    /// placement and shares its embedded audio asset with the source.
    pub fn duplicate_body_audio(
        &mut self,
        source_drawable_object_id: u64,
        anchor_character_index: usize,
    ) -> Result<PagesAudioInfo> {
        let source = body_audio_graph(self, source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Pages audio graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Pages audio object {identifier} is missing"))
                })?;
                clone_pages_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                Ok(archive.insert_object(cloned)?)
            })?;
        }

        let new_drawable_id = *remap.get(&source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Pages audio clone has no drawable identifier".to_owned())
        })?;
        let new_attachment_id = *remap.get(&source.attachment_id).ok_or_else(|| {
            Error::InvalidFormat("Pages audio clone has no attachment identifier".to_owned())
        })?;
        let geometry = offset_drawable_geometry(source.geometry, AUDIO_DUPLICATE_OFFSET)?;
        let expected_position = geometry
            .position
            .ok_or_else(|| Error::InvalidFormat("Pages audio clone has no position".to_owned()))?;
        set_audio_geometry(&mut staged, &source.archive_name, new_drawable_id, geometry)?;
        offset_pages_body_drawable_attachment_clone(
            &mut staged,
            new_attachment_id,
            AUDIO_DUPLICATE_OFFSET,
        )?;
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
            new_attachment_id,
        )?;
        patch_pages_zorder(&mut staged, None, Some(new_drawable_id))?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Pages audio graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages audio clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &new_uuid_object_ids)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            let new_object_identifier =
                remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages audio clone has no data-reference object for {object_identifier}"
                    ))
                })?;
            add_component_data_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                data_identifier,
                new_object_identifier,
            )?;
        }

        let verified = Self::from_package(staged)?;
        let created = verified
            .body_audio()?
            .into_iter()
            .find(|audio| audio.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages audio duplication failed validation".to_owned())
            })?;
        let created_graph = body_audio_graph(&verified, new_drawable_id)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        let expected_data_references = source
            .data_references
            .iter()
            .map(|&(data_identifier, object_identifier)| {
                let new_object_identifier = remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Pages audio clone has no validated data-reference object for {object_identifier}"
                    ))
                })?;
                Ok((data_identifier, new_object_identifier))
            })
            .collect::<Result<Vec<_>>>()?;
        if created.anchor_character_index != expected_anchor
            || created.audio_data_identifier != source.info.audio_data_identifier
            || created.position != expected_position
            || created.duration != source.info.duration
            || created_graph.object_ids.len() != source.object_ids.len()
            || created_graph.data_references != expected_data_references
        {
            return Err(Error::InvalidFormat(
                "Pages audio duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Replace the bytes referenced by one body-anchored audio clip.
    ///
    /// Audio controls duplicated with [`Self::duplicate_body_audio`] share
    /// their embedded asset, matching Pages' native Duplicate behavior.
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
        comments.clear_comment(crate::comments::DrawableObjectId::from_object_id(
            drawable_object_id,
        )?)?;
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
    use litchi_iwa_common::media::playback::{MediaLoopMode, MediaVolume};

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-pages-audio";
    const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFFreplacement-pages-audio";
    const POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };

    fn options() -> PagesAudioOptions {
        PagesAudioOptions::new(POSITION, Duration::from_millis(1_375))
    }

    fn properties(description: &str) -> DrawableProperties {
        DrawableProperties {
            hyperlink_url: Some("https://example.test/pages-audio".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some(description.to_owned()),
        }
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

        let changed_playback = MediaPlaybackSettings {
            loop_mode: Some(MediaLoopMode::Repeat),
            volume: Some(MediaVolume::new(0.75).unwrap()),
            ..created.playback
        };
        editor
            .set_body_audio_playback_settings(created.drawable_object_id, changed_playback)
            .unwrap();
        assert_eq!(
            editor
                .body_audio_playback_settings(created.drawable_object_id)
                .unwrap(),
            changed_playback
        );
        editor
            .set_body_audio_playback_settings(created.drawable_object_id, created.playback)
            .unwrap();

        let changed_properties = properties("Accessible Pages audio");
        editor
            .set_body_audio_properties(created.drawable_object_id, changed_properties.clone())
            .unwrap();
        assert_eq!(
            editor
                .body_audio_properties(created.drawable_object_id)
                .unwrap(),
            changed_properties
        );
        editor
            .set_body_audio_properties(created.drawable_object_id, DrawableProperties::default())
            .unwrap();
        assert_eq!(
            editor
                .body_audio_properties(created.drawable_object_id)
                .unwrap(),
            DrawableProperties::default()
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
    fn scratch_document_supports_native_audio_duplication() {
        let mut editor = PagesEditor::create_with_text("Audio notes").unwrap();
        let source = editor
            .add_body_audio(
                "Audio notes".encode_utf16().count(),
                "audio.aiff",
                AUDIO,
                options(),
            )
            .unwrap();
        let source_properties = properties("Duplicated Pages audio");
        editor
            .set_body_audio_properties(source.drawable_object_id, source_properties.clone())
            .unwrap();
        let duplicate_anchor = editor.body_text().unwrap().encode_utf16().count();

        let duplicate = editor
            .duplicate_body_audio(source.drawable_object_id, duplicate_anchor)
            .unwrap();
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        let source_graph = body_audio_graph(&editor, source.drawable_object_id).unwrap();
        let duplicate_graph = body_audio_graph(&editor, duplicate.drawable_object_id).unwrap();
        assert!(
            source_graph
                .object_ids
                .iter()
                .all(|identifier| !duplicate_graph.object_ids.contains(identifier))
        );
        assert!(
            source_graph
                .uuid_object_ids
                .iter()
                .all(|identifier| !duplicate_graph.uuid_object_ids.contains(identifier))
        );
        assert_eq!(duplicate.anchor_character_index, duplicate_anchor as u32);
        assert_eq!(
            duplicate.audio_data_identifier,
            source.audio_data_identifier
        );
        assert_eq!(
            duplicate.position,
            DrawablePoint {
                x: source.position.x + AUDIO_DUPLICATE_OFFSET,
                y: source.position.y + AUDIO_DUPLICATE_OFFSET,
            }
        );
        assert_eq!(duplicate.duration, source.duration);
        assert_eq!(duplicate.properties, source_properties);

        let moved_duplicate = DrawablePoint { x: 312.0, y: 264.0 };
        editor
            .set_body_audio_position(duplicate.drawable_object_id, moved_duplicate)
            .unwrap();
        assert_eq!(
            editor
                .body_audio_position(source.drawable_object_id)
                .unwrap(),
            source.position
        );
        assert_eq!(
            editor
                .body_audio_position(duplicate.drawable_object_id)
                .unwrap(),
            moved_duplicate
        );
        assert_eq!(
            editor
                .replace_body_audio_data(duplicate.drawable_object_id, REPLACEMENT_AUDIO)
                .unwrap(),
            AUDIO
        );
        assert_eq!(
            editor.extract_media(source.audio_data_identifier).unwrap(),
            REPLACEMENT_AUDIO
        );

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.body_audio().unwrap().len(), 2);
        assert_eq!(
            reopened
                .body_audio()
                .unwrap()
                .into_iter()
                .find(|audio| audio.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .position,
            moved_duplicate
        );

        let removed_source = editor.remove_body_audio(source.drawable_object_id).unwrap();
        assert!(removed_source.removed_data_identifiers.is_empty());
        assert_eq!(editor.body_audio().unwrap().len(), 1);
        let removed_duplicate = editor
            .remove_body_audio(duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(
            removed_duplicate.removed_data_identifiers,
            [source.audio_data_identifier]
        );
        assert!(editor.body_audio().unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_audio_creation_and_cross_type_edits_are_transactional() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(editor.duplicate_body_audio(999, 0).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
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
