//! Independently positioned audio-object CRUD for Numbers sheets.

use std::collections::HashMap;
use std::time::Duration;

use super::sheet_movies::graph::{
    MovieObjectIds, movie_creation_context, set_movie_geometry, set_movie_properties,
};
use super::*;
use crate::MediaPlaybackSettings;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::media_playback::replace_movie_playback_settings;
use crate::shapes::{
    DrawableGeometry, DrawablePoint, DrawableProperties, offset_drawable_geometry,
};

mod graph;

use graph::*;

/// One independently positioned audio clip owned directly by a Numbers sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersSheetAudioInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub audio_data_identifier: u64,
    /// Center point of Numbers' zero-size audio control, in sheet points.
    pub position: DrawablePoint,
    /// Shared drawable metadata, including accessibility description and lock state.
    pub properties: DrawableProperties,
    /// Trim, poster, repeat, and volume settings.
    pub playback: MediaPlaybackSettings,
    pub duration: Duration,
}

/// Typed placement and playback metadata for a newly created Numbers audio clip.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct NumbersSheetAudioOptions {
    /// Center point of Numbers' zero-size audio control, in sheet points.
    pub position: DrawablePoint,
    /// Playable duration of the source audio.
    pub duration: Duration,
}

impl NumbersSheetAudioOptions {
    pub const fn new(position: DrawablePoint, duration: Duration) -> Self {
        Self { position, duration }
    }
}

/// Result of removing one sheet-owned audio clip and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedNumbersSheetAudio {
    pub audio: NumbersSheetAudioInfo,
    /// Assets culled because the removed clip held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AudioClonePlacement {
    Offset,
    Preserve,
}

impl NumbersEditor {
    /// List independently positioned audio clips owned directly by one reachable sheet.
    pub fn sheet_audio(&self, sheet_id: u64) -> Result<Vec<NumbersSheetAudioInfo>> {
        audio_infos(self, sheet_id)
    }

    /// Add an independently editable audio clip to a reachable sheet.
    ///
    /// The audio archive, title/caption stand-ins, sheet ownership, style link,
    /// UUIDs, component data reference, and `Data/*` asset are constructed
    /// directly from typed values. No source drawable or package is copied.
    pub fn add_sheet_audio(
        &mut self,
        sheet_id: u64,
        preferred_filename: &str,
        data: &[u8],
        options: NumbersSheetAudioOptions,
    ) -> Result<NumbersSheetAudioInfo> {
        let (geometry, duration_seconds) = audio_creation_values(options)?;
        let context = movie_creation_context(self, sheet_id)?;
        let ids = MovieObjectIds::allocate(next_object_identifier(&self.package)?)?;

        let mut media = IWorkMediaEditor::from_package(self.package.clone())?;
        let asset = media.insert_unreferenced(preferred_filename, data)?;
        if asset.media_type != crate::MediaType::Audio {
            return Err(Error::ParseError(format!(
                "Numbers sheet audio requires audio data, not {}",
                asset.media_type.name()
            )));
        }
        let mut staged = media.into_package();
        let objects = audio_objects(
            ids,
            sheet_id,
            context.style_id,
            asset.data_identifier,
            geometry,
            duration_seconds,
        )?;
        staged.update_archive(&context.archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &context.archive_name,
            sheet_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, context.component_id, &ids.all())?;
        add_component_data_reference(
            &mut staged,
            context.component_id,
            asset.data_identifier,
            ids.drawable,
        )?;
        if context.stylesheet_component_id != context.component_id {
            add_component_external_reference(
                &mut staged,
                context.component_id,
                context.stylesheet_component_id,
                context.style_id,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .sheet_audio(sheet_id)?
            .into_iter()
            .find(|audio| audio.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers audio creation failed validation".to_owned())
            })?;
        let created_graph = audio_graph(&verified, sheet_id, ids.drawable)?;
        if created.audio_data_identifier != asset.data_identifier
            || created.position != options.position
            || created.duration.as_secs_f32() != duration_seconds
            || created_graph.object_ids != ids.all()
            || created_graph.uuid_object_ids != ids.all()
            || created_graph.data_references != [(asset.data_identifier, ids.drawable)]
            || verified.extract_media(asset.data_identifier)? != data
        {
            return Err(Error::InvalidFormat(
                "Numbers audio creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read the center position of one sheet-owned audio control.
    pub fn sheet_audio_position(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawablePoint> {
        Ok(audio_graph(self, sheet_id, drawable_object_id)?
            .info
            .position)
    }

    /// Move one sheet-owned audio control while preserving opaque media fields.
    pub fn set_sheet_audio_position(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: DrawablePoint,
    ) -> Result<()> {
        let source = audio_graph(self, sheet_id, drawable_object_id)?;
        let geometry = DrawableGeometry {
            position: Some(position),
            ..source.geometry
        }
        .validate()?;
        let mut staged = self.package.clone();
        set_movie_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_audio_position(sheet_id, drawable_object_id)? != position {
            return Err(Error::InvalidFormat(
                "Numbers audio position update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one sheet-owned audio control.
    pub fn sheet_audio_properties(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        Ok(audio_graph(self, sheet_id, drawable_object_id)?
            .info
            .properties)
    }

    /// Update audio accessibility, hyperlink, and lock properties.
    ///
    /// The typed update retains unknown native media fields and supports both
    /// clearing a property with `None` and encoding explicit boolean defaults.
    pub fn set_sheet_audio_properties(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = audio_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_movie_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_audio_properties(sheet_id, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Numbers audio properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read trim, poster, repeat, and volume settings for one sheet audio clip.
    pub fn sheet_audio_playback_settings(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<MediaPlaybackSettings> {
        Ok(audio_graph(self, sheet_id, drawable_object_id)?
            .info
            .playback)
    }

    /// Update playback settings while retaining unrelated and unknown audio fields.
    pub fn set_sheet_audio_playback_settings(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        settings: MediaPlaybackSettings,
    ) -> Result<()> {
        let source = audio_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        let expected = replace_movie_playback_settings(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            "Numbers audio",
            settings,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_audio_playback_settings(sheet_id, drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Numbers audio playback update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate one sheet audio control using Numbers' native placement.
    ///
    /// The audio and title/caption stand-ins receive fresh identifiers and
    /// UUIDs while retaining the source's style and unknown protobuf fields.
    /// The clone is added to the same sheet and shares its embedded audio asset
    /// with the source.
    pub fn duplicate_sheet_audio(
        &mut self,
        sheet_id: u64,
        source_drawable_object_id: u64,
    ) -> Result<NumbersSheetAudioInfo> {
        self.clone_sheet_audio(
            sheet_id,
            source_drawable_object_id,
            sheet_id,
            AudioClonePlacement::Offset,
        )
    }

    pub(super) fn duplicate_sheet_audio_to_sheet(
        &mut self,
        source_sheet_id: u64,
        source_drawable_object_id: u64,
        target_sheet_id: u64,
    ) -> Result<NumbersSheetAudioInfo> {
        self.clone_sheet_audio(
            source_sheet_id,
            source_drawable_object_id,
            target_sheet_id,
            AudioClonePlacement::Preserve,
        )
    }

    fn clone_sheet_audio(
        &mut self,
        source_sheet_id: u64,
        source_drawable_object_id: u64,
        target_sheet_id: u64,
        placement: AudioClonePlacement,
    ) -> Result<NumbersSheetAudioInfo> {
        let source = audio_graph(self, source_sheet_id, source_drawable_object_id)?;
        let (target_archive_name, _, _) = numbers_sheet(&self.package, target_sheet_id)?;
        if target_archive_name != source.archive_name {
            return Err(Error::InvalidFormat(format!(
                "Numbers audio source and target sheets must share a component: {} != {target_archive_name}",
                source.archive_name
            )));
        }
        let mut staged = self.package.clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len() + 1);
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Numbers audio graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }
        if source_sheet_id != target_sheet_id
            && remap.insert(source_sheet_id, target_sheet_id).is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers audio graph unexpectedly owns source sheet {source_sheet_id}"
            )));
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers audio object {identifier} is missing"))
                })?;
                clone_numbers_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }

        let new_drawable_id = *remap.get(&source_drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers audio clone has no drawable identifier".to_owned())
        })?;
        let geometry = match placement {
            AudioClonePlacement::Offset => {
                offset_drawable_geometry(source.geometry, DRAWABLE_DUPLICATE_OFFSET)?
            },
            AudioClonePlacement::Preserve => source.geometry,
        };
        let expected_position = geometry.position.ok_or_else(|| {
            Error::InvalidFormat("Numbers audio clone has no position".to_owned())
        })?;
        set_movie_geometry(&mut staged, &source.archive_name, new_drawable_id, geometry)?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            target_sheet_id,
            None,
            Some(new_drawable_id),
        )?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Numbers audio graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| {
                remap.get(identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers audio clone has no UUID identifier for {identifier}"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        add_component_object_uuids(&mut staged, source.component_id, &new_uuid_object_ids)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            let new_object_identifier =
                remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers audio clone has no data-reference object for {object_identifier}"
                    ))
                })?;
            add_component_data_reference(
                &mut staged,
                source.component_id,
                data_identifier,
                new_object_identifier,
            )?;
        }

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .sheet_audio(target_sheet_id)?
            .into_iter()
            .find(|audio| audio.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers audio duplication failed validation".to_owned())
            })?;
        let created_graph = audio_graph(&verified, target_sheet_id, new_drawable_id)?;
        let expected_data_references = source
            .data_references
            .iter()
            .map(|&(data_identifier, object_identifier)| {
                let new_object_identifier = remap.get(&object_identifier).copied().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers audio clone has no validated data-reference object for {object_identifier}"
                    ))
                })?;
                Ok((data_identifier, new_object_identifier))
            })
            .collect::<Result<Vec<_>>>()?;
        if created.sheet_id != target_sheet_id
            || created.audio_data_identifier != source.info.audio_data_identifier
            || created.position != expected_position
            || created.duration != source.info.duration
            || created_graph.object_ids.len() != source.object_ids.len()
            || created_graph.data_references != expected_data_references
        {
            return Err(Error::InvalidFormat(
                "Numbers audio duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Replace the bytes referenced by one sheet-owned audio clip.
    ///
    /// Audio controls duplicated with [`Self::duplicate_sheet_audio`] share
    /// their embedded asset, matching Numbers' native Duplicate behavior.
    pub fn replace_sheet_audio_data(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = audio_graph(self, sheet_id, drawable_object_id)?;
        self.replace_media(source.info.audio_data_identifier, replacement)
    }

    /// Remove one audio clip, its private graph, and its unshared asset.
    pub fn remove_sheet_audio(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetAudio> {
        let source = audio_graph(self, sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            source.sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                source.component_id,
                data_identifier,
                object_identifier,
            )?;
        }
        for identifier in &source.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                source.component_id,
                *identifier,
            )?;
        }
        staged.update_archive(&source.archive_name, |archive| {
            for identifier in &source.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers audio object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers audio object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
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
            .sheet_audio(sheet_id)?
            .iter()
            .any(|audio| audio.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_assets
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Numbers audio deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersSheetAudio {
            audio: source.info,
            removed_data_identifiers,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::DrawableSize;
    use crate::{MediaLoopMode, MediaVolume};

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-numbers-audio";
    const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFFreplacement-numbers-audio";
    const POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };

    fn options() -> NumbersSheetAudioOptions {
        NumbersSheetAudioOptions::new(POSITION, Duration::from_millis(1_375))
    }

    fn properties(description: &str) -> DrawableProperties {
        DrawableProperties {
            hyperlink_url: Some("https://example.test/numbers-audio".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some(description.to_owned()),
        }
    }

    #[test]
    fn scratch_spreadsheet_supports_audio_crud_without_a_source_package() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Audio")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        assert!(editor.sheet_audio(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());

        let created = editor
            .add_sheet_audio(sheet_id, "audio.aiff", AUDIO, options())
            .unwrap();
        assert_eq!(created.sheet_id, sheet_id);
        assert_eq!(created.position, POSITION);
        assert_eq!(created.duration, Duration::from_millis(1_375));
        assert_eq!(
            editor.extract_media(created.audio_data_identifier).unwrap(),
            AUDIO
        );
        assert!(editor.sheet_movies(sheet_id).unwrap().is_empty());

        let roundtripped = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.sheet_audio(sheet_id).unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_playback = MediaPlaybackSettings {
            loop_mode: Some(MediaLoopMode::Repeat),
            volume: Some(MediaVolume::new(0.75).unwrap()),
            ..created.playback
        };
        editor
            .set_sheet_audio_playback_settings(
                sheet_id,
                created.drawable_object_id,
                changed_playback,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_audio_playback_settings(sheet_id, created.drawable_object_id)
                .unwrap(),
            changed_playback
        );
        editor
            .set_sheet_audio_playback_settings(
                sheet_id,
                created.drawable_object_id,
                created.playback,
            )
            .unwrap();

        let changed_properties = properties("Accessible Numbers audio");
        editor
            .set_sheet_audio_properties(
                sheet_id,
                created.drawable_object_id,
                changed_properties.clone(),
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_audio_properties(sheet_id, created.drawable_object_id)
                .unwrap(),
            changed_properties
        );
        editor
            .set_sheet_audio_properties(
                sheet_id,
                created.drawable_object_id,
                DrawableProperties::default(),
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_audio_properties(sheet_id, created.drawable_object_id)
                .unwrap(),
            DrawableProperties::default()
        );

        let moved = DrawablePoint { x: 64.0, y: 96.0 };
        editor
            .set_sheet_audio_position(sheet_id, created.drawable_object_id, moved)
            .unwrap();
        assert_eq!(
            editor
                .sheet_audio_position(sheet_id, created.drawable_object_id)
                .unwrap(),
            moved
        );
        assert_eq!(
            editor
                .replace_sheet_audio_data(sheet_id, created.drawable_object_id, REPLACEMENT_AUDIO,)
                .unwrap(),
            AUDIO
        );
        editor
            .set_sheet_drawable_comment(sheet_id, created.drawable_object_id, "Remove after review")
            .unwrap();

        let removed = editor
            .remove_sheet_audio(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.audio.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [created.audio_data_identifier]
        );
        assert!(editor.sheet_audio(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn scratch_spreadsheet_supports_native_audio_duplication() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Audio")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source = editor
            .add_sheet_audio(sheet_id, "audio.aiff", AUDIO, options())
            .unwrap();
        let source_properties = properties("Duplicated Numbers audio");
        editor
            .set_sheet_audio_properties(
                sheet_id,
                source.drawable_object_id,
                source_properties.clone(),
            )
            .unwrap();

        let duplicate = editor
            .duplicate_sheet_audio(sheet_id, source.drawable_object_id)
            .unwrap();
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        let source_graph = audio_graph(&editor, sheet_id, source.drawable_object_id).unwrap();
        let duplicate_graph = audio_graph(&editor, sheet_id, duplicate.drawable_object_id).unwrap();
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
        assert_eq!(duplicate.sheet_id, sheet_id);
        assert_eq!(
            duplicate.audio_data_identifier,
            source.audio_data_identifier
        );
        assert_eq!(
            duplicate.position,
            DrawablePoint {
                x: source.position.x + DRAWABLE_DUPLICATE_OFFSET,
                y: source.position.y + DRAWABLE_DUPLICATE_OFFSET,
            }
        );
        assert_eq!(duplicate.duration, source.duration);
        assert_eq!(duplicate.properties, source_properties);

        let moved_duplicate = DrawablePoint { x: 512.0, y: 288.0 };
        editor
            .set_sheet_audio_position(sheet_id, duplicate.drawable_object_id, moved_duplicate)
            .unwrap();
        assert_eq!(
            editor
                .sheet_audio_position(sheet_id, source.drawable_object_id)
                .unwrap(),
            source.position
        );
        assert_eq!(
            editor
                .sheet_audio_position(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            moved_duplicate
        );
        assert_eq!(
            editor
                .replace_sheet_audio_data(
                    sheet_id,
                    duplicate.drawable_object_id,
                    REPLACEMENT_AUDIO,
                )
                .unwrap(),
            AUDIO
        );
        assert_eq!(
            editor.extract_media(source.audio_data_identifier).unwrap(),
            REPLACEMENT_AUDIO
        );

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.sheet_audio(sheet_id).unwrap().len(), 2);
        assert_eq!(
            reopened
                .sheet_audio(sheet_id)
                .unwrap()
                .into_iter()
                .find(|audio| audio.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .position,
            moved_duplicate
        );

        let removed_source = editor
            .remove_sheet_audio(sheet_id, source.drawable_object_id)
            .unwrap();
        assert!(removed_source.removed_data_identifiers.is_empty());
        assert_eq!(editor.sheet_audio(sheet_id).unwrap().len(), 1);
        let removed_duplicate = editor
            .remove_sheet_audio(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(
            removed_duplicate.removed_data_identifiers,
            [source.audio_data_identifier]
        );
        assert!(editor.sheet_audio(sheet_id).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_audio_creation_and_cross_type_edits_are_transactional() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        assert!(editor.duplicate_sheet_audio(sheet_id, 999).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        for result in [
            editor.add_sheet_audio(sheet_id, "payload.bin", b"not audio", options()),
            editor.add_sheet_audio(
                sheet_id,
                "audio.aiff",
                AUDIO,
                NumbersSheetAudioOptions::new(POSITION, Duration::ZERO),
            ),
            editor.add_sheet_audio(999, "audio.aiff", AUDIO, options()),
            editor.add_sheet_audio(
                sheet_id,
                "audio.aiff",
                AUDIO,
                NumbersSheetAudioOptions::new(
                    DrawablePoint {
                        x: f32::NAN,
                        y: 10.0,
                    },
                    Duration::from_secs(1),
                ),
            ),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }

        let movie = editor
            .add_sheet_movie(
                sheet_id,
                "movie.mov",
                b"\0\0\0\x18ftypqt  source-built-numbers-movie",
                "poster.png",
                b"\x89PNG\r\n\x1a\nsource-built-numbers-poster",
                NumbersSheetMovieOptions::new(
                    POSITION,
                    DrawableSize {
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
                .set_sheet_audio_position(sheet_id, movie.drawable_object_id, POSITION)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .replace_sheet_audio_data(sheet_id, movie.drawable_object_id, AUDIO)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn removing_one_sheet_audio_preserves_another() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let first = editor
            .add_sheet_audio(sheet_id, "first.aiff", AUDIO, options())
            .unwrap();
        let second = editor
            .add_sheet_audio(
                sheet_id,
                "second.aiff",
                REPLACEMENT_AUDIO,
                NumbersSheetAudioOptions::new(
                    DrawablePoint { x: 640.0, y: 360.0 },
                    Duration::from_secs(2),
                ),
            )
            .unwrap();

        editor
            .remove_sheet_audio(sheet_id, first.drawable_object_id)
            .unwrap();
        assert_eq!(
            editor.sheet_audio(sheet_id).unwrap(),
            std::slice::from_ref(&second)
        );
        assert_eq!(
            editor.extract_media(second.audio_data_identifier).unwrap(),
            REPLACEMENT_AUDIO
        );
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }
}
