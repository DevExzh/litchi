//! Native-style duplication of populated Numbers sheets.

use super::*;
use litchi_numbers::{SheetSelector, TableSelector};

mod wire;

use wire::clone_empty_sheet_object;

const TABLE_POSITION_X_PATH: &[u32] = &[1, 1, 1, 1];
const TABLE_POSITION_Y_PATH: &[u32] = &[1, 1, 1, 2];
const SHEET_COPY_TEXT_BOX_OFFSET: f32 = 0.0;

#[derive(Debug)]
enum SheetDrawableClone {
    Table {
        model_id: u64,
        info_id: u64,
        name: String,
    },
    TextBox {
        drawable_id: u64,
        text: String,
    },
    Image {
        drawable_id: u64,
    },
    Shape {
        drawable_id: u64,
    },
    Audio {
        drawable_id: u64,
    },
    Movie {
        drawable_id: u64,
    },
}

impl NumbersEditor {
    /// Duplicate a populated sheet immediately after its source.
    ///
    /// Sheet settings and unknown wire fields are retained. Populated tables,
    /// local formula dependency graphs, ordinary text boxes, images, shapes,
    /// audio, and movies receive fresh object identities. Image, audio, and
    /// movie assets and shape styles remain shared, matching Numbers, while
    /// writable object storage is independent.
    /// Unsupported drawable kinds and cross-table formula edges are rejected
    /// transactionally.
    pub fn duplicate_sheet(&mut self, selector: SheetSelector<'_>) -> Result<NumbersSheetInfo> {
        let sheet_id = super::selectors::sheet_id(self, selector)?;
        let sheets = self.sheets()?;
        let source = sheets
            .iter()
            .find(|sheet| sheet.object_id == sheet_id)
            .ok_or_else(|| Error::ParseError(format!("Numbers sheet {sheet_id} not found")))?;
        let existing_names = sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<HashSet<_>>();
        let new_sheet_name = duplicate_sheet_name(&source.name, &existing_names)?;
        let (archive_name, message_index, sheet) = numbers_sheet(&self.package, sheet_id)?;
        let drawables = classify_sheet_drawables(self, sheet_id, &sheet)?;
        for drawable in &drawables {
            if let SheetDrawableClone::Table {
                model_id, info_id, ..
            } = drawable
                && !table_formula_graph_is_self_contained(self.package(), *info_id)?
            {
                return Err(Error::ParseError(format!(
                    "Cannot duplicate Numbers sheet {sheet_id}: table {model_id} has cross-table formula dependencies"
                )));
            }
        }

        let new_sheet_id = next_object_identifier(&self.package)?;
        let cloned_sheet = {
            let archive = self.package.archive(&archive_name)?;
            let source_object = archive.object(sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing"))
            })?;
            clone_empty_sheet_object(
                source_object,
                message_index,
                new_sheet_id,
                &new_sheet_name,
                &sheet.drawable_infos,
            )?
        };

        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| {
            Ok(archive.insert_object(cloned_sheet)?)
        })?;
        update_numbers_document(&mut staged, |document| {
            let matches = document
                .sheets
                .iter()
                .enumerate()
                .filter(|(_, reference)| reference.identifier == sheet_id)
                .map(|(index, _)| index)
                .collect::<Vec<_>>();
            let [source_index] = matches.as_slice() else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers root must reference sheet {sheet_id} exactly once"
                )));
            };
            document.sheets.insert(
                *source_index + 1,
                tsp::Reference {
                    identifier: new_sheet_id,
                    ..Default::default()
                },
            );
            Ok(())
        })?;
        set_package_last_object_identifier(&mut staged, new_sheet_id)?;
        register_sheet_uuid_if_needed(
            &mut staged,
            &self.package,
            &archive_name,
            sheet_id,
            new_sheet_id,
        )?;

        let mut working = Self::from_package(staged)?;
        let mut cloned_drawable_ids = Vec::with_capacity(drawables.len());
        for drawable in drawables {
            match drawable {
                SheetDrawableClone::Table {
                    model_id,
                    name: table_name,
                    ..
                } => {
                    let source_index = super::selectors::table_index(&working, model_id)?;
                    let cloned = working.duplicate_table(TableSelector::index(source_index))?;
                    working.move_table(
                        TableSelector::name(&cloned.name),
                        SheetSelector::name(&new_sheet_name),
                    )?;
                    rename_attached_table_in_package(
                        &mut working.package,
                        cloned.native_id(),
                        &table_name,
                    )?;
                    restore_table_geometry(&mut working.package, model_id, cloned.object_id)?;
                    cloned_drawable_ids
                        .push(find_table_owner(working.package(), cloned.object_id)?.table_info_id);
                },
                SheetDrawableClone::TextBox { drawable_id, text } => {
                    let cloned = working.duplicate_text_box_to_sheet(
                        sheet_id,
                        drawable_id,
                        new_sheet_id,
                        &text,
                        SHEET_COPY_TEXT_BOX_OFFSET,
                    )?;
                    cloned_drawable_ids.push(cloned.drawable_object_id);
                },
                SheetDrawableClone::Image { drawable_id } => {
                    let cloned = working.duplicate_sheet_image_to_sheet(
                        sheet_id,
                        drawable_id,
                        new_sheet_id,
                    )?;
                    cloned_drawable_ids.push(cloned.drawable_object_id);
                },
                SheetDrawableClone::Shape { drawable_id } => {
                    let cloned = working.duplicate_sheet_shape_to_sheet(
                        sheet_id,
                        drawable_id,
                        new_sheet_id,
                    )?;
                    cloned_drawable_ids.push(cloned.drawable_object_id);
                },
                SheetDrawableClone::Audio { drawable_id } => {
                    let cloned = working.duplicate_sheet_audio_to_sheet(
                        sheet_id,
                        drawable_id,
                        new_sheet_id,
                    )?;
                    cloned_drawable_ids.push(cloned.drawable_object_id);
                },
                SheetDrawableClone::Movie { drawable_id } => {
                    let cloned = working.duplicate_sheet_movie_to_sheet(
                        sheet_id,
                        drawable_id,
                        new_sheet_id,
                    )?;
                    cloned_drawable_ids.push(cloned.drawable_object_id);
                },
            }
        }

        let (_, _, verified_sheet) = numbers_sheet(working.package(), new_sheet_id)?;
        let verified_drawables = verified_sheet
            .drawable_infos
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        if verified_sheet.name != new_sheet_name || verified_drawables != cloned_drawable_ids {
            return Err(Error::InvalidFormat(
                "Numbers sheet duplication failed structural validation".to_owned(),
            ));
        }
        let created = working
            .sheets()?
            .into_iter()
            .find(|sheet| sheet.object_id == new_sheet_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers duplicated sheet is unreachable".to_owned())
            })?;
        self.package = working.package;
        Ok(created)
    }
}

fn classify_sheet_drawables(
    editor: &NumbersEditor,
    sheet_id: u64,
    sheet: &tn::SheetArchive,
) -> Result<Vec<SheetDrawableClone>> {
    let package = editor.package();
    let mut tables = HashMap::new();
    for descriptor in table_models(package)? {
        let owner = find_table_owner(package, descriptor.object_id)?;
        if owner.sheet_id == sheet_id {
            tables.insert(
                owner.table_info_id,
                (descriptor.object_id, descriptor.model.table_name),
            );
        }
    }
    let text_boxes = editor
        .sheet_text_boxes(sheet_id)?
        .into_iter()
        .map(|text_box| {
            (
                text_box.drawable_object_id,
                text_box.storage.storage.into_text(),
            )
        })
        .collect::<HashMap<_, _>>();
    let images = editor
        .sheet_images(sheet_id)?
        .into_iter()
        .map(|image| image.drawable_object_id)
        .collect::<HashSet<_>>();
    let shapes = editor
        .sheet_shapes(sheet_id)?
        .into_iter()
        .map(|shape| shape.drawable_object_id)
        .collect::<HashSet<_>>();
    let audio = editor
        .sheet_audio(sheet_id)?
        .into_iter()
        .map(|audio| audio.drawable_object_id)
        .collect::<HashSet<_>>();
    let movies = editor
        .sheet_movies(sheet_id)?
        .into_iter()
        .map(|movie| movie.drawable_object_id)
        .collect::<HashSet<_>>();

    sheet
        .drawable_infos
        .iter()
        .map(|reference| {
            if let Some((model_id, name)) = tables.remove(&reference.identifier) {
                return Ok(SheetDrawableClone::Table {
                    model_id,
                    info_id: reference.identifier,
                    name,
                });
            }
            if let Some(text) = text_boxes.get(&reference.identifier) {
                return Ok(SheetDrawableClone::TextBox {
                    drawable_id: reference.identifier,
                    text: text.clone(),
                });
            }
            if images.contains(&reference.identifier) {
                return Ok(SheetDrawableClone::Image {
                    drawable_id: reference.identifier,
                });
            }
            if shapes.contains(&reference.identifier) {
                return Ok(SheetDrawableClone::Shape {
                    drawable_id: reference.identifier,
                });
            }
            if audio.contains(&reference.identifier) {
                return Ok(SheetDrawableClone::Audio {
                    drawable_id: reference.identifier,
                });
            }
            if movies.contains(&reference.identifier) {
                return Ok(SheetDrawableClone::Movie {
                    drawable_id: reference.identifier,
                });
            }
            Err(Error::ParseError(format!(
                "Cannot duplicate Numbers sheet {sheet_id}: drawable {} is not a supported table, ordinary text box, image, shape, audio, or movie",
                reference.identifier
            )))
        })
        .collect()
}

fn duplicate_sheet_name(source: &str, existing: &HashSet<&str>) -> Result<String> {
    validate_name(source, "sheet")?;
    for suffix in 1u32..=u32::MAX {
        let candidate = format!("{source}-{suffix}");
        if !existing.contains(candidate.as_str()) {
            return Ok(candidate);
        }
    }
    Err(Error::ParseError(
        "Unable to allocate a unique Numbers sheet name".to_owned(),
    ))
}

fn register_sheet_uuid_if_needed(
    staged: &mut IWorkPackage,
    source: &IWorkPackage,
    archive_name: &str,
    source_sheet_id: u64,
    new_sheet_id: u64,
) -> Result<()> {
    let Some(component_id) = component_identifier_for_entry(source, archive_name)? else {
        return Ok(());
    };
    if component_uuid_identifiers(source, component_id)?
        .is_some_and(|mapped| mapped.contains(&source_sheet_id))
    {
        add_component_object_uuids(staged, component_id, &[new_sheet_id])?;
    }
    Ok(())
}

fn restore_table_geometry(
    package: &mut IWorkPackage,
    source_table_id: u64,
    cloned_table_id: u64,
) -> Result<()> {
    let source_owner = find_table_owner(package, source_table_id)?;
    let cloned_owner = find_table_owner(package, cloned_table_id)?;
    let locations = object_locations(package)?;
    let source_archive_name = locations.get(&source_owner.table_info_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table info {} is missing",
            source_owner.table_info_id
        ))
    })?;
    let source_archive = package.archive(source_archive_name)?;
    let source_object = source_archive
        .object(source_owner.table_info_id)
        .ok_or_else(|| Error::InvalidFormat("Numbers source table info is missing".to_owned()))?;
    let (_, source_info) = decode_table_info(source_object)?;
    let source_position = source_info
        .super_
        .geometry
        .and_then(|geometry| geometry.position);

    let clone_archive_name = locations
        .get(&cloned_owner.table_info_id)
        .ok_or_else(|| Error::InvalidFormat("Numbers cloned table info is missing".to_owned()))?
        .to_owned();
    package.update_archive(&clone_archive_name, |archive| {
        let object = archive
            .object_mut(cloned_owner.table_info_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers cloned table info is missing".to_owned())
            })?;
        let (message_index, cloned_info) = decode_table_info(object)?;
        let cloned_position = cloned_info
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.position.as_ref());
        match (&source_position, cloned_position) {
            (None, None) => return Ok(()),
            (Some(_), None) | (None, Some(_)) => {
                return Err(Error::InvalidFormat(
                    "Numbers table clone changed positioned-geometry presence".to_owned(),
                ));
            },
            (Some(source_position), Some(_)) => {
                let message_type = object.messages[message_index].type_;
                let data = patch_nested_fixed32_field(
                    &object.messages[message_index].data,
                    TABLE_POSITION_X_PATH,
                    true,
                    Some(source_position.x.to_bits()),
                )?;
                let data = patch_nested_fixed32_field(
                    &data,
                    TABLE_POSITION_Y_PATH,
                    true,
                    Some(source_position.y.to_bits()),
                )?;
                let verified = tst::TableInfoArchive::decode(data.as_slice())?;
                let verified_position = verified
                    .super_
                    .geometry
                    .and_then(|geometry| geometry.position);
                if verified_position.as_ref() != Some(source_position) {
                    return Err(Error::InvalidFormat(
                        "Numbers sheet clone failed to preserve table position".to_owned(),
                    ));
                }
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
            },
        }
        Ok(())
    })
}

#[cfg(test)]
mod tests {
    use std::time::Duration;

    use super::*;
    use crate::numbers::{
        NumbersDocumentBuilder, NumbersSheetAudioOptions, NumbersSheetImageOptions,
        NumbersSheetMovieOptions,
    };
    use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize, RgbaColor, ShapeFill};
    use litchi_iwa_common::media::playback::{MediaLoopMode, MediaVolume};
    use litchi_iwa_common::shape::path::Preset;

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsheet-duplicate-audio";
    const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  sheet-duplicate-movie";
    const MOVIE_POSTER: &[u8] = b"\x89PNG\r\n\x1a\nsheet-duplicate-movie-poster";
    const AUDIO_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
    const MOVED_AUDIO_POSITION: DrawablePoint = DrawablePoint { x: 510.0, y: 225.0 };
    const MOVIE_POSITION: DrawablePoint = DrawablePoint { x: 320.0, y: 210.0 };
    const MOVIE_SIZE: DrawableSize = DrawableSize {
        width: 320.0,
        height: 180.0,
    };
    const MOVIE_NATURAL_SIZE: DrawableSize = DrawableSize {
        width: 640.0,
        height: 360.0,
    };
    const MOVED_MOVIE_POSITION: DrawablePoint = DrawablePoint { x: 480.0, y: 290.0 };
    const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 84.0, y: 126.0 };
    const IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 320.0,
        height: 155.0,
    };
    const MOVED_IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 440.0, y: 72.0 };
    const SHAPE_POSITION: DrawablePoint = DrawablePoint { x: 108.0, y: 244.0 };
    const SHAPE_SIZE: DrawableSize = DrawableSize {
        width: 180.0,
        height: 96.0,
    };
    const MOVED_SHAPE_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
    const SOURCE_SHAPE_FILL: ShapeFill = ShapeFill::Solid(RgbaColor::black());
    const COPIED_SHAPE_FILL: ShapeFill = ShapeFill::None;

    #[test]
    fn source_built_movie_sheet_duplicates_with_native_shared_asset_semantics() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Movie")
            .build()
            .unwrap();
        let source_sheet = editor.sheets().unwrap().remove(0);
        let source = editor
            .add_sheet_movie(
                source_sheet.object_id,
                "probe.mov",
                MOVIE,
                "probe.png",
                MOVIE_POSTER,
                NumbersSheetMovieOptions::new(MOVIE_POSITION, MOVIE_SIZE, Duration::from_secs(8))
                    .with_natural_size(MOVIE_NATURAL_SIZE),
            )
            .unwrap();
        editor
            .set_sheet_movie_title(
                source_sheet.object_id,
                source.drawable_object_id,
                "Native movie title",
            )
            .unwrap();
        editor
            .set_sheet_movie_caption(
                source_sheet.object_id,
                source.drawable_object_id,
                "Native movie caption",
            )
            .unwrap();

        let duplicate = editor
            .duplicate_sheet(test_sheet_selector(&editor, source_sheet.object_id))
            .unwrap();

        assert_eq!(duplicate.index, 1);
        assert_eq!(duplicate.name, "Movie-1");
        let copied = editor.sheet_movies(duplicate.object_id).unwrap().remove(0);
        assert_ne!(copied.drawable_object_id, source.drawable_object_id);
        assert_eq!(copied.sheet_id, duplicate.object_id);
        assert_eq!(copied.geometry, source.geometry);
        assert_eq!(copied.original_size, source.original_size);
        assert_eq!(copied.natural_size, source.natural_size);
        assert_eq!(copied.duration, source.duration);
        assert_eq!(copied.movie_data_identifier, source.movie_data_identifier);
        assert_eq!(
            copied.poster_image_data_identifier,
            source.poster_image_data_identifier
        );
        assert_eq!(
            editor
                .sheet_movie_title_caption(duplicate.object_id, copied.drawable_object_id)
                .unwrap(),
            crate::DrawableTitleCaption {
                title: Some("Native movie title".to_owned()),
                caption: Some("Native movie caption".to_owned()),
            }
        );
        assert_eq!(editor.media_assets().unwrap().len(), 2);

        let moved = DrawableGeometry {
            position: Some(MOVED_MOVIE_POSITION),
            ..copied.geometry
        };
        editor
            .set_sheet_movie_geometry(duplicate.object_id, copied.drawable_object_id, moved)
            .unwrap();
        let copied_playback = copied
            .playback
            .with_loop_mode(Some(MediaLoopMode::BackAndForth))
            .with_volume(Some(MediaVolume::new(0.4).unwrap()));
        editor
            .set_sheet_movie_playback_settings(
                duplicate.object_id,
                copied.drawable_object_id,
                copied_playback,
            )
            .unwrap();
        editor
            .set_sheet_movie_title(
                duplicate.object_id,
                copied.drawable_object_id,
                "Independent copy",
            )
            .unwrap();

        assert_eq!(
            editor
                .sheet_movie_geometry(source_sheet.object_id, source.drawable_object_id)
                .unwrap(),
            source.geometry
        );
        assert_eq!(
            editor
                .sheet_movie_playback_settings(source_sheet.object_id, source.drawable_object_id)
                .unwrap(),
            source.playback
        );
        assert_eq!(
            editor
                .sheet_movie_title_caption(source_sheet.object_id, source.drawable_object_id)
                .unwrap()
                .title
                .as_deref(),
            Some("Native movie title")
        );

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let reopened_copy = reopened
            .sheet_movies(duplicate.object_id)
            .unwrap()
            .remove(0);
        assert_eq!(reopened_copy.geometry, moved);
        assert_eq!(reopened_copy.playback, copied_playback);
        assert_eq!(
            reopened_copy.movie_data_identifier,
            source.movie_data_identifier
        );
        assert_eq!(
            reopened_copy.poster_image_data_identifier,
            source.poster_image_data_identifier
        );
        assert_eq!(reopened.media_assets().unwrap().len(), 2);
    }

    #[test]
    fn source_built_audio_sheet_duplicates_with_native_shared_asset_semantics() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Audio")
            .build()
            .unwrap();
        let source_sheet = editor.sheets().unwrap().remove(0);
        let source = editor
            .add_sheet_audio(
                source_sheet.object_id,
                "probe.aiff",
                AUDIO,
                NumbersSheetAudioOptions::new(AUDIO_POSITION, Duration::from_millis(2_200)),
            )
            .unwrap();

        let duplicate = editor
            .duplicate_sheet(test_sheet_selector(&editor, source_sheet.object_id))
            .unwrap();

        assert_eq!(duplicate.index, 1);
        assert_eq!(duplicate.name, "Audio-1");
        let copied = editor.sheet_audio(duplicate.object_id).unwrap().remove(0);
        assert_ne!(copied.drawable_object_id, source.drawable_object_id);
        assert_eq!(copied.sheet_id, duplicate.object_id);
        assert_eq!(copied.position, source.position);
        assert_eq!(copied.duration, source.duration);
        assert_eq!(copied.audio_data_identifier, source.audio_data_identifier);
        assert_eq!(editor.media_assets().unwrap().len(), 1);

        editor
            .set_sheet_audio_position(
                duplicate.object_id,
                copied.drawable_object_id,
                MOVED_AUDIO_POSITION,
            )
            .unwrap();
        let copied_playback = copied
            .playback
            .with_loop_mode(Some(MediaLoopMode::Repeat))
            .with_volume(Some(MediaVolume::new(0.4).unwrap()));
        editor
            .set_sheet_audio_playback_settings(
                duplicate.object_id,
                copied.drawable_object_id,
                copied_playback,
            )
            .unwrap();

        assert_eq!(
            editor
                .sheet_audio_position(source_sheet.object_id, source.drawable_object_id)
                .unwrap(),
            AUDIO_POSITION
        );
        assert_eq!(
            editor
                .sheet_audio_playback_settings(source_sheet.object_id, source.drawable_object_id)
                .unwrap(),
            source.playback
        );

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let reopened_copy = reopened.sheet_audio(duplicate.object_id).unwrap().remove(0);
        assert_eq!(reopened_copy.position, MOVED_AUDIO_POSITION);
        assert_eq!(reopened_copy.playback, copied_playback);
        assert_eq!(
            reopened_copy.audio_data_identifier,
            source.audio_data_identifier
        );
        assert_eq!(reopened.media_assets().unwrap().len(), 1);
    }

    #[test]
    fn source_built_shape_sheet_duplicates_with_native_independent_storage() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Shapes")
            .build()
            .unwrap();
        let source_sheet = editor.sheets().unwrap().remove(0);
        let source = editor
            .add_sheet_shape_with_fill(
                source_sheet.object_id,
                "Native-style shape",
                SHAPE_POSITION,
                SHAPE_SIZE,
                Preset::Rectangle,
                SOURCE_SHAPE_FILL,
            )
            .unwrap();

        let duplicate = editor
            .duplicate_sheet(test_sheet_selector(&editor, source_sheet.object_id))
            .unwrap();

        assert_eq!(duplicate.index, 1);
        assert_eq!(duplicate.name, "Shapes-1");
        let copied = editor.sheet_shapes(duplicate.object_id).unwrap().remove(0);
        assert_ne!(copied.drawable_object_id, source.drawable_object_id);
        assert_ne!(copied.storage.id, source.storage.id);
        assert_eq!(copied.sheet_id, duplicate.object_id);
        assert_eq!(copied.storage.storage, source.storage.storage);
        assert_eq!(copied.preset, source.preset);
        assert_eq!(copied.geometry, source.geometry);
        assert_eq!(
            editor
                .sheet_shape_fill(duplicate.object_id, copied.drawable_object_id)
                .unwrap(),
            SOURCE_SHAPE_FILL
        );

        editor
            .set_sheet_shape_text(
                duplicate.object_id,
                copied.drawable_object_id,
                "Independent copy",
            )
            .unwrap();
        editor
            .set_sheet_shape_geometry(
                duplicate.object_id,
                copied.drawable_object_id,
                DrawableGeometry {
                    position: Some(MOVED_SHAPE_POSITION),
                    ..copied.geometry
                },
            )
            .unwrap();
        editor
            .set_sheet_shape_fill(
                duplicate.object_id,
                copied.drawable_object_id,
                &COPIED_SHAPE_FILL,
            )
            .unwrap();

        let original = editor
            .sheet_shapes(source_sheet.object_id)
            .unwrap()
            .remove(0);
        assert_eq!(original.storage.storage.text(), "Native-style shape");
        assert_eq!(original.geometry, source.geometry);
        assert_eq!(
            editor
                .sheet_shape_fill(source_sheet.object_id, source.drawable_object_id)
                .unwrap(),
            SOURCE_SHAPE_FILL
        );

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let reopened_source = reopened
            .sheet_shapes(source_sheet.object_id)
            .unwrap()
            .remove(0);
        let reopened_copy = reopened
            .sheet_shapes(duplicate.object_id)
            .unwrap()
            .remove(0);
        assert_eq!(reopened_source.storage.storage.text(), "Native-style shape");
        assert_eq!(reopened_copy.storage.storage.text(), "Independent copy");
        assert_ne!(reopened_source.storage.id, reopened_copy.storage.id);
        assert_eq!(
            reopened
                .sheet_shape_fill(source_sheet.object_id, source.drawable_object_id)
                .unwrap(),
            SOURCE_SHAPE_FILL
        );
        assert_eq!(
            reopened
                .sheet_shape_fill(duplicate.object_id, copied.drawable_object_id)
                .unwrap(),
            COPIED_SHAPE_FILL
        );
    }

    #[test]
    fn source_built_image_sheet_duplicates_with_native_shared_asset_semantics() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Media")
            .build()
            .unwrap();
        let source_sheet = editor.sheets().unwrap().remove(0);
        let source = editor
            .add_sheet_image(
                source_sheet.object_id,
                "litchi_logo.png",
                include_bytes!("../../../../../media/litchi_logo.png"),
                NumbersSheetImageOptions::new(IMAGE_POSITION, IMAGE_SIZE),
            )
            .unwrap();

        let duplicate = editor
            .duplicate_sheet(test_sheet_selector(&editor, source_sheet.object_id))
            .unwrap();

        assert_eq!(duplicate.index, 1);
        assert_eq!(duplicate.name, "Media-1");
        let copied = editor.sheet_images(duplicate.object_id).unwrap().remove(0);
        assert_ne!(copied.drawable_object_id, source.drawable_object_id);
        assert_eq!(copied.sheet_id, duplicate.object_id);
        assert_eq!(copied.geometry, source.geometry);
        assert_eq!(copied.image_data_identifier, source.image_data_identifier);
        assert_eq!(
            copied.thumbnail_data_identifier,
            source.thumbnail_data_identifier
        );
        assert_eq!(editor.media_assets().unwrap().len(), 1);

        let moved = DrawableGeometry {
            position: Some(MOVED_IMAGE_POSITION),
            ..copied.geometry
        };
        editor
            .set_sheet_image_geometry(duplicate.object_id, copied.drawable_object_id, moved)
            .unwrap();
        assert_eq!(
            editor
                .sheet_image_geometry(source_sheet.object_id, source.drawable_object_id)
                .unwrap(),
            source.geometry
        );
        assert_eq!(
            editor
                .sheet_image_geometry(duplicate.object_id, copied.drawable_object_id)
                .unwrap(),
            moved
        );

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.sheet_images(source_sheet.object_id).unwrap().len(),
            1
        );
        assert_eq!(reopened.sheet_images(duplicate.object_id).unwrap().len(), 1);
        assert_eq!(reopened.media_assets().unwrap().len(), 1);
    }
}
