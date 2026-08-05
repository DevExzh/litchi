//! Transactional deletion of complete Numbers sheet object graphs.

use super::*;

#[derive(Debug, Default)]
struct SheetContents {
    charts: Vec<u64>,
    images: Vec<u64>,
    movies: Vec<u64>,
    audio: Vec<u64>,
    shapes: Vec<u64>,
    text_boxes: Vec<u64>,
    tables: Vec<u64>,
}

impl SheetContents {
    fn direct_object_ids(&self, editor: &NumbersEditor) -> Result<HashSet<u64>> {
        let mut identifiers = HashSet::new();
        for &identifier in self
            .charts
            .iter()
            .chain(&self.images)
            .chain(&self.movies)
            .chain(&self.audio)
            .chain(&self.shapes)
            .chain(&self.text_boxes)
        {
            if !identifiers.insert(identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet drawable {identifier} has multiple semantic representations"
                )));
            }
        }
        for &table_id in &self.tables {
            let owner = find_table_owner(editor.package(), table_id)?;
            if !identifiers.insert(owner.table_info_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers sheet drawable {} has multiple semantic representations",
                    owner.table_info_id
                )));
            }
        }
        Ok(identifiers)
    }
}

impl NumbersEditor {
    /// Remove a sheet and every supported object graph owned exclusively by it.
    ///
    /// Tables, charts, media, shapes, and text boxes are deleted through their
    /// regular typed lifecycle paths before the sheet itself is detached. This
    /// reclaims private objects, UUID registrations, component references, and
    /// unshared media assets. Unknown drawable kinds and incoming formula edges
    /// reject the operation transactionally instead of leaving unreachable data.
    pub fn remove_sheet(&mut self, sheet_id: u64) -> Result<NumbersSheetInfo> {
        let sheets = self.sheets()?;
        if sheets.len() <= 1 {
            return Err(Error::ParseError(
                "Cannot remove the final Numbers sheet".to_owned(),
            ));
        }
        let removed = sheets
            .iter()
            .find(|sheet| sheet.object_id == sheet_id)
            .cloned()
            .ok_or_else(|| Error::ParseError(format!("Numbers sheet {sheet_id} not found")))?;
        let contents = sheet_contents(self, sheet_id)?;

        let mut working = self.clone();
        delete_sheet_contents(&mut working, sheet_id, contents)?;
        detach_empty_sheet(&mut working, sheet_id)?;

        let verified = Self::from_bytes(&working.to_bytes()?)?;
        if verified
            .sheets()?
            .iter()
            .any(|sheet| sheet.object_id == sheet_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers sheet deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(removed)
    }
}

fn sheet_contents(editor: &NumbersEditor, sheet_id: u64) -> Result<SheetContents> {
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    let contents = SheetContents {
        charts: editor
            .sheet_charts(sheet_id)?
            .into_iter()
            .map(|chart| chart.drawable_object_id)
            .collect(),
        images: editor
            .sheet_images(sheet_id)?
            .into_iter()
            .map(|image| image.drawable_object_id)
            .collect(),
        movies: editor
            .sheet_movies(sheet_id)?
            .into_iter()
            .map(|movie| movie.drawable_object_id)
            .collect(),
        audio: editor
            .sheet_audio(sheet_id)?
            .into_iter()
            .map(|audio| audio.drawable_object_id)
            .collect(),
        shapes: editor
            .sheet_shapes(sheet_id)?
            .into_iter()
            .map(|shape| shape.drawable_object_id)
            .collect(),
        text_boxes: editor
            .sheet_text_boxes(sheet_id)?
            .into_iter()
            .map(|text_box| text_box.drawable_object_id)
            .collect(),
        tables: table_models(editor.package())?
            .into_iter()
            .map(|table| {
                let owner = find_table_owner(editor.package(), table.object_id)?;
                Ok((table.object_id, owner.sheet_id))
            })
            .collect::<Result<Vec<_>>>()?
            .into_iter()
            .filter_map(|(table_id, owner_sheet_id)| {
                (owner_sheet_id == sheet_id).then_some(table_id)
            })
            .collect(),
    };
    let expected = sheet
        .drawable_infos
        .into_iter()
        .map(|reference| reference.identifier)
        .collect::<HashSet<_>>();
    let classified = contents.direct_object_ids(editor)?;
    if classified != expected {
        let mut unclassified = expected
            .difference(&classified)
            .copied()
            .collect::<Vec<_>>();
        unclassified.sort_unstable();
        let mut unexpected = classified
            .difference(&expected)
            .copied()
            .collect::<Vec<_>>();
        unexpected.sort_unstable();
        return Err(Error::ParseError(format!(
            "Cannot remove Numbers sheet {sheet_id}: unclassified drawables {unclassified:?}, unexpected classified drawables {unexpected:?}"
        )));
    }
    Ok(contents)
}

fn delete_sheet_contents(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    contents: SheetContents,
) -> Result<()> {
    for identifier in contents.charts {
        editor.remove_sheet_chart(sheet_id, identifier)?;
    }
    for identifier in contents.images {
        editor.remove_sheet_image(sheet_id, identifier)?;
    }
    for identifier in contents.movies {
        editor.remove_sheet_movie(sheet_id, identifier)?;
    }
    for identifier in contents.audio {
        editor.remove_sheet_audio(sheet_id, identifier)?;
    }
    for identifier in contents.shapes {
        editor.remove_sheet_shape(sheet_id, identifier)?;
    }
    for identifier in contents.text_boxes {
        editor.remove_sheet_text_box(sheet_id, identifier)?;
    }
    for identifier in contents.tables {
        editor.remove_table(identifier)?;
    }
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if !sheet.drawable_infos.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "Numbers sheet {sheet_id} retained drawables after graph deletion"
        )));
    }
    Ok(())
}

fn detach_empty_sheet(editor: &mut NumbersEditor, sheet_id: u64) -> Result<()> {
    let locations = object_locations(editor.package())?;
    let archive_name = locations
        .get(&sheet_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?
        .to_owned();
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?;
    let mut staged = editor.package.clone();
    update_numbers_document(&mut staged, |document| {
        let previous_len = document.sheets.len();
        document
            .sheets
            .retain(|reference| reference.identifier != sheet_id);
        if document.sheets.len() + 1 != previous_len {
            return Err(Error::InvalidFormat(format!(
                "Numbers root does not reference sheet {sheet_id} exactly once"
            )));
        }
        Ok(())
    })?;
    if let Some(component_id) = component_id {
        remove_component_external_references_to_object(&mut staged, component_id, sheet_id)?;
        if component_uuid_identifiers(&staged, component_id)?
            .is_some_and(|identifiers| identifiers.contains(&sheet_id))
        {
            remove_component_object_uuids(&mut staged, component_id, &[sheet_id])?;
        }
    }
    remove_object_or_empty_entry(&mut staged, &locations, sheet_id)?;
    if !staged.contains_entry(&archive_name)
        && let Some(component_id) = component_id
    {
        remove_component_registration(&mut staged, component_id)?;
    }
    release_package_identifier_suffix(&mut staged, &[sheet_id])?;
    editor.package = staged;
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::time::Duration;

    use super::*;
    use crate::charts::{ChartData, Kind};
    use crate::numbers::{
        NumbersDocumentBuilder, NumbersSheetAudioOptions, NumbersSheetImageOptions,
        NumbersSheetMovieOptions,
    };
    use crate::shapes::{DrawablePoint, DrawableSize};

    const POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 120.0,
    };
    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsheet-delete-audio";
    const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  sheet-delete-movie";
    const POSTER: &[u8] = b"\x89PNG\r\n\x1a\nsheet-delete-poster";
    const IMAGE: &[u8] = include_bytes!("../../../../../media/litchi_logo.png");

    fn object_ids(package: &IWorkPackage) -> HashSet<u64> {
        package
            .iwa_entry_names()
            .flat_map(|entry| package.archive(entry).unwrap().objects)
            .filter_map(|object| object.archive_info.identifier)
            .collect()
    }

    #[test]
    fn source_built_sheet_deletion_reclaims_every_supported_private_graph() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let baseline_ids = object_ids(editor.package());
        let baseline_entries = editor
            .package()
            .entry_names()
            .map(str::to_owned)
            .collect::<HashSet<_>>();
        let sheet = editor.add_empty_sheet("Disposable").unwrap();
        editor
            .add_empty_table(sheet.object_id, "Data", 3, 2)
            .unwrap();
        editor
            .add_sheet_text_box(sheet.object_id, "Text", POSITION, SIZE)
            .unwrap();
        editor
            .add_sheet_rectangle(sheet.object_id, "Shape", POSITION, SIZE)
            .unwrap();
        editor
            .add_sheet_image(
                sheet.object_id,
                "logo.png",
                IMAGE,
                NumbersSheetImageOptions::new(POSITION, SIZE),
            )
            .unwrap();
        editor
            .add_sheet_audio(
                sheet.object_id,
                "audio.aiff",
                AUDIO,
                NumbersSheetAudioOptions::new(POSITION, Duration::from_secs(1)),
            )
            .unwrap();
        editor
            .add_sheet_movie(
                sheet.object_id,
                "movie.mov",
                MOVIE,
                "poster.png",
                POSTER,
                NumbersSheetMovieOptions::new(POSITION, SIZE, Duration::from_secs(1)),
            )
            .unwrap();
        editor
            .add_sheet_chart(
                sheet.object_id,
                Kind::Column2d,
                ChartData::new(
                    vec!["Series".to_owned()],
                    vec!["Item".to_owned()],
                    vec![vec![Some(1.0)]],
                )
                .unwrap(),
                POSITION,
                SIZE,
            )
            .unwrap();

        let removed = editor.remove_sheet(sheet.object_id).unwrap();
        assert_eq!(removed, sheet);
        assert_eq!(object_ids(editor.package()), baseline_ids);
        assert_eq!(
            editor
                .package()
                .entry_names()
                .map(str::to_owned)
                .collect::<HashSet<_>>(),
            baseline_entries
        );
        assert_eq!(editor.sheets().unwrap().len(), 1);
        assert_eq!(editor.tables().unwrap().len(), 1);
        assert!(editor.media_assets().unwrap().is_empty());
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn sheet_deletion_is_transactional_for_the_final_and_missing_sheet() {
        let mut editor = NumbersEditor::create().unwrap();
        let baseline = editor.to_bytes().unwrap();
        let only_sheet = editor.sheets().unwrap()[0].object_id;

        assert!(editor.remove_sheet(only_sheet).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor.add_empty_sheet("Second").unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(editor.remove_sheet(u64::MAX).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }
}
