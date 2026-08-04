//! Ordinary text-box graph duplication within or across Numbers sheets.

use super::*;

impl NumbersEditor {
    /// Duplicate an ordinary sheet-owned text box with independent storage.
    ///
    /// The shape, stand-in title and caption, and writable storage are cloned
    /// into the document component with fresh object identifiers and UUIDs.
    /// The clone is appended to the sheet's drawable list and offset by ten
    /// points, matching Numbers' native duplicate behavior.
    pub fn duplicate_sheet_text_box(
        &mut self,
        sheet_id: u64,
        source_drawable_object_id: u64,
        text: &str,
    ) -> Result<NumbersTextBoxInfo> {
        self.duplicate_text_box_to_sheet(
            sheet_id,
            source_drawable_object_id,
            sheet_id,
            text,
            DRAWABLE_DUPLICATE_OFFSET,
        )
    }

    pub(super) fn duplicate_text_box_to_sheet(
        &mut self,
        source_sheet_id: u64,
        source_drawable_object_id: u64,
        target_sheet_id: u64,
        text: &str,
        offset: f32,
    ) -> Result<NumbersTextBoxInfo> {
        if !offset.is_finite() {
            return Err(Error::ParseError(
                "Numbers text-box duplicate offset must be finite".to_owned(),
            ));
        }
        let source =
            numbers_text_box_graph(&self.package, source_sheet_id, source_drawable_object_id)?;
        let (target_archive_name, _, _) = numbers_sheet(&self.package, target_sheet_id)?;
        if target_archive_name != source.archive_name {
            return Err(Error::ParseError(format!(
                "Numbers text-box transfer from {} to {} crosses IWA components",
                source.archive_name, target_archive_name
            )));
        }

        let mut staged = self.package.clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len() + 1);
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Numbers text-box graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }
        if source_sheet_id != target_sheet_id {
            remap.insert(source_sheet_id, target_sheet_id);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers text-box object {identifier} is missing"))
                })?;
                clone_numbers_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                Ok(archive.insert_object(cloned)?)
            })?;
        }

        let new_drawable_id = remap[&source.drawable_id];
        let new_storage_id = remap[&source.storage_id];
        if offset != f32::default() {
            offset_numbers_drawable_clone(
                &mut staged,
                &source.archive_name,
                new_drawable_id,
                offset,
            )?;
        }
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.set_text(new_storage_id, text)?;
        staged = text_editor.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            target_sheet_id,
            None,
            Some(new_drawable_id),
        )?;
        let last_identifier = source
            .object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .max()
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers text-box graph has no object identifiers".to_owned())
            })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        add_component_object_uuids(
            &mut staged,
            DOCUMENT_COMPONENT_IDENTIFIER,
            &new_uuid_object_ids,
        )?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .sheet_text_boxes(target_sheet_id)?
            .into_iter()
            .find(|item| item.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers text-box duplication failed validation".to_owned())
            })?;
        let created_graph =
            numbers_text_box_graph(verified.package(), target_sheet_id, new_drawable_id)?;
        if created.storage.object_id != new_storage_id
            || created.storage.text != text
            || created_graph.object_ids.len() != source.object_ids.len()
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }
}
