//! Named paragraph-style lifecycle for sheet-owned text boxes.

use super::{NumbersEditor, numbers_text_box_graph};
use crate::text::{
    AppliedParagraphStyle, IWorkTextEditor, NamedParagraphStyle, ParagraphStyleId,
    ParagraphStyleName, applied_named_paragraph_style_in_storage,
    named_paragraph_styles_in_storage,
};
use crate::{Error, Result};

impl NumbersEditor {
    /// List named paragraph styles selectable for a sheet-owned text box.
    pub fn sheet_text_box_named_paragraph_styles(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<NamedParagraphStyle>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        named_paragraph_styles_in_storage(&self.package, graph.storage_id.get())
    }

    /// Read the named paragraph style selected for a sheet-owned text box.
    pub fn sheet_text_box_applied_named_paragraph_style(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<AppliedParagraphStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        applied_named_paragraph_style_in_storage(&self.package, graph.storage_id.get())
    }

    /// Redefine the selected named style from this text box's direct overrides.
    pub fn redefine_applied_sheet_text_box_named_paragraph_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<NamedParagraphStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let redefined = text.redefine_applied_named_paragraph_style(graph.storage_id)?;
        let verified = Self::from_package(text.into_package())?;
        let selection =
            verified.sheet_text_box_applied_named_paragraph_style(sheet_id, drawable_object_id)?;
        if selection.style() != &redefined || selection.has_overrides() {
            return Err(Error::InvalidFormat(
                "Numbers named paragraph-style redefinition failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(redefined)
    }

    /// Create a named paragraph style by cloning one selectable preset.
    pub fn create_sheet_text_box_named_paragraph_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        source: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let created = text.create_named_paragraph_style(graph.storage_id, source, name)?;
        let verified = Self::from_package(text.into_package())?;
        if !verified
            .sheet_text_box_named_paragraph_styles(sheet_id, drawable_object_id)?
            .contains(&created)
        {
            return Err(Error::InvalidFormat(
                "Numbers named paragraph-style creation failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Rename a selectable paragraph style without changing its identifier.
    pub fn rename_sheet_text_box_named_paragraph_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        target: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let renamed = text.rename_named_paragraph_style(graph.storage_id, target, name)?;
        let verified = Self::from_package(text.into_package())?;
        if !verified
            .sheet_text_box_named_paragraph_styles(sheet_id, drawable_object_id)?
            .contains(&renamed)
        {
            return Err(Error::InvalidFormat(
                "Numbers named paragraph-style rename failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(renamed)
    }

    /// Apply one named paragraph style and clear direct paragraph overrides.
    pub fn apply_sheet_text_box_named_paragraph_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        target: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let applied = text.apply_named_paragraph_style(graph.storage_id, target)?;
        let verified = Self::from_package(text.into_package())?;
        let selection =
            verified.sheet_text_box_applied_named_paragraph_style(sheet_id, drawable_object_id)?;
        if selection.style() != &applied || selection.has_overrides() {
            return Err(Error::InvalidFormat(
                "Numbers named paragraph-style application failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(applied)
    }

    /// Delete one unused named paragraph style.
    pub fn delete_sheet_text_box_named_paragraph_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        target: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let deleted = text.delete_named_paragraph_style(graph.storage_id, target)?;
        let verified = Self::from_package(text.into_package())?;
        if verified
            .sheet_text_box_named_paragraph_styles(sheet_id, drawable_object_id)?
            .iter()
            .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "Numbers named paragraph-style deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(deleted)
    }

    /// Replace the applied style and delete it as one transaction.
    pub fn delete_applied_sheet_text_box_named_paragraph_style_with_replacement(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        target: ParagraphStyleId,
        replacement: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let deleted = text.delete_applied_named_paragraph_style_with_replacement(
            graph.storage_id,
            target,
            replacement,
        )?;
        let verified = Self::from_package(text.into_package())?;
        let selection =
            verified.sheet_text_box_applied_named_paragraph_style(sheet_id, drawable_object_id)?;
        if selection.style().id() != replacement
            || selection.has_overrides()
            || verified
                .sheet_text_box_named_paragraph_styles(sheet_id, drawable_object_id)?
                .iter()
                .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "Numbers paragraph-style replacement deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(deleted)
    }
}
