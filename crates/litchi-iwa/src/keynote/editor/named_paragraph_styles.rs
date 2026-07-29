//! Named paragraph-style lifecycle for ordinary slide text boxes.

use super::KeynoteEditor;
use crate::text::{
    AppliedParagraphStyle, NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName,
};
use crate::{Error, Result};

impl KeynoteEditor {
    /// List named paragraph styles selectable for an ordinary slide text box.
    pub fn slide_text_box_named_paragraph_styles(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<NamedParagraphStyle>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.named_paragraph_styles(graph.storage_id)
    }

    /// Read the named paragraph style selected for an ordinary slide text box.
    pub fn slide_text_box_applied_named_paragraph_style(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<AppliedParagraphStyle> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.applied_named_paragraph_style(graph.storage_id)
    }

    /// Create a named paragraph style by cloning one selectable preset.
    pub fn create_slide_text_box_named_paragraph_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        source: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let created = staged.create_named_paragraph_style(graph.storage_id, source, name)?;
        let verified = Self::from_package(staged.package().clone())?;
        if !verified
            .slide_text_box_named_paragraph_styles(slide_index, drawable_object_id)?
            .contains(&created)
        {
            return Err(Error::InvalidFormat(
                "Keynote named paragraph-style creation failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(created)
    }

    /// Rename a selectable paragraph style without changing its identifier.
    pub fn rename_slide_text_box_named_paragraph_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        target: ParagraphStyleId,
        name: ParagraphStyleName,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let renamed = staged.rename_named_paragraph_style(graph.storage_id, target, name)?;
        let verified = Self::from_package(staged.package().clone())?;
        if !verified
            .slide_text_box_named_paragraph_styles(slide_index, drawable_object_id)?
            .contains(&renamed)
        {
            return Err(Error::InvalidFormat(
                "Keynote named paragraph-style rename failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(renamed)
    }

    /// Apply one named paragraph style and clear direct paragraph overrides.
    pub fn apply_slide_text_box_named_paragraph_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        target: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let applied = staged.apply_named_paragraph_style(graph.storage_id, target)?;
        let verified = Self::from_package(staged.package().clone())?;
        let selection = verified
            .slide_text_box_applied_named_paragraph_style(slide_index, drawable_object_id)?;
        if selection.style() != &applied || selection.has_overrides() {
            return Err(Error::InvalidFormat(
                "Keynote named paragraph-style application failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(applied)
    }

    /// Delete one unused named paragraph style.
    pub fn delete_slide_text_box_named_paragraph_style(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        target: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let deleted = staged.delete_named_paragraph_style(graph.storage_id, target)?;
        let verified = Self::from_package(staged.package().clone())?;
        if verified
            .slide_text_box_named_paragraph_styles(slide_index, drawable_object_id)?
            .iter()
            .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "Keynote named paragraph-style deletion failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(deleted)
    }

    /// Replace the applied style and delete it as one transaction.
    pub fn delete_applied_slide_text_box_named_paragraph_style_with_replacement(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        target: ParagraphStyleId,
        replacement: ParagraphStyleId,
    ) -> Result<NamedParagraphStyle> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let deleted = staged.delete_applied_named_paragraph_style_with_replacement(
            graph.storage_id,
            target,
            replacement,
        )?;
        let verified = Self::from_package(staged.package().clone())?;
        let selection = verified
            .slide_text_box_applied_named_paragraph_style(slide_index, drawable_object_id)?;
        if selection.style().id() != replacement
            || selection.has_overrides()
            || verified
                .slide_text_box_named_paragraph_styles(slide_index, drawable_object_id)?
                .iter()
                .any(|style| style.id() == target)
        {
            return Err(Error::InvalidFormat(
                "Keynote paragraph-style replacement deletion failed validation".to_owned(),
            ));
        }
        self.text = staged;
        Ok(deleted)
    }
}
