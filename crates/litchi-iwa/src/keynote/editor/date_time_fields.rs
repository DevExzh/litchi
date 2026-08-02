//! Date & Time smart-field CRUD for Keynote text boxes.

use super::KeynoteEditor;
use crate::Result;
use crate::text::{
    TextDateTimeDisplayText, TextDateTimeField, TextDateTimeFieldId, TextDateTimeFieldSettings,
    TextPosition, TextRange,
};

impl KeynoteEditor {
    /// Read every native Date & Time field in one Keynote text box.
    pub fn slide_text_box_date_time_fields(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Vec<TextDateTimeField>> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        self.text.text_date_time_fields(graph.storage_id)
    }

    /// Attach a Date & Time field to existing text in a Keynote text box.
    pub fn add_slide_text_box_date_time_field(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let field = staged.add_text_date_time_field(graph.storage_id, range, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(field)
    }

    /// Atomically insert visible text and a Date & Time field in a Keynote text box.
    pub fn insert_slide_text_box_date_time_field(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        position: TextPosition,
        display_text: TextDateTimeDisplayText,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let field = staged.insert_text_date_time_field(
            graph.storage_id,
            position,
            display_text,
            settings,
        )?;
        *self = Self::from_package(staged.into_package())?;
        Ok(field)
    }

    /// Atomically update a Keynote text-box Date & Time field.
    pub fn update_slide_text_box_date_time_field(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let field = staged.update_text_date_time_field(graph.storage_id, id, range, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(field)
    }

    /// Delete a Keynote text-box Date & Time field while retaining its text.
    pub fn remove_slide_text_box_date_time_field(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        id: TextDateTimeFieldId,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(slide_index, drawable_object_id)?;
        let mut staged = self.text.clone();
        let field = staged.remove_text_date_time_field(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(field)
    }
}
