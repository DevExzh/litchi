//! Date & Time smart-field CRUD for Pages text boxes.

use super::PagesEditor;
use crate::Result;
use crate::text::{
    TextDateTimeDisplayText, TextDateTimeField, TextDateTimeFieldId, TextDateTimeFieldSettings,
    TextPosition, TextRange,
};

impl PagesEditor {
    /// Read every native Date & Time field in one Pages text box.
    pub fn text_box_date_time_fields(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<TextDateTimeField>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_date_time_fields(graph.storage_id)
    }

    /// Attach a Date & Time field to existing text in a Pages text box.
    pub fn add_text_box_date_time_field(
        &mut self,
        drawable_object_id: u64,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let field = staged.add_text_date_time_field(graph.storage_id, range, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(field)
    }

    /// Atomically insert visible text and a Date & Time field in a Pages text box.
    pub fn insert_text_box_date_time_field(
        &mut self,
        drawable_object_id: u64,
        position: TextPosition,
        display_text: TextDateTimeDisplayText,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(drawable_object_id)?;
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

    /// Atomically update a Pages text-box Date & Time field.
    pub fn update_text_box_date_time_field(
        &mut self,
        drawable_object_id: u64,
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let field = staged.update_text_date_time_field(graph.storage_id, id, range, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(field)
    }

    /// Delete a Pages text-box Date & Time field while retaining its text.
    pub fn remove_text_box_date_time_field(
        &mut self,
        drawable_object_id: u64,
        id: TextDateTimeFieldId,
    ) -> Result<TextDateTimeField> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let field = staged.remove_text_date_time_field(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(field)
    }
}
