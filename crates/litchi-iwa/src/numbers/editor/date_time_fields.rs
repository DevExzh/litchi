//! Date & Time smart-field CRUD for Numbers text boxes.

use super::{IWorkTextEditor, NumbersEditor, numbers_text_box_graph};
use crate::Result;
use crate::text::{TextDateTimeField, TextDateTimeFieldId, TextPosition, TextRange};
use litchi_iwa_text::date_time::{DisplayText, Settings};

impl NumbersEditor {
    /// Read every native Date & Time field in one Numbers text box.
    pub fn sheet_text_box_date_time_fields(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextDateTimeField>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_date_time_fields(graph.storage_id)
    }

    /// Attach a Date & Time field to existing text in a Numbers text box.
    pub fn add_sheet_text_box_date_time_field(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: TextRange,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let field = text.add_text_date_time_field(graph.storage_id, range, settings)?;
        *self = Self::from_package(text.into_package())?;
        Ok(field)
    }

    /// Atomically insert visible text and a Date & Time field in a Numbers text box.
    pub fn insert_sheet_text_box_date_time_field(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: TextPosition,
        display_text: DisplayText,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let field =
            text.insert_text_date_time_field(graph.storage_id, position, display_text, settings)?;
        *self = Self::from_package(text.into_package())?;
        Ok(field)
    }

    /// Atomically update a Numbers text-box Date & Time field.
    pub fn update_sheet_text_box_date_time_field(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let field = text.update_text_date_time_field(graph.storage_id, id, range, settings)?;
        *self = Self::from_package(text.into_package())?;
        Ok(field)
    }

    /// Delete a Numbers text-box Date & Time field while retaining its text.
    pub fn remove_sheet_text_box_date_time_field(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextDateTimeFieldId,
    ) -> Result<TextDateTimeField> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let field = text.remove_text_date_time_field(graph.storage_id, id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(field)
    }
}
