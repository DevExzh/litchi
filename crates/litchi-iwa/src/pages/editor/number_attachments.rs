//! Native page-number/page-count attachment CRUD for Pages text storage.

use super::PagesEditor;
use crate::Result;
use crate::text::{
    TextNumberAttachment, TextNumberAttachmentId, TextNumberAttachmentSettings, TextPosition,
};

impl PagesEditor {
    /// Read every number attachment in the main Pages body.
    pub fn body_number_attachments(&self) -> Result<Vec<TextNumberAttachment>> {
        self.text.text_number_attachments(self.body_storage_id)
    }

    /// Atomically insert a native page number or page count in the body.
    pub fn insert_body_number_attachment(
        &mut self,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        self.text
            .insert_text_number_attachment(self.body_storage_id, position, settings)
    }

    /// Atomically update a body number attachment's payload.
    pub fn update_body_number_attachment(
        &mut self,
        id: TextNumberAttachmentId,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        self.text
            .update_text_number_attachment(self.body_storage_id, id, settings)
    }

    /// Delete a body number attachment and its U+FFFC placeholder.
    pub fn remove_body_number_attachment(
        &mut self,
        id: TextNumberAttachmentId,
    ) -> Result<TextNumberAttachment> {
        self.text
            .remove_text_number_attachment(self.body_storage_id, id)
    }

    /// Read every number attachment in a reachable header/footer storage.
    pub fn header_footer_number_attachments(
        &self,
        storage_id: u64,
    ) -> Result<Vec<TextNumberAttachment>> {
        self.require_header_footer(storage_id)?;
        self.text.text_number_attachments(storage_id)
    }

    /// Atomically insert a native page number or page count in a header/footer.
    pub fn insert_header_footer_number_attachment(
        &mut self,
        storage_id: u64,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        self.require_header_footer(storage_id)?;
        self.text
            .insert_text_number_attachment(storage_id, position, settings)
    }

    /// Atomically update a reachable header/footer number attachment.
    pub fn update_header_footer_number_attachment(
        &mut self,
        storage_id: u64,
        id: TextNumberAttachmentId,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        self.require_header_footer(storage_id)?;
        self.text
            .update_text_number_attachment(storage_id, id, settings)
    }

    /// Delete a header/footer number attachment and its U+FFFC placeholder.
    pub fn remove_header_footer_number_attachment(
        &mut self,
        storage_id: u64,
        id: TextNumberAttachmentId,
    ) -> Result<TextNumberAttachment> {
        self.require_header_footer(storage_id)?;
        self.text.remove_text_number_attachment(storage_id, id)
    }

    /// Read every number attachment in one ordinary Pages text box.
    pub fn text_box_number_attachments(
        &self,
        drawable_object_id: u64,
    ) -> Result<Vec<TextNumberAttachment>> {
        let graph = self.text_box_graph(drawable_object_id)?;
        self.text.text_number_attachments(graph.storage_id)
    }

    /// Atomically insert a number attachment in an ordinary Pages text box.
    pub fn insert_text_box_number_attachment(
        &mut self,
        drawable_object_id: u64,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let attachment =
            staged.insert_text_number_attachment(graph.storage_id, position, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(attachment)
    }

    /// Atomically update a Pages text-box number attachment.
    pub fn update_text_box_number_attachment(
        &mut self,
        drawable_object_id: u64,
        id: TextNumberAttachmentId,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let attachment = staged.update_text_number_attachment(graph.storage_id, id, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(attachment)
    }

    /// Delete a Pages text-box number attachment and its placeholder.
    pub fn remove_text_box_number_attachment(
        &mut self,
        drawable_object_id: u64,
        id: TextNumberAttachmentId,
    ) -> Result<TextNumberAttachment> {
        let graph = self.text_box_graph(drawable_object_id)?;
        let mut staged = self.text.clone();
        let attachment = staged.remove_text_number_attachment(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(attachment)
    }
}
