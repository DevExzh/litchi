//! Native page-number/page-count attachment CRUD for Pages text storage.
//!
//! Text-box selectors use [`DrawableObjectId`] as their canonical type. The
//! checked `TryInto` boundary preserves the existing raw-object-id call path
//! during migration without allowing zero or otherwise invalid identifiers.

use std::fmt;

use super::PagesEditor;
use crate::Result;
use crate::comments::DrawableObjectId;
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
    pub fn text_box_number_attachments<D>(
        &self,
        drawable_object_id: D,
    ) -> Result<Vec<TextNumberAttachment>>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_text_box_drawable_id(drawable_object_id)?;
        let graph = self.text_box_graph(drawable_object_id.object_id())?;
        self.text.text_number_attachments(graph.storage_id)
    }

    /// Atomically insert a number attachment in an ordinary Pages text box.
    pub fn insert_text_box_number_attachment<D>(
        &mut self,
        drawable_object_id: D,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_text_box_drawable_id(drawable_object_id)?;
        let graph = self.text_box_graph(drawable_object_id.object_id())?;
        let mut staged = self.text.clone();
        let attachment =
            staged.insert_text_number_attachment(graph.storage_id, position, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(attachment)
    }

    /// Atomically update a Pages text-box number attachment.
    pub fn update_text_box_number_attachment<D>(
        &mut self,
        drawable_object_id: D,
        id: TextNumberAttachmentId,
        settings: TextNumberAttachmentSettings,
    ) -> Result<TextNumberAttachment>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_text_box_drawable_id(drawable_object_id)?;
        let graph = self.text_box_graph(drawable_object_id.object_id())?;
        let mut staged = self.text.clone();
        let attachment = staged.update_text_number_attachment(graph.storage_id, id, settings)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(attachment)
    }

    /// Delete a Pages text-box number attachment and its placeholder.
    pub fn remove_text_box_number_attachment<D>(
        &mut self,
        drawable_object_id: D,
        id: TextNumberAttachmentId,
    ) -> Result<TextNumberAttachment>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_text_box_drawable_id(drawable_object_id)?;
        let graph = self.text_box_graph(drawable_object_id.object_id())?;
        let mut staged = self.text.clone();
        let attachment = staged.remove_text_number_attachment(graph.storage_id, id)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(attachment)
    }
}

fn normalize_text_box_drawable_id<D>(value: D) -> Result<DrawableObjectId>
where
    D: TryInto<DrawableObjectId>,
    D::Error: fmt::Debug,
{
    value.try_into().map_err(|error| {
        crate::Error::ParseError(format!(
            "invalid Pages text-box drawable selector: {error:?}"
        ))
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use crate::text::TextNumberAttachmentKind;

    const POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 72.0,
    };

    fn pages_with_text_box() -> (PagesEditor, DrawableObjectId) {
        let mut pages = PagesEditor::create_with_text("Body").unwrap();
        let text_box = pages.add_text_box(4, "Box", POSITION, SIZE).unwrap();
        let selector = DrawableObjectId::try_from(text_box.drawable_object_id).unwrap();
        (pages, selector)
    }

    fn settings() -> TextNumberAttachmentSettings {
        TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageNumber)
    }

    #[test]
    fn typed_selector_supports_text_box_attachment_crud() {
        let (mut pages, selector) = pages_with_text_box();
        assert!(
            pages
                .text_box_number_attachments(selector)
                .unwrap()
                .is_empty()
        );

        let attachment = pages
            .insert_text_box_number_attachment(
                selector,
                TextPosition::from_utf16_index(3).unwrap(),
                settings(),
            )
            .unwrap();
        assert_eq!(
            pages.text_box_number_attachments(selector).unwrap(),
            vec![attachment.clone()]
        );

        let updated = pages
            .update_text_box_number_attachment(
                selector,
                attachment.id,
                TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageCount),
            )
            .unwrap();
        assert_eq!(updated.id, attachment.id);
        assert_eq!(updated.settings.kind, TextNumberAttachmentKind::PageCount);
        assert_eq!(
            pages
                .remove_text_box_number_attachment(selector, updated.id)
                .unwrap(),
            updated
        );
        assert!(
            pages
                .text_box_number_attachments(selector)
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn zero_and_missing_selectors_fail_without_changing_bytes() {
        let (mut pages, selector) = pages_with_text_box();
        let baseline = pages.to_bytes().unwrap();

        assert!(pages.text_box_number_attachments(0_u64).is_err());
        assert_eq!(pages.to_bytes().unwrap(), baseline);
        assert!(
            pages
                .text_box_number_attachments(DrawableObjectId::try_from(999_999_u64).unwrap())
                .is_err()
        );
        assert_eq!(pages.to_bytes().unwrap(), baseline);
        assert!(
            pages
                .insert_text_box_number_attachment(
                    0_u64,
                    TextPosition::from_utf16_index(0).unwrap(),
                    settings(),
                )
                .is_err()
        );
        assert_eq!(pages.to_bytes().unwrap(), baseline);

        let attachment = pages
            .insert_text_box_number_attachment(
                selector,
                TextPosition::from_utf16_index(3).unwrap(),
                settings(),
            )
            .unwrap();
        let after_insert = pages.to_bytes().unwrap();
        let missing_attachment = TextNumberAttachmentId::from_object_id(999_999).unwrap();
        assert!(
            pages
                .update_text_box_number_attachment(
                    selector,
                    missing_attachment,
                    TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageCount),
                )
                .is_err()
        );
        assert_eq!(pages.to_bytes().unwrap(), after_insert);
        assert!(
            pages
                .remove_text_box_number_attachment(selector, missing_attachment)
                .is_err()
        );
        assert_eq!(pages.to_bytes().unwrap(), after_insert);
        assert!(
            pages
                .insert_text_box_number_attachment(
                    selector,
                    TextPosition::from_utf16_index(5).unwrap(),
                    settings(),
                )
                .is_err()
        );
        assert_eq!(pages.to_bytes().unwrap(), after_insert);
        assert_eq!(
            pages.text_box_number_attachments(selector).unwrap(),
            vec![attachment]
        );
    }
}
