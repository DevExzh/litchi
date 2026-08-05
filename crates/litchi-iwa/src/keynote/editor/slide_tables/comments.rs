//! Transactional comment-thread CRUD for Keynote slide-table cells.

use super::*;
use litchi_iwa_common::comment::{TableCellComment, TableCellReply};

impl KeynoteEditor {
    /// Read the comment attached to a reachable slide-table cell.
    pub fn slide_table_cell_comment(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<TableCellComment>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_comment_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a slide-table cell comment transactionally.
    pub fn set_slide_table_cell_comment(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let text = text.into();
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_cell_comment_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            text.clone(),
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let stored = verified
            .slide_table_cell_comment(slide_index, model_object_id, row, column)?
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote table comment was not persisted".to_owned())
            })?;
        if stored.comment.text != text {
            return Err(Error::InvalidFormat(
                "Keynote table comment failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Delete a slide-table cell comment and its reply thread transactionally.
    pub fn clear_slide_table_cell_comment(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::clear_table_cell_comment_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slide_table_cell_comment(slide_index, model_object_id, row, column)?
            .is_some()
        {
            return Err(Error::InvalidFormat(
                "Keynote table comment deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read direct replies to a slide-table cell comment in stored order.
    pub fn slide_table_cell_comment_replies(
        &self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<TableCellReply>> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_cell_comment_replies_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Append a direct reply to an existing slide-table cell comment.
    pub fn add_slide_table_cell_comment_reply(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<u64> {
        require_table_model(self, slide_index, model_object_id)?;
        let text = text.into();
        let mut staged = self.package().clone();
        let reply_id = crate::numbers::editor::add_table_cell_comment_reply_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            text.clone(),
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_reply(
            &verified,
            slide_index,
            model_object_id,
            row,
            column,
            reply_id,
            &text,
        )?;
        *self = verified;
        Ok(reply_id)
    }

    /// Replace one direct reply and return its new copy-on-write object ID.
    #[allow(clippy::too_many_arguments)]
    pub fn set_slide_table_cell_comment_reply(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        require_table_model(self, slide_index, model_object_id)?;
        let text = text.into();
        let mut staged = self.package().clone();
        let new_reply_id = crate::numbers::editor::set_table_cell_comment_reply_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            reply_storage_object_id,
            text.clone(),
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_reply(
            &verified,
            slide_index,
            model_object_id,
            row,
            column,
            new_reply_id,
            &text,
        )?;
        *self = verified;
        Ok(new_reply_id)
    }

    /// Remove one direct reply from a slide-table cell comment.
    pub fn remove_slide_table_cell_comment_reply(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::remove_table_cell_comment_reply_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
            reply_storage_object_id,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slide_table_cell_comment_replies(slide_index, model_object_id, row, column)?
            .iter()
            .any(|reply| reply.storage_id.get() == reply_storage_object_id)
        {
            return Err(Error::InvalidFormat(
                "Keynote table comment reply deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[allow(clippy::too_many_arguments)]
fn require_reply(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
    row: usize,
    column: usize,
    reply_id: u64,
    text: &str,
) -> Result<()> {
    let valid = editor
        .slide_table_cell_comment_replies(slide_index, model_object_id, row, column)?
        .iter()
        .any(|reply| reply.storage_id.get() == reply_id && reply.comment.text == text);
    if valid {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "Keynote table comment reply failed validation".to_owned(),
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;

    #[test]
    fn scratch_table_comment_threads_support_full_crud() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let table = editor
            .add_slide_table(
                0,
                "Review",
                2,
                2,
                DrawablePoint { x: 120.0, y: 180.0 },
                DrawableSize {
                    width: 840.0,
                    height: 360.0,
                },
            )
            .unwrap();
        let table_id = table.model_object_id;

        editor
            .set_slide_table_cell_comment(0, table_id, 1, 1, "Initial review")
            .unwrap();
        editor
            .set_slide_table_cell_comment(0, table_id, 1, 1, "Updated review")
            .unwrap();
        let root = editor
            .slide_table_cell_comment(0, table_id, 1, 1)
            .unwrap()
            .unwrap();
        assert_eq!(root.comment.text, "Updated review");
        assert!(root.comment.creation_date_seconds.is_some());
        assert!(root.comment.storage_uuid.is_some());
        assert_eq!(
            editor
                .slide_table(0, table_id)
                .unwrap()
                .get_comment(1, 1)
                .unwrap()
                .text,
            "Updated review"
        );

        let reply_id = editor
            .add_slide_table_cell_comment_reply(0, table_id, 1, 1, "First reply")
            .unwrap();
        let updated_reply_id = editor
            .set_slide_table_cell_comment_reply(0, table_id, 1, 1, reply_id, "Revised reply")
            .unwrap();
        assert_ne!(updated_reply_id, reply_id);
        assert_eq!(
            editor
                .slide_table_cell_comment_replies(0, table_id, 1, 1)
                .unwrap()[0]
                .comment
                .text,
            "Revised reply"
        );

        let before_invalid = editor.to_bytes().unwrap();
        assert!(
            editor
                .remove_slide_table_cell_comment_reply(0, table_id, 1, 1, u64::MAX)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_invalid);

        editor
            .remove_slide_table_cell_comment_reply(0, table_id, 1, 1, updated_reply_id)
            .unwrap();
        assert!(
            editor
                .slide_table_cell_comment_replies(0, table_id, 1, 1)
                .unwrap()
                .is_empty()
        );
        editor
            .clear_slide_table_cell_comment(0, table_id, 1, 1)
            .unwrap();
        assert!(
            editor
                .slide_table_cell_comment(0, table_id, 1, 1)
                .unwrap()
                .is_none()
        );

        let reparsed = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reparsed
                .slide_table_cell_comment(0, table_id, 1, 1)
                .unwrap()
                .is_none()
        );
    }
}
