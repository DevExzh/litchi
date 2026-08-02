//! Transactional comment-thread CRUD for Pages body-table cells.

use super::*;

/// A comment attached to a Pages table cell.
pub type PagesTableCellComment = crate::comments::IWorkComment;
/// Address and storage identity of a Pages table-cell comment.
pub type PagesTableCellCommentInfo = crate::comments::IWorkTableCellCommentInfo;
/// A resolved direct reply in a Pages table-cell comment thread.
pub type PagesTableCellCommentReplyInfo = crate::comments::IWorkTableCellCommentReplyInfo;

impl PagesEditor {
    /// Read the comment attached to a reachable body-table cell.
    pub fn table_cell_comment(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Option<PagesTableCellCommentInfo>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_comment_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Create or replace a body-table cell comment transactionally.
    pub fn set_table_cell_comment(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
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
            .table_cell_comment(model_object_id, row, column)?
            .ok_or_else(|| {
                Error::InvalidFormat("Pages table comment was not persisted".to_owned())
            })?;
        if stored.comment.text != text {
            return Err(Error::InvalidFormat(
                "Pages table comment failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Delete a body-table cell comment and its reply thread transactionally.
    pub fn clear_table_cell_comment(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::clear_table_cell_comment_in_package(
            &mut staged,
            model_object_id,
            row,
            column,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .table_cell_comment(model_object_id, row, column)?
            .is_some()
        {
            return Err(Error::InvalidFormat(
                "Pages table comment deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read direct replies to a body-table cell comment in stored order.
    pub fn table_cell_comment_replies(
        &self,
        model_object_id: u64,
        row: usize,
        column: usize,
    ) -> Result<Vec<PagesTableCellCommentReplyInfo>> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_cell_comment_replies_in_package(
            self.package(),
            model_object_id,
            row,
            column,
        )
    }

    /// Append a direct reply to an existing body-table cell comment.
    pub fn add_table_cell_comment_reply(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_body_table(model_object_id)?;
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
        require_reply(&verified, model_object_id, row, column, reply_id, &text)?;
        *self = verified;
        Ok(reply_id)
    }

    /// Replace one direct reply and return its new copy-on-write object ID.
    pub fn set_table_cell_comment_reply(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_body_table(model_object_id)?;
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
        require_reply(&verified, model_object_id, row, column, new_reply_id, &text)?;
        *self = verified;
        Ok(new_reply_id)
    }

    /// Remove one direct reply from a body-table cell comment.
    pub fn remove_table_cell_comment_reply(
        &mut self,
        model_object_id: u64,
        row: usize,
        column: usize,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
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
            .table_cell_comment_replies(model_object_id, row, column)?
            .iter()
            .any(|reply| reply.storage_object_id == reply_storage_object_id)
        {
            return Err(Error::InvalidFormat(
                "Pages table comment reply deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn require_reply(
    editor: &PagesEditor,
    model_object_id: u64,
    row: usize,
    column: usize,
    reply_id: u64,
    text: &str,
) -> Result<()> {
    let valid = editor
        .table_cell_comment_replies(model_object_id, row, column)?
        .iter()
        .any(|reply| reply.storage_object_id == reply_id && reply.comment.text == text);
    if valid {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "Pages table comment reply failed validation".to_owned(),
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesDocumentBuilder;

    #[test]
    fn source_built_table_comment_threads_support_full_crud() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Comment validation\n")
            .body_table("Review", 2, 2)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].model_object_id;

        editor
            .set_table_cell_comment(table_id, 1, 1, "Initial review")
            .unwrap();
        editor
            .set_table_cell_comment(table_id, 1, 1, "Updated review")
            .unwrap();
        let root = editor.table_cell_comment(table_id, 1, 1).unwrap().unwrap();
        assert_eq!(root.comment.text, "Updated review");
        assert!(root.comment.creation_date_seconds.is_some());
        assert!(root.comment.storage_uuid.is_some());
        assert_eq!(
            editor
                .table(table_id)
                .unwrap()
                .get_comment(1, 1)
                .unwrap()
                .text,
            "Updated review"
        );

        let reply_id = editor
            .add_table_cell_comment_reply(table_id, 1, 1, "First reply")
            .unwrap();
        let updated_reply_id = editor
            .set_table_cell_comment_reply(table_id, 1, 1, reply_id, "Revised reply")
            .unwrap();
        assert_ne!(updated_reply_id, reply_id);
        assert_eq!(
            editor.table_cell_comment_replies(table_id, 1, 1).unwrap()[0]
                .comment
                .text,
            "Revised reply"
        );

        let before_invalid = editor.to_bytes().unwrap();
        assert!(
            editor
                .remove_table_cell_comment_reply(table_id, 1, 1, u64::MAX)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_invalid);

        editor
            .remove_table_cell_comment_reply(table_id, 1, 1, updated_reply_id)
            .unwrap();
        assert!(
            editor
                .table_cell_comment_replies(table_id, 1, 1)
                .unwrap()
                .is_empty()
        );
        editor.clear_table_cell_comment(table_id, 1, 1).unwrap();
        assert!(editor.table_cell_comment(table_id, 1, 1).unwrap().is_none());

        let reparsed = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reparsed
                .table_cell_comment(table_id, 1, 1)
                .unwrap()
                .is_none()
        );
    }
}
