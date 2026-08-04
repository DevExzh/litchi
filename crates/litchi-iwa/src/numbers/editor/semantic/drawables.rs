//! Sheet-owned drawable comment semantics.

#![allow(unused_imports)]

use super::*;

impl NumbersEditor {
    /// List supported direct-comment drawables owned by one reachable sheet.
    pub fn sheet_drawables(&self, sheet_id: u64) -> Result<Vec<IWorkDrawableInfo>> {
        let owned = self.sheet_owned_drawable_ids(sheet_id)?;
        let mut drawables = IWorkDrawableCommentEditor::from_package(self.package.clone())?
            .drawables()?
            .into_iter()
            .filter(|drawable| owned.contains(&drawable.object_id.object_id()))
            .collect::<Vec<_>>();
        drawables.sort_by_key(|drawable| drawable.object_id.object_id());
        Ok(drawables)
    }

    /// Read a comment attached directly to a drawable owned by one sheet.
    pub fn sheet_drawable_comment(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<DrawableCommentInfo>> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package.clone())?.comment(drawable_object_id)
    }

    /// Create or replace a direct comment on a drawable owned by one sheet.
    pub fn set_sheet_drawable_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.set_comment(drawable_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Delete a direct comment from a drawable owned by one sheet.
    pub fn clear_sheet_drawable_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<()> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(drawable_object_id)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }

    /// Read direct replies in a comment thread on one sheet drawable.
    pub fn sheet_drawable_comment_replies(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<DrawableCommentReplyInfo>> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        IWorkDrawableCommentEditor::from_package(self.package.clone())?.replies(drawable_object_id)
    }

    /// Add a reply to a direct comment on one sheet drawable.
    pub fn add_sheet_drawable_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        let reply_id = comments.add_reply(drawable_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id.object_id())
    }

    /// Update a direct reply, returning its current storage identifier.
    pub fn set_sheet_drawable_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
        text: impl Into<String>,
    ) -> Result<u64> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        let reply_id = comments.set_reply(drawable_object_id, reply_storage_object_id, text)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(reply_id.object_id())
    }

    /// Remove a direct reply from a comment on one sheet drawable.
    pub fn remove_sheet_drawable_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        reply_storage_object_id: u64,
    ) -> Result<()> {
        self.require_sheet_drawable(sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.remove_reply(drawable_object_id, reply_storage_object_id)?;
        *self = Self::from_package(comments.into_package())?;
        Ok(())
    }
}
