//! Sheet-owned text-box editing semantics.

#![allow(unused_imports)]

use super::*;
use crate::text::layout::Layout;
use litchi_iwa_text::columns::Columns;

impl NumbersEditor {
    /// List ordinary text boxes owned by a reachable Numbers sheet.
    pub fn sheet_text_boxes(&self, sheet_id: u64) -> Result<Vec<NumbersTextBoxInfo>> {
        let mut catalog = NumbersObjectCatalog::build(&self.package)?;
        let drawable_ids = catalog
            .sheet_drawable_ids(&self.package, sheet_id)?
            .to_vec();
        let mut result = Vec::new();
        for drawable_id in drawable_ids {
            let Some(graph) =
                catalog.text_box_graph_if_supported(&self.package, sheet_id, drawable_id)?
            else {
                continue;
            };
            let storage = catalog.text_storage_info(&self.package, graph.storage_id)?;
            result.push(NumbersTextBoxInfo {
                sheet_id,
                drawable_object_id: drawable_id,
                storage,
            });
        }
        Ok(result)
    }

    /// Replace a UTF-16 range in an ordinary Numbers text box.
    pub fn replace_sheet_text_box_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.replace_text(graph.storage_id, range, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        numbers_text_box_graph(verified.package(), sheet_id, drawable_object_id)?;
        self.package = verified.package;
        Ok(())
    }

    /// Replace all text in an ordinary Numbers text box.
    pub fn set_sheet_text_box_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text(graph.storage_id, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        let updated = verified
            .sheet_text_boxes(sheet_id)?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers text-box update lost its drawable".to_owned())
            })?;
        if updated.storage.storage.text() != replacement {
            return Err(Error::InvalidFormat(
                "Numbers text-box update failed validation".to_owned(),
            ));
        }
        self.package = verified.package;
        Ok(())
    }

    /// Clear an ordinary Numbers text box without deleting it.
    pub fn clear_sheet_text_box_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<()> {
        self.set_sheet_text_box_text(sheet_id, drawable_object_id, "")
    }

    /// Read the geometry of an ordinary Numbers text box.
    pub fn sheet_text_box_geometry(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_geometry(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Update position, size, flags, and rotation on an ordinary Numbers text box.
    pub fn set_sheet_text_box_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_geometry(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_text_box_geometry(sheet_id, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Numbers text-box geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties from an ordinary Numbers text box.
    pub fn sheet_text_box_properties(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_properties(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Update shared drawable properties on an ordinary Numbers text box.
    pub fn set_sheet_text_box_properties(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_properties(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_text_box_properties(sheet_id, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Numbers text-box properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read vertical alignment, edge insets, and autosizing for a text box.
    pub fn sheet_text_box_text_layout(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Layout> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_text_layout(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Replace text-frame layout while preserving text, columns, and drawing style.
    pub fn set_sheet_text_box_text_layout(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        layout: Layout,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let staged = set_shape_text_layout(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
            layout,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_text_box_text_layout(sheet_id, drawable_object_id)? != layout {
            return Err(Error::InvalidFormat(
                "Numbers text-box layout update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove crate-authored text-frame layout overrides.
    pub fn reset_sheet_text_box_text_layout(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_layout(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the uniform column layout of an ordinary sheet text box.
    pub fn sheet_text_box_columns(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Columns> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        shape_text_columns(&self.package, &graph.archive_name, drawable_object_id)
    }

    /// Replace the uniform column layout of an ordinary sheet text box.
    pub fn set_sheet_text_box_columns(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        columns: &Columns,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let staged = set_shape_text_columns(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
            columns,
        )?;
        let verified = Self::from_package(staged)?;
        if &verified.sheet_text_box_columns(sheet_id, drawable_object_id)? != columns {
            return Err(Error::InvalidFormat(
                "Numbers text-box column update failed validation".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited column layout after a crate-authored override.
    pub fn reset_sheet_text_box_columns(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_columns(
            self.package.clone(),
            &graph.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective uniform font size, bold, and italic formatting.
    pub fn sheet_text_box_text_style(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextStyle> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_style(graph.storage_id)
    }

    /// Atomically set uniform font size, bold, and italic formatting.
    pub fn set_sheet_text_box_text_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        style: TextStyle,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_style(graph.storage_id, style)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_style(sheet_id, drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Numbers text-box character formatting update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited character formatting while preserving paragraph overrides.
    pub fn reset_sheet_text_box_text_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_style(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective PostScript font identity of a sheet-owned text box.
    pub fn sheet_text_box_text_font(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextFont> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_font(graph.storage_id)
    }

    /// Atomically set a typed font identity across a sheet-owned text box.
    pub fn set_sheet_text_box_text_font(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        font: TextFont,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_font(graph.storage_id, font)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore the inherited font while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_font(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_font(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read every explicit language boundary in a sheet-owned text box.
    pub fn sheet_text_box_text_languages(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextLanguageRun>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_languages(graph.storage_id)
    }

    /// Read the effective language at one UTF-16 text boundary.
    pub fn sheet_text_box_text_language(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<TextLanguage> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .text_language(graph.storage_id, position)
    }

    /// Atomically create or update one text-language boundary.
    pub fn set_sheet_text_box_text_language(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: TextPosition,
        language: TextLanguage,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_language(graph.storage_id, position, language)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Delete one nonzero language boundary so it inherits the preceding run.
    pub fn remove_sheet_text_box_text_language_boundary(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        position: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.remove_text_language_boundary(graph.storage_id, position)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Restore automatic language selection across a sheet-owned text box.
    pub fn reset_sheet_text_box_text_languages(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_languages(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read every hyperlink in a sheet-owned text box.
    pub fn sheet_text_box_hyperlinks(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextHyperlink>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_hyperlinks(graph.storage_id)
    }

    /// Create a hyperlink over a nonempty, unoccupied UTF-16 text range.
    pub fn add_sheet_text_box_hyperlink(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let hyperlink = text.add_text_hyperlink(graph.storage_id, range, target)?;
        *self = Self::from_package(text.into_package())?;
        Ok(hyperlink)
    }

    /// Update a text-box hyperlink's range and target without changing its ID.
    pub fn update_sheet_text_box_hyperlink(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHyperlinkId,
        range: TextRange,
        target: TextHyperlinkTarget,
    ) -> Result<TextHyperlink> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let hyperlink = text.update_text_hyperlink(graph.storage_id, id, range, target)?;
        *self = Self::from_package(text.into_package())?;
        Ok(hyperlink)
    }

    /// Delete a text-box hyperlink and its owned smart-field object.
    pub fn remove_sheet_text_box_hyperlink(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHyperlinkId,
    ) -> Result<TextHyperlink> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let hyperlink = text.remove_text_hyperlink(graph.storage_id, id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(hyperlink)
    }

    /// Read every plain highlight in a sheet-owned text box.
    pub fn sheet_text_box_highlights(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextHighlight>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_highlights(graph.storage_id)
    }

    /// Create a plain highlight over a nonempty UTF-16 text range.
    pub fn add_sheet_text_box_highlight(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let highlight = text.add_text_highlight(graph.storage_id, range)?;
        *self = Self::from_package(text.into_package())?;
        Ok(highlight)
    }

    /// Move a plain text-box highlight without changing its ID.
    pub fn update_sheet_text_box_highlight(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHighlightId,
        range: TextRange,
    ) -> Result<TextHighlight> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let highlight = text.update_text_highlight(graph.storage_id, id, range)?;
        *self = Self::from_package(text.into_package())?;
        Ok(highlight)
    }

    /// Delete a plain text-box highlight and its empty annotation graph.
    pub fn remove_sheet_text_box_highlight(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextHighlightId,
    ) -> Result<TextHighlight> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let highlight = text.remove_text_highlight(graph.storage_id, id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(highlight)
    }

    /// Read every ranged comment in a sheet-owned text box.
    pub fn sheet_text_box_comments(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<TextComment>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_comments(graph.storage_id)
    }

    /// Create a ranged comment in a sheet-owned text box.
    pub fn add_sheet_text_box_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let comment = text.add_text_comment(graph.storage_id, range, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(comment)
    }

    /// Update a text-box comment's range and body without changing its ID.
    pub fn update_sheet_text_box_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextCommentId,
        range: TextRange,
        body: TextCommentBody,
    ) -> Result<TextComment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let comment = text.update_text_comment(graph.storage_id, id, range, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(comment)
    }

    /// Delete a ranged text-box comment and its owned annotation graph.
    pub fn remove_sheet_text_box_comment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        id: TextCommentId,
    ) -> Result<TextComment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let comment = text.remove_text_comment(graph.storage_id, id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(comment)
    }

    /// Read every direct reply to a sheet text-box comment in stored order.
    pub fn sheet_text_box_comment_replies(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
    ) -> Result<Vec<TextCommentReply>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .text_comment_replies(graph.storage_id, comment_id)
    }

    /// Append a direct reply to a sheet text-box comment.
    pub fn add_sheet_text_box_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let reply = text.add_text_comment_reply(graph.storage_id, comment_id, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(reply)
    }

    /// Update a direct sheet text-box comment reply without changing its ID.
    pub fn update_sheet_text_box_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
        body: TextCommentReplyBody,
    ) -> Result<TextCommentReply> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let reply = text.update_text_comment_reply(graph.storage_id, comment_id, reply_id, body)?;
        *self = Self::from_package(text.into_package())?;
        Ok(reply)
    }

    /// Delete one direct sheet text-box comment reply and its storage.
    pub fn remove_sheet_text_box_comment_reply(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        comment_id: TextCommentId,
        reply_id: TextCommentReplyId,
    ) -> Result<TextCommentReply> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let reply = text.remove_text_comment_reply(graph.storage_id, comment_id, reply_id)?;
        *self = Self::from_package(text.into_package())?;
        Ok(reply)
    }

    /// Read the canonical list preset of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_list(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphList> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_list(graph.storage_id)
    }

    /// Atomically apply a canonical list preset to a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_list(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        list: ParagraphList,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list(graph.storage_id, list)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Remove list formatting from a sheet-owned text box.
    pub fn reset_sheet_text_box_paragraph_list(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read every list-level boundary in a sheet-owned text box.
    pub fn sheet_text_box_paragraph_list_levels(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ParagraphListLevelPlacement>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_list_levels(graph.storage_id)
    }

    /// Read one paragraph's effective list nesting level.
    pub fn sheet_text_box_paragraph_list_level(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLevel> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_level(graph.storage_id, paragraph)
    }

    /// Atomically set one paragraph's list nesting level.
    pub fn set_sheet_text_box_paragraph_list_level(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_level(graph.storage_id, paragraph, level)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore one paragraph to the top-level list nesting level.
    pub fn reset_sheet_text_box_paragraph_list_level(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_level(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read whether one text-box paragraph continues or restarts list numbering.
    pub fn sheet_text_box_paragraph_list_numbering(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumbering> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_numbering(graph.storage_id, paragraph)
    }

    /// Continue or restart numbered-list sequencing at one text-box paragraph.
    pub fn set_sheet_text_box_paragraph_list_numbering(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_numbering(graph.storage_id, paragraph, numbering)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Read one numbered sheet text-box paragraph's effective label format.
    pub fn sheet_text_box_paragraph_list_number_format(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberFormat> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_number_format(graph.storage_id, paragraph)
    }

    /// Set one numbered sheet text-box paragraph's locale-aware label format.
    pub fn set_sheet_text_box_paragraph_list_number_format(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        format: ParagraphListNumberFormat,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_number_format(graph.storage_id, paragraph, format)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore the standard decimal-period label format.
    pub fn reset_sheet_text_box_paragraph_list_number_format(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_number_format(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read whether one numbered sheet text-box paragraph displays hierarchical numbering.
    pub fn sheet_text_box_paragraph_list_number_tiering(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberTiering> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_number_tiering(graph.storage_id, paragraph)
    }

    /// Choose flat or hierarchical numbering for one sheet text-box list level.
    pub fn set_sheet_text_box_paragraph_list_number_tiering(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        tiering: ParagraphListNumberTiering,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_number_tiering(graph.storage_id, paragraph, tiering)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore flat numbering for one sheet text-box list level.
    pub fn reset_sheet_text_box_paragraph_list_number_tiering(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_number_tiering(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read one numbered sheet text-box paragraph's number-label size.
    pub fn sheet_text_box_paragraph_list_number_scale(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumberScale> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_number_scale(graph.storage_id, paragraph)
    }

    /// Set one numbered sheet text-box paragraph's number-label size.
    pub fn set_sheet_text_box_paragraph_list_number_scale(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        scale: ParagraphListNumberScale,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_number_scale(graph.storage_id, paragraph, scale)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore the standard 100% number-label size.
    pub fn reset_sheet_text_box_paragraph_list_number_scale(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_number_scale(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read one sheet text-box paragraph's effective text-bullet marker.
    pub fn sheet_text_box_paragraph_list_bullet(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListBullet> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_bullet(graph.storage_id, paragraph)
    }

    /// Set one sheet text-box paragraph's text-bullet marker.
    pub fn set_sheet_text_box_paragraph_list_bullet(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        bullet: &ParagraphListBullet,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_bullet(graph.storage_id, paragraph, bullet)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard `•` marker for one sheet text-box paragraph.
    pub fn reset_sheet_text_box_paragraph_list_bullet(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_bullet(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read one sheet text-box paragraph's effective bullet size and baseline.
    pub fn sheet_text_box_paragraph_list_bullet_geometry(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListBulletGeometry> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_bullet_geometry(graph.storage_id, paragraph)
    }

    /// Set one sheet text-box paragraph's bullet size and baseline.
    pub fn set_sheet_text_box_paragraph_list_bullet_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        geometry: ParagraphListBulletGeometry,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_bullet_geometry(graph.storage_id, paragraph, geometry)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard bullet size and baseline for this nesting level.
    pub fn reset_sheet_text_box_paragraph_list_bullet_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_bullet_geometry(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read one sheet text-box list paragraph's label and text-gap indentation.
    pub fn sheet_text_box_paragraph_list_indentation(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListIndentation> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_indentation(graph.storage_id, paragraph)
    }

    /// Set one sheet text-box list paragraph's label and text-gap indentation.
    pub fn set_sheet_text_box_paragraph_list_indentation(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        indentation: ParagraphListIndentation,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_indentation(graph.storage_id, paragraph, indentation)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore Apple's standard indentation for this list preset and level.
    pub fn reset_sheet_text_box_paragraph_list_indentation(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_indentation(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read one sheet text-box list paragraph's effective label color.
    pub fn sheet_text_box_paragraph_list_label_color(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLabelColor> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_list_label_color(graph.storage_id, paragraph)
    }

    /// Set one sheet text-box list paragraph's bullet or number color.
    pub fn set_sheet_text_box_paragraph_list_label_color(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
        color: ParagraphListLabelColor,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_list_label_color(graph.storage_id, paragraph, color)?;
        *self = Self::from_package(text.into_package())?;
        Ok(())
    }

    /// Restore the list label to the paragraph's automatic text color.
    pub fn reset_sheet_text_box_paragraph_list_label_color(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph: TextPosition,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_list_label_color(graph.storage_id, paragraph)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform underline and strikethrough formatting.
    pub fn sheet_text_box_text_decorations(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextDecorations> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_decorations(graph.storage_id)
    }

    /// Atomically set uniform underline and strikethrough formatting.
    pub fn set_sheet_text_box_text_decorations(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        decorations: TextDecorations,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_decorations(graph.storage_id, decorations)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_decorations(sheet_id, drawable_object_id)? != decorations {
            return Err(Error::InvalidFormat(
                "Numbers text-box decoration update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited decorations while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_decorations(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_decorations(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective uniform text color of a sheet-owned text box.
    pub fn sheet_text_box_text_color(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RgbaColor> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_color(graph.storage_id)
    }

    /// Atomically set one text color across a sheet-owned text box.
    pub fn set_sheet_text_box_text_color(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        color: RgbaColor,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_color(graph.storage_id, color)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_color(sheet_id, drawable_object_id)? != color {
            return Err(Error::InvalidFormat(
                "Numbers text-box color update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited text color while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_color(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_color(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform capitalization from a sheet-owned text box.
    pub fn sheet_text_box_text_capitalization(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextCapitalization> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_capitalization(graph.storage_id)
    }

    /// Atomically set one capitalization mode across a sheet-owned text box.
    pub fn set_sheet_text_box_text_capitalization(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        capitalization: TextCapitalization,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_capitalization(graph.storage_id, capitalization)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_capitalization(sheet_id, drawable_object_id)?
            != capitalization
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box capitalization update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited capitalization while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_capitalization(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_capitalization(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective uniform baseline script from a sheet-owned text box.
    pub fn sheet_text_box_text_script(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextScript> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_script(graph.storage_id)
    }

    /// Atomically set normal, superscript, or subscript formatting.
    pub fn set_sheet_text_box_text_script(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        script: TextScript,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_script(graph.storage_id, script)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_script(sheet_id, drawable_object_id)? != script {
            return Err(Error::InvalidFormat(
                "Numbers text-box script update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited baseline script while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_script(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_script(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective custom baseline displacement of a sheet-owned text box.
    pub fn sheet_text_box_text_baseline_shift(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextBaselineShift> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_baseline_shift(graph.storage_id)
    }

    /// Atomically set a signed custom baseline displacement.
    pub fn set_sheet_text_box_text_baseline_shift(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shift: TextBaselineShift,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_baseline_shift(graph.storage_id, shift)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_baseline_shift(sheet_id, drawable_object_id)? != shift {
            return Err(Error::InvalidFormat(
                "Numbers text-box baseline-shift update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited baseline displacement while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_baseline_shift(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_baseline_shift(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective character spacing of a sheet-owned text box.
    pub fn sheet_text_box_text_character_spacing(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextCharacterSpacing> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_character_spacing(graph.storage_id)
    }

    /// Atomically set character spacing across a sheet-owned text box.
    pub fn set_sheet_text_box_text_character_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        spacing: TextCharacterSpacing,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_character_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_character_spacing(sheet_id, drawable_object_id)? != spacing
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box character-spacing update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited character spacing while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_character_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_character_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ligature policy of a sheet-owned text box.
    pub fn sheet_text_box_text_ligatures(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextLigatures> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_ligatures(graph.storage_id)
    }

    /// Atomically set the ligature policy across a sheet-owned text box.
    pub fn set_sheet_text_box_text_ligatures(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        ligatures: TextLigatures,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_ligatures(graph.storage_id, ligatures)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_ligatures(sheet_id, drawable_object_id)? != ligatures {
            return Err(Error::InvalidFormat(
                "Numbers text-box ligature update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited ligatures while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_ligatures(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_ligatures(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective outline of a sheet-owned text box.
    pub fn sheet_text_box_text_outline(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextOutline> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_outline(graph.storage_id)
    }

    /// Atomically set a typed outline across a sheet-owned text box.
    pub fn set_sheet_text_box_text_outline(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        outline: TextOutline,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_outline(graph.storage_id, outline)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_outline(sheet_id, drawable_object_id)? != outline {
            return Err(Error::InvalidFormat(
                "Numbers text-box outline update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited outline while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_outline(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_outline(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective shadow of a sheet-owned text box.
    pub fn sheet_text_box_text_shadow(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextShadow> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_shadow(graph.storage_id)
    }

    /// Atomically set a typed drop shadow across a sheet-owned text box.
    pub fn set_sheet_text_box_text_shadow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shadow: TextShadow,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_shadow(graph.storage_id, shadow)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_shadow(sheet_id, drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Numbers text-box shadow update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited shadow while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_shadow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_shadow(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective solid background of a sheet-owned text box.
    pub fn sheet_text_box_text_background(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextBackground> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).text_background(graph.storage_id)
    }

    /// Atomically set a solid background across a sheet-owned text box.
    pub fn set_sheet_text_box_text_background(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        background: TextBackground,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text_background(graph.storage_id, background)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_text_background(sheet_id, drawable_object_id)? != background {
            return Err(Error::InvalidFormat(
                "Numbers text-box background update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited text background while preserving sibling overrides.
    pub fn reset_sheet_text_box_text_background(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_text_background(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective Text → Layout paragraph background.
    pub fn sheet_text_box_paragraph_background(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphBackground> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_background(graph.storage_id)
    }

    /// Atomically set the paragraph background across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_background(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        background: ParagraphBackground,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_background(graph.storage_id, background)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_background(sheet_id, drawable_object_id)? != background
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph background update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited paragraph background.
    pub fn reset_sheet_text_box_paragraph_background(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_background(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective Text → Layout paragraph borders.
    pub fn sheet_text_box_paragraph_borders(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphBorders> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_borders(graph.storage_id)
    }

    /// Atomically set paragraph borders across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_borders(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        borders: ParagraphBorders,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_borders(graph.storage_id, borders)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_borders(sheet_id, drawable_object_id)? != borders {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph border update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited paragraph borders.
    pub fn reset_sheet_text_box_paragraph_borders(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_borders(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph pagination and hyphenation controls.
    pub fn sheet_text_box_paragraph_flow(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphFlow> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_flow(graph.storage_id)
    }

    /// Atomically set paragraph pagination and hyphenation controls.
    pub fn set_sheet_text_box_paragraph_flow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        flow: ParagraphFlow,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_flow(graph.storage_id, flow)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_flow(sheet_id, drawable_object_id)? != flow {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph flow update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited paragraph pagination and hyphenation controls.
    pub fn reset_sheet_text_box_paragraph_flow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_flow(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective base-writing direction of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_writing_direction(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphWritingDirection> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_writing_direction(graph.storage_id)
    }

    /// Set one base-writing direction across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_writing_direction(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        direction: ParagraphWritingDirection,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_writing_direction(graph.storage_id, direction)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_writing_direction(sheet_id, drawable_object_id)?
            != direction
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph writing-direction update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited base-writing direction.
    pub fn reset_sheet_text_box_paragraph_writing_direction(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_writing_direction(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective paragraph alignment of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_alignment(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<TextAlignment> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_alignment(graph.storage_id)
    }

    /// Set one paragraph alignment across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_alignment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        alignment: TextAlignment,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_alignment(graph.storage_id, alignment)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_alignment(sheet_id, drawable_object_id)? != alignment {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph-alignment update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited paragraph alignment after a private minimal override.
    pub fn reset_sheet_text_box_paragraph_alignment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_alignment(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective line spacing of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_line_spacing(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphLineSpacing> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_line_spacing(graph.storage_id)
    }

    /// Set one typed line-spacing mode across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_line_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        spacing: ParagraphLineSpacing,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_line_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_line_spacing(sheet_id, drawable_object_id)? != spacing
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph line-spacing update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited line spacing while preserving sibling paragraph overrides.
    pub fn reset_sheet_text_box_paragraph_line_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_line_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective before/after paragraph spacing of a sheet-owned text box.
    pub fn sheet_text_box_paragraph_spacing(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphSpacing> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_spacing(graph.storage_id)
    }

    /// Atomically set before/after paragraph spacing across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        spacing: ParagraphSpacing,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_spacing(graph.storage_id, spacing)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_spacing(sheet_id, drawable_object_id)? != spacing {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph spacing update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited paragraph spacing while preserving sibling overrides.
    pub fn reset_sheet_text_box_paragraph_spacing(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_spacing(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read effective first-line, left, and right indentation of a sheet text box.
    pub fn sheet_text_box_paragraph_indents(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphIndents> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_indents(graph.storage_id)
    }

    /// Atomically set paragraph indentation across a sheet-owned text box.
    pub fn set_sheet_text_box_paragraph_indents(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        indents: ParagraphIndents,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_indents(graph.storage_id, indents)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_indents(sheet_id, drawable_object_id)? != indents {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph indentation update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited indentation while preserving sibling paragraph overrides.
    pub fn reset_sheet_text_box_paragraph_indents(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_indents(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the decimal-tab alignment character of a sheet text box.
    pub fn sheet_text_box_paragraph_decimal_tab_character(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphDecimalTabCharacter> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_decimal_tab_character(graph.storage_id)
    }

    /// Atomically set the decimal-tab alignment character of a sheet text box.
    pub fn set_sheet_text_box_paragraph_decimal_tab_character(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        character: ParagraphDecimalTabCharacter,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_decimal_tab_character(graph.storage_id, character)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_decimal_tab_character(sheet_id, drawable_object_id)?
            != character
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box decimal-tab character update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited decimal-tab alignment character.
    pub fn reset_sheet_text_box_paragraph_decimal_tab_character(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_decimal_tab_character(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the distance between implicit tab stops in a sheet text box.
    pub fn sheet_text_box_paragraph_default_tab_interval(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphDefaultTabInterval> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_default_tab_interval(graph.storage_id)
    }

    /// Atomically set the distance between implicit tab stops in a sheet text box.
    pub fn set_sheet_text_box_paragraph_default_tab_interval(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        interval: ParagraphDefaultTabInterval,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_default_tab_interval(graph.storage_id, interval)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_default_tab_interval(sheet_id, drawable_object_id)?
            != interval
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box default-tab interval update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore the inherited default-tab interval.
    pub fn reset_sheet_text_box_paragraph_default_tab_interval(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_default_tab_interval(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Read the effective ordered ruler tab stops of a sheet text box.
    pub fn sheet_text_box_paragraph_tab_stops(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ParagraphTabStops> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_tab_stops(graph.storage_id)
    }

    /// Atomically replace every explicit ruler tab stop of a sheet text box.
    pub fn set_sheet_text_box_paragraph_tab_stops(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        stops: ParagraphTabStops,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_tab_stops(graph.storage_id, stops)?;
        let expected = text.paragraph_tab_stops(graph.storage_id)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_tab_stops(sheet_id, drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Numbers text-box paragraph tab-stop update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Restore inherited tab stops while preserving sibling paragraph overrides.
    pub fn reset_sheet_text_box_paragraph_tab_stops(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.reset_paragraph_tab_stops(graph.storage_id)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// List every Drop Cap in a sheet-owned text box.
    pub fn sheet_text_box_paragraph_drop_caps(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Vec<ParagraphDropCapPlacement>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone()).paragraph_drop_caps(graph.storage_id)
    }

    /// Read the Drop Cap attached to one text-box paragraph.
    pub fn sheet_text_box_paragraph_drop_cap(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<Option<ParagraphDropCap>> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        IWorkTextEditor::from_package(self.package.clone())
            .paragraph_drop_cap(graph.storage_id, paragraph_start)
    }

    /// Atomically create or replace a text-box Drop Cap.
    pub fn set_sheet_text_box_paragraph_drop_cap(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
        drop_cap: ParagraphDropCap,
    ) -> Result<()> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_paragraph_drop_cap(graph.storage_id, paragraph_start, drop_cap)?;
        let verified = Self::from_package(text.into_package())?;
        if verified.sheet_text_box_paragraph_drop_cap(
            sheet_id,
            drawable_object_id,
            paragraph_start,
        )? != Some(drop_cap)
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box Drop Cap update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Atomically remove a text-box Drop Cap.
    pub fn remove_sheet_text_box_paragraph_drop_cap(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        paragraph_start: ParagraphStart,
    ) -> Result<bool> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        let changed = text.remove_paragraph_drop_cap(graph.storage_id, paragraph_start)?;
        if changed {
            *self = Self::from_package(text.into_package())?;
        }
        Ok(changed)
    }

    /// Remove an ordinary sheet-owned text box and its private object graph.
    pub fn remove_sheet_text_box(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersTextBox> {
        let graph = numbers_text_box_graph(&self.package, sheet_id, drawable_object_id)?;
        let text_box = self
            .sheet_text_boxes(sheet_id)?
            .into_iter()
            .find(|item| item.drawable_object_id == drawable_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers text box {drawable_object_id} lost its writable storage"
                ))
            })?;

        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(DrawableId::from_raw(drawable_object_id)?)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &graph.archive_name,
            graph.sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        staged.update_archive(&graph.archive_name, |archive| {
            for identifier in &graph.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers text-box object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &graph.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers text-box object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(
            &mut staged,
            DOCUMENT_COMPONENT_IDENTIFIER,
            &graph.uuid_object_ids,
        )?;
        release_package_identifier_suffix(&mut staged, &graph.object_ids)?;

        let verified = Self::from_package(staged)?;
        if verified
            .sheet_text_boxes(sheet_id)?
            .iter()
            .any(|item| item.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers text-box deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersTextBox { text_box })
    }
}
