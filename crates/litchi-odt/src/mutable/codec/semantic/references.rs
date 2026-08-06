//! Index, reference-mark, and bookmark semantic snapshot edits.

use super::super::super::model::MutableDocument;
use crate::{BookmarkTarget, ReferenceMark, TextIndex, TextIndexMark};
use litchi_core::Result;

impl MutableDocument {
    /// Return generated indexes from the current authoritative content XML.
    pub fn text_indexes(&self) -> Result<Vec<TextIndex>> {
        self.with_content_xml(crate::index::parse_text_indexes)
    }

    /// Append caller-authored index markup to `office:text` without refreshing its cache.
    pub fn insert_text_index(&mut self, index: &TextIndex) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::insert_text_index_xml(xml, index))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace one named index and return its previous inert representation.
    pub fn replace_text_index(&mut self, name: &str, replacement: &TextIndex) -> Result<TextIndex> {
        let old = self
            .text_indexes()?
            .into_iter()
            .find(|index| index.name() == name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!("text index {name:?} was not found"))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::replace_text_index_xml(xml, name, replacement))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove one named index and return its previous inert representation.
    pub fn remove_text_index(&mut self, name: &str) -> Result<TextIndex> {
        let old = self
            .text_indexes()?
            .into_iter()
            .find(|index| index.name() == name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!("text index {name:?} was not found"))
            })?;
        let updated = self.with_content_xml(|xml| crate::remove_text_index_xml(xml, name))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return typed point and resolved range index marks in document order.
    pub fn text_index_marks(&self) -> Result<Vec<TextIndexMark>> {
        self.with_content_xml(crate::index_mark::parse_text_index_marks)
    }

    /// Insert a point mark at a paragraph end, or wrap the paragraph with a range mark.
    pub fn insert_text_index_mark(
        &mut self,
        paragraph_index: usize,
        mark: &TextIndexMark,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_text_index_mark_xml(xml, paragraph_index, mark)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    pub fn replace_text_index_mark(
        &mut self,
        mark_index: usize,
        replacement: &TextIndexMark,
    ) -> Result<TextIndexMark> {
        let old = self
            .text_index_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "index mark {mark_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_text_index_mark_xml(xml, mark_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    pub fn remove_text_index_mark(&mut self, mark_index: usize) -> Result<TextIndexMark> {
        let old = self
            .text_index_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "index mark {mark_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_text_index_mark_xml(xml, mark_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return point and resolved range reference targets in document order.
    pub fn reference_marks(&self) -> Result<Vec<ReferenceMark>> {
        self.with_content_xml(crate::reference_mark::parse_reference_marks)
    }

    /// Insert a point reference at a paragraph end, or wrap the paragraph with a range reference.
    pub fn insert_reference_mark(
        &mut self,
        paragraph_index: usize,
        mark: &ReferenceMark,
    ) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| crate::insert_reference_mark_xml(xml, paragraph_index, mark))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a reference target selected in document order.
    pub fn replace_reference_mark(
        &mut self,
        mark_index: usize,
        replacement: &ReferenceMark,
    ) -> Result<ReferenceMark> {
        let old = self
            .reference_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "reference mark {mark_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_reference_mark_xml(xml, mark_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove marker elements while preserving enclosed text and markup.
    pub fn remove_reference_mark(&mut self, mark_index: usize) -> Result<ReferenceMark> {
        let old = self
            .reference_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "reference mark {mark_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_reference_mark_xml(xml, mark_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return point and resolved range bookmarks in document order.
    pub fn bookmark_targets(&self) -> Result<Vec<BookmarkTarget>> {
        self.with_content_xml(crate::parse_bookmark_targets)
    }

    /// Insert a point bookmark at a paragraph end, or wrap the paragraph with a range bookmark.
    pub fn insert_bookmark_target(
        &mut self,
        paragraph_index: usize,
        target: &BookmarkTarget,
    ) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::insert_bookmark_xml(xml, paragraph_index, target))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a bookmark target selected in document order.
    pub fn replace_bookmark_target(
        &mut self,
        target_index: usize,
        replacement: &BookmarkTarget,
    ) -> Result<BookmarkTarget> {
        let old = self
            .bookmark_targets()?
            .get(target_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "bookmark target {target_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::replace_bookmark_xml(xml, target_index, replacement))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove bookmark markers while preserving enclosed text and markup.
    pub fn remove_bookmark_target(&mut self, target_index: usize) -> Result<BookmarkTarget> {
        let old = self
            .bookmark_targets()?
            .get(target_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "bookmark target {target_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| crate::remove_bookmark_xml(xml, target_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }
}
