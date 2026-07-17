//! Section-scoped body text reading and mutation.

use std::ops::Range;

use super::PagesEditor;
use crate::text::{
    TextBookmark, TextBookmarkId, TextBookmarkSettings, TextDateTimeDisplayText, TextDateTimeField,
    TextDateTimeFieldId, TextDateTimeFieldSettings, TextPosition, TextRange,
};
use crate::{Error, Result};

impl PagesEditor {
    /// Read every native ranged bookmark in the main body.
    pub fn body_bookmarks(&self) -> Result<Vec<TextBookmark>> {
        self.text.text_bookmarks(self.body_storage_id)
    }

    /// Create a native body bookmark over a nonempty UTF-16 range.
    pub fn add_body_bookmark(
        &mut self,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        self.text
            .add_text_bookmark(self.body_storage_id, range, settings)
    }

    /// Atomically update a body bookmark's range and settings.
    pub fn update_body_bookmark(
        &mut self,
        id: TextBookmarkId,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        self.text
            .update_text_bookmark(self.body_storage_id, id, range, settings)
    }

    /// Delete one native body bookmark and reclaim its owned field object.
    pub fn remove_body_bookmark(&mut self, id: TextBookmarkId) -> Result<TextBookmark> {
        self.text.remove_text_bookmark(self.body_storage_id, id)
    }

    /// Read every native Date & Time field in the main body.
    pub fn body_date_time_fields(&self) -> Result<Vec<TextDateTimeField>> {
        self.text.text_date_time_fields(self.body_storage_id)
    }

    /// Attach a Date & Time field to existing body text.
    pub fn add_body_date_time_field(
        &mut self,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        self.text
            .add_text_date_time_field(self.body_storage_id, range, settings)
    }

    /// Atomically insert exact display text and its Date & Time field.
    pub fn insert_body_date_time_field(
        &mut self,
        position: TextPosition,
        display_text: TextDateTimeDisplayText,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        self.text.insert_text_date_time_field(
            self.body_storage_id,
            position,
            display_text,
            settings,
        )
    }

    /// Atomically update a body Date & Time field's range and formatter payload.
    pub fn update_body_date_time_field(
        &mut self,
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: TextDateTimeFieldSettings,
    ) -> Result<TextDateTimeField> {
        self.text
            .update_text_date_time_field(self.body_storage_id, id, range, settings)
    }

    /// Delete one body Date & Time field while retaining its visible text.
    pub fn remove_body_date_time_field(
        &mut self,
        id: TextDateTimeFieldId,
    ) -> Result<TextDateTimeField> {
        self.text
            .remove_text_date_time_field(self.body_storage_id, id)
    }

    /// Replace a UTF-16 range in the body without creating or deleting section boundaries.
    ///
    /// Ranges may edit content on either side of a boundary, but cannot consume the native
    /// U+0004 section-break marker. Use [`Self::insert_section`] or [`Self::remove_section`] to
    /// change the section graph.
    pub fn replace_body_text(&mut self, range: Range<usize>, replacement: &str) -> Result<()> {
        self.validate_body_edit(&range, replacement)?;
        let mut staged = self.text.clone();
        staged.replace_text(self.body_storage_id, range, replacement)?;
        *self = Self::from_package(staged.into_package())?;
        Ok(())
    }

    /// Replace the complete body of a single-section document.
    ///
    /// Multi-section documents must be edited with [`Self::set_section_text`] so their native
    /// section breaks cannot be discarded accidentally.
    pub fn set_body_text(&mut self, replacement: &str) -> Result<()> {
        if self.sections.len() > 1 {
            return Err(Error::ParseError(
                "Cannot replace a multi-section Pages body; use set_section_text".to_owned(),
            ));
        }
        let body_length = self.body_text()?.encode_utf16().count();
        self.replace_body_text(0..body_length, replacement)
    }

    /// Clear the complete body of a single-section document.
    pub fn clear_body(&mut self) -> Result<()> {
        self.set_body_text("")
    }

    /// Read the text owned by one reachable section, excluding native section-break markers.
    pub fn section_text(&self, section_id: u64) -> Result<String> {
        let body = self.body_text()?;
        let units = body.encode_utf16().collect::<Vec<_>>();
        let range = self.section_content_range(section_id, &units)?;
        String::from_utf16(&units[range]).map_err(|_| {
            Error::InvalidFormat(format!(
                "Pages section {section_id} boundary splits a UTF-16 surrogate pair"
            ))
        })
    }

    /// Replace a section-relative UTF-16 text range.
    pub fn replace_section_text(
        &mut self,
        section_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        if range.start > range.end {
            return Err(Error::ParseError(
                "Section text replacement range starts after it ends".to_owned(),
            ));
        }
        let body = self.body_text()?;
        let units = body.encode_utf16().collect::<Vec<_>>();
        let content = self.section_content_range(section_id, &units)?;
        let content_length = content.end - content.start;
        if range.end > content_length {
            return Err(Error::ParseError(format!(
                "Pages section {section_id} text range {}..{} exceeds its UTF-16 length {content_length}",
                range.start, range.end
            )));
        }
        let absolute_start = content
            .start
            .checked_add(range.start)
            .ok_or_else(|| Error::ParseError("Pages section text range overflow".to_owned()))?;
        let absolute_end = content
            .start
            .checked_add(range.end)
            .ok_or_else(|| Error::ParseError("Pages section text range overflow".to_owned()))?;
        self.replace_body_text(absolute_start..absolute_end, replacement)
    }

    /// Replace all text owned by one section while preserving its layout and neighboring sections.
    pub fn set_section_text(&mut self, section_id: u64, replacement: &str) -> Result<()> {
        let length = self.section_text(section_id)?.encode_utf16().count();
        self.replace_section_text(section_id, 0..length, replacement)
    }

    /// Clear all text owned by one section while preserving the section itself.
    pub fn clear_section_text(&mut self, section_id: u64) -> Result<()> {
        self.set_section_text(section_id, "")
    }

    fn validate_body_edit(&self, range: &Range<usize>, replacement: &str) -> Result<()> {
        if range.start > range.end {
            return Err(Error::ParseError(
                "Text replacement range starts after it ends".to_owned(),
            ));
        }
        if replacement.contains('\u{4}') {
            return Err(Error::ParseError(
                "Pages section breaks must be changed through section CRUD APIs".to_owned(),
            ));
        }
        let body = self.body_text()?;
        let units = body.encode_utf16().collect::<Vec<_>>();
        if range.end > units.len() {
            return Err(Error::ParseError(format!(
                "Text replacement range {}..{} exceeds body UTF-16 length {}",
                range.start,
                range.end,
                units.len()
            )));
        }
        for section in self.sections.iter().skip(1) {
            let boundary = usize::try_from(section.character_index).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Pages section {} boundary exceeds the platform index range",
                    section.object_id
                ))
            })?;
            let marker = boundary.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages section {} has an invalid zero boundary",
                    section.object_id
                ))
            })?;
            if units.get(marker) != Some(&0x0004) {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {} is not preceded by a native section-break marker",
                    section.object_id
                )));
            }
            if range.start <= marker && marker < range.end {
                return Err(Error::ParseError(format!(
                    "Text replacement range {}..{} crosses the section break before section {}",
                    range.start, range.end, section.object_id
                )));
            }
        }
        Ok(())
    }

    fn section_content_range(&self, section_id: u64, body: &[u16]) -> Result<Range<usize>> {
        let index = self
            .sections
            .iter()
            .position(|section| section.object_id == section_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Section {section_id} is not reachable from the Pages body"
                ))
            })?;
        let start = usize::try_from(self.sections[index].character_index).map_err(|_| {
            Error::InvalidFormat(format!(
                "Pages section {section_id} boundary exceeds the platform index range"
            ))
        })?;
        let end = if let Some(next) = self.sections.get(index + 1) {
            let boundary = usize::try_from(next.character_index).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Pages section {} boundary exceeds the platform index range",
                    next.object_id
                ))
            })?;
            let marker = boundary.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages section {} has an invalid zero boundary",
                    next.object_id
                ))
            })?;
            if body.get(marker) != Some(&0x0004) {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {} is not preceded by a native section-break marker",
                    next.object_id
                )));
            }
            marker
        } else {
            body.len()
        };
        if start > end || end > body.len() {
            return Err(Error::InvalidFormat(format!(
                "Pages section {section_id} has invalid body range {start}..{end}"
            )));
        }
        Ok(start..end)
    }
}
