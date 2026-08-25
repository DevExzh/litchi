//! Body text reading and mutation.

use std::ops::Range;

use super::PagesEditor;
use crate::text::{
    TextBookmark, TextBookmarkId, TextBookmarkSettings, TextDateTimeField, TextDateTimeFieldId,
    TextPosition, TextRange,
};
use crate::{Error, Result};
use litchi_iwa_text::date_time::{DisplayText, Settings};

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
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        self.text
            .add_text_date_time_field(self.body_storage_id, range, settings)
    }

    /// Atomically insert exact display text and its Date & Time field.
    pub fn insert_body_date_time_field(
        &mut self,
        position: TextPosition,
        display_text: DisplayText,
        settings: Settings,
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
        settings: Settings,
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
        let footnotes =
            super::footnotes::body_footnote_graphs(self.package(), self.body_storage_id.get())?;
        let mut staged = self.text.clone();
        staged.replace_text(self.body_storage_id, range, replacement)?;
        let mut package = staged.into_package();
        super::footnotes::cleanup_removed_body_footnotes(
            &mut package,
            self.body_storage_id.get(),
            &footnotes,
        )?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    /// Replace the complete body of a single-section document.
    ///
    /// Multi-section documents must be edited through the selector-first
    /// `litchi_pages::Package` section-text transaction so their native section
    /// breaks cannot be discarded accidentally.
    pub fn set_body_text(&mut self, replacement: &str) -> Result<()> {
        if self.sections.len() > 1 {
            return Err(Error::ParseError(
                "Cannot replace a multi-section Pages body; use the selector-first Pages section-text transaction".to_owned(),
            ));
        }
        let body_length = self.body_text()?.encode_utf16().count();
        self.replace_body_text(0..body_length, replacement)
    }

    /// Clear the complete body of a single-section document.
    pub fn clear_body(&mut self) -> Result<()> {
        self.set_body_text("")
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
        if replacement.contains('\u{e}') {
            return Err(Error::ParseError(
                "Pages footnote anchors must be changed through footnote CRUD APIs".to_owned(),
            ));
        }
        if replacement.contains('\u{fffc}') {
            return Err(Error::ParseError(
                "Pages inline-object markers must be changed through object CRUD APIs".to_owned(),
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
}
