//! Borrowed font-catalog views.

use crate::types::{
    EmbeddedFont as RawEmbedded, EmbeddedFontFormat, Font as RawFont, FontCharset, FontFamily,
    FontPage, FontPitch, FontRef, FontTable, FontTheme,
};
use std::fmt;
use std::iter::FusedIterator;

/// A borrowed document font definition.
///
/// This view deliberately hides the numeric RTF table reference. Resolve a
/// run's font through [`crate::text::Format::font`] or select a definition by
/// its semantic name through [`Catalog::find`].
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct Font<'a> {
    raw: &'a RawFont<'static>,
}

impl<'a> Font<'a> {
    pub(crate) const fn new(raw: &'a RawFont<'static>) -> Self {
        Self { raw }
    }

    /// Primary font name.
    pub fn name(self) -> &'a str {
        self.raw.name.as_ref()
    }

    /// Generic font family.
    pub const fn family(self) -> FontFamily {
        self.raw.family
    }

    /// Explicit RTF charset selector, if declared.
    pub const fn charset(self) -> Option<FontCharset> {
        self.raw.charset
    }

    /// Alternate font name, if declared.
    pub fn alternate_name(self) -> Option<&'a str> {
        self.raw.alternate_name.as_deref()
    }

    /// Non-tagged font name, if declared.
    pub fn non_tagged_name(self) -> Option<&'a str> {
        self.raw.non_tagged_name.as_deref()
    }

    /// Ten-byte PANOSE classification, if declared.
    pub const fn panose(self) -> Option<&'a [u8; 10]> {
        self.raw.panose.as_ref()
    }

    /// Font pitch preference.
    pub const fn pitch(self) -> FontPitch {
        self.raw.pitch
    }

    /// Explicit exact font code page, if declared.
    pub const fn code_page(self) -> Option<FontPage> {
        self.raw.code_page
    }

    /// Theme-font role, if declared.
    pub const fn theme(self) -> Option<FontTheme> {
        self.raw.theme
    }

    /// Inert embedded-font metadata and payload, if present.
    pub fn embedded(self) -> Option<Embedded<'a>> {
        self.raw.embedded.as_ref().map(Embedded::new)
    }
}

impl fmt::Debug for Font<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Font")
            .field("name", &self.name())
            .field("family", &self.family())
            .field("charset", &self.charset())
            .field("alternate_name", &self.alternate_name())
            .field("pitch", &self.pitch())
            .field("code_page", &self.code_page())
            .field("theme", &self.theme())
            .finish_non_exhaustive()
    }
}

/// Borrowed inert embedded-font data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Embedded<'a> {
    raw: &'a RawEmbedded<'static>,
}

impl<'a> Embedded<'a> {
    const fn new(raw: &'a RawEmbedded<'static>) -> Self {
        Self { raw }
    }

    /// Declared embedded-font format.
    pub const fn format(self) -> EmbeddedFontFormat {
        self.raw.format
    }

    /// External font-file name, if declared.
    pub fn file_name(self) -> Option<&'a str> {
        self.raw.file_name.as_deref()
    }

    /// Exact code page of the file name, if declared.
    pub const fn file_code_page(self) -> Option<FontPage> {
        self.raw.file_code_page
    }

    /// Inline embedded font bytes, if retained.
    pub fn data(self) -> Option<&'a [u8]> {
        self.raw.data.as_deref()
    }
}

/// A semantic font-name lookup failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LookupError {
    /// More than one font has the requested primary name.
    AmbiguousName,
}

impl fmt::Display for LookupError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::AmbiguousName => formatter.write_str("font name is ambiguous"),
        }
    }
}

impl std::error::Error for LookupError {}

/// A borrowed document font catalog.
///
/// Sparse numeric RTF slots are an encoding detail: `len`, `at`, and `iter`
/// operate only on definitions that actually appeared in the source.
#[derive(Clone, Copy)]
pub struct Catalog<'a> {
    raw: &'a FontTable<'static>,
}

impl<'a> Catalog<'a> {
    pub(crate) const fn new(raw: &'a FontTable<'static>) -> Self {
        Self { raw }
    }

    /// Number of defined fonts, excluding sparse RTF placeholder slots.
    pub fn len(self) -> usize {
        self.raw.defined.iter().filter(|defined| **defined).count()
    }

    /// Whether no font definitions were retained.
    pub fn is_empty(self) -> bool {
        !self.raw.defined.iter().any(|defined| *defined)
    }

    /// Return a checked zero-based logical catalog entry.
    ///
    /// Positions use stable catalog order rather than numeric RTF slots.
    /// Sparse placeholders are never returned.
    pub fn at(self, position: usize) -> Option<Font<'a>> {
        self.iter().nth(position)
    }

    /// Find a font by its exact primary name.
    ///
    /// `Ok(None)` means there is no match. Duplicate matching definitions are
    /// reported explicitly instead of silently selecting one.
    pub fn find(self, name: &str) -> Result<Option<Font<'a>>, LookupError> {
        let mut match_ = None;
        for font in self {
            if font.name() != name {
                continue;
            }
            if match_.is_some() {
                return Err(LookupError::AmbiguousName);
            }
            match_ = Some(font);
        }
        Ok(match_)
    }

    /// Lazily traverse defined fonts in stable catalog order.
    pub fn iter(self) -> Iter<'a> {
        Iter {
            catalog: self,
            front: 0,
            back: self.raw.fonts.len(),
        }
    }

    pub(crate) fn resolve(self, reference: FontRef) -> Option<Font<'a>> {
        self.raw.get(reference).map(Font::new)
    }
}

impl fmt::Debug for Catalog<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Catalog")
            .field("len", &self.len())
            .finish()
    }
}

impl<'a> IntoIterator for Catalog<'a> {
    type Item = Font<'a>;
    type IntoIter = Iter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

/// Lazy borrowed traversal over defined fonts.
#[derive(Clone)]
pub struct Iter<'a> {
    catalog: Catalog<'a>,
    front: usize,
    back: usize,
}

impl<'a> Iterator for Iter<'a> {
    type Item = Font<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.front < self.back {
            let index = self.front;
            self.front = self.front.saturating_add(1);
            let reference = FontRef::try_from(index).ok()?;
            if let Some(font) = self.catalog.resolve(reference) {
                return Some(font);
            }
        }
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (0, Some(self.back.saturating_sub(self.front)))
    }
}

impl DoubleEndedIterator for Iter<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        while self.front < self.back {
            self.back = self.back.saturating_sub(1);
            let reference = FontRef::try_from(self.back).ok()?;
            if let Some(font) = self.catalog.resolve(reference) {
                return Some(font);
            }
        }
        None
    }
}

impl FusedIterator for Iter<'_> {}
