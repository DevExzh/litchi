//! Public semantic value types used by the Pages editor.

use crate::text::{TextPosition, TextStorageInfo};
use crate::{Error, Result};
use litchi_iwa_common::color::Rgba;
use litchi_pages::header_footer::{Kind, Template};
use litchi_pages::section::{PageNumber, PageNumbering, Start};

/// A reachable header/footer slot and its current writable text storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesHeaderFooterInfo {
    pub section_id: u64,
    pub section_name: Option<String>,
    /// UTF-16 position where the section begins in the body storage.
    pub section_character_index: u32,
    pub template_id: u64,
    pub template: Template,
    pub kind: Kind,
    /// Archive order within the header/footer list, normally left/center/right.
    pub slot: usize,
    pub storage: TextStorageInfo,
}

/// A writable text storage owned by a drawable reachable from a Pages document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesDrawableTextInfo {
    pub drawable_object_id: u64,
    pub storage: TextStorageInfo,
}

/// Result of removing a body-anchored Pages text box.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedPagesTextBox {
    pub text: PagesDrawableTextInfo,
    /// UTF-16 body position formerly occupied by the object-replacement character.
    pub anchor_character_index: u32,
}

/// A section boundary reachable from the main Pages body storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesSectionInfo {
    pub object_id: u64,
    /// UTF-16 position where the section begins in the body storage.
    pub character_index: u32,
    pub name: Option<String>,
    pub first_template_id: Option<u64>,
    pub even_template_id: Option<u64>,
    pub odd_template_id: Option<u64>,
}

/// Identifier of a native Pages body-footnote reference attachment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct PagesFootnoteId(u64);

impl PagesFootnoteId {
    /// Construct an identifier previously obtained from [`PagesFootnote`].
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "Pages footnote object identifier cannot be zero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the underlying iWork package object identifier.
    pub const fn object_id(self) -> u64 {
        self.0
    }

    pub(crate) const fn from_native(identifier: u64) -> Self {
        Self(identifier)
    }
}

/// One native footnote attached to the Pages main body.
///
/// `position` is the UTF-16 index of Pages' native U+000E footnote anchor.
/// `text` excludes the internal footnote-mark attachment and its separator.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesFootnote {
    /// Stable native reference-attachment identifier.
    pub id: PagesFootnoteId,
    /// UTF-16 position of the body anchor.
    pub position: TextPosition,
    /// Footnote content, stored without Pages' internal leading marker.
    pub text: Box<str>,
    /// Optional custom marker written by Pages instead of automatic numbering.
    pub custom_mark: Option<Box<str>>,
}

impl PagesFootnote {
    pub(crate) fn new(
        id: PagesFootnoteId,
        position: TextPosition,
        text: impl Into<Box<str>>,
        custom_mark: Option<impl Into<Box<str>>>,
    ) -> Self {
        Self {
            id,
            position,
            text: text.into(),
            custom_mark: custom_mark.map(Into::into),
        }
    }
}

/// Writable settings stored directly on a Pages section.
///
/// Unknown native discriminants remain lossless through their typed `Unknown`
/// variants. `background_fill_payload`, when present, is the exact encoded
/// `TSD.FillArchive` payload.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PagesSectionSettings {
    pub name: Option<String>,
    pub inherit_previous_header_footer: Option<bool>,
    pub first_page_different: Option<bool>,
    pub even_odd_pages_different: Option<bool>,
    pub start: Option<Start>,
    pub page_numbering: Option<PageNumbering>,
    pub starting_page_number: Option<PageNumber>,
    pub first_page_hides_header_footer: Option<bool>,
    pub background_fill_payload: Option<Vec<u8>>,
}

/// Semantic Pages section background.
///
/// Gradient, image, extension, and future fills are exposed as `Opaque` so
/// callers can round-trip them losslessly through the same API.
#[derive(Debug, Clone, PartialEq)]
pub enum PagesSectionBackground {
    None,
    Solid(Rgba),
    Opaque(Vec<u8>),
}
