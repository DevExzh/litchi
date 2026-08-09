//! Inert source marks used to generate indexes and tables of contents.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_NAVIGATION_ENTRIES: usize = 65_536;
pub(crate) const MAX_NAVIGATION_ENTRY_TEXT_BYTES: usize = 1_048_576;
pub(crate) const MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES: usize = 16 * 1_048_576;
pub(crate) const MAX_NAVIGATION_ENTRY_DEPTH: usize = 64;

/// The page reference generated for an index entry.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum IndexPageReference<'a> {
    /// Generate the page containing the mark.
    #[default]
    CurrentPage,
    /// Use literal text instead of a page number (`\txe`).
    ReplacementText(Cow<'a, str>),
    /// Generate the page range covered by this inert bookmark name (`\rxe`).
    BookmarkRange(Cow<'a, str>),
}

/// One `\xe` index source mark.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct IndexEntry<'a> {
    /// UTF-8 byte offset in the extracted body text.
    pub position: usize,
    pub text: Cow<'a, str>,
    /// ASCII index identifier `A` through `Z` from `\xefN`.
    pub index_id: Option<u8>,
    pub bold_page_number: bool,
    pub italic_page_number: bool,
    pub page_reference: IndexPageReference<'a>,
    /// East Asian pronunciation text from `\yxe{\*\pxe ...}`.
    pub yomi: Option<Cow<'a, str>>,
}

impl<'a> IndexEntry<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(position: usize, text: impl Into<Cow<'a, str>>) -> RtfResult<Self> {
        let entry = Self {
            position,
            text: text.into(),
            index_id: None,
            bold_page_number: false,
            italic_page_number: false,
            page_reference: IndexPageReference::CurrentPage,
            yomi: None,
        };
        entry.validate()?;
        Ok(entry)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        validate_text("index-entry", &self.text, false)?;
        if self
            .index_id
            .is_some_and(|value| !value.is_ascii_uppercase())
        {
            return Err(RtfError::MalformedDocument(
                "RTF index identifier must be ASCII A through Z".to_string(),
            ));
        }
        match &self.page_reference {
            IndexPageReference::CurrentPage => {},
            IndexPageReference::ReplacementText(value) => {
                validate_text("index replacement", value, false)?;
            },
            IndexPageReference::BookmarkRange(value) => {
                validate_text("index bookmark range", value, false)?;
            },
        }
        if let Some(value) = &self.yomi {
            validate_text("index pronunciation", value, false)?;
        }
        Ok(())
    }

    pub(crate) fn text_bytes(&self) -> Option<usize> {
        self.text
            .len()
            .checked_add(match &self.page_reference {
                IndexPageReference::CurrentPage => 0,
                IndexPageReference::ReplacementText(value)
                | IndexPageReference::BookmarkRange(value) => value.len(),
            })?
            .checked_add(self.yomi.as_ref().map_or(0, |value| value.len()))
    }

    #[must_use]
    pub fn into_owned(self) -> IndexEntry<'static> {
        IndexEntry {
            position: self.position,
            text: Cow::Owned(self.text.into_owned()),
            index_id: self.index_id,
            bold_page_number: self.bold_page_number,
            italic_page_number: self.italic_page_number,
            page_reference: match self.page_reference {
                IndexPageReference::CurrentPage => IndexPageReference::CurrentPage,
                IndexPageReference::ReplacementText(value) => {
                    IndexPageReference::ReplacementText(Cow::Owned(value.into_owned()))
                },
                IndexPageReference::BookmarkRange(value) => {
                    IndexPageReference::BookmarkRange(Cow::Owned(value.into_owned()))
                },
            },
            yomi: self.yomi.map(|value| Cow::Owned(value.into_owned())),
        }
    }
}

/// One `\tc` or `\tcn` table-of-contents source mark.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableOfContentsEntry<'a> {
    /// UTF-8 byte offset in the extracted body text.
    pub position: usize,
    pub text: Cow<'a, str>,
    /// ASCII table identifier `A` through `Z`; the normative default is `C`.
    pub table_id: u8,
    /// Table-of-contents level, from 1 through 9.
    pub level: u8,
    pub suppress_page_number: bool,
}

impl<'a> TableOfContentsEntry<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(position: usize, text: impl Into<Cow<'a, str>>) -> RtfResult<Self> {
        let entry = Self {
            position,
            text: text.into(),
            table_id: b'C',
            level: 1,
            suppress_page_number: false,
        };
        entry.validate()?;
        Ok(entry)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        validate_text("table-of-contents entry", &self.text, false)?;
        if !self.table_id.is_ascii_uppercase() {
            return Err(RtfError::MalformedDocument(
                "RTF table-of-contents identifier must be ASCII A through Z".to_string(),
            ));
        }
        if !(1..=9).contains(&self.level) {
            return Err(RtfError::MalformedDocument(
                "RTF table-of-contents level must be between 1 and 9".to_string(),
            ));
        }
        Ok(())
    }

    #[must_use]
    pub fn into_owned(self) -> TableOfContentsEntry<'static> {
        TableOfContentsEntry {
            position: self.position,
            text: Cow::Owned(self.text.into_owned()),
            table_id: self.table_id,
            level: self.level,
            suppress_page_number: self.suppress_page_number,
        }
    }
}

/// An ordered, inert index or table-of-contents source mark.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum NavigationEntry<'a> {
    Index(IndexEntry<'a>),
    TableOfContents(TableOfContentsEntry<'a>),
}

impl NavigationEntry<'_> {
    #[must_use]
    pub fn position(&self) -> usize {
        match self {
            Self::Index(entry) => entry.position,
            Self::TableOfContents(entry) => entry.position,
        }
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        match self {
            Self::Index(entry) => entry.validate(),
            Self::TableOfContents(entry) => entry.validate(),
        }
    }

    pub(crate) fn text_bytes(&self) -> Option<usize> {
        match self {
            Self::Index(entry) => entry.text_bytes(),
            Self::TableOfContents(entry) => Some(entry.text.len()),
        }
    }

    #[must_use]
    pub fn into_owned(self) -> NavigationEntry<'static> {
        match self {
            Self::Index(entry) => NavigationEntry::Index(entry.into_owned()),
            Self::TableOfContents(entry) => NavigationEntry::TableOfContents(entry.into_owned()),
        }
    }
}

fn validate_text(kind: &str, text: &str, allow_empty: bool) -> RtfResult<()> {
    if !allow_empty && text.is_empty() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {kind} text cannot be empty"
        )));
    }
    if text.len() > MAX_NAVIGATION_ENTRY_TEXT_BYTES {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {kind} text exceeds {MAX_NAVIGATION_ENTRY_TEXT_BYTES} bytes"
        )));
    }
    Ok(())
}
