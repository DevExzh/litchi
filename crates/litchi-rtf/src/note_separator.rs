//! Inert footnote and endnote separator destinations.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_NOTE_SEPARATOR_TEXT_BYTES: usize = 65_536;
pub(crate) const MAX_NOTE_SEPARATOR_ELEMENTS: usize = 1_024;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum NoteSeparatorKind {
    FootnoteSeparator,
    FootnoteContinuationSeparator,
    FootnoteContinuationNotice,
    EndnoteSeparator,
    EndnoteContinuationSeparator,
    EndnoteContinuationNotice,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NoteSeparatorElement<'a> {
    Text(Cow<'a, str>),
    SeparatorMark,
    ContinuationSeparatorMark,
    ParagraphBreak,
    LineBreak,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NoteSeparator<'a> {
    pub kind: NoteSeparatorKind,
    pub elements: Vec<NoteSeparatorElement<'a>>,
}

impl<'a> NoteSeparator<'a> {
    pub fn validate(&self) -> RtfResult<()> {
        if self.elements.len() > MAX_NOTE_SEPARATOR_ELEMENTS {
            return Err(RtfError::MalformedDocument(
                "RTF note separator contains too many elements".to_string(),
            ));
        }
        let text_bytes = self.elements.iter().try_fold(0usize, |total, element| {
            total.checked_add(match element {
                NoteSeparatorElement::Text(text) => text.len(),
                _ => 0,
            })
        }).ok_or_else(|| RtfError::MalformedDocument("RTF note-separator size overflow".to_string()))?;
        if text_bytes > MAX_NOTE_SEPARATOR_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF note-separator text exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> NoteSeparator<'static> {
        NoteSeparator {
            kind: self.kind,
            elements: self.elements.into_iter().map(|element| match element {
                NoteSeparatorElement::Text(text) => NoteSeparatorElement::Text(Cow::Owned(text.into_owned())),
                NoteSeparatorElement::SeparatorMark => NoteSeparatorElement::SeparatorMark,
                NoteSeparatorElement::ContinuationSeparatorMark => NoteSeparatorElement::ContinuationSeparatorMark,
                NoteSeparatorElement::ParagraphBreak => NoteSeparatorElement::ParagraphBreak,
                NoteSeparatorElement::LineBreak => NoteSeparatorElement::LineBreak,
            }).collect(),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NoteSeparatorTable<'a> {
    entries: Vec<NoteSeparator<'a>>,
}

impl<'a> NoteSeparatorTable<'a> {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn entries(&self) -> &[NoteSeparator<'a>] {
        &self.entries
    }

    pub fn get(&self, kind: NoteSeparatorKind) -> Option<&NoteSeparator<'a>> {
        self.entries.iter().find(|entry| entry.kind == kind)
    }

    pub fn add(&mut self, separator: NoteSeparator<'a>) -> RtfResult<()> {
        separator.validate()?;
        if self.entries.len() >= 6
            || self.entries.last().is_some_and(|entry| entry.kind >= separator.kind)
        {
            return Err(RtfError::MalformedDocument(
                "RTF note-separator destinations are duplicated or out of order".to_string(),
            ));
        }
        self.entries.push(separator);
        Ok(())
    }

    pub fn validate(&self) -> RtfResult<()> {
        let mut previous = None;
        for entry in &self.entries {
            entry.validate()?;
            if previous.is_some_and(|kind| kind >= entry.kind) {
                return Err(RtfError::MalformedDocument(
                    "RTF note-separator destinations are duplicated or out of order".to_string(),
                ));
            }
            previous = Some(entry.kind);
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> NoteSeparatorTable<'static> {
        NoteSeparatorTable {
            entries: self.entries.into_iter().map(NoteSeparator::into_owned).collect(),
        }
    }
}
