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
    Drawing(crate::StoryDrawing),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NoteSeparator<'a> {
    pub kind: NoteSeparatorKind,
    pub elements: Vec<NoteSeparatorElement<'a>>,
    pub shapes: Vec<crate::Shape<'a>>,
    pub shape_groups: Vec<crate::ShapeGroup<'a>>,
}

impl<'a> NoteSeparator<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.elements.len() > MAX_NOTE_SEPARATOR_ELEMENTS {
            return Err(RtfError::MalformedDocument(
                "RTF note separator contains too many elements".to_string(),
            ));
        }
        let text_bytes = self
            .elements
            .iter()
            .try_fold(0usize, |total, element| {
                total.checked_add(match element {
                    NoteSeparatorElement::Text(text) => text.len(),
                    NoteSeparatorElement::SeparatorMark
                    | NoteSeparatorElement::ContinuationSeparatorMark
                    | NoteSeparatorElement::ParagraphBreak
                    | NoteSeparatorElement::LineBreak
                    | NoteSeparatorElement::Drawing(_) => 0,
                })
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF note-separator size overflow".to_string())
            })?;
        if text_bytes > MAX_NOTE_SEPARATOR_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF note-separator text exceeds the safety limit".to_string(),
            ));
        }
        let story = self.text();
        let order = self.drawing_order();
        crate::shape::validate_story_drawings(
            &story,
            &self.shapes,
            &self.shape_groups,
            &order,
            "note separator",
        )?;
        Ok(())
    }

    #[must_use]
    pub fn text(&self) -> String {
        let mut text = String::new();
        for element in &self.elements {
            match element {
                NoteSeparatorElement::Text(value) => text.push_str(value),
                NoteSeparatorElement::ParagraphBreak | NoteSeparatorElement::LineBreak => {
                    text.push('\n');
                },
                NoteSeparatorElement::SeparatorMark
                | NoteSeparatorElement::ContinuationSeparatorMark
                | NoteSeparatorElement::Drawing(_) => {},
            }
        }
        text
    }

    #[must_use]
    pub fn drawing_order(&self) -> Vec<crate::StoryDrawing> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                NoteSeparatorElement::Drawing(drawing) => Some(*drawing),
                NoteSeparatorElement::Text(_)
                | NoteSeparatorElement::SeparatorMark
                | NoteSeparatorElement::ContinuationSeparatorMark
                | NoteSeparatorElement::ParagraphBreak
                | NoteSeparatorElement::LineBreak => None,
            })
            .collect()
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_shape(&mut self, shape: crate::Shape<'a>) -> RtfResult<()> {
        self.elements
            .push(NoteSeparatorElement::Drawing(crate::StoryDrawing::Shape(
                self.shapes.len(),
            )));
        self.shapes.push(shape);
        if let Err(error) = self.validate() {
            self.shapes.pop();
            self.elements.pop();
            return Err(error);
        }
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'a>) -> RtfResult<()> {
        self.elements.push(NoteSeparatorElement::Drawing(
            crate::StoryDrawing::ShapeGroup(self.shape_groups.len()),
        ));
        self.shape_groups.push(group);
        if let Err(error) = self.validate() {
            self.shape_groups.pop();
            self.elements.pop();
            return Err(error);
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> NoteSeparator<'static> {
        NoteSeparator {
            kind: self.kind,
            shapes: self
                .shapes
                .into_iter()
                .map(crate::Shape::into_owned)
                .collect(),
            shape_groups: self
                .shape_groups
                .into_iter()
                .map(crate::ShapeGroup::into_owned)
                .collect(),
            elements: self
                .elements
                .into_iter()
                .map(|element| match element {
                    NoteSeparatorElement::Text(text) => {
                        NoteSeparatorElement::Text(Cow::Owned(text.into_owned()))
                    },
                    NoteSeparatorElement::SeparatorMark => NoteSeparatorElement::SeparatorMark,
                    NoteSeparatorElement::ContinuationSeparatorMark => {
                        NoteSeparatorElement::ContinuationSeparatorMark
                    },
                    NoteSeparatorElement::ParagraphBreak => NoteSeparatorElement::ParagraphBreak,
                    NoteSeparatorElement::LineBreak => NoteSeparatorElement::LineBreak,
                    NoteSeparatorElement::Drawing(drawing) => {
                        NoteSeparatorElement::Drawing(drawing)
                    },
                })
                .collect(),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NoteSeparatorTable<'a> {
    entries: Vec<NoteSeparator<'a>>,
}

impl<'a> NoteSeparatorTable<'a> {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn entries(&self) -> &[NoteSeparator<'a>] {
        &self.entries
    }

    #[must_use]
    pub fn get(&self, kind: NoteSeparatorKind) -> Option<&NoteSeparator<'a>> {
        self.entries.iter().find(|entry| entry.kind == kind)
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn add(&mut self, separator: NoteSeparator<'a>) -> RtfResult<()> {
        separator.validate()?;
        if self.entries.len() >= 6
            || self
                .entries
                .last()
                .is_some_and(|entry| entry.kind >= separator.kind)
        {
            return Err(RtfError::MalformedDocument(
                "RTF note-separator destinations are duplicated or out of order".to_string(),
            ));
        }
        self.entries.push(separator);
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
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
            entries: self
                .entries
                .into_iter()
                .map(NoteSeparator::into_owned)
                .collect(),
        }
    }
}
