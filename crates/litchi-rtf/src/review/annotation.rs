//! RTF annotation and comment support.
//!
//! This module provides support for comments, revisions, and other annotations
//! in RTF documents.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_ANNOTATIONS: usize = 65_536;
pub(crate) const MAX_ANNOTATION_METADATA_BYTES: usize = 65_536;
pub(crate) const MAX_ANNOTATION_BODY_BYTES: usize = 4 * 1_048_576;
pub(crate) const MAX_ANNOTATION_TEXT_TOTAL_BYTES: usize = 16 * 1_048_576;
pub(crate) const MAX_REVISION_AUTHORS: usize = 65_536;
pub(crate) const MAX_REVISION_AUTHOR_BYTES: usize = 65_536;
pub(crate) const MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES: usize = 16 * 1_048_576;
pub(crate) const MAX_REVISIONS: usize = 65_536;
pub(crate) const MAX_REVISION_TEXT_BYTES: usize = 4 * 1_048_576;
pub(crate) const MAX_REVISION_TEXT_TOTAL_BYTES: usize = 16 * 1_048_576;

/// Annotation type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnnotationType {
    /// Comment/note
    Comment,
    /// Revision mark (tracked change)
    Revision,
    /// Highlight
    Highlight,
}

/// Revision type (for tracked changes)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionType {
    /// Inserted text
    Insertion,
    /// Deleted text
    Deletion,
    /// Formatting change
    FormatChange,
    /// Moved from location
    MovedFrom,
    /// Moved to location
    MovedTo,
}

/// Comment or annotation
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Annotation<'a> {
    /// Annotation type
    pub annotation_type: AnnotationType,
    /// Annotation ID
    pub id: i32,
    /// Whether the source contained `atnref` and corresponding range identity.
    /// `LibreOffice` also emits valid point comments without a reference.
    pub has_reference: bool,
    /// Author name
    pub author: Cow<'a, str>,
    /// Author initials from the `atnid` destination.
    pub initials: Cow<'a, str>,
    /// Creation date (RTF datetime format)
    pub date: Option<Cow<'a, str>>,
    /// Comment text
    pub text: Cow<'a, str>,
    /// Positional root shapes owned by the comment story.
    pub shapes: Vec<crate::Shape<'a>>,
    /// Positional root shape groups owned by the comment story.
    pub shape_groups: Vec<crate::ShapeGroup<'a>>,
    /// Exact source order of drawings in the comment story.
    pub drawing_order: Vec<crate::StoryDrawing>,
    /// Exact source order of drawings, fields, and page breaks in the comment story.
    pub story_events: Vec<crate::StoryEvent>,
    /// UTF-8 byte offset where the annotated range starts.
    pub position: usize,
    /// UTF-8 byte offset where the annotated range ends.
    pub range_end: usize,
    /// Parent annotation identifier for threaded comments.
    pub parent_id: Option<Cow<'a, str>>,
    /// Annotation icon identifier, preserved from `atnicn`.
    pub icon: Option<Cow<'a, str>>,
    /// Annotation time value, preserved from `atntime`.
    pub time: Option<Cow<'a, str>>,
}

impl<'a> Annotation<'a> {
    /// Create a new comment
    #[inline]
    #[must_use]
    pub fn comment(id: i32, author: Cow<'a, str>, text: Cow<'a, str>) -> Self {
        Self {
            annotation_type: AnnotationType::Comment,
            id,
            has_reference: true,
            author,
            initials: Cow::Borrowed(""),
            date: None,
            text,
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            story_events: Vec::new(),
            position: 0,
            range_end: 0,
            parent_id: None,
            icon: None,
            time: None,
        }
    }

    /// Create a new revision mark
    #[inline]
    #[must_use]
    pub fn revision(id: i32, author: Cow<'a, str>) -> Self {
        Self {
            annotation_type: AnnotationType::Revision,
            id,
            has_reference: true,
            author,
            initials: Cow::Borrowed(""),
            date: None,
            text: Cow::Borrowed(""),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            story_events: Vec::new(),
            position: 0,
            range_end: 0,
            parent_id: None,
            icon: None,
            time: None,
        }
    }

    /// Validate this inert annotation without resolving any reference.
    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.annotation_type != AnnotationType::Comment {
            return Err(RtfError::MalformedDocument(
                "only comment annotations use the RTF annotation destination".to_string(),
            ));
        }
        if self.range_end < self.position {
            return Err(RtfError::MalformedDocument(
                "RTF annotation range end precedes its start".to_string(),
            ));
        }
        if self.text.len() > MAX_ANNOTATION_BODY_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF annotation body exceeds the safety limit".to_string(),
            ));
        }
        crate::field::validate_story_events(
            self.text.as_ref(),
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            &self.story_events,
            "annotation",
        )?;
        for (kind, value) in [
            ("author", Some(self.author.as_ref())),
            ("initials", Some(self.initials.as_ref())),
            ("date", self.date.as_deref()),
            ("parent", self.parent_id.as_deref()),
            ("icon", self.icon.as_deref()),
            ("time", self.time.as_deref()),
        ] {
            if value.is_some_and(|text| text.len() > MAX_ANNOTATION_METADATA_BYTES) {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF annotation {kind} exceeds the safety limit"
                )));
            }
        }
        Ok(())
    }

    /// Append a validated positional root shape to the comment story.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_shape(&mut self, shape: crate::Shape<'a>) -> RtfResult<()> {
        let mut shapes = self.shapes.clone();
        shapes.push(shape);
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::Shape(self.shapes.len()));
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &shapes,
            &self.shape_groups,
            &order,
            "annotation",
        )?;
        self.shapes = shapes;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::Shape(
                self.shapes.len() - 1,
            )));
        Ok(())
    }

    /// Append a validated positional root shape group to the comment story.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'a>) -> RtfResult<()> {
        let mut groups = self.shape_groups.clone();
        groups.push(group);
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::ShapeGroup(self.shape_groups.len()));
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &self.shapes,
            &groups,
            &order,
            "annotation",
        )?;
        self.shape_groups = groups;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(
                self.shape_groups.len() - 1,
            )));
        Ok(())
    }

    /// Clear all drawings owned by the comment story.
    pub fn clear_drawings(&mut self) {
        self.shapes.clear();
        self.shape_groups.clear();
        self.drawing_order.clear();
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::Drawing(_)));
    }

    pub fn page_breaks(&self) -> impl Iterator<Item = &crate::PageBreak> {
        self.story_events.iter().filter_map(|event| match event {
            crate::StoryEvent::PageBreak(page_break) => Some(page_break),
            crate::StoryEvent::Drawing(_) | crate::StoryEvent::Field(_) => None,
        })
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_page_break(&mut self, position: usize) -> RtfResult<()> {
        crate::field::push_story_page_break(
            &mut self.story_events,
            self.text.as_ref(),
            position,
            "annotation",
        )
    }

    pub fn clear_page_breaks(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::PageBreak(_)));
    }

    pub(crate) fn text_bytes(&self) -> Option<usize> {
        self.text
            .len()
            .checked_add(self.author.len())?
            .checked_add(self.initials.len())?
            .checked_add(self.date.as_ref().map_or(0, |value| value.len()))?
            .checked_add(self.parent_id.as_ref().map_or(0, |value| value.len()))?
            .checked_add(self.icon.as_ref().map_or(0, |value| value.len()))?
            .checked_add(self.time.as_ref().map_or(0, |value| value.len()))
    }

    pub fn into_owned(self) -> Annotation<'static> {
        Annotation {
            annotation_type: self.annotation_type,
            id: self.id,
            has_reference: self.has_reference,
            author: Cow::Owned(self.author.into_owned()),
            initials: Cow::Owned(self.initials.into_owned()),
            date: self.date.map(|value| Cow::Owned(value.into_owned())),
            text: Cow::Owned(self.text.into_owned()),
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
            drawing_order: self.drawing_order,
            story_events: self.story_events,
            position: self.position,
            range_end: self.range_end,
            parent_id: self.parent_id.map(|value| Cow::Owned(value.into_owned())),
            icon: self.icon.map(|value| Cow::Owned(value.into_owned())),
            time: self.time.map(|value| Cow::Owned(value.into_owned())),
        }
    }
}

/// Revision information (tracked change)
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Revision<'a> {
    /// Revision type
    pub revision_type: RevisionType,
    /// Author name
    pub author: Cow<'a, str>,
    /// Packed signed RTF DTTM value, stored as a decimal string
    pub date: Option<Cow<'a, str>>,
    /// Revision-author table index
    pub id: i32,
    /// Text content affected by revision
    pub content: Cow<'a, str>,
    /// UTF-8 byte offset where the revised range starts
    pub position: usize,
    /// UTF-8 byte offset where the revised range ends
    pub range_end: usize,
}

impl<'a> Revision<'a> {
    /// Create a new revision
    #[inline]
    #[must_use]
    pub fn new(revision_type: RevisionType, author: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self {
            revision_type,
            author,
            date: None,
            id: 0,
            content,
            position: 0,
            range_end: 0,
        }
    }

    /// Create an insertion revision
    #[inline]
    #[must_use]
    pub fn insertion(author: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self::new(RevisionType::Insertion, author, content)
    }

    /// Create a deletion revision
    #[inline]
    #[must_use]
    pub fn deletion(author: Cow<'a, str>, content: Cow<'a, str>) -> Self {
        Self::new(RevisionType::Deletion, author, content)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.id < 0 {
            return Err(RtfError::MalformedDocument(
                "RTF revision author index cannot be negative".to_string(),
            ));
        }
        if self.content.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF revision content cannot be empty".to_string(),
            ));
        }
        if self.content.len() > MAX_REVISION_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF revision content exceeds the safety limit".to_string(),
            ));
        }
        if self
            .date
            .as_deref()
            .is_some_and(|date| date.parse::<i32>().is_err())
        {
            return Err(RtfError::MalformedDocument(
                "RTF revision date must contain a packed signed DTTM value".to_string(),
            ));
        }
        match self.revision_type {
            RevisionType::Insertion if self.range_end <= self.position => {
                return Err(RtfError::MalformedDocument(
                    "RTF insertion revision must cover a non-empty body range".to_string(),
                ));
            },
            RevisionType::Deletion if self.range_end != self.position => {
                return Err(RtfError::MalformedDocument(
                    "RTF deletion revision must be positioned between visible body characters"
                        .to_string(),
                ));
            },
            RevisionType::FormatChange | RevisionType::MovedFrom | RevisionType::MovedTo => {
                return Err(RtfError::MalformedDocument(
                    "this RTF revision kind has no lossless scoped-run representation".to_string(),
                ));
            },
            RevisionType::Insertion | RevisionType::Deletion => {},
        }
        Ok(())
    }

    #[must_use]
    pub fn into_owned(self) -> Revision<'static> {
        Revision {
            revision_type: self.revision_type,
            author: Cow::Owned(self.author.into_owned()),
            date: self.date.map(|date| Cow::Owned(date.into_owned())),
            id: self.id,
            content: Cow::Owned(self.content.into_owned()),
            position: self.position,
            range_end: self.range_end,
        }
    }
}

/// One ordered entry in the inert RTF `revtbl` author table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RevisionAuthor<'a> {
    pub name: Cow<'a, str>,
}

impl<'a> RevisionAuthor<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(name: impl Into<Cow<'a, str>>) -> RtfResult<Self> {
        let author = Self { name: name.into() };
        author.validate()?;
        Ok(author)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.name.len() > MAX_REVISION_AUTHOR_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF revision author exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    #[must_use]
    pub fn into_owned(self) -> RevisionAuthor<'static> {
        RevisionAuthor {
            name: Cow::Owned(self.name.into_owned()),
        }
    }
}
