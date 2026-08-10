//! ODF list semantics.

use crate::paragraph::Paragraph;
use litchi_core::Position;

/// One list item.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Item {
    nested_lists: Vec<List>,
    paragraphs: Vec<Paragraph>,
    positions: Vec<Position>,
    start_value: Option<u32>,
}

impl Item {
    /// Creates a detached single-paragraph item.
    #[must_use]
    pub fn new(paragraph: Paragraph) -> Self {
        Self {
            nested_lists: Vec::new(),
            paragraphs: vec![paragraph],
            positions: Vec::new(),
            start_value: None,
        }
    }

    /// Sets an explicit item start value.
    #[must_use]
    pub const fn with_start_value(mut self, value: u32) -> Self {
        self.start_value = Some(value);
        self
    }

    pub(crate) const fn projected(
        nested_lists: Vec<List>,
        paragraphs: Vec<Paragraph>,
        positions: Vec<Position>,
        start_value: Option<u32>,
    ) -> Self {
        Self {
            nested_lists,
            paragraphs,
            positions,
            start_value,
        }
    }

    /// Adds a detached nested list to this item.
    #[must_use]
    pub fn with_nested_list(mut self, list: List) -> Self {
        self.nested_lists.push(list);
        self
    }

    /// Nested lists directly contained by this item.
    #[must_use]
    pub fn nested_lists(&self) -> &[List] {
        &self.nested_lists
    }

    /// Paragraphs directly contained by this item.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// Their positions in the body's paragraph projection.
    ///
    /// Detached items have no positions until published and reopened.
    #[must_use]
    pub fn paragraph_positions(&self) -> &[Position] {
        &self.positions
    }

    /// Explicit numbering restart for this item.
    #[must_use]
    pub const fn start_value(&self) -> Option<u32> {
        self.start_value
    }
}

/// One `text:list`, including nested lists as separate entries with a level.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct List {
    items: Vec<Item>,
    level: usize,
    style_name: Option<String>,
}

impl List {
    /// Creates a detached top-level list.
    #[must_use]
    pub fn new(items: impl IntoIterator<Item = Item>) -> Self {
        Self {
            items: items.into_iter().collect(),
            level: 1,
            style_name: None,
        }
    }

    /// Creates a detached styled top-level list.
    #[must_use]
    pub fn styled(style_name: impl Into<String>, items: impl IntoIterator<Item = Item>) -> Self {
        Self {
            items: items.into_iter().collect(),
            level: 1,
            style_name: Some(style_name.into()),
        }
    }

    pub(crate) const fn projected(
        items: Vec<Item>,
        level: usize,
        style_name: Option<String>,
    ) -> Self {
        Self {
            items,
            level,
            style_name,
        }
    }

    /// List nesting level, starting at one.
    #[must_use]
    pub const fn level(&self) -> usize {
        self.level
    }

    /// Referenced list style.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Items in source order.
    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.items
    }
}
