//! Semantic bookmark-summary values.

/// One summary bookmark and its link to a text bookmark.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Bookmark {
    /// Source `BookmarkEntityAtomContainer` record instance.
    pub container_instance: u16,
    pub id: u32,
    pub name: String,
    pub value: String,
}

/// The bookmark collection in the optional document `SummaryContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Summary {
    pub id_seed: u32,
    pub bookmarks: Vec<Bookmark>,
}
