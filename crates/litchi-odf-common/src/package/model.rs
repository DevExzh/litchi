//! Neutral ODF package and manifest models.

use soapberry_zip::office::{
    ArchiveReader, ArchiveReaderNames, IndexedArchive, IndexedArchiveNames,
};
use std::collections::HashMap;
use std::sync::Arc;

#[cfg(test)]
use std::cell::Cell;

#[cfg(test)]
thread_local! {
    static INDEX_BUILDS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
pub(crate) fn note_index_build() {
    INDEX_BUILDS.with(|count| count.set(count.get().saturating_add(1)));
}

#[cfg(test)]
pub(crate) fn reset_index_build_count() {
    INDEX_BUILDS.with(|count| count.set(0));
}

#[cfg(test)]
pub(crate) fn index_build_count() -> usize {
    INDEX_BUILDS.with(Cell::get)
}

/// A borrowed, lazily decoded ODF ZIP archive.
pub struct Archive<'data> {
    pub(super) reader: ArchiveReaderKind<'data>,
}

/// The archive index retained by an owned ODF package.
pub(crate) type PreparedArchive = Arc<IndexedArchive<Arc<Vec<u8>>>>;

/// A borrowed archive reader or an already-indexed owned archive.
pub(super) enum ArchiveReaderKind<'data> {
    Borrowed(ArchiveReader<'data>),
    Prepared(PreparedArchive),
}

pub enum ArchiveNames<'data> {
    Borrowed(ArchiveReaderNames<'data>),
    Prepared(IndexedArchiveNames<'data, Arc<Vec<u8>>>),
}

impl<'data> Iterator for ArchiveNames<'data> {
    type Item = &'data str;

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::Borrowed(names) => names.next(),
            Self::Prepared(names) => names.next(),
        }
    }
}

impl ExactSizeIterator for ArchiveNames<'_> {}

/// The family-neutral portion of `META-INF/manifest.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Manifest {
    /// The media type declared for the root entry.
    pub mimetype: String,
    /// Entries keyed by their normalized manifest path.
    pub entries: HashMap<String, Entry>,
}

/// One family-neutral ODF manifest file entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    /// The media type declared by `manifest:media-type`.
    pub media_type: String,
    /// The optional plaintext size declared by `manifest:size`.
    pub size: Option<u64>,
}

impl Manifest {
    /// Return the media type for a manifest path, if declared.
    #[must_use]
    pub fn get_media_type(&self, path: &str) -> Option<&str> {
        self.entries
            .get(path)
            .map(|entry| entry.media_type.as_str())
    }

    /// Check whether a manifest path is declared.
    #[must_use]
    pub fn has_path(&self, path: &str) -> bool {
        self.entries.contains_key(path)
    }

    /// Iterate over declared manifest paths.
    pub fn paths(&self) -> impl Iterator<Item = &str> {
        self.entries.keys().map(String::as_str)
    }

    /// Return a manifest entry by path.
    #[must_use]
    pub fn get_entry(&self, path: &str) -> Option<&Entry> {
        self.entries.get(path)
    }
}
