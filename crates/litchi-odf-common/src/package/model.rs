//! Neutral ODF package and manifest models.

use soapberry_zip::office::ArchiveReader;
use std::collections::HashMap;

/// A borrowed, lazily decoded ODF ZIP archive.
pub struct Archive<'data> {
    pub(super) reader: ArchiveReader<'data>,
}

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
    pub fn get_media_type(&self, path: &str) -> Option<&str> {
        self.entries
            .get(path)
            .map(|entry| entry.media_type.as_str())
    }

    /// Check whether a manifest path is declared.
    pub fn has_path(&self, path: &str) -> bool {
        self.entries.contains_key(path)
    }

    /// Iterate over declared manifest paths.
    pub fn paths(&self) -> impl Iterator<Item = &str> {
        self.entries.keys().map(String::as_str)
    }

    /// Return a manifest entry by path.
    pub fn get_entry(&self, path: &str) -> Option<&Entry> {
        self.entries.get(path)
    }
}
