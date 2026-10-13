//! Exact source provenance of retained destination text stories.
//!
//! The parser records, for every top-level table cell, header/footer
//! paragraph, footnote/endnote body and comment body, the source byte range
//! of the tokens that produced its text and the offset of the control or
//! group delimiter that terminates it.  Destination text edits use these ranges to splice
//! replacement text into the retained source instead of reserializing the
//! whole document, so untouched content stays byte-identical.
//!
//! Provenance is advisory: an editor must still validate the located bytes
//! and reparse the candidate before publication.

use std::ops::Range;

/// Source bytes that produced one retained destination text story.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct StorySource {
    /// Byte range from the first text-producing token through the last one.
    pub(crate) text: Option<Range<usize>>,
    /// Byte offset of the control word or closing brace that ends the story.
    pub(crate) terminator: Option<usize>,
}

/// Outer table-cell coordinate: table, row and cell positions.
pub(crate) type CellKey = (usize, usize, usize);

/// Header/footer paragraph coordinate: section, header/footer position within
/// the section, and paragraph position within the header/footer.
pub(crate) type StoryParagraphKey = (usize, usize, usize);

/// Source provenance for every retained destination text story.
///
/// Entries are appended in parse order, which is also ascending key order, so
/// lookups binary-search without an index.
#[derive(Debug, Clone, Default)]
pub(crate) struct TextProvenance {
    cells: Vec<(CellKey, StorySource)>,
    header_footer_paragraphs: Vec<(StoryParagraphKey, StorySource)>,
    notes: Vec<(usize, StorySource)>,
    annotations: Vec<(usize, StorySource)>,
    /// Body text offset of each picture opened directly in the main story.
    pictures: Vec<(usize, usize)>,
}

impl TextProvenance {
    pub(crate) fn push_cell(&mut self, key: CellKey, source: StorySource) {
        if self.cells.last().is_none_or(|(last, _)| *last < key) {
            self.cells.push((key, source));
        }
    }

    pub(crate) fn push_header_footer_paragraph(
        &mut self,
        key: StoryParagraphKey,
        source: StorySource,
    ) {
        if self
            .header_footer_paragraphs
            .last()
            .is_none_or(|(last, _)| *last < key)
        {
            self.header_footer_paragraphs.push((key, source));
        }
    }

    pub(crate) fn push_note(&mut self, index: usize, source: StorySource) {
        if self.notes.last().is_none_or(|(last, _)| *last < index) {
            self.notes.push((index, source));
        }
    }

    pub(crate) fn push_annotation(&mut self, index: usize, source: StorySource) {
        if self
            .annotations
            .last()
            .is_none_or(|(last, _)| *last < index)
        {
            self.annotations.push((index, source));
        }
    }

    pub(crate) fn push_picture(&mut self, index: usize, body_offset: usize) {
        if self.pictures.last().is_none_or(|(last, _)| *last < index) {
            self.pictures.push((index, body_offset));
        }
    }

    pub(crate) fn picture(&self, index: usize) -> Option<usize> {
        self.pictures
            .binary_search_by(|(candidate, _)| candidate.cmp(&index))
            .ok()
            .and_then(|position| self.pictures.get(position))
            .map(|(_, offset)| *offset)
    }

    pub(crate) fn cell(&self, key: CellKey) -> Option<&StorySource> {
        lookup(&self.cells, &key)
    }

    pub(crate) fn header_footer_paragraph(&self, key: StoryParagraphKey) -> Option<&StorySource> {
        lookup(&self.header_footer_paragraphs, &key)
    }

    pub(crate) fn note(&self, index: usize) -> Option<&StorySource> {
        lookup(&self.notes, &index)
    }

    pub(crate) fn annotation(&self, index: usize) -> Option<&StorySource> {
        lookup(&self.annotations, &index)
    }
}

fn lookup<'a, K: Ord>(entries: &'a [(K, StorySource)], key: &K) -> Option<&'a StorySource> {
    entries
        .binary_search_by(|(candidate, _)| candidate.cmp(key))
        .ok()
        .and_then(|index| entries.get(index))
        .map(|(_, source)| source)
}
