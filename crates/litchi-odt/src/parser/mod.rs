//! ODT-specific parsing utilities.
//!
//! This module provides parsing functionality that is specific to `OpenDocument` Text
//! documents (.odt). For generic ODF element parsing (paragraphs, tables, lists, etc.)
//! that works across all ODF formats, see `crate::elements::parser::Parser`.
//!
//! The owner keeps its public facade and semantic models stable while separating
//! XML codecs, document-level composition, and focused regression coverage.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

use litchi_core::Result;

pub(super) const MAX_SEMANTIC_DEPTH: usize = 4_096;
pub(super) const MAX_SEMANTIC_ITEMS: usize = 1_000_000;

pub use model::{
    ChangeType, Comment, Section, SectionDdeSource, SectionDisplay, SectionSource, TrackChange,
    TrackedChanges,
};

/// Parser for ODT-specific structures.
///
/// This provides parsing logic specific to text documents, such as:
/// - Track changes (insertions, deletions, formatting changes)
/// - Comments and annotations
/// - Sections (protected content, different formatting)
/// - Headers and footers
///
/// For generic element parsing (paragraphs, tables, etc.), use `Parser`
/// from `crate::elements::parser` instead.
pub struct Parser;

impl Parser {
    /// Parse track changes from content.
    pub fn parse_track_changes(content: &str) -> Result<Vec<TrackChange>> {
        package::parse_track_changes(content)
    }

    /// Parse the full tracked-change container, including inert policy metadata.
    pub fn parse_tracked_changes(content: &str) -> Result<TrackedChanges> {
        package::parse_tracked_changes(content)
    }

    /// Parse comments/annotations.
    pub fn parse_comments(content: &str) -> Result<Vec<Comment>> {
        package::parse_comments(content)
    }

    /// Parse sections.
    pub fn parse_sections(content: &str) -> Result<Vec<Section>> {
        package::parse_sections(content)
    }
}
