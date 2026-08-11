//! Namespace-aware ODT text block decoding.
//!
//! The low-level event state machine remains private to the text-element
//! owner; this module provides the typed codec facade used by callers. Keeping
//! the facade separate lets future XML writers and snapshot edits share the
//! semantic [`Block`] model without coupling to parser state or wire details.

use super::model::Block;
use super::{Heading, Paragraph, parse_paragraph_at, parse_text_blocks, parse_text_blocks_owned};
use litchi_core::Result;

/// Collection of typed text-element codec operations.
pub struct Elements;

impl Elements {
    /// Decode all `text:p` and `text:h` blocks in document order.
    pub fn parse(xml_content: &str) -> Result<Vec<Block>> {
        parse_text_blocks(xml_content)
    }

    /// Decode all block-level text elements in document order.
    pub fn parse_blocks(xml_content: &str) -> Result<Vec<Block>> {
        Self::parse(xml_content)
    }

    /// Decode all paragraphs from an XML reader.
    pub fn parse_paragraphs(xml_content: &str) -> Result<Vec<Paragraph>> {
        Self::parse(xml_content).map(|blocks| {
            blocks
                .into_iter()
                .filter_map(Block::into_paragraph)
                .collect()
        })
    }

    /// Decode the paragraph at `index` without retaining the other text blocks.
    ///
    /// The decoder still scans the complete input and applies the same XML and
    /// resource-limit validation as [`Self::parse_paragraphs`].
    pub fn parse_paragraph_at(xml_content: &str, index: usize) -> Result<Option<Paragraph>> {
        parse_paragraph_at(xml_content, index)
    }

    /// Decode all headings from XML content.
    pub fn parse_headings(xml_content: &str) -> Result<Vec<Heading>> {
        Self::parse(xml_content)
            .map(|blocks| blocks.into_iter().filter_map(Block::into_heading).collect())
    }

    /// Extract visible text from all decoded blocks, preserving block breaks.
    ///
    /// Namespace-aware parsing includes blocks nested in lists, sections, and
    /// text boxes. Tracked-change definitions, note bodies, and ruby
    /// pronunciation runs remain excluded by the decoder.
    pub fn extract_text(xml_content: &str) -> Result<String> {
        let mut blocks = parse_text_blocks_owned(xml_content)?.into_iter();
        let Some(first) = blocks.next() else {
            return Ok(String::new());
        };
        let mut output = first.into_text();
        for block in blocks {
            output.push('\n');
            output.push_str(&block.into_text());
        }
        Ok(output)
    }
}
