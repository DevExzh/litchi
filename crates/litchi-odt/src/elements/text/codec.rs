//! Namespace-aware ODT text block decoding.
//!
//! The low-level event state machine remains private to the text-element
//! owner; this module provides the typed codec facade used by callers. Keeping
//! the facade separate lets future XML writers and snapshot edits share the
//! semantic [`Block`] model without coupling to parser state or wire details.

use super::model::Block;
use super::{
    Heading, Paragraph, parse_block_at, parse_paragraph_at, parse_text_block_texts,
    parse_text_blocks,
};
use litchi_core::{
    Error, Result, SequentialTextWriter, TextOutputError, TextOutputOptions, TextOutputReport,
};
use std::io::Write;

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

    /// Decode one visible paragraph or heading without retaining the other
    /// text blocks.
    ///
    /// The decoder still scans the complete input and applies the same XML,
    /// text, and resource-limit validation as [`Self::parse`].
    pub fn parse_block_at(xml_content: &str, index: usize) -> Result<Option<Block>> {
        parse_block_at(xml_content, index)
    }

    /// Decode all paragraphs from an XML reader.
    pub fn parse_paragraphs(xml_content: &str) -> Result<Vec<Paragraph>> {
        Self::parse(xml_content).and_then(|blocks| {
            let mut paragraphs = Vec::new();
            paragraphs
                .try_reserve_exact(blocks.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT paragraph projection",
                    source,
                })?;
            for block in blocks {
                if let Some(paragraph) = block.into_paragraph() {
                    paragraphs.push(paragraph);
                }
            }
            Ok(paragraphs)
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
        Self::parse(xml_content).and_then(|blocks| {
            let mut headings = Vec::new();
            headings
                .try_reserve_exact(blocks.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT heading projection",
                    source,
                })?;
            for block in blocks {
                if let Some(heading) = block.into_heading() {
                    headings.push(heading);
                }
            }
            Ok(headings)
        })
    }

    /// Extract visible text from all decoded blocks, preserving block breaks.
    ///
    /// Namespace-aware parsing includes blocks nested in lists, sections, and
    /// text boxes. Tracked-change definitions, note bodies, and ruby
    /// pronunciation runs remain excluded by the decoder.
    ///
    /// The decoder validates every block's attributes exactly like the
    /// retained-element parsers but builds no `Element` tree; only the
    /// per-block text is retained (see `parse_text_block_texts`).
    pub fn extract_text(xml_content: &str) -> Result<String> {
        let mut texts = parse_text_block_texts(xml_content)?.into_iter();
        let Some(mut output) = texts.next() else {
            return Ok(String::new());
        };
        for text in texts {
            output
                .try_reserve(1usize.saturating_add(text.len()))
                .map_err(|source| Error::Allocation {
                    resource: "ODT full-text projection",
                    source,
                })?;
            output.push('\n');
            output.push_str(&text);
        }
        Ok(output)
    }

    /// Write visible text blocks directly to a bounded sequential sink.
    pub(crate) fn write_text_to<W: Write + ?Sized>(
        xml_content: &str,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<Error>> {
        let mut writer = SequentialTextWriter::new(output, options);
        Self::write_text_to_with_writer(xml_content, &mut writer)?;
        Ok(writer.finish())
    }

    /// Feed visible text blocks into an already-progressing sink writer.
    pub(crate) fn write_text_to_with_writer<'options, 'output, W: Write + ?Sized>(
        xml_content: &str,
        writer: &mut SequentialTextWriter<'options, 'output, W>,
    ) -> std::result::Result<(), TextOutputError<Error>> {
        super::write_text_blocks_to_writer(xml_content, writer)
    }
}
