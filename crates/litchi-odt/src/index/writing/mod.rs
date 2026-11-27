//! Typed index authoring and byte-preserving content-XML mutation.
//!
//! The owner is intentionally split by responsibility: semantic authoring
//! inputs live in `semantic`, ODF XML construction/validation and mutation
//! live in `xml`, and the public `TextIndex` authoring facade lives in
//! `package`.

mod package;
mod semantic;
mod xml;

#[cfg(test)]
mod tests;

pub(super) const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
pub(super) const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_TEMPLATES: usize = 4_096;
pub(super) const MAX_TOKENS: usize = 65_536;
pub(super) const MAX_BODY_PARAGRAPHS: usize = 65_536;
pub(super) const MAX_DEPTH: usize = 4_096;

pub use semantic::{
    AlphabeticalIndexSource, BibliographyIndexSource, IllustrationIndexSource, ObjectIndexSource,
    TableOfContentsSource, TextAlphabeticalIndexEntryTemplate, TextAlphabeticalIndexLevel,
    TextBibliographyEntryTemplate, TextBibliographyEntryToken, TextBibliographyType, TextIndexBody,
    TextIndexBodyParagraph, TextIndexBodyTitle, TextIndexCaptionSequenceFormat,
    TextIndexChapterDisplay, TextIndexEntryTemplate, TextIndexEntryToken, TextIndexScope,
    TextIndexSimpleEntryTemplate, TextIndexSourceStyles, TextIndexTabStop, TextIndexTitleTemplate,
    UserIndexSource,
};

pub use xml::{insert_text_index_xml, remove_text_index_xml, replace_text_index_xml};
