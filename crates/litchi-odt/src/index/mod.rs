//! Inert semantic access to generated `OpenDocument` text indexes.
//!
//! The owner keeps typed index snapshots in model, namespace-aware XML
//! decoding in codec, and authored index construction/mutation in writing.
//! The module facade intentionally preserves the existing contextual API.

mod codec;
mod model;
mod writing;

#[cfg(test)]
mod tests;

pub use model::{TextIndex, TextIndexAttribute, TextIndexContent, TextIndexElement, TextIndexKind};

pub(crate) use codec::{expanded_attributes, parse_text_indexes};

pub use writing::{
    AlphabeticalIndexSource, BibliographyIndexSource, IllustrationIndexSource, ObjectIndexSource,
    TableOfContentsSource, TextAlphabeticalIndexEntryTemplate, TextAlphabeticalIndexLevel,
    TextBibliographyEntryTemplate, TextBibliographyEntryToken, TextBibliographyType, TextIndexBody,
    TextIndexBodyParagraph, TextIndexBodyTitle, TextIndexCaptionSequenceFormat,
    TextIndexChapterDisplay, TextIndexEntryTemplate, TextIndexEntryToken, TextIndexScope,
    TextIndexSimpleEntryTemplate, TextIndexSourceStyles, TextIndexTabStop, TextIndexTitleTemplate,
    UserIndexSource, insert_text_index_xml, remove_text_index_xml, replace_text_index_xml,
};
