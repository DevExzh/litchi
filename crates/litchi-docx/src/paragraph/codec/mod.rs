//! Semantic layers for the paragraph codec.

mod content;
mod editing;
mod inlines;
mod paragraph_properties;
mod run;
mod run_contents;
mod run_properties;
mod runs;
mod text;
mod xml;

pub(crate) use text::extract_word_text;
pub(crate) use xml::is_fragment_word_name;
