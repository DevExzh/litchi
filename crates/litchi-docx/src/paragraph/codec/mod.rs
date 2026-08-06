//! Semantic layers for the paragraph codec.

mod content;
mod editing;
mod paragraph_properties;
mod run;
mod run_properties;
mod runs;
mod text;
mod xml;

pub(crate) use text::extract_word_text;
