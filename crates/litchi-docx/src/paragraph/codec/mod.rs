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

pub(crate) use text::{
    extract_word_text, semantic_text_raw_xml_limit, write_text_to,
    write_text_to_with_operation_check,
};
pub(crate) use xml::is_fragment_word_name;
