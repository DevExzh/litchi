//! Hierarchical XML parser for ODF data-pilot declarations.
//!
//! The parser facade remains a single crate-private entry point while each
//! semantic part of a data-pilot declaration owns its XML grammar in a
//! focused module.

mod field;
mod grouping;
mod level;
mod source;
mod support;
mod table;
mod tables;

pub(crate) use tables::parse_data_pilot_tables;
