//! XML codec layers for data-pilot declarations.

mod parser;
mod writer;
mod xml;

pub(crate) use parser::parse_data_pilot_tables;
pub(crate) use writer::{write_data_pilot_table_fragment, write_data_pilot_tables};
