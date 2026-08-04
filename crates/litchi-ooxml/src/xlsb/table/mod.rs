//! Table (ListObject) stream parsing (MS-XLSB 2.1.7.51).
//!
//! Parses a `tables/table*.bin` part into a typed [`Table`] model
//! covering the table identity and range, header/total row and insert-row
//! metadata, differential-formatting references, table columns with their
//! total-row functions and formulas, the applied table style, and alternate
//! text.
//!
//! Everything parsed here is an inert data snapshot: relationship
//! identifiers, external connection identifiers, differential-formatting
//! identifiers, and formula token streams are stored verbatim and are never
//! dereferenced, contacted, or evaluated. Display names are preserved exactly
//! as stored — no Excel display-name validation or sanitization is applied.
//!
//! The parser tolerates unknown record types (they are skipped) and rejects
//! structurally malformed payloads for the records it understands.

mod model;
mod parse;
pub(crate) mod write;

pub use model::*;
pub use parse::{parse_table_part, parse_table_part_rel_ids};

#[cfg(test)]
mod tests;
