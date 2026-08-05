//! Strict structural Word 97+ field-table (`Plcfld`) parsing.
//!
//! Implements [MS-DOC] sections 2.8.25, 2.9.88 through 2.9.90, and 2.9.110
//! for all seven field-bearing document stories. Field instructions and results
//! remain inert text ranges; this module never evaluates or executes them.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::*;
pub use package::{FieldStoryTable, FieldsTable};
