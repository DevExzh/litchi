//! DOC TAP (Table Properties) writer facade.
//!
//! TAP authoring is organized into three contextual layers:
//!
//! - model contains typed table and revision snapshot inputs;
//! - codec emits the corresponding Word SPRM streams; and
//! - validation contains deterministic, typed representability errors.
//!
//! The facade keeps the concise writer API while hiding the implementation
//! seams, so callers do not need to know the binary layout of a TAP.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{generate_table_style_sprms, generate_table_style_sprms_with_conditionals};
pub use model::{
    TableBorders, TableCell, TableRevisionMark, TableRow, TapBuilder, create_simple_table,
};
pub use validation::TapBuildError;

pub(crate) use codec::generate_row_sprms;
