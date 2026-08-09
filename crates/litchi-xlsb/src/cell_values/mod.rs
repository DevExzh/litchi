//! Lossless edits of existing IEEE-754 worksheet cell values.
//!
//! This deliberately narrow owner covers only existing `BrtCellReal` records
//! ([MS-XLSB] §2.4.111). It neither creates cells nor rewrites RK, formula,
//! string, shared-string, rich-text, error, Boolean, or style records.
//! Consequently all records outside an edited eight-byte value field remain
//! byte-exact, including unknown record kinds and their framing.

pub mod workbook;
mod worksheet;

pub use worksheet::{Commit, Edit, Number, Patch, Reference, Snapshot};

pub use crate::package::error::{Error, Result};
