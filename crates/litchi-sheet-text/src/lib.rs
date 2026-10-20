//! Delimited, SYLK, DIF and fixed-width text workbooks.
//!
//! These text formats carry rows of values, not Office documents, so they
//! live beside the format crates rather than in the unified `File` model:
//!
//! | Module | Format | Reader / writer |
//! |---|---|---|
//! | [`delimited`] | CSV, TSV, PRN (`;`), pipe or any single-byte delimiter | [`delimited::read`] / [`delimited::write`], the one-sheet [`Workbook`], and the shared [`delimited::Tokenizer`] |
//! | [`sylk`] | SYLK (`.slk`) | [`sylk::read`] / [`sylk::write`] |
//! | [`dif`] | DIF (`.dif`) | [`dif::read`] / [`dif::write`] |
//! | [`fixed_width`] | space-aligned columns (`.prn`) | [`fixed_width::read`] / [`fixed_width::write`] |
//!
//! Every reader decodes UTF-8 or UTF-16 by byte-order mark ([`Encoding`]),
//! bounds the bytes it reads ([`SourceLimits`]) and the cells it
//! materializes ([`SemanticLimits`]), and reports content-free typed
//! [`Error`]s. Values are the shared [`CellValue`]s; the readers infer
//! numbers and booleans with [`CellValue::infer_from_str`].
//!
//! The unified `File::import` of `litchi-ops` (CSV/TSV into a worksheet of
//! an Office workbook) tokenizes with [`delimited::Tokenizer`] and
//! decodes with [`Encoding`]; it infers its own cell types.
//!
//! ```
//! use litchi_core::sheet::CellValue;
//! use litchi_sheet_text::{Workbook, sylk};
//!
//! let workbook = Workbook::from_bytes(b"name,score\nAda,36\n")?;
//! let mut slk = Vec::new();
//! sylk::write(workbook.rows(), &mut slk, &sylk::Options::default())?;
//! let rows = sylk::read(&slk[..], &sylk::Options::default())?;
//! assert_eq!(rows[1], [CellValue::String("Ada".into()), CellValue::Int(36)]);
//! # Ok::<(), litchi_sheet_text::Error>(())
//! ```
//!
//! [`CellValue`]: litchi_core::sheet::CellValue
//! [`CellValue::infer_from_str`]: litchi_core::sheet::CellValue::infer_from_str

#![forbid(unsafe_code)]

mod common;
pub mod delimited;
pub mod dif;
mod encoding;
mod error;
pub mod fixed_width;
mod limits;
pub mod sylk;
mod workbook;
mod worksheet;

pub use encoding::Encoding;
pub use error::{Error, Format, OptionKind, Result};
pub use limits::{
    DEFAULT_CELLS, DEFAULT_INPUT_BYTES, LimitKind, SemanticLimitKind, SemanticLimits,
    SemanticLimitsBuilder, SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use workbook::Workbook;
pub use worksheet::{Cell, Cells, Rows, Worksheet, Worksheets};

#[cfg(test)]
mod tests;
