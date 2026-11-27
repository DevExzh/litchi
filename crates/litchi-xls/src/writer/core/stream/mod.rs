//! Workbook-stream generation facade.
//!
//! The stream owner keeps the BIFF workbook-stream coordinator separate from
//! its input validation and small semantic planning helpers.  Callers still
//! use the same `stream::generate_workbook_stream` boundary as before.

mod codec;
mod semantic;
mod shared_strings;
mod validation;

pub(crate) use self::codec::{generate_workbook_stream, write_pivot_cache, write_pivot_table_view};
pub(crate) use self::semantic::WorkbookStreams;
pub(crate) use self::shared_strings::SharedStringTable;
