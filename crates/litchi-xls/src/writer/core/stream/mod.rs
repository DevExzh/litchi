//! Workbook-stream generation facade.
//!
//! The stream owner keeps the BIFF workbook-stream coordinator separate from
//! its input validation and small semantic planning helpers.  Callers still
//! use the same `stream::generate_workbook_stream` boundary as before.

mod codec;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use self::codec::generate_workbook_stream;
pub(crate) use self::semantic::WorkbookStreams;
