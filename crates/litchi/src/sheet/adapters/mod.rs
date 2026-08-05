//! Local bridges from standalone spreadsheet models to the umbrella traits.

#[cfg(feature = "ooxml")]
mod xlsx;

#[cfg(feature = "ooxml")]
pub(super) use xlsx::Workbook;
