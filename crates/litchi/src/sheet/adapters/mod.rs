//! Local bridges from standalone spreadsheet models to the umbrella traits.

#[cfg(feature = "xlsx")]
mod xlsx;

#[cfg(feature = "xlsx")]
pub(super) use xlsx::Workbook;
