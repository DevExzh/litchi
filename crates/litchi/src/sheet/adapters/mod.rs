//! Local bridges from standalone spreadsheet models to the umbrella traits.

#[cfg(feature = "xlsx")]
mod xlsx;

#[cfg(feature = "xlsb")]
mod xlsb;

#[cfg(feature = "xlsx")]
pub(super) use xlsx::{Workbook, boxed_xlsx_error};

#[cfg(feature = "xlsb")]
pub(super) use xlsb::Workbook as XlsbWorkbook;
