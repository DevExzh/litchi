//! Chart sheet support for XLSB workbooks (MS-XLSB 2.1.7.7).
//!
//! A Chart Sheet part is the only chart-related BIFF12 binary stream in an
//! XLSB package: the embedded Chart part (2.1.7.5), the Chart Drawing part
//! (2.1.7.6), and the Drawings part (2.1.7.23) are standard DrawingML XML
//! parts, identical to XLSX. This module parses the binary chart sheet
//! stream into an inert typed model; the XML drawing inventory lives in
//! [`crate::package::drawing`] and the typed chart model is shared with the
//! other formats in [`litchi_drawingml::chart`].

mod model;
mod parse;

#[cfg(test)]
mod tests;

pub use model::{ChartSheet, Color, ColorType, PageSetup, Protection, State, View};
pub use parse::parse_chart_sheet_part;
