//! Layered SpreadsheetML worksheet header/footer metadata.
//!
//! Semantic values live in the model module, bounded XML decoding in the
//! codec module, and regression coverage in the tests module. The public
//! module remains the ergonomic litchi_xlsx::header_footer facade.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::parse_worksheet_header_footer;
pub use model::{SectionKind, Settings, Text};
