//! Glyph mapping and OpenType subsetting.

mod allsorts;
pub mod mapping;

use crate::model::{FontData, FontError};

pub use allsorts::Allsorts;
pub use mapping::glyph_ids;

/// A pluggable font-program reducer.
pub trait Subsetter {
    fn subset(&self, font: &FontData, glyph_ids: &[u16]) -> Result<Vec<u8>, FontError>;
}
