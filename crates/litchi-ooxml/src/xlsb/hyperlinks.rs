//! Compatibility path for typed XLSB hyperlink records.
//!
//! The canonical `BrtHLink` model and codec are owned by [`litchi_xlsb`].
//! This module preserves the historical host path while worksheet and OPC
//! orchestration remains in `litchi-ooxml`.

pub use litchi_xlsb::hyperlinks::*;
