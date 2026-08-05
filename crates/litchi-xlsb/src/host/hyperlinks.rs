//! Compatibility path for typed XLSB hyperlink records.
//!
//! The canonical `BrtHLink` model and codec are owned by [`litchi_xlsb`].
//! This module preserves the host path while worksheet and OPC orchestration
//! remains in the owning `litchi_xlsb` layers.

pub use crate::hyperlinks::*;
