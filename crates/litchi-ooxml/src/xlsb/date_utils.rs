//! Compatibility path for XLSB date and time utilities.
//!
//! The canonical implementation is owned by [`litchi_xlsb::date_utils`].
//! This module preserves the historical `litchi_ooxml::xlsb::date_utils` path.

pub use litchi_xlsb::date_utils::*;
