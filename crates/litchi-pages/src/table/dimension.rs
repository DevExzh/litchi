//! Checked, archive-free row and column dimension values.
//!
//! The value types are format-neutral and deliberately contain no package,
//! archive, protobuf, or native identifier state. Their implementation is
//! shared by the concrete iWork owners.

pub use litchi_iwa_common::table::dimension::{Dimension, Error, Points, Size};
