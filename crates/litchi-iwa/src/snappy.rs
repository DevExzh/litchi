//! Re-export of the bounded physical IWA Snappy codec.
//!
//! Framing, resource limits, and compression live in `litchi-iwa-core` so the
//! physical codec has one implementation and one test owner. The application
//! crate keeps this narrow module only as its format-layer import boundary.

pub use litchi_iwa_core::{SnappyLimits, SnappyStream};
