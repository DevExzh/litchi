//! Semantic Feature11/Feature12 translation for worksheet tables.
//!
//! The facade keeps Feature11/12 encoding, feature parsing, and List12 style
//! updates as separate semantic owners while preserving the existing
//! `ListObject` methods used by the package codec.

mod encode;
mod list12;
mod parse;
