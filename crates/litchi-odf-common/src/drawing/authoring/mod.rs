//! Neutral drawing-frame authoring primitives.
//!
//! This layer owns the values that are common to ODF drawing hosts. XML
//! element trees and document insertion policy remain in the owning family
//! crates.

mod model;

pub use model::{Anchor, Frame, Length, validate_text_box};
