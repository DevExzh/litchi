//! Layered value types used by PresentationML writers.
//!
//! Image detection and text styling are independent concerns.  They are kept
//! in focused modules while this facade exposes the two small, ergonomic
//! values used by the slide and shape APIs.

mod image;
mod text;

pub use image::ImageFormat;
pub use text::TextFormat;
