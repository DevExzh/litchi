//! Migration-only PPTX shape helpers.
//!
//! The public borrowed scene and data-bearing shape enum live in
//! [`crate::pptx::shape`]. Table and rich-text parsing remain here temporarily
//! until their shared DrawingML owners absorb them.

pub mod table;
pub mod textframe;
