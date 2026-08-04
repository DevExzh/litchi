//! Layered owner for package-independent PresentationML slide backgrounds.
//!
//! The semantic values live in [`model`], while the package-free `<p:bg>` XML
//! codec lives in [`codec`]. Relationship lookup and image-part ownership are
//! intentionally outside this owner and remain with the package host.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::{GradientStop, GradientType, PatternType, PictureStyle, SlideBackground};
