//! Archive-free Keynote slide-background values.

pub use litchi_iwa_common::shape::fill::{Angle, Gradient, Kind, Stop};

/// The semantic fill of a Keynote slide background.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
pub enum Background {
    /// No fill is applied.
    None,
    /// A single validated color fills the slide.
    Solid(litchi_iwa_common::color::Rgba),
    /// A validated native gradient fills the slide.
    Gradient(Gradient),
    /// A native fill unsupported by this semantic API.
    ///
    /// The exact native bytes remain private to the immutable package source.
    /// This marker is observable but cannot be authored or changed safely.
    Unsupported,
}
