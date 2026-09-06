//! Archive-free Numbers image selectors and adjustment values.

use litchi_core::Position;

pub use litchi_iwa_common::shape::image::{
    Error, ImageAdjustment, ImageAdjustments, ImageEnhancement,
};

/// Selects one image by its zero-based position in a sheet's source order.
///
/// The selector contains only a checked semantic position. Native object
/// identifiers, component names, and protobuf payloads remain private to the
/// Numbers package adapter that resolves it.
#[allow(
    clippy::module_name_repetitions,
    reason = "ImageSelector keeps the selected Numbers domain explicit at the public boundary"
)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageSelector {
    /// Select the image at this zero-based source-order position.
    Index(Position),
}

impl ImageSelector {
    /// Create a selector from a zero-based source-order index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(Position::new(index))
    }

    /// Create a selector from a typed zero-based source-order position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Index(position)
    }

    /// Return the selected typed source-order position.
    #[must_use]
    pub const fn as_position(self) -> Position {
        match self {
            Self::Index(position) => position,
        }
    }

    /// Return the selected zero-based source-order index.
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.as_position().get()
    }
}

impl From<usize> for ImageSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for ImageSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}
