//! Archive-free Pages drawable-order values and selectors.
//!
//! A [`BodyDrawableHandle`] is an opaque capability issued by one immutable
//! [`crate::Package`] snapshot. It carries no public native identifier; the
//! package adapter uses the private source binding and identity only after a
//! caller has selected a drawable semantically. Handles from another source
//! snapshot are rejected by the package transaction rather than being
//! interpreted against the current source.

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;

/// A native Arrange command for one existing body drawable.
///
/// Pages stores drawables from the back-most layer to the front-most layer.
/// These commands therefore have the same direction as the visible Arrange
/// controls and never create, remove, or otherwise mutate a drawable.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DrawableLayerMove {
    /// Move the selected drawable to the back-most layer.
    ToBack,
    /// Move the selected drawable one layer toward the back.
    Backward,
    /// Move the selected drawable one layer toward the front.
    Forward,
    /// Move the selected drawable to the front-most layer.
    ToFront,
}

/// An opaque, source-bound identity for one existing Pages body drawable.
///
/// Handles are returned by [`crate::Package::body_drawable_order`] and can be
/// supplied to the corresponding order transaction. They cannot be
/// constructed from a native object identifier, and their debug output does
/// not reveal the private identity or retained source bytes.
pub struct BodyDrawableHandle {
    pub(crate) source: Arc<[u8]>,
    pub(crate) identity: NonZeroU64,
    pub(crate) position: Position,
}

impl Clone for BodyDrawableHandle {
    fn clone(&self) -> Self {
        Self {
            source: Arc::clone(&self.source),
            identity: self.identity,
            position: self.position,
        }
    }
}

impl fmt::Debug for BodyDrawableHandle {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyDrawableHandle")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl PartialEq for BodyDrawableHandle {
    fn eq(&self, other: &Self) -> bool {
        self.identity == other.identity
            && (Arc::ptr_eq(&self.source, &other.source)
                || self.source.as_ref() == other.source.as_ref())
    }
}

impl Eq for BodyDrawableHandle {}

impl BodyDrawableHandle {
    pub(crate) fn new(source: Arc<[u8]>, identity: NonZeroU64, position: Position) -> Self {
        Self {
            source,
            identity,
            position,
        }
    }

    /// Return the checked zero-based position at which this handle was
    /// issued.
    ///
    /// The position is a source snapshot fact. After a successful reorder,
    /// obtain fresh handles from the returned package rather than reusing the
    /// old position as a new identity.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }
}

/// Selects one existing body drawable by a checked position or an opaque
/// source-bound handle.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyDrawableSelector {
    /// Select the drawable currently at this zero-based body order position.
    Position(Position),
    /// Select the drawable represented by a handle issued by the same source
    /// snapshot.
    Handle(BodyDrawableHandle),
}

impl BodyDrawableSelector {
    /// Create a checked semantic position selector.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Create a zero-based semantic position selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::position(Position::new(index))
    }

    /// Create a selector from an opaque source-bound handle.
    #[must_use]
    pub fn handle(handle: BodyDrawableHandle) -> Self {
        Self::Handle(handle)
    }

    /// Return the selected checked position, if this is a position selector.
    #[must_use]
    pub const fn as_position(&self) -> Option<Position> {
        match self {
            Self::Position(position) => Some(*position),
            Self::Handle(_) => None,
        }
    }
}

impl From<Position> for BodyDrawableSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

impl From<usize> for BodyDrawableSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<BodyDrawableHandle> for BodyDrawableSelector {
    fn from(handle: BodyDrawableHandle) -> Self {
        Self::handle(handle)
    }
}

impl From<&BodyDrawableHandle> for BodyDrawableSelector {
    fn from(handle: &BodyDrawableHandle) -> Self {
        Self::handle(handle.clone())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn handles_are_debug_redacted_and_source_bound() {
        let first_source: Arc<[u8]> = Arc::from([1_u8, 2, 3]);
        let same_source: Arc<[u8]> = Arc::from([1_u8, 2, 3]);
        let first = BodyDrawableHandle::new(
            Arc::clone(&first_source),
            NonZeroU64::new(7).unwrap(),
            Position::new(2),
        );
        let alias =
            BodyDrawableHandle::new(same_source, NonZeroU64::new(7).unwrap(), Position::new(0));
        assert_eq!(first, alias);
        assert_eq!(first.position(), Position::new(2));
        let debug = format!("{first:?}");
        assert!(debug.contains("BodyDrawableHandle"));
        assert!(!debug.contains("7"));
        assert!(!debug.contains("1, 2, 3"));
    }

    #[test]
    fn selectors_accept_checked_positions_and_owned_handles() {
        let position = BodyDrawableSelector::index(3);
        assert_eq!(position.as_position(), Some(Position::new(3)));
        let handle = BodyDrawableHandle::new(
            Arc::from([9_u8]),
            NonZeroU64::new(1).unwrap(),
            Position::new(0),
        );
        assert_eq!(
            BodyDrawableSelector::from(&handle),
            BodyDrawableSelector::Handle(handle)
        );
    }
}
