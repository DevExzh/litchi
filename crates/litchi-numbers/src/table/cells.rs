//! Presence-preserving semantic Numbers cell reads.
//!
//! Start with [`crate::Package::table_cell`] for one checked coordinate, or
//! [`crate::Package::table_cells`] for a checked dense range. Both are
//! selector-first: select a sheet with [`crate::SheetSelector`], then a table
//! on that sheet with [`crate::TableSelector`]. Name selectors use exact
//! visible names and reject malformed duplicate matches instead of silently
//! choosing one. Index selectors are checked zero-based source positions.
//!
//! [`State`](crate::table::cells::State) preserves the requested coordinate
//! and its native materialization state.
//! [`Storage::Missing`](crate::table::cells::Storage::Missing) is distinct
//! from [`Storage::Stored`](crate::table::cells::Storage::Stored) containing
//! [`Value::Empty`](crate::cell::Value::Empty), so a caller never needs to
//! infer presence from an empty-looking value. Stored values may also be the
//! semantic formula or error values already represented by
//! [`Value`](crate::cell::Value); this read slice neither evaluates nor
//! mutates them.
//!
//! A range is half-open and returned in row-major order. It is deliberately
//! dense: every addressed coordinate yields a
//! [`State`](crate::table::cells::State), including missing cells. Bounds are
//! checked against the selected table before allocation, and the requested
//! element count is bounded by the package's semantic materialized-cell
//! limit. An
//! over-limit, out-of-bounds, ambiguous, or malformed source fails as a typed
//! read error without returning a partial range.
//!
//! The reader is grounded in an Apple-authored Numbers workbook with
//! materialized text and number cells. Native identifiers, package members,
//! protobuf/wire records, and source bytes remain private to the package
//! reader. Cell editing, patching, cache changes, preview invalidation, and
//! native publication are intentionally outside this read-only surface.

use std::fmt;

use crate::cell::Value;

use super::CellPosition;

pub use crate::package::table_cells::{Error, LimitKind, Path};

/// Physical presence and semantic value at one requested coordinate.
#[derive(Clone, PartialEq)]
#[non_exhaustive]
pub enum Storage {
    /// No native cell is materialized at the coordinate.
    Missing,
    /// A native cell is materialized, including an explicitly stored
    /// [`Value::Empty`], formula, or error value.
    Stored(Value),
}

impl Storage {
    /// Return whether no native cell is materialized.
    #[must_use]
    pub const fn is_missing(&self) -> bool {
        matches!(self, Self::Missing)
    }

    /// Borrow the stored semantic value, if one is materialized.
    #[must_use]
    pub const fn value(&self) -> Option<&Value> {
        match self {
            Self::Missing => None,
            Self::Stored(value) => Some(value),
        }
    }
}

impl fmt::Debug for Storage {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Missing => formatter.write_str("Storage::Missing"),
            Self::Stored(value) => {
                write!(formatter, "Storage::Stored({})", value.cell_type().name())
            },
        }
    }
}

/// One presence-preserving semantic cell result.
#[derive(Clone, PartialEq)]
pub struct State {
    position: CellPosition,
    storage: Storage,
}

impl State {
    /// Return the requested semantic coordinate.
    #[must_use]
    pub const fn position(&self) -> CellPosition {
        self.position
    }

    /// Borrow the coordinate's presence-preserving storage state.
    #[must_use]
    pub const fn storage(&self) -> &Storage {
        &self.storage
    }

    pub(crate) const fn new(position: CellPosition, storage: Storage) -> Self {
        Self { position, storage }
    }
}

impl fmt::Debug for State {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("State")
            .field("position", &self.position)
            .field("storage", &self.storage)
            .finish()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn state_preserves_presence_and_redacts_authored_content() {
        assert_send_sync::<Storage>();
        assert_send_sync::<State>();

        let position = CellPosition::new(2, 1);
        let missing = State::new(position, Storage::Missing);
        assert_eq!(missing.position(), position);
        assert!(missing.storage().is_missing());
        assert_eq!(missing.storage().value(), None);

        let stored = State::new(
            position,
            Storage::Stored(Value::Text("private cell text".to_owned())),
        );
        assert!(matches!(stored.storage(), Storage::Stored(Value::Text(_))));
        assert!(!format!("{stored:?}").contains("private cell text"));
        assert_eq!(
            stored.storage().value().map(Value::cell_type),
            Some(crate::cell::Type::Text)
        );
    }
}
