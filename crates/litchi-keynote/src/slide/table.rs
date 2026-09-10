//! Archive-free semantic values for a Keynote slide table.
//!
//! The native IWA adapter owns object discovery, protobuf decoding, and
//! package mutation. These focused modules expose only the values callers use
//! to describe table formulas, sorting, names, and title settings.

use litchi_core::Position;

pub use self::number_format::CellPosition;

/// Effective appearance settings for a table's style graph.
pub mod appearance;
/// Archive-free values and comments returned by focused table reads.
pub mod cells;
/// Checked row and column point-size values.
pub mod dimension;
/// Formula values shared through the neutral iWork semantic model.
pub mod formula;
/// Lossless header, footer, freeze, and print-repetition settings.
pub mod headers;
/// Interactive lock state for a table.
pub mod lock;
/// Compact, archive-free merged-cell geometry and topology algebra.
pub mod merge;
/// Validated display names for a table.
pub mod name;
/// Decimal number display formats for existing table cells.
pub mod number_format;
/// Physical row-sort transaction values for a Keynote slide table.
pub mod physical_sort;
/// Sort values shared through the neutral iWork semantic model.
pub mod sort;
/// Lossless visibility and outline settings for a table title.
pub mod title;

/// Select one table by its checked zero-based position in slide z-order.
///
/// The selector deliberately carries no native object identifier. The
/// concrete package adapter resolves it through the selected slide's owned
/// drawable and z-order graph before a read or transaction is admitted.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableSelector(Position);

impl TableSelector {
    /// Create a selector from a typed zero-based position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self(position)
    }

    /// Create a selector from a zero-based source index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::position(Position::new(index))
    }

    /// Return the checked zero-based position.
    #[must_use]
    pub const fn as_position(self) -> Position {
        self.0
    }

    /// Return the zero-based source index.
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.0.get()
    }
}

impl From<Position> for TableSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

impl From<usize> for TableSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}
