//! Semantic values for the BIFF8 `ShrFmla` sequence.

use std::sync::Arc;

use crate::{Error, Result};

/// A checked zero-based BIFF8 cell coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Ord, PartialOrd)]
pub struct Cell {
    row: u16,
    col: u8,
}

impl Cell {
    /// Construct a coordinate from already checked BIFF8 fields.
    pub const fn new(row: u16, col: u8) -> Self {
        Self { row, col }
    }

    /// Construct a coordinate from the wider indices used by writer APIs.
    pub fn try_new(row: u32, col: u16) -> Result<Self> {
        let row = u16::try_from(row).map_err(|_| {
            Error::InvalidCellReference(format!(
                "shared-formula row {row} is outside the BIFF8 grid"
            ))
        })?;
        let col = u8::try_from(col).map_err(|_| {
            Error::InvalidCellReference(format!(
                "shared-formula column {col} is outside the BIFF8 grid"
            ))
        })?;
        Ok(Self { row, col })
    }

    /// Zero-based row.
    pub const fn row(self) -> u16 {
        self.row
    }

    /// Zero-based column, bounded to BIFF8's `A..IV` grid.
    pub const fn col(self) -> u8 {
        self.col
    }
}

/// An ordered inclusive BIFF8 cell range.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Range {
    first: Cell,
    last: Cell,
}

impl Range {
    /// Construct an ordered range from its upper-left and lower-right cells.
    pub fn new(first: Cell, last: Cell) -> Result<Self> {
        if first.row > last.row || first.col > last.col {
            return Err(Error::InvalidCellReference(
                "shared-formula range endpoints are reversed".to_string(),
            ));
        }
        Ok(Self { first, last })
    }

    /// Construct an ordered range from zero-based writer indices.
    pub fn try_new(first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> Result<Self> {
        Self::new(
            Cell::try_new(first_row, first_col)?,
            Cell::try_new(last_row, last_col)?,
        )
    }

    /// Upper-left endpoint.
    pub const fn first(self) -> Cell {
        self.first
    }

    /// Lower-right endpoint.
    pub const fn last(self) -> Cell {
        self.last
    }

    /// Whether a coordinate is inside this inclusive range.
    pub fn contains(self, cell: Cell) -> bool {
        self.first.row <= cell.row
            && cell.row <= self.last.row
            && self.first.col <= cell.col
            && cell.col <= self.last.col
    }
}

/// The semantic owner of one BIFF8 `ShrFmla` record.
///
/// The shared token stream and participant list are reference-counted so a
/// worksheet can stage the same owner on every participating Formula cell
/// without copying either collection.  The owner is still immutable after
/// construction; edits produce a new checked value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Owner {
    range: Range,
    anchor: Cell,
    tokens: Arc<[u8]>,
    participants: Arc<[Cell]>,
}

impl Owner {
    /// Construct an owner with only its anchor participating.
    pub fn new(range: Range, anchor: Cell, tokens: &[u8]) -> Result<Self> {
        let owner = Self {
            range,
            anchor,
            tokens: Arc::from(tokens),
            participants: Arc::from([anchor]),
        };
        owner.validate()?;
        Ok(owner)
    }

    /// Replace the participating-cell set.
    ///
    /// A non-empty set must contain the anchor exactly once. Passing an empty
    /// set is rejected so that `cUse` cannot silently disagree with the
    /// Formula records the caller intends to emit.
    pub fn with_participants(mut self, participants: &[Cell]) -> Result<Self> {
        if participants.is_empty() {
            return Err(Error::InvalidData(
                "shared-formula participants cannot be empty".to_string(),
            ));
        }
        let mut participants = participants.to_vec();
        participants.sort_unstable();
        if participants.windows(2).any(|cells| cells[0] == cells[1]) {
            return Err(Error::InvalidData(
                "shared-formula participants contain a duplicate cell".to_string(),
            ));
        }
        if !participants.contains(&self.anchor) {
            return Err(Error::InvalidData(
                "shared-formula participants do not include the anchor".to_string(),
            ));
        }
        if participants.iter().any(|cell| !self.range.contains(*cell)) {
            return Err(Error::InvalidData(
                "shared-formula participant is outside its RefU range".to_string(),
            ));
        }
        if participants.iter().any(|cell| *cell < self.anchor) {
            return Err(Error::InvalidData(
                "shared-formula participant precedes its anchor in worksheet order".to_string(),
            ));
        }
        self.participants = Arc::from(participants);
        self.validate()?;
        Ok(self)
    }

    /// Shared-formula range.
    pub const fn range(&self) -> Range {
        self.range
    }

    /// Cell whose Formula record owns the following ShrFmla record.
    pub const fn anchor(&self) -> Cell {
        self.anchor
    }

    /// The shared parsed formula (`rgce`) without the Formula-cell `PtgExp`.
    pub fn tokens(&self) -> &[u8] {
        &self.tokens
    }

    /// Formula cells declared as users of this owner, in row-major order.
    pub fn participants(&self) -> &[Cell] {
        &self.participants
    }

    /// Number of Formula records using this shared formula.
    pub fn count(&self) -> u8 {
        u8::try_from(self.participants.len()).unwrap_or(u8::MAX)
    }

    /// Number of Formula records using this shared formula as `cUse`.
    pub fn c_use(&self) -> Result<u8> {
        u8::try_from(self.participants.len()).map_err(|_| {
            Error::InvalidData("shared-formula participant count exceeds BIFF8 cUse".to_string())
        })
    }

    /// Whether this owner explicitly contains a Formula cell.
    pub fn is_participant(&self, cell: Cell) -> bool {
        self.participants.binary_search(&cell).is_ok()
    }

    /// Return the exact standalone Formula-record `PtgExp` for this owner.
    pub(crate) const fn anchor_tokens(&self) -> [u8; 5] {
        let row = self.anchor.row.to_le_bytes();
        [0x01, row[0], row[1], self.anchor.col, 0]
    }

    /// Validate that a Formula cell is covered by this shared formula.
    pub(crate) fn validate_cell(&self, row: u16, col: u16) -> Result<Cell> {
        let cell = Cell::try_new(u32::from(row), col)?;
        if !self.range.contains(cell) {
            return Err(Error::InvalidData(format!(
                "Formula cell ({row}, {col}) is outside its shared-formula RefU range"
            )));
        }
        if !self.is_participant(cell) {
            return Err(Error::InvalidData(format!(
                "Formula cell ({row}, {col}) is not a participant of its shared formula"
            )));
        }
        Ok(cell)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        super::validation::validate(self)
    }
}
