//! Semantic cell-watch values and bounded opaque records.

use super::validation;
use crate::package::error::Result;

/// The maximum zero-based worksheet row accepted by `[MS-XLSB]`.
pub(super) const MAX_ROW: u32 = 1_048_575;
/// The maximum zero-based worksheet column accepted by `[MS-XLSB]`.
pub(super) const MAX_COLUMN: u32 = 16_383;

/// A checked zero-based worksheet cell reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Reference {
    row: u32,
    column: u32,
}

impl Reference {
    /// Construct a reference using the `UncheckedRw`/`UncheckedCol` bounds.
    pub fn new(row: u32, column: u32) -> Result<Self> {
        let value = Self { row, column };
        validation::reference(&value)?;
        Ok(value)
    }

    /// Return the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Return the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// One `BrtCellWatch` reference in worksheet stream order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Watch {
    reference: Reference,
}

impl Watch {
    /// Construct a watch for one checked worksheet cell.
    pub fn new(row: u32, column: u32) -> Result<Self> {
        Ok(Self {
            reference: Reference::new(row, column)?,
        })
    }

    /// Construct a watch from an existing checked reference.
    #[must_use]
    pub const fn from_reference(reference: Reference) -> Self {
        Self { reference }
    }

    /// Return the watched cell reference.
    #[must_use]
    pub const fn reference(self) -> Reference {
        self.reference
    }

    /// Return the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.reference.row()
    }

    /// Return the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.reference.column()
    }
}

/// A checked, ordered collection of cell watches.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Watches {
    items: Vec<Watch>,
}

impl Watches {
    /// Construct and validate a collection in source order.
    pub fn new(items: Vec<Watch>) -> Result<Self> {
        let value = Self { items };
        validation::watches(&value)?;
        Ok(value)
    }

    /// Construct an empty watch collection.
    #[must_use]
    pub const fn empty() -> Self {
        Self { items: Vec::new() }
    }

    /// Return the ordered watch slice.
    #[must_use]
    pub fn as_slice(&self) -> &[Watch] {
        &self.items
    }

    /// Return the number of watches.
    #[must_use]
    pub fn len(&self) -> usize {
        self.items.len()
    }

    /// Whether the collection has no watches.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// Iterate in BIFF12 source order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Watch> {
        self.items.iter()
    }

    pub(crate) fn from_validated(items: Vec<Watch>) -> Self {
        Self { items }
    }

    pub(crate) fn items_mut(&mut self) -> &mut Vec<Watch> {
        &mut self.items
    }
}

impl AsRef<[Watch]> for Watches {
    fn as_ref(&self) -> &[Watch] {
        self.as_slice()
    }
}

/// One record not interpreted by the cell-watch owner.
///
/// Opaque records are retained in their original position inside the
/// `BrtBeginCellWatches` collection. Their payload includes reserved or
/// producer-extension bits, so editing typed watches never normalizes them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownRecord {
    kind: u16,
    payload: Box<[u8]>,
}

impl UnknownRecord {
    /// Return the numeric BIFF12 record kind.
    #[must_use]
    pub const fn kind(&self) -> u16 {
        self.kind
    }

    /// Lend the bounded opaque payload.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    pub(crate) fn new(kind: u16, payload: &[u8]) -> Result<Self> {
        let value = Self {
            kind,
            payload: payload.to_vec().into_boxed_slice(),
        };
        validation::unknown(&value)?;
        Ok(value)
    }
}

/// Internal collection item preserving the exact relative order of typed and
/// opaque records.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Item {
    Watch(Watch),
    Unknown(usize),
}
