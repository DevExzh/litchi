//! Semantic ownership for ODS cell annotations.
//!
//! The worksheet coordinates in this module are logical, zero-based positions
//! selected by the producer-visible sheet name.  The XML codec and package
//! transaction layers retain the physical source spans privately; callers do
//! not need to know table runs, qualified names, or package paths.

use litchi_core::{Error, Result};
use litchi_odf_common::annotation::Annotation;

/// One logical ODS cell selected by exact sheet name and zero-based indices.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Cell {
    sheet: String,
    row: usize,
    column: usize,
}

impl Cell {
    /// Construct a checked logical cell selector.
    pub fn new(sheet: impl Into<String>, row: usize, column: usize) -> Result<Self> {
        let cell = Self {
            sheet: sheet.into(),
            row,
            column,
        };
        super::validation::validate_cell(&cell)?;
        Ok(cell)
    }

    /// Exact ODF `table:name` of the selected worksheet.
    pub fn sheet(&self) -> &str {
        &self.sheet
    }

    /// Zero-based logical row.
    pub const fn row(&self) -> usize {
        self.row
    }

    /// Zero-based logical column.
    pub const fn column(&self) -> usize {
        self.column
    }
}

/// One annotation attached to one logical cell, in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Entry {
    index: usize,
    cell: Cell,
    annotation: Annotation,
}

impl Entry {
    pub(crate) fn new(index: usize, cell: Cell, annotation: Annotation) -> Self {
        Self {
            index,
            cell,
            annotation,
        }
    }

    /// Source-order index within the annotation owner.
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Logical sheet/cell anchor.
    pub fn cell(&self) -> &Cell {
        &self.cell
    }

    /// Common ODF annotation value, including rich body elements and metadata.
    pub fn annotation(&self) -> &Annotation {
        &self.annotation
    }
}

/// A compact semantic selection accepted by annotation lookup APIs.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Selector<'a> {
    /// A source-order annotation index.
    Index(usize),
    /// An exact sheet name and logical cell coordinate.
    Cell {
        sheet: &'a str,
        row: usize,
        column: usize,
    },
}

pub(crate) fn not_found(cell: &Cell) -> Error {
    Error::InvalidFormat(format!(
        "ODS annotation cell '{}!R{}C{}' was not found",
        cell.sheet,
        cell.row.saturating_add(1),
        cell.column.saturating_add(1)
    ))
}
