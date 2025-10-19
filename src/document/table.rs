//! Table implementation for Word documents.

use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A table in a Word document.
pub enum Table {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Table),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Table),
}

impl Table {
    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => t.row_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => t.row_count().map_err(Error::from),
        }
    }

    /// Get the rows in this table.
    pub fn rows(&self) -> Result<Vec<Row>> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(Row::Doc).collect())
            }
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(Row::Docx).collect())
            }
        }
    }
}

/// A table row in a Word document.
pub enum Row {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Row),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Row),
}

impl Row {
    /// Get the cells in this row.
    pub fn cells(&self) -> Result<Vec<Cell>> {
        match self {
            #[cfg(feature = "ole")]
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Doc).collect())
            }
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Docx).collect())
            }
        }
    }
}

/// A table cell in a Word document.
pub enum Cell {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Cell),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Cell),
}

impl Cell {
    /// Get the text content of the cell.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "ole")]
            Cell::Doc(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
        }
    }
}

