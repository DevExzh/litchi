//! Table implementation for Word documents.

use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A table in a Word document.
#[derive(Debug, Clone)]
pub enum Table {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Table),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Table),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Table<'static>),
}

impl Table {
    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => t.row_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => t.row_count().map_err(Error::from),
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => Ok(t.row_count()),
        }
    }

    /// Get the rows in this table.
    ///
    /// **Performance Note**: This method allocates and clones the entire row collection.
    /// For better performance when iterating, consider using `row_count()` and `row_at(index)`.
    pub fn rows(&self) -> Result<Vec<Row>> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(Row::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(Row::Docx).collect())
            },
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => {
                let rows = t.rows();
                Ok(rows.iter().cloned().map(Row::Rtf).collect())
            },
        }
    }

    /// Get a specific row by index without allocating a collection.
    ///
    /// This is more efficient than calling `rows()` and then indexing,
    /// as it avoids cloning the entire row collection.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn row_at(&self, index: usize) -> Result<Option<Row>> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.get(index).cloned().map(Row::Doc))
            },
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.get(index).cloned().map(Row::Docx))
            },
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => {
                let rows = t.rows();
                Ok(rows.get(index).cloned().map(Row::Rtf))
            },
        }
    }
}

/// A table row in a Word document.
#[derive(Debug, Clone)]
pub enum Row {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Row),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Row),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Row<'static>),
}

impl Row {
    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "ole")]
            Row::Doc(r) => r.cell_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => r.cell_count().map_err(Error::from),
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => Ok(r.cell_count()),
        }
    }

    /// Get the cells in this row.
    ///
    /// **Performance Note**: This method allocates and clones the entire cell collection.
    /// For better performance when iterating, consider using `cell_count()` and `cell_at(index)`.
    pub fn cells(&self) -> Result<Vec<Cell>> {
        match self {
            #[cfg(feature = "ole")]
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Docx).collect())
            },
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => {
                let cells = r.cells();
                Ok(cells.iter().cloned().map(Cell::Rtf).collect())
            },
        }
    }

    /// Get a specific cell by index without allocating a collection.
    ///
    /// This is more efficient than calling `cells()` and then indexing,
    /// as it avoids cloning the entire cell collection.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn cell_at(&self, index: usize) -> Result<Option<Cell>> {
        match self {
            #[cfg(feature = "ole")]
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.get(index).cloned().map(Cell::Doc))
            },
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.get(index).cloned().map(Cell::Docx))
            },
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => {
                let cells = r.cells();
                Ok(cells.get(index).cloned().map(Cell::Rtf))
            },
        }
    }
}

/// A table cell in a Word document.
#[derive(Debug, Clone)]
pub enum Cell {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Cell),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Cell),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Cell<'static>),
}

impl Cell {
    /// Get the text content of the cell.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "ole")]
            Cell::Doc(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "rtf")]
            Cell::Rtf(c) => Ok(c.text().to_string()),
        }
    }
}
