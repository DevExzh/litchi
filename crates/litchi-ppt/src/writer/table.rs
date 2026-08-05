//! Table authoring for PPT slides.
//!
//! A table in binary PPT is an OfficeArt group shape ([MS-PPT], [MS-ODRAW])
//! whose header SpContainer carries a TertiaryOpt (0xF122) record with the
//! `GroupTableProperties` (0x039F) flag set and a complex
//! `GroupTableRowProperties` (0x03A0) array of row heights. Each cell is a
//! rectangle SpContainer with a ChildAnchor inside the group coordinate
//! space and a ClientTextbox holding the cell text. This mirrors Apache
//! POI's `HSLFTable`/`HSLFTableCell` construction.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_ppt::writer::{Writer, Table};
//!
//! let mut writer = Writer::new();
//! let slide = writer.add_slide()?;
//!
//! let mut table = Table::new(2, 3)?;
//! table.set_cell_text(0, 0, "A1")?;
//! table.set_column_width(0, 120)?;
//! writer.add_table(slide, 50, 50, table)?;
//! writer.save("table.ppt")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use litchi_core::unit::{EMUS_PER_INCH, EMUS_PER_PT, PPT_MASTER_UNITS_PER_INCH, pt_to_emu_i32};
use zerocopy::IntoBytes;

use super::core::WriteError;
use super::escher::{
    ChildAnchor, Error, EscherBuilder, EscherProperty, EscherSpData, EscherSpgrData,
    PROPERTY_FLAG_COMPLEX, ShapeFlags, build_client_textbox, header_version, record_type,
    shape_type,
};

/// Default cell width in points (matches POI `HSLFTableCell.DEFAULT_WIDTH`).
pub const DEFAULT_COLUMN_WIDTH_PT: i32 = 100;
/// Default cell height in points (matches POI `HSLFTableCell.DEFAULT_HEIGHT`).
pub const DEFAULT_ROW_HEIGHT_PT: i32 = 40;

/// Upper bound for table rows or columns; keeps allocations and the u16
/// Escher property headers well within format limits.
pub const MAX_TABLE_DIMENSION: usize = 1000;

/// TextHeaderAtom text type used for table cells (4 = Other, same as
/// regular shapes written by this crate).
const CELL_TEXT_TYPE: u32 = 4;

/// OfficeArt TertiaryOpt record type (0xF122), which hosts the table
/// properties per [MS-ODRAW] and POI `EscherRecordTypes.USER_DEFINED`.
const RECORD_TYPE_TERTIARY_OPT: u16 = 0xF122;
/// OfficeArt ChildAnchor record type (0xF00F) per [MS-ODRAW].
const RECORD_TYPE_CHILD_ANCHOR: u16 = 0xF00F;
/// PowerPoint OfficeArtClientAnchor record type (0xF010).
const RECORD_TYPE_CLIENT_ANCHOR: u16 = 0xF010;
/// OfficeArt property id marking a group shape as a table ([MS-ODRAW]).
const PROP_GROUP_TABLE_PROPERTIES: u16 = 0x039F;
/// OfficeArt complex property id with one i32 row height per row.
const PROP_GROUP_TABLE_ROW_PROPERTIES: u16 = 0x03A0;

/// Convert EMUs to PPT master units (576 per inch) without sign loss.
fn emu_to_master_i32(emu: i32) -> i32 {
    ((i64::from(emu) * PPT_MASTER_UNITS_PER_INCH) / EMUS_PER_INCH) as i32
}

/// Convert EMUs to whole points.
fn emu_to_pt_i32(emu: i32) -> i32 {
    (i64::from(emu) * 72 / EMUS_PER_INCH) as i32
}

fn invalid(message: impl Into<String>) -> WriteError {
    WriteError::InvalidData(message.into())
}

/// A table to place on a slide: a grid of text cells with configurable
/// row heights and column widths.
#[derive(Debug, Clone)]
pub struct Table {
    /// Number of rows.
    rows: usize,
    /// Number of columns.
    columns: usize,
    /// Column widths in EMUs, one per column.
    column_widths: Vec<i32>,
    /// Row heights in EMUs, one per row.
    row_heights: Vec<i32>,
    /// Cell texts in row-major order (`rows * columns` entries).
    cells: Vec<String>,
}

impl Table {
    /// Create a table with `rows` rows and `columns` columns.
    ///
    /// Cells start empty; every column is [`DEFAULT_COLUMN_WIDTH_PT`] points
    /// wide and every row [`DEFAULT_ROW_HEIGHT_PT`] points high.
    pub fn new(rows: usize, columns: usize) -> Result<Self, WriteError> {
        if rows == 0 || columns == 0 {
            return Err(invalid("table requires at least one row and one column"));
        }
        if rows > MAX_TABLE_DIMENSION || columns > MAX_TABLE_DIMENSION {
            return Err(invalid(format!(
                "table dimensions are limited to {MAX_TABLE_DIMENSION} rows/columns"
            )));
        }
        Ok(Self {
            rows,
            columns,
            column_widths: vec![pt_to_emu_i32(DEFAULT_COLUMN_WIDTH_PT); columns],
            row_heights: vec![pt_to_emu_i32(DEFAULT_ROW_HEIGHT_PT); rows],
            cells: vec![String::new(); rows * columns],
        })
    }

    /// Number of rows.
    pub const fn rows(&self) -> usize {
        self.rows
    }

    /// Number of columns.
    pub const fn columns(&self) -> usize {
        self.columns
    }

    /// Text of the cell at (`row`, `column`), or `None` if out of range.
    pub fn cell(&self, row: usize, column: usize) -> Option<&str> {
        if row >= self.rows || column >= self.columns {
            return None;
        }
        self.cells
            .get(row * self.columns + column)
            .map(String::as_str)
    }

    /// Set the text of the cell at (`row`, `column`).
    pub fn set_cell_text(
        &mut self,
        row: usize,
        column: usize,
        text: impl Into<String>,
    ) -> Result<(), WriteError> {
        if row >= self.rows || column >= self.columns {
            return Err(invalid(format!(
                "cell ({row}, {column}) is outside a {}x{} table",
                self.rows, self.columns
            )));
        }
        self.cells[row * self.columns + column] = text.into();
        Ok(())
    }

    /// Width of a column in points, or `None` if out of range.
    pub fn column_width(&self, column: usize) -> Option<i32> {
        self.column_widths
            .get(column)
            .map(|width| emu_to_pt_i32(*width))
    }

    /// Set the width of a column in points.
    pub fn set_column_width(&mut self, column: usize, width_pt: i32) -> Result<(), WriteError> {
        let Some(slot) = self.column_widths.get_mut(column) else {
            return Err(invalid(format!(
                "column {column} is outside a {}-column table",
                self.columns
            )));
        };
        if width_pt <= 0 {
            return Err(invalid("column width must be positive"));
        }
        *slot = i32::try_from(i64::from(width_pt) * EMUS_PER_PT)
            .map_err(|_| invalid("column width exceeds the supported EMU range"))?;
        Ok(())
    }

    /// Height of a row in points, or `None` if out of range.
    pub fn row_height(&self, row: usize) -> Option<i32> {
        self.row_heights
            .get(row)
            .map(|height| emu_to_pt_i32(*height))
    }

    /// Set the height of a row in points.
    pub fn set_row_height(&mut self, row: usize, height_pt: i32) -> Result<(), WriteError> {
        let Some(slot) = self.row_heights.get_mut(row) else {
            return Err(invalid(format!(
                "row {row} is outside a {}-row table",
                self.rows
            )));
        };
        if height_pt <= 0 {
            return Err(invalid("row height must be positive"));
        }
        *slot = i32::try_from(i64::from(height_pt) * EMUS_PER_PT)
            .map_err(|_| invalid("row height exceeds the supported EMU range"))?;
        Ok(())
    }

    /// Total table width in EMUs.
    pub(crate) fn width_emu(&self) -> Option<i32> {
        self.column_widths
            .iter()
            .try_fold(0_i32, |total, width| total.checked_add(*width))
    }

    /// Total table height in EMUs.
    pub(crate) fn height_emu(&self) -> Option<i32> {
        self.row_heights
            .iter()
            .try_fold(0_i32, |total, height| total.checked_add(*height))
    }

    /// Number of OfficeArt shapes this table occupies in its drawing:
    /// one group shape plus one shape per cell.
    pub(crate) fn shape_count(&self) -> u32 {
        1 + (self.rows * self.columns) as u32
    }
}

/// A table with its top-left corner fixed on a slide.
#[derive(Debug, Clone)]
pub(crate) struct PositionedTable {
    /// X position in EMUs.
    pub x: i32,
    /// Y position in EMUs.
    pub y: i32,
    /// The table itself.
    pub table: Table,
}

/// Build the SpgrContainer for a positioned table.
///
/// `group_spid` is the shape id of the table group; cell shapes are numbered
/// consecutively starting at `group_spid + 1`. Layout:
///
/// ```text
/// SpgrContainer
/// ├── SpContainer (table group header)
/// │   ├── Spgr          child-coordinate bounding box of the grid
/// │   ├── Sp            NOT_PRIMITIVE, GROUP | HAVE_ANCHOR flags
/// │   ├── TertiaryOpt   GroupTableProperties=1, GroupTableRowProperties=[heights]
/// │   └── ClientAnchor  table position on the slide (master units)
/// └── SpContainer × rows*columns (cells, row-major)
///     ├── Sp            RECTANGLE, CHILD | HAVE_ANCHOR | HAVE_SPT flags
///     ├── ChildAnchor   cell rectangle in group child coordinates
///     └── ClientTextbox cell text
/// ```
pub(crate) fn build_table_spgr_container(
    placed: &PositionedTable,
    group_spid: u32,
) -> Result<Vec<u8>, Error> {
    let table = &placed.table;
    let width_emu = table.width_emu().ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "total table width exceeds the supported EMU range",
        )
    })?;
    let height_emu = table.height_emu().ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "total table height exceeds the supported EMU range",
        )
    })?;
    let width_master = emu_to_master_i32(width_emu);
    let height_master = emu_to_master_i32(height_emu);

    let mut group = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SPGR_CONTAINER);

    // Table group header SpContainer.
    let mut header = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut spgr = EscherBuilder::new(header_version::SPGR, 0, record_type::SPGR);
    spgr.add_data(
        EscherSpgrData {
            left: 0,
            top: 0,
            right: width_master,
            bottom: height_master,
        }
        .as_bytes(),
    );
    header.add_data(&spgr.build()?);

    let mut sp = EscherBuilder::new(
        header_version::SP,
        shape_type::NOT_PRIMITIVE,
        record_type::SP,
    );
    sp.add_data(
        EscherSpData::with_flags(group_spid, ShapeFlags::GROUP | ShapeFlags::HAVE_ANCHOR)
            .as_bytes(),
    );
    header.add_data(&sp.build()?);

    header.add_data(&build_table_properties(&table.row_heights)?);

    let left = emu_to_master_i32(placed.x);
    let top = emu_to_master_i32(placed.y);
    let right = left.checked_add(width_master).ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "table right edge exceeds the supported master-unit range",
        )
    })?;
    let bottom = top.checked_add(height_master).ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "table bottom edge exceeds the supported master-unit range",
        )
    })?;
    let top = i16::try_from(top).map_err(|_| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "table top edge exceeds the PPT short-anchor range",
        )
    })?;
    let left = i16::try_from(left).map_err(|_| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "table left edge exceeds the PPT short-anchor range",
        )
    })?;
    let right = i16::try_from(right).map_err(|_| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "table right edge exceeds the PPT short-anchor range",
        )
    })?;
    let bottom = i16::try_from(bottom).map_err(|_| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "table bottom edge exceeds the PPT short-anchor range",
        )
    })?;
    let mut anchor = EscherBuilder::new(header_version::SIMPLE, 0, RECORD_TYPE_CLIENT_ANCHOR);
    // PPT's eight-byte SmallRectStruct stores top, left, right, bottom.
    anchor.add_data(&top.to_le_bytes());
    anchor.add_data(&left.to_le_bytes());
    anchor.add_data(&right.to_le_bytes());
    anchor.add_data(&bottom.to_le_bytes());
    header.add_data(&anchor.build()?);

    group.add_data(&header.build()?);

    // Cell SpContainers in row-major order.
    let mut cell_spid = group_spid + 1;
    let mut cell_top = 0i32;
    for row in 0..table.rows {
        let row_height = emu_to_master_i32(table.row_heights[row]);
        let mut cell_left = 0i32;
        for column in 0..table.columns {
            let cell_width = emu_to_master_i32(table.column_widths[column]);
            let text = &table.cells[row * table.columns + column];
            group.add_data(&build_cell_container(
                cell_spid, text, cell_left, cell_top, cell_width, row_height,
            )?);
            cell_spid += 1;
            cell_left += cell_width;
        }
        cell_top += row_height;
    }

    group.build()
}

/// Build the TertiaryOpt record holding the table marker and row heights.
fn build_table_properties(row_heights_emu: &[i32]) -> Result<Vec<u8>, Error> {
    let rows = row_heights_emu.len();
    let rows_u16 = u16::try_from(rows).map_err(|_| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, "too many table rows")
    })?;

    // Complex property payload: 6-byte array header + one i32 per row.
    let mut row_data = Vec::with_capacity(6 + rows * 4);
    row_data.extend_from_slice(&rows_u16.to_le_bytes()); // nElems
    row_data.extend_from_slice(&rows_u16.to_le_bytes()); // nElemsAlloc
    row_data.extend_from_slice(&4u16.to_le_bytes()); // cbElem (i32)
    for height in row_heights_emu {
        row_data.extend_from_slice(&emu_to_master_i32(*height).to_le_bytes());
    }

    let properties = [
        EscherProperty::new(PROP_GROUP_TABLE_PROPERTIES, 1),
        EscherProperty::new(
            PROP_GROUP_TABLE_ROW_PROPERTIES | PROPERTY_FLAG_COMPLEX,
            row_data.len() as u32,
        ),
    ];

    let mut opt = EscherBuilder::new(
        header_version::OPT,
        properties.len() as u16,
        RECORD_TYPE_TERTIARY_OPT,
    );
    for property in &properties {
        opt.add_data(property.as_bytes());
    }
    // Complex property data follows the property table, in order.
    opt.add_data(&row_data);
    opt.build()
}

/// Build one cell SpContainer (rectangle shape + anchor + text).
fn build_cell_container(
    cell_spid: u32,
    text: &str,
    left: i32,
    top: i32,
    width: i32,
    height: i32,
) -> Result<Vec<u8>, Error> {
    let mut cell = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut sp = EscherBuilder::new(header_version::SP, shape_type::RECTANGLE, record_type::SP);
    sp.add_data(
        EscherSpData::with_flags(
            cell_spid,
            ShapeFlags::CHILD | ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT,
        )
        .as_bytes(),
    );
    cell.add_data(&sp.build()?);

    let mut anchor = EscherBuilder::new(header_version::SIMPLE, 0, RECORD_TYPE_CHILD_ANCHOR);
    anchor.add_data(
        ChildAnchor {
            left,
            top,
            right: left + width,
            bottom: top + height,
        }
        .as_bytes(),
    );
    cell.add_data(&anchor.build()?);

    cell.add_data(&build_client_textbox(text, CELL_TEXT_TYPE)?);

    cell.build()
}

#[cfg(test)]
mod tests;
