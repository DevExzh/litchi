//! Mutable XLSB worksheet for CRUD operations

use crate::ooxml::xlsb::comments::Comment;
use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::hyperlinks::Hyperlink;
use crate::ooxml::xlsb::merged_cells::MergedCell;
use crate::ooxml::xlsb::records::record_types;
use crate::ooxml::xlsb::writer::RecordWriter;
use crate::sheet::CellValue;
use std::collections::BTreeMap;
use std::io::Write;

/// Cell data for storage
#[derive(Debug, Clone)]
pub struct CellData {
    pub value: CellValue,
    pub style: u32, // Style XF index
}

/// Column information for a single 0-based column.
///
/// This writer-side structure drives `BrtColInfo` emission and mirrors the
/// semantics of [MS-XLSB] 2.4.323 and SheetJS' `write_BrtColInfo` helper.
#[derive(Debug, Clone)]
pub struct ColumnInfo {
    /// Column width in character units. `None` uses the sheet default.
    pub width: Option<f64>,
    /// Whether the column is hidden.
    pub hidden: bool,
    /// Whether the column width was inferred via best-fit.
    pub best_fit: bool,
}

/// Row information for a single 0-based row.
#[derive(Debug, Clone)]
pub struct RowInfo {
    /// Row height in points. `None` uses the sheet default.
    pub height: Option<f64>,
    /// Whether the row is hidden.
    pub hidden: bool,
}

/// Auto-filter configuration for a rectangular range.
///
/// The indices are 0-based and inclusive.
#[derive(Debug, Clone)]
pub struct AutoFilter {
    /// First row (0-based).
    pub row_first: u32,
    /// Last row (0-based, inclusive).
    pub row_last: u32,
    /// First column (0-based).
    pub col_first: u32,
    /// Last column (0-based, inclusive).
    pub col_last: u32,
}

/// Sheet protection options for XLSB.
///
/// This is a minimal representation used to drive the `BrtSheetProtection`
/// record. Individual flags are optional; when `None` the default from the
/// [MS-XLSB] examples / SheetJS writer is used.
#[derive(Debug, Clone, Default)]
pub struct SheetProtection {
    /// Optional password hash (Method 1). When `None`, no password is enforced.
    pub password_hash: Option<u16>,
    pub objects: Option<bool>,
    pub scenarios: Option<bool>,
    pub format_cells: Option<bool>,
    pub format_columns: Option<bool>,
    pub format_rows: Option<bool>,
    pub insert_columns: Option<bool>,
    pub insert_rows: Option<bool>,
    pub insert_hyperlinks: Option<bool>,
    pub delete_columns: Option<bool>,
    pub delete_rows: Option<bool>,
    pub select_locked_cells: Option<bool>,
    pub sort: Option<bool>,
    pub auto_filter: Option<bool>,
    pub pivot_tables: Option<bool>,
    pub select_unlocked_cells: Option<bool>,
}

/// Mutable XLSB worksheet supporting CRUD operations
#[derive(Debug, Clone)]
pub struct MutableXlsbWorksheet {
    name: String,
    cells: BTreeMap<(u32, u32), CellData>,
    max_row: u32,
    max_col: u32,
    merged_cells: Vec<MergedCell>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Comment>,
    /// Column information (0-based column index).
    columns: BTreeMap<u32, ColumnInfo>,
    /// Row information (0-based row index).
    rows: BTreeMap<u32, RowInfo>,
    /// Optional auto-filter configuration.
    auto_filter: Option<AutoFilter>,
    /// Optional sheet protection configuration.
    sheet_protection: Option<SheetProtection>,
}

impl MutableXlsbWorksheet {
    /// Create a new empty worksheet
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let sheet = MutableXlsbWorksheet::new("Sheet1");
    /// ```
    pub fn new<S: Into<String>>(name: S) -> Self {
        MutableXlsbWorksheet {
            name: name.into(),
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            columns: BTreeMap::new(),
            rows: BTreeMap::new(),
            auto_filter: None,
            sheet_protection: None,
        }
    }

    /// Get the worksheet name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Rename the worksheet
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Set a cell value
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, "Hello");
    /// sheet.set_cell(0, 1, 42.0);
    /// sheet.set_cell(1, 0, true);
    /// ```
    pub fn set_cell<V: Into<CellValue>>(&mut self, row: u32, col: u32, value: V) {
        self.set_cell_with_style(row, col, value, 0);
    }

    /// Set a cell value with style
    pub fn set_cell_with_style<V: Into<CellValue>>(
        &mut self,
        row: u32,
        col: u32,
        value: V,
        style: u32,
    ) {
        let cell_data = CellData {
            value: value.into(),
            style,
        };

        self.cells.insert((row, col), cell_data);
        self.max_row = self.max_row.max(row);
        self.max_col = self.max_col.max(col);
    }

    /// Get a cell value
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&CellValue> {
        self.cells.get(&(row, col)).map(|c| &c.value)
    }

    /// Delete a cell
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, "Hello");
    /// sheet.delete_cell(0, 0);
    /// assert!(sheet.get_cell(0, 0).is_none());
    /// ```
    pub fn delete_cell(&mut self, row: u32, col: u32) -> Option<CellValue> {
        self.cells.remove(&(row, col)).map(|c| c.value)
    }

    /// Clear all cells in the worksheet
    pub fn clear(&mut self) {
        self.cells.clear();
        self.max_row = 0;
        self.max_col = 0;
        self.merged_cells.clear();
        self.hyperlinks.clear();
        self.comments.clear();
        self.columns.clear();
        self.rows.clear();
        self.auto_filter = None;
        self.sheet_protection = None;
    }

    /// Set a custom column width (in character units) for a 0-based column.
    ///
    /// This controls the `BrtColInfo` width field. The default width from the
    /// sheet format properties (`BrtSheetFormatPr`) is used when no explicit
    /// width is set.
    pub fn set_column_width(&mut self, col: u32, width: f64) {
        let entry = self.columns.entry(col).or_insert(ColumnInfo {
            width: None,
            hidden: false,
            best_fit: false,
        });
        entry.width = Some(width);
    }

    /// Set a custom row height (in points) for a 0-based row.
    ///
    /// Heights are encoded in twips (1/20 of a point) in the `BrtRowHdr`
    /// records. When no explicit height is set, Excel's default of 15 points
    /// (300 twips) is used.
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        let entry = self.rows.entry(row).or_insert(RowInfo {
            height: None,
            hidden: false,
        });
        entry.height = Some(height);
    }

    /// Configure a basic auto-filter range for the worksheet.
    ///
    /// The indices are 0-based and inclusive.
    pub fn set_auto_filter(
        &mut self,
        row_first: u32,
        row_last: u32,
        col_first: u32,
        col_last: u32,
    ) {
        self.auto_filter = Some(AutoFilter {
            row_first,
            row_last,
            col_first,
            col_last,
        });
    }

    /// Set sheet protection options. Passing `None` clears protection.
    pub fn set_sheet_protection(&mut self, protection: Option<SheetProtection>) {
        self.sheet_protection = protection;
    }

    /// Add a merged cell range
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi::ooxml::xlsb::advanced_features::MergedCell;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.add_merged_cell(MergedCell::new(0, 1, 0, 1)); // Merge A1:B2
    /// ```
    pub fn add_merged_cell(&mut self, merged: MergedCell) {
        self.merged_cells.push(merged);
    }

    /// Add a hyperlink
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi::ooxml::xlsb::advanced_features::Hyperlink;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// let link = Hyperlink::new(0, 0, 0, 0, "rId1".to_string())
    ///     .with_tooltip("Visit website".to_string());
    /// sheet.add_hyperlink(link);
    /// ```
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) {
        self.hyperlinks.push(hyperlink);
    }

    /// Add a comment
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::MutableXlsbWorksheet;
    /// use litchi::ooxml::xlsb::advanced_features::Comment;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// let comment = Comment::new(0, 0, "John".to_string(), "Important note".to_string());
    /// sheet.add_comment(comment);
    /// ```
    pub fn add_comment(&mut self, comment: Comment) {
        self.comments.push(comment);
    }

    /// Get all merged cells
    pub fn merged_cells(&self) -> &[MergedCell] {
        &self.merged_cells
    }

    /// Get all hyperlinks
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }

    /// Get mutable access to all hyperlinks.
    ///
    /// This is primarily used by the workbook writer to inject concrete
    /// relationship IDs (`rId`) after creating external OPC relationships
    /// but before serializing `BrtHLink` records.
    pub(crate) fn hyperlinks_mut(&mut self) -> &mut [Hyperlink] {
        &mut self.hyperlinks
    }

    /// Get all comments
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Get the number of non-empty cells
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Get dimensions (min_row, min_col, max_row, max_col)
    pub fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        if self.cells.is_empty() {
            None
        } else {
            Some((0, 0, self.max_row, self.max_col))
        }
    }

    /// Delete a row (shifts remaining rows up)
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::MutableXlsbWorksheet;
    ///
    /// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    /// sheet.set_cell(0, 0, "Row 0");
    /// sheet.set_cell(1, 0, "Row 1");
    /// sheet.set_cell(2, 0, "Row 2");
    ///
    /// sheet.delete_row(1);
    ///
    /// // Row 2 becomes row 1
    /// assert_eq!(sheet.get_cell(1, 0).and_then(|v| v.as_str()), Some("Row 2"));
    /// ```
    pub fn delete_row(&mut self, row: u32) {
        // Remove all cells in the row
        self.cells.retain(|(r, _), _| *r != row);

        // Shift rows after the deleted row up
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((r, _), _)| *r > row)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r - 1, c), cell);
        }

        // Recalculate max_row
        self.max_row = self.cells.keys().map(|(r, _)| *r).max().unwrap_or(0);
    }

    /// Delete a column (shifts remaining columns left)
    pub fn delete_column(&mut self, col: u32) {
        // Remove all cells in the column
        self.cells.retain(|(_, c), _| *c != col);

        // Shift columns after the deleted column left
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((_, c), _)| *c > col)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r, c - 1), cell);
        }

        // Recalculate max_col
        self.max_col = self.cells.keys().map(|(_, c)| *c).max().unwrap_or(0);
    }

    /// Insert a row (shifts existing rows down)
    pub fn insert_row(&mut self, row: u32) {
        // Shift rows at and after the insert position down
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((r, _), _)| *r >= row)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r + 1, c), cell);
        }

        // Recalculate max_row
        self.max_row = self.cells.keys().map(|(r, _)| *r).max().unwrap_or(0);
    }

    /// Insert a column (shifts existing columns right)
    pub fn insert_column(&mut self, col: u32) {
        // Shift columns at and after the insert position right
        let cells_to_move: Vec<_> = self
            .cells
            .iter()
            .filter(|((_, c), _)| *c >= col)
            .map(|((r, c), cell)| (*r, *c, cell.clone()))
            .collect();

        for (r, c, cell) in cells_to_move {
            self.cells.remove(&(r, c));
            self.cells.insert((r, c + 1), cell);
        }

        // Recalculate max_col
        self.max_col = self.cells.keys().map(|(_, c)| *c).max().unwrap_or(0);
    }

    /// Write worksheet to binary format
    ///
    /// Following Excel's required structure
    pub(crate) fn write<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        shared_strings: &mut crate::ooxml::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        // Write BrtBeginSheet
        writer.write_record(record_types::BEGIN_SHEET, &[])?;

        // Write worksheet properties and basic formatting information.
        self.write_ws_properties(writer)?;

        // Write worksheet dimensions
        self.write_dimensions(writer)?;

        // Write worksheet views (minimal SheetJS-style layout)
        self.write_ws_views(writer)?;

        // Write sheet formatting properties (BrtSheetFormatPr)
        self.write_sheet_format_pr(writer)?;

        // Column information (BrtBeginColInfos / BrtColInfo / BrtEndColInfos)
        self.write_col_infos(writer)?;

        // Write sheet data
        writer.write_record(record_types::BEGIN_SHEET_DATA, &[])?;
        self.write_cells(writer, shared_strings)?;
        writer.write_record(record_types::END_SHEET_DATA, &[])?;

        // Sheet protection (BrtSheetProtection) - minimal skeleton mirroring
        // SheetJS and [MS-XLSB] examples.
        self.write_sheet_protection(writer)?;

        // AutoFilter skeleton (BrtBeginAFilter / BrtEndAFilter).
        self.write_auto_filter(writer)?;

        // Write merged cells if present
        if !self.merged_cells.is_empty() {
            self.write_merged_cells(writer)?;
        }

        // Write hyperlinks if present
        if !self.hyperlinks.is_empty() {
            self.write_hyperlinks(writer)?;
        }

        // Write BrtEndSheet
        writer.write_record(record_types::END_SHEET, &[])?;

        Ok(())
    }

    /// Write worksheet properties (BrtWsProp) - REQUIRED by Excel
    ///
    /// [MS-XLSB] 2.4.864 + spec example 3.7.21: 23 bytes total
    /// Structure: flags (3 bytes) + brtcolorTab (8 bytes) + rwSync (4) + colSync (4) + strName (4)
    fn write_ws_properties<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Flags (3 bytes per spec example 3.7.21):
        // Byte 0-1 (USHORT): flags A-O
        // Byte 2 (BYTE): flags P-Q + reserved
        //
        // From spec example: 0xC9, 0x04, 0x02
        // 0xC9 = fShowAutoBreaks(1) + fPublish(1) + fRowSumsBelow(1) + fColSumsRight(1) + fShowOutlineSymbols(1)
        // 0x04 = remaining bits
        // 0x02 = fCondFmtCalc(1) at bit 1
        temp_writer.write_u8(0xC9)?;
        temp_writer.write_u8(0x04)?;
        temp_writer.write_u8(0x02)?; // Third byte - fCondFmtCalc flag

        // brtcolorTab (8 bytes) - BrtColor structure
        // From spec example: xColorType=0x00 (auto), index=0x40
        temp_writer.write_u8(0x00)?; // fValidRGB(0) + xColorType(0x00)
        temp_writer.write_u8(0x40)?; // index
        temp_writer.write_u16(0)?; // nTintAndShade
        temp_writer.write_u8(0)?; // bRed
        temp_writer.write_u8(0)?; // bGreen
        temp_writer.write_u8(0)?; // bBlue
        temp_writer.write_u8(0)?; // bAlpha

        // rwSync (4 bytes) - RwNullable: 0xFFFFFFFF = no synchronization
        temp_writer.write_u32(0xFFFFFFFF)?;

        // colSync (4 bytes) - ColNullable: 0xFFFFFFFF = no synchronization
        temp_writer.write_u32(0xFFFFFFFF)?;

        // strName - CodeName (XLWideString): empty string
        temp_writer.write_u32(0)?;

        writer.write_record(record_types::WS_PROP, &data)?;
        Ok(())
    }

    /// Write worksheet views (REQUIRED by Excel)
    ///
    /// [MS-XLSB] 2.4.304: Specifies sheet view settings
    fn write_ws_views<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        writer.write_record(record_types::BEGIN_WS_VIEWS, &[])?;

        // BrtBeginWsView (30 bytes according to spec)
        let mut view_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut view_data);

        // Flags (2 bytes) - bits A-K + reserved
        // Default: fDspGrid=1, fDspRwCol=1, fDspZeros=1, fDefaultHdr=1
        // 0xDC = 11011100 = fDefaultHdr(1) + fDspGuts(1) + fSelected(1) + fDspZeros(1) + fDspRwCol(1) + fDspGrid(1)
        // 0x03 = 00000011 = reserved bits
        temp_writer.write_u8(0xDC)?;
        temp_writer.write_u8(0x03)?;

        // xlView (4 bytes) - XLView: 0 = normal view
        temp_writer.write_u32(0)?;

        // rwTop (4 bytes) - first row displayed
        temp_writer.write_u32(0)?;

        // colLeft (4 bytes) - first column displayed
        temp_writer.write_u32(0)?;

        // icvHdr (1 byte) - Icv: gridline color (0x40 = default)
        temp_writer.write_u8(0x40)?;

        // reserved2 (1 byte)
        temp_writer.write_u8(0)?;

        // reserved3 (2 bytes)
        temp_writer.write_u16(0)?;

        // wScale (2 bytes) - zoom level (100%)
        temp_writer.write_u16(100)?;

        // wScaleNormal (2 bytes) - per spec example: 0 means default 100
        temp_writer.write_u16(0)?;

        // wScaleSLV (2 bytes) - zoom for page break preview (0 = default 100%)
        temp_writer.write_u16(0)?;

        // wScalePLV (2 bytes) - zoom for page layout view (0 = default 100%)
        temp_writer.write_u16(0)?;

        // iWbkView (4 bytes) - workbook view index
        temp_writer.write_u32(0)?;

        // Minimal SheetJS-style view: BrtBeginWsViews / BrtBeginWsView / BrtEndWsView / BrtEndWsViews
        writer.write_record(record_types::BEGIN_WS_VIEW, &view_data)?;

        writer.write_record(record_types::END_WS_VIEW, &[])?;
        writer.write_record(record_types::END_WS_VIEWS, &[])?;

        Ok(())
    }

    /// Write SHEET_FORMAT_PR record (0x01E5) - sheet formatting properties
    /// REQUIRED by Excel
    ///
    /// [MS-XLSB] 2.4.862 + spec example 3.7.28: 12 bytes total
    fn write_sheet_format_pr<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // dxGCol (4 bytes) - 0xFFFFFFFF = use cchDefColWidth instead
        temp_writer.write_u32(0xFFFFFFFF)?;

        // cchDefColWidth (2 bytes) - default column width in characters
        // Spec example 3.7.28: 0x0008 (8 characters)
        temp_writer.write_u16(8)?;

        // miyDefRwHeight (2 bytes) - default row height in twips
        // Spec example 3.7.28: 0x012C (300 twips = 15 points)
        temp_writer.write_u16(300)?;

        // Flags (4 bytes): all zeros per spec example
        // fUnsynced=0, fDyZero=0, fExAsc=0, fExDesc=0, reserved=0, iOutLevelRw=0, iOutLevelCol=0
        temp_writer.write_u32(0)?;

        writer.write_record(0x01E5, &data)?;
        Ok(())
    }

    /// Write worksheet dimensions record
    fn write_dimensions<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        if let Some((min_row, min_col, max_row, max_col)) = self.dimensions() {
            temp_writer.write_u32(min_row)?;
            temp_writer.write_u32(max_row)?;
            temp_writer.write_u32(min_col)?;
            temp_writer.write_u32(max_col)?;
        } else {
            // Empty worksheet
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
            temp_writer.write_u32(0)?;
        }

        writer.write_record(record_types::WS_DIM, &data)?;
        Ok(())
    }

    /// Write all cells
    fn write_cells<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        shared_strings: &mut crate::ooxml::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        let mut current_row: Option<u32> = None;

        for ((row, col), cell_data) in &self.cells {
            // Write row header if row changed
            if current_row != Some(*row) {
                self.write_row_header(writer, *row)?;
                current_row = Some(*row);
            }

            // Write cell
            self.write_cell(writer, *row, *col, cell_data, shared_strings)?;
        }

        Ok(())
    }

    /// Write row header record with BrtColSpan elements
    ///
    /// BrtRowHdr structure (2.4.761):
    /// - rw (4 bytes): Row index
    /// - ixfe (4 bytes): Style index
    /// - miyRw (2 bytes): Row height in twips (1/20 of a point)
    /// - flags1 (1 byte): fExtraAsc | fExtraDsc | reserved
    /// - flags2 (1 byte): outline/visibility flags
    /// - phonetic (1 byte): phonetic guide flags
    /// - ccolspan (4 bytes): number of BrtColSpan elements
    /// - rgBrtColspan (variable): array of BrtColSpan, each 8 bytes
    ///   (colFirst (u32) + colLast (u32))
    fn write_row_header<W: Write>(&self, writer: &mut RecordWriter<W>, row: u32) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Fixed part
        temp_writer.write_u32(row)?; // rw: Row index
        temp_writer.write_u32(0)?; // ixfe: Style index (0 = default)

        // Row height in twips (1/20 of a point). When no explicit height is
        // configured, use Excel's default of 15 points (300 twips).
        let miy_rw: u16 = if let Some(info) = self.rows.get(&row) {
            if let Some(height_pts) = info.height {
                (height_pts * 20.0).round() as u16
            } else {
                0x012C
            }
        } else {
            0x012C
        };
        temp_writer.write_u16(miy_rw)?;

        // flags1: extra ascender/descender padding (unused here).
        temp_writer.write_u8(0)?;

        // flags2: outline / visibility / custom height flags.
        // Bits 0-2: outline level, 0x10: hidden, 0x20: custom height.
        let mut flags2: u8 = 0;
        if let Some(info) = self.rows.get(&row) {
            if info.hidden {
                flags2 |= 0x10;
            }
            if info.height.is_some() {
                flags2 |= 0x20;
            }
        }
        temp_writer.write_u8(flags2)?;

        // phonetic guide: 0 = no phonetic information
        temp_writer.write_u8(0)?;

        // Collect all columns that have cells in this row (BTreeMap preserves sorted order)
        let cells_in_row: Vec<u32> = self
            .cells
            .keys()
            .filter(|(r, _)| *r == row)
            .map(|(_, c)| *c)
            .collect();

        if cells_in_row.is_empty() {
            // No cells in row - write 0 colspans
            temp_writer.write_u32(0)?;
        } else {
            // Group columns by 1024-wide segments, as in [MS-XLSB] BrtColSpan and SheetJS
            let mut spans: Vec<(u32, u32)> = Vec::new();
            let mut current_segment = cells_in_row[0] / 1024;
            let mut segment_first = cells_in_row[0];
            let mut segment_last = cells_in_row[0];

            for &col in &cells_in_row[1..] {
                let segment = col / 1024;
                if segment == current_segment {
                    segment_last = col;
                } else {
                    spans.push((segment_first, segment_last));
                    current_segment = segment;
                    segment_first = col;
                    segment_last = col;
                }
            }
            spans.push((segment_first, segment_last));

            // Number of spans
            temp_writer.write_u32(spans.len() as u32)?;

            // Each span is a BrtColSpan: colFirst (u32) + colLast (u32)
            for (first, last) in spans {
                temp_writer.write_u32(first)?; // colFirst
                temp_writer.write_u32(last)?; // colLast
            }
        }

        writer.write_record(record_types::ROW_HDR, &data)?;
        Ok(())
    }

    /// Write a single cell record
    fn write_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        _row: u32,
        col: u32,
        cell_data: &CellData,
        shared_strings: &mut crate::ooxml::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        match &cell_data.value {
            CellValue::Empty => self.write_blank_cell(writer, col, cell_data.style)?,
            CellValue::String(s) => {
                self.write_shared_string_cell(writer, col, s, cell_data.style, shared_strings)?
            },
            CellValue::Int(i) => self.write_number_cell(writer, col, *i as f64, cell_data.style)?,
            CellValue::Float(f) => self.write_number_cell(writer, col, *f, cell_data.style)?,
            CellValue::Bool(b) => self.write_bool_cell(writer, col, *b, cell_data.style)?,
            CellValue::Error(e) => self.write_error_cell(writer, col, e, cell_data.style)?,
            CellValue::DateTime(dt) => {
                // Excel DateTime is already stored as serial number (days since epoch)
                // CellValue::DateTime stores the Excel serial number directly
                self.write_number_cell(writer, col, *dt, cell_data.style)?;
            },
            CellValue::Formula { cached_value, .. } => {
                // For formulas, write the cached value
                // TODO: Support writing formula bytes
                if let Some(cached) = cached_value {
                    match cached.as_ref() {
                        CellValue::Empty => self.write_blank_cell(writer, col, cell_data.style)?,
                        CellValue::String(s) => self.write_shared_string_cell(
                            writer,
                            col,
                            s,
                            cell_data.style,
                            shared_strings,
                        )?,
                        CellValue::Int(i) => {
                            self.write_number_cell(writer, col, *i as f64, cell_data.style)?
                        },
                        CellValue::Float(f) => {
                            self.write_number_cell(writer, col, *f, cell_data.style)?
                        },
                        CellValue::Bool(b) => {
                            self.write_bool_cell(writer, col, *b, cell_data.style)?
                        },
                        CellValue::Error(e) => {
                            self.write_error_cell(writer, col, e, cell_data.style)?
                        },
                        CellValue::DateTime(dt) => {
                            self.write_number_cell(writer, col, *dt, cell_data.style)?
                        },
                        CellValue::Formula { .. } => {
                            // Nested formula - shouldn't happen, but write as blank
                            self.write_blank_cell(writer, col, cell_data.style)?;
                        },
                    }
                } else {
                    // No cached value, write as blank
                    self.write_blank_cell(writer, col, cell_data.style)?;
                }
            },
        }
        Ok(())
    }

    /// Write the Cell structure (2.5.10) - 8 bytes
    ///
    /// Cell structure:
    /// - column (4 bytes): Column index
    /// - iStyleRef (3 bytes, 24-bit): Style XF index
    /// - fPhShow (1 bit): Phonetic info flag
    /// - reserved (7 bits): Reserved
    fn write_cell_structure<W: Write>(
        temp_writer: &mut RecordWriter<W>,
        col: u32,
        style: u32,
    ) -> XlsbResult<()> {
        // Column (4 bytes)
        temp_writer.write_u32(col)?;

        // iStyleRef (3 bytes) + flags (1 byte) = 4 bytes total
        temp_writer.write_u8((style & 0xFF) as u8)?;
        temp_writer.write_u8(((style >> 8) & 0xFF) as u8)?;
        temp_writer.write_u8(((style >> 16) & 0xFF) as u8)?;
        temp_writer.write_u8(0)?; // fPhShow=0, reserved=0

        Ok(())
    }

    /// Write a blank cell (BrtCellBlank - 8 bytes)
    fn write_blank_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        style: u32,
    ) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        Self::write_cell_structure(&mut temp_writer, col, style)?;

        writer.write_record(record_types::CELL_BLANK, &data)?;
        Ok(())
    }

    /// Write a shared string cell (BrtCellIsst - Cell + u32 = 12 bytes)
    fn write_shared_string_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        value: &str,
        style: u32,
        shared_strings: &mut crate::ooxml::xlsb::writer::MutableSharedStringsWriter,
    ) -> XlsbResult<()> {
        // Add string to shared strings table and get index
        let string_index = shared_strings.add_string(value.to_string());

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + isst index (4 bytes) = 12 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u32(string_index)?;

        writer.write_record(record_types::CELL_ISST, &data)?;
        Ok(())
    }

    /// Write a number cell (BrtCellReal - Cell + f64 = 16 bytes)
    fn write_number_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        value: f64,
        style: u32,
    ) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + Xnum value (8 bytes) = 16 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_f64(value)?;

        writer.write_record(record_types::CELL_REAL, &data)?;
        Ok(())
    }

    /// Write a boolean cell (BrtCellBool - Cell + u8 = 9 bytes)
    fn write_bool_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        value: bool,
        style: u32,
    ) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + fBool (1 byte) = 9 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u8(if value { 1 } else { 0 })?;

        writer.write_record(record_types::CELL_BOOL, &data)?;
        Ok(())
    }

    /// Write an error cell (BrtCellError - Cell + u8 = 9 bytes)
    fn write_error_cell<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        col: u32,
        error: &str,
        style: u32,
    ) -> XlsbResult<()> {
        let error_code = match error {
            "#NULL!" => 0x00,
            "#DIV/0!" => 0x07,
            "#VALUE!" => 0x0F,
            "#REF!" => 0x17,
            "#NAME?" => 0x1D,
            "#NUM!" => 0x24,
            "#N/A" => 0x2A,
            "#GETTING_DATA" => 0x2B,
            _ => 0x2A, // Default to #N/A
        };

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Cell structure (8 bytes) + bError (1 byte) = 9 bytes
        Self::write_cell_structure(&mut temp_writer, col, style)?;
        temp_writer.write_u8(error_code)?;

        writer.write_record(record_types::CELL_ERROR, &data)?;
        Ok(())
    }

    /// Write merged cells
    fn write_merged_cells<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        // BrtBeginMergeCells (0x00B1) payload is a single DWORD count of BrtMergeCell
        // records that follow. SheetJS writes this as write_BrtBeginMergeCells(cnt).
        let mut header = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut header);
        temp_writer.write_u32(self.merged_cells.len() as u32)?;

        writer.write_record(record_types::BEGIN_MERGE_CELLS, &header)?;

        for merged in &self.merged_cells {
            let data = merged.serialize();
            writer.write_record(record_types::MERGE_CELL, &data)?;
        }

        writer.write_record(record_types::END_MERGE_CELLS, &[])?;
        Ok(())
    }

    /// Write column information records.
    fn write_col_infos<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        if self.columns.is_empty() {
            return Ok(());
        }

        writer.write_record(record_types::BEGIN_COL_INFOS, &[])?;

        for (col, info) in &self.columns {
            let mut data = Vec::new();
            let mut temp_writer = RecordWriter::new(&mut data);

            // firstCol / lastCol (both 0-based inclusive).
            temp_writer.write_u32(*col)?;
            temp_writer.write_u32(*col)?;

            // Width is stored as 256ths of a character, mirroring SheetJS
            // write_BrtColInfo and [MS-XLSB] 2.4.323.
            let width_chars = info.width.unwrap_or(10.0);
            let width_raw = (width_chars * 256.0).round() as u32;
            temp_writer.write_u32(width_raw)?;

            // Style XF index (we currently do not support per-column styles).
            temp_writer.write_u32(0)?;

            // Flags (2 bytes): 0x0001 = hidden, 0x0002 = custom width,
            // 0x0008 = best fit.
            let mut flags: u16 = 0;
            if info.hidden {
                flags |= 0x0001;
            }
            if info.width.is_some() {
                flags |= 0x0002;
            }
            if info.best_fit {
                flags |= 0x0008;
            }
            temp_writer.write_u16(flags)?;

            writer.write_record(record_types::COL_INFO, &data)?;
        }

        writer.write_record(record_types::END_COL_INFOS, &[])?;
        Ok(())
    }

    /// Write hyperlinks
    fn write_hyperlinks<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        for hyperlink in &self.hyperlinks {
            let data = hyperlink.serialize();
            writer.write_record(record_types::H_LINK, &data)?;
        }
        Ok(())
    }

    /// Write sheet protection if configured.
    fn write_sheet_protection<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let Some(ref prot) = self.sheet_protection else {
            return Ok(());
        };

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // Password hash (Method 1). When absent, write 0.
        temp_writer.write_u16(prot.password_hash.unwrap_or(0))?;

        // Guard DWORD: this record should not be written if no protection.
        temp_writer.write_u32(1)?;

        fn flag(default_true: bool, value: Option<bool>) -> u32 {
            if default_true {
                if let Some(v) = value {
                    if !v { 1 } else { 0 }
                } else {
                    0
                }
            } else if let Some(v) = value {
                if v { 0 } else { 1 }
            } else {
                1
            }
        }

        temp_writer.write_u32(flag(false, prot.objects))?;
        temp_writer.write_u32(flag(false, prot.scenarios))?;
        temp_writer.write_u32(flag(true, prot.format_cells))?;
        temp_writer.write_u32(flag(true, prot.format_columns))?;
        temp_writer.write_u32(flag(true, prot.format_rows))?;
        temp_writer.write_u32(flag(true, prot.insert_columns))?;
        temp_writer.write_u32(flag(true, prot.insert_rows))?;
        temp_writer.write_u32(flag(true, prot.insert_hyperlinks))?;
        temp_writer.write_u32(flag(true, prot.delete_columns))?;
        temp_writer.write_u32(flag(true, prot.delete_rows))?;
        temp_writer.write_u32(flag(false, prot.select_locked_cells))?;
        temp_writer.write_u32(flag(true, prot.sort))?;
        temp_writer.write_u32(flag(true, prot.auto_filter))?;
        temp_writer.write_u32(flag(true, prot.pivot_tables))?;
        temp_writer.write_u32(flag(false, prot.select_unlocked_cells))?;

        writer.write_record(record_types::SHEET_PROTECTION, &data)?;
        Ok(())
    }

    /// Write basic auto-filter range if configured.
    fn write_auto_filter<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let Some(ref af) = self.auto_filter else {
            return Ok(());
        };

        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // UncheckedRfX: row_first, row_last, col_first, col_last
        temp_writer.write_u32(af.row_first)?;
        temp_writer.write_u32(af.row_last)?;
        temp_writer.write_u32(af.col_first)?;
        temp_writer.write_u32(af.col_last)?;

        writer.write_record(record_types::BEGIN_A_FILTER, &data)?;
        writer.write_record(record_types::END_A_FILTER, &[])?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_set_and_get_cell() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Hello");
        sheet.set_cell(1, 1, 42.0);

        assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Hello"));
        assert_eq!(sheet.get_cell(1, 1).and_then(|v| v.as_float()), Some(42.0));
    }

    #[test]
    fn test_delete_cell() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Hello");

        assert!(sheet.delete_cell(0, 0).is_some());
        assert!(sheet.get_cell(0, 0).is_none());
    }

    #[test]
    fn test_delete_row() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Row 0");
        sheet.set_cell(1, 0, "Row 1");
        sheet.set_cell(2, 0, "Row 2");

        sheet.delete_row(1);

        assert_eq!(sheet.get_cell(0, 0).and_then(|v| v.as_str()), Some("Row 0"));
        assert_eq!(sheet.get_cell(1, 0).and_then(|v| v.as_str()), Some("Row 2"));
        assert!(sheet.get_cell(2, 0).is_none());
    }

    #[test]
    fn test_dimensions() {
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        assert!(sheet.dimensions().is_none());

        sheet.set_cell(5, 10, "Test");
        assert_eq!(sheet.dimensions(), Some((0, 0, 5, 10)));
    }
}
