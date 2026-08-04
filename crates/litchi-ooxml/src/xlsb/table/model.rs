//! Typed model for the XLSB Table (ListObject) stream (MS-XLSB 2.1.7.51).
//!
//! All types are inert data snapshots: relationship identifiers, external
//! connection identifiers, differential-formatting identifiers, and formula
//! token streams are stored verbatim and are never dereferenced, contacted,
//! or evaluated. Display names are preserved exactly as stored — the parser
//! performs no Excel display-name validation (a display name containing
//! spaces, for example, is kept as-is even though Excel would reject it).

use crate::xlsb::error::XlsbError;

/// A structured table parsed from one `tables/table*.bin` part
/// (`BrtBeginList` and its record collection, MS-XLSB 2.4.100).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Table {
    /// Numeric identifier of the table (`idList`); unique within the workbook.
    pub id: u32,
    /// String identifier used for programmatic purposes (`stName`); when
    /// `None`, `display_name` takes that role.
    pub name: Option<String>,
    /// String identifier used in displayed formulas (`stDisplayName`),
    /// preserved verbatim without display-name validation.
    pub display_name: Option<String>,
    /// Comment about the table (`stComment`).
    pub comment: Option<String>,
    /// Cell range the table occupies (`rfxList`).
    pub range: Range,
    /// Table type (`lt`, a `ListType` value, MS-XLSB 2.5.89).
    pub table_type: Type,
    /// Header row count (`crwHeader`; `0` or `1` per the Boolean encoding).
    pub header_row_count: u32,
    /// Total row count (`crwTotals`; `0` or `1` per the Boolean encoding).
    pub totals_row_count: u32,
    /// The total row has ever been displayed (`fShownTotalRow`).
    pub totals_row_shown: bool,
    /// The table is a single cell table (`fSingleCell`).
    pub single_cell: bool,
    /// The table insert row is displayed (`fForceInsertToBeVisible`).
    pub insert_row_visible: bool,
    /// Cells were automatically inserted when the insert row was displayed
    /// (`fInsertRowInsCells`).
    pub insert_row_inserted_cells: bool,
    /// Publish-to-server state (`fPublished`).
    pub published: bool,
    /// Differential formatting of the header row (`nDxfHeader`); `None` = none.
    pub header_dxf_id: Option<u32>,
    /// Differential formatting of the data region (`nDxfData`); `None` = none.
    pub data_dxf_id: Option<u32>,
    /// Differential formatting of the total row (`nDxfAgg`); `None` = none.
    pub totals_dxf_id: Option<u32>,
    /// Differential formatting of the data-region borders (`nDxfBorder`);
    /// `None` = none.
    pub border_dxf_id: Option<u32>,
    /// Differential formatting of the header-row borders (`nDxfHeaderBorder`);
    /// `None` = none.
    pub header_border_dxf_id: Option<u32>,
    /// Differential formatting of the total-row borders (`nDxfAggBorder`);
    /// `None` = none.
    pub totals_border_dxf_id: Option<u32>,
    /// Inert identifier of the external connection used by this table
    /// (`dwConnID`); `None` when zero. Never resolved or contacted.
    pub connection_id: Option<u32>,
    /// Cell style applied to the header row (`stStyleHeader`).
    pub header_style: Option<String>,
    /// Cell style applied to the data region (`stStyleData`).
    pub data_style: Option<String>,
    /// Cell style applied to the total row (`stStyleAgg`).
    pub totals_style: Option<String>,
    /// Table columns in column index order (`BrtBeginListCols` collection).
    pub columns: Vec<Column>,
    /// Table style information (`BrtTableStyleClient`, MS-XLSB 2.4.847).
    pub style_info: Option<StyleInfo>,
    /// Alternate text of the table (`stAltText` of `BrtList14`,
    /// MS-XLSB 2.4.705).
    pub alternate_text: Option<String>,
    /// Alternate text summary of the table (`stAltTextSummary` of
    /// `BrtList14`).
    pub alternate_text_summary: Option<String>,
}

/// A cell range (`RfX`, MS-XLSB 2.5.118).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Range {
    /// First row (`rwFirst`).
    pub first_row: u32,
    /// Last row (`rwLast`).
    pub last_row: u32,
    /// First column (`colFirst`).
    pub first_column: u32,
    /// Last column (`colLast`).
    pub last_column: u32,
}

/// Table type (`ListType`, MS-XLSB 2.5.89).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[repr(u32)]
pub enum Type {
    /// Standard table (`LTRANGE`).
    #[default]
    Range = 0,
    /// XML table (`LTXML`).
    Xml = 2,
    /// Query table (`LTEXTDATA`).
    QueryTable = 3,
}

impl TryFrom<u32> for Type {
    type Error = XlsbError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Range),
            2 => Ok(Self::Xml),
            3 => Ok(Self::QueryTable),
            _ => Err(XlsbError::Unrecognized {
                typ: "table type".to_string(),
                val: format!("0x{value:08X}"),
            }),
        }
    }
}

/// Total row aggregation function of a table column (`ListTotalRowFunction`,
/// MS-XLSB 2.5.88).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[repr(u32)]
pub enum TotalsRowFunction {
    /// No operation (`ILTA_NONE`).
    #[default]
    None = 0,
    /// Arithmetic mean; subtotal function number 101 (`ILTA_AVERAGE`).
    Average = 1,
    /// Count of non-empty cells; subtotal function number 103 (`ILTA_COUNT`).
    Count = 2,
    /// Count of numeric cells; subtotal function number 102 (`ILTA_COUNTNUMS`).
    CountNums = 3,
    /// Largest value; subtotal function number 104 (`ILTA_MAX`).
    Max = 4,
    /// Smallest value; subtotal function number 105 (`ILTA_MIN`).
    Min = 5,
    /// Arithmetic sum; subtotal function number 109 (`ILTA_SUM`).
    Sum = 6,
    /// Estimated standard deviation; subtotal function number 107
    /// (`ILTA_STDDEV`).
    StdDev = 7,
    /// Estimated variance; subtotal function number 110 (`ILTA_VAR`).
    Var = 8,
    /// Custom formula carried by `BrtListTrFmla` (`ILTA_CUSTOM`).
    Custom = 9,
}

impl TryFrom<u32> for TotalsRowFunction {
    type Error = XlsbError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Average),
            2 => Ok(Self::Count),
            3 => Ok(Self::CountNums),
            4 => Ok(Self::Max),
            5 => Ok(Self::Min),
            6 => Ok(Self::Sum),
            7 => Ok(Self::StdDev),
            8 => Ok(Self::Var),
            9 => Ok(Self::Custom),
            _ => Err(XlsbError::Unrecognized {
                typ: "table totals-row function".to_string(),
                val: format!("0x{value:08X}"),
            }),
        }
    }
}

/// A column of a table (`BrtBeginListCol` collection, MS-XLSB 2.4.101).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Column {
    /// Numeric identifier of the column (`idField`); unique within the table.
    pub id: u32,
    /// Total row aggregation function (`ilta`).
    pub totals_row_function: TotalsRowFunction,
    /// Differential formatting of the column header (`nDxfHdr`); `None` = none.
    pub header_dxf_id: Option<u32>,
    /// Differential formatting of the column insert row (`nDxfInsertRow`);
    /// `None` = none.
    pub insert_row_dxf_id: Option<u32>,
    /// Differential formatting of the column total row (`nDxfAgg`);
    /// `None` = none.
    pub totals_dxf_id: Option<u32>,
    /// Query table column identifier (`idqsif`); `0` = no query table column.
    /// Inert; never resolved.
    pub query_table_field_id: u32,
    /// Textual identifier of the column (`stName`).
    pub name: Option<String>,
    /// Caption displayed in the sheet (`stCaption`).
    pub caption: Option<String>,
    /// Text displayed in the total row of the column (`stTotal`); inert.
    pub totals_row_label: Option<String>,
    /// Cell style applied to the column header (`stStyleHeader`).
    pub header_style: Option<String>,
    /// Cell style applied to the column insert row (`stStyleInsertRow`).
    pub insert_row_style: Option<String>,
    /// Cell style applied to the column total row (`stStyleAgg`).
    pub totals_style: Option<String>,
    /// Calculated column formula (`BrtListCCFmla`, MS-XLSB 2.4.706), stored
    /// as raw tokens; never evaluated.
    pub calculated_column_formula: Option<Formula>,
    /// Total row formula (`BrtListTrFmla`, MS-XLSB 2.4.708), stored as raw
    /// tokens; never evaluated.
    pub totals_row_formula: Option<Formula>,
}

/// A `ListParsedFormula` (MS-XLSB 2.5.98.11) stored verbatim; never evaluated.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Formula {
    /// The formula is an array formula (`fArray`).
    pub array: bool,
    /// Ptg token bytes (`rgce`).
    pub tokens: Vec<u8>,
    /// Ancillary bytes (`rgcb`).
    pub extra: Vec<u8>,
}

/// Table style applied to the table (`BrtTableStyleClient`, MS-XLSB 2.4.847).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleInfo {
    /// Name of the table style (`stStyleName`); a built-in style name or the
    /// `strName` of a `BrtBeginTableStyle` record in the Styles part.
    pub name: Option<String>,
    /// First-column table style elements are applied (`fFirstColumn`).
    pub show_first_column: bool,
    /// Last-column table style elements are applied (`fLastColumn`).
    pub show_last_column: bool,
    /// Row-stripe table style elements are applied (`fRowStripes`).
    pub show_row_stripes: bool,
    /// Column-stripe table style elements are applied (`fColumnStripes`).
    pub show_column_stripes: bool,
}
