//! Typed semantic values for the BIFF8 `WebPub` record.

use crate::{Error, Result};

use super::invalid;

/// The kind of Web source that was published (`WebPub.tws`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebSourceType {
    /// The source is undefined.
    Undefined,
    /// The whole workbook.
    Workbook,
    /// An entire sheet.
    Sheet,
    /// A print area.
    PrintArea,
    /// An `AutoFilter` range.
    AutoFilter,
    /// A range of cells; the record's `FrtRefHeaderU.ref8` applies.
    Range,
    /// A chart; the record carries the chart's shape identifier.
    Chart,
    /// A `PivotTable` report.
    PivotTable,
    /// A query table (external data range).
    QueryTable,
    /// A named range.
    NamedRange,
}

impl WebSourceType {
    pub(super) fn from_code(code: u8) -> Result<Self> {
        Ok(match code {
            0xFF => Self::Undefined,
            0x00 => Self::Workbook,
            0x01 => Self::Sheet,
            0x02 => Self::PrintArea,
            0x03 => Self::AutoFilter,
            0x04 => Self::Range,
            0x05 => Self::Chart,
            0x06 => Self::PivotTable,
            0x07 => Self::QueryTable,
            0x08 => Self::NamedRange,
            other => return Err(invalid(format!("unknown WebPub tws value 0x{other:02X}"))),
        })
    }

    /// Raw `tws` code; governs the conditional `srcName`/`crtID` fields.
    pub(super) const fn code(self) -> u8 {
        match self {
            Self::Undefined => 0xFF,
            Self::Workbook => 0x00,
            Self::Sheet => 0x01,
            Self::PrintArea => 0x02,
            Self::AutoFilter => 0x03,
            Self::Range => 0x04,
            Self::Chart => 0x05,
            Self::PivotTable => 0x06,
            Self::QueryTable => 0x07,
            Self::NamedRange => 0x08,
        }
    }
}

/// The kind of Web page created for a published item (`WebPub.twd`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebPageType {
    /// A non-interactive page, only for viewing.
    ViewOnly,
    /// An interactive page using workbook functionality.
    WorkbookFunctionality,
    /// An interactive page using `PivotTable` functionality.
    PivotTableFunctionality,
    /// An interactive page using chart functionality.
    ChartFunctionality,
}

impl WebPageType {
    pub(super) fn from_code(code: u8) -> Result<Self> {
        Ok(match code {
            0x00 => Self::ViewOnly,
            0x01 => Self::WorkbookFunctionality,
            0x02 => Self::PivotTableFunctionality,
            0x03 => Self::ChartFunctionality,
            other => return Err(invalid(format!("unknown WebPub twd value 0x{other:02X}"))),
        })
    }

    /// Raw `twd` code.
    pub(super) const fn code(self) -> u8 {
        match self {
            Self::ViewOnly => 0x00,
            Self::WorkbookFunctionality => 0x01,
            Self::PivotTableFunctionality => 0x02,
            Self::ChartFunctionality => 0x03,
        }
    }
}

/// The cell range a `WebPub` record publishes (`FrtRefHeaderU.ref8`),
/// present only when the source type is [`WebSourceType::Range`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WebPubRange {
    /// First row of the range.
    first_row: u16,
    /// Last row of the range.
    last_row: u16,
    /// First column of the range.
    first_column: u8,
    /// Last column of the range.
    last_column: u8,
}

impl WebPubRange {
    /// Create a checked, inclusive BIFF8 publication range.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(first_row: u32, last_row: u32, first_column: u16, last_column: u16) -> Result<Self> {
        let invalid = || {
            Error::InvalidCellReference(format!(
                "WebPub range ({first_row}, {first_column})..=({last_row}, {last_column}) is outside the BIFF8 grid"
            ))
        };
        let first_row = u16::try_from(first_row).map_err(|_error| invalid())?;
        let last_row = u16::try_from(last_row).map_err(|_error| invalid())?;
        let first_column = u8::try_from(first_column).map_err(|_error| invalid())?;
        let last_column = u8::try_from(last_column).map_err(|_error| invalid())?;
        if first_row > last_row || first_column > last_column {
            return Err(invalid());
        }
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }

    pub(super) fn decode(
        first_row: u16,
        last_row: u16,
        first_column: u16,
        last_column: u16,
    ) -> Result<Self> {
        Self::new(
            u32::from(first_row),
            u32::from(last_row),
            first_column,
            last_column,
        )
        .map_err(|_error| invalid("WebPub range is outside the BIFF8 grid"))
    }

    #[must_use]
    pub const fn first_row(self) -> u16 {
        self.first_row
    }

    #[must_use]
    pub const fn last_row(self) -> u16 {
        self.last_row
    }

    #[must_use]
    pub const fn first_column(self) -> u8 {
        self.first_column
    }

    #[must_use]
    pub const fn last_column(self) -> u8 {
        self.last_column
    }

    pub(super) const fn fields(self) -> (u16, u16, u8, u8) {
        (
            self.first_row,
            self.last_row,
            self.first_column,
            self.last_column,
        )
    }
}

/// Typed `WebPub` record content (MS-XLS 2.4.344).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebPub {
    /// What was published (`tws`).
    pub source: WebSourceType,
    /// The kind of Web page created (`twd`).
    pub page_type: WebPageType,
    /// The published cell range, present iff `source` is
    /// [`WebSourceType::Range`].
    pub range: Option<WebPubRange>,
    /// Whether the page is republished when the workbook is saved
    /// (`fAutoRepublish`).
    pub auto_republish: bool,
    /// Whether the page is published as a single Web page (MHTML) rather
    /// than a page with references to other files (`fMhtml`).
    pub single_file: bool,
    /// Unique identifier of the published content (`nStyleId`).
    pub style_id: u32,
    /// The named range to publish (`srcName`), present iff the `tws` code
    /// is greater than 4.
    pub source_name: Option<String>,
    /// URL or path of the published page (`stFileDest`).
    pub file_destination: String,
    /// Destination bookmark of the published page (`stDivId`).
    pub div_id: String,
    /// Title of the published item (`stTitle`).
    pub title: String,
    /// Shape identifier of the published chart object (`crtID`), present
    /// iff `source` is [`WebSourceType::Chart`].
    pub chart_shape_id: Option<u32>,
    /// Bytes reserved for future use (`frtRgb`), preserved verbatim.
    pub reserved: Vec<u8>,
}
