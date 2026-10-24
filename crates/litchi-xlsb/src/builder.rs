//! Typed workbook authoring, finalized as a validated package.
//!
//! [`Builder`] authors fresh content only: its [`Worksheet`] and
//! [`ChartSheet`] drafts stage cells, formulas, styles, drawings, charts and
//! the rest of their verbs, and [`Builder::build`] materializes them into a
//! validated [`crate::Package`]. The drafts reached from it live in this
//! module without a prefix; worksheet-chart drafts live under
//! [`chart`] (record 0351 §2.1 rule 7). An opened package
//! never hydrates into this writer: it changes through
//! [`crate::Package::edit`] and its scoped `edit_*` families (ADR 0034 §9).
//!
//! ```
//! use litchi_core::sheet::WorkbookTrait as _;
//! use litchi_xlsb::builder::{Builder, Worksheet};
//!
//! # fn main() -> litchi_xlsb::package::PackageResult<()> {
//! let mut builder = Builder::new();
//! let mut sheet = Worksheet::new("Sheet1");
//! sheet.set_cell(0, 0, "Hello");
//! sheet.set_cell(0, 1, 42.0);
//! builder.add_worksheet(sheet);
//! let package = builder.build()?;
//! assert_eq!(package.workbook()?.worksheet_names(), &["Sheet1"]);
//! # Ok(())
//! # }
//! ```

pub use crate::writer::{
    Anchor, AreaReference, Autofit, Body, Builder, CachedValue, CellData, CellLocation, CellMarker,
    CellReference, ChartSheet, Columns, ConnectionEndSpec, ConnectionShapeSpec, Coordinate32,
    DdeItem, DefinedName, Direction, DxfStyle, EditAs, Emu, EmuExtent, EmuOffset, ErrorValue,
    Geometry, GroupSpec, GroupTransform, Image, ImageFormat, Insets, Kind, Link,
    MAX_XLSB_EXTERNAL_CACHE_COLUMNS, MAX_XLSB_EXTERNAL_CACHE_ROWS, MAX_XLSB_EXTERNAL_CACHED_VALUES,
    NameFormula, NameFormulaKind, ObjectSpec, OleItem, Paragraph, Part, Preset, Properties, Run,
    ShapeSpec, SheetProtection, SheetRange, StylesWriter, TextSize, Underline, ValueMatrix,
    VerticalAnchor, Worksheet, Wrap,
};

/// Shared worksheet-chart authoring values.
pub mod chart {
    pub use litchi_spreadsheet_drawing::chart::builder::{
        Anchor, Chart, ExternalDataPart, ExternalDataTarget, Relationship, Series, Target,
        UserShapesPart,
    };
}
