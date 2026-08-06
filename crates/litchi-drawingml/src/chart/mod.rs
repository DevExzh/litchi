//! Chart support for Office Open XML formats.
//!
//! This module provides comprehensive support for reading and writing charts
//! in OOXML documents (XLSX, DOCX, PPTX). It includes:
//!
//! - Core chart types and enumerations
//! - Data models for series, axes, and legends
//! - XML readers and writers
//! - Full support for all major chart types
//!
//! Core chart XML follows `[MS-ODRAWXML]` chart structures. Package parts and
//! relationship graphs remain in the owning DOCX/PPTX/XLSX/XLSB facades under
//! `[MS-OI29500]`; host-specific placement is documented by `[MS-PPTX]`,
//! `[MS-XLSX]`, and `[MS-XLSB]`.
//!
//! # Chart Types Supported
//!
//! - Area charts (2D and 3D)
//! - Bar/Column charts (2D and 3D)
//! - Line charts (2D and 3D)
//! - Pie charts (2D and 3D)
//! - Scatter/XY charts
//! - Bubble charts
//! - Doughnut charts
//! - Radar charts
//! - Stock charts
//! - Surface charts (2D and 3D)
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_drawingml::chart::{Chart, PlotArea, Series, NumericData, StringData};
//! use litchi_drawingml::chart::plot_area::{BarTypeGroup, TypeGroup};
//! use litchi_drawingml::chart::types::{BarDirection, BarGrouping};
//! use litchi_drawingml::chart::axis::{CategoryAxis, ValueAxis, Axis};
//! use litchi_drawingml::chart::types::AxisPosition;
//! use litchi_drawingml::chart::legend::Legend;
//! use litchi_drawingml::chart::types::LegendPosition;
//!
//! // Create a bar chart
//! let mut chart = Chart::new()
//!     .with_title("Sales Report")
//!     .with_legend(Legend::new(LegendPosition::Right));
//!
//! // Create series
//! let series = Series::new(0)
//!     .with_title("Q1 Sales")
//!     .with_categories(StringData::from_values(vec![
//!         "Jan".to_string(),
//!         "Feb".to_string(),
//!         "Mar".to_string(),
//!     ]))
//!     .with_values(NumericData::from_values(vec![100.0, 150.0, 200.0]));
//!
//! // Create bar chart type group
//! let mut bar_group = BarTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
//! bar_group.common.add_series(series);
//!
//! // Create axes
//! let cat_axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
//! let val_axis = ValueAxis::new(2, AxisPosition::Left, 1);
//!
//! // Build plot area
//! let plot_area = PlotArea::new()
//!     .add_type_group(TypeGroup::Bar(bar_group))
//!     .add_axis(Axis::Category(cat_axis))
//!     .add_axis(Axis::Value(val_axis));
//!
//! chart.plot_area = plot_area;
//!
//! // Write to XML
//! let mut xml_output = Vec::new();
//! litchi_drawingml::chart::writer::write(&mut xml_output, &chart)?;
//! # Ok::<(), std::io::Error>(())
//! ```

pub mod axis;
pub mod bubble;
pub mod data;
pub mod edit;
pub mod legend;
pub mod model;
pub mod plot_area;
pub mod reader;
pub mod series;
pub mod style;
pub mod types;
pub mod writer;

pub use axis::{Axis, AxisCommon, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
pub use data::{
    Layout, MultiLevelStringData, NumberFormat, NumericData, RichText, StringData, TitleText,
};
pub use edit::{Commit, DataLabelFlag, Patch, Snapshot, Transaction};
pub use legend::Legend;
pub use model::{
    Chart, ColorMapOverride, ColorMapping, ColorSchemeIndex, ExtensionList, ExternalData,
    HeaderFooter, PageMargins, PageOrientation, PageSetup, PictureFormat, PictureOptions,
    PivotFormat, PivotSource, PrintSettings, Protection, ShapeProperties, TextProperties,
    UserShapes, View3D, WallFloor,
};
pub use plot_area::{
    Area3DTypeGroup, AreaTypeGroup, BandFormat, Bar3DTypeGroup, BarShape, BarTypeGroup,
    BubbleTypeGroup, DataTable, DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Lines,
    OfPieTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea, RadarTypeGroup, ScatterTypeGroup,
    StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup, TypeGroup, UpDownBars,
};
pub use series::{DataLabel, DataLabels, DataPoint, Marker, Series};
pub use types::{OfPieSplitType, OfPieType, Type};
