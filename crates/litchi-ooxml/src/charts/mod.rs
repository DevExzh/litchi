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
//! ```rust,ignore
//! use litchi::ooxml::charts::{Chart, PlotArea, Series, NumericData, StringData};
//! use litchi::ooxml::charts::plot_area::{BarTypeGroup, TypeGroup};
//! use litchi::ooxml::charts::types::{BarDirection, BarGrouping};
//! use litchi::ooxml::charts::axis::{CategoryAxis, ValueAxis, Axis};
//! use litchi::ooxml::charts::types::AxisPosition;
//! use litchi::ooxml::charts::legend::Legend;
//! use litchi::ooxml::charts::types::LegendPosition;
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
//! litchi::ooxml::charts::writer::write_chart(&mut xml_output, &chart)?;
//! ```

pub mod axis;
pub mod chart;
pub mod legend;
pub mod models;
pub mod plot_area;
pub mod reader;
pub mod series;
pub mod types;
pub mod writer;

pub use axis::{Axis, AxisCommon, CategoryAxis, DateAxis, SeriesAxis, ValueAxis};
pub use chart::{Chart, View3D, WallFloor};
pub use legend::Legend;
pub use models::{
    Layout, MultiLevelStringData, NumberFormat, NumericData, RichText, StringData, TitleText,
};
pub use plot_area::{
    Area3DTypeGroup, AreaTypeGroup, Bar3DTypeGroup, BarTypeGroup, BubbleTypeGroup,
    DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea,
    RadarTypeGroup, ScatterTypeGroup, StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup,
    TypeGroup,
};
pub use series::{DataLabels, DataPoint, Series};
pub use types::ChartType;
