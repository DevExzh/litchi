//! Core chart types and enumerations.
//!
//! This module defines all chart types, axis types, grouping types, and other
//! enumerations used throughout the chart API.

use std::fmt;

/// Chart type enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartType {
    /// Area chart (2D)
    Area,
    /// Area chart (3D)
    Area3D,
    /// Bar chart (horizontal bars)
    Bar,
    /// Bar chart (3D)
    Bar3D,
    /// Bubble chart
    Bubble,
    /// Doughnut chart
    Doughnut,
    /// Line chart
    Line,
    /// Line chart (3D)
    Line3D,
    /// Pie chart
    Pie,
    /// Pie chart (3D)
    Pie3D,
    /// Radar chart
    Radar,
    /// Scatter (XY) chart
    Scatter,
    /// Stock chart
    Stock,
    /// Surface chart
    Surface,
    /// Surface chart (3D)
    Surface3D,
    /// Unknown or unsupported chart type
    Unknown,
}

impl ChartType {
    /// Returns the XML element name for this chart type.
    #[inline]
    pub fn xml_element_name(&self) -> &'static str {
        match self {
            Self::Area => "areaChart",
            Self::Area3D => "area3DChart",
            Self::Bar => "barChart",
            Self::Bar3D => "bar3DChart",
            Self::Bubble => "bubbleChart",
            Self::Doughnut => "doughnutChart",
            Self::Line => "lineChart",
            Self::Line3D => "line3DChart",
            Self::Pie => "pieChart",
            Self::Pie3D => "pie3DChart",
            Self::Radar => "radarChart",
            Self::Scatter => "scatterChart",
            Self::Stock => "stockChart",
            Self::Surface => "surfaceChart",
            Self::Surface3D => "surface3DChart",
            Self::Unknown => "unknownChart",
        }
    }

    /// Returns true if this is a 3D chart type.
    #[inline]
    pub const fn is_3d(&self) -> bool {
        matches!(
            self,
            Self::Area3D | Self::Bar3D | Self::Line3D | Self::Pie3D | Self::Surface3D
        )
    }

    /// Returns true if this chart type supports categories.
    #[inline]
    pub const fn supports_categories(&self) -> bool {
        !matches!(self, Self::Scatter | Self::Bubble)
    }
}

impl fmt::Display for ChartType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.xml_element_name())
    }
}

/// Axis type identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AxisType {
    /// Category axis
    Category,
    /// Value axis
    Value,
    /// Date axis
    Date,
    /// Series axis (for 3D charts)
    Series,
}

impl AxisType {
    /// Returns the XML element name for this axis type.
    #[inline]
    pub const fn xml_element_name(&self) -> &'static str {
        match self {
            Self::Category => "catAx",
            Self::Value => "valAx",
            Self::Date => "dateAx",
            Self::Series => "serAx",
        }
    }
}

/// Axis position.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisPosition {
    /// Bottom position
    Bottom,
    /// Left position
    Left,
    /// Right position
    Right,
    /// Top position
    Top,
}

impl AxisPosition {
    /// Returns the XML value for this position.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Bottom => "b",
            Self::Left => "l",
            Self::Right => "r",
            Self::Top => "t",
        }
    }
}

/// Axis orientation (min to max or max to min).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisOrientation {
    /// Min to max (normal)
    MinMax,
    /// Max to min (reversed)
    MaxMin,
}

impl AxisOrientation {
    /// Returns the XML value for this orientation.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::MinMax => "minMax",
            Self::MaxMin => "maxMin",
        }
    }
}

/// Tick mark style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TickMark {
    /// Cross tick mark
    Cross,
    /// Inside tick mark
    In,
    /// No tick mark
    None,
    /// Outside tick mark
    Out,
}

impl TickMark {
    /// Returns the XML value for this tick mark style.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Cross => "cross",
            Self::In => "in",
            Self::None => "none",
            Self::Out => "out",
        }
    }
}

/// Tick label position.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TickLabelPosition {
    /// High position
    High,
    /// Low position
    Low,
    /// Next to axis
    NextTo,
    /// No tick labels
    None,
}

impl TickLabelPosition {
    /// Returns the XML value for this position.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::High => "high",
            Self::Low => "low",
            Self::NextTo => "nextTo",
            Self::None => "none",
        }
    }
}

/// Bar/column direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarDirection {
    /// Horizontal bars
    Bar,
    /// Vertical bars (columns)
    Column,
}

impl BarDirection {
    /// Returns the XML value for this direction.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Bar => "bar",
            Self::Column => "col",
        }
    }
}

/// Bar grouping type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarGrouping {
    /// Clustered bars
    Clustered,
    /// Stacked bars
    Stacked,
    /// 100% stacked bars
    PercentStacked,
    /// Standard grouping
    Standard,
}

impl BarGrouping {
    /// Returns the XML value for this grouping.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Clustered => "clustered",
            Self::Stacked => "stacked",
            Self::PercentStacked => "percentStacked",
            Self::Standard => "standard",
        }
    }
}

/// Legend position.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegendPosition {
    /// Bottom position
    Bottom,
    /// Left position
    Left,
    /// Right position
    Right,
    /// Top position
    Top,
    /// Top right corner
    TopRight,
}

impl LegendPosition {
    /// Returns the XML value for this position.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Bottom => "b",
            Self::Left => "l",
            Self::Right => "r",
            Self::Top => "t",
            Self::TopRight => "tr",
        }
    }
}

/// Data label position.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataLabelPosition {
    /// Best fit position
    BestFit,
    /// Center position
    Center,
    /// Inside base position
    InsideBase,
    /// Inside end position
    InsideEnd,
    /// Left position
    Left,
    /// Outside end position
    OutsideEnd,
    /// Right position
    Right,
    /// Top position
    Top,
    /// Bottom position
    Bottom,
}

impl DataLabelPosition {
    /// Returns the XML value for this position.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::BestFit => "bestFit",
            Self::Center => "ctr",
            Self::InsideBase => "inBase",
            Self::InsideEnd => "inEnd",
            Self::Left => "l",
            Self::OutsideEnd => "outEnd",
            Self::Right => "r",
            Self::Top => "t",
            Self::Bottom => "b",
        }
    }
}

/// Marker style for line and scatter charts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerStyle {
    /// Circle marker
    Circle,
    /// Dash marker
    Dash,
    /// Diamond marker
    Diamond,
    /// Dot marker
    Dot,
    /// No marker
    None,
    /// Picture marker
    Picture,
    /// Plus marker
    Plus,
    /// Square marker
    Square,
    /// Star marker
    Star,
    /// Triangle marker
    Triangle,
    /// X marker
    X,
    /// Automatic marker
    Auto,
}

impl MarkerStyle {
    /// Returns the XML value for this marker style.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Circle => "circle",
            Self::Dash => "dash",
            Self::Diamond => "diamond",
            Self::Dot => "dot",
            Self::None => "none",
            Self::Picture => "picture",
            Self::Plus => "plus",
            Self::Square => "square",
            Self::Star => "star",
            Self::Triangle => "triangle",
            Self::X => "x",
            Self::Auto => "auto",
        }
    }
}

/// Scatter chart style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScatterStyle {
    /// Line with markers
    LineMarker,
    /// Line only
    Line,
    /// Markers only
    Marker,
    /// No line or markers
    None,
    /// Smooth line with markers
    SmoothMarker,
    /// Smooth line only
    Smooth,
}

impl ScatterStyle {
    /// Returns the XML value for this scatter style.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::LineMarker => "lineMarker",
            Self::Line => "line",
            Self::Marker => "marker",
            Self::None => "none",
            Self::SmoothMarker => "smoothMarker",
            Self::Smooth => "smooth",
        }
    }
}

/// Radar chart style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RadarStyle {
    /// Standard radar
    Standard,
    /// Filled radar
    Filled,
    /// Marker radar
    Marker,
}

impl RadarStyle {
    /// Returns the XML value for this radar style.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Standard => "standard",
            Self::Filled => "filled",
            Self::Marker => "marker",
        }
    }
}

/// Display blanks mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DisplayBlanks {
    /// Gaps where there are blank cells
    Gap,
    /// Connect data points across blank cells
    Span,
    /// Treat blank cells as zero
    Zero,
}

impl DisplayBlanks {
    /// Returns the XML value for this display mode.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Gap => "gap",
            Self::Span => "span",
            Self::Zero => "zero",
        }
    }
}

/// Layout mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LayoutMode {
    /// Edge layout mode
    Edge,
    /// Factor layout mode
    Factor,
}

impl LayoutMode {
    /// Returns the XML value for this layout mode.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Edge => "edge",
            Self::Factor => "factor",
        }
    }
}

/// Layout target.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LayoutTarget {
    /// Inner target (plot area)
    Inner,
    /// Outer target (chart area)
    Outer,
}

impl LayoutTarget {
    /// Returns the XML value for this layout target.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Inner => "inner",
            Self::Outer => "outer",
        }
    }
}
