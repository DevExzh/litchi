//! Chart plot area models.
//!
//! This module contains structures for representing the plot area
//! and chart type groups.

use crate::ooxml::charts::axis::Axis;
use crate::ooxml::charts::models::Layout;
use crate::ooxml::charts::series::Series;
use crate::ooxml::charts::types::{BarDirection, BarGrouping, RadarStyle, ScatterStyle};

/// Plot area containing chart data and axes.
#[derive(Debug, Clone)]
pub struct PlotArea {
    /// Manual layout
    pub layout: Option<Layout>,
    /// Chart type groups
    pub type_groups: Vec<TypeGroup>,
    /// All axes in the plot area
    pub axes: Vec<Axis>,
}

impl PlotArea {
    /// Create a new plot area.
    #[inline]
    pub fn new() -> Self {
        Self {
            layout: None,
            type_groups: Vec::new(),
            axes: Vec::new(),
        }
    }

    /// Add a type group.
    #[inline]
    pub fn add_type_group(mut self, group: TypeGroup) -> Self {
        self.type_groups.push(group);
        self
    }

    /// Add an axis.
    #[inline]
    pub fn add_axis(mut self, axis: Axis) -> Self {
        self.axes.push(axis);
        self
    }
}

impl Default for PlotArea {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// A group of series with the same chart type.
#[derive(Debug, Clone)]
pub enum TypeGroup {
    /// Area chart
    Area(AreaTypeGroup),
    /// Area 3D chart
    Area3D(Area3DTypeGroup),
    /// Bar chart
    Bar(BarTypeGroup),
    /// Bar 3D chart
    Bar3D(Bar3DTypeGroup),
    /// Bubble chart
    Bubble(BubbleTypeGroup),
    /// Doughnut chart
    Doughnut(DoughnutTypeGroup),
    /// Line chart
    Line(LineTypeGroup),
    /// Line 3D chart
    Line3D(Line3DTypeGroup),
    /// Pie chart
    Pie(PieTypeGroup),
    /// Pie 3D chart
    Pie3D(Pie3DTypeGroup),
    /// Radar chart
    Radar(RadarTypeGroup),
    /// Scatter chart
    Scatter(ScatterTypeGroup),
    /// Stock chart
    Stock(StockTypeGroup),
    /// Surface chart
    Surface(SurfaceTypeGroup),
    /// Surface 3D chart
    Surface3D(Surface3DTypeGroup),
}

/// Common properties for type groups.
#[derive(Debug, Clone)]
pub struct TypeGroupCommon {
    /// Vary colors by point
    pub vary_colors: bool,
    /// Series in this group
    pub series: Vec<Series>,
}

impl TypeGroupCommon {
    /// Create new common properties.
    #[inline]
    pub fn new() -> Self {
        Self {
            vary_colors: false,
            series: Vec::new(),
        }
    }

    /// Add a series.
    #[inline]
    pub fn add_series(&mut self, series: Series) {
        self.series.push(series);
    }
}

impl Default for TypeGroupCommon {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Area chart type group.
#[derive(Debug, Clone)]
pub struct AreaTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Grouping type
    pub grouping: BarGrouping,
}

impl AreaTypeGroup {
    /// Create a new area type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
        }
    }
}

/// Area 3D chart type group.
#[derive(Debug, Clone)]
pub struct Area3DTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Grouping type
    pub grouping: BarGrouping,
}

impl Area3DTypeGroup {
    /// Create a new area 3D type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
        }
    }
}

/// Bar chart type group.
#[derive(Debug, Clone)]
pub struct BarTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Bar direction
    pub direction: BarDirection,
    /// Grouping type
    pub grouping: BarGrouping,
    /// Gap width (0-500%)
    pub gap_width: Option<u32>,
    /// Overlap (-100% to 100%)
    pub overlap: Option<i32>,
}

impl BarTypeGroup {
    /// Create a new bar type group.
    #[inline]
    pub fn new(direction: BarDirection, grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            direction,
            grouping,
            gap_width: None,
            overlap: None,
        }
    }
}

/// Bar 3D chart type group.
#[derive(Debug, Clone)]
pub struct Bar3DTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Bar direction
    pub direction: BarDirection,
    /// Grouping type
    pub grouping: BarGrouping,
    /// Gap width (0-500%)
    pub gap_width: Option<u32>,
    /// Gap depth (0-500%)
    pub gap_depth: Option<u32>,
    /// Shape type
    pub shape: Option<BarShape>,
}

/// 3D bar shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarShape {
    /// Box shape
    Box,
    /// Cone shape
    Cone,
    /// Cone to maximum
    ConeToMax,
    /// Cylinder shape
    Cylinder,
    /// Pyramid shape
    Pyramid,
    /// Pyramid to maximum
    PyramidToMax,
}

impl BarShape {
    /// Returns the XML value for this shape.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Box => "box",
            Self::Cone => "cone",
            Self::ConeToMax => "coneToMax",
            Self::Cylinder => "cylinder",
            Self::Pyramid => "pyramid",
            Self::PyramidToMax => "pyramidToMax",
        }
    }
}

impl Bar3DTypeGroup {
    /// Create a new bar 3D type group.
    #[inline]
    pub fn new(direction: BarDirection, grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            direction,
            grouping,
            gap_width: None,
            gap_depth: None,
            shape: None,
        }
    }
}

/// Bubble chart type group.
#[derive(Debug, Clone)]
pub struct BubbleTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Show bubbles in 3D
    pub bubble_3d: bool,
    /// Bubble scale (0-300%)
    pub bubble_scale: Option<u32>,
    /// Show negative bubbles
    pub show_negative_bubbles: bool,
    /// Size represents area or width ("area" or "w", default "area")
    pub size_represents: String,
}

impl BubbleTypeGroup {
    /// Create a new bubble type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            bubble_3d: false,
            bubble_scale: None,
            show_negative_bubbles: true,
            size_represents: "area".to_string(),
        }
    }
}

impl Default for BubbleTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Doughnut chart type group.
#[derive(Debug, Clone)]
pub struct DoughnutTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// First slice angle (0-360 degrees)
    pub first_slice_angle: u32,
    /// Hole size (10-90%)
    pub hole_size: u32,
}

impl DoughnutTypeGroup {
    /// Create a new doughnut type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            first_slice_angle: 0,
            hole_size: 50,
        }
    }
}

impl Default for DoughnutTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Line chart type group.
#[derive(Debug, Clone)]
pub struct LineTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Grouping type
    pub grouping: BarGrouping,
    /// Show markers
    pub marker: bool,
}

impl LineTypeGroup {
    /// Create a new line type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
            marker: true,
        }
    }
}

/// Line 3D chart type group.
#[derive(Debug, Clone)]
pub struct Line3DTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Grouping type
    pub grouping: BarGrouping,
}

impl Line3DTypeGroup {
    /// Create a new line 3D type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
        }
    }
}

/// Pie chart type group.
#[derive(Debug, Clone)]
pub struct PieTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// First slice angle (0-360 degrees)
    pub first_slice_angle: u32,
}

impl PieTypeGroup {
    /// Create a new pie type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            first_slice_angle: 0,
        }
    }
}

impl Default for PieTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Pie 3D chart type group.
#[derive(Debug, Clone)]
pub struct Pie3DTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
}

impl Pie3DTypeGroup {
    /// Create a new pie 3D type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
        }
    }
}

impl Default for Pie3DTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Radar chart type group.
#[derive(Debug, Clone)]
pub struct RadarTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Radar style
    pub style: RadarStyle,
}

impl RadarTypeGroup {
    /// Create a new radar type group.
    #[inline]
    pub fn new(style: RadarStyle) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            style,
        }
    }
}

/// Scatter chart type group.
#[derive(Debug, Clone)]
pub struct ScatterTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Scatter style
    pub style: ScatterStyle,
}

impl ScatterTypeGroup {
    /// Create a new scatter type group.
    #[inline]
    pub fn new(style: ScatterStyle) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            style,
        }
    }
}

/// Stock chart type group.
#[derive(Debug, Clone)]
pub struct StockTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
}

impl StockTypeGroup {
    /// Create a new stock type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
        }
    }
}

impl Default for StockTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Surface chart type group.
#[derive(Debug, Clone)]
pub struct SurfaceTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Wireframe mode
    pub wireframe: bool,
}

impl SurfaceTypeGroup {
    /// Create a new surface type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            wireframe: false,
        }
    }
}

impl Default for SurfaceTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Surface 3D chart type group.
#[derive(Debug, Clone)]
pub struct Surface3DTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Wireframe mode
    pub wireframe: bool,
}

impl Surface3DTypeGroup {
    /// Create a new surface 3D type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            wireframe: false,
        }
    }
}

impl Default for Surface3DTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}
