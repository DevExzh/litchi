//! Chart plot area models.
//!
//! This module contains structures for representing the plot area
//! and chart type groups.

use crate::chart::axis::Axis;
use crate::chart::bubble::{Scale as BubbleScale, Size as BubbleSize};
use crate::chart::data::Layout;
use crate::chart::model::{ChartExtensionList, ChartShapeProperties, ChartTextProperties};
use crate::chart::series::{DataLabels, Series};
use crate::chart::types::{
    BarDirection, BarGrouping, OfPieSplitType, OfPieType, RadarStyle, ScatterStyle,
};

/// Plot area containing chart data and axes.
#[derive(Debug, Clone)]
pub struct PlotArea {
    /// Manual layout
    pub layout: Option<Layout>,
    /// Chart type groups
    pub type_groups: Vec<TypeGroup>,
    /// All axes in the plot area
    pub axes: Vec<Axis>,
    /// Optional chart data table
    pub data_table: Option<DataTable>,
    /// DrawingML shape properties for the plot area
    pub shape_properties: Option<ChartShapeProperties>,
    /// Plot-area extension list
    pub extension_list: Option<ChartExtensionList>,
}

impl PlotArea {
    /// Create a new plot area.
    #[inline]
    pub fn new() -> Self {
        Self {
            layout: None,
            type_groups: Vec::new(),
            axes: Vec::new(),
            data_table: None,
            shape_properties: None,
            extension_list: None,
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

/// Visibility settings for the chart data table.
#[derive(Debug, Clone, Default)]
pub struct DataTable {
    /// Show horizontal cell borders
    pub show_horizontal_border: bool,
    /// Show vertical cell borders
    pub show_vertical_border: bool,
    /// Show the outside border
    pub show_outline: bool,
    /// Show legend keys beside series rows
    pub show_legend_keys: bool,
    /// DrawingML shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// DrawingML text properties
    pub text_properties: Option<ChartTextProperties>,
    /// Data-table extension list
    pub extension_list: Option<ChartExtensionList>,
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
    /// Pie-of-pie or bar-of-pie chart
    OfPie(OfPieTypeGroup),
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

impl TypeGroup {
    /// Return the properties shared by every classic chart-type group.
    pub fn common(&self) -> &TypeGroupCommon {
        match self {
            Self::Area(group) => &group.common,
            Self::Area3D(group) => &group.common,
            Self::Bar(group) => &group.common,
            Self::Bar3D(group) => &group.common,
            Self::Bubble(group) => &group.common,
            Self::Doughnut(group) => &group.common,
            Self::Line(group) => &group.common,
            Self::Line3D(group) => &group.common,
            Self::OfPie(group) => &group.common,
            Self::Pie(group) => &group.common,
            Self::Pie3D(group) => &group.common,
            Self::Radar(group) => &group.common,
            Self::Scatter(group) => &group.common,
            Self::Stock(group) => &group.common,
            Self::Surface(group) => &group.common,
            Self::Surface3D(group) => &group.common,
        }
    }

    /// Return mutable properties shared by every classic chart-type group.
    pub fn common_mut(&mut self) -> &mut TypeGroupCommon {
        match self {
            Self::Area(group) => &mut group.common,
            Self::Area3D(group) => &mut group.common,
            Self::Bar(group) => &mut group.common,
            Self::Bar3D(group) => &mut group.common,
            Self::Bubble(group) => &mut group.common,
            Self::Doughnut(group) => &mut group.common,
            Self::Line(group) => &mut group.common,
            Self::Line3D(group) => &mut group.common,
            Self::OfPie(group) => &mut group.common,
            Self::Pie(group) => &mut group.common,
            Self::Pie3D(group) => &mut group.common,
            Self::Radar(group) => &mut group.common,
            Self::Scatter(group) => &mut group.common,
            Self::Stock(group) => &mut group.common,
            Self::Surface(group) => &mut group.common,
            Self::Surface3D(group) => &mut group.common,
        }
    }
}

/// Common properties for type groups.
#[derive(Debug, Clone)]
pub struct TypeGroupCommon {
    /// Vary colors by point
    pub vary_colors: bool,
    /// Series in this group
    pub series: Vec<Series>,
    /// Data-label settings shared by the chart group
    pub data_labels: Option<DataLabels>,
    /// Identifiers of the axes used by this chart group
    pub axis_ids: Vec<u32>,
    /// Chart-type extension list
    pub extension_list: Option<ChartExtensionList>,
}

impl TypeGroupCommon {
    /// Create new common properties.
    #[inline]
    pub fn new() -> Self {
        Self {
            vary_colors: false,
            series: Vec::new(),
            data_labels: None,
            axis_ids: Vec::new(),
            extension_list: None,
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

/// A schema-defined chart line whose DrawingML styling is optional.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartLines {
    /// DrawingML shape properties for the line
    pub shape_properties: Option<ChartShapeProperties>,
}

impl ChartLines {
    /// Create an unformatted chart line.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }
}

/// Up/down bars shown between corresponding points in line or stock charts.
#[derive(Debug, Clone, Default)]
pub struct UpDownBars {
    /// Gap between bars (0-500 percent)
    pub gap_width: Option<u32>,
    /// Optional formatting container for rising-value bars
    pub up_bars: Option<ChartLines>,
    /// Optional formatting container for falling-value bars
    pub down_bars: Option<ChartLines>,
    /// Up/down-bar extension list
    pub extension_list: Option<ChartExtensionList>,
}

/// Area chart type group.
#[derive(Debug, Clone)]
pub struct AreaTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Grouping type
    pub grouping: BarGrouping,
    /// Drop lines connecting points to the category axis
    pub drop_lines: Option<ChartLines>,
}

impl AreaTypeGroup {
    /// Create a new area type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
            drop_lines: None,
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
    /// Gap depth (0-500%)
    pub gap_depth: Option<u32>,
    /// Drop lines connecting points to the category axis
    pub drop_lines: Option<ChartLines>,
}

impl Area3DTypeGroup {
    /// Create a new area 3D type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
            gap_depth: None,
            drop_lines: None,
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
    /// Connector-line formatting entries
    pub series_lines: Vec<ChartLines>,
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
            series_lines: Vec::new(),
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
    /// Bubble scale percentage.
    scale: BubbleScale,
    /// Show negative bubbles
    pub show_negative_bubbles: bool,
    /// How each bubble's numeric size is interpreted.
    size: BubbleSize,
}

impl BubbleTypeGroup {
    /// Create a new bubble type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            scale: BubbleScale::default(),
            show_negative_bubbles: true,
            size: BubbleSize::default(),
        }
    }

    /// Returns the bubble scale percentage.
    #[inline]
    pub const fn scale(&self) -> BubbleScale {
        self.scale
    }

    /// Sets the bubble scale percentage.
    #[inline]
    pub fn set_scale(&mut self, scale: BubbleScale) -> &mut Self {
        self.scale = scale;
        self
    }

    /// Sets the bubble scale percentage and returns this group.
    #[inline]
    pub fn with_scale(mut self, scale: BubbleScale) -> Self {
        self.set_scale(scale);
        self
    }

    /// Returns how each bubble's numeric size is interpreted.
    #[inline]
    pub const fn size(&self) -> BubbleSize {
        self.size
    }

    /// Sets how each bubble's numeric size is interpreted.
    #[inline]
    pub fn set_size(&mut self, size: BubbleSize) -> &mut Self {
        self.size = size;
        self
    }

    /// Sets how each bubble's numeric size is interpreted and returns this group.
    #[inline]
    pub fn with_size(mut self, size: BubbleSize) -> Self {
        self.set_size(size);
        self
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
    /// Hole size (1-90%)
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
    /// Smooth the chart-group lines
    pub smooth: bool,
    /// Drop lines connecting points to the category axis
    pub drop_lines: Option<ChartLines>,
    /// Lines connecting the highest and lowest values
    pub high_low_lines: Option<ChartLines>,
    /// Up/down bars between corresponding series points
    pub up_down_bars: Option<UpDownBars>,
}

impl LineTypeGroup {
    /// Create a new line type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
            marker: true,
            smooth: false,
            drop_lines: None,
            high_low_lines: None,
            up_down_bars: None,
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
    /// Gap depth (0-500%)
    pub gap_depth: Option<u32>,
    /// Drop lines connecting points to the category axis
    pub drop_lines: Option<ChartLines>,
}

impl Line3DTypeGroup {
    /// Create a new line 3D type group.
    #[inline]
    pub fn new(grouping: BarGrouping) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            grouping,
            gap_depth: None,
            drop_lines: None,
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

/// Pie-of-pie or bar-of-pie chart type group.
#[derive(Debug, Clone)]
pub struct OfPieTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Secondary plot type
    pub of_pie_type: OfPieType,
    /// Gap between the primary and secondary plots (0-500 percent)
    pub gap_width: Option<u32>,
    /// Rule used to select points for the secondary plot
    pub split_type: Option<OfPieSplitType>,
    /// Position, percentage, or value used by the selected split rule
    pub split_position: Option<f64>,
    /// Explicit point indexes in the secondary plot; `Some` preserves an empty custom split
    pub custom_split_points: Option<Vec<u32>>,
    /// Secondary plot size (5-200 percent)
    pub second_pie_size: Option<u32>,
    /// Connector-line formatting entries between the primary and secondary plots
    pub series_lines: Vec<ChartLines>,
}

impl OfPieTypeGroup {
    /// Create a new of-pie type group.
    #[inline]
    pub fn new(of_pie_type: OfPieType) -> Self {
        Self {
            common: TypeGroupCommon::new(),
            of_pie_type,
            gap_width: None,
            split_type: None,
            split_position: None,
            custom_split_points: None,
            second_pie_size: None,
            series_lines: Vec::new(),
        }
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
    /// Drop lines connecting points to the category axis
    pub drop_lines: Option<ChartLines>,
    /// Lines connecting the highest and lowest values
    pub high_low_lines: Option<ChartLines>,
    /// Up/down bars between corresponding series points
    pub up_down_bars: Option<UpDownBars>,
}

impl StockTypeGroup {
    /// Create a new stock type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            drop_lines: None,
            high_low_lines: None,
            up_down_bars: None,
        }
    }
}

impl Default for StockTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Formatting entry for one indexed surface-chart band.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BandFormat {
    /// Zero-based band index
    pub index: u32,
    /// DrawingML shape properties for the surface band
    pub shape_properties: Option<ChartShapeProperties>,
}

impl BandFormat {
    /// Create a surface band-format entry.
    #[inline]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            shape_properties: None,
        }
    }
}

/// Surface chart type group.
#[derive(Debug, Clone)]
pub struct SurfaceTypeGroup {
    /// Common properties
    pub common: TypeGroupCommon,
    /// Wireframe mode
    pub wireframe: bool,
    /// Optional indexed band-format collection; `Some` preserves an empty wrapper
    pub band_formats: Option<Vec<BandFormat>>,
}

impl SurfaceTypeGroup {
    /// Create a new surface type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            wireframe: false,
            band_formats: None,
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
    /// Optional indexed band-format collection; `Some` preserves an empty wrapper
    pub band_formats: Option<Vec<BandFormat>>,
}

impl Surface3DTypeGroup {
    /// Create a new surface 3D type group.
    #[inline]
    pub fn new() -> Self {
        Self {
            common: TypeGroupCommon::new(),
            wireframe: false,
            band_formats: None,
        }
    }
}

impl Default for Surface3DTypeGroup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}
