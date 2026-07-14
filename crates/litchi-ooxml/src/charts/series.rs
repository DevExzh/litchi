//! Chart series and data point models.
//!
//! This module contains structures for representing chart series,
//! data points, and their associated properties.

use crate::charts::chart::{ChartExtensionList, ChartShapeProperties, PictureOptions};
use crate::charts::models::{Layout, NumberFormat, NumericData, StringData, TitleText};
use crate::charts::plot_area::{BarShape, ChartLines};
use crate::charts::types::{DataLabelPosition, MarkerStyle};

/// Marker formatting shared by chart elements that support point symbols.
#[derive(Debug, Clone, Default)]
pub struct Marker {
    /// Marker symbol
    pub symbol: Option<MarkerStyle>,
    /// Marker size in points (2-72)
    pub size: Option<u32>,
    /// DrawingML shape properties for the marker
    pub shape_properties: Option<ChartShapeProperties>,
    /// Marker extension list
    pub extension_list: Option<ChartExtensionList>,
}

impl Marker {
    /// Create an empty marker override.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set marker symbol and size.
    #[inline]
    pub fn with_symbol_and_size(mut self, symbol: MarkerStyle, size: u32) -> Self {
        self.symbol = Some(symbol);
        self.size = Some(size);
        self
    }
}

/// A single data point with optional formatting.
#[derive(Debug, Clone)]
pub struct DataPoint {
    /// Index of this data point
    pub index: u32,
    /// Explosion (for pie/doughnut charts, in percent)
    pub explosion: Option<u32>,
    /// Marker size
    pub marker_size: Option<u32>,
    /// Marker symbol
    pub marker_symbol: Option<MarkerStyle>,
    /// Whether an explicit marker element is present, including an empty default marker
    pub marker_present: bool,
    /// DrawingML shape properties for the marker
    pub marker_shape_properties: Option<ChartShapeProperties>,
    /// Marker extension list
    pub marker_extension_list: Option<ChartExtensionList>,
    /// Invert colors if negative
    pub invert_if_negative: bool,
    /// Show bubble in 3D
    pub bubble_3d: Option<bool>,
    /// DrawingML shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// Picture-fill placement options
    pub picture_options: Option<PictureOptions>,
    /// Data-point extension list
    pub extension_list: Option<ChartExtensionList>,
}

impl DataPoint {
    /// Create a new data point.
    #[inline]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            explosion: None,
            marker_size: None,
            marker_symbol: None,
            marker_present: false,
            marker_shape_properties: None,
            marker_extension_list: None,
            invert_if_negative: false,
            bubble_3d: None,
            shape_properties: None,
            picture_options: None,
            extension_list: None,
        }
    }

    /// Set explosion percentage.
    #[inline]
    pub fn with_explosion(mut self, explosion: u32) -> Self {
        self.explosion = Some(explosion);
        self
    }

    /// Set marker properties.
    #[inline]
    pub fn with_marker(mut self, size: u32, symbol: MarkerStyle) -> Self {
        self.marker_present = true;
        self.marker_size = Some(size);
        self.marker_symbol = Some(symbol);
        self
    }
}

/// Data label settings.
#[derive(Debug, Clone)]
pub struct DataLabels {
    /// Point-specific data-label overrides
    pub labels: Vec<DataLabel>,
    /// Number format for label values
    pub number_format: Option<NumberFormat>,
    /// DrawingML shape properties for all labels
    pub shape_properties: Option<ChartShapeProperties>,
    /// DrawingML text properties for all labels
    pub text_properties: Option<crate::charts::chart::ChartTextProperties>,
    /// Position of data labels
    pub position: Option<DataLabelPosition>,
    /// Show legend key
    pub show_legend_key: bool,
    /// Show value
    pub show_value: bool,
    /// Show category name
    pub show_category_name: bool,
    /// Show series name
    pub show_series_name: bool,
    /// Show percentage (for pie charts)
    pub show_percent: bool,
    /// Show bubble size (for bubble charts)
    pub show_bubble_size: bool,
    /// Show leader lines between labels and data points
    pub show_leader_lines: bool,
    /// Leader-line formatting
    pub leader_lines: Option<ChartLines>,
    /// Separator between label components
    pub separator: Option<String>,
    /// Whether data labels are deleted
    pub deleted: bool,
    /// Data-label collection extension list
    pub extension_list: Option<ChartExtensionList>,
}

impl DataLabels {
    /// Create a new data labels configuration.
    #[inline]
    pub fn new() -> Self {
        Self {
            labels: Vec::new(),
            number_format: None,
            shape_properties: None,
            text_properties: None,
            position: None,
            show_legend_key: false,
            show_value: false,
            show_category_name: false,
            show_series_name: false,
            show_percent: false,
            show_bubble_size: false,
            show_leader_lines: false,
            leader_lines: None,
            separator: None,
            deleted: false,
            extension_list: None,
        }
    }

    /// Show values on labels.
    #[inline]
    pub fn with_show_value(mut self, show: bool) -> Self {
        self.show_value = show;
        self
    }

    /// Set label position.
    #[inline]
    pub fn with_position(mut self, position: DataLabelPosition) -> Self {
        self.position = Some(position);
        self
    }
}

/// Data-label settings for one data point.
#[derive(Debug, Clone)]
pub struct DataLabel {
    /// Zero-based data-point index
    pub index: u32,
    /// Whether this label is deleted
    pub deleted: bool,
    /// Manual layout for this label
    pub layout: Option<Layout>,
    /// Explicit label text or formula reference
    pub text: Option<TitleText>,
    /// Number format for the label value
    pub number_format: Option<NumberFormat>,
    /// DrawingML shape properties for this label
    pub shape_properties: Option<ChartShapeProperties>,
    /// DrawingML text properties for this label
    pub text_properties: Option<crate::charts::chart::ChartTextProperties>,
    /// Position of the label
    pub position: Option<DataLabelPosition>,
    /// Show legend key
    pub show_legend_key: bool,
    /// Show value
    pub show_value: bool,
    /// Show category name
    pub show_category_name: bool,
    /// Show series name
    pub show_series_name: bool,
    /// Show percentage
    pub show_percent: bool,
    /// Show bubble size
    pub show_bubble_size: bool,
    /// Separator between label components
    pub separator: Option<String>,
    /// Point data-label extension list
    pub extension_list: Option<ChartExtensionList>,
}

impl DataLabel {
    /// Create an empty override for one point.
    #[inline]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            deleted: false,
            layout: None,
            text: None,
            number_format: None,
            shape_properties: None,
            text_properties: None,
            position: None,
            show_legend_key: false,
            show_value: false,
            show_category_name: false,
            show_series_name: false,
            show_percent: false,
            show_bubble_size: false,
            separator: None,
            extension_list: None,
        }
    }
}

impl Default for DataLabels {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Error bar configuration.
#[derive(Debug, Clone)]
pub struct ErrorBar {
    /// Direction (X or Y axis)
    pub direction: ErrorBarDirection,
    /// Type (both, plus, minus)
    pub error_type: ErrorBarType,
    /// Value type (fixed, percentage, standard deviation, standard error, custom)
    pub value_type: ErrorBarValueType,
    /// Fixed value (for fixed value type)
    pub value: Option<f64>,
    /// Plus values (for custom)
    pub plus_values: Option<NumericData>,
    /// Minus values (for custom)
    pub minus_values: Option<NumericData>,
    /// No end cap on error bars
    pub no_end_cap: bool,
}

/// Error bar direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarDirection {
    /// X direction
    X,
    /// Y direction
    Y,
}

/// Error bar type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarType {
    /// Both directions
    Both,
    /// Positive direction only
    Plus,
    /// Negative direction only
    Minus,
}

/// Error bar value type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarValueType {
    /// Fixed value
    Fixed,
    /// Percentage
    Percentage,
    /// Standard deviation
    StdDev,
    /// Standard error
    StdErr,
    /// Custom values
    Custom,
}

/// Trendline configuration.
#[derive(Debug, Clone)]
pub struct Trendline {
    /// Trendline type
    pub trendline_type: TrendlineType,
    /// Name of the trendline
    pub name: Option<String>,
    /// Polynomial order (for polynomial trendlines, 2-6)
    pub order: Option<u32>,
    /// Moving average period (for moving average, 2-255)
    pub period: Option<u32>,
    /// Forward extrapolation
    pub forward: Option<f64>,
    /// Backward extrapolation
    pub backward: Option<f64>,
    /// Intercept value
    pub intercept: Option<f64>,
    /// Display equation on chart
    pub display_equation: bool,
    /// Display R-squared value on chart
    pub display_r_squared: bool,
    /// Whether a trendline label is present
    pub show_label: bool,
    /// Explicit trendline-label text or formula reference
    pub label: Option<TitleText>,
    /// Manual layout for the trendline label
    pub label_layout: Option<Layout>,
    /// Number format for the trendline label
    pub label_number_format: Option<NumberFormat>,
}

/// Trendline type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrendlineType {
    /// Exponential
    Exponential,
    /// Linear
    Linear,
    /// Logarithmic
    Logarithmic,
    /// Moving average
    MovingAverage,
    /// Polynomial
    Polynomial,
    /// Power
    Power,
}

impl Trendline {
    /// Create a linear trendline.
    #[inline]
    pub fn linear() -> Self {
        Self {
            trendline_type: TrendlineType::Linear,
            name: None,
            order: None,
            period: None,
            forward: None,
            backward: None,
            intercept: None,
            display_equation: false,
            display_r_squared: false,
            show_label: false,
            label: None,
            label_layout: None,
            label_number_format: None,
        }
    }
}

/// A data series in a chart.
#[derive(Debug, Clone)]
pub struct Series {
    /// Series index (for rendering order)
    pub index: u32,
    /// Series order (for legend order)
    pub order: u32,
    /// Series title
    pub title: Option<TitleText>,
    /// Category data (X-axis for scatter/bubble)
    pub categories: Option<StringData>,
    /// Value data (Y-axis)
    pub values: Option<NumericData>,
    /// X values (for scatter charts)
    pub x_values: Option<NumericData>,
    /// Y values (for scatter charts)
    pub y_values: Option<NumericData>,
    /// Bubble sizes (for bubble charts)
    pub bubble_sizes: Option<NumericData>,
    /// Individual data points with custom formatting
    pub data_points: Vec<DataPoint>,
    /// Data labels configuration
    pub data_labels: Option<DataLabels>,
    /// Marker size (2-72)
    pub marker_size: Option<u32>,
    /// Marker symbol
    pub marker_symbol: Option<MarkerStyle>,
    /// Whether an explicit series marker is present, including an empty default marker
    pub marker_present: bool,
    /// DrawingML shape properties for the series marker
    pub marker_shape_properties: Option<ChartShapeProperties>,
    /// Series-marker extension list
    pub marker_extension_list: Option<ChartExtensionList>,
    /// Explosion (for pie/doughnut, in percent)
    pub explosion: Option<u32>,
    /// Smooth line (for line/scatter charts)
    pub smooth: bool,
    /// Invert colors if negative
    pub invert_if_negative: bool,
    /// Show bubble in 3D
    pub bubble_3d: bool,
    /// Error bars
    pub error_bars: Vec<ErrorBar>,
    /// Trendlines
    pub trendlines: Vec<Trendline>,
    /// DrawingML shape properties
    pub shape_properties: Option<ChartShapeProperties>,
    /// Area- and bar-series picture-fill placement options
    pub picture_options: Option<PictureOptions>,
    /// Per-series shape override for bar and column charts
    pub bar_shape: Option<BarShape>,
    /// Series extension list
    pub extension_list: Option<ChartExtensionList>,
}

impl Series {
    /// Create a new series with index.
    #[inline]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            order: index,
            title: None,
            categories: None,
            values: None,
            x_values: None,
            y_values: None,
            bubble_sizes: None,
            data_points: Vec::new(),
            data_labels: None,
            marker_size: None,
            marker_symbol: None,
            marker_present: false,
            marker_shape_properties: None,
            marker_extension_list: None,
            explosion: None,
            smooth: false,
            invert_if_negative: false,
            bubble_3d: false,
            error_bars: Vec::new(),
            trendlines: Vec::new(),
            shape_properties: None,
            picture_options: None,
            bar_shape: None,
            extension_list: None,
        }
    }

    /// Set the series title.
    #[inline]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(TitleText::from_string(title));
        self
    }

    /// Set category data.
    #[inline]
    pub fn with_categories(mut self, categories: StringData) -> Self {
        self.categories = Some(categories);
        self
    }

    /// Set value data.
    #[inline]
    pub fn with_values(mut self, values: NumericData) -> Self {
        self.values = Some(values);
        self
    }

    /// Set X-Y values for scatter charts.
    #[inline]
    pub fn with_xy_values(mut self, x_values: NumericData, y_values: NumericData) -> Self {
        self.x_values = Some(x_values);
        self.y_values = Some(y_values);
        self
    }

    /// Add a data point.
    #[inline]
    pub fn add_data_point(mut self, point: DataPoint) -> Self {
        self.data_points.push(point);
        self
    }

    /// Set data labels.
    #[inline]
    pub fn with_data_labels(mut self, labels: DataLabels) -> Self {
        self.data_labels = Some(labels);
        self
    }

    /// Add a trendline.
    #[inline]
    pub fn add_trendline(mut self, trendline: Trendline) -> Self {
        self.trendlines.push(trendline);
        self
    }
}
