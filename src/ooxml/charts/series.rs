//! Chart series and data point models.
//!
//! This module contains structures for representing chart series,
//! data points, and their associated properties.

use crate::ooxml::charts::models::{NumericData, StringData, TitleText};
use crate::ooxml::charts::types::{DataLabelPosition, MarkerStyle};

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
    /// Invert colors if negative
    pub invert_if_negative: bool,
    /// Show bubble in 3D
    pub bubble_3d: Option<bool>,
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
            invert_if_negative: false,
            bubble_3d: None,
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
        self.marker_size = Some(size);
        self.marker_symbol = Some(symbol);
        self
    }
}

/// Data label settings.
#[derive(Debug, Clone)]
pub struct DataLabels {
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
    /// Separator between label components
    pub separator: Option<String>,
    /// Whether data labels are deleted
    pub deleted: bool,
}

impl DataLabels {
    /// Create a new data labels configuration.
    #[inline]
    pub fn new() -> Self {
        Self {
            position: None,
            show_legend_key: false,
            show_value: false,
            show_category_name: false,
            show_series_name: false,
            show_percent: false,
            show_bubble_size: false,
            separator: None,
            deleted: false,
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
            explosion: None,
            smooth: false,
            invert_if_negative: false,
            bubble_3d: false,
            error_bars: Vec::new(),
            trendlines: Vec::new(),
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
