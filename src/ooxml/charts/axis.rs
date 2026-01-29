//! Chart axis models.
//!
//! This module contains structures for representing chart axes,
//! including category, value, date, and series axes.

use crate::ooxml::charts::models::{Layout, NumberFormat, TitleText};
use crate::ooxml::charts::types::{
    AxisOrientation, AxisPosition, AxisType, TickLabelPosition, TickMark,
};

/// Axis crossing mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisCrossMode {
    /// Auto zero crossing
    AutoZero,
    /// Maximum value
    Max,
    /// Minimum value
    Min,
}

impl AxisCrossMode {
    /// Returns the XML value for this crossing mode.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::AutoZero => "autoZero",
            Self::Max => "max",
            Self::Min => "min",
        }
    }
}

/// Axis crossing position (between or mid-category).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisCrossBetween {
    /// Cross between categories
    Between,
    /// Cross at mid-category
    MidCategory,
}

impl AxisCrossBetween {
    /// Returns the XML value for this position.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::MidCategory => "midCat",
        }
    }
}

/// Time unit for date axes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeUnit {
    /// Days
    Days,
    /// Months
    Months,
    /// Years
    Years,
}

impl TimeUnit {
    /// Returns the XML value for this time unit.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Days => "days",
            Self::Months => "months",
            Self::Years => "years",
        }
    }
}

/// Built-in display units for value axes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuiltInUnit {
    /// Hundreds
    Hundreds,
    /// Thousands
    Thousands,
    /// Ten thousands
    TenThousands,
    /// Hundred thousands
    HundredThousands,
    /// Millions
    Millions,
    /// Ten millions
    TenMillions,
    /// Hundred millions
    HundredMillions,
    /// Billions
    Billions,
    /// Trillions
    Trillions,
}

impl BuiltInUnit {
    /// Returns the XML value for this built-in unit.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Hundreds => "hundreds",
            Self::Thousands => "thousands",
            Self::TenThousands => "tenThousands",
            Self::HundredThousands => "hundredThousands",
            Self::Millions => "millions",
            Self::TenMillions => "tenMillions",
            Self::HundredMillions => "hundredMillions",
            Self::Billions => "billions",
            Self::Trillions => "trillions",
        }
    }
}

/// Display units configuration for an axis.
#[derive(Debug, Clone)]
pub struct DisplayUnits {
    /// Built-in unit
    pub built_in_unit: Option<BuiltInUnit>,
    /// Custom unit scale
    pub custom_unit: Option<f64>,
    /// Display units label
    pub label: Option<TitleText>,
    /// Layout for the label
    pub layout: Option<Layout>,
}

impl DisplayUnits {
    /// Create display units with a built-in unit.
    #[inline]
    pub fn built_in(unit: BuiltInUnit) -> Self {
        Self {
            built_in_unit: Some(unit),
            custom_unit: None,
            label: None,
            layout: None,
        }
    }

    /// Create display units with a custom scale.
    #[inline]
    pub fn custom(scale: f64) -> Self {
        Self {
            built_in_unit: None,
            custom_unit: Some(scale),
            label: None,
            layout: None,
        }
    }
}

/// Common axis properties shared by all axis types.
#[derive(Debug, Clone)]
pub struct AxisCommon {
    /// Unique axis ID
    pub axis_id: u32,
    /// Axis position (bottom, left, right, top)
    pub position: AxisPosition,
    /// Axis title
    pub title: Option<TitleText>,
    /// Number format for tick labels
    pub number_format: Option<NumberFormat>,
    /// Scaling orientation
    pub orientation: AxisOrientation,
    /// Major tick mark style
    pub major_tick_mark: TickMark,
    /// Minor tick mark style
    pub minor_tick_mark: TickMark,
    /// Tick label position
    pub tick_label_position: TickLabelPosition,
    /// Whether axis is deleted
    pub deleted: bool,
    /// ID of crossing axis
    pub cross_axis_id: u32,
    /// Crossing mode
    pub cross_mode: AxisCrossMode,
    /// Specific crossing value (overrides cross_mode)
    pub crosses_at: Option<f64>,
    /// Show major gridlines
    pub show_major_gridlines: bool,
    /// Show minor gridlines
    pub show_minor_gridlines: bool,
    /// Axis layout
    pub layout: Option<Layout>,
}

impl AxisCommon {
    /// Create a new axis with default settings.
    #[inline]
    pub fn new(axis_id: u32, position: AxisPosition, cross_axis_id: u32) -> Self {
        Self {
            axis_id,
            position,
            title: None,
            number_format: None,
            orientation: AxisOrientation::MinMax,
            major_tick_mark: TickMark::Out,
            minor_tick_mark: TickMark::None,
            tick_label_position: TickLabelPosition::NextTo,
            deleted: false,
            cross_axis_id,
            cross_mode: AxisCrossMode::AutoZero,
            crosses_at: None,
            show_major_gridlines: false,
            show_minor_gridlines: false,
            layout: None,
        }
    }

    /// Set the axis title.
    #[inline]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(TitleText::from_string(title));
        self
    }

    /// Show major gridlines.
    #[inline]
    pub fn with_major_gridlines(mut self, show: bool) -> Self {
        self.show_major_gridlines = show;
        self
    }
}

/// Category axis (text-based categories).
#[derive(Debug, Clone)]
pub struct CategoryAxis {
    /// Common axis properties
    pub common: AxisCommon,
    /// Automatically determine text/date type
    pub auto: bool,
    /// Label alignment
    pub label_align: Option<AxisLabelAlign>,
    /// Label offset from axis (0-1000)
    pub label_offset: Option<u32>,
    /// Skip N tick labels
    pub tick_label_skip: Option<u32>,
    /// Skip N tick marks
    pub tick_mark_skip: Option<u32>,
    /// No multi-level categories
    pub no_multi_level: bool,
}

/// Axis label alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisLabelAlign {
    /// Center
    Center,
    /// Left
    Left,
    /// Right
    Right,
}

impl AxisLabelAlign {
    /// Returns the XML value for this alignment.
    #[inline]
    pub const fn xml_value(&self) -> &'static str {
        match self {
            Self::Center => "ctr",
            Self::Left => "l",
            Self::Right => "r",
        }
    }
}

impl CategoryAxis {
    /// Create a new category axis.
    #[inline]
    pub fn new(axis_id: u32, position: AxisPosition, cross_axis_id: u32) -> Self {
        Self {
            common: AxisCommon::new(axis_id, position, cross_axis_id),
            auto: true,
            label_align: None,
            label_offset: None,
            tick_label_skip: None,
            tick_mark_skip: None,
            no_multi_level: false,
        }
    }
}

/// Value axis (numeric values).
#[derive(Debug, Clone)]
pub struct ValueAxis {
    /// Common axis properties
    pub common: AxisCommon,
    /// Minimum value
    pub min: Option<f64>,
    /// Maximum value
    pub max: Option<f64>,
    /// Major unit
    pub major_unit: Option<f64>,
    /// Minor unit
    pub minor_unit: Option<f64>,
    /// Logarithmic scale base
    pub log_base: Option<f64>,
    /// Display units
    pub display_units: Option<DisplayUnits>,
    /// Cross between categories
    pub cross_between: AxisCrossBetween,
}

impl ValueAxis {
    /// Create a new value axis.
    #[inline]
    pub fn new(axis_id: u32, position: AxisPosition, cross_axis_id: u32) -> Self {
        Self {
            common: AxisCommon::new(axis_id, position, cross_axis_id),
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            log_base: None,
            display_units: None,
            cross_between: AxisCrossBetween::Between,
        }
    }

    /// Set the axis range.
    #[inline]
    pub fn with_range(mut self, min: f64, max: f64) -> Self {
        self.min = Some(min);
        self.max = Some(max);
        self
    }

    /// Set the major unit.
    #[inline]
    pub fn with_major_unit(mut self, unit: f64) -> Self {
        self.major_unit = Some(unit);
        self
    }

    /// Enable logarithmic scale.
    #[inline]
    pub fn with_log_scale(mut self, base: f64) -> Self {
        self.log_base = Some(base);
        self
    }
}

/// Date axis (date-based categories).
#[derive(Debug, Clone)]
pub struct DateAxis {
    /// Common axis properties
    pub common: AxisCommon,
    /// Minimum date value
    pub min: Option<f64>,
    /// Maximum date value
    pub max: Option<f64>,
    /// Major unit
    pub major_unit: Option<f64>,
    /// Minor unit
    pub minor_unit: Option<f64>,
    /// Major time unit
    pub major_time_unit: Option<TimeUnit>,
    /// Minor time unit
    pub minor_time_unit: Option<TimeUnit>,
    /// Base time unit
    pub base_time_unit: Option<TimeUnit>,
    /// Automatically determine text/date type
    pub auto: bool,
}

impl DateAxis {
    /// Create a new date axis.
    #[inline]
    pub fn new(axis_id: u32, position: AxisPosition, cross_axis_id: u32) -> Self {
        Self {
            common: AxisCommon::new(axis_id, position, cross_axis_id),
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            major_time_unit: None,
            minor_time_unit: None,
            base_time_unit: None,
            auto: true,
        }
    }
}

/// Series axis (for 3D charts).
#[derive(Debug, Clone)]
pub struct SeriesAxis {
    /// Common axis properties
    pub common: AxisCommon,
    /// Skip N tick labels
    pub tick_label_skip: Option<u32>,
    /// Skip N tick marks
    pub tick_mark_skip: Option<u32>,
}

impl SeriesAxis {
    /// Create a new series axis.
    #[inline]
    pub fn new(axis_id: u32, position: AxisPosition, cross_axis_id: u32) -> Self {
        Self {
            common: AxisCommon::new(axis_id, position, cross_axis_id),
            tick_label_skip: None,
            tick_mark_skip: None,
        }
    }
}

/// Axis enumeration for holding any axis type.
#[derive(Debug, Clone)]
pub enum Axis {
    /// Category axis
    Category(CategoryAxis),
    /// Value axis
    Value(ValueAxis),
    /// Date axis
    Date(DateAxis),
    /// Series axis
    Series(SeriesAxis),
}

impl Axis {
    /// Get the axis type.
    #[inline]
    pub const fn axis_type(&self) -> AxisType {
        match self {
            Self::Category(_) => AxisType::Category,
            Self::Value(_) => AxisType::Value,
            Self::Date(_) => AxisType::Date,
            Self::Series(_) => AxisType::Series,
        }
    }

    /// Get the common axis properties.
    #[inline]
    pub fn common(&self) -> &AxisCommon {
        match self {
            Self::Category(ax) => &ax.common,
            Self::Value(ax) => &ax.common,
            Self::Date(ax) => &ax.common,
            Self::Series(ax) => &ax.common,
        }
    }

    /// Get mutable common axis properties.
    #[inline]
    pub fn common_mut(&mut self) -> &mut AxisCommon {
        match self {
            Self::Category(ax) => &mut ax.common,
            Self::Value(ax) => &mut ax.common,
            Self::Date(ax) => &mut ax.common,
            Self::Series(ax) => &mut ax.common,
        }
    }

    /// Get the axis ID.
    #[inline]
    pub fn axis_id(&self) -> u32 {
        self.common().axis_id
    }
}
