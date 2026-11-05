//! Host-neutral chart authoring from tabular data.
//!
//! A [`Spec`] describes a new chart the way a caller thinks of it: a
//! [`Variant`] (column, stacked bar, 3D pie, ...), a title, the category
//! labels, the named [`Series`] values, the legend position and the axis
//! titles. [`Spec::to_chart`] turns it into the typed [`Chart`] model and
//! [`Spec::to_xml`] into a complete chart-space part, deterministically: the
//! same spec always produces the same bytes.
//!
//! The spec does not own package placement. Each host (PPTX slide frame,
//! XLSX drawing anchor, DOCX inline drawing) allocates the part, its
//! relationship and, where it wants one, the embedded workbook. The
//! [`Binding`] says where the chart's data lives:
//!
//! - [`Binding::Literal`] writes the values inline (`c:numLit`/`c:strLit`),
//!   which needs no workbook (an XLSX worksheet chart authored from inline
//!   data);
//! - [`Binding::Sheet`] writes references, with cached values, into a
//!   worksheet holding the layout [`Spec::sheet_layout`] describes (an
//!   embedded workbook's sheet, or a range of the host workbook) (category labels in column A from row 2, one column per
//!   series with its name in row 1; scatter and bubble charts give each
//!   series an X column, a Y column and, for bubbles, a size column).
//!
//! ```
//! use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
//! use litchi_drawingml::chart::types::LegendPosition;
//!
//! let spec = Spec::new(Variant::new(Kind::Column))
//!     .with_title("Sales")
//!     .with_categories(["Q1", "Q2", "Q3"])
//!     .with_series(Series::new("East", vec![120.0, 135.0, 148.0]))
//!     .with_legend(Some(LegendPosition::Bottom));
//! let xml = spec.to_xml(None)?;
//! let chart = litchi_drawingml::chart::reader::read(
//!     &xml[..],
//!     &litchi_drawingml::chart::Limits::DEFAULT,
//! )?;
//! assert_eq!(Variant::of(&chart), Some(Variant::new(Kind::Column)));
//! # Ok::<(), litchi_drawingml::Error>(())
//! ```

use crate::chart::axis::{Axis, CategoryAxis, SeriesAxis, ValueAxis};
use crate::chart::data::{NumericData, StringData, TitleText};
use crate::chart::legend::Legend;
use crate::chart::model::{Chart, View3D};
use crate::chart::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, Bar3DTypeGroup, BarTypeGroup, BubbleTypeGroup,
    DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea,
    RadarTypeGroup, ScatterTypeGroup, StockTypeGroup, TypeGroup, TypeGroupCommon,
};
use crate::chart::series::{DataLabels, Series as ChartSeries};
use crate::chart::types::{
    AxisPosition, BarDirection, BarGrouping, DataLabelPosition, LegendPosition, MarkerStyle,
    RadarStyle, ScatterStyle,
};
use crate::{Error, Result};

mod read;
mod workbook;

pub use read::{Parsed, read};

/// The most series one authored chart may carry (the spreadsheet limit).
pub const MAX_SERIES: usize = 255;
/// The most points one authored series may carry.
pub const MAX_POINTS: usize = 32_000;
/// The most points of all series together.
pub const MAX_TOTAL_POINTS: usize = 256_000;
/// The longest title, series name or category label, in bytes.
pub const MAX_TEXT_BYTES: usize = 32_767;

const CATEGORY_AXIS_ID: u32 = 500_000_001;
const VALUE_AXIS_ID: u32 = 500_000_002;
const SERIES_AXIS_ID: u32 = 500_000_003;

/// A chart family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Vertical bars.
    Column,
    /// Horizontal bars.
    Bar,
    /// Lines through category points.
    Line,
    /// One series as pie slices.
    Pie,
    /// Pie slices around a hole.
    Doughnut,
    /// Filled areas under lines.
    Area,
    /// X/Y points.
    Scatter,
    /// X/Y points sized by a third value.
    Bubble,
    /// Values on radial axes.
    Radar,
    /// High-low-close or open-high-low-close series.
    Stock,
    /// Column, line and area series on shared axes, one family per series
    /// ([`Spec::with_series_kinds`]).
    Combo,
    /// A running total of increases and decreases, its last point the total
    /// by default ([`Appearance::subtotals`]); an extended chart.
    Waterfall,
    /// Stages narrowing from the first value; an extended chart.
    Funnel,
    /// Nested rectangles sized by value; an extended chart.
    Treemap,
    /// Concentric rings sized by value; an extended chart.
    Sunburst,
    /// The distribution of the values in bins ([`Appearance::bins`]); an
    /// extended chart.
    Histogram,
    /// Columns in descending order with a cumulative-percentage line; an
    /// extended chart.
    Pareto,
    /// The quartiles, mean and outliers of each series; an extended chart.
    BoxWhisker,
}

impl Kind {
    /// Whether the family is an extended chart, written as a `cx:chartSpace`
    /// part ([`crate::chart::chartex`]) rather than a `c:chartSpace` one.
    ///
    /// ```
    /// use litchi_drawingml::chart::authoring::Kind;
    ///
    /// assert!(Kind::Waterfall.is_extended());
    /// assert!(!Kind::Column.is_extended());
    /// ```
    #[must_use]
    pub const fn is_extended(self) -> bool {
        matches!(
            self,
            Self::Waterfall
                | Self::Funnel
                | Self::Treemap
                | Self::Sunburst
                | Self::Histogram
                | Self::Pareto
                | Self::BoxWhisker
        )
    }
}

/// How the series of a bar, line or area chart combine.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Grouping {
    /// Side by side (bars) or overlapping (lines, areas).
    Standard,
    /// Stacked on each other.
    Stacked,
    /// Stacked to 100%.
    PercentStacked,
}

/// The chart family with its grouping and dimensionality.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Variant {
    /// The family.
    pub kind: Kind,
    /// The grouping; only column, bar, line and area charts stack.
    pub grouping: Grouping,
    /// Whether the chart is drawn in 3D; only column, bar, line, pie and
    /// area charts have a 3D form.
    pub three_d: bool,
}

impl Variant {
    /// The standard, flat variant of a family.
    #[must_use]
    pub const fn new(kind: Kind) -> Self {
        Self {
            kind,
            grouping: Grouping::Standard,
            three_d: false,
        }
    }

    /// The same variant with another grouping.
    #[must_use]
    pub const fn with_grouping(mut self, grouping: Grouping) -> Self {
        self.grouping = grouping;
        self
    }

    /// The same variant in 3D (`true`) or flat.
    #[must_use]
    pub const fn with_three_d(mut self, three_d: bool) -> Self {
        self.three_d = three_d;
        self
    }

    /// Whether the family admits this grouping and dimensionality.
    #[must_use]
    pub const fn is_valid(self) -> bool {
        let stacks = matches!(
            self.kind,
            Kind::Column | Kind::Bar | Kind::Line | Kind::Area
        );
        let solid = matches!(
            self.kind,
            Kind::Column | Kind::Bar | Kind::Line | Kind::Pie | Kind::Area
        );
        let stacks = stacks && !matches!(self.kind, Kind::Combo);
        (stacks || matches!(self.grouping, Grouping::Standard)) && (solid || !self.three_d)
    }

    /// Classify the first chart group of a chart, if it is one an authored
    /// spec can produce.
    #[must_use]
    pub fn of(chart: &Chart) -> Option<Self> {
        if chart.plot_area.type_groups.len() > 1 {
            return Some(Self::new(Kind::Combo));
        }
        let group = chart.plot_area.type_groups.first()?;
        let grouping = |value: BarGrouping| match value {
            BarGrouping::Stacked => Grouping::Stacked,
            BarGrouping::PercentStacked => Grouping::PercentStacked,
            BarGrouping::Clustered | BarGrouping::Standard => Grouping::Standard,
        };
        let direction = |value: BarDirection| match value {
            BarDirection::Bar => Kind::Bar,
            BarDirection::Column => Kind::Column,
        };
        Some(match group {
            TypeGroup::Bar(group) => {
                Self::new(direction(group.direction)).with_grouping(grouping(group.grouping))
            },
            TypeGroup::Bar3D(group) => Self::new(direction(group.direction))
                .with_grouping(grouping(group.grouping))
                .with_three_d(true),
            TypeGroup::Line(group) => Self::new(Kind::Line).with_grouping(grouping(group.grouping)),
            TypeGroup::Line3D(group) => Self::new(Kind::Line)
                .with_grouping(grouping(group.grouping))
                .with_three_d(true),
            TypeGroup::Area(group) => Self::new(Kind::Area).with_grouping(grouping(group.grouping)),
            TypeGroup::Area3D(group) => Self::new(Kind::Area)
                .with_grouping(grouping(group.grouping))
                .with_three_d(true),
            TypeGroup::Pie(_) => Self::new(Kind::Pie),
            TypeGroup::Pie3D(_) => Self::new(Kind::Pie).with_three_d(true),
            TypeGroup::Doughnut(_) => Self::new(Kind::Doughnut),
            TypeGroup::Scatter(_) => Self::new(Kind::Scatter),
            TypeGroup::Bubble(_) => Self::new(Kind::Bubble),
            TypeGroup::Radar(_) => Self::new(Kind::Radar),
            TypeGroup::Stock(_) => Self::new(Kind::Stock),
            TypeGroup::OfPie(_) | TypeGroup::Surface(_) | TypeGroup::Surface3D(_) => return None,
        })
    }
}

/// One named series of values.
#[derive(Debug, Clone, PartialEq)]
pub struct Series {
    /// The series name shown in the legend.
    pub name: String,
    /// The values, one per category (the Y values of a scatter or bubble
    /// chart).
    pub values: Vec<f64>,
    /// The series' own `RRGGBB` (or `AARRGGBB`) fill color; it takes
    /// precedence over this series' entry of [`Appearance::colors`]. On a
    /// single-series pie or doughnut it colors the whole series, while
    /// [`Appearance::colors`] still colors its points.
    pub color: Option<String>,
}

impl Series {
    /// A series with a name and its values.
    #[must_use]
    pub fn new(name: impl Into<String>, values: Vec<f64>) -> Self {
        Self {
            name: name.into(),
            values,
            color: None,
        }
    }

    /// The series with its own fill color, `RRGGBB` or `AARRGGBB`.
    ///
    /// ```
    /// use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
    ///
    /// let spec = Spec::new(Variant::new(Kind::Column))
    ///     .with_series(Series::new("Web", vec![1.0, 2.0]).with_color("4472C4"));
    /// let xml = String::from_utf8(spec.to_xml(None)?).expect("UTF-8");
    /// assert!(xml.contains(r#"<a:srgbClr val="4472C4"/>"#));
    /// # Ok::<(), litchi_drawingml::Error>(())
    /// ```
    #[must_use]
    pub fn with_color(mut self, color: impl Into<String>) -> Self {
        self.color = Some(color.into());
        self
    }
}

/// Where an authored chart's data lives.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Binding {
    /// Inline values, no workbook.
    Literal,
    /// References, with cached values, into a worksheet holding the
    /// [`Spec::sheet_layout`] layout: an embedded workbook's sheet (origin
    /// `A1`), or a host worksheet range.
    Sheet {
        /// The worksheet name.
        sheet: String,
        /// The zero-based row of the layout's first row (the series names).
        row: usize,
        /// The zero-based column of the layout's first column (the
        /// category labels).
        column: usize,
    },
}

impl Binding {
    /// References into `sheet` with the layout at `A1`, the form an
    /// embedded workbook uses.
    #[must_use]
    pub fn sheet(sheet: impl Into<String>) -> Self {
        Self::Sheet {
            sheet: sheet.into(),
            row: 0,
            column: 0,
        }
    }
}

/// One cell of the worksheet layout a [`Binding::Sheet`] chart refers to.
#[derive(Debug, Clone, PartialEq)]
pub enum Cell<'a> {
    /// A text cell (a series name or category label).
    Text(&'a str),
    /// A number.
    Number(f64),
}

/// Which values a chart's data labels show.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Labels {
    /// The point's value.
    pub value: bool,
    /// The point's category.
    pub category: bool,
    /// The series name.
    pub series: bool,
    /// The point's share of a pie or doughnut.
    pub percent: bool,
    /// Where the labels sit; `None` lets the application choose.
    pub position: Option<DataLabelPosition>,
}

/// A line's appearance: an `RRGGBB` color, a width in points and a
/// preset dash (`solid`, `dot`, `dash`, ...).
#[derive(Debug, Clone, PartialEq)]
pub struct Line {
    /// The `RRGGBB` color.
    pub color: String,
    /// The width in points.
    pub width: Option<f64>,
    /// The preset dash name.
    pub dash: Option<String>,
}

/// Optional formatting of an authored chart. Every field left `None` (or
/// empty) keeps the application default.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Appearance {
    /// `RRGGBB` (or `AARRGGBB`, with an alpha byte) fill colors, one per
    /// series in order (one per point of a single-series pie or doughnut).
    pub colors: Vec<String>,
    /// Data labels on every series; `Some` with nothing shown hides them.
    pub labels: Option<Labels>,
    /// Whether every point takes its own color.
    pub vary_colors: Option<bool>,
    /// The gap between bar clusters, in percent of a bar (0-500).
    pub gap_width: Option<u32>,
    /// The overlap of bars in a cluster, in percent (-100-100).
    pub overlap: Option<i32>,
    /// The hole of a doughnut, in percent of its radius (1-90).
    pub hole_size: Option<u32>,
    /// The angle of a pie's or doughnut's first slice, in degrees (0-360).
    pub first_slice_angle: Option<u32>,
    /// How a scatter chart draws its points.
    pub scatter_style: Option<ScatterStyle>,
    /// How a radar chart draws its series.
    pub radar_style: Option<RadarStyle>,
    /// Whether line and scatter series are smoothed.
    pub smooth: Option<bool>,
    /// How far pie slices are pulled out, in percent of the radius.
    pub explosion: Option<u32>,
    /// Whether the legend overlaps the plot area.
    pub legend_overlay: Option<bool>,
    /// Whether the automatic title of an untitled chart is hidden.
    pub auto_title_deleted: Option<bool>,
    /// The value axis minimum.
    pub axis_min: Option<f64>,
    /// The value axis maximum.
    pub axis_max: Option<f64>,
    /// The value axis major unit.
    pub major_unit: Option<f64>,
    /// The value axis major gridlines: `Some(None)` plain, `Some(Some(line))`
    /// styled, absent for the default (shown); see
    /// [`Self::hide_gridlines`].
    pub gridlines: Option<Option<Line>>,
    /// Hide the value axis major gridlines.
    pub hide_gridlines: bool,
    /// The value axis minor gridlines, shown when `Some`.
    pub minor_gridlines: Option<Option<Line>>,
    /// The marker of line, scatter and radar series.
    pub marker: Option<MarkerStyle>,
    /// The marker size in points (2-72).
    pub marker_size: Option<u32>,
    /// The marker's `RRGGBB` fill color.
    pub marker_color: Option<String>,
    /// The width of line, scatter and radar series lines, in points.
    pub line_width: Option<f64>,
    /// The 3D view: X rotation (0-90), Y rotation (0-360) and perspective
    /// (0-240).
    pub view_3d: Option<(u32, u32, Option<u32>)>,
    /// The zero-based points of a waterfall chart drawn as totals (the
    /// running sum) rather than as changes; `None` makes the last point the
    /// total, an empty list draws every point as a change.
    pub subtotals: Option<Vec<usize>>,
    /// The `RRGGBB` (or `AARRGGBB`) fill of a waterfall chart's increases.
    pub increase_color: Option<String>,
    /// The `RRGGBB` (or `AARRGGBB`) fill of a waterfall chart's decreases.
    pub decrease_color: Option<String>,
    /// The `RRGGBB` (or `AARRGGBB`) fill of a waterfall chart's totals.
    pub total_color: Option<String>,
    /// How a histogram (or a Pareto chart without categories) bins its
    /// values; `None` lets the application choose.
    pub bins: Option<Bins>,
    /// How a box-and-whisker chart computes its quartiles.
    pub quartile_method: Option<QuartileMethod>,
    /// How a treemap places the labels of parent categories.
    pub parent_labels: Option<ParentLabels>,
}

/// How a histogram groups its values into bins.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct Bins {
    /// The number of bins.
    pub count: Option<u32>,
    /// The width of each bin; exclusive with [`Self::count`].
    pub size: Option<f64>,
    /// Whether a bin includes its lower bound (`[a, b)`) rather than its
    /// upper bound (`(a, b]`, the default).
    pub closed_left: bool,
    /// Values above this go to one overflow bin.
    pub overflow: Option<f64>,
    /// Values below this go to one underflow bin.
    pub underflow: Option<f64>,
}

/// How a box-and-whisker chart computes quartiles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QuartileMethod {
    /// The median is excluded from both halves.
    Exclusive,
    /// The median is included in both halves.
    Inclusive,
}

/// How a treemap shows the labels of parent categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParentLabels {
    /// Over the rectangles of their children.
    Overlapping,
    /// In a banner above their children.
    Banner,
    /// Not shown.
    None,
}

/// A complete description of a new chart.
#[derive(Debug, Clone, PartialEq)]
pub struct Spec {
    variant: Variant,
    title: Option<String>,
    categories: Vec<String>,
    series: Vec<Series>,
    legend: Option<LegendPosition>,
    category_axis_title: Option<String>,
    value_axis_title: Option<String>,
    binding: Binding,
    appearance: Appearance,
    series_kinds: Vec<Kind>,
    /// Per-series worksheet references (record 0338): `(series index,
    /// values reference, categories reference)`.
    series_references: Vec<(usize, String, Option<String>)>,
}

const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn is_rgb(color: &str) -> bool {
    color.len() == 6 && color.bytes().all(|byte| byte.is_ascii_hexdigit())
}

fn is_argb(color: &str) -> bool {
    matches!(color.len(), 6 | 8) && color.bytes().all(|byte| byte.is_ascii_hexdigit())
}

/// An `a:srgbClr` for `RRGGBB` or `AARRGGBB` (the alpha byte as `a:alpha`).
fn srgb(color: &str) -> String {
    let color = color.to_ascii_uppercase();
    if color.len() == 8 {
        let alpha = u32::from_str_radix(&color[..2], 16).unwrap_or(255);
        format!(
            r#"<a:srgbClr val="{}"><a:alpha val="{}"/></a:srgbClr>"#,
            &color[2..],
            alpha * 100_000 / 255
        )
    } else {
        format!(r#"<a:srgbClr val="{color}"/>"#)
    }
}

fn line_xml(line: &Line) -> String {
    let width = line
        .width
        .map(|points| format!(r#" w="{}""#, (points * 12_700.0).round() as i64))
        .unwrap_or_default();
    let dash = line
        .dash
        .as_ref()
        .map(|dash| format!(r#"<a:prstDash val="{dash}"/>"#))
        .unwrap_or_default();
    format!(
        r#"<a:ln{width}><a:solidFill><a:srgbClr val="{}"/></a:solidFill>{dash}</a:ln>"#,
        line.color.to_ascii_uppercase()
    )
}

fn shape_properties(body: &str) -> Result<crate::chart::model::ShapeProperties> {
    crate::chart::model::ShapeProperties::from_xml(
        format!(r#"<c:spPr xmlns:c="{CHART_NS}" xmlns:a="{DRAWING_NS}">{body}</c:spPr>"#)
            .into_bytes(),
    )
}

/// The preset dash names `a:prstDash` admits.
const DASHES: &[&str] = &[
    "solid",
    "dot",
    "dash",
    "lgDash",
    "dashDot",
    "lgDashDot",
    "lgDashDotDot",
    "sysDash",
    "sysDot",
    "sysDashDot",
    "sysDashDotDot",
];

impl Spec {
    /// An empty chart of a variant, with a bottom legend and inline data.
    #[must_use]
    pub fn new(variant: Variant) -> Self {
        Self {
            variant,
            title: None,
            categories: Vec::new(),
            series: Vec::new(),
            legend: Some(LegendPosition::Bottom),
            category_axis_title: None,
            value_axis_title: None,
            binding: Binding::Literal,
            appearance: Appearance::default(),
            series_kinds: Vec::new(),
            series_references: Vec::new(),
        }
    }

    /// Set the family of each series of a combo chart, in order; series
    /// past the list are lines. Only column, line and area mix.
    #[must_use]
    pub fn with_series_kinds(mut self, kinds: Vec<Kind>) -> Self {
        self.series_kinds = kinds;
        self
    }

    /// The family of the series at `index` of a combo chart: the listed
    /// one, else column for the first series and line for the others.
    #[must_use]
    pub fn series_kind(&self, index: usize) -> Kind {
        match self.series_kinds.get(index) {
            Some(kind) => *kind,
            None if self.series_kinds.is_empty() && index == 0 => Kind::Column,
            None => Kind::Line,
        }
    }

    /// The formatting options.
    #[must_use]
    pub const fn appearance(&self) -> &Appearance {
        &self.appearance
    }

    /// Set the formatting options.
    #[must_use]
    pub fn with_appearance(mut self, appearance: Appearance) -> Self {
        self.appearance = appearance;
        self
    }

    /// The variant.
    #[must_use]
    pub const fn variant(&self) -> Variant {
        self.variant
    }

    /// The title, if any.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// The category labels (the X values of scatter and bubble charts, when
    /// they parse as numbers).
    #[must_use]
    pub fn categories(&self) -> &[String] {
        &self.categories
    }

    /// The series.
    #[must_use]
    pub fn series(&self) -> &[Series] {
        &self.series
    }

    /// The legend position; `None` hides the legend.
    #[must_use]
    pub const fn legend(&self) -> Option<LegendPosition> {
        self.legend
    }

    /// The data binding.
    #[must_use]
    pub const fn binding(&self) -> &Binding {
        &self.binding
    }

    /// The category (horizontal) axis title, if any.
    #[must_use]
    pub fn category_axis_title(&self) -> Option<&str> {
        self.category_axis_title.as_deref()
    }

    /// The value (vertical) axis title, if any.
    #[must_use]
    pub fn value_axis_title(&self) -> Option<&str> {
        self.value_axis_title.as_deref()
    }

    /// The series families of a combo chart, as given.
    #[must_use]
    pub fn series_kinds(&self) -> &[Kind] {
        &self.series_kinds
    }

    pub(crate) fn series_references(&self) -> &[(usize, String, Option<String>)] {
        &self.series_references
    }

    /// The same spec as another variant, keeping its data and formatting.
    #[must_use]
    pub const fn with_variant(mut self, variant: Variant) -> Self {
        self.variant = variant;
        self
    }

    /// The same spec without its title.
    #[must_use]
    pub fn without_title(mut self) -> Self {
        self.title = None;
        self
    }

    /// The same spec with other categories and series, keeping the rest.
    #[must_use]
    pub fn with_data(mut self, categories: Vec<String>, series: Vec<Series>) -> Self {
        self.categories = categories;
        self.series = series;
        self
    }

    /// The series, to change in place.
    pub fn series_mut(&mut self) -> &mut [Series] {
        &mut self.series
    }

    /// The formatting options, to change in place.
    pub const fn appearance_mut(&mut self) -> &mut Appearance {
        &mut self.appearance
    }

    /// Set the title.
    #[must_use]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Set the category labels.
    #[must_use]
    pub fn with_categories<I, S>(mut self, categories: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.categories = categories.into_iter().map(Into::into).collect();
        self
    }

    /// Append a series.
    #[must_use]
    pub fn with_series(mut self, series: Series) -> Self {
        self.series.push(series);
        self
    }

    /// Set the legend position; `None` hides the legend.
    #[must_use]
    pub const fn with_legend(mut self, legend: Option<LegendPosition>) -> Self {
        self.legend = legend;
        self
    }

    /// Set the category (horizontal) axis title.
    #[must_use]
    pub fn with_category_axis_title(mut self, title: impl Into<String>) -> Self {
        self.category_axis_title = Some(title.into());
        self
    }

    /// Set the value (vertical) axis title.
    #[must_use]
    pub fn with_value_axis_title(mut self, title: impl Into<String>) -> Self {
        self.value_axis_title = Some(title.into());
        self
    }

    /// Set the data binding.
    #[must_use]
    pub fn with_binding(mut self, binding: Binding) -> Self {
        self.binding = binding;
        self
    }

    /// Bind the values of series `index` to the worksheet range `values`
    /// (`Sheet1!$B$2:$B$13`), and its categories to `categories` when
    /// given; the series keeps its cached values (record 0338). Each
    /// reference overrides the [`Binding`]'s for that series, so a chart can
    /// refer to ranges anywhere in a workbook, another sheet included. A
    /// scatter or bubble chart ignores them.
    ///
    /// ```
    /// use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
    ///
    /// let spec = Spec::new(Variant::new(Kind::Line))
    ///     .with_categories(["Jan", "Feb"])
    ///     .with_series(Series::new("East", vec![120.0, 135.0]))
    ///     .with_series_reference(0, "Sheet1!$B$2:$B$3", Some("Sheet1!$A$2:$A$3".into()));
    /// let xml = String::from_utf8(spec.to_xml(None)?)?;
    /// assert!(xml.contains("<c:f>Sheet1!$B$2:$B$3</c:f>"));
    /// assert!(xml.contains("<c:f>Sheet1!$A$2:$A$3</c:f>"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[must_use]
    pub fn with_series_reference(
        mut self,
        index: usize,
        values: impl Into<String>,
        categories: Option<String>,
    ) -> Self {
        self.series_references.retain(|(held, _, _)| *held != index);
        self.series_references
            .push((index, values.into(), categories));
        self
    }

    /// Check the spec against the variant's data shape and the bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for an invalid variant, no series, a
    /// non-finite value, a stock chart without three or four series, a
    /// pie chart with more than one series or an empty value list, and
    /// [`Error::Limit`] past [`MAX_SERIES`], [`MAX_POINTS`],
    /// [`MAX_TOTAL_POINTS`] or [`MAX_TEXT_BYTES`].
    pub fn validate(&self) -> Result<()> {
        if !self.variant.is_valid() {
            return Err(Error::Invalid(
                "the chart family does not admit this grouping or 3D form".into(),
            ));
        }
        if self.series.is_empty() {
            return Err(Error::Invalid("a chart needs at least one series".into()));
        }
        for (index, values, categories) in &self.series_references {
            if *index >= self.series.len() {
                return Err(Error::Invalid("a series reference names no series".into()));
            }
            for reference in std::iter::once(values).chain(categories) {
                if reference.is_empty() || reference.len() > MAX_TEXT_BYTES {
                    return Err(Error::Invalid(
                        "a series reference needs 1 to MAX_TEXT_BYTES bytes".into(),
                    ));
                }
            }
        }
        if self.series.len() > MAX_SERIES {
            return Err(Error::limit(
                crate::LimitKind::ChartSeries,
                self.series.len(),
                MAX_SERIES,
            ));
        }
        if self.variant.kind == Kind::Stock && !matches!(self.series.len(), 3 | 4) {
            return Err(Error::Invalid(
                "a stock chart needs three (high, low, close) or four (open, high, low, close) series"
                    .into(),
            ));
        }
        if matches!(
            self.variant.kind,
            Kind::Waterfall | Kind::Funnel | Kind::Treemap | Kind::Sunburst | Kind::Pareto
        ) && self.series.len() != 1
        {
            return Err(Error::Invalid(
                "a waterfall, funnel, treemap, sunburst or Pareto chart takes one series".into(),
            ));
        }
        if self.categories.len() > MAX_POINTS {
            return Err(Error::limit(
                crate::LimitKind::ChartCategories,
                self.categories.len(),
                MAX_POINTS,
            ));
        }
        let texts = self
            .title
            .iter()
            .chain(&self.categories)
            .chain(self.category_axis_title.iter())
            .chain(self.value_axis_title.iter())
            .chain(self.series.iter().map(|series| &series.name));
        for text in texts {
            if text.len() > MAX_TEXT_BYTES {
                return Err(Error::limit(
                    crate::LimitKind::ChartTextBytes,
                    text.len(),
                    MAX_TEXT_BYTES,
                ));
            }
        }
        if let Binding::Sheet { sheet, row, column } = &self.binding {
            if sheet.is_empty() || sheet.chars().count() > 31 {
                return Err(Error::Invalid(
                    "a worksheet name needs 1 to 31 characters".into(),
                ));
            }
            let rows = self.rows().saturating_add(1);
            let columns = self.series.len().saturating_mul(3).saturating_add(1);
            if row.saturating_add(rows) > 1_048_576 || column.saturating_add(columns) > 16_384 {
                return Err(Error::Invalid(
                    "the chart data does not fit on a worksheet at this origin".into(),
                ));
            }
        }
        if self.variant.kind == Kind::Combo {
            if self.series.len() < 2 {
                return Err(Error::Invalid(
                    "a combo chart needs at least two series".into(),
                ));
            }
            if self.series_kinds.len() > self.series.len()
                || self
                    .series_kinds
                    .iter()
                    .any(|kind| !matches!(kind, Kind::Column | Kind::Line | Kind::Area))
            {
                return Err(Error::Invalid(
                    "a combo chart mixes column, line and area series, one per series".into(),
                ));
            }
        } else if !self.series_kinds.is_empty() {
            return Err(Error::Invalid(
                "only a combo chart lists series families".into(),
            ));
        }
        self.validate_appearance()?;
        let mut total = 0_usize;
        for series in &self.series {
            if series.values.is_empty() {
                return Err(Error::Invalid(format!(
                    "series '{}' has no values",
                    series.name
                )));
            }
            if series.color.as_deref().is_some_and(|color| !is_argb(color)) {
                return Err(Error::Invalid(format!(
                    "series '{}' has a color that is not RRGGBB or AARRGGBB",
                    series.name
                )));
            }
            if series.values.len() > MAX_POINTS {
                return Err(Error::limit(
                    crate::LimitKind::ChartPoints,
                    series.values.len(),
                    MAX_POINTS,
                ));
            }
            if series.values.iter().any(|value| !value.is_finite()) {
                return Err(Error::Invalid(format!(
                    "series '{}' has a non-finite value",
                    series.name
                )));
            }
            total = total.saturating_add(series.values.len());
        }
        if total > MAX_TOTAL_POINTS {
            return Err(Error::limit(
                crate::LimitKind::ChartPoints,
                total,
                MAX_TOTAL_POINTS,
            ));
        }
        Ok(())
    }

    fn validate_appearance(&self) -> Result<()> {
        let appearance = &self.appearance;
        let invalid = |message: &str| Err(Error::Invalid(message.into()));
        if appearance.colors.len() > MAX_POINTS
            || appearance.colors.iter().any(|color| !is_argb(color))
        {
            return invalid("chart colors must be RRGGBB hex values");
        }
        for line in appearance
            .gridlines
            .iter()
            .chain(appearance.minor_gridlines.iter())
            .flatten()
        {
            if !is_rgb(&line.color)
                || line
                    .width
                    .is_some_and(|width| !(0.0..=1584.0).contains(&width))
                || line
                    .dash
                    .as_deref()
                    .is_some_and(|dash| !DASHES.contains(&dash))
            {
                return invalid(
                    "a gridline needs an RRGGBB color, a 0-1584 pt width and a preset dash",
                );
            }
        }
        if appearance.gap_width.is_some_and(|value| value > 500)
            || appearance
                .overlap
                .is_some_and(|value| !(-100..=100).contains(&value))
            || appearance
                .hole_size
                .is_some_and(|value| !(1..=90).contains(&value))
            || appearance
                .first_slice_angle
                .is_some_and(|value| value > 360)
            || appearance.explosion.is_some_and(|value| value > 400)
        {
            return invalid("a chart format value is outside its range");
        }
        for value in [
            appearance.axis_min,
            appearance.axis_max,
            appearance.major_unit,
        ]
        .into_iter()
        .flatten()
        {
            if !value.is_finite() {
                return invalid("axis bounds must be finite");
            }
        }
        if appearance.major_unit.is_some_and(|unit| unit <= 0.0)
            || appearance
                .axis_min
                .zip(appearance.axis_max)
                .is_some_and(|(min, max)| min >= max)
        {
            return invalid("the axis minimum must be below its maximum and the unit positive");
        }
        if appearance
            .marker_size
            .is_some_and(|size| !(2..=72).contains(&size))
            || appearance
                .marker_color
                .as_deref()
                .is_some_and(|color| !is_rgb(color))
            || appearance
                .line_width
                .is_some_and(|width| !(0.0..=1584.0).contains(&width))
            || appearance.view_3d.is_some_and(|(x, y, perspective)| {
                x > 90 || y > 360 || perspective.is_some_and(|p| p > 240)
            })
        {
            return invalid("a marker, line width or 3D view value is outside its range");
        }
        if [
            &appearance.increase_color,
            &appearance.decrease_color,
            &appearance.total_color,
        ]
        .into_iter()
        .flatten()
        .any(|color| !is_argb(color))
        {
            return invalid("chart colors must be RRGGBB hex values");
        }
        if let Some(bins) = appearance.bins {
            let finite = [bins.size, bins.overflow, bins.underflow]
                .into_iter()
                .flatten()
                .all(f64::is_finite);
            if !finite
                || bins.count == Some(0)
                || bins.size.is_some_and(|size| size <= 0.0)
                || (bins.count.is_some() && bins.size.is_some())
                || bins
                    .underflow
                    .zip(bins.overflow)
                    .is_some_and(|(under, over)| under >= over)
            {
                return invalid(
                    "a histogram takes a positive bin count or a positive bin size, and an underflow below its overflow",
                );
            }
        }
        if let Some(subtotals) = &appearance.subtotals {
            let points = self.series.first().map_or(0, |series| series.values.len());
            if subtotals.iter().any(|point| *point >= points) {
                return invalid("a waterfall total names a point past the series");
            }
        }
        if let Some(position) = appearance.labels.and_then(|labels| labels.position) {
            use DataLabelPosition as P;
            let stacked = !matches!(self.variant.grouping, Grouping::Standard);
            let allowed = match self.variant.kind {
                Kind::Column | Kind::Bar if stacked => {
                    matches!(position, P::Center | P::InsideEnd | P::InsideBase)
                },
                Kind::Column | Kind::Bar => {
                    matches!(
                        position,
                        P::Center | P::InsideEnd | P::InsideBase | P::OutsideEnd
                    )
                },
                Kind::Line | Kind::Scatter | Kind::Stock | Kind::Bubble => matches!(
                    position,
                    P::Center | P::Left | P::Right | P::Top | P::Bottom
                ),
                Kind::Pie => matches!(
                    position,
                    P::Center | P::InsideEnd | P::OutsideEnd | P::BestFit
                ),
                Kind::Doughnut | Kind::Area | Kind::Radar | Kind::Combo => false,
                // `cx:dataLabels/@pos` admits every position.
                Kind::Waterfall
                | Kind::Funnel
                | Kind::Treemap
                | Kind::Sunburst
                | Kind::Histogram
                | Kind::Pareto
                | Kind::BoxWhisker => true,
            };
            if !allowed || self.variant.three_d {
                return invalid("this chart type does not take that data-label position");
            }
        }
        Ok(())
    }

    /// Whether the chart plots X/Y values rather than categories.
    const fn is_xy(&self) -> bool {
        matches!(self.variant.kind, Kind::Scatter | Kind::Bubble)
    }

    /// The X values of a scatter or bubble chart: the categories as numbers,
    /// or 1, 2, 3, ... where a label is absent or not a number.
    #[must_use]
    pub fn x_values(&self) -> Vec<f64> {
        let points = self
            .series
            .iter()
            .map(|series| series.values.len())
            .max()
            .unwrap_or(0);
        (0..points)
            .map(|index| {
                self.categories
                    .get(index)
                    .and_then(|label| label.trim().parse::<f64>().ok())
                    .filter(|value| value.is_finite())
                    .unwrap_or((index + 1) as f64)
            })
            .collect()
    }

    /// The number of category rows of the sheet layout.
    fn rows(&self) -> usize {
        self.series
            .iter()
            .map(|series| series.values.len())
            .chain(std::iter::once(self.categories.len()))
            .max()
            .unwrap_or(0)
    }

    /// The worksheet a [`Binding::Sheet`] chart refers to, as `(row, column,
    /// cell)` triples with zero-based coordinates relative to the binding's
    /// origin, in row-major order.
    ///
    /// Category charts put the category labels in column A from row 2 and
    /// series `n` in column `n + 1` with its name in row 1. Scatter charts
    /// give series `n` the X column `2n` and the Y column `2n + 1`; bubble
    /// charts add the size column `3n + 2` (with X `3n` and Y `3n + 1`),
    /// the sizes being the Y values.
    #[must_use]
    pub fn sheet_layout(&self) -> Vec<(usize, usize, Cell<'_>)> {
        let mut cells = Vec::new();
        let stride = match self.variant.kind {
            Kind::Bubble => 3,
            Kind::Scatter => 2,
            _ => 1,
        };
        let x_values = if self.is_xy() {
            self.x_values()
        } else {
            Vec::new()
        };
        for (index, series) in self.series.iter().enumerate() {
            let first = if self.is_xy() {
                index * stride
            } else {
                index + 1
            };
            if self.is_xy() {
                cells.push((0, first, Cell::Text(&series.name)));
            }
            cells.push((
                0,
                first + usize::from(self.is_xy()),
                Cell::Text(&series.name),
            ));
            if self.variant.kind == Kind::Bubble {
                cells.push((0, first + 2, Cell::Text(&series.name)));
            }
        }
        for row in 0..self.rows() {
            if !self.is_xy()
                && let Some(label) = self.categories.get(row)
            {
                cells.push((row + 1, 0, Cell::Text(label)));
            }
            for (index, series) in self.series.iter().enumerate() {
                let Some(&value) = series.values.get(row) else {
                    continue;
                };
                if self.is_xy() {
                    let first = index * stride;
                    if let Some(&x) = x_values.get(row) {
                        cells.push((row + 1, first, Cell::Number(x)));
                    }
                    cells.push((row + 1, first + 1, Cell::Number(value)));
                    if self.variant.kind == Kind::Bubble {
                        cells.push((row + 1, first + 2, Cell::Number(value)));
                    }
                } else {
                    cells.push((row + 1, index + 1, Cell::Number(value)));
                }
            }
        }
        cells
    }

    pub(crate) fn reference(
        &self,
        column: usize,
        first_row: usize,
        last_row: usize,
    ) -> Option<String> {
        let Binding::Sheet {
            sheet,
            row: row_origin,
            column: column_origin,
        } = &self.binding
        else {
            return None;
        };
        let name = quote_sheet(sheet);
        let column = column_name(column + column_origin);
        let (first_row, last_row) = (first_row + row_origin, last_row + row_origin);
        Some(if first_row == last_row {
            format!("{name}!${column}${}", first_row + 1)
        } else {
            format!(
                "{name}!${column}${}:${column}${}",
                first_row + 1,
                last_row + 1
            )
        })
    }

    fn numbers(&self, column: usize, values: Vec<f64>) -> NumericData {
        let last = values.len();
        let mut data = NumericData::from_values(values);
        data.format_code = Some("General".into());
        data.source_ref = self
            .reference(column, 1, last)
            .map(crate::chart::data::DataSourceRef::new);
        data
    }

    fn labels(&self, values: Vec<String>) -> StringData {
        let last = values.len();
        let mut data = StringData::from_values(values);
        data.source_ref = self
            .reference(0, 1, last)
            .map(crate::chart::data::DataSourceRef::new);
        data
    }

    fn build_series(&self) -> Vec<ChartSeries> {
        let stride = match self.variant.kind {
            Kind::Bubble => 3,
            _ => 2,
        };
        let x_values = if self.is_xy() {
            self.x_values()
        } else {
            Vec::new()
        };
        self.series
            .iter()
            .enumerate()
            .map(|(index, series)| {
                let position = u32::try_from(index).unwrap_or(u32::MAX);
                let mut built = ChartSeries::new(position);
                built.title = Some(TitleText::from_string(series.name.clone()));
                let points = series.values.len();
                let appearance = &self.appearance;
                let per_point = matches!(self.variant.kind, Kind::Pie | Kind::Doughnut)
                    && self.series.len() == 1
                    && appearance.colors.len() > 1;
                if per_point {
                    for (point, color) in appearance.colors.iter().take(points).enumerate() {
                        let mut data_point =
                            crate::chart::series::DataPoint::new(u32::try_from(point).unwrap_or(u32::MAX));
                        data_point.shape_properties =
                            shape_properties(&format!("<a:solidFill>{}</a:solidFill>", srgb(color)))
                                .ok();
                        built.data_points.push(data_point);
                    }
                    // The series' own color fills the whole series; the
                    // point colors above override it point by point.
                    if let Some(color) = &series.color {
                        built.shape_properties =
                            shape_properties(&format!("<a:solidFill>{}</a:solidFill>", srgb(color)))
                                .ok();
                    }
                } else {
                    let family = if self.variant.kind == Kind::Combo {
                        self.series_kind(index)
                    } else {
                        self.variant.kind
                    };
                    let lined = matches!(family, Kind::Line | Kind::Scatter)
                        || (family == Kind::Radar
                            && appearance.radar_style != Some(RadarStyle::Filled));
                    let color = series
                        .color
                        .as_ref()
                        .or_else(|| appearance.colors.get(index))
                        .map(|color| srgb(color));
                    let width = appearance
                        .line_width
                        .map_or(28_575, |points| (points * 12_700.0).round() as i64);
                    let body = match (&color, lined) {
                        (Some(color), true) => Some(format!(
                            r#"<a:ln w="{width}" cap="rnd"><a:solidFill>{color}</a:solidFill><a:round/></a:ln>"#
                        )),
                        (None, true) if appearance.line_width.is_some() => {
                            Some(format!(r#"<a:ln w="{width}" cap="rnd"><a:round/></a:ln>"#))
                        },
                        (Some(color), false) => Some(format!("<a:solidFill>{color}</a:solidFill>")),
                        (None, _) => None,
                    };
                    if let Some(body) = body {
                        built.shape_properties = shape_properties(&body).ok();
                    }
                    if matches!(family, Kind::Line | Kind::Scatter | Kind::Radar)
                        && (appearance.marker.is_some()
                            || appearance.marker_size.is_some()
                            || appearance.marker_color.is_some())
                    {
                        built.marker_present = true;
                        built.marker_symbol = appearance.marker;
                        built.marker_size = appearance.marker_size;
                        built.marker_shape_properties = appearance
                            .marker_color
                            .as_ref()
                            .and_then(|color| {
                                shape_properties(&format!(
                                    r#"<a:solidFill><a:srgbClr val="{}"/></a:solidFill>"#,
                                    color.to_ascii_uppercase()
                                ))
                                .ok()
                            });
                    }
                }
                let family = if self.variant.kind == Kind::Combo {
                    self.series_kind(index)
                } else {
                    self.variant.kind
                };
                if matches!(family, Kind::Line | Kind::Scatter) {
                    built.smooth = appearance.smooth.unwrap_or(false);
                }
                if matches!(self.variant.kind, Kind::Pie | Kind::Doughnut) {
                    built.explosion = appearance.explosion;
                }
                if self.is_xy() {
                    let first = index * stride;
                    let x = x_values.get(..points).unwrap_or(&x_values).to_vec();
                    built.x_values = Some(self.numbers(first, x));
                    built.y_values = Some(self.numbers(first + 1, series.values.clone()));
                    if self.variant.kind == Kind::Bubble {
                        built.bubble_sizes = Some(self.numbers(first + 2, series.values.clone()));
                    }
                } else {
                    let reference = self
                        .series_references
                        .iter()
                        .find(|(held, _, _)| *held == index);
                    if !self.categories.is_empty() {
                        let mut labels = self.labels(self.categories.clone());
                        if let Some((_, _, Some(categories))) = reference {
                            labels.source_ref =
                                Some(crate::chart::data::DataSourceRef::new(categories.clone()));
                        }
                        built.categories = Some(labels);
                    }
                    let mut values = self.numbers(index + 1, series.values.clone());
                    if let Some((_, reference, _)) = reference {
                        values.source_ref =
                            Some(crate::chart::data::DataSourceRef::new(reference.clone()));
                    }
                    built.values = Some(values);
                }
                built
            })
            .collect()
    }

    fn common(&self, vary_colors: bool) -> TypeGroupCommon {
        let mut common = TypeGroupCommon::new();
        common.vary_colors = self.appearance.vary_colors.unwrap_or(vary_colors);
        common.series = self.build_series();
        if let Some(labels) = self.appearance.labels {
            let mut data_labels = DataLabels::new();
            data_labels.show_value = labels.value;
            data_labels.show_category_name = labels.category;
            data_labels.show_series_name = labels.series;
            data_labels.show_percent = labels.percent;
            data_labels.position = labels.position;
            data_labels.deleted =
                !(labels.value || labels.category || labels.series || labels.percent);
            common.data_labels = Some(data_labels);
        }
        common
    }

    /// The chart groups: one, or for a combo chart one per run of series
    /// of the same family.
    fn type_groups(&self) -> Vec<TypeGroup> {
        if self.variant.kind != Kind::Combo {
            return vec![self.type_group()];
        }
        let series = self.build_series();
        let mut groups = Vec::new();
        let mut start = 0;
        while start < series.len() {
            let family = self.series_kind(start);
            let mut end = start + 1;
            while end < series.len() && self.series_kind(end) == family {
                end += 1;
            }
            let mut common = TypeGroupCommon::new();
            common.series = series[start..end].to_vec();
            common.axis_ids = vec![CATEGORY_AXIS_ID, VALUE_AXIS_ID];
            if let Some(labels) = self.common(false).data_labels {
                common.data_labels = Some(labels);
            }
            groups.push(match family {
                Kind::Line => {
                    let mut group = LineTypeGroup::new(BarGrouping::Standard);
                    group.common = common;
                    group.smooth = self.appearance.smooth.unwrap_or(false);
                    TypeGroup::Line(group)
                },
                Kind::Area => {
                    let mut group = AreaTypeGroup::new(BarGrouping::Standard);
                    group.common = common;
                    TypeGroup::Area(group)
                },
                _ => {
                    let mut group = BarTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
                    group.common = common;
                    group.gap_width = Some(self.appearance.gap_width.unwrap_or(150));
                    group.overlap = self.appearance.overlap;
                    TypeGroup::Bar(group)
                },
            });
            start = end;
        }
        // Areas draw behind columns, columns behind lines.
        groups.sort_by_key(|group| match group {
            TypeGroup::Area(_) => 0,
            TypeGroup::Bar(_) => 1,
            _ => 2,
        });
        groups
    }

    fn type_group(&self) -> TypeGroup {
        let grouping = match (self.variant.kind, self.variant.grouping) {
            (_, Grouping::Stacked) => BarGrouping::Stacked,
            (_, Grouping::PercentStacked) => BarGrouping::PercentStacked,
            (Kind::Column | Kind::Bar, Grouping::Standard) => BarGrouping::Clustered,
            (_, Grouping::Standard) => BarGrouping::Standard,
        };
        let axes = |mut common: TypeGroupCommon, series_axis: bool| {
            common.axis_ids = vec![CATEGORY_AXIS_ID, VALUE_AXIS_ID];
            if series_axis {
                common.axis_ids.push(SERIES_AXIS_ID);
            }
            common
        };
        let three_d = self.variant.three_d;
        match self.variant.kind {
            Kind::Column | Kind::Bar => {
                let direction = if self.variant.kind == Kind::Bar {
                    BarDirection::Bar
                } else {
                    BarDirection::Column
                };
                if three_d {
                    let mut group = Bar3DTypeGroup::new(direction, grouping);
                    group.common = axes(self.common(false), false);
                    group.gap_width = Some(self.appearance.gap_width.unwrap_or(150));
                    TypeGroup::Bar3D(group)
                } else {
                    let mut group = BarTypeGroup::new(direction, grouping);
                    group.common = axes(self.common(false), false);
                    group.gap_width = Some(self.appearance.gap_width.unwrap_or(150));
                    if grouping != BarGrouping::Clustered {
                        group.overlap = Some(100);
                    }
                    if let Some(overlap) = self.appearance.overlap {
                        group.overlap = Some(overlap);
                    }
                    TypeGroup::Bar(group)
                }
            },
            Kind::Line => {
                if three_d {
                    let mut group = Line3DTypeGroup::new(grouping);
                    group.common = axes(self.common(false), true);
                    TypeGroup::Line3D(group)
                } else {
                    let mut group = LineTypeGroup::new(grouping);
                    group.common = axes(self.common(false), false);
                    group.smooth = self.appearance.smooth.unwrap_or(false);
                    TypeGroup::Line(group)
                }
            },
            Kind::Area => {
                if three_d {
                    let mut group = Area3DTypeGroup::new(grouping);
                    group.common = axes(self.common(false), grouping == BarGrouping::Standard);
                    TypeGroup::Area3D(group)
                } else {
                    let mut group = AreaTypeGroup::new(grouping);
                    group.common = axes(self.common(false), false);
                    TypeGroup::Area(group)
                }
            },
            Kind::Pie => {
                if three_d {
                    let mut group = Pie3DTypeGroup::new();
                    group.common = self.common(true);
                    TypeGroup::Pie3D(group)
                } else {
                    let mut group = PieTypeGroup::new();
                    group.common = self.common(true);
                    group.first_slice_angle = self.appearance.first_slice_angle.unwrap_or(0);
                    TypeGroup::Pie(group)
                }
            },
            Kind::Doughnut => {
                let mut group = DoughnutTypeGroup::new();
                group.common = self.common(true);
                group.first_slice_angle = self.appearance.first_slice_angle.unwrap_or(0);
                group.hole_size = self.appearance.hole_size.unwrap_or(50);
                TypeGroup::Doughnut(group)
            },
            Kind::Scatter => {
                let mut group = ScatterTypeGroup::new(
                    self.appearance
                        .scatter_style
                        .unwrap_or(ScatterStyle::LineMarker),
                );
                group.common = axes(self.common(false), false);
                TypeGroup::Scatter(group)
            },
            Kind::Bubble => {
                let mut group = BubbleTypeGroup::new();
                group.common = axes(self.common(false), false);
                TypeGroup::Bubble(group)
            },
            Kind::Radar => {
                let mut group =
                    RadarTypeGroup::new(self.appearance.radar_style.unwrap_or(RadarStyle::Marker));
                group.common = axes(self.common(false), false);
                TypeGroup::Radar(group)
            },
            Kind::Combo => unreachable!("combo charts build several groups"),
            Kind::Waterfall
            | Kind::Funnel
            | Kind::Treemap
            | Kind::Sunburst
            | Kind::Histogram
            | Kind::Pareto
            | Kind::BoxWhisker => unreachable!("extended charts have no classic chart group"),
            Kind::Stock => {
                let mut group = StockTypeGroup::new();
                group.common = axes(self.common(false), false);
                group.high_low_lines = Some(crate::chart::plot_area::Lines::new());
                TypeGroup::Stock(group)
            },
        }
    }

    fn axes(&self, groups: &[TypeGroup]) -> Vec<Axis> {
        let titled = |title: &Option<String>| title.as_ref().map(TitleText::from_string);
        match self.variant.kind {
            Kind::Pie | Kind::Doughnut => Vec::new(),
            Kind::Scatter | Kind::Bubble => {
                let mut x = ValueAxis::new(CATEGORY_AXIS_ID, AxisPosition::Bottom, VALUE_AXIS_ID);
                x.common.title = titled(&self.category_axis_title);
                let mut y = ValueAxis::new(VALUE_AXIS_ID, AxisPosition::Left, CATEGORY_AXIS_ID);
                y.common.title = titled(&self.value_axis_title);
                self.format_value_axis(&mut y);
                vec![Axis::Value(x), Axis::Value(y)]
            },
            _ => {
                let horizontal = self.variant.kind == Kind::Bar;
                let (category_position, value_position) = if horizontal {
                    (AxisPosition::Left, AxisPosition::Bottom)
                } else {
                    (AxisPosition::Bottom, AxisPosition::Left)
                };
                let mut category =
                    CategoryAxis::new(CATEGORY_AXIS_ID, category_position, VALUE_AXIS_ID);
                category.common.title = titled(&self.category_axis_title);
                let mut value = ValueAxis::new(VALUE_AXIS_ID, value_position, CATEGORY_AXIS_ID);
                value.common.title = titled(&self.value_axis_title);
                self.format_value_axis(&mut value);
                let mut axes = vec![Axis::Category(category), Axis::Value(value)];
                if groups
                    .iter()
                    .any(|group| group.common().axis_ids.contains(&SERIES_AXIS_ID))
                {
                    let mut series =
                        SeriesAxis::new(SERIES_AXIS_ID, AxisPosition::Bottom, VALUE_AXIS_ID);
                    series.common.deleted = true;
                    axes.push(Axis::Series(series));
                }
                axes
            },
        }
    }

    fn format_value_axis(&self, axis: &mut ValueAxis) {
        let appearance = &self.appearance;
        axis.min = appearance.axis_min;
        axis.max = appearance.axis_max;
        axis.major_unit = appearance.major_unit;
        axis.common.show_major_gridlines = !appearance.hide_gridlines;
        if !appearance.hide_gridlines
            && let Some(Some(line)) = &appearance.gridlines
        {
            axis.common.major_gridlines = Some(crate::chart::plot_area::Lines {
                shape_properties: shape_properties(&line_xml(line)).ok(),
            });
        }
        if let Some(minor) = &appearance.minor_gridlines {
            axis.common.show_minor_gridlines = true;
            if let Some(line) = minor {
                axis.common.minor_gridlines = Some(crate::chart::plot_area::Lines {
                    shape_properties: shape_properties(&line_xml(line)).ok(),
                });
            }
        }
    }

    /// Build the typed chart model.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::validate`].
    pub fn to_chart(&self) -> Result<Chart> {
        self.validate()?;
        if self.variant.kind.is_extended() {
            return Err(Error::Invalid(
                "an extended chart has no classic chart model; write it with to_xml".into(),
            ));
        }
        let groups = self.type_groups();
        let mut plot_area = PlotArea::new();
        for axis in self.axes(&groups) {
            plot_area = plot_area.add_axis(axis);
        }
        for group in groups {
            plot_area = plot_area.add_type_group(group);
        }
        let mut chart = Chart::new().with_plot_area(plot_area);
        chart.title = self.title.as_ref().map(TitleText::from_string);
        chart.auto_title_deleted = self
            .appearance
            .auto_title_deleted
            .unwrap_or(self.title.is_none());
        chart.legend = self.legend.map(|position| {
            Legend::new(position).with_overlay(self.appearance.legend_overlay.unwrap_or(false))
        });
        chart.language = Some("en-US".into());
        chart.rounded_corners = false;
        if let Some((rot_x, rot_y, perspective)) = self.appearance.view_3d {
            let mut view = View3D::new();
            view.rot_x = Some(rot_x);
            view.rot_y = Some(rot_y);
            view.perspective = perspective;
            chart.view_3d = Some(view);
        } else if self.variant.three_d {
            let mut view = View3D::new();
            if self.variant.kind == Kind::Pie {
                view.rot_x = Some(30);
                view.rot_y = Some(0);
                view.right_angle_axes = false;
            } else {
                view.rot_x = Some(15);
                view.rot_y = Some(20);
            }
            chart.view_3d = Some(view);
        }
        Ok(chart)
    }

    /// Write the complete chart-space part: a `c:chartSpace`, or for an
    /// extended family ([`Kind::is_extended`]) a `cx:chartSpace`.
    ///
    /// `external_data` is the relationship ID of the embedded workbook the
    /// host attached, written as `c:externalData`; pass it with a
    /// [`Binding::Sheet`] spec.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::validate`] and of the chart writer.
    pub fn to_xml(&self, external_data: Option<&str>) -> Result<Vec<u8>> {
        self.to_xml_with_policy(external_data, &crate::chart::Limits::DEFAULT)
    }

    pub(crate) fn to_xml_with_policy(
        &self,
        external_data: Option<&str>,
        limits: &crate::chart::Limits,
    ) -> Result<Vec<u8>> {
        let limits = limits.validate()?;
        limits.validate_spec_for_chart(self)?;
        if self.variant.kind.is_extended() {
            return crate::chart::chartex::write_with_limits(self, external_data, &limits);
        }
        let mut chart = self.to_chart()?;
        if external_data.is_some() {
            chart.external_data = Some(crate::chart::model::ExternalData::pending());
        }
        limits.validate_chart(&chart)?;
        let mut xml = Vec::with_capacity(4096.min(limits.max_output_bytes));
        crate::chart::writer::write_with_rels(&mut xml, &chart, &limits, external_data, None)?;
        Ok(xml)
    }
}

/// Quote a worksheet name for a formula when it is not a plain identifier.
fn quote_sheet(name: &str) -> String {
    let plain = name
        .chars()
        .next()
        .is_some_and(|first| first.is_ascii_alphabetic() || first == '_')
        && name
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || character == '_');
    if plain {
        name.to_owned()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// The spreadsheet column name of a zero-based column index (`0` is `A`).
#[must_use]
pub fn column_name(mut index: usize) -> String {
    let mut name = Vec::new();
    loop {
        name.push(b'A' + (index % 26) as u8);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    name.reverse();
    String::from_utf8(name).unwrap_or_default()
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests;
