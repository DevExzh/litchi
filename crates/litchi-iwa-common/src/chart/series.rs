//! Archive-free per-series chart values: strokes, fills, connection lines,
//! data symbols and trendlines, and the native field families a chart kind
//! stores them in (record 0375).
//!
//! The field numbers of the families are those of the generated
//! `TSCH.Generated.ChartSeriesStyleArchive` (strokes, fills, symbol sizes,
//! fills and outlines) and `ChartSeriesNonStyleArchive` (connection lines,
//! symbol visibility and shape, trendlines) extensions, for the format owners'
//! private adapters over `litchi_iwa_protos::chart_style_codec`.

use crate::chart::kind::Kind;
use crate::color::Rgba;
use crate::shape::fill::Fill;
use crate::shape::stroke::{Join, Pattern, Stroke as ShapeStroke, Width};

/// Validation failures of the series values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A shape stroke pattern has no chart-series equivalent.
    #[error("the stroke pattern is not a chart series stroke pattern")]
    UnsupportedPattern,
    /// A stored enumeration holds an unknown value.
    #[error("unsupported native series value {0}")]
    UnknownNative(i32),
    /// A symbol size was not finite and positive.
    #[error("chart series symbol size must be finite and positive")]
    InvalidSymbolSize,
    /// The connection line is not valid for the family (only scatter charts
    /// hide connection lines).
    #[error("the series family cannot hide connection lines")]
    HiddenConnectionLine,
    /// A trendline polynomial order outside `2..=6`.
    #[error("trendline polynomial order must be 2..=6")]
    PolynomialOrder,
    /// A trendline moving-average period outside `2..=i32::MAX`.
    #[error("trendline moving-average period must be at least 2")]
    MovingAveragePeriod,
    /// A known trendline type spelled as `Unsupported`.
    #[error("a known trendline type must use its named representation")]
    KnownTrendlineAsUnsupported,
    /// Display options on a hidden trendline.
    #[error("a hidden trendline cannot expose display options")]
    HiddenTrendlineOptions,
    /// Equation or R² labels on a moving-average trendline.
    #[error("moving-average trendlines have no equation or R² labels")]
    MovingAverageLabels,
    /// A stored trendline is visible without a type, or hidden with one.
    #[error("the stored trendline visibility and type disagree")]
    TrendlineVisibility,
    /// A stored boolean is not 0 or 1.
    #[error("a stored trendline flag is not a canonical boolean")]
    NonCanonicalBool,
    /// Visible scatter connection lines store the hidden line type.
    #[error("visible scatter connection lines use the hidden line type")]
    VisibleHiddenLine,
    /// A symbol fill follows both the series fill and the series stroke.
    #[error("a data-symbol fill follows both the series fill and stroke")]
    BothSeriesSources,
}

/// Result type of the series values.
pub type Result<T> = std::result::Result<T, Error>;

/// Stroke patterns of the chart-series inspector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum StrokePattern {
    /// A solid line.
    #[default]
    Solid,
    /// Medium dashes.
    MediumDash,
    /// Rounded dots.
    RoundedDash,
}

/// The stroke of one chart series.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Stroke {
    /// Stroke color.
    pub color: Rgba,
    /// Stroke width.
    pub width: Width,
    /// Stroke pattern.
    pub pattern: StrokePattern,
}

impl Stroke {
    /// Construct a series stroke.
    #[must_use]
    pub const fn new(color: Rgba, width: Width, pattern: StrokePattern) -> Self {
        Self {
            color,
            width,
            pattern,
        }
    }

    /// The series stroke a stored shape stroke shows.
    ///
    /// # Errors
    ///
    /// [`Error::UnsupportedPattern`] for short and long dashes.
    pub fn from_shape_stroke(stroke: ShapeStroke) -> Result<Self> {
        let pattern = match stroke.pattern {
            Pattern::Solid => StrokePattern::Solid,
            Pattern::MediumDash => StrokePattern::MediumDash,
            Pattern::RoundedDash => StrokePattern::RoundedDash,
            Pattern::ShortDash | Pattern::LongDash => return Err(Error::UnsupportedPattern),
        };
        Ok(Self::new(stroke.color, stroke.width, pattern))
    }

    /// The shape stroke stored for this series stroke (miter join).
    #[must_use]
    pub fn to_shape_stroke(self) -> ShapeStroke {
        let pattern = match self.pattern {
            StrokePattern::Solid => Pattern::Solid,
            StrokePattern::MediumDash => Pattern::MediumDash,
            StrokePattern::RoundedDash => Pattern::RoundedDash,
        };
        ShapeStroke::new(self.color, self.width, pattern).with_join(Join::Miter)
    }
}

/// Where a chart kind stores its series strokes (series style).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StrokeFamily {
    /// The `Area2d` family.
    Area2d,
    /// The `BarOrColumn2d` family.
    BarOrColumn2d,
    /// The `Bubble2d` family.
    Bubble2d,
    /// The `Line2d` family.
    Line2d,
    /// The `Pie2d` family.
    Pie2d,
    /// The `Radar2d` family.
    Radar2d,
    /// The `Scatter2d` family.
    Scatter2d,
}

impl StrokeFamily {
    /// The extension field that stores the family's value.
    #[must_use]
    pub const fn field(self) -> u32 {
        match self {
            Self::Area2d => 45,
            Self::BarOrColumn2d => 46,
            Self::Bubble2d => 47,
            Self::Line2d => 48,
            Self::Pie2d => 52,
            Self::Radar2d => 172,
            Self::Scatter2d => 53,
        }
    }
}

impl StrokeFamily {
    /// The family an unambiguous chart kind uses, if any.
    #[must_use]
    pub const fn for_chart_kind(kind: Kind) -> Option<Self> {
        Some(match kind {
            Kind::Area2d | Kind::StackedArea2d => Self::Area2d,
            Kind::Bar2d
            | Kind::StackedBar2d
            | Kind::MultiDataBar2d
            | Kind::Column2d
            | Kind::StackedColumn2d
            | Kind::MultiDataColumn2d => Self::BarOrColumn2d,
            Kind::Bubble2d => Self::Bubble2d,
            Kind::Line2d => Self::Line2d,
            Kind::Pie2d | Kind::Donut2d => Self::Pie2d,
            Kind::Radar2d => Self::Radar2d,
            Kind::Scatter2d => Self::Scatter2d,
            _ => return None,
        })
    }
}

/// Where a chart kind stores its series fills (series style).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FillFamily {
    /// The `Area2d` family.
    Area2d,
    /// The `Bar2d` family.
    Bar2d,
    /// The `Column2d` family.
    Column2d,
    /// The `Pie2d` family.
    Pie2d,
    /// The `Radar2d` family.
    Radar2d,
    /// The `Area3d` family.
    Area3d,
    /// The `Bar3d` family.
    Bar3d,
    /// The `Column3d` family.
    Column3d,
    /// The `Line3d` family.
    Line3d,
    /// The `Pie3d` family.
    Pie3d,
}

impl FillFamily {
    /// The extension field that stores the family's value.
    #[must_use]
    pub const fn field(self) -> u32 {
        match self {
            Self::Area2d => 11,
            Self::Bar2d => 12,
            Self::Column2d => 13,
            Self::Pie2d => 17,
            Self::Radar2d => 165,
            Self::Area3d => 6,
            Self::Bar3d => 7,
            Self::Column3d => 8,
            Self::Line3d => 9,
            Self::Pie3d => 10,
        }
    }
}

impl FillFamily {
    /// The series-style field of the kind-independent fallback fill.
    pub const DEFAULT_FIELD: u32 = 14;

    /// The family an unambiguous chart kind uses, if any.
    #[must_use]
    pub const fn for_chart_kind(kind: Kind) -> Option<Self> {
        Some(match kind {
            Kind::Area2d | Kind::StackedArea2d => Self::Area2d,
            Kind::Bar2d | Kind::StackedBar2d | Kind::MultiDataBar2d => Self::Bar2d,
            Kind::Column2d | Kind::StackedColumn2d | Kind::MultiDataColumn2d => Self::Column2d,
            Kind::Pie2d | Kind::Donut2d => Self::Pie2d,
            Kind::Radar2d => Self::Radar2d,
            Kind::Area3d | Kind::StackedArea3d => Self::Area3d,
            Kind::Bar3d | Kind::StackedBar3d => Self::Bar3d,
            Kind::Column3d | Kind::StackedColumn3d => Self::Column3d,
            Kind::Line3d => Self::Line3d,
            Kind::Pie3d | Kind::Donut3d => Self::Pie3d,
            _ => return None,
        })
    }
}

/// How adjacent data points of one series are connected.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ConnectionLine {
    /// No connecting line (scatter charts only).
    Hidden,
    /// Straight segments.
    Straight,
    /// A smooth curve.
    Curved,
}

impl ConnectionLine {
    /// The stored value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Hidden => 0,
            Self::Straight => 1,
            Self::Curved => 2,
        }
    }

    /// The connection line of a stored value.
    ///
    /// # Errors
    ///
    /// [`Error::UnknownNative`].
    pub const fn from_native(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::Hidden),
            1 => Ok(Self::Straight),
            2 => Ok(Self::Curved),
            other => Err(Error::UnknownNative(other)),
        }
    }
}

/// Where a chart kind stores its connection lines (series non-style).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ConnectionLineFamily {
    /// The `Line2d` family.
    Line2d,
    /// The `Radar2d` family.
    Radar2d,
    /// The `Scatter2d` family.
    Scatter2d,
}

impl ConnectionLineFamily {
    /// The extension field that stores the family's value.
    #[must_use]
    pub const fn field(self) -> u32 {
        match self {
            Self::Line2d => 18,
            Self::Radar2d => 189,
            Self::Scatter2d => 20,
        }
    }
}

impl ConnectionLineFamily {
    /// The family an unambiguous chart kind uses, if any.
    #[must_use]
    pub const fn for_chart_kind(kind: Kind) -> Option<Self> {
        Some(match kind {
            Kind::Line2d => Self::Line2d,
            Kind::Radar2d => Self::Radar2d,
            Kind::Scatter2d => Self::Scatter2d,
            _ => return None,
        })
    }

    /// The connection line when none is stored.
    #[must_use]
    pub const fn default_value(self) -> ConnectionLine {
        match self {
            Self::Scatter2d => ConnectionLine::Hidden,
            _ => ConnectionLine::Straight,
        }
    }

    /// The series non-style field of the scatter line visibility.
    pub const SCATTER_SHOW_FIELD: u32 = 29;

    /// Read a stored connection line: the type field and, for scatter
    /// charts, the visibility field (`None` when absent).
    ///
    /// # Errors
    ///
    /// [`Error::UnknownNative`], [`Error::VisibleHiddenLine`] or
    /// [`Error::HiddenConnectionLine`].
    pub fn read(
        self,
        scatter_show: Option<bool>,
        line_type: Option<i32>,
    ) -> Result<ConnectionLine> {
        if self == Self::Scatter2d {
            if !scatter_show.unwrap_or(false) {
                if let Some(value) = line_type {
                    ConnectionLine::from_native(value)?;
                }
                return Ok(ConnectionLine::Hidden);
            }
            let line_type = line_type.unwrap_or(1);
            if line_type == 0 {
                return Err(Error::VisibleHiddenLine);
            }
            return ConnectionLine::from_native(line_type);
        }
        line_type.map_or(Ok(self.default_value()), |value| {
            let value = ConnectionLine::from_native(value)?;
            self.validate(value)?;
            Ok(value)
        })
    }

    /// Check that `value` is valid for the family.
    ///
    /// # Errors
    ///
    /// [`Error::HiddenConnectionLine`] for a hidden line outside scatter
    /// charts.
    pub fn validate(self, value: ConnectionLine) -> Result<()> {
        if value == ConnectionLine::Hidden && self != Self::Scatter2d {
            return Err(Error::HiddenConnectionLine);
        }
        Ok(())
    }
}

/// Marker shapes of the data-symbol menu.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum SymbolShape {
    /// A circle.
    #[default]
    Circle,
    /// A triangle.
    Triangle,
    /// An inverted triangle.
    InvertedTriangle,
    /// A diamond.
    Diamond,
    /// A square.
    Square,
}

impl SymbolShape {
    /// The stored value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Circle => 0,
            Self::Triangle => 2,
            Self::InvertedTriangle => 3,
            Self::Square => 4,
            Self::Diamond => 5,
        }
    }

    /// The shape of a stored value.
    ///
    /// # Errors
    ///
    /// [`Error::UnknownNative`].
    pub const fn from_native(value: i32) -> Result<Self> {
        match value {
            0 => Ok(Self::Circle),
            2 => Ok(Self::Triangle),
            3 => Ok(Self::InvertedTriangle),
            4 => Ok(Self::Square),
            5 => Ok(Self::Diamond),
            other => Err(Error::UnknownNative(other)),
        }
    }
}

/// An explicit marker size in points (absent means automatic sizing).
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct SymbolSize(f32);

impl SymbolSize {
    /// Construct a finite, positive size.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidSymbolSize`].
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() || points <= 0.0 {
            return Err(Error::InvalidSymbolSize);
        }
        Ok(Self(points))
    }

    /// The size in points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// One visible data symbol.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Symbol {
    /// The marker shape.
    pub shape: SymbolShape,
    /// The explicit size, or `None` for automatic sizing.
    pub size: Option<SymbolSize>,
}

impl Symbol {
    /// An automatically sized symbol.
    #[must_use]
    pub const fn automatic(shape: SymbolShape) -> Self {
        Self { shape, size: None }
    }

    /// An explicitly sized symbol.
    #[must_use]
    pub const fn sized(shape: SymbolShape, size: SymbolSize) -> Self {
        Self {
            shape,
            size: Some(size),
        }
    }
}

/// Where a chart kind stores its data symbols: visibility and shape (series
/// non-style) and size (series style).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SymbolFamily {
    /// Area charts.
    Area2d,
    /// Line charts.
    Line2d,
    /// The area series of mixed charts.
    MixedArea2d,
    /// The line series of mixed charts.
    MixedLine2d,
    /// Radar charts.
    Radar2d,
    /// Scatter charts.
    Scatter2d,
}

impl SymbolFamily {
    /// The family an unambiguous chart kind uses, if any.
    #[must_use]
    pub const fn for_chart_kind(kind: Kind) -> Option<Self> {
        Some(match kind {
            Kind::Area2d | Kind::StackedArea2d => Self::Area2d,
            Kind::Line2d => Self::Line2d,
            Kind::Radar2d => Self::Radar2d,
            Kind::Scatter2d => Self::Scatter2d,
            _ => return None,
        })
    }

    /// The non-style field of the symbol visibility.
    #[must_use]
    pub const fn show_field(self) -> u32 {
        match self {
            Self::Area2d => 32,
            Self::Line2d => 33,
            Self::MixedArea2d => 34,
            Self::MixedLine2d => 35,
            Self::Scatter2d => 36,
            Self::Radar2d => 160,
        }
    }

    /// The non-style field of the symbol shape.
    #[must_use]
    pub const fn type_field(self) -> u32 {
        match self {
            Self::Area2d => 47,
            Self::Line2d => 48,
            Self::MixedArea2d => 49,
            Self::MixedLine2d => 50,
            Self::Scatter2d => 51,
            Self::Radar2d => 163,
        }
    }

    /// The style field of the symbol size.
    #[must_use]
    pub const fn size_field(self) -> u32 {
        match self {
            Self::Area2d => 70,
            Self::Line2d => 71,
            Self::MixedArea2d => 72,
            Self::MixedLine2d => 73,
            Self::Scatter2d => 74,
            Self::Radar2d => 181,
        }
    }
}

/// The fill inside one data symbol.
#[derive(Debug, Clone, PartialEq)]
pub enum SymbolFill {
    /// A concrete fill.
    Custom(Fill),
    /// Follow the series fill color.
    SeriesFill,
    /// Follow the series stroke color.
    SeriesStroke,
}

impl From<Fill> for SymbolFill {
    fn from(fill: Fill) -> Self {
        Self::Custom(fill)
    }
}

/// Where a chart kind stores its symbol fills (series style): the fill and
/// the two "use series" switches (bubble charts have no series-fill switch).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SymbolFillFamily {
    /// Area charts.
    Area2d,
    /// Bubble charts.
    Bubble2d,
    /// Line charts.
    Line2d,
    /// Radar charts.
    Radar2d,
    /// Scatter charts.
    Scatter2d,
}

impl SymbolFillFamily {
    /// The family an unambiguous chart kind uses, if any.
    #[must_use]
    pub const fn for_chart_kind(kind: Kind) -> Option<Self> {
        Some(match kind {
            Kind::Area2d | Kind::StackedArea2d => Self::Area2d,
            Kind::Bubble2d => Self::Bubble2d,
            Kind::Line2d => Self::Line2d,
            Kind::Radar2d => Self::Radar2d,
            Kind::Scatter2d => Self::Scatter2d,
            _ => return None,
        })
    }

    /// The style field of the symbol fill.
    #[must_use]
    pub const fn fill_field(self) -> u32 {
        match self {
            Self::Area2d => 54,
            Self::Bubble2d => 55,
            Self::Line2d => 56,
            Self::Radar2d => 175,
            Self::Scatter2d => 59,
        }
    }

    /// The style field of the use-series-fill switch, if the family has one.
    #[must_use]
    pub const fn use_series_fill_field(self) -> Option<u32> {
        match self {
            Self::Area2d => Some(60),
            Self::Line2d => Some(61),
            Self::Radar2d => Some(177),
            Self::Bubble2d | Self::Scatter2d => None,
        }
    }

    /// The symbol fill when no style in the chain stores one.
    #[must_use]
    pub const fn default_fill(self) -> SymbolFill {
        match self {
            Self::Area2d | Self::Radar2d => SymbolFill::SeriesFill,
            Self::Bubble2d | Self::Line2d | Self::Scatter2d => SymbolFill::SeriesStroke,
        }
    }

    /// Read the symbol fill one style stores: the two switches and the fill
    /// (`None` when the style stores none of them).
    ///
    /// # Errors
    ///
    /// [`Error::BothSeriesSources`].
    pub fn read(
        self,
        use_series_fill: Option<bool>,
        use_series_stroke: Option<bool>,
        fill: Option<Fill>,
    ) -> Result<Option<SymbolFill>> {
        let use_fill = self.use_series_fill_field().and(use_series_fill);
        if use_fill == Some(true) && use_series_stroke == Some(true) {
            return Err(Error::BothSeriesSources);
        }
        if use_fill == Some(true) {
            return Ok(Some(SymbolFill::SeriesFill));
        }
        if use_series_stroke == Some(true) {
            return Ok(Some(SymbolFill::SeriesStroke));
        }
        Ok(fill.map(SymbolFill::Custom))
    }

    /// The style field of the use-series-stroke switch.
    #[must_use]
    pub const fn use_series_stroke_field(self) -> u32 {
        match self {
            Self::Area2d => 64,
            Self::Bubble2d => 65,
            Self::Line2d => 66,
            Self::Radar2d => 179,
            Self::Scatter2d => 69,
        }
    }
}

/// Where a chart kind stores its symbol outlines (series style).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SymbolOutlineFamily {
    /// The `Area2d` family.
    Area2d,
    /// The `Bubble2d` family.
    Bubble2d,
    /// The `Line2d` family.
    Line2d,
    /// The `Radar2d` family.
    Radar2d,
    /// The `Scatter2d` family.
    Scatter2d,
}

impl SymbolOutlineFamily {
    /// The extension field that stores the family's value.
    #[must_use]
    pub const fn field(self) -> u32 {
        match self {
            Self::Area2d => 75,
            Self::Bubble2d => 76,
            Self::Line2d => 77,
            Self::Radar2d => 183,
            Self::Scatter2d => 80,
        }
    }
}

impl SymbolOutlineFamily {
    /// The family an unambiguous chart kind uses, if any.
    #[must_use]
    pub const fn for_chart_kind(kind: Kind) -> Option<Self> {
        Some(match kind {
            Kind::Area2d | Kind::StackedArea2d => Self::Area2d,
            Kind::Bubble2d => Self::Bubble2d,
            Kind::Line2d => Self::Line2d,
            Kind::Radar2d => Self::Radar2d,
            Kind::Scatter2d => Self::Scatter2d,
            _ => return None,
        })
    }
}

/// The series non-style fields of a trendline.
pub mod trendline_field {
    /// `tschchartseriesshowtrendline`.
    pub const SHOW: u32 = 37;
    /// The custom label.
    pub const LABEL: u32 = 54;
    /// The polynomial order.
    pub const POLYNOMIAL_ORDER: u32 = 55;
    /// The moving-average period.
    pub const MOVING_AVERAGE_PERIOD: u32 = 56;
    /// Equation visibility.
    pub const SHOW_EQUATION: u32 = 59;
    /// Legend-label visibility.
    pub const SHOW_LABEL: u32 = 60;
    /// R² visibility.
    pub const SHOW_R_SQUARED: u32 = 61;
    /// The trendline type.
    pub const TYPE: u32 = 62;
}

const TRENDLINE_NONE: i32 = 0;
const TRENDLINE_LINEAR: i32 = 1;
const TRENDLINE_LOGARITHMIC: i32 = 2;
const TRENDLINE_POLYNOMIAL: i32 = 3;
const TRENDLINE_POWER: i32 = 4;
const TRENDLINE_EXPONENTIAL: i32 = 5;
const TRENDLINE_MOVING_AVERAGE: i32 = 6;

/// A trendline polynomial order (`2..=6`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PolynomialOrder(u8);

impl PolynomialOrder {
    /// The smallest order.
    pub const MIN: u8 = 2;
    /// The largest order.
    pub const MAX: u8 = 6;
    /// The inspector default.
    pub const DEFAULT: Self = Self(Self::MIN);

    /// Construct an order.
    ///
    /// # Errors
    ///
    /// [`Error::PolynomialOrder`].
    pub const fn new(value: u8) -> Result<Self> {
        if value < Self::MIN || value > Self::MAX {
            return Err(Error::PolynomialOrder);
        }
        Ok(Self(value))
    }

    /// The order.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for PolynomialOrder {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A trendline moving-average period (`2..=i32::MAX`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct MovingAveragePeriod(u32);

impl MovingAveragePeriod {
    /// The smallest period.
    pub const MIN: u32 = 2;
    /// The largest period (the native signed field).
    pub const MAX: u32 = i32::MAX as u32;
    /// The inspector default.
    pub const DEFAULT: Self = Self(Self::MIN);

    /// Construct a period.
    ///
    /// # Errors
    ///
    /// [`Error::MovingAveragePeriod`].
    pub const fn new(value: u32) -> Result<Self> {
        if value < Self::MIN || value > Self::MAX {
            return Err(Error::MovingAveragePeriod);
        }
        Ok(Self(value))
    }

    /// The period.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl Default for MovingAveragePeriod {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A trendline's curve family and fitting parameter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum TrendlineType {
    /// No trendline.
    #[default]
    None,
    /// A straight line.
    Linear,
    /// A logarithmic curve.
    Logarithmic,
    /// A polynomial curve.
    Polynomial(PolynomialOrder),
    /// A power curve.
    Power,
    /// An exponential curve.
    Exponential,
    /// A moving average.
    MovingAverage(MovingAveragePeriod),
    /// An unrecognized native type, kept.
    Unsupported(i32),
}

impl TrendlineType {
    /// The stored type value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::None => TRENDLINE_NONE,
            Self::Linear => TRENDLINE_LINEAR,
            Self::Logarithmic => TRENDLINE_LOGARITHMIC,
            Self::Polynomial(_) => TRENDLINE_POLYNOMIAL,
            Self::Power => TRENDLINE_POWER,
            Self::Exponential => TRENDLINE_EXPONENTIAL,
            Self::MovingAverage(_) => TRENDLINE_MOVING_AVERAGE,
            Self::Unsupported(value) => value,
        }
    }

    /// Whether a trendline is drawn.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        !matches!(self, Self::None)
    }

    /// Whether the type has equation and R² labels.
    #[must_use]
    pub const fn supports_equation(self) -> bool {
        !matches!(self, Self::None | Self::MovingAverage(_))
    }
}

/// The trendline of one series.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Trendline {
    trendline_type: TrendlineType,
    custom_name: Option<String>,
    show_name_in_legend: bool,
    show_equation: bool,
    show_r_squared: bool,
}

impl Trendline {
    /// No trendline.
    #[must_use]
    pub const fn none() -> Self {
        Self::from_type(TrendlineType::None)
    }

    /// A linear trendline.
    #[must_use]
    pub const fn linear() -> Self {
        Self::from_type(TrendlineType::Linear)
    }

    /// A logarithmic trendline.
    #[must_use]
    pub const fn logarithmic() -> Self {
        Self::from_type(TrendlineType::Logarithmic)
    }

    /// A polynomial trendline.
    #[must_use]
    pub const fn polynomial(order: PolynomialOrder) -> Self {
        Self::from_type(TrendlineType::Polynomial(order))
    }

    /// A power trendline.
    #[must_use]
    pub const fn power() -> Self {
        Self::from_type(TrendlineType::Power)
    }

    /// An exponential trendline.
    #[must_use]
    pub const fn exponential() -> Self {
        Self::from_type(TrendlineType::Exponential)
    }

    /// A moving-average trendline.
    #[must_use]
    pub const fn moving_average(period: MovingAveragePeriod) -> Self {
        Self::from_type(TrendlineType::MovingAverage(period))
    }

    /// Keep an unrecognized native type.
    ///
    /// # Errors
    ///
    /// [`Error::KnownTrendlineAsUnsupported`] for a known type.
    pub fn unsupported(native_type: i32) -> Result<Self> {
        let value = Self::from_type(TrendlineType::Unsupported(native_type));
        value.validate()?;
        Ok(value)
    }

    const fn from_type(trendline_type: TrendlineType) -> Self {
        Self {
            trendline_type,
            custom_name: None,
            show_name_in_legend: false,
            show_equation: false,
            show_r_squared: false,
        }
    }

    /// The curve family.
    #[must_use]
    pub const fn trendline_type(&self) -> TrendlineType {
        self.trendline_type
    }

    /// The custom legend name, if any.
    #[must_use]
    pub fn custom_name(&self) -> Option<&str> {
        self.custom_name.as_deref()
    }

    /// Whether the name is shown in the legend.
    #[must_use]
    pub const fn shows_name_in_legend(&self) -> bool {
        self.show_name_in_legend
    }

    /// Whether the equation is shown.
    #[must_use]
    pub const fn shows_equation(&self) -> bool {
        self.show_equation
    }

    /// Whether R² is shown.
    #[must_use]
    pub const fn shows_r_squared(&self) -> bool {
        self.show_r_squared
    }

    /// Show or hide the automatic name in the legend.
    ///
    /// # Errors
    ///
    /// The [`Self::validate`] errors.
    pub fn with_legend_visibility(mut self, visible: bool) -> Result<Self> {
        self.show_name_in_legend = visible;
        self.validate()?;
        Ok(self)
    }

    /// Use and show a custom legend name.
    ///
    /// # Errors
    ///
    /// The [`Self::validate`] errors.
    pub fn with_legend_name(mut self, name: impl Into<String>) -> Result<Self> {
        self.custom_name = Some(name.into());
        self.show_name_in_legend = true;
        self.validate()?;
        Ok(self)
    }

    /// Remove the custom name, keeping the legend visibility.
    #[must_use]
    pub fn without_custom_name(mut self) -> Self {
        self.custom_name = None;
        self
    }

    /// Show or hide the equation.
    ///
    /// # Errors
    ///
    /// The [`Self::validate`] errors.
    pub fn with_equation_visibility(mut self, visible: bool) -> Result<Self> {
        self.show_equation = visible;
        self.validate()?;
        Ok(self)
    }

    /// Show or hide R².
    ///
    /// # Errors
    ///
    /// The [`Self::validate`] errors.
    pub fn with_r_squared_visibility(mut self, visible: bool) -> Result<Self> {
        self.show_r_squared = visible;
        self.validate()?;
        Ok(self)
    }

    /// Check the native constraints between type and display options.
    ///
    /// # Errors
    ///
    /// [`Error::KnownTrendlineAsUnsupported`],
    /// [`Error::HiddenTrendlineOptions`] or [`Error::MovingAverageLabels`].
    pub fn validate(&self) -> Result<()> {
        if let TrendlineType::Unsupported(value) = self.trendline_type
            && (TRENDLINE_NONE..=TRENDLINE_MOVING_AVERAGE).contains(&value)
        {
            return Err(Error::KnownTrendlineAsUnsupported);
        }
        if !self.trendline_type.is_visible()
            && (self.custom_name.is_some()
                || self.show_name_in_legend
                || self.show_equation
                || self.show_r_squared)
        {
            return Err(Error::HiddenTrendlineOptions);
        }
        if !self.trendline_type.supports_equation() && (self.show_equation || self.show_r_squared) {
            return Err(Error::MovingAverageLabels);
        }
        Ok(())
    }

    /// Read a stored trendline: the raw varints of the show, type, order,
    /// period, show-label, show-equation and show-R² fields and the label.
    /// Equation and R² flags are ignored for types without them.
    ///
    /// # Errors
    ///
    /// [`Error::NonCanonicalBool`], [`Error::TrendlineVisibility`], an
    /// invalid order or period, or the [`Self::validate`] errors.
    #[allow(clippy::too_many_arguments, reason = "one argument per stored field")]
    pub fn from_native(
        show: Option<u64>,
        native_type: Option<u64>,
        polynomial_order: Option<u64>,
        moving_average_period: Option<u64>,
        show_label: Option<u64>,
        show_equation: Option<u64>,
        show_r_squared: Option<u64>,
        label: Option<String>,
    ) -> Result<Self> {
        let flag = |value: Option<u64>| -> Result<Option<bool>> {
            value
                .map(|value| match value {
                    0 => Ok(false),
                    1 => Ok(true),
                    _ => Err(Error::NonCanonicalBool),
                })
                .transpose()
        };
        let visible = flag(show)?.unwrap_or(false);
        let native_type = native_type.map_or(TRENDLINE_NONE, |value| value as i32);
        if !visible && native_type == TRENDLINE_NONE {
            return Ok(Self::none());
        }
        if !visible || native_type == TRENDLINE_NONE {
            return Err(Error::TrendlineVisibility);
        }
        let trendline_type = match native_type {
            TRENDLINE_LINEAR => TrendlineType::Linear,
            TRENDLINE_LOGARITHMIC => TrendlineType::Logarithmic,
            TRENDLINE_POLYNOMIAL => {
                let raw = polynomial_order.map_or(i32::from(PolynomialOrder::MIN), |v| v as i32);
                let value = u8::try_from(raw).map_err(|_| Error::PolynomialOrder)?;
                TrendlineType::Polynomial(PolynomialOrder::new(value)?)
            },
            TRENDLINE_POWER => TrendlineType::Power,
            TRENDLINE_EXPONENTIAL => TrendlineType::Exponential,
            TRENDLINE_MOVING_AVERAGE => {
                let raw =
                    moving_average_period.map_or(MovingAveragePeriod::MIN as i32, |v| v as i32);
                let value = u32::try_from(raw).map_err(|_| Error::MovingAveragePeriod)?;
                TrendlineType::MovingAverage(MovingAveragePeriod::new(value)?)
            },
            value => TrendlineType::Unsupported(value),
        };
        let supports_equation = trendline_type.supports_equation();
        let value = Self {
            trendline_type,
            custom_name: label,
            show_name_in_legend: flag(show_label)?.unwrap_or(false),
            show_equation: supports_equation && flag(show_equation)?.unwrap_or(false),
            show_r_squared: supports_equation && flag(show_r_squared)?.unwrap_or(false),
        };
        value.validate()?;
        Ok(value)
    }
}

#[cfg(test)]
mod tests {
    use super::{
        ConnectionLine, ConnectionLineFamily, Error, FillFamily, MovingAveragePeriod,
        PolynomialOrder, StrokeFamily, SymbolFamily, SymbolShape, SymbolSize, Trendline,
    };
    use crate::chart::kind::Kind;

    #[test]
    fn families_follow_chart_kinds() {
        assert_eq!(
            StrokeFamily::for_chart_kind(Kind::StackedBar2d),
            Some(StrokeFamily::BarOrColumn2d)
        );
        assert_eq!(StrokeFamily::Radar2d.field(), 172);
        assert_eq!(
            FillFamily::for_chart_kind(Kind::Donut3d),
            Some(FillFamily::Pie3d)
        );
        assert_eq!(FillFamily::for_chart_kind(Kind::Scatter2d), None);
        assert_eq!(SymbolFamily::Radar2d.size_field(), 181);
        let scatter = ConnectionLineFamily::Scatter2d;
        assert_eq!(scatter.default_value(), ConnectionLine::Hidden);
        assert_eq!(
            ConnectionLineFamily::Line2d.validate(ConnectionLine::Hidden),
            Err(Error::HiddenConnectionLine)
        );
        assert_eq!(ConnectionLine::from_native(2), Ok(ConnectionLine::Curved));
        assert_eq!(SymbolShape::from_native(5), Ok(SymbolShape::Diamond));
        assert_eq!(SymbolSize::new(0.0), Err(Error::InvalidSymbolSize));
    }

    #[test]
    fn trendlines_validate_and_read() {
        let line = Trendline::linear()
            .with_legend_name("Fit")
            .unwrap()
            .with_equation_visibility(true)
            .unwrap();
        assert_eq!(line.custom_name(), Some("Fit"));
        assert!(Trendline::none().with_legend_visibility(true).is_err());
        assert!(
            Trendline::moving_average(MovingAveragePeriod::DEFAULT)
                .with_r_squared_visibility(true)
                .is_err()
        );
        assert_eq!(
            Trendline::unsupported(3),
            Err(Error::KnownTrendlineAsUnsupported)
        );
        assert_eq!(PolynomialOrder::new(7), Err(Error::PolynomialOrder));
        let read = Trendline::from_native(
            Some(1),
            Some(3),
            Some(4),
            None,
            Some(1),
            Some(1),
            None,
            Some("p".to_owned()),
        )
        .unwrap();
        assert_eq!(read.trendline_type().native_value(), 3);
        assert!(read.shows_equation());
        assert_eq!(
            Trendline::from_native(None, None, None, None, None, None, None, None),
            Ok(Trendline::none())
        );
        assert_eq!(
            Trendline::from_native(Some(1), None, None, None, None, None, None, None),
            Err(Error::TrendlineVisibility)
        );
        assert_eq!(
            Trendline::from_native(Some(2), Some(1), None, None, None, None, None, None),
            Err(Error::NonCanonicalBool)
        );
    }
}
