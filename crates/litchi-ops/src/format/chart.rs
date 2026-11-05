//! Typed chart CRUD shared by the PPTX and XLSX adapters (record 0270).
//!
//! The chart model, writer, reader and source-preserving editor live in
//! `litchi-drawingml` ([`litchi_drawingml::chart`]); each host leaf places
//! the part (`litchi_pptx::opened::Edit::add_chart`,
//! `litchi_xlsx::chart::edit`). This module holds what both adapters share:
//!
//! - OfficeCLI's chart vocabulary: `chartType` names (`column`,
//!   `stackedBar`, `pie3d`, ...), `legend` positions, the `data` series
//!   grammar (`Name:1,2,3;Other:4,5,6`) and comma lists;
//! - [`AddProps`], the admitted values of a chart `add`;
//! - [`Summary`], the read projection of a chart part that the node views
//!   of a chart, its series and its axes print;
//! - [`ChartEdit`], the admitted values of a chart, series or axis `set`,
//!   applied through the source-preserving chart editor.
//!
//! XLSB (record 0322) reads worksheet charts through [`Summary`] only.
#![cfg_attr(
    not(any(feature = "pptx", feature = "xlsx", feature = "docx")),
    allow(
        dead_code,
        reason = "XLSB reads charts; chart authoring serves the XML hosts"
    )
)]

use litchi_drawingml::chart::authoring::{
    Appearance, Bins, Grouping, Kind, Labels, Line, ParentLabels, QuartileMethod, Series, Spec,
    Variant,
};
use litchi_drawingml::chart::data::TitleText;
use litchi_drawingml::chart::plot_area::TypeGroup;
use litchi_drawingml::chart::types::LegendPosition;
use litchi_drawingml::chart::types::{DataLabelPosition, MarkerStyle, RadarStyle, ScatterStyle};
use litchi_drawingml::chart::{Axis, Chart};

use crate::error::{Code, Error, Result};
use crate::node::spec::PropSpec;
use crate::value::{Scalar, ValueType};

/// The longest `data`, `categories` or `values` text the adapters admit.
pub(crate) const MAX_DATA_BYTES: usize = 4 * 1024 * 1024;

fn invalid(message: &'static str, property: &'static str) -> Error {
    Error::new(Code::InvalidValue, message).with_property(property)
}

/// A series value given as a cell range: the key is supported, the
/// reference form is not authored yet.
fn range_refused(property: &'static str) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "a series bound to a cell range is not authored yet",
    )
    .with_property(property)
    .with_suggestion("give the values inline (1,2,3), or bind a worksheet range with dataRange")
}

/// The OfficeCLI `chartType` name of a variant: the family, prefixed by
/// `stacked`/`percentStacked` and suffixed by `3d` (`stackedColumn3d`).
pub(crate) fn variant_name(variant: Variant) -> String {
    let family = match variant.kind {
        Kind::Column => "column",
        Kind::Bar => "bar",
        Kind::Line => "line",
        Kind::Pie => "pie",
        Kind::Doughnut => "doughnut",
        Kind::Area => "area",
        Kind::Scatter => "scatter",
        Kind::Bubble => "bubble",
        Kind::Radar => "radar",
        Kind::Stock => "stock",
        Kind::Combo => "combo",
        Kind::Waterfall => "waterfall",
        Kind::Funnel => "funnel",
        Kind::Treemap => "treemap",
        Kind::Sunburst => "sunburst",
        Kind::Histogram => "histogram",
        Kind::Pareto => "pareto",
        Kind::BoxWhisker => "boxWhisker",
    };
    let mut name = match variant.grouping {
        Grouping::Standard => family.to_owned(),
        Grouping::Stacked => format!("stacked{}{}", family[..1].to_uppercase(), &family[1..]),
        Grouping::PercentStacked => {
            format!(
                "percentStacked{}{}",
                family[..1].to_uppercase(),
                &family[1..]
            )
        },
    };
    if variant.three_d {
        name.push_str("3d");
    }
    name
}

/// Parse an OfficeCLI `chartType` name: a family (`column`, `col`, `bar`,
/// `line`, `pie`, `doughnut`/`donut`, `area`, `scatter`/`xy`, `bubble`,
/// `radar`/`spider`, `stock`/`ohlc`, `combo`, or an extended family:
/// `waterfall`/`wf`, `funnel`, `treemap`, `sunburst`, `histogram`,
/// `pareto`, `boxWhisker`) with an optional `stacked` or
/// `percentStacked` modifier (either side) and an optional `3d` suffix,
/// case-insensitively.
pub(crate) fn parse_variant(text: &str) -> Result<Variant> {
    let mut key: String = text
        .trim()
        .chars()
        .filter(|character| !matches!(character, '_' | '-' | ' '))
        .flat_map(char::to_lowercase)
        .collect();
    let three_d = key.contains("3d");
    key = key.replace("3d", "");
    let grouping = if key.contains("percentstacked") || key.contains("pstacked") {
        Grouping::PercentStacked
    } else if key.contains("stacked") {
        Grouping::Stacked
    } else {
        Grouping::Standard
    };
    let key = key
        .replace("percentstacked", "")
        .replace("pstacked", "")
        .replace("stacked", "");
    let kind = match key.as_str() {
        "column" | "col" => Kind::Column,
        "bar" => Kind::Bar,
        "line" => Kind::Line,
        "pie" => Kind::Pie,
        "doughnut" | "donut" => Kind::Doughnut,
        "area" => Kind::Area,
        "scatter" | "xy" => Kind::Scatter,
        "bubble" => Kind::Bubble,
        "radar" | "spider" => Kind::Radar,
        "stock" | "ohlc" => Kind::Stock,
        "combo" => Kind::Combo,
        "waterfall" | "wf" => Kind::Waterfall,
        "funnel" => Kind::Funnel,
        "treemap" => Kind::Treemap,
        "sunburst" => Kind::Sunburst,
        "histogram" => Kind::Histogram,
        "pareto" => Kind::Pareto,
        "boxwhisker" | "box" => Kind::BoxWhisker,
        "pieofpie" | "barofpie" => {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "this chart type is not authored yet",
            )
            .with_property("chartType")
            .with_suggestion(
                "use column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo or an extended type",
            ));
        },
        _ => {
            return Err(invalid("unknown chart type", "chartType").with_suggestion(
                "use column, bar, line, pie, doughnut, area, scatter, bubble, radar or stock, \
                 with an optional stacked/percentStacked prefix and 3d suffix",
            ));
        },
    };
    let variant = Variant::new(kind)
        .with_grouping(grouping)
        .with_three_d(three_d);
    if !variant.is_valid() {
        return Err(invalid(
            "the chart family has no such stacked or 3D form",
            "chartType",
        ));
    }
    Ok(variant)
}

/// The OfficeCLI `legend` value of a legend position (`none` when hidden).
pub(crate) const fn legend_name(position: Option<LegendPosition>) -> &'static str {
    match position {
        None => "none",
        Some(LegendPosition::Bottom) => "bottom",
        Some(LegendPosition::Top) => "top",
        Some(LegendPosition::Left) => "left",
        Some(LegendPosition::Right) => "right",
        Some(LegendPosition::TopRight) => "topRight",
    }
}

/// Parse an OfficeCLI `legend` value: a position (`top`/`t`, `bottom`/`b`,
/// `left`/`l`, `right`/`r`, `topRight`/`tr`), `true` (bottom), or
/// `false`/`none`/`hidden` (no legend).
pub(crate) fn parse_legend(text: &str) -> Result<Option<LegendPosition>> {
    Ok(Some(match text.trim().to_ascii_lowercase().as_str() {
        "true" | "bottom" | "b" | "show" => LegendPosition::Bottom,
        "top" | "t" => LegendPosition::Top,
        "left" | "l" => LegendPosition::Left,
        "right" | "r" => LegendPosition::Right,
        "topright" | "tr" => LegendPosition::TopRight,
        "false" | "none" | "hidden" | "hide" | "off" => return Ok(None),
        _ => {
            return Err(invalid("unknown legend position", "legend")
                .with_suggestion("use top, bottom, left, right, topRight, true or none"));
        },
    }))
}

fn bounded<'t>(text: &'t str, property: &'static str) -> Result<&'t str> {
    if text.len() > MAX_DATA_BYTES {
        return Err(
            Error::new(Code::InvalidValue, "the chart data text is too long")
                .with_property(property),
        );
    }
    if text.contains('\0') {
        return Err(invalid("chart text cannot contain NUL", property));
    }
    Ok(text)
}

/// Split a comma list (`categories`), trimming each item.
pub(crate) fn parse_list(text: &str, property: &'static str) -> Result<Vec<String>> {
    let text = bounded(text, property)?;
    if text.trim().is_empty() {
        return Err(invalid("the list is empty", property));
    }
    Ok(text.split(',').map(|item| item.trim().to_owned()).collect())
}

/// Parse a comma list of finite numbers (`values`).
pub(crate) fn parse_values(text: &str, property: &'static str) -> Result<Vec<f64>> {
    let text = bounded(text, property)?;
    text.split(',')
        .map(|item| {
            crate::value::parse_number(item.trim())
                .ok()
                .filter(|value| value.is_finite())
                .ok_or_else(|| invalid("a chart value is not a finite number", property))
        })
        .collect()
}

/// Parse OfficeCLI's `data` grammar: series separated by `;`, each
/// `Name:v1,v2,...` (the last `:` splits the name). Without `;`, a list
/// whose every comma item holds a `:` is one single-value series per item
/// (`Q1:40,Q2:55`).
pub(crate) fn parse_data(text: &str) -> Result<Vec<Series>> {
    let text = bounded(text, "data")?;
    let parts: Vec<&str> = if text.contains(';') {
        text.split(';')
            .filter(|part| !part.trim().is_empty())
            .collect()
    } else {
        let items: Vec<&str> = text.split(',').filter(|part| !part.is_empty()).collect();
        if items.len() > 1 && items.iter().all(|item| item.contains(':')) {
            items
        } else {
            vec![text]
        }
    };
    let mut series = Vec::with_capacity(parts.len());
    for (index, part) in parts.into_iter().enumerate() {
        let (name, values) = match part.rsplit_once(':') {
            Some((name, values)) => (name.trim().to_owned(), values),
            None => (format!("Series {}", index + 1), part),
        };
        if values.trim().is_empty() {
            return Err(invalid("a series has no values", "data")
                .with_suggestion("write data as Name:1,2,3;Other:4,5,6"));
        }
        series.push(Series::new(name, parse_values(values, "data")?));
    }
    if series.is_empty() {
        return Err(invalid("data names no series", "data")
            .with_suggestion("write data as Name:1,2,3;Other:4,5,6"));
    }
    Ok(series)
}

/// The chart formatting keys: written by `add`, rewritten by a `set` that
/// rebuilds an authored chart (record 0283), and read back from the chart
/// part where the spec carries them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FormatKey {
    Colors,
    DataLabels,
    LabelPos,
    VaryColors,
    GapWidth,
    Overlap,
    HoleSize,
    FirstSliceAngle,
    ScatterStyle,
    RadarStyle,
    Smooth,
    Explosion,
    LegendOverlay,
    AutoTitleDeleted,
    AxisMin,
    AxisMax,
    MajorUnit,
    Gridlines,
    MinorGridlines,
    ComboTypes,
    ComboSplit,
    Marker,
    MarkerSize,
    ShowMarkers,
    LineWidth,
    View3d,
    IncreaseColor,
    DecreaseColor,
    TotalColor,
    BinCount,
    BinSize,
    IntervalClosed,
    OverflowBin,
    UnderflowBin,
    QuartileMethod,
    ParentLabelLayout,
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

const fn read_write(name: &'static str, ty: ValueType) -> PropSpec {
    PropSpec::new(name, ty)
}

/// The formatting keys with their value types, OfficeCLI's spellings.
pub(crate) const FORMAT_PROPS: [(FormatKey, PropSpec); 36] = [
    (FormatKey::Colors, read_write("colors", ValueType::Text)),
    (
        FormatKey::DataLabels,
        read_write("dataLabels", ValueType::Text),
    ),
    (FormatKey::LabelPos, read_write("labelPos", ValueType::Text)),
    (
        FormatKey::VaryColors,
        read_write("varyColors", ValueType::Bool),
    ),
    (FormatKey::GapWidth, read_write("gapWidth", ValueType::Int)),
    (FormatKey::Overlap, read_write("overlap", ValueType::Int)),
    (FormatKey::HoleSize, read_write("holeSize", ValueType::Int)),
    (
        FormatKey::FirstSliceAngle,
        read_write("firstSliceAngle", ValueType::Int),
    ),
    (
        FormatKey::ScatterStyle,
        read_write("scatterStyle", ValueType::Text),
    ),
    (
        FormatKey::RadarStyle,
        read_write("radarStyle", ValueType::Text),
    ),
    (FormatKey::Smooth, read_write("smooth", ValueType::Bool)),
    (
        FormatKey::Explosion,
        read_write("explosion", ValueType::Int),
    ),
    (
        FormatKey::LegendOverlay,
        read_write("legend.overlay", ValueType::Bool),
    ),
    (
        FormatKey::AutoTitleDeleted,
        read_write("autoTitleDeleted", ValueType::Bool),
    ),
    (FormatKey::AxisMin, read_write("axisMin", ValueType::Number)),
    (FormatKey::AxisMax, read_write("axisMax", ValueType::Number)),
    (
        FormatKey::MajorUnit,
        read_write("majorUnit", ValueType::Number),
    ),
    (
        FormatKey::Gridlines,
        read_write("gridlines", ValueType::Text),
    ),
    (
        FormatKey::MinorGridlines,
        read_write("minorGridlines", ValueType::Text),
    ),
    (
        FormatKey::ComboTypes,
        read_write("comboTypes", ValueType::Text),
    ),
    (
        FormatKey::ComboSplit,
        write_only("comboSplit", ValueType::Int),
    ),
    (FormatKey::Marker, read_write("marker", ValueType::Text)),
    (
        FormatKey::MarkerSize,
        read_write("markerSize", ValueType::Int),
    ),
    (
        FormatKey::ShowMarkers,
        read_write("showMarkers", ValueType::Bool),
    ),
    (
        FormatKey::LineWidth,
        read_write("lineWidth", ValueType::Number),
    ),
    (FormatKey::View3d, read_write("view3d", ValueType::Text)),
    (
        FormatKey::IncreaseColor,
        read_write("increaseColor", ValueType::Text),
    ),
    (
        FormatKey::DecreaseColor,
        read_write("decreaseColor", ValueType::Text),
    ),
    (
        FormatKey::TotalColor,
        read_write("totalColor", ValueType::Text),
    ),
    (FormatKey::BinCount, read_write("binCount", ValueType::Int)),
    (FormatKey::BinSize, read_write("binSize", ValueType::Number)),
    (
        FormatKey::IntervalClosed,
        read_write("intervalClosed", ValueType::Text),
    ),
    (
        FormatKey::OverflowBin,
        read_write("overflowBin", ValueType::Number),
    ),
    (
        FormatKey::UnderflowBin,
        read_write("underflowBin", ValueType::Number),
    ),
    (
        FormatKey::QuartileMethod,
        read_write("quartileMethod", ValueType::Text),
    ),
    (
        FormatKey::ParentLabelLayout,
        read_write("parentLabelLayout", ValueType::Text),
    ),
];

/// A set of property-table indices, for duplicate-assignment checks on
/// tables of up to 256 entries without allocating.
#[derive(Default)]
pub(crate) struct Assigned([u64; 4]);

impl Assigned {
    /// Record `index`; `false` when it was already recorded.
    pub(crate) fn insert(&mut self, index: usize) -> bool {
        let (word, bit) = (index / 64, 1_u64 << (index % 64));
        match self.0.get_mut(word) {
            Some(slot) if *slot & bit == 0 => {
                *slot |= bit;
                true
            },
            Some(_) => false,
            None => true,
        }
    }
}

/// The largest chart property table [`Assigned`] tracks.
pub(crate) const MAX_TABLE: usize = 256;

/// The legacy numbered series keys (`series1=Name:1,2,3`).
pub(crate) const NUMBERED_SERIES: [&str; 20] = [
    "series1", "series2", "series3", "series4", "series5", "series6", "series7", "series8",
    "series9", "series10", "series11", "series12", "series13", "series14", "series15", "series16",
    "series17", "series18", "series19", "series20",
];

/// The numbered series keys as write-only add properties.
pub(crate) const NUMBERED_SERIES_PROPS: [PropSpec; 20] = {
    let mut props = [PropSpec::new("series1", ValueType::Text); 20];
    let mut index = 0;
    while index < 20 {
        props[index] = write_only(NUMBERED_SERIES[index], ValueType::Text);
        index += 1;
    }
    props
};

/// The part of one numbered series a dotted key (`series{N}.name`)
/// assigns, OfficeCLI's per-series add keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SeriesField {
    /// `series{N}.name`: the series name.
    Name,
    /// `series{N}.values`: the comma list of values.
    Values,
    /// `series{N}.color`: the series fill, `RRGGBB` or `AARRGGBB`.
    Color,
}

/// The fields in the order [`DOTTED_SERIES`] lists them for each number.
pub(crate) const SERIES_FIELDS: [SeriesField; 3] =
    [SeriesField::Name, SeriesField::Values, SeriesField::Color];

macro_rules! dotted_series {
    ($($number:literal),*) => {
        [$(
            concat!("series", $number, ".name"),
            concat!("series", $number, ".values"),
            concat!("series", $number, ".color"),
        )*]
    };
}

/// The dotted numbered series keys, `series1.name`, `series1.values`,
/// `series1.color`, `series2.name`, ... `series20.color`: entry
/// `(N - 1) * 3 + field`, fields in [`SERIES_FIELDS`] order.
pub(crate) const DOTTED_SERIES: [&str; 60] = dotted_series!(
    1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
);

/// The dotted numbered series keys as write-only properties.
pub(crate) const DOTTED_SERIES_PROPS: [PropSpec; 60] = {
    let mut props = [PropSpec::new("series1.name", ValueType::Text); 60];
    let mut index = 0;
    while index < 60 {
        props[index] = write_only(DOTTED_SERIES[index], ValueType::Text);
        index += 1;
    }
    props
};

/// The dotted key of a series number (one-based) and field.
fn dotted_key(number: usize, field: SeriesField) -> &'static str {
    let offset = match field {
        SeriesField::Name => 0,
        SeriesField::Values => 1,
        SeriesField::Color => 2,
    };
    DOTTED_SERIES[(number - 1) * 3 + offset]
}

/// Whether a series value is a cell range (`Sheet1!B2:B13`, `B2:B13`),
/// OfficeCLI's reference form, rather than a comma list of numbers.
fn is_range_reference(text: &str) -> bool {
    let text = text.trim();
    if text.contains('!') {
        return true;
    }
    let mut cells = text.split(':');
    let cell = |part: Option<&str>| {
        part.is_some_and(|part| {
            let part = part.trim().replace('$', "");
            let letters = part.bytes().take_while(u8::is_ascii_alphabetic).count();
            letters > 0
                && part.len() > letters
                && part.bytes().skip(letters).all(|byte| byte.is_ascii_digit())
        })
    };
    cell(cells.next()) && cell(cells.next()) && cells.next().is_none()
}

/// What the keys of one series number assigned on a chart `add`.
#[derive(Debug, Default)]
pub(crate) struct Numbered {
    /// The one-based series number.
    number: usize,
    /// `series{N}=[Name:]1,2,3`: the name, when the value has one, and
    /// the values.
    legacy: Option<(Option<String>, Vec<f64>)>,
    name: Option<String>,
    values: Option<Vec<f64>>,
    color: Option<String>,
}

impl Numbered {
    /// Whether the keys define the series itself (not only its color).
    fn defines(&self) -> bool {
        self.legacy.is_some() || self.name.is_some() || self.values.is_some()
    }

    /// The series the keys define, OfficeCLI's merge: the dotted name, else
    /// the legacy one, else `Series N`; the dotted values, else the legacy
    /// ones. A key that assigns what another already did is refused.
    fn into_series(self) -> Result<Series> {
        let number = self.number;
        let (legacy_name, legacy_values) = match self.legacy {
            Some((name, values)) => (name, Some(values)),
            None => (None, None),
        };
        if self.name.is_some() && legacy_name.is_some() {
            return Err(Error::new(
                Code::InvalidArgument,
                "the series is named by both series{N}=Name:... and series{N}.name",
            )
            .with_property(dotted_key(number, SeriesField::Name))
            .with_suggestion("name it once: drop the name from series{N}= or the .name key"));
        }
        if self.values.is_some() && legacy_values.is_some() {
            return Err(Error::new(
                Code::InvalidArgument,
                "the series values are given by both series{N}= and series{N}.values",
            )
            .with_property(dotted_key(number, SeriesField::Values))
            .with_suggestion("give the values once"));
        }
        let values = self.values.or(legacy_values).ok_or_else(|| {
            Error::new(Code::MissingProperty, "a numbered series has no values")
                .with_property(dotted_key(number, SeriesField::Values))
                .with_suggestion("add series{N}.values=1,2,3")
        })?;
        let name = self
            .name
            .or(legacy_name)
            .unwrap_or_else(|| format!("Series {number}"));
        let series = Series::new(name, values);
        Ok(match self.color {
            Some(color) => series.with_color(color),
            None => series,
        })
    }
}

fn label_position(text: &str) -> Option<DataLabelPosition> {
    Some(match text {
        "outsideend" | "outend" | "outside" => DataLabelPosition::OutsideEnd,
        "insideend" | "inend" | "inside" => DataLabelPosition::InsideEnd,
        "insidebase" | "inbase" => DataLabelPosition::InsideBase,
        "center" | "ctr" | "centre" => DataLabelPosition::Center,
        "bestfit" | "best" => DataLabelPosition::BestFit,
        "top" | "t" | "above" => DataLabelPosition::Top,
        "bottom" | "b" | "below" => DataLabelPosition::Bottom,
        "left" | "l" => DataLabelPosition::Left,
        "right" | "r" => DataLabelPosition::Right,
        _ => return None,
    })
}

fn marker(text: &str) -> Result<MarkerStyle> {
    Ok(match text.trim().to_ascii_lowercase().as_str() {
        "circle" => MarkerStyle::Circle,
        "dash" => MarkerStyle::Dash,
        "diamond" => MarkerStyle::Diamond,
        "dot" => MarkerStyle::Dot,
        "none" => MarkerStyle::None,
        "plus" => MarkerStyle::Plus,
        "square" => MarkerStyle::Square,
        "star" => MarkerStyle::Star,
        "triangle" => MarkerStyle::Triangle,
        "x" => MarkerStyle::X,
        "auto" => MarkerStyle::Auto,
        _ => return Err(invalid("unknown marker symbol", "marker")),
    })
}

fn rgb(text: &str, property: &'static str) -> Result<String> {
    let color = text.trim().trim_start_matches('#');
    if color.len() == 6 && color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Ok(color.to_ascii_uppercase())
    } else {
        Err(invalid("a chart color is not RRGGBB", property))
    }
}

/// Parse a series color: `RRGGBB` or `AARRGGBB`, `#` allowed.
#[cfg_attr(
    not(any(feature = "pptx", feature = "xlsx")),
    allow(dead_code, reason = "DOCX charts have no series or axis nodes yet")
)]
pub(crate) fn parse_color(text: &str, property: &'static str) -> Result<String> {
    argb(text, property)
}

/// `RRGGBB`, or `AARRGGBB` with an alpha byte (OfficeCLI's ARGB colors).
fn argb(text: &str, property: &'static str) -> Result<String> {
    let color = text.trim().trim_start_matches('#');
    if matches!(color.len(), 6 | 8) && color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Ok(color.to_ascii_uppercase())
    } else {
        Err(invalid("a chart color is not RRGGBB or AARRGGBB", property))
    }
}

/// `RRGGBB[:width[:dash]]` gridlines, or `true`/`false`/`none`.
fn gridline(text: &str, property: &'static str) -> Result<Option<Option<Line>>> {
    let text = text.trim();
    match text.to_ascii_lowercase().as_str() {
        "true" | "show" => return Ok(Some(None)),
        "false" | "none" | "hide" => return Ok(None),
        _ => {},
    }
    let mut parts = text.split(':');
    let color = rgb(parts.next().unwrap_or_default(), property)?;
    let width = match parts.next() {
        Some(width) => Some(
            crate::value::parse_number(width.trim())
                .ok()
                .filter(|width| (0.0..=1584.0).contains(width))
                .ok_or_else(|| invalid("a gridline width is not a number of points", property))?,
        ),
        None => None,
    };
    let dash = match parts.next() {
        Some(dash) => Some(
            match dash.trim().to_ascii_lowercase().as_str() {
                "solid" => "solid",
                "dot" => "sysDot",
                "dash" => "dash",
                "longdash" | "lgdash" => "lgDash",
                "dashdot" => "dashDot",
                "sysdash" => "sysDash",
                "sysdot" => "sysDot",
                _ => return Err(invalid("unknown gridline dash", property)),
            }
            .to_owned(),
        ),
        None => None,
    };
    if parts.next().is_some() {
        return Err(invalid("a gridline is color:width:dash", property));
    }
    Ok(Some(Some(Line { color, width, dash })))
}

/// The admitted values of a chart `add` the two hosts share; geometry and
/// placement stay with the host adapter.
#[derive(Debug, Default)]
pub(crate) struct AddProps {
    pub(crate) variant: Option<Variant>,
    pub(crate) title: Option<String>,
    pub(crate) legend: Option<Option<LegendPosition>>,
    pub(crate) categories: Option<Vec<String>>,
    pub(crate) data: Option<Vec<Series>>,
    pub(crate) category_axis_title: Option<String>,
    pub(crate) value_axis_title: Option<String>,
    pub(crate) appearance: Appearance,
    /// The `series{N}` and `series{N}.*` keys, one entry per number in
    /// first-assignment order.
    pub(crate) numbered: Vec<Numbered>,
    /// A combo chart's series families (`comboTypes`).
    pub(crate) combo_types: Option<Vec<Kind>>,
    /// How many leading series of a combo chart are columns (`comboSplit`).
    pub(crate) combo_split: Option<usize>,
}

impl AddProps {
    /// The entry of a series number, created on its first key.
    fn numbered_entry(&mut self, number: usize) -> &mut Numbered {
        let position = match self
            .numbered
            .iter()
            .position(|entry| entry.number == number)
        {
            Some(position) => position,
            None => {
                self.numbered.push(Numbered {
                    number,
                    ..Numbered::default()
                });
                self.numbered.len() - 1
            },
        };
        &mut self.numbered[position]
    }

    /// Admit a `series{N}` value: `Name:1,2,3` (the last `:` splits the
    /// name), or bare values named `Series N` (OfficeCLI).
    pub(crate) fn numbered_series(&mut self, number: usize, text: &str) -> Result<()> {
        let property = NUMBERED_SERIES[number - 1];
        if is_range_reference(text) {
            return Err(range_refused(property));
        }
        let (name, values) = match text.rsplit_once(':') {
            Some((name, values)) => (Some(name.trim().to_owned()), values),
            None => (None, text),
        };
        let values = parse_values(values, property)?;
        self.numbered_entry(number).legacy = Some((name, values));
        Ok(())
    }

    /// Admit a dotted `series{N}.name`, `.values` or `.color` value.
    pub(crate) fn dotted_series(
        &mut self,
        number: usize,
        field: SeriesField,
        text: &str,
    ) -> Result<()> {
        let property = dotted_key(number, field);
        match field {
            SeriesField::Name => {
                let name = bounded(text, property)?.to_owned();
                self.numbered_entry(number).name = Some(name);
            },
            SeriesField::Values => {
                if is_range_reference(text) {
                    return Err(range_refused(property));
                }
                let values = parse_values(text, property)?;
                self.numbered_entry(number).values = Some(values);
            },
            SeriesField::Color => {
                let color = argb(text, property)?;
                self.numbered_entry(number).color = Some(color);
            },
        }
        Ok(())
    }

    /// Admit one formatting key.
    pub(crate) fn format(&mut self, key: FormatKey, value: Option<Scalar<'_>>) -> Result<()> {
        let text = || match value {
            Some(Scalar::Text(text) | Scalar::Token(text)) => text.to_owned(),
            _ => String::new(),
        };
        let int = || match value {
            Some(Scalar::Int(value)) => value,
            _ => 0,
        };
        let number = || match value {
            Some(Scalar::Number(value)) => value,
            Some(Scalar::Int(value)) => value as f64,
            _ => 0.0,
        };
        let flag = || matches!(value, Some(Scalar::Bool(true)));
        let unsigned = |property: &'static str| {
            u32::try_from(int()).map_err(|_| invalid("the value cannot be negative", property))
        };
        let appearance = &mut self.appearance;
        match key {
            FormatKey::Colors => {
                appearance.colors = text()
                    .split(',')
                    .map(|color| argb(color, "colors"))
                    .collect::<Result<_>>()?;
            },
            FormatKey::DataLabels => {
                let text = text();
                let mut labels = Labels {
                    position: appearance.labels.and_then(|labels| labels.position),
                    ..Labels::default()
                };
                for item in text.split(',') {
                    match item.trim().to_ascii_lowercase().as_str() {
                        "true" | "show" | "value" | "values" | "val" => labels.value = true,
                        "category" | "cat" | "categoryname" => labels.category = true,
                        "series" | "seriesname" | "ser" => labels.series = true,
                        "percent" | "percentage" | "pct" => labels.percent = true,
                        "false" | "none" | "hide" => {},
                        // A position word shows values there (OfficeCLI).
                        other if label_position(other).is_some() => {
                            labels.value = true;
                            labels.position = label_position(other);
                        },
                        _ => {
                            return Err(invalid("unknown data-label content", "dataLabels")
                                .with_suggestion("use true, none, or a list of value, category, series and percent"));
                        },
                    }
                }
                appearance.labels = Some(labels);
            },
            FormatKey::LabelPos => {
                let position = label_position(&text().trim().to_ascii_lowercase())
                    .ok_or_else(|| invalid("unknown data-label position", "labelPos"))?;
                let mut labels = appearance.labels.unwrap_or(Labels {
                    value: true,
                    ..Labels::default()
                });
                labels.position = Some(position);
                appearance.labels = Some(labels);
            },
            FormatKey::VaryColors => appearance.vary_colors = Some(flag()),
            FormatKey::GapWidth => appearance.gap_width = Some(unsigned("gapWidth")?),
            FormatKey::Overlap => {
                appearance.overlap = Some(
                    i32::try_from(int())
                        .map_err(|_| invalid("overlap is -100 to 100", "overlap"))?,
                );
            },
            FormatKey::HoleSize => appearance.hole_size = Some(unsigned("holeSize")?),
            FormatKey::FirstSliceAngle => {
                appearance.first_slice_angle = Some(unsigned("firstSliceAngle")?);
            },
            FormatKey::ScatterStyle => {
                appearance.scatter_style =
                    Some(match text().trim().to_ascii_lowercase().as_str() {
                        "line" => ScatterStyle::Line,
                        "linemarker" => ScatterStyle::LineMarker,
                        "marker" => ScatterStyle::Marker,
                        "smooth" => ScatterStyle::Smooth,
                        "smoothmarker" => ScatterStyle::SmoothMarker,
                        "none" => ScatterStyle::None,
                        _ => return Err(invalid("unknown scatter style", "scatterStyle")),
                    });
            },
            FormatKey::RadarStyle => {
                appearance.radar_style = Some(match text().trim().to_ascii_lowercase().as_str() {
                    "standard" => RadarStyle::Standard,
                    "marker" => RadarStyle::Marker,
                    "filled" => RadarStyle::Filled,
                    _ => return Err(invalid("unknown radar style", "radarStyle")),
                });
            },
            FormatKey::Smooth => appearance.smooth = Some(flag()),
            FormatKey::Explosion => appearance.explosion = Some(unsigned("explosion")?),
            FormatKey::LegendOverlay => appearance.legend_overlay = Some(flag()),
            FormatKey::AutoTitleDeleted => appearance.auto_title_deleted = Some(flag()),
            FormatKey::AxisMin => appearance.axis_min = Some(number()),
            FormatKey::AxisMax => appearance.axis_max = Some(number()),
            FormatKey::MajorUnit => appearance.major_unit = Some(number()),
            FormatKey::Gridlines => match gridline(&text(), "gridlines")? {
                Some(line) => {
                    appearance.hide_gridlines = false;
                    appearance.gridlines = Some(line);
                },
                None => appearance.hide_gridlines = true,
            },
            FormatKey::MinorGridlines => {
                appearance.minor_gridlines = gridline(&text(), "minorGridlines")?;
            },
            FormatKey::ComboTypes => {
                let kinds = text()
                    .split(',')
                    .map(|kind| match kind.trim().to_ascii_lowercase().as_str() {
                        "column" | "col" | "bar" => Ok(Kind::Column),
                        "line" => Ok(Kind::Line),
                        "area" => Ok(Kind::Area),
                        _ => Err(invalid(
                            "a combo chart mixes column, line and area series",
                            "comboTypes",
                        )),
                    })
                    .collect::<Result<Vec<_>>>()?;
                self.combo_types = Some(kinds);
            },
            FormatKey::ComboSplit => {
                self.combo_split = Some(
                    usize::try_from(int())
                        .map_err(|_| invalid("comboSplit cannot be negative", "comboSplit"))?,
                );
            },
            FormatKey::Marker => {
                let text = text();
                let mut parts = text.split(':');
                appearance.marker = Some(marker(parts.next().unwrap_or_default())?);
                if let Some(size) = parts.next() {
                    appearance.marker_size =
                        Some(size.trim().parse::<u32>().map_err(|_| {
                            invalid("a marker size is a number of points", "marker")
                        })?);
                }
                if let Some(color) = parts.next() {
                    appearance.marker_color = Some(rgb(color, "marker")?);
                }
                if parts.next().is_some() {
                    return Err(invalid("a marker is symbol:size:color", "marker"));
                }
            },
            FormatKey::MarkerSize => appearance.marker_size = Some(unsigned("markerSize")?),
            FormatKey::ShowMarkers => {
                if !flag() {
                    appearance.marker = Some(MarkerStyle::None);
                }
            },
            FormatKey::LineWidth => appearance.line_width = Some(number()),
            FormatKey::View3d => {
                let text = text();
                let parts = text
                    .split(',')
                    .map(|part| {
                        part.trim()
                            .parse::<u32>()
                            .map_err(|_| invalid("view3d is rotX,rotY,perspective", "view3d"))
                    })
                    .collect::<Result<Vec<_>>>()?;
                appearance.view_3d = Some(match parts.as_slice() {
                    [x] => (*x, 0, None),
                    [x, y] => (*x, *y, None),
                    [x, y, perspective] => (*x, *y, Some(*perspective)),
                    _ => return Err(invalid("view3d is rotX,rotY,perspective", "view3d")),
                });
            },
            FormatKey::IncreaseColor => {
                appearance.increase_color = Some(argb(&text(), "increaseColor")?);
            },
            FormatKey::DecreaseColor => {
                appearance.decrease_color = Some(argb(&text(), "decreaseColor")?);
            },
            FormatKey::TotalColor => appearance.total_color = Some(argb(&text(), "totalColor")?),
            FormatKey::BinCount => {
                let bins = appearance.bins.get_or_insert_with(Bins::default);
                bins.count = Some(
                    u32::try_from(int())
                        .ok()
                        .filter(|count| *count > 0)
                        .ok_or_else(|| invalid("binCount is a positive count", "binCount"))?,
                );
                bins.size = None;
            },
            FormatKey::BinSize => {
                let bins = appearance.bins.get_or_insert_with(Bins::default);
                bins.size = Some(number()).filter(|size| *size > 0.0);
                if bins.size.is_none() {
                    return Err(invalid("binSize is a positive width", "binSize"));
                }
                bins.count = None;
            },
            FormatKey::IntervalClosed => {
                let bins = appearance.bins.get_or_insert_with(Bins::default);
                bins.closed_left = match text().trim().to_ascii_lowercase().as_str() {
                    "l" | "left" => true,
                    "r" | "right" => false,
                    _ => return Err(invalid("intervalClosed is l or r", "intervalClosed")),
                };
            },
            FormatKey::OverflowBin => {
                appearance.bins.get_or_insert_with(Bins::default).overflow = Some(number());
            },
            FormatKey::UnderflowBin => {
                appearance.bins.get_or_insert_with(Bins::default).underflow = Some(number());
            },
            FormatKey::QuartileMethod => {
                appearance.quartile_method =
                    Some(match text().trim().to_ascii_lowercase().as_str() {
                        "exclusive" => QuartileMethod::Exclusive,
                        "inclusive" => QuartileMethod::Inclusive,
                        _ => {
                            return Err(invalid(
                                "quartileMethod is exclusive or inclusive",
                                "quartileMethod",
                            ));
                        },
                    });
            },
            FormatKey::ParentLabelLayout => {
                appearance.parent_labels =
                    Some(match text().trim().to_ascii_lowercase().as_str() {
                        "overlapping" | "overlap" => ParentLabels::Overlapping,
                        "banner" => ParentLabels::Banner,
                        "none" => ParentLabels::None,
                        _ => {
                            return Err(invalid(
                                "parentLabelLayout is overlapping, banner or none",
                                "parentLabelLayout",
                            ));
                        },
                    });
            },
        }
        Ok(())
    }

    /// Build the chart spec; the chart type defaults to `column` (OfficeCLI)
    /// and the legend to `bottom`.
    pub(crate) fn into_spec(mut self) -> Result<Spec> {
        self.numbered.sort_by_key(|entry| entry.number);
        let defined = self.numbered.iter().any(Numbered::defines);
        if defined {
            if self.data.is_some() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "a chart takes data or series{N}, not both",
                )
                .with_property("data"));
            }
            let mut series = Vec::with_capacity(self.numbered.len());
            for entry in self.numbered.drain(..) {
                if !entry.defines() {
                    // A color for a number no other key defines.
                    return Err(invalid(
                        "the color names a series that no series{N} key defines",
                        dotted_key(entry.number, SeriesField::Color),
                    ));
                }
                series.push(entry.into_series()?);
            }
            self.data = Some(series);
        } else if let Some(data) = &mut self.data {
            // `series{N}.color` alone colors the Nth series of `data`
            // (OfficeCLI merges it over the `colors` palette).
            for entry in self.numbered.drain(..) {
                let property = dotted_key(entry.number, SeriesField::Color);
                let series = data.get_mut(entry.number - 1).ok_or_else(|| {
                    invalid("the color names a series beyond the chart's data", property)
                })?;
                series.color = entry.color;
            }
        }
        let mut data = self.data.ok_or_else(|| {
            Error::new(Code::MissingProperty, "a chart needs data")
                .with_property("data")
                .with_suggestion("pass data=Name:1,2,3;Other:4,5,6")
        })?;
        let variant = self.variant.unwrap_or(Variant::new(Kind::Column));
        // A one-series family given `Start:100,Q1:-30,...` (one single-value
        // series per item) takes the items as its points, their names as its
        // categories (OfficeCLI's waterfall flattening).
        if matches!(
            variant.kind,
            Kind::Waterfall | Kind::Funnel | Kind::Treemap | Kind::Sunburst | Kind::Pareto
        ) && data.len() > 1
            && data.iter().all(|series| series.values.len() == 1)
        {
            if self.categories.is_none() {
                self.categories = Some(data.iter().map(|series| series.name.clone()).collect());
            }
            let values = data.iter().map(|series| series.values[0]).collect();
            data = vec![Series::new("Series 1", values)];
        }
        let mut spec = Spec::new(variant);
        if let Some(title) = self.title {
            spec = spec.with_title(title);
        }
        if let Some(categories) = self.categories {
            spec = spec.with_categories(categories);
        }
        for series in data {
            spec = spec.with_series(series);
        }
        if let Some(legend) = self.legend {
            spec = spec.with_legend(legend);
        }
        if let Some(title) = self.category_axis_title {
            spec = spec.with_category_axis_title(title);
        }
        if let Some(title) = self.value_axis_title {
            spec = spec.with_value_axis_title(title);
        }
        spec = spec.with_appearance(self.appearance);
        if spec.variant().kind == Kind::Combo {
            let kinds = match (self.combo_types, self.combo_split) {
                (Some(kinds), _) => kinds,
                (None, Some(split)) => (0..spec.series().len())
                    .map(|index| {
                        if index < split {
                            Kind::Column
                        } else {
                            Kind::Line
                        }
                    })
                    .collect(),
                (None, None) => Vec::new(),
            };
            spec = spec.with_series_kinds(kinds);
        } else if self.combo_types.is_some() || self.combo_split.is_some() {
            return Err(invalid(
                "only a combo chart takes comboTypes or comboSplit",
                "comboTypes",
            ));
        }
        spec.validate().map_err(|error| {
            let message = error.to_string();
            // Attribute the leaf's refusal to the key that carries it.
            let property = [
                ("data-label position", "labelPos"),
                ("color", "colors"),
                ("gridline", "gridlines"),
                ("axis", "axisMin"),
                ("marker, line width or 3D view", "marker"),
                ("format value", "holeSize"),
                ("combo", "comboTypes"),
                ("series famil", "comboTypes"),
                ("grouping or 3D form", "chartType"),
                ("stock chart", "data"),
                ("takes one series", "data"),
                ("bin", "binCount"),
                ("waterfall total", "data"),
            ]
            .iter()
            .find(|(needle, _)| message.contains(needle))
            .map_or("data", |(_, property)| property);
            Error::new(Code::InvalidValue, message).with_property(property)
        })?;
        Ok(spec)
    }
}

/// The role of a chart axis, its OfficeCLI `role` key.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Role {
    Category,
    Value,
    Series,
}

#[cfg_attr(
    not(any(feature = "pptx", feature = "xlsx")),
    allow(dead_code, reason = "DOCX charts have no series or axis nodes yet")
)]
impl Role {
    pub(crate) const fn name(self) -> &'static str {
        match self {
            Self::Category => "category",
            Self::Value => "value",
            Self::Series => "series",
        }
    }

    pub(crate) fn parse(text: &str) -> Option<Self> {
        match text.trim().to_ascii_lowercase().as_str() {
            "category" | "cat" | "x" => Some(Self::Category),
            "value" | "val" | "y" => Some(Self::Value),
            "series" | "ser" | "z" => Some(Self::Series),
            _ => None,
        }
    }
}

/// One series of a chart, as its node prints it.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct SeriesSummary {
    /// The stable `c:idx` the editor selects the series by (its position
    /// for an extended chart).
    pub(crate) index: u32,
    pub(crate) name: Option<String>,
    pub(crate) values: Vec<f64>,
    pub(crate) categories: Vec<String>,
    /// The series' own fill, `RRGGBB` or `AARRGGBB`.
    pub(crate) color: Option<String>,
}

/// One axis of a chart, as its node prints it.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct AxisSummary {
    pub(crate) id: u32,
    pub(crate) role: Role,
    pub(crate) title: Option<String>,
    /// The scaling minimum of a value axis.
    pub(crate) min: Option<f64>,
    /// The scaling maximum of a value axis.
    pub(crate) max: Option<f64>,
    /// Whether the axis is a value axis (it has a numeric range).
    pub(crate) numeric: bool,
}

/// The read projection of one chart part.
#[derive(Debug, Clone)]
pub(crate) struct Summary {
    pub(crate) variant: Option<Variant>,
    pub(crate) title: Option<String>,
    pub(crate) legend: Option<LegendPosition>,
    pub(crate) series: Vec<SeriesSummary>,
    pub(crate) axes: Vec<AxisSummary>,
    /// The authoring spec the part reads as, when it has one; the source of
    /// the formatting keys' values and of rebuilt parts. Built on first use
    /// from the typed chart (an extended chart's is built at once).
    spec: std::sync::OnceLock<Option<Spec>>,
    /// The typed chart of an ordinary part, kept for [`Self::spec`].
    chart: Option<std::sync::Arc<Chart>>,
    /// Whether the part is an extended (`cx:chartSpace`) chart.
    pub(crate) extended: bool,
}

fn title_text(title: Option<&TitleText>) -> Option<String> {
    match title? {
        TitleText::Literal(rich) => Some(rich.text.clone()),
        TitleText::Reference(reference) => Some(reference.formula.clone()),
    }
}

/// The fill color of a series' shape properties.
fn series_color(series: &litchi_drawingml::chart::Series) -> Option<String> {
    let xml = String::from_utf8_lossy(series.shape_properties.as_ref()?.as_xml()).into_owned();
    let start = xml.find("srgbClr")?;
    xml[start..]
        .split("val=\"")
        .nth(1)?
        .get(..6)
        .map(str::to_ascii_uppercase)
}

fn drawing_code(error: &litchi_drawingml::Error, invalid: Code, fallback: Code) -> Code {
    match error {
        litchi_drawingml::Error::Limit { .. } => Code::LimitExceeded,
        litchi_drawingml::Error::InvalidLimit { .. } => Code::InvalidArgument,
        litchi_drawingml::Error::PatchConflict => Code::InvalidInput,
        litchi_drawingml::Error::Invalid(_) => invalid,
        _ => fallback,
    }
}

fn parsed(error: litchi_drawingml::Error, context: &'static str) -> Error {
    let code = drawing_code(&error, Code::CorruptFile, Code::CorruptFile);
    Error::new(code, format!("{context}: {error}"))
}

impl Summary {
    /// Read a chart part: a `c:chartSpace`, or an extended `cx:chartSpace`.
    pub(crate) fn read(xml: &[u8]) -> Result<Self> {
        if litchi_drawingml::chart::chartex::is_chartex(xml) {
            let parsed = litchi_drawingml::chart::chartex::read(
                xml,
                &litchi_drawingml::chart::Limits::DEFAULT,
            )
            .map_err(|error| parsed(error, "chart part"))?;
            return Ok(Self::of_extended(parsed.spec));
        }
        let chart =
            litchi_drawingml::chart::reader::read(xml, &litchi_drawingml::chart::Limits::DEFAULT)
                .map_err(|error| parsed(error, "chart part"))?;
        Ok(Self::of(chart))
    }

    /// The projection of an extended chart's spec: its series by position,
    /// and the category and value axes its family has.
    fn of_extended(spec: Spec) -> Self {
        let series = spec
            .series()
            .iter()
            .enumerate()
            .map(|(index, series)| SeriesSummary {
                index: u32::try_from(index).unwrap_or(u32::MAX),
                name: Some(series.name.clone()),
                values: series.values.clone(),
                categories: spec.categories().to_vec(),
                color: series.color.clone(),
            })
            .collect();
        let axes = match spec.variant().kind {
            Kind::Treemap | Kind::Sunburst => Vec::new(),
            Kind::Funnel => vec![AxisSummary {
                id: 0,
                role: Role::Category,
                title: None,
                min: None,
                max: None,
                numeric: false,
            }],
            _ => vec![
                AxisSummary {
                    id: 0,
                    role: Role::Category,
                    title: spec.category_axis_title().map(str::to_owned),
                    min: None,
                    max: None,
                    numeric: false,
                },
                AxisSummary {
                    id: 1,
                    role: Role::Value,
                    title: spec.value_axis_title().map(str::to_owned),
                    min: spec.appearance().axis_min,
                    max: spec.appearance().axis_max,
                    numeric: true,
                },
            ],
        };
        Self {
            variant: Some(spec.variant()),
            title: spec.title().map(str::to_owned),
            legend: spec.legend(),
            series,
            axes,
            spec: std::sync::OnceLock::from(Some(spec)),
            chart: None,
            extended: true,
        }
    }

    pub(crate) fn of(chart: Chart) -> Self {
        let series = chart
            .plot_area
            .type_groups
            .iter()
            .flat_map(|group: &TypeGroup| group.common().series.iter())
            .map(|series| SeriesSummary {
                index: series.index,
                name: title_text(series.title.as_ref()),
                values: series
                    .values
                    .as_ref()
                    .or(series.y_values.as_ref())
                    .map(|data| data.values.clone())
                    .unwrap_or_default(),
                categories: match (&series.categories, &series.x_values) {
                    (Some(categories), _) => categories.values.clone(),
                    (None, Some(x)) => x.values.iter().map(f64::to_string).collect(),
                    (None, None) => Vec::new(),
                },
                color: series_color(series),
            })
            .collect();
        let axes = chart
            .plot_area
            .axes
            .iter()
            .map(|axis| {
                let (role, common, range) = match axis {
                    Axis::Category(axis) => (Role::Category, &axis.common, (None, None)),
                    Axis::Date(axis) => (Role::Category, &axis.common, (None, None)),
                    Axis::Series(axis) => (Role::Series, &axis.common, (None, None)),
                    Axis::Value(axis) => (Role::Value, &axis.common, (axis.min, axis.max)),
                };
                AxisSummary {
                    id: common.axis_id,
                    role,
                    title: title_text(common.title.as_ref()),
                    min: range.0,
                    max: range.1,
                    numeric: matches!(axis, Axis::Value(_)),
                }
            })
            .collect::<Vec<_>>();
        // A scatter or bubble chart has two value axes: the first one is
        // its horizontal (category) axis.
        let mut axes = axes;
        if !axes.iter().any(|axis| axis.role == Role::Category)
            && let Some(first) = axes.iter_mut().find(|axis| axis.role == Role::Value)
        {
            first.role = Role::Category;
        }
        Self {
            variant: Variant::of(&chart),
            title: title_text(chart.title.as_ref()),
            legend: chart.legend.as_ref().map(|legend| legend.position),
            series,
            axes,
            spec: std::sync::OnceLock::new(),
            chart: Some(std::sync::Arc::new(chart)),
            extended: false,
        }
    }

    /// The authoring spec the part reads as, when it has one.
    pub(crate) fn spec(&self) -> Option<&Spec> {
        self.spec
            .get_or_init(|| self.chart.as_deref().and_then(|chart| Spec::of(chart).ok()))
            .as_ref()
    }

    /// The chart's category labels: those of its first series.
    pub(crate) fn categories(&self) -> Option<&[String]> {
        self.series
            .first()
            .map(|series| series.categories.as_slice())
            .filter(|categories| !categories.is_empty())
    }

    /// The axis of a role.
    pub(crate) fn axis(&self, role: Role) -> Option<&AxisSummary> {
        self.axes.iter().find(|axis| axis.role == role)
    }

    /// The `chartType` name, when the chart is one an authored spec can be.
    pub(crate) fn chart_type(&self) -> Option<String> {
        self.variant.map(variant_name)
    }

    /// The value of a formatting key, read from the chart's spec; `None`
    /// when the chart does not set it (the application default applies).
    pub(crate) fn format(&self, key: FormatKey) -> Option<crate::node::Value<'static>> {
        format_value(self.spec()?, key)
    }
}

fn text_value(text: impl Into<String>) -> crate::node::Value<'static> {
    crate::node::Value::Text(std::borrow::Cow::Owned(text.into()))
}

fn label_position_name(position: DataLabelPosition) -> &'static str {
    match position {
        DataLabelPosition::OutsideEnd => "outsideEnd",
        DataLabelPosition::InsideEnd => "insideEnd",
        DataLabelPosition::InsideBase => "insideBase",
        DataLabelPosition::Center => "center",
        DataLabelPosition::BestFit => "bestFit",
        DataLabelPosition::Top => "top",
        DataLabelPosition::Bottom => "bottom",
        DataLabelPosition::Left => "left",
        DataLabelPosition::Right => "right",
    }
}

fn gridline_text(line: Option<&Line>) -> String {
    match line {
        None => "true".to_owned(),
        Some(line) => {
            let mut text = line.color.clone();
            if let Some(width) = line.width {
                text.push(':');
                text.push_str(&width.to_string());
                if let Some(dash) = &line.dash {
                    text.push(':');
                    text.push_str(dash);
                }
            }
            text
        },
    }
}

/// The value of a formatting key in a spec, in the grammar `add` accepts.
#[allow(
    clippy::too_many_lines,
    reason = "one arm per formatting key, mirroring AddProps::format"
)]
pub(crate) fn format_value(spec: &Spec, key: FormatKey) -> Option<crate::node::Value<'static>> {
    use crate::node::Value;
    let appearance = spec.appearance();
    let int = |value: u32| Value::Int(i64::from(value));
    let number = |value: f64| Value::number(value);
    Some(match key {
        FormatKey::Colors => {
            let colors: Vec<String> = if appearance.colors.is_empty() {
                spec.series()
                    .iter()
                    .map(|series| series.color.clone())
                    .collect::<Option<Vec<_>>>()?
            } else {
                appearance.colors.clone()
            };
            if colors.is_empty() {
                return None;
            }
            text_value(colors.join(","))
        },
        FormatKey::DataLabels => {
            let labels = appearance.labels?;
            let parts: Vec<&str> = [
                (labels.value, "value"),
                (labels.category, "category"),
                (labels.series, "series"),
                (labels.percent, "percent"),
            ]
            .into_iter()
            .filter_map(|(shown, name)| shown.then_some(name))
            .collect();
            text_value(if parts.is_empty() {
                "none".to_owned()
            } else {
                parts.join(",")
            })
        },
        FormatKey::LabelPos => text_value(label_position_name(appearance.labels?.position?)),
        FormatKey::VaryColors => Value::Bool(appearance.vary_colors?),
        FormatKey::GapWidth => int(appearance.gap_width?),
        FormatKey::Overlap => Value::Int(i64::from(appearance.overlap?)),
        FormatKey::HoleSize => int(appearance.hole_size?),
        FormatKey::FirstSliceAngle => int(appearance.first_slice_angle?),
        FormatKey::ScatterStyle => text_value(match appearance.scatter_style? {
            ScatterStyle::Line => "line",
            ScatterStyle::LineMarker => "lineMarker",
            ScatterStyle::Marker => "marker",
            ScatterStyle::Smooth => "smooth",
            ScatterStyle::SmoothMarker => "smoothMarker",
            ScatterStyle::None => "none",
        }),
        FormatKey::RadarStyle => text_value(match appearance.radar_style? {
            RadarStyle::Standard => "standard",
            RadarStyle::Marker => "marker",
            RadarStyle::Filled => "filled",
        }),
        FormatKey::Smooth => Value::Bool(appearance.smooth?),
        FormatKey::Explosion => int(appearance.explosion?),
        FormatKey::LegendOverlay => Value::Bool(appearance.legend_overlay?),
        FormatKey::AutoTitleDeleted => Value::Bool(appearance.auto_title_deleted?),
        FormatKey::AxisMin => number(appearance.axis_min?)?,
        FormatKey::AxisMax => number(appearance.axis_max?)?,
        FormatKey::MajorUnit => number(appearance.major_unit?)?,
        FormatKey::Gridlines => {
            if appearance.hide_gridlines {
                text_value("none")
            } else {
                text_value(gridline_text(appearance.gridlines.as_ref()?.as_ref()))
            }
        },
        FormatKey::MinorGridlines => {
            text_value(gridline_text(appearance.minor_gridlines.as_ref()?.as_ref()))
        },
        FormatKey::ComboTypes => {
            if spec.variant().kind != Kind::Combo || spec.series_kinds().is_empty() {
                return None;
            }
            text_value(
                spec.series_kinds()
                    .iter()
                    .map(|kind| match kind {
                        Kind::Line => "line",
                        Kind::Area => "area",
                        _ => "column",
                    })
                    .collect::<Vec<_>>()
                    .join(","),
            )
        },
        FormatKey::ComboSplit => return None,
        FormatKey::Marker => text_value(marker_name(appearance.marker?)),
        FormatKey::MarkerSize => int(appearance.marker_size?),
        FormatKey::ShowMarkers => {
            if appearance.marker? == MarkerStyle::None {
                Value::Bool(false)
            } else {
                return None;
            }
        },
        FormatKey::LineWidth => number(appearance.line_width?)?,
        FormatKey::View3d => {
            let (x, y, perspective) = appearance.view_3d?;
            text_value(match perspective {
                Some(perspective) => format!("{x},{y},{perspective}"),
                None => format!("{x},{y}"),
            })
        },
        FormatKey::IncreaseColor => text_value(appearance.increase_color.clone()?),
        FormatKey::DecreaseColor => text_value(appearance.decrease_color.clone()?),
        FormatKey::TotalColor => text_value(appearance.total_color.clone()?),
        FormatKey::BinCount => int(appearance.bins?.count?),
        FormatKey::BinSize => number(appearance.bins?.size?)?,
        FormatKey::IntervalClosed => text_value(if appearance.bins?.closed_left {
            "l"
        } else {
            "r"
        }),
        FormatKey::OverflowBin => number(appearance.bins?.overflow?)?,
        FormatKey::UnderflowBin => number(appearance.bins?.underflow?)?,
        FormatKey::QuartileMethod => text_value(match appearance.quartile_method? {
            QuartileMethod::Exclusive => "exclusive",
            QuartileMethod::Inclusive => "inclusive",
        }),
        FormatKey::ParentLabelLayout => text_value(match appearance.parent_labels? {
            ParentLabels::Overlapping => "overlapping",
            ParentLabels::Banner => "banner",
            ParentLabels::None => "none",
        }),
    })
}

fn marker_name(style: MarkerStyle) -> &'static str {
    match style {
        MarkerStyle::Circle => "circle",
        MarkerStyle::Dash => "dash",
        MarkerStyle::Diamond => "diamond",
        MarkerStyle::Dot => "dot",
        MarkerStyle::None => "none",
        MarkerStyle::Plus => "plus",
        MarkerStyle::Square => "square",
        MarkerStyle::Star => "star",
        MarkerStyle::Triangle => "triangle",
        MarkerStyle::X => "x",
        MarkerStyle::Auto => "auto",
        _ => "auto",
    }
}

/// A chart part whose projection is read on first use: a path that only
/// edits the chart (a title `set`) never decodes it twice.
#[derive(Debug)]
pub(crate) struct LazySummary {
    xml: Vec<u8>,
    summary: std::sync::OnceLock<std::sync::Arc<Summary>>,
}

#[cfg_attr(
    not(feature = "xlsx"),
    allow(dead_code, reason = "only worksheet charts are resolved lazily")
)]
impl LazySummary {
    pub(crate) const fn new(xml: Vec<u8>) -> Self {
        Self {
            xml,
            summary: std::sync::OnceLock::new(),
        }
    }

    /// The projection, read once.
    pub(crate) fn get(&self) -> Result<&std::sync::Arc<Summary>> {
        if let Some(summary) = self.summary.get() {
            return Ok(summary);
        }
        let summary = std::sync::Arc::new(Summary::read(&self.xml)?);
        Ok(self.summary.get_or_init(|| summary))
    }
}

/// Join numbers the way OfficeCLI prints a series' values.
#[cfg_attr(
    not(any(feature = "pptx", feature = "xlsx")),
    allow(dead_code, reason = "DOCX charts have no series or axis nodes yet")
)]
pub(crate) fn join_values(values: &[f64]) -> String {
    let mut out = String::new();
    for (index, value) in values.iter().enumerate() {
        if index > 0 {
            out.push(',');
        }
        out.push_str(&value.to_string());
    }
    out
}

/// What one chart, series or axis `set` changes, applied in one chart edit:
/// source-preserving for titles, the legend, names, values, categories,
/// series colors and axis ranges; a rebuild from the chart's spec for a
/// changed family, replaced data, formatting keys, and every edit of an
/// extended chart (record 0283).
#[derive(Debug, Default)]
pub(crate) struct ChartEdit {
    pub(crate) title: Option<Option<String>>,
    pub(crate) legend: Option<Option<LegendPosition>>,
    pub(crate) categories: Option<Vec<String>>,
    pub(crate) series_names: Vec<(u32, Option<String>)>,
    pub(crate) series_values: Vec<(u32, Vec<f64>)>,
    pub(crate) series_colors: Vec<(u32, Option<String>)>,
    pub(crate) axis_titles: Vec<(u32, Option<String>)>,
    /// `(axis id, minimum, maximum)`: the complete new range.
    pub(crate) axis_ranges: Vec<(u32, Option<f64>, Option<f64>)>,
    /// The spec a rebuild writes, once a key needs one.
    pub(crate) rebuild: Option<Spec>,
    /// Whether the rebuild replaced the series (`data`).
    pub(crate) data_replaced: bool,
}

/// The result of applying a [`ChartEdit`].
#[cfg_attr(
    not(any(feature = "pptx", feature = "docx")),
    allow(
        dead_code,
        reason = "only PPTX and DOCX charts carry an embedded workbook"
    )
)]
pub(crate) struct Edited {
    /// The new chart-space XML.
    pub(crate) xml: Vec<u8>,
    /// Whether the edit changed the chart's data (values or categories),
    /// which an embedded workbook must follow.
    pub(crate) data_changed: bool,
    /// Whether the edit renamed a series, which an embedded workbook's
    /// header row may follow.
    pub(crate) names_changed: bool,
    /// The spec of a rebuilt part, from which an embedded workbook is
    /// regenerated.
    pub(crate) spec: Option<Spec>,
}

fn native(error: litchi_drawingml::Error, property: &'static str) -> Error {
    let code = drawing_code(&error, Code::InvalidValue, Code::CorruptFile);
    Error::new(code, error.to_string()).with_property(property)
}

fn not_authored(property: &'static str) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "this change rebuilds the chart, and the chart has content its spec cannot carry",
    )
    .with_property(property)
    .with_suggestion("remove the chart and add it again with the new values")
}

impl ChartEdit {
    /// The spec a rebuild starts from: the chart's own, read from its part.
    pub(crate) fn rebuild(
        &mut self,
        summary: &Summary,
        property: &'static str,
    ) -> Result<&mut Spec> {
        if self.rebuild.is_none() {
            let spec = summary
                .spec()
                .cloned()
                .ok_or_else(|| not_authored(property))?;
            self.rebuild = Some(spec);
        }
        Ok(self
            .rebuild
            .get_or_insert_with(|| unreachable!("set above")))
    }

    /// Admit a `chartType` set: the same part family (ordinary or extended).
    pub(crate) fn chart_type(&mut self, summary: &Summary, text: &str) -> Result<()> {
        let variant = parse_variant(text)?;
        if variant.kind.is_extended() != summary.extended {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "an ordinary chart and an extended chart are different parts",
            )
            .with_property("chartType")
            .with_suggestion("remove the chart and add it again as the new type"));
        }
        let spec = self.rebuild(summary, "chartType")?;
        let next = spec.clone().with_variant(variant);
        *spec = if variant.kind == Kind::Combo || spec.variant().kind == Kind::Combo {
            next.with_series_kinds(Vec::new())
        } else {
            next
        };
        Ok(())
    }

    /// Admit a `data` set: the series replaced; the categories kept unless
    /// the edit sets them too.
    pub(crate) fn data(&mut self, summary: &Summary, text: &str) -> Result<()> {
        let series = parse_data(text)?;
        let spec = self.rebuild(summary, "data")?;
        let categories = spec.categories().to_vec();
        *spec = spec.clone().with_data(categories, series);
        self.data_replaced = true;
        Ok(())
    }

    /// Admit a formatting key set.
    pub(crate) fn format(
        &mut self,
        summary: &Summary,
        key: FormatKey,
        property: &'static str,
        value: Option<Scalar<'_>>,
    ) -> Result<()> {
        let spec = self.rebuild(summary, property)?;
        let mut props = AddProps {
            appearance: spec.appearance().clone(),
            ..AddProps::default()
        };
        props.format(key, value)?;
        if let Some(kinds) = props.combo_types {
            *spec = spec.clone().with_series_kinds(kinds);
        }
        *spec.appearance_mut() = props.appearance;
        Ok(())
    }

    /// Admit a chart-level `series{N}.name`, `.values` or `.color` set: the
    /// `N`th series of the chart in reading order (the `series[N]` node).
    pub(crate) fn dotted_series(
        &mut self,
        summary: &Summary,
        number: usize,
        field: SeriesField,
        text: Option<String>,
    ) -> Result<()> {
        let property = dotted_key(number, field);
        let index = summary
            .series
            .get(number - 1)
            .map(|series| series.index)
            .ok_or_else(|| {
                Error::new(Code::OutOfRange, "the chart has no series with this number")
                    .with_property(property)
            })?;
        match field {
            SeriesField::Name => self.series_names.push((index, text)),
            SeriesField::Values => {
                let text = text.unwrap_or_default();
                if is_range_reference(&text) {
                    return Err(range_refused(property));
                }
                self.series_values
                    .push((index, parse_values(&text, property)?));
            },
            SeriesField::Color => {
                let color = text.map(|text| argb(&text, property)).transpose()?;
                self.series_colors.push((index, color));
            },
        }
        Ok(())
    }

    /// Admit an axis `min` or `max` set; the other bound keeps its value.
    #[cfg_attr(
        not(any(feature = "pptx", feature = "xlsx")),
        allow(dead_code, reason = "DOCX charts have no series or axis nodes yet")
    )]
    pub(crate) fn axis_bound(
        &mut self,
        axis: &AxisSummary,
        maximum: bool,
        value: Option<f64>,
    ) -> Result<()> {
        if !axis.numeric {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "only a value axis has a numeric range",
            )
            .with_property(if maximum { "max" } else { "min" }));
        }
        let entry = match self
            .axis_ranges
            .iter_mut()
            .find(|(id, _, _)| *id == axis.id)
        {
            Some(entry) => entry,
            None => {
                self.axis_ranges.push((axis.id, axis.min, axis.max));
                self.axis_ranges
                    .last_mut()
                    .unwrap_or_else(|| unreachable!("pushed"))
            },
        };
        if maximum {
            entry.2 = value;
        } else {
            entry.1 = value;
        }
        Ok(())
    }

    /// Apply the edit to a chart part; `None` when nothing changes.
    /// `embedded` says whether the host keeps the chart's data in an
    /// embedded workbook; without one, replaced data is written inline.
    pub(crate) fn apply(mut self, xml: Vec<u8>, embedded: bool) -> Result<Option<Edited>> {
        let extended = litchi_drawingml::chart::chartex::is_chartex(&xml);
        if extended && self.rebuild.is_none() && !self.is_empty() {
            let summary = Summary::read(&xml)?;
            self.rebuild(&summary, "chart")?;
        }
        if self.rebuild.is_some() {
            return self.apply_rebuild(&xml, embedded);
        }
        let snapshot = litchi_drawingml::chart::edit::Package::from_xml(xml)
            .map_err(|error| native(error, "chart"))?;
        let mut edit = snapshot.edit();
        if let Some(title) = self.title {
            edit.set_title(title.map(TitleText::from_string))
                .map_err(|error| native(error, "title"))?;
        }
        if let Some(legend) = self.legend {
            edit.set_legend(legend)
                .map_err(|error| native(error, "legend"))?;
        }
        let data_changed = self.categories.is_some() || !self.series_values.is_empty();
        let names_changed = !self.series_names.is_empty();
        if let Some(categories) = &self.categories {
            edit.set_categories(categories)
                .map_err(|error| native(error, "categories"))?;
        }
        for (index, name) in self.series_names {
            edit.set_series_title(index, name.map(TitleText::from_string))
                .map_err(|error| native(error, "name"))?;
        }
        for (index, values) in &self.series_values {
            edit.set_series_values(*index, values)
                .map_err(|error| native(error, "values"))?;
        }
        for (index, color) in &self.series_colors {
            edit.set_series_color(*index, color.as_deref())
                .map_err(|error| native(error, "color"))?;
        }
        for (id, title) in self.axis_titles {
            edit.set_axis_title(id, title.map(TitleText::from_string))
                .map_err(|error| native(error, "title"))?;
        }
        for (id, min, max) in self.axis_ranges {
            edit.set_axis_range(id, min, max)
                .map_err(|error| native(error, "min"))?;
        }
        let commit = edit.commit().map_err(|error| native(error, "chart"))?;
        if !commit.diagnostics().changed() {
            return Ok(None);
        }
        let snapshot = commit.into_package();
        Ok(Some(Edited {
            xml: snapshot.xml_bytes().to_vec(),
            data_changed,
            names_changed,
            spec: None,
        }))
    }

    fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.legend.is_none()
            && self.categories.is_none()
            && self.series_names.is_empty()
            && self.series_values.is_empty()
            && self.series_colors.is_empty()
            && self.axis_titles.is_empty()
            && self.axis_ranges.is_empty()
            && self.rebuild.is_none()
    }

    /// Rebuild the part from the edited spec, after checking that the part
    /// is exactly what its spec writes (so the rebuild loses nothing).
    fn apply_rebuild(self, xml: &[u8], embedded: bool) -> Result<Option<Edited>> {
        let parsed = litchi_drawingml::chart::authoring::read(
            xml,
            &litchi_drawingml::chart::Limits::DEFAULT,
        )
        .map_err(|error| parsed(error, "chart part"))?;
        if !parsed.is_authored(xml, &litchi_drawingml::chart::Limits::DEFAULT) {
            return Err(not_authored("chart"));
        }
        let mut spec = self.rebuild.unwrap_or_else(|| parsed.spec.clone());
        let original = parsed.spec;
        if let Some(title) = self.title {
            spec = match title {
                Some(title) => spec.with_title(title),
                None => spec.without_title(),
            };
        }
        if let Some(legend) = self.legend {
            spec = spec.with_legend(legend);
        }
        let mut data_changed = self.data_replaced || spec.variant() != original.variant();
        if let Some(categories) = self.categories {
            let series = spec.series().to_vec();
            spec = spec.with_data(categories, series);
            data_changed = true;
        }
        let position = |summary_index: u32, spec: &Spec| -> Option<usize> {
            // Extended charts number series by position; ordinary authored
            // charts write `c:idx` = position too.
            (usize::try_from(summary_index).ok()).filter(|index| *index < spec.series().len())
        };
        let names_changed = !self.series_names.is_empty();
        for (index, name) in self.series_names {
            let index = position(index, &spec).ok_or_else(|| missing_series("name"))?;
            spec.series_mut()[index].name = name.unwrap_or_default();
        }
        for (index, values) in self.series_values {
            let index = position(index, &spec).ok_or_else(|| missing_series("values"))?;
            spec.series_mut()[index].values = values;
            data_changed = true;
        }
        for (index, color) in self.series_colors {
            let index = position(index, &spec).ok_or_else(|| missing_series("color"))?;
            spec.series_mut()[index].color = color;
        }
        for (id, title) in self.axis_titles {
            let category = original.variant().kind.is_extended() && id == 0
                || (!original.variant().kind.is_extended() && id == CATEGORY_AXIS_ID);
            spec = match (category, title) {
                (true, Some(title)) => spec.with_category_axis_title(title),
                (false, Some(title)) => spec.with_value_axis_title(title),
                (_, None) => {
                    return Err(Error::new(
                        Code::UnsupportedCapability,
                        "an axis title of a rebuilt chart can be replaced, not cleared",
                    )
                    .with_property("title"));
                },
            };
        }
        for (id, min, max) in self.axis_ranges {
            let value_axis = if original.variant().kind.is_extended() {
                1
            } else {
                CATEGORY_AXIS_ID + 1
            };
            if id != value_axis {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "a rebuilt chart sets the range of its value axis only",
                )
                .with_property("min"));
            }
            let appearance = spec.appearance_mut();
            appearance.axis_min = min;
            appearance.axis_max = max;
        }
        if self.data_replaced && !embedded {
            spec = spec.with_binding(litchi_drawingml::chart::authoring::Binding::Literal);
        }
        let rebuilt = spec
            .to_xml(parsed.external_data.as_deref())
            .map_err(|error| {
                Error::new(Code::InvalidValue, error.to_string()).with_property("chart")
            })?;
        if rebuilt == xml {
            return Ok(None);
        }
        Ok(Some(Edited {
            xml: rebuilt,
            data_changed,
            names_changed,
            spec: Some(spec),
        }))
    }
}

/// The authored value axis id of an ordinary chart (`litchi-drawingml`'s
/// authoring constants: category 500000001, value 500000002).
const CATEGORY_AXIS_ID: u32 = 500_000_001;

fn missing_series(property: &'static str) -> Error {
    Error::new(Code::PathNotFound, "the series is missing").with_property(property)
}

#[cfg(any(feature = "pptx", feature = "docx"))]
/// Rebuild the spec an authored chart was made from, bound to an embedded
/// workbook's `Sheet1`, when the chart's references are exactly that
/// layout; `None` otherwise. A data edit regenerates the embedded workbook
/// from it so the workbook keeps matching the chart's caches.
pub(crate) fn embedded_spec(xml: &[u8]) -> Option<Spec> {
    use litchi_drawingml::chart::authoring::Binding;
    let chart =
        litchi_drawingml::chart::reader::read(xml, &litchi_drawingml::chart::Limits::DEFAULT)
            .ok()?;
    let summary = Summary::of(chart.clone());
    let variant = summary.variant?;
    let mut spec = Spec::new(variant).with_binding(Binding::sheet("Sheet1"));
    if let Some(title) = &summary.title {
        spec = spec.with_title(title.clone());
    }
    if let Some(categories) = summary.categories() {
        spec = spec.with_categories(categories.iter().cloned());
    }
    for series in &summary.series {
        spec = spec.with_series(Series::new(
            series.name.clone().unwrap_or_default(),
            series.values.clone(),
        ));
    }
    let rebuilt = spec.to_chart().ok()?;
    if !same_references(&chart, &rebuilt) {
        return None;
    }
    Some(spec)
}

/// Whether two charts' series refer to the same cells.
#[cfg(any(feature = "pptx", feature = "docx"))]
fn same_references(left: &Chart, right: &Chart) -> bool {
    let formulas = |chart: &Chart| -> Vec<Option<String>> {
        chart
            .plot_area
            .type_groups
            .iter()
            .flat_map(|group| group.common().series.iter())
            .flat_map(|series| {
                [
                    series.values.as_ref(),
                    series.y_values.as_ref(),
                    series.x_values.as_ref(),
                ]
                .into_iter()
                .map(|data| {
                    data.and_then(|data| data.source_ref.as_ref())
                        .map(|reference| reference.formula.clone())
                })
                .chain(std::iter::once(
                    series
                        .categories
                        .as_ref()
                        .and_then(|data| data.source_ref.as_ref())
                        .map(|reference| reference.formula.clone()),
                ))
                .collect::<Vec<_>>()
            })
            .collect()
    };
    formulas(left) == formulas(right)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn drawingml_native_errors_keep_typed_budget_and_replay_codes() {
        let limit = native(
            litchi_drawingml::Error::Limit {
                kind: litchi_drawingml::LimitKind::ChartTextBytes,
                observed: 9,
                maximum: 8,
            },
            "chart",
        );
        assert_eq!(limit.code(), Code::LimitExceeded);

        let invalid = native(
            litchi_drawingml::Error::InvalidLimit {
                kind: litchi_drawingml::LimitKind::ChartTextBytes,
                value: 0,
                maximum: 8,
            },
            "chart",
        );
        assert_eq!(invalid.code(), Code::InvalidArgument);

        let conflict = native(litchi_drawingml::Error::PatchConflict, "chart");
        assert_eq!(conflict.code(), Code::InvalidInput);

        let parsed_limit = parsed(
            litchi_drawingml::Error::Limit {
                kind: litchi_drawingml::LimitKind::ChartexXmlNodes,
                observed: 9,
                maximum: 8,
            },
            "chart part",
        );
        assert_eq!(parsed_limit.code(), Code::LimitExceeded);

        let parsed_invalid = parsed(
            litchi_drawingml::Error::InvalidLimit {
                kind: litchi_drawingml::LimitKind::XmlBytes,
                value: 0,
                maximum: 8,
            },
            "chart part",
        );
        assert_eq!(parsed_invalid.code(), Code::InvalidArgument);

        let parsed_conflict = parsed(litchi_drawingml::Error::PatchConflict, "chart part");
        assert_eq!(parsed_conflict.code(), Code::InvalidInput);
    }

    #[test]
    fn chart_type_names_round_trip() {
        for name in [
            "column",
            "bar",
            "stackedColumn",
            "percentStackedBar",
            "column3d",
            "stackedBar3d",
            "percentStackedColumn3d",
            "pie3d",
            "line3d",
            "area3d",
            "doughnut",
            "scatter",
            "bubble",
            "radar",
            "stock",
            "stackedArea",
            "percentStackedLine",
        ] {
            let variant = parse_variant(name).unwrap();
            assert_eq!(variant_name(variant), name);
        }
        assert_eq!(
            parse_variant("columnStacked").unwrap(),
            parse_variant("stackedColumn").unwrap()
        );
        assert_eq!(parse_variant("donut").unwrap().kind, Kind::Doughnut);
        assert_eq!(
            parse_variant("bogus").unwrap_err().code(),
            Code::InvalidValue
        );
        assert_eq!(parse_variant("combo").unwrap().kind, Kind::Combo);
        assert_eq!(parse_variant("waterfall").unwrap().kind, Kind::Waterfall);
        assert_eq!(parse_variant("boxWhisker").unwrap().kind, Kind::BoxWhisker);
        assert_eq!(
            parse_variant("pieOfPie").unwrap_err().code(),
            Code::UnsupportedCapability
        );
        assert_eq!(
            parse_variant("radar3d").unwrap_err().code(),
            Code::InvalidValue
        );
    }

    #[test]
    fn legend_and_data_grammar() {
        assert_eq!(parse_legend("b").unwrap(), Some(LegendPosition::Bottom));
        assert_eq!(parse_legend("TR").unwrap(), Some(LegendPosition::TopRight));
        assert_eq!(parse_legend("none").unwrap(), None);
        assert!(parse_legend("middle").is_err());
        let data = parse_data("East:1,2,3;West:4,5,6").unwrap();
        assert_eq!(data.len(), 2);
        assert_eq!(data[1].name, "West");
        assert_eq!(data[1].values, vec![4.0, 5.0, 6.0]);
        let single = parse_data("Start:1000,Revenue:500,Cost:-200").unwrap();
        assert_eq!(single.len(), 3);
        let colon = parse_data("Persons (year: 2021):1,2").unwrap();
        assert_eq!(colon[0].name, "Persons (year: 2021)");
        assert!(parse_data("A:1,x").is_err());
        assert!(parse_data("A:").is_err());
        assert_eq!(
            parse_list("Q1, Q2 ,Q3", "categories").unwrap(),
            ["Q1", "Q2", "Q3"]
        );
    }

    fn spec_of(keys: &[(&str, &str)]) -> Result<Spec> {
        let mut props = AddProps::default();
        for (key, value) in keys {
            if *key == "data" {
                props.data = Some(parse_data(value)?);
            } else if let Some(number) = NUMBERED_SERIES.iter().position(|name| name == key) {
                props.numbered_series(number + 1, value)?;
            } else {
                let index = DOTTED_SERIES
                    .iter()
                    .position(|name| name == key)
                    .expect("a series key");
                props.dotted_series(index / 3 + 1, SERIES_FIELDS[index % 3], value)?;
            }
        }
        props.into_spec()
    }

    fn named(spec: &Spec) -> Vec<(&str, &[f64], Option<&str>)> {
        spec.series()
            .iter()
            .map(|series| {
                (
                    series.name.as_str(),
                    series.values.as_slice(),
                    series.color.as_deref(),
                )
            })
            .collect()
    }

    #[test]
    fn dotted_series_keys_follow_officecli() {
        assert_eq!(DOTTED_SERIES[0], "series1.name");
        assert_eq!(DOTTED_SERIES[59], "series20.color");
        assert_eq!(dotted_key(7, SeriesField::Values), "series7.values");
        // The corpus form: names, values and colors, in number order
        // whatever the key order.
        let spec = spec_of(&[
            ("series2.name", "Mobile"),
            ("series1.name", "Web"),
            ("series1.values", "50,60"),
            ("series1.color", "#4472c4"),
            ("series2.values", "30, 35"),
            ("series2.color", "80ED7D31"),
        ])
        .unwrap();
        assert_eq!(
            named(&spec),
            [
                ("Web", &[50.0, 60.0][..], Some("4472C4")),
                ("Mobile", &[30.0, 35.0][..], Some("80ED7D31")),
            ]
        );
        // Values without a name are `Series N`; numbers may skip.
        let spec = spec_of(&[("series3.values", "1,2"), ("series1.values", "3")]).unwrap();
        assert_eq!(
            named(&spec),
            [
                ("Series 1", &[3.0][..], None),
                ("Series 3", &[1.0, 2.0][..], None)
            ]
        );
        // Legacy keys: `Name:values`, or bare values named `Series N`; a
        // dotted name names bare legacy values (OfficeCLI's mixed form).
        let spec = spec_of(&[
            ("series1", "Sales:1,2"),
            ("series2", "3,4"),
            ("series2.name", "Cost"),
            ("series2.color", "C00000"),
        ])
        .unwrap();
        assert_eq!(
            named(&spec),
            [
                ("Sales", &[1.0, 2.0][..], None),
                ("Cost", &[3.0, 4.0][..], Some("C00000"))
            ]
        );
        // A color alone colors the Nth series of `data`.
        let spec = spec_of(&[("data", "A:1;B:2"), ("series2.color", "00FF00")]).unwrap();
        assert_eq!(named(&spec)[1].2, Some("00FF00"));
        assert_eq!(named(&spec)[0].2, None);
    }

    #[test]
    fn dotted_series_refusals_name_their_key() {
        let refused = |keys: &[(&str, &str)]| {
            let error = spec_of(keys).unwrap_err();
            (error.code(), error.property().map(str::to_owned))
        };
        // A name without values.
        assert_eq!(
            refused(&[("series1.name", "Web")]),
            (Code::MissingProperty, Some("series1.values".into()))
        );
        // Mixing data with series definitions stays refused.
        assert_eq!(
            refused(&[
                ("data", "A:1"),
                ("series1.name", "Web"),
                ("series1.values", "1")
            ]),
            (Code::InvalidArgument, Some("data".into()))
        );
        // Two assignments of one part of a series.
        assert_eq!(
            refused(&[("series1", "A:1"), ("series1.values", "2")]),
            (Code::InvalidArgument, Some("series1.values".into()))
        );
        assert_eq!(
            refused(&[("series1", "A:1"), ("series1.name", "B")]),
            (Code::InvalidArgument, Some("series1.name".into()))
        );
        // A color for a series nothing defines.
        assert_eq!(
            refused(&[("series1.values", "1"), ("series2.color", "FF0000")]),
            (Code::InvalidValue, Some("series2.color".into()))
        );
        assert_eq!(
            refused(&[("data", "A:1"), ("series2.color", "FF0000")]),
            (Code::InvalidValue, Some("series2.color".into()))
        );
        // Only colors: the chart still has no data.
        assert_eq!(
            refused(&[("series1.color", "FF0000")]),
            (Code::MissingProperty, Some("data".into()))
        );
        // Bad values and colors.
        assert_eq!(
            refused(&[("series1.values", "1,x")]),
            (Code::InvalidValue, Some("series1.values".into()))
        );
        assert_eq!(
            refused(&[("series1.values", "1"), ("series1.color", "red")]),
            (Code::InvalidValue, Some("series1.color".into()))
        );
        // A cell range is OfficeCLI's reference form, not authored yet.
        for range in ["Sheet1!B2:B13", "B2:B13", "$B$2:$B$13"] {
            assert_eq!(
                refused(&[("series1.values", range)]),
                (Code::UnsupportedCapability, Some("series1.values".into()))
            );
        }
        assert_eq!(
            refused(&[("series4", "Sheet1!B2:B13")]),
            (Code::UnsupportedCapability, Some("series4".into()))
        );
        // `Q1:40` is a name and a value, not a range.
        assert!(!is_range_reference("Q1:40"));
        assert!(!is_range_reference("1,2,3"));
    }

    #[test]
    fn dotted_series_sets_select_series_by_number() {
        let xml = AddProps {
            data: Some(parse_data("East:1,2;West:3,4").unwrap()),
            ..AddProps::default()
        }
        .into_spec()
        .unwrap()
        .to_xml(None)
        .unwrap();
        let summary = Summary::read(&xml).unwrap();
        let mut edit = ChartEdit::default();
        edit.dotted_series(&summary, 2, SeriesField::Name, Some("North".into()))
            .unwrap();
        edit.dotted_series(&summary, 1, SeriesField::Values, Some("7,8".into()))
            .unwrap();
        let error = edit
            .dotted_series(&summary, 3, SeriesField::Name, Some("x".into()))
            .unwrap_err();
        assert_eq!(
            (error.code(), error.property()),
            (Code::OutOfRange, Some("series3.name"))
        );
        let edited = edit.apply(xml, true).unwrap().unwrap();
        let after = Summary::read(&edited.xml).unwrap();
        assert_eq!(after.series[1].name.as_deref(), Some("North"));
        assert_eq!(after.series[0].values, [7.0, 8.0]);
    }

    #[test]
    fn summaries_and_edits_follow_the_chart() {
        let spec = AddProps {
            variant: Some(parse_variant("stackedBar").unwrap()),
            title: Some("Sales".into()),
            legend: Some(Some(LegendPosition::Right)),
            categories: Some(vec!["Q1".into(), "Q2".into()]),
            data: Some(parse_data("East:1,2;West:3,4").unwrap()),
            category_axis_title: Some("Quarter".into()),
            value_axis_title: None,
            ..AddProps::default()
        }
        .into_spec()
        .unwrap()
        .with_binding(litchi_drawingml::chart::authoring::Binding::sheet("Sheet1"));
        let xml = spec.to_xml(Some("rId1")).unwrap();
        let summary = Summary::read(&xml).unwrap();
        assert_eq!(summary.chart_type().as_deref(), Some("stackedBar"));
        assert_eq!(summary.title.as_deref(), Some("Sales"));
        assert_eq!(legend_name(summary.legend), "right");
        assert_eq!(summary.categories().unwrap(), ["Q1", "Q2"]);
        assert_eq!(summary.series[1].name.as_deref(), Some("West"));
        let category = summary.axis(Role::Category).unwrap();
        assert_eq!(category.title.as_deref(), Some("Quarter"));
        #[cfg(any(feature = "pptx", feature = "docx"))]
        assert!(embedded_spec(&xml).is_some());

        let edited = ChartEdit {
            title: Some(None),
            legend: Some(None),
            series_values: vec![(1, vec![9.0, 8.0, 7.0])],
            categories: Some(vec!["A".into(), "B".into(), "C".into()]),
            axis_titles: vec![(category.id, Some("Period".into()))],
            ..ChartEdit::default()
        }
        .apply(xml.clone(), true)
        .unwrap()
        .unwrap();
        assert!(edited.data_changed);
        let after = Summary::read(&edited.xml).unwrap();
        assert_eq!(after.title, None);
        assert_eq!(after.legend, None);
        assert_eq!(after.series[1].values, vec![9.0, 8.0, 7.0]);
        assert_eq!(
            after.axis(Role::Category).unwrap().title.as_deref(),
            Some("Period")
        );
        // The workbook layout still matches, but series 0 now has fewer
        // points than the categories: the spec rebuild reflects the chart.
        #[cfg(any(feature = "pptx", feature = "docx"))]
        assert!(embedded_spec(&edited.xml).is_some());
        // An exact no-op.
        assert!(ChartEdit::default().apply(xml, true).unwrap().is_none());
    }
}
