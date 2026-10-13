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

use litchi_drawingml::chart::authoring::{
    Appearance, Grouping, Kind, Labels, Line, Series, Spec, Variant,
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
/// `radar`/`spider`, `stock`/`ohlc`) with an optional `stacked` or
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
        "waterfall" | "wf" | "funnel" | "treemap" | "sunburst" | "boxwhisker" | "histogram"
        | "pareto" | "pieofpie" | "barofpie" => {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "this chart type is not authored yet",
            )
            .with_property("chartType")
            .with_suggestion(
                "use column, bar, line, pie, doughnut, area, scatter, bubble, radar or stock",
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

/// The chart formatting keys `add` accepts on both hosts (write-only:
/// the node view does not read them back yet).
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
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

/// The formatting keys with their value types, OfficeCLI's spellings.
pub(crate) const FORMAT_PROPS: [(FormatKey, PropSpec); 26] = [
    (FormatKey::Colors, write_only("colors", ValueType::Text)),
    (
        FormatKey::DataLabels,
        write_only("dataLabels", ValueType::Text),
    ),
    (FormatKey::LabelPos, write_only("labelPos", ValueType::Text)),
    (
        FormatKey::VaryColors,
        write_only("varyColors", ValueType::Bool),
    ),
    (FormatKey::GapWidth, write_only("gapWidth", ValueType::Int)),
    (FormatKey::Overlap, write_only("overlap", ValueType::Int)),
    (FormatKey::HoleSize, write_only("holeSize", ValueType::Int)),
    (
        FormatKey::FirstSliceAngle,
        write_only("firstSliceAngle", ValueType::Int),
    ),
    (
        FormatKey::ScatterStyle,
        write_only("scatterStyle", ValueType::Text),
    ),
    (
        FormatKey::RadarStyle,
        write_only("radarStyle", ValueType::Text),
    ),
    (FormatKey::Smooth, write_only("smooth", ValueType::Bool)),
    (
        FormatKey::Explosion,
        write_only("explosion", ValueType::Int),
    ),
    (
        FormatKey::LegendOverlay,
        write_only("legend.overlay", ValueType::Bool),
    ),
    (
        FormatKey::AutoTitleDeleted,
        write_only("autoTitleDeleted", ValueType::Bool),
    ),
    (FormatKey::AxisMin, write_only("axisMin", ValueType::Number)),
    (FormatKey::AxisMax, write_only("axisMax", ValueType::Number)),
    (
        FormatKey::MajorUnit,
        write_only("majorUnit", ValueType::Number),
    ),
    (
        FormatKey::Gridlines,
        write_only("gridlines", ValueType::Text),
    ),
    (
        FormatKey::MinorGridlines,
        write_only("minorGridlines", ValueType::Text),
    ),
    (
        FormatKey::ComboTypes,
        write_only("comboTypes", ValueType::Text),
    ),
    (
        FormatKey::ComboSplit,
        write_only("comboSplit", ValueType::Int),
    ),
    (FormatKey::Marker, write_only("marker", ValueType::Text)),
    (
        FormatKey::MarkerSize,
        write_only("markerSize", ValueType::Int),
    ),
    (
        FormatKey::ShowMarkers,
        write_only("showMarkers", ValueType::Bool),
    ),
    (
        FormatKey::LineWidth,
        write_only("lineWidth", ValueType::Number),
    ),
    (FormatKey::View3d, write_only("view3d", ValueType::Text)),
];

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
    /// `series{N}` values by their number.
    pub(crate) numbered: Vec<(usize, Series)>,
    /// A combo chart's series families (`comboTypes`).
    pub(crate) combo_types: Option<Vec<Kind>>,
    /// How many leading series of a combo chart are columns (`comboSplit`).
    pub(crate) combo_split: Option<usize>,
}

impl AddProps {
    /// Admit a `series{N}` value (`Name:1,2,3`).
    pub(crate) fn numbered_series(&mut self, number: usize, text: &str) -> Result<()> {
        let property = NUMBERED_SERIES[number - 1];
        let (name, values) = text.rsplit_once(':').ok_or_else(|| {
            invalid("a numbered series is Name:1,2,3", property)
                .with_suggestion("write series1=Name:1,2,3")
        })?;
        self.numbered.push((
            number,
            Series::new(name.trim().to_owned(), parse_values(values, property)?),
        ));
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
        }
        Ok(())
    }

    /// Build the chart spec; the chart type defaults to `column` (OfficeCLI)
    /// and the legend to `bottom`.
    pub(crate) fn into_spec(mut self) -> Result<Spec> {
        if !self.numbered.is_empty() {
            if self.data.is_some() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "a chart takes data or series{N}, not both",
                )
                .with_property("data"));
            }
            self.numbered.sort_by_key(|(number, _)| *number);
            self.data = Some(self.numbered.drain(..).map(|(_, series)| series).collect());
        }
        let data = self.data.ok_or_else(|| {
            Error::new(Code::MissingProperty, "a chart needs data")
                .with_property("data")
                .with_suggestion("pass data=Name:1,2,3;Other:4,5,6")
        })?;
        let mut spec = Spec::new(self.variant.unwrap_or(Variant::new(Kind::Column)));
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
    /// The stable `c:idx` the editor selects the series by.
    pub(crate) index: u32,
    pub(crate) name: Option<String>,
    pub(crate) values: Vec<f64>,
    pub(crate) categories: Vec<String>,
}

/// One axis of a chart, as its node prints it.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct AxisSummary {
    pub(crate) id: u32,
    pub(crate) role: Role,
    pub(crate) title: Option<String>,
}

/// The read projection of one chart part.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Summary {
    pub(crate) variant: Option<Variant>,
    pub(crate) title: Option<String>,
    pub(crate) legend: Option<LegendPosition>,
    pub(crate) series: Vec<SeriesSummary>,
    pub(crate) axes: Vec<AxisSummary>,
}

fn title_text(title: Option<&TitleText>) -> Option<String> {
    match title? {
        TitleText::Literal(rich) => Some(rich.text.clone()),
        TitleText::Reference(reference) => Some(reference.formula.clone()),
    }
}

impl Summary {
    /// Read a chart part.
    pub(crate) fn read(xml: &[u8]) -> Result<Self> {
        let chart = litchi_drawingml::chart::reader::read(xml)
            .map_err(|error| Error::new(Code::CorruptFile, format!("chart part: {error}")))?;
        Ok(Self::of(&chart))
    }

    pub(crate) fn of(chart: &Chart) -> Self {
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
            })
            .collect();
        let axes = chart
            .plot_area
            .axes
            .iter()
            .map(|axis| {
                let (role, common) = match axis {
                    Axis::Category(axis) => (Role::Category, &axis.common),
                    Axis::Date(axis) => (Role::Category, &axis.common),
                    Axis::Series(axis) => (Role::Series, &axis.common),
                    Axis::Value(axis) => (Role::Value, &axis.common),
                };
                AxisSummary {
                    id: common.axis_id,
                    role,
                    title: title_text(common.title.as_ref()),
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
            variant: Variant::of(chart),
            title: title_text(chart.title.as_ref()),
            legend: chart.legend.as_ref().map(|legend| legend.position),
            series,
            axes,
        }
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
}

/// Join numbers the way OfficeCLI prints a series' values.
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

/// What one chart, series or axis `set` changes, applied in one
/// source-preserving chart edit.
#[derive(Debug, Default)]
pub(crate) struct ChartEdit {
    pub(crate) title: Option<Option<String>>,
    pub(crate) legend: Option<Option<LegendPosition>>,
    pub(crate) categories: Option<Vec<String>>,
    pub(crate) series_names: Vec<(u32, Option<String>)>,
    pub(crate) series_values: Vec<(u32, Vec<f64>)>,
    pub(crate) axis_titles: Vec<(u32, Option<String>)>,
}

/// The result of applying a [`ChartEdit`].
#[cfg_attr(
    not(feature = "pptx"),
    allow(dead_code, reason = "only PPTX charts carry an embedded workbook")
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
}

fn native(error: litchi_drawingml::Error, property: &'static str) -> Error {
    let code = match error {
        litchi_drawingml::Error::Limit { .. } => Code::InvalidValue,
        litchi_drawingml::Error::Invalid(_) => Code::InvalidValue,
        _ => Code::CorruptFile,
    };
    Error::new(code, error.to_string()).with_property(property)
}

impl ChartEdit {
    /// Apply the edit to a chart part; `None` when nothing changes.
    pub(crate) fn apply(self, xml: Vec<u8>) -> Result<Option<Edited>> {
        let snapshot = litchi_drawingml::chart::edit::Snapshot::from_xml(xml)
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
        for (id, title) in self.axis_titles {
            edit.set_axis_title(id, title.map(TitleText::from_string))
                .map_err(|error| native(error, "title"))?;
        }
        if !edit.is_changed() {
            return Ok(None);
        }
        let snapshot = edit
            .commit()
            .map_err(|error| native(error, "chart"))?
            .into_snapshot();
        Ok(Some(Edited {
            xml: snapshot.xml_bytes().to_vec(),
            data_changed,
            names_changed,
        }))
    }
}

#[cfg(feature = "pptx")]
/// Rebuild the spec an authored chart was made from, bound to an embedded
/// workbook's `Sheet1`, when the chart's references are exactly that
/// layout; `None` otherwise. A data edit regenerates the embedded workbook
/// from it so the workbook keeps matching the chart's caches.
pub(crate) fn embedded_spec(xml: &[u8]) -> Option<Spec> {
    use litchi_drawingml::chart::authoring::Binding;
    let chart = litchi_drawingml::chart::reader::read(xml).ok()?;
    let summary = Summary::of(&chart);
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
#[cfg(feature = "pptx")]
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
        assert_eq!(
            parse_variant("waterfall").unwrap_err().code(),
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
        #[cfg(feature = "pptx")]
        assert!(embedded_spec(&xml).is_some());

        let edited = ChartEdit {
            title: Some(None),
            legend: Some(None),
            series_values: vec![(1, vec![9.0, 8.0, 7.0])],
            categories: Some(vec!["A".into(), "B".into(), "C".into()]),
            axis_titles: vec![(category.id, Some("Period".into()))],
            ..ChartEdit::default()
        }
        .apply(xml.clone())
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
        #[cfg(feature = "pptx")]
        assert!(embedded_spec(&edited.xml).is_some());
        // An exact no-op.
        assert!(ChartEdit::default().apply(xml).unwrap().is_none());
    }
}
