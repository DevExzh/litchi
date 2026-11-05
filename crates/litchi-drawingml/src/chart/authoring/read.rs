//! Reading a chart part back into the [`Spec`] that writes it.
//!
//! A host that changes what a source-preserving edit cannot (a chart's
//! family, all of its data) rebuilds the part from its spec. [`read`]
//! recovers that spec from a `c:chartSpace` or `cx:chartSpace` part, with
//! its formatting where the spec has a field for it, and
//! [`Parsed::is_authored`] tells whether the rebuild keeps everything the
//! chart model reads, so a host can refuse to rebuild a chart it would
//! otherwise simplify.

use super::{Appearance, Binding, Kind, Labels, Series, Spec, Variant};
use crate::chart::Limits;
use crate::chart::axis::Axis;
use crate::chart::data::TitleText;
use crate::chart::model::{Chart, ShapeProperties};
use crate::chart::plot_area::TypeGroup;
use crate::chart::types::{RadarStyle, ScatterStyle};
use crate::{Error, Result};

/// A chart part read back as the spec that produces it.
#[derive(Debug, Clone, PartialEq)]
pub struct Parsed {
    /// The chart as an authoring spec.
    pub spec: Spec,
    /// The relationship ID of the chart's embedded workbook
    /// (`c:externalData` or `cx:externalData`), if any.
    pub external_data: Option<String>,
}

impl Parsed {
    /// Whether rebuilding the chart from [`Self::spec`] keeps everything the
    /// chart model reads from `xml`: for a `cx:chartSpace` part, whether the
    /// spec writes exactly `xml`; for a `c:chartSpace` part, whether the
    /// two parts read as the same chart (so a part that a source-preserving
    /// edit retouched still counts).
    #[must_use]
    pub fn is_authored(&self, xml: &[u8], limits: &Limits) -> bool {
        let Ok(limits) = limits.validate() else {
            return false;
        };
        if limits.validate_spec_for_chart(&self.spec).is_err() {
            return false;
        }
        let Ok(rebuilt) = self
            .spec
            .to_xml_with_policy(self.external_data.as_deref(), &limits)
        else {
            return false;
        };
        if rebuilt == xml {
            return true;
        }
        if self.spec.variant().kind.is_extended() {
            return false;
        }
        let normalized = |xml: &[u8]| -> Option<Vec<u8>> {
            let chart = crate::chart::reader::read(xml, &limits).ok()?;
            let mut out = Vec::new();
            crate::chart::writer::write_with_rels(
                &mut out,
                &chart,
                &limits,
                self.external_data.as_deref(),
                None,
            )
            .ok()?;
            Some(out)
        };
        normalized(xml).is_some_and(|original| normalized(&rebuilt) == Some(original))
    }
}

/// Read a chart part — `c:chartSpace`, or `cx:chartSpace` for an extended
/// family — back into the spec that writes it.
///
/// ```
/// use litchi_drawingml::chart::authoring::{self, Kind, Series, Spec, Variant};
///
/// let spec = Spec::new(Variant::new(Kind::Bar))
///     .with_title("Sales")
///     .with_categories(["Q1", "Q2"])
///     .with_series(Series::new("East", vec![1.0, 2.0]).with_color("4472C4"));
/// let xml = spec.to_xml(None)?;
/// let parsed = authoring::read(&xml, &litchi_drawingml::chart::Limits::DEFAULT)?;
/// assert_eq!(parsed.spec.variant().kind, Kind::Bar);
/// assert_eq!(parsed.spec.series()[0].color.as_deref(), Some("4472C4"));
/// assert!(parsed.is_authored(&xml, &litchi_drawingml::chart::Limits::DEFAULT));
/// # Ok::<(), litchi_drawingml::Error>(())
/// ```
///
/// # Errors
///
/// Returns the chart reader's errors, and [`Error::Invalid`] for a chart
/// with no authoring form (a surface, bar-of-pie or region-map chart).
pub fn read(xml: &[u8], limits: &Limits) -> Result<Parsed> {
    if crate::chart::chartex::is_chartex(xml) {
        return crate::chart::chartex::read(xml, limits);
    }
    let chart = crate::chart::reader::read(xml, limits)?;
    let external_data = chart
        .external_data
        .as_ref()
        .and_then(|data| data.relationship_id.clone());
    let parsed = Parsed {
        spec: of_chart(&chart)?,
        external_data,
    };
    limits.validate_spec_for_chart(&parsed.spec)?;
    Ok(parsed)
}

fn title_text(title: Option<&TitleText>) -> Option<String> {
    match title? {
        TitleText::Literal(rich) => Some(rich.text.clone()),
        TitleText::Reference(reference) => Some(reference.formula.clone()),
    }
}

/// The first `a:srgbClr` of shape properties, as `RRGGBB` or `AARRGGBB`.
pub(crate) fn shape_color(properties: Option<&ShapeProperties>) -> Option<String> {
    let xml = String::from_utf8_lossy(properties?.as_xml()).into_owned();
    let start = xml.find("srgbClr")?;
    let rest = &xml[start..];
    let value = rest.split("val=\"").nth(1)?.get(..6)?.to_ascii_uppercase();
    let element_end = rest.find("</a:srgbClr>").or_else(|| rest.find("/>"))?;
    let alpha = rest[..element_end]
        .split("alpha val=\"")
        .nth(1)
        .and_then(|tail| tail.split('"').next())
        .and_then(|alpha| alpha.parse::<u32>().ok());
    Some(match alpha {
        Some(alpha) => format!("{:02X}{value}", (alpha * 255 + 50_000) / 100_000),
        None => value,
    })
}

/// The line width, in points, of shape properties' `a:ln/@w`.
fn line_width(properties: Option<&ShapeProperties>) -> Option<f64> {
    let xml = String::from_utf8_lossy(properties?.as_xml()).into_owned();
    let line = &xml[xml.find("<a:ln")?..];
    let width = line
        .split_once('>')?
        .0
        .split("w=\"")
        .nth(1)?
        .split('"')
        .next()?
        .parse::<f64>()
        .ok()?;
    Some(width / 12_700.0)
}

fn binding(chart: &Chart) -> Binding {
    let first = chart
        .plot_area
        .type_groups
        .iter()
        .flat_map(|group| group.common().series.iter())
        .min_by_key(|series| series.index);
    first
        .and_then(|series| series.values.as_ref().or(series.y_values.as_ref()))
        .and_then(|data| data.source_ref.as_ref())
        .and_then(|reference| crate::chart::chartex::binding_of(&reference.formula, 0))
        .unwrap_or(Binding::Literal)
}

impl Spec {
    /// The authoring spec of a typed chart (see [`read`]), with its binding
    /// inferred from the first series' references.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for a chart with no authoring form.
    pub fn of(chart: &Chart) -> Result<Self> {
        of_chart(chart)
    }
}

/// The authoring spec of a typed chart.
#[allow(
    clippy::too_many_lines,
    reason = "one mapping per authored field, read in the writer's order"
)]
fn of_chart(chart: &Chart) -> Result<Spec> {
    let variant = Variant::of(chart)
        .ok_or_else(|| Error::Invalid("the chart family has no authoring form".into()))?;
    let mut spec = Spec::new(variant).with_legend(chart.legend.as_ref().map(|l| l.position));
    if let Some(title) = title_text(chart.title.as_ref()) {
        spec = spec.with_title(title);
    }
    let mut appearance = Appearance::default();
    if chart.auto_title_deleted != spec.title.is_none() {
        appearance.auto_title_deleted = Some(chart.auto_title_deleted);
    }
    if chart.legend.as_ref().is_some_and(|legend| legend.overlay) {
        appearance.legend_overlay = Some(true);
    }
    // Every series with its family, in index order.
    let mut series_list = Vec::new();
    for group in &chart.plot_area.type_groups {
        let family = match group {
            TypeGroup::Line(_) | TypeGroup::Line3D(_) => Kind::Line,
            TypeGroup::Area(_) | TypeGroup::Area3D(_) => Kind::Area,
            _ => Kind::Column,
        };
        for series in &group.common().series {
            series_list.push((series, family));
        }
    }
    series_list.sort_by_key(|(series, _)| series.index);
    let xy = matches!(variant.kind, Kind::Scatter | Kind::Bubble);
    if let Some((first, _)) = series_list.first() {
        let categories = match (&first.categories, &first.x_values) {
            (Some(categories), _) => categories.values.clone(),
            (None, Some(x)) if xy => x.values.iter().map(f64::to_string).collect(),
            _ => Vec::new(),
        };
        spec.categories = categories;
    }
    let single_pie = matches!(variant.kind, Kind::Pie | Kind::Doughnut) && series_list.len() == 1;
    for (series, _) in &series_list {
        let values = series
            .values
            .as_ref()
            .or(series.y_values.as_ref())
            .map(|data| data.values.clone())
            .unwrap_or_default();
        let mut built = Series::new(
            title_text(series.title.as_ref()).unwrap_or_default(),
            values,
        );
        if let Some(color) = shape_color(series.shape_properties.as_ref()) {
            built = built.with_color(color);
        }
        if single_pie && !series.data_points.is_empty() {
            appearance.colors = series
                .data_points
                .iter()
                .filter_map(|point| shape_color(point.shape_properties.as_ref()))
                .collect();
        }
        if series.smooth {
            appearance.smooth = Some(true);
        }
        if series.explosion.is_some() {
            appearance.explosion = series.explosion;
        }
        if series.marker_present {
            appearance.marker = series.marker_symbol;
            appearance.marker_size = series.marker_size;
            appearance.marker_color = shape_color(series.marker_shape_properties.as_ref());
        }
        if let Some(width) = line_width(series.shape_properties.as_ref())
            && (width - 2.25).abs() > f64::EPSILON
        {
            appearance.line_width = Some(width);
        }
        spec.series.push(built);
    }
    if variant.kind == Kind::Combo {
        spec.series_kinds = series_list.iter().map(|(_, family)| *family).collect();
    }
    if let Some(group) = chart.plot_area.type_groups.first() {
        let common = group.common();
        let default_vary = matches!(variant.kind, Kind::Pie | Kind::Doughnut);
        if common.vary_colors != default_vary {
            appearance.vary_colors = Some(common.vary_colors);
        }
        // The writer's default container shows nothing and is not deleted;
        // an authored `Labels` writes a deleted container when it shows
        // nothing.
        if let Some(labels) = common.data_labels.as_ref().filter(|labels| {
            labels.deleted
                || labels.show_value
                || labels.show_category_name
                || labels.show_series_name
                || labels.show_percent
        }) {
            appearance.labels = Some(Labels {
                value: labels.show_value,
                category: labels.show_category_name,
                series: labels.show_series_name,
                percent: labels.show_percent,
                position: labels.position,
            });
        }
        match group {
            TypeGroup::Bar(bar) => {
                if bar.gap_width != Some(150) {
                    appearance.gap_width = bar.gap_width;
                }
                let default_overlap = (variant.grouping != super::Grouping::Standard
                    && variant.kind != Kind::Combo)
                    .then_some(100);
                if bar.overlap != default_overlap {
                    appearance.overlap = bar.overlap;
                }
            },
            TypeGroup::Bar3D(bar) => {
                if bar.gap_width != Some(150) {
                    appearance.gap_width = bar.gap_width;
                }
            },
            TypeGroup::Line(line) if line.smooth => appearance.smooth = Some(true),
            TypeGroup::Pie(pie) if pie.first_slice_angle != 0 => {
                appearance.first_slice_angle = Some(pie.first_slice_angle);
            },
            TypeGroup::Doughnut(doughnut) => {
                if doughnut.first_slice_angle != 0 {
                    appearance.first_slice_angle = Some(doughnut.first_slice_angle);
                }
                if doughnut.hole_size != 50 {
                    appearance.hole_size = Some(doughnut.hole_size);
                }
            },
            TypeGroup::Scatter(scatter) if scatter.style != ScatterStyle::LineMarker => {
                appearance.scatter_style = Some(scatter.style);
            },
            TypeGroup::Radar(radar) if radar.style != RadarStyle::Marker => {
                appearance.radar_style = Some(radar.style);
            },
            _ => {},
        }
    }
    let mut value_axes = 0;
    for axis in &chart.plot_area.axes {
        match axis {
            Axis::Category(axis) => {
                if let Some(title) = title_text(axis.common.title.as_ref()) {
                    spec.category_axis_title = Some(title);
                }
            },
            Axis::Value(axis) => {
                value_axes += 1;
                // A scatter chart's first value axis is its horizontal one.
                if xy && value_axes == 1 {
                    if let Some(title) = title_text(axis.common.title.as_ref()) {
                        spec.category_axis_title = Some(title);
                    }
                    continue;
                }
                if let Some(title) = title_text(axis.common.title.as_ref()) {
                    spec.value_axis_title = Some(title);
                }
                appearance.axis_min = axis.min;
                appearance.axis_max = axis.max;
                appearance.major_unit = axis.major_unit;
                appearance.hide_gridlines = !axis.common.show_major_gridlines;
                if axis.common.show_minor_gridlines {
                    appearance.minor_gridlines = Some(None);
                }
            },
            Axis::Date(_) | Axis::Series(_) => {},
        }
    }
    if let Some(view) = &chart.view_3d {
        let default = if variant.kind == Kind::Pie {
            (Some(30), Some(0))
        } else {
            (Some(15), Some(20))
        };
        if !variant.three_d || (view.rot_x, view.rot_y) != default || view.perspective.is_some() {
            appearance.view_3d = Some((
                view.rot_x.unwrap_or(0),
                view.rot_y.unwrap_or(0),
                view.perspective,
            ));
        }
    }
    spec.appearance = appearance;
    spec.binding = binding(chart);
    Ok(spec)
}
