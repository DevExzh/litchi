//! Extended charts: the `cx:chartSpace` (`ChartEx`) part of `[MS-ODRAWXML]`
//! §2.24, written from an authoring [`Spec`] and read back into one.
//!
//! Waterfall, funnel, treemap, sunburst, histogram, Pareto and
//! box-and-whisker charts are not `c:chartSpace` chart groups: Office 2016
//! stores them in a separate part family with its own content type,
//! relationship type and graphic-frame URI ([`CONTENT_TYPE`],
//! [`RELATIONSHIP_TYPE`], [`GRAPHIC_DATA_URI`]). [`Spec::to_xml`] writes this
//! part for an extended [`Kind`]; the writer is deterministic, so the same
//! spec always produces the same bytes. Each host relates the part to its
//! chart-style and color-style companions ([`style_xml`], [`colors_xml`])
//! and, in PPTX and DOCX, to an embedded workbook.
//!
//! [`read`] turns a part back into the spec that produces it, for the
//! families above. [`Parsed::is_authored`] tells whether the part is exactly
//! what that spec writes, so a host can rebuild it after an edit without
//! losing anything.
//!
//! ```
//! use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
//! use litchi_drawingml::chart::chartex;
//!
//! let spec = Spec::new(Variant::new(Kind::Waterfall))
//!     .with_title("Cash")
//!     .with_categories(["Start", "Q1", "Q2", "End"])
//!     .with_series(Series::new("Cash", vec![100.0, -30.0, 50.0, 0.0]));
//! let xml = spec.to_xml(None)?;
//! let chart = chartex::read(&xml, &litchi_drawingml::chart::Limits::DEFAULT)?;
//! assert_eq!(chart.spec.variant().kind, Kind::Waterfall);
//! assert!(chart.is_authored(&xml, &litchi_drawingml::chart::Limits::DEFAULT));
//! # Ok::<(), litchi_drawingml::Error>(())
//! ```

use std::fmt::{self, Write as _};

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::chart::Limits;
use crate::chart::authoring::{
    Appearance, Binding, Bins, Kind, Labels, ParentLabels, Parsed, QuartileMethod, Series, Spec,
    Variant, column_name,
};
use crate::chart::limits::{
    decoded_attribute_value_bytes, decoded_text_value_bytes, lossy_utf8_value_bytes,
};
use crate::chart::types::{DataLabelPosition, LegendPosition};
use crate::{Error, Result};

/// The `ChartEx` namespace (`cx:`).
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
/// The content type of a `ChartEx` part.
pub const CONTENT_TYPE: &str = "application/vnd.ms-office.chartex+xml";
/// The relationship type from a host part to a `ChartEx` part.
pub const RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2014/relationships/chartEx";
/// The `a:graphicData/@uri` of a graphic frame holding a `ChartEx` part.
pub const GRAPHIC_DATA_URI: &str = NAMESPACE;
/// The content type of a chart-style companion part.
pub const STYLE_CONTENT_TYPE: &str = "application/vnd.ms-office.chartstyle+xml";
/// The relationship type from a chart part to its chart-style part.
pub const STYLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
/// The content type of a color-style companion part.
pub const COLORS_CONTENT_TYPE: &str = "application/vnd.ms-office.chartcolorstyle+xml";
/// The relationship type from a chart part to its color-style part.
pub const COLORS_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STYLE_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

/// A formatting sink that refuses to grow beyond the caller's output budget.
/// The ChartEx writer is intentionally streaming at the formatting boundary:
/// it does not first build an unbounded `String` and only then compare its
/// length with the active policy.
struct BoundedXml {
    inner: String,
    maximum: usize,
    attempted: usize,
    exceeded: bool,
}

impl BoundedXml {
    fn new(maximum: usize) -> Self {
        Self {
            inner: String::with_capacity(4096.min(maximum)),
            maximum,
            attempted: 0,
            exceeded: false,
        }
    }

    fn push_str(&mut self, value: &str) {
        let _ = self.write_str(value);
    }

    fn finish(self) -> Result<Vec<u8>> {
        if self.exceeded {
            Err(Error::limit(
                crate::LimitKind::OutputBytes,
                self.attempted,
                self.maximum,
            ))
        } else {
            Ok(self.inner.into_bytes())
        }
    }
}

impl fmt::Write for BoundedXml {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        let attempted = self.attempted.saturating_add(value.len());
        self.attempted = attempted;
        if attempted > self.maximum {
            self.exceeded = true;
            return Err(fmt::Error);
        }
        self.inner.push_str(value);
        Ok(())
    }
}

/// XML-escape text directly into the bounded formatting sink.  Returning an
/// owned escaped `String` first would let a small output policy be bypassed by
/// a large text value before the sink sees its first byte.
struct Escaped<'a>(&'a str);

impl fmt::Display for Escaped<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut start = 0usize;
        for (index, character) in self.0.char_indices() {
            let replacement = match character {
                '&' => "&amp;",
                '<' => "&lt;",
                '>' => "&gt;",
                '"' => "&quot;",
                '\'' => "&apos;",
                _ => continue,
            };
            formatter.write_str(&self.0[start..index])?;
            formatter.write_str(replacement)?;
            start = index + character.len_utf8();
        }
        formatter.write_str(&self.0[start..])
    }
}

/// The largest `ChartEx` part [`read`] accepts.
pub const MAX_XML_BYTES: usize = 16 * 1024 * 1024;

/// An `a:solidFill` of `RRGGBB` or `AARRGGBB`.
fn solid_fill(color: &str) -> String {
    let color = color.to_ascii_uppercase();
    if color.len() == 8 {
        let alpha = u32::from_str_radix(&color[..2], 16).unwrap_or(255);
        format!(
            r#"<a:solidFill><a:srgbClr val="{}"><a:alpha val="{}"/></a:srgbClr></a:solidFill>"#,
            &color[2..],
            alpha * 100_000 / 255
        )
    } else {
        format!(r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#)
    }
}

/// The `cx:series/@layoutId` of a family.
const fn layout_id(kind: Kind) -> &'static str {
    match kind {
        Kind::Waterfall => "waterfall",
        Kind::Funnel => "funnel",
        Kind::Treemap => "treemap",
        Kind::Sunburst => "sunburst",
        Kind::BoxWhisker => "boxWhisker",
        _ => "clusteredColumn",
    }
}

/// Whether a family's data carries a category dimension.
fn has_categories(spec: &Spec) -> bool {
    !spec.categories().is_empty()
        && !matches!(spec.variant().kind, Kind::Histogram | Kind::BoxWhisker)
}

/// The zero-based points a waterfall chart draws as totals.
fn subtotals(spec: &Spec) -> Vec<usize> {
    match &spec.appearance().subtotals {
        Some(points) => {
            let mut points = points.clone();
            points.sort_unstable();
            points.dedup();
            points
        },
        None => spec
            .series()
            .first()
            .map(|series| series.values.len().saturating_sub(1))
            .into_iter()
            .collect(),
    }
}

/// The class of a waterfall point: 0 increase, 1 decrease, 2 total.
fn waterfall_class(totals: &[usize], point: usize, value: f64) -> usize {
    if totals.contains(&point) {
        2
    } else if value < 0.0 {
        1
    } else {
        0
    }
}

/// The fill of each point of a single-series chart, by point.
fn point_colors(spec: &Spec, series: &Series) -> Vec<Option<String>> {
    let appearance = spec.appearance();
    let per_point = spec.series().len() == 1 && appearance.colors.len() > 1;
    let totals = subtotals(spec);
    (0..series.values.len())
        .map(|point| {
            if spec.variant().kind == Kind::Waterfall {
                let class = waterfall_class(&totals, point, series.values[point]);
                let color = [
                    &appearance.increase_color,
                    &appearance.decrease_color,
                    &appearance.total_color,
                ][class];
                if color.is_some() {
                    return color.clone();
                }
            }
            per_point
                .then(|| appearance.colors.get(point).cloned())
                .flatten()
        })
        .collect()
}

fn gap(value: u32) -> String {
    (f64::from(value) / 100.0).to_string()
}

/// A chart or axis title, as rich text (`cx:tx/cx:rich`): the form
/// applications other than Excel read as well.
fn write_title(xml: &mut BoundedXml, chart: bool, title: &str) {
    let _ = write!(
        xml,
        r#"<cx:title{}><cx:tx><cx:rich><a:bodyPr/><a:p><a:r><a:t>{}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>"#,
        if chart {
            r#" pos="t" align="ctr" overlay="0""#
        } else {
            ""
        },
        Escaped(title)
    );
}

fn write_level_text(xml: &mut BoundedXml, points: &[String]) {
    let _ = write!(xml, r#"<cx:lvl ptCount="{}">"#, points.len());
    for (index, point) in points.iter().enumerate() {
        let _ = write!(xml, r#"<cx:pt idx="{index}">{}</cx:pt>"#, Escaped(point));
    }
    xml.push_str("</cx:lvl>");
}

fn write_data(xml: &mut BoundedXml, spec: &Spec, index: usize, series: &Series) {
    let _ = write!(xml, r#"<cx:data id="{index}">"#);
    if has_categories(spec) {
        xml.push_str(r#"<cx:strDim type="cat">"#);
        if let Some(formula) = spec.reference(0, 1, spec.categories().len()) {
            let _ = write!(xml, "<cx:f>{}</cx:f>", Escaped(&formula));
        }
        write_level_text(xml, spec.categories());
        xml.push_str("</cx:strDim>");
    }
    xml.push_str(r#"<cx:numDim type="val">"#);
    if let Some(formula) = spec.reference(index + 1, 1, series.values.len()) {
        let _ = write!(xml, "<cx:f>{}</cx:f>", Escaped(&formula));
    }
    let _ = write!(
        xml,
        r#"<cx:lvl ptCount="{}" formatCode="General">"#,
        series.values.len()
    );
    for (point, value) in series.values.iter().enumerate() {
        let _ = write!(xml, r#"<cx:pt idx="{point}">{value}</cx:pt>"#);
    }
    xml.push_str("</cx:lvl></cx:numDim></cx:data>");
}

fn write_labels(xml: &mut BoundedXml, labels: &Labels) {
    if !(labels.value || labels.category || labels.series || labels.percent) {
        return;
    }
    xml.push_str("<cx:dataLabels");
    if let Some(position) = labels.position {
        let _ = write!(xml, r#" pos="{}""#, position.xml_value());
    }
    let _ = write!(
        xml,
        r#"><cx:visibility seriesName="{}" categoryName="{}" value="{}"/></cx:dataLabels>"#,
        u8::from(labels.series),
        u8::from(labels.category),
        u8::from(labels.value)
    );
}

fn write_layout(xml: &mut BoundedXml, spec: &Spec) {
    let appearance = spec.appearance();
    match spec.variant().kind {
        Kind::Waterfall => {
            let totals = subtotals(spec);
            if !totals.is_empty() {
                xml.push_str("<cx:layoutPr><cx:subtotals>");
                for point in totals {
                    let _ = write!(xml, r#"<cx:idx val="{point}"/>"#);
                }
                xml.push_str("</cx:subtotals></cx:layoutPr>");
            }
        },
        Kind::Treemap => {
            if let Some(labels) = appearance.parent_labels {
                let _ = write!(
                    xml,
                    r#"<cx:layoutPr><cx:parentLabelLayout val="{}"/></cx:layoutPr>"#,
                    match labels {
                        ParentLabels::Overlapping => "overlapping",
                        ParentLabels::Banner => "banner",
                        ParentLabels::None => "none",
                    }
                );
            }
        },
        Kind::Pareto if has_categories(spec) => {
            xml.push_str("<cx:layoutPr><cx:aggregation/></cx:layoutPr>");
        },
        Kind::Histogram | Kind::Pareto => {
            let bins = appearance.bins.unwrap_or_default();
            let _ = write!(
                xml,
                r#"<cx:layoutPr><cx:binning intervalClosed="{}""#,
                if bins.closed_left { "l" } else { "r" }
            );
            if let Some(underflow) = bins.underflow {
                let _ = write!(xml, r#" underflow="{underflow}""#);
            }
            if let Some(overflow) = bins.overflow {
                let _ = write!(xml, r#" overflow="{overflow}""#);
            }
            match (bins.count, bins.size) {
                (Some(count), _) => {
                    let _ = write!(xml, r#"><cx:binCount val="{count}"/></cx:binning>"#);
                },
                (None, Some(size)) => {
                    let _ = write!(xml, r#"><cx:binSize val="{size}"/></cx:binning>"#);
                },
                (None, None) => xml.push_str("/>"),
            }
            xml.push_str("</cx:layoutPr>");
        },
        Kind::BoxWhisker => {
            let _ = write!(
                xml,
                concat!(
                    r#"<cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>"#,
                    r#"<cx:statistics quartileMethod="{}"/></cx:layoutPr>"#
                ),
                match appearance.quartile_method {
                    Some(QuartileMethod::Inclusive) => "inclusive",
                    _ => "exclusive",
                }
            );
        },
        _ => {},
    }
}

fn write_series(xml: &mut BoundedXml, spec: &Spec, index: usize, series: &Series) {
    let kind = spec.variant().kind;
    let appearance = spec.appearance();
    let _ = write!(
        xml,
        r#"<cx:series layoutId="{}" uniqueId="{{4C495443-4849-4000-8000-{:012X}}}">"#,
        layout_id(kind),
        index
    );
    xml.push_str("<cx:tx><cx:txData>");
    if let Some(formula) = spec.reference(index + 1, 0, 0) {
        let _ = write!(xml, "<cx:f>{}</cx:f>", Escaped(&formula));
    }
    let _ = write!(
        xml,
        "<cx:v>{}</cx:v></cx:txData></cx:tx>",
        Escaped(&series.name)
    );
    let single = spec.series().len() == 1 && appearance.colors.len() > 1;
    let color = series
        .color
        .as_ref()
        .or_else(|| (!single).then(|| appearance.colors.get(index)).flatten());
    if let Some(color) = color {
        let _ = write!(xml, "<cx:spPr>{}</cx:spPr>", solid_fill(color));
    }
    if spec.series().len() == 1 {
        for (point, color) in point_colors(spec, series).into_iter().enumerate() {
            if let Some(color) = color {
                let _ = write!(
                    xml,
                    r#"<cx:dataPt idx="{point}"><cx:spPr>{}</cx:spPr></cx:dataPt>"#,
                    solid_fill(&color)
                );
            }
        }
    }
    if let Some(labels) = &appearance.labels {
        write_labels(xml, labels);
    }
    let _ = write!(xml, r#"<cx:dataId val="{index}"/>"#);
    write_layout(xml, spec);
    if kind == Kind::Pareto {
        xml.push_str(r#"<cx:axisId val="1"/>"#);
    }
    xml.push_str("</cx:series>");
}

fn write_axes(xml: &mut BoundedXml, spec: &Spec) {
    let kind = spec.variant().kind;
    let appearance = spec.appearance();
    match kind {
        Kind::Treemap | Kind::Sunburst => return,
        Kind::Funnel => {
            let _ = write!(
                xml,
                r#"<cx:axis id="0"><cx:catScaling gapWidth="{}"/><cx:tickLabels/></cx:axis>"#,
                gap(appearance.gap_width.unwrap_or(6))
            );
            return;
        },
        _ => {},
    }
    let default_gap = match kind {
        Kind::Waterfall => 50,
        Kind::BoxWhisker => 100,
        _ => 0,
    };
    let _ = write!(
        xml,
        r#"<cx:axis id="0"><cx:catScaling gapWidth="{}"/>"#,
        gap(appearance.gap_width.unwrap_or(default_gap))
    );
    if let Some(title) = spec.category_axis_title() {
        write_title(xml, false, title);
    }
    xml.push_str("<cx:tickLabels/></cx:axis>");
    xml.push_str(r#"<cx:axis id="1"><cx:valScaling"#);
    if let Some(max) = appearance.axis_max {
        let _ = write!(xml, r#" max="{max}""#);
    }
    if let Some(min) = appearance.axis_min {
        let _ = write!(xml, r#" min="{min}""#);
    }
    if let Some(unit) = appearance.major_unit {
        let _ = write!(xml, r#" majorUnit="{unit}""#);
    }
    xml.push_str("/>");
    if let Some(title) = spec.value_axis_title() {
        write_title(xml, false, title);
    }
    if !appearance.hide_gridlines {
        xml.push_str("<cx:majorGridlines/>");
    }
    xml.push_str("<cx:tickLabels/></cx:axis>");
    if kind == Kind::Pareto {
        xml.push_str(concat!(
            r#"<cx:axis id="2"><cx:valScaling max="1" min="0"/>"#,
            r#"<cx:units unit="percentage"/><cx:tickLabels/></cx:axis>"#
        ));
    }
}

/// Write a ChartEx part under a caller-selected resource policy.
pub(crate) fn write_with_limits(
    spec: &Spec,
    external_data: Option<&str>,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    if let Some(id) = external_data {
        if id.len() > limits.max_attribute_value_bytes {
            return Err(Error::limit(
                crate::LimitKind::AttributeValueBytes,
                id.len(),
                limits.max_attribute_value_bytes,
            ));
        }
    }
    spec.validate()?;
    limits.validate_spec_for_chart(spec)?;
    let kind = spec.variant().kind;
    if !kind.is_extended() {
        return Err(Error::Invalid(
            "only an extended chart is written as a chartex part".into(),
        ));
    }
    let mut xml = BoundedXml::new(limits.max_output_bytes);
    let _ = write!(
        xml,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:a="{DRAWING_NS}" xmlns:r="{RELATIONSHIPS_NS}" xmlns:cx="{NAMESPACE}"><cx:chartData>"#
    );
    if let Some(id) = external_data {
        let _ = write!(xml, r#"<cx:externalData r:id="{}"/>"#, Escaped(id));
    }
    for (index, series) in spec.series().iter().enumerate() {
        write_data(&mut xml, spec, index, series);
    }
    xml.push_str("</cx:chartData><cx:chart>");
    if let Some(title) = spec.title() {
        write_title(&mut xml, true, title);
    }
    xml.push_str("<cx:plotArea><cx:plotAreaRegion>");
    for (index, series) in spec.series().iter().enumerate() {
        write_series(&mut xml, spec, index, series);
    }
    if kind == Kind::Pareto {
        xml.push_str(concat!(
            r#"<cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="{4C495443-4849-4000-8000-0000000000FF}">"#,
            r#"<cx:axisId val="2"/></cx:series>"#
        ));
    }
    xml.push_str("</cx:plotAreaRegion>");
    write_axes(&mut xml, spec);
    xml.push_str("</cx:plotArea>");
    if let Some(position) = spec.legend() {
        let (pos, align) = match position {
            LegendPosition::Bottom => ("b", "ctr"),
            LegendPosition::Top => ("t", "ctr"),
            LegendPosition::Left => ("l", "ctr"),
            LegendPosition::Right => ("r", "ctr"),
            LegendPosition::TopRight => ("r", "min"),
        };
        let _ = write!(
            xml,
            r#"<cx:legend pos="{pos}" align="{align}" overlay="{}"/>"#,
            u8::from(spec.appearance().legend_overlay.unwrap_or(false))
        );
    }
    xml.push_str("</cx:chart></cx:chartSpace>");
    xml.finish()
}

/// Whether a chart part's root element is a `cx:chartSpace` (an extended
/// chart) rather than a `c:chartSpace`. Only the start of the part is
/// inspected.
///
/// ```
/// use litchi_drawingml::chart::chartex;
///
/// assert!(chartex::is_chartex(
///     br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>"#
/// ));
/// assert!(!chartex::is_chartex(
///     br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#
/// ));
/// ```
#[must_use]
pub fn is_chartex(xml: &[u8]) -> bool {
    let mut reader = NsReader::from_reader(&xml[..xml.len().min(64 * 1024)]);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                return local.as_ref() == b"chartSpace"
                    && matches!(namespace, ResolveResult::Bound(namespace)
                        if namespace.as_ref() == NAMESPACE.as_bytes());
            },
            Ok(Event::Eof) | Err(_) => return false,
            Ok(_) => {},
        }
    }
}

/// A deterministic chart-style companion part (`cs:chartStyle`): every
/// entry the schema requires, with theme references only.
#[must_use]
pub fn style_xml() -> Vec<u8> {
    const ENTRIES: [&str; 30] = [
        "axisTitle",
        "categoryAxis",
        "chartArea",
        "dataLabel",
        "dataLabelCallout",
        "dataPoint",
        "dataPoint3D",
        "dataPointLine",
        "dataPointMarker",
        "dataPointWireframe",
        "dataTable",
        "downBar",
        "dropLine",
        "errorBar",
        "floor",
        "gridlineMajor",
        "gridlineMinor",
        "hiLoLine",
        "leaderLine",
        "legend",
        "plotArea",
        "plotArea3D",
        "seriesAxis",
        "seriesLine",
        "title",
        "trendline",
        "trendlineLabel",
        "upBar",
        "valueAxis",
        "wall",
    ];
    let mut xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:chartStyle xmlns:cs="{STYLE_NS}" xmlns:a="{DRAWING_NS}" id="410">"#
    );
    for entry in ENTRIES {
        let fill = if entry.starts_with("dataPoint") && entry != "dataPointLine" {
            r#"<cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef>"#
        } else {
            r#"<cs:fillRef idx="0"/>"#
        };
        let line = if entry == "dataPointLine" {
            r#"<cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef>"#
        } else {
            r#"<cs:lnRef idx="0"/>"#
        };
        let shape = match entry {
            "dataPoint" | "dataPoint3D" | "dataPointMarker" => {
                r#"<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>"#
            },
            "dataPointLine" | "dataPointWireframe" => {
                r#"<cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr>"#
            },
            "gridlineMajor" | "categoryAxis" | "valueAxis" => {
                r#"<cs:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr>"#
            },
            "chartArea" => {
                r#"<cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></cs:spPr>"#
            },
            _ => "",
        };
        let size = match entry {
            "title" => r#"<cs:defRPr sz="1862"/>"#,
            "chartArea" => r#"<cs:defRPr sz="1330"/>"#,
            _ => r#"<cs:defRPr sz="1197"/>"#,
        };
        let _ = write!(
            xml,
            concat!(
                "<cs:{entry}>{line}{fill}",
                r#"<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>"#,
                "{shape}{size}</cs:{entry}>"
            ),
            entry = entry,
            line = line,
            fill = fill,
            shape = shape,
            size = size
        );
    }
    xml.push_str("</cs:chartStyle>");
    xml.into_bytes()
}

/// A deterministic color-style companion part (`cs:colorStyle`): the six
/// theme accents, cycled with lighter and darker variations.
#[must_use]
pub fn colors_xml() -> Vec<u8> {
    let mut xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:colorStyle xmlns:cs="{STYLE_NS}" xmlns:a="{DRAWING_NS}" meth="cycle" id="10">"#
    );
    for accent in 1..=6 {
        let _ = write!(xml, r#"<a:schemeClr val="accent{accent}"/>"#);
    }
    xml.push_str(concat!(
        "<cs:variation/>",
        r#"<cs:variation><a:lumMod val="60000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="80000"/></cs:variation>"#,
        r#"<cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation>"#,
        "</cs:colorStyle>"
    ));
    xml.into_bytes()
}

// ---------------------------------------------------------------- reading

#[derive(Debug, Default)]
struct Node {
    /// Whether the element is in the `cx:` namespace.
    cx: bool,
    /// Whether the element is in the `a:` namespace.
    drawing: bool,
    name: String,
    attributes: Vec<(String, String)>,
    text: String,
    children: Vec<Node>,
}

impl Node {
    fn child(&self, name: &str) -> Option<&Self> {
        self.children
            .iter()
            .find(|child| child.cx && child.name == name)
    }

    fn all<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Self> + 'a {
        self.children
            .iter()
            .filter(move |child| child.cx && child.name == name)
    }

    fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.as_str())
    }

    fn drawing_child(&self, name: &str) -> Option<&Self> {
        self.children
            .iter()
            .find(|child| child.drawing && child.name == name)
    }

    /// The text of a `cx:tx` (or `cx:title`): its `txData/v`, or the runs
    /// of its rich text.
    fn text_body(&self, limits: Limits) -> Result<Option<(Option<String>, String)>> {
        let Some(tx) = self.child("tx") else {
            return Ok(None);
        };
        if let Some(data) = tx.child("txData") {
            let formula = data.child("f").map(|f| f.text.clone());
            let Some(value) = data.child("v").map(|v| v.text.clone()) else {
                return Ok(None);
            };
            return Ok(Some((formula, value)));
        }
        let Some(rich) = tx.child("rich") else {
            return Ok(None);
        };
        let mut text = String::new();
        collect_runs(rich, &mut text, limits)?;
        Ok(Some((None, text)))
    }
}

fn collect_runs(node: &Node, out: &mut String, limits: Limits) -> Result<()> {
    for child in &node.children {
        if child.drawing && child.name == "t" {
            let length = out.len().checked_add(child.text.len()).ok_or_else(|| {
                Error::limit(
                    crate::LimitKind::ChartTextBytes,
                    usize::MAX,
                    limits.max_text_bytes,
                )
            })?;
            if length > limits.max_text_bytes {
                return Err(Error::limit(
                    crate::LimitKind::ChartTextBytes,
                    length,
                    limits.max_text_bytes,
                ));
            }
            out.push_str(&child.text);
        } else {
            collect_runs(child, out, limits)?;
        }
    }
    Ok(())
}

fn append_text(node: &mut Node, text: &str, limits: Limits) -> Result<()> {
    let length = node.text.len().checked_add(text.len()).ok_or_else(|| {
        Error::limit(
            crate::LimitKind::ChartTextBytes,
            usize::MAX,
            limits.max_text_bytes,
        )
    })?;
    if length > limits.max_text_bytes {
        return Err(Error::limit(
            crate::LimitKind::ChartTextBytes,
            length,
            limits.max_text_bytes,
        ));
    }
    node.text.push_str(text);
    Ok(())
}

fn attribute_name(
    reader: &NsReader<&[u8]>,
    attribute: &quick_xml::events::attributes::Attribute<'_>,
) -> String {
    let local_name = attribute.key.local_name();
    let local = String::from_utf8_lossy(local_name.as_ref());
    let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
    match namespace {
        ResolveResult::Unbound => local.into_owned(),
        ResolveResult::Bound(namespace)
            if namespace.as_ref() == RELATIONSHIPS_NS.as_bytes()
                || namespace.as_ref() == STRICT_RELATIONSHIPS_NS.as_bytes() =>
        {
            format!("r:{local}")
        },
        ResolveResult::Bound(namespace) => {
            format!("{{{}}}{local}", String::from_utf8_lossy(namespace.as_ref()))
        },
        ResolveResult::Unknown(prefix) => {
            format!("{{?{}}}{local}", String::from_utf8_lossy(prefix.as_ref()))
        },
    }
}

fn parse_tree(xml: &[u8], limits: Limits) -> Result<Node> {
    let maximum = limits.max_chartex_bytes.min(MAX_XML_BYTES);
    if xml.len() > maximum {
        return Err(Error::limit(
            crate::LimitKind::ChartexPartBytes,
            xml.len(),
            maximum,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0_usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut pre_root_event_seen = false;
    let open = |reader: &NsReader<&[u8]>, element: &BytesStart<'_>| -> Result<Node> {
        let (namespace, local) = reader.resolver().resolve_element(element.name());
        let uri = match namespace {
            ResolveResult::Bound(namespace) => namespace.as_ref().to_vec(),
            _ => Vec::new(),
        };
        let mut attribute_count = 0usize;
        // Admit the complete attribute budget before decoding values into
        // owned strings. The checked iterator below retains the codec's
        // duplicate/namespace validation behavior.
        for item in element.unchecked_attributes() {
            let item = item.map_err(|error| Error::Xml(error.to_string()))?;
            attribute_count = attribute_count.checked_add(1).ok_or_else(|| {
                Error::limit(
                    crate::LimitKind::XmlAttributes,
                    usize::MAX,
                    limits.max_xml_attributes,
                )
            })?;
            if attribute_count > limits.max_xml_attributes {
                return Err(Error::limit(
                    crate::LimitKind::XmlAttributes,
                    attribute_count,
                    limits.max_xml_attributes,
                ));
            }
            let value_length = decoded_attribute_value_bytes(item.value.as_ref(), reader.decoder())
                .map_err(Error::Xml)?;
            if value_length > limits.max_attribute_value_bytes {
                return Err(Error::limit(
                    crate::LimitKind::AttributeValueBytes,
                    value_length,
                    limits.max_attribute_value_bytes,
                ));
            }
        }
        let mut attributes = Vec::new();
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
            let name = attribute_name(reader, &attribute);
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            attributes.push((name, value));
        }
        Ok(Node {
            cx: uri == NAMESPACE.as_bytes(),
            drawing: uri == DRAWING_NS.as_bytes()
                || uri == b"http://purl.oclc.org/ooxml/drawingml/main",
            name: String::from_utf8_lossy(local.as_ref()).into_owned(),
            attributes,
            text: String::new(),
            children: Vec::new(),
        })
    };
    loop {
        let event = reader.read_event()?;
        if !root_seen && !matches!(&event, Event::Decl(_) | Event::Eof) {
            pre_root_event_seen = true;
        }
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(Error::Invalid(
                        "chartex XML contains more than one root element".into(),
                    ));
                }
                root_seen = true;
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::limit(
                        crate::LimitKind::ChartexXmlNodes,
                        usize::MAX,
                        limits.max_chartex_nodes,
                    )
                })?;
                if nodes > limits.max_chartex_nodes {
                    return Err(Error::limit(
                        crate::LimitKind::ChartexXmlNodes,
                        nodes,
                        limits.max_chartex_nodes,
                    ));
                }
                let depth = stack.len().checked_add(1).ok_or_else(|| {
                    Error::limit(
                        crate::LimitKind::ChartexXmlDepth,
                        usize::MAX,
                        limits.max_chartex_depth,
                    )
                })?;
                if depth > limits.max_chartex_depth {
                    return Err(Error::limit(
                        crate::LimitKind::ChartexXmlDepth,
                        depth,
                        limits.max_chartex_depth,
                    ));
                }
                let node = open(&reader, &element)?;
                stack.push(node);
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(Error::Invalid(
                        "chartex XML contains more than one root element".into(),
                    ));
                }
                root_seen = true;
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::limit(
                        crate::LimitKind::ChartexXmlNodes,
                        usize::MAX,
                        limits.max_chartex_nodes,
                    )
                })?;
                if nodes > limits.max_chartex_nodes {
                    return Err(Error::limit(
                        crate::LimitKind::ChartexXmlNodes,
                        nodes,
                        limits.max_chartex_nodes,
                    ));
                }
                let depth = stack.len().checked_add(1).ok_or_else(|| {
                    Error::limit(
                        crate::LimitKind::ChartexXmlDepth,
                        usize::MAX,
                        limits.max_chartex_depth,
                    )
                })?;
                if depth > limits.max_chartex_depth {
                    return Err(Error::limit(
                        crate::LimitKind::ChartexXmlDepth,
                        depth,
                        limits.max_chartex_depth,
                    ));
                }
                let node = open(&reader, &element)?;
                match stack.last_mut() {
                    Some(parent) => parent.children.push(node),
                    None => {
                        root = Some(node);
                        root_closed = true;
                    },
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| Error::Xml("unbalanced chartex XML".into()))?;
                match stack.last_mut() {
                    Some(parent) => parent.children.push(node),
                    None => {
                        root = Some(node);
                        root_closed = true;
                    },
                }
            },
            Event::Text(text) => {
                let added = decoded_text_value_bytes(
                    text.as_ref(),
                    reader.decoder(),
                    XmlVersion::Implicit1_0,
                )
                .map_err(Error::Xml)?;
                if let Some(node) = stack.last() {
                    admit_text_bytes(node.text.len(), added, limits.max_text_bytes)?;
                }
                let text = text
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if let Some(node) = stack.last_mut() {
                    append_text(node, &text, limits)?;
                } else if !text.trim().is_empty() {
                    return Err(Error::Invalid(
                        "chartex XML contains text outside its root".into(),
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                let Some(node) = stack.last_mut() else {
                    return Err(Error::Invalid(
                        "chartex XML contains a reference outside its root".into(),
                    ));
                };
                let resolved = match reference
                    .resolve_char_ref()
                    .map_err(|error| Error::Xml(error.to_string()))?
                {
                    Some(character) => character.to_string(),
                    None => {
                        let name = reference
                            .decode()
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        quick_xml::escape::resolve_predefined_entity(&name)
                            .ok_or_else(|| Error::Xml("unknown XML entity".into()))?
                            .to_owned()
                    },
                };
                append_text(node, &resolved, limits)?;
            },
            Event::CData(data) => {
                let Some(node) = stack.last_mut() else {
                    return Err(Error::Invalid(
                        "chartex XML contains data outside its root".into(),
                    ));
                };
                let added = lossy_utf8_value_bytes(data.as_ref());
                admit_text_bytes(node.text.len(), added, limits.max_text_bytes)?;
                let text = String::from_utf8_lossy(&data);
                append_text(node, &text, limits)?;
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen || pre_root_event_seen {
                    return Err(Error::Invalid(
                        "chartex XML declaration is misplaced".into(),
                    ));
                }
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(Error::Invalid(
                    "chartex XML cannot contain a document type".into(),
                ));
            },
            Event::PI(_) => {
                return Err(Error::Invalid(
                    "chartex XML cannot contain processing instructions".into(),
                ));
            },
            Event::Eof => break,
            Event::Comment(_) => {},
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(Error::Invalid("chartex XML has no complete root".into()));
    }
    let root = root.ok_or_else(|| Error::Xml("empty chartex XML".into()))?;
    if !(root.cx && root.name == "chartSpace") {
        return Err(Error::Invalid("the part is not a cx:chartSpace".into()));
    }
    Ok(root)
}

fn admit_text_bytes(current: usize, added: usize, maximum: usize) -> Result<()> {
    let observed = current
        .checked_add(added)
        .ok_or_else(|| Error::limit(crate::LimitKind::ChartTextBytes, usize::MAX, maximum))?;
    if observed > maximum {
        return Err(Error::limit(
            crate::LimitKind::ChartTextBytes,
            observed,
            maximum,
        ));
    }
    Ok(())
}

/// The `RRGGBB` or `AARRGGBB` of a `cx:spPr` solid fill.
fn fill_color(node: &Node) -> Option<String> {
    let color = node
        .child("spPr")?
        .drawing_child("solidFill")?
        .drawing_child("srgbClr")?;
    let rgb = color.attribute("val")?.to_ascii_uppercase();
    match color
        .drawing_child("alpha")
        .and_then(|alpha| alpha.attribute("val")?.parse::<u32>().ok())
    {
        Some(alpha) => Some(format!("{:02X}{rgb}", (alpha * 255 + 50_000) / 100_000)),
        None => Some(rgb),
    }
}

fn number(node: &Node, name: &str) -> Option<f64> {
    node.attribute(name)?
        .trim()
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

#[derive(Debug, Default)]
struct Data {
    categories: Vec<String>,
    category_formula: Option<String>,
    values: Vec<f64>,
    value_formula: Option<String>,
}

fn level_points(level: &Node) -> impl Iterator<Item = (usize, String)> + '_ {
    level
        .all("pt")
        .filter_map(|pt| Some((pt.attribute("idx")?.parse().ok()?, pt.text.clone())))
}

fn read_data(node: &Node, limits: Limits, total_points: &mut usize) -> Result<Data> {
    let mut data = Data::default();
    if let Some(dimension) = node
        .all("strDim")
        .find(|dimension| dimension.attribute("type") == Some("cat"))
    {
        data.category_formula = dimension.child("f").map(|f| f.text.clone());
        if let Some(level) = dimension.child("lvl") {
            let count = level
                .attribute("ptCount")
                .and_then(|count| count.parse::<usize>().ok())
                .unwrap_or(0);
            if count > limits.max_categories {
                return Err(Error::limit(
                    crate::LimitKind::ChartCategories,
                    count,
                    limits.max_categories,
                ));
            }
            data.categories = vec![String::new(); count];
            for (index, text) in level_points(level) {
                if let Some(slot) = data.categories.get_mut(index) {
                    *slot = text;
                }
            }
        }
    }
    if let Some(dimension) = node
        .all("numDim")
        .find(|dimension| dimension.attribute("type") == Some("val"))
    {
        data.value_formula = dimension.child("f").map(|f| f.text.clone());
        if let Some(level) = dimension.child("lvl") {
            let count = level
                .attribute("ptCount")
                .and_then(|count| count.parse::<usize>().ok())
                .unwrap_or(0);
            if count > limits.max_points {
                return Err(Error::limit(
                    crate::LimitKind::ChartPoints,
                    count,
                    limits.max_points,
                ));
            }
            let total = total_points.checked_add(count).ok_or_else(|| {
                Error::limit(
                    crate::LimitKind::ChartPoints,
                    usize::MAX,
                    limits.max_total_points,
                )
            })?;
            if total > limits.max_total_points {
                return Err(Error::limit(
                    crate::LimitKind::ChartPoints,
                    total,
                    limits.max_total_points,
                ));
            }
            *total_points = total;
            data.values = vec![0.0; count];
            for (index, text) in level_points(level) {
                if let (Some(slot), Ok(value)) = (data.values.get_mut(index), text.trim().parse()) {
                    *slot = value;
                }
            }
        }
    }
    Ok(data)
}

/// The binding a value reference (`Sheet1!$B$2:$B$6`) implies for the
/// series at `index`, in the layout of [`Spec::sheet_layout`].
pub(crate) fn binding_of(formula: &str, index: usize) -> Option<Binding> {
    let (sheet, cells) = formula.rsplit_once('!')?;
    let sheet = sheet
        .strip_prefix('\'')
        .and_then(|name| name.strip_suffix('\''))
        .map_or_else(|| sheet.to_owned(), |name| name.replace("''", "'"));
    let first = cells.split(':').next()?.replace('$', "");
    let letters = first.bytes().take_while(u8::is_ascii_alphabetic).count();
    let (column, row) = first.split_at(letters);
    let row: usize = row.parse().ok()?;
    let mut column_index = 0_usize;
    for byte in column.bytes() {
        column_index = column_index * 26 + usize::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    let column = column_index.checked_sub(1)?.checked_sub(index + 1)?;
    let row = row.checked_sub(2)?;
    let binding = Binding::Sheet { sheet, row, column };
    // Only a layout whose first value cell this formula is.
    let expected = format!("{}{}", column_name(column + index + 1), row + 2);
    (expected == first.to_ascii_uppercase()).then_some(binding)
}

fn label_position(text: &str) -> Option<DataLabelPosition> {
    Some(match text {
        "bestFit" => DataLabelPosition::BestFit,
        "b" => DataLabelPosition::Bottom,
        "ctr" => DataLabelPosition::Center,
        "inBase" => DataLabelPosition::InsideBase,
        "inEnd" => DataLabelPosition::InsideEnd,
        "l" => DataLabelPosition::Left,
        "outEnd" => DataLabelPosition::OutsideEnd,
        "r" => DataLabelPosition::Right,
        "t" => DataLabelPosition::Top,
        _ => return None,
    })
}

/// Read a `ChartEx` part of an authored family back into its spec.
///
/// # Errors
///
/// Returns [`Error::Xml`] for malformed XML, [`Error::Invalid`] for a part
/// that is not a `cx:chartSpace` or whose series family has no authoring
/// form (a region map), and [`Error::Limit`] past [`MAX_XML_BYTES`].
pub fn read(xml: &[u8], limits: &Limits) -> Result<Parsed> {
    let limits = limits.validate()?;
    let root = parse_tree(xml, limits)?;
    let chart_data = root
        .child("chartData")
        .ok_or_else(|| Error::Invalid("the chartex part has no chart data".into()))?;
    let external_data = chart_data
        .child("externalData")
        .and_then(|data| data.attribute("r:id"))
        .map(str::to_owned);
    let data_count = chart_data.all("data").count();
    if data_count > limits.max_series {
        return Err(Error::limit(
            crate::LimitKind::ChartSeries,
            data_count,
            limits.max_series,
        ));
    }
    let mut total_points = 0usize;
    let data: Vec<(String, Data)> = chart_data
        .all("data")
        .map(|node| {
            Ok((
                node.attribute("id").unwrap_or("0").to_owned(),
                read_data(node, limits, &mut total_points)?,
            ))
        })
        .collect::<Result<_>>()?;
    let chart = root
        .child("chart")
        .ok_or_else(|| Error::Invalid("the chartex part has no chart".into()))?;
    let plot_area = chart
        .child("plotArea")
        .ok_or_else(|| Error::Invalid("the chartex chart has no plot area".into()))?;
    let region = plot_area
        .child("plotAreaRegion")
        .ok_or_else(|| Error::Invalid("the chartex chart has no plot region".into()))?;
    let mut pareto = false;
    let mut series_nodes = Vec::new();
    for series in region.all("series") {
        if series.attribute("layoutId") == Some("paretoLine") {
            pareto = true;
            continue;
        }
        if series_nodes.len() >= limits.max_series {
            return Err(Error::limit(
                crate::LimitKind::ChartSeries,
                series_nodes.len() + 1,
                limits.max_series,
            ));
        }
        series_nodes.push(series);
    }
    let first = series_nodes
        .first()
        .ok_or_else(|| Error::Invalid("the chartex chart has no series".into()))?;
    let kind = match (pareto, first.attribute("layoutId")) {
        (true, _) => Kind::Pareto,
        (false, Some("waterfall")) => Kind::Waterfall,
        (false, Some("funnel")) => Kind::Funnel,
        (false, Some("treemap")) => Kind::Treemap,
        (false, Some("sunburst")) => Kind::Sunburst,
        (false, Some("boxWhisker")) => Kind::BoxWhisker,
        (false, Some("clusteredColumn")) => Kind::Histogram,
        _ => {
            return Err(Error::Invalid(
                "the chartex series family has no authoring form".into(),
            ));
        },
    };
    let mut spec = Spec::new(Variant::new(kind)).with_legend(None);
    if let Some((_, title)) = chart
        .child("title")
        .map(|node| node.text_body(limits))
        .transpose()?
        .flatten()
    {
        spec = spec.with_title(title);
    }
    let data_of = |node: &Node| {
        let id = node
            .child("dataId")
            .and_then(|data| data.attribute("val"))
            .unwrap_or("0");
        data.iter()
            .find(|(candidate, _)| candidate == id)
            .map(|(_, data)| data)
    };
    let mut appearance = Appearance::default();
    let mut binding = None;
    let mut categories = Vec::new();
    let mut series = Vec::with_capacity(series_nodes.len());
    for (index, node) in series_nodes.iter().enumerate() {
        let data = data_of(node);
        if index == 0
            && let Some(data) = data
        {
            categories.clone_from(&data.categories);
            binding = data
                .value_formula
                .as_deref()
                .and_then(|formula| binding_of(formula, 0));
        }
        let (_, name) = node.text_body(limits)?.unwrap_or((None, String::new()));
        let values = data.map(|data| data.values.clone()).unwrap_or_default();
        let mut built = Series::new(name, values);
        if let Some(color) = fill_color(node) {
            built = built.with_color(color);
        }
        series.push(built);
        if index == 0 {
            if let Some(labels) = node.child("dataLabels") {
                let visibility = labels.child("visibility");
                let flag = |name: &str| {
                    visibility
                        .and_then(|visibility| visibility.attribute(name))
                        .is_some_and(|value| value == "1" || value == "true")
                };
                appearance.labels = Some(Labels {
                    value: flag("value"),
                    category: flag("categoryName"),
                    series: flag("seriesName"),
                    percent: false,
                    position: labels.attribute("pos").and_then(label_position),
                });
            }
            if let Some(layout) = node.child("layoutPr") {
                read_layout(layout, kind, &mut appearance, limits)?;
            } else if kind == Kind::Waterfall {
                appearance.subtotals = Some(Vec::new());
            }
        }
    }
    if series.len() == 1
        && let Some(node) = series_nodes.first()
    {
        let point_count = node.all("dataPt").count();
        if point_count > limits.max_points {
            return Err(Error::limit(
                crate::LimitKind::ChartPoints,
                point_count,
                limits.max_points,
            ));
        }
        let mut points: Vec<(usize, String)> = node
            .all("dataPt")
            .filter_map(|point| Some((point.attribute("idx")?.parse().ok()?, fill_color(point)?)))
            .collect();
        points.sort_by_key(|(index, _)| *index);
        read_point_colors(&points, &series[0].values, kind, &mut appearance);
    }
    for axis in plot_area.all("axis") {
        if let Some(scaling) = axis.child("catScaling") {
            if let Some(gap) = number(scaling, "gapWidth") {
                appearance.gap_width = Some((gap * 100.0).round() as u32);
            }
            if let Some((_, title)) = axis
                .child("title")
                .map(|node| node.text_body(limits))
                .transpose()?
                .flatten()
            {
                spec = spec.with_category_axis_title(title);
            }
        } else if let Some(scaling) = axis.child("valScaling") {
            if axis.child("units").is_some() {
                // The Pareto chart's percentage axis.
                continue;
            }
            appearance.axis_max = number(scaling, "max");
            appearance.axis_min = number(scaling, "min");
            appearance.major_unit = number(scaling, "majorUnit");
            appearance.hide_gridlines = axis.child("majorGridlines").is_none();
            if let Some((_, title)) = axis
                .child("title")
                .map(|node| node.text_body(limits))
                .transpose()?
                .flatten()
            {
                spec = spec.with_value_axis_title(title);
            }
        }
    }
    // A default gap is written whatever the spec says; read it back only
    // when it differs, so a spec's `None` survives.
    let default_gap = match kind {
        Kind::Waterfall => Some(50),
        Kind::BoxWhisker => Some(100),
        Kind::Funnel => Some(6),
        Kind::Histogram | Kind::Pareto => Some(0),
        _ => None,
    };
    if appearance.gap_width == default_gap {
        appearance.gap_width = None;
    }
    if let Some(legend) = chart.child("legend") {
        let position = match (legend.attribute("pos"), legend.attribute("align")) {
            (Some("t"), _) => LegendPosition::Top,
            (Some("l"), _) => LegendPosition::Left,
            (Some("r"), Some("min")) => LegendPosition::TopRight,
            (Some("r"), _) => LegendPosition::Right,
            _ => LegendPosition::Bottom,
        };
        spec = spec.with_legend(Some(position));
        if matches!(legend.attribute("overlay"), Some("1" | "true")) {
            appearance.legend_overlay = Some(true);
        }
    }
    spec = spec
        .with_data(categories, series)
        .with_appearance(appearance);
    if let Some(binding) = binding {
        spec = spec.with_binding(binding);
    }
    limits.validate_spec_for_chart(&spec)?;
    Ok(Parsed {
        spec,
        external_data,
    })
}

fn read_layout(
    layout: &Node,
    kind: Kind,
    appearance: &mut Appearance,
    limits: Limits,
) -> Result<()> {
    match kind {
        Kind::Waterfall => {
            let mut subtotals = Vec::new();
            if let Some(totals) = layout.child("subtotals") {
                for idx in totals.all("idx") {
                    let Some(value) = idx.attribute("val").and_then(|value| value.parse().ok())
                    else {
                        continue;
                    };
                    let count = subtotals.len().checked_add(1).ok_or_else(|| {
                        Error::limit(crate::LimitKind::ChartPoints, usize::MAX, limits.max_points)
                    })?;
                    if count > limits.max_points {
                        return Err(Error::limit(
                            crate::LimitKind::ChartPoints,
                            count,
                            limits.max_points,
                        ));
                    }
                    subtotals.push(value);
                }
            }
            appearance.subtotals = Some(subtotals);
        },
        Kind::Treemap => {
            appearance.parent_labels = layout
                .child("parentLabelLayout")
                .and_then(|node| node.attribute("val"))
                .and_then(|value| match value {
                    "overlapping" => Some(ParentLabels::Overlapping),
                    "banner" => Some(ParentLabels::Banner),
                    "none" => Some(ParentLabels::None),
                    _ => None,
                });
        },
        Kind::Histogram | Kind::Pareto => {
            if let Some(binning) = layout.child("binning") {
                let bins = Bins {
                    count: binning
                        .child("binCount")
                        .and_then(|node| node.attribute("val")?.parse().ok()),
                    size: binning
                        .child("binSize")
                        .and_then(|node| number(node, "val")),
                    closed_left: binning.attribute("intervalClosed") == Some("l"),
                    overflow: number(binning, "overflow"),
                    underflow: number(binning, "underflow"),
                };
                appearance.bins = (bins != Bins::default()).then_some(bins);
            }
        },
        Kind::BoxWhisker => {
            appearance.quartile_method = layout
                .child("statistics")
                .and_then(|node| node.attribute("quartileMethod"))
                .and_then(|value| match value {
                    "inclusive" => Some(QuartileMethod::Inclusive),
                    "exclusive" => Some(QuartileMethod::Exclusive),
                    _ => None,
                });
            if appearance.quartile_method == Some(QuartileMethod::Exclusive) {
                appearance.quartile_method = None;
            }
        },
        _ => {},
    }
    Ok(())
}

/// Recover per-point fills: a waterfall's increase, decrease and total
/// colors when each class has one color, else a color per point.
fn read_point_colors(
    points: &[(usize, String)],
    values: &[f64],
    kind: Kind,
    appearance: &mut Appearance,
) {
    if points.is_empty() {
        return;
    }
    if kind == Kind::Waterfall {
        let totals = appearance.subtotals.clone().unwrap_or_default();
        let mut classes: [Option<Option<String>>; 3] = [None, None, None];
        let mut consistent = true;
        for (point, value) in values.iter().enumerate() {
            let class = waterfall_class(&totals, point, *value);
            let color = points
                .iter()
                .find(|(index, _)| *index == point)
                .map(|(_, color)| color.clone());
            match &classes[class] {
                None => classes[class] = Some(color),
                Some(seen) if *seen == color => {},
                Some(_) => consistent = false,
            }
        }
        if consistent {
            let [increase, decrease, total] = classes;
            appearance.increase_color = increase.flatten();
            appearance.decrease_color = decrease.flatten();
            appearance.total_color = total.flatten();
            return;
        }
    }
    // A color per point, from the first; a gap ends the list.
    let mut colors = Vec::new();
    for (expected, (index, color)) in points.iter().enumerate() {
        if *index != expected {
            break;
        }
        colors.push(color.clone());
    }
    if colors.len() > 1 {
        appearance.colors = colors;
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests;
