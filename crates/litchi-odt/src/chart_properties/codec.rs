//! Bounded XML parsing, serialization, and lossless mutation for chart styles.

use super::model::{
    Angle, AxisLabelPosition, AxisPosition, DataLabelNumber, Direction, Double, EmptyCellTreatment,
    ErrorCategory, Integer, Interpolation, LabelArrangement, LabelPosition, LabelSeparator,
    NonNegativeInteger, NonNegativeLength, Percent, PositiveInteger, RegressionType, SeriesSource,
    SolidType, StyleProperties, StylePropertiesSet, StyleRecord, SymbolImage, SymbolName,
    SymbolType, TickMarkPosition, bad, safe,
};
use super::{
    CHART, CHART_NS, MAX_ATTRIBUTES, MAX_DEPTH, MAX_EVENTS, MAX_STYLES, MAX_TOTAL, MAX_VALUE,
    MAX_XML, OFFICE, OFFICE_NS, STYLE, STYLE_NS, TEXT, TEXT_NS, XLINK, XLINK_NS,
};
use litchi_core::{Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

impl SymbolImage {
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        Ok(format!(
            r#"<chart:symbol-image xmlns:chart="{CHART_NS}" xmlns:xlink="{XLINK_NS}" xlink:href="{}"/>"#,
            escape_xml(&self.href)
        ))
    }
}
impl LabelSeparator {
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_paragraph(self.paragraph_xml())?;
        Ok(format!(
            r#"<chart:label-separator xmlns:chart="{CHART_NS}">{} </chart:label-separator>"#,
            self.paragraph_xml()
        )
        .replace("</text:p> </chart:", "</text:p></chart:"))
    }
}

pub(super) fn validate_paragraph(xml: &str) -> Result<()> {
    if xml.len() > MAX_VALUE {
        return Err(bad("chart label paragraph is too large"));
    }
    let wrapped = format!(
        r#"<wrapper xmlns:text="{TEXT_NS}" xmlns:chart="{CHART_NS}" xmlns:style="{STYLE_NS}" xmlns:xlink="{XLINK_NS}">{xml}</wrapper>"#
    );
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    let mut depth = 0usize;
    let mut paragraph = false;
    let mut events = 0;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("chart label paragraph has too many events"));
        }
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if depth >= MAX_DEPTH {
                    return Err(bad("chart label paragraph is too deep"));
                }
                let current = element(&reader, start.name());
                if depth == 0 {
                    if current.0 != Ns::Other || current.1 != b"wrapper" {
                        return Err(bad("invalid chart label wrapper"));
                    }
                } else if depth == 1 {
                    if paragraph || current.0 != Ns::Text || current.1 != b"p" {
                        return Err(bad("chart label separator requires one text:p"));
                    }
                    paragraph = true
                }
                depth += 1
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                if depth == 1 {
                    if paragraph || current.0 != Ns::Text || current.1 != b"p" {
                        return Err(bad("chart label separator requires one text:p"));
                    }
                    paragraph = true
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if depth == 1 && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad("chart label separator allows only one text:p"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if depth == 1 && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad("chart label separator allows only one text:p"));
                }
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| bad("invalid chart label paragraph"))?
            },
            Ok(Event::Decl(_)) | Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad(
                    "declarations, DTDs, and processing instructions are not allowed in chart labels",
                ));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid chart label paragraph: {error}"))),
        }
    }
    if !paragraph || depth != 0 {
        return Err(bad("truncated chart label paragraph"));
    }
    Ok(())
}

impl StyleProperties {
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> {
        let xml = format!(
            r#"<office:document xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}"><office:styles><style:style style:name="fragment" style:family="chart">{fragment}</style:style></office:styles></office:document>"#
        );
        let mut set = parse_chart_style_properties(&xml)?;
        set.styles
            .pop()
            .and_then(|style| style.properties)
            .ok_or_else(|| bad("fragment does not contain style:chart-properties"))
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:chart-properties xmlns:style="{STYLE_NS}" xmlns:chart="{CHART_NS}" xmlns:text="{TEXT_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        macro_rules! a {
            ($field:expr,$name:literal,$render:expr) => {
                if let Some(value) = $field {
                    let rendered = ($render)(value);
                    xml.push_str(&format!(concat!(" ", $name, "=\"{}\""), rendered))
                }
            };
        }
        macro_rules! b {
            ($field:expr,$name:literal) => {
                a!($field, $name, |value: bool| if value {
                    "true"
                } else {
                    "false"
                })
            };
        }
        b!(self.scale_text, "chart:scale-text");
        b!(self.three_dimensional, "chart:three-dimensional");
        b!(self.deep, "chart:deep");
        b!(self.right_angled_axes, "chart:right-angled-axes");
        a!(self.symbol_type, "chart:symbol-type", |v: SymbolType| v
            .xml());
        a!(self.symbol_name, "chart:symbol-name", |v: SymbolName| v
            .xml());
        a!(
            self.symbol_width.as_ref(),
            "chart:symbol-width",
            |v: &NonNegativeLength| v.as_str().to_owned()
        );
        a!(
            self.symbol_height.as_ref(),
            "chart:symbol-height",
            |v: &NonNegativeLength| v.as_str().to_owned()
        );
        b!(self.sort_by_x_values, "chart:sort-by-x-values");
        b!(self.vertical, "chart:vertical");
        b!(self.connect_bars, "chart:connect-bars");
        a!(self.gap_width.as_ref(), "chart:gap-width", |v: &Integer| v
            .as_str()
            .to_owned());
        a!(self.overlap.as_ref(), "chart:overlap", |v: &Integer| v
            .as_str()
            .to_owned());
        b!(self.group_bars_per_axis, "chart:group-bars-per-axis");
        b!(self.japanese_candle_stick, "chart:japanese-candle-stick");
        a!(
            self.interpolation,
            "chart:interpolation",
            |v: Interpolation| v.xml()
        );
        a!(
            self.spline_order.as_ref(),
            "chart:spline-order",
            |v: &PositiveInteger| v.as_str().to_owned()
        );
        a!(
            self.spline_resolution.as_ref(),
            "chart:spline-resolution",
            |v: &PositiveInteger| v.as_str().to_owned()
        );
        a!(
            self.pie_offset.as_ref(),
            "chart:pie-offset",
            |v: &NonNegativeInteger| v.as_str().to_owned()
        );
        a!(
            self.angle_offset.as_ref(),
            "chart:angle-offset",
            |v: &Angle| escape_xml(v.as_str())
        );
        a!(self.hole_size.as_ref(), "chart:hole-size", |v: &Percent| v
            .as_str()
            .to_owned());
        b!(self.lines, "chart:lines");
        a!(self.solid_type, "chart:solid-type", |v: SolidType| v.xml());
        b!(self.stacked, "chart:stacked");
        b!(self.percentage, "chart:percentage");
        a!(
            self.treat_empty_cells,
            "chart:treat-empty-cells",
            |v: EmptyCellTreatment| v.xml()
        );
        b!(
            self.link_data_style_to_source,
            "chart:link-data-style-to-source"
        );
        b!(self.logarithmic, "chart:logarithmic");
        a!(self.maximum.as_ref(), "chart:maximum", |v: &Double| v
            .as_str()
            .to_owned());
        a!(self.minimum.as_ref(), "chart:minimum", |v: &Double| v
            .as_str()
            .to_owned());
        a!(self.origin.as_ref(), "chart:origin", |v: &Double| v
            .as_str()
            .to_owned());
        a!(
            self.interval_major.as_ref(),
            "chart:interval-major",
            |v: &Double| v.as_str().to_owned()
        );
        a!(
            self.interval_minor_divisor.as_ref(),
            "chart:interval-minor-divisor",
            |v: &PositiveInteger| v.as_str().to_owned()
        );
        b!(self.tick_marks_major_inner, "chart:tick-marks-major-inner");
        b!(self.tick_marks_major_outer, "chart:tick-marks-major-outer");
        b!(self.tick_marks_minor_inner, "chart:tick-marks-minor-inner");
        b!(self.tick_marks_minor_outer, "chart:tick-marks-minor-outer");
        b!(self.reverse_direction, "chart:reverse-direction");
        b!(self.display_label, "chart:display-label");
        b!(self.text_overlap, "chart:text-overlap");
        b!(self.line_break, "text:line-break");
        a!(
            self.label_arrangement,
            "chart:label-arrangement",
            |v: LabelArrangement| v.xml()
        );
        a!(self.direction, "style:direction", |v: Direction| v.xml());
        a!(
            self.rotation_angle.as_ref(),
            "style:rotation-angle",
            |v: &Angle| escape_xml(v.as_str())
        );
        a!(
            self.data_label_number,
            "chart:data-label-number",
            |v: DataLabelNumber| v.xml()
        );
        b!(self.data_label_text, "chart:data-label-text");
        b!(self.data_label_symbol, "chart:data-label-symbol");
        a!(
            self.label_position,
            "chart:label-position",
            |v: LabelPosition| v.xml()
        );
        a!(
            self.label_position_negative,
            "chart:label-position-negative",
            |v: LabelPosition| v.xml()
        );
        b!(self.visible, "chart:visible");
        b!(self.auto_position, "chart:auto-position");
        b!(self.auto_size, "chart:auto-size");
        b!(self.mean_value, "chart:mean-value");
        a!(
            self.error_category,
            "chart:error-category",
            |v: ErrorCategory| v.xml()
        );
        a!(
            self.error_percentage.as_ref(),
            "chart:error-percentage",
            |v: &Double| v.as_str().to_owned()
        );
        a!(
            self.error_margin.as_ref(),
            "chart:error-margin",
            |v: &Double| v.as_str().to_owned()
        );
        a!(
            self.error_lower_limit.as_ref(),
            "chart:error-lower-limit",
            |v: &Double| v.as_str().to_owned()
        );
        a!(
            self.error_upper_limit.as_ref(),
            "chart:error-upper-limit",
            |v: &Double| v.as_str().to_owned()
        );
        b!(self.error_upper_indicator, "chart:error-upper-indicator");
        b!(self.error_lower_indicator, "chart:error-lower-indicator");
        a!(
            self.series_source,
            "chart:series-source",
            |v: SeriesSource| v.xml()
        );
        a!(
            self.regression_type,
            "chart:regression-type",
            |v: RegressionType| v.xml()
        );
        a!(
            self.axis_position.as_ref(),
            "chart:axis-position",
            |v: &AxisPosition| v.xml().to_owned()
        );
        a!(
            self.axis_label_position,
            "chart:axis-label-position",
            |v: AxisLabelPosition| v.xml()
        );
        a!(
            self.tick_mark_position,
            "chart:tick-mark-position",
            |v: TickMarkPosition| v.xml()
        );
        b!(self.include_hidden_cells, "chart:include-hidden-cells");
        if self.symbol_image.is_some() || self.label_separator.is_some() {
            xml.push('>');
            if let Some(image) = &self.symbol_image {
                xml.push_str(&image.to_xml_fragment()?)
            }
            if let Some(label) = &self.label_separator {
                xml.push_str(&label.to_xml_fragment()?)
            }
            xml.push_str("</style:chart-properties>")
        } else {
            xml.push_str("/>")
        }
        Ok(xml)
    }
}

impl StyleRecord {
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml = format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="chart""#);
        if let Some(value) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)))
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ))
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"))
        } else {
            xml.push_str("/>")
        }
        Ok(xml)
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Chart,
    Text,
    Xlink,
    Other,
}
fn ns(value: ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == CHART => Ns::Chart,
        ResolveResult::Bound(value) if value.as_ref() == TEXT => Ns::Text,
        ResolveResult::Bound(value) if value.as_ref() == XLINK => Ns::Xlink,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (ns(namespace), local.as_ref().to_vec())
}
fn attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Vec<(Ns, Vec<u8>, String)>> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid chart property attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if out.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many chart property attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate chart property attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid chart property value: {error}")))?
            .into_owned();
        safe(&value, "chart property value", true)?;
        out.push((key.0, key.1, value))
    }
    Ok(out)
}
fn take(a: &mut Vec<(Ns, Vec<u8>, String)>, namespace: Ns, local: &[u8]) -> Option<String> {
    a.iter()
        .position(|value| value.0 == namespace && value.1 == local)
        .map(|index| a.remove(index).2)
}
fn boolean(value: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(bad("ODF boolean must be true or false")),
    }
}
fn e<T>(value: Option<String>, parse: fn(&str) -> Result<T>) -> Result<Option<T>> {
    value.map(|value| parse(&value)).transpose()
}
fn header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<StyleRecord>> {
    let mut a = attrs(reader, version, start)?;
    if take(&mut a, Ns::Style, b"family").as_deref() != Some("chart") {
        return Ok(None);
    }
    let style = StyleRecord {
        name: take(&mut a, Ns::Style, b"name"),
        parent_style_name: take(&mut a, Ns::Style, b"parent-style-name"),
        is_default_style: default,
        properties: None,
    };
    style.validate()?;
    Ok(Some(style))
}
fn properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<StyleProperties> {
    let mut a = attrs(reader, version, start)?;
    macro_rules! b {
        ($n:literal) => {
            take(&mut a, Ns::Chart, $n)
                .map(|v| boolean(&v))
                .transpose()?
        };
    }
    let value = StyleProperties {
        scale_text: b!(b"scale-text"),
        three_dimensional: b!(b"three-dimensional"),
        deep: b!(b"deep"),
        right_angled_axes: b!(b"right-angled-axes"),
        symbol_type: e(take(&mut a, Ns::Chart, b"symbol-type"), SymbolType::parse)?,
        symbol_name: e(take(&mut a, Ns::Chart, b"symbol-name"), SymbolName::parse)?,
        symbol_image: None,
        symbol_width: take(&mut a, Ns::Chart, b"symbol-width")
            .map(NonNegativeLength::new)
            .transpose()?,
        symbol_height: take(&mut a, Ns::Chart, b"symbol-height")
            .map(NonNegativeLength::new)
            .transpose()?,
        sort_by_x_values: b!(b"sort-by-x-values"),
        vertical: b!(b"vertical"),
        connect_bars: b!(b"connect-bars"),
        gap_width: take(&mut a, Ns::Chart, b"gap-width")
            .map(Integer::new)
            .transpose()?,
        overlap: take(&mut a, Ns::Chart, b"overlap")
            .map(Integer::new)
            .transpose()?,
        group_bars_per_axis: b!(b"group-bars-per-axis"),
        japanese_candle_stick: b!(b"japanese-candle-stick"),
        interpolation: e(
            take(&mut a, Ns::Chart, b"interpolation"),
            Interpolation::parse,
        )?,
        spline_order: take(&mut a, Ns::Chart, b"spline-order")
            .map(PositiveInteger::new)
            .transpose()?,
        spline_resolution: take(&mut a, Ns::Chart, b"spline-resolution")
            .map(PositiveInteger::new)
            .transpose()?,
        pie_offset: take(&mut a, Ns::Chart, b"pie-offset")
            .map(NonNegativeInteger::new)
            .transpose()?,
        angle_offset: take(&mut a, Ns::Chart, b"angle-offset")
            .map(Angle::new)
            .transpose()?,
        hole_size: take(&mut a, Ns::Chart, b"hole-size")
            .map(Percent::new)
            .transpose()?,
        lines: b!(b"lines"),
        solid_type: e(take(&mut a, Ns::Chart, b"solid-type"), SolidType::parse)?,
        stacked: b!(b"stacked"),
        percentage: b!(b"percentage"),
        treat_empty_cells: e(
            take(&mut a, Ns::Chart, b"treat-empty-cells"),
            EmptyCellTreatment::parse,
        )?,
        link_data_style_to_source: b!(b"link-data-style-to-source"),
        logarithmic: b!(b"logarithmic"),
        maximum: take(&mut a, Ns::Chart, b"maximum")
            .map(Double::new)
            .transpose()?,
        minimum: take(&mut a, Ns::Chart, b"minimum")
            .map(Double::new)
            .transpose()?,
        origin: take(&mut a, Ns::Chart, b"origin")
            .map(Double::new)
            .transpose()?,
        interval_major: take(&mut a, Ns::Chart, b"interval-major")
            .map(Double::new)
            .transpose()?,
        interval_minor_divisor: take(&mut a, Ns::Chart, b"interval-minor-divisor")
            .map(PositiveInteger::new)
            .transpose()?,
        tick_marks_major_inner: b!(b"tick-marks-major-inner"),
        tick_marks_major_outer: b!(b"tick-marks-major-outer"),
        tick_marks_minor_inner: b!(b"tick-marks-minor-inner"),
        tick_marks_minor_outer: b!(b"tick-marks-minor-outer"),
        reverse_direction: b!(b"reverse-direction"),
        display_label: b!(b"display-label"),
        text_overlap: b!(b"text-overlap"),
        line_break: take(&mut a, Ns::Text, b"line-break")
            .map(|v| boolean(&v))
            .transpose()?,
        label_arrangement: e(
            take(&mut a, Ns::Chart, b"label-arrangement"),
            LabelArrangement::parse,
        )?,
        direction: e(take(&mut a, Ns::Style, b"direction"), Direction::parse)?,
        rotation_angle: take(&mut a, Ns::Style, b"rotation-angle")
            .map(Angle::new)
            .transpose()?,
        data_label_number: e(
            take(&mut a, Ns::Chart, b"data-label-number"),
            DataLabelNumber::parse,
        )?,
        data_label_text: b!(b"data-label-text"),
        data_label_symbol: b!(b"data-label-symbol"),
        label_separator: None,
        label_position: e(
            take(&mut a, Ns::Chart, b"label-position"),
            LabelPosition::parse,
        )?,
        label_position_negative: e(
            take(&mut a, Ns::Chart, b"label-position-negative"),
            LabelPosition::parse,
        )?,
        visible: b!(b"visible"),
        auto_position: b!(b"auto-position"),
        auto_size: b!(b"auto-size"),
        mean_value: b!(b"mean-value"),
        error_category: e(
            take(&mut a, Ns::Chart, b"error-category"),
            ErrorCategory::parse,
        )?,
        error_percentage: take(&mut a, Ns::Chart, b"error-percentage")
            .map(Double::new)
            .transpose()?,
        error_margin: take(&mut a, Ns::Chart, b"error-margin")
            .map(Double::new)
            .transpose()?,
        error_lower_limit: take(&mut a, Ns::Chart, b"error-lower-limit")
            .map(Double::new)
            .transpose()?,
        error_upper_limit: take(&mut a, Ns::Chart, b"error-upper-limit")
            .map(Double::new)
            .transpose()?,
        error_upper_indicator: b!(b"error-upper-indicator"),
        error_lower_indicator: b!(b"error-lower-indicator"),
        series_source: e(
            take(&mut a, Ns::Chart, b"series-source"),
            SeriesSource::parse,
        )?,
        regression_type: e(
            take(&mut a, Ns::Chart, b"regression-type"),
            RegressionType::parse,
        )?,
        axis_position: take(&mut a, Ns::Chart, b"axis-position")
            .map(|v| AxisPosition::parse(&v))
            .transpose()?,
        axis_label_position: e(
            take(&mut a, Ns::Chart, b"axis-label-position"),
            AxisLabelPosition::parse,
        )?,
        tick_mark_position: e(
            take(&mut a, Ns::Chart, b"tick-mark-position"),
            TickMarkPosition::parse,
        )?,
        include_hidden_cells: b!(b"include-hidden-cells"),
    };
    if !a.is_empty() {
        return Err(bad(
            "unknown style:chart-properties attribute or wrong namespace",
        ));
    }
    Ok(value)
}
fn symbol_image(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<SymbolImage> {
    let mut a = attrs(reader, version, start)?;
    let href = take(&mut a, Ns::Xlink, b"href")
        .ok_or_else(|| bad("chart:symbol-image requires xlink:href"))?;
    if !a.is_empty() {
        return Err(bad(
            "unknown chart:symbol-image attribute or wrong namespace",
        ));
    }
    SymbolImage::new(href)
}
fn no_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    name: &str,
) -> Result<()> {
    if !attrs(reader, version, start)?.is_empty() {
        return Err(bad(format!("{name} does not allow attributes")));
    }
    Ok(())
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}

struct Active {
    depth: usize,
    style: StyleRecord,
    seen: bool,
    property_depth: Option<usize>,
    symbol_depth: Option<usize>,
    label_depth: Option<usize>,
    paragraph_depth: Option<usize>,
    paragraph_start: Option<usize>,
}
fn push(out: &mut Vec<StyleRecord>, style: StyleRecord, total: &mut usize) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out.iter().any(|value| {
            value.name == style.name && value.is_default_style == style.is_default_style
        })
    {
        return Err(bad("duplicate or excessive chart style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("chart style data is too large"));
    }
    out.push(style);
    Ok(())
}
/// Parse direct chart-family styles in standard style containers.
pub fn parse_chart_style_properties(xml: &str) -> Result<StylePropertiesSet> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut out = Vec::new();
    let mut total = 0;
    let mut events = 0;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("styles XML has too many events"));
        }
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML is too deep"));
                }
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        active = Some(Active {
                            depth,
                            style,
                            seen: false,
                            property_depth: None,
                            symbol_depth: None,
                            label_depth: None,
                            paragraph_depth: None,
                            paragraph_start: None,
                        })
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if value.paragraph_depth.is_some() {
                        continue;
                    }
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"chart-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:chart-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(properties(&reader, version, &start)?);
                        value.property_depth = Some(depth)
                    } else if current.1 == b"chart-properties" {
                        return Err(bad(
                            "style:chart-properties has invalid namespace or parent",
                        ));
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"symbol-image"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .symbol_image
                            .is_some()
                        {
                            return Err(bad("duplicate chart:symbol-image"));
                        }
                        value.style.properties.as_mut().unwrap().symbol_image =
                            Some(symbol_image(&reader, version, &start)?);
                        value.symbol_depth = Some(depth)
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"label-separator"
                    {
                        if value.label_depth.is_some()
                            || value
                                .style
                                .properties
                                .as_ref()
                                .unwrap()
                                .label_separator
                                .is_some()
                        {
                            return Err(bad("duplicate chart:label-separator"));
                        }
                        no_attrs(&reader, version, &start, "chart:label-separator")?;
                        value.label_depth = Some(depth)
                    } else if value.label_depth.is_some_and(|l| depth == l + 1)
                        && current.0 == Ns::Text
                        && current.1 == b"p"
                    {
                        if value.paragraph_start.is_some() {
                            return Err(bad("duplicate text:p in chart:label-separator"));
                        }
                        value.paragraph_depth = Some(depth);
                        value.paragraph_start = Some(begin)
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:chart-properties child"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        push(&mut out, style, &mut total)?
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if value.paragraph_depth.is_some() {
                        continue;
                    }
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"chart-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:chart-properties"));
                        }
                        value.seen = true;
                        let parsed = properties(&reader, version, &start)?;
                        parsed.validate()?;
                        value.style.properties = Some(parsed)
                    } else if current.1 == b"chart-properties" {
                        return Err(bad(
                            "style:chart-properties has invalid namespace or parent",
                        ));
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"symbol-image"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .symbol_image
                            .is_some()
                        {
                            return Err(bad("duplicate chart:symbol-image"));
                        }
                        value.style.properties.as_mut().unwrap().symbol_image =
                            Some(symbol_image(&reader, version, &start)?)
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"label-separator"
                    {
                        return Err(bad("chart:label-separator requires text:p"));
                    } else if value.label_depth.is_some_and(|l| depth == l + 1)
                        && current.0 == Ns::Text
                        && current.1 == b"p"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .label_separator
                            .is_some()
                        {
                            return Err(bad("duplicate text:p in chart:label-separator"));
                        }
                        value.style.properties.as_mut().unwrap().label_separator =
                            Some(LabelSeparator::from_paragraph_xml(&xml[begin..end])?)
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:chart-properties child"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|v| v.property_depth.is_some() && v.paragraph_depth.is_none())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:chart-properties"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|v| v.property_depth.is_some() && v.paragraph_depth.is_none())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:chart-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let depth = stack.len();
                if let Some(value) = active.as_mut() {
                    if value.paragraph_depth == Some(depth) {
                        let begin = value.paragraph_start.take().unwrap();
                        value.style.properties.as_mut().unwrap().label_separator =
                            Some(LabelSeparator::from_paragraph_xml(&xml[begin..end])?);
                        value.paragraph_depth = None
                    }
                    if value.symbol_depth == Some(depth) {
                        value.symbol_depth = None
                    }
                    if value.label_depth == Some(depth) {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .label_separator
                            .is_none()
                        {
                            return Err(bad("chart:label-separator requires text:p"));
                        }
                        value.label_depth = None
                    }
                    if value.property_depth == Some(depth) {
                        value.style.properties.as_ref().unwrap().validate()?;
                        value.property_depth = None
                    }
                }
                if active.as_ref().is_some_and(|value| value.depth == depth) {
                    push(&mut out, active.take().unwrap().style, &mut total)?
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(StylePropertiesSet { styles: out })
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct Target {
    style: Span,
    properties: Option<Span>,
}
fn replace(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw.rfind("/>").ok_or_else(|| bad("invalid empty style"))?;
    Ok(replace(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}
/// Losslessly replace, insert, or remove one existing chart style property element.
pub fn set_chart_style_properties_xml(xml: &str, requested: &StyleRecord) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut target_depth = None;
    let mut active: Option<Target> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|v| {
                    v.0 == Ns::Office && matches!(v.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target chart style"));
                        }
                        target_depth = Some(depth);
                        active = Some(Target {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"chart-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:chart-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|v| {
                    v.0 == Ns::Office && matches!(v.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target chart style"));
                        }
                        found = Some(Target {
                            style: span,
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"chart-properties"
                    && active.as_mut().unwrap().properties.replace(span).is_some()
                {
                    return Err(bad("duplicate style:chart-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.properties.as_ref().is_some_and(|s| s.end == 0)
                        && target_depth.is_some_and(|d| depth == d + 1)
                    {
                        let span = spans.properties.as_mut().unwrap();
                        span.end_start = begin;
                        span.end = end
                    }
                    if target_depth == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        target_depth = None
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target chart style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(StyleProperties::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        return Ok(replace(
            xml,
            properties,
            replacement.as_deref().unwrap_or(""),
        ));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if spans.style.empty {
        return expand(xml, &spans.style, &replacement);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &replacement);
    Ok(out)
}
