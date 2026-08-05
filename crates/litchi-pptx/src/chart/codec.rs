//! Chart parts for PowerPoint presentations.
//!
//! This module provides types for working with charts in PPTX files.

use super::model::{Chart, Info, Series, Type};
use super::package::Part;
use crate::writer::shape::escape_xml;
use crate::{Error, Result};
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_chart_name, is_drawingml_name, unqualified_attribute_value,
};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

impl<'a> Part<'a> {
    fn plot_type(name: &[u8]) -> Option<Type> {
        match name {
            b"barChart" | b"bar3DChart" => Some(Type::Bar),
            b"lineChart" | b"line3DChart" => Some(Type::Line),
            b"pieChart" | b"pie3DChart" | b"ofPieChart" => Some(Type::Pie),
            b"areaChart" | b"area3DChart" => Some(Type::Area),
            b"scatterChart" => Some(Type::Scatter),
            b"bubbleChart" => Some(Type::Bubble),
            b"doughnutChart" => Some(Type::Doughnut),
            b"radarChart" => Some(Type::Radar),
            b"surfaceChart" | b"surface3DChart" => Some(Type::Surface),
            b"stockChart" => Some(Type::Stock),
            _ => None,
        }
    }

    /// Get the XML bytes of the chart.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Parse and return basic chart information.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let chart_part = Part::from_part(part)?;
    /// let info = chart_part.chart_info()?;
    /// println!("Chart type: {:?}", info.chart_type);
    /// ```
    pub fn chart_info(&self) -> Result<Info> {
        let xml = litchi_ooxml_common::mce::process_ooxml(self.xml_bytes())?;
        let mut reader = NsReader::from_reader(xml.as_ref());

        let mut state = ChartScanState::new();

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    state.depth = state.depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("chart XML nesting is too deep".to_string())
                    })?;
                    if state.depth == 1 {
                        if state.root_depth.is_some()
                            || !is_drawingml_chart_name(&namespace, element.name(), b"chartSpace")
                        {
                            return Err(Error::Invalid(
                                "chart XML must have one DrawingML chartSpace root".to_string(),
                            ));
                        }
                        state.root_depth = Some(state.depth);
                    }
                    inspect_chart_start(&namespace, &element, decoder, &mut state)?;
                },
                Event::Empty(element) => {
                    let child_depth = state.depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("chart XML nesting is too deep".to_string())
                    })?;
                    inspect_chart_empty(&namespace, &element, decoder, child_depth, &mut state)?;
                },
                Event::Text(e) if state.title_text_depth.is_some() => {
                    let decoded = e
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    let text = quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    append_title(&mut state.title, &text);
                },
                Event::CData(e) if state.title_text_depth.is_some() => {
                    let text = e
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    append_title(&mut state.title, &text);
                },
                Event::GeneralRef(reference) if state.title_text_depth.is_some() => {
                    let text = decode_xml_reference(&reference)?;
                    append_title(&mut state.title, &text);
                },
                Event::End(element) => {
                    if state.title_text_depth == Some(state.depth)
                        && (is_drawingml_name(&namespace, element.name(), b"t")
                            || is_drawingml_chart_name(&namespace, element.name(), b"v"))
                    {
                        state.title_text_depth = None;
                    }
                    if state.title_depth == Some(state.depth)
                        && is_drawingml_chart_name(&namespace, element.name(), b"title")
                    {
                        state.title_depth = None;
                    }
                    if state.primary_bar_depth == Some(state.depth) {
                        state.primary_bar_depth = None;
                    }
                    if state.plot_area_depth == Some(state.depth)
                        && is_drawingml_chart_name(&namespace, element.name(), b"plotArea")
                    {
                        state.plot_area_depth = None;
                    }
                    if state.chart_depth == Some(state.depth)
                        && is_drawingml_chart_name(&namespace, element.name(), b"chart")
                    {
                        state.chart_depth = None;
                    }
                    if state.root_depth == Some(state.depth)
                        && is_drawingml_chart_name(&namespace, element.name(), b"chartSpace")
                    {
                        state.closed_root = true;
                    }
                    state.depth = state
                        .depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::Invalid("invalid chart XML nesting".to_string()))?;
                },
                Event::Eof
                    if state.depth != 0 || state.root_depth.is_none() || !state.closed_root =>
                {
                    return Err(Error::Invalid(
                        "missing or unterminated DrawingML chartSpace XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !state.seen_chart || !state.seen_plot_area {
            return Err(Error::Invalid(
                "DrawingML chart is missing its chart or plotArea element".to_string(),
            ));
        }

        Ok(Info {
            chart_type: state.chart_type,
            title: state.title,
            has_legend: state.has_legend,
        })
    }
}

struct ChartScanState {
    chart_type: Type,
    title: Option<String>,
    has_legend: bool,
    depth: usize,
    root_depth: Option<usize>,
    closed_root: bool,
    chart_depth: Option<usize>,
    plot_area_depth: Option<usize>,
    title_depth: Option<usize>,
    title_text_depth: Option<usize>,
    primary_bar_depth: Option<usize>,
    seen_chart: bool,
    seen_plot_area: bool,
    seen_title: bool,
    seen_legend: bool,
}

impl ChartScanState {
    fn new() -> Self {
        Self {
            chart_type: Type::Unknown,
            title: None,
            has_legend: false,
            depth: 0,
            root_depth: None,
            closed_root: false,
            chart_depth: None,
            plot_area_depth: None,
            title_depth: None,
            title_text_depth: None,
            primary_bar_depth: None,
            seen_chart: false,
            seen_plot_area: false,
            seen_title: false,
            seen_legend: false,
        }
    }
}

fn inspect_chart_start(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    state: &mut ChartScanState,
) -> Result<()> {
    let depth = state.depth;
    if state.root_depth.is_some_and(|root| depth == root + 1)
        && is_drawingml_chart_name(namespace, element.name(), b"chart")
    {
        if state.seen_chart {
            return Err(Error::Invalid(
                "duplicate DrawingML chart element".to_string(),
            ));
        }
        state.seen_chart = true;
        state.chart_depth = Some(depth);
    } else if state.chart_depth.is_some_and(|chart| depth == chart + 1) {
        if is_drawingml_chart_name(namespace, element.name(), b"title") {
            if state.seen_title {
                return Err(Error::Invalid(
                    "duplicate DrawingML chart title".to_string(),
                ));
            }
            state.seen_title = true;
            state.title_depth = Some(depth);
        } else if is_drawingml_chart_name(namespace, element.name(), b"legend") {
            if state.seen_legend {
                return Err(Error::Invalid(
                    "duplicate DrawingML chart legend".to_string(),
                ));
            }
            state.seen_legend = true;
            state.has_legend = true;
        } else if is_drawingml_chart_name(namespace, element.name(), b"plotArea") {
            if state.seen_plot_area {
                return Err(Error::Invalid("duplicate DrawingML plot area".to_string()));
            }
            state.seen_plot_area = true;
            state.plot_area_depth = Some(depth);
        }
    }

    if state.title_depth.is_some_and(|title| depth > title)
        && (is_drawingml_name(namespace, element.name(), b"t")
            || is_drawingml_chart_name(namespace, element.name(), b"v"))
    {
        if state.title_text_depth.is_some() {
            return Err(Error::Invalid(
                "nested chart title text is invalid".to_string(),
            ));
        }
        state.title_text_depth = Some(depth);
    }

    inspect_plot_element(namespace, element, decoder, depth, state)
}

fn inspect_chart_empty(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    depth: usize,
    state: &mut ChartScanState,
) -> Result<()> {
    if state.root_depth.is_some_and(|root| depth == root + 1)
        && is_drawingml_chart_name(namespace, element.name(), b"chart")
    {
        if state.seen_chart {
            return Err(Error::Invalid(
                "duplicate DrawingML chart element".to_string(),
            ));
        }
        state.seen_chart = true;
    } else if state.chart_depth.is_some_and(|chart| depth == chart + 1) {
        if is_drawingml_chart_name(namespace, element.name(), b"title") {
            if state.seen_title {
                return Err(Error::Invalid(
                    "duplicate DrawingML chart title".to_string(),
                ));
            }
            state.seen_title = true;
        } else if is_drawingml_chart_name(namespace, element.name(), b"legend") {
            if state.seen_legend {
                return Err(Error::Invalid(
                    "duplicate DrawingML chart legend".to_string(),
                ));
            }
            state.seen_legend = true;
            state.has_legend = true;
        } else if is_drawingml_chart_name(namespace, element.name(), b"plotArea") {
            if state.seen_plot_area {
                return Err(Error::Invalid("duplicate DrawingML plot area".to_string()));
            }
            state.seen_plot_area = true;
        }
    }
    inspect_plot_element(namespace, element, decoder, depth, state)?;
    if state.primary_bar_depth == Some(depth) {
        state.primary_bar_depth = None;
    }
    Ok(())
}

fn inspect_plot_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    depth: usize,
    state: &mut ChartScanState,
) -> Result<()> {
    if state
        .plot_area_depth
        .is_some_and(|plot_area| depth == plot_area + 1)
        && is_drawingml_chart_name(namespace, element.name(), element.local_name().as_ref())
        && let Some(detected) = Part::plot_type(element.local_name().as_ref())
        && state.chart_type == Type::Unknown
    {
        state.chart_type = detected;
        if matches!(element.local_name().as_ref(), b"barChart" | b"bar3DChart") {
            state.primary_bar_depth = Some(depth);
        }
    }

    if state
        .primary_bar_depth
        .is_some_and(|bar_chart| depth == bar_chart + 1)
        && is_drawingml_chart_name(namespace, element.name(), b"barDir")
    {
        let direction = unqualified_attribute_value(element, b"val", decoder)?
            .ok_or_else(|| Error::Invalid("barDir is missing its val attribute".to_string()))?;
        state.chart_type = match direction.as_str() {
            "col" => Type::Column,
            "bar" => Type::Bar,
            _ => {
                return Err(Error::Invalid(format!(
                    "invalid barDir value '{direction}'"
                )));
            },
        };
    }
    Ok(())
}

fn append_title(title: &mut Option<String>, text: &str) {
    match title {
        Some(current) => current.push_str(text),
        None => *title = Some(text.to_string()),
    }
}

/// Generate chart XML from Chart.
pub fn encode(chart: &Chart) -> Result<String> {
    validate(chart)?;
    let mut xml = String::with_capacity(8192);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
    );
    xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#);
    xml.push_str(
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );

    xml.push_str("<c:date1904 val=\"0\"/>");
    xml.push_str("<c:lang val=\"en-US\"/>");
    xml.push_str("<c:roundedCorners val=\"0\"/>");

    xml.push_str("<c:chart>");

    // Title
    if let Some(ref title) = chart.title {
        xml.push_str("<c:title>");
        xml.push_str("<c:tx><c:rich>");
        xml.push_str("<a:bodyPr/><a:lstStyle/>");
        xml.push_str("<a:p><a:pPr><a:defRPr/></a:pPr>");
        xml.push_str(&format!(
            "<a:r><a:rPr lang=\"en-US\"/><a:t>{}</a:t></a:r>",
            escape_xml(title)
        ));
        xml.push_str("</a:p></c:rich></c:tx>");
        xml.push_str("<c:overlay val=\"0\"/>");
        xml.push_str("</c:title>");
    }

    xml.push_str("<c:autoTitleDeleted val=\"0\"/>");

    // Plot area
    xml.push_str("<c:plotArea>");
    xml.push_str("<c:layout/>");

    // Chart type-specific content
    match chart.chart_type {
        Type::Bar | Type::Column => {
            generate_bar_chart(&mut xml, chart, chart.chart_type == Type::Bar);
        },
        Type::Line => {
            generate_line_chart(&mut xml, chart);
        },
        Type::Pie => {
            generate_pie_chart(&mut xml, chart);
        },
        Type::Area => {
            generate_area_chart(&mut xml, chart);
        },
        Type::Scatter => {
            generate_scatter_chart(&mut xml, chart);
        },
        Type::Doughnut => {
            generate_doughnut_chart(&mut xml, chart);
        },
        Type::Bubble => {
            generate_bubble_chart(&mut xml, chart);
        },
        Type::Radar => {
            generate_radar_chart(&mut xml, chart);
        },
        Type::Surface => {
            generate_surface_chart(&mut xml, chart);
        },
        Type::Stock => {
            generate_stock_chart(&mut xml, chart);
        },
        Type::Unknown => unreachable!("validated above"),
    }

    match chart.chart_type {
        Type::Pie | Type::Doughnut => {},
        Type::Scatter | Type::Bubble => write_value_axes(&mut xml),
        Type::Surface => write_surface_axes(&mut xml),
        _ => write_category_value_axes(&mut xml),
    }

    xml.push_str("</c:plotArea>");

    // Legend
    if chart.show_legend {
        xml.push_str(r#"<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>"#);
    }

    xml.push_str("<c:plotVisOnly val=\"1\"/>");
    xml.push_str("<c:dispBlanksAs val=\"gap\"/>");
    xml.push_str("</c:chart>");

    xml.push_str("<c:printSettings>");
    xml.push_str("<c:headerFooter/>");
    xml.push_str(r#"<c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>"#);
    xml.push_str("<c:pageSetup/>");
    xml.push_str("</c:printSettings>");

    xml.push_str("</c:chartSpace>");

    Ok(xml)
}

pub(crate) fn validate(chart: &Chart) -> Result<()> {
    if chart.chart_type == Type::Unknown {
        return Err(Error::Invalid(
            "cannot generate XML for an unknown chart type".to_string(),
        ));
    }

    if chart.width <= 0 || chart.height <= 0 {
        return Err(Error::Invalid(
            "chart width and height must be positive".to_string(),
        ));
    }

    for series in &chart.series {
        if series.values.iter().any(|value| !value.is_finite())
            || series.x_values.iter().any(|value| !value.is_finite())
            || series.bubble_sizes.iter().any(|value| !value.is_finite())
        {
            return Err(Error::Invalid(format!(
                "chart series {:?} contains a non-finite number",
                series.name
            )));
        }
    }

    if chart.chart_type == Type::Bubble {
        for series in &chart.series {
            if series.x_values.len() != series.values.len()
                || series.bubble_sizes.len() != series.values.len()
            {
                return Err(Error::Invalid(format!(
                    "bubble series {:?} must have equal X, Y, and bubble-size lengths",
                    series.name
                )));
            }
            if series.bubble_sizes.iter().any(|size| *size < 0.0) {
                return Err(Error::Invalid(format!(
                    "bubble series {:?} contains a negative bubble size",
                    series.name
                )));
            }
        }
    }

    if chart.chart_type == Type::Scatter {
        for series in &chart.series {
            if !series.x_values.is_empty() && series.x_values.len() != series.values.len() {
                return Err(Error::Invalid(format!(
                    "scatter series {:?} must have equal X and Y lengths",
                    series.name
                )));
            }
        }
    }

    if chart.chart_type == Type::Stock && !matches!(chart.series.len(), 3 | 4) {
        return Err(Error::Invalid(
            "stock charts require three (high-low-close) or four (open-high-low-close) series"
                .to_string(),
        ));
    }

    Ok(())
}

fn generate_bar_chart(xml: &mut String, chart: &Chart, horizontal: bool) {
    let chart_tag = "barChart";
    let dir = if horizontal { "bar" } else { "col" };

    xml.push_str(&format!("<c:{}>", chart_tag));
    xml.push_str(&format!(r#"<c:barDir val="{}"/>"#, dir));
    xml.push_str(r#"<c:grouping val="clustered"/>"#);
    xml.push_str(r#"<c:varyColors val="0"/>"#);

    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, false);
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:gapWidth val="150"/>"#);
    xml.push_str(r#"<c:axId val="1"/><c:axId val="2"/>"#);
    xml.push_str(&format!("</c:{}>", chart_tag));
}

fn generate_line_chart(xml: &mut String, chart: &Chart) {
    xml.push_str("<c:lineChart>");
    xml.push_str(r#"<c:grouping val="standard"/>"#);
    xml.push_str(r#"<c:varyColors val="0"/>"#);

    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, false);
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:marker val="1"/>"#);
    xml.push_str(r#"<c:axId val="1"/><c:axId val="2"/>"#);
    xml.push_str("</c:lineChart>");
}

fn generate_pie_chart(xml: &mut String, chart: &Chart) {
    xml.push_str("<c:pieChart>");
    xml.push_str(r#"<c:varyColors val="1"/>"#);

    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, true);
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:firstSliceAng val="0"/>"#);
    xml.push_str("</c:pieChart>");
}

fn generate_area_chart(xml: &mut String, chart: &Chart) {
    xml.push_str("<c:areaChart>");
    xml.push_str(r#"<c:grouping val="standard"/>"#);
    xml.push_str(r#"<c:varyColors val="0"/>"#);

    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, false);
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:axId val="1"/><c:axId val="2"/>"#);
    xml.push_str("</c:areaChart>");
}

fn generate_scatter_chart(xml: &mut String, chart: &Chart) {
    xml.push_str("<c:scatterChart>");
    xml.push_str(r#"<c:scatterStyle val="lineMarker"/>"#);
    xml.push_str(r#"<c:varyColors val="0"/>"#);

    for (idx, series) in chart.series.iter().enumerate() {
        xml.push_str("<c:ser>");
        xml.push_str(&format!(r#"<c:idx val="{}"/>"#, idx));
        xml.push_str(&format!(r#"<c:order val="{}"/>"#, idx));
        xml.push_str(&format!(
            "<c:tx><c:v>{}</c:v></c:tx>",
            escape_xml(&series.name)
        ));

        // X values (prefer explicit values, then numeric categories, then indexes)
        let x_column = idx * 2;
        let y_column = x_column + 1;
        xml.push_str("<c:xVal><c:numRef><c:f>");
        xml.push_str(&worksheet_range(x_column, series.values.len()));
        xml.push_str("</c:f><c:numCache>");
        xml.push_str(&format!(
            r#"<c:formatCode>General</c:formatCode><c:ptCount val="{}"/>"#,
            series.values.len()
        ));
        for (i, _) in series.values.iter().enumerate() {
            let x_val = if let Some(value) = series.x_values.get(i) {
                *value
            } else if i < series.categories.len() {
                series.categories[i].parse::<f64>().unwrap_or(i as f64)
            } else {
                i as f64
            };
            xml.push_str(&format!(r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#, i, x_val));
        }
        xml.push_str("</c:numCache></c:numRef></c:xVal>");

        // Y values
        xml.push_str("<c:yVal><c:numRef><c:f>");
        xml.push_str(&worksheet_range(y_column, series.values.len()));
        xml.push_str("</c:f><c:numCache>");
        xml.push_str(&format!(
            r#"<c:formatCode>General</c:formatCode><c:ptCount val="{}"/>"#,
            series.values.len()
        ));
        for (i, val) in series.values.iter().enumerate() {
            xml.push_str(&format!(r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#, i, val));
        }
        xml.push_str("</c:numCache></c:numRef></c:yVal>");

        xml.push_str("</c:ser>");
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:axId val="1"/><c:axId val="2"/>"#);
    xml.push_str("</c:scatterChart>");
}

fn generate_doughnut_chart(xml: &mut String, chart: &Chart) {
    xml.push_str("<c:doughnutChart>");
    xml.push_str(r#"<c:varyColors val="1"/>"#);

    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, true);
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:firstSliceAng val="0"/>"#);
    xml.push_str(r#"<c:holeSize val="50"/>"#);
    xml.push_str("</c:doughnutChart>");
}

fn generate_bubble_chart(xml: &mut String, chart: &Chart) {
    xml.push_str("<c:bubbleChart><c:varyColors val=\"0\"/>");

    for (idx, series) in chart.series.iter().enumerate() {
        let x_column = idx * 3;
        write_xy_size_series(xml, series, idx as u32, x_column);
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:bubble3D val="0"/><c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/>"#);
    xml.push_str(r#"<c:axId val="1"/><c:axId val="2"/></c:bubbleChart>"#);
}

fn generate_radar_chart(xml: &mut String, chart: &Chart) {
    xml.push_str(r#"<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>"#);
    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, false);
    }
    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:axId val="1"/><c:axId val="2"/></c:radarChart>"#);
}

fn generate_stock_chart(xml: &mut String, chart: &Chart) {
    xml.push_str("<c:stockChart>");
    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, false);
    }
    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str("<c:hiLowLines/>");
    if chart.series.len() == 4 {
        xml.push_str(
            r#"<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>"#,
        );
    }
    xml.push_str(r#"<c:axId val="1"/><c:axId val="2"/></c:stockChart>"#);
}

fn generate_surface_chart(xml: &mut String, chart: &Chart) {
    xml.push_str(r#"<c:surfaceChart><c:wireframe val="0"/>"#);
    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, false);
    }
    xml.push_str(
        r#"<c:bandFmts/><c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surfaceChart>"#,
    );
}

fn write_xy_size_series(xml: &mut String, series: &Series, idx: u32, x_column: usize) {
    xml.push_str("<c:ser>");
    xml.push_str(&format!(
        r#"<c:idx val="{}"/><c:order val="{}"/>"#,
        idx, idx
    ));
    xml.push_str(&format!(
        "<c:tx><c:v>{}</c:v></c:tx>",
        escape_xml(&series.name)
    ));
    write_numeric_reference(
        xml,
        "xVal",
        &worksheet_range(x_column, series.x_values.len()),
        &series.x_values,
    );
    write_numeric_reference(
        xml,
        "yVal",
        &worksheet_range(x_column + 1, series.values.len()),
        &series.values,
    );
    write_numeric_reference(
        xml,
        "bubbleSize",
        &worksheet_range(x_column + 2, series.bubble_sizes.len()),
        &series.bubble_sizes,
    );
    xml.push_str(r#"<c:bubble3D val="0"/></c:ser>"#);
}

fn write_numeric_reference(xml: &mut String, tag: &str, formula: &str, values: &[f64]) {
    xml.push_str(&format!(
        "<c:{tag}><c:numRef><c:f>{formula}</c:f><c:numCache>"
    ));
    xml.push_str(&format!(
        r#"<c:formatCode>General</c:formatCode><c:ptCount val="{}"/>"#,
        values.len()
    ));
    for (index, value) in values.iter().enumerate() {
        xml.push_str(&format!(
            r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#,
            index, value
        ));
    }
    xml.push_str(&format!("</c:numCache></c:numRef></c:{tag}>"));
}

fn write_category_value_axes(xml: &mut String) {
    xml.push_str(
        r#"<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>"#,
    );
    xml.push_str(r#"<c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/>"#);
    xml.push_str(
        r#"<c:minorTickMark val="none"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:catAx>"#,
    );
    write_value_axis(xml, 2, "l", 1);
}

fn write_value_axes(xml: &mut String) {
    write_value_axis(xml, 1, "b", 2);
    write_value_axis(xml, 2, "l", 1);
}

fn write_surface_axes(xml: &mut String) {
    write_category_value_axes(xml);
    xml.push_str(
        r#"<c:serAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling>"#,
    );
    xml.push_str(r#"<c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:serAx>"#);
}

fn write_value_axis(xml: &mut String, id: u32, position: &str, cross_axis: u32) {
    xml.push_str(&format!(
        r#"<c:valAx><c:axId val="{}"/><c:scaling><c:orientation val="minMax"/></c:scaling>"#,
        id
    ));
    xml.push_str(&format!(
        r#"<c:delete val="0"/><c:axPos val="{}"/><c:majorGridlines/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:crossAx val="{}"/><c:crosses val="autoZero"/></c:valAx>"#,
        position, cross_axis
    ));
}

fn worksheet_range(column: usize, len: usize) -> String {
    let column = worksheet_column(column);
    let end_row = len.saturating_add(1).max(2);
    format!("Sheet1!${column}$2:${column}${end_row}")
}

fn worksheet_column(mut index: usize) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (index % 26) as u8) as char);
        if index < 26 {
            return result;
        }
        index = index / 26 - 1;
    }
}

fn write_series(xml: &mut String, series: &Series, idx: u32, is_pie: bool) {
    xml.push_str("<c:ser>");
    xml.push_str(&format!(r#"<c:idx val="{}"/>"#, idx));
    xml.push_str(&format!(r#"<c:order val="{}"/>"#, idx));
    xml.push_str(&format!(
        "<c:tx><c:v>{}</c:v></c:tx>",
        escape_xml(&series.name)
    ));

    // Categories (if present)
    if !series.categories.is_empty() {
        xml.push_str("<c:cat><c:strRef><c:f>");
        xml.push_str(&worksheet_range(0, series.categories.len()));
        xml.push_str("</c:f><c:strCache>");
        xml.push_str(&format!(
            r#"<c:ptCount val="{}"/>"#,
            series.categories.len()
        ));
        for (i, cat) in series.categories.iter().enumerate() {
            xml.push_str(&format!(
                r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#,
                i,
                escape_xml(cat)
            ));
        }
        xml.push_str("</c:strCache></c:strRef></c:cat>");
    }

    // Values
    xml.push_str("<c:val><c:numRef><c:f>");
    xml.push_str(&worksheet_range(idx as usize + 1, series.values.len()));
    xml.push_str("</c:f><c:numCache>");
    xml.push_str(&format!(
        r#"<c:formatCode>General</c:formatCode><c:ptCount val="{}"/>"#,
        series.values.len()
    ));
    for (i, val) in series.values.iter().enumerate() {
        xml.push_str(&format!(r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#, i, val));
    }
    xml.push_str("</c:numCache></c:numRef></c:val>");

    // Explosion for pie charts
    if is_pie {
        xml.push_str(r#"<c:explosion val="0"/>"#);
    }

    xml.push_str("</c:ser>");
}

/// Generate graphic frame XML for embedding a chart on a slide.
pub fn write_graphic_frame(shape_id: u32, chart_rel_id: &str, chart: &Chart) -> String {
    let mut xml = String::with_capacity(1024);

    xml.push_str("<p:graphicFrame>");
    xml.push_str("<p:nvGraphicFramePr>");
    xml.push_str(&format!(
        r#"<p:cNvPr id="{}" name="Chart {}"/>"#,
        shape_id, shape_id
    ));
    xml.push_str(r#"<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>"#);
    xml.push_str("<p:nvPr/>");
    xml.push_str("</p:nvGraphicFramePr>");

    xml.push_str("<p:xfrm>");
    xml.push_str(&format!(r#"<a:off x="{}" y="{}"/>"#, chart.x, chart.y));
    xml.push_str(&format!(
        r#"<a:ext cx="{}" cy="{}"/>"#,
        chart.width, chart.height
    ));
    xml.push_str("</p:xfrm>");

    xml.push_str("<a:graphic>");
    xml.push_str(r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#);
    xml.push_str(&format!(
        r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{}"/>"#,
        escape_xml(chart_rel_id)
    ));
    xml.push_str("</a:graphicData>");
    xml.push_str("</a:graphic>");

    xml.push_str("</p:graphicFrame>");

    xml
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{packuri::PackURI, part::BlobPart};

    const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    fn parse_chart_info(xml: &str) -> Info {
        let part = BlobPart::new(
            PackURI::new("/ppt/charts/chart1.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml".to_string(),
            xml.as_bytes().to_vec(),
        );
        Part::from_part(&part).unwrap().chart_info().unwrap()
    }

    #[test]
    fn chart_info_distinguishes_column_and_decodes_title() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue &amp; <![CDATA[Growth < 2027]]></a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/></c:barChart></c:plotArea><c:legend/></c:chart></c:chartSpace>"#
        );
        let info = parse_chart_info(&xml);

        assert_eq!(info.chart_type, Type::Column);
        assert_eq!(info.title.as_deref(), Some("Revenue & Growth < 2027"));
        assert!(info.has_legend);
    }

    #[test]
    fn chart_info_uses_first_plot_in_combination_chart() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:lineChart/><c:barChart><c:barDir val="col"/></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let info = parse_chart_info(&xml);

        assert_eq!(info.chart_type, Type::Line);
    }

    #[test]
    fn empty_title_does_not_capture_later_text() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:title/><c:plotArea><c:barChart><c:barDir val="bar"/></c:barChart><a:t>not a title</a:t></c:plotArea></c:chart></c:chartSpace>"#
        );
        let info = parse_chart_info(&xml);

        assert_eq!(info.chart_type, Type::Bar);
        assert!(info.title.is_none());
    }

    #[test]
    fn strict_cached_titles_and_foreign_lookalikes_are_supported() {
        let xml = r#"<x:chartSpace xmlns:x="http://purl.oclc.org/ooxml/drawingml/chart"
            xmlns:f="urn:foreign"><x:chart><x:title><x:tx><x:strRef><x:strCache>
            <x:pt idx="0"><x:v>Cached &amp; Strict</x:v></x:pt></x:strCache></x:strRef></x:tx></x:title>
            <x:plotArea><f:pieChart/><x:doughnutChart/></x:plotArea><f:legend/></x:chart></x:chartSpace>"#;
        let info = parse_chart_info(xml);
        assert_eq!(info.chart_type, Type::Doughnut);
        assert_eq!(info.title.as_deref(), Some("Cached & Strict"));
        assert!(!info.has_legend);
    }

    #[test]
    fn malformed_chart_info_is_rejected() {
        let invalid = [
            format!(
                r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:barChart><c:barDir val="bad"/></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
            ),
            format!(r#"<c:chartSpace xmlns:c="{C}"><c:chart>"#),
            "<chartSpace/>".to_string(),
        ];
        for xml in invalid {
            let part = BlobPart::new(
                PackURI::new("/ppt/charts/chart1.xml").unwrap(),
                litchi_opc::constants::content_type::DML_CHART.to_string(),
                xml.into_bytes(),
            );
            assert!(Part::from_part(&part).unwrap().chart_info().is_err());
        }
    }

    #[test]
    fn test_chart_series() {
        let series = Series::new("Sales")
            .with_categories(vec![
                "Q1".to_string(),
                "Q2".to_string(),
                "Q3".to_string(),
                "Q4".to_string(),
            ])
            .with_values(vec![100.0, 150.0, 200.0, 175.0]);

        assert_eq!(series.name, "Sales");
        assert_eq!(series.values.len(), 4);
        assert_eq!(series.categories.len(), 4);
    }

    #[test]
    fn test_chart_data() {
        let chart = Chart::new(Type::Bar, 914400, 914400, 4572000, 2743200)
            .with_title("Quarterly Sales")
            .add_series(Series::new("2023").with_values(vec![100.0, 200.0]));

        assert_eq!(chart.chart_type, Type::Bar);
        assert_eq!(chart.title, Some("Quarterly Sales".to_string()));
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn test_encode() {
        let chart = Chart::new(Type::Bar, 0, 0, 100, 100)
            .with_title("Test Chart")
            .add_series(Series::new("Data").with_values(vec![1.0, 2.0, 3.0]));

        let xml = encode(&chart).unwrap();
        assert!(xml.contains("<c:chartSpace"));
        assert!(xml.contains("Test Chart"));
        assert!(xml.contains("<c:barChart>"));
    }

    fn chart_with_series(chart_type: Type, series_count: usize) -> Chart {
        (0..series_count).fold(Chart::new(chart_type, 0, 0, 100, 100), |chart, index| {
            chart.add_series(
                Series::new(format!("Series {index}"))
                    .with_categories(vec!["A".to_string(), "B".to_string()])
                    .with_values(vec![index as f64 + 1.0, index as f64 + 2.0]),
            )
        })
    }

    #[test]
    fn declared_chart_types_are_not_downgraded_to_columns() {
        for (chart_type, expected_tag, series_count) in [
            (Type::Radar, "radarChart", 2),
            (Type::Surface, "surfaceChart", 2),
            (Type::Stock, "stockChart", 3),
        ] {
            let xml = encode(&chart_with_series(chart_type, series_count)).unwrap();
            assert!(xml.contains(&format!("<c:{expected_tag}>")));
            assert!(!xml.contains("<c:barChart>"));
            assert_eq!(parse_chart_info(&xml).chart_type, chart_type);
        }
    }

    #[test]
    fn bubble_chart_writes_all_three_numeric_dimensions() {
        let chart = Chart::new(Type::Bubble, 0, 0, 100, 100).add_series(
            Series::new("Reach")
                .with_x_values(vec![1.0, 2.0])
                .with_values(vec![10.0, 20.0])
                .with_bubble_sizes(vec![5.0, 8.0]),
        );
        let xml = encode(&chart).unwrap();

        assert!(xml.contains("<c:bubbleChart>"));
        assert!(xml.contains("<c:xVal>"));
        assert!(xml.contains("<c:yVal>"));
        assert!(xml.contains("<c:bubbleSize>"));
        assert!(xml.contains("Sheet1!$A$2:$A$3"));
        assert!(xml.contains("Sheet1!$B$2:$B$3"));
        assert!(xml.contains("Sheet1!$C$2:$C$3"));
        assert_eq!(parse_chart_info(&xml).chart_type, Type::Bubble);
    }

    #[test]
    fn surface_chart_writes_its_series_axis() {
        let xml = encode(&chart_with_series(Type::Surface, 2)).unwrap();
        assert!(xml.contains("<c:serAx>"));
        assert!(xml.contains(r#"<c:axId val="3"/>"#));
    }

    #[test]
    fn chart_formulas_reference_real_workbook_ranges() {
        let chart = chart_with_series(Type::Column, 2);
        let xml = encode(&chart).unwrap();

        assert!(xml.contains("Sheet1!$A$2:$A$3"));
        assert!(xml.contains("Sheet1!$B$2:$B$3"));
        assert!(xml.contains("Sheet1!$C$2:$C$3"));
        assert!(!xml.contains("Sheet1!$A$1"));
    }

    #[test]
    fn doughnut_xml_is_balanced_and_relationship_ids_are_escaped() {
        let chart = chart_with_series(Type::Doughnut, 1);
        let xml = encode(&chart).unwrap();
        assert_eq!(xml.matches("<c:doughnutChart>").count(), 1);
        assert_eq!(xml.matches("</c:doughnutChart>").count(), 1);
        assert_eq!(parse_chart_info(&xml).chart_type, Type::Doughnut);

        let frame = write_graphic_frame(1, "rId1&spoof", &chart);
        assert!(frame.contains(r#"r:id="rId1&amp;spoof""#));
    }

    #[test]
    fn invalid_specialized_chart_data_is_rejected() {
        let unknown = Chart::new(Type::Unknown, 0, 0, 100, 100);
        assert!(encode(&unknown).is_err());

        let bad_bubble = Chart::new(Type::Bubble, 0, 0, 100, 100).add_series(
            Series::new("Bad")
                .with_x_values(vec![1.0])
                .with_values(vec![2.0, 3.0])
                .with_bubble_sizes(vec![4.0, 5.0]),
        );
        assert!(encode(&bad_bubble).is_err());

        let bad_stock = chart_with_series(Type::Stock, 2);
        assert!(encode(&bad_stock).is_err());

        let bad_scatter = Chart::new(Type::Scatter, 0, 0, 100, 100).add_series(
            Series::new("Bad")
                .with_x_values(vec![1.0])
                .with_values(vec![2.0, 3.0]),
        );
        assert!(encode(&bad_scatter).is_err());

        let non_finite = Chart::new(Type::Line, 0, 0, 100, 100)
            .add_series(Series::new("Bad").with_values(vec![f64::NAN]));
        assert!(encode(&non_finite).is_err());

        let invalid_geometry = Chart::new(Type::Line, 0, 0, 0, 100);
        assert!(encode(&invalid_geometry).is_err());
    }
}
