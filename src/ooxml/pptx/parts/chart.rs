/// Chart parts for PowerPoint presentations.
///
/// This module provides types for working with charts in PPTX files.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Chart type enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartType {
    /// Bar chart
    Bar,
    /// Column chart (vertical bars)
    Column,
    /// Line chart
    Line,
    /// Pie chart
    Pie,
    /// Area chart
    Area,
    /// Scatter (XY) chart
    Scatter,
    /// Bubble chart
    Bubble,
    /// Doughnut chart
    Doughnut,
    /// Radar chart
    Radar,
    /// Surface chart
    Surface,
    /// Stock chart
    Stock,
    /// Unknown or unsupported chart type
    Unknown,
}

/// Basic chart information extracted from a chart part.
#[derive(Debug, Clone)]
pub struct ChartInfo {
    /// Chart type
    pub chart_type: ChartType,
    /// Chart title if available
    pub title: Option<String>,
    /// Whether the chart has a legend
    pub has_legend: bool,
}

/// Chart part - contains chart data and visualization.
///
/// Corresponds to `/ppt/charts/chartN.xml` in the package.
pub struct ChartPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> ChartPart<'a> {
    /// Create a ChartPart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
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
    /// let chart_part = ChartPart::from_part(part)?;
    /// let info = chart_part.chart_info()?;
    /// println!("Chart type: {:?}", info.chart_type);
    /// ```
    pub fn chart_info(&self) -> Result<ChartInfo> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut chart_type = ChartType::Unknown;
        let mut title: Option<String> = None;
        let mut has_legend = false;

        let mut in_title = false;
        let mut in_title_text = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    // Detect chart type from plot area elements
                    chart_type = match tag_name.as_ref() {
                        b"barChart" => ChartType::Bar,
                        b"bar3DChart" => ChartType::Bar,
                        b"lineChart" => ChartType::Line,
                        b"line3DChart" => ChartType::Line,
                        b"pieChart" => ChartType::Pie,
                        b"pie3DChart" => ChartType::Pie,
                        b"areaChart" => ChartType::Area,
                        b"area3DChart" => ChartType::Area,
                        b"scatterChart" => ChartType::Scatter,
                        b"bubbleChart" => ChartType::Bubble,
                        b"doughnutChart" => ChartType::Doughnut,
                        b"radarChart" => ChartType::Radar,
                        b"surfaceChart" => ChartType::Surface,
                        b"surface3DChart" => ChartType::Surface,
                        b"stockChart" => ChartType::Stock,
                        b"title" => {
                            in_title = true;
                            chart_type
                        },
                        b"legend" => {
                            has_legend = true;
                            chart_type
                        },
                        b"t" if in_title => {
                            in_title_text = true;
                            chart_type
                        },
                        _ => chart_type,
                    };
                },
                Ok(Event::Text(e)) if in_title_text => {
                    let text = std::str::from_utf8(e.as_ref())
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    match &mut title {
                        Some(t) => t.push_str(text),
                        None => title = Some(text.to_string()),
                    }
                },
                Ok(Event::End(e)) => {
                    let tag_name = e.local_name();
                    match tag_name.as_ref() {
                        b"title" => in_title = false,
                        b"t" => in_title_text = false,
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(ChartInfo {
            chart_type,
            title,
            has_legend,
        })
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

// ============================================================================
// Chart Writing Support
// ============================================================================

/// A data series for a chart.
#[derive(Debug, Clone)]
pub struct ChartSeries {
    /// Series name
    pub name: String,
    /// Series values
    pub values: Vec<f64>,
    /// Category names (for the X-axis)
    pub categories: Vec<String>,
}

impl ChartSeries {
    /// Create a new chart series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
            categories: Vec::new(),
        }
    }

    /// Set the series values.
    pub fn with_values(mut self, values: Vec<f64>) -> Self {
        self.values = values;
        self
    }

    /// Set the category names.
    pub fn with_categories(mut self, categories: Vec<String>) -> Self {
        self.categories = categories;
        self
    }
}

/// Chart data for writing charts.
#[derive(Debug, Clone)]
pub struct ChartData {
    /// Chart type
    pub chart_type: ChartType,
    /// Chart title
    pub title: Option<String>,
    /// Data series
    pub series: Vec<ChartSeries>,
    /// Whether to show legend
    pub show_legend: bool,
    /// X position in EMUs
    pub x: i64,
    /// Y position in EMUs
    pub y: i64,
    /// Width in EMUs
    pub width: i64,
    /// Height in EMUs
    pub height: i64,
}

impl ChartData {
    /// Create a new chart data structure.
    pub fn new(chart_type: ChartType, x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            chart_type,
            title: None,
            series: Vec::new(),
            show_legend: true,
            x,
            y,
            width,
            height,
        }
    }

    /// Set the chart title.
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Add a data series.
    pub fn add_series(mut self, series: ChartSeries) -> Self {
        self.series.push(series);
        self
    }

    /// Set whether to show the legend.
    pub fn with_legend(mut self, show: bool) -> Self {
        self.show_legend = show;
        self
    }
}

/// Generate chart XML from ChartData.
pub fn generate_chart_xml(chart: &ChartData) -> String {
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
        ChartType::Bar | ChartType::Column => {
            generate_bar_chart(&mut xml, chart, chart.chart_type == ChartType::Bar);
        },
        ChartType::Line => {
            generate_line_chart(&mut xml, chart);
        },
        ChartType::Pie => {
            generate_pie_chart(&mut xml, chart);
        },
        ChartType::Area => {
            generate_area_chart(&mut xml, chart);
        },
        ChartType::Scatter => {
            generate_scatter_chart(&mut xml, chart);
        },
        ChartType::Doughnut => {
            generate_doughnut_chart(&mut xml, chart);
        },
        _ => {
            // Default to bar chart for unsupported types
            generate_bar_chart(&mut xml, chart, false);
        },
    }

    // Category axis (for non-scatter charts)
    if chart.chart_type != ChartType::Scatter
        && chart.chart_type != ChartType::Pie
        && chart.chart_type != ChartType::Doughnut
    {
        xml.push_str(
            r#"<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>"#,
        );
        xml.push_str(r#"<c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/>"#);
        xml.push_str(r#"<c:minorTickMark val="none"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:catAx>"#);

        // Value axis
        xml.push_str(
            r#"<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>"#,
        );
        xml.push_str(r#"<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>"#);
        xml.push_str(r#"<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:crossAx val="1"/><c:crosses val="autoZero"/></c:valAx>"#);
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

    xml
}

fn generate_bar_chart(xml: &mut String, chart: &ChartData, horizontal: bool) {
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

fn generate_line_chart(xml: &mut String, chart: &ChartData) {
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

fn generate_pie_chart(xml: &mut String, chart: &ChartData) {
    xml.push_str("<c:pieChart>");
    xml.push_str(r#"<c:varyColors val="1"/>"#);

    for (idx, series) in chart.series.iter().enumerate() {
        write_series(xml, series, idx as u32, true);
    }

    xml.push_str(r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#);
    xml.push_str(r#"<c:firstSliceAng val="0"/>"#);
    xml.push_str("</c:pieChart>");
}

fn generate_area_chart(xml: &mut String, chart: &ChartData) {
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

fn generate_scatter_chart(xml: &mut String, chart: &ChartData) {
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

        // X values (use categories or index)
        xml.push_str("<c:xVal><c:numRef><c:f>Sheet1!$A$1</c:f><c:numCache>");
        xml.push_str(&format!(
            r#"<c:formatCode>General</c:formatCode><c:ptCount val="{}"/>"#,
            series.values.len()
        ));
        for (i, _) in series.values.iter().enumerate() {
            let x_val = if i < series.categories.len() {
                series.categories[i].parse::<f64>().unwrap_or(i as f64)
            } else {
                i as f64
            };
            xml.push_str(&format!(r#"<c:pt idx="{}"><c:v>{}</c:v></c:pt>"#, i, x_val));
        }
        xml.push_str("</c:numCache></c:numRef></c:xVal>");

        // Y values
        xml.push_str("<c:yVal><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache>");
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

    // Add X and Y axes for scatter chart
    xml.push_str(
        r#"<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>"#,
    );
    xml.push_str(r#"<c:delete val="0"/><c:axPos val="b"/><c:majorGridlines/>"#);
    xml.push_str(r#"<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:valAx>"#);

    xml.push_str(
        r#"<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>"#,
    );
    xml.push_str(r#"<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>"#);
    xml.push_str(r#"<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:crossAx val="1"/><c:crosses val="autoZero"/></c:valAx>"#);
}

fn generate_doughnut_chart(xml: &mut String, chart: &ChartData) {
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

fn write_series(xml: &mut String, series: &ChartSeries, idx: u32, is_pie: bool) {
    xml.push_str("<c:ser>");
    xml.push_str(&format!(r#"<c:idx val="{}"/>"#, idx));
    xml.push_str(&format!(r#"<c:order val="{}"/>"#, idx));
    xml.push_str(&format!(
        "<c:tx><c:v>{}</c:v></c:tx>",
        escape_xml(&series.name)
    ));

    // Categories (if present)
    if !series.categories.is_empty() {
        xml.push_str("<c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>");
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
    xml.push_str("<c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache>");
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

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Generate graphic frame XML for embedding a chart on a slide.
pub fn generate_chart_graphic_frame(
    shape_id: u32,
    chart_rel_id: &str,
    chart: &ChartData,
) -> String {
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
        chart_rel_id
    ));
    xml.push_str("</a:graphicData>");
    xml.push_str("</a:graphic>");

    xml.push_str("</p:graphicFrame>");

    xml
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_chart_series() {
        let series = ChartSeries::new("Sales")
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
        let chart = ChartData::new(ChartType::Bar, 914400, 914400, 4572000, 2743200)
            .with_title("Quarterly Sales")
            .add_series(ChartSeries::new("2023").with_values(vec![100.0, 200.0]));

        assert_eq!(chart.chart_type, ChartType::Bar);
        assert_eq!(chart.title, Some("Quarterly Sales".to_string()));
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn test_generate_chart_xml() {
        let chart = ChartData::new(ChartType::Bar, 0, 0, 100, 100)
            .with_title("Test Chart")
            .add_series(ChartSeries::new("Data").with_values(vec![1.0, 2.0, 3.0]));

        let xml = generate_chart_xml(&chart);
        assert!(xml.contains("<c:chartSpace"));
        assert!(xml.contains("Test Chart"));
        assert!(xml.contains("<c:barChart>"));
    }
}
