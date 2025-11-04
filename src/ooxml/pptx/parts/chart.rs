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

        let mut buf = Vec::new();
        let mut in_title = false;
        let mut in_title_text = false;

        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
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
