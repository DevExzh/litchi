//! Excel chart integration module.
//!
//! This module provides integration between the XLSX worksheet API and the
//! comprehensive chart implementation in `crate::ooxml::charts`.

use crate::ooxml::charts::{
    axis::{Axis, CategoryAxis, ValueAxis},
    chart::Chart as ChartModel,
    legend::Legend,
    models::{DataSourceRef, NumericData, RichText, StringData, TitleText},
    plot_area::{
        AreaTypeGroup, BarTypeGroup, LineTypeGroup, PieTypeGroup, PlotArea, ScatterTypeGroup,
        TypeGroup,
    },
    series::Series,
    types::{
        AxisPosition, BarDirection, BarGrouping, ChartType as ChartModelType, LegendPosition,
        ScatterStyle,
    },
};
use crate::ooxml::error::{OoxmlError, Result};

/// Chart anchor position in a worksheet.
///
/// Specifies the position and size of a chart using cell anchors and offsets.
#[derive(Debug, Clone)]
pub struct ChartAnchor {
    /// Starting column (0-based)
    pub from_col: u32,
    /// Offset from the left edge of from_col (in EMUs)
    pub from_col_offset: i64,
    /// Starting row (0-based)
    pub from_row: u32,
    /// Offset from the top edge of from_row (in EMUs)
    pub from_row_offset: i64,
    /// Ending column (0-based)
    pub to_col: u32,
    /// Offset from the left edge of to_col (in EMUs)
    pub to_col_offset: i64,
    /// Ending row (0-based)
    pub to_row: u32,
    /// Offset from the top edge of to_row (in EMUs)
    pub to_row_offset: i64,
}

impl ChartAnchor {
    /// Create a new chart anchor from cell positions.
    ///
    /// # Arguments
    ///
    /// * `from_col` - Starting column (0-based)
    /// * `from_row` - Starting row (0-based)
    /// * `to_col` - Ending column (0-based)
    /// * `to_row` - Ending row (0-based)
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// // Chart spanning from B2 to H15
    /// let anchor = ChartAnchor::new(1, 1, 7, 14);
    /// ```
    pub fn new(from_col: u32, from_row: u32, to_col: u32, to_row: u32) -> Self {
        Self {
            from_col,
            from_col_offset: 0,
            from_row,
            from_row_offset: 0,
            to_col,
            to_col_offset: 0,
            to_row,
            to_row_offset: 0,
        }
    }

    /// Create a chart anchor with precise offsets.
    #[allow(clippy::too_many_arguments)]
    pub fn with_offsets(
        from_col: u32,
        from_col_offset: i64,
        from_row: u32,
        from_row_offset: i64,
        to_col: u32,
        to_col_offset: i64,
        to_row: u32,
        to_row_offset: i64,
    ) -> Self {
        Self {
            from_col,
            from_col_offset,
            from_row,
            from_row_offset,
            to_col,
            to_col_offset,
            to_row,
            to_row_offset,
        }
    }
}

impl Default for ChartAnchor {
    fn default() -> Self {
        Self::new(0, 0, 10, 15)
    }
}

/// Chart reference for managing charts in a worksheet.
///
/// Contains the chart model and its positioning information.
#[derive(Debug, Clone)]
pub struct WorksheetChart {
    /// The chart model containing all chart data and configuration
    pub chart: ChartModel,
    /// Position and size of the chart in the worksheet
    pub anchor: ChartAnchor,
}

impl WorksheetChart {
    /// Create a new worksheet chart.
    pub fn new(chart: ChartModel, anchor: ChartAnchor) -> Self {
        Self { chart, anchor }
    }

    /// Create a simple bar chart from data ranges.
    ///
    /// # Arguments
    ///
    /// * `title` - Chart title
    /// * `categories` - Range for category labels (e.g., "Sheet1!$A$2:$A$10")
    /// * `values` - Range for data values (e.g., "Sheet1!$B$2:$B$10")
    /// * `anchor` - Chart position
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let chart = WorksheetChart::bar_chart(
    ///     "Sales Data",
    ///     "Sheet1!$A$2:$A$10",
    ///     "Sheet1!$B$2:$B$10",
    ///     ChartAnchor::new(1, 1, 7, 14),
    /// );
    /// ```
    pub fn bar_chart(
        title: &str,
        categories: &str,
        values: &str,
        anchor: ChartAnchor,
    ) -> Result<Self> {
        Self::bar_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    /// Create a bar chart with cached data values.
    pub fn bar_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: ChartAnchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(|s| s.to_string()).collect(),
            })
            .with_values(NumericData {
                source_ref: Some(DataSourceRef {
                    formula: values.to_string(),
                }),
                values: cached_values.to_vec(),
                format_code: None,
            });

        let mut bar_group = BarTypeGroup::new(BarDirection::Column, BarGrouping::Clustered);
        bar_group.common.series.push(series);

        let cat_axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
        let val_axis = ValueAxis::new(2, AxisPosition::Left, 1);

        chart.plot_area = PlotArea::new()
            .add_type_group(TypeGroup::Bar(bar_group))
            .add_axis(Axis::Category(cat_axis))
            .add_axis(Axis::Value(val_axis));

        Ok(Self::new(chart, anchor))
    }

    /// Create a simple line chart from data ranges.
    pub fn line_chart(
        title: &str,
        categories: &str,
        values: &str,
        anchor: ChartAnchor,
    ) -> Result<Self> {
        Self::line_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    /// Create a line chart with cached data values.
    pub fn line_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: ChartAnchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(|s| s.to_string()).collect(),
            })
            .with_values(NumericData {
                source_ref: Some(DataSourceRef {
                    formula: values.to_string(),
                }),
                values: cached_values.to_vec(),
                format_code: None,
            });

        let mut line_group = LineTypeGroup::new(BarGrouping::Standard);
        line_group.common.series.push(series);

        let cat_axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
        let val_axis = ValueAxis::new(2, AxisPosition::Left, 1);

        chart.plot_area = PlotArea::new()
            .add_type_group(TypeGroup::Line(line_group))
            .add_axis(Axis::Category(cat_axis))
            .add_axis(Axis::Value(val_axis));

        Ok(Self::new(chart, anchor))
    }

    /// Create a simple pie chart from data ranges.
    pub fn pie_chart(
        title: &str,
        categories: &str,
        values: &str,
        anchor: ChartAnchor,
    ) -> Result<Self> {
        Self::pie_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    pub fn pie_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: ChartAnchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(|s| s.to_string()).collect(),
            })
            .with_values(NumericData {
                source_ref: Some(DataSourceRef {
                    formula: values.to_string(),
                }),
                values: cached_values.to_vec(),
                format_code: None,
            });

        let mut pie_group = PieTypeGroup::new();
        pie_group.common.series.push(series);

        chart.plot_area = PlotArea::new().add_type_group(TypeGroup::Pie(pie_group));

        Ok(Self::new(chart, anchor))
    }

    /// Create a simple area chart from data ranges.
    pub fn area_chart(
        title: &str,
        categories: &str,
        values: &str,
        anchor: ChartAnchor,
    ) -> Result<Self> {
        Self::area_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    /// Create an area chart with cached data values.
    pub fn area_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: ChartAnchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(|s| s.to_string()).collect(),
            })
            .with_values(NumericData {
                source_ref: Some(DataSourceRef {
                    formula: values.to_string(),
                }),
                values: cached_values.to_vec(),
                format_code: None,
            });

        let mut area_group = AreaTypeGroup::new(BarGrouping::Standard);
        area_group.common.series.push(series);

        let cat_axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
        let val_axis = ValueAxis::new(2, AxisPosition::Left, 1);

        chart.plot_area = PlotArea::new()
            .add_type_group(TypeGroup::Area(area_group))
            .add_axis(Axis::Category(cat_axis))
            .add_axis(Axis::Value(val_axis));

        Ok(Self::new(chart, anchor))
    }

    /// Create a simple scatter (XY) chart from data ranges.
    pub fn scatter_chart(
        title: &str,
        x_values: &str,
        y_values: &str,
        anchor: ChartAnchor,
    ) -> Result<Self> {
        Self::scatter_chart_with_cache(title, x_values, &[], y_values, &[], anchor)
    }

    pub fn scatter_chart_with_cache(
        title: &str,
        x_values: &str,
        cached_x_values: &[f64],
        y_values: &str,
        cached_y_values: &[f64],
        anchor: ChartAnchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let mut series = Series::new(0);
        series.x_values = Some(NumericData {
            source_ref: Some(DataSourceRef {
                formula: x_values.to_string(),
            }),
            values: cached_x_values.to_vec(),
            format_code: None,
        });
        series.y_values = Some(NumericData {
            source_ref: Some(DataSourceRef {
                formula: y_values.to_string(),
            }),
            values: cached_y_values.to_vec(),
            format_code: None,
        });

        let mut scatter_group = ScatterTypeGroup::new(ScatterStyle::LineMarker);
        scatter_group.common.series.push(series);

        let cat_axis = ValueAxis::new(1, AxisPosition::Bottom, 2);
        let val_axis = ValueAxis::new(2, AxisPosition::Left, 1);

        chart.plot_area = PlotArea::new()
            .add_type_group(TypeGroup::Scatter(scatter_group))
            .add_axis(Axis::Value(cat_axis))
            .add_axis(Axis::Value(val_axis));

        Ok(Self::new(chart, anchor))
    }

    /// Get the chart type.
    pub fn chart_type(&self) -> ChartModelType {
        if let Some(type_group) = self.chart.plot_area.type_groups.first() {
            match type_group {
                TypeGroup::Area(_) => ChartModelType::Area,
                TypeGroup::Area3D(_) => ChartModelType::Area3D,
                TypeGroup::Bar(_) | TypeGroup::Bar3D(_) => ChartModelType::Bar,
                TypeGroup::Bubble(_) => ChartModelType::Bubble,
                TypeGroup::Doughnut(_) => ChartModelType::Doughnut,
                TypeGroup::Line(_) | TypeGroup::Line3D(_) => ChartModelType::Line,
                TypeGroup::Pie(_) | TypeGroup::Pie3D(_) => ChartModelType::Pie,
                TypeGroup::Radar(_) => ChartModelType::Radar,
                TypeGroup::Scatter(_) => ChartModelType::Scatter,
                TypeGroup::Stock(_) => ChartModelType::Stock,
                TypeGroup::Surface(_) | TypeGroup::Surface3D(_) => ChartModelType::Surface,
            }
        } else {
            ChartModelType::Unknown
        }
    }

    /// Add a series to the chart.
    ///
    /// # Note
    ///
    /// This adds the series to the first type group in the plot area.
    pub fn add_series(&mut self, series: Series) -> Result<()> {
        if let Some(type_group) = self.chart.plot_area.type_groups.first_mut() {
            match type_group {
                TypeGroup::Area(g) => g.common.series.push(series),
                TypeGroup::Area3D(g) => g.common.series.push(series),
                TypeGroup::Bar(g) => g.common.series.push(series),
                TypeGroup::Bar3D(g) => g.common.series.push(series),
                TypeGroup::Bubble(g) => g.common.series.push(series),
                TypeGroup::Doughnut(g) => g.common.series.push(series),
                TypeGroup::Line(g) => g.common.series.push(series),
                TypeGroup::Line3D(g) => g.common.series.push(series),
                TypeGroup::Pie(g) => g.common.series.push(series),
                TypeGroup::Pie3D(g) => g.common.series.push(series),
                TypeGroup::Radar(g) => g.common.series.push(series),
                TypeGroup::Scatter(g) => g.common.series.push(series),
                TypeGroup::Stock(g) => g.common.series.push(series),
                TypeGroup::Surface(g) => g.common.series.push(series),
                TypeGroup::Surface3D(g) => g.common.series.push(series),
            }
            Ok(())
        } else {
            Err(OoxmlError::Xml(
                "Chart has no type groups to add series to".to_string(),
            ))
        }
    }

    /// Get the number of series in the chart.
    pub fn series_count(&self) -> usize {
        if let Some(type_group) = self.chart.plot_area.type_groups.first() {
            match type_group {
                TypeGroup::Area(g) => g.common.series.len(),
                TypeGroup::Area3D(g) => g.common.series.len(),
                TypeGroup::Bar(g) => g.common.series.len(),
                TypeGroup::Bar3D(g) => g.common.series.len(),
                TypeGroup::Bubble(g) => g.common.series.len(),
                TypeGroup::Doughnut(g) => g.common.series.len(),
                TypeGroup::Line(g) => g.common.series.len(),
                TypeGroup::Line3D(g) => g.common.series.len(),
                TypeGroup::Pie(g) => g.common.series.len(),
                TypeGroup::Pie3D(g) => g.common.series.len(),
                TypeGroup::Radar(g) => g.common.series.len(),
                TypeGroup::Scatter(g) => g.common.series.len(),
                TypeGroup::Stock(g) => g.common.series.len(),
                TypeGroup::Surface(g) => g.common.series.len(),
                TypeGroup::Surface3D(g) => g.common.series.len(),
            }
        } else {
            0
        }
    }
}

/// Parse a chart from chart XML and drawing anchor.
pub fn parse_chart_from_xml(chart_xml: &[u8], anchor: ChartAnchor) -> Result<WorksheetChart> {
    let chart = crate::ooxml::charts::reader::parse_chart(chart_xml)?;
    Ok(WorksheetChart::new(chart, anchor))
}

/// Generate chart XML for a worksheet chart.
pub fn generate_chart_xml(chart: &ChartModel) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    crate::ooxml::charts::writer::write_chart(&mut output, chart)
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
    Ok(output)
}
