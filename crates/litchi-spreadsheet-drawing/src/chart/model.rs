//! Spreadsheet chart model and host-neutral authoring behavior.
//!
//! This module provides integration between spreadsheet hosts and the
//! comprehensive chart implementation in `litchi_drawingml::chart`.

use super::anchor::Anchor;
use super::relationship::{ExternalDataPart, Relationship, UserShapesPart};
use crate::{Error, Result};
use litchi_drawingml::chart::{
    axis::{Axis, CategoryAxis, ValueAxis},
    data::{DataSourceRef, NumericData, RichText, StringData, TitleText},
    legend::Legend,
    model::Chart as ChartModel,
    plot_area::{
        AreaTypeGroup, BarTypeGroup, LineTypeGroup, PieTypeGroup, PlotArea, ScatterTypeGroup,
        TypeGroup,
    },
    types::{
        AxisPosition, BarDirection, BarGrouping, LegendPosition, ScatterStyle,
        Type as ChartModelType,
    },
};

pub use litchi_drawingml::chart::series::Series;

/// Chart reference for managing charts in a worksheet.
///
/// Contains the chart model and its positioning information.
#[derive(Debug, Clone)]
pub struct Chart {
    /// The chart model containing all chart data and configuration
    pub chart: ChartModel,
    /// Position and size of the chart in the worksheet
    pub anchor: Anchor,
    /// Optional package payload targeted by `chart.external_data`
    pub external_data_part: Option<ExternalDataPart>,
    /// Optional chart user-shapes drawing and its direct related resources
    pub user_shapes_part: Option<UserShapesPart>,
    /// Other direct relationships owned by the chart part
    pub additional_relationships: Vec<Relationship>,
}

impl Chart {
    /// Create a new worksheet chart.
    #[must_use]
    pub fn new(chart: ChartModel, anchor: Anchor) -> Self {
        Self {
            chart,
            anchor,
            external_data_part: None,
            user_shapes_part: None,
            additional_relationships: Vec::new(),
        }
    }

    /// Attach an embedded OOXML workbook as the chart's external data.
    #[must_use]
    pub fn with_embedded_workbook(mut self, data: Vec<u8>) -> Self {
        self.chart.external_data = Some(litchi_drawingml::chart::ExternalData::pending());
        self.external_data_part = Some(ExternalDataPart::embedded_workbook(data));
        self
    }

    /// Attach a package-backed external-data relationship.
    #[must_use]
    pub fn with_external_data_part(
        mut self,
        part: ExternalDataPart,
        auto_update: Option<bool>,
    ) -> Self {
        let mut metadata = litchi_drawingml::chart::ExternalData::pending();
        metadata.auto_update = auto_update;
        self.chart.external_data = Some(metadata);
        self.external_data_part = Some(part);
        self
    }

    /// Attach a chart user-shapes drawing part.
    #[must_use]
    pub fn with_user_shapes_part(mut self, part: UserShapesPart) -> Self {
        self.chart.user_shapes = Some(litchi_drawingml::chart::UserShapes::pending());
        self.user_shapes_part = Some(part);
        self
    }

    /// Add a direct chart-part relationship retained with the worksheet chart.
    #[must_use]
    pub fn with_additional_relationship(mut self, relationship: Relationship) -> Self {
        self.additional_relationships.push(relationship);
        self
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
    /// let chart = Chart::bar_chart(
    ///     "Sales Data",
    ///     "Sheet1!$A$2:$A$10",
    ///     "Sheet1!$B$2:$B$10",
    ///     Anchor::new(1, 1, 7, 14),
    /// );
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn bar_chart(title: &str, categories: &str, values: &str, anchor: Anchor) -> Result<Self> {
        Self::bar_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    /// Create a bar chart with cached data values.
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn bar_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: Anchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(ToString::to_string).collect(),
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
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn line_chart(title: &str, categories: &str, values: &str, anchor: Anchor) -> Result<Self> {
        Self::line_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    /// Create a line chart with cached data values.
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn line_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: Anchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(ToString::to_string).collect(),
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
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn pie_chart(title: &str, categories: &str, values: &str, anchor: Anchor) -> Result<Self> {
        Self::pie_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    /// Create a pie chart with cached data values.
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn pie_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: Anchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(ToString::to_string).collect(),
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
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn area_chart(title: &str, categories: &str, values: &str, anchor: Anchor) -> Result<Self> {
        Self::area_chart_with_cache(title, categories, &[], values, &[], anchor)
    }

    /// Create an area chart with cached data values.
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn area_chart_with_cache(
        title: &str,
        categories: &str,
        cached_categories: &[&str],
        values: &str,
        cached_values: &[f64],
        anchor: Anchor,
    ) -> Result<Self> {
        let mut chart = ChartModel::new();
        chart.title = Some(TitleText::Literal(RichText::new(title)));
        chart.legend = Some(Legend::new(LegendPosition::Right));

        let series = Series::new(0)
            .with_categories(StringData {
                source_ref: Some(DataSourceRef {
                    formula: categories.to_string(),
                }),
                values: cached_categories.iter().map(ToString::to_string).collect(),
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
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn scatter_chart(
        title: &str,
        x_values: &str,
        y_values: &str,
        anchor: Anchor,
    ) -> Result<Self> {
        Self::scatter_chart_with_cache(title, x_values, &[], y_values, &[], anchor)
    }

    /// Create a scatter chart with cached data values.
    ///
    /// # Errors
    ///
    /// Returns an error if the chart cannot be constructed.
    pub fn scatter_chart_with_cache(
        title: &str,
        x_values: &str,
        cached_x_values: &[f64],
        y_values: &str,
        cached_ordinate_values: &[f64],
        anchor: Anchor,
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
            values: cached_ordinate_values.to_vec(),
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
    #[must_use]
    pub fn chart_type(&self) -> ChartModelType {
        if let Some(type_group) = self.chart.plot_area.type_groups.first() {
            match type_group {
                TypeGroup::Area(_) => ChartModelType::Area,
                TypeGroup::Area3D(_) => ChartModelType::Area3D,
                TypeGroup::Bar(_) | TypeGroup::Bar3D(_) => ChartModelType::Bar,
                TypeGroup::Bubble(_) => ChartModelType::Bubble,
                TypeGroup::Doughnut(_) => ChartModelType::Doughnut,
                TypeGroup::Line(_) | TypeGroup::Line3D(_) => ChartModelType::Line,
                TypeGroup::OfPie(_) => ChartModelType::OfPie,
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
    ///
    /// # Errors
    ///
    /// Returns an error when the chart has no type group.
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
                TypeGroup::OfPie(g) => g.common.series.push(series),
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
            Err(Error::Invalid(
                "Chart has no type groups to add series to".to_string(),
            ))
        }
    }

    /// Get the number of series in the chart.
    #[must_use]
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
                TypeGroup::OfPie(g) => g.common.series.len(),
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

pub(super) fn for_each_series_mut(chart: &mut ChartModel, mut apply: impl FnMut(&mut Series)) {
    for type_group in &mut chart.plot_area.type_groups {
        match type_group {
            TypeGroup::Area(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Area3D(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Bar(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Bar3D(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Bubble(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Doughnut(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Line(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Line3D(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::OfPie(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Pie(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Pie3D(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Radar(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Scatter(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Stock(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Surface(g) => g.common.series.iter_mut().for_each(&mut apply),
            TypeGroup::Surface3D(g) => g.common.series.iter_mut().for_each(&mut apply),
        }
    }
}
