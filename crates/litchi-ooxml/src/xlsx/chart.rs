//! Excel chart integration module.
//!
//! This module provides integration between the XLSX worksheet API and the
//! comprehensive chart implementation in `crate::charts`.

use crate::charts::{
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
use crate::error::{OoxmlError, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const WORKSHEET_ROW_COUNT: u32 = 1_048_576;
const WORKSHEET_COLUMN_COUNT: u32 = 16_384;

/// Storage target for a chart's external-data relationship.
#[derive(Debug, Clone)]
pub enum ChartExternalDataTarget {
    /// A part embedded in the containing OOXML package
    Embedded {
        /// Complete bytes of the embedded object
        data: Vec<u8>,
        /// OPC content type for the embedded part
        content_type: String,
        /// Filename extension without a leading dot
        extension: String,
    },
    /// An externally linked object
    Linked {
        /// External relationship target
        target: String,
    },
}

/// Package payload and relationship type for chart external data.
#[derive(Debug, Clone)]
pub struct ChartExternalDataPart {
    /// Relationship type, normally package or OLE object
    pub relationship_type: String,
    /// Embedded or linked relationship target
    pub target: ChartExternalDataTarget,
}

/// Target of a relationship owned by a chart or chart user-shapes part.
#[derive(Debug, Clone)]
pub enum ChartRelationshipTarget {
    /// A directly related part embedded in the containing package
    Embedded {
        /// Complete target-part bytes
        data: Vec<u8>,
        /// OPC content type for the target part
        content_type: String,
        /// Filename extension without a leading dot
        extension: String,
    },
    /// An external relationship target
    External {
        /// External target URI
        target: String,
    },
}

/// One relationship owned by a chart or chart user-shapes part.
#[derive(Debug, Clone)]
pub struct ChartRelationship {
    /// Relationship identifier referenced by the owning part's XML
    pub relationship_id: String,
    /// Relationship type URI
    pub relationship_type: String,
    /// Internal payload or external target
    pub target: ChartRelationshipTarget,
}

/// Backwards-compatible name for a chart user-shapes relationship target.
pub type ChartUserShapesRelationshipTarget = ChartRelationshipTarget;

/// Backwards-compatible name for a chart user-shapes relationship.
pub type ChartUserShapesRelationship = ChartRelationship;

/// Lossless chart user-shapes XML and its direct relationship targets.
#[derive(Debug, Clone)]
pub struct ChartUserShapesPart {
    /// Complete chart user-shapes XML document
    pub xml: Vec<u8>,
    /// Relationships owned by the chart user-shapes part
    pub relationships: Vec<ChartRelationship>,
}

impl ChartUserShapesPart {
    /// Create a relationship-free user-shapes drawing part.
    pub fn new(xml: Vec<u8>) -> Self {
        Self {
            xml,
            relationships: Vec::new(),
        }
    }
}

impl ChartExternalDataPart {
    /// Create an embedded OOXML spreadsheet payload.
    pub fn embedded_workbook(data: Vec<u8>) -> Self {
        Self {
            relationship_type: litchi_opc::constants::relationship_type::PACKAGE.to_string(),
            target: ChartExternalDataTarget::Embedded {
                data,
                content_type: litchi_opc::constants::content_type::OFC_PACKAGE.to_string(),
                extension: "xlsx".to_string(),
            },
        }
    }

    /// Create a linked OOXML package relationship.
    pub fn linked_package(target: impl Into<String>) -> Self {
        Self {
            relationship_type: litchi_opc::constants::relationship_type::PACKAGE.to_string(),
            target: ChartExternalDataTarget::Linked {
                target: target.into(),
            },
        }
    }
}

pub(crate) fn is_chart_external_data_relationship_type(relationship_type: &str) -> bool {
    chart_external_data_content_type(relationship_type).is_some()
}

pub(crate) fn chart_external_data_content_type(relationship_type: &str) -> Option<&'static str> {
    match relationship_type {
        litchi_opc::constants::relationship_type::PACKAGE
        | "http://purl.oclc.org/ooxml/officeDocument/relationships/package" => {
            Some(litchi_opc::constants::content_type::OFC_PACKAGE)
        },
        litchi_opc::constants::relationship_type::OLE_OBJECT
        | "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject" => {
            Some(litchi_opc::constants::content_type::OFC_OLE_OBJECT)
        },
        _ => None,
    }
}

pub(crate) fn is_chart_user_shapes_relationship_type(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        litchi_opc::constants::relationship_type::CHART_USER_SHAPES
            | "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes"
    )
}

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

pub(crate) fn validate_chart_anchor(anchor: &ChartAnchor) -> Result<()> {
    if anchor.to_row < anchor.from_row || anchor.to_col < anchor.from_col {
        return Err(OoxmlError::InvalidFormat(
            "chart anchor cannot be descending".to_string(),
        ));
    }
    if anchor.to_row >= WORKSHEET_ROW_COUNT || anchor.to_col >= WORKSHEET_COLUMN_COUNT {
        return Err(OoxmlError::InvalidFormat(
            "chart anchor exceeds worksheet bounds".to_string(),
        ));
    }
    if [
        anchor.from_col_offset,
        anchor.from_row_offset,
        anchor.to_col_offset,
        anchor.to_row_offset,
    ]
    .iter()
    .any(|offset| *offset < 0)
    {
        return Err(OoxmlError::InvalidFormat(
            "chart anchor offsets cannot be negative".to_string(),
        ));
    }
    Ok(())
}

pub(crate) fn write_worksheet_chart_anchors(
    xml: &mut String,
    charts: &[WorksheetChart],
    object_id_offset: usize,
    relationship_id_offset: usize,
) -> Result<()> {
    use std::fmt::Write as _;

    for (index, chart) in charts.iter().enumerate() {
        validate_chart_anchor(&chart.anchor)?;
        let object_id = object_id_offset
            .checked_add(index)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| OoxmlError::InvalidFormat("chart object ID overflow".to_string()))?;
        let relationship_id = relationship_id_offset
            .checked_add(index)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat("chart relationship ID overflow".to_string())
            })?;
        let anchor = &chart.anchor;
        xml.push_str("<xdr:twoCellAnchor>");
        write!(
            xml,
            "<xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>",
            anchor.from_col,
            anchor.from_col_offset,
            anchor.from_row,
            anchor.from_row_offset
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        write!(
            xml,
            "<xdr:to><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:to>",
            anchor.to_col,
            anchor.to_col_offset,
            anchor.to_row,
            anchor.to_row_offset
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        write!(
            xml,
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{object_id}" name="Chart {}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
            index + 1
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        xml.push_str(
            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#,
        );
        write!(
            xml,
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{relationship_id}"/>"#
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        xml.push_str(
            "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>",
        );
    }
    Ok(())
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
    /// Optional package payload targeted by `chart.external_data`
    pub external_data_part: Option<ChartExternalDataPart>,
    /// Optional chart user-shapes drawing and its direct related resources
    pub user_shapes_part: Option<ChartUserShapesPart>,
    /// Other direct relationships owned by the chart part
    pub additional_relationships: Vec<ChartRelationship>,
}

impl WorksheetChart {
    /// Create a new worksheet chart.
    pub fn new(chart: ChartModel, anchor: ChartAnchor) -> Self {
        Self {
            chart,
            anchor,
            external_data_part: None,
            user_shapes_part: None,
            additional_relationships: Vec::new(),
        }
    }

    /// Attach an embedded OOXML workbook as the chart's external data.
    pub fn with_embedded_workbook(mut self, data: Vec<u8>) -> Self {
        self.chart.external_data = Some(crate::charts::ChartExternalData::pending());
        self.external_data_part = Some(ChartExternalDataPart::embedded_workbook(data));
        self
    }

    /// Attach a package-backed external-data relationship.
    pub fn with_external_data_part(
        mut self,
        part: ChartExternalDataPart,
        auto_update: Option<bool>,
    ) -> Self {
        let mut metadata = crate::charts::ChartExternalData::pending();
        metadata.auto_update = auto_update;
        self.chart.external_data = Some(metadata);
        self.external_data_part = Some(part);
        self
    }

    /// Attach a chart user-shapes drawing part.
    pub fn with_user_shapes_part(mut self, part: ChartUserShapesPart) -> Self {
        self.chart.user_shapes = Some(crate::charts::ChartUserShapes::pending());
        self.user_shapes_part = Some(part);
        self
    }

    /// Add a direct chart-part relationship retained with the worksheet chart.
    pub fn with_additional_relationship(mut self, relationship: ChartRelationship) -> Self {
        self.additional_relationships.push(relationship);
        self
    }

    /// Convert this chart into a pivot chart bound to a pivot table by name.
    ///
    /// Sets the chart's pivot source to `pivot_table_name` (with the default
    /// format ID) and gives every existing series without an extension list
    /// the default all-visible drop-zone options. The name is validated
    /// against the workbook's pivot tables and normalized to its
    /// sheet-qualified form when the workbook is saved.
    pub fn into_pivot_chart(mut self, pivot_table_name: &str) -> Self {
        self.chart.pivot_source = Some(crate::charts::chart::PivotSource::new(
            pivot_table_name,
            crate::xlsx::pivot_chart::DEFAULT_PIVOT_CHART_FORMAT_ID,
        ));
        let extension = crate::charts::ChartExtensionList::from_xml(
            crate::xlsx::pivot_chart::default_pivot_options_extension_xml(),
        )
        .expect("default pivot-options extension XML is a valid extLst fragment");
        for_each_series_mut(&mut self.chart, |series| {
            if series.extension_list.is_none() {
                series.extension_list = Some(extension.clone());
            }
        });
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

/// Parse a chart from chart XML and drawing anchor.
pub fn parse_chart_from_xml(chart_xml: &[u8], anchor: ChartAnchor) -> Result<WorksheetChart> {
    let chart = crate::charts::reader::parse_chart(chart_xml)?;
    Ok(WorksheetChart::new(chart, anchor))
}

fn for_each_series_mut(chart: &mut ChartModel, mut apply: impl FnMut(&mut Series)) {
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

/// Generate chart XML for a worksheet chart.
pub fn generate_chart_xml(chart: &ChartModel) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    crate::charts::writer::write_chart(&mut output, chart)
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
    Ok(output)
}

pub(crate) fn generate_chart_xml_with_external_data_id(
    chart: &ChartModel,
    external_data_relationship_id: Option<&str>,
    user_shapes_relationship_id: Option<&str>,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    crate::charts::writer::write_chart_with_relationship_ids(
        &mut output,
        chart,
        external_data_relationship_id,
        user_shapes_relationship_id,
    )
    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
    Ok(output)
}

pub(crate) fn chart_user_shapes_relationship_ids(xml: &[u8]) -> Result<HashSet<String>> {
    const RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let xml = crate::common::mce::process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut relationship_ids = HashSet::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if depth == 0 {
                    if saw_root
                        || !crate::common::xml::is_drawingml_chart_name(
                            &namespace,
                            element.name(),
                            b"userShapes",
                        )
                    {
                        return Err(OoxmlError::InvalidFormat(
                            "chart user-shapes XML must have one chart userShapes root".into(),
                        ));
                    }
                    saw_root = true;
                }
                for attribute in element.attributes() {
                    let attribute =
                        attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    let (attribute_namespace, _) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if matches!(
                        attribute_namespace,
                        ResolveResult::Bound(Namespace(value))
                            if value == RELATIONSHIPS_NAMESPACE
                                || value == STRICT_RELATIONSHIPS_NAMESPACE
                    ) {
                        relationship_ids.insert(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                                .into_owned(),
                        );
                    }
                }
                if matches!(event, Event::Start(_)) {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "chart user-shapes XML nesting is too deep".into(),
                        )
                    })?;
                } else if depth == 0 {
                    closed_root = true;
                }
            },
            Event::End(ref element) => {
                if depth == 0 {
                    return Err(OoxmlError::InvalidFormat(
                        "chart user-shapes XML has an unmatched closing element".into(),
                    ));
                }
                depth -= 1;
                if depth == 0 {
                    if !crate::common::xml::is_drawingml_chart_name(
                        &namespace,
                        element.name(),
                        b"userShapes",
                    ) {
                        return Err(OoxmlError::InvalidFormat(
                            "chart user-shapes XML has an invalid root closing element".into(),
                        ));
                    }
                    closed_root = true;
                }
            },
            Event::Text(ref text)
                if depth == 0
                    && !text
                        .decode()
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?
                        .trim()
                        .is_empty() =>
            {
                return Err(OoxmlError::InvalidFormat(
                    "chart user-shapes XML contains text outside its root".into(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(OoxmlError::InvalidFormat(
                    "chart user-shapes XML contains data outside its root".into(),
                ));
            },
            Event::DocType(_) => {
                return Err(OoxmlError::InvalidFormat(
                    "chart user-shapes XML cannot contain a document type".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !saw_root || !closed_root || depth != 0 {
        return Err(OoxmlError::InvalidFormat(
            "chart user-shapes XML has no complete root".into(),
        ));
    }
    Ok(relationship_ids)
}

fn append_chart_line_fragment<'a>(
    fragments: &mut Vec<&'a [u8]>,
    lines: Option<&'a crate::charts::ChartLines>,
) {
    if let Some(shape_properties) = lines.and_then(|lines| lines.shape_properties.as_ref()) {
        fragments.push(shape_properties.as_xml());
    }
}

fn append_up_down_bar_fragments<'a>(
    fragments: &mut Vec<&'a [u8]>,
    bars: Option<&'a crate::charts::UpDownBars>,
) {
    let Some(bars) = bars else {
        return;
    };
    append_chart_line_fragment(fragments, bars.up_bars.as_ref());
    append_chart_line_fragment(fragments, bars.down_bars.as_ref());
    if let Some(extension_list) = bars.extension_list.as_ref() {
        fragments.push(extension_list.as_xml());
    }
}

fn append_point_data_label_fragments<'a>(
    fragments: &mut Vec<&'a [u8]>,
    label: Option<&'a crate::charts::DataLabel>,
) {
    let Some(label) = label else {
        return;
    };
    if let Some(shape_properties) = label.shape_properties.as_ref() {
        fragments.push(shape_properties.as_xml());
    }
    if let Some(text_properties) = label.text_properties.as_ref() {
        fragments.push(text_properties.as_xml());
    }
    if let Some(extension_list) = label.extension_list.as_ref() {
        fragments.push(extension_list.as_xml());
    }
}

fn append_data_label_fragments<'a>(
    fragments: &mut Vec<&'a [u8]>,
    labels: Option<&'a crate::charts::DataLabels>,
) {
    let Some(labels) = labels else {
        return;
    };
    if let Some(shape_properties) = labels.shape_properties.as_ref() {
        fragments.push(shape_properties.as_xml());
    }
    if let Some(text_properties) = labels.text_properties.as_ref() {
        fragments.push(text_properties.as_xml());
    }
    append_chart_line_fragment(fragments, labels.leader_lines.as_ref());
    if let Some(extension_list) = labels.extension_list.as_ref() {
        fragments.push(extension_list.as_xml());
    }
    for label in &labels.labels {
        append_point_data_label_fragments(fragments, Some(label));
    }
}

pub(crate) fn chart_fragment_relationship_ids(chart: &ChartModel) -> Result<HashSet<String>> {
    const RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
        b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let mut relationship_ids = HashSet::new();
    let mut fragments = Vec::new();
    fragments.extend(
        [
            chart
                .shape_properties
                .as_ref()
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .text_properties
                .as_ref()
                .map(crate::charts::ChartTextProperties::as_xml),
            chart
                .extension_list
                .as_ref()
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .chart_extension_list
                .as_ref()
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .title
                .as_ref()
                .and(chart.title_shape_properties.as_ref())
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .title
                .as_ref()
                .and(chart.title_text_properties.as_ref())
                .map(crate::charts::ChartTextProperties::as_xml),
            chart
                .title
                .as_ref()
                .and(chart.title_extension_list.as_ref())
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .plot_area
                .shape_properties
                .as_ref()
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .plot_area
                .extension_list
                .as_ref()
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .plot_area
                .data_table
                .as_ref()
                .and_then(|table| table.shape_properties.as_ref())
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .plot_area
                .data_table
                .as_ref()
                .and_then(|table| table.text_properties.as_ref())
                .map(crate::charts::ChartTextProperties::as_xml),
            chart
                .plot_area
                .data_table
                .as_ref()
                .and_then(|table| table.extension_list.as_ref())
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .floor
                .as_ref()
                .and_then(|surface| surface.shape_properties.as_ref())
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .back_wall
                .as_ref()
                .and_then(|surface| surface.shape_properties.as_ref())
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .side_wall
                .as_ref()
                .and_then(|surface| surface.shape_properties.as_ref())
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .floor
                .as_ref()
                .and_then(|surface| surface.extension_list.as_ref())
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .back_wall
                .as_ref()
                .and_then(|surface| surface.extension_list.as_ref())
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .side_wall
                .as_ref()
                .and_then(|surface| surface.extension_list.as_ref())
                .map(crate::charts::ChartExtensionList::as_xml),
            chart
                .legend
                .as_ref()
                .and_then(|legend| legend.shape_properties.as_ref())
                .map(crate::charts::ChartShapeProperties::as_xml),
            chart
                .legend
                .as_ref()
                .and_then(|legend| legend.text_properties.as_ref())
                .map(crate::charts::ChartTextProperties::as_xml),
            chart
                .legend
                .as_ref()
                .and_then(|legend| legend.extension_list.as_ref())
                .map(crate::charts::ChartExtensionList::as_xml),
        ]
        .into_iter()
        .flatten(),
    );
    if let Some(legend) = chart.legend.as_ref() {
        for entry in &legend.entries {
            fragments.extend(
                [
                    entry
                        .text_properties
                        .as_ref()
                        .map(crate::charts::ChartTextProperties::as_xml),
                    entry
                        .extension_list
                        .as_ref()
                        .map(crate::charts::ChartExtensionList::as_xml),
                ]
                .into_iter()
                .flatten(),
            );
        }
    }
    if let Some(formats) = chart.pivot_formats.as_ref() {
        for format in formats {
            fragments.extend(
                [
                    format
                        .shape_properties
                        .as_ref()
                        .map(crate::charts::ChartShapeProperties::as_xml),
                    format
                        .text_properties
                        .as_ref()
                        .map(crate::charts::ChartTextProperties::as_xml),
                    format
                        .extension_list
                        .as_ref()
                        .map(crate::charts::ChartExtensionList::as_xml),
                    format
                        .marker
                        .as_ref()
                        .and_then(|marker| marker.shape_properties.as_ref())
                        .map(crate::charts::ChartShapeProperties::as_xml),
                    format
                        .marker
                        .as_ref()
                        .and_then(|marker| marker.extension_list.as_ref())
                        .map(crate::charts::ChartExtensionList::as_xml),
                ]
                .into_iter()
                .flatten(),
            );
            append_point_data_label_fragments(&mut fragments, format.data_label.as_ref());
        }
    }
    for axis in &chart.plot_area.axes {
        let common = axis.common();
        fragments.extend(
            [
                common
                    .title
                    .as_ref()
                    .and(common.title_shape_properties.as_ref())
                    .map(crate::charts::ChartShapeProperties::as_xml),
                common
                    .title
                    .as_ref()
                    .and(common.title_text_properties.as_ref())
                    .map(crate::charts::ChartTextProperties::as_xml),
                common
                    .title
                    .as_ref()
                    .and(common.title_extension_list.as_ref())
                    .map(crate::charts::ChartExtensionList::as_xml),
                common
                    .major_gridlines
                    .as_ref()
                    .and_then(|lines| lines.shape_properties.as_ref())
                    .map(crate::charts::ChartShapeProperties::as_xml),
                common
                    .minor_gridlines
                    .as_ref()
                    .and_then(|lines| lines.shape_properties.as_ref())
                    .map(crate::charts::ChartShapeProperties::as_xml),
                common
                    .shape_properties
                    .as_ref()
                    .map(crate::charts::ChartShapeProperties::as_xml),
                common
                    .text_properties
                    .as_ref()
                    .map(crate::charts::ChartTextProperties::as_xml),
                common
                    .scaling_extension_list
                    .as_ref()
                    .map(crate::charts::ChartExtensionList::as_xml),
                common
                    .extension_list
                    .as_ref()
                    .map(crate::charts::ChartExtensionList::as_xml),
            ]
            .into_iter()
            .flatten(),
        );
        if let Axis::Value(axis) = axis {
            if let Some(display_units) = axis.display_units.as_ref() {
                fragments.extend(
                    [
                        display_units
                            .label_shape_properties
                            .as_ref()
                            .map(crate::charts::ChartShapeProperties::as_xml),
                        display_units
                            .label_text_properties
                            .as_ref()
                            .map(crate::charts::ChartTextProperties::as_xml),
                        display_units
                            .extension_list
                            .as_ref()
                            .map(crate::charts::ChartExtensionList::as_xml),
                    ]
                    .into_iter()
                    .flatten(),
                );
            }
        }
    }
    for group in &chart.plot_area.type_groups {
        match group {
            crate::charts::TypeGroup::Area(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
            },
            crate::charts::TypeGroup::Area3D(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
            },
            crate::charts::TypeGroup::Bar(group) => {
                for lines in &group.series_lines {
                    append_chart_line_fragment(&mut fragments, Some(lines));
                }
            },
            crate::charts::TypeGroup::Line(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
                append_chart_line_fragment(&mut fragments, group.high_low_lines.as_ref());
                append_up_down_bar_fragments(&mut fragments, group.up_down_bars.as_ref());
            },
            crate::charts::TypeGroup::Line3D(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
            },
            crate::charts::TypeGroup::OfPie(group) => {
                for lines in &group.series_lines {
                    append_chart_line_fragment(&mut fragments, Some(lines));
                }
            },
            crate::charts::TypeGroup::Stock(group) => {
                append_chart_line_fragment(&mut fragments, group.drop_lines.as_ref());
                append_chart_line_fragment(&mut fragments, group.high_low_lines.as_ref());
                append_up_down_bar_fragments(&mut fragments, group.up_down_bars.as_ref());
            },
            crate::charts::TypeGroup::Surface(group) => {
                if let Some(formats) = group.band_formats.as_ref() {
                    for format in formats {
                        if let Some(shape_properties) = format.shape_properties.as_ref() {
                            fragments.push(shape_properties.as_xml());
                        }
                    }
                }
            },
            crate::charts::TypeGroup::Surface3D(group) => {
                if let Some(formats) = group.band_formats.as_ref() {
                    for format in formats {
                        if let Some(shape_properties) = format.shape_properties.as_ref() {
                            fragments.push(shape_properties.as_xml());
                        }
                    }
                }
            },
            _ => {},
        }
        for series in &group.common().series {
            fragments.extend(
                [
                    series
                        .shape_properties
                        .as_ref()
                        .map(crate::charts::ChartShapeProperties::as_xml),
                    series
                        .extension_list
                        .as_ref()
                        .map(crate::charts::ChartExtensionList::as_xml),
                    series
                        .marker_shape_properties
                        .as_ref()
                        .map(crate::charts::ChartShapeProperties::as_xml),
                    series
                        .marker_extension_list
                        .as_ref()
                        .map(crate::charts::ChartExtensionList::as_xml),
                ]
                .into_iter()
                .flatten(),
            );
            append_data_label_fragments(&mut fragments, series.data_labels.as_ref());
            for error_bar in &series.error_bars {
                if let Some(shape_properties) = error_bar.shape_properties.as_ref() {
                    fragments.push(shape_properties.as_xml());
                }
                if let Some(extension_list) = error_bar.extension_list.as_ref() {
                    fragments.push(extension_list.as_xml());
                }
            }
            for trendline in &series.trendlines {
                fragments.extend(
                    [
                        trendline
                            .shape_properties
                            .as_ref()
                            .map(crate::charts::ChartShapeProperties::as_xml),
                        trendline
                            .label_shape_properties
                            .as_ref()
                            .map(crate::charts::ChartShapeProperties::as_xml),
                        trendline
                            .label_text_properties
                            .as_ref()
                            .map(crate::charts::ChartTextProperties::as_xml),
                        trendline
                            .label_extension_list
                            .as_ref()
                            .map(crate::charts::ChartExtensionList::as_xml),
                        trendline
                            .extension_list
                            .as_ref()
                            .map(crate::charts::ChartExtensionList::as_xml),
                    ]
                    .into_iter()
                    .flatten(),
                );
            }
            for point in &series.data_points {
                fragments.extend(
                    [
                        point
                            .shape_properties
                            .as_ref()
                            .map(crate::charts::ChartShapeProperties::as_xml),
                        point
                            .extension_list
                            .as_ref()
                            .map(crate::charts::ChartExtensionList::as_xml),
                        point
                            .marker_shape_properties
                            .as_ref()
                            .map(crate::charts::ChartShapeProperties::as_xml),
                        point
                            .marker_extension_list
                            .as_ref()
                            .map(crate::charts::ChartExtensionList::as_xml),
                    ]
                    .into_iter()
                    .flatten(),
                );
            }
        }
        if let Some(extension_list) = group.common().extension_list.as_ref() {
            fragments.push(extension_list.as_xml());
        }
    }
    for xml in fragments {
        let mut reader = NsReader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            let (_, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    for attribute in element.attributes() {
                        let attribute =
                            attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        let (attribute_namespace, _) =
                            reader.resolver().resolve_attribute(attribute.key);
                        if matches!(
                            attribute_namespace,
                            ResolveResult::Bound(Namespace(value))
                                if value == RELATIONSHIPS_NAMESPACE
                                    || value == STRICT_RELATIONSHIPS_NAMESPACE
                        ) {
                            relationship_ids.insert(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Explicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                                    .into_owned(),
                            );
                        }
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
    }
    Ok(relationship_ids)
}
