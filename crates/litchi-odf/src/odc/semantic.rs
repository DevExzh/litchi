//! Zero-copy typed views over retained standalone-chart XML.

use super::{ChartDocument, ChartElement, ChartElementKind};
use litchi_core::{Error, Result};

const CHART_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisDimension {
    X,
    Y,
    Z,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartDataSourceLabels {
    None,
    Row,
    Column,
    Both,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartLegendPosition {
    Start,
    End,
    Top,
    Bottom,
    TopStart,
    TopEnd,
    BottomStart,
    BottomEnd,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartGridClass {
    Major,
    Minor,
}

#[derive(Debug, Clone, Copy)]
pub struct ChartPlotArea<'a> {
    element: &'a ChartElement,
}

impl<'a> ChartPlotArea<'a> {
    pub fn element(self) -> &'a ChartElement {
        self.element
    }

    pub fn cell_range_address(self) -> Option<&'a str> {
        self.element
            .attribute(Some(TABLE_NAMESPACE), "cell-range-address")
    }

    pub fn data_source_labels(self) -> Result<ChartDataSourceLabels> {
        match chart_attribute(self.element, "data-source-has-labels") {
            None | Some("none") => Ok(ChartDataSourceLabels::None),
            Some("row") => Ok(ChartDataSourceLabels::Row),
            Some("column") => Ok(ChartDataSourceLabels::Column),
            Some("both") => Ok(ChartDataSourceLabels::Both),
            Some(value) => Err(invalid(format!(
                "invalid chart:data-source-has-labels '{value}'"
            ))),
        }
    }

    pub fn axes(self) -> impl Iterator<Item = ChartAxis<'a>> + 'a {
        self.element
            .children_of_kind(ChartElementKind::Axis)
            .map(|element| ChartAxis { element })
    }

    pub fn series(self) -> impl Iterator<Item = ChartSeries<'a>> + 'a {
        self.element
            .children_of_kind(ChartElementKind::Series)
            .map(|element| ChartSeries { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ChartAxis<'a> {
    element: &'a ChartElement,
}

impl<'a> ChartAxis<'a> {
    pub fn element(self) -> &'a ChartElement {
        self.element
    }

    pub fn dimension(self) -> Result<ChartAxisDimension> {
        match chart_attribute(self.element, "dimension") {
            Some("x") => Ok(ChartAxisDimension::X),
            Some("y") => Ok(ChartAxisDimension::Y),
            Some("z") => Ok(ChartAxisDimension::Z),
            Some(value) => Err(invalid(format!("invalid chart:dimension '{value}'"))),
            None => Err(invalid("chart:axis requires chart:dimension")),
        }
    }

    pub fn name(self) -> Option<&'a str> {
        chart_attribute(self.element, "name")
    }

    pub fn style_name(self) -> Option<&'a str> {
        chart_attribute(self.element, "style-name")
    }

    pub fn categories_range(self) -> Option<&'a str> {
        self.element
            .children_of_kind(ChartElementKind::Categories)
            .next()
            .and_then(|categories| {
                categories.attribute(Some(TABLE_NAMESPACE), "cell-range-address")
            })
    }

    pub fn grids(self) -> impl Iterator<Item = ChartGrid<'a>> + 'a {
        self.element
            .children_of_kind(ChartElementKind::Grid)
            .map(|element| ChartGrid { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ChartGrid<'a> {
    element: &'a ChartElement,
}

impl ChartGrid<'_> {
    pub fn class(self) -> Result<ChartGridClass> {
        match chart_attribute(self.element, "class") {
            Some("major") => Ok(ChartGridClass::Major),
            Some("minor") => Ok(ChartGridClass::Minor),
            Some(value) => Err(invalid(format!("invalid chart:grid class '{value}'"))),
            None => Err(invalid("chart:grid requires chart:class")),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ChartSeries<'a> {
    element: &'a ChartElement,
}

impl<'a> ChartSeries<'a> {
    pub fn element(self) -> &'a ChartElement {
        self.element
    }

    /// Return the chart class QName lexically; namespace resolution remains inert.
    pub fn class_name(self) -> Option<&'a str> {
        chart_attribute(self.element, "class")
    }

    pub fn values_range(self) -> Option<&'a str> {
        chart_attribute(self.element, "values-cell-range-address")
    }

    pub fn label_cell(self) -> Option<&'a str> {
        chart_attribute(self.element, "label-cell-address")
    }

    pub fn attached_axis(self) -> Option<&'a str> {
        chart_attribute(self.element, "attached-axis")
    }

    pub fn domains(self) -> impl Iterator<Item = &'a str> + 'a {
        self.element
            .children_of_kind(ChartElementKind::Domain)
            .filter_map(|domain| {
                domain.attribute(Some(TABLE_NAMESPACE), "cell-range-address")
            })
    }

    pub fn data_points(self) -> impl Iterator<Item = ChartDataPoint<'a>> + 'a {
        self.element
            .children_of_kind(ChartElementKind::DataPoint)
            .map(|element| ChartDataPoint { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ChartDataPoint<'a> {
    element: &'a ChartElement,
}

impl<'a> ChartDataPoint<'a> {
    pub fn element(self) -> &'a ChartElement {
        self.element
    }

    pub fn repeated(self) -> Result<u32> {
        let Some(value) = chart_attribute(self.element, "repeated") else {
            return Ok(1);
        };
        let repeated = value
            .parse::<u32>()
            .map_err(|_| invalid(format!("invalid chart:repeated '{value}'")))?;
        if repeated == 0 {
            return Err(invalid("chart:repeated must be positive"));
        }
        Ok(repeated)
    }

    pub fn style_name(self) -> Option<&'a str> {
        chart_attribute(self.element, "style-name")
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ChartLegend<'a> {
    element: &'a ChartElement,
}

impl ChartLegend<'_> {
    pub fn position(self) -> Result<ChartLegendPosition> {
        match chart_attribute(self.element, "legend-position") {
            None | Some("end") => Ok(ChartLegendPosition::End),
            Some("start") => Ok(ChartLegendPosition::Start),
            Some("top") => Ok(ChartLegendPosition::Top),
            Some("bottom") => Ok(ChartLegendPosition::Bottom),
            Some("top-start") => Ok(ChartLegendPosition::TopStart),
            Some("top-end") => Ok(ChartLegendPosition::TopEnd),
            Some("bottom-start") => Ok(ChartLegendPosition::BottomStart),
            Some("bottom-end") => Ok(ChartLegendPosition::BottomEnd),
            Some(value) => Err(invalid(format!("invalid chart:legend-position '{value}'"))),
        }
    }
}

impl ChartDocument {
    pub fn plot_area(&self) -> Option<ChartPlotArea<'_>> {
        self.chart()
            .children_of_kind(ChartElementKind::PlotArea)
            .next()
            .map(|element| ChartPlotArea { element })
    }

    pub fn legend(&self) -> Option<ChartLegend<'_>> {
        self.chart()
            .children_of_kind(ChartElementKind::Legend)
            .next()
            .map(|element| ChartLegend { element })
    }
}

fn chart_attribute<'a>(element: &'a ChartElement, local_name: &str) -> Option<&'a str> {
    element.attribute(Some(CHART_NAMESPACE), local_name)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;

    fn document(body: &str) -> ChartDocument {
        let content = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="{CHART_NAMESPACE}" xmlns:t="{TABLE_NAMESPACE}"><o:body><o:chart><c:chart><c:legend c:legend-position="bottom-end"/><c:plot-area t:cell-range-address="Data.A1:C4" c:data-source-has-labels="both">{body}</c:plot-area></c:chart></o:chart></o:body></o:document-content>"#
        );
        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_CHART).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        ChartDocument::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn exposes_typed_zero_copy_chart_semantics() {
        let document = document(r#"<c:axis c:dimension="x" c:name="x"><c:categories t:cell-range-address="Data.A2:A4"/><c:grid c:class="major"/></c:axis><c:series c:values-cell-range-address="Data.B2:B4" c:label-cell-address="Data.B1"><c:domain t:cell-range-address="Data.A2:A4"/><c:data-point c:repeated="3"/></c:series>"#);
        assert_eq!(document.legend().unwrap().position().unwrap(), ChartLegendPosition::BottomEnd);
        let plot = document.plot_area().unwrap();
        assert_eq!(plot.cell_range_address(), Some("Data.A1:C4"));
        assert_eq!(plot.data_source_labels().unwrap(), ChartDataSourceLabels::Both);
        let axis = plot.axes().next().unwrap();
        assert_eq!(axis.dimension().unwrap(), ChartAxisDimension::X);
        assert_eq!(axis.categories_range(), Some("Data.A2:A4"));
        assert_eq!(axis.grids().next().unwrap().class().unwrap(), ChartGridClass::Major);
        let series = plot.series().next().unwrap();
        assert_eq!(series.values_range(), Some("Data.B2:B4"));
        assert_eq!(series.domains().next(), Some("Data.A2:A4"));
        assert_eq!(series.data_points().next().unwrap().repeated().unwrap(), 3);
    }

    #[test]
    fn validates_typed_enumerations_and_counts_lazily() {
        let document = document(r#"<c:axis c:dimension="time"><c:grid c:class="micro"/></c:axis><c:series><c:data-point c:repeated="0"/></c:series>"#);
        let plot = document.plot_area().unwrap();
        let axis = plot.axes().next().unwrap();
        assert!(axis.dimension().is_err());
        assert!(axis.grids().next().unwrap().class().is_err());
        assert!(plot.series().next().unwrap().data_points().next().unwrap().repeated().is_err());
    }
}
