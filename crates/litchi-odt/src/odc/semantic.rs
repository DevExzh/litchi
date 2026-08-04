//! Zero-copy typed views over retained standalone-chart XML.

use super::{Document, Element, ElementKind};
use litchi_core::{Error, Result};

const CHART_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Dimension {
    X,
    Y,
    Z,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataSourceLabels {
    None,
    Row,
    Column,
    Both,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegendPosition {
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
pub enum GridClass {
    Major,
    Minor,
}

#[derive(Debug, Clone, Copy)]
pub struct PlotArea<'a> {
    element: &'a Element,
}

impl<'a> PlotArea<'a> {
    pub fn element(self) -> &'a Element {
        self.element
    }

    pub fn cell_range_address(self) -> Option<&'a str> {
        self.element
            .attribute(Some(TABLE_NAMESPACE), "cell-range-address")
    }

    pub fn data_source_labels(self) -> Result<DataSourceLabels> {
        match chart_attribute(self.element, "data-source-has-labels") {
            None | Some("none") => Ok(DataSourceLabels::None),
            Some("row") => Ok(DataSourceLabels::Row),
            Some("column") => Ok(DataSourceLabels::Column),
            Some("both") => Ok(DataSourceLabels::Both),
            Some(value) => Err(invalid(format!(
                "invalid chart:data-source-has-labels '{value}'"
            ))),
        }
    }

    pub fn axes(self) -> impl Iterator<Item = Axis<'a>> + 'a {
        self.element
            .children_of_kind(ElementKind::Axis)
            .map(|element| Axis { element })
    }

    pub fn series(self) -> impl Iterator<Item = Series<'a>> + 'a {
        self.element
            .children_of_kind(ElementKind::Series)
            .map(|element| Series { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Axis<'a> {
    element: &'a Element,
}

impl<'a> Axis<'a> {
    pub fn element(self) -> &'a Element {
        self.element
    }

    pub fn dimension(self) -> Result<Dimension> {
        match chart_attribute(self.element, "dimension") {
            Some("x") => Ok(Dimension::X),
            Some("y") => Ok(Dimension::Y),
            Some("z") => Ok(Dimension::Z),
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
            .children_of_kind(ElementKind::Categories)
            .next()
            .and_then(|categories| {
                categories.attribute(Some(TABLE_NAMESPACE), "cell-range-address")
            })
    }

    pub fn grids(self) -> impl Iterator<Item = Grid<'a>> + 'a {
        self.element
            .children_of_kind(ElementKind::Grid)
            .map(|element| Grid { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Grid<'a> {
    element: &'a Element,
}

impl Grid<'_> {
    pub fn class(self) -> Result<GridClass> {
        match chart_attribute(self.element, "class") {
            Some("major") => Ok(GridClass::Major),
            Some("minor") => Ok(GridClass::Minor),
            Some(value) => Err(invalid(format!("invalid chart:grid class '{value}'"))),
            None => Err(invalid("chart:grid requires chart:class")),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Series<'a> {
    element: &'a Element,
}

impl<'a> Series<'a> {
    pub fn element(self) -> &'a Element {
        self.element
    }

    pub fn xml_id(self) -> Option<&'a str> {
        self.element
            .attribute(Some("http://www.w3.org/XML/1998/namespace"), "id")
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
            .children_of_kind(ElementKind::Domain)
            .filter_map(|domain| domain.attribute(Some(TABLE_NAMESPACE), "cell-range-address"))
    }

    pub fn data_points(self) -> impl Iterator<Item = DataPoint<'a>> + 'a {
        self.element
            .children_of_kind(ElementKind::DataPoint)
            .map(|element| DataPoint { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct DataPoint<'a> {
    element: &'a Element,
}

impl<'a> DataPoint<'a> {
    pub fn element(self) -> &'a Element {
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
pub struct Legend<'a> {
    element: &'a Element,
}

impl Legend<'_> {
    pub fn position(self) -> Result<LegendPosition> {
        match chart_attribute(self.element, "legend-position") {
            None | Some("end") => Ok(LegendPosition::End),
            Some("start") => Ok(LegendPosition::Start),
            Some("top") => Ok(LegendPosition::Top),
            Some("bottom") => Ok(LegendPosition::Bottom),
            Some("top-start") => Ok(LegendPosition::TopStart),
            Some("top-end") => Ok(LegendPosition::TopEnd),
            Some("bottom-start") => Ok(LegendPosition::BottomStart),
            Some("bottom-end") => Ok(LegendPosition::BottomEnd),
            Some(value) => Err(invalid(format!("invalid chart:legend-position '{value}'"))),
        }
    }
}

impl Document {
    pub fn plot_area(&self) -> Option<PlotArea<'_>> {
        self.chart()
            .children_of_kind(ElementKind::PlotArea)
            .next()
            .map(|element| PlotArea { element })
    }

    pub fn legend(&self) -> Option<Legend<'_>> {
        self.chart()
            .children_of_kind(ElementKind::Legend)
            .next()
            .map(|element| Legend { element })
    }
}

fn chart_attribute<'a>(element: &'a Element, local_name: &str) -> Option<&'a str> {
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

    fn document(body: &str) -> Document {
        let content = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="{CHART_NAMESPACE}" xmlns:t="{TABLE_NAMESPACE}"><o:body><o:chart><c:chart><c:legend c:legend-position="bottom-end"/><c:plot-area t:cell-range-address="Data.A1:C4" c:data-source-has-labels="both">{body}</c:plot-area></c:chart></o:chart></o:body></o:document-content>"#
        );
        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_CHART).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn exposes_typed_zero_copy_chart_semantics() {
        let document = document(
            r#"<c:axis c:dimension="x" c:name="x"><c:categories t:cell-range-address="Data.A2:A4"/><c:grid c:class="major"/></c:axis><c:series c:values-cell-range-address="Data.B2:B4" c:label-cell-address="Data.B1"><c:domain t:cell-range-address="Data.A2:A4"/><c:data-point c:repeated="3"/></c:series>"#,
        );
        assert_eq!(
            document.legend().unwrap().position().unwrap(),
            LegendPosition::BottomEnd
        );
        let plot = document.plot_area().unwrap();
        assert_eq!(plot.cell_range_address(), Some("Data.A1:C4"));
        assert_eq!(plot.data_source_labels().unwrap(), DataSourceLabels::Both);
        let axis = plot.axes().next().unwrap();
        assert_eq!(axis.dimension().unwrap(), Dimension::X);
        assert_eq!(axis.categories_range(), Some("Data.A2:A4"));
        assert_eq!(
            axis.grids().next().unwrap().class().unwrap(),
            GridClass::Major
        );
        let series = plot.series().next().unwrap();
        assert_eq!(series.values_range(), Some("Data.B2:B4"));
        assert_eq!(series.domains().next(), Some("Data.A2:A4"));
        assert_eq!(series.data_points().next().unwrap().repeated().unwrap(), 3);
    }

    #[test]
    fn validates_typed_enumerations_and_counts_lazily() {
        let document = document(
            r#"<c:axis c:dimension="time"><c:grid c:class="micro"/></c:axis><c:series><c:data-point c:repeated="0"/></c:series>"#,
        );
        let plot = document.plot_area().unwrap();
        let axis = plot.axes().next().unwrap();
        assert!(axis.dimension().is_err());
        assert!(axis.grids().next().unwrap().class().is_err());
        assert!(
            plot.series()
                .next()
                .unwrap()
                .data_points()
                .next()
                .unwrap()
                .repeated()
                .is_err()
        );
    }
}
