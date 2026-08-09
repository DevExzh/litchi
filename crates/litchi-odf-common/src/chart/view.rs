//! Borrowed semantic views over retained chart content.

use super::axis::Dimension;
use super::grid::Class;
use super::legend::Position;
use super::plot_area::Labels;
use super::reader::{Element, Kind};
use crate::namespace::{CHARTNS, TABLENS, XMLNS};
use litchi_core::{Error, Result};

#[derive(Debug, Clone, Copy)]
pub struct PlotArea<'a> {
    element: &'a Element,
}

impl<'a> PlotArea<'a> {
    #[must_use]
    pub fn element(self) -> &'a Element {
        self.element
    }

    #[must_use]
    pub fn cell_range_address(self) -> Option<&'a str> {
        self.element.attribute(Some(TABLENS), "cell-range-address")
    }

    /// Decode the plot area's label orientation.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error when the attribute uses an unsupported
    /// ODF value.
    pub fn data_source_labels(self) -> Result<Labels> {
        match chart_attribute(self.element, "data-source-has-labels") {
            None | Some("none") => Ok(Labels::None),
            Some("row") => Ok(Labels::Row),
            Some("column") => Ok(Labels::Column),
            Some("both") => Ok(Labels::Both),
            Some(value) => Err(invalid(format!(
                "invalid chart:data-source-has-labels '{value}'"
            ))),
        }
    }

    pub fn axes(self) -> impl Iterator<Item = Axis<'a>> + 'a {
        self.element
            .children_of_kind(Kind::Axis)
            .map(|element| Axis { element })
    }

    pub fn series(self) -> impl Iterator<Item = Series<'a>> + 'a {
        self.element
            .children_of_kind(Kind::Series)
            .map(|element| Series { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Axis<'a> {
    element: &'a Element,
}

impl<'a> Axis<'a> {
    #[must_use]
    pub fn element(self) -> &'a Element {
        self.element
    }

    /// Decode the axis dimension.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error when the axis has no supported
    /// `chart:dimension` value.
    pub fn dimension(self) -> Result<Dimension> {
        match chart_attribute(self.element, "dimension") {
            Some("x") => Ok(Dimension::X),
            Some("y") => Ok(Dimension::Y),
            Some("z") => Ok(Dimension::Z),
            Some(value) => Err(invalid(format!("invalid chart:dimension '{value}'"))),
            None => Err(invalid("chart:axis requires chart:dimension")),
        }
    }

    #[must_use]
    pub fn name(self) -> Option<&'a str> {
        chart_attribute(self.element, "name")
    }

    #[must_use]
    pub fn style_name(self) -> Option<&'a str> {
        chart_attribute(self.element, "style-name")
    }

    #[must_use]
    pub fn categories_range(self) -> Option<&'a str> {
        self.element
            .children_of_kind(Kind::Categories)
            .next()
            .and_then(|categories| categories.attribute(Some(TABLENS), "cell-range-address"))
    }

    pub fn grids(self) -> impl Iterator<Item = Grid<'a>> + 'a {
        self.element
            .children_of_kind(Kind::Grid)
            .map(|element| Grid { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Grid<'a> {
    element: &'a Element,
}

impl<'a> Grid<'a> {
    #[must_use]
    pub fn element(self) -> &'a Element {
        self.element
    }

    /// Decode the grid class.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error when the grid has no supported
    /// `chart:class` value.
    pub fn class(self) -> Result<Class> {
        match chart_attribute(self.element, "class") {
            Some("major") => Ok(Class::Major),
            Some("minor") => Ok(Class::Minor),
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
    #[must_use]
    pub fn element(self) -> &'a Element {
        self.element
    }

    #[must_use]
    pub fn xml_id(self) -> Option<&'a str> {
        self.element.attribute(Some(XMLNS), "id")
    }

    /// Return the class `QName` lexically; namespace resolution remains inert.
    #[must_use]
    pub fn class_name(self) -> Option<&'a str> {
        chart_attribute(self.element, "class")
    }

    #[must_use]
    pub fn values_range(self) -> Option<&'a str> {
        chart_attribute(self.element, "values-cell-range-address")
    }

    #[must_use]
    pub fn label_cell(self) -> Option<&'a str> {
        chart_attribute(self.element, "label-cell-address")
    }

    #[must_use]
    pub fn attached_axis(self) -> Option<&'a str> {
        chart_attribute(self.element, "attached-axis")
    }

    pub fn domains(self) -> impl Iterator<Item = &'a str> + 'a {
        self.element
            .children_of_kind(Kind::Domain)
            .filter_map(|domain| domain.attribute(Some(TABLENS), "cell-range-address"))
    }

    pub fn data_points(self) -> impl Iterator<Item = DataPoint<'a>> + 'a {
        self.element
            .children_of_kind(Kind::DataPoint)
            .map(|element| DataPoint { element })
    }
}

#[derive(Debug, Clone, Copy)]
pub struct DataPoint<'a> {
    element: &'a Element,
}

impl<'a> DataPoint<'a> {
    #[must_use]
    pub fn element(self) -> &'a Element {
        self.element
    }

    /// Decode the number of repeated data points.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error for a non-positive or non-numeric
    /// `chart:repeated` value.
    pub fn repeated(self) -> Result<u32> {
        let Some(value) = chart_attribute(self.element, "repeated") else {
            return Ok(1);
        };
        let repeated = value
            .parse::<u32>()
            .map_err(|error| invalid(format!("invalid chart:repeated '{value}': {error}")))?;
        if repeated == 0 {
            return Err(invalid("chart:repeated must be positive"));
        }
        Ok(repeated)
    }

    #[must_use]
    pub fn style_name(self) -> Option<&'a str> {
        chart_attribute(self.element, "style-name")
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Legend<'a> {
    element: &'a Element,
}

impl<'a> Legend<'a> {
    #[must_use]
    pub fn element(self) -> &'a Element {
        self.element
    }

    /// Decode the legend placement.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error when the placement token is unknown.
    pub fn position(self) -> Result<Position> {
        match chart_attribute(self.element, "legend-position") {
            None | Some("end") => Ok(Position::End),
            Some("start") => Ok(Position::Start),
            Some("top") => Ok(Position::Top),
            Some("bottom") => Ok(Position::Bottom),
            Some("top-start") => Ok(Position::TopStart),
            Some("top-end") => Ok(Position::TopEnd),
            Some("bottom-start") => Ok(Position::BottomStart),
            Some("bottom-end") => Ok(Position::BottomEnd),
            Some(value) => Err(invalid(format!("invalid chart:legend-position '{value}'"))),
        }
    }
}

impl Element {
    /// Borrow the first direct plot area from this chart element.
    #[must_use]
    pub fn plot_area(&self) -> Option<PlotArea<'_>> {
        self.children_of_kind(Kind::PlotArea)
            .next()
            .map(|element| PlotArea { element })
    }

    /// Borrow the first direct legend from this chart element.
    #[must_use]
    pub fn legend(&self) -> Option<Legend<'_>> {
        self.children_of_kind(Kind::Legend)
            .next()
            .map(|element| Legend { element })
    }
}

fn chart_attribute<'a>(element: &'a Element, local_name: &str) -> Option<&'a str> {
    element.attribute(Some(CHARTNS), local_name)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::read;

    const XML: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:chart><c:chart><c:legend c:legend-position="bottom-end"/><c:plot-area t:cell-range-address="Data.A1:C4" c:data-source-has-labels="both"><c:axis c:dimension="x" c:name="x"><c:categories t:cell-range-address="Data.A2:A4"/><c:grid c:class="major"/></c:axis><c:series c:values-cell-range-address="Data.B2:B4" c:label-cell-address="Data.B1"><c:domain t:cell-range-address="Data.A2:A4"/><c:data-point c:repeated="3"/></c:series></c:plot-area></c:chart></o:chart></o:body></o:document-content>"#;

    fn required<T>(value: Option<T>, description: &str) -> Result<T> {
        value.ok_or_else(|| Error::InvalidFormat(format!("missing {description}")))
    }

    #[test]
    fn provides_zero_copy_semantic_views() -> Result<()> {
        let chart = read(XML)?;
        let legend = required(chart.legend(), "chart legend")?;
        assert_eq!(legend.position()?, Position::BottomEnd);
        let plot = required(chart.plot_area(), "chart plot area")?;
        assert_eq!(plot.cell_range_address(), Some("Data.A1:C4"));
        assert_eq!(plot.data_source_labels()?, Labels::Both);
        let axis = required(plot.axes().next(), "plot axis")?;
        assert_eq!(axis.dimension()?, Dimension::X);
        assert_eq!(axis.categories_range(), Some("Data.A2:A4"));
        let grid = required(axis.grids().next(), "axis grid")?;
        assert_eq!(grid.class()?, Class::Major);
        let series = required(plot.series().next(), "plot series")?;
        assert_eq!(series.values_range(), Some("Data.B2:B4"));
        assert_eq!(series.domains().next(), Some("Data.A2:A4"));
        let data_point = required(series.data_points().next(), "series data point")?;
        assert_eq!(data_point.repeated()?, 3);
        Ok(())
    }

    #[test]
    fn validates_domains_at_access_time() -> Result<()> {
        let xml = XML
            .replace("c:dimension=\"x\"", "c:dimension=\"time\"")
            .replace("c:class=\"major\"", "c:class=\"micro\"")
            .replace("c:repeated=\"3\"", "c:repeated=\"0\"");
        let chart = read(&xml)?;
        let plot = required(chart.plot_area(), "chart plot area")?;
        let axis = required(plot.axes().next(), "plot axis")?;
        assert!(axis.dimension().is_err());
        let grid = required(axis.grids().next(), "axis grid")?;
        assert!(grid.class().is_err());
        let series = required(plot.series().next(), "plot series")?;
        let data_point = required(series.data_points().next(), "series data point")?;
        assert!(data_point.repeated().is_err());
        Ok(())
    }
}
