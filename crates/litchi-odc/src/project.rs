//! Strict projection of canonical opened chart XML into the typed definition.

use crate::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, DataLabelSpec, DataPointSpec,
    Definition, DomainSpec, EquationSpec, Extensions, GridSpec, LegendSpec, RegressionSpec,
    SeriesSpec, StyleElement, Text,
};
use litchi_core::{Error, Result};
use litchi_odf_common::chart::{Element, Kind};

const CHART: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

pub(crate) fn definition(chart: &crate::Chart) -> Result<Definition> {
    let root = chart.chart();
    let mut value = Definition::new(root.chart_class()?);
    value.style_name = attribute(root, CHART, "style-name");
    value.width = attribute(root, SVG, "width");
    value.height = attribute(root, SVG, "height");
    value.title = child(root, Kind::Title).map(text);
    value.subtitle = child(root, Kind::Subtitle).map(text);
    value.footer = child(root, Kind::Footer).map(text);
    value.legend = root.legend().map(legend).transpose()?;
    value.plot_area = plot_area(
        root.plot_area()
            .ok_or_else(|| invalid("ODC chart has no projectable plot area"))?,
    )?;
    value.cached_table = root
        .children()
        .iter()
        .find(|element| element.namespace_uri() == Some(TABLE) && element.local_name() == "table")
        .map(cached_table)
        .transpose()?;
    crate::validation::validate_definition(&value, chart.limits())?;
    let canonical = crate::serialize_content_with_limits(&value, chart.limits())?;
    if canonical != chart.content_xml() {
        return Err(Error::Unsupported(
            "opened ODC contains lexical or semantic content outside the lossless typed edit surface"
                .into(),
        ));
    }
    Ok(value)
}

fn text(element: &Element) -> Text {
    Text {
        text: element.all_text(),
        cell_range: attribute(element, TABLE, "cell-range"),
        style_name: attribute(element, CHART, "style-name"),
        x: attribute(element, SVG, "x"),
        y: attribute(element, SVG, "y"),
        ..Text::default()
    }
}

fn legend(value: litchi_odf_common::chart::Legend<'_>) -> Result<LegendSpec> {
    let element = value.element();
    Ok(LegendSpec {
        position: value.position()?,
        style_name: attribute(element, CHART, "style-name"),
        title: element
            .children()
            .iter()
            .find(|child| child.local_name() == "p")
            .map(Element::all_text),
        x: attribute(element, SVG, "x"),
        y: attribute(element, SVG, "y"),
        expansion: attribute(element, STYLE, "legend-expansion"),
        expansion_aspect_ratio: attribute(element, STYLE, "legend-expansion-aspect-ratio"),
        ..LegendSpec::default()
    })
}

fn plot_area(value: litchi_odf_common::chart::PlotArea<'_>) -> Result<crate::PlotAreaSpec> {
    let element = value.element();
    let mut plot = crate::PlotAreaSpec {
        cell_range_address: value.cell_range_address().map(str::to_owned),
        data_source_labels: element
            .attribute(Some(CHART), "data-source-has-labels")
            .map(|_| value.data_source_labels())
            .transpose()?,
        style_name: attribute(element, CHART, "style-name"),
        x: attribute(element, SVG, "x"),
        y: attribute(element, SVG, "y"),
        width: attribute(element, SVG, "width"),
        height: attribute(element, SVG, "height"),
        ..crate::PlotAreaSpec::default()
    };
    plot.axes = value.axes().map(axis).collect::<Result<Vec<_>>>()?;
    plot.series = value.series().map(series).collect::<Result<Vec<_>>>()?;
    plot.wall = child(element, Kind::Wall).map(style_element);
    plot.floor = child(element, Kind::Floor).map(style_element);
    plot.stock_gain_marker = child(element, Kind::StockGainMarker).map(style_element);
    plot.stock_loss_marker = child(element, Kind::StockLossMarker).map(style_element);
    plot.stock_range_line = child(element, Kind::StockRangeLine).map(style_element);
    Ok(plot)
}

fn axis(value: litchi_odf_common::chart::Axis<'_>) -> Result<AxisSpec> {
    let element = value.element();
    Ok(AxisSpec {
        dimension: value.dimension()?,
        name: value.name().map(str::to_owned),
        style_name: value.style_name().map(str::to_owned),
        title: child(element, Kind::Title).map(text),
        categories_cell_range_address: value.categories_range().map(str::to_owned),
        grids: value
            .grids()
            .map(|grid| {
                Ok(GridSpec {
                    class: grid.class()?,
                    style_name: attribute(grid.element(), CHART, "style-name"),
                    extensions: Extensions::default(),
                })
            })
            .collect::<Result<Vec<_>>>()?,
        extensions: Extensions::default(),
    })
}

fn series(value: litchi_odf_common::chart::Series<'_>) -> Result<SeriesSpec> {
    let element = value.element();
    let mut output = SeriesSpec {
        xml_id: value.xml_id().map(str::to_owned),
        class: element
            .attribute(Some(CHART), "class")
            .map(|_| element.chart_class())
            .transpose()?,
        values_cell_range_address: value.values_range().map(str::to_owned),
        label_cell_address: value.label_cell().map(str::to_owned),
        attached_axis: value.attached_axis().map(str::to_owned),
        style_name: attribute(element, CHART, "style-name"),
        domains: value
            .domains()
            .map(|address| DomainSpec {
                cell_range_address: address.to_owned(),
                extensions: Extensions::default(),
            })
            .collect(),
        data_points: value
            .data_points()
            .map(|point| {
                Ok(DataPointSpec {
                    repeated: point.repeated()?,
                    style_name: point.style_name().map(str::to_owned),
                    label: child(point.element(), Kind::DataLabel).map(data_label),
                    extensions: Extensions::default(),
                })
            })
            .collect::<Result<Vec<_>>>()?,
        data_label: child(element, Kind::DataLabel).map(data_label),
        mean_value: child(element, Kind::MeanValue).map(style_element),
        error_indicator: child(element, Kind::ErrorIndicator).map(style_element),
        ..SeriesSpec::default()
    };
    output.regression_curves = element
        .children_of_kind(Kind::RegressionCurve)
        .map(regression)
        .collect();
    Ok(output)
}

fn data_label(element: &Element) -> DataLabelSpec {
    let text = element.all_text();
    DataLabelSpec {
        text: (!text.is_empty()).then_some(text),
        style_name: attribute(element, CHART, "style-name"),
        x: attribute(element, SVG, "x"),
        y: attribute(element, SVG, "y"),
        ..DataLabelSpec::default()
    }
}

fn regression(element: &Element) -> RegressionSpec {
    RegressionSpec {
        style_name: attribute(element, CHART, "style-name"),
        equation: child(element, Kind::Equation).map(|equation| EquationSpec {
            display_equation: boolean(equation, "display-equation"),
            display_r_square: boolean(equation, "display-r-square"),
            style_name: attribute(equation, CHART, "style-name"),
            x: attribute(equation, SVG, "x"),
            y: attribute(equation, SVG, "y"),
            ..EquationSpec::default()
        }),
        ..RegressionSpec::default()
    }
}

fn style_element(element: &Element) -> StyleElement {
    StyleElement {
        style_name: attribute(element, CHART, "style-name"),
        ..StyleElement::default()
    }
}

fn cached_table(element: &Element) -> Result<CachedTable> {
    let name = required_attribute(element, TABLE, "name")?;
    let header_columns = repeated_columns(element, "table-header-columns")?;
    let columns = header_columns
        .checked_add(repeated_columns(element, "table-columns")?)
        .ok_or_else(|| invalid("ODC cached-table column count overflow"))?;
    let mut table = CachedTable::new(name, columns);
    table.header_columns = header_columns;
    table.header_rows = rows(element, "table-header-rows")?;
    table.rows = rows(element, "table-rows")?;
    Ok(table)
}

fn repeated_columns(element: &Element, container: &str) -> Result<u32> {
    element
        .children()
        .iter()
        .find(|child| child.namespace_uri() == Some(TABLE) && child.local_name() == container)
        .map_or(Ok(0), |value| {
            value.children().iter().try_fold(0u32, |total, column| {
                total
                    .checked_add(parse_u32(column, TABLE, "number-columns-repeated", 1)?)
                    .ok_or_else(|| invalid("ODC cached-table column count overflow"))
            })
        })
}

fn rows(element: &Element, container_name: &str) -> Result<Vec<CachedRow>> {
    let Some(container_element) = element
        .children()
        .iter()
        .find(|child| child.namespace_uri() == Some(TABLE) && child.local_name() == container_name)
    else {
        return Ok(Vec::new());
    };
    container_element
        .children()
        .iter()
        .filter(|row| row.namespace_uri() == Some(TABLE) && row.local_name() == "table-row")
        .map(|row| {
            Ok(CachedRow {
                cells: row
                    .children()
                    .iter()
                    .filter(|cell| {
                        cell.namespace_uri() == Some(TABLE) && cell.local_name() == "table-cell"
                    })
                    .map(cached_cell)
                    .collect::<Result<Vec<_>>>()?,
                repeated: parse_u32(row, TABLE, "number-rows-repeated", 1)?,
            })
        })
        .collect()
}

fn cached_cell(element: &Element) -> Result<CachedCell> {
    let value_type = element.attribute(Some(OFFICE), "value-type");
    let value = match value_type {
        None => CachedValue::Empty,
        Some("float") => CachedValue::Float(parse_f64(element, "value")?),
        Some("percentage") => CachedValue::Percentage(parse_f64(element, "value")?),
        Some("currency") => CachedValue::Currency {
            value: parse_f64(element, "value")?,
            currency: required_attribute(element, OFFICE, "currency")?,
        },
        Some("boolean") => {
            CachedValue::Boolean(required_attribute(element, OFFICE, "boolean-value")? == "true")
        },
        Some("date") => CachedValue::Date(required_attribute(element, OFFICE, "date-value")?),
        Some("time") => CachedValue::Time(required_attribute(element, OFFICE, "time-value")?),
        Some("string") => CachedValue::String(element.all_text()),
        Some(other) => {
            return Err(invalid(format!(
                "unsupported ODC cached value type '{other}'"
            )));
        },
    };
    Ok(CachedCell {
        value,
        formula: attribute(element, TABLE, "formula"),
        repeated: parse_u32(element, TABLE, "number-columns-repeated", 1)?,
    })
}

fn parse_f64(element: &Element, local: &str) -> Result<f64> {
    let lexical = required_attribute(element, OFFICE, local)?;
    lexical
        .parse()
        .map_err(|error| invalid(format!("invalid ODC cached number '{lexical}': {error}")))
}

fn parse_u32(element: &Element, namespace: &str, local: &str, default: u32) -> Result<u32> {
    element
        .attribute(Some(namespace), local)
        .map_or(Ok(default), |value| {
            value
                .parse()
                .map_err(|error| invalid(format!("invalid ODC repeat count '{value}': {error}")))
        })
}

fn boolean(element: &Element, local: &str) -> bool {
    element.attribute(Some(CHART), local) == Some("true")
}

fn child(element: &Element, kind: Kind) -> Option<&Element> {
    element.children_of_kind(kind).next()
}

fn attribute(element: &Element, namespace: &str, local: &str) -> Option<String> {
    element.attribute(Some(namespace), local).map(str::to_owned)
}

fn required_attribute(element: &Element, namespace: &str, local: &str) -> Result<String> {
    attribute(element, namespace, local)
        .ok_or_else(|| invalid(format!("ODC element requires {local}")))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
