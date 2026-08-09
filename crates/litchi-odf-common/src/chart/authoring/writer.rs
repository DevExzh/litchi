//! Deterministic, bounded XML serialization and validation for ODF charts.

use super::data::{CachedCell, CachedRow, CachedTable, CachedValue};
use super::extensions::{ExtensionAttribute, ExtensionElement, Extensions};
use super::model::{
    AxisSpec, DataLabelSpec, DataPointSpec, Definition, LegendSpec, PlotAreaSpec, RegressionSpec,
    SeriesSpec, StyleElement, Text,
};
use crate::calculation::write;
use crate::chart::{ChartClass, Class, Dimension, Labels, Position};
use litchi_core::{Error, Result};
use std::collections::{BTreeMap, BTreeSet, HashSet};

pub(crate) const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(crate) const CHART_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
pub(crate) const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(crate) const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(crate) const STYLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(crate) const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(crate) const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const MAX_EXPANDED_CELLS: u64 = 16_777_216;

impl Definition {
    pub fn validate(&self) -> Result<()> {
        validate_optional_name(self.style_name.as_deref(), "chart style name")?;
        validate_optional_scalar(self.width.as_deref(), "chart width")?;
        validate_optional_scalar(self.height.as_deref(), "chart height")?;
        for text in [&self.title, &self.subtitle, &self.footer]
            .into_iter()
            .flatten()
        {
            validate_text(text)?;
        }
        if let Some(legend) = &self.legend {
            validate_optional_name(legend.style_name.as_deref(), "legend style name")?;
            validate_optional_scalar(legend.x.as_deref(), "legend x")?;
            validate_optional_scalar(legend.y.as_deref(), "legend y")?;
            validate_extensions(&legend.extensions)?;
        }
        validate_plot_area(&self.plot_area)?;
        if let Some(table) = &self.cached_table {
            validate_table(table)?;
        }
        if self.calculation_settings.is_some() {
            write(&mut String::new(), self.calculation_settings.as_ref())?;
        }
        validate_extensions(&self.extensions)
    }
}

/// Serialize a chart as canonical ODF 1.2 `content.xml`.
///
/// The output contains no executable behavior. Formula attributes in cached
/// cells are emitted only as escaped, opaque strings.
pub fn serialize_content(definition: &Definition) -> Result<String> {
    definition.validate()?;
    let namespaces = NamespaceMap::for_definition(definition)?;
    let mut out = String::with_capacity(4096);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    out.push_str("<office:document-content office:version=\"1.2\"");
    for (prefix, uri) in namespaces.declarations() {
        out.push_str(" xmlns:");
        out.push_str(prefix);
        out.push_str("=\"");
        escape_attribute(&mut out, uri)?;
        out.push('"');
    }
    out.push_str("><office:body><office:chart>");
    write_definition(&mut out, definition, &namespaces)?;
    write(&mut out, definition.calculation_settings.as_ref())?;
    out.push_str("</office:chart></office:body></office:document-content>");
    Ok(out)
}

fn write_definition(out: &mut String, chart: &Definition, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:chart");
    attr(out, "chart:class", chart.class.lexical())?;
    opt_attr(out, "chart:style-name", chart.style_name.as_deref())?;
    opt_attr(out, "svg:width", chart.width.as_deref())?;
    opt_attr(out, "svg:height", chart.height.as_deref())?;
    write_extension_attributes(out, &chart.extensions.attributes, ns)?;
    out.push('>');
    if let Some(value) = &chart.title {
        write_text_element(out, "title", value, ns)?;
    }
    if let Some(value) = &chart.subtitle {
        write_text_element(out, "subtitle", value, ns)?;
    }
    if let Some(value) = &chart.footer {
        write_text_element(out, "footer", value, ns)?;
    }
    if let Some(value) = &chart.legend {
        write_legend(out, value, ns)?;
    }
    write_plot_area(out, &chart.plot_area, ns)?;
    if let Some(table) = &chart.cached_table {
        write_table(out, table, ns)?;
    }
    write_extension_children(out, &chart.extensions.children, ns)?;
    out.push_str("</chart:chart>");
    Ok(())
}

fn write_text_element(
    out: &mut String,
    local: &str,
    value: &Text,
    ns: &NamespaceMap,
) -> Result<()> {
    out.push_str("<chart:");
    out.push_str(local);
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    opt_attr(out, "table:cell-range", value.cell_range.as_deref())?;
    opt_attr(out, "svg:x", value.x.as_deref())?;
    opt_attr(out, "svg:y", value.y.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    out.push('>');
    out.push_str("<text:p>");
    escape_text(out, &value.text)?;
    out.push_str("</text:p>");
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:");
    out.push_str(local);
    out.push('>');
    Ok(())
}

fn write_legend(out: &mut String, value: &LegendSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:legend");
    attr(
        out,
        "chart:legend-position",
        legend_position(value.position),
    )?;
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    opt_attr(out, "svg:x", value.x.as_deref())?;
    opt_attr(out, "svg:y", value.y.as_deref())?;
    opt_attr(out, "style:legend-expansion", value.expansion.as_deref())?;
    opt_attr(
        out,
        "style:legend-expansion-aspect-ratio",
        value.expansion_aspect_ratio.as_deref(),
    )?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    if value.title.is_none() && value.extensions.children.is_empty() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(title) = &value.title {
        out.push_str("<text:p>");
        escape_text(out, title)?;
        out.push_str("</text:p>");
    }
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:legend>");
    Ok(())
}

fn write_plot_area(out: &mut String, value: &PlotAreaSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:plot-area");
    opt_attr(
        out,
        "table:cell-range-address",
        value.cell_range_address.as_deref(),
    )?;
    if let Some(labels) = value.data_source_labels {
        attr(
            out,
            "chart:data-source-has-labels",
            data_source_labels(labels),
        )?;
    }
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    opt_attr(out, "svg:x", value.x.as_deref())?;
    opt_attr(out, "svg:y", value.y.as_deref())?;
    opt_attr(out, "svg:width", value.width.as_deref())?;
    opt_attr(out, "svg:height", value.height.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    out.push('>');
    for axis in &value.axes {
        write_axis(out, axis, ns)?;
    }
    for series in &value.series {
        write_series(out, series, ns)?;
    }
    if let Some(v) = &value.stock_gain_marker {
        write_style_element(out, "stock-gain-marker", v, ns)?;
    }
    if let Some(v) = &value.stock_loss_marker {
        write_style_element(out, "stock-loss-marker", v, ns)?;
    }
    if let Some(v) = &value.stock_range_line {
        write_style_element(out, "stock-range-line", v, ns)?;
    }
    if let Some(v) = &value.wall {
        write_style_element(out, "wall", v, ns)?;
    }
    if let Some(v) = &value.floor {
        write_style_element(out, "floor", v, ns)?;
    }
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:plot-area>");
    Ok(())
}

fn write_axis(out: &mut String, value: &AxisSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:axis");
    attr(out, "chart:dimension", axis_dimension(value.dimension))?;
    opt_attr(out, "chart:name", value.name.as_deref())?;
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    out.push('>');
    if let Some(title) = &value.title {
        write_text_element(out, "title", title, ns)?;
    }
    if let Some(range) = &value.categories_cell_range_address {
        out.push_str("<chart:categories");
        attr(out, "table:cell-range-address", range)?;
        out.push_str("/>");
    }
    for grid in &value.grids {
        out.push_str("<chart:grid");
        attr(out, "chart:class", grid_class(grid.class))?;
        opt_attr(out, "chart:style-name", grid.style_name.as_deref())?;
        write_extension_attributes(out, &grid.extensions.attributes, ns)?;
        if grid.extensions.children.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            write_extension_children(out, &grid.extensions.children, ns)?;
            out.push_str("</chart:grid>");
        }
    }
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:axis>");
    Ok(())
}

fn write_series(out: &mut String, value: &SeriesSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:series");
    opt_attr(out, "xml:id", value.xml_id.as_deref())?;
    if let Some(class) = &value.class {
        attr(out, "chart:class", class.lexical())?;
    }
    opt_attr(
        out,
        "chart:values-cell-range-address",
        value.values_cell_range_address.as_deref(),
    )?;
    opt_attr(
        out,
        "chart:label-cell-address",
        value.label_cell_address.as_deref(),
    )?;
    opt_attr(out, "chart:attached-axis", value.attached_axis.as_deref())?;
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    out.push('>');
    for domain in &value.domains {
        out.push_str("<chart:domain");
        attr(out, "table:cell-range-address", &domain.cell_range_address)?;
        write_extension_attributes(out, &domain.extensions.attributes, ns)?;
        if domain.extensions.children.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            write_extension_children(out, &domain.extensions.children, ns)?;
            out.push_str("</chart:domain>");
        }
    }
    if let Some(label) = &value.data_label {
        write_data_label(out, label, ns)?;
    }
    for point in &value.data_points {
        write_data_point(out, point, ns)?;
    }
    if let Some(v) = &value.mean_value {
        write_style_element(out, "mean-value", v, ns)?;
    }
    if let Some(v) = &value.error_indicator {
        write_style_element(out, "error-indicator", v, ns)?;
    }
    for curve in &value.regression_curves {
        write_regression(out, curve, ns)?;
    }
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:series>");
    Ok(())
}

pub fn serialize_axis_fragment(value: &AxisSpec) -> Result<String> {
    let mut definition = Definition::new(ChartClass::line());
    definition.plot_area.axes.push(value.clone());
    definition.validate()?;
    let namespaces = NamespaceMap::for_definition(&definition)?;
    let mut output = String::with_capacity(512);
    write_axis(&mut output, value, &namespaces)?;
    add_fragment_namespaces(&mut output, "<chart:axis".len(), &namespaces)?;
    Ok(output)
}

pub fn serialize_series_fragment(value: &SeriesSpec) -> Result<String> {
    let mut definition = Definition::new(value.class.clone().unwrap_or_else(ChartClass::line));
    if let Some(axis) = &value.attached_axis {
        let mut attached = AxisSpec::new(Dimension::Y);
        attached.name = Some(axis.clone());
        definition.plot_area.axes.push(attached);
    }
    definition.plot_area.series.push(value.clone());
    definition.validate()?;
    let namespaces = NamespaceMap::for_definition(&definition)?;
    let mut output = String::with_capacity(1024);
    write_series(&mut output, value, &namespaces)?;
    add_fragment_namespaces(&mut output, "<chart:series".len(), &namespaces)?;
    Ok(output)
}

fn add_fragment_namespaces(
    output: &mut String,
    position: usize,
    namespaces: &NamespaceMap,
) -> Result<()> {
    let mut declarations = String::new();
    for (prefix, uri) in namespaces.declarations() {
        declarations.push_str(" xmlns:");
        declarations.push_str(prefix);
        declarations.push_str("=\"");
        escape_attribute(&mut declarations, uri)?;
        declarations.push('"');
    }
    output.insert_str(position, &declarations);
    Ok(())
}

fn write_data_point(out: &mut String, value: &DataPointSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:data-point");
    if value.repeated != 1 {
        attr(out, "chart:repeated", &value.repeated.to_string())?;
    }
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    if value.label.is_none() && value.extensions.children.is_empty() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(label) = &value.label {
        write_data_label(out, label, ns)?;
    }
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:data-point>");
    Ok(())
}

fn write_data_label(out: &mut String, value: &DataLabelSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:data-label");
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    opt_attr(out, "svg:x", value.x.as_deref())?;
    opt_attr(out, "svg:y", value.y.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    if value.text.is_none() && value.extensions.children.is_empty() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(text) = &value.text {
        out.push_str("<text:p>");
        escape_text(out, text)?;
        out.push_str("</text:p>");
    }
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:data-label>");
    Ok(())
}

fn write_regression(out: &mut String, value: &RegressionSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:regression-curve");
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    if value.equation.is_none() && value.extensions.children.is_empty() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(eq) = &value.equation {
        out.push_str("<chart:equation");
        attr(out, "chart:display-equation", bool_xml(eq.display_equation))?;
        attr(out, "chart:display-r-square", bool_xml(eq.display_r_square))?;
        opt_attr(out, "chart:style-name", eq.style_name.as_deref())?;
        opt_attr(out, "svg:x", eq.x.as_deref())?;
        opt_attr(out, "svg:y", eq.y.as_deref())?;
        write_extension_attributes(out, &eq.extensions.attributes, ns)?;
        if eq.extensions.children.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            write_extension_children(out, &eq.extensions.children, ns)?;
            out.push_str("</chart:equation>");
        }
    }
    write_extension_children(out, &value.extensions.children, ns)?;
    out.push_str("</chart:regression-curve>");
    Ok(())
}

fn write_style_element(
    out: &mut String,
    local: &str,
    value: &StyleElement,
    ns: &NamespaceMap,
) -> Result<()> {
    out.push_str("<chart:");
    out.push_str(local);
    opt_attr(out, "chart:style-name", value.style_name.as_deref())?;
    write_extension_attributes(out, &value.extensions.attributes, ns)?;
    if value.extensions.children.is_empty() {
        out.push_str("/>");
    } else {
        out.push('>');
        write_extension_children(out, &value.extensions.children, ns)?;
        out.push_str("</chart:");
        out.push_str(local);
        out.push('>');
    }
    Ok(())
}

fn write_table(out: &mut String, table: &CachedTable, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<table:table");
    attr(out, "table:name", &table.name)?;
    write_extension_attributes(out, &table.extensions.attributes, ns)?;
    out.push('>');
    if table.header_columns > 0 {
        out.push_str("<table:table-header-columns><table:table-column");
        if table.header_columns != 1 {
            attr(
                out,
                "table:number-columns-repeated",
                &table.header_columns.to_string(),
            )?;
        }
        out.push_str("/></table:table-header-columns>");
    }
    let regular_columns = table.columns - table.header_columns;
    if regular_columns > 0 {
        out.push_str("<table:table-columns><table:table-column");
        if regular_columns != 1 {
            attr(
                out,
                "table:number-columns-repeated",
                &regular_columns.to_string(),
            )?;
        }
        out.push_str("/></table:table-columns>");
    }
    if !table.header_rows.is_empty() {
        out.push_str("<table:table-header-rows>");
        for row in &table.header_rows {
            write_row(out, row)?;
        }
        out.push_str("</table:table-header-rows>");
    }
    out.push_str("<table:table-rows>");
    for row in &table.rows {
        write_row(out, row)?;
    }
    out.push_str("</table:table-rows>");
    write_extension_children(out, &table.extensions.children, ns)?;
    out.push_str("</table:table>");
    Ok(())
}

fn write_row(out: &mut String, row: &CachedRow) -> Result<()> {
    out.push_str("<table:table-row");
    if row.repeated != 1 {
        attr(out, "table:number-rows-repeated", &row.repeated.to_string())?;
    }
    out.push('>');
    for cell in &row.cells {
        write_cell(out, cell)?;
    }
    out.push_str("</table:table-row>");
    Ok(())
}

fn write_cell(out: &mut String, cell: &CachedCell) -> Result<()> {
    out.push_str("<table:table-cell");
    if cell.repeated != 1 {
        attr(
            out,
            "table:number-columns-repeated",
            &cell.repeated.to_string(),
        )?;
    }
    opt_attr(out, "table:formula", cell.formula.as_deref())?;
    let text = match &cell.value {
        CachedValue::Empty => None,
        CachedValue::Float(value) => {
            attr(out, "office:value-type", "float")?;
            attr(out, "office:value", &value.to_string())?;
            None
        },
        CachedValue::Percentage(value) => {
            attr(out, "office:value-type", "percentage")?;
            attr(out, "office:value", &value.to_string())?;
            None
        },
        CachedValue::Currency { value, currency } => {
            attr(out, "office:value-type", "currency")?;
            attr(out, "office:value", &value.to_string())?;
            attr(out, "office:currency", currency)?;
            None
        },
        CachedValue::Boolean(value) => {
            attr(out, "office:value-type", "boolean")?;
            attr(out, "office:boolean-value", bool_xml(*value))?;
            None
        },
        CachedValue::Date(value) => {
            attr(out, "office:value-type", "date")?;
            attr(out, "office:date-value", value)?;
            None
        },
        CachedValue::Time(value) => {
            attr(out, "office:value-type", "time")?;
            attr(out, "office:time-value", value)?;
            None
        },
        CachedValue::String(value) => {
            attr(out, "office:value-type", "string")?;
            Some(value)
        },
    };
    if let Some(text) = text {
        out.push_str("><text:p>");
        escape_text(out, text)?;
        out.push_str("</text:p></table:table-cell>");
    } else {
        out.push_str("/>");
    }
    Ok(())
}

fn validate_plot_area(plot: &PlotAreaSpec) -> Result<()> {
    validate_optional_range(plot.cell_range_address.as_deref(), "plot-area range")?;
    validate_optional_name(plot.style_name.as_deref(), "plot-area style name")?;
    for scalar in [
        plot.x.as_deref(),
        plot.y.as_deref(),
        plot.width.as_deref(),
        plot.height.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_scalar(scalar, "plot-area coordinate")?;
    }
    let mut axis_names = HashSet::new();
    for axis in &plot.axes {
        validate_optional_name(axis.style_name.as_deref(), "axis style name")?;
        if let Some(name) = &axis.name {
            validate_name(name, "axis name")?;
            if !axis_names.insert(name.as_str()) {
                return invalid(format!("duplicate chart axis name '{name}'"));
            }
        }
        validate_optional_range(
            axis.categories_cell_range_address.as_deref(),
            "category range",
        )?;
        if let Some(title) = &axis.title {
            validate_text(title)?;
        }
        for grid in &axis.grids {
            validate_optional_name(grid.style_name.as_deref(), "grid style name")?;
            validate_extensions(&grid.extensions)?;
        }
        validate_extensions(&axis.extensions)?;
    }
    let mut series_ids = HashSet::new();
    for series in &plot.series {
        validate_optional_name(series.xml_id.as_deref(), "series xml:id")?;
        if let Some(xml_id) = series.xml_id.as_deref()
            && !series_ids.insert(xml_id)
        {
            return invalid(format!("duplicate chart series xml:id '{xml_id}'"));
        }
        validate_optional_range(
            series.values_cell_range_address.as_deref(),
            "series values range",
        )?;
        validate_optional_range(series.label_cell_address.as_deref(), "series label cell")?;
        validate_optional_name(series.style_name.as_deref(), "series style name")?;
        if let Some(axis) = &series.attached_axis {
            validate_name(axis, "attached axis")?;
            if !axis_names.contains(axis.as_str()) {
                return invalid(format!("series references unknown axis '{axis}'"));
            }
        }
        for domain in &series.domains {
            validate_range(&domain.cell_range_address, "domain range")?;
            validate_extensions(&domain.extensions)?;
        }
        let mut points = 0u64;
        for point in &series.data_points {
            if point.repeated == 0 {
                return invalid("chart:data-point repeat count must be nonzero");
            }
            points = points
                .checked_add(u64::from(point.repeated))
                .ok_or_else(|| {
                    Error::InvalidFormat("chart data-point count overflow".to_string())
                })?;
            if points > MAX_EXPANDED_CELLS {
                return invalid("expanded chart data-point count exceeds safety limit");
            }
            validate_optional_name(point.style_name.as_deref(), "data-point style name")?;
            if let Some(label) = &point.label {
                validate_data_label(label)?;
            }
            validate_extensions(&point.extensions)?;
        }
        if let Some(label) = &series.data_label {
            validate_data_label(label)?;
        }
        for style in [&series.mean_value, &series.error_indicator]
            .into_iter()
            .flatten()
        {
            validate_style_element(style)?;
        }
        for curve in &series.regression_curves {
            validate_optional_name(curve.style_name.as_deref(), "regression style name")?;
            if let Some(eq) = &curve.equation {
                validate_optional_name(eq.style_name.as_deref(), "equation style name")?;
                validate_extensions(&eq.extensions)?;
            }
            validate_extensions(&curve.extensions)?;
        }
        validate_extensions(&series.extensions)?;
    }
    for style in [
        &plot.wall,
        &plot.floor,
        &plot.stock_gain_marker,
        &plot.stock_loss_marker,
        &plot.stock_range_line,
    ]
    .into_iter()
    .flatten()
    {
        validate_style_element(style)?;
    }
    validate_extensions(&plot.extensions)
}

fn validate_table(table: &CachedTable) -> Result<()> {
    validate_name(&table.name, "cached table name")?;
    if table.columns == 0 {
        return invalid("cached table column count must be nonzero");
    }
    if table.header_columns > table.columns {
        return invalid("cached table header column count exceeds total columns");
    }
    let mut expanded_rows = 0u64;
    for row in table.header_rows.iter().chain(&table.rows) {
        if row.repeated == 0 {
            return invalid("cached table row repeat count must be nonzero");
        }
        expanded_rows = expanded_rows
            .checked_add(u64::from(row.repeated))
            .ok_or_else(|| Error::InvalidFormat("cached table row count overflow".to_string()))?;
        let mut columns = 0u64;
        for cell in &row.cells {
            if cell.repeated == 0 {
                return invalid("cached table cell repeat count must be nonzero");
            }
            columns = columns
                .checked_add(u64::from(cell.repeated))
                .ok_or_else(|| {
                    Error::InvalidFormat("cached table column count overflow".to_string())
                })?;
            if let CachedValue::Float(v)
            | CachedValue::Percentage(v)
            | CachedValue::Currency { value: v, .. } = &cell.value
                && !v.is_finite()
            {
                return invalid("cached numeric chart values must be finite");
            }
            if let CachedValue::Currency { currency, .. } = &cell.value {
                validate_name(currency, "currency code")?;
            }
            if let CachedValue::Date(v) | CachedValue::Time(v) = &cell.value {
                validate_scalar(v, "cached date/time value")?;
            }
            if let Some(formula) = &cell.formula {
                validate_scalar(formula, "cached formula")?;
            }
        }
        if columns > u64::from(table.columns) {
            return invalid("cached table row exceeds declared column count");
        }
    }
    if expanded_rows.saturating_mul(u64::from(table.columns)) > MAX_EXPANDED_CELLS {
        return invalid("expanded cached table exceeds safety limit");
    }
    validate_extensions(&table.extensions)
}

fn validate_text(value: &Text) -> Result<()> {
    validate_xml_chars(&value.text, "chart text")?;
    validate_optional_range(value.cell_range.as_deref(), "chart text cell range")?;
    validate_optional_name(value.style_name.as_deref(), "chart text style name")?;
    validate_optional_scalar(value.x.as_deref(), "chart text x")?;
    validate_optional_scalar(value.y.as_deref(), "chart text y")?;
    validate_extensions(&value.extensions)
}

fn validate_data_label(value: &DataLabelSpec) -> Result<()> {
    if let Some(text) = &value.text {
        validate_xml_chars(text, "data label text")?;
    }
    validate_optional_name(value.style_name.as_deref(), "data label style name")?;
    validate_extensions(&value.extensions)
}

fn validate_style_element(value: &StyleElement) -> Result<()> {
    validate_optional_name(value.style_name.as_deref(), "chart style name")?;
    validate_extensions(&value.extensions)
}

fn validate_extensions(value: &Extensions) -> Result<()> {
    let mut names = BTreeSet::new();
    for attribute in &value.attributes {
        validate_local_name(&attribute.local_name, "extension attribute")?;
        validate_xml_chars(&attribute.value, "extension attribute value")?;
        let key = (
            attribute.namespace_uri.as_deref().unwrap_or(""),
            attribute.local_name.as_str(),
        );
        if !names.insert(key) {
            return invalid("duplicate extension attribute expanded name");
        }
    }
    for child in &value.children {
        validate_extension_element(child, 0)?;
    }
    Ok(())
}

fn validate_extension_element(value: &ExtensionElement, depth: usize) -> Result<()> {
    if depth >= 128 {
        return invalid("extension subtree exceeds 128 levels");
    }
    validate_local_name(&value.local_name, "extension element")?;
    validate_xml_chars(&value.text, "extension text")?;
    let extensions = Extensions {
        attributes: value.attributes.clone(),
        children: Vec::new(),
    };
    validate_extensions(&extensions)?;
    for child in &value.children {
        validate_extension_element(child, depth + 1)?;
    }
    Ok(())
}

fn validate_optional_range(value: Option<&str>, kind: &str) -> Result<()> {
    if let Some(value) = value {
        validate_range(value, kind)?;
    }
    Ok(())
}
fn validate_range(value: &str, kind: &str) -> Result<()> {
    validate_scalar(value, kind)?;
    let mut quoted = false;
    let mut saw_dot = false;
    let mut saw_digit = false;
    let mut chars = value.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\'' {
            if quoted && chars.peek() == Some(&'\'') {
                chars.next();
            } else {
                quoted = !quoted;
            }
        } else if !quoted {
            if ch.is_whitespace() {
                return invalid(format!("{kind} contains unquoted whitespace"));
            }
            saw_dot |= ch == '.';
            saw_digit |= ch.is_ascii_digit();
        }
    }
    if quoted || !saw_dot || !saw_digit {
        return invalid(format!("invalid {kind} '{value}'"));
    }
    Ok(())
}

fn validate_optional_name(value: Option<&str>, kind: &str) -> Result<()> {
    if let Some(value) = value {
        validate_name(value, kind)?;
    }
    Ok(())
}
fn validate_name(value: &str, kind: &str) -> Result<()> {
    validate_local_name(value, kind)
}
fn validate_local_name(value: &str, kind: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return invalid(format!("{kind} cannot be empty"));
    };
    if !(first == '_' || first.is_alphabetic())
        || chars.any(|ch| !(ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric()))
    {
        return invalid(format!("invalid {kind} '{value}'"));
    }
    Ok(())
}
fn validate_optional_scalar(value: Option<&str>, kind: &str) -> Result<()> {
    if let Some(value) = value {
        validate_scalar(value, kind)?;
    }
    Ok(())
}
fn validate_scalar(value: &str, kind: &str) -> Result<()> {
    if value.trim().is_empty() {
        return invalid(format!("{kind} cannot be empty"));
    }
    validate_xml_chars(value, kind)
}
fn validate_xml_chars(value: &str, kind: &str) -> Result<()> {
    if value.chars().any(|ch| !matches!(ch as u32, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)) { return invalid(format!("{kind} contains a character forbidden by XML 1.0")); }
    Ok(())
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

struct NamespaceMap {
    by_uri: BTreeMap<String, String>,
    aliases: BTreeMap<String, String>,
}
impl NamespaceMap {
    fn for_definition(value: &Definition) -> Result<Self> {
        let mut uris = BTreeSet::new();
        collect_extensions(&value.extensions, &mut uris);
        collect_plot_extensions(&value.plot_area, &mut uris);
        if let Some(v) = &value.title {
            collect_extensions(&v.extensions, &mut uris);
        }
        if let Some(v) = &value.subtitle {
            collect_extensions(&v.extensions, &mut uris);
        }
        if let Some(v) = &value.footer {
            collect_extensions(&v.extensions, &mut uris);
        }
        if let Some(v) = &value.legend {
            collect_extensions(&v.extensions, &mut uris);
        }
        if let Some(v) = &value.cached_table {
            collect_extensions(&v.extensions, &mut uris);
        }
        let standards = [
            (OFFICE_NAMESPACE, "office"),
            (CHART_NAMESPACE, "chart"),
            (TABLE_NAMESPACE, "table"),
            (TEXT_NAMESPACE, "text"),
            (STYLE_NAMESPACE, "style"),
            (SVG_NAMESPACE, "svg"),
            (XLINK_NAMESPACE, "xlink"),
        ];
        let mut by_uri: BTreeMap<String, String> = standards
            .into_iter()
            .map(|(u, p)| (u.to_string(), p.to_string()))
            .collect();
        let mut index = 1usize;
        for uri in uris {
            if uri != XML_NAMESPACE && !by_uri.contains_key(&uri) {
                by_uri.insert(uri, format!("ns{index}"));
                index += 1;
            }
        }
        let mut aliases = BTreeMap::new();
        collect_class_alias(&value.class, &mut aliases)?;
        for series in &value.plot_area.series {
            if let Some(class) = &series.class {
                collect_class_alias(class, &mut aliases)?;
            }
        }
        for (prefix, uri) in &aliases {
            if let Some((existing_uri, _)) =
                by_uri.iter().find(|(_, canonical)| *canonical == prefix)
                && existing_uri != uri
            {
                return invalid(format!(
                    "chart class namespace prefix '{prefix}' conflicts with a standard prefix"
                ));
            }
        }
        Ok(Self { by_uri, aliases })
    }
    fn declarations(&self) -> impl Iterator<Item = (&str, &str)> {
        self.by_uri
            .iter()
            .map(|(u, p)| (p.as_str(), u.as_str()))
            .chain(self.aliases.iter().filter_map(|(prefix, uri)| {
                (!self.by_uri.values().any(|canonical| canonical == prefix))
                    .then_some((prefix.as_str(), uri.as_str()))
            }))
    }
    fn prefix(&self, uri: &str) -> Result<&str> {
        if uri == XML_NAMESPACE {
            Ok("xml")
        } else {
            self.by_uri.get(uri).map(String::as_str).ok_or_else(|| {
                Error::InvalidFormat(format!("unregistered extension namespace '{uri}'"))
            })
        }
    }
}

fn collect_class_alias(value: &ChartClass, aliases: &mut BTreeMap<String, String>) -> Result<()> {
    let Some((prefix, uri)) = value.namespace_alias() else {
        return Ok(());
    };
    match aliases.insert(prefix.to_owned(), uri.to_owned()) {
        Some(existing) if existing != uri => invalid(format!(
            "chart class namespace prefix '{prefix}' resolves to multiple URIs"
        )),
        _ => Ok(()),
    }
}

fn collect_plot_extensions(value: &PlotAreaSpec, uris: &mut BTreeSet<String>) {
    collect_extensions(&value.extensions, uris);
    for axis in &value.axes {
        collect_extensions(&axis.extensions, uris);
        if let Some(v) = &axis.title {
            collect_extensions(&v.extensions, uris);
        }
        for grid in &axis.grids {
            collect_extensions(&grid.extensions, uris);
        }
    }
    for series in &value.series {
        collect_extensions(&series.extensions, uris);
        for domain in &series.domains {
            collect_extensions(&domain.extensions, uris);
        }
        for point in &series.data_points {
            collect_extensions(&point.extensions, uris);
            if let Some(v) = &point.label {
                collect_extensions(&v.extensions, uris);
            }
        }
        if let Some(v) = &series.data_label {
            collect_extensions(&v.extensions, uris);
        }
        for v in [&series.mean_value, &series.error_indicator]
            .into_iter()
            .flatten()
        {
            collect_extensions(&v.extensions, uris);
        }
        for curve in &series.regression_curves {
            collect_extensions(&curve.extensions, uris);
            if let Some(eq) = &curve.equation {
                collect_extensions(&eq.extensions, uris);
            }
        }
    }
    for v in [
        &value.wall,
        &value.floor,
        &value.stock_gain_marker,
        &value.stock_loss_marker,
        &value.stock_range_line,
    ]
    .into_iter()
    .flatten()
    {
        collect_extensions(&v.extensions, uris);
    }
}
fn collect_extensions(value: &Extensions, uris: &mut BTreeSet<String>) {
    for a in &value.attributes {
        if let Some(uri) = &a.namespace_uri {
            uris.insert(uri.clone());
        }
    }
    for c in &value.children {
        collect_extension_element(c, uris);
    }
}
fn collect_extension_element(value: &ExtensionElement, uris: &mut BTreeSet<String>) {
    if let Some(uri) = &value.namespace_uri {
        uris.insert(uri.clone());
    }
    for a in &value.attributes {
        if let Some(uri) = &a.namespace_uri {
            uris.insert(uri.clone());
        }
    }
    for c in &value.children {
        collect_extension_element(c, uris);
    }
}

fn write_extension_attributes(
    out: &mut String,
    values: &[ExtensionAttribute],
    ns: &NamespaceMap,
) -> Result<()> {
    for value in values {
        out.push(' ');
        if let Some(uri) = &value.namespace_uri {
            out.push_str(ns.prefix(uri)?);
            out.push(':');
        }
        out.push_str(&value.local_name);
        out.push_str("=\"");
        escape_attribute(out, &value.value)?;
        out.push('"');
    }
    Ok(())
}
fn write_extension_children(
    out: &mut String,
    values: &[ExtensionElement],
    ns: &NamespaceMap,
) -> Result<()> {
    for value in values {
        write_extension_element(out, value, ns)?;
    }
    Ok(())
}
fn write_extension_element(
    out: &mut String,
    value: &ExtensionElement,
    ns: &NamespaceMap,
) -> Result<()> {
    out.push('<');
    if let Some(uri) = &value.namespace_uri {
        out.push_str(ns.prefix(uri)?);
        out.push(':');
    }
    out.push_str(&value.local_name);
    write_extension_attributes(out, &value.attributes, ns)?;
    if value.text.is_empty() && value.children.is_empty() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    escape_text(out, &value.text)?;
    write_extension_children(out, &value.children, ns)?;
    out.push_str("</");
    if let Some(uri) = &value.namespace_uri {
        out.push_str(ns.prefix(uri)?);
        out.push(':');
    }
    out.push_str(&value.local_name);
    out.push('>');
    Ok(())
}

fn attr(out: &mut String, name: &str, value: &str) -> Result<()> {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    escape_attribute(out, value)?;
    out.push('"');
    Ok(())
}
fn opt_attr(out: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        attr(out, name, value)?;
    }
    Ok(())
}
fn escape_text(out: &mut String, value: &str) -> Result<()> {
    validate_xml_chars(value, "XML text")?;
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
    Ok(())
}
fn escape_attribute(out: &mut String, value: &str) -> Result<()> {
    validate_xml_chars(value, "XML attribute")?;
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    Ok(())
}
fn bool_xml(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}
fn axis_dimension(value: Dimension) -> &'static str {
    match value {
        Dimension::X => "x",
        Dimension::Y => "y",
        Dimension::Z => "z",
    }
}
fn grid_class(value: Class) -> &'static str {
    match value {
        Class::Major => "major",
        Class::Minor => "minor",
    }
}
fn data_source_labels(value: Labels) -> &'static str {
    match value {
        Labels::None => "none",
        Labels::Row => "row",
        Labels::Column => "column",
        Labels::Both => "both",
    }
}
fn legend_position(value: Position) -> &'static str {
    match value {
        Position::Start => "start",
        Position::End => "end",
        Position::Top => "top",
        Position::Bottom => "bottom",
        Position::TopStart => "top-start",
        Position::TopEnd => "top-end",
        Position::BottomStart => "bottom-start",
        Position::BottomEnd => "bottom-end",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::ChartClass;
    use crate::chart::authoring::model::{EquationSpec, GridSpec};
    use crate::chart::read;

    fn sample() -> Definition {
        let mut chart = Definition::new(ChartClass::line());
        chart.title = Some(Text::new("Quarterly revenue"));
        chart.legend = Some(LegendSpec::default());
        let mut x = AxisSpec::new(Dimension::X);
        x.name = Some("x-axis".to_string());
        x.categories_cell_range_address = Some("local-table.A2:.A4".to_string());
        let mut y = AxisSpec::new(Dimension::Y);
        y.name = Some("y-axis".to_string());
        y.grids.push(GridSpec {
            class: Class::Major,
            style_name: None,
            extensions: Extensions::default(),
        });
        chart.plot_area.axes = vec![x, y];
        chart.plot_area.series.push(SeriesSpec {
            values_cell_range_address: Some("local-table.B2:.B4".to_string()),
            label_cell_address: Some("local-table.B1".to_string()),
            attached_axis: Some("y-axis".to_string()),
            data_points: vec![DataPointSpec {
                repeated: 3,
                ..DataPointSpec::default()
            }],
            mean_value: Some(StyleElement::default()),
            regression_curves: vec![RegressionSpec {
                equation: Some(EquationSpec {
                    display_equation: true,
                    display_r_square: true,
                    ..EquationSpec::default()
                }),
                ..RegressionSpec::default()
            }],
            ..SeriesSpec::default()
        });
        let mut table = CachedTable::new("local-table", 2);
        table.header_rows.push(CachedRow::new(vec![
            CachedCell::new(CachedValue::String("Quarter".into())),
            CachedCell::new(CachedValue::String("Revenue".into())),
        ]));
        table.rows.push(CachedRow::new(vec![
            CachedCell::new(CachedValue::String("Q1".into())),
            CachedCell::new(CachedValue::Float(10.0)),
        ]));
        chart.cached_table = Some(table);
        chart
    }

    #[test]
    fn canonical_content_roundtrip_retains_typed_data() {
        let content = serialize_content(&sample()).unwrap();
        let chart = read(&content).unwrap();
        assert_eq!(
            chart.attribute(Some(CHART_NAMESPACE), "class"),
            Some("chart:line")
        );
        assert!(
            chart
                .children()
                .iter()
                .any(|node| node.namespace_uri() == Some(TABLE_NAMESPACE)
                    && node.local_name() == "table")
        );
        assert_eq!(chart.plot_area().unwrap().series().count(), 1);
    }

    #[test]
    fn cached_formula_is_escaped_and_inert() {
        let mut definition = sample();
        definition.cached_table.as_mut().unwrap().rows[0].cells[1].formula =
            Some("of:=SUM([.B2:.B4])&\"x\"".into());
        let content = serialize_content(&definition).unwrap();
        assert!(content.contains("table:formula=\"of:=SUM([.B2:.B4])&amp;&quot;x&quot;\""));
        assert_eq!(
            read(&content).unwrap().all_text(),
            "Quarterly revenueQuarterRevenueQ1"
        );
    }

    #[test]
    fn rejects_dangling_axis_and_zero_counts() {
        let mut definition = sample();
        definition.plot_area.series[0].attached_axis = Some("missing".into());
        assert!(serialize_content(&definition).is_err());
        definition.plot_area.series[0].attached_axis = Some("y-axis".into());
        definition.plot_area.series[0].data_points[0].repeated = 0;
        assert!(serialize_content(&definition).is_err());
    }

    #[test]
    fn extension_namespaces_are_stable_and_retained() {
        let mut definition = Definition::new(ChartClass::line());
        definition.extensions.attributes.push(ExtensionAttribute {
            namespace_uri: Some("urn:z".into()),
            local_name: "zeta".into(),
            value: "1".into(),
        });
        definition.extensions.attributes.push(ExtensionAttribute {
            namespace_uri: Some("urn:a".into()),
            local_name: "alpha".into(),
            value: "2".into(),
        });
        definition.extensions.children.push(ExtensionElement {
            namespace_uri: Some("urn:z".into()),
            local_name: "extension".into(),
            attributes: Vec::new(),
            text: "opaque".into(),
            children: Vec::new(),
        });
        let first = serialize_content(&definition).unwrap();
        let second = serialize_content(&definition).unwrap();
        assert_eq!(first, second);
        assert!(first.contains("xmlns:ns1=\"urn:a\""));
        assert!(first.contains("xmlns:ns2=\"urn:z\""));
        assert!(first.contains("<ns2:extension>opaque</ns2:extension>"));
        assert_eq!(
            read(&first)
                .unwrap()
                .children()
                .last()
                .unwrap()
                .namespace_uri(),
            Some("urn:z")
        );
    }
}
