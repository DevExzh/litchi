//! Typed creation and canonical serialization for standalone ODF charts.

use super::document::{ChartDocument, ChartElement, parse_chart_content};
use super::semantic::{
    ChartAxisDimension, ChartDataSourceLabels, ChartGridClass, ChartLegendPosition,
};
use crate::{constants, core::PackageWriter};
use litchi_core::{Error, Result};
use litchi_odf_common::calculation::{Settings, write};
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

/// An extension attribute retained by expanded XML name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExtensionAttribute {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub value: String,
}

/// An extension subtree retained without interpreting vendor behavior.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExtensionElement {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub attributes: Vec<ChartExtensionAttribute>,
    pub text: String,
    pub children: Vec<ChartExtensionElement>,
}

impl ChartExtensionElement {
    /// Clone a retained read-only element into an owned extension subtree.
    pub fn from_retained(element: &ChartElement) -> Self {
        Self {
            namespace_uri: element.namespace_uri().map(str::to_string),
            local_name: element.local_name().to_string(),
            attributes: element
                .attributes()
                .iter()
                .map(|attribute| ChartExtensionAttribute {
                    namespace_uri: attribute.namespace_uri().map(str::to_string),
                    local_name: attribute.local_name().to_string(),
                    value: attribute.value().to_string(),
                })
                .collect(),
            text: element.text().to_string(),
            children: element.children().iter().map(Self::from_retained).collect(),
        }
    }
}

/// Unknown attributes and child elements attached to a typed chart node.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartExtensions {
    pub attributes: Vec<ChartExtensionAttribute>,
    pub children: Vec<ChartExtensionElement>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartText {
    pub text: String,
    pub cell_range: Option<String>,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub extensions: ChartExtensions,
}

impl ChartText {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartLegendSpec {
    pub position: ChartLegendPosition,
    pub style_name: Option<String>,
    pub title: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub expansion: Option<String>,
    pub expansion_aspect_ratio: Option<String>,
    pub extensions: ChartExtensions,
}

impl Default for ChartLegendSpec {
    fn default() -> Self {
        Self {
            position: ChartLegendPosition::End,
            style_name: None,
            title: None,
            x: None,
            y: None,
            expansion: None,
            expansion_aspect_ratio: None,
            extensions: ChartExtensions::default(),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartStyleElement {
    pub style_name: Option<String>,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartGridSpec {
    pub class: ChartGridClass,
    pub style_name: Option<String>,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartDataLabelSpec {
    pub text: Option<String>,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartDataPointSpec {
    pub repeated: u32,
    pub style_name: Option<String>,
    pub label: Option<ChartDataLabelSpec>,
    pub extensions: ChartExtensions,
}

impl Default for ChartDataPointSpec {
    fn default() -> Self {
        Self {
            repeated: 1,
            style_name: None,
            label: None,
            extensions: ChartExtensions::default(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartDomainSpec {
    pub cell_range_address: String,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartEquationSpec {
    pub display_equation: bool,
    pub display_r_square: bool,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartRegressionSpec {
    pub style_name: Option<String>,
    pub equation: Option<ChartEquationSpec>,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartSeriesSpec {
    pub xml_id: Option<String>,
    pub class: Option<String>,
    pub values_cell_range_address: Option<String>,
    pub label_cell_address: Option<String>,
    pub attached_axis: Option<String>,
    pub style_name: Option<String>,
    pub domains: Vec<ChartDomainSpec>,
    pub data_points: Vec<ChartDataPointSpec>,
    pub data_label: Option<ChartDataLabelSpec>,
    pub mean_value: Option<ChartStyleElement>,
    pub error_indicator: Option<ChartStyleElement>,
    pub regression_curves: Vec<ChartRegressionSpec>,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartAxisSpec {
    pub dimension: ChartAxisDimension,
    pub name: Option<String>,
    pub style_name: Option<String>,
    pub title: Option<ChartText>,
    pub categories_cell_range_address: Option<String>,
    pub grids: Vec<ChartGridSpec>,
    pub extensions: ChartExtensions,
}

impl ChartAxisSpec {
    pub fn new(dimension: ChartAxisDimension) -> Self {
        Self {
            dimension,
            name: None,
            style_name: None,
            title: None,
            categories_cell_range_address: None,
            grids: Vec::new(),
            extensions: ChartExtensions::default(),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartPlotAreaSpec {
    pub cell_range_address: Option<String>,
    pub data_source_labels: Option<ChartDataSourceLabels>,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    pub axes: Vec<ChartAxisSpec>,
    pub series: Vec<ChartSeriesSpec>,
    pub wall: Option<ChartStyleElement>,
    pub floor: Option<ChartStyleElement>,
    pub stock_gain_marker: Option<ChartStyleElement>,
    pub stock_loss_marker: Option<ChartStyleElement>,
    pub stock_range_line: Option<ChartStyleElement>,
    pub extensions: ChartExtensions,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub enum ChartCachedValue {
    #[default]
    Empty,
    Float(f64),
    Percentage(f64),
    Currency {
        value: f64,
        currency: String,
    },
    Boolean(bool),
    Date(String),
    Time(String),
    String(String),
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartCachedCell {
    pub value: ChartCachedValue,
    /// An OpenDocument formula stored as inert text; this crate never evaluates it.
    pub formula: Option<String>,
    pub repeated: u32,
}

impl ChartCachedCell {
    pub fn new(value: ChartCachedValue) -> Self {
        Self {
            value,
            formula: None,
            repeated: 1,
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartCachedRow {
    pub cells: Vec<ChartCachedCell>,
    pub repeated: u32,
}

impl ChartCachedRow {
    pub fn new(cells: Vec<ChartCachedCell>) -> Self {
        Self { cells, repeated: 1 }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartCachedTable {
    pub name: String,
    pub columns: u32,
    pub header_columns: u32,
    pub header_rows: Vec<ChartCachedRow>,
    pub rows: Vec<ChartCachedRow>,
    pub extensions: ChartExtensions,
}

impl ChartCachedTable {
    pub fn new(name: impl Into<String>, columns: u32) -> Self {
        Self {
            name: name.into(),
            columns,
            header_columns: 0,
            header_rows: Vec::new(),
            rows: Vec::new(),
            extensions: ChartExtensions::default(),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartDefinition {
    pub class: String,
    pub style_name: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    pub title: Option<ChartText>,
    pub subtitle: Option<ChartText>,
    pub footer: Option<ChartText>,
    pub legend: Option<ChartLegendSpec>,
    pub plot_area: ChartPlotAreaSpec,
    pub cached_table: Option<ChartCachedTable>,
    /// Inert cached-formula recalculation metadata; never executed by this crate.
    pub calculation_settings: Option<Settings>,
    pub extensions: ChartExtensions,
}

impl ChartDefinition {
    pub fn new(class: impl Into<String>) -> Self {
        Self {
            class: class.into(),
            style_name: None,
            width: None,
            height: None,
            title: None,
            subtitle: None,
            footer: None,
            legend: None,
            plot_area: ChartPlotAreaSpec::default(),
            cached_table: None,
            calculation_settings: None,
            extensions: ChartExtensions::default(),
        }
    }

    pub fn validate(&self) -> Result<()> {
        validate_qname(&self.class, "chart class")?;
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

impl ChartDocument {
    /// Create a new packaged `.odc` document from a typed chart definition.
    pub fn create(definition: &ChartDefinition) -> Result<Self> {
        Self::create_with_mimetype(definition, constants::ODF_CHART)
    }

    /// Create a new packaged `.otc` chart template.
    pub fn create_template(definition: &ChartDefinition) -> Result<Self> {
        Self::create_with_mimetype(definition, constants::ODF_CHART_TEMPLATE)
    }

    fn create_with_mimetype(definition: &ChartDefinition, mimetype: &str) -> Result<Self> {
        let content = serialize_chart_content(definition)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype)?;
        writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
        Self::from_bytes(writer.finish_to_bytes()?)
    }

    /// Replace the typed chart content while preserving safe package entries.
    pub fn set_definition(&mut self, definition: &ChartDefinition) -> Result<()> {
        let content = serialize_chart_content(definition)?;
        let parsed = parse_chart_content(&content)?;
        self.package.replace_content_xml(content)?;
        self.chart = parsed;
        Ok(())
    }
}

/// Serialize a chart as canonical ODF 1.2 `content.xml`.
///
/// The output contains no executable behavior. Formula attributes in cached
/// cells are emitted only as escaped, opaque strings.
pub fn serialize_chart_content(definition: &ChartDefinition) -> Result<String> {
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

fn write_definition(out: &mut String, chart: &ChartDefinition, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:chart");
    attr(out, "chart:class", &chart.class)?;
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
    value: &ChartText,
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

fn write_legend(out: &mut String, value: &ChartLegendSpec, ns: &NamespaceMap) -> Result<()> {
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

fn write_plot_area(out: &mut String, value: &ChartPlotAreaSpec, ns: &NamespaceMap) -> Result<()> {
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

fn write_axis(out: &mut String, value: &ChartAxisSpec, ns: &NamespaceMap) -> Result<()> {
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

fn write_series(out: &mut String, value: &ChartSeriesSpec, ns: &NamespaceMap) -> Result<()> {
    out.push_str("<chart:series");
    opt_attr(out, "xml:id", value.xml_id.as_deref())?;
    opt_attr(out, "chart:class", value.class.as_deref())?;
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

pub(crate) fn serialize_chart_axis_fragment(value: &ChartAxisSpec) -> Result<String> {
    let mut definition = ChartDefinition::new("chart:line");
    definition.plot_area.axes.push(value.clone());
    definition.validate()?;
    let namespaces = NamespaceMap::for_definition(&definition)?;
    let mut output = String::with_capacity(512);
    write_axis(&mut output, value, &namespaces)?;
    add_fragment_namespaces(&mut output, "<chart:axis".len(), &namespaces)?;
    Ok(output)
}

pub(crate) fn serialize_chart_series_fragment(value: &ChartSeriesSpec) -> Result<String> {
    let mut definition = ChartDefinition::new(
        value
            .class
            .clone()
            .unwrap_or_else(|| "chart:line".to_string()),
    );
    if let Some(axis) = &value.attached_axis {
        let mut attached = ChartAxisSpec::new(ChartAxisDimension::Y);
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

fn write_data_point(out: &mut String, value: &ChartDataPointSpec, ns: &NamespaceMap) -> Result<()> {
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

fn write_data_label(out: &mut String, value: &ChartDataLabelSpec, ns: &NamespaceMap) -> Result<()> {
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

fn write_regression(
    out: &mut String,
    value: &ChartRegressionSpec,
    ns: &NamespaceMap,
) -> Result<()> {
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
    value: &ChartStyleElement,
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

fn write_table(out: &mut String, table: &ChartCachedTable, ns: &NamespaceMap) -> Result<()> {
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

fn write_row(out: &mut String, row: &ChartCachedRow) -> Result<()> {
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

fn write_cell(out: &mut String, cell: &ChartCachedCell) -> Result<()> {
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
        ChartCachedValue::Empty => None,
        ChartCachedValue::Float(value) => {
            attr(out, "office:value-type", "float")?;
            attr(out, "office:value", &value.to_string())?;
            None
        },
        ChartCachedValue::Percentage(value) => {
            attr(out, "office:value-type", "percentage")?;
            attr(out, "office:value", &value.to_string())?;
            None
        },
        ChartCachedValue::Currency { value, currency } => {
            attr(out, "office:value-type", "currency")?;
            attr(out, "office:value", &value.to_string())?;
            attr(out, "office:currency", currency)?;
            None
        },
        ChartCachedValue::Boolean(value) => {
            attr(out, "office:value-type", "boolean")?;
            attr(out, "office:boolean-value", bool_xml(*value))?;
            None
        },
        ChartCachedValue::Date(value) => {
            attr(out, "office:value-type", "date")?;
            attr(out, "office:date-value", value)?;
            None
        },
        ChartCachedValue::Time(value) => {
            attr(out, "office:value-type", "time")?;
            attr(out, "office:time-value", value)?;
            None
        },
        ChartCachedValue::String(value) => {
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

fn validate_plot_area(plot: &ChartPlotAreaSpec) -> Result<()> {
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
        if let Some(class) = &series.class {
            validate_qname(class, "series class")?;
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

fn validate_table(table: &ChartCachedTable) -> Result<()> {
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
            if let ChartCachedValue::Float(v)
            | ChartCachedValue::Percentage(v)
            | ChartCachedValue::Currency { value: v, .. } = &cell.value
                && !v.is_finite()
            {
                return invalid("cached numeric chart values must be finite");
            }
            if let ChartCachedValue::Currency { currency, .. } = &cell.value {
                validate_name(currency, "currency code")?;
            }
            if let ChartCachedValue::Date(v) | ChartCachedValue::Time(v) = &cell.value {
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

fn validate_text(value: &ChartText) -> Result<()> {
    validate_xml_chars(&value.text, "chart text")?;
    validate_optional_range(value.cell_range.as_deref(), "chart text cell range")?;
    validate_optional_name(value.style_name.as_deref(), "chart text style name")?;
    validate_optional_scalar(value.x.as_deref(), "chart text x")?;
    validate_optional_scalar(value.y.as_deref(), "chart text y")?;
    validate_extensions(&value.extensions)
}

fn validate_data_label(value: &ChartDataLabelSpec) -> Result<()> {
    if let Some(text) = &value.text {
        validate_xml_chars(text, "data label text")?;
    }
    validate_optional_name(value.style_name.as_deref(), "data label style name")?;
    validate_extensions(&value.extensions)
}

fn validate_style_element(value: &ChartStyleElement) -> Result<()> {
    validate_optional_name(value.style_name.as_deref(), "chart style name")?;
    validate_extensions(&value.extensions)
}

fn validate_extensions(value: &ChartExtensions) -> Result<()> {
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

fn validate_extension_element(value: &ChartExtensionElement, depth: usize) -> Result<()> {
    if depth >= 128 {
        return invalid("extension subtree exceeds 128 levels");
    }
    validate_local_name(&value.local_name, "extension element")?;
    validate_xml_chars(&value.text, "extension text")?;
    let extensions = ChartExtensions {
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

fn validate_qname(value: &str, kind: &str) -> Result<()> {
    let mut parts = value.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if first.is_empty() || parts.next().is_some() {
        return invalid(format!("invalid {kind} '{value}'"));
    }
    validate_local_name(first, kind)?;
    if let Some(second) = second {
        validate_local_name(second, kind)?;
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
}
impl NamespaceMap {
    fn for_definition(value: &ChartDefinition) -> Result<Self> {
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
        Ok(Self { by_uri })
    }
    fn declarations(&self) -> impl Iterator<Item = (&str, &str)> {
        self.by_uri.iter().map(|(u, p)| (p.as_str(), u.as_str()))
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

fn collect_plot_extensions(value: &ChartPlotAreaSpec, uris: &mut BTreeSet<String>) {
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
fn collect_extensions(value: &ChartExtensions, uris: &mut BTreeSet<String>) {
    for a in &value.attributes {
        if let Some(uri) = &a.namespace_uri {
            uris.insert(uri.clone());
        }
    }
    for c in &value.children {
        collect_extension_element(c, uris);
    }
}
fn collect_extension_element(value: &ChartExtensionElement, uris: &mut BTreeSet<String>) {
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
    values: &[ChartExtensionAttribute],
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
    values: &[ChartExtensionElement],
    ns: &NamespaceMap,
) -> Result<()> {
    for value in values {
        write_extension_element(out, value, ns)?;
    }
    Ok(())
}
fn write_extension_element(
    out: &mut String,
    value: &ChartExtensionElement,
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
fn axis_dimension(value: ChartAxisDimension) -> &'static str {
    match value {
        ChartAxisDimension::X => "x",
        ChartAxisDimension::Y => "y",
        ChartAxisDimension::Z => "z",
    }
}
fn grid_class(value: ChartGridClass) -> &'static str {
    match value {
        ChartGridClass::Major => "major",
        ChartGridClass::Minor => "minor",
    }
}
fn data_source_labels(value: ChartDataSourceLabels) -> &'static str {
    match value {
        ChartDataSourceLabels::None => "none",
        ChartDataSourceLabels::Row => "row",
        ChartDataSourceLabels::Column => "column",
        ChartDataSourceLabels::Both => "both",
    }
}
fn legend_position(value: ChartLegendPosition) -> &'static str {
    match value {
        ChartLegendPosition::Start => "start",
        ChartLegendPosition::End => "end",
        ChartLegendPosition::Top => "top",
        ChartLegendPosition::Bottom => "bottom",
        ChartLegendPosition::TopStart => "top-start",
        ChartLegendPosition::TopEnd => "top-end",
        ChartLegendPosition::BottomStart => "bottom-start",
        ChartLegendPosition::BottomEnd => "bottom-end",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> ChartDefinition {
        let mut chart = ChartDefinition::new("chart:line");
        chart.title = Some(ChartText::new("Quarterly revenue"));
        chart.legend = Some(ChartLegendSpec::default());
        let mut x = ChartAxisSpec::new(ChartAxisDimension::X);
        x.name = Some("x-axis".to_string());
        x.categories_cell_range_address = Some("local-table.A2:.A4".to_string());
        let mut y = ChartAxisSpec::new(ChartAxisDimension::Y);
        y.name = Some("y-axis".to_string());
        y.grids.push(ChartGridSpec {
            class: ChartGridClass::Major,
            style_name: None,
            extensions: ChartExtensions::default(),
        });
        chart.plot_area.axes = vec![x, y];
        chart.plot_area.series.push(ChartSeriesSpec {
            values_cell_range_address: Some("local-table.B2:.B4".to_string()),
            label_cell_address: Some("local-table.B1".to_string()),
            attached_axis: Some("y-axis".to_string()),
            data_points: vec![ChartDataPointSpec {
                repeated: 3,
                ..ChartDataPointSpec::default()
            }],
            mean_value: Some(ChartStyleElement::default()),
            regression_curves: vec![ChartRegressionSpec {
                equation: Some(ChartEquationSpec {
                    display_equation: true,
                    display_r_square: true,
                    ..ChartEquationSpec::default()
                }),
                ..ChartRegressionSpec::default()
            }],
            ..ChartSeriesSpec::default()
        });
        let mut table = ChartCachedTable::new("local-table", 2);
        table.header_rows.push(ChartCachedRow::new(vec![
            ChartCachedCell::new(ChartCachedValue::String("Quarter".into())),
            ChartCachedCell::new(ChartCachedValue::String("Revenue".into())),
        ]));
        table.rows.push(ChartCachedRow::new(vec![
            ChartCachedCell::new(ChartCachedValue::String("Q1".into())),
            ChartCachedCell::new(ChartCachedValue::Float(10.0)),
        ]));
        chart.cached_table = Some(table);
        chart
    }

    #[test]
    fn standalone_odc_package_roundtrip() {
        let document = ChartDocument::create(&sample()).unwrap();
        let bytes = document.to_bytes();
        let reopened = ChartDocument::from_bytes(bytes).unwrap();
        assert_eq!(reopened.mimetype(), constants::ODF_CHART);
        assert_eq!(
            reopened.chart().attribute(Some(CHART_NAMESPACE), "class"),
            Some("chart:line")
        );
        assert!(
            reopened
                .chart()
                .children()
                .iter()
                .any(|node| node.namespace_uri() == Some(TABLE_NAMESPACE)
                    && node.local_name() == "table")
        );
        assert_eq!(reopened.plot_area().unwrap().series().count(), 1);
    }

    #[test]
    fn mutation_preserves_auxiliary_parts_and_formula_is_inert() {
        let mut definition = sample();
        definition.cached_table.as_mut().unwrap().rows[0].cells[1].formula =
            Some("of:=SUM([.B2:.B4])&\"x\"".into());
        let mut document = ChartDocument::create(&definition).unwrap();
        definition.title = Some(ChartText::new("Changed"));
        document.set_definition(&definition).unwrap();
        let content = crate::OpenDocumentPackage::from_bytes(document.to_bytes())
            .unwrap()
            .content_xml()
            .unwrap();
        assert!(content.contains("table:formula=\"of:=SUM([.B2:.B4])&amp;&quot;x&quot;\""));
        assert_eq!(document.text(), "ChangedQuarterRevenueQ1");
    }

    #[test]
    fn rejects_dangling_axis_and_zero_counts() {
        let mut definition = sample();
        definition.plot_area.series[0].attached_axis = Some("missing".into());
        assert!(serialize_chart_content(&definition).is_err());
        definition.plot_area.series[0].attached_axis = Some("y-axis".into());
        definition.plot_area.series[0].data_points[0].repeated = 0;
        assert!(serialize_chart_content(&definition).is_err());
    }
}
