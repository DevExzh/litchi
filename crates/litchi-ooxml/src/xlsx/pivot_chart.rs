//! Typed, inert pivot-chart bindings for XLSX workbooks.
//!
//! A pivot chart is a classic chart (`c:chartSpace`) whose `c:pivotSource`
//! element binds it to a pivot table hosted on a worksheet (ECMA-376 part 1,
//! `CT_PivotSource`). Series-level `c:extLst` entries carry the MS-XLSX
//! `c14:pivotOptions` extension (drop-zone visibility per PivotTable field
//! type). This module parses both, resolves the pivot-table name through the
//! workbook part graph to the typed pivot-table model, and validates the
//! binding in the same style as the slicer-cache and timeline modules.
//!
//! Everything here is read-only and inert: no pivot refresh, no cache
//! rebuild, and no rendering.

use crate::common::xml::{
    decode_xml_reference, is_drawingml_chart_name, unqualified_attribute_value,
};
use crate::error::{OoxmlError, Result};
use crate::pivot::PivotTable;
use crate::xlsx::drawing::parse_drawing_xml;
use crate::xlsx::parsers::workbook_parser;
use crate::xlsx::pivot::read_pivot_tables;
use crate::xlsx::worksheet::WorksheetInfo;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::str::FromStr;

/// Extension URI of the MS-XLSX `c14:pivotOptions` series extension that
/// records drop-zone visibility for a pivot chart.
pub const PIVOT_OPTIONS_EXTENSION_URI: &str = "{781A3756-C4B2-4CAC-9D66-4F8C8630D5DC}";

const C14_CHART_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
const MAX_CHART_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_WORKBOOK_BYTES: usize = 32 * 1024 * 1024;
const MAX_DRAWING_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_DRAWINGS_PER_WORKSHEET: usize = 1024;
const MAX_PIVOT_CHARTS_PER_WORKSHEET: usize = 4096;
const MAX_SERIES_PER_CHART: usize = 16_384;
const MAX_EXTENSION_URIS: usize = 256;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_DEPTH: usize = 256;

/// PivotTable field-type zone addressed by a pivot-chart drop-zone switch.
///
/// The variants parse from the ECMA-376 axis identifiers (`axisRow`,
/// `axisCol`, `axisPage`, `axisValues`, `dataFields`) and map to the
/// `c14:dropZone*` element names used by the series pivot-options extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotChartFieldType {
    /// Row axis fields (`c14:dropZoneCategories`)
    AxisRow,
    /// Column axis fields (`c14:dropZoneSeries`)
    AxisCol,
    /// Page (report filter) fields (`c14:dropZoneAxis`)
    AxisPage,
    /// Value axis fields (`c14:dropZoneValues`)
    AxisValues,
    /// Data fields (`c14:dropZoneData`)
    DataFields,
}

impl PivotChartFieldType {
    /// ECMA-376 axis identifier for this field type.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::AxisRow => "axisRow",
            Self::AxisCol => "axisCol",
            Self::AxisPage => "axisPage",
            Self::AxisValues => "axisValues",
            Self::DataFields => "dataFields",
        }
    }

    /// Map a `c14:dropZone*` element name to its field type.
    fn from_drop_zone_element(local_name: &[u8]) -> Option<Self> {
        match local_name {
            b"dropZoneCategories" => Some(Self::AxisRow),
            b"dropZoneSeries" => Some(Self::AxisCol),
            b"dropZoneAxis" => Some(Self::AxisPage),
            b"dropZoneValues" => Some(Self::AxisValues),
            b"dropZoneData" => Some(Self::DataFields),
            _ => None,
        }
    }
}

impl FromStr for PivotChartFieldType {
    type Err = OoxmlError;

    fn from_str(value: &str) -> Result<Self> {
        Ok(match value {
            "axisRow" => Self::AxisRow,
            "axisCol" => Self::AxisCol,
            "axisPage" => Self::AxisPage,
            "axisValues" => Self::AxisValues,
            "dataFields" => Self::DataFields,
            _ => {
                return Err(invalid(format!("unknown pivot-chart field type '{value}'")));
            },
        })
    }
}

/// Visibility of one drop zone in a pivot chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PivotChartDropZoneVisibility {
    /// PivotTable field type occupying the drop zone
    pub field_type: PivotChartFieldType,
    /// Whether the drop zone's field buttons are visible
    pub visible: bool,
}

/// Drop-zone metadata parsed from one series' `c14:pivotOptions` extension.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotChartPivotOptions {
    /// `c14:dropZoneVisible` master switch; `None` when omitted
    pub drop_zone_visible: Option<bool>,
    /// Per field-type drop-zone switches, in document order
    pub drop_zones: Vec<PivotChartDropZoneVisibility>,
}

impl PivotChartPivotOptions {
    /// Visibility recorded for one field-type zone, if present.
    pub fn visibility(&self, field_type: PivotChartFieldType) -> Option<bool> {
        self.drop_zones
            .iter()
            .find(|zone| zone.field_type == field_type)
            .map(|zone| zone.visible)
    }
}

/// Typed `c:pivotSource` metadata binding a chart to a pivot table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotChartSource {
    /// Pivot-table reference as written in `c:name`, optionally qualified
    /// with a `[workbook]` prefix and a sheet name (for example
    /// `[Book1.xlsx]Sheet1!PivotTable1`)
    pub name: String,
    /// Pivot format identifier from `c:fmtId`
    pub format_id: u32,
    /// Extension URIs recorded under `c:pivotSource/c:extLst`, retained
    /// inertly in document order
    pub extension_uris: Vec<String>,
}

/// Pivot metadata for one chart series.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotChartSeries {
    /// Series index from the series' `c:idx` element
    pub index: u32,
    /// Drop-zone options parsed from the series' `c14:pivotOptions`
    /// extension; `None` when the series carries no such extension
    pub pivot_options: Option<PivotChartPivotOptions>,
}

/// Unresolved pivot-chart payload of one chart part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotChartBinding {
    /// Pivot-table source metadata
    pub pivot_source: PivotChartSource,
    /// Per-series pivot options, keyed by series index as written
    pub series: Vec<PivotChartSeries>,
}

/// A pivot chart anchored on a worksheet, with its pivot-table binding
/// resolved against the workbook's typed pivot-table models.
#[derive(Debug, Clone)]
pub struct PivotChart {
    /// Relationship ID from the worksheet drawing part to the chart part
    pub relationship_id: String,
    /// Chart part name (for example `/xl/charts/chart1.xml`)
    pub part_name: String,
    /// Pivot-table source metadata from the chart part
    pub pivot_source: PivotChartSource,
    /// Per-series pivot options parsed from chart extensions
    pub series: Vec<PivotChartSeries>,
    /// Resolved typed pivot-table model named by `pivot_source`
    pub pivot_table: PivotTable,
}

/// Pivot charts anchored on one worksheet.
#[derive(Debug, Clone)]
pub struct WorksheetPivotCharts {
    /// Worksheet name from the workbook
    pub worksheet_name: String,
    /// Worksheet part name (for example `/xl/worksheets/sheet1.xml`)
    pub worksheet_part_name: String,
    /// Pivot charts anchored on the worksheet, in drawing order
    pub pivot_charts: Vec<PivotChart>,
}

/// Parse the pivot-chart payload of one chart part.
///
/// Returns `Ok(None)` for ordinary charts that have no `c:pivotSource`.
/// Extension lists with unknown URIs and unknown `c14` children are skipped
/// without failing; structurally invalid pivot sources and pivot options are
/// rejected.
pub fn parse_pivot_chart_binding(xml: &[u8]) -> Result<Option<PivotChartBinding>> {
    if xml.len() > MAX_CHART_PART_BYTES {
        return Err(limit("chart part bytes"));
    }
    let xml = crate::common::mce::process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut state = BindingState::default();
    let mut stack: Vec<Context> = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let decoder = reader.decoder();
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("content after pivot-chart root"));
                }
                if stack.is_empty() {
                    if root_seen
                        || !is_drawingml_chart_name(&namespace, element.name(), b"chartSpace")
                    {
                        return Err(invalid("expected c:chartSpace root"));
                    }
                    root_seen = true;
                    stack.push(Context::ChartSpace);
                    continue;
                }
                let context = classify_start(&namespace, &element, &stack, &mut state, decoder)?;
                stack.push(context);
                if stack.len() > MAX_DEPTH {
                    return Err(limit("chart XML depth"));
                }
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("content after pivot-chart root"));
                }
                if stack.is_empty() {
                    return Err(invalid("pivot-chart root cannot be empty"));
                }
                handle_empty(&namespace, &element, &stack, &mut state, decoder)?;
            },
            Event::End(_) => {
                let Some(context) = stack.pop() else {
                    return Err(invalid("unexpected pivot-chart closing element"));
                };
                finalize_end(context, &mut state)?;
                if stack.is_empty() {
                    if context != Context::ChartSpace || !root_seen {
                        return Err(invalid("mismatched pivot-chart root"));
                    }
                    root_closed = true;
                }
            },
            Event::Text(text) if stack.last() == Some(&Context::PivotSourceName) => {
                append_name_text(&mut state, &text.decode().map_err(xml_error)?)?;
            },
            Event::CData(text) if stack.last() == Some(&Context::PivotSourceName) => {
                append_name_text(&mut state, &text.decode().map_err(xml_error)?)?;
            },
            Event::GeneralRef(reference) if stack.last() == Some(&Context::PivotSourceName) => {
                append_name_text(&mut state, &decode_xml_reference(&reference)?)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("incomplete pivot-chart XML"));
    }
    Ok(state.pivot_source.map(|pivot_source| PivotChartBinding {
        pivot_source,
        series: state.series,
    }))
}

/// Load all worksheet-anchored pivot charts in a workbook package.
///
/// One entry is returned per worksheet that anchors at least one pivot
/// chart; worksheets without pivot charts are omitted. Every returned chart
/// has its `c:pivotSource` name resolved to the typed pivot-table model read
/// from the package graph; broken or dangling bindings are errors.
pub fn load_pivot_charts(package: &OpcPackage) -> Result<Vec<WorksheetPivotCharts>> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let tables = read_pivot_tables(package).map_err(|error| invalid(error.to_string()))?;
    let mut output = Vec::new();
    for sheet in &sheets {
        let (part_name, charts) = load_sheet_pivot_charts(package, workbook_part, sheet, &tables)?;
        if !charts.is_empty() {
            output.push(WorksheetPivotCharts {
                worksheet_name: sheet.name.clone(),
                worksheet_part_name: part_name.to_string(),
                pivot_charts: charts,
            });
        }
    }
    Ok(output)
}

/// Load the pivot charts anchored on one worksheet, addressed by sheet name.
pub fn load_worksheet_pivot_charts(
    package: &OpcPackage,
    sheet_name: &str,
) -> Result<Vec<PivotChart>> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let tables = read_pivot_tables(package).map_err(|error| invalid(error.to_string()))?;
    let sheet = sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
        .ok_or_else(|| invalid(format!("worksheet '{sheet_name}' not found")))?;
    let (_, charts) = load_sheet_pivot_charts(package, workbook_part, sheet, &tables)?;
    Ok(charts)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    ChartSpace,
    Chart,
    PivotSource,
    PivotSourceName,
    PivotSourceExtensionList,
    Extension,
    Series,
    SeriesExtensionList,
    SeriesPivotExtension,
    PivotOptions,
    Other,
}

#[derive(Default)]
struct BindingState {
    pivot_source: Option<PivotChartSource>,
    source_name: Option<String>,
    source_format_id: Option<u32>,
    extension_uris: Vec<String>,
    name_text: String,
    series: Vec<PivotChartSeries>,
    pending_series: Option<PendingSeries>,
    pending_options: Option<PivotChartPivotOptions>,
}

struct PendingSeries {
    index: Option<u32>,
    pivot_options: Option<PivotChartPivotOptions>,
}

fn classify_start(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    stack: &[Context],
    state: &mut BindingState,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Context> {
    // Extension payloads (known and unknown) are inert below the `c:ext`
    // boundary; only the series pivot-options extension is interpreted.
    if stack
        .iter()
        .any(|context| matches!(context, Context::Extension))
    {
        return Ok(Context::Other);
    }
    let in_chart = stack
        .iter()
        .any(|context| matches!(context, Context::Chart));
    let in_series = stack
        .iter()
        .any(|context| matches!(context, Context::Series));
    let in_pivot_source = stack
        .iter()
        .any(|context| matches!(context, Context::PivotSource));
    if in_chart
        && !in_series
        && !in_pivot_source
        && is_drawingml_chart_name(namespace, element.name(), b"ser")
    {
        state.pending_series = Some(PendingSeries {
            index: None,
            pivot_options: None,
        });
        return Ok(Context::Series);
    }
    Ok(match stack.last() {
        Some(Context::ChartSpace) => {
            if is_drawingml_chart_name(namespace, element.name(), b"chart") {
                Context::Chart
            } else if is_drawingml_chart_name(namespace, element.name(), b"pivotSource") {
                if state.pivot_source.is_some() {
                    return Err(invalid("pivot chart contains duplicate pivot sources"));
                }
                Context::PivotSource
            } else {
                Context::Other
            }
        },
        Some(Context::PivotSource) => {
            if is_drawingml_chart_name(namespace, element.name(), b"name") {
                if state.source_name.is_some() {
                    return Err(invalid("pivot source contains duplicate names"));
                }
                Context::PivotSourceName
            } else if is_drawingml_chart_name(namespace, element.name(), b"fmtId") {
                set_format_id(state, element, decoder)?;
                Context::Other
            } else if is_drawingml_chart_name(namespace, element.name(), b"extLst") {
                Context::PivotSourceExtensionList
            } else {
                Context::Other
            }
        },
        Some(Context::PivotSourceExtensionList) => {
            if is_drawingml_chart_name(namespace, element.name(), b"ext") {
                capture_extension_uri(state, element, decoder)?;
                Context::Extension
            } else {
                Context::Other
            }
        },
        Some(Context::Series) => {
            if is_drawingml_chart_name(namespace, element.name(), b"idx") {
                set_series_index(state, element, decoder)?;
                Context::Other
            } else if is_drawingml_chart_name(namespace, element.name(), b"extLst") {
                Context::SeriesExtensionList
            } else {
                Context::Other
            }
        },
        Some(Context::SeriesExtensionList) => {
            if is_drawingml_chart_name(namespace, element.name(), b"ext") {
                let uri = unqualified_attribute_value(element, b"uri", decoder)?;
                if uri.as_deref() == Some(PIVOT_OPTIONS_EXTENSION_URI) {
                    Context::SeriesPivotExtension
                } else {
                    // Unknown series extensions degrade gracefully.
                    Context::Extension
                }
            } else {
                Context::Other
            }
        },
        Some(Context::SeriesPivotExtension) => {
            if is_c14_name(namespace, element.name(), b"pivotOptions") {
                if state.pending_options.is_some() {
                    return Err(invalid("series contains duplicate pivot options"));
                }
                state.pending_options = Some(PivotChartPivotOptions::default());
                Context::PivotOptions
            } else {
                Context::Other
            }
        },
        Some(Context::PivotOptions) => {
            if is_c14(namespace) {
                apply_drop_zone(state, element, decoder)?;
            }
            Context::Other
        },
        _ => Context::Other,
    })
}

fn handle_empty(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    stack: &[Context],
    state: &mut BindingState,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    if stack
        .iter()
        .any(|context| matches!(context, Context::Extension))
    {
        return Ok(());
    }
    match stack.last() {
        Some(Context::ChartSpace)
            if is_drawingml_chart_name(namespace, element.name(), b"pivotSource") =>
        {
            return Err(invalid("pivot source requires a name and format ID"));
        },
        Some(Context::PivotSource) => {
            if is_drawingml_chart_name(namespace, element.name(), b"name") {
                if state.source_name.replace(String::new()).is_some() {
                    return Err(invalid("pivot source contains duplicate names"));
                }
            } else if is_drawingml_chart_name(namespace, element.name(), b"fmtId") {
                set_format_id(state, element, decoder)?;
            }
        },
        Some(Context::PivotSourceExtensionList)
            if is_drawingml_chart_name(namespace, element.name(), b"ext") =>
        {
            capture_extension_uri(state, element, decoder)?;
        },
        Some(Context::Series) if is_drawingml_chart_name(namespace, element.name(), b"idx") => {
            set_series_index(state, element, decoder)?;
        },
        Some(Context::SeriesPivotExtension)
            if is_c14_name(namespace, element.name(), b"pivotOptions") =>
        {
            attach_pivot_options(state, PivotChartPivotOptions::default())?;
        },
        Some(Context::PivotOptions) if is_c14(namespace) => {
            apply_drop_zone(state, element, decoder)?;
        },
        _ => {},
    }
    Ok(())
}

fn finalize_end(context: Context, state: &mut BindingState) -> Result<()> {
    match context {
        Context::PivotSourceName => {
            let name = std::mem::take(&mut state.name_text);
            if state.source_name.replace(name).is_some() {
                return Err(invalid("pivot source contains duplicate names"));
            }
        },
        Context::PivotSource => {
            let name = state
                .source_name
                .take()
                .ok_or_else(|| invalid("pivot source requires a name"))?;
            let format_id = state
                .source_format_id
                .take()
                .ok_or_else(|| invalid("pivot source requires a format ID"))?;
            state.pivot_source = Some(PivotChartSource {
                name,
                format_id,
                extension_uris: std::mem::take(&mut state.extension_uris),
            });
        },
        Context::Series => {
            let pending = state
                .pending_series
                .take()
                .ok_or_else(|| invalid("mismatched series close"))?;
            // Series without a valid c:idx are dropped instead of failing.
            if let Some(index) = pending.index {
                if state.series.len() >= MAX_SERIES_PER_CHART {
                    return Err(limit("series per chart"));
                }
                state.series.push(PivotChartSeries {
                    index,
                    pivot_options: pending.pivot_options,
                });
            }
        },
        Context::PivotOptions => {
            let options = state
                .pending_options
                .take()
                .ok_or_else(|| invalid("mismatched pivot-options close"))?;
            attach_pivot_options(state, options)?;
        },
        _ => {},
    }
    Ok(())
}

fn attach_pivot_options(state: &mut BindingState, options: PivotChartPivotOptions) -> Result<()> {
    let pending = state
        .pending_series
        .as_mut()
        .ok_or_else(|| invalid("pivot options outside a series"))?;
    if pending.pivot_options.replace(options).is_some() {
        return Err(invalid("series contains duplicate pivot options"));
    }
    Ok(())
}

fn apply_drop_zone(
    state: &mut BindingState,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let local = element.local_name();
    let local = local.as_ref();
    let value = unqualified_attribute_value(element, b"val", decoder)?;
    let visible = match value.as_deref() {
        Some(value) => parse_bool(value, "drop-zone visibility")?,
        // CT_Boolean defaults to true when val is omitted.
        None => true,
    };
    let options = state
        .pending_options
        .as_mut()
        .ok_or_else(|| invalid("drop zone outside pivot options"))?;
    if local == b"dropZoneVisible" {
        if options.drop_zone_visible.replace(visible).is_some() {
            return Err(invalid("duplicate dropZoneVisible switch"));
        }
        return Ok(());
    }
    let Some(field_type) = PivotChartFieldType::from_drop_zone_element(local) else {
        // Unknown c14 children degrade gracefully.
        return Ok(());
    };
    if options
        .drop_zones
        .iter()
        .any(|zone| zone.field_type == field_type)
    {
        return Err(invalid(format!(
            "duplicate drop-zone switch for '{}'",
            field_type.as_str()
        )));
    }
    options.drop_zones.push(PivotChartDropZoneVisibility {
        field_type,
        visible,
    });
    Ok(())
}

fn set_format_id(
    state: &mut BindingState,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let value = unqualified_attribute_value(element, b"val", decoder)?
        .ok_or_else(|| invalid("pivot-source format ID requires val"))?;
    let format_id = parse_u32(&value, "pivot-source format ID")?;
    if state.source_format_id.replace(format_id).is_some() {
        return Err(invalid("pivot source contains duplicate format IDs"));
    }
    Ok(())
}

fn set_series_index(
    state: &mut BindingState,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let value = unqualified_attribute_value(element, b"val", decoder)?
        .ok_or_else(|| invalid("series index requires val"))?;
    let index = parse_u32(&value, "series index")?;
    let pending = state
        .pending_series
        .as_mut()
        .ok_or_else(|| invalid("series index outside a series"))?;
    if pending.index.replace(index).is_some() {
        return Err(invalid("series contains duplicate indices"));
    }
    Ok(())
}

fn capture_extension_uri(
    state: &mut BindingState,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let Some(uri) = unqualified_attribute_value(element, b"uri", decoder)? else {
        return Ok(());
    };
    if state.extension_uris.len() >= MAX_EXTENSION_URIS {
        return Err(limit("extension URIs"));
    }
    state.extension_uris.push(uri);
    Ok(())
}

fn append_name_text(state: &mut BindingState, text: &str) -> Result<()> {
    if state.name_text.len() + text.len() > MAX_TEXT_BYTES {
        return Err(limit("pivot-source name bytes"));
    }
    state.name_text.push_str(text);
    Ok(())
}

fn is_c14(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == C14_CHART_NAMESPACE.as_bytes()
    )
}

fn is_c14_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name && is_c14(namespace)
}

fn parse_bool(value: &str, description: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid {description} '{value}'"))),
    }
}

fn parse_u32(value: &str, description: &str) -> Result<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid(format!("invalid {description} '{value}'")));
    }
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {description} '{value}'")))
}

fn parse_workbook_sheets(xml: &[u8]) -> Result<Vec<WorksheetInfo>> {
    if xml.len() > MAX_WORKBOOK_BYTES {
        return Err(limit("workbook XML bytes"));
    }
    let content = std::str::from_utf8(xml).map_err(xml_error)?;
    Ok(workbook_parser::parse_workbook_details(content)
        .map_err(|error| invalid(error.to_string()))?
        .sheets)
}

fn load_sheet_pivot_charts(
    package: &OpcPackage,
    workbook_part: &dyn Part,
    sheet: &WorksheetInfo,
    tables: &[PivotTable],
) -> Result<(PackURI, Vec<PivotChart>)> {
    let relationship = workbook_part
        .rels()
        .get(&sheet.relationship_id)
        .ok_or_else(|| {
            invalid(format!(
                "worksheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
    if !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
        return Err(invalid(format!(
            "worksheet '{}' relationship has invalid type '{}'",
            sheet.name,
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(invalid(format!(
            "worksheet '{}' relationship cannot be external",
            sheet.name
        )));
    }
    let sheet_uri = relationship.target_partname()?;
    let sheet_part = package.get_part(&sheet_uri)?;
    if sheet_part.content_type() != ct::SML_WORKSHEET {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::SML_WORKSHEET.into(),
            got: sheet_part.content_type().into(),
        });
    }
    let mut charts = Vec::new();
    let drawings: Vec<_> = sheet_part
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING))
        .collect();
    if drawings.len() > MAX_DRAWINGS_PER_WORKSHEET {
        return Err(limit("drawings per worksheet"));
    }
    for drawing_relationship in drawings {
        if drawing_relationship.is_external() {
            return Err(invalid("worksheet drawing relationship cannot be external"));
        }
        let drawing_uri = drawing_relationship.target_partname()?;
        let drawing_part = package.get_part(&drawing_uri)?;
        if drawing_part.content_type() != ct::OFC_DRAWING {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::OFC_DRAWING.into(),
                got: drawing_part.content_type().into(),
            });
        }
        if drawing_part.blob().len() > MAX_DRAWING_PART_BYTES {
            return Err(limit("drawing part bytes"));
        }
        let drawing_xml = std::str::from_utf8(drawing_part.blob()).map_err(xml_error)?;
        let drawing = parse_drawing_xml(drawing_xml)?
            .ok_or_else(|| invalid(format!("drawing part '{drawing_uri}' has no wsDr root")))?;
        for chart in &drawing.charts {
            let chart_relationship =
                drawing_part
                    .rels()
                    .get(&chart.relationship_id)
                    .ok_or_else(|| {
                        invalid(format!(
                            "drawing chart references missing relationship '{}'",
                            chart.relationship_id
                        ))
                    })?;
            if !matches!(chart_relationship.reltype(), rt::CHART | rt::STRICT_CHART) {
                return Err(invalid(format!(
                    "drawing chart relationship '{}' has invalid type '{}'",
                    chart.relationship_id,
                    chart_relationship.reltype()
                )));
            }
            if chart_relationship.is_external() {
                return Err(invalid("drawing chart relationship cannot be external"));
            }
            let chart_uri = chart_relationship.target_partname()?;
            let chart_part = package.get_part(&chart_uri)?;
            if chart_part.content_type() != ct::DML_CHART {
                return Err(OoxmlError::InvalidContentType {
                    expected: ct::DML_CHART.into(),
                    got: chart_part.content_type().into(),
                });
            }
            // Ordinary charts have no pivot source and are excluded.
            let Some(binding) = parse_pivot_chart_binding(chart_part.blob())? else {
                continue;
            };
            if charts.len() >= MAX_PIVOT_CHARTS_PER_WORKSHEET {
                return Err(limit("pivot charts per worksheet"));
            }
            let pivot_table =
                resolve_pivot_table(&chart_uri, &binding.pivot_source, sheet, tables)?;
            charts.push(PivotChart {
                relationship_id: chart.relationship_id.clone(),
                part_name: chart_uri.to_string(),
                pivot_source: binding.pivot_source,
                series: binding.series,
                pivot_table: pivot_table.clone(),
            });
        }
    }
    Ok((sheet_uri, charts))
}

/// Resolve a `c:pivotSource` name to the typed pivot-table model.
///
/// Names written by Excel are sheet-qualified (`[Book1.xlsx]Sheet1!Pivot1`);
/// an unqualified name resolves against the chart's own worksheet first and
/// then against the whole workbook, with ambiguity reported as an error.
fn resolve_pivot_table<'a>(
    chart_uri: &PackURI,
    pivot_source: &PivotChartSource,
    sheet: &WorksheetInfo,
    tables: &'a [PivotTable],
) -> Result<&'a PivotTable> {
    let (sheet_prefix, table_name) = split_pivot_source_name(&pivot_source.name);
    if table_name.is_empty() {
        return Err(invalid(format!(
            "pivot chart '{chart_uri}' has an empty pivot-table name"
        )));
    }
    let folded = table_name.to_lowercase();
    let candidates = || {
        tables
            .iter()
            .filter(|table| table.name.to_lowercase() == folded)
    };
    match sheet_prefix {
        Some(prefix) => {
            let wanted = prefix.to_lowercase();
            candidates()
                .find(|table| table.sheet_name.to_lowercase() == wanted)
                .ok_or_else(|| {
                    invalid(format!(
                        "pivot chart '{chart_uri}' references pivot table '{table_name}' on sheet '{prefix}', which does not host it"
                    ))
                })
        },
        None => {
            let host = sheet.name.to_lowercase();
            if let Some(table) = candidates().find(|table| table.sheet_name.to_lowercase() == host)
            {
                return Ok(table);
            }
            let mut matches = candidates();
            match (matches.next(), matches.next()) {
                (Some(table), None) => Ok(table),
                (None, _) => Err(invalid(format!(
                    "pivot chart '{chart_uri}' references missing pivot table '{table_name}'"
                ))),
                (Some(_), Some(_)) => Err(invalid(format!(
                    "pivot chart '{chart_uri}' pivot-table name '{table_name}' is ambiguous"
                ))),
            }
        },
    }
}

/// Split a `c:pivotSource` name into its optional sheet qualifier and the
/// pivot-table name, stripping any `[workbook]` prefix and sheet quoting.
fn split_pivot_source_name(name: &str) -> (Option<String>, String) {
    let Some(bang) = name.rfind('!') else {
        return (None, name.to_string());
    };
    let mut sheet = &name[..bang];
    if let Some(rest) = sheet.strip_prefix('[') {
        if let Some(end) = rest.find(']') {
            sheet = &rest[end + 1..];
        }
    }
    (
        Some(unquote_sheet_name(sheet)),
        name[bang + 1..].to_string(),
    )
}

fn unquote_sheet_name(sheet: &str) -> String {
    if sheet.len() >= 2 && sheet.starts_with('\'') && sheet.ends_with('\'') {
        sheet[1..sheet.len() - 1].replace("''", "'")
    } else {
        sheet.to_string()
    }
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(name: &str) -> OoxmlError {
    invalid(format!("pivot chart {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::BlobPart;

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const POI_PIVOT_CHART: &[u8] =
        include_bytes!("../../../../test-data/poi/test-data/spreadsheet/WithChartSheet.xlsx");

    fn pivot_chart_xml(name: &str) -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{C}" xmlns:c14="{C14_CHART_NAMESPACE}">
                <c:lang val="en-US"/>
                <c:pivotSource>
                    <c:name>{name}</c:name>
                    <c:fmtId val="7"/>
                    <c:extLst>
                        <c:ext uri="{{9EF4A71B-D1ED-4E66-9E1E-2A10C2C8F28B}}"/>
                        <c:ext uri="{{00000000-0000-0000-0000-000000000000}}"><x:payload xmlns:x="urn:example:source"/></c:ext>
                    </c:extLst>
                </c:pivotSource>
                <c:chart>
                    <c:plotArea>
                        <c:barChart>
                            <c:ser>
                                <c:idx val="0"/>
                                <c:extLst>
                                    <c:ext uri="{PIVOT_OPTIONS_EXTENSION_URI}">
                                        <c14:pivotOptions>
                                            <c14:dropZoneVisible val="0"/>
                                            <c14:dropZoneCategories val="0"/>
                                            <c14:dropZoneData val="1"/>
                                            <c14:dropZoneSeries val="0"/>
                                            <c14:dropZoneAxis val="1"/>
                                            <c14:dropZoneValues val="1"/>
                                            <c14:futureSwitch val="1"/>
                                        </c14:pivotOptions>
                                    </c:ext>
                                    <c:ext uri="{{11111111-2222-3333-4444-555555555555}}">
                                        <x:payload xmlns:x="urn:example:series"><x:c:ser xmlns:x:c="{C}"><x:c:idx val="9"/></x:c:ser></x:payload>
                                    </c:ext>
                                </c:extLst>
                            </c:ser>
                            <c:ser><c:idx val="1"/></c:ser>
                        </c:barChart>
                    </c:plotArea>
                </c:chart>
            </c:chartSpace>"#
        )
    }

    #[test]
    fn parses_pivot_source_and_series_pivot_options() {
        let binding = parse_pivot_chart_binding(pivot_chart_xml("Pivot!PivotOne").as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(binding.pivot_source.name, "Pivot!PivotOne");
        assert_eq!(binding.pivot_source.format_id, 7);
        assert_eq!(
            binding.pivot_source.extension_uris,
            [
                "{9EF4A71B-D1ED-4E66-9E1E-2A10C2C8F28B}",
                "{00000000-0000-0000-0000-000000000000}"
            ]
        );
        assert_eq!(binding.series.len(), 2);
        let options = binding.series[0].pivot_options.as_ref().unwrap();
        assert_eq!(options.drop_zone_visible, Some(false));
        assert_eq!(
            options.visibility(PivotChartFieldType::AxisRow),
            Some(false)
        );
        assert_eq!(
            options.visibility(PivotChartFieldType::AxisCol),
            Some(false)
        );
        assert_eq!(
            options.visibility(PivotChartFieldType::AxisPage),
            Some(true)
        );
        assert_eq!(
            options.visibility(PivotChartFieldType::AxisValues),
            Some(true)
        );
        assert_eq!(
            options.visibility(PivotChartFieldType::DataFields),
            Some(true)
        );
        // A series without the pivot-options extension reports None.
        assert_eq!(binding.series[1].index, 1);
        assert!(binding.series[1].pivot_options.is_none());
        // Unknown series extensions and their nested payload stay inert.
        assert!(!binding.series.iter().any(|series| series.index == 9));
    }

    #[test]
    fn ordinary_chart_has_no_binding() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:barChart>
                <c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        assert!(parse_pivot_chart_binding(xml.as_bytes()).unwrap().is_none());
    }

    #[test]
    fn field_type_parses_axis_identifiers() {
        assert_eq!(
            "axisRow".parse::<PivotChartFieldType>().unwrap(),
            PivotChartFieldType::AxisRow
        );
        assert_eq!(
            "axisCol".parse::<PivotChartFieldType>().unwrap(),
            PivotChartFieldType::AxisCol
        );
        assert_eq!(
            "axisPage".parse::<PivotChartFieldType>().unwrap(),
            PivotChartFieldType::AxisPage
        );
        assert_eq!(
            "axisValues".parse::<PivotChartFieldType>().unwrap(),
            PivotChartFieldType::AxisValues
        );
        assert_eq!(
            "dataFields".parse::<PivotChartFieldType>().unwrap(),
            PivotChartFieldType::DataFields
        );
        assert!("axisRow".parse::<PivotChartFieldType>().unwrap().as_str() == "axisRow");
        assert!("bogus".parse::<PivotChartFieldType>().is_err());
    }

    #[test]
    fn rejects_malformed_pivot_chart_parts() {
        let head = format!(r#"<c:chartSpace xmlns:c="{C}" xmlns:c14="{C14_CHART_NAMESPACE}">"#);
        let cases = [
            // Duplicate pivot sources.
            format!(
                "{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource>\
                 <c:pivotSource><c:name>B</c:name><c:fmtId val=\"2\"/></c:pivotSource></c:chartSpace>"
            ),
            // Missing format ID.
            format!("{head}<c:pivotSource><c:name>A</c:name></c:pivotSource></c:chartSpace>"),
            // Missing name.
            format!("{head}<c:pivotSource><c:fmtId val=\"1\"/></c:pivotSource></c:chartSpace>"),
            // Empty pivot source.
            format!("{head}<c:pivotSource/></c:chartSpace>"),
            // Invalid boolean on a drop zone.
            format!(
                "{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource><c:chart>\
                 <c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:extLst>\
                 <c:ext uri=\"{PIVOT_OPTIONS_EXTENSION_URI}\"><c14:pivotOptions>\
                 <c14:dropZoneAxis val=\"maybe\"/></c14:pivotOptions></c:ext></c:extLst></c:ser>\
                 </c:barChart></c:plotArea></c:chart></c:chartSpace>"
            ),
            // Duplicate drop-zone switch.
            format!(
                "{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource><c:chart>\
                 <c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:extLst>\
                 <c:ext uri=\"{PIVOT_OPTIONS_EXTENSION_URI}\"><c14:pivotOptions>\
                 <c14:dropZoneData val=\"1\"/><c14:dropZoneData val=\"0\"/>\
                 </c14:pivotOptions></c:ext></c:extLst></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
            ),
            // DTDs are rejected.
            format!(
                "<!DOCTYPE c:chartSpace>{head}<c:pivotSource><c:name>A</c:name><c:fmtId val=\"1\"/></c:pivotSource></c:chartSpace>"
            ),
            // Wrong root.
            format!("<c:chart xmlns:c=\"{C}\"/>"),
        ];
        for xml in cases {
            assert!(
                parse_pivot_chart_binding(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        assert!(parse_pivot_chart_binding(&vec![b' '; MAX_CHART_PART_BYTES + 1]).is_err());
    }

    fn drawing_xml(chart_relationship_id: &str) -> String {
        format!(
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:c="{C}" xmlns:r="{R}">
                <xdr:twoCellAnchor>
                    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
                        <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff>
                        <xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                    <xdr:graphicFrame><a:graphic><a:graphicData>
                        <c:chart r:id="{chart_relationship_id}"/>
                    </a:graphicData></a:graphic></xdr:graphicFrame>
                    <xdr:clientData/>
                </xdr:twoCellAnchor>
            </xdr:wsDr>"#
        )
    }

    fn package_with_pivot_chart(chart_xml: &str) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let chart_uri = PackURI::new("/xl/charts/chart1.xml").unwrap();
        let mut workbook_part = BlobPart::new(
            PackURI::new("/xl/workbook.xml").unwrap(),
            ct::SML_SHEET_MAIN.to_string(),
            format!(
                r#"<workbook xmlns="{SML}" xmlns:r="{R}">
                    <sheets><sheet name="Pivot" sheetId="1" r:id="rId1"/>
                        <sheet name="Source" sheetId="2" r:id="rId2"/></sheets>
                    <pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches>
                </workbook>"#
            )
            .into_bytes(),
        );
        workbook_part.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        workbook_part.relate_to("worksheets/sheet2.xml", rt::WORKSHEET);
        workbook_part.relate_to(
            "pivotCache/pivotCacheDefinition1.xml",
            rt::PIVOT_CACHE_DEFINITION,
        );
        let mut sheet_part = BlobPart::new(
            PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
        );
        sheet_part.relate_to("../pivotTables/pivotTable1.xml", rt::PIVOT_TABLE);
        sheet_part.relate_to("../drawings/drawing1.xml", rt::DRAWING);
        let mut drawing_part = BlobPart::new(
            PackURI::new("/xl/drawings/drawing1.xml").unwrap(),
            ct::OFC_DRAWING.to_string(),
            drawing_xml("rId1").into_bytes(),
        );
        drawing_part.relate_to("../charts/chart1.xml", rt::CHART);
        let mut cache_part = BlobPart::new(
            PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap(),
            ct::SML_PIVOT_CACHE_DEFINITION.to_string(),
            format!(
                r#"<pivotCacheDefinition xmlns="{SML}" xmlns:r="{R}" r:id="rId1" recordCount="2">
                    <cacheSource type="worksheet"><worksheetSource ref="$A$1:$B$3" r:id="rId2"/></cacheSource>
                    <cacheFields count="1"><cacheField name="Cache Region"/></cacheFields>
                </pivotCacheDefinition>"#
            )
            .into_bytes(),
        );
        cache_part.relate_to("pivotCacheRecords1.xml", rt::PIVOT_CACHE_RECORDS);
        cache_part.relate_to("../worksheets/sheet2.xml", rt::WORKSHEET);
        let mut table_part = BlobPart::new(
            PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap(),
            ct::SML_PIVOT_TABLE.to_string(),
            format!(
                r#"<pivotTableDefinition xmlns="{SML}" name="PivotOne" cacheId="7" dataCaption="Values">
                    <location ref="A1:C5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="1"><pivotField/></pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                </pivotTableDefinition>"#
            )
            .into_bytes(),
        );
        table_part.relate_to(
            "../pivotCache/pivotCacheDefinition1.xml",
            rt::PIVOT_CACHE_DEFINITION,
        );
        package.relate_to(
            "xl/workbook.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(sheet_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/worksheets/sheet2.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
        )));
        package.add_part(Box::new(drawing_part));
        package.add_part(Box::new(BlobPart::new(
            chart_uri.clone(),
            ct::DML_CHART.to_string(),
            chart_xml.as_bytes().to_vec(),
        )));
        package.add_part(Box::new(cache_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/pivotCache/pivotCacheRecords1.xml").unwrap(),
            ct::SML_PIVOT_CACHE_RECORDS.to_string(),
            format!(
                r#"<pivotCacheRecords xmlns="{SML}" count="2">
                    <r><s v="North"/></r><r><s v="South"/></r>
                </pivotCacheRecords>"#
            )
            .into_bytes(),
        )));
        package.add_part(Box::new(table_part));
        (package, chart_uri)
    }

    #[test]
    fn resolves_qualified_and_plain_pivot_table_names() {
        let (package, chart_uri) =
            package_with_pivot_chart(&pivot_chart_xml("[Book1.xlsx]Pivot!PivotOne"));
        let sheets = load_pivot_charts(&package).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].worksheet_name, "Pivot");
        assert_eq!(sheets[0].worksheet_part_name, "/xl/worksheets/sheet1.xml");
        assert_eq!(sheets[0].pivot_charts.len(), 1);
        let chart = &sheets[0].pivot_charts[0];
        assert_eq!(chart.part_name, chart_uri.to_string());
        assert_eq!(chart.relationship_id, "rId1");
        assert_eq!(chart.pivot_source.format_id, 7);
        assert_eq!(chart.pivot_table.name, "PivotOne");
        assert_eq!(chart.pivot_table.sheet_name, "Pivot");
        assert_eq!(chart.pivot_table.cache_id, 7);
        let options = chart.series[0].pivot_options.as_ref().unwrap();
        assert_eq!(options.drop_zone_visible, Some(false));
        assert_eq!(
            options.visibility(PivotChartFieldType::AxisPage),
            Some(true)
        );

        // Per-worksheet accessor with a plain, unqualified pivot-table name.
        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("PivotOne"));
        let charts = load_worksheet_pivot_charts(&package, "Pivot").unwrap();
        assert_eq!(charts.len(), 1);
        assert_eq!(charts[0].pivot_table.name, "PivotOne");
        assert!(load_worksheet_pivot_charts(&package, "Missing").is_err());
    }

    #[test]
    fn rejects_dangling_and_foreign_sheet_bindings() {
        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("Pivot!NoSuchTable"));
        assert!(load_pivot_charts(&package).is_err());

        // The table exists but is hosted on a different sheet than qualified.
        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("Source!PivotOne"));
        assert!(load_pivot_charts(&package).is_err());

        let (package, _) = package_with_pivot_chart(&pivot_chart_xml("MissingTable"));
        assert!(load_pivot_charts(&package).is_err());
    }

    #[test]
    fn excludes_ordinary_charts_and_validates_chart_graph() {
        let ordinary = format!(
            r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:barChart>
                <c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let (package, _) = package_with_pivot_chart(&ordinary);
        assert!(load_pivot_charts(&package).unwrap().is_empty());
        assert!(
            load_worksheet_pivot_charts(&package, "Pivot")
                .unwrap()
                .is_empty()
        );

        // A drawing anchor whose chart relationship is missing is an error.
        let (mut package, _) = package_with_pivot_chart(&pivot_chart_xml("PivotOne"));
        let drawing_uri = PackURI::new("/xl/drawings/drawing1.xml").unwrap();
        package
            .get_part_mut(&drawing_uri)
            .unwrap()
            .rels_mut()
            .remove("rId1")
            .unwrap();
        assert!(load_pivot_charts(&package).is_err());

        // A chart part with the wrong content type is an error.
        let (mut package, chart_uri) = package_with_pivot_chart(&pivot_chart_xml("PivotOne"));
        package.add_part(Box::new(BlobPart::new(
            chart_uri,
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        )));
        assert!(load_pivot_charts(&package).is_err());
    }

    #[test]
    fn poi_fixture_pivot_chart_part_parses() {
        let package = OpcPackage::from_bytes(POI_PIVOT_CHART).unwrap();
        let chart_part = package
            .get_part(&PackURI::new("/xl/charts/chart1.xml").unwrap())
            .unwrap();
        let binding = parse_pivot_chart_binding(chart_part.blob())
            .unwrap()
            .expect("fixture chart is a pivot chart");
        assert_eq!(binding.pivot_source.name, "[CVT23.tmp]Sheet2!PivotTable2");
        assert_eq!(binding.pivot_source.format_id, 0);
        let (sheet, table) = split_pivot_source_name(&binding.pivot_source.name);
        assert_eq!(sheet.as_deref(), Some("Sheet2"));
        assert_eq!(table, "PivotTable2");
        assert!(!binding.series.is_empty());
    }
}
