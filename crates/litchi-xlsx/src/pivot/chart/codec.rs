//! Bounded `DrawingML` pivot-source and c14 pivot-options codec.

use super::{
    Binding, C14_CHART_NAMESPACE, DropZoneVisibility, FieldType, MAX_CHART_PART_BYTES, MAX_DEPTH,
    MAX_EXTENSION_URIS, MAX_SERIES_PER_CHART, MAX_TEXT_BYTES, OPTIONS_EXTENSION_URI, Options,
    Series, Source, invalid, limit,
};
use crate::error::{Error, Result};
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_chart_name, unqualified_attribute_value,
};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

/// Parse the pivot-chart payload of one chart part.
///
/// Returns `Ok(None)` for ordinary charts that have no `c:pivotSource`.
/// Extension lists with unknown URIs and unknown `c14` children are skipped
/// without failing; structurally invalid pivot sources and pivot options are
/// rejected.
pub fn parse_binding(xml: &[u8]) -> Result<Option<Binding>> {
    if xml.len() > MAX_CHART_PART_BYTES {
        return Err(limit("chart part bytes"));
    }
    let xml = litchi_ooxml_common::mce::process_ooxml(xml)?;
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
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("incomplete pivot-chart XML"));
    }
    Ok(state.source.map(|source| Binding {
        source,
        series: state.series,
    }))
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
    source: Option<Source>,
    source_name: Option<String>,
    source_format_id: Option<u32>,
    extension_uris: Vec<String>,
    name_text: String,
    series: Vec<Series>,
    pending_series: Option<PendingSeries>,
    pending_options: Option<Options>,
}

struct PendingSeries {
    index: Option<u32>,
    options: Option<Options>,
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
    let in_source = stack
        .iter()
        .any(|context| matches!(context, Context::PivotSource));
    if in_chart
        && !in_series
        && !in_source
        && is_drawingml_chart_name(namespace, element.name(), b"ser")
    {
        state.pending_series = Some(PendingSeries {
            index: None,
            options: None,
        });
        return Ok(Context::Series);
    }
    Ok(match stack.last() {
        Some(Context::ChartSpace) => {
            if is_drawingml_chart_name(namespace, element.name(), b"chart") {
                Context::Chart
            } else if is_drawingml_chart_name(namespace, element.name(), b"pivotSource") {
                if state.source.is_some() {
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
                if uri.as_deref() == Some(OPTIONS_EXTENSION_URI) {
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
                state.pending_options = Some(Options::default());
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
            attach_options(state, Options::default())?;
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
            state.source = Some(Source {
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
                state.series.push(Series {
                    index,
                    options: pending.options,
                });
            }
        },
        Context::PivotOptions => {
            let options = state
                .pending_options
                .take()
                .ok_or_else(|| invalid("mismatched pivot-options close"))?;
            attach_options(state, options)?;
        },
        Context::ChartSpace
        | Context::Chart
        | Context::PivotSourceExtensionList
        | Context::Extension
        | Context::SeriesExtensionList
        | Context::SeriesPivotExtension
        | Context::Other => {},
    }
    Ok(())
}

fn attach_options(state: &mut BindingState, options: Options) -> Result<()> {
    let pending = state
        .pending_series
        .as_mut()
        .ok_or_else(|| invalid("pivot options outside a series"))?;
    if pending.options.replace(options).is_some() {
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
    let Some(field_type) = FieldType::from_drop_zone_element(local) else {
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
    options.drop_zones.push(DropZoneVisibility {
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
        .map_err(|_source| invalid(format!("invalid {description} '{value}'")))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Invalid(error.to_string())
}
