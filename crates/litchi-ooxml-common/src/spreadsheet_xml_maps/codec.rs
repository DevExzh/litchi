//! Bounded `SpreadsheetML` Custom XML Maps XML codec.

use std::collections::{HashMap, HashSet};

use super::invalid;
use super::model::{
    DataBindingRef, NS, NS_TEXT, ParsedXmlMapInfo, STRICT_NS, STRICT_NS_TEXT, XmlMap,
    XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapInfoRef, XmlMapLimits, XmlMapRef,
    XmlMapSchema, XmlSchemaRef,
};
use super::validation::{validate_xml_map_info_ref_with_limits, validate_xml_map_info_with_limits};
use crate::Result;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};

impl XmlMapInfo {
    pub(crate) fn parse(xml: &[u8]) -> Result<Self> {
        Self::parse_with_limits(xml, &XmlMapLimits::DEFAULT)
    }

    pub(crate) fn parse_with_limits(xml: &[u8], limits: &XmlMapLimits) -> Result<Self> {
        Ok(parse_xml_map_info_with_conformance_and_limits(xml, limits)?.info)
    }

    #[allow(dead_code)]
    pub(crate) fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        self.to_xml_with_limits(strict, &XmlMapLimits::DEFAULT)
    }

    #[allow(dead_code)]
    pub(crate) fn to_xml_with_limits(
        &self,
        strict: bool,
        limits: &XmlMapLimits,
    ) -> Result<Vec<u8>> {
        if limits.max_part_bytes == 0 {
            return Err(invalid(
                "serialized custom XML maps part exceeds configured limit",
            ));
        }
        validate_xml_map_info_with_limits(self, limits)?;
        let namespace = if strict { STRICT_NS_TEXT } else { NS_TEXT };
        let mut xml = BoundedXml::new(limits.max_part_bytes);
        xml.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><MapInfo xmlns=\"",
        )?;
        escape_attr(&mut xml, namespace)?;
        xml.push_str("\" SelectionNamespaces=\"")?;
        escape_attr(&mut xml, &self.selection_namespaces)?;
        xml.push_str("\">")?;
        for schema in &self.schemas {
            xml.push_str("<Schema ID=\"")?;
            escape_attr(&mut xml, &schema.id)?;
            xml.push_char('"')?;
            optional_string_attr(&mut xml, "SchemaRef", schema.schema_reference.as_deref())?;
            optional_string_attr(&mut xml, "Namespace", schema.namespace.as_deref())?;
            if let Some(payload) = &schema.payload_xml {
                xml.push_char('>')?;
                xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?)?;
                xml.push_str("</Schema>")?;
            } else {
                xml.push_str("/>")?;
            }
        }
        for map in &self.maps {
            xml.push_str("<Map ID=\"")?;
            xml.push_str(&map.id.to_string())?;
            xml.push_str("\" Name=\"")?;
            escape_attr(&mut xml, &map.name)?;
            xml.push_str("\" RootElement=\"")?;
            escape_attr(&mut xml, &map.root_element)?;
            xml.push_str("\" SchemaID=\"")?;
            escape_attr(&mut xml, &map.schema_id)?;
            xml.push_char('"')?;
            bool_attr(
                &mut xml,
                "ShowImportExportValidationErrors",
                map.show_import_export_validation_errors,
            )?;
            bool_attr(&mut xml, "AutoFit", map.auto_fit)?;
            bool_attr(&mut xml, "Append", map.append)?;
            bool_attr(
                &mut xml,
                "PreserveSortAFLayout",
                map.preserve_sort_auto_filter_layout,
            )?;
            bool_attr(&mut xml, "PreserveFormat", map.preserve_format)?;
            if let Some(binding) = &map.data_binding {
                xml.push_str("><DataBinding")?;
                optional_string_attr(
                    &mut xml,
                    "DataBindingName",
                    binding.data_binding_name.as_deref(),
                )?;
                optional_bool_attr(&mut xml, "FileBinding", binding.file_binding)?;
                optional_u32_attr(&mut xml, "ConnectionID", binding.connection_id)?;
                optional_string_attr(
                    &mut xml,
                    "FileBindingName",
                    binding.file_binding_name.as_deref(),
                )?;
                optional_u32_attr(&mut xml, "DataBindingLoadMode", Some(binding.load_mode))?;
                if let Some(payload) = &binding.payload_xml {
                    xml.push_char('>')?;
                    xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?)?;
                    xml.push_str("</DataBinding></Map>")?;
                } else {
                    xml.push_str("/></Map>")?;
                }
            } else {
                xml.push_str("/>")?;
            }
        }
        xml.push_str("</MapInfo>")?;
        Ok(xml.finish())
    }
}

fn parse_xml_map_info_with_conformance_and_limits_impl(
    xml: &[u8],
    limits: &XmlMapLimits,
) -> Result<ParsedXmlMapInfo> {
    if xml.len() > limits.max_part_bytes {
        return Err(invalid("custom XML maps part exceeds 32 MiB"));
    }
    let processed = crate::mce::process_ooxml(xml)?;
    if processed.len() > limits.max_part_bytes {
        return Err(invalid("processed custom XML maps part exceeds 32 MiB"));
    }
    std::str::from_utf8(processed.as_ref()).map_err(xml_error)?;
    parse_processed(processed.as_ref(), limits)
}

/*
        let mut xml = BoundedXml::new(MAX_PART_BYTES);
        xml.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><MapInfo xmlns=\"",
        )?;
        escape_attr(&mut xml, namespace)?;
        xml.push_str("\" SelectionNamespaces=\"")?;
        escape_attr(&mut xml, &self.selection_namespaces)?;
        xml.push_str("\">")?;
        for schema in &self.schemas {
            xml.push_str("<Schema ID=\"")?;
            escape_attr(&mut xml, &schema.id)?;
            xml.push_char('"')?;
            optional_string_attr(&mut xml, "SchemaRef", schema.schema_reference.as_deref())?;
            optional_string_attr(&mut xml, "Namespace", schema.namespace.as_deref())?;
            if let Some(payload) = &schema.payload_xml {
                xml.push_char('>')?;
                xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?)?;
                xml.push_str("</Schema>")?;
            } else {
                xml.push_str("/>")?;
            }
        }
        for map in &self.maps {
            xml.push_str("<Map ID=\"")?;
            xml.push_str(&map.id.to_string())?;
            xml.push_str("\" Name=\"")?;
            escape_attr(&mut xml, &map.name)?;
            xml.push_str("\" RootElement=\"")?;
            escape_attr(&mut xml, &map.root_element)?;
            xml.push_str("\" SchemaID=\"")?;
            escape_attr(&mut xml, &map.schema_id)?;
            xml.push_char('"')?;
            bool_attr(
                &mut xml,
                "ShowImportExportValidationErrors",
                map.show_import_export_validation_errors,
            )?;
            bool_attr(&mut xml, "AutoFit", map.auto_fit)?;
            bool_attr(&mut xml, "Append", map.append)?;
            bool_attr(
                &mut xml,
                "PreserveSortAFLayout",
                map.preserve_sort_auto_filter_layout,
            )?;
            bool_attr(&mut xml, "PreserveFormat", map.preserve_format)?;
            if let Some(binding) = &map.data_binding {
                xml.push_str("><DataBinding")?;
                optional_string_attr(
                    &mut xml,
                    "DataBindingName",
                    binding.data_binding_name.as_deref(),
                )?;
                optional_bool_attr(&mut xml, "FileBinding", binding.file_binding)?;
                optional_u32_attr(&mut xml, "ConnectionID", binding.connection_id)?;
                optional_string_attr(
                    &mut xml,
                    "FileBindingName",
                    binding.file_binding_name.as_deref(),
                )?;
                optional_u32_attr(&mut xml, "DataBindingLoadMode", Some(binding.load_mode))?;
                if let Some(payload) = &binding.payload_xml {
                    xml.push_char('>')?;
                    xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?)?;
                    xml.push_str("</DataBinding></Map>")?;
                } else {
                    xml.push_str("/></Map>")?;
                }
            } else {
                xml.push_str("/>")?;
            }
        }
        xml.push_str("</MapInfo>")?;
        Ok(xml.finish())
*/

#[derive(Clone, Copy)]
enum Context {
    Root,
    Schema(usize),
    Map(usize),
    Binding(usize),
}
#[derive(Clone, Copy)]
enum Owner {
    Schema(usize),
    Binding(usize),
}
struct Capture {
    depth: usize,
    events: usize,
    owner: Owner,
    writer: Writer<Vec<u8>>,
}
struct SchemaBuilder {
    value: XmlMapSchema,
    bindings: Vec<(String, String)>,
}
struct BindingBuilder {
    value: XmlMapDataBinding,
    bindings: Vec<(String, String)>,
}
struct MapBuilder {
    value: XmlMap,
    binding: Option<BindingBuilder>,
}

fn parse_processed(xml: &[u8], limits: &XmlMapLimits) -> Result<ParsedXmlMapInfo> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut root_closed = false;
    let mut root_bindings = Vec::new();
    let mut selection = None;
    let mut conformance = None;
    let mut schemas: Vec<SchemaBuilder> = Vec::new();
    let mut maps: Vec<MapBuilder> = Vec::new();
    let mut capture: Option<Capture> = None;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        if let Some(active) = capture.as_mut() {
            active.events += 1;
            if active.events > limits.max_events {
                return Err(invalid("opaque XML event limit exceeded"));
            }
            match &event {
                Event::Start(_) => {
                    active.depth += 1;
                    if active.depth > limits.max_depth {
                        return Err(invalid("opaque XML depth limit exceeded"));
                    }
                },
                Event::End(_) => {
                    active.depth = active
                        .depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid opaque XML nesting"))?;
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("DTD and processing instructions are rejected"));
                },
                Event::Eof => return Err(invalid("unterminated opaque XML payload")),
                _ => {},
            }
            active
                .writer
                .write_event(event.clone())
                .map_err(xml_error)?;
            if active.writer.get_ref().len() > limits.max_opaque_bytes {
                return Err(invalid("opaque XML payload exceeds 16 MiB"));
            }
            if active.depth == 0 {
                let completed = capture
                    .take()
                    .ok_or_else(|| invalid("opaque XML capture state was lost"))?;
                assign_payload(
                    &mut schemas,
                    &mut maps,
                    completed.owner,
                    completed.writer.into_inner(),
                )?;
            }
            continue;
        }
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(e) if stack.is_empty() => {
                let observed = root_conformance(&namespace, &e, b"MapInfo");
                if root_closed || observed.is_none() {
                    return Err(invalid("expected one SpreadsheetML MapInfo root"));
                }
                conformance = observed;
                root_bindings = namespace_attributes(&e, decoder)?;
                selection = Some(required_attr(&e, decoder, b"SelectionNamespaces")?);
                only_attrs(&e, &[b"SelectionNamespaces"])?;
                stack.push(Context::Root);
            },
            Event::Start(e) => handle_start(
                &mut stack,
                &mut schemas,
                &mut maps,
                &root_bindings,
                &namespace,
                e,
                decoder,
                &mut capture,
                limits,
            )?,
            Event::Empty(e) => handle_empty(
                &mut stack,
                &mut schemas,
                &mut maps,
                &root_bindings,
                &namespace,
                e,
                decoder,
                limits,
            )?,
            Event::Text(t) => {
                if !t.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("text outside opaque XML payload"));
                }
            },
            Event::CData(t) => {
                if !t.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("CDATA outside opaque XML payload"));
                }
            },
            Event::GeneralRef(_) => {
                return Err(invalid("entity reference outside opaque XML payload"));
            },
            Event::End(_) => {
                let ended = stack
                    .pop()
                    .ok_or_else(|| invalid("closing element outside MapInfo"))?;
                if let Context::Binding(index) = ended {
                    let value = maps
                        .get(index)
                        .and_then(|map| map.binding.as_ref())
                        .map(|binding| binding.value.clone())
                        .ok_or_else(|| invalid("DataBinding context has no binding"))?;
                    maps.get_mut(index)
                        .ok_or_else(|| invalid("DataBinding context has no map"))?
                        .value
                        .data_binding = Some(value);
                }
                if matches!(ended, Context::Root) {
                    root_closed = true;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !root_closed || !stack.is_empty() {
        return Err(invalid("unterminated custom XML maps XML"));
    }
    let result = XmlMapInfo {
        selection_namespaces: selection.ok_or_else(|| invalid("missing MapInfo root"))?,
        schemas: schemas.into_iter().map(|v| v.value).collect(),
        maps: maps.into_iter().map(|v| v.value).collect(),
    };
    validate_xml_map_info_with_limits(&result, limits)?;
    Ok(ParsedXmlMapInfo {
        info: result,
        conformance: conformance.ok_or_else(|| invalid("missing MapInfo conformance"))?,
    })
}

#[allow(clippy::too_many_arguments)]
fn handle_start(
    stack: &mut Vec<Context>,
    schemas: &mut Vec<SchemaBuilder>,
    maps: &mut Vec<MapBuilder>,
    root_bindings: &[(String, String)],
    ns: &ResolveResult<'_>,
    e: BytesStart<'static>,
    decoder: Decoder,
    capture: &mut Option<Capture>,
    limits: &XmlMapLimits,
) -> Result<()> {
    match stack
        .last()
        .copied()
        .ok_or_else(|| invalid("element outside MapInfo"))?
    {
        Context::Root if core_name(ns, &e, b"Schema") => {
            if !maps.is_empty() || schemas.len() >= limits.max_schemas {
                return Err(invalid("invalid Schema order or limit"));
            }
            let value = parse_schema(&e, decoder)?;
            let bindings = merged_bindings(root_bindings, &namespace_attributes(&e, decoder)?);
            schemas.push(SchemaBuilder { value, bindings });
            stack.push(Context::Schema(schemas.len() - 1));
        },
        Context::Root if core_name(ns, &e, b"Map") => {
            if schemas.is_empty() || maps.len() >= limits.max_maps {
                return Err(invalid("invalid Map order or limit"));
            }
            maps.push(MapBuilder {
                value: parse_map(&e, decoder)?,
                binding: None,
            });
            stack.push(Context::Map(maps.len() - 1));
        },
        Context::Map(index) if core_name(ns, &e, b"DataBinding") => {
            if maps[index].binding.is_some() {
                return Err(invalid("duplicate DataBinding"));
            }
            let value = parse_binding(&e, decoder)?;
            let bindings = merged_bindings(root_bindings, &namespace_attributes(&e, decoder)?);
            maps[index].binding = Some(BindingBuilder { value, bindings });
            stack.push(Context::Binding(index));
        },
        Context::Schema(index) => {
            if schemas[index].value.payload_xml.is_some() {
                return Err(invalid("Schema permits at most one opaque child"));
            }
            *capture = Some(begin_capture(
                Owner::Schema(index),
                e,
                &schemas[index].bindings,
            )?);
        },
        Context::Binding(index) => {
            let binding = maps
                .get(index)
                .and_then(|map| map.binding.as_ref())
                .ok_or_else(|| invalid("DataBinding context has no binding"))?;
            if binding.value.payload_xml.is_some() {
                return Err(invalid("DataBinding permits at most one opaque child"));
            }
            *capture = Some(begin_capture(Owner::Binding(index), e, &binding.bindings)?);
        },
        _ => return Err(invalid("unexpected custom XML maps element")),
    }
    Ok(())
}

fn handle_empty(
    stack: &mut [Context],
    schemas: &mut Vec<SchemaBuilder>,
    maps: &mut Vec<MapBuilder>,
    root_bindings: &[(String, String)],
    ns: &ResolveResult<'_>,
    e: BytesStart<'static>,
    decoder: Decoder,
    limits: &XmlMapLimits,
) -> Result<()> {
    match stack.last().copied() {
        Some(Context::Root) if core_name(ns, &e, b"Schema") => {
            if !maps.is_empty() || schemas.len() >= limits.max_schemas {
                return Err(invalid("invalid Schema order or limit"));
            }
            schemas.push(SchemaBuilder {
                value: parse_schema(&e, decoder)?,
                bindings: merged_bindings(root_bindings, &namespace_attributes(&e, decoder)?),
            });
        },
        Some(Context::Root) if core_name(ns, &e, b"Map") => {
            if schemas.is_empty() || maps.len() >= limits.max_maps {
                return Err(invalid("invalid Map order or limit"));
            }
            maps.push(MapBuilder {
                value: parse_map(&e, decoder)?,
                binding: None,
            });
        },
        Some(Context::Map(index)) if core_name(ns, &e, b"DataBinding") => {
            if maps[index].binding.is_some() {
                return Err(invalid("duplicate DataBinding"));
            }
            let value = parse_binding(&e, decoder)?;
            maps[index].value.data_binding = Some(value);
        },
        Some(Context::Schema(index)) => {
            if schemas[index].value.payload_xml.is_some() {
                return Err(invalid("Schema permits at most one opaque child"));
            }
            schemas[index].value.payload_xml =
                Some(capture_empty(e, &schemas[index].bindings, limits)?);
        },
        Some(Context::Binding(index)) => {
            let binding = maps
                .get_mut(index)
                .and_then(|map| map.binding.as_mut())
                .ok_or_else(|| invalid("DataBinding context has no binding"))?;
            if binding.value.payload_xml.is_some() {
                return Err(invalid("DataBinding permits at most one opaque child"));
            }
            binding.value.payload_xml = Some(capture_empty(e, &binding.bindings, limits)?);
        },
        _ => return Err(invalid("unexpected empty custom XML maps element")),
    }
    Ok(())
}

fn parse_schema(e: &BytesStart<'_>, d: Decoder) -> Result<XmlMapSchema> {
    let id = required_attr(e, d, b"ID")?;
    let schema_reference = optional_attr(e, d, b"SchemaRef")?;
    let namespace = optional_attr(e, d, b"Namespace")?;
    only_attrs(e, &[b"ID", b"SchemaRef", b"Namespace", b"SchemaLanguage"])?;
    Ok(XmlMapSchema {
        id,
        schema_reference,
        namespace,
        payload_xml: None,
    })
}
fn parse_map(e: &BytesStart<'_>, d: Decoder) -> Result<XmlMap> {
    let id = required_u32_attr(e, d, b"ID")?;
    let name = required_attr(e, d, b"Name")?;
    let root_element = required_attr(e, d, b"RootElement")?;
    let schema_id = required_attr(e, d, b"SchemaID")?;
    let show_import_export_validation_errors =
        required_bool_attr(e, d, b"ShowImportExportValidationErrors")?;
    let auto_fit = required_bool_attr(e, d, b"AutoFit")?;
    let append = required_bool_attr(e, d, b"Append")?;
    let preserve_sort_auto_filter_layout = required_bool_attr(e, d, b"PreserveSortAFLayout")?;
    let preserve_format = required_bool_attr(e, d, b"PreserveFormat")?;
    only_attrs(
        e,
        &[
            b"ID",
            b"Name",
            b"RootElement",
            b"SchemaID",
            b"ShowImportExportValidationErrors",
            b"AutoFit",
            b"Append",
            b"PreserveSortAFLayout",
            b"PreserveFormat",
        ],
    )?;
    Ok(XmlMap {
        id,
        name,
        root_element,
        schema_id,
        show_import_export_validation_errors,
        auto_fit,
        append,
        preserve_sort_auto_filter_layout,
        preserve_format,
        data_binding: None,
    })
}
fn parse_binding(e: &BytesStart<'_>, d: Decoder) -> Result<XmlMapDataBinding> {
    let data_binding_name = optional_attr(e, d, b"DataBindingName")?;
    let file_binding = parse_bool_attr(e, d, b"FileBinding", false)?;
    let connection_id = parse_u32_attr(e, d, b"ConnectionID", false)?;
    let file_binding_name = optional_attr(e, d, b"FileBindingName")?;
    let load_mode = required_u32_attr(e, d, b"DataBindingLoadMode")?;
    only_attrs(
        e,
        &[
            b"DataBindingName",
            b"FileBinding",
            b"ConnectionID",
            b"FileBindingName",
            b"DataBindingLoadMode",
        ],
    )?;
    Ok(XmlMapDataBinding {
        data_binding_name,
        file_binding,
        connection_id,
        file_binding_name,
        load_mode,
        payload_xml: None,
    })
}

fn begin_capture(
    owner: Owner,
    e: BytesStart<'static>,
    bindings: &[(String, String)],
) -> Result<Capture> {
    let e = add_inherited_bindings(e, bindings)?;
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(e)).map_err(xml_error)?;
    Ok(Capture {
        depth: 1,
        events: 1,
        owner,
        writer,
    })
}
fn capture_empty(
    e: BytesStart<'static>,
    bindings: &[(String, String)],
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    let e = add_inherited_bindings(e, bindings)?;
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Empty(e)).map_err(xml_error)?;
    let bytes = writer.into_inner();
    if bytes.len() > limits.max_opaque_bytes {
        return Err(invalid("opaque XML payload exceeds 16 MiB"));
    }
    Ok(bytes)
}
fn assign_payload(
    schemas: &mut [SchemaBuilder],
    maps: &mut [MapBuilder],
    owner: Owner,
    payload: Vec<u8>,
) -> Result<()> {
    match owner {
        Owner::Schema(i) => schemas[i].value.payload_xml = Some(payload),
        Owner::Binding(i) => {
            let map = maps
                .get_mut(i)
                .ok_or_else(|| invalid("DataBinding payload has no map"))?;
            let binding = map
                .binding
                .as_mut()
                .ok_or_else(|| invalid("DataBinding payload has no binding"))?;
            binding.value.payload_xml = Some(payload);
            map.value.data_binding = Some(binding.value.clone());
        },
    }
    Ok(())
}

fn root_conformance(
    ns: &ResolveResult<'_>,
    e: &BytesStart<'_>,
    local: &[u8],
) -> Option<XmlMapConformance> {
    if e.local_name().as_ref() != local {
        return None;
    }
    match ns {
        ResolveResult::Bound(Namespace(value)) => {
            let bytes: &[u8] = value;
            match bytes {
                NS => Some(XmlMapConformance::Transitional),
                STRICT_NS => Some(XmlMapConformance::Strict),
                _ => None,
            }
        },
        _ => None,
    }
}
fn core_name(ns: &ResolveResult<'_>, e: &BytesStart<'_>, local: &[u8]) -> bool {
    (namespace_matches(ns) || matches!(ns, ResolveResult::Unbound))
        && e.local_name().as_ref() == local
}
fn namespace_matches(ns: &ResolveResult<'_>) -> bool {
    match ns {
        ResolveResult::Bound(Namespace(v)) => {
            let bytes: &[u8] = v;
            bytes == NS || bytes == STRICT_NS
        },
        _ => false,
    }
}
fn required_attr(e: &BytesStart<'_>, d: Decoder, n: &[u8]) -> Result<String> {
    optional_attr(e, d, n)?.ok_or_else(|| {
        invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(n)
        ))
    })
}
fn optional_attr(e: &BytesStart<'_>, d: Decoder, n: &[u8]) -> Result<Option<String>> {
    let mut value = None;
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        if a.key.as_ref() == n {
            if value.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            value = Some(
                a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                    .map_err(xml_error)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}
fn parse_bool_attr(
    e: &BytesStart<'_>,
    d: Decoder,
    n: &[u8],
    required: bool,
) -> Result<Option<bool>> {
    let value = optional_attr(e, d, n)?;
    match value.as_deref() {
        Some("true") => Ok(Some(true)),
        Some("false") => Ok(Some(false)),
        Some(_) => Err(invalid(format!(
            "invalid boolean attribute '{}'",
            String::from_utf8_lossy(n)
        ))),
        None if required => Err(invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(n)
        ))),
        None => Ok(None),
    }
}
fn parse_u32_attr(e: &BytesStart<'_>, d: Decoder, n: &[u8], required: bool) -> Result<Option<u32>> {
    match optional_attr(e, d, n)? {
        Some(v) => Ok(Some(v.parse().map_err(|_| {
            invalid(format!(
                "invalid unsigned integer attribute '{}'",
                String::from_utf8_lossy(n)
            ))
        })?)),
        None if required => Err(invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(n)
        ))),
        None => Ok(None),
    }
}
fn required_bool_attr(e: &BytesStart<'_>, d: Decoder, n: &[u8]) -> Result<bool> {
    parse_bool_attr(e, d, n, true)?.ok_or_else(|| {
        invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(n)
        ))
    })
}
fn required_u32_attr(e: &BytesStart<'_>, d: Decoder, n: &[u8]) -> Result<u32> {
    parse_u32_attr(e, d, n, true)?.ok_or_else(|| {
        invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(n)
        ))
    })
}
fn only_attrs(e: &BytesStart<'_>, allowed: &[&[u8]]) -> Result<()> {
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        let k = a.key.as_ref();
        if k == b"xmlns" || k.starts_with(b"xmlns:") {
            continue;
        }
        if k.contains(&b':') || !allowed.contains(&k) {
            return Err(invalid(format!(
                "unexpected attribute '{}'",
                String::from_utf8_lossy(k)
            )));
        }
    }
    Ok(())
}
fn namespace_attributes(e: &BytesStart<'_>, d: Decoder) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        let key = std::str::from_utf8(a.key.as_ref()).map_err(xml_error)?;
        if key == "xmlns" || key.starts_with("xmlns:") {
            values.push((
                key.to_string(),
                a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(values)
}
fn merged_bindings(
    parent: &[(String, String)],
    local: &[(String, String)],
) -> Vec<(String, String)> {
    let mut result = parent.to_vec();
    for (key, value) in local {
        if let Some(existing) = result.iter_mut().find(|v| v.0 == *key) {
            existing.1 = value.clone();
        } else {
            result.push((key.clone(), value.clone()));
        }
    }
    result
}
fn add_inherited_bindings(
    mut e: BytesStart<'static>,
    bindings: &[(String, String)],
) -> Result<BytesStart<'static>> {
    let declared: HashSet<Vec<u8>> = e
        .attributes()
        .with_checks(true)
        .filter_map(std::result::Result::ok)
        .filter(|a| a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:"))
        .map(|a| a.key.as_ref().to_vec())
        .collect();
    for (key, value) in bindings {
        if !declared.contains(key.as_bytes()) {
            e.push_attribute((key.as_str(), value.as_str()));
        }
    }
    Ok(e)
}
struct BoundedXml {
    bytes: Vec<u8>,
    max_part_bytes: usize,
}

impl BoundedXml {
    fn new(max_part_bytes: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max_part_bytes,
        }
    }

    fn push_str(&mut self, value: &str) -> Result<()> {
        self.push_bytes(value.as_bytes())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        let mut encoded = [0; 4];
        let length = value.encode_utf8(&mut encoded).len();
        self.push_bytes(&encoded[..length])
    }

    fn push_bytes(&mut self, value: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid("serialized custom XML maps length overflows"))?;
        if length > self.max_part_bytes {
            return Err(invalid("serialized custom XML maps part exceeds 32 MiB"));
        }
        self.bytes
            .try_reserve(value.len())
            .map_err(|_| invalid("serialized custom XML maps output allocation failed"))?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn optional_string_attr(xml: &mut BoundedXml, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        xml.push_char(' ')?;
        xml.push_str(name)?;
        xml.push_str("=\"")?;
        escape_attr(xml, value)?;
        xml.push_char('"')?;
    }
    Ok(())
}
fn optional_bool_attr(xml: &mut BoundedXml, name: &str, value: Option<bool>) -> Result<()> {
    if let Some(value) = value {
        bool_attr(xml, name, value)?;
    }
    Ok(())
}
fn optional_u32_attr(xml: &mut BoundedXml, name: &str, value: Option<u32>) -> Result<()> {
    if let Some(value) = value {
        xml.push_char(' ')?;
        xml.push_str(name)?;
        xml.push_str("=\"")?;
        xml.push_str(&value.to_string())?;
        xml.push_char('"')?;
    }
    Ok(())
}
fn bool_attr(xml: &mut BoundedXml, name: &str, value: bool) -> Result<()> {
    xml.push_char(' ')?;
    xml.push_str(name)?;
    xml.push_str(if value { "=\"true\"" } else { "=\"false\"" })?;
    Ok(())
}
fn escape_attr(out: &mut BoundedXml, value: &str) -> Result<()> {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;")?,
            '<' => out.push_str("&lt;")?,
            '"' => out.push_str("&quot;")?,
            '\r' => out.push_str("&#xD;")?,
            '\n' => out.push_str("&#xA;")?,
            '\t' => out.push_str("&#x9;")?,
            _ => out.push_char(c)?,
        }
    }
    Ok(())
}
fn xml_error(e: impl std::fmt::Display) -> crate::Error {
    invalid(e.to_string())
}

/// Patch known Custom XML Maps fields inside their source XML spans.
///
/// The semantic parser intentionally projects schema and binding payloads into
/// bounded inert bytes. Transactions use this source patcher so changing a
/// typed attribute does not discard producer extensions, comments, namespace
/// spelling, or the original payload markup. Structural catalog replacement
/// falls back to the already validated canonical writer.
fn patch_xml_map_info_source_impl(
    source: &[u8],
    before: &XmlMapInfoRef<'_>,
    after: &XmlMapInfoRef<'_>,
    before_strict: bool,
    after_strict: bool,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    if source.len() > limits.max_part_bytes {
        return Err(invalid(
            "custom XML maps source exceeds configured part limit",
        ));
    }
    validate_xml_map_info_ref_with_limits(before, limits)?;
    validate_xml_map_info_ref_with_limits(after, limits)?;
    if before == after && before_strict == after_strict {
        return Ok(source.to_vec());
    }

    let tree = SourceTree::parse(source, limits)?;
    let after_schemas: HashMap<&str, &XmlSchemaRef<'_>> = after
        .schemas
        .iter()
        .map(|schema| (schema.id, schema))
        .collect();
    let after_maps: HashMap<u32, &XmlMapRef<'_>> =
        after.maps.iter().map(|map| (map.id, map)).collect();
    let root = tree.root;
    if tree.nodes[root].local != "MapInfo" || tree.nodes[root].self_closing {
        return serialize_xml_map_info_ref_with_limits(
            after,
            if after_strict {
                XmlMapConformance::Strict
            } else {
                XmlMapConformance::Transitional
            },
            limits,
        );
    }

    let schema_nodes = tree.direct_children(root, "Schema");
    let map_nodes = tree.direct_children(root, "Map");
    if schema_nodes.len() != before.schemas.len() || map_nodes.len() != before.maps.len() {
        return serialize_xml_map_info_ref_with_limits(after, conformance(after_strict), limits);
    }
    for (node, schema) in schema_nodes.iter().zip(&before.schemas) {
        if tree.attribute_value(source, *node, "ID")?.as_deref() != Some(schema.id) {
            return serialize_xml_map_info_ref_with_limits(
                after,
                conformance(after_strict),
                limits,
            );
        }
        if tree.nodes[*node].self_closing
            && schema.payload_xml.is_none()
            && after_schemas
                .get(schema.id)
                .and_then(|candidate| candidate.payload_xml.as_ref())
                .is_some()
        {
            return serialize_xml_map_info_ref_with_limits(
                after,
                conformance(after_strict),
                limits,
            );
        }
    }
    for (node, map) in map_nodes.iter().zip(&before.maps) {
        let id = tree
            .attribute_value(source, *node, "ID")?
            .and_then(|value| value.parse::<u32>().ok());
        if id != Some(map.id) {
            return serialize_xml_map_info_ref_with_limits(
                after,
                conformance(after_strict),
                limits,
            );
        }
        if tree.nodes[*node].self_closing
            && map.data_binding.is_none()
            && after_maps
                .get(&map.id)
                .and_then(|candidate| candidate.data_binding.as_ref())
                .is_some()
        {
            return serialize_xml_map_info_ref_with_limits(
                after,
                conformance(after_strict),
                limits,
            );
        }
    }

    let mut edits = Vec::new();
    patch_attribute(
        &tree,
        root,
        "SelectionNamespaces",
        Some(before.selection_namespaces),
        Some(after.selection_namespaces),
        &mut edits,
    )?;
    if before_strict != after_strict {
        patch_attribute(
            &tree,
            root,
            "xmlns",
            Some(if before_strict {
                STRICT_NS_TEXT
            } else {
                NS_TEXT
            }),
            Some(if after_strict {
                STRICT_NS_TEXT
            } else {
                NS_TEXT
            }),
            &mut edits,
        )?;
    }

    for (node, before_schema) in schema_nodes.iter().zip(&before.schemas) {
        let after_schema = after_schemas
            .get(before_schema.id)
            .copied()
            .ok_or_else(|| invalid("Schema identity changed during source patching"));
        let Ok(after_schema) = after_schema else {
            return serialize_xml_map_info_ref_with_limits(
                after,
                conformance(after_strict),
                limits,
            );
        };
        patch_attribute(
            &tree,
            *node,
            "ID",
            Some(before_schema.id),
            Some(after_schema.id),
            &mut edits,
        )?;
        patch_attribute(
            &tree,
            *node,
            "SchemaRef",
            before_schema.schema_reference,
            after_schema.schema_reference,
            &mut edits,
        )?;
        patch_attribute(
            &tree,
            *node,
            "Namespace",
            before_schema.namespace,
            after_schema.namespace,
            &mut edits,
        )?;
        patch_payload(
            &tree,
            source,
            *node,
            before_schema.payload_xml,
            after_schema.payload_xml,
            &mut edits,
        )?;
    }

    for (node, before_map) in map_nodes.iter().zip(&before.maps) {
        let after_map = after_maps
            .get(&before_map.id)
            .copied()
            .ok_or_else(|| invalid("Map identity changed during source patching"));
        let Ok(after_map) = after_map else {
            return serialize_xml_map_info_ref_with_limits(
                after,
                conformance(after_strict),
                limits,
            );
        };
        patch_attribute(
            &tree,
            *node,
            "ID",
            Some(&before_map.id.to_string()),
            Some(&after_map.id.to_string()),
            &mut edits,
        )?;
        patch_attribute(
            &tree,
            *node,
            "Name",
            Some(before_map.name),
            Some(after_map.name),
            &mut edits,
        )?;
        patch_attribute(
            &tree,
            *node,
            "RootElement",
            Some(before_map.root_element),
            Some(after_map.root_element),
            &mut edits,
        )?;
        patch_attribute(
            &tree,
            *node,
            "SchemaID",
            Some(before_map.schema_id),
            Some(after_map.schema_id),
            &mut edits,
        )?;
        patch_bool_attribute(
            &tree,
            *node,
            "ShowImportExportValidationErrors",
            before_map.show_import_export_validation_errors,
            after_map.show_import_export_validation_errors,
            &mut edits,
        )?;
        patch_bool_attribute(
            &tree,
            *node,
            "AutoFit",
            before_map.auto_fit,
            after_map.auto_fit,
            &mut edits,
        )?;
        patch_bool_attribute(
            &tree,
            *node,
            "Append",
            before_map.append,
            after_map.append,
            &mut edits,
        )?;
        patch_bool_attribute(
            &tree,
            *node,
            "PreserveSortAFLayout",
            before_map.preserve_sort_auto_filter_layout,
            after_map.preserve_sort_auto_filter_layout,
            &mut edits,
        )?;
        patch_bool_attribute(
            &tree,
            *node,
            "PreserveFormat",
            before_map.preserve_format,
            after_map.preserve_format,
            &mut edits,
        )?;

        let binding_node = tree.direct_child(*node, "DataBinding")?;
        match (
            before_map.data_binding.as_ref(),
            after_map.data_binding.as_ref(),
            binding_node,
        ) {
            (Some(before_binding), Some(after_binding), Some(binding_node)) => {
                patch_attribute(
                    &tree,
                    binding_node,
                    "DataBindingName",
                    before_binding.data_binding_name,
                    after_binding.data_binding_name,
                    &mut edits,
                )?;
                patch_bool_optional_attribute(
                    &tree,
                    binding_node,
                    "FileBinding",
                    before_binding.file_binding,
                    after_binding.file_binding,
                    &mut edits,
                )?;
                patch_number_optional_attribute(
                    &tree,
                    binding_node,
                    "ConnectionID",
                    before_binding.connection_id,
                    after_binding.connection_id,
                    &mut edits,
                )?;
                patch_attribute(
                    &tree,
                    binding_node,
                    "FileBindingName",
                    before_binding.file_binding_name,
                    after_binding.file_binding_name,
                    &mut edits,
                )?;
                patch_number_attribute(
                    &tree,
                    binding_node,
                    "DataBindingLoadMode",
                    before_binding.load_mode,
                    after_binding.load_mode,
                    &mut edits,
                )?;
                patch_payload(
                    &tree,
                    source,
                    binding_node,
                    before_binding.payload_xml,
                    after_binding.payload_xml,
                    &mut edits,
                )?;
            },
            (None, Some(after_binding), None) => {
                let replacement = binding_fragment(after_binding, limits)?;
                insert_child(&tree, *node, replacement, &mut edits)?;
            },
            (Some(_), None, Some(binding_node)) => edits.push(SourceEdit {
                range: tree.nodes[binding_node].start..tree.nodes[binding_node].end,
                replacement: Vec::new(),
            }),
            (None, None, None) => {},
            _ => {
                return serialize_xml_map_info_ref_with_limits(
                    after,
                    conformance(after_strict),
                    limits,
                );
            },
        }
    }

    apply_source_edits(source, edits, limits)
}

fn patch_payload(
    tree: &SourceTree,
    source: &[u8],
    parent: usize,
    before: Option<&[u8]>,
    after: Option<&[u8]>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    let existing = tree.element_children(parent).first().copied();
    match (existing, after) {
        (Some(node), Some(payload)) => edits.push(SourceEdit {
            range: tree.nodes[node].start..tree.nodes[node].end,
            replacement: payload.to_vec(),
        }),
        (Some(node), None) => edits.push(SourceEdit {
            range: tree.nodes[node].start..tree.nodes[node].end,
            replacement: Vec::new(),
        }),
        (None, Some(payload)) => insert_child(tree, parent, payload.to_vec(), edits)?,
        (None, None) => {},
    }
    let _ = source;
    Ok(())
}

fn insert_child(
    tree: &SourceTree,
    parent: usize,
    replacement: Vec<u8>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if tree.nodes[parent].self_closing {
        return Err(invalid(
            "custom XML maps source needs a structural child insertion",
        ));
    }
    edits.push(SourceEdit {
        range: tree.nodes[parent].end_start..tree.nodes[parent].end_start,
        replacement,
    });
    Ok(())
}

fn patch_bool_attribute(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: bool,
    after: bool,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    patch_attribute(
        tree,
        node,
        name,
        Some(if before { "true" } else { "false" }),
        Some(if after { "true" } else { "false" }),
        edits,
    )
}

fn patch_bool_optional_attribute(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<bool>,
    after: Option<bool>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    patch_attribute(
        tree,
        node,
        name,
        before.map(|value| if value { "true" } else { "false" }),
        after.map(|value| if value { "true" } else { "false" }),
        edits,
    )
}

fn patch_number_attribute(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: u32,
    after: u32,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    patch_attribute(
        tree,
        node,
        name,
        Some(&before.to_string()),
        Some(&after.to_string()),
        edits,
    )
}

fn patch_number_optional_attribute(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<u32>,
    after: Option<u32>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    let before = before.map(|value| value.to_string());
    let after = after.map(|value| value.to_string());
    patch_attribute(tree, node, name, before.as_deref(), after.as_deref(), edits)
}

fn patch_attribute(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<&str>,
    after: Option<&str>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    let attribute = tree.attribute(node, name);
    match (attribute, after) {
        (Some(attribute), Some(value)) => edits.push(SourceEdit {
            range: attribute.value_start..attribute.value_end,
            replacement: escape_source_attribute(value),
        }),
        (Some(attribute), None) => edits.push(SourceEdit {
            range: attribute.start..attribute.value_end + 1,
            replacement: Vec::new(),
        }),
        (None, Some(value)) if before.is_none() => edits.push(SourceEdit {
            range: tree.nodes[node].close_pos..tree.nodes[node].close_pos,
            replacement: format!(
                " {name}=\"{}\"",
                String::from_utf8_lossy(&escape_source_attribute(value))
            )
            .into_bytes(),
        }),
        (None, Some(_)) => {
            return Err(invalid(format!(
                "custom XML maps source is missing attribute '{name}'"
            )));
        },
        (None, None) => {
            if before.is_some() {
                return Err(invalid(format!(
                    "custom XML maps source is missing attribute '{name}'"
                )));
            }
        },
    }
    Ok(())
}

fn binding_fragment(binding: &DataBindingRef<'_>, limits: &XmlMapLimits) -> Result<Vec<u8>> {
    let mut xml = BoundedXml::new(limits.max_part_bytes);
    xml.push_str("<DataBinding")?;
    optional_string_attr(&mut xml, "DataBindingName", binding.data_binding_name)?;
    optional_bool_attr(&mut xml, "FileBinding", binding.file_binding)?;
    optional_u32_attr(&mut xml, "ConnectionID", binding.connection_id)?;
    optional_string_attr(&mut xml, "FileBindingName", binding.file_binding_name)?;
    optional_u32_attr(&mut xml, "DataBindingLoadMode", Some(binding.load_mode))?;
    if let Some(payload) = &binding.payload_xml {
        xml.push_char('>')?;
        xml.push_bytes(payload)?;
        xml.push_str("</DataBinding>")?;
    } else {
        xml.push_str("/>")?;
    }
    Ok(xml.finish())
}

fn escape_source_attribute(value: &str) -> Vec<u8> {
    let mut result = Vec::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => result.extend_from_slice(b"&amp;"),
            '<' => result.extend_from_slice(b"&lt;"),
            '"' => result.extend_from_slice(b"&quot;"),
            '\r' => result.extend_from_slice(b"&#xD;"),
            '\n' => result.extend_from_slice(b"&#xA;"),
            '\t' => result.extend_from_slice(b"&#x9;"),
            _ => {
                let mut encoded = [0; 4];
                result.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
    result
}

#[derive(Debug)]
struct SourceEdit {
    range: std::ops::Range<usize>,
    replacement: Vec<u8>,
}

fn apply_source_edits(
    source: &[u8],
    mut edits: Vec<SourceEdit>,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    edits.sort_by_key(|edit| edit.range.start);
    for pair in edits.windows(2) {
        if pair[0].range.end > pair[1].range.start {
            return Err(invalid("custom XML maps source edits overlap"));
        }
    }
    let mut final_len = source.len();
    for edit in &edits {
        final_len = final_len
            .checked_sub(edit.range.len())
            .and_then(|length| length.checked_add(edit.replacement.len()))
            .ok_or_else(|| invalid("serialized custom XML maps length overflows"))?;
    }
    if final_len > limits.max_part_bytes {
        return Err(invalid(
            "serialized custom XML maps part exceeds configured limit",
        ));
    }
    let mut result = Vec::new();
    result
        .try_reserve(final_len)
        .map_err(|_| invalid("serialized custom XML maps output allocation failed"))?;
    let mut copied = 0usize;
    for edit in edits {
        if edit.range.end > source.len() {
            return Err(invalid("custom XML maps source edit is out of bounds"));
        }
        result.extend_from_slice(&source[copied..edit.range.start]);
        result.extend_from_slice(&edit.replacement);
        copied = edit.range.end;
    }
    result.extend_from_slice(&source[copied..]);
    debug_assert_eq!(result.len(), final_len);
    Ok(result)
}

#[derive(Debug)]
struct SourceTree {
    nodes: Vec<SourceNode>,
    root: usize,
}

#[derive(Debug)]
struct SourceNode {
    local: String,
    start: usize,
    end_start: usize,
    end: usize,
    close_pos: usize,
    self_closing: bool,
    attrs: Vec<SourceAttribute>,
    children: Vec<usize>,
}

#[derive(Debug)]
struct SourceAttribute {
    name: String,
    local: String,
    start: usize,
    value_start: usize,
    value_end: usize,
}

impl SourceTree {
    fn parse(source: &[u8], limits: &XmlMapLimits) -> Result<Self> {
        let mut nodes = Vec::<SourceNode>::new();
        let mut stack = Vec::<usize>::new();
        let mut root = None::<usize>;
        let mut position = 0;
        while position < source.len() {
            if source[position] != b'<' {
                position += 1;
                continue;
            }
            if source[position..].starts_with(b"<?") {
                position = find_source_bytes(source, position + 2, b"?>")? + 2;
                continue;
            }
            if source[position..].starts_with(b"<!--") {
                position = find_source_bytes(source, position + 4, b"-->")? + 3;
                continue;
            }
            if source[position..].starts_with(b"<![CDATA[") {
                position = find_source_bytes(source, position + 9, b"]]>")? + 3;
                continue;
            }
            if source[position..].starts_with(b"<!") {
                position = source_tag_end(source, position)? + 1;
                continue;
            }
            if source[position..].starts_with(b"</") {
                let end = source_tag_end(source, position)?;
                let name_start = position + 2;
                let name_end = source_name_end(source, name_start);
                let node = stack.pop().ok_or_else(|| {
                    invalid("custom XML maps source has an unmatched closing tag")
                })?;
                if nodes[node].local != source_local(&source[name_start..name_end])? {
                    return Err(invalid("custom XML maps source has mismatched tags"));
                }
                nodes[node].end_start = position;
                nodes[node].end = end + 1;
                position = end + 1;
                continue;
            }
            let end = source_tag_end(source, position)?;
            let (local, attrs, close_pos, self_closing) = source_start_tag(source, position, end)?;
            let node = nodes.len();
            if node >= limits.max_events {
                return Err(invalid("custom XML maps source node limit exceeded"));
            }
            nodes.push(SourceNode {
                local,
                start: position,
                end_start: end + 1,
                end: end + 1,
                close_pos,
                self_closing,
                attrs,
                children: Vec::new(),
            });
            if let Some(parent) = stack.last().copied() {
                nodes[parent].children.push(node);
            } else if root.replace(node).is_some() {
                return Err(invalid("custom XML maps source has multiple roots"));
            }
            if !self_closing {
                if stack.len() >= limits.max_depth {
                    return Err(invalid("custom XML maps source depth limit exceeded"));
                }
                stack.push(node);
            }
            position = end + 1;
        }
        if !stack.is_empty() {
            return Err(invalid("custom XML maps source has unterminated markup"));
        }
        Ok(Self {
            nodes,
            root: root.ok_or_else(|| invalid("custom XML maps source has no root"))?,
        })
    }

    fn attribute(&self, node: usize, name: &str) -> Option<&SourceAttribute> {
        self.nodes[node].attrs.iter().find(|attribute| {
            attribute.name == name
                || (!name.eq_ignore_ascii_case("xmlns") && attribute.local == name)
        })
    }

    fn attribute_value(&self, source: &[u8], node: usize, name: &str) -> Result<Option<String>> {
        self.attribute(node, name)
            .map(|attribute| {
                let raw = std::str::from_utf8(&source[attribute.value_start..attribute.value_end])
                    .map_err(xml_error)?;
                quick_xml::escape::unescape(raw)
                    .map(std::borrow::Cow::into_owned)
                    .map_err(xml_error)
            })
            .transpose()
    }

    fn direct_children(&self, node: usize, local: &str) -> Vec<usize> {
        self.nodes[node]
            .children
            .iter()
            .copied()
            .filter(|child| self.nodes[*child].local == local)
            .collect()
    }

    fn direct_child(&self, node: usize, local: &str) -> Result<Option<usize>> {
        let mut result = None;
        for child in self.nodes[node]
            .children
            .iter()
            .copied()
            .filter(|child| self.nodes[*child].local == local)
        {
            if result.replace(child).is_some() {
                return Err(invalid(format!(
                    "custom XML maps source has duplicate '{local}'"
                )));
            }
        }
        Ok(result)
    }

    fn element_children(&self, node: usize) -> Vec<usize> {
        self.nodes[node].children.clone()
    }
}

fn find_source_bytes(source: &[u8], start: usize, needle: &[u8]) -> Result<usize> {
    source[start..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|position| start + position)
        .ok_or_else(|| invalid("custom XML maps source has an unterminated declaration"))
}

fn source_tag_end(source: &[u8], start: usize) -> Result<usize> {
    let mut quote = None;
    for (offset, byte) in source[start + 1..].iter().enumerate() {
        match (quote, byte) {
            (Some(value), byte) if *byte == value => quote = None,
            (None, b'\'' | b'\"') => quote = Some(*byte),
            (None, b'>') => return Ok(start + 1 + offset),
            _ => {},
        }
    }
    Err(invalid("custom XML maps source has an unterminated tag"))
}

fn source_start_tag(
    source: &[u8],
    start: usize,
    end: usize,
) -> Result<(String, Vec<SourceAttribute>, usize, bool)> {
    let name_start = start + 1;
    let name_end = source_name_end(source, name_start);
    if name_start == name_end {
        return Err(invalid("custom XML maps source has an empty element name"));
    }
    let mut attrs = Vec::new();
    let mut position = name_end;
    while position < end {
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= end || source[position] == b'/' {
            break;
        }
        let attribute_start = position;
        let attribute_end = source_name_end(source, position);
        if attribute_end == attribute_start {
            return Err(invalid("custom XML maps source has an invalid attribute"));
        }
        position = attribute_end;
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= end || source[position] != b'=' {
            return Err(invalid("custom XML maps source attribute is missing '='"));
        }
        position += 1;
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        let quote = *source
            .get(position)
            .ok_or_else(|| invalid("custom XML maps source attribute is missing quotes"))?;
        if !matches!(quote, b'\'' | b'\"') {
            return Err(invalid(
                "custom XML maps source attribute is missing quotes",
            ));
        }
        position += 1;
        let value_start = position;
        while position < end && source[position] != quote {
            position += 1;
        }
        if position >= end {
            return Err(invalid("custom XML maps source attribute is unterminated"));
        }
        let value_end = position;
        position += 1;
        let name = std::str::from_utf8(&source[attribute_start..attribute_end])
            .map_err(xml_error)?
            .to_owned();
        attrs.push(SourceAttribute {
            local: source_local(name.as_bytes())?,
            name,
            start: attribute_start,
            value_start,
            value_end,
        });
    }
    let self_closing = source[..end]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .is_some_and(|position| source[position] == b'/');
    let close_pos = if self_closing {
        source[..end]
            .iter()
            .rposition(|byte| !byte.is_ascii_whitespace())
            .unwrap_or(end)
    } else {
        end
    };
    Ok((
        source_local(&source[name_start..name_end])?,
        attrs,
        close_pos,
        self_closing,
    ))
}

fn source_name_end(source: &[u8], mut position: usize) -> usize {
    while position < source.len()
        && !source[position].is_ascii_whitespace()
        && !matches!(source[position], b'/' | b'>' | b'=')
    {
        position += 1;
    }
    position
}

fn source_local(value: &[u8]) -> Result<String> {
    let value = std::str::from_utf8(value).map_err(xml_error)?;
    Ok(value.rsplit(':').next().unwrap_or(value).to_owned())
}

/// Parse a bounded, namespace-aware, MCE-processed `SpreadsheetML` `MapInfo` part.
pub fn parse_xml_map_info(xml: &[u8]) -> Result<XmlMapInfo> {
    XmlMapInfo::parse(xml)
}

/// Parse with caller-selected resource ceilings.
pub fn parse_xml_map_info_with_limits(xml: &[u8], limits: &XmlMapLimits) -> Result<XmlMapInfo> {
    XmlMapInfo::parse_with_limits(xml, limits)
}

/// Parse and report the `SpreadsheetML` namespace family observed at the root.
pub fn parse_xml_map_info_with_conformance(xml: &[u8]) -> Result<ParsedXmlMapInfo> {
    parse_xml_map_info_with_conformance_and_limits(xml, &XmlMapLimits::DEFAULT)
}

/// Parse with caller ceilings and report the root namespace family.
pub fn parse_xml_map_info_with_conformance_and_limits(
    xml: &[u8],
    limits: &XmlMapLimits,
) -> Result<ParsedXmlMapInfo> {
    parse_xml_map_info_with_conformance_and_limits_impl(xml, limits)
}

/// Serialize `MapInfo` canonically for the selected OOXML conformance family.
pub fn serialize_xml_map_info(
    info: &XmlMapInfo,
    conformance: XmlMapConformance,
) -> Result<Vec<u8>> {
    serialize_xml_map_info_with_limits(info, conformance, &XmlMapLimits::DEFAULT)
}

/// Serialize using caller-selected resource ceilings.
pub fn serialize_xml_map_info_with_limits(
    info: &XmlMapInfo,
    conformance: XmlMapConformance,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    let info = XmlMapInfoRef::from_owned_with_limits(info, limits)?;
    serialize_xml_map_info_ref_with_limits(&info, conformance, limits)
}

/// Serialize a borrowed `MapInfo` projection using default resource ceilings.
pub fn serialize_xml_map_info_ref(
    info: &XmlMapInfoRef<'_>,
    conformance: XmlMapConformance,
) -> Result<Vec<u8>> {
    serialize_xml_map_info_ref_with_limits(info, conformance, &XmlMapLimits::DEFAULT)
}

/// Serialize a borrowed `MapInfo` projection without cloning referenced data.
pub fn serialize_xml_map_info_ref_with_limits(
    info: &XmlMapInfoRef<'_>,
    conformance: XmlMapConformance,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    validate_xml_map_info_ref_with_limits(info, limits)?;
    let mut xml = BoundedXml::new(limits.max_part_bytes);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><MapInfo xmlns=\"")?;
    escape_attr(
        &mut xml,
        if conformance.is_strict() {
            STRICT_NS_TEXT
        } else {
            NS_TEXT
        },
    )?;
    xml.push_str("\" SelectionNamespaces=\"")?;
    escape_attr(&mut xml, info.selection_namespaces)?;
    xml.push_str("\">")?;
    for schema in &info.schemas {
        xml.push_str("<Schema ID=\"")?;
        escape_attr(&mut xml, schema.id)?;
        xml.push_char('"')?;
        optional_string_attr(&mut xml, "SchemaRef", schema.schema_reference)?;
        optional_string_attr(&mut xml, "Namespace", schema.namespace)?;
        if let Some(payload) = schema.payload_xml {
            xml.push_char('>')?;
            xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?)?;
            xml.push_str("</Schema>")?;
        } else {
            xml.push_str("/>")?;
        }
    }
    for map in &info.maps {
        xml.push_str("<Map ID=\"")?;
        xml.push_str(&map.id.to_string())?;
        xml.push_str("\" Name=\"")?;
        escape_attr(&mut xml, map.name)?;
        xml.push_str("\" RootElement=\"")?;
        escape_attr(&mut xml, map.root_element)?;
        xml.push_str("\" SchemaID=\"")?;
        escape_attr(&mut xml, map.schema_id)?;
        xml.push_char('"')?;
        bool_attr(
            &mut xml,
            "ShowImportExportValidationErrors",
            map.show_import_export_validation_errors,
        )?;
        bool_attr(&mut xml, "AutoFit", map.auto_fit)?;
        bool_attr(&mut xml, "Append", map.append)?;
        bool_attr(
            &mut xml,
            "PreserveSortAFLayout",
            map.preserve_sort_auto_filter_layout,
        )?;
        bool_attr(&mut xml, "PreserveFormat", map.preserve_format)?;
        if let Some(binding) = map.data_binding {
            xml.push_str("><DataBinding")?;
            optional_string_attr(&mut xml, "DataBindingName", binding.data_binding_name)?;
            optional_bool_attr(&mut xml, "FileBinding", binding.file_binding)?;
            optional_u32_attr(&mut xml, "ConnectionID", binding.connection_id)?;
            optional_string_attr(&mut xml, "FileBindingName", binding.file_binding_name)?;
            optional_u32_attr(&mut xml, "DataBindingLoadMode", Some(binding.load_mode))?;
            if let Some(payload) = binding.payload_xml {
                xml.push_char('>')?;
                xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?)?;
                xml.push_str("</DataBinding></Map>")?;
            } else {
                xml.push_str("/></Map>")?;
            }
        } else {
            xml.push_str("/>")?;
        }
    }
    xml.push_str("</MapInfo>")?;
    Ok(xml.finish())
}

fn conformance(strict: bool) -> XmlMapConformance {
    if strict {
        XmlMapConformance::Strict
    } else {
        XmlMapConformance::Transitional
    }
}

/// Patch modeled fields while preserving unaffected source spelling and markup.
pub fn patch_xml_map_info_source(
    source: &[u8],
    before: &XmlMapInfo,
    after: &XmlMapInfo,
    before_conformance: XmlMapConformance,
    after_conformance: XmlMapConformance,
) -> Result<Vec<u8>> {
    patch_xml_map_info_source_with_limits(
        source,
        before,
        after,
        before_conformance,
        after_conformance,
        &XmlMapLimits::DEFAULT,
    )
}

/// Patch modeled fields using caller-selected validation and output ceilings.
pub fn patch_xml_map_info_source_with_limits(
    source: &[u8],
    before: &XmlMapInfo,
    after: &XmlMapInfo,
    before_conformance: XmlMapConformance,
    after_conformance: XmlMapConformance,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    let before = XmlMapInfoRef::from_owned_with_limits(before, limits)?;
    let after = XmlMapInfoRef::from_owned_with_limits(after, limits)?;
    patch_xml_map_info_source_ref_with_limits(
        source,
        &before,
        &after,
        before_conformance,
        after_conformance,
        limits,
    )
}

/// Patch source from borrowed projections using default resource ceilings.
pub fn patch_xml_map_info_source_ref(
    source: &[u8],
    before: &XmlMapInfoRef<'_>,
    after: &XmlMapInfoRef<'_>,
    before_conformance: XmlMapConformance,
    after_conformance: XmlMapConformance,
) -> Result<Vec<u8>> {
    patch_xml_map_info_source_ref_with_limits(
        source,
        before,
        after,
        before_conformance,
        after_conformance,
        &XmlMapLimits::DEFAULT,
    )
}

/// Patch source from borrowed projections using caller-selected ceilings.
pub fn patch_xml_map_info_source_ref_with_limits(
    source: &[u8],
    before: &XmlMapInfoRef<'_>,
    after: &XmlMapInfoRef<'_>,
    before_conformance: XmlMapConformance,
    after_conformance: XmlMapConformance,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    patch_xml_map_info_source_impl(
        source,
        before,
        after,
        before_conformance.is_strict(),
        after_conformance.is_strict(),
        limits,
    )
}
