//! SpreadsheetML Custom XML Maps with inert, bounded inline schema payloads.

use std::collections::HashSet;

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer, XmlVersion};

const NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/xmlMaps";
const CONTENT_TYPE: &str = "application/xml";
const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_SCHEMAS: usize = 4_096;
const MAX_MAPS: usize = 65_536;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_OPAQUE_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;

/// Namespace family used for a Custom XML Maps part and its workbook relationship.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum XmlMapConformance {
    #[default]
    Transitional,
    Strict,
}

impl XmlMapConformance {
    const fn relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }

    /// Whether this conformance uses ISO/IEC 29500 Strict namespace URIs.
    pub const fn is_strict(self) -> bool {
        matches!(self, Self::Strict)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapSchema {
    pub id: String,
    pub schema_reference: Option<String>,
    pub namespace: Option<String>,
    /// One schema-valid `xsd:any` element, stored without interpretation or resolution.
    pub payload_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapDataBinding {
    pub data_binding_name: Option<String>,
    pub file_binding: Option<bool>,
    pub connection_id: Option<u32>,
    pub file_binding_name: Option<String>,
    pub load_mode: u32,
    /// One schema-valid `xsd:any` element, stored without interpretation or execution.
    pub payload_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMap {
    pub id: u32,
    pub name: String,
    pub root_element: String,
    pub schema_id: String,
    pub show_import_export_validation_errors: bool,
    pub auto_fit: bool,
    pub append: bool,
    pub preserve_sort_auto_filter_layout: bool,
    pub preserve_format: bool,
    pub data_binding: Option<XmlMapDataBinding>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XmlMapInfo {
    pub selection_namespaces: String,
    pub schemas: Vec<XmlMapSchema>,
    pub maps: Vec<XmlMap>,
}

impl XmlMapInfo {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_PART_BYTES {
            return Err(invalid("custom XML maps part exceeds 32 MiB"));
        }
        let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
        if processed.len() > MAX_PART_BYTES {
            return Err(invalid("processed custom XML maps part exceeds 32 MiB"));
        }
        std::str::from_utf8(processed.as_ref()).map_err(xml_error)?;
        parse_processed(processed.as_ref())
    }

    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate(self)?;
        let namespace = if strict {
            std::str::from_utf8(STRICT_NS).map_err(xml_error)?
        } else {
            std::str::from_utf8(NS).map_err(xml_error)?
        };
        let mut xml = String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><MapInfo xmlns=\"",
        );
        escape_attr(&mut xml, namespace);
        xml.push_str("\" SelectionNamespaces=\"");
        escape_attr(&mut xml, &self.selection_namespaces);
        xml.push_str("\">");
        for schema in &self.schemas {
            xml.push_str("<Schema ID=\"");
            escape_attr(&mut xml, &schema.id);
            xml.push('"');
            optional_string_attr(&mut xml, "SchemaRef", schema.schema_reference.as_deref());
            optional_string_attr(&mut xml, "Namespace", schema.namespace.as_deref());
            if let Some(payload) = &schema.payload_xml {
                xml.push('>');
                xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?);
                xml.push_str("</Schema>");
            } else {
                xml.push_str("/>");
            }
        }
        for map in &self.maps {
            xml.push_str("<Map ID=\"");
            xml.push_str(&map.id.to_string());
            xml.push_str("\" Name=\"");
            escape_attr(&mut xml, &map.name);
            xml.push_str("\" RootElement=\"");
            escape_attr(&mut xml, &map.root_element);
            xml.push_str("\" SchemaID=\"");
            escape_attr(&mut xml, &map.schema_id);
            xml.push('"');
            bool_attr(
                &mut xml,
                "ShowImportExportValidationErrors",
                map.show_import_export_validation_errors,
            );
            bool_attr(&mut xml, "AutoFit", map.auto_fit);
            bool_attr(&mut xml, "Append", map.append);
            bool_attr(
                &mut xml,
                "PreserveSortAFLayout",
                map.preserve_sort_auto_filter_layout,
            );
            bool_attr(&mut xml, "PreserveFormat", map.preserve_format);
            if let Some(binding) = &map.data_binding {
                xml.push_str("><DataBinding");
                optional_string_attr(
                    &mut xml,
                    "DataBindingName",
                    binding.data_binding_name.as_deref(),
                );
                optional_bool_attr(&mut xml, "FileBinding", binding.file_binding);
                optional_u32_attr(&mut xml, "ConnectionID", binding.connection_id);
                optional_string_attr(
                    &mut xml,
                    "FileBindingName",
                    binding.file_binding_name.as_deref(),
                );
                optional_u32_attr(&mut xml, "DataBindingLoadMode", Some(binding.load_mode));
                if let Some(payload) = &binding.payload_xml {
                    xml.push('>');
                    xml.push_str(std::str::from_utf8(payload).map_err(xml_error)?);
                    xml.push_str("</DataBinding></Map>");
                } else {
                    xml.push_str("/></Map>");
                }
            } else {
                xml.push_str("/>");
            }
        }
        xml.push_str("</MapInfo>");
        if xml.len() > MAX_PART_BYTES {
            return Err(invalid("serialized custom XML maps part exceeds 32 MiB"));
        }
        Ok(xml.into_bytes())
    }
}

/// Discovers and parses the single Custom XML Maps part related to the workbook.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<XmlMapInfo>> {
    Ok(load_from_package_with_conformance(package)?.map(|(value, _)| value))
}

/// Discovers the workbook's Custom XML Maps part together with its namespace family.
///
/// The schema payload and data-binding payloads remain opaque. This function does
/// not resolve schema locations, open bound files, or import/export mapped data.
pub fn load_from_package_with_conformance(
    package: &OpcPackage,
) -> Result<Option<(XmlMapInfo, XmlMapConformance)>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_for_workbook(package, &workbook_uri)
}

/// Store caller-authored Custom XML Maps metadata in a SpreadsheetML package.
///
/// Existing malformed XML Maps relationships are rejected before mutation. The
/// writer never resolves inline schema references, opens bound files, or applies
/// a mapping to worksheet cells.
pub fn store_in_package(
    package: &mut OpcPackage,
    value: &XmlMapInfo,
    conformance: XmlMapConformance,
) -> Result<()> {
    let xml = value.to_xml(conformance.is_strict())?;
    let workbook_uri = main_workbook_uri(package)?;
    let existing = xml_maps_relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_xml_maps_graph(package, &workbook_uri, Some(&existing))?;
        validate_xml_maps_part(package, &existing.part_name)?;
        package.get_part_mut(&existing.part_name)?.set_blob(xml);
        if existing.conformance != conformance {
            let workbook = package.get_part_mut(&workbook_uri)?;
            workbook.rels_mut().remove(&existing.relationship_id);
            workbook.rels_mut().add_relationship(
                conformance.relationship_type().into(),
                existing.target_reference,
                existing.relationship_id,
                false,
            );
        }
    } else {
        validate_xml_maps_graph(package, &workbook_uri, None)?;
        let part_name = next_xml_maps_part_name(package)?;
        let relationship_id = next_xml_maps_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name,
            CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&workbook_uri)?
            .rels_mut()
            .add_relationship(
                conformance.relationship_type().into(),
                target,
                relationship_id,
                false,
            );
    }

    package.unsign();
    Ok(())
}

/// Remove the workbook's Custom XML Maps relationship and its unreferenced part.
///
/// No mapping is applied to worksheet data. A target that remains referenced by
/// another package part is retained.
pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = xml_maps_relationship(package, &workbook_uri)? else {
        validate_xml_maps_graph(package, &workbook_uri, None)?;
        return Ok(false);
    };
    validate_xml_maps_graph(package, &workbook_uri, Some(&existing))?;
    validate_xml_maps_part(package, &existing.part_name)?;

    package
        .get_part_mut(&workbook_uri)?
        .rels_mut()
        .remove(&existing.relationship_id);
    if !package_part_is_referenced(package, &existing.part_name) {
        package.remove_part(&existing.part_name);
    }
    package.unsign();
    Ok(true)
}

fn load_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<(XmlMapInfo, XmlMapConformance)>> {
    let Some(relationship) = xml_maps_relationship(package, workbook_uri)? else {
        validate_xml_maps_graph(package, workbook_uri, None)?;
        return Ok(None);
    };
    validate_xml_maps_graph(package, workbook_uri, Some(&relationship))?;
    validate_xml_maps_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    Ok(Some((
        XmlMapInfo::parse(part.blob())?,
        relationship.conformance,
    )))
}

#[derive(Clone, Debug)]
struct XmlMapsRelationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: XmlMapConformance,
}

fn xml_maps_relationship(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<XmlMapsRelationship>> {
    let workbook = package.get_part(workbook_uri)?;
    let mut relationships = workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL));
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "workbook has multiple custom XML maps relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("custom XML maps relationship cannot be external"));
    }
    let conformance = if relationship.reltype() == REL {
        XmlMapConformance::Transitional
    } else {
        XmlMapConformance::Strict
    };
    Ok(Some(XmlMapsRelationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_xml_maps_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "custom XML maps part '{part_name}' has content type '{}', expected '{CONTENT_TYPE}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("custom XML maps part must not have relationships"));
    }
    Ok(())
}

fn validate_xml_maps_graph(
    package: &OpcPackage,
    workbook_uri: &PackURI,
    expected: Option<&XmlMapsRelationship>,
) -> Result<()> {
    let mut found = 0usize;
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        {
            if part.partname() != workbook_uri {
                return Err(invalid(
                    "custom XML maps relationships may only originate from the workbook",
                ));
            }
            if relationship.is_external() {
                return Err(invalid("custom XML maps relationship cannot be external"));
            }
            let target = relationship.target_partname()?;
            let Some(expected) = expected else {
                return Err(invalid(
                    "workbook has an unexpected custom XML maps relationship",
                ));
            };
            if relationship.r_id() != expected.relationship_id || target != expected.part_name {
                return Err(invalid(
                    "custom XML maps relationship graph is inconsistent",
                ));
            }
            found += 1;
        }
    }
    if package
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
    {
        return Err(invalid(
            "custom XML maps relationships may not originate from the package root",
        ));
    }
    match (expected, found) {
        (None, 0) | (Some(_), 1) => {},
        (None, _) => {
            return Err(invalid(
                "workbook has an unexpected custom XML maps relationship",
            ));
        },
        (Some(_), _) => {
            return Err(invalid(
                "workbook custom XML maps relationship graph is incomplete",
            ));
        },
    }
    Ok(())
}

fn main_workbook_uri(package: &OpcPackage) -> Result<PackURI> {
    use litchi_opc::constants::content_type as ct;

    let workbook = package.main_document_part()?;
    if !matches!(
        workbook.content_type(),
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid(format!(
            "main document part '{}' is not an XML workbook",
            workbook.partname()
        )));
    }
    Ok(workbook.partname().clone())
}

fn next_xml_maps_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/xmlMaps.xml".to_string()
        } else {
            format!("/xl/xmlMaps{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free custom XML maps part name"))
}

fn next_xml_maps_relationship_id(package: &OpcPackage, workbook_uri: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdXmlMaps{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free custom XML maps relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part_name| part_name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|part_name| part_name == *target)
    })
}

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

fn parse_processed(xml: &[u8]) -> Result<XmlMapInfo> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut root_closed = false;
    let mut root_bindings = Vec::new();
    let mut selection = None;
    let mut schemas: Vec<SchemaBuilder> = Vec::new();
    let mut maps: Vec<MapBuilder> = Vec::new();
    let mut capture: Option<Capture> = None;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        if let Some(active) = capture.as_mut() {
            active.events += 1;
            if active.events > MAX_EVENTS {
                return Err(invalid("opaque XML event limit exceeded"));
            }
            match &event {
                Event::Start(_) => {
                    active.depth += 1;
                    if active.depth > MAX_DEPTH {
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
            if active.writer.get_ref().len() > MAX_OPAQUE_BYTES {
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
                if root_closed || !root_name(&namespace, &e, b"MapInfo") {
                    return Err(invalid("expected one SpreadsheetML MapInfo root"));
                }
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
            )?,
            Event::Empty(e) => handle_empty(
                &mut stack,
                &mut schemas,
                &mut maps,
                &root_bindings,
                &namespace,
                e,
                decoder,
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
    validate(&result)?;
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn handle_start(
    stack: &mut Vec<Context>,
    schemas: &mut Vec<SchemaBuilder>,
    maps: &mut Vec<MapBuilder>,
    root_bindings: &[(String, String)],
    ns: &ResolveResult,
    e: BytesStart<'static>,
    decoder: Decoder,
    capture: &mut Option<Capture>,
) -> Result<()> {
    match stack
        .last()
        .copied()
        .ok_or_else(|| invalid("element outside MapInfo"))?
    {
        Context::Root if core_name(ns, &e, b"Schema") => {
            if !maps.is_empty() || schemas.len() >= MAX_SCHEMAS {
                return Err(invalid("invalid Schema order or limit"));
            }
            let value = parse_schema(&e, decoder)?;
            let bindings = merged_bindings(root_bindings, &namespace_attributes(&e, decoder)?);
            schemas.push(SchemaBuilder { value, bindings });
            stack.push(Context::Schema(schemas.len() - 1));
        },
        Context::Root if core_name(ns, &e, b"Map") => {
            if schemas.is_empty() || maps.len() >= MAX_MAPS {
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
    ns: &ResolveResult,
    e: BytesStart<'static>,
    decoder: Decoder,
) -> Result<()> {
    match stack.last().copied() {
        Some(Context::Root) if core_name(ns, &e, b"Schema") => {
            if !maps.is_empty() || schemas.len() >= MAX_SCHEMAS {
                return Err(invalid("invalid Schema order or limit"));
            }
            schemas.push(SchemaBuilder {
                value: parse_schema(&e, decoder)?,
                bindings: merged_bindings(root_bindings, &namespace_attributes(&e, decoder)?),
            });
        },
        Some(Context::Root) if core_name(ns, &e, b"Map") => {
            if schemas.is_empty() || maps.len() >= MAX_MAPS {
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
            schemas[index].value.payload_xml = Some(capture_empty(e, &schemas[index].bindings)?);
        },
        Some(Context::Binding(index)) => {
            let binding = maps
                .get_mut(index)
                .and_then(|map| map.binding.as_mut())
                .ok_or_else(|| invalid("DataBinding context has no binding"))?;
            if binding.value.payload_xml.is_some() {
                return Err(invalid("DataBinding permits at most one opaque child"));
            }
            binding.value.payload_xml = Some(capture_empty(e, &binding.bindings)?);
        },
        _ => return Err(invalid("unexpected empty custom XML maps element")),
    }
    Ok(())
}

fn parse_schema(e: &BytesStart<'_>, d: Decoder) -> Result<XmlMapSchema> {
    let id = required_attr(e, d, b"ID")?;
    let schema_reference = optional_attr(e, d, b"SchemaRef")?;
    let namespace = optional_attr(e, d, b"Namespace")?;
    only_attrs(e, &[b"ID", b"SchemaRef", b"Namespace"])?;
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
fn capture_empty(e: BytesStart<'static>, bindings: &[(String, String)]) -> Result<Vec<u8>> {
    let e = add_inherited_bindings(e, bindings)?;
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Empty(e)).map_err(xml_error)?;
    let bytes = writer.into_inner();
    if bytes.len() > MAX_OPAQUE_BYTES {
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

fn validate(info: &XmlMapInfo) -> Result<()> {
    bounded(&info.selection_namespaces)?;
    if info.schemas.is_empty() || info.schemas.len() > MAX_SCHEMAS {
        return Err(invalid("MapInfo requires 1..4096 Schema children"));
    }
    if info.maps.is_empty() || info.maps.len() > MAX_MAPS {
        return Err(invalid("MapInfo requires 1..65536 Map children"));
    }
    let mut schema_ids = HashSet::new();
    for schema in &info.schemas {
        bounded(&schema.id)?;
        if !schema_ids.insert(schema.id.as_str()) {
            return Err(invalid("duplicate Schema ID"));
        }
        optional_bounded(schema.schema_reference.as_deref())?;
        optional_bounded(schema.namespace.as_deref())?;
        if let Some(payload) = &schema.payload_xml {
            validate_opaque(payload)?;
        }
    }
    let mut map_ids = HashSet::new();
    for map in &info.maps {
        if !map_ids.insert(map.id) {
            return Err(invalid("duplicate Map ID"));
        }
        for value in [&map.name, &map.root_element, &map.schema_id] {
            bounded(value)?;
        }
        if !schema_ids.contains(map.schema_id.as_str()) {
            return Err(invalid("Map references an unknown SchemaID"));
        }
        if let Some(binding) = &map.data_binding {
            optional_bounded(binding.data_binding_name.as_deref())?;
            optional_bounded(binding.file_binding_name.as_deref())?;
            if let Some(payload) = &binding.payload_xml {
                validate_opaque(payload)?;
            }
        }
    }
    Ok(())
}
fn validate_opaque(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_OPAQUE_BYTES {
        return Err(invalid("opaque XML payload exceeds 16 MiB"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut events = 0usize;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(invalid("opaque XML event limit exceeded"));
        }
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                if depth == 0 {
                    roots += 1;
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(invalid("opaque XML depth limit exceeded"));
                }
            },
            Ok(Event::Empty(_)) if depth == 0 => {
                roots += 1;
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid opaque XML nesting"))?;
            },
            Ok(Event::DocType(_) | Event::PI(_) | Event::Decl(_)) => {
                return Err(invalid(
                    "DTD, declarations, and processing instructions are rejected in opaque XML",
                ));
            },
            Ok(Event::Text(t))
                if depth == 0 && !t.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("text outside opaque XML root"));
            },
            Ok(Event::CData(_) | Event::GeneralRef(_)) if depth == 0 => {
                return Err(invalid("data outside opaque XML root"));
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
            _ => {},
        }
    }
    if roots != 1 || depth != 0 {
        return Err(invalid(
            "opaque XML must contain exactly one complete element",
        ));
    }
    Ok(())
}

fn root_name(ns: &ResolveResult, e: &BytesStart<'_>, local: &[u8]) -> bool {
    namespace_matches(ns) && e.local_name().as_ref() == local
}
fn core_name(ns: &ResolveResult, e: &BytesStart<'_>, local: &[u8]) -> bool {
    (namespace_matches(ns) || matches!(ns, ResolveResult::Unbound))
        && e.local_name().as_ref() == local
}
fn namespace_matches(ns: &ResolveResult) -> bool {
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
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
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
        .filter_map(|a| a.ok())
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
fn bounded(value: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        Err(invalid("custom XML maps string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn optional_bounded(value: Option<&str>) -> Result<()> {
    if let Some(v) = value {
        bounded(v)
    } else {
        Ok(())
    }
}
fn optional_string_attr(xml: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        xml.push(' ');
        xml.push_str(name);
        xml.push_str("=\"");
        escape_attr(xml, value);
        xml.push('"');
    }
}
fn optional_bool_attr(xml: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        bool_attr(xml, name, value);
    }
}
fn optional_u32_attr(xml: &mut String, name: &str, value: Option<u32>) {
    if let Some(value) = value {
        xml.push(' ');
        xml.push_str(name);
        xml.push_str("=\"");
        xml.push_str(&value.to_string());
        xml.push('"');
    }
}
fn bool_attr(xml: &mut String, name: &str, value: bool) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str(if value { "=\"1\"" } else { "=\"0\"" });
}
fn escape_attr(out: &mut String, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            '\r' => out.push_str("&#xD;"),
            '\n' => out.push_str("&#xA;"),
            '\t' => out.push_str("&#x9;"),
            _ => out.push(c),
        }
    }
}
fn invalid(message: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, message.into()).into()
}
fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::part::{BlobPart, Part};

    fn package(bytes: &[u8]) -> XmlMapInfo {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        load_from_package(&package).unwrap().unwrap()
    }

    fn fixture_info() -> XmlMapInfo {
        XmlMapInfo {
            selection_namespaces: "xmlns:xs='http://www.w3.org/2001/XMLSchema'".into(),
            schemas: vec![XmlMapSchema {
                id: "schema-1".into(),
                schema_reference: Some("urn:litchi:example".into()),
                namespace: Some("urn:litchi:example".into()),
                payload_xml: Some(
                    br#"<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec(),
                ),
            }],
            maps: vec![XmlMap {
                id: 1,
                name: "Example map".into(),
                root_element: "example".into(),
                schema_id: "schema-1".into(),
                show_import_export_validation_errors: true,
                auto_fit: true,
                append: false,
                preserve_sort_auto_filter_layout: true,
                preserve_format: true,
                data_binding: Some(XmlMapDataBinding {
                    data_binding_name: Some("inert binding".into()),
                    file_binding: Some(false),
                    connection_id: Some(7),
                    file_binding_name: None,
                    load_mode: 1,
                    payload_xml: Some(br#"<binding xmlns="urn:litchi:binding"/>"#.to_vec()),
                }),
            }],
        }
    }

    fn workbook_package() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let workbook = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.into(),
            format!(
                r#"<workbook xmlns="{}"><sheets/></workbook>"#,
                std::str::from_utf8(NS).unwrap()
            )
            .into_bytes(),
        );
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook));
        package
    }

    fn synthetic_package(
        relationship_type: &str,
        external: bool,
        content_type: &str,
        outbound: bool,
    ) -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let mut workbook = BlobPart::new(
            workbook_uri.clone(),
            ct::SML_SHEET_MAIN.into(),
            format!(
                r#"<workbook xmlns="{}"><sheets/></workbook>"#,
                std::str::from_utf8(NS).unwrap()
            )
            .into_bytes(),
        );
        if external {
            workbook.relate_to_ext("https://example.invalid/xmlMaps.xml", relationship_type);
        } else {
            workbook.relate_to("xmlMaps.xml", relationship_type);
        }
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook));
        if !external {
            let mut maps = BlobPart::new(
                PackURI::new("/xl/xmlMaps.xml").unwrap(),
                content_type.into(),
                fixture_info().to_xml(false).unwrap(),
            );
            if outbound {
                maps.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
            }
            package.add_part(Box::new(maps));
        }
        package
    }
    #[test]
    fn reads_poi_real_fixture_and_round_trips_strict() {
        let maps = package(include_bytes!(
            "../../../test-data/poi/test-data/spreadsheet/CustomXMLMappings.xlsx"
        ));
        assert_eq!(maps.schemas.len(), 1);
        assert_eq!(maps.maps[0].name, "CORSO_mapping");
        let strict = maps.to_xml(true).unwrap();
        assert_eq!(XmlMapInfo::parse(&strict).unwrap(), maps);
    }
    #[test]
    fn keeps_poi_xxe_schema_inert() {
        let maps = package(include_bytes!(
            "../../../test-data/poi/test-data/spreadsheet/xxe_in_schema.xlsx"
        ));
        let payload = maps.schemas[0].payload_xml.as_deref().unwrap();
        let text = std::str::from_utf8(payload).unwrap();
        assert!(text.contains("schemaLocation=\"http://localhost\""));
        assert!(text.contains("redefine"));
    }
    #[test]
    fn reads_libreoffice_unqualified_children_and_binding() {
        let maps = package(include_bytes!(
            "../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf167689_xmlMaps_and_xmlColumnPr.xlsx"
        ));
        let binding = maps.maps[0].data_binding.as_ref().unwrap();
        assert_eq!(binding.file_binding, Some(true));
        assert_eq!(binding.connection_id, Some(1));
        assert_eq!(binding.load_mode, 1);
    }
    #[test]
    fn handles_strict_and_mce_fallback() {
        let strict = std::str::from_utf8(STRICT_NS).unwrap();
        let xml = format!(
            r#"<MapInfo xmlns="{strict}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u" SelectionNamespaces=""><Schema ID="s"><x:schema xmlns:x="urn:x"/></Schema><mc:AlternateContent><mc:Choice Requires="u"><u:Map/></mc:Choice><mc:Fallback><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="0" AutoFit="1" Append="0" PreserveSortAFLayout="1" PreserveFormat="1"/></mc:Fallback></mc:AlternateContent></MapInfo>"#
        );
        let parsed = XmlMapInfo::parse(xml.as_bytes()).unwrap();
        assert_eq!(parsed.maps.len(), 1);
        assert_eq!(
            XmlMapInfo::parse(&parsed.to_xml(false).unwrap()).unwrap(),
            parsed
        );
    }
    #[test]
    fn rejects_malformed_unsafe_and_invalid_models() {
        let ns = std::str::from_utf8(NS).unwrap();
        for xml in [
            format!(
                r#"<MapInfo xmlns="{ns}" SelectionNamespaces=""><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="0" AutoFit="1" Append="0" PreserveSortAFLayout="1" PreserveFormat="1"/></MapInfo>"#
            ),
            format!(
                r#"<MapInfo xmlns="{ns}" SelectionNamespaces=""><Schema ID="s"/><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="maybe" AutoFit="1" Append="0" PreserveSortAFLayout="1" PreserveFormat="1"/></MapInfo>"#
            ),
            format!(
                r#"<!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><MapInfo xmlns="{ns}" SelectionNamespaces=""><Schema ID="s"><x:schema xmlns:x="urn:x">&e;</x:schema></Schema><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="0" AutoFit="1" Append="0" PreserveSortAFLayout="1" PreserveFormat="1"/></MapInfo>"#
            ),
        ] {
            assert!(XmlMapInfo::parse(xml.as_bytes()).is_err(), "accepted {xml}");
        }
        let mut valid = package(include_bytes!(
            "../../../test-data/poi/test-data/spreadsheet/CustomXMLMappings.xlsx"
        ));
        valid.schemas[0].payload_xml = Some(b"<?unsafe?><x/>".to_vec());
        assert!(valid.to_xml(false).is_err());
    }

    #[test]
    fn stores_rewrites_and_removes_inert_xml_maps_parts() {
        let mut package = workbook_package();
        let value = fixture_info();

        store_in_package(&mut package, &value, XmlMapConformance::Transitional).unwrap();
        assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));
        assert_eq!(
            load_from_package_with_conformance(&package).unwrap(),
            Some((value.clone(), XmlMapConformance::Transitional))
        );

        let workbook = package.main_document_part().unwrap();
        let relationship = workbook
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == REL)
            .unwrap();
        let relationship_id = relationship.r_id().to_string();
        let part_name = relationship.target_partname().unwrap();
        assert_eq!(part_name, PackURI::new("/xl/xmlMaps.xml").unwrap());
        assert!(
            std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
                .unwrap()
                .contains(std::str::from_utf8(NS).unwrap())
        );

        let mut replacement = value.clone();
        replacement.maps[0].name = "Strict replacement".into();
        store_in_package(&mut package, &replacement, XmlMapConformance::Strict).unwrap();
        let workbook = package.main_document_part().unwrap();
        let relationship = workbook
            .rels()
            .iter()
            .find(|relationship| relationship.r_id() == relationship_id)
            .unwrap();
        assert_eq!(relationship.reltype(), STRICT_REL);
        assert_eq!(relationship.target_partname().unwrap(), part_name);
        assert!(
            std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
                .unwrap()
                .contains(std::str::from_utf8(STRICT_NS).unwrap())
        );
        assert_eq!(
            load_from_package_with_conformance(&package).unwrap(),
            Some((replacement, XmlMapConformance::Strict))
        );

        assert!(remove_from_package(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_err());
        assert_eq!(load_from_package(&package).unwrap(), None);
        assert!(!remove_from_package(&mut package).unwrap());
    }

    #[test]
    fn preserves_unrelated_references_when_removing_xml_maps() {
        let mut package = workbook_package();
        let value = fixture_info();
        store_in_package(&mut package, &value, XmlMapConformance::Transitional).unwrap();

        let part_name = PackURI::new("/xl/xmlMaps.xml").unwrap();
        let mut referring_part = BlobPart::new(
            PackURI::new("/xl/retained-reference.xml").unwrap(),
            ct::XML.into(),
            b"<reference/>".to_vec(),
        );
        referring_part.relate_to("xmlMaps.xml", "urn:litchi:test:xml-maps-reference");
        package.add_part(Box::new(referring_part));

        assert!(remove_from_package(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_ok());
        assert_eq!(load_from_package(&package).unwrap(), None);

        store_in_package(&mut package, &value, XmlMapConformance::Transitional).unwrap();
        let relationship = package
            .main_document_part()
            .unwrap()
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == REL)
            .unwrap();
        assert_eq!(
            relationship.target_partname().unwrap(),
            PackURI::new("/xl/xmlMaps1.xml").unwrap()
        );
    }

    #[test]
    fn writes_real_poi_xml_maps_package_without_resolving_schema_payloads() {
        let mut package = OpcPackage::from_bytes(include_bytes!(
            "../../../test-data/poi/test-data/spreadsheet/CustomXMLMappings.xlsx"
        ))
        .unwrap();
        let (value, conformance) = load_from_package_with_conformance(&package)
            .unwrap()
            .unwrap();
        store_in_package(&mut package, &value, conformance).unwrap();

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("xml-maps.xlsx");
        package.save(&path).unwrap();
        let reopened = OpcPackage::open(&path).unwrap();
        assert_eq!(
            load_from_package_with_conformance(&reopened).unwrap(),
            Some((value, conformance))
        );
    }

    #[test]
    fn package_xml_maps_mutators_reject_invalid_existing_graphs_before_replacement() {
        let value = fixture_info();
        let mut wrong_content_type = synthetic_package(REL, false, ct::SML_STYLES, false);
        let part_name = PackURI::new("/xl/xmlMaps.xml").unwrap();
        let original = wrong_content_type
            .get_part(&part_name)
            .unwrap()
            .blob()
            .to_vec();
        assert!(
            store_in_package(
                &mut wrong_content_type,
                &value,
                XmlMapConformance::Transitional,
            )
            .is_err()
        );
        assert_eq!(
            wrong_content_type.get_part(&part_name).unwrap().blob(),
            original
        );
        assert!(remove_from_package(&mut wrong_content_type).is_err());

        let mut duplicate = synthetic_package(REL, false, CONTENT_TYPE, false);
        duplicate
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                REL.into(),
                "xmlMaps.xml".into(),
                "rIdDuplicateXmlMaps".into(),
                false,
            );
        assert!(store_in_package(&mut duplicate, &value, XmlMapConformance::Transitional).is_err());
        assert!(remove_from_package(&mut duplicate).is_err());

        let mut external = synthetic_package(REL, true, CONTENT_TYPE, false);
        assert!(store_in_package(&mut external, &value, XmlMapConformance::Transitional).is_err());
        assert!(remove_from_package(&mut external).is_err());

        let mut outbound = synthetic_package(REL, false, CONTENT_TYPE, true);
        assert!(store_in_package(&mut outbound, &value, XmlMapConformance::Transitional).is_err());
        assert!(remove_from_package(&mut outbound).is_err());

        let mut root_relationship = workbook_package();
        root_relationship.relate_to("xl/xmlMaps.xml", REL);
        assert!(
            store_in_package(
                &mut root_relationship,
                &value,
                XmlMapConformance::Transitional,
            )
            .is_err()
        );
    }
}
