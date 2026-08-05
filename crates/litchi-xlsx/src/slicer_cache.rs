//! Bounded SpreadsheetML slicer-cache definition codec.
///
/// This module owns the inert XML grammar rooted at `x14:slicerCacheDefinition`
/// from the checked-in [MS-XLSX] anchors §§2.1.4, 2.2.4.8, 2.3.2.1, 2.4.38,
/// 2.4.60, 2.6.70--2.6.85, 2.6.97, and 2.6.103--2.6.104. It deliberately
/// does not open or resolve package relationships; the OOXML host retains
/// workbook extension and OPC graph operations.
///
/// Cache data and extension payloads are validated and retained inertly with
/// bounded sizes. Serialization is deterministic and never interprets
/// relationship-looking content inside retained subtrees.
use crate::error::{Error, Result};
use litchi_ooxml_common::custom_xml::valid_guid;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashSet, TryReserveError};

#[path = "slicer_cache/crud.rs"]
pub mod crud;
#[path = "slicer_cache/package.rs"]
pub mod package;
#[path = "slicer_cache/views.rs"]
pub mod views;

pub use crud::*;
pub use package::{load_slicer_caches, store_slicer_cache};
pub use views::{
    Slicer, SlicerExtensionList, SlicerPart, Slicers, load_slicer_parts, parse_slicers,
    store_slicer_part, write_slicers,
};

pub const SLICER_CACHE_CONTENT_TYPE: &str = "application/vnd.ms-excel.slicerCache+xml";
pub const SLICER_CACHE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicerCache";

const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XR10: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_PIVOT_TABLES: usize = 65_536;
const MAX_DEPTH: usize = 256;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_INERT_SUBTREE_BYTES: usize = 8 * 1024 * 1024;
const MAX_TOTAL_INERT_BYTES: usize = 16 * 1024 * 1024;
const MAX_EXTENSION_ATTRIBUTES: usize = 128;
const MAX_EXTENSION_ATTRIBUTE_BYTES: usize = 64 * 1024;
const MAX_EVENTS: usize = 1_000_000;
const MAX_NODES: usize = 500_000;

#[derive(Debug, Clone, PartialEq, Eq)]
struct XmlAttribute {
    qualified_name: String,
    value: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataKind {
    Olap,
    Tabular,
}

/// The complete `data` subtree, retained inertly after validating its exact
/// one-of OLAP/tabular discriminator.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Data {
    kind: DataKind,
    xml: Vec<u8>,
}

impl Data {
    pub fn new(xml: Vec<u8>) -> Result<Self> {
        let wrapped = wrap_fragment(&xml)?;
        let definition = parse(&wrapped)?;
        definition
            .data
            .ok_or_else(|| invalid("expected one Slicer Cache data fragment"))
    }
    pub fn kind(&self) -> DataKind {
        self.kind
    }
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }
}

/// The complete optional `extLst` subtree, retained inertly and never used to
/// resolve relationships.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList(Vec<u8>);

impl ExtensionList {
    pub fn new(xml: Vec<u8>) -> Result<Self> {
        let wrapped = wrap_fragment(&xml)?;
        let definition = parse(&wrapped)?;
        definition
            .extension_list
            .ok_or_else(|| invalid("expected one Slicer Cache extLst fragment"))
    }
    pub fn xml(&self) -> &[u8] {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTable {
    pub tab_id: u32,
    pub name: String,
    xml_attributes: Vec<XmlAttribute>,
}

impl PivotTable {
    pub fn new(tab_id: u32, name: impl Into<String>) -> Self {
        Self {
            tab_id,
            name: name.into(),
            xml_attributes: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Definition {
    pub name: String,
    pub source_name: String,
    pub uid: Option<String>,
    pub pivot_tables: Vec<PivotTable>,
    pub data: Option<Data>,
    pub extension_list: Option<ExtensionList>,
    xml_attributes: Vec<XmlAttribute>,
}

impl Definition {
    pub fn new(name: impl Into<String>, source_name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            source_name: source_name.into(),
            uid: None,
            pivot_tables: Vec::new(),
            data: None,
            extension_list: None,
            xml_attributes: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cache {
    pub relationship_id: String,
    pub part_name: String,
    pub definition: Definition,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CaptureKind {
    Data,
    Extension,
}

#[derive(Debug, Clone, Copy)]
struct Capture {
    kind: CaptureKind,
    start: usize,
    parent_depth: usize,
    data_kind: Option<DataKind>,
}

/// A byte sink that makes every growth and output-size check explicit.
struct BoundedXml {
    bytes: Vec<u8>,
    limit: usize,
    resource: &'static str,
}

impl BoundedXml {
    fn new(limit: usize, resource: &'static str) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
            resource,
        }
    }

    fn append(&mut self, bytes: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| limit("serialized XML bytes"))?;
        if length > self.limit {
            return Err(limit(self.resource));
        }
        self.bytes
            .try_reserve(bytes.len())
            .map_err(|source| allocation(self.resource, source))?;
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }

    fn push(&mut self, byte: u8) -> Result<()> {
        self.append(&[byte])
    }

    fn escape(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            match character {
                '&' => self.append(b"&amp;")?,
                '<' => self.append(b"&lt;")?,
                '"' => self.append(b"&quot;")?,
                '\t' => self.append(b"&#x9;")?,
                '\n' => self.append(b"&#xA;")?,
                '\r' => self.append(b"&#xD;")?,
                _ => {
                    let mut bytes = [0; 4];
                    self.append(character.encode_utf8(&mut bytes).as_bytes())?;
                },
            }
        }
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

pub fn parse(xml: &[u8]) -> Result<Definition> {
    if xml.len() > MAX_PART_BYTES {
        return Err(limit("part bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let decoder = reader.decoder();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut definition: Option<Definition> = None;
    let mut stage = 0u8;
    let mut in_pivots = false;
    let mut open_pivot: Option<(usize, PivotTable)> = None;
    let mut capture: Option<Capture> = None;
    let mut total_inert = 0usize;
    let mut events = 0usize;
    let mut nodes = 0usize;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| limit("XML event count"))?;
        if events > MAX_EVENTS {
            return Err(limit("XML event count"));
        }
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("Slicer Cache XML offset overflow"))?;
        let borrowed = reader.read_event().map_err(xml_error)?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("Slicer Cache XML offset overflow"))?;
        let event_bytes = end
            .checked_sub(start)
            .ok_or_else(|| invalid("Slicer Cache XML offsets moved backwards"))?;
        if event_bytes > MAX_INERT_SUBTREE_BYTES {
            return Err(limit("XML event bytes"));
        }
        let event = borrowed.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| limit("XML node count"))?;
            if nodes > MAX_NODES {
                return Err(limit("XML node count"));
            }
        }
        validate_event_xml(&event, &resolver, decoder, capture.is_some())?;
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("content after Slicer Cache root"));
                }
                if let Some(active) = capture.as_mut() {
                    if active.kind == CaptureKind::Data && depth == active.parent_depth + 1 {
                        let kind = data_kind(&namespace, element.local_name().as_ref())?;
                        if active.data_kind.replace(kind).is_some() {
                            return Err(invalid("Slicer Cache data requires exactly one source"));
                        }
                    }
                } else if let Some((pivot_depth, _)) = &open_pivot {
                    if depth > *pivot_depth {
                        return Err(invalid("pivotTable must be empty"));
                    }
                } else if depth == 0 {
                    if root_seen
                        || !exact(&namespace, X14)
                        || element.local_name().as_ref() != b"slicerCacheDefinition"
                    {
                        return Err(invalid("expected x14:slicerCacheDefinition root"));
                    }
                    root_seen = true;
                    definition = Some(parse_root_attributes(&element, &resolver, decoder)?);
                } else if depth == 1 {
                    match element.local_name().as_ref() {
                        b"pivotTables" if exact(&namespace, X14) && stage == 0 => {
                            reject_attributes(&element, &resolver, decoder, &[])?;
                            in_pivots = true;
                            stage = 1;
                        },
                        b"data" if exact(&namespace, X14) && stage <= 1 => {
                            reject_attributes(&element, &resolver, decoder, &[])?;
                            capture = Some(Capture {
                                kind: CaptureKind::Data,
                                start,
                                parent_depth: depth,
                                data_kind: None,
                            });
                            stage = 2;
                        },
                        b"extLst" if exact(&namespace, X14) && stage <= 2 => {
                            capture = Some(Capture {
                                kind: CaptureKind::Extension,
                                start,
                                parent_depth: depth,
                                data_kind: None,
                            });
                            stage = 3;
                        },
                        _ => return Err(invalid("unexpected or out-of-order Slicer Cache child")),
                    }
                } else if depth == 2 && in_pivots {
                    if !exact(&namespace, X14) || element.local_name().as_ref() != b"pivotTable" {
                        return Err(invalid("pivotTables may contain only pivotTable"));
                    }
                    let pivot = parse_pivot_table(&element, &resolver, decoder)?;
                    open_pivot = Some((depth, pivot));
                } else {
                    return Err(invalid("unexpected Slicer Cache element depth"));
                }
                depth = depth.checked_add(1).ok_or_else(|| limit("XML depth"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("content after Slicer Cache root"));
                }
                if let Some(active) = capture.as_mut() {
                    if active.kind == CaptureKind::Data && depth == active.parent_depth + 1 {
                        let kind = data_kind(&namespace, element.local_name().as_ref())?;
                        if active.data_kind.replace(kind).is_some() {
                            return Err(invalid("Slicer Cache data requires exactly one source"));
                        }
                    }
                } else if depth == 0 {
                    return Err(invalid("Slicer Cache root cannot be empty"));
                } else if depth == 1 {
                    match element.local_name().as_ref() {
                        b"pivotTables" if exact(&namespace, X14) && stage == 0 => {
                            return Err(invalid("pivotTables cannot be empty"));
                        },
                        b"data" if exact(&namespace, X14) && stage <= 1 => {
                            return Err(invalid("data cannot be empty"));
                        },
                        b"extLst" if exact(&namespace, X14) && stage <= 2 => {
                            let bytes = retained(xml, start, end, &mut total_inert)?;
                            let definition = definition
                                .as_mut()
                                .ok_or_else(|| invalid("missing Slicer Cache definition"))?;
                            definition.extension_list = Some(ExtensionList(bytes));
                            stage = 3;
                        },
                        _ => return Err(invalid("unexpected or out-of-order Slicer Cache child")),
                    }
                } else if depth == 2 && in_pivots {
                    if !exact(&namespace, X14) || element.local_name().as_ref() != b"pivotTable" {
                        return Err(invalid("pivotTables may contain only pivotTable"));
                    }
                    let definition = definition
                        .as_mut()
                        .ok_or_else(|| invalid("missing Slicer Cache definition"))?;
                    push_pivot(definition, parse_pivot_table(&element, &resolver, decoder)?)?;
                } else {
                    return Err(invalid("unexpected empty Slicer Cache element"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected Slicer Cache closing element"));
                }
                depth -= 1;
                if let Some(active) = capture {
                    if depth == active.parent_depth {
                        let bytes = retained(xml, active.start, end, &mut total_inert)?;
                        let target = definition
                            .as_mut()
                            .ok_or_else(|| invalid("missing Slicer Cache definition"))?;
                        match active.kind {
                            CaptureKind::Data => {
                                let kind = active.data_kind.ok_or_else(|| {
                                    invalid("Slicer Cache data requires an olap or tabular child")
                                })?;
                                target.data = Some(Data { kind, xml: bytes });
                            },
                            CaptureKind::Extension => {
                                target.extension_list = Some(ExtensionList(bytes))
                            },
                        }
                        capture = None;
                    }
                    continue;
                }
                if let Some((pivot_depth, _)) = &open_pivot {
                    if depth == *pivot_depth {
                        if element.local_name().as_ref() != b"pivotTable" {
                            return Err(invalid("mismatched pivotTable close"));
                        }
                        let (_, pivot) = open_pivot
                            .take()
                            .ok_or_else(|| invalid("missing Slicer Cache pivotTable"))?;
                        let definition = definition
                            .as_mut()
                            .ok_or_else(|| invalid("missing Slicer Cache definition"))?;
                        push_pivot(definition, pivot)?;
                    }
                    continue;
                }
                if depth == 1 && in_pivots {
                    if element.local_name().as_ref() != b"pivotTables" {
                        return Err(invalid("mismatched pivotTables close"));
                    }
                    if definition
                        .as_ref()
                        .ok_or_else(|| invalid("missing Slicer Cache definition"))?
                        .pivot_tables
                        .is_empty()
                    {
                        return Err(invalid("pivotTables requires at least one pivotTable"));
                    }
                    in_pivots = false;
                } else if depth == 0 {
                    if element.local_name().as_ref() != b"slicerCacheDefinition" || !root_seen {
                        return Err(invalid("mismatched Slicer Cache root"));
                    }
                    root_closed = true;
                } else {
                    return Err(invalid("unexpected Slicer Cache closing depth"));
                }
            },
            Event::Text(text) if capture.is_none() => {
                let value = text.decode().map_err(xml_error)?;
                if !value.trim().is_empty() {
                    return Err(invalid("unexpected text in Slicer Cache part"));
                }
            },
            Event::CData(_) if capture.is_none() => {
                return Err(invalid("unexpected CDATA in Slicer Cache part"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen
        || !root_closed
        || depth != 0
        || capture.is_some()
        || in_pivots
        || open_pivot.is_some()
    {
        return Err(invalid("incomplete Slicer Cache XML"));
    }
    let value = definition.ok_or_else(|| invalid("missing Slicer Cache definition"))?;
    validate(&value)?;
    Ok(value)
}

pub fn write(value: &Definition) -> Result<Vec<u8>> {
    validate(value)?;
    let mut output = BoundedXml::new(MAX_PART_BYTES, "serialized part bytes");
    output.append(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")?;
    output.append(b"<x14:slicerCacheDefinition xmlns:x14=\"")?;
    output.escape(X14)?;
    output.push(b'\"')?;
    write_xml_attributes(&mut output, &value.xml_attributes)?;
    output.append(b" name=\"")?;
    output.escape(&value.name)?;
    output.append(b"\" sourceName=\"")?;
    output.escape(&value.source_name)?;
    output.push(b'\"')?;
    if let Some(uid) = &value.uid {
        output.append(b" xmlns:xr10=\"")?;
        output.escape(XR10)?;
        output.append(b"\" xr10:uid=\"")?;
        output.escape(uid)?;
        output.push(b'\"')?;
    }
    output.push(b'>')?;
    if !value.pivot_tables.is_empty() {
        output.append(b"<x14:pivotTables>")?;
        for pivot in &value.pivot_tables {
            output.append(b"<x14:pivotTable tabId=\"")?;
            let tab_id = pivot.tab_id.to_string();
            output.append(tab_id.as_bytes())?;
            output.append(b"\" name=\"")?;
            output.escape(&pivot.name)?;
            output.push(b'\"')?;
            write_xml_attributes(&mut output, &pivot.xml_attributes)?;
            output.append(b"/>")?;
        }
        output.append(b"</x14:pivotTables>")?;
    }
    if let Some(data) = &value.data {
        output.append(data.xml())?;
    }
    if let Some(extension) = &value.extension_list {
        output.append(extension.xml())?;
    }
    output.append(b"</x14:slicerCacheDefinition>")?;
    Ok(output.finish())
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<Definition> {
    let mut name = None;
    let mut source_name = None;
    let mut uid = None;
    let mut xml_attributes = Vec::new();
    let mut retained_bytes = 0usize;
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = std::str::from_utf8(item.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if raw == "xmlns:x14" && value == X14 {
            continue;
        }
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            validate_reserved_namespace(&raw, &value)?;
            retain_attribute(&mut xml_attributes, &mut retained_bytes, raw, value)?;
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(item.key);
        let namespace = namespace_string(&namespace)?;
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        match (namespace.as_str(), local) {
            ("", "name") => set_once(&mut name, value, "name")?,
            ("", "sourceName") => set_once(&mut source_name, value, "sourceName")?,
            (XR10, "uid") => set_once(&mut uid, value, "uid")?,
            ("", _) => {
                return Err(invalid(format!(
                    "unexpected Slicer Cache attribute '{local}'"
                )));
            },
            _ => retain_attribute(&mut xml_attributes, &mut retained_bytes, raw, value)?,
        }
    }
    Ok(Definition {
        name: name.ok_or_else(|| invalid("Slicer Cache requires name"))?,
        source_name: source_name.ok_or_else(|| invalid("Slicer Cache requires sourceName"))?,
        uid,
        pivot_tables: Vec::new(),
        data: None,
        extension_list: None,
        xml_attributes,
    })
}

fn parse_pivot_table(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<PivotTable> {
    let mut tab_id = None;
    let mut name = None;
    let mut xml_attributes = Vec::new();
    let mut retained_bytes = 0usize;
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = std::str::from_utf8(item.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            validate_reserved_namespace(&raw, &value)?;
            retain_attribute(&mut xml_attributes, &mut retained_bytes, raw, value)?;
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(item.key);
        let namespace = namespace_string(&namespace)?;
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        match (namespace.as_str(), local) {
            ("", "tabId") => set_once(&mut tab_id, parse_u32(&value, "tabId")?, "tabId")?,
            ("", "name") => set_once(&mut name, value, "name")?,
            ("", _) => {
                return Err(invalid(format!(
                    "unexpected pivotTable attribute '{local}'"
                )));
            },
            _ => retain_attribute(&mut xml_attributes, &mut retained_bytes, raw, value)?,
        }
    }
    let value = PivotTable {
        tab_id: tab_id.ok_or_else(|| invalid("pivotTable requires tabId"))?,
        name: name.ok_or_else(|| invalid("pivotTable requires name"))?,
        xml_attributes,
    };
    validate_pivot(&value)?;
    Ok(value)
}

fn reject_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    allowed: &[(&str, &str)],
) -> Result<()> {
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            let value = item
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?;
            validate_reserved_namespace(std::str::from_utf8(raw).map_err(xml_error)?, &value)?;
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(item.key);
        let namespace = namespace_string(&namespace)?;
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        let _ = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        if !allowed.contains(&(namespace.as_str(), local)) {
            return Err(invalid(format!("unexpected attribute '{local}'")));
        }
    }
    Ok(())
}

pub fn validate(value: &Definition) -> Result<()> {
    validate_name(&value.name)?;
    validate_required_string(&value.source_name, "Slicer Cache sourceName")?;
    if let Some(uid) = &value.uid {
        validate_guid(uid)?;
    }
    validate_xml_attributes(&value.xml_attributes)?;
    if value.pivot_tables.len() > MAX_PIVOT_TABLES {
        return Err(limit("pivot table count"));
    }
    let mut pivots = HashSet::new();
    for pivot in &value.pivot_tables {
        validate_pivot(pivot)?;
        if !pivots.insert((pivot.tab_id, pivot.name.to_lowercase())) {
            return Err(invalid("duplicate Slicer Cache pivotTable binding"));
        }
    }
    let mut total = 0usize;
    if let Some(data) = &value.data {
        if data.xml.len() > MAX_INERT_SUBTREE_BYTES {
            return Err(limit("data subtree bytes"));
        }
        total = total
            .checked_add(data.xml.len())
            .ok_or_else(|| limit("inert subtree bytes"))?;
    }
    if let Some(extension) = &value.extension_list {
        if extension.0.len() > MAX_INERT_SUBTREE_BYTES {
            return Err(limit("extension subtree bytes"));
        }
        total = total
            .checked_add(extension.0.len())
            .ok_or_else(|| limit("inert subtree bytes"))?;
    }
    if total > MAX_TOTAL_INERT_BYTES {
        return Err(limit("inert subtree bytes"));
    }
    Ok(())
}

fn validate_pivot(value: &PivotTable) -> Result<()> {
    validate_required_string(&value.name, "pivotTable name")?;
    validate_xml_attributes(&value.xml_attributes)
}

fn validate_name(value: &str) -> Result<()> {
    validate_required_string(value, "Slicer Cache name")?;
    if value.chars().count() > 32_767 {
        return Err(invalid("Slicer Cache name exceeds 32767 characters"));
    }
    let mut chars = value.chars();
    let first = chars
        .next()
        .ok_or_else(|| invalid("Slicer Cache name cannot be empty"))?;
    if !(first == '_' || first == '\\' || first.is_alphabetic())
        || !chars.all(|character| {
            character == '_'
                || character == '\\'
                || character == '.'
                || character == '?'
                || character.is_alphanumeric()
        })
    {
        return Err(invalid(format!(
            "invalid Slicer Cache defined name '{value}'"
        )));
    }
    Ok(())
}

fn validate_guid(value: &str) -> Result<()> {
    if !valid_guid(value) {
        Err(invalid(format!("invalid Slicer Cache uid '{value}'")))
    } else {
        Ok(())
    }
}

fn validate_required_string(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{name} cannot be empty")));
    }
    if value.len() > MAX_TEXT_BYTES {
        return Err(limit(name));
    }
    if value.chars().any(|character| !is_xml_character(character)) {
        return Err(invalid(format!("{name} contains an invalid XML character")));
    }
    Ok(())
}

fn validate_xml_value(value: &str, max_bytes: usize, name: &str) -> Result<()> {
    if value.len() > max_bytes {
        return Err(limit(name));
    }
    if value.chars().any(|character| !is_xml_character(character)) {
        return Err(invalid(format!("{name} contains an invalid XML character")));
    }
    Ok(())
}

fn validate_attribute_name(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    let prefix = parts.next().unwrap_or_default();
    let local = parts.next();
    if parts.next().is_some()
        || prefix.is_empty()
        || !valid_xml_name_part(prefix, true)
        || local.is_some_and(|part| !valid_xml_name_part(part, false))
    {
        return Err(invalid(format!("invalid XML attribute name '{value}'")));
    }
    Ok(())
}

fn validate_reserved_namespace(name: &str, value: &str) -> Result<()> {
    let expected = match name {
        "xmlns:x14" => Some(X14),
        "xmlns:xr10" => Some(XR10),
        _ => None,
    };
    if let Some(expected) = expected
        && value != expected
    {
        return Err(invalid(format!(
            "reserved XML namespace '{name}' is incorrect"
        )));
    }
    Ok(())
}

fn valid_xml_name_part(value: &str, allow_colon: bool) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first == '_' || first.is_alphabetic() || (allow_colon && first == ':')) {
        return false;
    }
    chars.all(|character| {
        character == '_'
            || character == '-'
            || character == '.'
            || character.is_alphanumeric()
            || (allow_colon && character == ':')
    })
}

fn validate_xml_attributes(values: &[XmlAttribute]) -> Result<()> {
    if values.len() > MAX_EXTENSION_ATTRIBUTES {
        return Err(limit("extension attributes"));
    }
    let mut total = 0usize;
    let mut names = HashSet::new();
    names
        .try_reserve(values.len())
        .map_err(|source| allocation("Slicer Cache attribute names", source))?;
    for value in values {
        validate_attribute_name(&value.qualified_name)?;
        validate_xml_value(&value.value, MAX_TEXT_BYTES, "retained XML attribute")?;
        if !names.insert(value.qualified_name.as_str()) {
            return Err(invalid("duplicate retained XML attribute"));
        }
        total = total
            .checked_add(
                value
                    .qualified_name
                    .len()
                    .checked_add(value.value.len())
                    .ok_or_else(|| limit("extension attribute bytes"))?,
            )
            .ok_or_else(|| limit("extension attribute bytes"))?;
        if total > MAX_EXTENSION_ATTRIBUTE_BYTES {
            return Err(limit("extension attribute bytes"));
        }
    }
    Ok(())
}

fn write_xml_attributes(output: &mut BoundedXml, values: &[XmlAttribute]) -> Result<()> {
    validate_xml_attributes(values)?;
    for value in values {
        if value.qualified_name == "xmlns:x14" || value.qualified_name == "xmlns:xr10" {
            continue;
        }
        output.push(b' ')?;
        output.append(value.qualified_name.as_bytes())?;
        output.append(b"=\"")?;
        output.escape(&value.value)?;
        output.push(b'\"')?;
    }
    Ok(())
}

fn retain_attribute(
    values: &mut Vec<XmlAttribute>,
    total: &mut usize,
    name: String,
    value: String,
) -> Result<()> {
    validate_attribute_name(&name)?;
    validate_xml_value(&value, MAX_TEXT_BYTES, "retained XML attribute")?;
    if values.len() >= MAX_EXTENSION_ATTRIBUTES {
        return Err(limit("extension attributes"));
    }
    *total = total
        .checked_add(name.len() + value.len())
        .ok_or_else(|| limit("extension attribute bytes"))?;
    if *total > MAX_EXTENSION_ATTRIBUTE_BYTES {
        return Err(limit("extension attribute bytes"));
    }
    values
        .try_reserve(1)
        .map_err(|source| allocation("Slicer Cache attributes", source))?;
    values.push(XmlAttribute {
        qualified_name: name,
        value,
    });
    Ok(())
}

fn push_pivot(value: &mut Definition, pivot: PivotTable) -> Result<()> {
    if value.pivot_tables.len() >= MAX_PIVOT_TABLES {
        return Err(limit("pivot table count"));
    }
    value
        .pivot_tables
        .try_reserve(1)
        .map_err(|source| allocation("Slicer Cache pivot tables", source))?;
    value.pivot_tables.push(pivot);
    Ok(())
}

fn retained(xml: &[u8], start: usize, end: usize, total: &mut usize) -> Result<Vec<u8>> {
    let bytes = xml
        .get(start..end)
        .ok_or_else(|| invalid("invalid retained XML offsets"))?;
    if bytes.len() > MAX_INERT_SUBTREE_BYTES {
        return Err(limit("inert subtree bytes"));
    }
    *total = total
        .checked_add(bytes.len())
        .ok_or_else(|| limit("inert subtree bytes"))?;
    if *total > MAX_TOTAL_INERT_BYTES {
        return Err(limit("inert subtree bytes"));
    }
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(bytes.len())
        .map_err(|source| allocation("Slicer Cache retained XML", source))?;
    retained.extend_from_slice(bytes);
    Ok(retained)
}

fn data_kind(namespace: &ResolveResult<'_>, local: &[u8]) -> Result<DataKind> {
    if !exact(namespace, X14) {
        return Err(invalid("Slicer Cache data source has the wrong namespace"));
    }
    match local {
        b"olap" => Ok(DataKind::Olap),
        b"tabular" => Ok(DataKind::Tabular),
        _ => Err(invalid("Slicer Cache data requires olap or tabular")),
    }
}

fn wrap_fragment(fragment: &[u8]) -> Result<Vec<u8>> {
    let mut wrapped = BoundedXml::new(MAX_PART_BYTES, "wrapped XML bytes");
    wrapped.append(b"<x14:slicerCacheDefinition xmlns:x14=\"")?;
    wrapped.escape(X14)?;
    wrapped.append(b"\" name=\"Cache\" sourceName=\"Source\">")?;
    wrapped.append(fragment)?;
    wrapped.append(b"</x14:slicerCacheDefinition>")?;
    Ok(wrapped.finish())
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {name} '{value}'")))
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        Err(invalid(format!("duplicate '{name}' attribute")))
    } else {
        Ok(())
    }
}

fn namespace_string(namespace: &ResolveResult<'_>) -> Result<String> {
    match namespace {
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Bound(Namespace(value)) => std::str::from_utf8(value)
            .map(str::to_owned)
            .map_err(xml_error),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unknown XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn exact(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected.as_bytes())
}

fn validate_event_xml(
    event: &Event<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    retained: bool,
) -> Result<()> {
    match event {
        Event::Start(element) | Event::Empty(element) => {
            validate_element_attributes(element, resolver, decoder, retained)
        },
        Event::Text(text) => {
            let value = text.decode().map_err(xml_error)?;
            validate_xml_value(
                &value,
                if retained {
                    MAX_INERT_SUBTREE_BYTES
                } else {
                    MAX_TEXT_BYTES
                },
                "XML text",
            )
        },
        Event::CData(text) => {
            let value = text.decode().map_err(xml_error)?;
            validate_xml_value(
                &value,
                if retained {
                    MAX_INERT_SUBTREE_BYTES
                } else {
                    MAX_TEXT_BYTES
                },
                "XML CDATA",
            )
        },
        _ => Ok(()),
    }
}

fn validate_element_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    retained: bool,
) -> Result<()> {
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = std::str::from_utf8(item.key.as_ref()).map_err(xml_error)?;
        validate_attribute_name(raw)?;
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        validate_xml_value(
            &value,
            if retained {
                MAX_INERT_SUBTREE_BYTES
            } else {
                MAX_TEXT_BYTES
            },
            "XML attribute",
        )?;
        if raw != "xmlns" && !raw.starts_with("xmlns:") {
            let (namespace, _) = resolver.resolve_attribute(item.key);
            namespace_string(&namespace)?;
        }
    }
    Ok(())
}
fn is_xml_character(character: char) -> bool {
    matches!(character as u32, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(name: &str) -> Error {
    invalid(format!("Slicer Cache {name} limit exceeded"))
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}
#[cfg(test)]
mod tests {
    use super::*;

    fn microsoft() -> String {
        format!(
            r#"<slicerCacheDefinition xmlns="{X14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Slicer_State" sourceName="State"><pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables><data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>"#
        )
    }

    #[test]
    fn microsoft_example_is_typed_and_round_trips() {
        let value = parse(microsoft().as_bytes()).unwrap();
        assert_eq!(
            (&value.name[..], &value.source_name[..]),
            ("Slicer_State", "State")
        );
        assert_eq!(value.pivot_tables, vec![PivotTable::new(1, "PivotTable1")]);
        assert_eq!(value.data.as_ref().unwrap().kind(), DataKind::Tabular);
        assert_eq!(parse(&write(&value).unwrap()).unwrap(), value);
    }

    #[test]
    fn rejects_hostile_grammar_and_bounds() {
        let cases = [
            format!(
                r#"<!DOCTYPE x><slicerCacheDefinition xmlns="{X14}" name="Cache" sourceName="Field"/>"#
            ),
            format!(r#"<slicerCacheDefinition xmlns="{X14}" sourceName="Field"/>"#),
            format!(r#"<slicerCacheDefinition xmlns="{X14}" name="9bad" sourceName="Field"/>"#),
            format!(
                r#"<slicerCacheDefinition xmlns="{X14}" name="Cache" sourceName="Field" xmlns:xr10="{XR10}" xr10:uid="bad"/>"#
            ),
            format!(
                r#"<slicerCacheDefinition xmlns="{X14}" name="Cache" sourceName="Field"><pivotTables/></slicerCacheDefinition>"#
            ),
            format!(
                r#"<slicerCacheDefinition xmlns="{X14}" name="Cache" sourceName="Field"><pivotTables><pivotTable tabId="x" name="P"/></pivotTables></slicerCacheDefinition>"#
            ),
            format!(
                r#"<slicerCacheDefinition xmlns="{X14}" name="Cache" sourceName="Field"><data><olap/><tabular/></data></slicerCacheDefinition>"#
            ),
            format!(
                r#"<slicerCacheDefinition xmlns="{X14}" name="Cache" sourceName="Field"><extLst/><data><tabular/></data></slicerCacheDefinition>"#
            ),
        ];
        for xml in cases {
            assert!(parse(xml.as_bytes()).is_err(), "accepted {xml}");
        }
        assert!(parse(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
        let mut deep = String::new();
        for _ in 0..=MAX_DEPTH {
            deep.push_str("<x14:x>");
        }
        for _ in 0..=MAX_DEPTH {
            deep.push_str("</x14:x>");
        }
        let xml = format!(
            r#"<x14:slicerCacheDefinition xmlns:x14="{X14}" name="Cache" sourceName="Field"><x14:data><x14:tabular>{deep}</x14:tabular></x14:data></x14:slicerCacheDefinition>"#
        );
        assert!(parse(xml.as_bytes()).is_err());
    }

    #[test]
    fn rejects_invalid_xml_values_and_event_bombs() {
        let malformed = format!(
            r#"<x14:slicerCacheDefinition xmlns:x14="{X14}" name="Cache" sourceName="Field"><x14:pivotTables><x14:pivotTable tabId="1" name="P"></x14:pivotTables></x14:slicerCacheDefinition>"#
        );
        assert!(parse(malformed.as_bytes()).is_err());
        let wrong_namespace = r#"<x14:slicerCacheDefinition xmlns:x14="wrong" name="Cache" sourceName="Field"></x14:slicerCacheDefinition>"#;
        assert!(parse(wrong_namespace.as_bytes()).is_err());
        let mut bomb = format!(
            r#"<x14:slicerCacheDefinition xmlns:x14="{X14}" name="Cache" sourceName="Field">"#
        );
        for _ in 0..=MAX_EVENTS {
            bomb.push_str("<!--x-->");
        }
        bomb.push_str("</x14:slicerCacheDefinition>");
        assert!(parse(bomb.as_bytes()).is_err());
    }

    #[test]
    fn writer_rejects_invalid_and_oversized_retained_values() {
        let mut invalid = Definition::new("Cache", "Field");
        invalid.name = "bad\u{0}".into();
        assert!(write(&invalid).is_err());
        let mut oversized = Definition::new("Cache", "Field");
        oversized.data = Some(Data {
            kind: DataKind::Tabular,
            xml: vec![b'x'; MAX_INERT_SUBTREE_BYTES + 1],
        });
        assert!(write(&oversized).is_err());
    }
}
