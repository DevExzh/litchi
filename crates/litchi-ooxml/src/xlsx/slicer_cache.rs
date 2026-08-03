//! Typed support for the MS-XLSX Slicer Cache part.

use crate::error::{OoxmlError, Result};
use crate::xlsx::slicers::{
    SLICERS_CONTENT_TYPE, parse_slicers, validate_package_graph as validate_slicers_graph,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashSet, TryReserveError};

pub const SLICER_CACHE_CONTENT_TYPE: &str = "application/vnd.ms-excel.slicerCache+xml";
pub const SLICER_CACHE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicerCache";

const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XR10: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const CACHE_EXTENSION_URI: &str = "{BBE1A952-AA13-448E-AADC-164F8A28A991}";
const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_WORKBOOK_BYTES: usize = 32 * 1024 * 1024;
const MAX_CACHE_COUNT: usize = 65_536;
const MAX_PIVOT_TABLES: usize = 65_536;
const MAX_DEPTH: usize = 256;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_INERT_SUBTREE_BYTES: usize = 8 * 1024 * 1024;
const MAX_TOTAL_INERT_BYTES: usize = 16 * 1024 * 1024;
const MAX_EXTENSION_ATTRIBUTES: usize = 128;
const MAX_EXTENSION_ATTRIBUTE_BYTES: usize = 64 * 1024;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;
const MAX_EVENTS: usize = 1_000_000;
const MAX_NODES: usize = 500_000;

#[derive(Debug, Clone, PartialEq, Eq)]
struct XmlAttribute {
    qualified_name: String,
    value: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SlicerCacheDataKind {
    Olap,
    Tabular,
}

/// The complete `data` subtree, retained inertly after validating its exact
/// one-of OLAP/tabular discriminator.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlicerCacheData {
    kind: SlicerCacheDataKind,
    xml: Vec<u8>,
}

impl SlicerCacheData {
    pub fn new(xml: Vec<u8>) -> Result<Self> {
        let wrapped = wrap_fragment(&xml)?;
        let definition = parse_slicer_cache_definition(&wrapped)?;
        definition
            .data
            .ok_or_else(|| invalid("expected one Slicer Cache data fragment"))
    }
    pub fn kind(&self) -> SlicerCacheDataKind {
        self.kind
    }
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }
}

/// The complete optional `extLst` subtree, retained inertly and never used to
/// resolve relationships.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlicerCacheExtensionList(Vec<u8>);

impl SlicerCacheExtensionList {
    pub fn new(xml: Vec<u8>) -> Result<Self> {
        let wrapped = wrap_fragment(&xml)?;
        let definition = parse_slicer_cache_definition(&wrapped)?;
        definition
            .extension_list
            .ok_or_else(|| invalid("expected one Slicer Cache extLst fragment"))
    }
    pub fn xml(&self) -> &[u8] {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlicerCachePivotTable {
    pub tab_id: u32,
    pub name: String,
    xml_attributes: Vec<XmlAttribute>,
}

impl SlicerCachePivotTable {
    pub fn new(tab_id: u32, name: impl Into<String>) -> Self {
        Self {
            tab_id,
            name: name.into(),
            xml_attributes: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlicerCacheDefinition {
    pub name: String,
    pub source_name: String,
    pub uid: Option<String>,
    pub pivot_tables: Vec<SlicerCachePivotTable>,
    pub data: Option<SlicerCacheData>,
    pub extension_list: Option<SlicerCacheExtensionList>,
    xml_attributes: Vec<XmlAttribute>,
}

impl SlicerCacheDefinition {
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
pub struct WorkbookSlicerCache {
    pub relationship_id: String,
    pub part_name: String,
    pub definition: SlicerCacheDefinition,
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
    data_kind: Option<SlicerCacheDataKind>,
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

pub fn parse_slicer_cache_definition(xml: &[u8]) -> Result<SlicerCacheDefinition> {
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
    let mut definition: Option<SlicerCacheDefinition> = None;
    let mut stage = 0u8;
    let mut in_pivots = false;
    let mut open_pivot: Option<(usize, SlicerCachePivotTable)> = None;
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
                            definition.extension_list = Some(SlicerCacheExtensionList(bytes));
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
                                target.data = Some(SlicerCacheData { kind, xml: bytes });
                            },
                            CaptureKind::Extension => {
                                target.extension_list = Some(SlicerCacheExtensionList(bytes))
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
    validate_definition(&value)?;
    Ok(value)
}

pub fn write_slicer_cache_definition(value: &SlicerCacheDefinition) -> Result<Vec<u8>> {
    validate_definition(value)?;
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

pub fn load_slicer_caches(package: &OpcPackage) -> Result<Vec<WorkbookSlicerCache>> {
    let workbook = package.main_document_part()?;
    let references = parse_workbook_references(workbook.blob())?;
    validate_cache_graph(package, workbook, &references)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(references.len())
        .map_err(|source| allocation("Slicer Cache collection", source))?;
    for relationship_id in references {
        let relationship = workbook.rels().get(&relationship_id).ok_or_else(|| {
            invalid(format!(
                "missing Slicer Cache relationship '{relationship_id}'"
            ))
        })?;
        let target = relationship.target_partname()?;
        output.push(WorkbookSlicerCache {
            relationship_id,
            part_name: target.to_string(),
            definition: parse_slicer_cache_definition(package.get_part(&target)?.blob())?,
        });
    }
    validate_cache_collection(&output)?;
    validate_slicer_names(
        package,
        output.iter().map(|cache| cache.definition.name.as_str()),
    )?;
    Ok(output)
}

pub fn store_slicer_cache(package: &mut OpcPackage, value: &WorkbookSlicerCache) -> Result<()> {
    validate_relationship_id(&value.relationship_id)?;
    validate_definition(&value.definition)?;
    let xml = write_slicer_cache_definition(&value.definition)?;
    let part_name = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let workbook = package.get_part(&workbook_name)?;
    let references = parse_workbook_references(workbook.blob())?;
    validate_cache_graph(package, workbook, &references)?;
    if references.len() >= MAX_CACHE_COUNT {
        return Err(limit("cache count"));
    }
    if workbook.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(format!(
            "workbook relationship ID '{}' already exists",
            value.relationship_id
        )));
    }
    if package
        .iter_parts()
        .any(|part| part.partname() == &part_name)
    {
        return Err(invalid(format!(
            "Slicer Cache part '{part_name}' already exists"
        )));
    }
    let mut existing = Vec::new();
    existing
        .try_reserve_exact(
            references
                .len()
                .checked_add(1)
                .ok_or_else(|| limit("cache count"))?,
        )
        .map_err(|source| allocation("Slicer Cache collection", source))?;
    for id in &references {
        let relationship = workbook.rels().get(id).ok_or_else(|| {
            invalid(format!(
                "missing Slicer Cache relationship '{id}' during store"
            ))
        })?;
        let target = relationship.target_partname()?;
        existing.push(WorkbookSlicerCache {
            relationship_id: id.clone(),
            part_name: target.to_string(),
            definition: parse_slicer_cache_definition(package.get_part(&target)?.blob())?,
        });
    }
    existing.push(value.clone());
    validate_cache_collection(&existing)?;
    validate_slicers_graph(package)?;
    validate_slicer_names(
        package,
        existing.iter().map(|cache| cache.definition.name.as_str()),
    )?;
    let updated_workbook = add_workbook_reference(workbook.blob(), &value.relationship_id)?;
    let target = part_name.relative_ref(workbook_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(
        part_name,
        SLICER_CACHE_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(&workbook_name)?
        .set_blob(updated_workbook);
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .add_relationship(
            SLICER_CACHE_RELATIONSHIP_TYPE.into(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

fn validate_cache_graph(
    package: &OpcPackage,
    workbook: &dyn Part,
    references: &[String],
) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == SLICER_CACHE_RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "package root cannot source a Slicer Cache relationship",
        ));
    }
    let mut reference_set = HashSet::new();
    reference_set
        .try_reserve(references.len())
        .map_err(|source| allocation("Slicer Cache relationship IDs", source))?;
    for reference in references {
        reference_set.insert(reference.as_str());
    }
    if reference_set.len() != references.len() {
        return Err(invalid("duplicate workbook Slicer Cache reference"));
    }
    let mut targets = HashSet::new();
    let mut cache_relationships = 0usize;
    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == SLICER_CACHE_RELATIONSHIP_TYPE)
        {
            if source.partname() != workbook.partname() {
                return Err(invalid(format!(
                    "Slicer Cache relationship has non-workbook source '{}'",
                    source.partname()
                )));
            }
            if !reference_set.contains(relationship.r_id()) {
                return Err(invalid(format!(
                    "unreferenced Slicer Cache relationship '{}'",
                    relationship.r_id()
                )));
            }
            if relationship.is_external() {
                return Err(invalid("Slicer Cache relationship must be internal"));
            }
            cache_relationships = cache_relationships
                .checked_add(1)
                .ok_or_else(|| limit("relationship count"))?;
            if cache_relationships > MAX_CACHE_COUNT {
                return Err(limit("relationship count"));
            }
            let target = relationship.target_partname()?;
            targets
                .try_reserve(1)
                .map_err(|source| allocation("Slicer Cache targets", source))?;
            if !targets.insert(target.to_string()) {
                return Err(invalid(format!(
                    "Slicer Cache target '{target}' is used more than once"
                )));
            }
            let part = package.get_part(&target)?;
            if part.content_type() != SLICER_CACHE_CONTENT_TYPE {
                return Err(OoxmlError::InvalidContentType {
                    expected: SLICER_CACHE_CONTENT_TYPE.into(),
                    got: part.content_type().into(),
                });
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "Slicer Cache part '{target}' has forbidden outbound relationships"
                )));
            }
        }
    }
    for id in references {
        let relationship = workbook.rels().get(id).ok_or_else(|| {
            invalid(format!(
                "workbook Slicer Cache reference '{id}' has no relationship"
            ))
        })?;
        if relationship.reltype() != SLICER_CACHE_RELATIONSHIP_TYPE {
            return Err(invalid(format!(
                "workbook Slicer Cache reference '{id}' has the wrong relationship type"
            )));
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == SLICER_CACHE_CONTENT_TYPE)
    {
        if !targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "orphan Slicer Cache part '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_cache_collection(values: &[WorkbookSlicerCache]) -> Result<()> {
    if values.len() > MAX_CACHE_COUNT {
        return Err(limit("cache count"));
    }
    let mut names = HashSet::new();
    let mut uids = HashSet::new();
    names
        .try_reserve(values.len())
        .map_err(|source| allocation("Slicer Cache names", source))?;
    uids.try_reserve(values.len())
        .map_err(|source| allocation("Slicer Cache UIDs", source))?;
    let any_uid = values.iter().any(|value| value.definition.uid.is_some());
    for value in values {
        validate_definition(&value.definition)?;
        let folded = value.definition.name.to_lowercase();
        if !names.insert(folded) {
            return Err(invalid(format!(
                "duplicate Slicer Cache name '{}'",
                value.definition.name
            )));
        }
        if let Some(uid) = &value.definition.uid {
            if !uids.insert(uid.to_ascii_lowercase()) {
                return Err(invalid(format!("duplicate Slicer Cache uid '{uid}'")));
            }
        } else if any_uid {
            return Err(invalid(
                "Slicer Cache uid must be present on every cache or none",
            ));
        }
    }
    Ok(())
}

fn validate_slicer_names<'a>(
    package: &OpcPackage,
    cache_names: impl Iterator<Item = &'a str>,
) -> Result<()> {
    let mut names = HashSet::new();
    names
        .try_reserve(MAX_CACHE_COUNT.min(64))
        .map_err(|source| allocation("Slicer Cache names", source))?;
    for name in cache_names {
        names
            .try_reserve(1)
            .map_err(|source| allocation("Slicer Cache names", source))?;
        names.insert(name.to_lowercase());
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == SLICERS_CONTENT_TYPE)
    {
        for slicer in parse_slicers(part.blob())?.slicers {
            if !names.contains(&slicer.cache.to_lowercase()) {
                return Err(invalid(format!(
                    "slicer '{}' references missing cache '{}'",
                    slicer.name, slicer.cache
                )));
            }
        }
    }
    Ok(())
}

fn parse_workbook_references(xml: &[u8]) -> Result<Vec<String>> {
    if xml.len() > MAX_WORKBOOK_BYTES {
        return Err(limit("workbook XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let decoder = reader.decoder();
    let mut depth = 0usize;
    let mut root = false;
    let mut closed = false;
    let mut ext_lst_depth = None;
    let mut target_ext_depth = None;
    let mut caches_depth = None;
    let mut open_reference_depth = None;
    let mut found_target = false;
    let mut found_caches = false;
    let mut references = Vec::new();
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| limit("workbook XML event count"))?;
        if events > MAX_EVENTS {
            return Err(limit("workbook XML event count"));
        }
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("workbook XML offset overflow"))?;
        let borrowed = reader.read_event().map_err(xml_error)?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("workbook XML offset overflow"))?;
        let event_bytes = end
            .checked_sub(start)
            .ok_or_else(|| invalid("workbook XML offsets moved backwards"))?;
        if event_bytes > MAX_INERT_SUBTREE_BYTES {
            return Err(limit("workbook XML event bytes"));
        }
        let event = borrowed.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| limit("workbook XML node count"))?;
            if nodes > MAX_NODES {
                return Err(limit("workbook XML node count"));
            }
        }
        validate_event_xml(&event, &resolver, decoder, false)?;
        match event {
            Event::Start(element) => {
                if closed {
                    return Err(invalid("content after workbook root"));
                }
                if depth == 0 {
                    if !is_core(&namespace) || element.local_name().as_ref() != b"workbook" {
                        return Err(invalid("expected SpreadsheetML workbook root"));
                    }
                    root = true;
                } else if depth == 1
                    && is_core(&namespace)
                    && element.local_name().as_ref() == b"extLst"
                {
                    if ext_lst_depth.replace(depth).is_some() {
                        return Err(invalid("workbook has multiple direct extLst elements"));
                    }
                } else if ext_lst_depth.is_some()
                    && depth == 2
                    && is_core(&namespace)
                    && element.local_name().as_ref() == b"ext"
                {
                    if unqualified_attribute(&element, &resolver, decoder, "uri")?.as_deref()
                        == Some(CACHE_EXTENSION_URI)
                    {
                        if found_target {
                            return Err(invalid("workbook has multiple Slicer Cache extensions"));
                        }
                        found_target = true;
                        target_ext_depth = Some(depth);
                    }
                } else if target_ext_depth.is_some() && depth == 3 {
                    if !exact(&namespace, X14)
                        || element.local_name().as_ref() != b"slicerCaches"
                        || found_caches
                    {
                        return Err(invalid("invalid workbook Slicer Cache extension payload"));
                    }
                    reject_attributes(&element, &resolver, decoder, &[])?;
                    found_caches = true;
                    caches_depth = Some(depth);
                } else if caches_depth.is_some() && depth == 4 {
                    if !exact(&namespace, X14) || element.local_name().as_ref() != b"slicerCache" {
                        return Err(invalid("slicerCaches may contain only slicerCache"));
                    }
                    let id = relationship_id(&element, &resolver, decoder)?;
                    validate_relationship_id(&id)?;
                    if references.len() >= MAX_CACHE_COUNT {
                        return Err(limit("cache references"));
                    }
                    references
                        .try_reserve(1)
                        .map_err(|source| allocation("Slicer Cache references", source))?;
                    references.push(id);
                    open_reference_depth = Some(depth);
                } else if target_ext_depth.is_some() && depth <= 4 {
                    return Err(invalid("invalid workbook Slicer Cache extension structure"));
                } else if open_reference_depth.is_some() {
                    return Err(invalid("workbook slicerCache reference must be empty"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("workbook XML depth"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("workbook XML depth"));
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Err(invalid("workbook root cannot be empty"));
                }
                if ext_lst_depth.is_some()
                    && depth == 2
                    && is_core(&namespace)
                    && element.local_name().as_ref() == b"ext"
                    && unqualified_attribute(&element, &resolver, decoder, "uri")?.as_deref()
                        == Some(CACHE_EXTENSION_URI)
                {
                    return Err(invalid("workbook Slicer Cache extension cannot be empty"));
                }
                if target_ext_depth.is_some() && depth == 3 {
                    return Err(invalid("workbook slicerCaches cannot be empty"));
                }
                if caches_depth.is_some() && depth == 4 {
                    if !exact(&namespace, X14) || element.local_name().as_ref() != b"slicerCache" {
                        return Err(invalid("slicerCaches may contain only slicerCache"));
                    }
                    let id = relationship_id(&element, &resolver, decoder)?;
                    validate_relationship_id(&id)?;
                    if references.len() >= MAX_CACHE_COUNT {
                        return Err(limit("cache references"));
                    }
                    references
                        .try_reserve(1)
                        .map_err(|source| allocation("Slicer Cache references", source))?;
                    references.push(id);
                } else if target_ext_depth.is_some() && depth <= 4 {
                    return Err(invalid("invalid workbook Slicer Cache extension structure"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook closing element"));
                }
                depth -= 1;
                if open_reference_depth == Some(depth) {
                    open_reference_depth = None;
                }
                if caches_depth == Some(depth) {
                    caches_depth = None;
                }
                if target_ext_depth == Some(depth) {
                    if !found_caches || references.is_empty() {
                        return Err(invalid(
                            "workbook Slicer Cache extension requires references",
                        ));
                    }
                    target_ext_depth = None;
                }
                if ext_lst_depth == Some(depth) {
                    ext_lst_depth = None;
                }
                if depth == 0 {
                    if element.local_name().as_ref() != b"workbook" || !root {
                        return Err(invalid("mismatched workbook root"));
                    }
                    closed = true;
                }
            },
            Event::Text(text)
                if target_ext_depth.is_some()
                    && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("text in workbook Slicer Cache extension"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root || !closed || depth != 0 {
        return Err(invalid("incomplete workbook XML"));
    }
    Ok(references)
}

#[derive(Debug)]
struct WorkbookLayout {
    core: &'static str,
    root_close: usize,
    ext_close: Option<usize>,
    empty_ext: Option<(usize, usize)>,
    caches_close: Option<usize>,
}

fn add_workbook_reference(xml: &[u8], relationship_id: &str) -> Result<Vec<u8>> {
    validate_relationship_id(relationship_id)?;
    let existing = parse_workbook_references(xml)?;
    if existing.iter().any(|id| id == relationship_id) {
        return Err(invalid(format!(
            "duplicate Slicer Cache reference '{relationship_id}'"
        )));
    }
    let layout = workbook_layout(xml)?;
    let relationship_namespace = if layout.core == STRICT_SML {
        STRICT_R
    } else {
        R
    };
    let (start, end, insertion) = if let Some(position) = layout.caches_close {
        let mut insertion = BoundedXml::new(MAX_WORKBOOK_BYTES, "rewritten workbook bytes");
        write_workbook_reference(&mut insertion, relationship_id, relationship_namespace)?;
        (position, position, insertion.finish())
    } else if let Some((start, end)) = layout.empty_ext {
        let mut insertion = BoundedXml::new(MAX_WORKBOOK_BYTES, "rewritten workbook bytes");
        insertion.append(b"<x:extLst xmlns:x=\"")?;
        insertion.escape(layout.core)?;
        insertion.append(b"\">")?;
        write_workbook_extension(
            &mut insertion,
            layout.core,
            relationship_id,
            relationship_namespace,
        )?;
        insertion.append(b"</x:extLst>")?;
        (start, end, insertion.finish())
    } else if let Some(position) = layout.ext_close {
        let mut insertion = BoundedXml::new(MAX_WORKBOOK_BYTES, "rewritten workbook bytes");
        write_workbook_extension(
            &mut insertion,
            layout.core,
            relationship_id,
            relationship_namespace,
        )?;
        (position, position, insertion.finish())
    } else {
        let mut insertion = BoundedXml::new(MAX_WORKBOOK_BYTES, "rewritten workbook bytes");
        insertion.append(b"<x:extLst xmlns:x=\"")?;
        insertion.escape(layout.core)?;
        insertion.append(b"\">")?;
        write_workbook_extension(
            &mut insertion,
            layout.core,
            relationship_id,
            relationship_namespace,
        )?;
        insertion.append(b"</x:extLst>")?;
        (layout.root_close, layout.root_close, insertion.finish())
    };
    let removed = end
        .checked_sub(start)
        .ok_or_else(|| invalid("workbook insertion offsets moved backwards"))?;
    let length = xml
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(insertion.len()))
        .ok_or_else(|| limit("rewritten workbook bytes"))?;
    if length > MAX_WORKBOOK_BYTES {
        return Err(limit("rewritten workbook bytes"));
    }
    let mut output = BoundedXml::new(MAX_WORKBOOK_BYTES, "rewritten workbook bytes");
    output.append(
        xml.get(..start)
            .ok_or_else(|| invalid("invalid workbook insertion start"))?,
    )?;
    output.append(&insertion)?;
    output.append(
        xml.get(end..)
            .ok_or_else(|| invalid("invalid workbook insertion end"))?,
    )?;
    Ok(output.finish())
}

fn write_workbook_reference(
    output: &mut BoundedXml,
    relationship_id: &str,
    relationship_namespace: &str,
) -> Result<()> {
    output.append(b"<x14:slicerCache xmlns:x14=\"")?;
    output.escape(X14)?;
    output.append(b"\" xmlns:r=\"")?;
    output.escape(relationship_namespace)?;
    output.append(b"\" r:id=\"")?;
    output.escape(relationship_id)?;
    output.append(b"\"/>")
}

fn write_workbook_extension(
    output: &mut BoundedXml,
    core_namespace: &str,
    relationship_id: &str,
    relationship_namespace: &str,
) -> Result<()> {
    output.append(b"<x:ext xmlns:x=\"")?;
    output.escape(core_namespace)?;
    output.append(b"\" uri=\"")?;
    output.escape(CACHE_EXTENSION_URI)?;
    output.append(b"\"><x14:slicerCaches xmlns:x14=\"")?;
    output.escape(X14)?;
    output.append(b"\">")?;
    write_workbook_reference(output, relationship_id, relationship_namespace)?;
    output.append(b"</x14:slicerCaches></x:ext>")
}

fn workbook_layout(xml: &[u8]) -> Result<WorkbookLayout> {
    if xml.len() > MAX_WORKBOOK_BYTES {
        return Err(limit("workbook XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let decoder = reader.decoder();
    let mut depth = 0usize;
    let mut core = None;
    let mut root_close = None;
    let mut ext_depth = None;
    let mut target_depth = None;
    let mut caches_depth = None;
    let mut ext_close = None;
    let mut empty_ext = None;
    let mut caches_close = None;
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| limit("workbook XML event count"))?;
        if events > MAX_EVENTS {
            return Err(limit("workbook XML event count"));
        }
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("workbook XML offset overflow"))?;
        let borrowed = reader.read_event().map_err(xml_error)?;
        let resolver = reader.resolver().clone();
        let event = borrowed.into_owned();
        let (namespace, event) = resolver.resolve_event(event);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("workbook XML offset overflow"))?;
        let event_bytes = end
            .checked_sub(start)
            .ok_or_else(|| invalid("workbook XML offsets moved backwards"))?;
        if event_bytes > MAX_INERT_SUBTREE_BYTES {
            return Err(limit("workbook XML event bytes"));
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| limit("workbook XML node count"))?;
            if nodes > MAX_NODES {
                return Err(limit("workbook XML node count"));
            }
        }
        validate_event_xml(&event, &resolver, decoder, false)?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    core = Some(core_namespace(&namespace)?);
                } else if depth == 1
                    && is_core(&namespace)
                    && element.local_name().as_ref() == b"extLst"
                {
                    ext_depth = Some(depth);
                } else if ext_depth.is_some()
                    && depth == 2
                    && is_core(&namespace)
                    && element.local_name().as_ref() == b"ext"
                    && unqualified_attribute(&element, &resolver, decoder, "uri")?.as_deref()
                        == Some(CACHE_EXTENSION_URI)
                {
                    target_depth = Some(depth);
                } else if target_depth.is_some()
                    && depth == 3
                    && exact(&namespace, X14)
                    && element.local_name().as_ref() == b"slicerCaches"
                {
                    caches_depth = Some(depth);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("workbook XML depth"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("workbook XML depth"));
                }
            },
            Event::Empty(element)
                if depth == 1
                    && is_core(&namespace)
                    && element.local_name().as_ref() == b"extLst" =>
            {
                empty_ext = Some((start, end))
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook close"));
                }
                if caches_depth == Some(depth - 1) {
                    caches_close = Some(start);
                    caches_depth = None;
                }
                if target_depth == Some(depth - 1) {
                    target_depth = None;
                }
                if ext_depth == Some(depth - 1) {
                    ext_close = Some(start);
                    ext_depth = None;
                }
                depth -= 1;
                if depth == 0 {
                    root_close = Some(start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("incomplete workbook XML"));
    }
    Ok(WorkbookLayout {
        core: core.ok_or_else(|| invalid("missing workbook root"))?,
        root_close: root_close.ok_or_else(|| invalid("missing workbook close"))?,
        ext_close,
        empty_ext,
        caches_close,
    })
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<SlicerCacheDefinition> {
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
    Ok(SlicerCacheDefinition {
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
) -> Result<SlicerCachePivotTable> {
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
    let value = SlicerCachePivotTable {
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

fn validate_definition(value: &SlicerCacheDefinition) -> Result<()> {
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

fn validate_pivot(value: &SlicerCachePivotTable) -> Result<()> {
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
    let bytes = value.as_bytes();
    if bytes.len() != 38
        || bytes[0] != b'{'
        || bytes[37] != b'}'
        || ![9, 14, 19, 24]
            .iter()
            .all(|position| bytes[*position] == b'-')
        || bytes[1..37]
            .iter()
            .enumerate()
            .any(|(index, byte)| ![8, 13, 18, 23].contains(&index) && !byte.is_ascii_hexdigit())
    {
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

fn push_pivot(value: &mut SlicerCacheDefinition, pivot: SlicerCachePivotTable) -> Result<()> {
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

fn data_kind(namespace: &ResolveResult<'_>, local: &[u8]) -> Result<SlicerCacheDataKind> {
    if !exact(namespace, X14) {
        return Err(invalid("Slicer Cache data source has the wrong namespace"));
    }
    match local {
        b"olap" => Ok(SlicerCacheDataKind::Olap),
        b"tabular" => Ok(SlicerCacheDataKind::Tabular),
        _ => Err(invalid("Slicer Cache data requires olap or tabular")),
    }
}

fn relationship_id(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<String> {
    let mut id = None;
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(item.key);
        let namespace = namespace_string(&namespace)?;
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if matches!(namespace.as_str(), R | STRICT_R) && local == "id" {
            set_once(&mut id, value, "r:id")?;
        } else {
            return Err(invalid(format!(
                "unexpected workbook slicerCache attribute '{local}'"
            )));
        }
    }
    id.ok_or_else(|| invalid("workbook slicerCache requires r:id"))
}

fn unqualified_attribute(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    expected: &str,
) -> Result<Option<String>> {
    let mut result = None;
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(item.key);
        let namespace = namespace_string(&namespace)?;
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if namespace.is_empty() && local == expected {
            set_once(&mut result, value, expected)?;
        }
    }
    Ok(result)
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

fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid("invalid Slicer Cache relationship ID length"));
    }
    let mut bytes = value.bytes();
    let first = bytes
        .next()
        .ok_or_else(|| invalid("invalid Slicer Cache relationship ID"))?;
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!(
            "invalid Slicer Cache relationship ID '{value}'"
        )))
    } else {
        Ok(())
    }
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

fn core_namespace(namespace: &ResolveResult<'_>) -> Result<&'static str> {
    if exact(namespace, SML) {
        Ok(SML)
    } else if exact(namespace, STRICT_SML) {
        Ok(STRICT_SML)
    } else {
        Err(invalid("unsupported workbook namespace"))
    }
}
fn is_core(namespace: &ResolveResult<'_>) -> bool {
    exact(namespace, SML) || exact(namespace, STRICT_SML)
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
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("Slicer Cache {name} limit exceeded"))
}

fn allocation(resource: &'static str, source: TryReserveError) -> OoxmlError {
    OoxmlError::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::{Slicer, Slicers, WorksheetSlicers, store_worksheet_slicers};

    fn microsoft() -> String {
        format!(
            r#"<slicerCacheDefinition xmlns="{X14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="{SML}" name="Slicer_State" sourceName="State"><pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables><data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>"#
        )
    }

    fn package() -> (OpcPackage, PackURI, PackURI) {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!("<workbook xmlns=\"{SML}\"><sheets/></workbook>").into_bytes(),
        )));
        package.relate_to(
            "xl/workbook.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            worksheet.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml".into(),
            format!("<worksheet xmlns=\"{SML}\"><sheetData/></worksheet>").into_bytes(),
        )));
        (package, workbook, worksheet)
    }

    fn cache() -> WorkbookSlicerCache {
        WorkbookSlicerCache {
            relationship_id: "rIdCache".into(),
            part_name: "/xl/slicerCaches/slicerCache1.xml".into(),
            definition: parse_slicer_cache_definition(microsoft().as_bytes()).unwrap(),
        }
    }

    #[test]
    fn microsoft_example_is_typed_and_round_trips() {
        let value = parse_slicer_cache_definition(microsoft().as_bytes()).unwrap();
        assert_eq!(
            (&value.name[..], &value.source_name[..]),
            ("Slicer_State", "State")
        );
        assert_eq!(
            value.pivot_tables,
            vec![SlicerCachePivotTable::new(1, "PivotTable1")]
        );
        assert_eq!(
            value.data.as_ref().unwrap().kind(),
            SlicerCacheDataKind::Tabular
        );
        assert_eq!(
            parse_slicer_cache_definition(&write_slicer_cache_definition(&value).unwrap()).unwrap(),
            value
        );
    }

    #[test]
    fn package_store_load_and_slicer_name_cross_validation() {
        let (mut package, _, worksheet) = package();
        store_worksheet_slicers(
            &mut package,
            &worksheet,
            &WorksheetSlicers {
                relationship_id: "rIdSlicers".into(),
                part_name: "/xl/slicers/slicer1.xml".into(),
                slicers: Slicers::new(vec![Slicer::new("StateView", "Slicer_State", 228600)]),
            },
        )
        .unwrap();
        let expected = cache();
        store_slicer_cache(&mut package, &expected).unwrap();
        assert_eq!(load_slicer_caches(&package).unwrap(), vec![expected]);
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
            assert!(
                parse_slicer_cache_definition(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        assert!(parse_slicer_cache_definition(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
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
        assert!(parse_slicer_cache_definition(xml.as_bytes()).is_err());
    }

    #[test]
    fn rejects_invalid_xml_values_and_event_bombs() {
        let malformed = format!(
            r#"<x14:slicerCacheDefinition xmlns:x14="{X14}" name="Cache" sourceName="Field"><x14:pivotTables><x14:pivotTable tabId="1" name="P"></x14:pivotTables></x14:slicerCacheDefinition>"#
        );
        assert!(parse_slicer_cache_definition(malformed.as_bytes()).is_err());

        let wrong_namespace = r#"<x14:slicerCacheDefinition xmlns:x14="wrong" name="Cache" sourceName="Field"></x14:slicerCacheDefinition>"#.to_string();
        assert!(parse_slicer_cache_definition(wrong_namespace.as_bytes()).is_err());

        let invalid_id = format!(
            r#"<workbook xmlns="{SML}"><extLst><ext uri="{CACHE_EXTENSION_URI}"><x14:slicerCaches xmlns:x14="{X14}"><x14:slicerCache xmlns:r="{R}" r:id="1bad"/></x14:slicerCaches></ext></extLst></workbook>"#
        );
        assert!(parse_workbook_references(invalid_id.as_bytes()).is_err());

        let mut bomb = format!(
            r#"<x14:slicerCacheDefinition xmlns:x14="{X14}" name="Cache" sourceName="Field">"#
        );
        for _ in 0..=MAX_EVENTS {
            bomb.push_str("<!--x-->");
        }
        bomb.push_str("</x14:slicerCacheDefinition>");
        assert!(parse_slicer_cache_definition(bomb.as_bytes()).is_err());
    }

    #[test]
    fn writer_rejects_invalid_and_oversized_retained_values() {
        let mut invalid = SlicerCacheDefinition::new("Cache", "Field");
        invalid.name = "bad\u{0}".into();
        assert!(write_slicer_cache_definition(&invalid).is_err());

        let mut oversized = SlicerCacheDefinition::new("Cache", "Field");
        oversized.data = Some(SlicerCacheData {
            kind: SlicerCacheDataKind::Tabular,
            xml: vec![b'x'; MAX_INERT_SUBTREE_BYTES + 1],
        });
        assert!(write_slicer_cache_definition(&oversized).is_err());
    }

    #[test]
    fn rejects_graph_reference_identity_and_cross_name_errors() {
        let (mut missing_ref, workbook, _) = package();
        let part = PackURI::new("/xl/slicerCaches/slicerCache1.xml").unwrap();
        missing_ref.add_part(Box::new(BlobPart::new(
            part.clone(),
            SLICER_CACHE_CONTENT_TYPE.into(),
            microsoft().into_bytes(),
        )));
        missing_ref
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .add_relationship(
                SLICER_CACHE_RELATIONSHIP_TYPE.into(),
                "slicerCaches/slicerCache1.xml".into(),
                "rIdCache".into(),
                false,
            );
        assert!(load_slicer_caches(&missing_ref).is_err());

        let (mut outbound, _, _) = package();
        store_slicer_cache(&mut outbound, &cache()).unwrap();
        outbound
            .get_part_mut(&part)
            .unwrap()
            .rels_mut()
            .add_relationship("urn:forbidden".into(), "x".into(), "rId9".into(), false);
        assert!(load_slicer_caches(&outbound).is_err());

        let (mut mismatch, _, worksheet) = package();
        store_worksheet_slicers(
            &mut mismatch,
            &worksheet,
            &WorksheetSlicers {
                relationship_id: "rIdS".into(),
                part_name: "/xl/slicers/slicer1.xml".into(),
                slicers: Slicers::new(vec![Slicer::new("View", "MissingCache", 1)]),
            },
        )
        .unwrap();
        assert!(store_slicer_cache(&mut mismatch, &cache()).is_err());

        let (mut duplicate, _, _) = package();
        let first = cache();
        store_slicer_cache(&mut duplicate, &first).unwrap();
        let mut second = first.clone();
        second.relationship_id = "rIdCache2".into();
        second.part_name = "/xl/slicerCaches/slicerCache2.xml".into();
        assert!(store_slicer_cache(&mut duplicate, &second).is_err());
    }
}
