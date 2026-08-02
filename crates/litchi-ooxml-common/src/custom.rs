//! Typed custom document properties shared by every OOXML host format.
//!
//! The public surface stays deliberately small: [`Props`] owns a bounded set of
//! named [`Value`]s and can read or write the package-level custom-properties
//! part. Parsing is namespace-aware and rejects ambiguous package graphs and
//! malformed property records instead of treating corruption as absence.

use crate::{Error, Result};
use caseless::Caseless;
use chrono::{DateTime, SecondsFormat, Utc};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;
use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashSet};
use unicode_normalization::UnicodeNormalization;

const PART_NAME: &str = "/docProps/custom.xml";
const PART_TARGET: &str = "docProps/custom.xml";
const FORMAT_ID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
const SUMMARY_FORMAT_ID: &str = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}";
const DOCUMENT_SUMMARY_FORMAT_ID: &str = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}";
const CUSTOM_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
const VT_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 3;
const MAX_XML_NODES: usize = 4_096;
const MAX_ATTRIBUTES: usize = 32;
const MAX_ATTRIBUTE_BYTES: usize = 4_096;
const MAX_PROPERTIES: usize = 1_024;
const MAX_NAME_CHARS: usize = 255;
const MAX_NAME_BYTES: usize = 1_024;
const MAX_TOTAL_NAME_BYTES: usize = 256 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_TOTAL_TEXT_BYTES: usize = 1024 * 1024;

/// A custom document-property value.
///
/// `I64` and `F32` preserve vocabulary accepted by older Litchi releases even
/// though Microsoft Office's standard custom-property producer profile uses
/// `I32` and `F64`. Values are never silently narrowed or widened.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    /// An explicit `vt:empty` value.
    Empty,
    /// XML text. New values are written as `vt:lpwstr`; a parsed `vt:lpstr`
    /// retains that wire kind on subsequent writes.
    Text(String),
    /// A signed 32-bit integer (`vt:i4`).
    I32(i32),
    /// A signed 64-bit integer (`vt:i8`).
    I64(i64),
    /// A finite 32-bit float (`vt:r4`).
    F32(f32),
    /// A finite 64-bit float (`vt:r8`).
    F64(f64),
    /// A Boolean (`vt:bool`).
    Bool(bool),
    /// A UTC instant (`vt:filetime`) serialized as RFC3339/XML date-time text.
    Time(DateTime<Utc>),
}

impl From<()> for Value {
    fn from((): ()) -> Self {
        Self::Empty
    }
}

impl From<String> for Value {
    fn from(value: String) -> Self {
        Self::Text(value)
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Self::Text(value.to_owned())
    }
}

impl From<i32> for Value {
    fn from(value: i32) -> Self {
        Self::I32(value)
    }
}

impl From<i64> for Value {
    fn from(value: i64) -> Self {
        Self::I64(value)
    }
}

impl From<f32> for Value {
    fn from(value: f32) -> Self {
        Self::F32(value)
    }
}

impl From<f64> for Value {
    fn from(value: f64) -> Self {
        Self::F64(value)
    }
}

impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<DateTime<Utc>> for Value {
    fn from(value: DateTime<Utc>) -> Self {
        Self::Time(value)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WireKind {
    Empty,
    Lpstr,
    Lpwstr,
    I4,
    I8,
    R4,
    R8,
    Bool,
    Filetime,
}

impl WireKind {
    fn qualified_name(self) -> &'static str {
        match self {
            Self::Empty => "vt:empty",
            Self::Lpstr => "vt:lpstr",
            Self::Lpwstr => "vt:lpwstr",
            Self::I4 => "vt:i4",
            Self::I8 => "vt:i8",
            Self::R4 => "vt:r4",
            Self::R8 => "vt:r8",
            Self::Bool => "vt:bool",
            Self::Filetime => "vt:filetime",
        }
    }

    fn from_local_name(name: &[u8]) -> Option<Self> {
        match name {
            b"empty" => Some(Self::Empty),
            b"lpstr" => Some(Self::Lpstr),
            b"lpwstr" => Some(Self::Lpwstr),
            b"i4" => Some(Self::I4),
            b"i8" => Some(Self::I8),
            b"r4" => Some(Self::R4),
            b"r8" => Some(Self::R8),
            b"bool" => Some(Self::Bool),
            b"filetime" => Some(Self::Filetime),
            _ => None,
        }
    }

    fn for_value(value: &Value) -> Self {
        match value {
            Value::Empty => Self::Empty,
            Value::Text(_) => Self::Lpwstr,
            Value::I32(_) => Self::I4,
            Value::I64(_) => Self::I8,
            Value::F32(_) => Self::R4,
            Value::F64(_) => Self::R8,
            Value::Bool(_) => Self::Bool,
            Value::Time(_) => Self::Filetime,
        }
    }

    fn is_text(self) -> bool {
        matches!(self, Self::Lpstr | Self::Lpwstr)
    }
}

#[derive(Debug)]
struct Property {
    pid: i32,
    format_id: String,
    wire: WireKind,
    value: Value,
}

/// A bounded collection of custom document properties.
///
/// Names are unique case-insensitively. Exact-name insertion replaces a value
/// while retaining its PID; insertion with different casing is rejected as
/// ambiguous. Iterators use lexical name order, while XML output uses PID order.
#[derive(Debug)]
pub struct Props {
    properties: BTreeMap<String, Property>,
    folded_names: BTreeMap<String, String>,
    next_pid: Option<i32>,
    name_bytes: usize,
    text_bytes: usize,
}

impl Default for Props {
    fn default() -> Self {
        Self::new()
    }
}

impl Props {
    /// Creates an empty property collection.
    #[must_use]
    pub fn new() -> Self {
        Self {
            properties: BTreeMap::new(),
            folded_names: BTreeMap::new(),
            next_pid: Some(2),
            name_bytes: 0,
            text_bytes: 0,
        }
    }

    /// Inserts or replaces a property.
    ///
    /// New PIDs are allocated monotonically with checked arithmetic. The old
    /// value is moved out when an exact-name property is replaced.
    pub fn insert(
        &mut self,
        name: impl Into<String>,
        value: impl Into<Value>,
    ) -> Result<Option<Value>> {
        let name = name.into();
        let value = value.into();
        validate_name(&name)?;
        validate_value(&value)?;

        if let Some(property) = self.properties.get_mut(&name) {
            let old_text = value_text_bytes(&property.value);
            let new_text = value_text_bytes(&value);
            let base = self
                .text_bytes
                .checked_sub(old_text)
                .ok_or_else(|| invalid("custom-property text accounting is inconsistent"))?;
            let updated = checked_total(
                base,
                new_text,
                MAX_TOTAL_TEXT_BYTES,
                "custom-property text bytes",
            )?;
            let wire = if property.wire.is_text() && matches!(value, Value::Text(_)) {
                property.wire
            } else {
                WireKind::for_value(&value)
            };
            self.text_bytes = updated;
            property.wire = wire;
            return Ok(Some(std::mem::replace(&mut property.value, value)));
        }

        let folded = fold_name(&name);
        if self.folded_names.contains_key(&folded) {
            return Err(invalid(format!(
                "custom property name '{name}' duplicates an existing name case-insensitively"
            )));
        }
        let actual_count = checked_increment(self.properties.len(), "custom properties")?;
        if actual_count > MAX_PROPERTIES {
            return Err(limit("custom properties", MAX_PROPERTIES, actual_count));
        }
        let names = checked_total(
            self.name_bytes,
            name.len(),
            MAX_TOTAL_NAME_BYTES,
            "custom-property name bytes",
        )?;
        let texts = checked_total(
            self.text_bytes,
            value_text_bytes(&value),
            MAX_TOTAL_TEXT_BYTES,
            "custom-property text bytes",
        )?;
        let pid = self
            .next_pid
            .ok_or_else(|| invalid("custom property PID space is exhausted"))?;
        self.next_pid = pid.checked_add(1);
        let property = Property {
            pid,
            format_id: FORMAT_ID.to_owned(),
            wire: WireKind::for_value(&value),
            value,
        };
        self.folded_names.insert(folded, name.clone());
        self.properties.insert(name, property);
        self.name_bytes = names;
        self.text_bytes = texts;
        Ok(None)
    }

    /// Borrows a property by its case-insensitive name.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Value> {
        self.folded_names
            .get(&fold_name(name))
            .and_then(|stored| self.properties.get(stored))
            .map(|property| &property.value)
    }

    /// Removes a property and moves out its value.
    pub fn remove(&mut self, name: &str) -> Option<Value> {
        let stored = self.folded_names.remove(&fold_name(name))?;
        let property = self.properties.remove(&stored)?;
        self.name_bytes = self.name_bytes.saturating_sub(stored.len());
        self.text_bytes = self
            .text_bytes
            .saturating_sub(value_text_bytes(&property.value));
        Some(property.value)
    }

    /// Returns property names in lexical order.
    pub fn names(&self) -> impl ExactSizeIterator<Item = &str> {
        self.properties.keys().map(String::as_str)
    }

    /// Returns name/value pairs in lexical name order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = (&str, &Value)> {
        self.properties
            .iter()
            .map(|(name, property)| (name.as_str(), &property.value))
    }

    /// Removes every property and resets PID allocation to 2.
    pub fn clear(&mut self) {
        self.properties.clear();
        self.folded_names.clear();
        self.next_pid = Some(2);
        self.name_bytes = 0;
        self.text_bytes = 0;
    }

    /// Returns whether a case-insensitive property name is present.
    #[must_use]
    pub fn contains(&self, name: &str) -> bool {
        self.folded_names.contains_key(&fold_name(name))
    }

    /// Returns the number of properties.
    #[must_use]
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Returns whether no properties are present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    /// Reads the custom-properties relationship and target part from a package.
    ///
    /// A genuinely absent relationship and part produce empty properties.
    /// Orphans, duplicate relationships, external targets, missing targets,
    /// wrong content types, and malformed XML are returned as typed errors.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        let graph = inspect_graph(package)?;
        let Some(part_name) = graph.part_name else {
            return Ok(Self::new());
        };
        let part = package.get_part(&part_name)?;
        decode(part.blob())
    }

    /// Writes this collection to a package.
    ///
    /// Empty properties remove both the target part and its package-level
    /// relationship. Non-empty properties update the existing validated target
    /// or create the canonical `/docProps/custom.xml` part and relationship.
    pub fn write(&self, package: &mut OpcPackage) -> Result<()> {
        let graph = inspect_graph(package)?;
        if self.is_empty() {
            if graph.part_name.is_none() && graph.relationship_id.is_none() {
                return Ok(());
            }
            package.unsign();
            if let Some(part_name) = graph.part_name {
                let removed = package.remove_part(&part_name);
                if !removed {
                    return Err(Error::Missing(part_name.to_string()));
                }
            }
            if let Some(relationship_id) = graph.relationship_id {
                let removed = package.rels_mut().remove(&relationship_id);
                if removed.is_none() {
                    return Err(Error::Relationship(format!(
                        "custom-properties relationship '{relationship_id}' disappeared during removal"
                    )));
                }
            }
            return Ok(());
        }

        let xml = encode(self)?;
        match graph.part_name {
            Some(part_name) => {
                if package.get_part(&part_name)?.blob() == xml.as_slice() {
                    return Ok(());
                }
                package.get_part_mut(&part_name)?.set_blob(xml);
                package.unsign();
            },
            None => {
                let part_name = custom_part_name()?;
                package.validate_new_part_name(&part_name)?;
                let part = BlobPart::new(part_name, ct::OFC_CUSTOM_PROPERTIES.to_owned(), xml);
                package.unsign();
                package.add_part(Box::new(part));
                package.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
            },
        }
        Ok(())
    }

    fn insert_parsed(
        &mut self,
        name: String,
        pid: i32,
        format_id: String,
        wire: WireKind,
        value: Value,
    ) -> Result<()> {
        validate_name(&name)?;
        validate_value(&value)?;
        if pid < 2 {
            return Err(invalid(format!(
                "custom property '{name}' has PID {pid}; PIDs must be at least 2"
            )));
        }
        let folded = fold_name(&name);
        if self.folded_names.contains_key(&folded) {
            return Err(invalid(format!(
                "duplicate custom property name '{name}' (names are case-insensitive)"
            )));
        }
        let actual_count = checked_increment(self.properties.len(), "custom properties")?;
        if actual_count > MAX_PROPERTIES {
            return Err(limit("custom properties", MAX_PROPERTIES, actual_count));
        }
        let names = checked_total(
            self.name_bytes,
            name.len(),
            MAX_TOTAL_NAME_BYTES,
            "custom-property name bytes",
        )?;
        let texts = checked_total(
            self.text_bytes,
            value_text_bytes(&value),
            MAX_TOTAL_TEXT_BYTES,
            "custom-property text bytes",
        )?;
        self.folded_names.insert(folded, name.clone());
        self.properties.insert(
            name,
            Property {
                pid,
                format_id,
                wire,
                value,
            },
        );
        self.name_bytes = names;
        self.text_bytes = texts;
        self.next_pid = match self.next_pid {
            Some(next) if pid >= next => pid.checked_add(1),
            next => next,
        };
        Ok(())
    }
}

struct PackageGraph {
    part_name: Option<PackURI>,
    relationship_id: Option<String>,
}

fn inspect_graph(package: &OpcPackage) -> Result<PackageGraph> {
    let canonical = custom_part_name()?;
    let mut custom_parts = Vec::new();
    for part in package.iter_parts() {
        if part
            .partname()
            .as_str()
            .eq_ignore_ascii_case(canonical.as_str())
            && part.content_type() != ct::OFC_CUSTOM_PROPERTIES
        {
            return Err(Error::ContentType {
                expected: ct::OFC_CUSTOM_PROPERTIES.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        if part.content_type() == ct::OFC_CUSTOM_PROPERTIES {
            custom_parts.push(part.partname().clone());
            if custom_parts.len() > 1 {
                return Err(invalid("package contains multiple custom-properties parts"));
            }
        }
        if part
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
        {
            return Err(Error::Relationship(format!(
                "custom-properties relationship must be package-level, not owned by '{}'",
                part.partname().as_str()
            )));
        }
    }

    let relationships: Vec<_> = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
        .collect();
    if relationships.len() > 1 {
        return Err(Error::Relationship(
            "package contains multiple custom-properties relationships".to_owned(),
        ));
    }

    let Some(relationship) = relationships.first().copied() else {
        if let Some(part_name) = custom_parts.first() {
            return Err(Error::Relationship(format!(
                "custom-properties part '{}' is orphaned",
                part_name.as_str()
            )));
        }
        return Ok(PackageGraph {
            part_name: None,
            relationship_id: None,
        });
    };
    if relationship.is_external() {
        return Err(Error::Relationship(
            "custom-properties relationship cannot be external".to_owned(),
        ));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Relationship(
            "custom-properties relationship target cannot contain a query or fragment".to_owned(),
        ));
    }
    let target = relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "invalid custom-properties relationship target: {error}"
        ))
    })?;
    let target_part = package
        .iter_parts()
        .find(|part| {
            part.partname()
                .as_str()
                .eq_ignore_ascii_case(target.as_str())
        })
        .ok_or_else(|| Error::Missing(target.to_string()))?;
    if target_part.content_type() != ct::OFC_CUSTOM_PROPERTIES {
        return Err(Error::ContentType {
            expected: ct::OFC_CUSTOM_PROPERTIES.to_owned(),
            actual: target_part.content_type().to_owned(),
        });
    }
    let part_name = target_part.partname().clone();
    if custom_parts
        .first()
        .is_some_and(|candidate| !candidate.as_str().eq_ignore_ascii_case(part_name.as_str()))
    {
        return Err(Error::Relationship(
            "custom-properties relationship does not target the unique custom-properties part"
                .to_owned(),
        ));
    }

    Ok(PackageGraph {
        part_name: Some(part_name),
        relationship_id: Some(relationship.r_id().to_owned()),
    })
}

fn custom_part_name() -> Result<PackURI> {
    PackURI::new(PART_NAME).map_err(|error| Error::Uri(error.to_string()))
}

struct PendingProperty {
    name: String,
    pid: i32,
    format_id: String,
    value: Option<(WireKind, Value)>,
}

fn decode(xml: &[u8]) -> Result<Props> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit(
            "custom-properties XML bytes",
            MAX_XML_BYTES,
            xml.len(),
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut props = Props::new();
    let mut pids = HashSet::new();
    let mut pending: Option<PendingProperty> = None;
    let mut value_kind: Option<WireKind> = None;
    let mut value_text = String::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Decl(_) => {
                count_node(&mut nodes)?;
                if declaration_seen || root_seen {
                    return Err(invalid("XML declaration must occur once before the root"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "DTD declarations are forbidden in custom-properties XML",
                ));
            },
            Event::PI(_) => {
                return Err(invalid(
                    "processing instructions are forbidden in custom-properties XML",
                ));
            },
            Event::Start(element) => {
                count_node(&mut nodes)?;
                let child_depth = checked_depth(depth)?;
                match child_depth {
                    1 => {
                        if root_seen || root_closed {
                            return Err(invalid(
                                "custom-properties XML must contain exactly one root",
                            ));
                        }
                        validate_root(&namespace, &element, decoder)?;
                        root_seen = true;
                    },
                    2 => {
                        if !is_name(&namespace, &element, CUSTOM_NS, b"property") {
                            return Err(invalid(format!(
                                "unexpected element '{}' below custom-properties root",
                                display_name(element.name().as_ref())
                            )));
                        }
                        if pending.is_some() {
                            return Err(invalid("custom properties cannot be nested"));
                        }
                        let parsed = parse_property_attributes(&element, decoder)?;
                        if !pids.insert(parsed.pid) {
                            return Err(invalid(format!(
                                "duplicate custom property PID {}",
                                parsed.pid
                            )));
                        }
                        pending = Some(parsed);
                    },
                    3 => {
                        let property = pending.as_ref().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        if property.value.is_some() || value_kind.is_some() {
                            return Err(invalid(format!(
                                "custom property '{}' must contain exactly one value",
                                property.name
                            )));
                        }
                        value_kind = Some(parse_value_element(&namespace, &element)?);
                        value_text.clear();
                        validate_value_attributes(&element)?;
                    },
                    _ => {
                        return Err(limit(
                            "custom-properties XML depth",
                            MAX_XML_DEPTH,
                            child_depth,
                        ));
                    },
                }
                depth = child_depth;
            },
            Event::Empty(element) => {
                count_node(&mut nodes)?;
                let child_depth = checked_depth(depth)?;
                match child_depth {
                    1 => {
                        if root_seen || root_closed {
                            return Err(invalid(
                                "custom-properties XML must contain exactly one root",
                            ));
                        }
                        validate_root(&namespace, &element, decoder)?;
                        root_seen = true;
                        root_closed = true;
                    },
                    2 => {
                        return Err(invalid(
                            "custom property must contain exactly one typed value",
                        ));
                    },
                    3 => {
                        let property = pending.as_mut().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        if property.value.is_some() || value_kind.is_some() {
                            return Err(invalid(format!(
                                "custom property '{}' must contain exactly one value",
                                property.name
                            )));
                        }
                        let kind = parse_value_element(&namespace, &element)?;
                        validate_value_attributes(&element)?;
                        property.value = Some((kind, parse_value(kind, "")?));
                    },
                    _ => {
                        return Err(limit(
                            "custom-properties XML depth",
                            MAX_XML_DEPTH,
                            child_depth,
                        ));
                    },
                }
            },
            Event::End(_) => {
                count_node(&mut nodes)?;
                match depth {
                    3 => {
                        let kind = value_kind
                            .take()
                            .ok_or_else(|| invalid("custom-property value state is incomplete"))?;
                        let value = parse_value(kind, &value_text)?;
                        let property = pending.as_mut().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        property.value = Some((kind, value));
                        value_text.clear();
                    },
                    2 => {
                        let property = pending
                            .take()
                            .ok_or_else(|| invalid("custom-property record is incomplete"))?;
                        let (wire, value) = property.value.ok_or_else(|| {
                            invalid(format!(
                                "custom property '{}' must contain exactly one value",
                                property.name
                            ))
                        })?;
                        props.insert_parsed(
                            property.name,
                            property.pid,
                            property.format_id,
                            wire,
                            value,
                        )?;
                    },
                    1 => {
                        root_closed = true;
                    },
                    _ => return Err(invalid("unexpected XML end element")),
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("custom-properties XML depth underflow"))?;
            },
            Event::Text(text) => {
                count_node(&mut nodes)?;
                let text = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(format!("invalid XML text: {error}")))?;
                if depth == 3 {
                    append_value_text(&mut value_text, &text)?;
                } else if !text.trim().is_empty() {
                    return Err(invalid(
                        "non-whitespace text is not allowed outside a property value",
                    ));
                }
            },
            Event::CData(text) => {
                count_node(&mut nodes)?;
                if depth != 3 {
                    return Err(invalid("CDATA is only allowed inside a property value"));
                }
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(format!("invalid CDATA text: {error}")))?;
                append_value_text(&mut value_text, &text)?;
            },
            Event::GeneralRef(reference) => {
                count_node(&mut nodes)?;
                if depth != 3 {
                    return Err(invalid(
                        "entity references are only allowed inside a property value",
                    ));
                }
                let decoded = decode_reference(&reference)?;
                append_value_text(&mut value_text, &decoded)?;
            },
            Event::Comment(_) => count_node(&mut nodes)?,
            Event::Eof => break,
        }
    }

    if !root_seen || !root_closed || depth != 0 || pending.is_some() || value_kind.is_some() {
        return Err(invalid(
            "custom-properties XML must contain one complete Properties root",
        ));
    }
    Ok(props)
}

fn encode(props: &Props) -> Result<Vec<u8>> {
    validate_collection(props)?;
    let estimated = estimated_xml_size(props)?;
    let mut writer = Writer::new(Vec::with_capacity(estimated));
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    let mut root = BytesStart::new("Properties");
    root.push_attribute((
        "xmlns",
        std::str::from_utf8(CUSTOM_NS).map_err(|error| {
            Error::Xml(format!(
                "invalid built-in custom-properties namespace: {error}"
            ))
        })?,
    ));
    root.push_attribute((
        "xmlns:vt",
        std::str::from_utf8(VT_NS).map_err(|error| {
            Error::Xml(format!("invalid built-in variant-types namespace: {error}"))
        })?,
    ));
    writer.write_event(Event::Start(root))?;

    let mut ordered: Vec<_> = props.properties.iter().collect();
    ordered.sort_unstable_by_key(|(_, property)| property.pid);
    for (name, property) in ordered {
        let pid = property.pid.to_string();
        let mut element = BytesStart::new("property");
        element.push_attribute(("fmtid", property.format_id.as_str()));
        element.push_attribute(("pid", pid.as_str()));
        element.push_attribute(("name", name.as_str()));
        writer.write_event(Event::Start(element))?;

        let value_name = property.wire.qualified_name();
        if property.wire == WireKind::Empty {
            writer.write_event(Event::Empty(BytesStart::new(value_name)))?;
        } else {
            writer.write_event(Event::Start(BytesStart::new(value_name)))?;
            let value = value_lexical(&property.value)?;
            writer.write_event(Event::Text(BytesText::new(&value)))?;
            writer.write_event(Event::End(BytesEnd::new(value_name)))?;
        }
        writer.write_event(Event::End(BytesEnd::new("property")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("Properties")))?;
    let xml = writer.into_inner();
    if xml.len() > MAX_XML_BYTES {
        return Err(limit(
            "custom-properties XML bytes",
            MAX_XML_BYTES,
            xml.len(),
        ));
    }
    Ok(xml)
}

fn validate_collection(props: &Props) -> Result<()> {
    if props.properties.len() > MAX_PROPERTIES {
        return Err(limit(
            "custom properties",
            MAX_PROPERTIES,
            props.properties.len(),
        ));
    }
    let mut pids = HashSet::with_capacity(props.properties.len());
    let mut folded = BTreeSet::new();
    let mut name_bytes = 0usize;
    let mut text_bytes = 0usize;
    for (name, property) in &props.properties {
        validate_name(name)?;
        validate_value(&property.value)?;
        validate_wire_value(property.wire, &property.value)?;
        if property.pid < 2 || !pids.insert(property.pid) {
            return Err(invalid(format!(
                "invalid or duplicate custom property PID {}",
                property.pid
            )));
        }
        validate_format_id(&property.format_id)?;
        if !folded.insert(fold_name(name)) {
            return Err(invalid(format!(
                "duplicate custom property name '{name}' (names are case-insensitive)"
            )));
        }
        name_bytes = checked_total(
            name_bytes,
            name.len(),
            MAX_TOTAL_NAME_BYTES,
            "custom-property name bytes",
        )?;
        text_bytes = checked_total(
            text_bytes,
            value_text_bytes(&property.value),
            MAX_TOTAL_TEXT_BYTES,
            "custom-property text bytes",
        )?;
    }
    Ok(())
}

fn estimated_xml_size(props: &Props) -> Result<usize> {
    let mut size = 256usize;
    for (name, property) in &props.properties {
        size = size
            .checked_add(256)
            .and_then(|size| {
                name.len()
                    .checked_mul(6)
                    .and_then(|name| size.checked_add(name))
            })
            .and_then(|size| {
                value_text_bytes(&property.value)
                    .checked_mul(6)
                    .and_then(|text| size.checked_add(text))
            })
            .ok_or_else(|| limit("custom-properties XML bytes", MAX_XML_BYTES, usize::MAX))?;
        if size > MAX_XML_BYTES {
            return Err(limit("custom-properties XML bytes", MAX_XML_BYTES, size));
        }
    }
    Ok(size)
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    if !is_name(namespace, element, CUSTOM_NS, b"Properties") {
        return Err(invalid(format!(
            "custom-properties root must be Properties in namespace '{}'",
            String::from_utf8_lossy(CUSTOM_NS)
        )));
    }
    let mut count = 0usize;
    for attribute in element.attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if !is_namespace_declaration(attribute.key.as_ref()) {
            let name = attribute.key.local_name().as_ref().to_vec();
            let _ = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(format!("invalid root attribute: {error}")))?;
            return Err(invalid(format!(
                "unexpected custom-properties root attribute '{}'",
                display_name(&name)
            )));
        }
    }
    Ok(())
}

fn parse_property_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<PendingProperty> {
    let mut name = None;
    let mut pid = None;
    let mut format_id = None;
    let mut count = 0usize;
    for attribute in element.attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        if attribute.key.prefix().is_some() {
            return Err(invalid(format!(
                "custom property has unexpected qualified attribute '{}'",
                display_name(attribute.key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(format!("invalid property attribute: {error}")))?
            .into_owned();
        match attribute.key.local_name().as_ref() {
            b"name" if name.is_none() => name = Some(value),
            b"pid" if pid.is_none() => {
                let parsed = value.parse::<i32>().map_err(|error| {
                    invalid(format!("invalid custom property PID '{value}': {error}"))
                })?;
                pid = Some(parsed);
            },
            b"fmtid" if format_id.is_none() => format_id = Some(normalize_format_id(&value)?),
            local => {
                return Err(invalid(format!(
                    "duplicate or unexpected custom property attribute '{}'",
                    display_name(local)
                )));
            },
        }
    }
    let name = name.ok_or_else(|| invalid("custom property is missing its name attribute"))?;
    validate_name(&name)?;
    let pid = pid.ok_or_else(|| invalid(format!("custom property '{name}' is missing its PID")))?;
    if pid < 2 {
        return Err(invalid(format!(
            "custom property '{name}' has PID {pid}; PIDs must be at least 2"
        )));
    }
    let format_id = format_id
        .ok_or_else(|| invalid(format!("custom property '{name}' is missing its format ID")))?;
    Ok(PendingProperty {
        name,
        pid,
        format_id,
        value: None,
    })
}

fn validate_value_attributes(element: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err(invalid(format!(
                "custom-property value has unexpected attribute '{}'",
                display_name(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn parse_value_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Result<WireKind> {
    if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == VT_NS) {
        return Err(invalid(format!(
            "custom-property value '{}' is not in the variant-types namespace",
            display_name(element.name().as_ref())
        )));
    }
    WireKind::from_local_name(element.local_name().as_ref()).ok_or_else(|| {
        invalid(format!(
            "unsupported custom-property value type '{}'",
            display_name(element.local_name().as_ref())
        ))
    })
}

fn parse_value(kind: WireKind, text: &str) -> Result<Value> {
    match kind {
        WireKind::Empty => {
            if text.is_empty() {
                Ok(Value::Empty)
            } else {
                Err(invalid("vt:empty cannot contain text"))
            }
        },
        WireKind::Lpstr | WireKind::Lpwstr => {
            validate_xml_text(text, "custom-property text")?;
            Ok(Value::Text(text.to_owned()))
        },
        WireKind::I4 => text
            .trim()
            .parse::<i32>()
            .map(Value::I32)
            .map_err(|error| invalid(format!("invalid vt:i4 value '{text}': {error}"))),
        WireKind::I8 => text
            .trim()
            .parse::<i64>()
            .map(Value::I64)
            .map_err(|error| invalid(format!("invalid vt:i8 value '{text}': {error}"))),
        WireKind::R4 => {
            let value = text
                .trim()
                .parse::<f32>()
                .map_err(|error| invalid(format!("invalid vt:r4 value '{text}': {error}")))?;
            if !value.is_finite() {
                return Err(invalid("vt:r4 custom property must be finite"));
            }
            Ok(Value::F32(value))
        },
        WireKind::R8 => {
            let value = text
                .trim()
                .parse::<f64>()
                .map_err(|error| invalid(format!("invalid vt:r8 value '{text}': {error}")))?;
            if !value.is_finite() {
                return Err(invalid("vt:r8 custom property must be finite"));
            }
            Ok(Value::F64(value))
        },
        WireKind::Bool => match text.trim() {
            "true" | "1" => Ok(Value::Bool(true)),
            "false" | "0" => Ok(Value::Bool(false)),
            value => Err(invalid(format!("invalid vt:bool value '{value}'"))),
        },
        WireKind::Filetime => {
            let value = DateTime::parse_from_rfc3339(text.trim()).map_err(|error| {
                invalid(format!(
                    "invalid vt:filetime RFC3339 date-time '{text}': {error}"
                ))
            })?;
            Ok(Value::Time(value.with_timezone(&Utc)))
        },
    }
}

fn value_lexical(value: &Value) -> Result<Cow<'_, str>> {
    validate_value(value)?;
    Ok(match value {
        Value::Empty => Cow::Borrowed(""),
        Value::Text(text) => Cow::Borrowed(text),
        Value::I32(value) => Cow::Owned(value.to_string()),
        Value::I64(value) => Cow::Owned(value.to_string()),
        Value::F32(value) => Cow::Owned(value.to_string()),
        Value::F64(value) => Cow::Owned(value.to_string()),
        Value::Bool(true) => Cow::Borrowed("true"),
        Value::Bool(false) => Cow::Borrowed("false"),
        Value::Time(value) => Cow::Owned(value.to_rfc3339_opts(SecondsFormat::AutoSi, true)),
    })
}

fn validate_wire_value(wire: WireKind, value: &Value) -> Result<()> {
    let matches = matches!(
        (wire, value),
        (WireKind::Empty, Value::Empty)
            | (WireKind::Lpstr | WireKind::Lpwstr, Value::Text(_))
            | (WireKind::I4, Value::I32(_))
            | (WireKind::I8, Value::I64(_))
            | (WireKind::R4, Value::F32(_))
            | (WireKind::R8, Value::F64(_))
            | (WireKind::Bool, Value::Bool(_))
            | (WireKind::Filetime, Value::Time(_))
    );
    if matches {
        Ok(())
    } else {
        Err(invalid(
            "custom-property wire type does not match its value",
        ))
    }
}

fn validate_value(value: &Value) -> Result<()> {
    match value {
        Value::Text(text) => {
            if text.len() > MAX_TEXT_BYTES {
                return Err(limit(
                    "custom-property text bytes",
                    MAX_TEXT_BYTES,
                    text.len(),
                ));
            }
            validate_xml_text(text, "custom-property text")
        },
        Value::F32(value) if !value.is_finite() => {
            Err(invalid("F32 custom property must be finite"))
        },
        Value::F64(value) if !value.is_finite() => {
            Err(invalid("F64 custom property must be finite"))
        },
        _ => Ok(()),
    }
}

fn validate_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(invalid("custom property name cannot be empty"));
    }
    if name.len() > MAX_NAME_BYTES {
        return Err(limit(
            "custom-property name bytes",
            MAX_NAME_BYTES,
            name.len(),
        ));
    }
    let chars = name.chars().count();
    if chars > MAX_NAME_CHARS {
        return Err(limit(
            "custom-property name characters",
            MAX_NAME_CHARS,
            chars,
        ));
    }
    validate_xml_text(name, "custom-property name")
}

fn validate_xml_text(value: &str, label: &str) -> Result<()> {
    if let Some(character) = value.chars().find(|character| !is_xml_10_char(*character)) {
        return Err(invalid(format!(
            "{label} contains XML 1.0-forbidden character U+{:04X}",
            u32::from(character)
        )));
    }
    Ok(())
}

fn is_xml_10_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(u32::from(character), 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn validate_format_id(format_id: &str) -> Result<()> {
    let normalized = normalize_format_id(format_id)?;
    if normalized == SUMMARY_FORMAT_ID || normalized == DOCUMENT_SUMMARY_FORMAT_ID {
        return Err(invalid(format!(
            "format ID {normalized} is forbidden for custom properties"
        )));
    }
    Ok(())
}

fn normalize_format_id(format_id: &str) -> Result<String> {
    let bytes = format_id.as_bytes();
    let valid = bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && [9, 14, 19, 24]
            .iter()
            .all(|position| bytes.get(*position) == Some(&b'-'))
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || byte.is_ascii_hexdigit()
        });
    if !valid {
        return Err(invalid(format!(
            "invalid custom property format ID '{format_id}'"
        )));
    }
    let normalized = format_id.to_ascii_uppercase();
    if normalized == SUMMARY_FORMAT_ID || normalized == DOCUMENT_SUMMARY_FORMAT_ID {
        return Err(invalid(format!(
            "format ID {normalized} is forbidden for custom properties"
        )));
    }
    Ok(normalized)
}

fn fold_name(name: &str) -> String {
    name.chars().nfd().default_case_fold().nfd().collect()
}

fn value_text_bytes(value: &Value) -> usize {
    match value {
        Value::Text(text) => text.len(),
        _ => 0,
    }
}

fn append_value_text(buffer: &mut String, text: &str) -> Result<()> {
    let actual = buffer
        .len()
        .checked_add(text.len())
        .ok_or_else(|| limit("custom-property text bytes", MAX_TEXT_BYTES, usize::MAX))?;
    if actual > MAX_TEXT_BYTES {
        return Err(limit("custom-property text bytes", MAX_TEXT_BYTES, actual));
    }
    buffer.push_str(text);
    Ok(())
}

fn decode_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| Error::Xml(format!("invalid character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::Xml(format!("invalid entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_owned()),
        "lt" => Ok("<".to_owned()),
        "gt" => Ok(">".to_owned()),
        "quot" => Ok("\"".to_owned()),
        "apos" => Ok("'".to_owned()),
        _ => Err(invalid(format!("unsupported entity reference '&{name};'"))),
    }
}

fn is_name(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    element.local_name().as_ref() == expected_local_name
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected_namespace)
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn display_name(name: &[u8]) -> String {
    String::from_utf8_lossy(name).into_owned()
}

fn checked_depth(depth: usize) -> Result<usize> {
    let actual = checked_increment(depth, "custom-properties XML depth")?;
    if actual > MAX_XML_DEPTH {
        return Err(limit("custom-properties XML depth", MAX_XML_DEPTH, actual));
    }
    Ok(actual)
}

fn count_node(nodes: &mut usize) -> Result<()> {
    let actual = checked_increment(*nodes, "custom-properties XML nodes")?;
    if actual > MAX_XML_NODES {
        return Err(limit("custom-properties XML nodes", MAX_XML_NODES, actual));
    }
    *nodes = actual;
    Ok(())
}

fn checked_increment(value: usize, resource: &'static str) -> Result<usize> {
    value
        .checked_add(1)
        .ok_or_else(|| limit(resource, usize::MAX, usize::MAX))
}

fn checked_total(
    current: usize,
    added: usize,
    max: usize,
    resource: &'static str,
) -> Result<usize> {
    let actual = current
        .checked_add(added)
        .ok_or_else(|| limit(resource, max, usize::MAX))?;
    if actual > max {
        return Err(limit(resource, max, actual));
    }
    Ok(actual)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, max: usize, actual: usize) -> Error {
    Error::Limit {
        resource,
        max,
        actual,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::rel::TargetMode;
    use std::sync::Arc;

    const PREFIX: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" "#,
        r#"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"#
    );

    fn property(pid: i32, name: &str, value: &str) -> String {
        format!(r#"<property fmtid="{FORMAT_ID}" pid="{pid}" name="{name}">{value}</property>"#)
    }

    fn document(body: &str) -> String {
        format!("{PREFIX}{body}</Properties>")
    }

    #[test]
    fn concise_crud_moves_values_and_orders_names() {
        let mut props = Props::new();
        assert_eq!(props.insert("Version", 1_i32).expect("insert"), None);
        props.insert("Author", "Ada").expect("insert");
        assert_eq!(props.names().collect::<Vec<_>>(), ["Author", "Version"]);
        assert_eq!(props.get("VERSION"), Some(&Value::I32(1)));
        assert!(props.contains("author"));
        assert_eq!(
            props.insert("Version", 2_i32).expect("replace"),
            Some(Value::I32(1))
        );
        assert_eq!(props.remove("vErSiOn"), Some(Value::I32(2)));
        assert!(!props.contains("Version"));
        props.clear();
        assert!(props.is_empty());
        props.insert("AfterClear", true).expect("PID reset insert");
        let xml = encode(&props).expect("encode");
        assert!(String::from_utf8_lossy(&xml).contains(r#"pid="2""#));
    }

    #[test]
    fn case_insensitive_duplicate_names_are_rejected_on_insert_and_read() {
        let mut props = Props::new();
        props.insert("Project", "one").expect("first insert");
        assert!(props.insert("project", "two").is_err());

        let xml = document(&format!(
            "{}{}",
            property(2, "Project", "<vt:lpwstr>one</vt:lpwstr>"),
            property(3, "PROJECT", "<vt:lpwstr>two</vt:lpwstr>")
        ));
        assert!(decode(xml.as_bytes()).is_err());
    }

    #[test]
    fn names_use_canonical_unicode_caseless_identity() {
        let mut props = Props::new();
        props.insert("Straße", "stored spelling").expect("insert");
        assert_eq!(
            props.get("STRASSE"),
            Some(&Value::Text("stored spelling".to_owned()))
        );
        assert!(props.contains("strasse"));
        assert!(props.insert("STRASSE", "duplicate").is_err());
        assert_eq!(
            props.remove("STRASSE"),
            Some(Value::Text("stored spelling".to_owned()))
        );
        assert!(props.is_empty());
    }

    #[test]
    fn all_supported_values_round_trip_deterministically_in_pid_order() {
        let xml = document(&format!(
            "{}{}{}{}{}{}{}{}{}",
            property(9, "Last", "<vt:lpstr>narrow &amp; exact</vt:lpstr>"),
            property(2, "Empty", "<vt:empty/>"),
            property(3, "Text", "<vt:lpwstr>wide</vt:lpwstr>"),
            property(4, "I32", "<vt:i4>-7</vt:i4>"),
            property(5, "I64", "<vt:i8>9000000000</vt:i8>"),
            property(6, "F32", "<vt:r4>1.25</vt:r4>"),
            property(7, "F64", "<vt:r8>2.5</vt:r8>"),
            property(8, "Bool", "<vt:bool>1</vt:bool>"),
            property(
                10,
                "Time",
                "<vt:filetime>2024-05-06T07:08:09.123Z</vt:filetime>"
            )
        ));
        let props = decode(xml.as_bytes()).expect("decode");
        assert_eq!(props.get("Empty"), Some(&Value::Empty));
        assert_eq!(props.get("I64"), Some(&Value::I64(9_000_000_000)));
        assert_eq!(props.get("F32"), Some(&Value::F32(1.25)));
        let first = encode(&props).expect("first encode");
        let second = encode(&props).expect("second encode");
        assert_eq!(first, second);
        let output = String::from_utf8(first.clone()).expect("UTF-8 XML");
        assert!(output.contains("<vt:lpstr>narrow &amp; exact</vt:lpstr>"));
        assert!(output.find(r#"pid="2""#) < output.find(r#"pid="9""#));
        assert_eq!(
            encode(&decode(&first).expect("round-trip decode")).expect("encode"),
            first
        );
    }

    #[test]
    fn filetime_is_rfc3339_not_a_numeric_windows_counter() {
        let valid = document(&property(
            2,
            "When",
            "<vt:filetime>2020-01-02T03:04:05+02:30</vt:filetime>",
        ));
        let props = decode(valid.as_bytes()).expect("RFC3339 filetime");
        let output = String::from_utf8(encode(&props).expect("encode")).expect("UTF-8");
        assert!(output.contains("2020-01-02T00:34:05Z"));

        let numeric = document(&property(
            2,
            "When",
            "<vt:filetime>132223104000000000</vt:filetime>",
        ));
        assert!(decode(numeric.as_bytes()).is_err());
    }

    #[test]
    fn non_finite_floats_are_rejected_on_insert_and_read() {
        let mut props = Props::new();
        assert!(props.insert("NaN", f64::NAN).is_err());
        assert!(props.insert("Infinity", f32::INFINITY).is_err());
        let xml = document(&property(2, "NaN", "<vt:r8>NaN</vt:r8>"));
        assert!(decode(xml.as_bytes()).is_err());
    }

    #[test]
    fn root_namespace_and_property_cardinality_are_strict() {
        let wrong_namespace = r#"<Properties xmlns="urn:wrong"/>"#;
        assert!(decode(wrong_namespace.as_bytes()).is_err());
        let missing = document(&property(2, "Missing", ""));
        assert!(decode(missing.as_bytes()).is_err());
        let duplicate = document(&property(
            2,
            "Duplicate",
            "<vt:i4>1</vt:i4><vt:i4>2</vt:i4>",
        ));
        assert!(decode(duplicate.as_bytes()).is_err());
        let wrong_value_namespace = document(&property(
            2,
            "Wrong",
            r#"<x:i4 xmlns:x="urn:wrong">1</x:i4>"#,
        ));
        assert!(decode(wrong_value_namespace.as_bytes()).is_err());
    }

    #[test]
    fn malformed_and_duplicate_pids_are_rejected() {
        let below_minimum = document(&property(1, "Low", "<vt:i4>1</vt:i4>"));
        assert!(decode(below_minimum.as_bytes()).is_err());
        let duplicate = document(&format!(
            "{}{}",
            property(2, "One", "<vt:i4>1</vt:i4>"),
            property(2, "Two", "<vt:i4>2</vt:i4>")
        ));
        assert!(decode(duplicate.as_bytes()).is_err());
        let malformed = document(&property(2, "Bad", "<vt:i4>not-an-int</vt:i4>"));
        assert!(decode(malformed.as_bytes()).is_err());
    }

    #[test]
    fn exhausted_pid_space_allows_replacement_but_not_allocation() {
        let xml = document(&property(i32::MAX, "Last", "<vt:lpwstr>value</vt:lpwstr>"));
        let mut props = decode(xml.as_bytes()).expect("maximum PID is valid");
        assert_eq!(
            props.insert("Last", "replacement").expect("replace"),
            Some(Value::Text("value".to_owned()))
        );
        assert!(props.insert("New", "cannot allocate").is_err());
    }

    #[test]
    fn forbidden_and_malformed_format_ids_are_rejected() {
        for format_id in [SUMMARY_FORMAT_ID, DOCUMENT_SUMMARY_FORMAT_ID, "not-a-guid"] {
            let xml = document(&format!(
                r#"<property fmtid="{format_id}" pid="2" name="Bad"><vt:i4>1</vt:i4></property>"#
            ));
            assert!(decode(xml.as_bytes()).is_err());
        }
    }

    #[test]
    fn dtd_unknown_entities_and_malformed_xml_are_rejected() {
        let dtd = format!(
            r#"<!DOCTYPE Properties [<!ENTITY x "expanded">]>{PREFIX}{} </Properties>"#,
            property(2, "X", "<vt:lpwstr>&x;</vt:lpwstr>")
        );
        assert!(decode(dtd.as_bytes()).is_err());
        let unknown = document(&property(2, "X", "<vt:lpwstr>&unknown;</vt:lpwstr>"));
        assert!(decode(unknown.as_bytes()).is_err());
        assert!(decode(b"<Properties><property></Properties>").is_err());
    }

    #[test]
    fn byte_depth_node_name_and_text_limits_are_enforced() {
        let oversized = vec![b' '; MAX_XML_BYTES + 1];
        assert!(matches!(decode(&oversized), Err(Error::Limit { .. })));

        let deep = document(&property(
            2,
            "Deep",
            r#"<vt:lpwstr><x xmlns="urn:x"/></vt:lpwstr>"#,
        ));
        assert!(matches!(decode(deep.as_bytes()), Err(Error::Limit { .. })));

        let comments = "<!--x-->".repeat(MAX_XML_NODES + 1);
        let noisy = document(&comments);
        assert!(matches!(decode(noisy.as_bytes()), Err(Error::Limit { .. })));

        let mut props = Props::new();
        assert!(matches!(
            props.insert("n".repeat(MAX_NAME_CHARS + 1), "x"),
            Err(Error::Limit { .. })
        ));
        assert!(matches!(
            props.insert("Text", "x".repeat(MAX_TEXT_BYTES + 1)),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn property_count_is_bounded() {
        let mut body = String::new();
        for index in 0..=MAX_PROPERTIES {
            body.push_str(&property(
                i32::try_from(index).expect("test PID") + 2,
                &format!("P{index}"),
                "<vt:empty/>",
            ));
        }
        assert!(matches!(
            decode(document(&body).as_bytes()),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn absent_package_properties_are_empty_but_orphans_are_errors() {
        let package = OpcPackage::new();
        assert!(Props::read(&package).expect("absence is valid").is_empty());

        let mut orphan = OpcPackage::new();
        orphan.add_part(Box::new(BlobPart::new(
            custom_part_name().expect("URI"),
            ct::OFC_CUSTOM_PROPERTIES.to_owned(),
            document("").into_bytes(),
        )));
        assert!(matches!(Props::read(&orphan), Err(Error::Relationship(_))));
    }

    #[test]
    fn package_write_read_and_clear_remove_the_complete_graph() {
        let mut package = OpcPackage::new();
        let mut props = Props::new();
        props.insert("Project", "Litchi").expect("insert");
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(package.is_signed());
        props.write(&mut package).expect("write");
        assert!(!package.is_signed());
        let first_blob = package
            .get_part(&custom_part_name().expect("URI"))
            .expect("custom part")
            .blob_arc();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(package.is_signed());
        props.write(&mut package).expect("byte-identical no-op");
        assert!(package.is_signed(), "a true no-op preserves signatures");
        let second_blob = package
            .get_part(&custom_part_name().expect("URI"))
            .expect("custom part")
            .blob_arc();
        assert!(Arc::ptr_eq(&first_blob, &second_blob));
        assert_eq!(
            Props::read(&package).expect("read").get("Project"),
            props.get("Project")
        );
        assert!(package.get_part(&custom_part_name().expect("URI")).is_ok());
        assert_eq!(
            package
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == rt::CUSTOM_PROPERTIES)
                .count(),
            1
        );

        Props::new().write(&mut package).expect("clear graph");
        assert!(!package.is_signed());
        assert!(package.get_part(&custom_part_name().expect("URI")).is_err());
        assert!(
            package
                .rels()
                .iter()
                .all(|relationship| relationship.reltype() != rt::CUSTOM_PROPERTIES)
        );
        assert!(Props::read(&package).expect("read cleared").is_empty());
    }

    #[test]
    fn malformed_package_graphs_are_never_treated_as_absent() {
        let mut missing = OpcPackage::new();
        missing.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
        assert!(matches!(Props::read(&missing), Err(Error::Missing(_))));

        let mut wrong_type = OpcPackage::new();
        wrong_type.add_part(Box::new(BlobPart::new(
            custom_part_name().expect("URI"),
            "application/xml".to_owned(),
            document("").into_bytes(),
        )));
        wrong_type.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
        assert!(matches!(
            Props::read(&wrong_type),
            Err(Error::ContentType { .. })
        ));

        let mut duplicate = OpcPackage::new();
        duplicate.add_part(Box::new(BlobPart::new(
            custom_part_name().expect("URI"),
            ct::OFC_CUSTOM_PROPERTIES.to_owned(),
            document("").into_bytes(),
        )));
        duplicate.relate_to(PART_TARGET, rt::CUSTOM_PROPERTIES);
        duplicate
            .rels_mut()
            .try_add_relationship(
                rt::CUSTOM_PROPERTIES.to_owned(),
                PART_TARGET.to_owned(),
                "rId2".to_owned(),
                TargetMode::Internal,
            )
            .expect("second relationship");
        assert!(matches!(
            Props::read(&duplicate),
            Err(Error::Relationship(_))
        ));

        let mut external = OpcPackage::new();
        external.relate_to_external("https://example.invalid/custom.xml", rt::CUSTOM_PROPERTIES);
        assert!(matches!(
            Props::read(&external),
            Err(Error::Relationship(_))
        ));
    }

    #[test]
    fn illegal_xml_characters_are_rejected_before_writing() {
        let mut props = Props::new();
        assert!(props.insert("Bad\0Name", "value").is_err());
        assert!(props.insert("BadText", "value\u{1}").is_err());
    }
}
