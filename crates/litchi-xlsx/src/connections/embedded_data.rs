//! Source-bound x14 connection references to Custom Data storage IDs.
//!
//! MS-XLSX 2.4.21, 2.6.34 and 2.2.4.1. Only embeddedDataId values are
//! rewritten; queries, credentials, unknown XML and connection identities
//! remain source bytes. No connection is executed or refreshed.

use std::borrow::Cow;
use std::collections::HashSet;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::Resource;
use litchi_opc::{OpcPackage, OwnedRelationships, PackURI, Part, ReadLimits, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::model::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP, CORE_NAMESPACE, MAX_CONNECTIONS,
    STRICT_CONNECTIONS_RELATIONSHIP, STRICT_NAMESPACE,
};
use crate::custom_data::Limits as CustomDataLimits;
use crate::error::{Error, Result};
use crate::source_attributes::{
    escaped_xstring_len, try_escaped_xstring, validate_xml_characters, value_span,
};

const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_NODES: usize = 200_000;
const MAX_DEPTH: usize = 128;

fn invalid(value: impl std::fmt::Display) -> Error {
    crate::error::invalid(value.to_string())
}

fn limit(resource: Resource, name: &'static str, observed: usize, maximum: usize) -> Error {
    Error::ResourceLimit(litchi_core::ResourceLimit {
        resource,
        observed: observed as u64,
        limit: maximum as u64,
        scope: Arc::<str>::from(format!("XLSX Custom Data {name}")),
    })
}

#[derive(Clone, PartialEq, Eq)]
struct Edge {
    source: String,
    id: String,
    kind: String,
    target: String,
    mode: TargetMode,
}

#[derive(Clone, PartialEq, Eq)]
struct Slot {
    connection_id: u32,
    range: Range<usize>,
}

#[derive(PartialEq, Eq)]
struct Source {
    name: PackURI,
    bytes: Arc<Vec<u8>>,
    proof: litchi_opc::OwnedXmlPart,
    relationships: OwnedRelationships,
    edges: Vec<Edge>,
    slots: Vec<Slot>,
    values: Arc<[String]>,
    unmodeled_context: bool,
}

/// A cheap immutable source plus staged values; XML is emitted once at commit.
#[derive(Clone, Default)]
pub(crate) struct Bindings {
    source: Option<Arc<Source>>,
    values: Arc<[String]>,
}

impl std::fmt::Debug for Bindings {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("EmbeddedDataBindings")
            .field("references", &self.values.len())
            .finish_non_exhaustive()
    }
}

impl PartialEq for Bindings {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source && self.values == other.values
    }
}
impl Eq for Bindings {}

impl Bindings {
    pub(crate) fn load_with_limits(
        package: &OpcPackage,
        limits: CustomDataLimits,
        read_limits: ReadLimits,
    ) -> Result<Self> {
        // Admit the XML projection before the complete typed graph walk. The
        // latter has its own schema parser and therefore must not be allowed
        // to materialize a connections DOM when this owner has already
        // received a tighter caller profile.
        for part in package
            .iter_parts()
            .filter(|part| part.content_type() == CONNECTIONS_CONTENT_TYPE)
        {
            // Only a connections part's payload is read, so only it is
            // decoded (ADR 0030).
            parse_with_limits(package.get_part(part.partname())?.blob(), limits)?;
        }
        // Validate the complete workbook connection/query-table graph before
        // the owner/part projection below. This must also reject malformed
        // graphs when no connections content-type part is present, rather
        // than letting the `(None, None)` fast path admit a root edge or an
        // orphan query table.
        super::package::validate_graph_with_limits(package, &limits)?;
        let workbook = package.main_document_part()?;
        let mut owners = workbook.rels().iter().filter(|r| {
            matches!(
                r.reltype(),
                CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
            )
        });
        let owner = owners.next();
        if owners.next().is_some() {
            return Err(invalid("multiple workbook connections owners"));
        }
        let mut parts = package
            .iter_parts()
            .filter(|p| p.content_type() == CONNECTIONS_CONTENT_TYPE);
        let part = parts.next();
        if parts.next().is_some() {
            return Err(invalid("multiple connections parts"));
        }
        let (owner, part) = match (owner, part) {
            (None, None) => return Ok(Self::default()),
            (Some(owner), Some(part)) => (owner, part),
            _ => return Err(invalid("connections part and workbook owner do not agree")),
        };
        if owner.is_external()
            || owner.target_query().is_some()
            || owner.target_fragment().is_some()
            || !owner.target_partname()?.is_equivalent_to(part.partname())
        {
            return Err(invalid(
                "connections owner must target the internal connections part",
            ));
        }
        // The connections payload is read below, so it is decoded here
        // (ADR 0030).
        let part = package.get_part(part.partname())?;
        // The lightweight scanner below only projects embeddedDataId slots;
        // the complete graph and typed connection catalog were validated
        // above. Keep this as a second pass: it retains source locations and
        // deliberately refuses unmodeled MCE contexts.
        if part.blob().len() > limits.max_connections_bytes() {
            return Err(limit(
                Resource::InputBytes,
                "connections XML bytes",
                part.blob().len(),
                limits.max_connections_bytes(),
            ));
        }
        let (slots, values, unmodeled_context) = parse_with_limits(part.blob(), limits)?;
        let values = Arc::<[String]>::from(values);
        let relationships =
            package.source_relationships_with_limits(part.partname(), read_limits)?;
        if relationships.bytes().len() > limits.max_relationship_xml_bytes() {
            return Err(limit(
                Resource::InputBytes,
                "connections relationship XML bytes",
                relationships.bytes().len(),
                limits.max_relationship_xml_bytes(),
            ));
        }
        let source = Source {
            name: part.partname().clone(),
            bytes: part.blob_arc(),
            proof: package.source_xml_part(part.partname())?,
            relationships,
            edges: capture_edges(package, part, limits)?,
            slots,
            values: Arc::clone(&values),
            unmodeled_context,
        };
        Ok(Self {
            source: Some(Arc::new(source)),
            values,
        })
    }

    pub(crate) fn validate_ids(&self, ids: &HashSet<&str>) -> Result<()> {
        for id in self.values.iter().filter(|value| !value.is_empty()) {
            if !ids.contains(id.as_str()) {
                return Err(invalid(
                    "connection embeddedDataId has no matching Custom Data storage",
                ));
            }
        }
        Ok(())
    }

    pub(crate) fn count(&self, id: &str) -> usize {
        if id.is_empty() {
            return 0;
        }
        self.values
            .iter()
            .filter(|value| value.as_str() == id)
            .count()
    }

    #[cfg(test)]
    pub(crate) fn rebind(&self, from: &str, to: &str) -> Result<Self> {
        self.rebind_with_limits(from, to, CustomDataLimits::default())
    }

    pub(crate) fn rebind_with_limits(
        &self,
        from: &str,
        to: &str,
        limits: CustomDataLimits,
    ) -> Result<Self> {
        if from == to {
            return Ok(self.clone());
        }
        if self
            .source
            .as_ref()
            .is_some_and(|source| source.unmodeled_context)
        {
            return Err(Error::Unsupported {
                feature: "Custom Data identity changes with unmodeled connection markup compatibility",
            });
        }
        if self.count(from) == 0 {
            return Ok(self.clone());
        }
        // Charge replacement multiplicity before cloning or growing any values.
        let size = self
            .values
            .iter()
            .try_fold(0usize, |total, value| {
                total.checked_add(if value == from { to.len() } else { value.len() })
            })
            .ok_or_else(|| invalid("embeddedDataId aggregate size overflow"))?;
        if size > limits.max_connections_bytes() {
            return Err(limit(
                Resource::OutputBytes,
                "connections XML bytes",
                size,
                limits.max_connections_bytes(),
            ));
        }
        if let Some(source) = self.source.as_ref() {
            if source.slots.len() != source.values.len() || source.slots.len() != self.values.len()
            {
                return Err(invalid("connections source binding slot count changed"));
            }
            let mut output_bytes = source.bytes.len();
            for ((slot, before), after) in source
                .slots
                .iter()
                .zip(source.values.iter())
                .zip(self.values.iter())
            {
                if before == after {
                    continue;
                }
                let replacement = escaped_xstring_len(after)?;
                output_bytes = output_bytes
                    .checked_sub(slot.range.len())
                    .and_then(|value| value.checked_add(replacement))
                    .ok_or_else(|| invalid("connections XML output size overflow"))?;
            }
            let maximum = limits.max_connections_bytes();
            if output_bytes > maximum {
                return Err(limit(
                    Resource::OutputBytes,
                    "connections output bytes",
                    output_bytes,
                    maximum,
                ));
            }
        }
        let mut values = Vec::new();
        values
            .try_reserve(self.values.len())
            .map_err(|source| Error::Allocation {
                resource: "Custom Data connection values",
                source,
            })?;
        for value in self.values.iter() {
            values.push(if value == from {
                to.to_owned()
            } else {
                value.clone()
            });
        }
        Ok(Self {
            source: self.source.clone(),
            values: Arc::from(values),
        })
    }

    pub(crate) fn is_changed(&self) -> bool {
        self.source
            .as_ref()
            .is_some_and(|source| source.values != self.values)
    }

    pub(crate) fn source_bytes_len(&self) -> Option<usize> {
        self.source.as_ref().map(|source| source.bytes.len())
    }

    pub(crate) fn source_part_name(&self) -> Option<&PackURI> {
        self.source.as_ref().map(|source| &source.name)
    }

    pub(crate) fn source_relationships_bytes_len(&self) -> Option<usize> {
        self.source
            .as_ref()
            .map(|source| source.relationships.bytes().len())
    }

    /// Exact source-member byte length after the currently staged attribute
    /// replacements, computed without allocating the replacement XML buffer.
    pub(crate) fn projected_output_bytes(&self, limits: CustomDataLimits) -> Result<Option<usize>> {
        let Some(source) = &self.source else {
            return Ok(None);
        };
        if !self.is_changed() {
            return Ok(Some(source.bytes.len()));
        }
        if source.slots.len() != source.values.len() || source.slots.len() != self.values.len() {
            return Err(invalid("connections source binding slot count changed"));
        }
        let mut removed = 0usize;
        let mut replacement = 0usize;
        for ((slot, before), after) in source
            .slots
            .iter()
            .zip(source.values.iter())
            .zip(self.values.iter())
        {
            if before == after {
                continue;
            }
            removed = removed
                .checked_add(slot.range.len())
                .ok_or_else(|| invalid("connections XML replacement size overflow"))?;
            replacement = replacement
                .checked_add(escaped_xstring_len(after)?)
                .ok_or_else(|| invalid("connections XML replacement size overflow"))?;
        }
        let output = source
            .bytes
            .len()
            .checked_sub(removed)
            .and_then(|size| size.checked_add(replacement))
            .ok_or_else(|| invalid("connections XML output size overflow"))?;
        if output > limits.max_connections_bytes() {
            return Err(limit(
                Resource::OutputBytes,
                "connections output bytes",
                output,
                limits.max_connections_bytes(),
            ));
        }
        Ok(Some(output))
    }

    pub(crate) fn same_values(&self, other: &Self) -> bool {
        self.values == other.values
            && self.source.as_ref().map(|s| &s.name) == other.source.as_ref().map(|s| &s.name)
    }

    pub(crate) fn publish_with_limits(
        &self,
        package: &mut OpcPackage,
        limits: CustomDataLimits,
        output_limit: usize,
    ) -> Result<()> {
        let Some(source) = &self.source else {
            return Ok(());
        };
        if !self.is_changed() {
            return Ok(());
        }
        let mut edits = Vec::new();
        edits
            .try_reserve(source.slots.len())
            .map_err(|source| Error::Allocation {
                resource: "Custom Data connection edits",
                source,
            })?;
        let mut replacement_bytes = 0usize;
        for ((slot, before), after) in source
            .slots
            .iter()
            .zip(source.values.iter())
            .zip(self.values.iter())
        {
            if before != after {
                let escaped_len = escaped_xstring_len(after)?;
                replacement_bytes = replacement_bytes
                    .checked_add(escaped_len)
                    .ok_or_else(|| invalid("embeddedDataId replacement size overflow"))?;
                if replacement_bytes > limits.max_connections_bytes() {
                    return Err(limit(
                        Resource::OutputBytes,
                        "connections XML bytes",
                        replacement_bytes,
                        limits.max_connections_bytes(),
                    ));
                }
                if escaped_len > limits.max_temporary_bytes() {
                    return Err(limit(
                        Resource::Memory,
                        "temporary connection attribute bytes",
                        escaped_len,
                        limits.max_temporary_bytes(),
                    ));
                }
                let escaped = try_escaped_xstring(after)?;
                edits.push((slot.range.clone(), escaped));
            }
        }
        let removed = edits
            .iter()
            .try_fold(0usize, |total, (range, _)| total.checked_add(range.len()))
            .ok_or_else(|| invalid("connections XML replacement size overflow"))?;
        let output_bytes = source
            .bytes
            .len()
            .checked_sub(removed)
            .and_then(|size| size.checked_add(replacement_bytes))
            .ok_or_else(|| invalid("connections XML output size overflow"))?;
        let maximum = limits
            .max_connections_bytes()
            .min(output_limit)
            .min(limits.max_temporary_bytes());
        if output_bytes > maximum {
            return Err(limit(
                Resource::OutputBytes,
                "connections output bytes",
                output_bytes,
                maximum,
            ));
        }
        let token = source.proof.replace_attributes(&edits)?;
        package.try_replace_owned_xml_part(&source.bytes, token)?;
        Ok(())
    }

    pub(crate) fn restore_with_limits(
        &self,
        package: &mut OpcPackage,
        limits: CustomDataLimits,
        read_limits: ReadLimits,
    ) -> Result<()> {
        if let Some(source) = &self.source {
            if source.bytes.len() > limits.max_connections_bytes() {
                return Err(limit(
                    Resource::InputBytes,
                    "connections XML bytes",
                    source.bytes.len(),
                    limits.max_connections_bytes(),
                ));
            }
            let current = package.get_part(&source.name)?.blob_arc();
            package.try_replace_owned_xml_part(&current, source.proof.clone())?;
            let current_relationships =
                package.source_relationships_with_limits(&source.name, read_limits)?;
            package.try_replace_relationships_with_limits(
                &current_relationships,
                &source.relationships,
                read_limits,
            )?;
        }
        Ok(())
    }
}

fn capture_edges(
    package: &OpcPackage,
    target: &dyn Part,
    limits: CustomDataLimits,
) -> Result<Vec<Edge>> {
    let mut edges = Vec::new();
    for (source, relationships) in std::iter::once(("/", package.rels())).chain(
        package
            .iter_parts()
            .map(|part| (part.partname().as_str(), part.rels())),
    ) {
        let source_uri = PackURI::new(source).map_err(invalid)?;
        for relationship in relationships.iter() {
            if source_uri.is_equivalent_to(target.partname())
                || !relationship.is_external()
                    && relationship
                        .target_partname()?
                        .is_equivalent_to(target.partname())
            {
                if edges.len() >= MAX_CONNECTIONS.min(limits.max_relationships()) {
                    return Err(limit(
                        Resource::Objects,
                        "connections relationships",
                        edges.len().saturating_add(1),
                        MAX_CONNECTIONS.min(limits.max_relationships()),
                    ));
                }
                edges.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "Custom Data connection relationships",
                    source,
                })?;
                edges.push(Edge {
                    source: source.into(),
                    id: relationship.r_id().into(),
                    kind: relationship.reltype().into(),
                    target: relationship.target_ref().into(),
                    mode: relationship.target_mode(),
                });
            }
        }
    }
    edges.sort_by(|left, right| {
        left.source
            .cmp(&right.source)
            .then_with(|| left.id.cmp(&right.id))
    });
    Ok(edges)
}

#[derive(Clone, Copy)]
enum Context {
    Root,
    Connection(u32, Option<u32>),
    List,
    Extension(bool),
    Other,
}

fn namespace(value: ResolveResult<'_>) -> Result<Cow<'_, str>> {
    match value {
        ResolveResult::Bound(value) => {
            let value = std::str::from_utf8(value.0).map_err(invalid)?;
            quick_xml::escape::unescape(value).map_err(invalid)
        },
        ResolveResult::Unbound => Ok(Cow::Borrowed("")),
        ResolveResult::Unknown(_) => Err(invalid("unbound prefix in connections XML")),
    }
}

#[cfg(test)]
fn parse(xml: &[u8]) -> Result<(Vec<Slot>, Vec<String>, bool)> {
    parse_with_limits(xml, CustomDataLimits::default())
}

fn parse_with_limits(
    xml: &[u8],
    limits: CustomDataLimits,
) -> Result<(Vec<Slot>, Vec<String>, bool)> {
    if xml.len() > limits.max_connections_bytes() {
        return Err(limit(
            Resource::InputBytes,
            "connections XML bytes",
            xml.len(),
            limits.max_connections_bytes(),
        ));
    }
    validate_xml_characters(xml)?;
    super::codec::preflight_xml_with_limits(
        xml,
        limits.max_xml_nodes(),
        limits.max_xml_events(),
        limits.max_xml_depth(),
        limits.max_xml_string_bytes(),
        limits.max_xml_namespace_bytes(),
        limits.max_xml_attributes(),
    )?;
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut namespace_stack = Vec::new();
    let mut root_seen = false;
    let mut core = String::new();
    let mut nodes = 0usize;
    let mut ids = HashSet::new();
    let mut extended = HashSet::new();
    let mut slots = Vec::new();
    let mut values = Vec::new();
    let mut unmodeled = false;
    let mut events = 0usize;
    let mut namespace_bytes = 0usize;
    loop {
        events = events.saturating_add(1);
        if events > limits.max_xml_events() {
            return Err(limit(
                Resource::Objects,
                "connections XML events",
                events,
                limits.max_xml_events(),
            ));
        }
        let event = reader.read_event().map_err(invalid)?;
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("connections XML node count overflow"))?;
                if nodes > MAX_NODES.min(limits.max_xml_nodes()) {
                    return Err(limit(
                        Resource::Objects,
                        "connections XML nodes",
                        nodes,
                        MAX_NODES.min(limits.max_xml_nodes()),
                    ));
                }
                if stack.len() >= MAX_DEPTH.min(limits.max_xml_depth()) {
                    return Err(limit(
                        Resource::Depth,
                        "connections XML depth",
                        stack.len().saturating_add(1),
                        MAX_DEPTH.min(limits.max_xml_depth()),
                    ));
                }
                let ns = namespace(reader.resolver().resolve_element(element.name()).0)?;
                let name = element.local_name();
                let name = name.as_ref();
                unmodeled |= ns == MCE;
                let mut id = None;
                let mut kind = None;
                let mut uri = None;
                let mut embedded = None;
                let mut attributes = HashSet::new();
                let mut attribute_count = 0usize;
                let inherited_namespace_bytes = namespace_bytes;
                let mut local_namespace_bytes = 0usize;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(invalid)?;
                    attribute_count = attribute_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("connections XML attribute count overflow"))?;
                    if attribute_count > limits.max_xml_attributes() {
                        return Err(limit(
                            Resource::Objects,
                            "connections XML attributes",
                            attribute_count,
                            limits.max_xml_attributes(),
                        ));
                    }
                    if attribute.value.as_ref().len() > limits.max_xml_string_bytes() {
                        return Err(limit(
                            Resource::InputBytes,
                            "connections XML string bytes",
                            attribute.value.as_ref().len(),
                            limits.max_xml_string_bytes(),
                        ));
                    }
                    if attribute.key.as_ref() == b"xmlns"
                        || attribute.key.as_ref().starts_with(b"xmlns:")
                    {
                        let prefix_bytes = attribute
                            .key
                            .as_ref()
                            .strip_prefix(b"xmlns:")
                            .map_or(0, |prefix| prefix.len());
                        local_namespace_bytes = local_namespace_bytes
                            .checked_add(prefix_bytes)
                            .and_then(|bytes| bytes.checked_add(attribute.value.as_ref().len()))
                            .ok_or_else(|| {
                                invalid("connections XML namespace byte count overflow")
                            })?;
                        namespace_bytes = inherited_namespace_bytes
                            .checked_add(local_namespace_bytes)
                            .ok_or_else(|| {
                                invalid("connections XML namespace byte count overflow")
                            })?;
                        if namespace_bytes > limits.max_xml_namespace_bytes() {
                            return Err(limit(
                                Resource::InputBytes,
                                "connections XML namespace bytes",
                                namespace_bytes,
                                limits.max_xml_namespace_bytes(),
                            ));
                        }
                    }
                    if attribute.key.as_ref() == b"xmlns"
                        || attribute.key.as_ref().starts_with(b"xmlns:")
                    {
                        continue;
                    }
                    let (attribute_ns, local) = reader.resolver().resolve_attribute(attribute.key);
                    let attribute_ns = namespace(attribute_ns)?;
                    let namespaced = !attribute_ns.is_empty();
                    attributes
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "Custom Data connection attributes",
                            source,
                        })?;
                    if !attributes.insert((attribute_ns.into_owned(), local.as_ref().to_vec())) {
                        return Err(invalid("duplicate expanded attribute in connections XML"));
                    }
                    // Directives do not change a recognized owner's attribute
                    // semantics. AlternateContent and unrecognized owner paths
                    // are tracked separately and cannot be rewritten.
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(invalid)?;
                    if value.len() > limits.max_xml_string_bytes() {
                        return Err(limit(
                            Resource::InputBytes,
                            "connections XML string bytes",
                            value.len(),
                            limits.max_xml_string_bytes(),
                        ));
                    }
                    validate_xml_characters(value.as_bytes())?;
                    if namespaced {
                        continue;
                    }
                    match local.as_ref() {
                        b"id" => id = Some(value.into_owned()),
                        b"type" => kind = Some(value.into_owned()),
                        b"uri" => uri = Some(value.into_owned()),
                        b"embeddedDataId" if ns == X14 && name == b"connection" => {
                            let value = crate::raw::strings::decode_spreadsheet_text(&value)?;
                            let units = value.encode_utf16().count();
                            if units > limits.max_uid_units() {
                                return Err(limit(
                                    Resource::InputBytes,
                                    "UID UTF-16 units",
                                    units,
                                    limits.max_uid_units(),
                                ));
                            }
                            embedded = Some((value_span(xml, attribute.value.as_ref())?, value));
                        },
                        b"culture" if ns == X14 && name == b"connection" => {
                            if crate::raw::strings::decode_spreadsheet_text(&value)?
                                .chars()
                                .count()
                                >= 85
                            {
                                return Err(invalid("connection culture is too long"));
                            }
                        },
                        _ => {},
                    }
                }
                let context = if stack.is_empty() {
                    if root_seen
                        || name != b"connections"
                        || (ns.as_ref() != CORE_NAMESPACE && ns.as_ref() != STRICT_NAMESPACE)
                    {
                        return Err(invalid("expected one SpreadsheetML connections root"));
                    }
                    root_seen = true;
                    core = ns.as_ref().to_owned();
                    Context::Root
                } else if ns == core
                    && name == b"connection"
                    && matches!(stack.last(), Some(Context::Root))
                {
                    let id = id
                        .ok_or_else(|| invalid("connection ID is absent"))?
                        .parse::<u32>()
                        .map_err(invalid)?;
                    if ids.len() >= MAX_CONNECTIONS {
                        return Err(invalid("duplicate or excessive connection IDs"));
                    }
                    ids.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "Custom Data connection IDs",
                        source,
                    })?;
                    if !ids.insert(id) {
                        return Err(invalid("duplicate or excessive connection IDs"));
                    }
                    Context::Connection(
                        id,
                        kind.map(|v| v.parse::<u32>().map_err(invalid))
                            .transpose()?,
                    )
                } else if ns == core
                    && name == b"extLst"
                    && matches!(stack.last(), Some(Context::Connection(..)))
                {
                    Context::List
                } else if ns == core
                    && name == b"ext"
                    && matches!(stack.last(), Some(Context::List))
                {
                    Context::Extension(matches!(
                        uri.as_deref(),
                        Some(
                            "{D79990A0-CA42-45E3-83F4-45C500A0EAA5}"
                                | "{DE250136-89BD-433C-8126-D09CA5730AF9}"
                        )
                    ))
                } else {
                    Context::Other
                };
                if ns == X14 && name == b"connection" {
                    let parent = stack.iter().rev().find_map(|value| match value {
                        Context::Connection(id, kind) => Some((*id, *kind)),
                        _ => None,
                    });
                    let recognized = stack.len() == 4
                        && matches!(stack.last(), Some(Context::Extension(true)))
                        && parent.is_some();
                    if recognized {
                        let (id, kind) =
                            parent.ok_or_else(|| invalid("missing extended connection owner"))?;
                        extended
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "Custom Data extended connections",
                                source,
                            })?;
                        if kind != Some(5) || !extended.insert(id) {
                            return Err(invalid("x14 connection requires one type-5 owner"));
                        }
                    } else {
                        unmodeled = true;
                    }
                    if let Some((range, value)) = embedded {
                        if slots.len() >= MAX_CONNECTIONS {
                            return Err(invalid("embeddedDataId reference limit exceeded"));
                        }
                        slots.try_reserve(1).map_err(|source| Error::Allocation {
                            resource: "Custom Data connection slots",
                            source,
                        })?;
                        values.try_reserve(1).map_err(|source| Error::Allocation {
                            resource: "Custom Data connection values",
                            source,
                        })?;
                        slots.push(Slot {
                            connection_id: parent.map_or(0, |v| v.0),
                            range,
                        });
                        values.push(value);
                    }
                }
                if is_start {
                    stack.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "Custom Data connection nesting",
                        source,
                    })?;
                    namespace_stack
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "Custom Data connection namespace nesting",
                            source,
                        })?;
                    stack.push(context);
                    namespace_stack.push(local_namespace_bytes);
                } else {
                    namespace_bytes = inherited_namespace_bytes;
                }
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected connections closing element"))?;
                let local_namespace_bytes = namespace_stack
                    .pop()
                    .ok_or_else(|| invalid("connections namespace nesting is unbalanced"))?;
                namespace_bytes = namespace_bytes
                    .checked_sub(local_namespace_bytes)
                    .ok_or_else(|| invalid("connections namespace byte count underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DTDs are rejected in connections XML")),
            Event::Text(text) => {
                if text.as_ref().len() > limits.max_xml_string_bytes() {
                    return Err(limit(
                        Resource::InputBytes,
                        "connections XML string bytes",
                        text.as_ref().len(),
                        limits.max_xml_string_bytes(),
                    ));
                }
                if stack.is_empty() && !text.decode().map_err(invalid)?.trim().is_empty() {
                    return Err(invalid("text outside connections root"));
                }
            },
            Event::GeneralRef(reference) => {
                if reference.as_ref().len() > limits.max_xml_string_bytes() {
                    return Err(limit(
                        Resource::InputBytes,
                        "connections XML string bytes",
                        reference.as_ref().len(),
                        limits.max_xml_string_bytes(),
                    ));
                }
                let value = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
                validate_xml_characters(value.as_bytes())?;
                if stack.is_empty() {
                    return Err(invalid("entity outside connections root"));
                }
            },
            Event::CData(_) if stack.is_empty() => {
                return Err(invalid("CDATA outside connections root"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !stack.is_empty() {
        return Err(invalid("incomplete connections XML"));
    }
    Ok((slots, values, unmodeled))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn xml(count: usize, attributes: &str) -> Vec<u8> {
        let mut xml = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:e="{X14}" xmlns:f="urn:foreign">"#
        );
        for id in 1..=count {
            xml.push_str(&format!(r#"<connection id="{id}" type="5"><extLst><ext uri="{{D79990A0-CA42-45E3-83F4-45C500A0EAA5}}"><e:connection {attributes}/></ext></extLst></connection>"#));
        }
        xml.push_str("</connections>");
        xml.into_bytes()
    }

    #[test]
    fn custom_data_binding_growth_is_bounded_before_replacement_copies() {
        let bytes = xml(300, r#"embeddedDataId="x""#);
        let (slots, values, unmodeled_context) = parse(&bytes).unwrap();
        let values = Arc::<[String]>::from(values);
        let mut package = OpcPackage::new();
        let name = PackURI::new("/xl/connections.xml").unwrap();
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            CONNECTIONS_CONTENT_TYPE.into(),
            bytes.clone(),
        )));
        let proof = package.source_xml_part(&name).unwrap();
        let relationships = package
            .source_relationships_with_limits(&name, package.read_limits())
            .unwrap();
        let bindings = Bindings {
            source: Some(Arc::new(Source {
                name: PackURI::new("/xl/connections.xml").unwrap(),
                bytes: Arc::new(bytes),
                proof,
                relationships,
                edges: Vec::new(),
                slots,
                values: Arc::clone(&values),
                unmodeled_context,
            })),
            values: Arc::clone(&values),
        };
        assert!(bindings.rebind("x", &"y".repeat(65_535)).is_err());
        assert!(Arc::ptr_eq(&bindings.values, &values));
        assert!(!bindings.is_changed());
    }

    #[test]
    fn custom_data_reference_parser_checks_namespaces_framing_and_entities() {
        assert!(
            parse(&xml(1, r#"f:embeddedDataId="foreign""#))
                .unwrap()
                .1
                .is_empty()
        );
        let escaped = parse(&xml(1, r#"embeddedDataId="_x005F_x0041_""#)).unwrap();
        assert_eq!(escaped.1, ["_x0041_"]);
        for attributes in [
            r#"embeddedDataId="&unknown;""#,
            r#"embeddedDataId="&#0;""#,
            r#"embeddedDataId="_xD800_""#,
            r#"embeddedDataId="one" embeddedDataId="two""#,
            r#"xmlns:g="urn:foreign" f:value="one" g:value="two""#,
            r#"unknown:embeddedDataId="one""#,
        ] {
            assert!(parse(&xml(1, attributes)).is_err(), "accepted {attributes}");
        }
        let source = xml(1, r#"embeddedDataId="x""#);
        let mut dtd = b"<!DOCTYPE connections>".to_vec();
        dtd.extend_from_slice(&source);
        assert!(parse(&dtd).is_err());
        assert!(parse(&source[..source.len() - 1]).is_err());
        let mut roots = source.clone();
        roots.extend_from_slice(&source);
        assert!(parse(&roots).is_err());
        let deep = format!(
            "<connections xmlns='{CORE_NAMESPACE}'>{}{}</connections>",
            "<opaque>".repeat(MAX_DEPTH),
            "</opaque>".repeat(MAX_DEPTH)
        );
        assert!(parse(deep.as_bytes()).is_err());
    }
}
