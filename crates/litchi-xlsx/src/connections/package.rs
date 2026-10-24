//! Workbook package integration for the `SpreadsheetML` connections owner.

use super::codec::{normalize_connections_source_projection, patch_connections_source};
use super::model::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP, Conformance, Connection, Connections,
    MAX_DOM_DEPTH, MAX_DOM_NODES, MAX_XML_BYTES, QUERY_TABLE_CONTENT_TYPE,
    STRICT_CONNECTIONS_RELATIONSHIP, STRICT_NAMESPACE,
};
use super::{codec, invalid};
use crate::error::Result;
use crate::error::{Error as XlsxError, Result as XlsxResult};
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part, SourceLimits};

const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const QUERY_TABLE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
const STRICT_QUERY_TABLE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable";

type GraphResult<T> = std::result::Result<T, GraphError>;

#[derive(Debug)]
enum GraphError {
    Host(XlsxError),
    Invalid(String),
}

impl From<XlsxError> for GraphError {
    fn from(error: XlsxError) -> Self {
        Self::Host(error)
    }
}

impl From<litchi_opc::OpcError> for GraphError {
    fn from(error: litchi_opc::OpcError) -> Self {
        Self::Host(XlsxError::Package(error))
    }
}

impl GraphError {
    fn into_xlsx(self) -> XlsxError {
        match self {
            Self::Host(error) => error,
            Self::Invalid(error) => XlsxError::Invalid(error),
        }
    }
}

macro_rules! graph_invalid {
    ($value:expr $(,)?) => {
        GraphError::Invalid($value.to_string())
    };
}

fn validate_query_table_connection_ids<F>(
    package: &OpcPackage,
    value: &Connections,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let ids = value
        .connections
        .iter()
        .map(|connection| connection.id)
        .collect::<HashSet<_>>();
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        // Only a query-table part's payload is parsed, so only a query-table
        // part's payload is decoded (ADR 0030).
        let connection_id = query_table_connection_id(package.get_part(part.partname())?.blob())?;
        if !ids.contains(&connection_id) {
            return Err(invalid(format!(
                "query-table part '{}' references missing connection ID {}",
                part.partname(),
                connection_id
            )));
        }
    }
    Ok(())
}

fn validate_query_table_connection_ids_with_parser<Q>(
    package: &OpcPackage,
    value: &Connections,
    parse_query_table: &Q,
) -> GraphResult<()>
where
    Q: Fn(&[u8]) -> GraphResult<u32>,
{
    let ids = value
        .connections
        .iter()
        .map(|connection| connection.id)
        .collect::<HashSet<_>>();
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        // Only a query-table part's payload is parsed, so only a query-table
        // part's payload is decoded (ADR 0030).
        let connection_id = parse_query_table(package.get_part(part.partname())?.blob())?;
        if !ids.contains(&connection_id) {
            return Err(GraphError::Invalid(format!(
                "query-table part '{}' references missing connection ID {}",
                part.partname(),
                connection_id
            )));
        }
    }
    Ok(())
}

fn query_table_connection_id(xml: &[u8], limits: &SourceLimits) -> Result<u32> {
    let (
        max_bytes,
        max_nodes,
        max_events,
        max_depth,
        max_string_bytes,
        max_namespace_bytes,
        max_attributes,
        max_temporary_bytes,
    ) = codec::source_profile(limits);
    query_table_connection_id_with_limits(
        xml,
        max_bytes,
        max_nodes,
        max_events,
        max_depth,
        max_string_bytes,
        max_namespace_bytes,
        max_attributes,
        max_temporary_bytes,
    )
}

fn query_table_connection_id_with_limits(
    xml: &[u8],
    max_bytes: usize,
    max_nodes: usize,
    max_events: usize,
    max_depth: usize,
    max_string_bytes: usize,
    max_namespace_bytes: usize,
    max_attributes: usize,
    max_temporary_bytes: usize,
) -> XlsxResult<u32> {
    let maximum = max_bytes.min(MAX_XML_BYTES);
    if xml.len() > maximum {
        return Err(super::limit(LimitKind::XmlBytes, xml.len(), maximum));
    }
    codec::preflight_xml_with_limits(
        xml,
        max_nodes.min(MAX_DOM_NODES),
        max_events,
        max_depth.min(MAX_DOM_DEPTH),
        max_string_bytes,
        max_namespace_bytes,
        max_attributes,
    )?;
    // The MCE processor intentionally handles schema-bearing markup and
    // rejects processing instructions. PIs are legal inert query-table XML;
    // remove only those events for validation while the source snapshot keeps
    // the original bytes for publication.
    let without_pi = codec::strip_processing_instructions_with_limit(xml, max_temporary_bytes)?;
    let depth = max_depth.min(MAX_DOM_DEPTH);
    let mce_output_limit = if codec::contains_mce_markup(without_pi.as_ref()) {
        maximum.min(max_temporary_bytes)
    } else {
        maximum
    };
    let processed = litchi_ooxml_common::mce::process_markup_compatibility(
        without_pi.as_ref(),
        &litchi_ooxml_common::mce::Capabilities::default(),
        &litchi_ooxml_common::mce::Limits {
            max_input_bytes: maximum,
            max_output_bytes: mce_output_limit,
            max_depth: depth,
            max_namespace_bindings: max_attributes.min(4096),
            max_directive_tokens: max_events.min(4096),
            max_choices_per_alternate: max_events.min(1024),
            max_attributes_per_element:
                litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
        },
    )
    .map_err(codec::map_mce_error)?;
    if processed.xml.len() > maximum {
        return Err(super::limit(
            LimitKind::OutputBytes,
            processed.xml.len(),
            maximum,
        ));
    }
    let root = codec::parse_dom_with_limits(
        processed.xml.as_ref(),
        max_nodes.min(MAX_DOM_NODES),
        max_events,
        depth,
        max_string_bytes,
        max_namespace_bytes,
        max_attributes,
    )?;
    codec::expect(&root, "queryTable")?;
    let _name = codec::req(&root, "name")?;
    let connection_id = codec::u32req(&root, "connectionId")?;
    codec::only_unqualified(
        &root,
        &[
            "name",
            "headers",
            "rowNumbers",
            "disableRefresh",
            "backgroundRefresh",
            "firstBackgroundRefresh",
            "refreshOnLoad",
            "growShrinkType",
            "fillFormulas",
            "removeDataOnSave",
            "disableEdit",
            "preserveFormatting",
            "adjustColumnWidth",
            "intermediate",
            "connectionId",
            "autoFormatId",
            "applyNumberFormats",
            "applyBorderFormats",
            "applyFontFormats",
            "applyPatternFormats",
            "applyAlignmentFormats",
            "applyWidthHeightFormats",
        ],
    )?;
    codec::kids(&root)?;
    Ok(connection_id)
}

fn next_connections_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/connections.xml".into()
        } else {
            format!("/xl/connections{suffix}.xml")
        };
        let candidate = PackURI::new(&name).map_err(invalid)?;
        match package.get_part(&candidate) {
            Ok(_) => {},
            Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(candidate),
            Err(error) => return Err(error.into()),
        }
    }
    Err(invalid("no free connections part name"))
}

fn next_connections_relationship_id(package: &OpcPackage, workbook: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdConnections{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free connections relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name.is_equivalent_to(target))
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name.is_equivalent_to(target))
    })
}
pub fn load_from_package(
    package: &OpcPackage,
    limits: &SourceLimits,
) -> Result<Option<Connections>> {
    load_from_package_with_parser(package, &|xml| {
        Connections::parse(xml, limits).map_err(GraphError::from)
    })
    .map_err(GraphError::into_xlsx)
}

fn load_from_package_with_parser<P>(
    package: &OpcPackage,
    parse_connections: &P,
) -> GraphResult<Option<Connections>>
where
    P: Fn(&[u8]) -> GraphResult<Connections>,
{
    let workbook = package.main_document_part()?;
    require_workbook_content_type(workbook)?;
    let mut found = workbook.rels().iter().filter(|x| {
        matches!(
            x.reltype(),
            CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
        )
    });
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(graph_invalid!(
            "workbook has multiple connections relationships"
        ));
    }
    if rel.is_external() || target_has_suffix(rel) {
        return Err(graph_invalid!(
            "connections relationship must target an internal part URI",
        ));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(graph_invalid!(format!(
            "connections part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(graph_invalid!(
            "connections part must not have relationships"
        ));
    }
    Ok(Some(parse_connections(part.blob())?))
}

/// Validate the workbook connection/query-table graph and return its typed
/// connection catalog. No query-table payload is interpreted beyond its
/// inert connection ID, and no target is opened or refreshed.
pub fn validate_graph(package: &OpcPackage, limits: &SourceLimits) -> Result<Option<Connections>> {
    validate_graph_with_parsers(
        package,
        &|xml| Connections::parse(xml, limits).map_err(GraphError::from),
        &|xml| query_table_connection_id(xml, limits).map_err(GraphError::from),
    )
    .map_err(GraphError::into_xlsx)
}

fn validate_graph_with_parsers<CP, QP>(
    package: &OpcPackage,
    parse_connections: &CP,
    parse_query_table: &QP,
) -> GraphResult<Option<Connections>>
where
    CP: Fn(&[u8]) -> GraphResult<Connections>,
    QP: Fn(&[u8]) -> GraphResult<u32>,
{
    if package
        .rels()
        .iter()
        .any(|relationship| is_connections_relationship(relationship.reltype()))
    {
        return Err(graph_invalid!(
            "package root cannot source a workbook connections relationship",
        ));
    }
    let workbook = package.main_document_part()?;
    require_workbook_content_type(workbook)?;
    for part in package
        .iter_parts()
        .filter(|part| !part.partname().is_equivalent_to(workbook.partname()))
    {
        if part
            .rels()
            .iter()
            .any(|relationship| is_connections_relationship(relationship.reltype()))
        {
            return Err(graph_invalid!(
                "only the workbook may source a connections relationship",
            ));
        }
    }
    let owners = workbook
        .rels()
        .iter()
        .filter(|relationship| is_connections_relationship(relationship.reltype()))
        .collect::<Vec<_>>();
    if owners.len() > 1 {
        return Err(graph_invalid!(
            "workbook has multiple connections relationships"
        ));
    }
    let owner_target = owners
        .first()
        .map(|relationship| {
            if relationship.is_external() || target_has_suffix(relationship) {
                return Err(graph_invalid!(
                    "connections relationship must target an internal part URI",
                ));
            }
            Ok(relationship.target_partname()?)
        })
        .transpose()?;
    if let Some(target) = owner_target.as_ref() {
        let part = package.get_part(target)?;
        if part.content_type() != CONNECTIONS_CONTENT_TYPE {
            return Err(graph_invalid!(
                "connections relationship targets an invalid part"
            ));
        }
        if part.rels().iter().next().is_some() {
            return Err(graph_invalid!(
                "connections part must not have relationships"
            ));
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == CONNECTIONS_CONTENT_TYPE)
    {
        if !owner_target
            .as_ref()
            .is_some_and(|target| target.is_equivalent_to(part.partname()))
        {
            return Err(graph_invalid!(format!(
                "connections part '{}' has no workbook owner",
                part.partname()
            )));
        }
    }

    let query_parts = package
        .iter_parts()
        .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
        .collect::<Vec<_>>();
    if package
        .rels()
        .iter()
        .any(|relationship| is_query_table_relationship(relationship.reltype()))
    {
        return Err(graph_invalid!(
            "package root cannot source a query-table relationship",
        ));
    }
    let mut query_owner_counts = HashMap::<String, usize>::new();
    query_owner_counts
        .try_reserve(query_parts.len())
        .map_err(|_| invalid("query-table owner allocation failed"))?;
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if !is_query_table_relationship(relationship.reltype()) {
                continue;
            }
            if source.content_type() != WORKSHEET_CONTENT_TYPE {
                return Err(graph_invalid!(
                    "only worksheet parts may source a query-table relationship",
                ));
            }
            if relationship.is_external() || target_has_suffix(relationship) {
                return Err(graph_invalid!(
                    "query-table relationship must target an internal part URI",
                ));
            }
            let target = relationship.target_partname()?;
            let target_part = package.get_part(&target)?;
            if target_part.content_type() != QUERY_TABLE_CONTENT_TYPE {
                return Err(graph_invalid!(
                    "query-table relationship targets an invalid part"
                ));
            }
            let key = target.as_str().to_ascii_lowercase();
            let count = query_owner_counts.entry(key).or_insert(0);
            *count += 1;
        }
    }
    let values = if owner_target.is_some() {
        Some(
            load_from_package_with_parser(package, parse_connections)?
                .ok_or_else(|| invalid("connections owner disappeared"))?,
        )
    } else {
        None
    };
    if !query_parts.is_empty() && values.is_none() {
        return Err(graph_invalid!(
            "query-table parts require a workbook connections part",
        ));
    }
    if let Some(values) = values.as_ref() {
        validate_query_table_connection_ids_with_parser(package, values, parse_query_table)?;
    }
    for part in query_parts {
        if part.rels().iter().next().is_some() {
            return Err(graph_invalid!(
                "query-table parts must not have relationships"
            ));
        }
        let owners = query_owner_counts
            .get(&part.partname().as_str().to_ascii_lowercase())
            .copied()
            .unwrap_or(0);
        if owners != 1 {
            return Err(graph_invalid!(format!(
                "query-table part '{}' must have exactly one worksheet owner",
                part.partname()
            )));
        }
    }
    Ok(values)
}

/// Validate the connection graph after admitting every XML-bearing member to
/// the Custom Data owner’s caller-selected profile. The bounded closures are
/// used by the complete graph walk, so neither typed connections parsing nor
/// query-table MCE/DOM expansion can bypass the owner limits.
pub(crate) fn validate_graph_with_limits(
    package: &OpcPackage,
    limits: &crate::custom_data::Limits,
) -> XlsxResult<Option<Connections>> {
    for part in package.iter_parts().filter(|part| {
        matches!(
            part.content_type(),
            CONNECTIONS_CONTENT_TYPE | QUERY_TABLE_CONTENT_TYPE
        )
    }) {
        // The owner byte ceiling is charged against the inflated payload, so
        // each connections and query-table part is decoded here (ADR 0030).
        let part = package.get_part(part.partname())?;
        if part.blob().len() > limits.max_connections_bytes() {
            return Err(super::limit(
                LimitKind::XmlBytes,
                part.blob().len(),
                limits.max_connections_bytes(),
            ));
        }
    }
    let parse_connections = |xml: &[u8]| {
        Connections::parse_with_limits(
            xml,
            limits.max_connections_bytes(),
            limits.max_xml_nodes(),
            limits.max_xml_events(),
            limits.max_xml_depth(),
            limits.max_xml_string_bytes(),
            limits.max_xml_namespace_bytes(),
            limits.max_xml_attributes(),
            limits.max_temporary_bytes(),
        )
        .map_err(GraphError::from)
    };
    let parse_query_table = |xml: &[u8]| {
        query_table_connection_id_with_limits(
            xml,
            limits.max_connections_bytes(),
            limits.max_xml_nodes(),
            limits.max_xml_events(),
            limits.max_xml_depth(),
            limits.max_xml_string_bytes(),
            limits.max_xml_namespace_bytes(),
            limits.max_xml_attributes(),
            limits.max_temporary_bytes(),
        )
        .map_err(GraphError::from)
    };
    validate_graph_with_parsers(package, &parse_connections, &parse_query_table)
        .map_err(GraphError::into_xlsx)
}

fn is_connections_relationship(value: &str) -> bool {
    matches!(
        value,
        CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
    )
}

fn is_query_table_relationship(value: &str) -> bool {
    matches!(
        value,
        QUERY_TABLE_RELATIONSHIP | STRICT_QUERY_TABLE_RELATIONSHIP
    )
}

fn target_has_suffix(relationship: &litchi_opc::Relationship) -> bool {
    relationship.target_query().is_some() || relationship.target_fragment().is_some()
}

fn require_workbook_content_type(workbook: &dyn Part) -> Result<()> {
    if matches!(
        workbook.content_type(),
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main document part '{}' is not an XML workbook",
            workbook.partname()
        )))
    }
}

/// Immutable source snapshot used by connection transactions and patches.
#[derive(Clone, Debug, PartialEq)]
pub struct Snapshot {
    connections: Option<Connections>,
    source: SourceState,
    conformance: Conformance,
}

impl Snapshot {
    pub fn load(package: &OpcPackage, limits: &SourceLimits) -> Result<Self> {
        let connections = validate_graph(package, limits)?;
        let source = SourceState::capture(package)?;
        let conformance = if let Some(part) = &source.connection {
            detect_conformance(part.bytes())
        } else {
            detect_conformance(&source.workbook_bytes)
        };
        Ok(Self {
            connections,
            source,
            conformance,
        })
    }

    #[must_use]
    pub fn connections(&self) -> Option<&Connections> {
        self.connections.as_ref()
    }

    pub fn source_xml(&self) -> Option<&[u8]> {
        self.source.connection.as_ref().map(SourcePart::bytes)
    }

    pub fn query_table_xml(&self, part_uri: &PackURI) -> Option<&[u8]> {
        self.source
            .query_tables
            .iter()
            .find(|part| part.part_uri.is_equivalent_to(part_uri))
            .map(SourcePart::bytes)
    }

    pub fn query_table_parts(&self) -> impl Iterator<Item = &PackURI> {
        self.source.query_tables.iter().map(|part| &part.part_uri)
    }

    #[must_use]
    pub fn workbook_part_name(&self) -> &str {
        &self.source.workbook_part_name
    }

    #[must_use]
    pub fn conformance(&self) -> Conformance {
        self.conformance
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.connections.is_none()
    }

    fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

#[derive(Clone, Debug, PartialEq)]
struct SourceState {
    workbook_part_name: String,
    workbook_content_type: String,
    workbook_bytes: Arc<Vec<u8>>,
    workbook_relationships: Vec<SourceRelationship>,
    root_relationships: Vec<SourceRelationship>,
    connection: Option<SourcePart>,
    query_tables: Vec<SourcePart>,
}

impl SourceState {
    fn capture(package: &OpcPackage) -> Result<Self> {
        let workbook = package.main_document_part()?;
        let mut workbook_relationships = workbook
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        workbook_relationships.sort_by(|left, right| left.id.cmp(&right.id));
        let mut root_relationships = package
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        root_relationships.sort_by(|left, right| left.id.cmp(&right.id));
        let connection_uri = workbook
            .rels()
            .iter()
            .find(|relationship| is_connections_relationship(relationship.reltype()))
            .map(litchi_opc::Relationship::target_partname)
            .transpose()?;
        let connection = connection_uri
            .as_ref()
            .map(|uri| package.get_part(uri).map(SourcePart::from_part))
            .transpose()?;
        let query_table_names = package
            .iter_parts()
            .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
            .map(|part| part.partname().clone())
            .collect::<Vec<_>>();
        let mut query_tables = Vec::new();
        for name in &query_table_names {
            query_tables.push(SourcePart::from_part(package.get_part(name)?));
        }
        query_tables.sort_by(|left, right| left.part_uri.as_str().cmp(right.part_uri.as_str()));
        Ok(Self {
            workbook_part_name: workbook.partname().to_string(),
            workbook_content_type: workbook.content_type().to_owned(),
            workbook_bytes: workbook.blob_arc(),
            workbook_relationships,
            root_relationships,
            connection,
            query_tables,
        })
    }

    fn connection_relationship(&self) -> Option<&SourceRelationship> {
        self.workbook_relationships
            .iter()
            .find(|relationship| is_connections_relationship(&relationship.relationship_type))
    }
}

#[derive(Clone, Debug, PartialEq)]
struct SourcePart {
    part_uri: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    relationships: Vec<SourceRelationship>,
}

impl SourcePart {
    fn from_part(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(SourceRelationship::from_relationship)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            part_uri: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            bytes: part.blob_arc(),
            relationships,
        }
    }

    fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }
}

#[derive(Clone, Debug, PartialEq)]
struct SourceRelationship {
    id: String,
    relationship_type: String,
    target: String,
    external: bool,
}

impl SourceRelationship {
    fn from_relationship(relationship: &litchi_opc::Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        }
    }
}

fn detect_conformance(xml: &[u8]) -> Conformance {
    let mut reader = NsReader::from_reader(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                let strict = match reader.resolver().resolve_element(element.name()).0 {
                    ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.0)
                        .is_ok_and(|value| value == STRICT_NAMESPACE),
                    ResolveResult::Unbound | ResolveResult::Unknown(_) => false,
                };
                return if strict {
                    Conformance::Strict
                } else {
                    Conformance::Transitional
                };
            },
            Ok(Event::Eof) | Err(_) => return Conformance::Transitional,
            Ok(_) => {},
        }
    }
}

crate::record::record_family!("connections");

crate::record::limit_kinds!("connections", Connections {
    XmlBytes => "connections XML bytes",
    ExtensionBytes => "connections extension bytes",
    Depth => "connections XML depth",
    Nodes => "connections XML nodes",
    Connections => "connections",
    Parameters => "connection parameters",
    WebTables => "connection web tables",
    TextFields => "connection text fields",
    StringBytes => "connection string bytes",
    TemporaryBytes => "connections temporary bytes",
    OutputBytes => "connections output bytes",
    RelationshipBytes => "connection relationship bytes",
    RelationshipCount => "connection relationships",
    Events => "connections XML events",
    Attributes => "connections XML attributes",
    NamespaceBindings => "connections XML namespace bindings",
    NamespaceBytes => "connections XML namespace bytes",
});

/// A staged, failure-atomic edit of the workbook's inert connection catalog
/// of one [`crate::Package`].
///
/// Starting one retains only the source package and performs no graph read or
/// payload copy. The connection graph is materialized by the first getter,
/// mutation or commit. Dropping the edit discards it; [`Self::commit`]
/// publishes a new package.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    loaded: Option<Loaded>,
}

#[derive(Debug, Clone)]
struct Loaded {
    before: Snapshot,
    draft: Option<Connections>,
    strict: bool,
}

impl crate::Package {
    /// Start an edit of the workbook's connection catalog.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused. Connection-graph, XML and source-limit errors are deferred to
    /// the first getter, mutation or commit.
    pub fn edit_connections(&self) -> Result<Edit> {
        Edit::start(crate::record::Source::new(self, "edit_connections")?)
    }

    /// Apply an exact connections patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_connections(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_connections")
    }
}

impl Edit {
    fn start(target: crate::record::Source) -> Result<Self> {
        Ok(Self {
            target,
            loaded: None,
        })
    }

    fn ensure_loaded(&mut self) -> Result<&mut Loaded> {
        if self.loaded.is_none() {
            let limits = self.target.read_limits();
            let before = Snapshot::load(&self.target, &limits)?;
            let strict = before.conformance.strict();
            self.loaded = Some(Loaded {
                draft: before.connections.clone(),
                before,
                strict,
            });
        }
        Ok(self.loaded.as_mut().expect("connections graph loaded"))
    }

    /// Materialize and return the source snapshot used by this edit.
    pub fn before(&mut self) -> Result<&Snapshot> {
        Ok(&self.ensure_loaded()?.before)
    }

    /// The staged connection catalog, if any.
    pub fn connections(&mut self) -> Result<Option<&Connections>> {
        Ok(self.ensure_loaded()?.draft.as_ref())
    }

    /// Replace the complete catalog; `None` removes the connections part.
    pub fn replace(&mut self, value: Option<Connections>) -> Result<bool> {
        self.ensure_loaded()?;
        let limits = self.target.read_limits();
        validate_draft(&self.target, value.as_ref(), &limits)?;
        if self
            .loaded
            .as_ref()
            .expect("connections graph loaded")
            .draft
            == value
        {
            return Ok(false);
        }
        self.loaded
            .as_mut()
            .expect("connections graph loaded")
            .draft = value;
        Ok(true)
    }

    /// Edit one connection in place; its ID is immutable.
    pub fn edit(
        &mut self,
        id: u32,
        edit: impl FnOnce(&mut Connection) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .ensure_loaded()?
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent connections part"))?;
        let connection = draft
            .connections
            .iter_mut()
            .find(|connection| connection.id == id)
            .ok_or_else(|| invalid(format!("connection ID {id} was not found")))?;
        edit(connection)?;
        if connection.id != id {
            return Err(invalid("connection ID is immutable inside an edit"));
        }
        let limits = self.target.read_limits();
        validate_draft(&self.target, Some(&draft), &limits)?;
        if self
            .loaded
            .as_ref()
            .expect("connections graph loaded")
            .draft
            .as_ref()
            == Some(&draft)
        {
            return Ok(false);
        }
        self.loaded
            .as_mut()
            .expect("connections graph loaded")
            .draft = Some(draft);
        Ok(true)
    }

    /// Add a connection or replace the one with the same ID.
    pub fn set(&mut self, connection: Connection) -> Result<bool> {
        let mut draft = self.ensure_loaded()?.draft.clone().unwrap_or(Connections {
            connections: Vec::new(),
        });
        if let Some(existing) = draft
            .connections
            .iter_mut()
            .find(|existing| existing.id == connection.id)
        {
            if *existing == connection {
                return Ok(false);
            }
            *existing = connection;
        } else {
            draft.add(connection)?;
        }
        let limits = self.target.read_limits();
        validate_draft(&self.target, Some(&draft), &limits)?;
        self.loaded
            .as_mut()
            .expect("connections graph loaded")
            .draft = Some(draft);
        Ok(true)
    }

    /// Remove one connection and return it; removing the last one removes
    /// the connections part.
    pub fn remove(&mut self, id: u32) -> Result<Option<Connection>> {
        let mut draft = self
            .ensure_loaded()?
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot remove from an absent connections part"))?;
        let Some(index) = draft
            .connections
            .iter()
            .position(|connection| connection.id == id)
        else {
            return Ok(None);
        };
        let removed = draft.connections.remove(index);
        if draft.connections.is_empty() {
            let limits = self.target.read_limits();
            validate_draft(&self.target, None, &limits)?;
            self.loaded
                .as_mut()
                .expect("connections graph loaded")
                .draft = None;
        } else {
            let limits = self.target.read_limits();
            validate_draft(&self.target, Some(&draft), &limits)?;
            self.loaded
                .as_mut()
                .expect("connections graph loaded")
                .draft = Some(draft);
        }
        Ok(Some(removed))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        let loaded = self.ensure_loaded()?;
        loaded.draft.clone_from(&loaded.before.connections);
        Ok(())
    }

    fn changed(loaded: &Loaded) -> bool {
        loaded.before.connections.as_ref() != loaded.draft.as_ref()
    }

    /// Validate and publish the staged connections as a new package.
    ///
    /// An unchanged edit returns the source package, so a signed package
    /// keeps its signature byte for byte. A changed commit over a signed
    /// package is refused with [`Error::Signed`]: remove the signature
    /// explicitly first ([`crate::Package::unsign`]).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`], a stale-source error and the publication's
    /// validation errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], keeping the published snapshot.
    fn publish(mut self) -> Result<(Commit, Snapshot)> {
        let unchanged = {
            let loaded = self.ensure_loaded()?;
            !Self::changed(loaded)
        };
        if unchanged {
            let loaded = self.loaded.take().expect("connections graph loaded");
            let patch = Patch::new(loaded.before.clone(), loaded.before.clone());
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                loaded.before,
            ));
        }
        crate::signature_policy::ensure_unsigned(&self.target)?;
        let loaded = self.loaded.take().expect("connections graph loaded");
        let Loaded {
            before,
            draft,
            strict,
        } = loaded;
        let limits = self.target.read_limits();
        let current = Snapshot::load(&self.target, &limits)?;
        if !current.same_source(&before) {
            return Err(invalid("connections edit source is stale"));
        }
        let mut candidate = OpcPackage::clone(&self.target);
        apply_connections(
            &mut candidate,
            before.connections.as_ref(),
            draft.as_ref(),
            strict,
            &limits,
        )?;
        let snapshot = Snapshot::load(&candidate, &limits)?;
        let mut staged = draft.clone();
        if let (Some(actual), Some(staged)) = (snapshot.connections.as_ref(), staged.as_mut()) {
            normalize_connections_source_projection(actual, staged)?;
        }
        if snapshot.connections.as_ref() != staged.as_ref() {
            return Err(invalid("connection publication changed the staged model"));
        }
        let patch = Patch::new(before, snapshot.clone());
        let package = self.target.publish(candidate)?;
        Ok((
            Commit::new(package, patch, Diagnostics::new(true)),
            snapshot,
        ))
    }
}

/// A reversible source-checked connections edit.
#[derive(Clone, Debug, PartialEq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source state required before application.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// The reverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically after checking the connections source graph.
    ///
    /// An empty patch publishes nothing and keeps a signed target exactly as
    /// it was. Any other patch changes the target, so a signed target is
    /// refused up front with [`Error::Signed`]. The inverse never restores a
    /// signature.
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let limits = target.read_limits();
        let current = Snapshot::load(target, &limits)?;
        if !current.same_source(&self.before) {
            return Err(invalid("connections patch source is stale"));
        }
        if self.is_noop() {
            return Ok(());
        }
        crate::signature_policy::ensure_unsigned(target)?;
        let mut candidate = target.clone();
        restore_snapshot(&mut candidate, &self.after, &limits)?;
        let resulting = Snapshot::load(&candidate, &limits)?;
        if !resulting.same_source(&self.after) {
            return Err(invalid("connections patch publication changed its source"));
        }
        *target = candidate;
        Ok(())
    }
}

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}

crate::record::raw_test_route!(edit_connections);

fn validate_draft(
    package: &OpcPackage,
    value: Option<&Connections>,
    limits: &SourceLimits,
) -> Result<()> {
    if let Some(value) = value {
        value.to_xml(false, limits)?;
        validate_query_table_connection_ids(package, value, |xml| {
            query_table_connection_id(xml, limits)
        })?;
    } else if package
        .iter_parts()
        .any(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        return Err(invalid(
            "cannot remove connections while query-table parts remain",
        ));
    }
    Ok(())
}

fn apply_connections(
    package: &mut OpcPackage,
    before: Option<&Connections>,
    after: Option<&Connections>,
    strict: bool,
    limits: &SourceLimits,
) -> Result<()> {
    validate_draft(package, after, limits)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let owner = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| is_connections_relationship(relationship.reltype()))
        .map(|relationship| {
            if relationship.is_external() || target_has_suffix(relationship) {
                return Err(invalid(
                    "connections relationship must target an internal part URI",
                ));
            }
            Ok(relationship
                .target_partname()
                .map(|target| (relationship.r_id().to_owned(), target))?)
        })
        .transpose()?;
    if owner.is_none() && after.is_none() {
        // Nothing to publish: an absent owner stays absent, signed or not.
        return validate_graph(package, limits).map(|_| ());
    }
    crate::signature_policy::ensure_unsigned(package)?;
    match (owner, after) {
        (Some((relationship_id, part_name)), Some(after)) => {
            let source = package.get_part(&part_name)?.blob().to_vec();
            let updated = if let Some(before) = before {
                patch_connections_source(&source, before, after, strict, limits)?
            } else {
                after.to_xml(strict, limits)?
            };
            package.get_part_mut(&part_name)?.set_blob(updated);
            if package.get_part(&part_name)?.rels().iter().next().is_some() {
                return Err(invalid("connections part must not have relationships"));
            }
            let _ = relationship_id;
        },
        (None, Some(after)) => {
            let part_name = next_connections_part_name(package)?;
            let relationship_id = next_connections_relationship_id(package, &workbook_name)?;
            package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
                part_name.clone(),
                CONNECTIONS_CONTENT_TYPE.into(),
                after.to_xml(strict, limits)?,
            )))?;
            package
                .get_part_mut(&workbook_name)?
                .rels_mut()
                .add_relationship(
                    if strict {
                        STRICT_CONNECTIONS_RELATIONSHIP
                    } else {
                        CONNECTIONS_RELATIONSHIP
                    }
                    .into(),
                    part_name.relative_ref(workbook_name.base_uri()),
                    relationship_id,
                    false,
                );
        },
        (Some((relationship_id, part_name)), None) => {
            package
                .get_part_mut(&workbook_name)?
                .rels_mut()
                .remove(&relationship_id);
            if !package_part_is_referenced(package, &part_name) {
                package.remove_part(&part_name);
            }
        },
        (None, None) => {},
    }
    validate_graph(package, limits).map(|_| ())
}

fn restore_snapshot(
    package: &mut OpcPackage,
    snapshot: &Snapshot,
    limits: &SourceLimits,
) -> Result<()> {
    let workbook_name = package.main_document_part()?.partname().clone();
    let existing = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| is_connections_relationship(relationship.reltype()))
        .map(|relationship| {
            if relationship.is_external() || target_has_suffix(relationship) {
                return Err(invalid(
                    "connections relationship must target an internal part URI",
                ));
            }
            Ok(relationship
                .target_partname()
                .map(|target| (relationship.r_id().to_owned(), target))?)
        })
        .transpose()?;
    if let Some((relationship_id, part_name)) = existing {
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .remove(&relationship_id);
        if !package_part_is_referenced(package, &part_name) {
            package.remove_part(&part_name);
        }
    }
    if let Some(part) = &snapshot.source.connection {
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part.part_uri.clone(),
            part.content_type.clone(),
            part.bytes().to_vec(),
        )))?;
        let relationship = snapshot
            .source
            .connection_relationship()
            .ok_or_else(|| invalid("connection snapshot is missing its workbook relationship"))?;
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                relationship.external,
            );
    }
    validate_graph(package, limits).map(|_| ())
}
