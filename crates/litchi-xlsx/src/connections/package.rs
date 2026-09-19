//! Workbook package integration for the `SpreadsheetML` connections owner.

use super::codec::{normalize_connections_source_projection, patch_connections_source};
use super::model::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP, Conformance, Connection, Connections,
    MAX_DOM_DEPTH, MAX_DOM_NODES, MAX_STRING_BYTES, MAX_XML_BYTES, QUERY_TABLE_CONTENT_TYPE,
    STRICT_CONNECTIONS_RELATIONSHIP, STRICT_NAMESPACE,
};
use super::{codec, invalid};
use crate::error::{Error as XlsxError, Result as XlsxResult};
use litchi_core::Resource;
use litchi_core::sheet::Result;
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};

const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const QUERY_TABLE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
const STRICT_QUERY_TABLE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable";

type SheetError = Box<dyn std::error::Error + Send + Sync>;
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

impl From<SheetError> for GraphError {
    fn from(error: SheetError) -> Self {
        Self::Invalid(error.to_string())
    }
}

impl GraphError {
    fn into_xlsx(self) -> XlsxError {
        match self {
            Self::Host(error) => error,
            Self::Invalid(error) => XlsxError::Invalid(error),
        }
    }

    fn into_sheet(self) -> SheetError {
        match self {
            Self::Host(error) => Box::new(error),
            Self::Invalid(error) => invalid(error),
        }
    }
}

macro_rules! graph_invalid {
    ($value:expr $(,)?) => {
        GraphError::Invalid($value.to_string())
    };
}

pub fn store_in_package(package: &mut OpcPackage, value: &Connections, strict: bool) -> Result<()> {
    store_in_package_with_query_table_validator(package, value, strict, query_table_connection_id)
}

/// Store connections while allowing the migration host to retain its complete
/// query-table parser for cross-part validation.
#[doc(hidden)]
pub fn store_in_package_with_query_table_validator<F>(
    package: &mut OpcPackage,
    value: &Connections,
    strict: bool,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let xml = value.to_xml(strict)?;
    validate_query_table_connection_ids(package, value, query_table_connection_id)?;
    // Validate the existing graph after the caller-supplied query-table
    // callback has established the value-level reference invariant. This
    // preserves the callback's diagnostic precedence while rejecting foreign
    // recognized edges and orphan/wrong-content-type parts before mutation.
    validate_graph(package)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let existing = {
        let workbook = package.get_part(&workbook_name)?;
        let mut found = workbook.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
            )
        });
        let first = found
            .next()
            .map(|relationship| {
                if relationship.is_external() || target_has_suffix(relationship) {
                    return Err(invalid(
                        "connections relationship must target an internal part URI",
                    ));
                }
                Ok((
                    relationship.r_id().to_string(),
                    relationship.target_partname()?,
                ))
            })
            .transpose()?;
        if found.next().is_some() {
            return Err(invalid("workbook has multiple connections relationships"));
        }
        first
    };
    if let Some((_, part_name)) = existing {
        let part = package.get_part(&part_name)?;
        if part.content_type() != CONNECTIONS_CONTENT_TYPE {
            return Err(invalid(
                "existing connections part has invalid content type",
            ));
        }
        package.get_part_mut(&part_name)?.set_blob(xml);
    } else {
        let part_name = next_connections_part_name(package)?;
        let relationship_id = next_connections_relationship_id(package, &workbook_name)?;
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            CONNECTIONS_CONTENT_TYPE.into(),
            xml,
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
    }
    package.unsign();
    Ok(())
}

pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    // Refuse to repair or partially remove an invalid graph. In particular,
    // this keeps orphan query-table parts and foreign recognized edges from
    // being hidden by deleting only the workbook owner.
    validate_graph(package)?;
    if package
        .iter_parts()
        .any(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        return Err(invalid(
            "cannot remove connections while query-table parts remain",
        ));
    }
    let workbook_name = package.main_document_part()?.partname().clone();
    let relationship = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
            )
        })
        .map(|relationship| {
            if relationship.is_external() || target_has_suffix(relationship) {
                return Err(invalid(
                    "connections relationship must target an internal part URI",
                ));
            }
            Ok(relationship
                .target_partname()
                .map(|part_name| (relationship.r_id().to_string(), part_name))?)
        })
        .transpose()?;
    let Some((relationship_id, part_name)) = relationship else {
        return Ok(false);
    };
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .remove(&relationship_id);
    if !package_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
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
        let connection_id = query_table_connection_id(part.blob())?;
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
        let connection_id = parse_query_table(part.blob())?;
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

fn query_table_connection_id(xml: &[u8]) -> Result<u32> {
    query_table_connection_id_with_limits(
        xml,
        MAX_XML_BYTES,
        MAX_DOM_NODES,
        MAX_DOM_NODES.saturating_mul(4),
        MAX_DOM_DEPTH,
        MAX_STRING_BYTES,
        MAX_XML_BYTES,
        MAX_DOM_NODES,
        MAX_XML_BYTES,
    )
    .map_err(|error| invalid(error.to_string()))
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
        return Err(XlsxError::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::InputBytes,
            observed: xml.len() as u64,
            limit: maximum as u64,
            scope: Arc::<str>::from("XLSX Custom Data query-table input XML bytes"),
        }));
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
        },
    )
    .map_err(|error| {
        let mapped = codec::map_mce_error(
            error,
            maximum,
            mce_output_limit,
            depth,
            max_attributes.min(4096),
            max_events.min(4096),
            max_events.min(1024),
            "query-table",
        );
        if mce_output_limit < maximum
            && matches!(
                mapped,
                XlsxError::ResourceLimit(litchi_core::ResourceLimit {
                    resource: Resource::OutputBytes,
                    ..
                })
            )
        {
            query_table_temporary_limit(
                mce_output_limit.saturating_add(1),
                mce_output_limit,
                "query-table MCE output bytes",
            )
        } else {
            mapped
        }
    })?;
    if processed.xml.len() > maximum {
        return Err(XlsxError::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::OutputBytes,
            observed: processed.xml.len() as u64,
            limit: maximum as u64,
            scope: Arc::<str>::from("XLSX Custom Data query-table processed XML bytes"),
        }));
    }
    let root = codec::parse_dom_with_limits(
        processed.xml.as_ref(),
        max_nodes.min(MAX_DOM_NODES),
        max_events,
        depth,
        max_string_bytes,
        max_namespace_bytes,
        max_attributes,
    )
    .map_err(|error| XlsxError::Invalid(error.to_string()))?;
    codec::expect(&root, "queryTable").map_err(|error| XlsxError::Invalid(error.to_string()))?;
    let _name = codec::req(&root, "name").map_err(|error| XlsxError::Invalid(error.to_string()))?;
    let connection_id = codec::u32req(&root, "connectionId")
        .map_err(|error| XlsxError::Invalid(error.to_string()))?;
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
    )
    .map_err(|error| XlsxError::Invalid(error.to_string()))?;
    codec::kids(&root).map_err(|error| XlsxError::Invalid(error.to_string()))?;
    Ok(connection_id)
}

fn query_table_temporary_limit(observed: usize, maximum: usize, name: &'static str) -> XlsxError {
    XlsxError::ResourceLimit(litchi_core::ResourceLimit {
        resource: Resource::Memory,
        observed: observed as u64,
        limit: maximum as u64,
        scope: Arc::<str>::from(format!("XLSX Custom Data {name}")),
    })
}

fn next_connections_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/connections.xml".into()
        } else {
            format!("/xl/connections{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
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
pub fn load_from_package(package: &OpcPackage) -> Result<Option<Connections>> {
    load_from_package_with_parser(package, &|xml| {
        Connections::parse(xml).map_err(GraphError::from)
    })
    .map_err(GraphError::into_sheet)
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
pub fn validate_graph(package: &OpcPackage) -> Result<Option<Connections>> {
    validate_graph_with_parsers(
        package,
        &|xml| Connections::parse(xml).map_err(GraphError::from),
        &|xml| query_table_connection_id(xml).map_err(GraphError::from),
    )
    .map_err(GraphError::into_sheet)
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
        if part.blob().len() > limits.max_connections_bytes() {
            return Err(XlsxError::ResourceLimit(litchi_core::ResourceLimit {
                resource: Resource::InputBytes,
                observed: part.blob().len() as u64,
                limit: limits.max_connections_bytes() as u64,
                scope: Arc::<str>::from(format!("XLSX Custom Data graph XML {}", part.partname())),
            }));
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
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let connections = validate_graph(package)?;
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

    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    #[must_use]
    pub fn connections(&self) -> Option<&Connections> {
        self.connections.as_ref()
    }

    #[must_use]
    pub fn catalog(&self) -> Option<&Connections> {
        self.connections()
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
        let mut query_tables = package
            .iter_parts()
            .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
            .map(SourcePart::from_part)
            .collect::<Vec<_>>();
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

/// Failure-atomic edits over the workbook's inert connection catalog.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Option<Connections>,
    strict: bool,
}

impl<'a> Transaction<'a> {
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        let strict = before.conformance.strict();
        Ok(Self {
            draft: before.connections.clone(),
            target,
            before,
            strict,
        })
    }

    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    pub fn connections(&self) -> Option<&Connections> {
        self.draft.as_ref()
    }

    pub fn replace(&mut self, value: Option<Connections>) -> Result<bool> {
        validate_draft(self.target, value.as_ref())?;
        if self.draft == value {
            return Ok(false);
        }
        self.draft = value;
        Ok(true)
    }

    pub fn edit(
        &mut self,
        id: u32,
        edit: impl FnOnce(&mut Connection) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
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
            return Err(invalid("connection ID is immutable inside a transaction"));
        }
        validate_draft(self.target, Some(&draft))?;
        if self.draft.as_ref() == Some(&draft) {
            return Ok(false);
        }
        self.draft = Some(draft);
        Ok(true)
    }

    pub fn set(&mut self, connection: Connection) -> Result<bool> {
        let mut draft = self.draft.clone().unwrap_or(Connections {
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
        validate_draft(self.target, Some(&draft))?;
        self.draft = Some(draft);
        Ok(true)
    }

    pub fn remove(&mut self, id: u32) -> Result<Option<Connection>> {
        let mut draft = self
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
            validate_draft(self.target, None)?;
            self.draft = None;
        } else {
            validate_draft(self.target, Some(&draft))?;
            self.draft = Some(draft);
        }
        Ok(Some(removed))
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.connections.as_ref() != self.draft.as_ref()
    }

    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let current = Snapshot::load(self.target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("connections transaction source is stale"));
        }
        let mut candidate = self.target.clone();
        apply_connections(
            &mut candidate,
            self.before.connections.as_ref(),
            self.draft.as_ref(),
            self.strict,
        )?;
        let snapshot = Snapshot::load(&candidate)?;
        let mut staged = self.draft.clone();
        if let (Some(actual), Some(staged)) = (snapshot.connections.as_ref(), staged.as_mut()) {
            normalize_connections_source_projection(actual, staged)?;
        }
        if snapshot.connections.as_ref() != staged.as_ref() {
            return Err(invalid("connection publication changed the staged model"));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}

/// A reversible source-checked package edit.
#[derive(Clone, Debug, PartialEq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("connections patch source is stale"));
        }
        if self.is_empty() {
            return Ok(());
        }
        let mut candidate = target.clone();
        restore_snapshot(&mut candidate, &self.after)?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) {
            return Err(invalid("connections patch publication changed its source"));
        }
        *target = candidate;
        Ok(())
    }
}

#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

fn validate_draft(package: &OpcPackage, value: Option<&Connections>) -> Result<()> {
    if let Some(value) = value {
        value.to_xml(false)?;
        validate_query_table_connection_ids(package, value, query_table_connection_id)?;
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
) -> Result<()> {
    validate_draft(package, after)?;
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
    match (owner, after) {
        (Some((relationship_id, part_name)), Some(after)) => {
            let source = package.get_part(&part_name)?.blob().to_vec();
            let updated = if let Some(before) = before {
                patch_connections_source(&source, before, after, strict)?
            } else {
                after.to_xml(strict)?
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
                after.to_xml(strict)?,
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
    package.unsign();
    validate_graph(package).map(|_| ())
}

fn restore_snapshot(package: &mut OpcPackage, snapshot: &Snapshot) -> Result<()> {
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
    validate_graph(package).map(|_| ())
}
